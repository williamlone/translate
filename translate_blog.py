#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Translate Chinese fields into English in the configured Access database.

Features:
  * Writes results to the configured target table with the same table structure.
  * Resumes from the first record whose translation is missing, failed or stale.
  * Records per-article status and a per-run progress log.
  * Supports qwen-mt-lite / qwen-mt-plus, doubao-seed-translation and DeepSeek.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import os
import re
import sys
import threading
import time
import urllib.error
import urllib.request
import uuid
from collections import Counter, OrderedDict
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

import pyodbc


FIELDS_TO_TRANSLATE = ("log_Title", "log_Intro", "log_Content")
REQUIRED_COLUMNS = {
    "log_ID",
    *FIELDS_TO_TRANSLATE,
    "log_TransHash",
    "log_TransStatus",
    "log_TransModel",
    "log_TransError",
    "log_TransTime",
}

PLACEHOLDER_RE = re.compile(r"__BLOGMT(\d+)__")
TOKEN_RE = re.compile(r"__BLOGMT\d+__")
SENTENCE_SPLIT_RE = re.compile(r"(?<=[。！？!?；;])")
HTML_TAG_RE = re.compile(r"<[^>]+>", re.DOTALL)
ENTITY_RE = re.compile(r"&(?:#\d+|#x[0-9A-Fa-f]+|[A-Za-z][A-Za-z0-9]+);")
URL_RE = re.compile(r"https?://[^\s<>\"']+", re.IGNORECASE)
OPAQUE_BLOCK_RE = re.compile(
    r"<\s*(script|style|pre|code)\b[^>]*>.*?<\s*/\s*\1\s*>",
    re.IGNORECASE | re.DOTALL,
)
PROTECT_RE = re.compile(
    r"(<\s*(script|style|pre|code)\b[^>]*>.*?<\s*/\s*\2\s*>"
    r"|<[^>]+>"
    r"|https?://[^\s<>\"']+)",
    re.IGNORECASE | re.DOTALL,
)

DEFAULT_GLOSSARY: Dict[str, str] = {
    "月光博客": "Moonlight Blog",
    "Google Earth": "Google Earth",
    "Web 2.0": "Web 2.0",
    "SEO": "SEO",
    "QQ": "QQ",
    "MSN": "MSN",
}

DOMAIN_PROMPT = (
    "Translate Chinese into faithful, natural and elegant English for a personal "
    "technology blog. Keep a readable, professional tone. Preserve HTML tags, "
    "attributes, URLs, placeholder tokens, code snippets and technical terms exactly. "
    "Do not add explanations or commentary."
)


class TranslateError(RuntimeError):
    pass


@dataclass
class ArticleResult:
    log_id: int
    ok: bool
    translated: Optional[Dict[str, Any]] = None
    error: Optional[str] = None
    seconds: float = 0.0


class ProgressBar:
    def __init__(self, total: int, enabled: bool = True, width: int = 28):
        self.total = total
        self.width = width
        self.enabled = enabled and total > 0
        self.interactive = bool(getattr(sys.stdout, "isatty", lambda: False)())
        self.count = 0
        self.succeeded = 0
        self.failed = 0
        self.started = time.monotonic()
        self.last_len = 0
        self.lock = threading.Lock()
        if self.enabled:
            self._render()

    def update(self, result: ArticleResult) -> None:
        if not self.enabled:
            return
        with self.lock:
            self.count += 1
            if result.ok:
                self.succeeded += 1
            else:
                self.failed += 1
            self._render()

    def _render(self) -> None:
        pct = self.count / self.total if self.total else 0.0
        filled = int(self.width * pct)
        bar = "#" * filled + "-" * (self.width - filled)
        elapsed = time.monotonic() - self.started
        text = (
            f"[{bar}] {self.count}/{self.total} ({pct * 100:3.0f}%) "
            f"ok={self.succeeded} fail={self.failed} {elapsed:.1f}s"
        )
        if self.interactive:
            padding = " " * max(0, self.last_len - len(text))
            sys.stdout.write("\r" + text + padding)
        else:
            sys.stdout.write(text + "\n")
        sys.stdout.flush()
        self.last_len = len(text)

    def finish(self) -> None:
        if not self.enabled:
            return
        with self.lock:
            if self.interactive:
                self._render()
                sys.stdout.write("\n")
            sys.stdout.flush()


def now() -> dt.datetime:
    return dt.datetime.now()


def q(identifier: str) -> str:
    return "[" + str(identifier).replace("]", "]]") + "]"


def load_env_file(path: Optional[str]) -> None:
    if not path:
        return
    if not os.path.isfile(path):
        raise FileNotFoundError(f"env file not found: {path}")
    with open(path, "r", encoding="utf-8-sig") as fh:
        for raw in fh:
            line = raw.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip()
            if len(value) >= 2 and value[0] == value[-1] and value[0] in ("'", '"'):
                value = value[1:-1]
            if key and key not in os.environ:
                os.environ[key] = value


def resolve_local_path(path: str) -> str:
    if os.path.isabs(path):
        return os.path.abspath(os.path.expanduser(path))
    return os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), path))


def load_glossary(path: Optional[str]) -> Dict[str, str]:
    glossary = dict(DEFAULT_GLOSSARY)
    if not path:
        return glossary
    with open(path, "r", encoding="utf-8-sig") as fh:
        data = json.load(fh)
    if not isinstance(data, dict):
        raise ValueError("glossary file must contain a JSON object")
    glossary.update({str(k): str(v) for k, v in data.items()})
    return glossary


def connect_access(db_path: str, readonly: bool = False) -> pyodbc.Connection:
    driver = os.environ.get("ACCESS_ODBC_DRIVER", "Microsoft Access Driver (*.mdb, *.accdb)")
    abs_path = resolve_local_path(db_path)
    if not os.path.isfile(abs_path):
        raise FileNotFoundError(f"database not found: {abs_path}")
    conn_str = f"DRIVER={{{driver}}};DBQ={abs_path};Timeout=10;"
    if readonly:
        conn_str += "ReadOnly=True;"
    return pyodbc.connect(conn_str, autocommit=False)


def table_exists(cur: pyodbc.Cursor, table: str) -> bool:
    return any(t.table_name.lower() == table.lower() for t in cur.tables(table=table, tableType="TABLE"))


def get_columns(cur: pyodbc.Cursor, table: str) -> List[Any]:
    cols = list(cur.columns(table=table))
    cols.sort(key=lambda c: c.ordinal_position or 0)
    return cols


def column_names(cols: Sequence[Any]) -> List[str]:
    return [c.column_name for c in cols]


def row_to_dict(description: Sequence[Any], row: Sequence[Any]) -> Dict[str, Any]:
    return {d[0]: row[i] for i, d in enumerate(description)}


def access_type(col: Any) -> str:
    type_name = str(col.type_name or "").upper()
    size = int(col.column_size or 0)
    if type_name == "COUNTER":
        return "COUNTER"
    if type_name in {"INTEGER", "INT", "SMALLINT", "TINYINT"}:
        return "INTEGER"
    if type_name in {"VARCHAR", "NVARCHAR", "CHAR", "NCHAR", "TEXT"}:
        return f"VARCHAR({size or 255})"
    if type_name in {"LONGCHAR", "MEMO", "LONGTEXT", "NTEXT"}:
        return "LONGCHAR"
    if type_name in {"DATETIME", "TIMESTAMP"}:
        return "DATETIME"
    if type_name in {"BIT", "BOOLEAN"}:
        return "BIT"
    if type_name in {"DECIMAL", "NUMERIC"}:
        precision = getattr(col, "decimal_digits", None)
        scale = getattr(col, "num_precision_radix", None)
        if precision and scale:
            return f"DECIMAL({precision},{scale})"
        return "DECIMAL(18,2)"
    if type_name == "DOUBLE":
        return "DOUBLE"
    if type_name == "REAL":
        return "REAL"
    if type_name == "CURRENCY":
        return "CURRENCY"
    if size:
        return f"VARCHAR({size})"
    return "LONGCHAR"


def get_index_rows(cur: pyodbc.Cursor, table: str) -> List[Any]:
    return list(cur.statistics(table=table))


def create_target_table(cur: pyodbc.Cursor, source: str, target: str) -> bool:
    if table_exists(cur, target):
        return False

    cols = get_columns(cur, source)
    if not cols:
        raise RuntimeError(f"source table has no columns: {source}")

    stats = get_index_rows(cur, source)
    pk_rows = sorted(
        [s for s in stats if s.index_name and not s.non_unique],
        key=lambda s: s.ordinal_position or 0,
    )

    definitions: List[str] = []
    for col in cols:
        sql_type = access_type(col)
        if sql_type == "COUNTER":
            definitions.append(f"{q(col.column_name)} COUNTER NOT NULL")
        elif not col.nullable:
            definitions.append(f"{q(col.column_name)} {sql_type} NOT NULL")
        else:
            definitions.append(f"{q(col.column_name)} {sql_type}")

    if pk_rows:
        pk_cols = ", ".join(q(s.column_name) for s in pk_rows)
        definitions.append(f"CONSTRAINT {q('PrimaryKey')} PRIMARY KEY ({pk_cols})")

    cur.execute(f"CREATE TABLE {q(target)} ({', '.join(definitions)})")

    grouped: "OrderedDict[str, List[Any]]" = OrderedDict()
    for stat in stats:
        if not stat.index_name or not stat.non_unique:
            continue
        grouped.setdefault(stat.index_name, []).append(stat)

    for index_name, entries in grouped.items():
        entries.sort(key=lambda s: s.ordinal_position or 0)
        unique = "UNIQUE " if entries[0].non_unique == 0 else ""
        direction = {"A": "ASC", "D": "DESC"}
        index_cols = ", ".join(
            f"{q(e.column_name)} {direction.get((e.asc_or_desc or 'A').upper(), 'ASC')}"
            for e in entries
        )
        cur.execute(
            f"CREATE {unique}INDEX {q(index_name)} ON {q(target)} ({index_cols})"
        )
    return True


def create_progress_table(cur: pyodbc.Cursor, table: str) -> bool:
    if table_exists(cur, table):
        return False
    ddl = f"""
        CREATE TABLE {q(table)} (
            [id] COUNTER NOT NULL,
            [run_id] VARCHAR(64) NOT NULL,
            [engine] VARCHAR(64),
            [start_log_id] INTEGER,
            [end_log_id] INTEGER,
            [requested] INTEGER,
            [succeeded] INTEGER,
            [failed] INTEGER,
            [skipped] INTEGER,
            [started_at] DATETIME,
            [finished_at] DATETIME,
            [status] VARCHAR(32),
            [message] LONGCHAR,
            [last_success_log_id] INTEGER,
            [last_success_hash] VARCHAR(64),
            CONSTRAINT {q('PrimaryKey_TransProgress')} PRIMARY KEY ([id])
        )
    """
    cur.execute(ddl)
    return True


def validate_schema(cur: pyodbc.Cursor, source: str, target: str) -> Tuple[List[Any], List[Any]]:
    src_cols = get_columns(cur, source)
    tgt_cols = get_columns(cur, target)
    src_names = set(column_names(src_cols))
    tgt_names = set(column_names(tgt_cols))
    missing = REQUIRED_COLUMNS - src_names
    if missing:
        raise RuntimeError(f"source table lacks required columns: {sorted(missing)}")
    if src_names != tgt_names:
        raise RuntimeError(
            "target table structure differs from source; "
            f"only in source={sorted(src_names - tgt_names)}, only in target={sorted(tgt_names - src_names)}"
        )
    return src_cols, tgt_cols


def hash_values(title: Any, intro: Any, content: Any) -> str:
    payload = json.dumps(
        [title or "", intro or "", content or ""],
        ensure_ascii=False,
        separators=(",", ":"),
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def hash_source_row(row: Dict[str, Any]) -> str:
    return hash_values(
        row.get("log_Title"), row.get("log_Intro"), row.get("log_Content")
    )


def upsert_row(
    conn: pyodbc.Connection,
    cursor: pyodbc.Cursor,
    table: str,
    cols: Sequence[str],
    values: Sequence[Any],
    log_id: Any,
    key_col: str = "log_ID",
) -> None:
    cursor.execute(f"SELECT 1 FROM {q(table)} WHERE {q(key_col)} = ?", (log_id,))
    exists = cursor.fetchone() is not None
    if exists:
        set_cols = [c for c in cols if c != key_col]
        assignments = ", ".join(f"{q(c)} = ?" for c in set_cols)
        by_col = dict(zip(cols, values))
        params = [by_col[c] for c in set_cols] + [log_id]
        cursor.execute(
            f"UPDATE {q(table)} SET {assignments} WHERE {q(key_col)} = ?", params
        )
    else:
        placeholders = ", ".join("?" for _ in cols)
        cursor.execute(
            f"INSERT INTO {q(table)} ({', '.join(q(c) for c in cols)}) VALUES ({placeholders})",
            list(values),
        )
    conn.commit()


def bootstrap_existing_translations(
    conn: pyodbc.Connection,
    cur: pyodbc.Cursor,
    source: str,
    target: str,
    cols: Sequence[Any],
) -> int:
    cur.execute(f"SELECT COUNT(*) FROM {q(target)}")
    if cur.fetchone()[0]:
        return 0

    names = column_names(cols)
    en_cols = [f"{f}_EN" for f in FIELDS_TO_TRANSLATE]
    if not all(c in names for c in en_cols):
        return 0

    cur.execute(
        f"SELECT * FROM {q(source)} WHERE log_TransStatus = ? ORDER BY log_ID", ("ok",)
    )
    description = cur.description
    copied = 0
    for raw in cur.fetchall():
        row = row_to_dict(description, raw)
        translations = {f: row.get(f"{f}_EN") for f in FIELDS_TO_TRANSLATE}
        if not any(translations.values()):
            continue

        output = dict(row)
        for field, translated in translations.items():
            output[field] = translated
            output[f"{field}_EN"] = translated
        output["log_TransHash"] = hash_source_row(row)
        output["log_TransStatus"] = "ok"
        output["log_TransModel"] = f"bootstrap:{row.get('log_TransModel') or 'source'}"
        output["log_TransError"] = None
        output["log_TransTime"] = now()
        upsert_row(
            conn,
            cur,
            target,
            names,
            [output.get(c) for c in names],
            row["log_ID"],
        )
        copied += 1
    return copied


def load_completed(cur: pyodbc.Cursor, target: str) -> Dict[Any, Dict[str, Any]]:
    cur.execute(
        f"SELECT log_ID, log_Title, log_Intro, log_Content, log_TransHash, log_TransStatus "
        f"FROM {q(target)} WHERE log_TransStatus = ?",
        ("ok",),
    )
    description = cur.description
    result: Dict[Any, Dict[str, Any]] = {}
    for raw in cur.fetchall():
        row = row_to_dict(description, raw)
        result[row["log_ID"]] = row
    return result


def completed_matches_source(
    source_row: Dict[str, Any], completed: Optional[Dict[str, Any]], source_hash: str
) -> bool:
    if not completed or completed.get("log_TransStatus") != "ok":
        return False
    if (completed.get("log_TransHash") or "") != source_hash:
        return False
    for field in FIELDS_TO_TRANSLATE:
        if (source_row.get(field) or "") and not (completed.get(field) or ""):
            return False
    return True


def fetch_pending(
    cur: pyodbc.Cursor,
    source: str,
    completed: Dict[Any, Dict[str, Any]],
    limit: int,
) -> List[Tuple[Dict[str, Any], str]]:
    cur.execute(
        f"SELECT log_ID, log_Title, log_Intro, log_Content FROM {q(source)} ORDER BY log_ID"
    )
    ids_and_hashes: List[Tuple[Any, str]] = []
    for raw in cur:
        row_id, title, intro, content = raw
        source_hash = hash_values(title, intro, content)
        lightweight = {
            "log_ID": row_id,
            "log_Title": title,
            "log_Intro": intro,
            "log_Content": content,
        }
        if completed_matches_source(lightweight, completed.get(row_id), source_hash):
            continue
        ids_and_hashes.append((row_id, source_hash))
        if len(ids_and_hashes) >= limit:
            break

    full_rows: List[Tuple[Dict[str, Any], str]] = []
    for row_id, source_hash in ids_and_hashes:
        cur.execute(f"SELECT * FROM {q(source)} WHERE log_ID = ?", (row_id,))
        raw = cur.fetchone()
        if raw is None:
            continue
        row = row_to_dict(cur.description, raw)
        full_rows.append((row, source_hash))
    return full_rows


def post_json(
    url: str,
    payload: Dict[str, Any],
    headers: Dict[str, str],
    timeout: float,
    retries: int,
    min_interval: float,
    rate_lock: threading.Lock,
    rate_state: Dict[str, float],
) -> Dict[str, Any]:
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    all_headers = {"Content-Type": "application/json", **headers}
    last_error = ""

    for attempt in range(retries + 1):
        with rate_lock:
            current = time.monotonic()
            wait = rate_state.get("next") - current
            if wait > 0:
                time.sleep(wait)
            rate_state["next"] = time.monotonic() + min_interval

        request = urllib.request.Request(url, data=body, headers=all_headers, method="POST")
        try:
            with urllib.request.urlopen(request, timeout=timeout) as response:
                raw = response.read().decode("utf-8", "replace")
            data = json.loads(raw)
            if not isinstance(data, dict):
                raise TranslateError(f"unexpected JSON response: {raw[:300]}")
            return data
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", "replace")
            last_error = f"HTTP {exc.code}: {detail[:1000]}"
            if exc.code not in {408, 409, 425, 429, 500, 502, 503, 504} or attempt == retries:
                raise TranslateError(last_error) from exc
        except (urllib.error.URLError, TimeoutError, OSError) as exc:
            last_error = f"{type(exc).__name__}: {exc}"
            if attempt == retries:
                raise TranslateError(last_error) from exc
        except json.JSONDecodeError as exc:
            last_error = f"invalid JSON: {exc}"
            if attempt == retries:
                raise TranslateError(last_error) from exc

        backoff = min(30.0, (2 ** attempt) * 1.5)
        time.sleep(backoff)

    raise TranslateError(last_error or "request failed")


def extract_chat_content(data: Dict[str, Any]) -> str:
    try:
        text = data["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError) as exc:
        raise TranslateError(f"missing choices[0].message.content: {str(data)[:500]}") from exc
    if not isinstance(text, str) or not text.strip():
        raise TranslateError("translation result is empty")
    return text.strip()


class BaseTranslator:
    name = "base"

    def translate(self, text: str, context: Dict[str, Any]) -> str:
        raise NotImplementedError


class QwenTranslator(BaseTranslator):
    name = "qwen-mt"

    def __init__(self, args: argparse.Namespace, glossary: Dict[str, str]):
        self.model = args.qwen_model
        self.base_url = args.qwen_base_url.rstrip("/")
        self.url = f"{self.base_url}/chat/completions"
        self.api_key = os.environ.get("DASHSCOPE_API_KEY", "")
        self.timeout = args.timeout
        self.retries = args.retries
        self.min_interval = args.min_interval
        self.rate_lock = args.rate_lock
        self.rate_state = args.rate_state
        self.translation_options: Dict[str, Any] = {
            "source_lang": "Chinese",
            "target_lang": "English",
            "domains": DOMAIN_PROMPT,
        }
        terms = [{"source": k, "target": v} for k, v in glossary.items()]
        if terms:
            self.translation_options["terms"] = terms

    def translate(self, text: str, context: Dict[str, Any]) -> str:
        payload = {
            "model": self.model,
            "messages": [{"role": "user", "content": text}],
            "translation_options": self.translation_options,
        }
        data = post_json(
            self.url,
            payload,
            {"Authorization": f"Bearer {self.api_key}"},
            self.timeout,
            self.retries,
            self.min_interval,
            self.rate_lock,
            self.rate_state,
        )
        return extract_chat_content(data)


class DoubaoTranslator(BaseTranslator):
    name = "doubao-seed-translation"

    def __init__(self, args: argparse.Namespace, glossary: Dict[str, str]):
        self.model = args.doubao_model
        self.base_url = args.doubao_base_url.rstrip("/")
        self.url = f"{self.base_url}/responses"
        self.api_key = os.environ.get("ARK_API_KEY", "")
        self.timeout = args.timeout
        self.retries = args.retries
        self.min_interval = args.min_interval
        self.rate_lock = args.rate_lock
        self.rate_state = args.rate_state

    def translate(self, text: str, context: Dict[str, Any]) -> str:
        payload = {
            "model": self.model,
            "input": [
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "input_text",
                            "text": text,
                            "translation_options": {
                                "source_language": "zh",
                                "target_language": "en",
                            },
                        }
                    ],
                }
            ],
        }
        data = post_json(
            self.url,
            payload,
            {"Authorization": f"Bearer {self.api_key}"},
            self.timeout,
            self.retries,
            self.min_interval,
            self.rate_lock,
            self.rate_state,
        )
        for item in data.get("output", []):
            if not isinstance(item, dict) or item.get("type") != "message":
                continue
            for part in item.get("content", []):
                if isinstance(part, dict) and part.get("type") == "output_text":
                    translated = part.get("text")
                    if isinstance(translated, str) and translated.strip():
                        return translated.strip()
        raise TranslateError(f"cannot find output_text in Doubao response: {str(data)[:500]}")


def deepseek_api_key(base_url: str) -> str:
    normalized = (base_url or "").lower()
    if "volces.com" in normalized or "ark.cn-beijing" in normalized:
        return os.environ.get("ARK_API_KEY", "")
    return os.environ.get("DEEPSEEK_API_KEY", "")


class DeepSeekTranslator(BaseTranslator):
    name = "deepseek"

    def __init__(self, args: argparse.Namespace, glossary: Dict[str, str]):
        self.model = args.deepseek_model
        self.base_url = args.deepseek_base_url.rstrip("/")
        self.url = f"{self.base_url}/chat/completions"
        self.api_key = deepseek_api_key(self.base_url)
        self.timeout = args.timeout
        self.retries = args.retries
        self.min_interval = args.min_interval
        self.rate_lock = args.rate_lock
        self.rate_state = args.rate_state
        self.max_tokens = args.deepseek_max_tokens
        self.disable_thinking = args.deepseek_disable_thinking
        glossary_lines = "\n".join(f"- {k} => {v}" for k, v in glossary.items())
        self.system_prompt_template = (
            "You are a professional Chinese-to-English translator for a personal IT blog. "
            "Translate faithfully (信), express naturally (达), and write elegantly (雅). "
            "Keep the original meaning, tone and paragraph structure. Use idiomatic English. "
            "Preserve HTML tags, attributes, URLs, numbers, code snippets and every placeholder "
            "of the form __BLOGMTn__ exactly. A placeholder may represent only one character or "
            "a piece of punctuation; never omit, merge or rewrite it even if it looks insignificant. "
            "Do not add explanations, notes or summaries.\n"
            "Use these fixed translations where applicable:\n"
            f"{glossary_lines}"
        )

    def translate(self, text: str, context: Dict[str, Any]) -> str:
        title = (context.get("log_Title") or "").strip()
        system_prompt = self.system_prompt_template
        if title:
            system_prompt += f"\nArticle title context: {title}"
        missing = context.get("missing_placeholders")
        if missing:
            system_prompt += (
                "\nCRITICAL RETRY: your previous output omitted protected placeholder token(s): "
                + ", ".join(str(token) for token in missing)
                + ". Return the COMPLETE translation again and include EVERY placeholder "
                "exactly once, in the same form as in the input. Never omit, merge, renumber "
                "or rewrite them, even if they look insignificant."
            )
        unexpected = context.get("unexpected_placeholders")
        if unexpected:
            system_prompt += (
                "\nCRITICAL RETRY: your previous output contained placeholder token(s) that are "
                "not in the input: "
                + ", ".join(str(token) for token in unexpected)
                + ". Remove them completely. Use ONLY the placeholders present in the input, "
                "each exactly once. Never invent, duplicate or renumber them."
            )
        payload = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": text},
            ],
            "temperature": 0.2,
            "max_tokens": self.max_tokens,
            "stream": False,
        }
        if self.disable_thinking:
            payload["thinking"] = {"type": "disabled"}
        data = post_json(
            self.url,
            payload,
            {"Authorization": f"Bearer {self.api_key}"},
            self.timeout,
            self.retries,
            self.min_interval,
            self.rate_lock,
            self.rate_state,
        )
        return extract_chat_content(data)


def build_translator(args: argparse.Namespace, glossary: Dict[str, str]) -> BaseTranslator:
    engine = args.engine.lower()
    if engine in {"qwen", "qwen-mt", "qwen-mt-lite", "qwen-mt-plus"}:
        if not os.environ.get("DASHSCOPE_API_KEY"):
            raise RuntimeError("qwen engine requires DASHSCOPE_API_KEY")
        return QwenTranslator(args, glossary)
    if engine in {"doubao", "doubao-seed-translation"}:
        if not os.environ.get("ARK_API_KEY"):
            raise RuntimeError("doubao engine requires ARK_API_KEY")
        return DoubaoTranslator(args, glossary)
    if engine == "deepseek":
        if not deepseek_api_key(args.deepseek_base_url):
            raise RuntimeError(
                "deepseek engine requires DEEPSEEK_API_KEY "
                "(or ARK_API_KEY when --deepseek-base-url points to Volcengine Ark)"
            )
        return DeepSeekTranslator(args, glossary)
    raise RuntimeError(f"unsupported engine: {args.engine}")


def protect_html(text: str) -> Tuple[str, Dict[str, str]]:
    mapping: Dict[str, str] = {}

    def mask(match: re.Match) -> str:
        token = f"__BLOGMT{len(mapping)}__"
        mapping[token] = match.group(0)
        return token

    masked = PROTECT_RE.sub(mask, text)
    return masked, mapping


def restore_html(text: str, mapping: Dict[str, str]) -> str:
    def canonicalize(match: re.Match) -> str:
        return f"__BLOGMT{int(match.group(1))}__"

    text = PLACEHOLDER_RE.sub(canonicalize, text)
    missing = [token for token in mapping if token not in text]
    if missing:
        raise TranslateError(f"translator dropped placeholder tokens: {missing[:3]}")

    restored = text
    for token, original in mapping.items():
        restored = restored.replace(token, original)

    leftover = PLACEHOLDER_RE.findall(restored)
    if leftover:
        raise TranslateError(f"unresolved placeholder tokens: {leftover[:3]}")
    return restored


def split_avoiding_tokens(text: str, max_chars: int) -> List[str]:
    if len(text) <= max_chars:
        return [text]
    pieces: List[str] = []
    current: List[str] = []
    current_len = 0
    index = 0
    for match in TOKEN_RE.finditer(text):
        start, end = match.span()
        if start > index:
            plain = text[index:start]
            for char in plain:
                if current_len >= max_chars:
                    pieces.append("".join(current))
                    current, current_len = [], 0
                current.append(char)
                current_len += 1
        token = match.group(0)
        if current and current_len + len(token) > max_chars:
            pieces.append("".join(current))
            current, current_len = [], 0
        current.append(token)
        current_len += len(token)
        index = end
    if index < len(text):
        for char in text[index:]:
            if current_len >= max_chars:
                pieces.append("".join(current))
                current, current_len = [], 0
            current.append(char)
            current_len += 1
    if current:
        pieces.append("".join(current))
    return pieces


def split_text(text: str, max_chars: int) -> List[str]:
    if len(text) <= max_chars:
        return [text]

    sentences = SENTENCE_SPLIT_RE.split(text)
    grouped: List[str] = []
    buffer = ""
    for sentence in sentences:
        if not sentence:
            continue
        if len(sentence) > max_chars:
            if buffer:
                grouped.append(buffer)
                buffer = ""
            grouped.extend(split_avoiding_tokens(sentence, max_chars))
            continue
        if buffer and len(buffer) + len(sentence) > max_chars:
            grouped.append(buffer)
            buffer = sentence
        else:
            buffer += sentence
    if buffer:
        grouped.append(buffer)
    return grouped or [text]


def split_edge_tokens(chunk: str) -> Tuple[List[str], str, List[str]]:
    matches = list(TOKEN_RE.finditer(chunk))
    leading: List[str] = []
    pos = 0
    for match in matches:
        if match.start() != pos:
            break
        leading.append(match.group(0))
        pos = match.end()
    if pos == len(chunk):
        return leading, "", []
    trailing: List[str] = []
    end = len(chunk)
    for match in reversed(matches):
        if match.end() != end:
            break
        trailing.insert(0, match.group(0))
        end = match.start()
    return leading, chunk[pos:end], trailing


def translate_chunk_checked(
    chunk: str,
    translator: BaseTranslator,
    context: Dict[str, Any],
    attempts: int = 5,
) -> str:
    leading, core, trailing = split_edge_tokens(chunk)
    if not core.strip():
        return chunk
    expected = Counter(f"__BLOGMT{int(n)}__" for n in PLACEHOLDER_RE.findall(core))
    last_error = "placeholder tokens do not match source"
    retry_context = dict(context)
    for _ in range(attempts):
        result = translator.translate(core, retry_context)
        actual = Counter(f"__BLOGMT{int(n)}__" for n in PLACEHOLDER_RE.findall(result))
        missing = expected - actual
        unexpected = actual - expected
        if not missing and not unexpected:
            return "".join(leading) + result + "".join(trailing)

        problems: List[str] = []
        if missing:
            problems.append(
                f"translator dropped placeholder tokens: {sorted(missing.elements())[:3]}"
            )
            retry_context["missing_placeholders"] = sorted(missing.elements())
        else:
            retry_context.pop("missing_placeholders", None)
        if unexpected:
            problems.append(
                f"translator added placeholder tokens: {sorted(unexpected.elements())[:3]}"
            )
            retry_context["unexpected_placeholders"] = sorted(unexpected.elements())
        else:
            retry_context.pop("unexpected_placeholders", None)
        last_error = "; ".join(problems)
    raise TranslateError(f"{last_error} (after {attempts} attempts)")


def translate_text(
    text: str,
    translator: BaseTranslator,
    context: Dict[str, Any],
    chunk_chars: int,
    workers: int,
) -> str:
    chunks = split_text(text, chunk_chars)
    if not chunks:
        return text
    if len(chunks) == 1 or workers <= 1:
        return "".join(translate_chunk_checked(chunk, translator, context) for chunk in chunks)

    with ThreadPoolExecutor(max_workers=min(workers, len(chunks))) as pool:
        futures = [
            pool.submit(translate_chunk_checked, chunk, translator, context)
            for chunk in chunks
        ]
        return "".join(future.result() for future in futures)


def translate_field(
    value: Any,
    translator: BaseTranslator,
    context: Dict[str, Any],
    chunk_chars: int,
    workers: int,
) -> Any:
    if value is None or value == "":
        return value
    source = str(value)
    masked, mapping = protect_html(source)
    if not masked.strip():
        return source
    translated = translate_text(masked, translator, context, chunk_chars, workers)
    return restore_html(translated, mapping)


def translate_article(
    row: Dict[str, Any],
    translator: BaseTranslator,
    args: argparse.Namespace,
) -> ArticleResult:
    started = time.monotonic()
    context = {"log_Title": row.get("log_Title")}
    try:
        translated: Dict[str, Any] = {}
        for field in FIELDS_TO_TRANSLATE:
            translated[field] = translate_field(
                row.get(field),
                translator,
                context,
                args.chunk_chars,
                args.workers,
            )
        return ArticleResult(
            log_id=row["log_ID"],
            ok=True,
            translated=translated,
            seconds=time.monotonic() - started,
        )
    except Exception as exc:
        return ArticleResult(
            log_id=row["log_ID"],
            ok=False,
            error=format_error(exc),
            seconds=time.monotonic() - started,
        )


def format_error(exc: BaseException) -> str:
    return f"{type(exc).__name__}: {exc}"[:2000]


def build_output_row(
    source_row: Dict[str, Any],
    translated: Optional[Dict[str, Any]],
    source_hash: str,
    engine: str,
    status: str,
    error: Optional[str],
    cols: Sequence[str],
) -> List[Any]:
    output = dict(source_row)
    if translated:
        for field, value in translated.items():
            output[field] = value
            en_field = f"{field}_EN"
            if en_field in output:
                output[en_field] = value
    output["log_TransHash"] = source_hash
    output["log_TransStatus"] = status
    output["log_TransModel"] = engine
    output["log_TransError"] = error
    output["log_TransTime"] = now()
    return [output.get(c) for c in cols]


def start_run(
    conn: pyodbc.Connection,
    cur: pyodbc.Cursor,
    progress_table: str,
    run_id: str,
    engine: str,
    pending: Sequence[Tuple[Dict[str, Any], str]],
) -> None:
    first_id = pending[0][0]["log_ID"] if pending else None
    last_id = pending[-1][0]["log_ID"] if pending else None
    cur.execute(
        f"INSERT INTO {q(progress_table)} "
        "(run_id, engine, start_log_id, end_log_id, requested, succeeded, failed, skipped, "
        "started_at, status, message, last_success_log_id, last_success_hash) "
        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
        (
            run_id,
            engine,
            first_id,
            last_id,
            len(pending),
            0,
            0,
            0,
            now(),
            "running",
            None,
            None,
            None,
        ),
    )
    conn.commit()


def finish_run(
    conn: pyodbc.Connection,
    cur: pyodbc.Cursor,
    progress_table: str,
    run_id: str,
    succeeded: int,
    failed: int,
    status: str,
    message: str,
    last_success_id: Optional[int],
    last_success_hash: Optional[str],
) -> None:
    cur.execute(
        f"UPDATE {q(progress_table)} SET "
        "succeeded=?, failed=?, finished_at=?, status=?, message=?, "
        "last_success_log_id=?, last_success_hash=? WHERE run_id=?",
        (
            succeeded,
            failed,
            now(),
            status,
            message[:2000],
            last_success_id,
            last_success_hash,
            run_id,
        ),
    )
    conn.commit()


def translate_pending(
    pending: Sequence[Tuple[Dict[str, Any], str]],
    translator: BaseTranslator,
    article_workers: int,
    on_result: Optional[Callable[[ArticleResult], None]] = None,
) -> Dict[Any, ArticleResult]:
    results: Dict[Any, ArticleResult] = {}
    if article_workers <= 1 or len(pending) <= 1:
        for row, _ in pending:
            result = translate_article(row, translator, _ARTICLE_ARGS)
            results[row["log_ID"]] = result
            if on_result:
                on_result(result)
        return results

    with ThreadPoolExecutor(max_workers=min(article_workers, len(pending))) as pool:
        futures = {pool.submit(translate_article, row, translator, _ARTICLE_ARGS): row["log_ID"] for row, _ in pending}
        for future in as_completed(futures):
            result = future.result()
            results[result.log_id] = result
            if on_result:
                on_result(result)
    return results


_ARTICLE_ARGS: argparse.Namespace


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Translate the configured source table into the target table in the same Access database."
    )
    parser.add_argument("--db", default=os.environ.get("BLOG_DB"))
    parser.add_argument("--source-table", default=os.environ.get("SOURCE_TABLE"))
    parser.add_argument("--target-table", default=os.environ.get("TARGET_TABLE"))
    parser.add_argument("--progress-table", default=os.environ.get("PROGRESS_TABLE"))
    parser.add_argument(
        "--engine",
        default=os.environ.get("TRANSLATE_ENGINE", "qwen-mt-lite"),
        choices=["qwen-mt-lite", "qwen-mt-plus", "doubao-seed-translation", "deepseek"],
        help="translation engine",
    )
    parser.add_argument("--limit", type=int, default=10, help="number of articles to translate")
    parser.add_argument("--env-file", default=os.environ.get("ENV_FILE"))
    parser.add_argument("--glossary", default=os.environ.get("GLOSSARY_FILE"))
    parser.add_argument("--workers", type=int, default=4, help="parallel chunks per article")
    parser.add_argument("--article-workers", type=int, default=1, help="parallel articles in one batch")
    parser.add_argument("--chunk-chars", type=int, default=600, help="max characters per API request")
    parser.add_argument("--timeout", type=float, default=120.0)
    parser.add_argument("--retries", type=int, default=4)
    parser.add_argument("--min-interval", type=float, default=0.0, help="minimum seconds between requests")
    parser.add_argument("--dry-run", action="store_true", help="show pending records without calling APIs")
    parser.add_argument("--init-only", action="store_true", help="create tables and copy existing translations only")

    parser.add_argument("--qwen-model", default=os.environ.get("QWEN_MODEL"))
    parser.add_argument(
        "--qwen-base-url",
        default=os.environ.get("QWEN_BASE_URL", "https://dashscope.aliyuncs.com/compatible-mode/v1"),
    )
    parser.add_argument(
        "--doubao-model",
        default=os.environ.get("DOUBAO_MODEL", "doubao-seed-translation-250915"),
    )
    parser.add_argument(
        "--doubao-base-url",
        default=os.environ.get("DOUBAO_BASE_URL", "https://ark.cn-beijing.volces.com/api/v3"),
    )
    parser.add_argument("--deepseek-model", default=os.environ.get("DEEPSEEK_MODEL", "deepseek-chat"))
    parser.add_argument(
        "--deepseek-base-url",
        default=os.environ.get("DEEPSEEK_BASE_URL", "https://api.deepseek.com"),
    )
    parser.add_argument(
        "--deepseek-max-tokens", type=int, default=int(os.environ.get("DEEPSEEK_MAX_TOKENS", "4096"))
    )
    parser.add_argument(
        "--deepseek-disable-thinking",
        action="store_true",
        default=os.environ.get("DEEPSEEK_DISABLE_THINKING", "").lower() in ("1", "true", "yes"),
        help="disable reasoning tokens (needed for Ark reasoning models such as DeepSeek-V4.1-Flash)",
    )

    args = parser.parse_args(argv)
    if args.limit < 1:
        parser.error("--limit must be >= 1")
    if args.workers < 1 or args.article_workers < 1:
        parser.error("--workers and --article-workers must be >= 1")
    if args.chunk_chars < 100:
        parser.error("--chunk-chars must be >= 100")
    return args


def resolve_env_file(argv: Optional[Sequence[str]] = None) -> Optional[str]:
    env_file = os.environ.get("ENV_FILE")
    values = list(argv if argv is not None else sys.argv[1:])
    for index, value in enumerate(values):
        if value == "--env-file" and index + 1 < len(values):
            env_file = values[index + 1]
        elif value.startswith("--env-file="):
            env_file = value.split("=", 1)[1]
    if env_file:
        return env_file
    default_env = os.path.join(os.path.dirname(os.path.abspath(__file__)), "translate.env")
    return default_env if os.path.isfile(default_env) else None


def main(argv: Optional[Sequence[str]] = None) -> int:
    global _ARTICLE_ARGS
    load_env_file(resolve_env_file(argv))
    args = parse_args(argv)
    _ARTICLE_ARGS = args

    missing = [
        name
        for name, value in (
            ("BLOG_DB", args.db),
            ("SOURCE_TABLE", args.source_table),
            ("TARGET_TABLE", args.target_table),
            ("PROGRESS_TABLE", args.progress_table),
        )
        if not value
    ]
    if missing:
        print(
            "缺少必要配置: " + ", ".join(missing) + "；请复制 config.example.env 为 translate.env 并填写",
            file=sys.stderr,
        )
        return 1

    glossary = load_glossary(args.glossary)

    if args.engine.startswith("qwen") and not args.qwen_model:
        args.qwen_model = args.engine

    args.rate_lock = threading.Lock()
    args.rate_state = {"next": 0.0}

    translator: Optional[BaseTranslator] = None
    if not args.dry_run:
        translator = build_translator(args, glossary)

    conn = connect_access(args.db)
    try:
        cur = conn.cursor()
        created_target = create_target_table(cur, args.source_table, args.target_table)
        create_progress_table(cur, args.progress_table)
        conn.commit()

        src_cols, tgt_cols = validate_schema(cur, args.source_table, args.target_table)
        names = column_names(tgt_cols)
        copied = bootstrap_existing_translations(
            conn, cur, args.source_table, args.target_table, src_cols
        )
        if copied:
            print(f"Bootstrapped {copied} existing translated rows from source.")

        completed = load_completed(cur, args.target_table)
        pending = fetch_pending(cur, args.source_table, completed, args.limit)

        if not pending:
            print("No pending records. Translation progress is complete.")
            return 0

        id_list = ", ".join(str(row["log_ID"]) for row, _ in pending)
        print(f"Pending {len(pending)} record(s), log_ID: {id_list}")

        if args.dry_run or args.init_only:
            print("Dry-run / init-only: no translation was called.")
            return 0

        run_id = uuid.uuid4().hex
        engine_name = translator.name if translator else args.engine
        start_run(conn, cur, args.progress_table, run_id, engine_name, pending)

        progress = ProgressBar(len(pending))
        try:
            results = translate_pending(
                pending, translator, args.article_workers, on_result=progress.update
            )
        finally:
            progress.finish()

        succeeded = 0
        failed = 0
        last_success_id: Optional[int] = None
        last_success_hash: Optional[str] = None

        for row, source_hash in pending:
            row_id = row["log_ID"]
            result = results.get(row_id)
            if result is None:
                result = ArticleResult(row_id, False, error="missing translation result")

            if result.ok:
                values = build_output_row(
                    row, result.translated, source_hash, engine_name, "ok", None, names
                )
                status = "ok"
                succeeded += 1
                last_success_id = row_id
                last_success_hash = source_hash
                print(
                    f"[OK] log_ID={row_id} ({result.seconds:.1f}s) -> "
                    f"{(result.translated or {}).get('log_Title', '')[:80]}"
                )
            else:
                values = build_output_row(
                    row, None, source_hash, engine_name, "failed", result.error, names
                )
                status = "failed"
                failed += 1
                print(f"[FAIL] log_ID={row_id}: {result.error}")

            upsert_row(
                conn,
                cur,
                args.target_table,
                names,
                values,
                row_id,
            )

        message = f"{engine_name}: requested={len(pending)}, succeeded={succeeded}, failed={failed}"
        finish_run(
            conn,
            cur,
            args.progress_table,
            run_id,
            succeeded,
            failed,
            "ok" if failed == 0 else "partial",
            message,
            last_success_id,
            last_success_hash,
        )
        print(message)
        return 0 if failed == 0 else 2
    except Exception as exc:
        print(f"Fatal error: {format_error(exc)}", file=sys.stderr)
        return 1
    finally:
        conn.close()


if __name__ == "__main__":
    raise SystemExit(main())
