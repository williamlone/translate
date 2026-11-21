import argparse
import os

import pyodbc

from translate_blog import load_env_file, q, resolve_env_file, resolve_local_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Inspect translation progress records.")
    parser.add_argument("--engine", default=None, help="engine name to filter by (default: INSPECT_ENGINE)")
    args = parser.parse_args()

    load_env_file(resolve_env_file())
    engine = args.engine or os.environ.get("INSPECT_ENGINE")
    db_path = os.environ.get("BLOG_DB")
    progress_table = os.environ.get("PROGRESS_TABLE")
    if not db_path or not progress_table:
        raise SystemExit("缺少 BLOG_DB / PROGRESS_TABLE 配置，请检查 translate.env")

    driver = os.environ.get("ACCESS_ODBC_DRIVER", "Microsoft Access Driver (*.mdb, *.accdb)")
    conn_str = f"DRIVER={{{driver}}};DBQ={resolve_local_path(db_path)};ReadOnly=True;"
    conn = pyodbc.connect(conn_str)
    try:
        cur = conn.cursor()
        if engine:
            cur.execute(
                f"SELECT run_id, engine, requested, succeeded, failed, status, message, last_success_log_id "
                f"FROM {q(progress_table)} WHERE engine = ? ORDER BY id",
                (engine,),
            )
        else:
            cur.execute(
                f"SELECT run_id, engine, requested, succeeded, failed, status, message, last_success_log_id "
                f"FROM {q(progress_table)} ORDER BY id"
            )
        for row in cur.fetchall():
            print("run:", row[0], row[1], row[2], row[3], row[4], row[5], repr(row[6]), row[7])
    finally:
        conn.close()


if __name__ == "__main__":
    main()
