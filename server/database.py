"""
database.py - SQLite CRUD 처리
"""
import sqlite3
import os
from typing import List, Dict, Any, Optional
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(__file__), "db", "inspection.db")


def get_conn():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """테이블 초기화 (최초 1회)"""
    conn = get_conn()
    c = conn.cursor()

    c.executescript("""
    CREATE TABLE IF NOT EXISTS incoming_data (
        id              TEXT PRIMARY KEY,
        company_code    TEXT DEFAULT '',
        company_name    TEXT NOT NULL,
        date            TEXT NOT NULL,
        time            TEXT DEFAULT '',
        tm_no           TEXT NOT NULL,
        product_name    TEXT NOT NULL,
        quantity        INTEGER DEFAULT 0,
        pdf_url         TEXT DEFAULT '',
        local_pdf_path  TEXT DEFAULT '',
        created_at      TEXT DEFAULT '',
        created_by      TEXT DEFAULT '',
        updated_at      TEXT DEFAULT '',
        synced_at       TEXT DEFAULT ''
    );

    CREATE TABLE IF NOT EXISTS inspection_results (
        id                  TEXT PRIMARY KEY,
        company_code        TEXT DEFAULT '',
        date                TEXT NOT NULL,
        company_name        TEXT NOT NULL,
        tm_no               TEXT NOT NULL,
        product_name        TEXT DEFAULT '',
        inspection_item     TEXT DEFAULT '',
        inspection_type     TEXT DEFAULT '정량',
        measurement_method  TEXT DEFAULT '',
        lower_limit         TEXT DEFAULT '',
        upper_limit         TEXT DEFAULT '',
        sample1             TEXT DEFAULT '',
        sample2             TEXT DEFAULT '',
        sample3             TEXT DEFAULT '',
        sample4             TEXT DEFAULT '',
        sample5             TEXT DEFAULT '',
        sample6             TEXT DEFAULT '',
        sample7             TEXT DEFAULT '',
        sample8             TEXT DEFAULT '',
        sample9             TEXT DEFAULT '',
        sample10            TEXT DEFAULT '',
        pass_fail_result    TEXT DEFAULT '',
        registered_at       TEXT DEFAULT '',
        registered_by       TEXT DEFAULT '',
        synced_at           TEXT DEFAULT ''
    );

    CREATE TABLE IF NOT EXISTS item_list (
        id                   INTEGER PRIMARY KEY AUTOINCREMENT,
        company_code         TEXT DEFAULT '',
        tm_no                TEXT NOT NULL,
        company_name         TEXT NOT NULL,
        product_name         TEXT DEFAULT '',
        inspection_type      TEXT DEFAULT '검사',
        inspection_standard  TEXT DEFAULT '',
        synced_at            TEXT DEFAULT ''
    );

    CREATE TABLE IF NOT EXISTS sync_log (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        synced_at       TEXT NOT NULL,
        data_count      INTEGER DEFAULT 0,
        result_count    INTEGER DEFAULT 0,
        item_count      INTEGER DEFAULT 0,
        pdf_count       INTEGER DEFAULT 0,
        status          TEXT DEFAULT 'success',
        message         TEXT DEFAULT ''
    );

    CREATE INDEX IF NOT EXISTS idx_incoming_company ON incoming_data(company_name);
    CREATE INDEX IF NOT EXISTS idx_incoming_date    ON incoming_data(date);
    CREATE INDEX IF NOT EXISTS idx_incoming_tmno    ON incoming_data(tm_no);
    CREATE INDEX IF NOT EXISTS idx_result_key       ON inspection_results(date, company_name, tm_no);
    CREATE INDEX IF NOT EXISTS idx_item_company     ON item_list(company_name);
    """)

    conn.commit()

    # 마이그레이션: 기존 DB에 inspection_standard 컬럼 추가 (없을 때만)
    try:
        c.execute("ALTER TABLE item_list ADD COLUMN inspection_standard TEXT DEFAULT ''")
        conn.commit()
    except Exception:
        pass  # 이미 존재하면 무시

    conn.close()


# ─────────────────────────────────────────
#  동기화: 일괄 저장 (upsert)
# ─────────────────────────────────────────

def upsert_incoming_data(rows: List[Dict]) -> int:
    conn = get_conn()
    c = conn.cursor()
    synced_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    count = 0
    for row in rows:
        c.execute("""
            INSERT INTO incoming_data
                (id, company_code, company_name, date, time, tm_no,
                 product_name, quantity, pdf_url, created_at, created_by, updated_at, synced_at)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(id) DO UPDATE SET
                company_code  = excluded.company_code,
                company_name  = excluded.company_name,
                date          = excluded.date,
                time          = excluded.time,
                tm_no         = excluded.tm_no,
                product_name  = excluded.product_name,
                quantity      = excluded.quantity,
                pdf_url       = excluded.pdf_url,
                created_at    = excluded.created_at,
                created_by    = excluded.created_by,
                updated_at    = excluded.updated_at,
                synced_at     = excluded.synced_at
        """, (
            row["id"], row.get("company_code",""), row["company_name"],
            row["date"], row.get("time",""), row["tm_no"],
            row["product_name"], row.get("quantity", 0), row.get("pdf_url",""),
            row.get("created_at",""), row.get("created_by",""),
            row.get("updated_at",""), synced_at
        ))
        count += 1
    conn.commit()
    conn.close()
    return count


def upsert_inspection_results(rows: List[Dict]) -> int:
    conn = get_conn()
    c = conn.cursor()
    synced_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    count = 0
    for row in rows:
        samples = row.get("samples", [])
        while len(samples) < 10:
            samples.append("")
        c.execute("""
            INSERT INTO inspection_results
                (id, company_code, date, company_name, tm_no, product_name,
                 inspection_item, inspection_type, measurement_method,
                 lower_limit, upper_limit,
                 sample1, sample2, sample3, sample4, sample5,
                 sample6, sample7, sample8, sample9, sample10,
                 pass_fail_result, registered_at, registered_by, synced_at)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(id) DO UPDATE SET
                company_code        = excluded.company_code,
                date                = excluded.date,
                company_name        = excluded.company_name,
                tm_no               = excluded.tm_no,
                product_name        = excluded.product_name,
                inspection_item     = excluded.inspection_item,
                inspection_type     = excluded.inspection_type,
                measurement_method  = excluded.measurement_method,
                lower_limit         = excluded.lower_limit,
                upper_limit         = excluded.upper_limit,
                sample1 = excluded.sample1, sample2 = excluded.sample2,
                sample3 = excluded.sample3, sample4 = excluded.sample4,
                sample5 = excluded.sample5, sample6 = excluded.sample6,
                sample7 = excluded.sample7, sample8 = excluded.sample8,
                sample9 = excluded.sample9, sample10 = excluded.sample10,
                pass_fail_result    = excluded.pass_fail_result,
                registered_at       = excluded.registered_at,
                registered_by       = excluded.registered_by,
                synced_at           = excluded.synced_at
        """, (
            row["id"], row.get("company_code",""), row["date"],
            row["company_name"], row["tm_no"], row.get("product_name",""),
            row.get("inspection_item",""), row.get("inspection_type","정량"),
            row.get("measurement_method",""), row.get("lower_limit",""), row.get("upper_limit",""),
            samples[0], samples[1], samples[2], samples[3], samples[4],
            samples[5], samples[6], samples[7], samples[8], samples[9],
            row.get("pass_fail_result",""), row.get("registered_at",""),
            row.get("registered_by",""), synced_at
        ))
        count += 1
    conn.commit()
    conn.close()
    return count


def upsert_item_list(rows: List[Dict]) -> int:
    conn = get_conn()
    c = conn.cursor()
    synced_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # 전체 재동기화: 기존 데이터 삭제 후 삽입
    if rows:
        company_names = list(set(r["company_name"] for r in rows))
        for cn in company_names:
            c.execute("DELETE FROM item_list WHERE company_name=?", (cn,))
    count = 0
    for row in rows:
        c.execute("""
            INSERT INTO item_list
                (company_code, tm_no, company_name, product_name,
                 inspection_type, inspection_standard, synced_at)
            VALUES (?,?,?,?,?,?,?)
        """, (
            row.get("company_code",""), row["tm_no"], row["company_name"],
            row.get("product_name",""), row.get("inspection_type","검사"),
            row.get("inspection_standard",""), synced_at
        ))
        count += 1
    conn.commit()
    conn.close()
    return count


def remove_incoming_duplicates() -> int:
    """
    자연키 (date, company_name, tm_no, quantity, created_at) 기준 중복 제거.
    같은 날짜·업체·TM-NO·수량·등록일시가 동일한 행 중 rowid가 가장 작은 행(최초 입력)만 보존.
    보존할 행에 pdf_url이 없으나 중복 행 중 pdf_url이 있는 경우 pdf_url을 복사 후 삭제.
    """
    conn = get_conn()
    c = conn.cursor()

    # 1단계: pdf_url을 가진 행의 값을 같은 그룹의 대표(min rowid) 행에 복사
    c.execute("""
        UPDATE incoming_data
        SET pdf_url = (
            SELECT i2.pdf_url FROM incoming_data i2
            WHERE i2.date = incoming_data.date
              AND i2.company_name = incoming_data.company_name
              AND i2.tm_no = incoming_data.tm_no
              AND i2.quantity = incoming_data.quantity
              AND COALESCE(NULLIF(i2.created_at,''), '__') = COALESCE(NULLIF(incoming_data.created_at,''), '__')
              AND i2.pdf_url != '' AND i2.pdf_url IS NOT NULL
            LIMIT 1
        )
        WHERE (pdf_url = '' OR pdf_url IS NULL)
          AND rowid = (
            SELECT MIN(rowid) FROM incoming_data i3
            WHERE i3.date = incoming_data.date
              AND i3.company_name = incoming_data.company_name
              AND i3.tm_no = incoming_data.tm_no
              AND i3.quantity = incoming_data.quantity
              AND COALESCE(NULLIF(i3.created_at,''), '__') = COALESCE(NULLIF(incoming_data.created_at,''), '__')
          )
    """)

    # 2단계: 중복 행 삭제 (그룹별 min rowid만 유지)
    c.execute("""
        DELETE FROM incoming_data
        WHERE rowid NOT IN (
            SELECT MIN(rowid)
            FROM incoming_data
            GROUP BY date, company_name, tm_no, quantity,
                     COALESCE(NULLIF(created_at,''), '__')
        )
    """)
    removed = c.rowcount
    conn.commit()
    conn.close()
    return removed


def get_unmatched_pdf_records() -> List[Dict]:
    """pdf_url은 있으나 local_pdf_path가 없는 레코드 반환 (Drive 직접 다운로드용)"""
    conn = get_conn()
    rows = conn.execute(
        """SELECT id, company_name, tm_no, date, pdf_url
           FROM incoming_data
           WHERE pdf_url != '' AND pdf_url IS NOT NULL
             AND (local_pdf_path = '' OR local_pdf_path IS NULL)
           ORDER BY date DESC"""
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def update_local_pdf_path(incoming_id: str, local_path: str):
    conn = get_conn()
    conn.execute(
        "UPDATE incoming_data SET local_pdf_path=? WHERE id=?",
        (local_path, incoming_id)
    )
    conn.commit()
    conn.close()


def add_sync_log(data_count: int, result_count: int, item_count: int,
                 pdf_count: int, status: str = "success", message: str = "") -> int:
    conn = get_conn()
    c = conn.cursor()
    c.execute("""
        INSERT INTO sync_log (synced_at, data_count, result_count, item_count, pdf_count, status, message)
        VALUES (?,?,?,?,?,?,?)
    """, (datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
          data_count, result_count, item_count, pdf_count, status, message))
    row_id = c.lastrowid
    conn.commit()
    conn.close()
    return row_id


def get_last_sync() -> Optional[Dict]:
    conn = get_conn()
    row = conn.execute(
        "SELECT * FROM sync_log ORDER BY id DESC LIMIT 1"
    ).fetchone()
    conn.close()
    return dict(row) if row else None


# ─────────────────────────────────────────
#  검사성적서 이력 조회
# ─────────────────────────────────────────

def search_history(company_name: str = "", date_from: str = "", date_to: str = "",
                   tm_no: str = "", inspection_type: str = "",
                   page: int = 1, page_size: int = 50) -> Dict:
    conn = get_conn()

    # item_list에서 검사형태 map 구성
    item_map = {}
    for r in conn.execute("SELECT company_name, tm_no, inspection_type FROM item_list").fetchall():
        item_map[(r["company_name"], r["tm_no"])] = r["inspection_type"]

    # result map: key=(date, company_name, tm_no) → {pass, fail}
    result_map = {}
    for r in conn.execute(
        "SELECT date, company_name, tm_no, pass_fail_result FROM inspection_results"
    ).fetchall():
        key = (r["date"], r["company_name"], r["tm_no"])
        if key not in result_map:
            result_map[key] = {"pass": 0, "fail": 0, "exists": True}
        if r["pass_fail_result"] == "합격":
            result_map[key]["pass"] += 1
        elif r["pass_fail_result"] == "불합격":
            result_map[key]["fail"] += 1

    # incoming_data 조회
    sql = "SELECT * FROM incoming_data WHERE 1=1"
    params = []
    if company_name:
        sql += " AND company_name LIKE ?"
        params.append(f"%{company_name}%")
    if date_from:
        sql += " AND date >= ?"
        params.append(date_from)
    if date_to:
        sql += " AND date <= ?"
        params.append(date_to)
    if tm_no:
        sql += " AND tm_no LIKE ?"
        params.append(f"%{tm_no}%")
    sql += " ORDER BY date DESC, time ASC"

    rows = conn.execute(sql, params).fetchall()
    conn.close()

    results = []
    for row in rows:
        row = dict(row)
        insp_type = item_map.get((row["company_name"], row["tm_no"]), "검사")
        if inspection_type and insp_type != inspection_type:
            continue

        key = (row["date"], row["company_name"], row["tm_no"])
        ri = result_map.get(key, {"exists": False, "pass": 0, "fail": 0})

        overall = ""
        if ri["exists"]:
            overall = "불합격" if ri["fail"] > 0 else ("합격" if ri["pass"] > 0 else "")

        results.append({
            "id": row["id"],
            "company_code": row["company_code"],
            "company_name": row["company_name"],
            "date": row["date"],
            "time": row["time"],
            "tm_no": row["tm_no"],
            "product_name": row["product_name"],
            "quantity": row["quantity"],
            "pdf_url": row["pdf_url"],
            "local_pdf_path": row["local_pdf_path"],
            "inspection_type": insp_type,
            "has_result": ri["exists"],
            "overall_pass_fail": overall,
        })

    total = len(results)
    start = (page - 1) * page_size
    paged = results[start: start + page_size]

    return {
        "total": total,
        "page": page,
        "page_size": page_size,
        "total_pages": (total + page_size - 1) // page_size if total > 0 else 0,
        "data": paged
    }


def get_inspection_results_by_key(date: str, company_name: str, tm_no: str) -> List[Dict]:
    conn = get_conn()
    rows = conn.execute("""
        SELECT * FROM inspection_results
        WHERE date=? AND company_name=? AND tm_no=?
        ORDER BY registered_at ASC
    """, (date, company_name, tm_no)).fetchall()
    conn.close()
    results = []
    for r in rows:
        r = dict(r)
        r["samples"] = [
            r.pop(f"sample{i}", "") for i in range(1, 11)
        ]
        results.append(r)
    return results


def get_item_list(company_name: str = "", tm_no: str = "") -> List[Dict]:
    """TM-NO 자동완성용 품목 목록 조회"""
    conn = get_conn()
    sql = "SELECT company_code, tm_no, company_name, product_name, inspection_type, inspection_standard FROM item_list WHERE 1=1"
    params = []
    if company_name:
        sql += " AND company_name = ?"
        params.append(company_name)
    if tm_no:
        sql += " AND tm_no LIKE ?"
        params.append(f"%{tm_no}%")
    sql += " ORDER BY tm_no ASC"
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_companies() -> List[str]:
    conn = get_conn()
    rows = conn.execute(
        "SELECT DISTINCT company_name FROM incoming_data ORDER BY company_name"
    ).fetchall()
    conn.close()
    return [r["company_name"] for r in rows]
