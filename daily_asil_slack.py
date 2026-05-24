import os
import re
import time
import requests
import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta

SAVE_DIR = Path("data")
REPORT_DIR = Path("reports")
SAVE_DIR.mkdir(parents=True, exist_ok=True)
REPORT_DIR.mkdir(parents=True, exist_ok=True)

SLACK_WEBHOOK_URL = os.getenv("SLACK_WEBHOOK_URL", "")
BASE_URL = "https://realty.asil.kr/api_asil/data_sale_of_member.aspx"
TARGET_GROUPS = ["31평형", "34평형", "39평형"]


def tag(text, pattern):
    return "O" if re.search(pattern, text or "") else ""


def get_deal_category(text):
    text = str(text)
    if "매매" in text:
        return "매매"
    if "전세" in text:
        return "전세"
    if "월세" in text:
        return "월세"
    return "기타"


def collect_asil():
    today = datetime.now().strftime("%Y%m%d")
    save_path = SAVE_DIR / f"아실_매교역팰루시드_전체매물_{today}.xlsx"
    headers = {"User-Agent": "Mozilla/5.0", "Referer": "https://realty.asil.kr/"}

    rows = []
    last_mm_num = 0
    page_count = 0

    while True:
        params = {
            "member": "50074",
            "oidx": "1",
            "oby": "down",
            "total": "20",
            "dealmode": "A01,B01,B02,B03",
            "rlsttype_cd": "A01,B01",
            "hscp_no": "",
            "last_mm_num": str(last_mm_num),
            "mm_uid": ""
        }

        res = requests.get(BASE_URL, params=params, headers=headers, timeout=30)
        res.raise_for_status()
        data = res.json()
        items = data.get("list_result", [])

        if not items:
            break

        for item in items:
            if "매교역팰루시드" not in item.get("BLDNM", ""):
                continue

            memo = item.get("FETR_DESC", "")
            deal_type = item.get("DEALTYPE_NM", "")
            category = get_deal_category(deal_type)

            rows.append({
                "수집일": today,
                "구분": category,
                "매물ID": item.get("mm_uid", ""),
                "단지명": item.get("BLDNM", ""),
                "동": f"{item.get('BDONG_NM', '')}동",
                "공급면적": item.get("SPLY_SPC", ""),
                "전용면적": item.get("EXCLS_SPC", ""),
                "평형": item.get("spc_py_v1", ""),
                "층": item.get("CORES_FLR_CNT_NM") or item.get("CORES_FLR_CNT", ""),
                "실제층": item.get("CORES_FLR_CNT", ""),
                "총층": item.get("TOT_FLR_CNT", ""),
                "거래유형": deal_type,
                "매매가": item.get("DEAL_AMT", ""),
                "보증금/전세가": item.get("WRRNT_AMT", ""),
                "월세": item.get("LEASE_AMT", ""),
                "확인일": item.get("SVC_DATE_STRT", ""),
                "중개사": item.get("BRKG_NM", ""),
                "확장": tag(memo, r"확장|확장형"),
                "시스템에어컨": tag(memo, r"시에|시스템에어컨|에어컨4|에4|풀에어컨|시스템풀"),
                "입주관련": tag(memo, r"입주|입주협의|입주일"),
                "대출": tag(memo, r"대출"),
                "옵션": tag(memo, r"옵션|식세기|냉장고|중문|조합원|붙박이|풀옵션|휴젠트"),
                "뷰": tag(memo, r"뷰|조망|채광|트인|뻥뷰"),
                "비고": memo
            })

        page_count += 1
        print(f"{page_count}페이지 수집 완료 / 누적 {len(rows)}건")

        if not data.get("next_page"):
            break

        last_mm_num += 20
        time.sleep(0.5)

    df = pd.DataFrame(rows)

    for col in ["매매가", "보증금/전세가", "월세"]:
        df[col + "_숫자"] = (
            df[col]
            .astype(str)
            .str.replace(",", "", regex=False)
            .replace(["", "nan", "None"], pd.NA)
        )
        df[col + "_숫자"] = pd.to_numeric(df[col + "_숫자"], errors="coerce")

    sale_df = df[df["구분"] == "매매"].copy()
    jeonse_df = df[df["구분"] == "전세"].copy()
    monthly_df = df[df["구분"] == "월세"].copy()

    with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="전체", index=False)
        sale_df.to_excel(writer, sheet_name="매매", index=False)
        jeonse_df.to_excel(writer, sheet_name="전세", index=False)
        monthly_df.to_excel(writer, sheet_name="월세", index=False)

    print("저장 완료:", save_path)
    return save_path


def load_sheet(file, sheet):
    if not file.exists():
        return None
    try:
        return pd.read_excel(file, sheet_name=sheet)
    except Exception:
        return pd.DataFrame()


def filter_area(df):
    if df is None or df.empty:
        return df
    area = pd.to_numeric(df["전용면적"], errors="coerce")
    return df[area >= 65].copy()


def money(v):
    try:
        v = float(v)
        if pd.isna(v):
            return "-"
        return f"{v / 10000:.1f}억".replace(".0억", "억")
    except Exception:
        return "-"


def money_diff(v):
    if v is None:
        return "데이터 없음"
    try:
        v = float(v)
        sign = "+" if v >= 0 else ""
        return f"{sign}{v / 10000:.1f}억".replace(".0억", "억")
    except Exception:
        return "데이터 없음"


def count_diff(cur, old):
    if old is None:
        return "데이터 없음"
    return f"{cur - old:+}개"


def price_col(deal):
    return "매매가_숫자" if deal == "매매" else "보증금/전세가_숫자"


def group_df(df, group):
    if df is None or df.empty:
        return pd.DataFrame()
    p = pd.to_numeric(df["평형"], errors="coerce")
    target = int(group.replace("평형", ""))
    return df[p == target].copy()


def id_set(df):
    if df is None or df.empty:
        return set()
    return set(df["매물ID"].astype(str))


def item_line(row, deal, before=None, after=None):
    floor = str(row.get("층", ""))
    dong = str(row.get("동", ""))
    memo = str(row.get("비고", ""))

    if before is not None and after is not None:
        price = f"{money(before)}→{money(after)}"
    elif deal == "월세":
        deposit = money(row.get("보증금/전세가_숫자", ""))
        try:
            monthly = f"{int(row.get('월세_숫자', 0))}만"
        except Exception:
            monthly = "-"
        price = f"{deposit}/{monthly}"
    else:
        price = money(row.get(price_col(deal), ""))

    return f"{floor}층_{price}_{dong}_{memo}"


def list_lines(df, deal, before_col=None, after_col=None):
    if df is None or df.empty:
        return []

    marks = ["①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩"]
    lines = []

    for i, (_, r) in enumerate(df.iterrows()):
        mark = marks[i] if i < len(marks) else f"{i + 1}."
        if before_col and after_col:
            lines.append(f"      {mark} {item_line(r, deal, r[before_col], r[after_col])}")
        else:
            lines.append(f"      {mark} {item_line(r, deal)}")

    return lines


def deal_summary_sentence(deal, cur_df, prev7_df, prev1_df):
    cur_cnt = len(cur_df) if cur_df is not None else 0
    prev7_cnt = len(prev7_df) if prev7_df is not None else None

    if prev1_df is not None:
        added_today = len(id_set(cur_df) - id_set(prev1_df))
        removed_today = len(id_set(prev1_df) - id_set(cur_df))
    else:
        added_today = "데이터 없음"
        removed_today = "데이터 없음"

    if prev7_cnt is None:
        return f"{deal}는 현재 {cur_cnt}개이며, 7일 전 데이터가 없어 추세 비교는 불가합니다."

    return f"{deal}는 현재 {cur_cnt}개이며, 7일 전 대비 {cur_cnt - prev7_cnt:+}개입니다. 오늘 추가 {added_today}개, 사라진 매물 {removed_today}개입니다."


def section_for_deal(deal, cur_df, prev7_df, prev1_df, section_no):
    lines = [f"{section_no}. {deal}"]
    pcol = price_col(deal)

    for idx, group in enumerate(TARGET_GROUPS, start=1):
        cur = group_df(cur_df, group)
        prev7 = group_df(prev7_df, group) if prev7_df is not None else None
        prev1 = group_df(prev1_df, group) if prev1_df is not None else None

        cur_cnt = len(cur)
        prev7_cnt = len(prev7) if prev7 is not None else None
        cur_avg = cur[pcol].mean() if cur_cnt else 0
        prev7_avg = prev7[pcol].mean() if prev7 is not None and len(prev7) else None

        lines.append(f"  {idx}) {group}")
        lines.append("    [7일 추세]")
        lines.append(f"    - 현재 매물 : {cur_cnt}개 (7일 전 대비 {count_diff(cur_cnt, prev7_cnt)})")
        lines.append(f"    - 평균가 : {money(cur_avg)} (7일 전 대비 {money_diff(cur_avg - prev7_avg) if prev7_avg is not None else '데이터 없음'})")

        if prev7 is None:
            lines.append("    - 가격 올린 매물 : 데이터 없음")
            lines.append("    - 가격 내린 매물 : 데이터 없음")
        else:
            merged = prev7.merge(cur, on="매물ID", suffixes=("_이전", "_현재"))
            changed = merged[merged[f"{pcol}_이전"] != merged[f"{pcol}_현재"]]
            up = changed[changed[f"{pcol}_현재"] > changed[f"{pcol}_이전"]]
            down = changed[changed[f"{pcol}_현재"] < changed[f"{pcol}_이전"]]

            lines.append(f"    - 가격 올린 매물 : {len(up)}개")
            lines.extend(list_lines(up, deal, f"{pcol}_이전", f"{pcol}_현재"))
            lines.append(f"    - 가격 내린 매물 : {len(down)}개")
            lines.extend(list_lines(down, deal, f"{pcol}_이전", f"{pcol}_현재"))

        lines.append("")
        lines.append("    [오늘 변동]")

        if prev1 is None:
            lines.append("    - 오늘 추가된 매물 : 데이터 없음")
            lines.append("    - 오늘 사라진 매물 : 데이터 없음")
        else:
            cur_ids = id_set(cur)
            prev1_ids = id_set(prev1)
            added = cur[cur["매물ID"].astype(str).isin(cur_ids - prev1_ids)]
            removed = prev1[prev1["매물ID"].astype(str).isin(prev1_ids - cur_ids)]

            lines.append(f"    - 오늘 추가된 매물 : {len(added)}개")
            lines.extend(list_lines(added, deal))
            lines.append(f"    - 오늘 사라진 매물 : {len(removed)}개")
            lines.extend(list_lines(removed, deal))

        lines.append("")

    return lines


def monthly_section(cur_df, prev7_df, prev1_df):
    lines = []
    cur_cnt = len(cur_df) if cur_df is not None else 0
    prev7_cnt = len(prev7_df) if prev7_df is not None else None

    lines.append(f"3. 월세 전체 매물 : {cur_cnt}개 (7일 전 대비 {count_diff(cur_cnt, prev7_cnt)})")

    if prev1_df is not None:
        added = cur_df[cur_df["매물ID"].astype(str).isin(id_set(cur_df) - id_set(prev1_df))]
        removed = prev1_df[prev1_df["매물ID"].astype(str).isin(id_set(prev1_df) - id_set(cur_df))]

        lines.append("  [오늘 변동]")
        lines.append(f"  - 오늘 추가된 월세 : {len(added)}개")
        lines.extend(list_lines(added, "월세"))
        lines.append(f"  - 오늘 사라진 월세 : {len(removed)}개")
        lines.extend(list_lines(removed, "월세"))

    lines.append("")
    lines.append("  [전체 월세 매물]")

    for i, (_, r) in enumerate(cur_df.iterrows(), start=1):
        lines.append(f"  {i}) {r.get('평형', '')}평형_{item_line(r, '월세')}")

    return lines


def split_message(text, max_len=2800):
    lines = text.split("\n")
    chunks = []
    current = ""

    for line in lines:
        if len(current) + len(line) + 1 > max_len:
            chunks.append(current)
            current = line
        else:
            current += ("\n" if current else "") + line

    if current:
        chunks.append(current)

    return chunks


def send_slack(text):
    if not SLACK_WEBHOOK_URL:
        raise Exception("GitHub Secret SLACK_WEBHOOK_URL이 설정되지 않았습니다.")
    res = requests.post(SLACK_WEBHOOK_URL, json={"text": text}, timeout=30)
    res.raise_for_status()


def make_report_and_send(current_file):
    date_str = current_file.stem.split("_")[-1]
    current_date = datetime.strptime(date_str, "%Y%m%d")

    file_7d = SAVE_DIR / f"아실_매교역팰루시드_전체매물_{(current_date - timedelta(days=7)).strftime('%Y%m%d')}.xlsx"
    file_1d = SAVE_DIR / f"아실_매교역팰루시드_전체매물_{(current_date - timedelta(days=1)).strftime('%Y%m%d')}.xlsx"

    result_path = REPORT_DIR / f"아실_매교역팰루시드_텍스트리포트_{date_str}.txt"

    cur_sale = filter_area(load_sheet(current_file, "매매"))
    cur_jeonse = filter_area(load_sheet(current_file, "전세"))
    cur_monthly = filter_area(load_sheet(current_file, "월세"))

    prev7_sale = filter_area(load_sheet(file_7d, "매매"))
    prev7_jeonse = filter_area(load_sheet(file_7d, "전세"))
    prev7_monthly = filter_area(load_sheet(file_7d, "월세"))

    prev1_sale = filter_area(load_sheet(file_1d, "매매"))
    prev1_jeonse = filter_area(load_sheet(file_1d, "전세"))
    prev1_monthly = filter_area(load_sheet(file_1d, "월세"))

    summary = [
        f"📊 매교역팰루시드 매물 리포트 ({date_str})",
        "",
        "□ 요약",
        deal_summary_sentence("매매", cur_sale, prev7_sale, prev1_sale),
        deal_summary_sentence("전세", cur_jeonse, prev7_jeonse, prev1_jeonse),
        deal_summary_sentence("월세", cur_monthly, prev7_monthly, prev1_monthly),
    ]

    messages = [
        "\n".join(summary),
        "\n".join(section_for_deal("매매", cur_sale, prev7_sale, prev1_sale, 1)),
        "\n".join(section_for_deal("전세", cur_jeonse, prev7_jeonse, prev1_jeonse, 2)),
        "\n".join(monthly_section(cur_monthly, prev7_monthly, prev1_monthly)),
    ]

    full_text = "\n\n".join(messages)

    with open(result_path, "w", encoding="utf-8") as f:
        f.write(full_text)

    print(full_text)

    for msg in messages:
        for chunk in split_message(msg):
            send_slack(chunk)
            time.sleep(1)

    print("Slack 전송 완료")
    print("리포트 저장:", result_path)


if __name__ == "__main__":
    current_file = collect_asil()
    make_report_and_send(current_file)
