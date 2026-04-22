import dart_fss as dart
import requests, zipfile, io
from lxml import etree
import pandas as pd
import os

API_KEY = os.getenv("DART_API_KEY")
dart.set_api_key(api_key=API_KEY)


def parse_table_with_span(table):
    rows = table.xpath(".//tr")
    grid = []
    span_map = {}

    for r_idx, tr in enumerate(rows):
        cells = tr.xpath("./th|./td")
        row = []
        c_idx = 0

        while (r_idx, c_idx) in span_map:
            row.append(span_map[(r_idx, c_idx)])
            c_idx += 1

        for cell in cells:
            while (r_idx, c_idx) in span_map:
                row.append(span_map[(r_idx, c_idx)])
                c_idx += 1

            text = cell.xpath("string(.)").strip()
            rowspan = int(cell.get("rowspan", 1))
            colspan = int(cell.get("colspan", 1))

            for i in range(colspan):
                row.append(text)

            for i in range(rowspan):
                for j in range(colspan):
                    if i == 0 and j == 0:
                        continue
                    span_map[(r_idx + i, c_idx + j)] = text

            c_idx += colspan

        grid.append(row)

    return grid


def extract_table(root, keyword):
    keyword = keyword.strip()

    nodes = root.xpath(
        f"//*[contains(normalize-space(text()), '{keyword}')]"
    )

    if not nodes:
        raise Exception("❌ 섹션을 찾을 수 없음")

    node = nodes[0]

    tables = node.xpath("following::table[position()<=5]")

    for table in tables:
        text = "".join(table.xpath(".//text()"))

        if len(text) > 100:
            data = parse_table_with_span(table)
            df = pd.DataFrame(data)

            if len(df.columns) >= 2:
                df.columns = df.iloc[0]
                df = df[1:].reset_index(drop=True)
                return df

    raise Exception("❌ 적절한 표를 찾지 못함")


def get_table(corp_name, section, pblntf_ty, target_date=None):
    corp_list = dart.get_corp_list()
    corps = corp_list.find_by_corp_name(corp_name, exactly=False)

    if not corps:
        raise Exception("❌ 기업을 찾을 수 없음")

    corp = corps[0]

    reports = corp.search_filings(
        pblntf_ty=pblntf_ty,
        bgn_de='20200101',
        end_de='20261231'
    )

    if not reports:
        raise Exception("❌ 공시 없음")

    # 🔥 날짜 필터
    if target_date:
        target_date = target_date.replace("-", "").replace(".", "")

        matched = [r for r in reports if target_date in r.rcept_dt]

        if not matched:
            raise Exception("❌ 해당 날짜 공시 없음")

        target_report = matched[0]
    else:
        target_report = reports[0]

    res = requests.get(
        "https://opendart.fss.or.kr/api/document.xml",
        params={
            "crtfc_key": API_KEY,
            "rcept_no": target_report.rcp_no
        }
    )

    try:
        zf = zipfile.ZipFile(io.BytesIO(res.content))
    except zipfile.BadZipFile:
        raise Exception("❌ ZIP 파일 아님")

    full_text = ""

    for f in zf.namelist():
        if f.endswith(".xml"):
            try:
                full_text += zf.read(f).decode("euc-kr")
            except:
                full_text += zf.read(f).decode("utf-8", errors="ignore")

    root = etree.fromstring(full_text.encode(), etree.HTMLParser())

    return extract_table(root, section)