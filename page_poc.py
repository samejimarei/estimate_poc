import io
import zipfile
from pathlib import Path

import pdfplumber
import pandas as pd


# ============================================================
# このスクリプトの目的
# ------------------------------------------------------------
# いきなり全PDF対応を目指さず、
# 「1つのPDFの1ページだけ」を対象にして、
# ヘッダー位置から列境界を決め、
# 明細を6列に分けて読めるかを確認するためのPoCです。
#
# 今回の対象
# - ZIP: 見積一覧.zip
# - PDF: 0213近藤様_近藤様邸1階改装工事_御見積書.pdf
# - Page: 2ページ目（Python上は index=1）
#
# この段階では、完璧な抽出器は作りません。
# まずは
#   no / item_spec / quantity / unit / unit_price / amount
# が「位置ベース」で取れるかどうかだけを確認します。
# ============================================================


# ============================================================
# 設定値
# ------------------------------------------------------------
# 必要に応じてあとで変更できます。
# まずはこのままでOKです。
# ============================================================

ZIP_PATH = "見積一覧.zip"
TARGET_PDF_NAME = "0213近藤様_近藤様邸1階改装工事_御見積書.pdf"
TARGET_PAGE_INDEX = 1  # 2ページ目 = index 1

# ページ下部のフッターっぽい領域は最初から無視する
# 日付やページ番号ノイズが入るのを防ぐため
FOOTER_CUTOFF_FROM_BOTTOM = 24

# 同じ行とみなす y座標の許容差
ROW_Y_TOLERANCE = 3

# 列境界を少し広めに取るための余白
COLUMN_MARGIN = 6


# ============================================================
# ユーティリティ
# ============================================================

def clean_text(text: str) -> str:
    """
    文字列の基本クリーニング。
    全角スペースや連続空白を整えて扱いやすくします。
    """
    text = str(text or "")
    text = text.replace("\u3000", " ")
    text = text.replace("\xa0", " ")
    text = " ".join(text.split())
    return text.strip()


def load_target_page_from_zip(zip_path: str, pdf_name: str, page_index: int):
    """
    ZIPの中から対象PDFを取り出し、指定ページを返します。

    注意:
    - pdfplumberの page_index は0始まり
    - 今回は検証用なので、対象PDFを1つに固定しています
    """
    with zipfile.ZipFile(zip_path, "r") as zf:
        file_bytes = zf.read(pdf_name)

    pdf = pdfplumber.open(io.BytesIO(file_bytes))
    page = pdf.pages[page_index]
    return pdf, page


def extract_words_from_page(page):
    """
    PDFページから words を抽出します。

    ここでいう words は、
    PDF上の「文字のかたまり」と、その座標情報です。

    例:
    {
        "text": "数量",
        "x0": 380.1,
        "x1": 402.3,
        "top": 145.0,
        "bottom": 154.8
    }

    重要:
    今回のPoCでは、文字列そのものよりも、
    「その文字がどこにあるか」を重視します。
    """
    words = page.extract_words(
        keep_blank_chars=False,
        use_text_flow=True,
        x_tolerance=2,
        y_tolerance=2,
    )

    footer_cutoff = page.height - FOOTER_CUTOFF_FROM_BOTTOM

    cleaned = []
    for w in words:
        text = clean_text(w.get("text", ""))
        if not text:
            continue

        # ページ下端のフッター付近は除外
        if float(w["bottom"]) >= footer_cutoff:
            continue

        cleaned.append({
            "text": text,
            "x0": float(w["x0"]),
            "x1": float(w["x1"]),
            "top": float(w["top"]),
            "bottom": float(w["bottom"]),
        })

    return cleaned


def group_words_into_rows(words):
    """
    抽出した words を、y座標ベースで「行」にまとめます。

    今回の見積PDFは、明細表の行が水平方向に並んでいるので、
    top 座標が近いものを同じ行として扱います。
    """
    words_sorted = sorted(words, key=lambda w: (round(w["top"], 1), w["x0"]))

    rows = []
    current_row = []
    current_top = None

    for w in words_sorted:
        if current_top is None:
            current_row = [w]
            current_top = w["top"]
            continue

        if abs(w["top"] - current_top) <= ROW_Y_TOLERANCE:
            current_row.append(w)
        else:
            rows.append(sorted(current_row, key=lambda x: x["x0"]))
            current_row = [w]
            current_top = w["top"]

    if current_row:
        rows.append(sorted(current_row, key=lambda x: x["x0"]))

    return rows


def row_to_text(row_words):
    """
    行に含まれる words を左から結合し、人間が見やすい1行テキストにします。
    デバッグ表示用です。
    """
    return " ".join(w["text"] for w in sorted(row_words, key=lambda x: x["x0"]))


def find_header_row(rows):
    """
    ヘッダー行を探します。

    今回の見積PDFでは、典型的に
    "NO. 項目 仕様・規格/型番 数量 単位 単価 金額"
    のような行がヘッダーです。

    ここでは厳密一致ではなく、
    ヘッダーらしいキーワードを複数含む行を採用します。
    """
    header_keywords = ["NO.", "項目", "数量", "単位", "単価", "金額"]

    for row in rows:
        text = row_to_text(row)
        score = sum(1 for k in header_keywords if k in text)
        if score >= 5:
            return row

    return None


def get_word_x_by_text(row_words, target_text):
    """
    ヘッダー行から、指定した語の x0 / x1 を取得します。

    完全一致が基本ですが、
    多少の揺れがあっても取れるように contains で見ます。
    """
    for w in row_words:
        if target_text in w["text"]:
            return w["x0"], w["x1"]
    return None, None


def build_column_boundaries_from_header(header_row):
    """
    ヘッダー行の位置から、列境界を動的に作ります。

    今回の考え方:
    - 縦罫線に頼らない
    - 毎ページのヘッダー語の位置から列帯を決める
    - 右側4列（数量・単位・単価・金額）はかなり安定しているので特に重要

    戻り値は各列の x範囲です。
    """
    # ヘッダー語の位置を取得
    no_x0, no_x1 = get_word_x_by_text(header_row, "NO.")
    item_x0, item_x1 = get_word_x_by_text(header_row, "項目")
    qty_x0, qty_x1 = get_word_x_by_text(header_row, "数量")
    unit_x0, unit_x1 = get_word_x_by_text(header_row, "単位")
    unit_price_x0, unit_price_x1 = get_word_x_by_text(header_row, "単価")
    amount_x0, amount_x1 = get_word_x_by_text(header_row, "金額")

    # 念のため、ヘッダーが取れなかった場合はエラーにする
    required = [no_x0, item_x0, qty_x0, unit_x0, unit_price_x0, amount_x0]
    if any(v is None for v in required):
        raise ValueError("ヘッダー位置の取得に失敗しました。")

    # 列境界の基本方針
    # 左列の開始はそのまま採用し、
    # 列の切れ目は隣のヘッダー語の中間くらいで決める
    #
    # 例:
    # item列の終端 = (数量列の開始 + 項目列の終端) / 2
    #
    # こうすると帳票ごとの多少のズレに追従しやすい
    boundaries = {
        "no": (
            max(0, no_x0 - COLUMN_MARGIN),
            (item_x0 + no_x1) / 2
        ),
        "item_spec": (
            (item_x0 - COLUMN_MARGIN),
            (qty_x0 + item_x1) / 2
        ),
        "quantity": (
            (qty_x0 - COLUMN_MARGIN),
            (unit_x0 + qty_x1) / 2
        ),
        "unit": (
            (unit_x0 - COLUMN_MARGIN),
            (unit_price_x0 + unit_x1) / 2
        ),
        "unit_price": (
            (unit_price_x0 - COLUMN_MARGIN),
            (amount_x0 + unit_price_x1) / 2
        ),
        "amount": (
            (amount_x0 - COLUMN_MARGIN),
            10000  # 十分大きい右端値
        ),
    }

    return boundaries


def assign_word_to_column(word, boundaries):
    """
    1つの word がどの列に属するかを判定します。

    基本は word の中心 x座標で判定します。
    """
    center_x = (word["x0"] + word["x1"]) / 2

    for col_name, (x_min, x_max) in boundaries.items():
        if x_min <= center_x < x_max:
            return col_name

    return None


def row_words_to_record(row_words, boundaries):
    """
    1行分の words を、6列のレコードに変換します。

    ここでは意味解釈よりも「どの列帯にいるか」を優先します。
    """
    record = {
        "no": [],
        "item_spec": [],
        "quantity": [],
        "unit": [],
        "unit_price": [],
        "amount": [],
        "raw_row": row_to_text(row_words),
    }

    for w in row_words:
        col = assign_word_to_column(w, boundaries)
        if col is None:
            continue
        record[col].append(w["text"])

    # 各列を文字列にまとめる
    # item_spec は単語をつなぎ、他の列も空白区切りでまとめる
    result = {
        "no": " ".join(record["no"]).strip(),
        "item_spec": " ".join(record["item_spec"]).strip(),
        "quantity": " ".join(record["quantity"]).strip(),
        "unit": " ".join(record["unit"]).strip(),
        "unit_price": " ".join(record["unit_price"]).strip(),
        "amount": " ".join(record["amount"]).strip(),
        "raw_row": record["raw_row"],
    }

    return result


def is_detail_like_record(record):
    """
    明細っぽい行だけを残すための簡易判定。

    今回はPoCなので厳しすぎる条件にしません。
    以下のような最低限の条件だけ置きます。
    - no が数字で始まる
    - amount が何か入っている
    """
    no_text = record["no"].strip()
    amount_text = record["amount"].strip()

    if not no_text:
        return False

    if not no_text[0].isdigit():
        return False

    if not amount_text:
        return False

    return True


def main():
    """
    実行本体。
    ここで順番に
    1) PDF読み込み
    2) words取得
    3) 行化
    4) ヘッダー検出
    5) 列境界作成
    6) 明細化
    7) CSV保存
    を行います。
    """
    print("=== page_poc.py を開始します ===")
    print(f"対象ZIP: {ZIP_PATH}")
    print(f"対象PDF: {TARGET_PDF_NAME}")
    print(f"対象ページ: {TARGET_PAGE_INDEX + 1}ページ目")

    pdf, page = load_target_page_from_zip(ZIP_PATH, TARGET_PDF_NAME, TARGET_PAGE_INDEX)

    try:
        words = extract_words_from_page(page)
        rows = group_words_into_rows(words)
        header_row = find_header_row(rows)

        if header_row is None:
            raise ValueError("ヘッダー行が見つかりませんでした。")

        print("\n=== 検出したヘッダー行 ===")
        print(row_to_text(header_row))

        boundaries = build_column_boundaries_from_header(header_row)

        print("\n=== 列境界 ===")
        for k, v in boundaries.items():
            print(f"{k}: {v}")

        # ヘッダーより下の行だけを対象にする
        header_top = min(w["top"] for w in header_row)

        detail_rows = []
        for row in rows:
            row_top = min(w["top"] for w in row)
            if row_top <= header_top:
                continue

            record = row_words_to_record(row, boundaries)

            if is_detail_like_record(record):
                detail_rows.append(record)

        df = pd.DataFrame(detail_rows, columns=[
            "no", "item_spec", "quantity", "unit", "unit_price", "amount", "raw_row"
        ])

        print("\n=== 抽出結果 ===")
        print(df.to_string(index=False))

        output_csv = Path("page_poc_output.csv")
        df.to_csv(output_csv, index=False, encoding="utf-8-sig")

        print(f"\nCSV保存完了: {output_csv.resolve()}")

    finally:
        pdf.close()


if __name__ == "__main__":
    main()
