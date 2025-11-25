import csv
import os
import urllib.request

# ===== 設定 =====
CSV_PATH = r"C:\Users\YourName\Documents\pdf_links.csv"
OUTPUT_DIR = r"C:\Users\YourName\Documents\downloaded_pdfs"


def download_pdf(url, save_path):
    """PDFをダウンロードし、成功ならTrue、失敗ならFalseを返す"""
    try:
        request = urllib.request.Request(
            url,
            headers={"User-Agent": "Mozilla/5.0"}  # ブロック防止
        )
        with urllib.request.urlopen(request, timeout=10) as response:
            if response.status != 200:
                return False

            with open(save_path, "wb") as f:
                f.write(response.read())

        return True

    except Exception:
        return False


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with open(CSV_PATH, newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)

        for row_idx, row in enumerate(reader, start=1):

            # C列 = 3番目
            if len(row) < 3:
                continue

            url = row[2].strip()
            if not url:
                continue

            # PDFで終わらないURLはスキップ（必要なら削除OK）
            if not url.lower().endswith(".pdf"):
                print(f"行{row_idx}: PDFリンクではない → スキップ: {url}")
                continue

            # 保存先
            save_path = os.path.join(OUTPUT_DIR, f"row_{row_idx}.pdf")

            print(f"行{row_idx}: ダウンロード中 → {url}")

            # ここで失敗時にスキップ
            if download_pdf(url, save_path):
                print(f"  ✔ 保存成功: {save_path}")
            else:
                print(f"  ✖ ダウンロード失敗（URLが無効？）→ スキップ: {url}")

    print("\nすべての処理が完了しました。")


if __name__ == "__main__":
    main()

#C3~C5のみにする
for row_idx, row in enumerate(reader, start=1):
    if row_idx < 3 or row_idx > 5:
        continue


