import firebase_admin
from firebase_admin import credentials, firestore
import openpyxl
import os
import json
import sys
from datetime import datetime, time
import logging

# サービスアカウントキーはコマンドライン引数から受け取ります

# 出力ディレクトリのパス (スクリプトと同じディレクトリ)
OUTPUT_DIR = os.path.dirname(__file__)
# 出力ファイル名を固定
OUTPUT_FILENAME = os.path.join(OUTPUT_DIR, "production_records.xlsx")

# ログファイルの設定
LOG_FILE = os.path.join(OUTPUT_DIR, "export_log.txt")
logging.basicConfig(filename=LOG_FILE, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def main(start_date_str, end_date_str, service_account_path):
    logging.info(f"スクリプト開始: {start_date_str} ～ {end_date_str}")
    # --- Firestore Admin SDK の初期化 ---
    try:
        if not firebase_admin._apps:
            cred = credentials.Certificate(service_account_path)
            firebase_admin.initialize_app(cred)
        db = firestore.client()
        logging.info("Firebase Admin SDK の初期化に成功しました。")
    except Exception as e:
        print(f"Firebase Admin SDK の初期化に失敗しました: {e}", file=sys.stderr, flush=True)
        logging.error(f"Firebase Admin SDK の初期化に失敗しました: {e}")
        sys.exit(1)

    # --- 日付文字列の検証 ---
    try:
        # 日付形式が 'YYYY-MM-DD' であることを確認
        datetime.strptime(start_date_str, '%Y-%m-%d')
        datetime.strptime(end_date_str, '%Y-%m-%d')
        logging.info("日付文字列の検証に成功しました。")
    except ValueError:
        print(f"エラー: 日付の形式が不正です。'YYYY-MM-DD' 形式で指定してください。", file=sys.stderr, flush=True)
        logging.error(f"日付の形式が不正です。'YYYY-MM-DD' 形式で指定してください。")
        sys.exit(1)

    # --- Firestore からデータを取得 ---
    logging.info(f"Firestore から {start_date_str} ～ {end_date_str} のデータを取得中...")
    try:
        # 'date' フィールドが 'YYYY-MM-DD' 形式の文字列であることを想定してクエリを実行
        docs = db.collection('productionRecords') \
                 .where('date', '>=', start_date_str) \
                 .where('date', '<=', end_date_str) \
                 .stream()
        
        records = []
        for doc in docs:
            record = doc.to_dict()
            record['documentId'] = doc.id
            records.append(record)
            
        logging.info(f"{len(records)} 件のレコードを取得しました。")
    except Exception as e:
        print(f"Firestore からデータの取得に失敗しました: {e}", file=sys.stderr, flush=True)
        logging.error(f"Firestore からデータの取得に失敗しました: {e}")
        sys.exit(1)

    if not records:
        logging.info("取得したレコードがありません。空のExcelファイルを生成します。")

    # --- Excel ファイルの作成 ---
    logging.info("Excel ファイルを作成中...")
    try:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Production Records"

        if records:
            # ヘッダーの作成
            all_keys = set()
            for record in records:
                all_keys.update(record.keys())
            
            sorted_keys = sorted(list(all_keys))
            if 'documentId' in sorted_keys:
                sorted_keys.remove('documentId')
                sorted_keys.insert(0, 'documentId')
            
            # 'date'列をdocumentIdの隣に移動
            if 'date' in sorted_keys:
                sorted_keys.remove('date')
                sorted_keys.insert(1, 'date')

            sheet.append(sorted_keys)

            for record in records:
                row_data = []
                for key in sorted_keys:
                    value = record.get(key, '')
                    # FirestoreのTimestampをPythonのdatetimeに変換
                    if isinstance(value, datetime):
                        # タイムゾーン情報を削除（Excelで扱いやすくするため）
                        value = value.replace(tzinfo=None)
                    elif isinstance(value, (list, dict)):
                        value = json.dumps(value, ensure_ascii=False)
                    row_data.append(value)
                sheet.append(row_data)
        else:
            # レコードがない場合もヘッダーなしの空のファイルを作成
            pass

        workbook.save(OUTPUT_FILENAME)
        logging.info(f"データを '{OUTPUT_FILENAME}' に正常にエクスポートしました。")

    except Exception as e:
        print(f"Excel ファイルの作成または書き込みに失敗しました: {e}", file=sys.stderr, flush=True)
        logging.error(f"Excel ファイルの作成または書き込みに失敗しました: {e}")
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("エラー: 開始日、終了日、サービスアカウントキーのパスをコマンドライン引数として指定してください。", file=sys.stderr, flush=True)
        print("例: python export_production_records.py 2024-07-29 2024-07-30 /path/to/key.json", file=sys.stderr, flush=True)
        sys.exit(1)
    
    main(sys.argv[1], sys.argv[2], sys.argv[3])