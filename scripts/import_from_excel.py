import json
import pandas as pd
from pathlib import Path
import collections

# パス設定
# このスクリプトはプロジェクトのルートから実行されることを想定
JSON_PATH = Path("./data/name_master.json")
EXCEL_PATH = Path("./name_master_editor.xlsx")

def import_from_excel():
    """
    編集用のExcelファイルを読み込み、name_master.jsonを更新する。
    """
    if not EXCEL_PATH.exists():
        print(f"エラー: {EXCEL_PATH} が見つかりません。先にエクスポートを実行してください。")
        return

    final_json = collections.defaultdict(dict)

    # Excelから両シートを読み込む
    try:
        excel_file = pd.ExcelFile(EXCEL_PATH)
        sheet_names = excel_file.sheet_names
    except Exception as e:
        print(f"Excelファイルの読み込み中にエラーが発生しました: {e}")
        return

    # お客様名マスタを処理
    if 'お客様名マスタ' in sheet_names:
        customer_df = pd.read_excel(EXCEL_PATH, sheet_name='お客様名マスタ')
        # NaN（空のセル）を空文字列に変換
        customer_df['別名'] = customer_df['別名'].fillna('')
        # 正式名称でグループ化し、別名をリストにまとめる
        customer_grouped = customer_df.groupby('正式名称')['別名'].apply(lambda x: list(x) if x.iloc[0] != '' else []).to_dict()
        final_json["お客様名"] = customer_grouped
        print("お客様名マスタをインポートしました。")

    # 商品名マスタを処理
    if '商品名マスタ' in sheet_names:
        product_df = pd.read_excel(EXCEL_PATH, sheet_name='商品名マスタ')
        product_df['別名'] = product_df['別名'].fillna('')
        product_grouped = product_df.groupby('正式名称')['別名'].apply(lambda x: list(x) if x.iloc[0] != '' else []).to_dict()
        final_json["商品名"] = product_grouped
        print("商品名マスタをインポートしました。")

    # JSONファイルに書き出す
    try:
        with open(JSON_PATH, 'w', encoding='utf-8') as f:
            json.dump(final_json, f, indent=2, ensure_ascii=False)
        print(f"\nインポートが完了しました。'{JSON_PATH}' が更新されました。")
    except Exception as e:
        print(f"JSONファイルへの書き込み中にエラーが発生しました: {e}")

if __name__ == "__main__":
    import_from_excel()
