import json
import pandas as pd
from pathlib import Path

# パス設定
# このスクリプトはプロジェクトのルートから実行されることを想定
JSON_PATH = Path("./data/name_master.json")
EXCEL_PATH = Path("./name_master_editor.xlsx")

def export_to_excel():
    """
    name_master.jsonを読み込み、編集用のExcelファイルに出力する。
    """
    # JSONファイルを読み込む
    try:
        with open(JSON_PATH, 'r', encoding='utf-8') as f:
            name_master = json.load(f)
    except FileNotFoundError:
        print(f"エラー: {JSON_PATH} が見つかりません。")
        return
    except json.JSONDecodeError:
        print(f"エラー: {JSON_PATH} は不正なJSON形式です。")
        return

    # Excelライターを作成
    with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl') as writer:
        # お客様名マスタを処理
        customer_data = []
        for official_name, aliases in name_master.get("お客様名", {}).items():
            if not aliases:  # 別名がない場合
                customer_data.append({"正式名称": official_name, "別名": ""})
            else:
                for alias in aliases:
                    customer_data.append({"正式名称": official_name, "別名": alias})
        
        if customer_data:
            customer_df = pd.DataFrame(customer_data)
            customer_df.to_excel(writer, sheet_name='お客様名マスタ', index=False)
            print("お客様名マスタをExcelに出力しました。")

        # 商品名マスタを処理
        product_data = []
        for official_name, aliases in name_master.get("商品名", {}).items():
            if not aliases:
                product_data.append({"正式名称": official_name, "別名": ""})
            else:
                for alias in aliases:
                    product_data.append({"正式名称": official_name, "別名": alias})

        if product_data:
            product_df = pd.DataFrame(product_data)
            product_df.to_excel(writer, sheet_name='商品名マスタ', index=False)
            print("商品名マスタをExcelに出力しました。")

    print(f"\nエクスポートが完了しました。'{EXCEL_PATH}' を確認してください。")

if __name__ == "__main__":
    export_to_excel()
