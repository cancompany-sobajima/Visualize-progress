import pandas as pd
import json
from pathlib import Path
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import numpy as np
import subprocess
import tempfile
import os
import sys
from datetime import timedelta

# --- 1. Google Sheets APIのスコープを設定 ---
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
]

# --- 2. 認証情報ファイルへのパス ---
PRODUCT_MASTER_PATH = Path(__file__).parent / "商品マスタ.xlsx"

# --- 3. スプレッドシートのIDとシート名 ---
PLAN_SHEET_ID = "1Nx1cIlaBToKLdl_d7DK-5JWNKEcBGxtiEd0AajWKFTc"
PLAN_WORKSHEET_NAME = "抽出先"  # 生産予定データのシート名
# -------------------------------------------------------------

# 名寄せマスタのパス
DATA_DIR = Path(__file__).parent / "data"
NAME_MASTER_PATH = DATA_DIR / "name_master.json"

@st.cache_resource(ttl=600)
def _get_gsheet_client():
    """gspreadクライアントを認証・初期化して返す。"""
    service_account_info = st.secrets["service_account"]
    creds = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client

@st.cache_data(ttl=60)
def _load_data_from_gsheet(_client, sheet_id: str, worksheet_name: str) -> pd.DataFrame:
    """
    指定されたシートからデータを安定的に読み込み、DataFrameとして返す。
    get_all_values()を使い、データの中身に左右されずに常にヘッダーに基づいたDataFrameを作成する。
    """
    try:
        spreadsheet = _client.open_by_key(sheet_id)
        worksheet = spreadsheet.worksheet(worksheet_name)
        
        values = worksheet.get_all_values()
        if not values:
            return pd.DataFrame()

        header = values[0]
        data = values[1:]
        
        num_columns = len(header)
        cleaned_data = [row[:num_columns] for row in data]

        df = pd.DataFrame(cleaned_data, columns=header)
        return df

    except gspread.exceptions.SpreadsheetNotFound:
        st.error(f"スプレッドシートが見つかりません。ID: {sheet_id}")
        st.stop()
    except gspread.exceptions.WorksheetNotFound:
        st.error(f"ワークシート '{worksheet_name}' が見つかりません。")
        st.stop()
    except Exception as e:
        st.error(f"データの読み込み中にエラーが発生しました: {e}")
        st.stop()


@st.cache_data(ttl=3600)
def load_product_master() -> pd.DataFrame:
    """商品マスタ.xlsxを読み込む。"""
    if not PRODUCT_MASTER_PATH.exists():
        st.error(f"商品マスタが見つかりません: {PRODUCT_MASTER_PATH}")
        return pd.DataFrame()
    
    df = pd.read_excel(PRODUCT_MASTER_PATH)
    df.rename(columns={'お客様': 'お客様名', 'ライン': '担当設備'}, inplace=True)
    return df


def load_plan_data(date) -> pd.DataFrame:
    """生産予定データをGoogleスプレッドシートから読み込み、整形して返す。"""
    if PLAN_SHEET_ID.startswith("ここに"):
        st.warning("生産予定シートのIDが設定されていません。data_loader.pyを編集してください。")
        return pd.DataFrame()

    client = _get_gsheet_client()
    df = _load_data_from_gsheet(client, PLAN_SHEET_ID, PLAN_WORKSHEET_NAME)
    
    if df.empty:
        st.info(f"DEBUG: load_plan_data - スプレッドシートからデータが読み込まれませんでした。") # デバッグ用
        return pd.DataFrame()

    try:
        # --- データクレンジングと型変換 ---
        # 1. 列名から前後の空白を除去
        df.columns = df.columns.str.strip()

        # 2. 必要な列が存在するか確認
        required_cols = ['日付', '開始時間', '終了時間', '予定数量', 'ライン', '顧客名（型替え）', '商品名（型の名前）']
        if not all(col in df.columns for col in required_cols):
            missing = [col for col in required_cols if col not in df.columns]
            st.error(f"スプレッドシートの列名に問題があります。必要な列が見つかりません: {missing}")
            st.info(f"プログラムが認識している現在の列名: {df.columns.tolist()}")
            st.stop()

        # 3. データフレーム全体の空文字列 '' を NaN に置換
        df.replace('', np.nan, inplace=True)

        # 4. 日付列をdatetimeに変換し、不正な行は除外
        df['日付'] = pd.to_datetime(df['日付'], errors='coerce')
        df.dropna(subset=['日付'], inplace=True)

        # 5. 選択された日付でフィルタリング
        df = df[df['日付'].dt.date == date].copy()
        if df.empty:
            st.info(f"{date} の生産予定データはありません。")
            return pd.DataFrame()

        # --- アプリ用データフレームへの変換 ---
        df_mapped = pd.DataFrame()

        # 6. 時刻・数量をテキストとして準備
        start_time_str = df['開始時間'].astype(str)
        end_time_str = df['終了時間'].astype(str)
        
        # 7. 日付と時刻の文字列を結合
        start_datetime_str = df['日付'].dt.strftime('%Y-%m-%d') + ' ' + start_time_str
        end_datetime_str = df['日付'].dt.strftime('%Y-%m-%d') + ' ' + end_time_str
        
        # 8. アプリ内部形式へ変換
        df_mapped['予定開始時刻'] = pd.to_datetime(start_datetime_str, errors='coerce')
        df_mapped['予定終了時刻'] = pd.to_datetime(end_datetime_str, errors='coerce')
        df_mapped['担当設備'] = df['ライン']
        df_mapped['お客様名'] = df['顧客名（型替え）']
        df_mapped['商品名'] = df['商品名（型の名前）']
        df_mapped['予定数'] = pd.to_numeric(df['予定数量'], errors='coerce')
        
<<<<<<< HEAD
        # 9. 必須データ（時刻と数量）がない行を最終的に除外
        df_mapped.dropna(subset=['予定開始時刻', '予定数'], inplace=True)
=======
        st.write(f"DEBUG: load_plan_data - Before dropna shape: {df_mapped.shape}") # デバッグ用
        # 予定開始時刻または予定数が空欄のデータを除外
        df_mapped.dropna(subset=['予定開始時刻', '予定数'], inplace=True)
        st.write(f"DEBUG: load_plan_data - After dropna shape: {df_mapped.shape}") # デバッグ用
>>>>>>> c586b6b38341dc837fc24378c5cc6d64a0d3d0e3

        return df_mapped

    except Exception as e:
        st.error(f"データの処理中に予期せぬエラーが発生しました: {e}")
        st.info("スプレッドシートのデータ形式が想定と違う可能性があります。")
        st.stop()


# --- 実績データ取得関連の関数（変更なし） ---
_APP_DIR = Path(__file__).parent
_EXTRACT_SCRIPT_DIR = _APP_DIR / "record_exporter"
_EXTRACT_SCRIPT_PATH = _EXTRACT_SCRIPT_DIR / "export_production_records.py"
_GENERATED_EXCEL_PATH = _EXTRACT_SCRIPT_DIR / "production_records.xlsx"

def load_results_data(date) -> pd.DataFrame:
    """生産実績データをFirestoreから取得し、整形して返す。"""
    end_date = date
    start_date = end_date - timedelta(days=1)
    start_date_str = start_date.strftime('%Y-%m-%d')
    end_date_str = end_date.strftime('%Y-%m-%d')

    if not _EXTRACT_SCRIPT_PATH.is_file():
        st.error(f"実績取得スクリプトが見つかりません: {_EXTRACT_SCRIPT_PATH}")
        return pd.DataFrame()

    service_account_info = st.secrets["service_account"]
    
    temp_file_path = None
    try:
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.json', dir=str(_EXTRACT_SCRIPT_DIR)) as temp_file:
            json.dump(dict(service_account_info), temp_file)
            temp_file_path = temp_file.name
        
        try:
            command = [sys.executable, str(_EXTRACT_SCRIPT_PATH), start_date_str, end_date_str, temp_file_path]
            result = subprocess.run(
                command, capture_output=True, text=True, encoding='utf-8', cwd=str(_EXTRACT_SCRIPT_DIR)
            )
            if result.returncode != 0:
                st.error("生産実績の取得に失敗しました。")
                st.code(result.stderr if result.stderr else result.stdout)
                return pd.DataFrame()
        finally:
            if temp_file_path and Path(temp_file_path).exists():
                Path(temp_file_path).unlink()

    except Exception as e:
        st.error(f"実績取得スクリプトの実行準備中に予期せぬエラーが発生しました: {e}")
        if temp_file_path and Path(temp_file_path).exists():
            Path(temp_file_path).unlink()
        return pd.DataFrame()

    if not _GENERATED_EXCEL_PATH.is_file():
        st.warning(f"生成されたExcelファイルが見つかりません: {_GENERATED_EXCEL_PATH}")
        return pd.DataFrame()
    try:
        df = pd.read_excel(_GENERATED_EXCEL_PATH)
    except Exception as e:
        st.error(f"Excelファイルの読み込みに失敗しました: {e}")
        return pd.DataFrame()

    if df.empty:
        return pd.DataFrame()

    df['date'] = pd.to_datetime(df['date']).dt.date
    df = df[df['date'] == date].copy()
    
    if df.empty:
        st.info(f"{date} の生産実績データはありません。")
        return pd.DataFrame()

    try:
        rename_map = {
            'line': '担当設備', 'customer': 'お客様名', 'product': '商品名',
            'actualQuantity': '実生産数', 'date': '日付'
        }
        df.rename(columns=rename_map, inplace=True)

        required_cols = ['担当設備', 'お客様名', '商品名', '実生産数', 'editSessions', '日付']
        if not all(col in df.columns for col in required_cols):
            missing_cols = [col for col in required_cols if col not in df.columns]
            st.error(f"実績Excelファイルに必要な列がありません: {', '.join(missing_cols)}")
            return pd.DataFrame()

        df_mapped = df[required_cols].copy()
        
        start_times, end_times, total_durations_min = [], [], []
        all_session_starts, all_session_ends = [], []

        for idx, row in df_mapped.iterrows():
            item = row['editSessions']
            record_date = row['日付']

            current_record_session_starts = []
            current_record_session_ends = []
            current_record_total_duration_sec = 0

            try:
                sessions = json.loads(item) if isinstance(item, str) else item
                if sessions and isinstance(sessions, list):
                    for s in sessions:
                        s_start = s.get('startTime')
                        s_end = s.get('endTime')
                        
                        if s_start and s_end:
                            full_start_dt = pd.to_datetime(f"{str(record_date)} {s_start}", format="%Y-%m-%d %H:%M")
                            full_end_dt = pd.to_datetime(f"{str(record_date)} {s_end}", format="%Y-%m-%d %H:%M")
                            
                            current_record_session_starts.append(full_start_dt)
                            current_record_session_ends.append(full_end_dt)
                            current_record_total_duration_sec += (full_end_dt - full_start_dt).total_seconds()
                
                start_times.append(min(current_record_session_starts) if current_record_session_starts else pd.NaT)
                end_times.append(max(current_record_session_ends) if current_record_session_ends else pd.NaT)
                total_durations_min.append(current_record_total_duration_sec / 60 if current_record_total_duration_sec > 0 else 0)
                all_session_starts.append(current_record_session_starts)
                all_session_ends.append(current_record_session_ends)

            except (json.JSONDecodeError, TypeError, ValueError):
                start_times.append(pd.NaT)
                end_times.append(pd.NaT)
                total_durations_min.append(0)
                all_session_starts.append([])
                all_session_ends.append([])

        df_mapped['実生産開始時刻'] = start_times
        df_mapped['実生産終了時刻'] = end_times
        df_mapped['実績総生産時間_分'] = total_durations_min
        df_mapped['実セッション開始時刻リスト'] = all_session_starts
        df_mapped['実セッション終了時刻リスト'] = all_session_ends
        df_mapped['実生産数'] = pd.to_numeric(df_mapped['実生産数'], errors='coerce')
        
        return df_mapped.drop(columns=['editSessions'])

    except Exception as e:
        st.error(f"実績Excelデータの整形中にエラーが発生しました: {e}")
        st.dataframe(df.head())
        return pd.DataFrame()


def load_name_master() -> dict:
    """名寄せマスタを読み込む。"""
    try:
        with open(NAME_MASTER_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        return {"お客様名": {}, "商品名": {}}