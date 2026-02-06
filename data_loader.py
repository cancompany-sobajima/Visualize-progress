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
# Streamlit Secretsからサービスアカウント情報を読み込む
# SERVICE_ACCOUNT_FILE = Path(__file__).parent / "service_account.json"
PRODUCT_MASTER_PATH = Path(__file__).parent / "商品マスタ.xlsx"

# --- 3. ★★★ ここにスプレッドシートのIDとシート名を入力してください ★★★ ---
PLAN_SHEET_ID = "1Nx1cIlaBToKLdl_d7DK-5JWNKEcBGxtiEd0AajWKFTc"
PLAN_WORKSHEET_NAME = "抽出先"  # 生産予定データのシート名
# -------------------------------------------------------------

# 名寄せマスタのパスは変更なし
DATA_DIR = Path(__file__).parent / "data"
NAME_MASTER_PATH = DATA_DIR / "name_master.json"

@st.cache_resource(ttl=600)
def _get_gsheet_client():
    """gspreadクライアントを認証・初期化して返す。結果はStreamlitでキャッシュする。"""
    # Streamlit Secretsからサービスアカウント情報を取得
    service_account_info = st.secrets["service_account"]
    
    # サービスアカウント情報を使って認証
    creds = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client

@st.cache_data(ttl=60)
def _load_data_from_gsheet(_client, sheet_id: str, worksheet_name: str) -> pd.DataFrame:
    """
    指定されたシートIDとワークシート名からデータを読み込み、DataFrameとして返す。
    結果はStreamlitでキャッシュする。
    """
    try:
        spreadsheet = _client.open_by_key(sheet_id)
        worksheet = spreadsheet.worksheet(worksheet_name)
        records = worksheet.get_all_records()
        return pd.DataFrame(records)
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
    """商品マスタ.xlsxを読み込む。結果は1時間キャッシュする。"""
    if not PRODUCT_MASTER_PATH.exists():
        st.error(f"商品マスタが見つかりません: {PRODUCT_MASTER_PATH}")
        return pd.DataFrame()
    
    df = pd.read_excel(PRODUCT_MASTER_PATH)
    # アプリ内部で使う列名に統一する
    df.rename(columns={'お客様': 'お客様名', 'ライン': '担当設備'}, inplace=True)
    return df


def load_plan_data(date) -> pd.DataFrame:
    """生産予定データをGoogleスプレッドシートから読み込む。"""
    if PLAN_SHEET_ID.startswith("ここに"):
        st.warning("生産予定シートのIDが設定されていません。data_loader.pyを編集してください。")
        return pd.DataFrame()

    client = _get_gsheet_client()
    df = _load_data_from_gsheet(client, PLAN_SHEET_ID, PLAN_WORKSHEET_NAME)
    
    if df.empty:
        return pd.DataFrame()

    try:
        # --- 「生産予定」シートは年が含まれているため、そのまま日付に変換 ---
        df['日付'] = pd.to_datetime(df['日付'], errors='coerce')

        # 日付でフィルタリング
        df = df[df['日付'].dt.date == date].copy()
        if df.empty:
            st.info(f"{date} の生産予定データはありません。")
            return pd.DataFrame()

        # 列名をプログラムの内部名に変換
        df_mapped = pd.DataFrame()
        df_mapped['予定開始時刻'] = pd.to_datetime(df['日付'].dt.strftime('%Y-%m-%d') + ' ' + df['開始時間'], errors='coerce')
        df_mapped['予定終了時刻'] = pd.to_datetime(df['日付'].dt.strftime('%Y-%m-%d') + ' ' + df['終了時間'], errors='coerce')
        df_mapped['担当設備'] = df['ライン']
        df_mapped['お客様名'] = df['顧客名（型替え）']
        df_mapped['商品名'] = df['商品名（型の名前）']
        df_mapped['予定数'] = pd.to_numeric(df['予定数量'], errors='coerce')
        
        # 予定開始時刻が空欄のデータを除外
        df_mapped.dropna(subset=['予定開始時刻'], inplace=True)

        return df_mapped

    except KeyError as e:
        st.error(f"「生産予定」シートで列のマッピング中にエラーが発生しました。必要な列が見つかりません: {e}")
        st.error("シートの列名が変更されていないか確認してください。")
        st.stop()


# --- パス定義 ---
# スクリプトのあるディレクトリ（＝アプリのルート）を基準にする
_APP_DIR = Path(__file__).parent
_EXTRACT_SCRIPT_DIR = _APP_DIR / "record_exporter"
_EXTRACT_SCRIPT_PATH = _EXTRACT_SCRIPT_DIR / "export_production_records.py"
_GENERATED_EXCEL_PATH = _EXTRACT_SCRIPT_DIR / "production_records.xlsx"

def load_results_data(date) -> pd.DataFrame:
    """
    外部スクリプトを実行してFirestoreから生産実績をExcelに抽出し、
    そのExcelファイルを読み込んで整形する。
    認証情報は一時ファイル経由で安全に渡す。
    """
    # --- 1. 日付設定とパスの検証 ---
    end_date = date
    start_date = end_date - timedelta(days=1)
    start_date_str = start_date.strftime('%Y-%m-%d')
    end_date_str = end_date.strftime('%Y-%m-%d')

    if not _EXTRACT_SCRIPT_PATH.is_file():
        st.error(f"実績取得スクリプトが見つかりません: {_EXTRACT_SCRIPT_PATH}")
        
        # --- デバッグ情報 ---
        st.warning("デバッグ情報を表示します。")
        
        # 親ディレクトリの存在確認
        parent_dir = _EXTRACT_SCRIPT_PATH.parent
        st.info(f"スクリプトの親ディレクトリ: '{parent_dir}'")
        st.info(f"親ディレクトリは存在しますか？ -> {parent_dir.exists()}")
        st.info(f"親ディレクトリは 'ディレクトリ' ですか？ -> {parent_dir.is_dir()}")

        # 親ディレクトリの中身をリストアップ
        if parent_dir.exists() and parent_dir.is_dir():
            try:
                st.info(f"'{parent_dir}' の中身:")
                st.code(str(os.listdir(parent_dir)))
            except Exception as e:
                st.error(f"ディレクトリの中身を取得中にエラー: {e}")

        # アプリのルートディレクトリの中身もリストアップ
        app_root = Path(__file__).parent
        st.info(f"アプリのルートディレクトリ '{app_root}' の中身:")
        try:
            st.code(str(os.listdir(app_root)))
        except Exception as e:
            st.error(f"アプリルートの中身を取得中にエラー: {e}")
        # --- デバッグ情報ここまで ---

        return pd.DataFrame()

    # --- 2. 外部スクリプトを実行してExcelを生成 ---
    service_account_info = st.secrets["service_account"]
    
    temp_file_path = None
    try:
        # 一時ファイルに認証情報を書き込む
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.json', dir=str(_EXTRACT_SCRIPT_DIR)) as temp_file:
            json.dump(service_account_info, temp_file)
            temp_file_path = temp_file.name
        
        st.info(f"最新の生産実績を取得しています ({start_date_str} ～ {end_date_str})...")
        
        try:
            # コマンドに一時ファイルのパスを追加
            command = [sys.executable, str(_EXTRACT_SCRIPT_PATH), start_date_str, end_date_str, temp_file_path]
            result = subprocess.run(
                command, capture_output=True, text=True, encoding='utf-8', cwd=str(_EXTRACT_SCRIPT_DIR)
            )
            if result.returncode != 0:
                st.error("生産実績の取得に失敗しました。")
                st.text("---エラー詳細---")
                if result.stderr:
                    st.code(result.stderr)
                elif result.stdout:
                    st.code(result.stdout)
                else:
                    st.error("外部スクリプトはエラーコードを返しましたが、詳細なエラーメッセージがありませんでした。")
                return pd.DataFrame()
        finally:
            # 一時ファイルを確実に削除
            if temp_file_path and Path(temp_file_path).exists():
                Path(temp_file_path).unlink()

    except Exception as e:
        st.error(f"実績取得スクリプトの実行準備中に予期せぬエラーが発生しました: {e}")
        # 一時ファイルが残っていたら削除
        if temp_file_path and Path(temp_file_path).exists():
            Path(temp_file_path).unlink()
        return pd.DataFrame()

    # --- 3. 生成されたExcelを読み込む ---
    if not _GENERATED_EXCEL_PATH.is_file():
        st.warning(f"生成されたExcelファイルが見つかりません: {_GENERATED_EXCEL_PATH}")
        return pd.DataFrame()
    try:
        df = pd.read_excel(_GENERATED_EXCEL_PATH)
    except Exception as e:
        st.error(f"Excelファイルの読み込みに失敗しました: {e}")
        return pd.DataFrame()

    if df.empty:
        st.info(f"{start_date_str} ～ {end_date_str} の生産実績データはありません。")
        return pd.DataFrame()

    # 'date' 列を datetime.date オブジェクトに変換
    df['date'] = pd.to_datetime(df['date']).dt.date

    # 選択された日付でフィルタリング
    df = df[df['date'] == date].copy()
    
    if df.empty:
        st.info(f"{date} の生産実績データはありません。")
        return pd.DataFrame()

    # --- 4. データをアプリの形式に変換 ---
    try:
        rename_map = {
            'line': '担当設備', 'customer': 'お客様名', 'product': '商品名',
            'actualQuantity': '実生産数', 'date': '日付'
        }
        df.rename(columns=rename_map, inplace=True)

        required_cols = ['担当設備', 'お客様名', '商品名', '実生産数', 'editSessions', '日付']
        if not all(col in df.columns for col in required_cols):
            missing_cols = [col for col in required_cols if col not in df.columns]
            st.error(f"Excelファイルに必要な列がありません: {', '.join(missing_cols)}")
            st.code(f"要求列: {required_cols}")
            st.code(f"実際の列: {df.columns.to_list()}")
            return pd.DataFrame()

        df['日付'] = pd.to_datetime(df['日付']).dt.date # 日付列をdatetime.dateオブジェクトに変換
        df_mapped = df[required_cols].copy()
        
        start_times, end_times, total_durations_min = [], [], []
        all_session_starts, all_session_ends = [], [] # 各セッションの開始・終了時刻リスト

        for idx, row in df_mapped.iterrows():
            item = row['editSessions']
            record_date = row['日付'] # この実績レコードの日付

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
                            # 日付情報と結合してdatetimeオブジェクトを作成
                            # formatを明示的に指定してパースを堅牢化
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
        df_mapped['実セッション開始時刻リスト'] = all_session_starts # 新しい列
        df_mapped['実セッション終了時刻リスト'] = all_session_ends   # 新しい列
        df_mapped['実生産数'] = pd.to_numeric(df_mapped['実生産数'], errors='coerce')
        
        return df_mapped.drop(columns=['editSessions'])

    except Exception as e:
        st.error(f"Excelデータの整形中にエラーが発生しました: {e}")
        st.dataframe(df.head())
        return pd.DataFrame()


def load_name_master() -> dict:
    """名寄せマスタを読み込む。"""
    try:
        with open(NAME_MASTER_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        return {"お客様名": {}, "商品名": {}}