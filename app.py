import streamlit as st
import pandas as pd
from datetime import date
from pandas.io.formats.style import Styler

# --- 各レイヤーのモジュールをインポート ---
import data_loader
import progress_logic

# --- UIの表示設定 ---
st.set_page_config(page_title="今日の生産進捗", layout="wide")

def style_progress_table(df: pd.DataFrame) -> Styler:
    """進捗状態に応じてテーブルの行を色付けする。"""
    def get_color(status):
        # デフォルトは白背景
        default_style = 'background-color: #ffffff; color: #000000;'
        
        if pd.isna(status):
            return default_style
        
        # ステータスに応じた配色
        if "遅延" in status:
            return 'background-color: #fff0f0; color: #000000;'  # Light Red
        elif "未開始" in status or status == '予定外':
            return 'background-color: #fafafa; color: #000000;'  # Light Grey
        
        # それ以外（進行中、完了など）はすべて白背景
        return default_style

    return df.style.apply(
        lambda row: [get_color(row['予定'])] * len(row), 
        axis=1
    )

def style_timeline(df: pd.DataFrame) -> Styler:
    """タイムラインDataFrameをスタイリングする。"""
    def get_style(val):
        if val == "予定":
            # 薄い灰色
            return 'background-color: #f0f0f0; color: #f0f0f0;'
        elif val == "実績(予定内)":
            # 青色
            return 'background-color: #339af0; color: #339af0;'
        elif val == "実績(超過)":
            # 赤色
            return 'background-color: #ff6b6b; color: #ff6b6b;'
        else:
            # デフォルトのスタイル（背景・文字を白に）
            return 'background-color: #ffffff; color: #ffffff;'
    
    # インデックス以外の列にスタイルを適用
    subset_cols = [col for col in df.columns if col not in ['担当設備', 'お客様名', '商品名']]
    return df.style.apply(lambda col: col.map(get_style), subset=subset_cols)

def main():
    """メインのアプリケーション処理。"""
    # --- カスタムCSSを適用 ---
    st.markdown("""
        <style>
            /* アプリ全体の背景を白に設定 */
            [data-testid="stAppViewContainer"] > .main {
                background-color: #ffffff;
            }
            /* データテーブルのフォントサイズを少し大きくする */
            div[data-testid="stDataFrame"] {
                font-size: 1.1rem;
            }
        </style>
    """, unsafe_allow_html=True)

    # --- 1. ヘッダー部分のレイアウト ---
    col1, col2 = st.columns([3, 1])
    with col1:
        st.title("今日の生産進捗")
    with col2:
        # labelを非表示にしてスペースを節約
        selected_date = st.date_input("対象日を選択", date.today(), label_visibility="collapsed", key="selected_date_input")

    # --- 2. データ取得 ---
    master_df = data_loader.load_product_master()
    plan_df = data_loader.load_plan_data(selected_date)

    results_df = data_loader.load_results_data(selected_date)
    name_master = data_loader.load_name_master()

    # --- 3. ロジック実行 ---
    progress_df = progress_logic.create_progress_table(plan_df, results_df, master_df, name_master)

    # --- 4. UI表示 ---
    if progress_df.empty:
        st.info("表示対象のデータがありません。")
        st.stop()

    # 並び替え（ライン昇順 → 開始予定昇順）
    progress_df.sort_values(by=['担当設備', '予定開始時刻'], ascending=True, na_position='first', inplace=True)

    # フィルタリング機能は削除済

    # 表示する列の名称変更と順序指定
    rename_map = {
        '進捗状態': '予定',
        '担当設備': 'ライン',
        '生産数差異': '差異(数)',
        '予定開始時刻': '開始予定',
        '予定終了時刻': '終了予定',
        '実生産開始時刻': '実開始',
        '実生産終了時刻': '実終了',
        '生産時間差異(分)': '差異(分)'
    }
    display_df = progress_df.rename(columns=rename_map)

    # 表示したい列を、指定された順序で定義
    display_order = [
        '予定', 'ライン', 'お客様名', '商品名', '予定数', '実生産数', '差異(数)',
        '開始予定', '終了予定', '実開始', '実終了', '差異(分)'
    ]
    
    # 存在しない列を除外しつつ、表示用にDFを再構成
    final_display_cols = [col for col in display_order if col in display_df.columns]
    display_df = display_df[final_display_cols]
    
    st.caption("並び順：ライン > 開始予定")
    
    # スタイルを適用してテーブル表示
    styled_df = style_progress_table(display_df)

    # 差異セルのスタイルを適用
    def style_diff_cells(s):
        """差異(数)はマイナス、差異(分)はプラスの場合に背景色を変更する。"""
        styles = []
        # Seriesの名前で処理を分岐
        if s.name == '差異(数)':
            # 差異(数)はマイナスで背景色を変更
            is_negative = pd.to_numeric(s, errors='coerce') < 0
            for v in is_negative:
                styles.append('background-color: red; color: white' if v else '')
        elif s.name == '差異(分)':
            # 差異(分)はプラスで背景色を変更
            is_positive = pd.to_numeric(s, errors='coerce') > 0
            for v in is_positive:
                styles.append('background-color: red; color: white' if v else '')
        else:
            # その他の列はスタイルを適用しない
            styles = ['' for _ in s]
        return styles

    style_columns = [col for col in ['差異(数)', '差異(分)'] if col in display_df.columns]
    if style_columns:
        styled_df.apply(style_diff_cells, subset=style_columns, axis=0)

    st.dataframe(styled_df.format({
        '開始予定': '{:%H:%M}',
        '終了予定': '{:%H:%M}',
        '実開始': '{:%H:%M}',
        '実終了': '{:%H:%M}',
        '予定数': '{:.0f}',
        '実生産数': '{:.0f}',
        '差異(数)': '{:+.0f}',
        '差異(分)': '{:+.1f}',
    }, na_rep="-"))

    # --- タイムライン表示 ---
    st.markdown("---")
    st.subheader("タイムライン")

    # 凡例を追加
    legend_html = """
    <div style="display: flex; align-items: center; gap: 20px; margin-bottom: 10px; font-size: 0.9rem;">
        <div style="display: flex; align-items: center; gap: 5px;">
            <div style="width: 15px; height: 15px; background-color: #f0f0f0; border: 1px solid #ccc;"></div>
            <span>予定</span>
        </div>
        <div style="display: flex; align-items: center; gap: 5px;">
            <div style="width: 15px; height: 15px; background-color: #339af0; border: 1px solid #ccc;"></div>
            <span>実績 (予定内)</span>
        </div>
        <div style="display: flex; align-items: center; gap: 5px;">
            <div style="width: 15px; height: 15px; background-color: #ff6b6b; border: 1px solid #ccc;"></div>
            <span>実績 (超過)</span>
        </div>
    </div>
    """
    st.markdown(legend_html, unsafe_allow_html=True)
    
    timeline_df = progress_logic.create_timeline_dataframe(progress_df, selected_date)

    if not timeline_df.empty:
        # インデックス名をリセットして表示列に含める
        timeline_display_df = timeline_df.reset_index()
        
        # スタイリングを適用
        styled_timeline = style_timeline(timeline_display_df)
        
        # 時間列のコンフィグを作成
        time_column_config = {
            col: st.column_config.TextColumn(
                label=col,
                width=20, # ピクセル単位で指定
            ) for col in timeline_df.columns
        }
        # 品目情報列のコンフィグ
        info_column_config = {
            "担当設備": st.column_config.TextColumn(label="ライン", width=40),
            "お客様名": st.column_config.TextColumn(label="お客様名", width=120),
            "商品名": st.column_config.TextColumn(label="商品名", width=150),
        }
        # コンフィグをマージ
        column_config = {**info_column_config, **time_column_config}

        # データフレームを表示
        st.dataframe(
            styled_timeline, 
            hide_index=True, 
            height=400,
            use_container_width=True,
            column_config=column_config
        )
    else:
        st.info("タイムラインを表示するデータがありません。")

    

if __name__ == "__main__":
    main()
