import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime

import name_matching 

from name_matching import apply_name_matching, get_name_similarity_score

def create_progress_table(plan_df, results_df, master_df, name_master):
    """
    新しいメインロジック：
    1. 予定表の名称を商品マスタでクリーンナップ
    2. クリーンになった予定表と実績表を突合
    """
    # 1. 予定表の名称を商品マスタを使いクリーンナップ
    cleaned_plan_df = _clean_plan_with_master(plan_df, master_df, name_master)

    # 2. 名称がクリーンになった予定表と実績表を突合
    final_df = _merge_plan_and_results(cleaned_plan_df, results_df)
    
    # 3. 差異と進捗状態を計算
    if not final_df.empty:
        final_df = calculate_differences_and_status(final_df)

    # 4. 予定開始または終了がNaTの行を削除 (主に予定外生産で時刻がないもの)
    if not final_df.empty:
        # 予定がない（NaT）が、実績がある（NotNaT）ものは残す
        is_unplanned_but_valid = final_df['予定開始時刻'].isna() & final_df['実生産開始時刻'].notna()
        # 予定があるものは、開始・終了の両方が揃っている必要がある
        is_planned_and_valid = final_df['予定開始時刻'].notna() & final_df['予定終了時刻'].notna()
        
        final_df = final_df[is_unplanned_but_valid | is_planned_and_valid]

    return final_df

def _clean_plan_with_master(plan_df, master_df, name_master): # 戻り値型修正
    """予定表の各行を、商品マスタと照合し、お客様名・商品名をクリーンなものに更新する"""
    if plan_df.empty:
        return pd.DataFrame()

    # まず、予定表の表記揺れを「振れ幅表(name_master)」で吸収する
    plan_df_matched = name_matching.apply_name_matching(plan_df, name_master) # apply_name_matching の戻り値に対応
    
    cleaned_rows = []
    for idx, plan_row in plan_df_matched.iterrows():
        new_row = plan_row.to_dict()
        
        # この予定に最も一致するマスタ品目を探す
        best_master_row = _find_best_master_for_plan(plan_row, master_df) # _find_best_master_for_plan の戻り値に対応
        
        if not best_master_row.empty: # pd.Series.empty で判定
            # マッチしたら、マスタの綺麗な名称で上書き
            new_row['お客様名'] = best_master_row['お客様名']
            new_row['商品名'] = best_master_row['商品名']
        
        cleaned_rows.append(new_row)
        
    return pd.DataFrame(cleaned_rows)

def _find_best_master_for_plan(plan_row, master_df):
    """特定の予定に最も一致するマスタ品目を、スコアリングに基づいて見つける"""
    debug_log = []
    debug_log.append(f"\nDEBUG: --- _find_best_master_for_plan for plan_row: {plan_row.get('お客様名', '')} - {plan_row.get('商品名', '')} ---")

    best_match_row = pd.Series(dtype='object') # Noneではなく空のSeriesで初期化

    # 補助関数: 完全一致判定
    def _is_exact_match(name1_norm, name2_norm):
        return name1_norm == name2_norm

    # 補助関数: 部分一致判定
    def _is_partial_match(name1_norm, name2_norm):
        return name1_norm in name2_norm or name2_norm in name1_norm

    # 1. ラインが一致するマスタ品目に候補を絞る
    candidate_masters_by_line = master_df[master_df['担当設備'] == plan_row['担当設備']]
    debug_log.append(f"DEBUG:   Candidates by line '{plan_row['担当設備']}': {len(candidate_masters_by_line)} rows")
    if candidate_masters_by_line.empty:
        debug_log.append(f"DEBUG:   No candidates found for line '{plan_row['担当設備']}'. Returning empty series.")
        return best_match_row, debug_log # このラインの候補がない

    # 予定の正規化済みお客様名と商品名を取得
    normalized_plan_customer = name_matching.normalize_text(plan_row['お客様名'])
    normalized_plan_product = name_matching.normalize_text(plan_row['商品名'])

    # --- 優先度1: 顧客名 完全一致 & 商品名 完全一致 ---
    for _, master_row in candidate_masters_by_line.iterrows():
        normalized_master_customer = name_matching.normalize_text(master_row['お客様名'])
        normalized_master_product = name_matching.normalize_text(master_row['商品名'])

        if _is_exact_match(normalized_plan_customer, normalized_master_customer) and \
           _is_exact_match(normalized_plan_product, normalized_master_product):
            return master_row # 完璧な一致が見つかったら即座に返す

    # --- 優先度2: 顧客名 完全一致 & 商品名 部分一致 ---
    # 顧客名が完全一致する候補をフィルタリング
    exact_customer_candidates = []
    for _, master_row in candidate_masters_by_line.iterrows():
        normalized_master_customer = name_matching.normalize_text(master_row['お客様名'])
        if _is_exact_match(normalized_plan_customer, normalized_master_customer):
            exact_customer_candidates.append(master_row)

    if exact_customer_candidates:
        # その中で商品名が部分一致するものを探す
        for _, master_row in pd.DataFrame(exact_customer_candidates).iterrows():
            normalized_master_product = name_matching.normalize_text(master_row['商品名'])
            if _is_partial_match(normalized_plan_product, normalized_master_product):
                return master_row # 見つかったら即座に返す (最初に見つかった部分一致)

    # --- 優先度3: 顧客名 部分一致 & 商品名 完全一致 ---
    # 顧客名が部分一致する候補をフィルタリング (ただし、完全一致は既に処理済み)
    partial_customer_candidates = []
    for _, master_row in candidate_masters_by_line.iterrows():
        normalized_master_customer = name_matching.normalize_text(master_row['お客様名'])
        if not _is_exact_match(normalized_plan_customer, normalized_master_customer) and \
           _is_partial_match(normalized_plan_customer, normalized_master_customer):
            partial_customer_candidates.append(master_row)

    if partial_customer_candidates:
        # その中で商品名が完全一致するものを探す
        for _, master_row in pd.DataFrame(partial_customer_candidates).iterrows():
            normalized_master_product = name_matching.normalize_text(master_row['商品名'])
            if _is_exact_match(normalized_plan_product, normalized_master_product):
                return master_row # 見つかったら即座に返す

    # --- 優先度4: 顧客名 部分一致 & 商品名 部分一致 ---
    if partial_customer_candidates: # 優先度3でフィルタリングした候補を再利用
        # その中で商品名が部分一致するものを探す
        for _, master_row in pd.DataFrame(partial_customer_candidates).iterrows():
            normalized_master_product = name_matching.normalize_text(master_row['商品名'])
            if _is_partial_match(normalized_plan_product, normalized_master_product):
                return master_row # 見つかったら即座に返す

    # --- どの条件にも合致しない場合 ---
    return best_match_row

def _merge_plan_and_results(cleaned_plan_df, results_df):
    """クリーンな予定表と実績表をマージする。日付も考慮する。"""
    
    # 予定データに日付列を追加 (マージキーとして使用)
    if not cleaned_plan_df.empty and '予定開始時刻' in cleaned_plan_df.columns:
        cleaned_plan_df['日付'] = pd.to_datetime(cleaned_plan_df['予定開始時刻']).dt.date

    # 実績が同じ日に複数ある可能性を考慮し、キーで集計しておく
    agg_results_df = pd.DataFrame()
    if not results_df.empty:
        key_cols = ['日付', '担当設備', 'お客様名', '商品名']
        
        agg_dict = {
            '実生産開始時刻': ('実生産開始時刻', 'min'),
            '実生産終了時刻': ('実生産終了時刻', 'max'),
            '実生産数': ('実生産数', 'sum'),
            '実績総生産時間_分': ('実績総生産時間_分', 'sum'),
            '実セッション開始時刻リスト': ('実セッション開始時刻リスト', 'sum'),
            '実セッション終了時刻リスト': ('実セッション終了時刻リスト', 'sum'),
        }
        
        agg_dict_filtered = {k: v for k, v in agg_dict.items() if v[0] in results_df.columns}
        if agg_dict_filtered:
            agg_results_df = results_df.groupby(key_cols).agg(**agg_dict_filtered).reset_index()

    # 予定と実績をouter joinで結合
    if not cleaned_plan_df.empty and not agg_results_df.empty:
        key_cols = ['日付', '担当設備', 'お客様名', '商品名']
        merged_df = pd.merge(cleaned_plan_df, agg_results_df, on=key_cols, how='outer')
    elif not cleaned_plan_df.empty:
        merged_df = cleaned_plan_df.copy()
    elif not agg_results_df.empty:
        merged_df = agg_results_df.copy()
    else:
        merged_df = pd.DataFrame()

    if not merged_df.empty:
        # マージ後にデータ型を明示的に変換し、安全性を高める
        for col in ['予定数', '実生産数']:
            if col in merged_df.columns:
                merged_df[col] = pd.to_numeric(merged_df[col], errors='coerce')
        
        # 実績関連の列が存在しない場合に備えて、列を確保する
        result_cols = ['実生産開始時刻', '実生産終了時刻', '実生産数', '実績総生産時間_分']
        for col in result_cols:
            if col not in merged_df.columns:
                if '時刻' in col:
                    merged_df[col] = pd.NaT
                else:
                    merged_df[col] = np.nan
    
    return merged_df


def calculate_differences_and_status(df):
    """生産数差異、時間差異、進捗状態を計算する。"""
    # datetime型への変換を確実に行う
    time_cols = ['予定開始時刻', '予定終了時刻', '実生産開始時刻', '実生産終了時刻']
    for col in time_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')

    # 生産数差異
    df['生産数差異'] = pd.to_numeric(df['実生産数'], errors='coerce').fillna(0) - pd.to_numeric(df['予定数'], errors='coerce').fillna(0)

    # 生産時間差異
    planned_duration = np.nan
    if '予定終了時刻' in df.columns and '予定開始時刻' in df.columns:
        valid_times = df['予定終了時刻'].notna() & df['予定開始時刻'].notna()
        planned_duration = (df.loc[valid_times, '予定終了時刻'] - df.loc[valid_times, '予定開始時刻']).dt.total_seconds() / 60
    
    # 実績総生産時間_分 は data_loader で計算済み。NaNは0で埋める。
    actual_duration = df['実績総生産時間_分'].fillna(0)

    df['生産時間差異(分)'] = actual_duration - planned_duration.fillna(0)

    # 進捗状態
    df['進捗状態'] = df.apply(get_status, axis=1)
    
    return df

def get_status(row):
    """行データから進捗状態を判定する。"""
    now = datetime.now()
    
    is_planned = pd.notna(row.get('予定開始時刻'))
    is_resulted = pd.notna(row.get('実生産開始時刻'))

    if is_planned and not is_resulted:
        if now > row['予定終了時刻']:
            return "遅延(未開始)"
        else:
            return "未開始"
    elif is_planned and is_resulted:
        if pd.isna(row.get('実生産終了時刻')):
            if now > row['予定終了時刻']:
                return "遅延(進行中)"
            else:
                return "進行中"
        else:
            if row['実生産終了時刻'] > row['予定終了時刻']:
                return "完了(遅延)"
            else:
                return "完了"
    elif not is_planned and is_resulted:
        return "予定外"
    
    return "---" # 予定も実績もない（マスタのみ）

def create_timeline_dataframe(progress_df, target_date):
    """タイムライン表示用のDataFrameを生成する。セルの状態を3つに分ける。"""
    if progress_df.empty:
        return pd.DataFrame()

    # 時間軸を15分単位で生成 (指定された日付の8:30から17:00まで)
    date_str = target_date.strftime('%Y-%m-%d')
    time_slots = pd.to_datetime(pd.date_range(f"{date_str} 08:30", f"{date_str} 17:01", freq="15min"))
    time_labels = [t.strftime('%H:%M') for t in time_slots]

    # 表示するレコードを特定（予定または実績があるもの）
    # 実セッション開始時刻リストが存在しない場合は空リストとして扱う
    df = progress_df.copy()
    if '実セッション開始時刻リスト' not in df.columns:
        df['実セッション開始時刻リスト'] = [[] for _ in range(len(df))]
    if '実セッション終了時刻リスト' not in df.columns:
        df['実セッション終了時刻リスト'] = [[] for _ in range(len(df))]

    # 予定または実績セッションがある行のみを対象とする
    df = df[
        (df['予定開始時刻'].notna() & df['予定終了時刻'].notna()) | 
        (df['実セッション開始時刻リスト'].apply(lambda x: isinstance(x, list) and len(x) > 0))
    ].copy()

    if df.empty:
        return pd.DataFrame(columns=time_labels)

    # MultiIndexを作成
    key_cols = ['担当設備', 'お客様名', '商品名']
    for col in key_cols:
        if col not in df.columns:
            df[col] = 'N/A'
    df = df.fillna({col: 'N/A' for col in key_cols})

    df.set_index(key_cols, inplace=True)
    df.sort_index(inplace=True)
    
    # タイムライン用の空のDataFrameを作成
    timeline_df = pd.DataFrame("", index=df.index, columns=time_labels)

    for idx, row in df.iterrows():
        # 1. 予定をプロット
        if pd.notna(row['予定開始時刻']) and pd.notna(row['予定終了時刻']):
            plan_start_time = row['予定開始時刻']
            plan_end_time = row['予定終了時刻']
            for i, slot_start in enumerate(time_slots):
                slot_end = slot_start + pd.Timedelta(minutes=15)
                # 予定の期間がタイムスロットと重なる場合
                if plan_start_time < slot_end and plan_end_time > slot_start:
                    timeline_df.loc[idx, time_labels[i]] = "予定"

        # 2. 実績をプロット (予定を上書き) - 各セッションを個別にプロット
        if isinstance(row['実セッション開始時刻リスト'], list) and isinstance(row['実セッション終了時刻リスト'], list):
            for session_start, session_end in zip(row['実セッション開始時刻リスト'], row['実セッション終了時刻リスト']):
                # セッションが有効な時刻範囲内にあるか確認
                if pd.notna(session_start) and pd.notna(session_end):
                    plan_end_time = row.get('予定終了時刻') # 予定終了時刻は各行で共通

                    for i, slot_start in enumerate(time_slots):
                        slot_end = slot_start + pd.Timedelta(minutes=15)
                        # 実績セッションの期間がタイムスロットと重なる場合
                        if session_start < slot_end and session_end > slot_start:
                            # この時間帯が実績の範囲内である
                            # それが予定時刻を超過しているかを判定
                            # 予定終了時刻がNaTでない、かつ、スロット開始が予定終了時刻以降の場合
                            if pd.notna(plan_end_time) and slot_start >= plan_end_time:
                                timeline_df.loc[idx, time_labels[i]] = "実績(超過)"
                            else:
                                timeline_df.loc[idx, time_labels[i]] = "実績(予定内)"
    
    return timeline_df
