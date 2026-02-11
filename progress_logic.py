import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime

import name_matching 

from name_matching import apply_name_matching, get_name_similarity_score

def create_progress_table(plan_df, results_df, master_df, name_master) -> (pd.DataFrame, list[str]): # 戻り値の型を修正
    """
    新しいメインロジック：
    1. 予定表の名称を商品マスタでクリーンナップ
    2. クリーンになった予定表と実績表を突合
    """
    all_debug_logs = []
    all_debug_logs.append(f"\nDEBUG: --- create_progress_table ---")

    # 1. 予定表の名称を商品マスタを使いクリーンナップ
    cleaned_plan_df, log_cpm = _clean_plan_with_master(plan_df, master_df, name_master) # _clean_plan_with_master の戻り値に対応
    all_debug_logs.extend(log_cpm)

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

    all_debug_logs.append(f"DEBUG: --- create_progress_table finished ---")
    return final_df, all_debug_logs
logs

def _clean_plan_with_master(plan_df, master_df, name_master) -> (pd.DataFrame, list[str]): # 戻り値の型を修正
    """予定表の各行を、商品マスタと照合し、お客様名・商品名をクリーンなものに更新する"""
    debug_log = []
    debug_log.append(f"\nDEBUG: --- _clean_plan_with_master ---")

    if plan_df.empty:
        debug_log.append(f"DEBUG:   plan_df is empty.")
        return pd.DataFrame(), debug_log

    # まず、予定表の表記揺れを「振れ幅表(name_master)」で吸収する
    plan_df_matched, log_apm = name_matching.apply_name_matching(plan_df, name_master) # apply_name_matching の戻り値に対応
    debug_log.extend(log_apm)
    
    cleaned_rows = []
    for idx, plan_row in plan_df_matched.iterrows():
        new_row = plan_row.to_dict()
        
        # この予定に最も一致するマスタ品目を探す
        best_master_row, log_fbm = _find_best_master_for_plan(plan_row, master_df) # _find_best_master_for_plan の戻り値に対応
        debug_log.extend(log_fbm)
        
        if not best_master_row.empty: # pd.Series.empty で判定
            # マッチしたら、マスタの綺麗な名称で上書き
            new_row['お客様名'] = best_master_row['お客様名']
            new_row['商品名'] = best_master_row['商品名']
            debug_log.append(f"DEBUG:   Cleaned plan row {idx}: Customer='{new_row['お客様名']}', Product='{new_row['商品名']}'")
        else:
            debug_log.append(f"DEBUG:   No master match for plan row {idx}. Keeping original names.")
        
        cleaned_rows.append(new_row)
        
    return pd.DataFrame(cleaned_rows), debug_log

def _find_best_master_for_plan(plan_row, master_df) -> (pd.Series, list[str]):
    """特定の予定に最も一致するマスタ品目を、スコアリングに基づいて見つける"""
    debug_log = []
    debug_log.append(f"\nDEBUG: --- _find_best_master_for_plan for plan_row: {plan_row.get('お客様名', '')} - {plan_row.get('商品名', '')} ---")

    best_candidate_row = pd.Series(dtype='object') # Noneではなく空のSeriesで初期化
    highest_score = 0 # 全体で最も高いスコアを追跡

    # 1. ラインが一致するマスタ品目に候補を絞る
    candidate_masters_by_line = master_df[master_df['担当設備'] == plan_row['担当設備']]
    debug_log.append(f"DEBUG:   Candidates by line '{plan_row['担当設備']}': {len(candidate_masters_by_line)} rows")

    # 予定の正規化済みお客様名と商品名を取得
    normalized_plan_customer, log_npc = name_matching.normalize_text(plan_row['お客様名']) # normalize_text の戻り値に対応
    debug_log.extend(log_npc)
    normalized_plan_product, log_npp = name_matching.normalize_text(plan_row['商品名']) # normalize_text の戻り値に対応
    debug_log.extend(log_npp)

    # --- フェーズ1: 顧客名が完全に一致する候補を探し、その中で商品名をスコアリング ---
    exact_customer_candidates = []
    for _, master_row in candidate_masters_by_line.iterrows():
        normalized_master_customer, log_nmc = name_matching.normalize_text(master_row['お客様名']) # normalize_text の戻り値に対応
        debug_log.extend(log_nmc)
        if normalized_plan_customer == normalized_master_customer:
            exact_customer_candidates.append(master_row)
    debug_log.append(f"DEBUG:   Phase 1: Exact customer matches found: {len(exact_customer_candidates)} rows")

    if exact_customer_candidates:
        # 顧客名が完全に一致する候補が見つかった場合、その中で商品名をスコアリング
        for master_row in exact_customer_candidates:
            score_this_candidate = 0
            normalized_master_product, log_nmp = name_matching.normalize_text(master_row['商品名']) # normalize_text の戻り値に対応
            debug_log.extend(log_nmp)

            if normalized_plan_product == normalized_master_product:
                score_this_candidate = 1000 # 商品名も完全一致なら最高スコア
                debug_log.append(f"DEBUG:     Exact Product Match (Phase 1): '{master_row['商品名']}' -> Score: {score_this_candidate}")
            else:
                # 顧客名が一致しているので、商品名のみでスコアリング (重み100%)
                prod_score, log_ps = name_matching.get_name_similarity_score( # get_name_similarity_score の戻り値に対応
                    master_name=master_row['商品名'],
                    plan_master_name=plan_row['正規_商品名'],
                    master_original=master_row['商品名'],
                    plan_original=plan_row['商品名']
                )
                debug_log.extend(log_ps)
                score_this_candidate = (prod_score / 100) * 100 # 商品名スコアをそのまま利用
                debug_log.append(f"DEBUG:     Product Scoring (Phase 1): '{master_row['商品名']}' -> Score: {score_this_candidate}")

            if score_this_candidate > highest_score:
                highest_score = score_this_candidate
                best_candidate_row = master_row
        
        # フェーズ1で有効なマッチが見つかった場合、ここで確定し、フェーズ2は実行しない
        if highest_score > 0: # 顧客名一致かつ商品名マッチがあった場合
            debug_log.append(f"DEBUG:   Phase 1 Best Match: '{best_candidate_row.get('お客様名', '')}' - '{best_candidate_row.get('商品名', '')}' with score {highest_score}")
            return best_candidate_row, debug_log

    # --- フェーズ2: 顧客名が完全に一致する候補が見つからない、またはフェーズ1で商品名がマッチしなかった場合 ---
    debug_log.append(f"DEBUG:   Phase 2: Falling back to general scoring.")
    # 従来通りの方法で、ラインが一致する全ての候補に対してスコアリング
    # highest_score はフェーズ1で更新されている可能性があるので、それを引き継ぐ
    for _, master_row in candidate_masters_by_line.iterrows():
        score_this_candidate = 0
        
        # お客様名スコア (配点60)
        cust_score, log_cs = name_matching.get_name_similarity_score( # get_name_similarity_score の戻り値に対応
            master_name=master_row['お客様名'],
            plan_master_name=plan_row['正規_お客様名'],
            master_original=master_row['お客様名'],
            plan_original=plan_row['お客様名']
        )
        debug_log.extend(log_cs)
        score_this_candidate += (cust_score / 100) * 60

        # 商品名スコア (配点40)
        prod_score, log_ps = name_matching.get_name_similarity_score( # get_name_similarity_score の戻り値に対応
            master_name=master_row['商品名'],
            plan_master_name=plan_row['正規_商品名'],
            master_original=master_row['商品名'],
            plan_original=plan_row['商品名']
        )
        debug_log.extend(log_ps)
        score_this_candidate += (prod_score / 100) * 40
        
        debug_log.append(f"DEBUG:     Candidate (Phase 2): '{master_row.get('お客様名', '')}' - '{master_row.get('商品名', '')}' -> Cust Score: {cust_score}, Prod Score: {prod_score}, Total Score: {score_this_candidate}")

        # 閾値: お客様名が部分一致(70*0.6=42)すれば候補
        if score_this_candidate > highest_score and score_this_candidate > 40:
            highest_score = score_this_candidate
            best_candidate_row = master_row
            
    debug_log.append(f"DEBUG:   Final Best Match: '{best_candidate_row.get('お客様名', '')}' - '{best_candidate_row.get('商品名', '')}' with score {highest_score}")
    return best_candidate_row, debug_log

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
