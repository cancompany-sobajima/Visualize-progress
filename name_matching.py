import re
import unicodedata
import pandas as pd

def normalize_text(text: str) -> (str, list[str]): # 戻り値の型を修正
    """テキストを正規化する（全角→半角、大文字→小文字、記号除去）。"""
    debug_log = [] # デバッグメッセージを収集するリスト
    if not isinstance(text, str):
        debug_log.append(f"DEBUG: normalize_text input not string: {text}")
        return "", debug_log
    original_text = text
    # 全角を半角に
    text = unicodedata.normalize('NFKC', text)
    # 小文字に統一
    text = text.lower()
    # 一般的な記号や空白を削除
    text = re.sub(r'[\s\-,.()\[\]株式会社]', '', text)
    debug_log.append(f"DEBUG: normalize_text('{original_text}') -> '{text}'")
    return text, debug_log # テキストとデバッグログを返す

def find_best_match(name: str, master_dict: dict) -> (str, int, list[str]): # 戻り値の型を修正
    """
    与えられた名称に最も一致するマスタ名を返す。

    Args:
        name (str): 照合したい名称。
        master_dict (dict): {正規名: [別名リスト], ...} の辞書。

    Returns:
        tuple: (最も一致した正規名, 一致度スコア, デバッグログ)。一致しない場合は (None, 0, デバッグログ)。
    """
    debug_log = []
    debug_log.append(f"DEBUG: find_best_match for name: '{name}'")

    if not name or not master_dict:
        debug_log.append(f"DEBUG:   Input name or master_dict is empty.")
        return None, 0, debug_log

    normalized_name, norm_log = normalize_text(name) # normalize_text の戻り値に対応
    debug_log.extend(norm_log)

    best_match = None
    highest_score = 0

    for master_name, aliases in master_dict.items():
        # マスタ名自体との比較
        normalized_master_name, norm_log = normalize_text(master_name) # normalize_text の戻り値に対応
        debug_log.extend(norm_log)
        
        score, score_log = get_match_score(normalized_name, normalized_master_name) # get_match_score の戻り値に対応
        debug_log.extend(score_log)

        if score > highest_score:
            highest_score = score
            best_match = master_name

        # 別名リストとの比較
        for alias in aliases:
            normalized_alias, norm_log = normalize_text(alias) # normalize_text の戻り値に対応
            debug_log.extend(norm_log)
            
            score, score_log = get_match_score(normalized_name, normalized_alias) # get_match_score の戻り値に対応
            debug_log.extend(score_log)

            if score > highest_score:
                highest_score = score
                best_match = master_name
    
    debug_log.append(f"DEBUG:   Best match for '{name}': '{best_match}' with score {highest_score}")
    return best_match, highest_score, debug_log

def get_match_score(str1: str, str2: str) -> (int, list[str]): # 戻り値の型を修正
    """
    2つの正規化済み文字列の一致度スコアを計算する（単純な部分一致）。
    """
    debug_log = []
    debug_log.append(f"DEBUG: get_match_score('{str1}', '{str2}')")

    if not str1 or not str2:
        debug_log.append(f"DEBUG:   Condition: str1 or str2 empty -> Score: 0")
        return 0, debug_log
    
    # 完全一致
    if str1 == str2:
        debug_log.append(f"DEBUG:   Condition: str1 == str2 -> Score: 100")
        return 100, debug_log
    
    # 部分一致
    if str1 in str2 or str2 in str1:
        score = 85 + int(15 * (1 - abs(len(str1) - len(str2)) / max(len(str1), len(str2))))
        debug_log.append(f"DEBUG:   Condition: str1 partial match -> Score: {score}")
        return score, debug_log

    debug_log.append(f"DEBUG:   Condition: No match -> Score: 0")
    return 0, debug_log

def get_name_similarity_score(master_name: str, plan_master_name: str, master_original: str, plan_original: str) -> (int, list[str]): # 戻り値の型を修正
    """
    マスタ名と予定の名称を比較し、名称の一致度スコアを返す。
    - マスタ名どうしが完全一致: 100点
    - 予定に名寄せマスタ名があり、それがマスタの名称と部分一致: 80点
    - 元名称どうしが部分一致: 70点
    - それ以外: 0点
    """
    debug_log = []
    debug_log.append(f"\nDEBUG: --- get_name_similarity_score ---")
    debug_log.append(f"DEBUG:   master_name: '{master_name}'")
    debug_log.append(f"DEBUG:   plan_master_name: '{plan_master_name}'")
    debug_log.append(f"DEBUG:   master_original: '{master_original}'")
    debug_log.append(f"DEBUG:   plan_original: '{plan_original}'")

    # 予定の名寄せ後とマスタ名が完全一致
    if plan_master_name and plan_master_name == master_name:
        debug_log.append(f"DEBUG:   Condition: plan_master_name == master_name -> Score: 100")
        return 100, debug_log
    
    # 予定の名寄せ後とマスタ名が部分一致
    if plan_master_name and (plan_master_name in master_name or master_name in plan_master_name):
        debug_log.append(f"DEBUG:   Condition: plan_master_name partial match -> Score: 80")
        return 80, debug_log

    # 元名称どうしで部分一致
    norm_master_original, log_m = normalize_text(master_original) # normalize_text の戻り値に対応
    debug_log.extend(log_m)
    norm_plan_original, log_p = normalize_text(plan_original) # normalize_text の戻り値に対応
    debug_log.extend(log_p)

    if not norm_master_original or not norm_plan_original:
        debug_log.append(f"DEBUG:   Condition: normalized originals empty -> Score: 0")
        return 0, debug_log

    if norm_master_original in norm_plan_original or norm_plan_original in norm_master_original:
        debug_log.append(f"DEBUG:   Condition: normalized originals partial match -> Score: 70")
        return 70, debug_log
    
    debug_log.append(f"DEBUG:   Condition: No match -> Score: 0")
    return 0, debug_log


def apply_name_matching(df: pd.DataFrame, master: dict) -> (pd.DataFrame, list[str]): # 戻り値の型を修正
    """DataFrameに名寄せを適用し、正規化された名前とスコアの列を追加する。"""
    debug_log = []
    df_copy = df.copy()
    
    # お客様名
    customer_matches_data = []
    for _, row in df_copy.iterrows():
        match, score, log = find_best_match(row['お客様名'], master.get('お客様名', {}))
        customer_matches_data.append((match, score))
        debug_log.extend(log)
    
    df_copy['正規_お客様名'] = [data[0] for data in customer_matches_data]
    df_copy['お客様名スコア'] = [data[1] for data in customer_matches_data]

    # 商品名
    product_matches_data = []
    for _, row in df_copy.iterrows():
        match, score, log = find_best_match(row['商品名'], master.get('商品名', {}))
        product_matches_data.append((match, score))
        debug_log.extend(log)

    df_copy['正規_商品名'] = [data[0] for data in product_matches_data]
    df_copy['商品名スコア'] = [data[1] for data in product_matches_data]
    
    return df_copy, debug_log


def find_matching_product(plan_row: pd.Series, master_df: pd.DataFrame, name_master: dict) -> (pd.Series, list[str]): # 戻り値の型を修正
    """
    生産予定の行情報に最も一致する商品マスタの行を返す。
    """
    debug_log = []
    debug_log.append(f"\nDEBUG: --- find_matching_product for plan_row: {plan_row.get('お客様名', '')} - {plan_row.get('商品名', '')} ---")

    # 1. 名寄せマスタを使って、予定のお客様名と商品名を正規化
    plan_customer_name = plan_row.get('お客様名', '')
    plan_product_name = plan_row.get('商品名', '')

    matched_customer, _, log_cust = find_best_match(plan_customer_name, name_master.get('お客様名', {})) # find_best_match の戻り値に対応
    debug_log.extend(log_cust)
    matched_product, _, log_prod = find_best_match(plan_product_name, name_master.get('商品名', {})) # find_best_match の戻り値に対応
    debug_log.extend(log_prod)

    # 候補を絞り込むためのスコアリング
    scores = []
    for idx, master_row in master_df.iterrows():
        current_log = []
        score = 0
        
        # お客様名の一致度を評価
        customer_score, log_cs = get_name_similarity_score( # get_name_similarity_score の戻り値に対応
            master_name=master_row.get('お客様名', ''),
            plan_master_name=matched_customer,
            master_original=master_row.get('お客様名', ''),
            plan_original=plan_customer_name
        )
        current_log.extend(log_cs)
        
        # 商品名の一致度を評価
        product_score, log_ps = get_name_similarity_score( # get_name_similarity_score の戻り値に対応
            master_name=master_row.get('商品名', ''),
            plan_master_name=matched_product,
            master_original=master_row.get('商品名', ''),
            plan_original=plan_product_name
        )
        current_log.extend(log_ps)

        # お客様名が完全一致ならスコアを大きく加算
        if customer_score == 100:
            score += 100
        elif customer_score > 0:
            score += customer_score
        
        # 商品名が完全一致ならスコアを加算
        if product_score == 100:
            score += 50
        elif product_score > 0:
            score += product_score / 2 # 商品名の一致度は少し低めに評価

        scores.append(score)
        debug_log.append(f"DEBUG:   Master Row {idx} ('{master_row.get('お客様名', '')}' - '{master_row.get('商品名', '')}'): Customer Score={customer_score}, Product Score={product_score}, Total Score={score}")
        debug_log.extend(current_log) # 各候補のスコアリングログも追加

    if not scores or max(scores) == 0:
        debug_log.append(f"DEBUG:   No matching product found or all scores are 0.")
        return pd.Series(dtype='object'), debug_log # 一致するものがなければ空のSeriesとログ

    # 最もスコアの高いマスタ行を返す
    best_match_index = scores.index(max(scores))
    best_match_row = master_df.iloc[best_match_index]
    debug_log.append(f"DEBUG:   Best match found: '{best_match_row.get('お客様名', '')}' - '{best_match_row.get('商品名', '')}' with score {max(scores)}")
    return best_match_row, debug_log