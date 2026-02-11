import re
import unicodedata
import pandas as pd
from fuzzywuzzy import fuzz

def normalize_text(text: str) -> str:
    """テキストを正規化する（全角→半角、大文字→小文字、記号除去）。"""
    if not isinstance(text, str):
        return ""
    # 全角を半角に
    text = unicodedata.normalize('NFKC', text)
    # 小文字に統一
    text = text.lower()
    # 一般的な記号や空白を削除
    text = re.sub(r'[\s\-,.()\[\]株式会社]', '', text)
    return text

def find_best_match(name: str, master_dict: dict) -> (str, int):
    """
    与えられた名称に最も一致するマスタ名を返す。

    Args:
        name (str): 照合したい名称。
        master_dict (dict): {正規名: [別名リスト], ...} の辞書。

    Returns:
        tuple: (最も一致した正規名, 一致度スコア)。一致しない場合は (None, 0)。
    """
    if not name or not master_dict:
        return None, 0

    normalized_name = normalize_text(name)
    best_match = None
    highest_score = 0

    for master_name, aliases in master_dict.items():
        # マスタ名自体との比較
        score = get_match_score(normalized_name, normalize_text(master_name))
        if score > highest_score:
            highest_score = score
            best_match = master_name

        # 別名リストとの比較
        for alias in aliases:
            score = get_match_score(normalized_name, normalize_text(alias))
            if score > highest_score:
                highest_score = score
                best_match = master_name
    
    return best_match, highest_score

def get_match_score(str1: str, str2: str) -> int:
    """
    2つの正規化済み文字列の一致度スコアを計算する（fuzzywuzzyを使用）。
    """
    if not str1 or not str2:
        return 0
    
    # fuzzywuzzyのtoken_set_ratioを使用してスコアを計算
    # これは、文字列内の単語の順序や重複を考慮しつつ、最も高い類似度を返す
    score = fuzz.token_set_ratio(str1, str2)
    
    return score

def get_name_similarity_score(master_name: str, plan_master_name: str, master_original: str, plan_original: str) -> int:
    """
    マスタ名と予定の名称を比較し、名称の一致度スコアを返す。
    - マスタ名どうしが完全一致: 100点
    - 予定に名寄せマスタ名があり、それがマスタの名称と部分一致: 80点
    - 元名称どうしが部分一致: 70点
    - それ以外: 0点
    """
    # 予定の名寄せ後とマスタ名が完全一致
    if plan_master_name and plan_master_name == master_name:
        return 100
    
    # 予定の名寄せ後とマスタ名が部分一致
    if plan_master_name and (plan_master_name in master_name or master_name in plan_master_name):
        return 80

    # 元名称どうしで部分一致
    norm_master_original = normalize_text(master_original)
    norm_plan_original = normalize_text(plan_original)

    if not norm_master_original or not norm_plan_original:
        return 0

    if norm_master_original in norm_plan_original or norm_plan_original in norm_master_original:
        return 70
    
    return 0

def apply_name_matching(df: pd.DataFrame, master: dict) -> pd.DataFrame:
    """DataFrameに名寄せを適用し、正規化された名前とスコアの列を追加する。"""
    df_copy = df.copy()
    
    # お客様名
    customer_matches = df_copy['お客様名'].apply(lambda x: find_best_match(x, master.get('お客様名', {})))
    df_copy['正規_お客様名'] = [match[0] for match in customer_matches]
    df_copy['お客様名スコア'] = [match[1] for match in customer_matches]

    # 商品名
    product_matches = df_copy['商品名'].apply(lambda x: find_best_match(x, master.get('商品名', {})))
    df_copy['正規_商品名'] = [match[0] for match in product_matches]
    df_copy['商品名スコア'] = [match[1] for match in product_matches]
    
    return df_copy


def find_matching_product(plan_row: pd.Series, master_df: pd.DataFrame, name_master: dict) -> pd.Series:
    """
    生産予定の行情報に最も一致する商品マスタの行を返す。
    """
    # 1. 名寄せマスタを使って、予定のお客様名と商品名を正規化
    plan_customer_name = plan_row.get('お客様名', '')
    plan_product_name = plan_row.get('商品名', '')

    matched_customer, _ = find_best_match(plan_customer_name, name_master.get('お客様名', {}))
    matched_product, _ = find_best_match(plan_product_name, name_master.get('商品名', {}))

    # 候補を絞り込むためのスコアリング
    scores = []
    for _, master_row in master_df.iterrows():
        score = 0
        
        # お客様名の一致度を評価
        customer_score = get_name_similarity_score(
            master_name=master_row.get('お客様名', ''),
            plan_master_name=matched_customer,
            master_original=master_row.get('お客様名', ''),
            plan_original=plan_customer_name
        )
        
        # 商品名の一致度を評価
        product_score = get_name_similarity_score(
            master_name=master_row.get('商品名', ''),
            plan_master_name=matched_product,
            master_original=master_row.get('商品名', ''),
            plan_original=plan_product_name
        )

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

    if not scores or max(scores) == 0:
        return pd.Series(dtype='object') # 一致するものがなければ空のSeries

    # 最もスコアの高いマスタ行を返す
    best_match_index = scores.index(max(scores))
    return master_df.iloc[best_match_index]