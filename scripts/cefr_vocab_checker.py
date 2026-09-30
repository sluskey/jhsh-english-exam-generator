"""
cefr_vocab_checker.py
用於檢查英語試題文章是否超出教育部 1200/2000 字表及 CEFR A2 範疇，
並自動提取超綱生詞以供試卷末尾加註中文。
"""

import re

# 範例核心高頻單字集合（可擴充完整 2000 字庫）
COMMON_STOPWORDS = {
    "a", "an", "the", "in", "on", "at", "by", "for", "with", "about", "against", "between",
    "into", "through", "during", "before", "after", "above", "below", "to", "from", "up", "down",
    "is", "am", "are", "was", "were", "be", "been", "being", "have", "has", "had", "having",
    "do", "does", "did", "doing", "would", "should", "could", "ought", "i", "you", "he", "she",
    "it", "we", "they", "me", "him", "her", "us", "them", "my", "your", "his", "their", "our",
    "mine", "yours", "hers", "ours", "theirs", "this", "that", "these", "those", "what", "which",
    "who", "whom", "whose", "where", "when", "why", "how", "all", "any", "both", "each", "few",
    "more", "most", "other", "some", "such", "no", "nor", "not", "only", "own", "same", "so",
    "than", "too", "very", "can", "will", "just", "don", "should", "now", "and", "but", "or", "because",
    "if", "as", "until", "while", "of", "off", "out"
}

def clean_and_tokenize(text: str):
    """清理文字並切分為單字"""
    words = re.findall(r"\b[A-Za-z]+(?:'[A-Za-z]+)?\b", text.lower())
    return words

def check_text_vocabulary(text: str, allowed_vocab_set: set = None):
    """
    分析文章詞彙，找出非基礎核心字彙
    """
    tokens = clean_and_tokenize(text)
    if allowed_vocab_set is None:
        allowed_vocab_set = set()
    
    unique_words = sorted(list(set(tokens)))
    potential_advanced_words = []
    
    for w in unique_words:
        if len(w) <= 2:
            continue
        if w in COMMON_STOPWORDS:
            continue
        if allowed_vocab_set and w not in allowed_vocab_set:
            potential_advanced_words.append(w)
            
    return potential_advanced_words

if __name__ == "__main__":
    sample_text = """
    Many people love to visit national parks. However, climate change and habitat destruction
    threaten endangered species around the globe. We need to preserve nature.
    """
    advanced = check_text_vocabulary(sample_text)
    print("Detected potential advanced words requiring footnote check:")
    print(advanced)
