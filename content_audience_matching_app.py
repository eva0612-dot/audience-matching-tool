import pandas as pd
import jieba
import re
from collections import defaultdict
from flask import Flask, request, render_template_string
import random

app = Flask(__name__)

# 讀取 Excel 文件
def load_data(file_path):
    data_df = pd.read_excel(file_path, sheet_name="Data")
    keywords_df = pd.read_excel(file_path, sheet_name="Keywords")
    return data_df, keywords_df

# 文本預處理：清理並使用 jieba 分詞
def preprocess_text(text):
    text = re.sub(r'[^\w\s]', '', text)
    words = jieba.cut(text, cut_all=False)
    return [word.strip() for word in words if word.strip()]

# 提取文章關鍵字
def extract_keywords(text, top_n=10):
    words = preprocess_text(text)
    word_freq = defaultdict(int)
    for word in words:
        word_freq[word] += 1
    return [word for word, _ in sorted(word_freq.items(), key=lambda x: x[1], reverse=True)[:top_n]]

# 匹配受眾
def match_audience(article_keywords, keywords_df):
    audience_scores = defaultdict(float)
    for _, row in keywords_df.iterrows():
        audience = row['受眾分群']
        keywords = row['關鍵字'].split(',')
        keywords = [kw.strip() for kw in keywords]
        overlap = len(set(article_keywords) & set(keywords))
        score =
