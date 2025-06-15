```python
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
        score = overlap / max(len(keywords), 1)
        audience_scores[audience] = score
    return sorted(audience_scores.items(), key=lambda x: x[1], reverse=True)

# 計算文章熱度分數
def calculate_heat_score(keywords, keywords_df):
    # 簡單熱度分數：基於關鍵字重疊和受眾市場預測
    max_score = 0
    for _, row in keywords_df.iterrows():
        audience_keywords = [kw.strip() for kw in row['關鍵字'].split(',')]
        overlap = len(set(keywords) & set(audience_keywords))
        # 假設市場預測影響熱度（可根據「歷史資料」中的市場預測加權）
        market_weight = 1.0  # 未來可替換為實際市場數據
        score = (overlap / max(len(audience_keywords), 1)) * 100 * market_weight
        max_score = max(max_score, score)
    return min(max_score, 100)  # 限制分數在 0-100

# 自動生成文章
def generate_article(audience, keywords_df):
    # 找到指定受眾的關鍵字
    audience_row = keywords_df[keywords_df['受眾分群'] == audience]
    if audience_row.empty:
        return "找不到指定受眾的關鍵字！"
    keywords = audience_row.iloc[0]['關鍵字'].split(',')
    keywords = [kw.strip() for kw in keywords]
    # 簡單模板生成文章
    template = (
        f"探索如何透過{'、'.join(random.sample(keywords, min(3, len(keywords))))}改善您的生活！\n\n"
        f"您是否感到{keywords[0]}的困擾？別擔心！根據專家建議，您可以嘗試以下方法：\n"
        f"1. 結合{random.choice(keywords)}，例如每天花10分鐘進行簡單的放鬆練習。\n"
        f"2. 注重{random.choice(keywords)}，這有助於提升您的健康與幸福感。\n"
        f"3. 尋求專業建議，針對{random.choice(keywords)}進行個人化調整。\n\n"
        f"立即行動，讓您的生活更美好！"
    )
    return template

# 網頁模板
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <title>內容分析與受眾匹配工具</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #333; }
        textarea { width: 100%; height: 200px; }
        select { padding: 10px; width: 200px; }
        button { padding: 10px 20px; background-color: #007bff; color: white; border: none; cursor: pointer; }
        button:hover { background-color: #0056b3; }
        .result { margin-top: 20px; }
        .section { margin-bottom: 20px; }
    </style>
</head>
<body>
    <h1>內容分析與受眾匹配工具</h1>
    
    <div class="section">
        <h2>分析文章受眾</h2>
        <form method="POST" action="/analyze">
            <label for="content">請輸入文章內容：</label><br>
            <textarea id="content" name="content"></textarea><br>
            <button type="submit">分析受眾與熱度</button>
        </form>
    </div>
    
    <div class="section">
        <h2>生成文章草稿</h2>
        <form method="POST" action="/generate">
            <label for="audience">選擇目標受眾：</label><br>
            <select id="audience" name="audience">
                {% for audience in audiences %}
                <option value="{{ audience }}">{{ audience }}</option>
                {% endfor %}
            </select><br>
            <button type="submit">生成文章</button>
        </form>
    </div>
    
    {% if result %}
    <div class="result">
        <h2>分析結果</h2>
        <p><strong>推薦受眾：</strong> {{ result.top_audience }} (分數: {{ result.score|round(2) }})</p>
        <p><strong>文章熱度分數：</strong> {{ result.heat_score|round(2) }} / 100</p>
        <p><strong>前三名匹配受眾：</strong></p>
        <ul>
        {% for audience, score in result.top_three %}
            <li>{{ audience }}: {{ score|round(2) }}</li>
        {% endfor %}
        </ul>
        <p><strong>提取的關鍵字：</strong> {{ result.keywords|join(', ') }}</p>
    </div>
    {% endif %}
    
    {% if generated_article %}
    <div class="result">
        <h2>生成文章草稿</h2>
        <p>{{ generated_article|replace('\n', '<br>') }}</p>
    </div>
    {% endif %}
</body>
</html>
"""

# 載入受眾列表
_, keywords_df = load_data('0612內容分析與受眾匹配的智慧行銷工具.xlsx')
audiences = keywords_df['受眾分群'].tolist()

# 分析路由
@app.route('/analyze', methods=['POST'])
def analyze():
    content = request.form['content']
    if not content:
        return render_template_string(HTML_TEMPLATE, audiences=audiences, error="請輸入文章內容！")
    keywords = extract_keywords(content)
    audience_ranking = match_audience(keywords, keywords_df)
    heat_score = calculate_heat_score(keywords, keywords_df)
    result = {
        'top_audience': audience_ranking[0][0],
        'score': audience_ranking[0][1],
        'top_three': audience_ranking[:3],
        'keywords': keywords,
        'heat_score': heat_score
    }
    return render_template_string(HTML_TEMPLATE, audiences=audiences, result=result)

# 生成文章路由
@app.route('/generate', methods=['POST'])
def generate():
    audience = request.form['audience']
    generated_article = generate_article(audience, keywords_df)
    return render_template_string(HTML_TEMPLATE, audiences=audiences, generated_article=generated_article)

# 首頁路由
@app.route('/', methods=['GET'])
def index():
    return render_template_string(HTML_TEMPLATE, audiences=audiences)

if __name__ == '__main__':
    app.run(debug=True)
```