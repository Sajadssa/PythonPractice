import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.cluster import KMeans
import re

# خواندن داده‌ها
data = pd.read_excel('D:\Sepher_Pasargad\works\فایل های تعمیرات\Ai Template Title.xlsx', sheet_name='Sheet2')

# پیش‌پردازش متن
def preprocess_text(text):
    text = re.sub(r'\d+', '', text)  # حذف اعداد
    text = re.sub(r'[^\w\s]', '', text)  # حذف علائم
    text = re.sub(r'\s+', ' ', text).strip()  # نرمال‌سازی فاصله‌ها
    return text

data['Processed_Title'] = data['Title'].apply(preprocess_text)

# استخراج ویژگی‌های TF-IDF
vectorizer = TfidfVectorizer()
X = vectorizer.fit_transform(data['Processed_Title'])

# خوشه‌بندی با KMeans
num_clusters = 500  # تعداد خوشه‌ها، قابل تنظیم
kmeans = KMeans(n_clusters=num_clusters, random_state=42)
data['Cluster'] = kmeans.fit_predict(X)

# ایجاد شیت title
unique_titles = data.groupby('Cluster').first()[['Title']].reset_index()
unique_titles['idTit'] = unique_titles.index + 1
title_df = unique_titles[['idTit', 'Title']]

# ایجاد شیت DatabaseTemplate
data_template = data[['Row', 'Work Order NO']].copy()
data_template['Title'] = data.apply(lambda row: unique_titles[unique_titles['Cluster'] == row['Cluster']]['Title'].iloc[0], axis=1)

# ذخیره در فایل اکسل
with pd.ExcelWriter('output.xlsx') as writer:
    title_df.to_excel(writer, sheet_name='title', index=False)
    data_template.to_excel(writer, sheet_name='DatabaseTemplate', index=False)