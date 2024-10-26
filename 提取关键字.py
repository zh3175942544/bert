from keybert import KeyBERT
from sklearn.feature_extraction.text import CountVectorizer
import jieba
import docx

# 本地模型路径
model_path = r"C:\Users\15711\.cache\huggingface\hub\models--sentence-transformers--all-MiniLM-L6-v2\snapshots\8b3219a92973c328a8e22fadcfa821b5dc75636a"

# 初始化KeyBERT模型
kw_model = KeyBERT(model=model_path)

# 读取Word文档内容
def read_docx(file_path):
    doc = docx.Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return '\n'.join(full_text)

# Word文档路径
doc_path = r"C:\Users\15711\Desktop\PyProject\pytest\transformer实战\读取不同格式\444.docx"
text = read_docx(doc_path)

# 使用jieba进行中文分词
def chinese_tokenizer(text):
    return list(jieba.cut(text))

# 使用CountVectorizer进行分词
vectorizer = CountVectorizer(tokenizer=chinese_tokenizer, ngram_range=(1, 1))

# 提取关键词
keywords = kw_model.extract_keywords(text, vectorizer=vectorizer, keyphrase_ngram_range=(1, 1), top_n=30)

# 输出关键词
print(keywords)