
import warnings
warnings.filterwarnings("ignore")
from transformers import pipeline, AutoTokenizer, AutoModelForSequenceClassification

# 指定本地模型的路径
local_model_path = "C:\\Users\\15711\\.cache\\huggingface\\hub\\models--ahmedrachidFinancialBERT-Sentiment-Analysis"

# 加载分词器
tokenizer = AutoTokenizer.from_pretrained(local_model_path)

# 加载模型
model = AutoModelForSequenceClassification.from_pretrained(local_model_path)

# 创建情感分析 pipeline，并指定模型和分词器
classifier = pipeline("sentiment-analysis", model=model, tokenizer=tokenizer)

# 示例句子
sentences = [
    "Operating profit rose to EUR 13.1 mn from EUR 8.7 mn in the corresponding period in 2007 representing 7.7 % of net sales.",
    "Bids or offers include at least 1,000 shares and the value of the shares must correspond to at least EUR 4,000.",
    "Raute reported a loss per share of EUR 0.86 for the first half of 2009 , against EPS of EUR 0.74 in the corresponding period of 2008.",
]

# 进行情感分析
results = classifier(sentences)

# 打印结果
print(results)