import jieba
from collections import Counter

# 示例短文
text = """
自然语言处理（NLP）是人工智能的一个重要领域，专注于计算机与人类之间的交互。
它涉及文本分析、语音识别、机器翻译等多种技术，广泛应用于智能助手、搜索引擎和推荐系统中。
NLP的目标是让计算机能够理解、生成和处理人类语言。
"""

# 分词
words = jieba.lcut(text)

# 去除停用词（示例停用词列表）
stop_words = set(["是", "的", "一个", "之间", "与", "它", "和", "多种", "能够"])
filtered_words = [word for word in words if word not in stop_words and len(word) > 1]  # 去除单字词

# 统计词频
word_freq = Counter(filtered_words)

# 提取前5个高频词作为关键词
keywords = [word for word, freq in word_freq.most_common(5)]

# 生成总结
summary = "本文介绍了自然语言处理（NLP），它是人工智能的重要领域，涉及文本分析、语音识别等技术，广泛应用于智能助手和推荐系统。"
print("总结:", summary)
print("关键词:", ", ".join(keywords))