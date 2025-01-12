from snownlp import SnowNLP
import jieba
from tkinter import Tk, filedialog  # 用于文件选择对话框
from docx import Document  # 用于读取Word文档

# 1. 文本预处理（分词）
def preprocess_text(text):
    # 使用Jieba进行分词
    words = jieba.lcut(text)
    return " ".join(words)  # 返回分词后的文本

# 2. 情感分析
def analyze_sentiment(text):
    # 使用SnowNLP进行情感分析
    s = SnowNLP(text)
    sentiment_score = s.sentiments  # 情感得分（0-1，越接近1表示越正面）
    return sentiment_score

# 3. 情感分类
def classify_sentiment(score):
    if score > 0.6:
        return "正面情感"
    elif score < 0.4:
        return "负面情感"
    else:
        return "中性情感"

# 4. 打开文件选择对话框
def open_file_dialog():
    root = Tk()
    root.withdraw()  # 隐藏Tkinter主窗口
    file_path = filedialog.askopenfilename(
        title="选择Word文档",
        filetypes=[("Word文件", "*.docx")]
    )
    return file_path

# 5. 读取Word文档内容
def read_docx(file_path):
    try:
        doc = Document(file_path)
        text = "\n".join([paragraph.text for paragraph in doc.paragraphs])
        return text
    except Exception as e:
        print(f"读取文件时出错: {e}")
        return None

# 6. 主函数
def main():
    # 打开文件选择对话框
    file_path = open_file_dialog()
    if not file_path:
        print("未选择文件，程序退出。")
        return

    # 读取Word文档内容
    text = read_docx(file_path)
    if not text:
        print("文档内容为空，程序退出。")
        return

    # 文本预处理
    processed_text = preprocess_text(text)
    print("分词结果:", processed_text)

    # 情感分析
    sentiment_score = analyze_sentiment(text)
    print("情感得分:", sentiment_score)

    # 情感分类
    sentiment_category = classify_sentiment(sentiment_score)
    print("情感分类:", sentiment_category)

# 运行主函数
if __name__ == "__main__":
    main()