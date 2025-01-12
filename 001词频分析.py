import jieba
from collections import Counter
import matplotlib.pyplot as plt
plt.switch_backend('TkAgg')  # 使用 TkAgg 后端
from wordcloud import WordCloud
from docx import Document
from tkinter import Tk, filedialog
import os
import matplotlib

# 全局设置 Matplotlib 支持中文显示
matplotlib.rcParams['font.sans-serif'] = ['SimHei']  # 使用黑体
matplotlib.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题

# 打开文件选择对话框
def open_file_dialog():
    root = Tk()
    root.withdraw()  # 隐藏主窗口
    file_path = filedialog.askopenfilename(
        title="选择 Word 文档",
        filetypes=[("Word 文件", "*.docx"), ("所有文件", "*.*")]
    )
    return file_path

# 读取 Word 文档
def read_docx(file_path):
    doc = Document(file_path)
    text = ""
    for paragraph in doc.paragraphs:
        text += paragraph.text + "\n"
    return text

# 加载停用词表
def load_stopwords(stopwords_path):
    if not os.path.exists(stopwords_path):
        print(f"停用词文件 {stopwords_path} 不存在，将使用默认停用词。")
        return set()  # 返回空集合
    with open(stopwords_path, "r", encoding="utf-8") as f:
        stopwords = set([line.strip() for line in f.readlines()])
    return stopwords

# 过滤停用词
def filter_stopwords(words, stopwords):
    return [word for word in words if word not in stopwords]

# 主程序
def main():
    # 打开文件选择对话框
    file_path = open_file_dialog()
    if not file_path:
        print("未选择文件，程序退出。")
        return

    # 读取 Word 文档
    text = read_docx(file_path)
    if not text:
        print("文档内容为空，程序退出。")
        return

    # 使用 jieba 进行分词
    words = jieba.lcut(text)
    if not words:
        print("分词结果为空，程序退出。")
        return

    # 加载停用词表
    stopwords_path = "stopwords.txt"  # 停用词文件路径
    stopwords = load_stopwords(stopwords_path)
    if not stopwords:
        print("未加载停用词，将使用默认停用词。")
        stopwords = set(["的", "是", "在", "了", "和", "有", "我", "你", "他", "这", "，", "。", "、", "\n", " "])  # 默认停用词

    # 过滤停用词
    filtered_words = filter_stopwords(words, stopwords)

    # 统计词频
    word_counts = Counter(filtered_words)

    # 打印最常见的 15 个词
    print("最常见的 15 个词：")
    print(word_counts.most_common(15))

    # 可视化词频
    # 使用 matplotlib 绘制柱状图
    top_words = word_counts.most_common(15)
    labels, values = zip(*top_words)
    plt.bar(labels, values)
    plt.xticks(rotation=45)
    plt.xlabel('词语')
    plt.ylabel('频率')
    plt.title('最常见的 15 个词')
    plt.show()

    # 使用 wordcloud 生成词云
    font_path = "C:/Windows/Fonts/simhei.ttf"  # 字体文件路径（黑体）
    if not os.path.exists(font_path):
        print(f"字体文件 {font_path} 不存在，请确保路径正确。")
        return

    try:
        wordcloud = WordCloud(font_path=font_path, width=800, height=400, background_color='white').generate_from_frequencies(word_counts)
        plt.figure(figsize=(10, 5))
        plt.imshow(wordcloud, interpolation='bilinear')
        plt.axis('off')
        plt.title('词云图')
        plt.show()
    except Exception as e:
        print(f"生成词云时出错：{e}")

# 运行主程序
if __name__ == "__main__":
    main()