import jieba
from collections import defaultdict, Counter
import networkx as nx
import matplotlib.pyplot as plt
from docx import Document
from tkinter import Tk, filedialog
import os
# 设置 matplotlib 全局中文字体
plt.rcParams['font.sans-serif'] = ['SimHei']  # 设置字体为 SimHei
plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题

# 1. 打开文件选择窗口，选择Word文档
def select_word_file():
    root = Tk()
    root.withdraw()  # 隐藏主窗口
    file_path = filedialog.askopenfilename(
        title="选择Word文档",
        filetypes=[("Word文件", "*.docx")]
    )
    return file_path

# 2. 从Word文档中读取文本
def read_text_from_word(file_path):
    doc = Document(file_path)
    text = ""
    for paragraph in doc.paragraphs:
        text += paragraph.text + "\n"
    return text

# 3. 分词
def tokenize_text(text):
    words = list(jieba.cut(text))
    print("分词结果:", words)
    return words

# 4. 过滤停用词并选取前20个高频词
def filter_top_n_words(words, stopwords, n=20):
    # 过滤停用词
    filtered_words = [word for word in words if word not in stopwords]
    
    # 统计词频
    word_freq = Counter(filtered_words)
    
    # 选取前n个高频词
    top_n_words = [word for word, _ in word_freq.most_common(n)]
    
    print("前20个高频词:", top_n_words)
    return top_n_words

# 5. 构建共现矩阵（仅限高频词）
def build_co_occurrence_matrix(words, top_n_words, window_size=2):
    co_occurrence = defaultdict(int)
    for i in range(len(words)):
        if words[i] not in top_n_words:
            continue  # 跳过非高频词
        for j in range(i + 1, min(i + window_size + 1, len(words))):
            if words[j] not in top_n_words:
                continue  # 跳过非高频词
            word1, word2 = words[i], words[j]
            if word1 != word2:
                co_occurrence[(word1, word2)] += 1
    print("共现关系:", co_occurrence)
    return co_occurrence

# 6. 构建共现语义网络
def build_co_occurrence_network(co_occurrence):
    G = nx.Graph()
    for (word1, word2), weight in co_occurrence.items():
        G.add_edge(word1, word2, weight=weight)
    return G

# 7. 可视化网络
def visualize_network(G):
    plt.figure(figsize=(10, 8))
    pos = nx.spring_layout(G)  # 布局算法
    nx.draw_networkx_nodes(G, pos, node_size=2000, node_color='lightblue')
    nx.draw_networkx_edges(G, pos, width=1.5, alpha=0.6, edge_color='gray')
    nx.draw_networkx_labels(G, pos, font_size=12)  # 不再需要单独设置字体
    plt.title("中文共现语义网络（前20个高频词）", fontsize=15)
    plt.axis('off')
    plt.show()



# 加载停用词表
def load_stopwords(stopwords_path):
    if not os.path.exists(stopwords_path):
        print(f"停用词文件 {stopwords_path} 不存在，将使用默认停用词。")
        return set()  # 返回空集合
    with open(stopwords_path, "r", encoding="utf-8") as f:
        stopwords = set([line.strip() for line in f.readlines()])
    return stopwords


# 主程序
def main():
    # 选择Word文档
    file_path = select_word_file()
    if not file_path:
        print("未选择文件，程序退出。")
        return

    # 读取文本
    text = read_text_from_word(file_path)
    print("读取的文本内容:", text)

    # 分词
    words = tokenize_text(text)
    
    # 加载停用词表
    stopwords_path = "stopwords.txt"  # 停用词文件路径
    stopwords = load_stopwords(stopwords_path)
    if not stopwords:
        print("未加载停用词，将使用默认停用词。")
        stopwords = set(["的", "是", "在", "了", "和", "有", "我", "你", "他", "这", "，", "。", "、", "\n", " ","："])  # 默认停用词
    
    # 选取前20个高频词

    top_n_words = filter_top_n_words(words, stopwords, n=20)

    # 构建共现矩阵（仅限高频词）
    co_occurrence = build_co_occurrence_matrix(words, top_n_words)

    # 构建共现语义网络
    G = build_co_occurrence_network(co_occurrence)

    # 可视化网络
    visualize_network(G)

if __name__ == "__main__":
    main()