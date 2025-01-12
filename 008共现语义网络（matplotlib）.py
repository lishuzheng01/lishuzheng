import jieba
from collections import defaultdict
import networkx as nx
import matplotlib.pyplot as plt
from docx import Document
from tkinter import Tk, filedialog

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

# 4. 构建共现矩阵
def build_co_occurrence_matrix(words, window_size=2):
    co_occurrence = defaultdict(int)
    for i in range(len(words)):
        for j in range(i + 1, min(i + window_size + 1, len(words))):
            word1, word2 = words[i], words[j]
            if word1 != word2:
                co_occurrence[(word1, word2)] += 1
    print("共现关系:", co_occurrence)
    return co_occurrence

# 5. 构建共现语义网络
def build_co_occurrence_network(co_occurrence):
    G = nx.Graph()
    for (word1, word2), weight in co_occurrence.items():
        G.add_edge(word1, word2, weight=weight)
    return G

# 6. 可视化网络
def visualize_network(G):
    plt.figure(figsize=(10, 8))
    pos = nx.spring_layout(G)  # 布局算法
    nx.draw_networkx_nodes(G, pos, node_size=2000, node_color='lightblue')
    nx.draw_networkx_edges(G, pos, width=1.5, alpha=0.6, edge_color='gray')
    nx.draw_networkx_labels(G, pos, font_size=12, font_family='SimHei')  # 设置中文字体
    plt.title("中文共现语义网络", fontsize=15)
    plt.axis('off')
    plt.show()

# 主函数
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

    # 构建共现矩阵
    co_occurrence = build_co_occurrence_matrix(words)

    # 构建共现语义网络
    G = build_co_occurrence_network(co_occurrence)

    # 可视化网络
    visualize_network(G)

if __name__ == "__main__":
    main()