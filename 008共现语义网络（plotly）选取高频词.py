import jieba
from collections import defaultdict, Counter
import networkx as nx
from docx import Document
from tkinter import Tk, filedialog
import plotly.graph_objects as go
import os
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

# 5. 构建共现矩阵
def build_co_occurrence_matrix(words, window_size=2):
    co_occurrence = defaultdict(int)
    for i in range(len(words)):
        for j in range(i + 1, min(i + window_size + 1, len(words))):
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

# 7. 使用 Plotly 可视化网络
def visualize_network_with_plotly(G):
    # 获取节点位置
    pos = nx.spring_layout(G)

    # 创建边的轨迹
    edge_trace = []
    for edge in G.edges:
        x0, y0 = pos[edge[0]]
        x1, y1 = pos[edge[1]]
        edge_trace.append(go.Scatter(
            x=[x0, x1, None], y=[y0, y1, None],
            line=dict(width=G[edge[0]][edge[1]]['weight'] * 1.5, color='#888'),
            hoverinfo='none',
            mode='lines'
        ))

    # 创建节点的轨迹
    node_trace = go.Scatter(
        x=[], y=[], text=[], mode='markers+text', textposition="top center",
        hoverinfo='text', marker=dict(
            color=[G.degree(node) for node in G.nodes],  # 节点颜色基于度数
            size=[G.degree(node) * 10 for node in G.nodes],  # 节点大小基于度数
            colorscale='YlGnBu',  # 颜色方案
            line_width=2
        )
    )

    # 添加节点位置和标签
    for node in G.nodes:
        x, y = pos[node]
        node_trace['x'] += (x,)
        node_trace['y'] += (y,)
        node_trace['text'] += (node,)

    # 创建图
    fig = go.Figure(data=edge_trace + [node_trace],
                    layout=go.Layout(
                        title='中文共现语义网络',
                        titlefont_size=16,
                        showlegend=False,
                        hovermode='closest',
                        margin=dict(b=20, l=5, r=5, t=40),
                        xaxis=dict(showgrid=False, zeroline=False, showticklabels=False),
                        yaxis=dict(showgrid=False, zeroline=False, showticklabels=False)
                    ))

    # 显示图
    fig.show()

# 加载停用词表
def load_stopwords(stopwords_path):
    if not os.path.exists(stopwords_path):
        print(f"停用词文件 {stopwords_path} 不存在，将使用默认停用词。")
        return set()  # 返回空集合
    with open(stopwords_path, "r", encoding="utf-8") as f:
        stopwords = set([line.strip() for line in f.readlines()])
    return stopwords

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

    # 加载停用词表
    stopwords_path = "stopwords.txt"  # 停用词文件路径
    stopwords = load_stopwords(stopwords_path)
    if not stopwords:
        print("未加载停用词，将使用默认停用词。")
        stopwords = set(["的", "是", "在", "了", "和", "有", "我", "你", "他", "这", "，", "。", "、", "\n", " ","："])  # 默认停用词
    
    # 选取前20个高频词
    top_n_words = filter_top_n_words(words, stopwords, n=20)

    # 构建共现矩阵
    co_occurrence = build_co_occurrence_matrix(top_n_words)

    # 构建共现语义网络
    G = build_co_occurrence_network(co_occurrence)

    # 使用 Plotly 可视化网络
    visualize_network_with_plotly(G)

if __name__ == "__main__":
    main()