import jieba
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.decomposition import LatentDirichletAllocation
from tkinter import Tk, filedialog
from docx import Document

# 1. 打开文件选择窗口，选择 Word 文档
def select_word_file():
    root = Tk()
    root.withdraw()  # 隐藏主窗口
    file_path = filedialog.askopenfilename(
        title="选择 Word 文档",
        filetypes=[("Word 文件", "*.docx")]
    )
    return file_path

# 2. 读取 Word 文档内容
def read_word_file(file_path):
    doc = Document(file_path)
    text = "\n".join([paragraph.text for paragraph in doc.paragraphs if paragraph.text.strip()])
    return text

# 3. 中文分词
def chinese_tokenizer(text):
    return " ".join(jieba.lcut(text))  # 使用 jieba 分词并用空格连接

# 4. 主题分析
def thematic_analysis(documents):
    # 分词
    tokenized_docs = [chinese_tokenizer(doc) for doc in documents]

    # TF-IDF 向量化
    tfidf_vectorizer = TfidfVectorizer()
    tfidf_matrix = tfidf_vectorizer.fit_transform(tokenized_docs)

    # LDA 主题建模
    n_topics = 2  # 假设提取 2 个主题
    lda = LatentDirichletAllocation(n_components=n_topics, random_state=42)
    lda.fit(tfidf_matrix)

    # 输出每个主题的关键词
    def print_top_words(model, feature_names, n_top_words):
        for topic_idx, topic in enumerate(model.components_):
            print(f"主题 {topic_idx + 1}:")
            print(" ".join([feature_names[i] for i in topic.argsort()[:-n_top_words - 1:-1]]))

    n_top_words = 5  # 每个主题显示前 5 个关键词
    print_top_words(lda, tfidf_vectorizer.get_feature_names_out(), n_top_words)

# 主程序
if __name__ == "__main__":
    # 选择 Word 文档
    file_path = select_word_file()
    if not file_path:
        print("未选择文件，程序退出。")
        exit()

    # 读取文档内容
    text = read_word_file(file_path)
    print("读取的文档内容：")
    print(text)

    # 将文档内容按段落分割为列表
    documents = [paragraph for paragraph in text.split("\n") if paragraph.strip()]

    # 进行主题分析
    print("\n主题分析结果：")
    thematic_analysis(documents)