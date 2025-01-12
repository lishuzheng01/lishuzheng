
import tkinter as tk
from tkinter import filedialog
from docx import Document
from sumy.parsers.plaintext import PlaintextParser
from sumy.nlp.tokenizers import Tokenizer
from sumy.summarizers.lsa import LsaSummarizer
from sumy.summarizers.text_rank import TextRankSummarizer

# 1. 创建文件选择窗口
def select_word_file():
    """
    打开文件选择窗口，选择 Word 文档。
    
    :return: 选择的文件路径
    """
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    file_path = filedialog.askopenfilename(
        title="选择 Word 文档",
        filetypes=[("Word 文件", "*.docx"), ("所有文件", "*.*")]
    )
    return file_path

# 2. 读取 Word 文档内容
def read_word_file(file_path):
    """
    读取 Word 文档内容。
    
    :param file_path: Word 文档路径
    :return: 文档内容（字符串）
    """
    doc = Document(file_path)
    content = []
    for paragraph in doc.paragraphs:
        content.append(paragraph.text)
    return "\n".join(content)

# 3. 使用 sumy 生成文本摘要
def summarize_with_sumy(text, language="chinese", sentences_count=2, algorithm="lsa"):
    """
    使用 sumy 库生成文本摘要。
    
    :param text: 输入的文本
    :param language: 文本语言（默认为中文）
    :param sentences_count: 摘要的句子数量
    :param algorithm: 摘要算法（lsa 或 textrank）
    :return: 生成的摘要
    """
    # 创建解析器
    parser = PlaintextParser.from_string(text, Tokenizer(language))
    
    # 选择摘要算法
    if algorithm == "lsa":
        summarizer = LsaSummarizer()
    elif algorithm == "textrank":
        summarizer = TextRankSummarizer()
    else:
        raise ValueError("不支持的摘要算法。请选择 'lsa' 或 'textrank'。")
    
    # 生成摘要
    summary = summarizer(parser.document, sentences_count)
    return " ".join([str(sentence) for sentence in summary])

# 4. 主程序
def main():
    # 选择 Word 文档
    file_path = select_word_file()
    if not file_path:
        print("未选择文件。")
        return

    # 读取文档内容
    try:
        content = read_word_file(file_path)
        print("文档内容:\n", content)

        # 使用 sumy 生成摘要
        summary = summarize_with_sumy(content, language="chinese", sentences_count=5, algorithm="lsa")
        print("\n摘要:\n", summary)
        # algorithm 参数可选 "lsa" 或 "textrank"，sentences_count 参数可选摘要句子数量
    except Exception as e:
        print("读取文档或生成摘要失败:", e)

# 运行主程序
if __name__ == "__main__":
    main()