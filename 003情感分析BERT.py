from docx import Document
from transformers import BertTokenizer, BertForSequenceClassification
import torch
from tkinter import Tk, filedialog  # 用于文件选择对话框

# 加载预训练的中文BERT模型和分词器
tokenizer = BertTokenizer.from_pretrained("bert-base-chinese")
model = BertForSequenceClassification.from_pretrained("bert-base-chinese", num_labels=2)

# 情感分析函数
def bert_sentiment_analysis(text):
    inputs = tokenizer(text, return_tensors="pt", truncation=True, padding=True)
    outputs = model(**inputs)
    probs = torch.softmax(outputs.logits, dim=-1)
    return probs[0][1].item()  # 返回正面情感的概率

# 打开文件选择对话框
def open_file_dialog():
    root = Tk()
    root.withdraw()  # 隐藏Tkinter主窗口
    file_path = filedialog.askopenfilename(
        title="选择Word文档",
        filetypes=[("Word文件", "*.docx")]
    )
    return file_path

# 读取 Word 文档内容
def read_docx(file_path):
    try:
        doc = Document(file_path)
        text = "\n".join([paragraph.text for paragraph in doc.paragraphs])
        return text
    except Exception as e:
        print(f"读取文件时出错: {e}")
        return None

# 主函数
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

    # 进行情感分析
    sentiment_score = bert_sentiment_analysis(text)
    print("BERT情感得分:", sentiment_score)

    # 根据情感得分输出分类结果
    if sentiment_score > 0.6:
        print("情感分类: 正面情感")
    elif sentiment_score < 0.4:
        print("情感分类: 负面情感")
    else:
        print("情感分类: 中性情感")

# 运行主函数
if __name__ == "__main__":
    main()