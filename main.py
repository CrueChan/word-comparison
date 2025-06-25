import docx
from difflib import SequenceMatcher

def extract_text_from_docx(file_path):
    """
    从Word文档中提取文本内容
    
    参数:
    file_path (str): Word文档的路径
    
    返回:
    str: 文档中的所有文本内容
    """
    doc = docx.Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return '\n'.join(full_text)

def compare_documents(file_paths):
    """
    比较多个Word文档的内容
    
    参数:
    file_paths (list): Word文档路径列表
    
    返回:
    tuple: (是否完全相同, 相似度矩阵, 差异详情)
    """
    if len(file_paths) < 2:
        raise ValueError("至少需要2个文件进行比较")
    
    # 提取所有文档的文本
    texts = []
    for path in file_paths:
        try:
            text = extract_text_from_docx(path)
            texts.append(text)
        except Exception as e:
            print(f"读取文件 {path} 时发生错误: {str(e)}")
            return None
    
    # 计算文档间的相似度
    n = len(texts)
    similarity_matrix = [[0 for _ in range(n)] for _ in range(n)]
    differences = []
    
    for i in range(n):
        for j in range(i+1, n):
            similarity = SequenceMatcher(None, texts[i], texts[j]).ratio()
            similarity_matrix[i][j] = similarity
            similarity_matrix[j][i] = similarity
            
            if similarity < 1.0:
                differences.append(f"文档 {i+1} 和文档 {j+1} 的相似度为: {similarity:.2%}")
    
    # 判断是否完全相同
    all_identical = all(similarity_matrix[i][j] == 1.0 
                       for i in range(n) 
                       for j in range(i+1, n))
    
    return all_identical, similarity_matrix, differences

def main():
    """
    主函数，用于演示如何使用比较功能
    """
    # 替换为实际的文件路径
    files = [
        "培训中心物联网连接业务服务协议（无纸化办公）-250117.docx",
        "培训中心物联网连接业务服务协议（无纸化办公）-250120.docx"
    ]
    
    result = compare_documents(files)
    if result is None:
        return
    
    all_identical, similarity_matrix, differences = result
    
    print("\n比较结果:")
    if all_identical:
        print("所有文档的内容完全相同！")
    else:
        print("文档内容存在差异:")
        for diff in differences:
            print(diff)
        
        print("\n相似度矩阵:")
        for i, row in enumerate(similarity_matrix):
            print(f"文档 {i+1}:", end=" ")
            print([f"{x:.2%}" for x in row])

if __name__ == "__main__":
    main()
