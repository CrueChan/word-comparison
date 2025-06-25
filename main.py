import docx
from difflib import SequenceMatcher
import sys
import os

def extract_text_from_docx(file_path):
    """
    Extract text content from a Word document
    
    Args:
        file_path (str): Path to the Word document
    
    Returns:
        str: All text content from the document
    """
    doc = docx.Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return '\n'.join(full_text)

def compare_documents(file_paths):
    """
    Compare multiple Word documents content
    
    Args:
        file_paths (list): List of Word document paths
    
    Returns:
        tuple: (is_completely_identical, similarity_matrix, difference_details)
    """
    if len(file_paths) < 2:
        raise ValueError("At least 2 files are required for comparison")
    
    # Extract text from all documents
    texts = []
    for path in file_paths:
        try:
            text = extract_text_from_docx(path)
            texts.append(text)
        except Exception as e:
            print(f"Error reading file {path}: {str(e)}")
            return None
    
    # Calculate similarity between documents
    n = len(texts)
    similarity_matrix = [[0 for _ in range(n)] for _ in range(n)]
    differences = []
    
    for i in range(n):
        for j in range(i+1, n):
            similarity = SequenceMatcher(None, texts[i], texts[j]).ratio()
            similarity_matrix[i][j] = similarity
            similarity_matrix[j][i] = similarity
            
            if similarity < 1.0:
                differences.append(f"Similarity between Document {i+1} and Document {j+1}: {similarity:.2%}")
    
    # Check if all documents are identical
    all_identical = all(similarity_matrix[i][j] == 1.0 
                       for i in range(n) 
                       for j in range(i+1, n))
    
    return all_identical, similarity_matrix, differences

def main():
    """
    Main function to demonstrate how to use the comparison functionality
    """
    # Replace with actual file paths
    files = [
        "document1.docx",
        "document2.docx"
    ]
    
    result = compare_documents(files)
    if result is None:
        return
    
    all_identical, similarity_matrix, differences = result
    
    print("\nComparison Results:")
    if all_identical:
        print("All documents are completely identical!")
    else:
        print("Documents have differences:")
        for diff in differences:
            print(diff)
        
        print("\nSimilarity Matrix:")
        for i, row in enumerate(similarity_matrix):
            print(f"Document {i+1}:", end=" ")
            print([f"{x:.2%}" for x in row])

if __name__ == "__main__":
    main()
