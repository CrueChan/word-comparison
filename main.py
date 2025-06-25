import docx
from difflib import SequenceMatcher, ndiff
import re

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
        if para.text.strip():  # Only include non-empty paragraphs
            full_text.append(para.text.strip())
    return '\n'.join(full_text)

def get_detailed_differences(text1, text2):
    """
    Get detailed differences between two texts
    
    Args:
        text1 (str): First text
        text2 (str): Second text
    
    Returns:
        list: List of differences with context
    """
    # Split texts into lines
    lines1 = text1.split('\n')
    lines2 = text2.split('\n')
    
    # Get differences using ndiff
    diff_list = list(ndiff(lines1, lines2))
    
    # Process and format differences
    detailed_differences = []
    current_diff = []
    
    for line in diff_list:
        if line.startswith('- ') or line.startswith('+ ') or line.startswith('? '):
            current_diff.append(line)
        else:
            if current_diff:
                # Process the collected difference
                if len(current_diff) > 1:  # Only include actual differences
                    detailed_differences.append('\n'.join(current_diff))
                current_diff = []
    
    # Add any remaining differences
    if current_diff:
        detailed_differences.append('\n'.join(current_diff))
    
    return detailed_differences

def compare_documents(file_paths):
    """
    Compare multiple Word documents content
    
    Args:
        file_paths (list): List of Word document paths
    
    Returns:
        tuple: (is_completely_identical, similarity_matrix, difference_details, detailed_diffs)
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
    detailed_diffs = []
    
    for i in range(n):
        for j in range(i+1, n):
            # Calculate similarity
            similarity = SequenceMatcher(None, texts[i], texts[j]).ratio()
            similarity_matrix[i][j] = similarity
            similarity_matrix[j][i] = similarity
            
            if similarity < 1.0:
                differences.append(f"Similarity between Document {i+1} and Document {j+1}: {similarity:.2%}")
                
                # Get detailed differences
                detailed_diff = get_detailed_differences(texts[i], texts[j])
                if detailed_diff:
                    diff_entry = {
                        'doc1': f"Document {i+1}",
                        'doc2': f"Document {j+1}",
                        'differences': detailed_diff
                    }
                    detailed_diffs.append(diff_entry)
    
    # Check if all documents are identical
    all_identical = all(similarity_matrix[i][j] == 1.0 
                       for i in range(n) 
                       for j in range(i+1, n))
    
    return all_identical, similarity_matrix, differences, detailed_diffs

def main():
    """
    Main function to demonstrate how to use the comparison functionality
    """
    # Replace with actual file paths
    files = [
        "Appendix_2_24.docx",
        "Appendix_2_25.docx"
    ]
    
    result = compare_documents(files)
    if result is None:
        return
    
    all_identical, similarity_matrix, differences, detailed_diffs = result
    
    print("\n=== Comparison Results ===")
    if all_identical:
        print("All documents are identical!")
    else:
        print("Documents have differences:")
        for diff in differences:
            print(diff)
        
        print("\nSimilarity Matrix:")
        for i, row in enumerate(similarity_matrix):
            print(f"document {i+1}:", end=" ")
            print([f"{x:.2%}" for x in row])
        
        print("\nDetailed differences:")
        for diff in detailed_diffs:
            print(f"\nDifferences between {diff['doc1']} and {diff['doc2']} :")
            for difference in diff['differences']:
                print("\n" + difference)
                print("-" * 50)

if __name__ == "__main__":
    main()
