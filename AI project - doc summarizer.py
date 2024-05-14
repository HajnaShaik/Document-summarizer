from docx import Document
import re
import os

def load_word_document(file_path):
    """
    Load a Word document and return the text content.
    """
    try:
        doc = Document(file_path)
        text = ""
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
        return text
    except Exception as e:
        print("Error:", e)
        
        return None

def clean_text(text):
    """
    Clean the text by removing special characters and extra spaces.
    """
    text = re.sub(r'\s+', ' ', text)  # Replace multiple spaces with single space
    text = re.sub(r'[^\w\s]', '', text)  # Remove special characters
    return text.lower()  # Convert text to lowercase

def summarize_text(text, ratio=0.6):
    """
    Summarize the text content based on word frequency.
    """
    cleaned_text = clean_text(text)
    words = cleaned_text.split()
    # Calculate the number of words to include in the summary
    summary_length = int(len(words) * ratio)
    # Reconstruct the summary
    summary = ' '.join(words[:summary_length])
    return summary

def main():
    # Ask for the Word document path
    file_path = input("Enter the path to the Word document: ")

    # Check if the file exists
    if not os.path.exists(file_path):
        print("Error: File not found.")
        return

    # Load the Word document
    text = load_word_document(file_path)
    if text is None:
        return

    # Ask for the summary ratio
    ratio = float(input("Enter the summary ratio (e.g., 0.2 for 20% summary length): "))

    # Summarize the text
    summary = summarize_text(text, ratio)
    
    # Print the summary
    print("\nSummary:")
    print(summary)

if __name__ == "__main__":
    main()
