# PeerReviewWordCounter README
_Automate word count for paper rebuttals or multiple authors by distinguishing assuming each author uses a distinct color-coded text in a Word document._

This Python snippet is designed to read a Microsoft Word document and count the number of words corresponding to each color used in the document. The output provides the word count for every distinct color in the document, with colors identified by their closest named color (e.g., "Red", "Blue").

## Prerequisites

1. Ensure you have Python installed. The code was written for Python 3.
2. Install the necessary Python packages using the following command:
   ```
   pip install -r requirements.txt
   ```

## Steps to Use:

1. **Prepare Your Document**:
   - Create a Microsoft Word document. 
   - Ensure its filename matches the one specified in the Python code (`'clipboard_content.docx'` by default). You can change the filename in the code if necessary.
   - Paste your content into this Word document.
   - Save the document.

2. **Run the Script**:
   - Navigate to the directory containing the Python script and the Word document.
   - Execute the script using:
     ```
     python word_counter.py
     ```
     Replace `<script_name>` with the name of the Python file containing the code (e.g., `color_counter.py`).

3. **Review the Results**:
   - Once the script completes, it will print the word count for each color used in the document. The output will associate word counts with the closest named color (e.g., `'Red': 150`).

## Notes:

- The color names are derived from the closest matching named CSS3 color
