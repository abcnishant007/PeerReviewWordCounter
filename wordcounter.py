from docx import Document
from collections import defaultdict
import webcolors

def closest_color(requested_color):
    """
    Determines the closest CSS3 named color for a given RGB value.

    Parameters:
    - requested_color (tuple): The RGB value for which to find the closest named color.

    Returns:
    - str: The closest named color.
    """
    min_colors = {}
    for key, name in webcolors.CSS3_HEX_TO_NAMES.items():
        r_c, g_c, b_c = webcolors.hex_to_rgb(key)
        rd = (r_c - requested_color[0]) ** 2
        gd = (g_c - requested_color[1]) ** 2
        bd = (b_c - requested_color[2]) ** 2
        min_colors[(rd + gd + bd)] = name
    return min_colors[min(min_colors.keys())]

def get_color_name(rgb_value):
    """
    Fetches the name of the color for a given RGB value. If an exact name isn't found, 
    computes the closest named color.

    Parameters:
    - rgb_value (tuple): The RGB value for which to find the color name.

    Returns:
    - str: The color name (or the closest approximation).
    """
    try:
        return webcolors.rgb_to_name(rgb_value)
    except ValueError:
        return closest_color(rgb_value)

def extract_color_and_word_count(doc_path):
    """
    Reads a .docx file, extracts text based on its color, and counts the words for each color.

    Parameters:
    - doc_path (str): Path to the .docx file.

    Returns:
    - dict: A dictionary with color names as keys and word counts as values.
    """
    doc = Document(doc_path)
    color_word_count = defaultdict(int)
    
    # Loop through each paragraph in the document
    for paragraph in doc.paragraphs:
        # Loop through each 'run' (a segment of text with same style) in the paragraph
        for run in paragraph.runs:
            # Check if the run has a specific RGB color value
            if run.font.color.rgb:
                color = run.font.color.rgb
                # Split the text of the run into words and count them
                words = run.text.split()
                color_word_count[color] += len(words)

    # Convert the RGB values to color names for better readability
    color_word_count_str = {get_color_name(color): count for color, count in color_word_count.items()}
    return color_word_count_str

# Assuming the file name is 'clipboard_content.docx'. 
# You can change this to the name/path of your actual file.
print(extract_color_and_word_count('clipboard_content.docx'))
