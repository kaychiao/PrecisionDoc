"""
Configuration settings for document formatting.
"""

# Word document font settings
WORD_FONT_SETTINGS = {
    'default_font_name': 'Microsoft YaHei',  # Default font for regular text
    'east_asia_font_name': 'Microsoft YaHei',  # Font for East Asian characters
    'default_font_size': 10,  # Default font size in points
    'heading_font_name': 'Microsoft YaHei',  # Font for headings
    'heading_font_size': 16,  # Font size for headings in points
    'subheading_font_size': 12,  # Font size for subheadings in points
    'run_font_name': 'Microsoft YaHei',  # Font for regular text runs
    'run_font_size': 10,  # Font size for regular text runs
    'page_number_font_name': 'Microsoft YaHei',  # Font for page numbers
    'separator_font_name': 'Microsoft YaHei',  # Font for separator lines
}

# Table formatting settings
TABLE_SETTINGS = {
    'show_borders_default': True,  # Default setting for showing table borders
    'multi_line_text_default': False,  # Default setting for multi-line text in cells
}
