import re

def clean_label(text: str):
    text = text.replace('<', 'LessThan')
    text = text.replace('>', 'GreaterThan')
    text = text.replace('+', 'Plus')
    text = text.replace('-', 'To')
    return re.sub(r'[^\w]', '', text)