import re

def set_nitty_gritty(text):
    # Strip 2 or more spaces
    text = re.sub(' {2,}', ' ',  text)
    # Put bullet point for Sub-subitem preceded by '-' or '~'.
    text = re.sub('^(-|~)', '•', text)
    # Put bullet point for Sub-subitem preceded by a single * followed by space.
    text = re.sub('^[*?]\s', ' • ', text)
    # Instead of ';' at the end of line, use hyphen instead.
    text = re.sub(';$', ':', text)
    text = set_comma_space(text)
    text = set_x(text)
    return text

# Set space after comma
def set_comma_space(text):
    x = re.compile(',\w+')
    if x.search(text):
        substring = re.findall(',\w+', text)
        for word in substring:
            text = re.sub(word, ', ' + word[1:], text)
    return text

# Function to replace description such as 1x, 20x, 10X , x1, x20, X20 into 1 x, 20 x, 10 x, x 1, x 20, X 10 etc.
def set_x(text):
    # For cases such as 20x, 30X
    x = re.compile('(\d+x|\d+X)')
    if x.search(text):
        substring = re.findall('(\d+x|\d+X)', text)
        for word in substring:
            text = re.sub(word, (word[:-1] + ' x'), text)
    # For cases such as x20, X30
    x = re.compile('(x\d+|X\d+)')
    if x.search(text):
        substring = re.findall('(x\d+|X\d+)', text)
        for word in substring:
            text = re.sub(word, ('x ' + word[1:]), text)
    # For cases such as 20 X, 30 X
    x = re.compile('(\d+ X)')
    if x.search(text):
        substring = re.findall('(\d+ X)', text)
        for word in substring:
            text = re.sub(word, (word[:-1] + 'x'), text)
    # For cases such as X 20, X 30
    x = re.compile('(X \d+)')
    if x.search(text):
        substring = re.findall('(X \d+)', text)
        for word in substring:
            text = re.sub(word, ('x' + word[1:]), text)
    return text

