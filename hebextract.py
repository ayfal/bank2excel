'''
this module extracts hebrew text from a pdf file in the correct order. 
'''
import pymupdf, re

'''
About rotation:
the "line" dictionary has a "dir" tuple that holds the (cos, sin) of the rotation angle of the text line.
when dir is (1,0) (0 deg.) we sort the lines in ascending vertical direction (y) and then the characters by descending horizontal direction (-x).
when dir is (0,1) (90 deg.) we sort the lines in descending horizontal direction (-x) and then the characters by descending vertical direction (-y).
when dir is (-1,0) (180 deg.) we sort the lines in descending vertical direction (-y) and then the characters by ascending horizontal direction (x).
when dir is (0,-1) (270 deg.) we sort the lines in ascending horizontal direction (x) and then the characters by ascending vertical direction (y).
so this means that the sorting of the lines is cos*y - sin*x, and the sorting of the characters is -cos*x - sin*y.
'''

'''
We round the values to the nearest 15 pixels because sometimes characters' height is not exactly the same, and we want to group them together.
'''
def lines_direction(cos, sin, x, y):
    return round(round(cos*y - sin*x)/15)*15

def chars_direction(cos, sin, x, y):
    return -cos*x - sin*y

'''
Sort the blocks of a TextPage in hebrew reading order.
'''
def SortBlocks(blocks):
    sorted_blocks = []
    for b in blocks:
        cos = b["lines"][0]["dir"][0]
        sin = b["lines"][0]["dir"][1]
        x = b["bbox"][0]
        y = b["bbox"][1]
        sorted_blocks.append([lines_direction(cos, sin, x, y), chars_direction(cos, sin, x, y), b])
    sorted_blocks.sort(key= lambda x: (x[0], x[1]))
    return [b[2] for b in sorted_blocks]


''' 
Sort the lines of a block in hebrew reading order.
'''
def SortLines(lines):
    sorted_lines = []
    for l in lines:
        sorted_lines.append([lines_direction(l["dir"][0], l["dir"][1], l["bbox"][0], l["bbox"][1]), l])
    sorted_lines.sort(key=lambda x: x[0])
    return [l[1] for l in sorted_lines]

#todo: the next two functions are the same, remove one of them.
''' 
Sort the spans of a line in hebrew reading order.
'''
def SortSpans(spans, dir):
    sorted_spans = []
    for s in spans:
        sorted_spans.append([chars_direction(dir[0], dir[1], s["bbox"][0], s["bbox"][1]), s])
    sorted_spans.sort(key=lambda x: x[0])
    return [s[1] for s in sorted_spans]


''' 
Sort the characters of a span in hebew reading order.
'''
def SortChars(chars,dir):
    sorted_chars = []
    for c in chars:
        sorted_chars.append([chars_direction(dir[0], dir[1], c["bbox"][0], c["bbox"][1]), c])
    sorted_chars.sort(key=lambda x: x[0])
    return [c[1] for c in sorted_chars]

'''
sort the characters of a page in hebrew reading order.
'''
def SortCharsPage(chars, dir):
    sorted_chars = []
    for c in chars:
        cos=dir[0]
        sin=dir[1]
        x = c["bbox"][0]
        y = c["bbox"][1]
        sorted_chars.append([lines_direction(cos, sin, x, y), chars_direction(cos, sin, x, y), c])
    sorted_chars.sort(key=lambda x: (x[0], x[1]))
    return [c[2] for c in sorted_chars]


def get_hebrew_text(fname):
    with pymupdf.open(fname) as doc:  # open document
        pages = [page.get_textpage().extractRAWDICT() for page in doc]
        words = [page.get_textpage().extractWORDS() for page in doc]
    full_text = ''
    #sort the document:
    for page in pages:
        page_text=''
        blocks = SortBlocks(page["blocks"])
        for b in blocks:
            lines = SortLines(b["lines"])
            for l in lines:
                spans = SortSpans(l["spans"], l["dir"])
                for s in spans:
                    chars = SortChars(s["chars"], l["dir"])
                    for c in chars:
                        page_text += c["c"]
                        if c['c'] == ':':
                            page_text += ' ' #because sometimes the colon is attached to the next word. this is a quick fix. todo: find a better solution.
                    #ensure that spans are separated by a space:
                    if page_text[-1] != " ":
                        page_text += " "
                page_text += "\n"
        full_text += page_text
    #find english words and numbers, and reverse them
    full_text = re.sub(r'[a-zA-Z\d\.,\-/]+', lambda match: match.group()[::-1], full_text)
    return full_text, words

'''
this function sorts by characters only, and not by lines or blocks.
'''
def get_hebrew_text_chars(fname):
    with pymupdf.open(fname) as doc:  # open document
        pages = [page.get_textpage().extractRAWDICT() for page in doc]
        words = [page.get_textpage().extractWORDS() for page in doc]
    full_text = ''
    #sort the document:
    for page in pages:
        chars = []
        for b in page["blocks"]:
            for l in b["lines"]:
                for s in l["spans"]:
                    chars.extend(s["chars"])
        chars = SortCharsPage(chars, (1,0)) #for now, we don't support rotation in this function.
        #append the chars to the full text:
        for i in range(len(chars)-1):
            full_text += chars[i]["c"]
            #add a space where needed:
            if chars[i]["bbox"][0] - chars[i+1]["bbox"][2] > 2:
                full_text += " "
            #add a newline where needed:
            if chars[i]["bbox"][3] < chars[i+1]["bbox"][1]:
                full_text += "\n"
        full_text += chars[-1]["c"] + "\n"
#find english words and numbers, and reverse them
    full_text = re.sub(r'[a-zA-Z\d\.,\-/]+', lambda match: match.group()[::-1], full_text)
    return full_text, words