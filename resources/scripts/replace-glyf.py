from fontTools import ttLib
import sys
import copy

font_path = sys.argv[1]
font_output = sys.argv[2]
font = ttLib.TTFont(font_path)
original_font = copy.deepcopy(font)
for key in original_font['glyf'].keys():
    font['glyf'][key] = original_font['glyf']['space']

font.save(font_output)