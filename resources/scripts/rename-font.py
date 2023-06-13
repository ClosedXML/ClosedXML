from fontTools import ttLib
import sys

font_path = sys.argv[1]
font_output = sys.argv[2]

font = ttLib.TTFont(font_path)

table = font["name"];

for rec in table.names:
    # Record 0 point to the website
    if rec.nameID > 0 and 'Carlito' in rec.toUnicode():
        table.setName(rec.toUnicode().replace('Carlito', 'CarlitoBare'), rec.nameID, rec.platformID, rec.platEncID, rec.langID)

font.save(font_output)