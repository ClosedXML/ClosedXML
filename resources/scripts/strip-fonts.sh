# Strip the largest tables: ttf instructions and various positionsins

strip_font() {
  pyftsubset Carlito-${1}.ttf '*' --hinting-tables+=GPOS --no-hinting
  python3 replace-glyf.py Carlito-${1}.subset.ttf Carlito-${1}.subset.glyf.ttf
  rm Carlito-${1}.subset.ttf

  python3 rename-font.py Carlito-${1}.subset.glyf.ttf CarlitoBare-${1}.ttf
  rm Carlito-${1}.subset.glyf.ttf
}

strip_font Bold
strip_font Italic
strip_font BoldItalic
strip_font Regular