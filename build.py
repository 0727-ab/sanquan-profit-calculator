"""把 src/ 里的界面和三个脚本合成一个可双击打开的 HTML：dist/三全利润与人工测算器.html"""
from pathlib import Path
root = Path(__file__).parent
h = (root / 'src/app.html').read_text(encoding='utf-8')
for tag, f in (('/*DEFAULTS*/', 'defaults.js'), ('/*ENGINE*/', 'engine.js'), ('/*XLSX*/', 'xlsx.js')):
    h = h.replace(tag, (root / 'src' / f).read_text(encoding='utf-8').replace('</script', '<\\/script'))
(root / 'dist').mkdir(exist_ok=True)
(root / 'dist/三全利润与人工测算器.html').write_text(h, encoding='utf-8')
print('built', len(h), 'chars')
