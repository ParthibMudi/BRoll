import sys, os, json
sys.path.insert(0, '.')
from app import render_broll_item_to_png

plan = json.load(open('output/broll_plan_5127cde4.json'))
for item in plan['broll_items']:
    png = f"output/test_{item['id']}.png"
    render_broll_item_to_png(item, png)
    size = os.path.getsize(png)
    print(f"{item['id']} ({item['template_type']}) -> {png}  [{size} bytes]")
print('All PNGs rendered OK')
