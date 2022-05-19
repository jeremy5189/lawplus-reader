import os
import json
from xlsx_to_json import xlsx_to_json
from jinja2 import Template

OUTPUT = './public'
INPUT = './xlsx'

# store output dir name (filename without extension)
output_dirs = []
iframe_url = os.environ.get('IFRAME_URL').replace("\\", "")
print(f"iframe_url from env var: {iframe_url}")

if not os.path.exists(OUTPUT):
    os.makedirs(OUTPUT)

print(f"Scan input files in {INPUT}...")

dir = os.fsencode(INPUT)
for file in os.listdir(dir):
    filename = os.fsdecode(file)
    if filename.endswith(".xlsx"):
        print(os.path.join(INPUT, filename))
        output_dirs.append(filename.replace('.xlsx', ''))

# create OUTPUT/<filename> if not exist
for output_dir in output_dirs:
    if not os.path.exists(os.path.join(OUTPUT, output_dir)):
        print(f"Create output dir {output_dir}...")
        os.makedirs(os.path.join(OUTPUT, output_dir))

# convert xlsx to json
for output_dir in output_dirs:
    order = xlsx_to_json(os.path.join(INPUT, f"{output_dir}.xlsx"), os.path.join(OUTPUT, output_dir))

    print(f"Generate order.json in {output_dir}...")
    with open(os.path.join(OUTPUT, output_dir, 'order.json'), 'w', encoding='utf8') as w:
        w.write(json.dumps(order, indent=2, ensure_ascii=False))

    print(f"Generate index.html in {output_dir}...")
    with open('reader.html', 'r', encoding='utf8') as r:
        reader_template = Template(r.read())
        with open(os.path.join(OUTPUT, output_dir, 'index.html'), 'w', encoding='utf8') as w:
            w.write(reader_template.render(
                dirname=output_dir,
                iframe_url=iframe_url,))
