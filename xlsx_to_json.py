import json
import hashlib
import openpyxl


def sha1(str):
    hash = hashlib.sha1()
    hash.update(str.encode('utf8'))
    return hash.hexdigest()


def xlsx_to_json(path, output):
    order = []
    sh = openpyxl.load_workbook(path).worksheets[0]

    for c in sh.values:
        if c[0] == None:
            continue
        struct = {
            "casename": c[1],
            "place": c[2],
            "judge_date": c[3],
            "issue": c[4],
            "type": c[5],
            "content": c[6],
            "related_statue": None,
        }
        try:
            struct["related_statue"] = json.loads(c[7].replace('\'', '"'))
        except:
            struct["related_statue"] = c[7]

        digest = sha1(struct['content'])
        name = f"{struct['casename']}_{struct['place']}_{struct['judge_date']}_{digest[0:6]}"
        struct['uuid'] = sha1(name + digest)
        order.append(name)

        print(f"Writing to {output}/{name}.json")
        with open(f"{output}/{name}.json", 'w', encoding='utf8') as w:
            w.write(json.dumps(struct, indent=2, ensure_ascii=False))

    return order
