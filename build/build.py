import json, re, datetime, pathlib
root = pathlib.Path(__file__).resolve().parent.parent
d = json.load(open(root/'data/data.json'))
d['generatedAt'] = d.get('generatedAt') or datetime.datetime.now(datetime.timezone.utc).isoformat()
json.dump(d, open(root/'data/data.json','w'), ensure_ascii=False, indent=1)
t = (root/'build/index.template.html').read_text(encoding='utf-8')
t = t.replace('/*DATA_JSON*/null', json.dumps(d, ensure_ascii=False, separators=(',',':')))
t = t.replace('/*STATES_JSON*/null', (root/'build/states.json').read_text())
(root/'index.html').write_text(t, encoding='utf-8')
print('ok', len(t))
