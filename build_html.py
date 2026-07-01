import sys, os, json
from datetime import datetime, timezone, timedelta
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8")

from merge_warehouses import extract_all_items
from parse_specs import build_specs_map, SPEC_LABELS

DIR = os.path.dirname(os.path.abspath(__file__))

def load_from_xls():
    print("Загрузка данных из XLS-файлов...")
    minsk_path = os.path.join(DIR, "Минск.xls")
    moscow_path = os.path.join(DIR, "Москва.xls")
    minsk = extract_all_items(minsk_path) if os.path.exists(minsk_path) else {}
    moscow = extract_all_items(moscow_path) if os.path.exists(moscow_path) else {}
    print(f"  Минск: {len(minsk)} позиций")
    print(f"  Москва: {len(moscow)} позиций")
    return minsk, moscow, "XLS"


# Источник данных — только XLS-файлы (Минск.xls / Москва.xls)
stock1, stock2, source = load_from_xls()
col1_name = "Склад Минск"
col2_name = "Склад Москва"

all_names = sorted(set(list(stock1.keys()) + list(stock2.keys())))

data = []
for name in all_names:
    m = stock1.get(name, 0)
    k = stock2.get(name, 0)
    m = int(m) if m == int(m) else m
    k = int(k) if k == int(k) else k
    data.append([name, m, k])

BRANDS = ["Dahua", "Tiandy", "ZKTeco", "Temid", "Belpark", "Hikvision", "HiWatch", "Optimus", "EZVIZ"]
brand_counts = Counter()
for name, _, _ in data:
    nl = name.lower()
    for b in BRANDS:
        if nl.startswith(b.lower()):
            brand_counts[b] += 1
            break

brand_buttons = '<button class="active" data-brand="">Все</button>\n'
for b in BRANDS:
    c = brand_counts.get(b, 0)
    if c > 0:
        brand_buttons += f'<button data-brand="{b}">{b} <span class="cnt">{c}</span></button>\n'

data_json = json.dumps(data, ensure_ascii=False, separators=(',', ':'))
now = datetime.now(timezone(timedelta(hours=3))).strftime("%d.%m.%Y %H:%M")

print("Загрузка характеристик из прайсов...")
specs_map = build_specs_map(all_names)
specs_json = json.dumps(specs_map, ensure_ascii=False, separators=(',', ':'))
spec_labels_json = json.dumps(SPEC_LABELS, ensure_ascii=False, separators=(',', ':'))

# Загрузка изображений из fetched_specs.json (если есть)
images_map = {}
fetched_path = os.path.join(DIR, "fetched_specs.json")
if os.path.exists(fetched_path):
    with open(fetched_path, encoding="utf-8") as _f:
        for entry in json.load(_f):
            img = entry.get("image_url", "")
            if img:
                images_map[entry["name"]] = img
    print(f"  Изображений загружено: {len(images_map)}")
images_json = json.dumps(images_map, ensure_ascii=False, separators=(',', ':'))

html = f'''<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Общие остатки — Электронный мир</title>
<meta http-equiv="X-Frame-Options" content="DENY">
<meta http-equiv="Content-Security-Policy" content="frame-ancestors 'none'">
<style>
body{{-webkit-user-select:none;-moz-user-select:none;-ms-user-select:none;user-select:none}}
.toolbar input[type=text],.product-name{{-webkit-user-select:text;-moz-user-select:text;-ms-user-select:text;user-select:text;cursor:text}}
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:'Segoe UI',Tahoma,sans-serif;background:#f0f2f5;color:#1a1a2e;-webkit-text-size-adjust:100%}}
.header{{background:linear-gradient(135deg,#0d47a1,#1565c0,#1976d2);color:#fff;padding:18px 30px;display:flex;align-items:center;justify-content:space-between;box-shadow:0 2px 8px rgba(0,0,0,.2);gap:10px}}
.header h1{{font-size:22px;font-weight:600;letter-spacing:.5px}}
.header .subtitle{{font-size:13px;opacity:.8;margin-top:2px}}
.header .logo{{display:flex;align-items:center;gap:10px;min-width:0}}
.header .logo img{{height:36px;flex-shrink:0}}
.toolbar{{background:#fff;padding:12px 30px;display:flex;gap:12px;align-items:center;flex-wrap:wrap;border-bottom:1px solid #ddd;box-shadow:0 1px 3px rgba(0,0,0,.05)}}
.toolbar input[type=text]{{flex:1;min-width:180px;padding:8px 14px;border:1px solid #ccc;border-radius:4px;font-size:16px;outline:none;transition:border .2s;-webkit-appearance:none}}
.toolbar input:focus{{border-color:#1976d2}}
.toolbar select,.toolbar button{{padding:8px 16px;border:1px solid #ccc;border-radius:4px;font-size:14px;cursor:pointer;background:#fff}}
.toolbar button{{background:#1976d2;color:#fff;border:none;font-weight:500}}
.toolbar button:hover{{background:#1565c0}}
.toolbar .info{{font-size:13px;color:#666;margin-left:auto;white-space:nowrap}}
.container{{padding:20px 30px;overflow-x:auto;-webkit-overflow-scrolling:touch}}
table{{width:100%;border-collapse:collapse;background:#fff;border-radius:6px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.08);min-width:480px}}
thead{{background:#1565c0;color:#fff;position:sticky;top:0;z-index:2}}
th{{padding:10px 14px;text-align:left;font-size:13px;font-weight:600;cursor:pointer;user-select:none;white-space:nowrap}}
th:hover{{background:#0d47a1}}
th.num{{text-align:right;width:100px}}
th .arrow{{margin-left:4px;font-size:10px;opacity:.6}}
th.sorted .arrow{{opacity:1}}
td{{padding:8px 14px;font-size:13px;border-bottom:1px solid #eee}}
td.num{{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}}
tr:hover{{background:#e3f2fd}}
tr.zero{{color:#aaa}}
.highlight{{background:#fff59d;padding:1px 0;border-radius:2px}}
.pagination{{display:flex;justify-content:space-between;align-items:center;padding:14px 0;font-size:13px;color:#555;flex-wrap:wrap;gap:8px}}
.pagination button{{padding:8px 16px;border:1px solid #ccc;border-radius:4px;background:#fff;cursor:pointer;font-size:14px;min-height:44px}}
.pagination button:disabled{{opacity:.4;cursor:default}}
.pagination button:not(:disabled):hover{{background:#e3f2fd;border-color:#1976d2}}
.brands{{background:#fff;padding:14px 30px;border-bottom:1px solid #ddd;display:flex;gap:8px;flex-wrap:wrap;align-items:center}}
.brands .label{{font-size:12px;color:#666;text-transform:uppercase;letter-spacing:.5px;font-weight:600;margin-right:4px}}
.brands button{{padding:6px 14px;border-radius:20px;border:1px solid #ccc;background:#f5f7fa;color:#333;font-size:13px;cursor:pointer;transition:all .15s;display:flex;align-items:center;gap:6px;min-height:36px}}
.brands button:hover{{border-color:#1976d2;color:#1976d2;background:#e3f2fd}}
.brands button.active{{background:#1976d2;color:#fff;border-color:#1976d2}}
.brands button .cnt{{font-size:11px;background:rgba(0,0,0,.1);padding:1px 6px;border-radius:10px}}
.brands button.active .cnt{{background:rgba(255,255,255,.25)}}
.contact-bar{{background:#e8edf3;padding:8px 30px;display:flex;align-items:center;gap:16px;font-size:13px;color:#444;flex-wrap:wrap}}
.contact-bar a{{color:#1565c0;text-decoration:none;font-weight:500}}
.contact-bar a:hover{{text-decoration:underline}}
.contact-bar .sep{{color:#bbb}}
.footer{{text-align:center;padding:16px;font-size:12px;color:#999}}
td.cmp-cell{{text-align:center;width:50px}}
th.cmp-cell{{text-align:center;width:50px;cursor:default}}
th.cmp-cell:hover{{background:#1565c0}}
td.cmp-cell input[type=checkbox]{{width:18px;height:18px;cursor:pointer;accent-color:#1976d2}}
.compare-bar{{background:#e3f2fd;padding:10px 30px;display:none;align-items:center;gap:12px;border-bottom:1px solid #bbdefb;font-size:13px;flex-wrap:wrap}}
.compare-bar.visible{{display:flex}}
.compare-bar button{{padding:8px 18px;border:none;border-radius:4px;font-size:14px;cursor:pointer;font-weight:500}}
.compare-bar .btn-compare{{background:#1565c0;color:#fff}}
.compare-bar .btn-compare:hover{{background:#0d47a1}}
.compare-bar .btn-clear{{background:#eee;color:#333;border:1px solid #ccc}}
.compare-bar .btn-clear:hover{{background:#ddd}}
.compare-bar .cnt-info{{color:#1565c0;font-weight:600}}
.modal-overlay{{display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,.5);z-index:100;justify-content:center;align-items:flex-start;padding:40px 20px;overflow-y:auto}}
.modal-overlay.visible{{display:flex}}
.modal{{background:#fff;border-radius:8px;max-width:95vw;width:auto;box-shadow:0 8px 32px rgba(0,0,0,.3);overflow:hidden;max-height:85vh;display:flex;flex-direction:column}}
.modal-header{{background:#1565c0;color:#fff;padding:14px 20px;display:flex;justify-content:space-between;align-items:center;flex-shrink:0}}
.modal-header h2{{font-size:18px;font-weight:600}}
.modal-close{{background:none;border:none;color:#fff;font-size:24px;cursor:pointer;padding:0 4px}}
.modal-body{{padding:16px;overflow:auto;flex:1}}
.spec-table{{border-collapse:collapse;width:100%}}
.spec-table th,.spec-table td{{padding:8px 12px;border:1px solid #ddd;font-size:13px;text-align:left;white-space:nowrap}}
.spec-table thead th{{background:#f5f5f5;position:sticky;top:0;z-index:1;font-weight:600;max-width:200px;overflow:hidden;text-overflow:ellipsis}}
.spec-table tbody th{{background:#fafafa;font-weight:600;white-space:nowrap;min-width:120px}}
.spec-table td{{min-width:100px;max-width:220px;overflow:hidden;text-overflow:ellipsis}}
.spec-table tr:hover td,.spec-table tr:hover th{{background:#e3f2fd}}
.spec-val-match{{color:#2e7d32;font-weight:600}}
.spec-val-diff{{color:#c62828}}
.spec-val-empty{{color:#bbb;font-style:italic}}
.has-specs{{color:#1565c0;font-size:10px;vertical-align:super;margin-left:3px}}
.spec-img-row th{{background:#f5f5f5;text-align:center;font-size:12px;color:#666;padding:6px 8px}}
.spec-img-row td{{text-align:center;padding:10px 8px;background:#fafafa}}
.spec-img-row img{{max-width:160px;max-height:120px;object-fit:contain;border:1px solid #eee;border-radius:4px;background:#fff}}
@media(max-width:768px){{
.header{{padding:14px 16px;flex-wrap:wrap}}
.header h1{{font-size:18px}}
.header .subtitle{{font-size:12px}}
.header .logo img{{height:28px}}
.toolbar{{padding:10px 16px;gap:8px}}
.toolbar input[type=text]{{min-width:0;width:100%;flex-basis:100%}}
.toolbar select,.toolbar button{{flex:1;min-width:0;text-align:center}}
.toolbar .info{{width:100%;text-align:center;margin-left:0}}
.container{{padding:10px 8px}}
table{{min-width:420px;font-size:12px}}
th{{padding:8px 6px;font-size:11px}}
th.num{{width:70px}}
td{{padding:6px 6px;font-size:12px}}
.brands{{padding:10px 16px;gap:6px}}
.brands button{{padding:5px 10px;font-size:12px;min-height:32px}}
.contact-bar{{padding:8px 16px;font-size:12px;gap:8px}}
.contact-bar .sep{{display:none}}
.compare-bar{{padding:8px 16px}}
.pagination button{{padding:10px 14px;flex:1;text-align:center}}
.footer{{font-size:11px}}
}}
@media(max-width:400px){{
.header{{padding:10px 12px}}
.header h1{{font-size:16px}}
.container{{padding:8px 4px}}
table{{min-width:360px}}
th{{padding:6px 4px;font-size:10px}}
td{{padding:5px 4px;font-size:11px}}
.brands{{padding:8px 12px}}
.brands .label{{width:100%;margin-bottom:4px}}
}}
</style>
</head>
<body>
<div class="header">
<div class="logo">
<a href="https://electronmir.ru/for_partners/" target="_blank"><img src="https://electronmir.ru/upload/slam.options/ff2/sjtd8wuaptpti4hhomaeucqlijm4zjr5/header_logo-_2_.svg" alt="Электронный мир" style="height:36px"></a>
<div><h1>ЭЛЕКТРОННЫЙ МИР</h1><div class="subtitle">Общие остатки по складам</div></div>
</div>
<div style="font-size:13px;opacity:.8" id="dateEl"></div>
</div>
<div class="contact-bar">
<span>&#9993; Менеджер: <a href="mailto:anikin_s@elektronmir.com">anikin_s@elektronmir.com</a></span>
<span class="sep">|</span>
<span>&#9742; <a href="tel:+79885315067">+7 988 531 50 67</a></span>
<span class="sep">|</span>
<span>Аникин Сергей Александрович</span>
<span class="sep">|</span>
<span><a href="https://t.me/Asa88serwer" target="_blank">Telegram</a></span>
<span class="sep">|</span>
<span><a href="https://max.ru/u/f9LHodD0cOLxsJKEgdxMwcWarCIrWLrkGd4onIZR9HJBtCEZ_" target="_blank">MAX</a></span>
</div>
<div class="brands">
<span class="label">Бренды:</span>
{brand_buttons}</div>
<div class="toolbar">
<input type="text" id="search" placeholder="Поиск по наименованию..." autocomplete="off">
<select id="filter"><option value="all">Все позиции</option><option value="minsk">Только {col1_name} > 0</option><option value="moscow">Только {col2_name} > 0</option><option value="both">Оба склада > 0</option><option value="zero">Нулевые остатки</option></select>

<div class="info" id="info"></div>
</div>
<div class="compare-bar" id="compareBar">
<span class="cnt-info" id="cmpCount">0 выбрано</span>
<button class="btn-compare" onclick="showCompare()">Сравнить выбранные</button>
<button class="btn-clear" onclick="clearCompare()">Сбросить</button>
</div>
<div class="container">
<table><thead><tr>
<th class="cmp-cell">&#9745;</th>
<th style="width:50px">#</th>
<th id="thName" onclick="sortBy(0)">Наименование <span class="arrow">&#9650;</span></th>
<th class="num" id="thMinsk" onclick="sortBy(1)">{col1_name} <span class="arrow">&#9650;</span></th>
<th class="num" id="thMoscow" onclick="sortBy(2)">{col2_name} <span class="arrow">&#9650;</span></th>
<th class="num" style="width:100px">Итого</th>
</tr></thead><tbody id="tbody"></tbody></table>
<div class="pagination">
<div><button id="prevBtn" onclick="changePage(-1)">&larr; Назад</button>
<button id="nextBtn" onclick="changePage(1)">Вперёд &rarr;</button></div>
<div id="pageInfo"></div>
</div>
</div>
<div class="modal-overlay" id="specModal">
<div class="modal">
<div class="modal-header"><h2>Сравнение характеристик</h2><button class="modal-close" onclick="closeSpecModal()">&times;</button></div>
<div class="modal-body" id="specModalBody"></div>
</div>
</div>
<div class="footer">Электронный мир &copy; 2026 | Данные обновлены: {now} | Источник: {source}</div>
<script>
document.getElementById('dateEl').textContent='Обновлено: {now}';
var DATA={data_json};
var SPECS={specs_json};
var SPEC_LABELS={spec_labels_json};
var IMAGES={images_json};
var PAGE=200,page=0,sortCol=0,sortAsc=true,filtered=DATA.slice(),curBrand='';
var checked={{}},compareMode=false;
var searchEl=document.getElementById('search'),filterEl=document.getElementById('filter');
document.querySelectorAll('.brands button').forEach(function(btn){{
btn.addEventListener('click',function(){{
document.querySelectorAll('.brands button').forEach(function(b){{b.classList.remove('active');}});
this.classList.add('active');
curBrand=this.dataset.brand;
searchEl.value='';
applyFilter();
}});
}});
function applyFilter(){{
var q=searchEl.value.toLowerCase(),f=filterEl.value;
filtered=DATA.filter(function(r){{
if(curBrand&&r[0].toLowerCase().indexOf(curBrand.toLowerCase())!==0)return false;
if(q&&r[0].toLowerCase().indexOf(q)===-1)return false;
if(f==='minsk')return r[1]>0;
if(f==='moscow')return r[2]>0;
if(f==='both')return r[1]>0&&r[2]>0;
if(f==='zero')return r[1]===0&&r[2]===0;
return true;
}});
doSort();page=0;render();
}}
function doSort(){{
var c=sortCol,a=sortAsc;
filtered.sort(function(x,y){{var v=x[c]<y[c]?-1:x[c]>y[c]?1:0;return a?v:-v;}});
}}
function sortBy(c){{if(sortCol===c)sortAsc=!sortAsc;else{{sortCol=c;sortAsc=c===0;}}doSort();page=0;render();}}
function escHtml(s){{return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}}
function highlightText(text,q){{if(!q)return escHtml(text);var i=text.toLowerCase().indexOf(q);if(i===-1)return escHtml(text);return escHtml(text.substring(0,i))+'<span class="highlight">'+escHtml(text.substring(i,i+q.length))+'</span>'+escHtml(text.substring(i+q.length));}}
function updateCompareBar(){{
var cnt=Object.keys(checked).length;
document.getElementById('cmpCount').textContent=cnt+' выбрано';
var bar=document.getElementById('compareBar');
if(cnt>0)bar.classList.add('visible');else{{bar.classList.remove('visible');if(compareMode)exitCompare();}}
}}
function toggleCheck(name){{
if(checked[name])delete checked[name];else checked[name]=true;
updateCompareBar();render();
}}
function showCompare(){{
var names=Object.keys(checked);
if(names.length===0)return;
var hasAnySpecs=names.some(function(n){{return SPECS[n];}});
if(hasAnySpecs){{openSpecModal(names);}}
compareMode=true;
filtered=DATA.filter(function(r){{return checked[r[0]];}});
doSort();page=0;render();
document.getElementById('info').textContent='Сравнение: '+filtered.length+' позиций';
}}
function openSpecModal(names){{
var hasImages=names.some(function(n){{return IMAGES[n];}});
var html='<table class="spec-table"><thead>';
if(hasImages){{
html+='<tr class="spec-img-row"><th>Фото</th>';
for(var i=0;i<names.length;i++){{
var imgUrl=IMAGES[names[i]]||'';
html+='<td>'+(imgUrl?'<img src="'+imgUrl+'" alt="'+escHtml(names[i])+'" loading="lazy">':'<span style="color:#ccc;font-size:12px">нет фото</span>')+'</td>';
}}
html+='</tr>';
}}
html+='<tr><th>Характеристика</th>';
for(var i=0;i<names.length;i++){{
var short=names[i].length>30?names[i].substring(0,30)+'...':names[i];
html+='<th title="'+escHtml(names[i])+'">'+escHtml(short)+'</th>';
}}
html+='</tr></thead><tbody>';
for(var s=0;s<SPEC_LABELS.length;s++){{
var key=SPEC_LABELS[s][0],label=SPEC_LABELS[s][1];
var vals=[];
var hasAny=false;
for(var i=0;i<names.length;i++){{
var sp=SPECS[names[i]];
var v=sp?sp[key]||'':'';
vals.push(v);
if(v)hasAny=true;
}}
if(!hasAny)continue;
html+='<tr><th>'+escHtml(label)+'</th>';
var allSame=vals.every(function(v){{return v===vals[0];}});
for(var i=0;i<vals.length;i++){{
var cls='';
if(!vals[i])cls=' class="spec-val-empty"';
else if(vals.filter(function(v){{return v;}}).length>1){{
cls=allSame?' class="spec-val-match"':' class="spec-val-diff"';
}}
html+='<td'+cls+'>'+(vals[i]?escHtml(vals[i]):'—')+'</td>';
}}
html+='</tr>';
}}
html+='</tbody></table>';
document.getElementById('specModalBody').innerHTML=html;
document.getElementById('specModal').classList.add('visible');
}}
function closeSpecModal(){{
document.getElementById('specModal').classList.remove('visible');
}}
function exitCompare(){{
compareMode=false;applyFilter();
}}
function clearCompare(){{
checked={{}};compareMode=false;updateCompareBar();applyFilter();
}}
function render(){{
var start=page*PAGE,end=Math.min(start+PAGE,filtered.length);
var q=searchEl.value.toLowerCase();
var html='';
for(var i=start;i<end;i++){{
var r=filtered[i],total=r[1]+r[2],cls=total===0?' class="zero"':'';
var chk=checked[r[0]]?' checked':'';
var specMark=SPECS[r[0]]?'<span class="has-specs" title="Есть характеристики">&#9679;</span>':'';
html+='<tr'+cls+'><td class="cmp-cell"><input type="checkbox" data-idx="'+i+'"'+chk+'></td><td>'+(i+1)+'</td><td class="product-name">'+highlightText(r[0],q)+specMark+'</td><td class="num">'+r[1].toLocaleString('ru-RU')+'</td><td class="num">'+r[2].toLocaleString('ru-RU')+'</td><td class="num"><b>'+total.toLocaleString('ru-RU')+'</b></td></tr>';
}}
document.getElementById('tbody').innerHTML=html;
if(!compareMode)document.getElementById('info').textContent='Найдено: '+filtered.length.toLocaleString('ru-RU')+' из '+DATA.length.toLocaleString('ru-RU');
var pages=Math.ceil(filtered.length/PAGE)||1;
document.getElementById('pageInfo').textContent='Страница '+(page+1)+' из '+pages;
document.getElementById('prevBtn').disabled=page===0;
document.getElementById('nextBtn').disabled=page>=pages-1;
document.querySelectorAll('th').forEach(function(th,idx){{th.classList.remove('sorted');}});
var thIds=['thName','thMinsk','thMoscow'];
if(sortCol<3){{var el=document.getElementById(thIds[sortCol]);el.classList.add('sorted');el.querySelector('.arrow').innerHTML=sortAsc?'&#9650;':'&#9660;';}}
}}
function changePage(d){{page+=d;render();window.scrollTo(0,0);}}

document.getElementById('tbody').addEventListener('change',function(e){{
if(e.target.type==='checkbox'&&e.target.dataset.idx!==undefined){{
var idx=parseInt(e.target.dataset.idx);
var name=filtered[idx][0];
toggleCheck(name);
}}
}});
searchEl.addEventListener('input',function(){{clearTimeout(this._t);this._t=setTimeout(applyFilter,200);}});
filterEl.addEventListener('change',applyFilter);
applyFilter();
document.addEventListener('contextmenu',function(e){{e.preventDefault();}});
document.addEventListener('keydown',function(e){{
var k=e.key&&e.key.toLowerCase();
if(e.key==='F12'){{e.preventDefault();return false;}}
if((e.ctrlKey||e.metaKey)&&(k==='u'||k==='s'||k==='p')){{e.preventDefault();return false;}}
if((e.ctrlKey||e.metaKey)&&e.shiftKey&&(k==='i'||k==='j'||k==='c')){{e.preventDefault();return false;}}
}});
document.addEventListener('dragstart',function(e){{e.preventDefault();}});
document.getElementById('specModal').addEventListener('click',function(e){{if(e.target===this)closeSpecModal();}});
document.addEventListener('keydown',function(e){{if(e.key==='Escape')closeSpecModal();}});
</script>
</body>
</html>'''

# index.html — публикуемая на GitHub Pages версия; второй файл — для просмотра/отправки
out_paths = [
    os.path.join(DIR, "index.html"),
    os.path.join(DIR, "Общие остатки - Корпоративный.html"),
]
for out_path in out_paths:
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"HTML обновлён: {out_path}")

print(f"Товаров: {len(data)}, Дата: {now}")
for b in BRANDS:
    c = brand_counts.get(b, 0)
    if c > 0:
        print(f"  {b}: {c}")
