import sys, os, json
from datetime import datetime
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8")

from merge_warehouses import extract_all_items

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
now = datetime.now().strftime("%d.%m.%Y %H:%M")

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
<span>Аникин Сергей Александрович</span>
</div>
<div class="brands">
<span class="label">Бренды:</span>
{brand_buttons}</div>
<div class="toolbar">
<input type="text" id="search" placeholder="Поиск по наименованию..." autocomplete="off">
<select id="filter"><option value="all">Все позиции</option><option value="minsk">Только {col1_name} > 0</option><option value="moscow">Только {col2_name} > 0</option><option value="both">Оба склада > 0</option><option value="zero">Нулевые остатки</option></select>

<div class="info" id="info"></div>
</div>
<div class="container">
<table><thead><tr>
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
<div class="footer">Электронный мир &copy; 2026 | Данные обновлены: {now} | Источник: {source}</div>
<script>
document.getElementById('dateEl').textContent='Обновлено: {now}';
var DATA={data_json};
var PAGE=200,page=0,sortCol=0,sortAsc=true,filtered=DATA.slice(),curBrand='';
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
function render(){{
var start=page*PAGE,end=Math.min(start+PAGE,filtered.length);
var q=searchEl.value.toLowerCase();
var html='';
for(var i=start;i<end;i++){{
var r=filtered[i],total=r[1]+r[2],cls=total===0?' class="zero"':'';
html+='<tr'+cls+'><td>'+(i+1)+'</td><td class="product-name">'+highlightText(r[0],q)+'</td><td class="num">'+r[1].toLocaleString('ru-RU')+'</td><td class="num">'+r[2].toLocaleString('ru-RU')+'</td><td class="num"><b>'+total.toLocaleString('ru-RU')+'</b></td></tr>';
}}
document.getElementById('tbody').innerHTML=html;
document.getElementById('info').textContent='Найдено: '+filtered.length.toLocaleString('ru-RU')+' из '+DATA.length.toLocaleString('ru-RU');
var pages=Math.ceil(filtered.length/PAGE)||1;
document.getElementById('pageInfo').textContent='Страница '+(page+1)+' из '+pages;
document.getElementById('prevBtn').disabled=page===0;
document.getElementById('nextBtn').disabled=page>=pages-1;
document.querySelectorAll('th').forEach(function(th,idx){{th.classList.remove('sorted');}});
var thIds=['thName','thMinsk','thMoscow'];
if(sortCol<3){{var el=document.getElementById(thIds[sortCol]);el.classList.add('sorted');el.querySelector('.arrow').innerHTML=sortAsc?'&#9650;':'&#9660;';}}
}}
function changePage(d){{page+=d;render();window.scrollTo(0,0);}}

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
