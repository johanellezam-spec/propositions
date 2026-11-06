"""Génère une page web autonome (un seul fichier .html) à partir d'un Excel réseau MediaHub Campus.
Usage : python3 reseau_html.py <reseau.xlsx> <sortie.html> "<Client>" "<date en toutes lettres>" [lien_excel]
"""
import sys, json, html
import openpyxl

src, out, client, date = sys.argv[1:5]
xlsx_link = sys.argv[5] if len(sys.argv) > 5 else ""

ws = openpyxl.load_workbook(src, read_only=True).active
rows = [["" if v is None else str(v).strip() for v in r] for r in ws.iter_rows(values_only=True)]
head, data = rows[0], [r for r in rows[1:] if any(r)]
idx = {h: i for i, h in enumerate(head)}
def col(name):
    return idx.get(name)
ville_i, region_i, cat_i = col("Ville"), col("Région"), col("Catégorie")
nb_villes = len({r[ville_i] for r in data}) if ville_i is not None else 0
nb_regions = len({r[region_i] for r in data}) if region_i is not None else 0
cats = sorted({r[cat_i] for r in data if r[cat_i]}) if cat_i is not None else []
villes = sorted({r[ville_i] for r in data if r[ville_i]}) if ville_i is not None else []

title = f"Réseau {client} - {date}"
dl = (f'<a class="btn" href="{html.escape(xlsx_link)}" download>Télécharger en Excel</a>'
      if xlsx_link else "")

page = f"""<!doctype html>
<html lang="fr"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="robots" content="noindex,nofollow">
<title>{html.escape(title)} | MediaHub Campus</title>
<style>
:root{{--ink:#14213d;--muted:#5c6b80;--line:#e3e8ef;--bg:#f6f8fb;--card:#fff;--accent:#ff5a1f}}
*{{box-sizing:border-box}}body{{margin:0;font-family:Tahoma,Geneva,Verdana,sans-serif;color:var(--ink);background:var(--bg);font-size:14px}}
header{{background:var(--ink);color:#fff;padding:22px 16px}}
.wrap{{max-width:1200px;margin:0 auto}}
.brand{{font-size:12px;letter-spacing:.12em;text-transform:uppercase;opacity:.75}}
h1{{margin:6px 0 4px;font-size:24px}} .sub{{opacity:.8}}
.kpis{{display:flex;gap:12px;flex-wrap:wrap;margin:18px 0 0}}
.kpi{{background:rgba(255,255,255,.1);border-radius:10px;padding:10px 14px;min-width:120px}}
.kpi b{{display:block;font-size:22px}}
main{{padding:16px}}
.tools{{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-bottom:12px}}
input,select{{font:inherit;padding:9px 10px;border:1px solid var(--line);border-radius:8px;background:#fff;min-width:0}}
input{{flex:1 1 220px}} select{{flex:0 1 200px}}
.btn{{background:var(--accent);color:#fff;text-decoration:none;padding:9px 14px;border-radius:8px;font-weight:bold;white-space:nowrap}}
.count{{color:var(--muted);margin:0 0 8px}}
table{{width:100%;border-collapse:collapse;background:var(--card);border-radius:10px;overflow:hidden;box-shadow:0 1px 3px rgba(20,33,61,.08)}}
th{{position:sticky;top:0;background:#eef2f7;text-align:left;font-size:12px;text-transform:uppercase;letter-spacing:.04em;color:var(--muted);padding:10px;cursor:pointer}}
td{{padding:9px 10px;border-top:1px solid var(--line);vertical-align:top}}
tr:hover td{{background:#fafbfd}}
footer{{color:var(--muted);font-size:12px;padding:20px 16px;text-align:center}}
footer a{{color:var(--ink)}}
@media (max-width:720px){{
 thead{{display:none}} table,tbody,tr,td{{display:block;width:100%}}
 tr{{margin-bottom:10px;border-radius:10px;background:#fff;box-shadow:0 1px 3px rgba(20,33,61,.08)}}
 td{{border:0;padding:4px 12px}} td:first-child{{padding-top:10px}} td:last-child{{padding-bottom:10px}}
 td::before{{content:attr(data-h);display:block;font-size:11px;color:var(--muted);text-transform:uppercase}}
 h1{{font-size:20px}}
}}
</style></head><body>
<header><div class="wrap">
<div class="brand">MediaHub Campus · Régie publicitaire OOH Campus &amp; Lycées</div>
<h1>{html.escape(title)}</h1>
<div class="sub">Réseau d'affichage proposé</div>
<div class="kpis"><div class="kpi"><b>{len(data)}</b>établissements</div><div class="kpi"><b>{nb_villes}</b>villes</div><div class="kpi"><b>{nb_regions}</b>régions</div></div>
</div></header>
<main><div class="wrap">
<div class="tools">
<input id="q" type="search" placeholder="Rechercher une ville, une école, une filière…">
<select id="v"><option value="">Toutes les villes</option>{''.join(f'<option>{html.escape(v)}</option>' for v in villes)}</select>
<select id="c"><option value="">Toutes les catégories</option>{''.join(f'<option>{html.escape(c)}</option>' for c in cats)}</select>
{dl}
</div>
<p class="count" id="n"></p>
<table><thead><tr>{''.join(f'<th data-k="{i}">{html.escape(h)}</th>' for i, h in enumerate(head))}</tr></thead><tbody id="b"></tbody></table>
</div></main>
<footer>MediaHub Campus · <a href="https://mediahubcampus.com/">mediahubcampus.com</a></footer>
<script>
const H={json.dumps(head, ensure_ascii=False)},D={json.dumps(data, ensure_ascii=False)},VI={json.dumps(ville_i)},CI={json.dumps(cat_i)};
const b=document.getElementById('b'),q=document.getElementById('q'),v=document.getElementById('v'),c=document.getElementById('c'),n=document.getElementById('n');
let sortK=null,asc=true;
const esc=s=>s.replace(/[&<>"]/g,m=>({{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}}[m]));
function draw(){{const t=q.value.toLowerCase();let r=D.filter(x=>(!v.value||x[VI]===v.value)&&(!c.value||x[CI]===c.value)&&(!t||x.join(' ').toLowerCase().includes(t)));
if(sortK!==null)r=r.slice().sort((a,z)=>a[sortK].localeCompare(z[sortK],'fr')*(asc?1:-1));
n.textContent=r.length+' établissement'+(r.length>1?'s':'')+' affiché'+(r.length>1?'s':'');
b.innerHTML=r.map(x=>'<tr>'+x.map((y,i)=>'<td data-h="'+esc(H[i])+'">'+esc(y)+'</td>').join('')+'</tr>').join('');}}
[q,v,c].forEach(e=>e.addEventListener('input',draw));
document.querySelectorAll('th').forEach(th=>th.onclick=()=>{{const k=+th.dataset.k;asc=sortK===k?!asc:true;sortK=k;draw();}});
draw();
</script></body></html>"""
open(out, "w", encoding="utf-8").write(page)
print(out, len(data), "lignes", round(len(page) / 1024), "Ko")
