"""
HTML-дашборд (Chart.js): data/reports/Dashboard_2025.html

Вкладки: Результат (план-факт), Деньги (ДДС и мост прибыль -> деньги), Врачи (грейды, возвраты),
Факторный мост (валовая маржа план -> факт). Фильтр клиники общий. Данные встроены в страницу.
Все числа берутся из тех же функций, что и Excel-отчёты (build_factor, build_cashflow, build_motivation).
"""
import json
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from build_cashflow import load as load_cash  # noqa: E402
from build_factor import DIRS, effects, prepare, prepare_staff, staff_effects  # noqa: E402
from build_motivation import doctor_table  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
CORR = ROOT / "data" / "corrected"
OUT = ROOT / "data" / "reports" / "Dashboard_2025.html"
MLN = 1e6
GRADES = ["lead", "senior", "middle", "junior"]
BR = [("99", "Прибыль до налога"), ("02", "Амортизация"), ("10", "Запасы материалов"), ("62.02", "Дебиторка ДМС"),
      ("62.01", "Расчёты с пациентами"), ("60", "Долг поставщикам"), ("70", "Долг по зарплате"), ("68", "Долг по НДФЛ"),
      ("69", "Долг по взносам"), ("01", "Капвложения")]
FACT_LBL = [("volume", "Объём"), ("mix", "Структура услуг"), ("price_list", "Цена: прайс"), ("discount", "Цена: скидки"),
            ("materials", "Материалы"), ("labs", "Лаборатории"), ("doc_vol", "Врачи: объём"), ("doc_price", "Врачи: цена и скидки"),
            ("doc_grade", "Врачи: грейды и доля"), ("nurse_eff", "Медперсонал"), ("depr_eff", "Амортизация")]


def r2(x):
    return round(float(x) / MLN, 3)


def monthly(s, per):
    return [r2(s.get(p, 0.0)) for p in per]


def main():
    d, _ = prepare()
    e = effects(d)
    se = staff_effects(prepare_staff(e))
    opex = pd.read_csv(CORR / "budget_opex.csv")
    opex = opex[opex.pnl_block != "cost_of_sales"]
    post = pd.read_csv(CORR / "postings_1c.csv", dtype={"debit": str, "credit": str})
    post = post[post.doc_type != "Закрытие месяца"].copy()
    post["period"] = post.date.str[:7]
    dds, osv = load_cash()
    dt = doctor_table()
    per = [f"2025-{m:02d}" for m in range(1, 13)]
    out = {}
    for key in ["ALL", "C01", "C02", "C03"]:
        f = (lambda x: x) if key == "ALL" else (lambda x, k=key: x[x.clinic_id == k])
        ek, sk, ok, pk, dk = f(e), f(se), f(opex), f(post), f(dt)
        rev_p = ek.groupby("period").plan_revenue.sum()
        rev_f = ek.groupby("period").fact_revenue.sum()
        gm_p = ek.groupby("period").plan_margin.sum() - sk.groupby("period").plan_costs.sum()
        gm_f = ek.groupby("period").fact_margin.sum() - sk.groupby("period").fact_costs.sum()
        pr_p = gm_p - ok.groupby("period").plan_amount.sum()
        cost = pk[pk.debit.isin(["20", "26", "44"])].groupby("period").amount.sum()
        pr_f = pk[pk.credit == "90.01"].groupby("period").amount.sum() - cost
        dc, oc = f(dds.rename(columns={"clinic_id": "clinic_id"})), f(osv)
        net = dc.groupby("m").signed_amount.sum()
        open51 = oc[(oc.acct == "сч. 51") & (oc.m == 1)].open_net.sum()
        close = [open51 + net.loc[:m].sum() for m in range(1, 13)]
        opn = dc[dc.activity == "operating"].signed_amount.sum()
        j, dcl = oc[oc.m == 1].groupby("acct").open_net.sum(), oc[oc.m == 12].groupby("acct").close_net.sum()
        bridge = [[lbl, r2(-(dcl.get(f"сч. {a}", 0) - j.get(f"сч. {a}", 0)))] for a, lbl in BR]
        tot = sk[["doc_vol", "doc_price", "doc_grade", "nurse_eff", "depr_eff"]].sum()
        te = ek[["volume", "mix", "price_list", "discount", "materials", "labs"]].sum()
        fac = {**te.to_dict(), **tot.to_dict()}
        dn = ek.groupby("direction_code")[["plan_revenue", "fact_revenue"]].sum()
        dk = dk.assign(margin=dk.revenue - dk.norm_cost - dk.pay_total, ret_n=dk.cohort_n * dk.returned)
        gr = dk.groupby("grade").agg(n=("doctor_id", "count"), rev=("revenue", "sum"), pay=("pay_total", "sum"), mg=("margin", "sum"))
        rt = dk.groupby("direction_code").agg(c=("cohort_n", "sum"), r=("ret_n", "sum"))
        out[key] = dict(
            rev_p=monthly(rev_p, per), rev_f=monthly(rev_f, per), gm_p=monthly(gm_p, per), gm_f=monthly(gm_f, per),
            pr_p=monthly(pr_p, per), pr_f=monthly(pr_f, per), cash_net=[r2(net.get(m, 0)) for m in range(1, 13)],
            cash_close=[r2(c) for c in close], op=r2(opn), bridge=bridge,
            factors=[[lbl, r2(fac[k])] for k, lbl in FACT_LBL],
            dirs={c: [r2(dn.loc[c, "plan_revenue"]), r2(dn.loc[c, "fact_revenue"])] for c in dn.index},
            grades={g: [int(gr.loc[g, "n"]), r2(gr.loc[g, "rev"]), r2(gr.loc[g, "pay"]), r2(gr.loc[g, "mg"])] for g in GRADES if g in gr.index},
            ret={c: round(float(rt.loc[c, "r"] / rt.loc[c, "c"]), 4) for c in rt.index})
    a = out["ALL"]
    print(f"Контроль ALL: выручка план {sum(a['rev_p']):.2f} факт {sum(a['rev_f']):.2f}; валовая план {sum(a['gm_p']):.2f} факт {sum(a['gm_f']):.2f}; "
          f"прибыль план {sum(a['pr_p']):.2f} факт {sum(a['pr_f']):.2f}; деньги {sum(a['cash_net']):.2f}, остаток {a['cash_close'][-1]:.2f}; "
          f"мост {sum(v for _, v in a['bridge']):.2f}")
    html = TEMPLATE.replace("__DATA__", json.dumps({"d": out, "dirs": DIRS}, ensure_ascii=False, separators=(",", ":")))
    OUT.write_text(html, encoding="utf-8")
    print(f"Сохранено: {OUT} ({len(html) / 1024:.0f} КБ)")


TEMPLATE = r"""<!DOCTYPE html>
<html lang="ru"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>Сеть клиник: дашборд 2025</title>
<style>
:root{box-sizing:border-box;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px);color-scheme:light;
--bg:#f6f7f9;--card:#fff;--tx:#1b2430;--mu:#5f6b7a;--bd:#dde2e8;--ac:#1f5fbf;--ok:#1a8a4a;--bad:#c0392b;--grid:#e6eaef}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){color-scheme:dark;--bg:#12161c;--card:#1b222b;--tx:#e8ecf1;--mu:#9aa6b4;--bd:#2c3642;--ac:#6ea8ff;--ok:#4cc27f;--bad:#ff7a6b;--grid:#2a3440}}
:root[data-theme="dark"]{color-scheme:dark;--bg:#12161c;--card:#1b222b;--tx:#e8ecf1;--mu:#9aa6b4;--bd:#2c3642;--ac:#6ea8ff;--ok:#4cc27f;--bad:#ff7a6b;--grid:#2a3440}
html{scroll-padding-top:env(safe-area-inset-top,0px)}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--tx);font:14px/1.45 -apple-system,"Segoe UI",Roboto,Arial,sans-serif}
.w{max-width:1100px;margin:0 auto;padding:16px}h1{font-size:20px;margin:0 0 4px}.sub{color:var(--mu);font-size:12px;margin-bottom:12px}
.bar{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:12px}.bar button{border:1px solid var(--bd);background:var(--card);color:var(--tx);padding:7px 12px;border-radius:8px;cursor:pointer;font:inherit}
.bar button.on{background:var(--ac);border-color:var(--ac);color:#fff}.sp{flex:1}
.kp{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:10px;margin-bottom:12px}
.k,.c{background:var(--card);border:1px solid var(--bd);border-radius:10px;padding:12px}.k b{display:block;font-size:20px}.k span{color:var(--mu);font-size:12px}.k i{font-style:normal;font-size:12px}
.g{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,460px),1fr));gap:12px}.c h3{margin:0 0 8px;font-size:14px}.c canvas{max-height:300px}
.ok{color:var(--ok)}.bad{color:var(--bad)}.n{color:var(--mu);font-size:12px;margin-top:8px}
table{border-collapse:collapse;width:100%}th,td{padding:5px 8px;text-align:right;border-bottom:1px solid var(--bd)}th:first-child,td:first-child{text-align:left}.t{overflow-x:auto}
.p{display:none}.p.on{display:block}
</style></head><body><div class="w">
<h1>Сеть клиник: управленческий дашборд, 2025</h1>
<div class="sub">Синтетические данные (демо). Те же числа, что в Excel-отчётах: P&amp;L, план-факт, факторный анализ, ДДС. Суммы в млн ₽.</div>
<div class="bar" id="cl"></div><div class="bar" id="tb"></div>
<div id="e1" class="p on"><div class="kp" id="k1"></div><div class="g">
<div class="c"><h3>Выручка: план и факт по месяцам</h3><canvas id="a1"></canvas></div>
<div class="c"><h3>Прибыль до налога: план и факт</h3><canvas id="a2"></canvas></div>
<div class="c"><h3>Валовая маржа, % выручки</h3><canvas id="a3"></canvas></div>
<div class="c"><h3>Факт к плану выручки по направлениям, %</h3><canvas id="a4"></canvas></div></div></div>
<div id="e2" class="p"><div class="kp" id="k2"></div><div class="g">
<div class="c"><h3>Чистое изменение денег и остаток на конец месяца</h3><canvas id="b1"></canvas></div>
<div class="c"><h3>Мост: прибыль до налога → изменение денег за год</h3><canvas id="b2"></canvas>
<div class="n">Деньги ≠ прибыль: амортизация не требует оттока, долги по зарплате и налогам ещё не оплачены, дебиторка ДМС и капвложения связывают деньги.</div></div></div></div>
<div id="e3" class="p"><div class="kp" id="k3"></div><div class="g">
<div class="c"><h3>Маржа после оплаты врачей по грейдам, % выручки</h3><canvas id="c1"></canvas>
<div class="n">Маржа = выручка − материалы и лаб. (норматив) − оплата врача с взносами. Допущение демо: взносы 22%.</div></div>
<div class="c"><h3>Повторные визиты за 90 дней по направлениям, %</h3><canvas id="c2"></canvas>
<div class="n">Малая выборка (6-12 врачей на направление): выводы о конкретном враче по одному году делать нельзя.</div></div></div>
<div class="c" style="margin-top:12px"><h3>Грейды</h3><div class="t" id="gt"></div></div></div>
<div id="e4" class="p"><div class="kp" id="k4"></div><div class="g">
<div class="c" style="grid-column:1/-1"><h3>Факторный мост: валовая маржа, план → факт</h3><canvas id="d1"></canvas></div></div>
<div class="c" style="margin-top:12px"><h3>Как читать</h3><div class="n" style="margin:0">«Врачи: объём» и «цена и скидки» положительны потому, что оплата врачей считается от выручки и снижается вместе с ней: это следствие недобора, а не экономия.
Реальная эффективность затрат: «Врачи: грейды и доля».</div></div></div>
<div class="sub" style="margin-top:12px">Допущения: взносы 22% плоской ставкой; без налога на прибыль и НДС; план по деньгам не составлялся; эффект объёма смешивает фоновый недобор и разовые события.</div>
</div>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.js"></script>
<script>
const DATA=__DATA__;const D=DATA.d,DN=DATA.dirs;
const MS=["янв","фев","мар","апр","май","июн","июл","авг","сен","окт","ноя","дек"],CN={ALL:"Вся сеть",C01:"C01 «Невский»",C02:"C02 «Север»",C03:"C03 «Юг»"};
const TABS=[["e1","Результат"],["e2","Деньги"],["e3","Врачи"],["e4","Факторный мост"]];
let cl="ALL",tab="e1",ch={};
const sum=a=>a.reduce((x,y)=>x+y,0),f1=v=>v.toLocaleString("ru-RU",{minimumFractionDigits:1,maximumFractionDigits:1}),
pc=v=>(v*100).toLocaleString("ru-RU",{minimumFractionDigits:1,maximumFractionDigits:1})+"%",
sg=v=>(v>0?"+":v<0?"−":"")+f1(Math.abs(v)),cs=n=>getComputedStyle(document.documentElement).getPropertyValue(n).trim();
function card(l,v,s,cls){return `<div class="k"><span>${l}</span><b>${v}</b><i class="${cls||""}">${s||""}</i></div>`}
function mk(id,cfg){if(ch[id])ch[id].destroy();const tx=cs("--tx"),gr=cs("--grid");
cfg.options=Object.assign({responsive:true,maintainAspectRatio:true,aspectRatio:1.7,plugins:{legend:{labels:{color:tx}}}},cfg.options||{});
cfg.options.scales=cfg.options.scales||{};for(const k of ["x","y","y1"]){if(cfg.options.scales[k]||k!="y1"){cfg.options.scales[k]=Object.assign({ticks:{color:tx},grid:{color:gr}},cfg.options.scales[k]||{})}}
ch[id]=new Chart(document.getElementById(id),cfg)}
function wf(id,items){let cum=0,base=[],up=[],dn=[],tt=[];items.forEach(([l,v,t])=>{if(t){base.push(0);up.push(0);dn.push(0);tt.push(v);cum=v}else{base.push(Math.min(cum,cum+v));up.push(v>=0?v:0);dn.push(v<0?-v:0);tt.push(0);cum+=v}});
mk(id,{type:"bar",data:{labels:items.map(i=>i[0]),datasets:[{label:"b",data:base,backgroundColor:"transparent",stack:"s"},{label:"Итого",data:tt,backgroundColor:cs("--ac"),stack:"s"},{label:"Рост",data:up,backgroundColor:cs("--ok"),stack:"s"},{label:"Снижение",data:dn,backgroundColor:cs("--bad"),stack:"s"}]},
options:{plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>c.datasetIndex?f1(c.raw)+" млн ₽":null}}},scales:{x:{ticks:{maxRotation:60,minRotation:45}},y:{}}}})}
function render(){const x=D[cl],tx=cs("--tx");
const rf=sum(x.rev_f),rp=sum(x.rev_p),gf=sum(x.gm_f),gp=sum(x.gm_p),pf=sum(x.pr_f),pp=sum(x.pr_p),nt=sum(x.cash_net),cl2=x.cash_close[11];
document.getElementById("k1").innerHTML=card("Выручка",f1(rf),`план ${f1(rp)}, ${pc(rf/rp-1)}`,rf<rp?"bad":"ok")+card("Валовая маржа",pc(gf/rf),`план ${pc(gp/rp)}, ${sg(gf-gp)} млн ₽`,gf<gp?"bad":"ok")+card("Прибыль до налога",f1(pf),`план ${f1(pp)}, ${sg(pf-pp)} млн ₽`,pf<pp?"bad":"ok")+card("Остаток денег на конец года",f1(cl2),`изменение ${sg(nt)} млн ₽`);
mk("a1",{type:"bar",data:{labels:MS,datasets:[{label:"Факт",data:x.rev_f,backgroundColor:cs("--ac")},{label:"План",data:x.rev_p,type:"line",borderColor:cs("--mu"),backgroundColor:cs("--mu"),pointRadius:2}]}});
mk("a2",{type:"bar",data:{labels:MS,datasets:[{label:"Факт",data:x.pr_f,backgroundColor:cs("--ac")},{label:"План",data:x.pr_p,backgroundColor:cs("--mu")}]}});
mk("a3",{type:"line",data:{labels:MS,datasets:[{label:"Факт",data:x.gm_f.map((v,i)=>v/x.rev_f[i]*100),borderColor:cs("--ac"),backgroundColor:cs("--ac")},{label:"План",data:x.gm_p.map((v,i)=>v/x.rev_p[i]*100),borderColor:cs("--mu"),backgroundColor:cs("--mu"),borderDash:[5,4]}]},options:{scales:{y:{ticks:{callback:v=>v+"%"}}}}});
const dk=Object.keys(x.dirs);mk("a4",{type:"bar",data:{labels:dk.map(k=>DN[k]),datasets:[{label:"Факт к плану, %",data:dk.map(k=>x.dirs[k][1]/x.dirs[k][0]*100),backgroundColor:cs("--ac")}]},options:{indexAxis:"y",plugins:{legend:{display:false}},scales:{x:{min:80,max:100,ticks:{callback:v=>v+"%"}}}}});
document.getElementById("k2").innerHTML=card("Операционный денежный поток",f1(x.op),"")+card("Капвложения",f1(x.bridge[9][1]),"")+card("Чистое изменение денег",f1(nt),"",nt<0?"bad":"ok")+card("Остаток на конец года",f1(cl2),"");
mk("b1",{type:"bar",data:{labels:MS,datasets:[{label:"Чистое изменение",data:x.cash_net,backgroundColor:cs("--ac"),yAxisID:"y"},{label:"Остаток на конец",data:x.cash_close,type:"line",borderColor:cs("--ok"),backgroundColor:cs("--ok"),yAxisID:"y1",pointRadius:2}]},options:{scales:{y1:{position:"right",grid:{drawOnChartArea:false}}}}});
wf("b2",x.bridge.map((b,i)=>[b[0],b[1],i==0]).concat([["Изменение денег",nt,true]]));
const G=x.grades,gk=Object.keys(G),gl={lead:"lead",senior:"senior",middle:"middle",junior:"junior"};
const ar=gk.map(g=>G[g]),tr=sum(ar.map(a=>a[1])),tp=sum(ar.map(a=>a[2])),tm=sum(ar.map(a=>a[3]));
document.getElementById("k3").innerHTML=card("Врачей",sum(ar.map(a=>a[0])),"")+card("Выручка врачей",f1(tr),"")+card("Оплата с взносами, % выручки",pc(tp/tr),"")+card("Маржа после оплаты врачей",pc(tm/tr),"");
mk("c1",{type:"bar",data:{labels:gk,datasets:[{label:"Маржа, %",data:gk.map(g=>G[g][3]/G[g][1]*100),backgroundColor:cs("--ac")},{label:"Оплата, %",data:gk.map(g=>G[g][2]/G[g][1]*100),backgroundColor:cs("--mu")}]},options:{scales:{y:{ticks:{callback:v=>v+"%"}}}}});
const rk=Object.keys(x.ret);mk("c2",{type:"bar",data:{labels:rk.map(k=>DN[k]),datasets:[{label:"Вернулись, %",data:rk.map(k=>x.ret[k]*100),backgroundColor:cs("--ac")}]},options:{indexAxis:"y",plugins:{legend:{display:false}},scales:{x:{ticks:{callback:v=>v+"%"}}}}});
document.getElementById("gt").innerHTML="<table><tr><th>Грейд</th><th>Врачей</th><th>Выручка</th><th>Оплата с взносами</th><th>Доля оплаты</th><th>Маржа</th><th>Маржа, %</th></tr>"+gk.map(g=>{const a=G[g];return `<tr><td>${g}</td><td>${a[0]}</td><td>${f1(a[1])}</td><td>${f1(a[2])}</td><td>${pc(a[2]/a[1])}</td><td>${f1(a[3])}</td><td>${pc(a[3]/a[1])}</td></tr>`}).join("")+"</table>";
const F=x.factors,fs=sum(F.map(a=>a[1])),dv=F.filter((a,i)=>i>=6&&i<=8),gr2=F[8][1];
document.getElementById("k4").innerHTML=card("Валовая маржа: план",f1(gp),pc(gp/rp)+" выручки")+card("Валовая маржа: факт",f1(gf),pc(gf/rf)+" выручки")+card("Отклонение",sg(gf-gp),"млн ₽",gf<gp?"bad":"ok")+card("Эффективность оплаты врачей",sg(gr2),"грейды и доля, млн ₽",gr2<0?"bad":"ok");
wf("d1",[["План",gp,true]].concat(F.map(a=>[a[0],a[1],false]),[["Факт",gf,true]]))}
function ui(){document.getElementById("cl").innerHTML=Object.keys(CN).map(k=>`<button class="${k==cl?"on":""}" data-c="${k}">${CN[k]}</button>`).join("");
document.getElementById("tb").innerHTML=TABS.map(t=>`<button class="${t[0]==tab?"on":""}" data-t="${t[0]}">${t[1]}</button>`).join("");
TABS.forEach(t=>document.getElementById(t[0]).classList.toggle("on",t[0]==tab))}
document.addEventListener("click",ev=>{const b=ev.target.closest("button");if(!b)return;if(b.dataset.c)cl=b.dataset.c;if(b.dataset.t){tab=b.dataset.t;history.replaceState(null,"","#"+tab)}ui();render()});
if(location.hash&&TABS.some(t=>t[0]==location.hash.slice(1)))tab=location.hash.slice(1);
ui();if(typeof Chart==="undefined"){document.getElementById("k1").innerHTML="<div class='k'>Не загрузился Chart.js: проверьте подключение к сети.</div>"}else{render()}
</script></body></html>"""

if __name__ == "__main__":
    main()
