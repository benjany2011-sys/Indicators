if(!window.__PANEL_APP_LOADED){ window.__PANEL_APP_LOADED = true;
(function(){
// ===== mejoras visuales (rediseño) =====
const TT_THEME = { backgroundColor:"#1b1b1f", borderColor:"#3a3a42", borderWidth:1, titleColor:"#FFC400", bodyColor:"#ECECEC", padding:12, cornerRadius:8, boxPadding:5, caretSize:6 };
const DESTACADAS = new Set(["MXN por USD","Henry Hub ($/MMBtu)","S&P 500","WTI ($/bbl)"]);
const SPARKS = [];
function hexA(hex, a){ const n=parseInt(hex.slice(1),16); return "rgba("+(n>>16&255)+","+(n>>8&255)+","+(n&255)+","+a+")"; }
function colorDelta(c){ if(!c) return "#9aa0a6"; return c.val>0.0005 ? "#37d67a" : (c.val<-0.0005 ? "#ff5b5b" : "#9aa0a6"); }
function registrarSpark(card, vals, color){
  if(!vals || !vals.some(v=>v!==null && v!==undefined)) return;
  const cv=document.createElement("canvas"); cv.className="spark";
  card.appendChild(cv);
  SPARKS.push({cv, vals, color, drawn:false});
}
function dibujarSparks(){
  SPARKS.forEach(s=>{
    if(s.drawn) return;
    const w=s.cv.offsetWidth, h=s.cv.offsetHeight;
    if(!w || !h) return;
    const pts=[]; s.vals.slice(-120).forEach(v=>{ if(v!==null && v!==undefined) pts.push(v); });
    if(pts.length<2){ s.drawn=true; return; }
    const dpr=window.devicePixelRatio||1;
    s.cv.width=w*dpr; s.cv.height=h*dpr;
    const ctx=s.cv.getContext("2d"); ctx.scale(dpr,dpr);
    let min=Math.min(...pts), max=Math.max(...pts);
    if(max-min<1e-12){ max+=1; min-=1; }
    const X=i=> 1 + i/(pts.length-1)*(w-2);
    const Y=v=> h-2-((v-min)/(max-min))*(h-8);
    ctx.beginPath(); pts.forEach((v,i)=> i? ctx.lineTo(X(i),Y(v)) : ctx.moveTo(X(i),Y(v)));
    ctx.lineTo(X(pts.length-1),h); ctx.lineTo(X(0),h); ctx.closePath();
    const g=ctx.createLinearGradient(0,0,0,h);
    g.addColorStop(0,hexA(s.color,.25)); g.addColorStop(1,hexA(s.color,0));
    ctx.fillStyle=g; ctx.fill();
    ctx.beginPath(); pts.forEach((v,i)=> i? ctx.lineTo(X(i),Y(v)) : ctx.moveTo(X(i),Y(v)));
    ctx.strokeStyle=s.color; ctx.lineWidth=1.6; ctx.lineJoin="round"; ctx.lineCap="round"; ctx.stroke();
    s.drawn=true;
  });
}
window.addEventListener("resize", ()=>{ SPARKS.forEach(s=>{ s.drawn=false; }); requestAnimationFrame(dibujarSparks); });
function gradFill(color){
  return function(context){
    const chart=context.chart, area=chart.chartArea;
    if(!area) return hexA(color,.12);
    const g=chart.ctx.createLinearGradient(0,area.top,0,area.bottom);
    g.addColorStop(0,hexA(color,.20)); g.addColorStop(1,hexA(color,0));
    return g;
  };
}
// ===== fin mejoras =====

const DATA_URL = "resultados/datos.json";
let DATA = null;
const charts = {};        // instancias por id
const construido = {};    // pestañas ya construidas

function fmtNum(v, fmt){
  if(v===null || v===undefined) return "—";
  if(fmt==="0.0%") return (v*100).toFixed(1)+"%";
  let dec = 2;
  if(fmt && fmt.indexOf(".")>-1) dec = fmt.split(".")[1].length;
  return v.toLocaleString("es-MX",{minimumFractionDigits:dec, maximumFractionDigits:dec});
}

// últimos dos valores no nulos de un arreglo -> [actual, previo]
function ultimosDos(arr){
  const r=[];
  for(let i=arr.length-1;i>=0 && r.length<2;i--){ if(arr[i]!==null && arr[i]!==undefined) r.push(arr[i]); }
  return r;
}

// promedio móvil sobre una ventana de posiciones (ignora huecos/nulos)
function mediaMovil(arr, ventana){
  const out = new Array(arr.length).fill(null);
  for(let i=0;i<arr.length;i++){
    let suma=0, cuenta=0;
    for(let j=Math.max(0,i-ventana+1); j<=i; j++){
      const v=arr[j];
      if(v!==null && v!==undefined){ suma+=v; cuenta++; }
    }
    if(cuenta>=5) out[i]=suma/cuenta;
  }
  return out;
}

// cambio vs. dato previo disponible (solo para series de resumen: commodities + divisas)
function cambio(serie){
  const arr = DATA.series[serie];
  if(!arr) return null;
  const v = ultimosDos(arr);
  if(v.length<2 || v[1]===0) return null;
  return {pp:false, val:(v[0]-v[1])/v[1]*100};
}

function htmlDelta(c){
  if(!c) return "<span class='delta flat'>—</span>";
  const arriba = c.val > 0.0005, abajo = c.val < -0.0005;
  const cls = arriba ? "up" : (abajo ? "down" : "flat");
  const flecha = arriba ? "▲" : (abajo ? "▼" : "■");
  const signo = c.val>0 ? "+" : "";
  const txt = c.pp ? (signo+c.val.toFixed(1)+" pp")
                   : (signo+c.val.toFixed(2)+"%");
  return "<span class='delta "+cls+"'>"+flecha+" "+txt+"</span>";
}

function tarjeta(r){
  const c = cambio(r.serie);
  const el=document.createElement("div"); el.className="card";
  let interp = "";
  if(esFX(r.serie)){
    const t = interpretaFX(r.serie, c);
    if(t) interp = "<div class='interp'>"+t+"</div>";
  }
  el.innerHTML =
    "<div class='nombre'>"+r.serie+"</div>"+
    "<div class='valor'>"+fmtNum(r.valor, r.fmt)+"</div>"+
    "<div class='delta-linea'>"+htmlDelta(c)+"</div>"+
    interp+
    "<div class='fecha'>"+r.fecha+"</div>";
  if(DESTACADAS.has(r.serie)) el.classList.add("destacada");
  registrarSpark(el, (DATA.series && DATA.series[r.serie])||[], colorDelta(c));
  return el;
}

function esFX(nombre){ return nombre.indexOf(" por USD")>-1 || nombre.indexOf("USD por ")===0; }

// nombre amigable de cada divisa para el texto interpretativo
// dominio de cada acerera, para pedirle el logo a Clearbit (gratis, sin key).
// si una empresa no está aquí, o el logo falla, la tarjeta se ve igual que antes.
const LOGO_DOMINIO_ACERERA = {
  "ArcelorMittal":         "arcelormittal.com",
  "POSCO":                 "posco.co.kr",
  "Ternium":               "ternium.com",
  "Gerdau":                "gerdau.com",
  "Nippon Steel":          "nipponsteel.com",
  "thyssenkrupp":          "thyssenkrupp.com",
  "SSAB":                  "ssab.com",
  "Nucor":                 "nucor.com",
  "Steel Dynamics":        "steeldynamics.com",
  "Cleveland-Cliffs":      "clevelandcliffs.com",
  "Commercial Metals":     "cmc.com",
  "Reliance":              "reliancesteel.com",
  "Worthington":           "worthingtonindustries.com",
  "ATI":                   "atimaterials.com",
  "Carpenter Technology":  "cartech.com",
  "Baoshan Iron & Steel":  "baosteel.com",
  "Angang Steel":          "ansteel.cn",
  "Maanshan Iron & Steel": "magang.com.cn",
  "HBIS":                  "hbisco.com",
  "Shougang":              "shougang.com.cn",
  "Tata Steel":            "tatasteel.com",
  "JSW Steel":             "jsw.in",
  "SAIL":                  "sail.co.in",
  "Jindal Steel":          "jindalsteelpower.com",
  "Insteel Industries":    "insteel.com",
  "Metallus":              "metallusinc.com",
  "Friedman Industries":   "friedmanindustries.com"
};
// <img> del logo, o "" si no tenemos dominio para esa empresa.
// onerror la quita sola si Clearbit no tiene el logo (empresa chica, ADR OTC, etc.)
function logoAcerera(nombre){
  const dom = LOGO_DOMINIO_ACERERA[nombre];
  if(!dom) return "";
  return "<img class='logo-acerera' src='https://img.logo.dev/"+dom+"?token=pk_ViN51m9GS9qe0Tbrjv4Sqw&size=64' alt='' onerror=\"this.remove()\">";
}

const NOMBRE_DIVISA = {
  EUR:"el euro", JPY:"el yen", CNY:"el yuan", GBP:"la libra", MXN:"el peso",
  CAD:"el dólar canadiense", TRY:"la lira turca", BRL:"el real brasileño", AUD:"el dólar australiano",
  INR:"la rupia india", KRW:"el won coreano", COP:"el peso colombiano", HKD:"el dólar de Hong Kong",
};
function capitaliza(s){ return s.charAt(0).toUpperCase()+s.slice(1); }

// interpreta qué le pasó a la DIVISA (no al número) según la columna y el cambio
function interpretaFX(serie, c){
  if(!c) return "";
  let iso, subeEsFuerte;
  if(serie.indexOf("USD por ")===0){ iso = serie.slice(8,11); subeEsFuerte = true; }  // USD por X
  else { iso = serie.slice(0,3); subeEsFuerte = false; }                                // X por USD
  const nombre = NOMBRE_DIVISA[iso] || "la divisa";
  const subio = c.val > 0.0005, bajo = c.val < -0.0005;
  if(!subio && !bajo) return capitaliza(nombre)+" sin cambio";
  const fuerte = subio ? subeEsFuerte : !subeEsFuerte;
  return capitaliza(nombre)+(fuerte ? " se fortaleció" : " se debilitó");
}

// registro el plugin de zoom (si el <script> lo auto-registró, no lo duplico)
function registrarZoom(){
  try{
    let ya=false;
    try{ ya = !!Chart.registry.plugins.get("zoom"); }catch(_){ ya=false; }
    if(!ya && window.ChartZoom) Chart.register(window.ChartZoom);
  }catch(e){}
}

// rueda = zoom; arrastrar = caja de zoom; ctrl+arrastrar = mover; doble clic = reset
const ZOOM_OPTS = {
  zoom:{
    wheel:{enabled:true},
    pinch:{enabled:true},
    drag:{enabled:true, backgroundColor:"rgba(255,157,46,.15)", borderColor:"#FF9D2E", borderWidth:1},
    mode:"xy"
  },
  pan:{ enabled:true, mode:"xy", modifierKey:"ctrl" }
};

const BASE_OPTS = {
  responsive:true, maintainAspectRatio:false,
  interaction:{mode:"index", intersect:false},
  plugins:{ legend:{display:false, labels:{color:"#cfcfcf"}},
    tooltip:{ callbacks:{}, backgroundColor:"#1b1b1f", borderColor:"#3a3a42", borderWidth:1,
      titleColor:"#FFC400", bodyColor:"#ECECEC", padding:12, cornerRadius:8, boxPadding:5,
      titleFont:{weight:"700"}, caretSize:6, boxWidth:8, boxHeight:8 },
    zoom:ZOOM_OPTS },
  elements:{ point:{radius:0, hoverRadius:3.5, hitRadius:8}, line:{borderWidth:2, tension:.15} },
  scales:{
    x:{ ticks:{maxTicksLimit:8, font:{size:11}, color:"#9aa0a6"}, grid:{display:false}, border:{color:"#3a3a42"} },
    y:{ ticks:{font:{size:11}, color:"#9aa0a6"}, grid:{color:"rgba(255,255,255,.06)"}, border:{color:"#3a3a42"} }
  }
};
function lineChart(id, labels, datasets, opts){
  const options = JSON.parse(JSON.stringify(BASE_OPTS));
  if(opts){
    if(opts.plugins && opts.plugins.tooltip) opts.plugins.tooltip = Object.assign({}, options.plugins.tooltip, opts.plugins.tooltip);
    if(opts.plugins) options.plugins = Object.assign(options.plugins, opts.plugins);
    if(opts.scales){
      options.scales.x = Object.assign(options.scales.x||{}, opts.scales.x||{});
      options.scales.y = Object.assign(options.scales.y||{}, opts.scales.y||{});
    }
    for(const k in opts){ if(k!=="plugins" && k!=="scales") options[k]=opts[k]; }
  }
  // Reparte ~8 etiquetas parejas en el eje X pero garantiza que la ÚLTIMA
  // fecha siempre se muestre (si no, Chart.js a veces corta antes de 2026).
  options.scales.x.ticks = options.scales.x.ticks || {};
  options.scales.x.ticks.autoSkip = false;
  options.scales.x.ticks.maxRotation = 0;
  options.scales.x.afterBuildTicks = (axis)=>{
    const n = axis.chart.data.labels.length;
    if(n <= 1) return;
    // si hay zoom, reparto las ~8 etiquetas dentro del rango visible
    let lo = 0, hi = n - 1;
    if(typeof axis.min === "number") lo = Math.ceil(axis.min);
    if(typeof axis.max === "number") hi = Math.floor(axis.max);
    if(hi <= lo){ axis.ticks = [{value:Math.max(0,Math.min(lo,n-1))}]; return; }
    const max = 8, paso = (hi - lo) / (max - 1), idx = [];
    for(let i=0;i<max;i++) idx.push(Math.round(lo + i*paso));
    const uniq = [...new Set(idx)].sort((a,b)=>a-b);
    if(uniq[uniq.length-1] !== hi) uniq.push(hi);
    axis.ticks = uniq.map(v=>({value:v}));
  };
  const c = new Chart(document.getElementById(id), {
    type:"line",
    data:{ labels, datasets:datasets.map((d,i)=>{
      const base = Object.assign({spanGaps:true}, d);
      if(i===0 && !d.borderDash && typeof d.borderColor==="string" && d.borderColor.charAt(0)==="#"){
        base.fill = "origin"; base.backgroundColor = gradFill(d.borderColor);
      }
      return base;
    }) },
    options
  });
  c.canvas.addEventListener("dblclick", ()=>{ try{ c.resetZoom(); }catch(e){} });  // doble clic = reset zoom
  charts[id]=c; return c;
}

function ultimosDosMacro(valores){
  const r=[];
  for(let i=valores.length-1;i>=0 && r.length<2;i--){ if(valores[i]!==null && valores[i]!==undefined) r.push(valores[i]); }
  return r;
}

// busca el índice de un indicador macro por palabra clave en el nombre
function _macroIdx(kw){ return (DATA.macro||[]).findIndex(x=>(x.nombre||"").toLowerCase().includes(kw)); }

function construirMacro(){
  const cont=document.getElementById("cards-macro");
  const graf=document.getElementById("graf-macro");
  const macro = DATA.macro||[];
  // tarjetas: una por indicador, como siempre
  macro.forEach((ind)=>{
    const v = ultimosDosMacro(ind.valores);
    const ultVal = v.length ? v[0]*ind.escala : null;
    let delta=null;
    if(v.length>=2) delta={pp:true, val:(v[0]-v[1])*ind.escala};
    const ultFecha = ind.fecha.length ? ind.fecha[ind.fecha.length-1] : "";
    const card=document.createElement("div"); card.className="card";
    card.innerHTML =
      "<div class='nombre'>"+ind.nombre+"</div>"+
      "<div class='valor'>"+(ultVal===null?"—":ultVal.toFixed(1)+"%")+"</div>"+
      "<div class='delta-linea'>"+htmlDelta(delta)+"</div>"+
      "<div class='fecha'>"+ultFecha+"</div>";
    cont.appendChild(card);
    registrarSpark(card, ind.valores.map(x=> x===null||x===undefined ? null : x*ind.escala), colorDelta(delta));
  });
  // paneles: inflación + desempleo en UNA gráfica; el resto cada uno por su lado
  const iInfl=_macroIdx("inflaci"), iDes=_macroIdx("desemple");
  if(iInfl>=0 && iDes>=0){
    const panel=document.createElement("div"); panel.className="panel ancho";
    panel.innerHTML="<h3>Inflación vs. Desempleo (%)</h3><div class='lienzo'><canvas id='c-macro-combo'></canvas></div>";
    graf.appendChild(panel);
  }
  macro.forEach((ind,k)=>{
    if(k===iInfl || k===iDes) return;   // esos dos ya van juntos arriba
    const panel=document.createElement("div"); panel.className="panel ancho";
    panel.innerHTML="<h3>"+ind.nombre+" (%)</h3><div class='lienzo'><canvas id='c-macro-"+k+"'></canvas></div>";
    graf.appendChild(panel);
  });
}

function graficarMacro(){
  const macro = DATA.macro||[];
  const iInfl=_macroIdx("inflaci"), iDes=_macroIdx("desemple");
  if(iInfl>=0 && iDes>=0){
    const inf=macro[iInfl], des=macro[iDes];
    const vInf=inf.valores.map(x=> x===null?null:x*inf.escala);
    const vDes=des.valores.map(x=> x===null?null:x*des.escala);
    lineChart("c-macro-combo", inf.fecha, [
      {label:inf.nombre, data:vInf, borderColor:"#FF9D2E"},
      {label:des.nombre, data:vDes, borderColor:"#3b82f6"}
    ], {
      plugins:{ legend:{display:true, labels:{boxWidth:12, font:{size:11}, color:"#cfcfcf"}},
        tooltip:{ callbacks:{ label:c=> c.dataset.label+": "+(c.parsed.y!=null? c.parsed.y.toFixed(1)+"%":"—") } } }
    });
  }
  macro.forEach((ind,k)=>{
    if(k===iInfl || k===iDes) return;
    const vals = ind.valores.map(x=> x===null?null:x*ind.escala);
    lineChart("c-macro-"+k, ind.fecha, [{data:vals, borderColor:"#FF9D2E"}], {
      plugins:{ legend:{display:false},
        tooltip:{ callbacks:{ label:c=> (c.parsed.y!=null? c.parsed.y.toFixed(1)+"%":"—") } } }
    });
  });
}

// formatea un valor de construcción según su unidad (prefijo, decimales, sufijo)
function fmtConstr(v, ind){
  if(v===null || v===undefined) return "—";
  const n = v.toLocaleString("es-MX",{minimumFractionDigits:ind.dec, maximumFractionDigits:ind.dec});
  return (ind.prefijo||"")+n+(ind.sufijo||"");
}

function construirConstruccion(){
  const cont=document.getElementById("cards-construccion");
  const graf=document.getElementById("graf-construccion");
  const lista = DATA.construccion||[];
  if(!lista.length){
    if(cont) cont.innerHTML="<div class='card'><div class='nombre'>Construcción</div>"+
      "<div class='interp'>Sin datos en esta corrida.</div></div>";
    return;
  }
  lista.forEach((ind, k)=>{
    const v = ultimosDosMacro(ind.valores);   // [último, previo] no nulos
    const ultVal = v.length ? v[0] : null;
    // estos son niveles, así que el cambio se muestra en % mensual
    let delta=null;
    if(v.length>=2 && v[1]!==0) delta={pp:false, val:(v[0]-v[1])/v[1]*100};
    const ultFecha = ind.fecha.length ? ind.fecha[ind.fecha.length-1] : "";
    const card=document.createElement("div"); card.className="card";
    card.innerHTML =
      "<div class='nombre'>"+ind.nombre+"</div>"+
      "<div class='valor'>"+fmtConstr(ultVal, ind)+"</div>"+
      "<div class='delta-linea'>"+htmlDelta(delta)+"</div>"+
      "<div class='interp'>"+(ind.unidad||"")+"</div>"+
      "<div class='fecha'>"+ultFecha+"</div>";
    cont.appendChild(card);
    registrarSpark(card, ind.valores, colorDelta(delta));
    const panel=document.createElement("div"); panel.className="panel ancho";
    panel.innerHTML="<h3>"+ind.nombre+"  ·  "+(ind.unidad||"")+"</h3>"+
      "<div class='lienzo'><canvas id='c-constr-"+k+"'></canvas></div>";
    graf.appendChild(panel);
  });
}

function graficarConstruccion(){
  (DATA.construccion||[]).forEach((ind, k)=>{
    lineChart("c-constr-"+k, ind.fecha, [{data:ind.valores, borderColor:"#FF9D2E"}], {
      plugins:{ legend:{display:false},
        tooltip:{ callbacks:{ label:c=> fmtConstr(c.parsed.y, ind) } } }
    });
  });
}

function construirChatarra(){
  const cont=document.getElementById("cards-chatarra");
  const graf=document.getElementById("graf-chatarra");
  const lista = DATA.chatarra||[];
  if(!lista.length){
    if(cont) cont.innerHTML="<div class='card'><div class='nombre'>Chatarra</div>"+
      "<div class='interp'>Sin datos en esta corrida.</div></div>";
    return;
  }
  lista.forEach((ind, k)=>{
    const v = ultimosDosMacro(ind.valores);   // [último, previo] no nulos
    const ultVal = v.length ? v[0] : null;
    // son índices: la lectura útil es el % de cambio mensual
    let delta=null;
    if(v.length>=2 && v[1]!==0) delta={pp:false, val:(v[0]-v[1])/v[1]*100};
    const ultFecha = ind.fecha.length ? ind.fecha[ind.fecha.length-1] : "";
    const card=document.createElement("div"); card.className="card";
    card.innerHTML =
      "<div class='nombre'>"+ind.nombre+"</div>"+
      "<div class='valor'>"+fmtConstr(ultVal, ind)+"</div>"+
      "<div class='delta-linea'>"+htmlDelta(delta)+"</div>"+
      "<div class='interp'>"+(ind.unidad||"")+"</div>"+
      "<div class='fecha'>"+ultFecha+"</div>";
    cont.appendChild(card);
    registrarSpark(card, ind.valores, colorDelta(delta));
    const panel=document.createElement("div"); panel.className="panel ancho";
    panel.innerHTML="<h3>"+ind.nombre+"  ·  "+(ind.unidad||"")+"</h3>"+
      "<div class='lienzo'><canvas id='c-chat-"+k+"'></canvas></div>";
    graf.appendChild(panel);
  });
}

function graficarChatarra(){
  (DATA.chatarra||[]).forEach((ind, k)=>{
    lineChart("c-chat-"+k, ind.fecha, [{data:ind.valores, borderColor:"#FF9D2E"}], {
      plugins:{ legend:{display:false},
        tooltip:{ callbacks:{ label:c=>{
          let txt = fmtConstr(c.parsed.y, ind);
          // busca el dato previo no nulo para calcular el % de cambio mensual
          const v = ind.valores;
          let prev = null;
          for(let i=c.dataIndex-1; i>=0; i--){
            if(v[i]!==null && v[i]!==undefined){ prev = v[i]; break; }
          }
          if(prev!==null && prev!==0 && c.parsed.y!=null){
            const pct = (c.parsed.y - prev)/prev*100;
            txt += "  ·  " + (pct>=0?"+":"") + pct.toFixed(1) + "% vs mes previo";
          }
          return txt;
        } } } }
    });
  });
}

function etiquetaCorr(r){
  if(r===null || r===undefined) return "Sin dato suficiente";
  const ar=Math.abs(r);
  const fuerza = ar>=0.7 ? "fuerte" : (ar>=0.4 ? "moderada" : "baja");
  const signo = r>=0 ? "positiva" : "negativa";
  const sentido = r>=0 ? "se mueven en el mismo sentido" : "se mueven en sentido opuesto";
  return "Correlación "+signo+" "+fuerza+" · "+sentido;
}

function construirAcereras(){
  const a = DATA.acereras;
  const cont = document.getElementById("cards-acereras");
  const cob = document.getElementById("ac-cobertura");
  if(!a){
    cont.innerHTML = "<div class='card'><div class='nombre'>Acereras</div>"+
      "<div class='interp'>Sin datos en esta corrida.</div></div>";
    return;
  }
  const r = a.corr_global;
  const card=document.createElement("div"); card.className="card";
  card.innerHTML =
    "<div class='nombre'>Correlación mundial vs. EE. UU. (desde 2021)</div>"+
    "<div class='valor'>"+(r===null||r===undefined?"—":r.toFixed(2))+"</div>"+
    "<div class='interp'>"+etiquetaCorr(r)+"</div>";
  cont.appendChild(card);
  registrarSpark(card, a.corr||[], "#FF7A18");
  cob.innerHTML =
    "<b>Grupo mundial:</b> "+((a.ok_mundial||[]).join(", ")||"—")+
    "<br><b>Grupo EE. UU.:</b> "+((a.ok_eeuu||[]).join(", ")||"—")+
    ((a.fallaron&&a.fallaron.length)?"<br><span class='gris'>Sin datos (omitidas): "+a.fallaron.join(", ")+"</span>":"");

  // Tarjetas de precio por acción (estilo divisas): último cierre + % de cambio.
  const cp = document.getElementById("cards-acereras-precios");
  const pr = a.precios;
  if(cp){
    if(pr && pr.resumen && pr.resumen.length){
      pr.resumen.forEach(p=> cp.appendChild(tarjetaPrecio(p)));
    } else {
      cp.innerHTML = "<div class='card'><div class='nombre'>Precios</div>"+
        "<div class='interp'>Sin datos en esta corrida.</div></div>";
    }
  }
}

// tarjeta de una acción: nombre, último precio (USD) y % vs. día previo
function tarjetaPrecio(p){
  const c = (p.cambio===null || p.cambio===undefined) ? null : {pp:false, val:p.cambio};
  const el=document.createElement("div"); el.className="card";
  const etiq = (p.pais || p.grupo) ? "<div class='interp'>"+(p.pais || p.grupo)+"</div>" : "";
  el.innerHTML =
    "<div class='nombre'>"+logoAcerera(p.nombre)+p.nombre+"</div>"+
    "<div class='valor'>$"+(p.valor!=null? p.valor.toLocaleString("es-MX",{minimumFractionDigits:2, maximumFractionDigits:2}) : "—")+"</div>"+
    "<div class='delta-linea'>"+htmlDelta(c)+"</div>"+
    etiq+
    "<div class='fecha'>"+p.fecha+"</div>";
  try{ registrarSpark(el, DATA.acereras.precios.series[p.nombre]||[], colorDelta(c)); }catch(_e){}
  return el;
}

function graficarAcereras(){
  const a = DATA.acereras;
  if(!a) return;
  const idxDatasets = [
    {data:a.mundial, borderColor:"#FFC400", label:"Mundial"},
    {data:a.eeuu,    borderColor:"#B8BCC0", label:"EE. UU."}
  ];
  if(Array.isArray(a.indice_real) && a.indice_real.some(v=>v!=null)){
    idxDatasets.push({data:a.indice_real, borderColor:"#3b82f6",
                      label:(a.indice_real_nombre||"NYSE Arca Steel Index")});
  }
  lineChart("c-ac-idx", a.fecha, idxDatasets,
    {plugins:{legend:{display:true, labels:{boxWidth:12, font:{size:11}, color:"#cfcfcf"}}}});
  lineChart("c-ac-corr", a.corr_fecha, [{data:a.corr, borderColor:"#FF7A18"}], {
    plugins:{ legend:{display:false},
      tooltip:{ callbacks:{ label:c=> (c.parsed.y!=null? c.parsed.y.toFixed(2):"—") } } },
    scales:{ y:{ suggestedMin:-1, suggestedMax:1 } }
  });

  // Gráfica de precio por acción con selector (estilo divisas).
  const pr = a.precios;
  const sel = document.getElementById("sel-accion");
  if(sel && pr && pr.series && pr.resumen && pr.resumen.length){
    sel.innerHTML = "";
    pr.resumen.forEach(p=>{
      const o=document.createElement("option");
      o.value=p.nombre; o.textContent=p.nombre + ((p.pais||p.grupo)? "  ·  "+(p.pais||p.grupo) : "");
      sel.appendChild(o);
    });
    const inicial = pr.resumen[0].nombre;
    sel.value = inicial;
    const ch = lineChart("c-ac-precio", pr.fecha,
      [{data:pr.series[inicial], borderColor:"#FF7A18", label:inicial}],
      {plugins:{ tooltip:{ callbacks:{ label:c=> (c.parsed.y!=null? "$"+c.parsed.y.toFixed(2):"—") } } }});
    sel.addEventListener("change", ()=>{
      ch.data.datasets[0].data = pr.series[sel.value];
      ch.data.datasets[0].label = sel.value;
      ch.update();
    });
  }
}

let _prodMetrica = "ytd";   // "ytd" o "mes"
let _prodMesIdx = 0;        // índice dentro de meses[] (0 = más reciente)

// nombres de país en español (worldsteel los manda en inglés). Si alguno no está
// en el mapa, se queda como viene, así nunca aparece vacío.
const _PROD_ES = {
  "China":"China", "India":"India", "United States":"Estados Unidos",
  "Japan":"Japón", "South Korea":"Corea del Sur", "Russia":"Rusia",
  "Türkiye":"Turquía", "Turkey":"Turquía", "Germany":"Alemania",
  "Brazil":"Brasil", "Viet Nam":"Vietnam", "Vietnam":"Vietnam", "Iran":"Irán",
  "Italy":"Italia", "Taiwan (China)":"Taiwán", "Taiwan, China":"Taiwán",
  "Ukraine":"Ucrania", "Mexico":"México", "Spain":"España", "France":"Francia",
};
function _prodNombreES(p){ return _PROD_ES[p] || p; }

function _prodPeriodoBonito(p){
  if(!p) return "";
  const meses=["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio",
               "Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
  const [a,m]=p.split("-"); const i=parseInt(m,10)-1;
  return (meses[i]||p)+" "+a;
}

// lista de meses disponibles; si el JSON es viejo (sin 'meses') armo uno solo
function _prodMesesArr(){
  const pr = DATA.productores;
  if(!pr) return [];
  if(Array.isArray(pr.meses) && pr.meses.length) return pr.meses;
  if(Array.isArray(pr.paises) && pr.paises.length)
    return [{periodo:pr.periodo, fuente_url:pr.fuente_url, paises:pr.paises}];
  return [];
}
function _prodSnap(){
  const arr=_prodMesesArr();
  return arr[Math.min(_prodMesIdx, arr.length-1)] || null;
}

function construirProductores(){
  const pr = DATA.productores;
  const cont = document.getElementById("cards-productores");
  const arr = _prodMesesArr();
  if(!arr.length){
    if(cont) cont.innerHTML="<div class='card'><div class='nombre'>Productores</div>"+
      "<div class='interp'>Sin datos en esta corrida.</div></div>";
    return;
  }
  // lleno el selector de mes (más reciente primero)
  const sel = document.getElementById("prod-mes");
  if(sel){
    sel.innerHTML = "";
    arr.forEach((mm, i)=>{
      const o=document.createElement("option");
      o.value=i; o.textContent=_prodPeriodoBonito(mm.periodo);
      sel.appendChild(o);
    });
    sel.value = 0;
    sel.style.display = arr.length>1 ? "" : "none";   // si solo hay un mes, lo escondo
    sel.addEventListener("change", ()=>{ _prodMesIdx = parseInt(sel.value,10)||0; _prodActualizar(); });
  }
  // toggle YTD/Mes
  document.querySelectorAll(".prod-btn").forEach(btn=>{
    btn.classList.toggle("activo", btn.dataset.metrica===_prodMetrica);
    btn.addEventListener("click", ()=>{
      _prodMetrica = btn.dataset.metrica;
      document.querySelectorAll(".prod-btn").forEach(b=> b.classList.toggle("activo", b===btn));
      _prodActualizar();
    });
  });
  _prodActualizar();
}

// repinta subtítulo + tarjetas + gráfica según el mes y la métrica activos
function _prodActualizar(){
  const pr = DATA.productores;
  const snap = _prodSnap();
  if(!snap) return;
  const periodoTxt = _prodPeriodoBonito(snap.periodo);

  const spanP = document.getElementById("prod-periodo");
  if(spanP){
    spanP.innerHTML = "· "+periodoTxt+
      (pr && pr.respaldo ? " <span style='color:var(--gris);'>(respaldo)</span>" : "")+
      (snap.fuente_url ? "  ·  <a href='"+snap.fuente_url+"' target='_blank' rel='noopener' style='color:var(--gold);'>fuente</a>" : "");
  }

  // tarjetas: top 3 por acumulado del mes seleccionado
  const cont = document.getElementById("cards-productores");
  cont.innerHTML = "";
  [...snap.paises].sort((a,b)=>b.ytd_mt-a.ytd_mt).slice(0,3).forEach((p,k)=>{
    const card=document.createElement("div"); card.className="card";
    card.innerHTML =
      "<div class='nombre'>#"+(k+1)+" "+_prodNombreES(p.pais)+"</div>"+
      "<div class='valor'>"+p.ytd_mt.toLocaleString("es-MX")+" Mt</div>"+
      "<div class='delta-linea'>"+htmlDelta({pp:false, val:p.var_ytd_pct})+"</div>"+
      "<div class='interp'>acumulado del año</div>"+
      "<div class='fecha'>"+periodoTxt+"</div>";
    cont.appendChild(card);
  });

  graficarProductores();
}

function graficarProductores(){
  const snap = _prodSnap();
  if(!snap || !Array.isArray(snap.paises) || !snap.paises.length) return;
  const esYTD = _prodMetrica==="ytd";
  const campoMt  = esYTD ? "ytd_mt"      : "mes_mt";
  const campoVar = esYTD ? "var_ytd_pct" : "var_mes_pct";

  const filas = [...snap.paises].sort((a,b)=> b[campoMt]-a[campoMt]);
  const labels  = filas.map(p=>_prodNombreES(p.pais));
  const valores = filas.map(p=>p[campoMt]);
  const cambios = filas.map(p=>p[campoVar]);
  const colores = cambios.map(v=> v>=0 ? "#37d67a" : "#ff5b5b");

  const id="c-productores";
  if(charts[id]){
    charts[id].data.labels = labels;
    charts[id].data.datasets[0].data = valores;
    charts[id].data.datasets[0].backgroundColor = colores;
    charts[id]._cambios = cambios;
    charts[id].update();
    return;
  }
  const c = new Chart(document.getElementById(id), {
    type:"bar",
    data:{ labels, datasets:[{ data:valores, backgroundColor:colores, borderRadius:6, borderWidth:0 }] },
    options:{
      indexAxis:"y", responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:false},
        tooltip: Object.assign({}, TT_THEME, { callbacks:{ label:ctx=>{
          const v = charts[id]._cambios[ctx.dataIndex];
          const s = v>=0 ? "+" : "";
          return ctx.parsed.x.toLocaleString("es-MX")+" Mt  ·  "+s+v+"% vs año previo";
        } } }) },
      scales:{
        x:{ title:{display:true, text:"Millones de toneladas (Mt)", color:"#9aa0a6", font:{size:11}},
            ticks:{font:{size:11}, color:"#9aa0a6"}, grid:{color:"rgba(255,255,255,.06)"}, border:{color:"#3a3a42"} },
        y:{ ticks:{font:{size:12}, color:"#ECECEC"}, grid:{display:false}, border:{color:"#3a3a42"} }
      }
    }
  });
  c._cambios = cambios;
  charts[id]=c;
}

function construirGraficas(tab){
  const f=DATA.fechas, s=DATA.series;
  if(tab==="commodities"){
    if(s["Henry Hub ($/MMBtu)"]){
      const hh = s["Henry Hub ($/MMBtu)"];
      lineChart("c-hh", f, [
        {data:hh, borderColor:"#FFC400", label:"Precio"},
        {data:mediaMovil(hh,30), borderColor:"#B8BCC0", borderDash:[5,4], label:"Promedio 30 días"}
      ], {plugins:{legend:{display:true, labels:{boxWidth:12, font:{size:11}, color:"#cfcfcf"}}}});
    }
    lineChart("c-oil", f, [
      {data:s["WTI ($/bbl)"],   borderColor:"#FF7A18", label:"WTI"},
      {data:s["Brent ($/bbl)"], borderColor:"#E62315", label:"Brent"}
    ], {plugins:{legend:{display:true, labels:{boxWidth:12, font:{size:11}, color:"#cfcfcf"}}}});
    if(s["S&P 500"]) lineChart("c-sp", f, [{data:s["S&P 500"], borderColor:"#B8BCC0"}]);
  }
  if(tab==="macro"){
    graficarMacro();
  }
  if(tab==="construccion"){
    graficarConstruccion();
  }
  if(tab==="chatarra"){
    graficarChatarra();
  }
  if(tab==="productores"){
    graficarProductores();
  }
  if(tab==="acereras"){
    graficarAcereras();
  }
  if(tab==="divisas"){
    const sel=document.getElementById("sel-moneda");
    DATA.monedas.forEach(m=>{ const o=document.createElement("option"); o.value=m; o.textContent=m; sel.appendChild(o); });
    const inicial = DATA.monedas.indexOf("MXN por USD")>-1 ? "MXN por USD" : DATA.monedas[0];
    sel.value = inicial;
    const fx = lineChart("c-fx", f, [{data:s[inicial], borderColor:"#FF7A18"}]);
    sel.addEventListener("change", ()=>{ fx.data.datasets[0].data = s[sel.value]; fx.update(); });
  }
}

function mostrarTab(tab){
  document.querySelectorAll(".tab").forEach(b=> b.classList.toggle("activo", b.dataset.tab===tab));
  document.querySelectorAll("section").forEach(sec=> sec.classList.toggle("activa", sec.id==="sec-"+tab));
  requestAnimationFrame(dibujarSparks);
  if(tab==="swap" && !construido.swap){ iniciarSwap(); iniciarComparador(); construido.swap=true; }
  if(!construido[tab]){ construirGraficas(tab); construido[tab]=true; }
  else { Object.values(charts).forEach(c=> c.resize()); }
}

// ---- Comparación entre las dos simulaciones (fair value vs swap/collar) ----
const COMPARACION = { fair:null, swap:null, none:null, collar:null };

function actualizarComparacion(){
  const { fair, swap, none, collar } = COMPARACION;
  if(fair===null || swap===null || none===null || collar===null) return;
  const fmt = (n)=> "$"+n.toLocaleString(undefined,{maximumFractionDigits:0});
  document.getElementById("cmp-fair").textContent = fmt(fair);
  document.getElementById("cmp-swap").textContent = fmt(swap);
  document.getElementById("cmp-none").textContent = fmt(none);
  document.getElementById("cmp-collar").textContent = fmt(collar);

  const opciones = [
    { nombre:"Fair value WMC", valor:fair },
    { nombre:"Swap cotizado", valor:swap },
    { nombre:"Sin cobertura (al spot de referencia)", valor:none },
    { nombre:"Collar", valor:collar },
  ];
  opciones.sort((a,b)=> a.valor-b.valor);
  const mejor = opciones[0];
  const veredicto = document.getElementById("cmp-veredicto");
  veredicto.innerHTML = `<b>Más barata: ${mejor.nombre}</b> (${fmt(mejor.valor)}). Orden completo, de menor a mayor costo: ${opciones.map(o=>o.nombre+" "+fmt(o.valor)).join(" · ")}.`;
}

// ---- SWAP GAS: fair value calculator ----
const SWAP_MESES = ["Jun","Jul","Ago","Sep","Oct","Nov","Dic","Ene","Feb","Mar","Abr","May"];
const SWAP_FWD_DEFAULT = [2.65,2.70,2.72,2.75,2.90,3.15,3.30,3.10,2.95,2.80,2.70,2.68];

let SW_SELECTED_K = 1;
let SW_CURRENT_MESES = SWAP_MESES.slice();

function iniciarSwap(){
  const nInput = document.getElementById("sw-nmonths");
  const fwdReal = (DATA && Array.isArray(DATA.curva_forward_hh)) ? DATA.curva_forward_hh : null;
  const fuenteEl = document.getElementById("sw-fuente");
  if(fwdReal && fwdReal.length){
    nInput.value = fwdReal.length;
    fuenteEl.textContent = "— Henry Hub real (NYMEX vía Yahoo)";
    fuenteEl.style.color = "var(--sube)";
  } else {
    fuenteEl.textContent = "— valores de ejemplo, edítalos manualmente";
    fuenteEl.style.color = "var(--baja)";
  }
  const rebuildMonths = ()=>{
    const n = Math.max(1, Math.min(12, parseInt(nInput.value)||7));
    const cont = document.getElementById("swap-months");
    cont.innerHTML = "";
    SW_CURRENT_MESES = [];
    for(let i=0;i<n;i++){
      const usarReal = fwdReal && fwdReal[i];
      const etiqueta = usarReal ? fwdReal[i].mes : SWAP_MESES[i];
      const valor = usarReal ? fwdReal[i].precio : SWAP_FWD_DEFAULT[i];
      SW_CURRENT_MESES.push(etiqueta);
      const div = document.createElement("div");
      div.innerHTML = `<label style="display:block; font-size:11px; color:var(--gris); margin-bottom:4px; text-transform:uppercase;">${etiqueta}</label>
        <input type="number" step="0.001" class="sw-fwd prod-sel" style="width:100%;" value="${valor}">`;
      cont.appendChild(div);
    }
    cont.querySelectorAll(".sw-fwd").forEach(el=> el.addEventListener("input", calcularSwap));
    calcularSwap();
  };
  nInput.addEventListener("change", rebuildMonths);
  ["sw-beta","sw-int","sw-vol","sw-margin","sw-price"].forEach(id=>{
    document.getElementById(id).addEventListener("input", calcularSwap);
  });

  const seeEl = document.getElementById("sw-see");
  const marginEl = document.getElementById("sw-margin");
  const applyTargetMarginFromK = ()=>{
    const see = parseFloat(seeEl.value)||0;
    marginEl.value = (SW_SELECTED_K*see).toFixed(3);
    calcularSwap();
  };
  document.querySelectorAll(".sw-kbtn").forEach(btn=>{
    btn.addEventListener("click", ()=>{
      document.querySelectorAll(".sw-kbtn").forEach(b=>b.classList.remove("activo"));
      btn.classList.add("activo");
      SW_SELECTED_K = parseFloat(btn.dataset.k);
      applyTargetMarginFromK();
    });
  });
  seeEl.addEventListener("input", applyTargetMarginFromK);
  applyTargetMarginFromK();

  rebuildMonths();
}

function calcularSwap(){
  const beta = parseFloat(document.getElementById("sw-beta").value)||0;
  const intercept = parseFloat(document.getElementById("sw-int").value)||0;
  const volume = parseFloat(document.getElementById("sw-vol").value)||0;
  const swapPrice = parseFloat(document.getElementById("sw-price").value)||0;
  const targetMargin = parseFloat(document.getElementById("sw-margin").value)||0;
  const fwds = Array.from(document.querySelectorAll(".sw-fwd")).map(el=> parseFloat(el.value)||0);

  let totalVol=0, totalFairCost=0, totalSwapCost=0, weightedFairSum=0, rows="";
  fwds.forEach((hsc,i)=>{
    const fairWMC = beta*hsc + intercept;
    const cost = fairWMC*volume;
    const swapCost = swapPrice*volume;
    totalVol+=volume; totalFairCost+=cost; totalSwapCost+=swapCost; weightedFairSum+=fairWMC*volume;
    rows += `<tr style="border-bottom:1px solid var(--linea);">
      <td style="padding:6px 8px;">${SW_CURRENT_MESES[i] || SWAP_MESES[i]}</td>
      <td style="padding:6px 8px; text-align:right;">$${hsc.toFixed(3)}</td>
      <td style="padding:6px 8px; text-align:right;">$${fairWMC.toFixed(3)}</td>
      <td style="padding:6px 8px; text-align:right;">$${swapPrice.toFixed(3)}</td>
      <td style="padding:6px 8px; text-align:right;">${volume.toLocaleString()}</td>
      <td style="padding:6px 8px; text-align:right;">$${cost.toLocaleString(undefined,{maximumFractionDigits:0})}</td>
      <td style="padding:6px 8px; text-align:right;">$${swapCost.toLocaleString(undefined,{maximumFractionDigits:0})}</td>
    </tr>`;
  });
  document.getElementById("sw-tbody").innerHTML = rows;

  const fairAvg = totalVol>0 ? weightedFairSum/totalVol : 0;
  const impliedMargin = swapPrice - fairAvg;
  const marginDelta = impliedMargin - targetMargin;

  document.getElementById("sw-fair").textContent = "$"+fairAvg.toFixed(3);
  const impliedEl = document.getElementById("sw-implied");
  impliedEl.textContent = (impliedMargin>=0?"+":"")+"$"+impliedMargin.toFixed(3);
  impliedEl.style.color = impliedMargin>0 ? "var(--baja)" : "var(--sube)";

  const deltaEl = document.getElementById("sw-delta");
  deltaEl.textContent = (marginDelta>=0?"+":"")+"$"+marginDelta.toFixed(3);
  deltaEl.style.color = marginDelta>0 ? "var(--baja)" : "var(--sube)";

  const verdict = document.getElementById("sw-verdict");
  if(impliedMargin <= targetMargin){
    verdict.innerHTML = `<span style="color:var(--sube); font-weight:700;">✓ Favorable —</span> el swap cotizado a $${swapPrice.toFixed(3)} queda dentro de tu margen objetivo sobre el fair value.`;
  } else {
    verdict.innerHTML = `<span style="color:var(--baja); font-weight:700;">⚠ Revisar —</span> el swap cotizado a $${swapPrice.toFixed(3)} implica un margen de $${impliedMargin.toFixed(3)} sobre el fair value, arriba de tu objetivo de $${targetMargin.toFixed(3)}. Hay espacio para negociar.`;
  }

  COMPARACION.fair = totalFairCost;
  COMPARACION.swap = totalSwapCost;
  actualizarComparacion();

  if(!b76FwdManual){
    const avg = promedioCurvaHSC();
    const fwdInput = document.getElementById("b76-fwd");
    if(avg!=null && fwdInput){ fwdInput.value = avg.toFixed(3); calcularBlack76(); }
  }
}

function mostrarAviso(msg){ const a=document.getElementById("aviso"); a.style.display="block"; a.innerHTML=msg; }

// ---- SWAP vs COLLAR: comparador de escenarios ----
let cpModo = "swap";

function cpFmtMM(n){
  const abs = Math.abs(n);
  const signo = n<0 ? "-" : "";
  if(abs >= 1e6) return signo+"$"+(abs/1e6).toFixed(2)+"M";
  if(abs >= 1e3) return signo+"$"+(abs/1e3).toFixed(1)+"K";
  return signo+"$"+abs.toFixed(0);
}

function cpCosto(hsc, swapPrice, floor, cap, modo){
  if(modo==="none") return hsc;
  if(modo==="swap") return swapPrice;
  if(modo==="collar") return Math.min(Math.max(hsc, floor), cap);
  return hsc;
}

function iniciarComparador(){
  document.querySelectorAll("#cp-toggle .prod-btn").forEach(b=>{
    b.classList.toggle("activo", b.dataset.modo===cpModo);
    b.addEventListener("click", ()=>{
      cpModo = b.dataset.modo;
      document.querySelectorAll("#cp-toggle .prod-btn").forEach(x=> x.classList.toggle("activo", x.dataset.modo===cpModo));
      calcularComparador();
    });
  });
  ["cp-vol","cp-nmeses","cp-spot","cp-swap","cp-floor","cp-cap"].forEach(id=>{
    document.getElementById(id).addEventListener("input", calcularComparador);
  });
  ["cp-rmin","cp-rmax"].forEach(id=>{
    document.getElementById(id).addEventListener("input", ()=>{
      document.getElementById("cp-rminlbl").textContent = parseFloat(document.getElementById("cp-rmin").value).toFixed(2);
      document.getElementById("cp-rmaxlbl").textContent = parseFloat(document.getElementById("cp-rmax").value).toFixed(2);
      calcularComparador();
    });
  });
  calcularComparador();
}

function calcularComparador(){
  const vol = parseFloat(document.getElementById("cp-vol").value)||0;
  const nmeses = parseFloat(document.getElementById("cp-nmeses").value)||1;
  const spot = parseFloat(document.getElementById("cp-spot").value)||0;
  const swapPrice = parseFloat(document.getElementById("cp-swap").value)||0;
  const floor = parseFloat(document.getElementById("cp-floor").value)||0;
  const cap = parseFloat(document.getElementById("cp-cap").value)||0;
  const rMin = parseFloat(document.getElementById("cp-rmin").value)||1;
  const rMax = parseFloat(document.getElementById("cp-rmax").value)||5;
  const totalVol = vol*nmeses;

  const N = 60, xs=[];
  for(let i=0;i<=N;i++){ xs.push(rMin + (rMax-rMin)*i/N); }
  const noHedgePL = xs.map(h=> -(h-spot)*totalVol);
  const hedgePL = xs.map(h=> -(cpCosto(h, swapPrice, floor, cap, cpModo) - spot)*totalVol);

  let breakevenTxt;
  if(cpModo==="swap") breakevenTxt = "$"+swapPrice.toFixed(3);
  else if(cpModo==="collar") breakevenTxt = `$${floor.toFixed(2)}–$${cap.toFixed(2)}`;
  else breakevenTxt = "$"+spot.toFixed(3);

  const worst = Math.min(...hedgePL);
  const ceilingPrice = cpModo==="swap" ? swapPrice : (cpModo==="collar" ? cap : rMax);

  document.getElementById("cp-maxloss").textContent = cpModo==="none" ? "sin límite" : cpFmtMM(worst);
  document.getElementById("cp-maxloss").style.color = cpModo==="none" ? "var(--baja)" : "var(--tinta)";
  document.getElementById("cp-breakeven").textContent = breakevenTxt;
  document.getElementById("cp-ceiling").textContent = "$"+ceilingPrice.toFixed(3);
  document.getElementById("cp-premium").textContent = "$0";

  const datasets = [
    { label:"Sin cobertura", data: xs.map((x,i)=>({x,y:noHedgePL[i]})), borderColor:"#B8BCC0", borderDash:[6,4], borderWidth:2, pointRadius:0 }
  ];
  if(cpModo!=="none"){
    datasets.push({
      label: cpModo==="swap" ? "Swap" : "Collar",
      data: xs.map((x,i)=>({x,y:hedgePL[i]})),
      borderColor:"#FFC400", borderWidth:2.5, pointRadius:0
    });
  }

  if(charts["c-swapcollar"]) charts["c-swapcollar"].destroy();
  charts["c-swapcollar"] = new Chart(document.getElementById("c-swapcollar"), {
    type:"line",
    data:{ datasets },
    options:{
      responsive:true, maintainAspectRatio:false,
      interaction:{ mode:"index", intersect:false },
      plugins:{
        legend:{ display:true, labels:{ boxWidth:12, font:{size:11}, color:"#cfcfcf" } },
        tooltip: Object.assign({}, TT_THEME, { callbacks:{
          title:(items)=> "HSC: $"+items[0].parsed.x.toFixed(2),
          label:(item)=> item.dataset.label+": "+cpFmtMM(item.parsed.y)
        }})
      },
      scales:{
        x:{ type:"linear", min:rMin, max:rMax,
            title:{ display:true, text:"HSC al vencimiento ($/MMBtu)", color:"#9aa0a6" },
            ticks:{ color:"#9aa0a6", callback:(v)=>"$"+v.toFixed(1) },
            grid:{ color:"rgba(255,255,255,0.06)" } },
        y:{ title:{ display:true, text:"Costo vs. spot de referencia (USD)", color:"#9aa0a6" },
            ticks:{ color:"#9aa0a6", callback:(v)=>cpFmtMM(v) },
            grid:{ color:"rgba(255,255,255,0.06)" } }
      }
    }
  });

  const escenarios = [
    { nombre:"Muy desfavorable", hsc: Math.min(rMax, spot*1.4) },
    { nombre:"Desfavorable", hsc: Math.min(rMax, spot*1.2) },
    { nombre:"Moderadamente bajo", hsc: Math.max(rMin, spot*0.9) },
    { nombre:"Precio actual", hsc: spot, hl:true },
    { nombre:"Favorable (baja fuerte)", hsc: Math.max(rMin, spot*0.7) },
  ];
  let rows="";
  const csvRows = [["Escenario","HSC","Costo sin cobertura","Costo con cobertura","Diferencia"]];
  escenarios.forEach(e=>{
    const costoNo = (e.hsc-spot)*totalVol;
    const costoHedgeTotal = (cpCosto(e.hsc, swapPrice, floor, cap, cpModo)-spot)*totalVol;
    const diff = costoNo - costoHedgeTotal;
    const cNo = costoNo>0?"delta down":costoNo<0?"delta up":"delta flat";
    const cHedge = cpModo==="none" ? "" : (costoHedgeTotal>0?"delta down":costoHedgeTotal<0?"delta up":"delta flat");
    const cDiff = cpModo==="none" ? "" : (diff>0?"delta up":diff<0?"delta down":"delta flat");
    rows += `<tr style="border-bottom:1px solid var(--linea); ${e.hl?'background:var(--tarjeta2);':''}">
      <td style="padding:8px; ${e.hl?'font-weight:700;':''}">${e.hl?'★ ':''}${e.nombre}</td>
      <td style="padding:8px; text-align:right;">$${e.hsc.toFixed(2)}</td>
      <td style="padding:8px; text-align:right;" class="${cNo}">${costoNo===0?'$0':(costoNo>0?'+':'')+cpFmtMM(costoNo)}</td>
      <td style="padding:8px; text-align:right;" class="${cHedge}">${cpModo==="none"?"—":(costoHedgeTotal===0?'$0':(costoHedgeTotal>0?'+':'')+cpFmtMM(costoHedgeTotal))}</td>
      <td style="padding:8px; text-align:right;" class="${cDiff}">${cpModo==="none"?"—":(diff===0?'$0':(diff>0?'+':'')+cpFmtMM(diff))}</td>
    </tr>`;
    csvRows.push([e.nombre, e.hsc.toFixed(3), costoNo.toFixed(2), cpModo==="none"?"":costoHedgeTotal.toFixed(2), cpModo==="none"?"":diff.toFixed(2)]);
  });
  document.getElementById("cp-tbody").innerHTML = rows;
  window.ESCENARIOS_CSV = csvRows;

  COMPARACION.none = spot*totalVol;
  COMPARACION.collar = cpCosto(spot, swapPrice, floor, cap, "collar")*totalVol;
  actualizarComparacion();
}

// ---- Candado de acceso (Coberturas + descarga de Excel) ----
const COB_HASH = "772b17f5d551de57fcdfe7e3ec743aa6c3351081a65158729d6333e91fa74b0a";
let cobAuth = sessionStorage.getItem("cob_auth") === "1";

async function sha256Hex(str){
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(str));
  return Array.from(new Uint8Array(buf)).map(b=> b.toString(16).padStart(2,"0")).join("");
}

async function pedirPassword(){
  if(cobAuth) return true;
  const intento = prompt("Esta sección requiere contraseña:");
  if(intento===null || intento==="") return false;
  const hash = await sha256Hex(intento);
  if(hash===COB_HASH){
    cobAuth = true;
    sessionStorage.setItem("cob_auth","1");
    return true;
  }
  alert("Contraseña incorrecta.");
  return false;
}

// ---- Descarga CSV de Escenarios ----
function descargarEscenariosCSV(){
  const rows = window.ESCENARIOS_CSV;
  if(!rows) return;
  const csv = rows.map(r=> r.map(c=> `"${String(c).replace(/"/g,'""')}"`).join(",")).join("\n");
  const blob = new Blob([csv], {type:"text/csv;charset=utf-8;"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = "escenarios_cobertura.csv";
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ---- Black-76: valuación de opciones sobre futuros ----
function normCDF(x){
  const t = 1/(1+0.2316419*Math.abs(x));
  const d = 0.3989423*Math.exp(-x*x/2);
  let p = d*t*(0.3193815+t*(-0.3565638+t*(1.781478+t*(-1.821256+t*1.330274))));
  if(x>0) p = 1-p;
  return p;
}

function black76(fwd, strike, volPct, dias, rPct){
  const sigma = volPct/100, T = dias/365, r = rPct/100;
  if(fwd<=0 || strike<=0 || sigma<=0 || T<=0) return {call:0, put:0};
  const d1 = (Math.log(fwd/strike) + (sigma*sigma/2)*T) / (sigma*Math.sqrt(T));
  const d2 = d1 - sigma*Math.sqrt(T);
  const disc = Math.exp(-r*T);
  const call = disc*(fwd*normCDF(d1) - strike*normCDF(d2));
  const put = disc*(strike*normCDF(-d2) - fwd*normCDF(-d1));
  return {call, put};
}

// se pone en true en cuanto el usuario edita el forward de Black-76 a mano;
// mientras siga en false, lo mantenemos sincronizado con el promedio de la curva HSC.
let b76FwdManual = false;

function promedioCurvaHSC(){
  const inputs = document.querySelectorAll(".sw-fwd");
  if(inputs.length){
    const vals = Array.from(inputs).map(el=>parseFloat(el.value)||0).filter(v=>v>0);
    if(vals.length) return vals.reduce((a,b)=>a+b,0)/vals.length;
  }
  // si aún no se ha construido la pestaña de Coberturas, usamos la misma fuente que usaría iniciarSwap()
  const nMonths = parseInt(document.getElementById("sw-nmonths")?.value)||7;
  const fwdReal = (DATA && Array.isArray(DATA.curva_forward_hh)) ? DATA.curva_forward_hh : null;
  const arr = (fwdReal && fwdReal.length) ? fwdReal.slice(0,nMonths).map(x=>x.precio) : SWAP_FWD_DEFAULT.slice(0,nMonths);
  if(!arr.length) return null;
  return arr.reduce((a,b)=>a+b,0)/arr.length;
}

function calcularBlack76(){
  const fwd = parseFloat(document.getElementById("b76-fwd").value)||0;
  const vol = parseFloat(document.getElementById("b76-vol").value)||0;
  const kPut = parseFloat(document.getElementById("b76-kput").value)||0;
  const kCall = parseFloat(document.getElementById("b76-kcall").value)||0;
  const dias = parseFloat(document.getElementById("b76-dias").value)||0;
  const r = parseFloat(document.getElementById("b76-r").value)||0;

  const put = black76(fwd, kPut, vol, dias, r).put;
  const call = black76(fwd, kCall, vol, dias, r).call;
  const diff = put - call;

  document.getElementById("b76-pput").textContent = "$"+put.toFixed(4);
  document.getElementById("b76-pcall").textContent = "$"+call.toFixed(4);
  const diffEl = document.getElementById("b76-diff");
  diffEl.textContent = (diff>=0?"+":"")+"$"+diff.toFixed(4);
  diffEl.style.color = Math.abs(diff)<0.01 ? "var(--tinta)" : (diff>0 ? "var(--baja)" : "var(--sube)");

  const veredicto = document.getElementById("b76-veredicto");
  if(Math.abs(diff) < 0.01){
    veredicto.innerHTML = `<span style="color:var(--sube); font-weight:700;">✓ Aprox. costo cero —</span> el put y el call teóricos están prácticamente balanceados con esta volatilidad.`;
  } else if(diff>0){
    veredicto.innerHTML = `<span style="color:var(--baja); font-weight:700;">⚠ Put &gt; Call —</span> el piso teóricamente vale más que el techo; para que el collar sea costo cero, el bróker debería pagarte una prima neta de $${diff.toFixed(4)}/MMBtu, o tendrías que subir el piso / bajar el techo.`;
  } else {
    veredicto.innerHTML = `<span style="color:var(--baja); font-weight:700;">⚠ Call &gt; Put —</span> el techo teóricamente vale más que el piso; tendrías que pagar una prima neta de $${(-diff).toFixed(4)}/MMBtu, o bajar el piso / subir el techo para llegar a costo cero.`;
  }
}

function iniciarBlack76(){
  const fwdInput = document.getElementById("b76-fwd");
  const avg = promedioCurvaHSC();
  if(avg!=null) fwdInput.value = avg.toFixed(3);
  fwdInput.addEventListener("input", ()=>{ b76FwdManual = true; });
  ["b76-fwd","b76-vol","b76-kput","b76-kcall","b76-dias","b76-r"].forEach(id=>{
    document.getElementById(id).addEventListener("input", calcularBlack76);
  });
  calcularBlack76();
}

window.PANEL_INIT = function(){
  registrarZoom();
fetch(DATA_URL)
  .then(r=>{ if(!r.ok) throw new Error("HTTP "+r.status); return r.json(); })
  .then(d=>render(d))
  .catch(e=>{
    document.getElementById("meta").textContent="No se pudieron cargar los datos.";
    mostrarAviso("<b>No se encontró <code>"+DATA_URL+"</code>.</b><br>"+
      "Esto funciona desde el enlace de GitHub Pages (no con doble clic local, por las "+
      "restricciones del navegador). Verifica también que el workflow ya haya generado el archivo. ("+e.message+")");
  });
};

function haceCuanto(antes, ahora){
  let seg = Math.round((ahora - antes)/1000);
  if(seg < 60) return "hace un momento";
  let min = Math.round(seg/60);
  if(min < 60) return "hace " + min + " min";
  let h = Math.floor(min/60);
  if(h < 24) return "hace " + h + (h===1?" hora":" horas");
  let dias = Math.floor(h/24);
  return "hace " + dias + (dias===1?" día":" días");
}

function render(d){
  DATA = d;
  const gen = d.generado_iso ? new Date(d.generado_iso) : null;

  function pintarMeta(){
    const ahora = new Date();
    let s;
    if(gen && !isNaN(gen)){
      const local = gen.toLocaleString("es-MX", {dateStyle:"medium", timeStyle:"short"});
      s = "Actualizado: " + local + " (tu hora) · " + haceCuanto(gen, ahora);
    } else {
      s = "Actualizado: " + d.generado;   // respaldo si el JSON aún es viejo
    }
    s += "  ·  Cambio vs. dato previo";
    document.getElementById("meta").textContent = s;
    document.getElementById("reloj").textContent =
      "Tu hora local: " + ahora.toLocaleString("es-MX", {dateStyle:"medium", timeStyle:"medium"});
  }
  pintarMeta();
  setInterval(pintarMeta, 1000);

  document.getElementById("footer").innerHTML =
    "Fuentes — " + d.fuentes + ".<br>Generado automáticamente el " + d.generado +
    ". El % junto a cada cifra compara el último dato con el anterior disponible; "+
    "las fechas pueden diferir según el rezago de cada fuente.";

  // repartir tarjetas por sección (resumen = commodities + divisas)
  const cDiv=document.getElementById("cards-divisas");
  const cCom=document.getElementById("cards-commodities");
  d.resumen.forEach(r=>{
    if(esFX(r.serie)) cDiv.appendChild(tarjeta(r));
    else cCom.appendChild(tarjeta(r));
  });
  construirMacro();   // tarjetas + paneles de Macro EE. UU.
  construirConstruccion(); // tarjetas + paneles de Construcción
  construirChatarra();     // tarjetas + paneles de Chatarra
  construirProductores();  // tarjetas + toggle de Productores (worldsteel)
  construirAcereras();// tarjeta de correlación + cobertura

  // pestañas (Coberturas pide contraseña la primera vez por sesión)
  document.querySelectorAll(".tab").forEach(b=> b.addEventListener("click", async ()=>{
    if(b.dataset.tab==="swap" && !cobAuth){
      const ok = await pedirPassword();
      if(!ok) return;
    }
    mostrarTab(b.dataset.tab);
  }));
  mostrarTab("divisas");   // pestaña inicial

  // descarga de Excel (misma contraseña que Coberturas)
  const dlExcel = document.getElementById("dl-excel");
  if(dlExcel){
    dlExcel.addEventListener("click", async (e)=>{
      if(!cobAuth){
        e.preventDefault();
        const ok = await pedirPassword();
        if(ok) window.location.href = dlExcel.getAttribute("href");
      }
    });
  }

  // botón de descarga de Escenarios
  const cpDl = document.getElementById("cp-download");
  if(cpDl) cpDl.addEventListener("click", descargarEscenariosCSV);

  // calculadora Black-76
  iniciarBlack76();
}

})();
}
