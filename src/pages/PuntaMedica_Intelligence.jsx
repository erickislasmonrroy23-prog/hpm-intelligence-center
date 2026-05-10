import { useState, useMemo, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, AreaChart, Area,
  PieChart, Pie, Cell, ComposedChart,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  ResponsiveContainer, ReferenceLine
} from "recharts";

// ─── BRAND PUNTA MÉDICA ─────────────────────────────────────────────
const B = {
  // Primarios
  cyan:     "#00B4D8",
  cyanD:    "#006D8F",
  cyanDD:   "#004D6E",
  cyanL:    "#E0F7FA",
  cyanLL:   "#F0FBFD",
  // Secundarios
  teal:     "#0097A7",
  navy:     "#003D52",
  // Semáforos
  verde:    "#00897B",
  verdeL:   "#E0F2F1",
  amber:    "#F59E0B",
  amberL:   "#FEF3C7",
  rojo:     "#E53935",
  rojoL:    "#FFEBEE",
  violeta:  "#5C6BC0",
  violetaL: "#E8EAF6",
  // Neutros
  gris:     "#37474F",
  grisM:    "#78909C",
  grisL:    "#F8FAFB",
  borde:    "#E0E8EB",
  blanco:   "#FFFFFF",
  // Charts
  chart:    ["#00B4D8","#0097A7","#006D8F","#00BCD4","#26C6DA","#4DD0E1","#80DEEA","#B2EBF2","#00897B","#26A69A"],
};

// ─── SVG LOGO + ECG ─────────────────────────────────────────────────
const LOGO_SRC = "data:image/jpeg;base64,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1s/iLaxGgnTSimpzpMfFNJN0p0ppWkm9060v0ikwaIv3TU0pMvOLbjHF9jh3DDumBgN6d8mlJN2xbqTSk+c4v1S/VsaTBvjQIuAAAAAAAAAAAAAAAAAAAAAAAAAAAAAy6visnRk+b8Gmll1ZdgV1AATNnK+HCnj6mSlNpQZNeSLdO/D0nF869L4amXVn0plOXfLez/YfPvtMds3MjzSk9Zrrn9+qfGjF2idM/XlKQp0ccdKbfCjN0k56WDVDFSm0hNpTRSmUz5NpzTXh90Jrm64tqeOzHsrfJ69c5rrFdwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMOC5x05ensz6QSABNo+Ok5fJ1LmT9Pb5anH539xbnloeE/fZXSZT8e5ol1OSOH3qmuT1t8y4cNnWE6lz6Rdw7/ItO79V8efHX1i2JoIy+KXMwU+fSL+MNHmidp7/LU48NvpM/Ro8oytv2Jxb/AJ9jTH5285ph79vs1w0fSugRoAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAB//aAAwDAQACAAMAAAAh888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888N8888888880888888888888888888888888888888888888888888888888888uz38888888jd8888888888888888888888888lO6r+/889i388qCyw5yeU28888uL888018oBzCWvOez1i+xzyU98v5888888888V88Nd+888ELS8rV89/88pW8888e2987BY8rc88c8csbOg5yP8Kf8kE188888888Szz50+888V+H+rV89/8AI3kNfPPMPOPPrjfK/P8Arz3zzwQCn/zzzz3T37zzzzzzzwXz0zr7zzhz396hXz3/ANc8ub887S8ZQl8p8r98u898880Ap38888pc8R98888888V888tgx0rU98fgU89/7P8ALWPPO/8Ay93by0ivDHHXzXXPmi3rvmpBnzx9LzzzzzzxLzzzywaN/wB0889U8+1X9887283U8tC888KpuOObeO3O8J0dS21a8889r888888888888888888888888888888N8k8888888888888888888888888888888888888888888888888888888888881i388862129584/wCcN8f9s+/t9ffutfvuPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPLrdPPPEIF6/RnvO0J61Ep3UbDl63uv6f8ALzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzzV3zzy8+2rSyL+/Af6q0dKe7/AEo16U9hG8888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888888//2gAMAwEAAgADAAAAEPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPMPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPOGdPPPPPPPOPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPHHPPPPPPPAfPPPPPPPPPPPPPPPPPPPPPPPPOAAsum4PPFnZ9PEict+MiJ0PPPPPU2fOD1fOPX3uBjnhmnGKMChOvFidPPPPPPPLFDDPoFfPBOFrdCrzFx3Ah4/PPOP62dIDsfKpjDLBZHH7nr2fTLQOG7HvPPPPPPPFTQBK1fPFKebH4//F1/AocnfPKFvNVG6j9KggRVFFPLFfrb/PPDLhVJNPPPPPPKBv0wwGfPEKKyJgP/ABde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";
const LogoPM = ({size=32, dark=false}) => (
  <img src={LOGO_SRC} alt="Hospital Punta Médica" style={{height: size, width: "auto", filter: dark ? "brightness(0) saturate(100%) invert(14%) sepia(48%) saturate(1200%) hue-rotate(177deg) brightness(95%) contrast(97%)" : "none", maxWidth: "100%"}}/>
);

const ECGLine = ({width=200,height=30,color=B.cyan,animated=false}) => (
  <svg width={width} height={height} viewBox={`0 0 ${width} ${height}`}>
    <path d={`M0 ${height/2} L${width*0.2} ${height/2} L${width*0.28} ${height/2} L${width*0.32} 4 L${width*0.38} ${height-4} L${width*0.42} ${height/2-8} L${width*0.46} ${height/2} L${width*0.7} ${height/2} L${width*0.78} ${height/2} L${width*0.82} 4 L${width*0.88} ${height-4} L${width*0.92} ${height/2-8} L${width*0.96} ${height/2} L${width} ${height/2}`}
      stroke={color} strokeWidth="1.5" fill="none" strokeLinecap="round" opacity="0.5"/>
  </svg>
);

// ─── UTILS ──────────────────────────────────────────────────────────
const f$ = n => !n&&n!==0?"—":"$"+Math.round(n).toLocaleString("es-MX");
const fM = n => !n&&n!==0?"—":"$"+(n/1e6).toFixed(2)+"M";
const fP = n => !n&&n!==0?"—":(n*100).toFixed(1)+"%";
const fN = n => !n&&n!==0?"—":Math.round(n).toLocaleString("es-MX");
const fDelta = (a,b) => b===0?"—":`${a>=b?"+":""}${((a/b-1)*100).toFixed(1)}%`;
const dC = v => { const n=parseFloat(v); return isNaN(n)?B.grisM:n>=0?B.verde:B.rojo; };

function detectTipo(headers){
  const h=headers.map(x=>String(x||"").toLowerCase());
  if(h.some(x=>x.includes("cirug")||x.includes("quirof")||x.includes("cirujano"))) return "cirugias";
  if(h.some(x=>x.includes("existencia")||x.includes("stock")||x.includes("inventario"))) return "inventario";
  if(h.some(x=>x.includes("costo")||x.includes("cost"))) return "cargos_costo";
  if(h.some(x=>x.includes("cargo")||x.includes("importe")||x.includes("monto"))) return "cargos";
  if(h.some(x=>x.includes("abono")||x.includes("devolucion"))) return "abonos";
  return "generico";
}
function autoMap(headers){
  const find=kws=>{const h=headers.map(x=>String(x||"").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,""));for(const k of kws){const i=h.findIndex(x=>x.includes(k));if(i>=0)return headers[i];}return null;};
  return {
    fecha:    find(["fecha","date","periodo","mes","fec","dia"]),
    medico:   find(["medico","doctor","cirujano","med_","nombre_med"]),
    paciente: find(["nombre_pac","paciente","patient","nombre_p","pac_"]),
    servicio: find(["servicio","ubicacion","depto","area","tipo","departamento"]),
    cargo:    find(["cargo","importe","monto","total_cargo","precio","charge","amount"]),
    costo:    find(["costo","cost","coste","precio_costo"]),
    convenio: find(["convenio","seguro","insurance","payer","aseguradora","pagador"]),
    folio:    find(["folio","ecp","id_ecp","num_exp","expediente","id_ti"]),
    cantidad: find(["cantidad","qty","piezas","unidades","existencia"]),
  };
}
const parseDate=v=>{if(!v)return null;const s=String(v);if(/^\d{5}$/.test(s)){const d=new Date(Date.UTC(1899,11,30)+parseInt(s)*86400000);return d.toISOString().substring(0,10);}if(/^\d{4}-\d{2}-\d{2}/.test(s))return s.substring(0,10);if(/^\d{2}\/\d{2}\/\d{4}/.test(s)){const[d,m,y]=s.split("/");return`${y}-${m}-${d}`;}return s.substring(0,10);};
const getMes=f=>f?String(f).substring(0,7):"Sin fecha";
const getQ=f=>{if(!f)return"S/D";const[y,m]=String(f).split("-");return`${y}-Q${Math.ceil(parseInt(m)/3)}`;};

// ─── COMPONENTES BRAND ───────────────────────────────────────────────
const Card = ({children,style={}}) => (
  <div style={{background:B.blanco,border:`1px solid ${B.borde}`,borderRadius:14,padding:"18px 22px",...style}}>{children}</div>
);

const KpiCard = ({label,value,sub,delta,color=B.cyan,badge,badgeOk=true,icon}) => (
  <div style={{background:B.blanco,border:`1px solid ${B.borde}`,borderRadius:14,padding:"18px 22px",borderTop:`3px solid ${color}`,transition:"box-shadow 0.2s"}} onMouseEnter={e=>e.currentTarget.style.boxShadow="0 6px 20px rgba(0,180,216,0.1)"} onMouseLeave={e=>e.currentTarget.style.boxShadow="none"}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
      <div style={{fontSize:10,textTransform:"uppercase",letterSpacing:"0.1em",color:B.grisM,fontWeight:600}}>{label}</div>
      {icon && <span style={{fontSize:18,opacity:0.6}}>{icon}</span>}
    </div>
    <div style={{fontFamily:"Georgia,serif",fontSize:30,fontWeight:700,color:B.navy,lineHeight:1}}>{value}</div>
    <div style={{display:"flex",alignItems:"center",gap:8,marginTop:8}}>
      {sub && <span style={{fontSize:11,color:B.grisM}}>{sub}</span>}
      {delta && <span style={{fontSize:11,fontWeight:700,color:dC(delta),background:parseFloat(delta)>=0?B.verdeL:B.rojoL,padding:"1px 7px",borderRadius:10}}>{delta}</span>}
    </div>
    {badge && <span style={{display:"inline-block",marginTop:8,padding:"3px 10px",borderRadius:20,fontSize:10,fontWeight:600,background:badgeOk?B.cyanL:B.rojoL,color:badgeOk?B.cyanD:B.rojo}}>{badge}</span>}
  </div>
);

const Tip = ({active,payload,label}) => {
  if(!active||!payload?.length) return null;
  return <div style={{background:B.blanco,border:`1px solid ${B.borde}`,borderRadius:10,padding:"10px 14px",fontSize:12,boxShadow:"0 8px 24px rgba(0,180,216,0.15)"}}>
    <div style={{fontWeight:700,marginBottom:6,color:B.navy,borderBottom:`1px solid ${B.borde}`,paddingBottom:6}}>{label}</div>
    {payload.map((p,i)=><div key={i} style={{color:p.color,marginBottom:2}}>{p.name}: <strong>{typeof p.value==="number"&&p.value>100?f$(p.value):p.value}</strong></div>)}
  </div>;
};

const Alerta = ({tipo,titulo,mensaje}) => {
  const cfg={critico:{bg:B.rojoL,bl:B.rojo,c:B.rojo,icon:"🔴"},alerta:{bg:B.amberL,bl:B.amber,c:"#92400E",icon:"🟡"},ok:{bg:B.verdeL,bl:B.verde,c:B.verde,icon:"🟢"}};
  const c=cfg[tipo]||cfg.alerta;
  return <div style={{background:c.bg,borderLeft:`4px solid ${c.bl}`,borderRadius:"0 12px 12px 0",padding:"14px 18px",marginBottom:10}}>
    <div style={{fontWeight:600,fontSize:13,color:c.c,marginBottom:3}}>{c.icon} {titulo}</div>
    <div style={{fontSize:12,color:B.gris}}>{mensaje}</div>
  </div>;
};

const TipoBadge = ({tipo}) => {
  const m={cargos_costo:{label:"Cargos + Costo",c:"cyan"},cirugias:{label:"Cirugías",c:"teal"},inventario:{label:"Inventario",c:"amber"},abonos:{label:"Abonos",c:"violeta"},generico:{label:"Genérico",c:"gris"}};
  const t=m[tipo]||m.generico;
  const cols={cyan:{bg:B.cyanL,c:B.cyanD},teal:{bg:B.cyanLL,c:B.teal},amber:{bg:B.amberL,c:"#92400E"},violeta:{bg:B.violetaL,c:B.violeta},gris:{bg:B.grisL,c:B.grisM}};
  const cl=cols[t.c];
  return <span style={{display:"inline-block",padding:"2px 9px",borderRadius:20,fontSize:10,fontWeight:600,background:cl.bg,color:cl.c}}>{t.label}</span>;
};

// ─── ROLES ──────────────────────────────────────────────────────────
const ROLES = {
  accionista:{label:"Accionista",icon:"👔",color:B.navy,views:["ejecutivo","financiero","tendencias","alertas"]},
  director:  {label:"Director General",icon:"🏛",color:B.cyanD,views:["ejecutivo","financiero","operaciones","medicos","tendencias","alertas","ia"]},
  gerente:   {label:"Gerente Operativo",icon:"📋",color:B.teal,views:["operaciones","medicos","servicios","convenios","tendencias","alertas"]},
  analista:  {label:"Analista",icon:"📊",color:B.gris,views:["financiero","medicos","servicios","convenios","tendencias","ia"]},
};

// ─── CARGA ───────────────────────────────────────────────────────────
function Carga({onData}) {
  const [drag,setDrag]=useState(false);
  const [prog,setProg]=useState([]);
  const [loading,setLoading]=useState(false);
  const ref=useRef();

  const procesar=async(files)=>{
    setLoading(true);
    const fuentes=[];
    for(let i=0;i<files.length;i++){
      const f=files[i];
      setProg(p=>[...p,{nombre:f.name,estado:"⏳ Procesando..."}]);
      await new Promise(r=>setTimeout(r,30));
      try{
        const buf=await f.arrayBuffer();
        const wb=XLSX.read(buf,{type:"array",cellDates:true});
        const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:null,raw:false});
        if(rows.length>0){
          const hdrs=Object.keys(rows[0]);
          const tipo=detectTipo(hdrs);
          const mapa=autoMap(hdrs);
          fuentes.push({nombre:f.name,tipo,mapa,rows,hdrs,filas:rows.length});
          setProg(p=>p.map(x=>x.nombre===f.name?{...x,estado:`✅ ${rows.length.toLocaleString()} filas`}:x));
        }
      }catch(e){setProg(p=>p.map(x=>x.nombre===f.name?{...x,estado:"❌ Error"}:x));}
    }
    setLoading(false); setProg([]);
    if(fuentes.length>0) onData(fuentes);
  };

  return (
    <div style={{minHeight:"100vh",background:B.navy,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:40,position:"relative",overflow:"hidden"}}>
      {/* Fondo ECG decorativo */}
      <div style={{position:"absolute",bottom:40,left:0,right:0,opacity:0.06}}>
        <svg width="100%" height="60" viewBox="0 0 1400 60" preserveAspectRatio="none">
          <path d="M0 30 L200 30 L250 30 L270 5 L290 55 L310 20 L330 30 L600 30 L650 30 L670 5 L690 55 L710 20 L730 30 L1000 30 L1050 30 L1070 5 L1090 55 L1110 20 L1130 30 L1400 30"
            stroke={B.cyan} strokeWidth="2" fill="none"/>
        </svg>
      </div>

      <div style={{maxWidth:680,width:"100%",position:"relative",zIndex:1}}>
        {/* Logo */}
        <div style={{textAlign:"center",marginBottom:48}}>
          <div style={{fontSize:28,fontWeight:700,color:B.cyan,letterSpacing:"0.05em",fontFamily:"Georgia,serif",marginBottom:12}}>Hospital Punta Médica</div>
          <ECGLine width={280} height={22} color={B.cyan}/>
          <div style={{color:"rgba(255,255,255,0.4)",fontSize:11,letterSpacing:"0.15em",textTransform:"uppercase",marginTop:10}}>Centro de Inteligencia Empresarial</div>
        </div>

        {/* Drop zone */}
        <div onDrop={e=>{e.preventDefault();setDrag(false);procesar([...e.dataTransfer.files])}}
          onDragOver={e=>{e.preventDefault();setDrag(true)}} onDragLeave={()=>setDrag(false)}
          onClick={()=>!loading&&ref.current?.click()}
          style={{border:`2px dashed ${drag?B.cyan:"rgba(0,180,216,0.3)"}`,borderRadius:20,padding:"44px 32px",textAlign:"center",cursor:"pointer",background:drag?"rgba(0,180,216,0.06)":"rgba(255,255,255,0.03)",transition:"all 0.2s",marginBottom:24}}>
          <input ref={ref} type="file" multiple accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={e=>procesar([...e.target.files])}/>
          {loading?(
            <div>
              <div style={{fontSize:13,color:B.cyan,fontWeight:500,marginBottom:14}}>Procesando archivos ECP...</div>
              {prog.map((p,i)=><div key={i} style={{fontSize:12,color:"rgba(255,255,255,0.5)",marginBottom:5}}>
                <span style={{color:"rgba(255,255,255,0.8)"}}>{p.nombre}</span> → {p.estado}
              </div>)}
            </div>
          ):(
            <>
              <div style={{fontSize:38,marginBottom:10}}>📂</div>
              <div style={{fontSize:16,fontWeight:600,color:"white",marginBottom:6}}>Arrastra tus reportes ECP aquí</div>
              <div style={{fontSize:13,color:"rgba(255,255,255,0.4)",marginBottom:20}}>Soporte múltiple · .xlsx · .xls · .csv</div>
              <div style={{display:"inline-block",background:B.cyan,color:B.navy,padding:"10px 28px",borderRadius:8,fontSize:13,fontWeight:700,letterSpacing:"0.05em"}}>SELECCIONAR ARCHIVOS</div>
            </>
          )}
        </div>

        {/* Tipos compatibles */}
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10}}>
          {[{icon:"💰",l:"Cargos con Costo",d:"Admisión → Reporte"},
            {icon:"👥",l:"Censo Pacientes",d:"Admisión → Histórico"},
            {icon:"🔪",l:"Cirugías",d:"Quirófano → General"},
            {icon:"📦",l:"Inventario",d:"Almacén → Existencias"}
          ].map((t,i)=>(
            <div key={i} style={{background:"rgba(0,180,216,0.05)",border:"1px solid rgba(0,180,216,0.15)",borderRadius:12,padding:"14px 12px",textAlign:"center"}}>
              <div style={{fontSize:22,marginBottom:6}}>{t.icon}</div>
              <div style={{fontSize:11,fontWeight:600,color:"rgba(255,255,255,0.8)",marginBottom:3}}>{t.l}</div>
              <div style={{fontSize:10,color:"rgba(255,255,255,0.3)"}}>{t.d}</div>
            </div>
          ))}
        </div>
      </div>
      {/* ══ SETTINGS MODAL ══ */}
      {showSettings && (
        <div onClick={e=>{if(e.target===e.currentTarget)setShowSettings(false)}} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:999,display:"flex",alignItems:"center",justifyContent:"center",padding:24}}>
          <div style={{background:B.blanco,borderRadius:20,width:"100%",maxWidth:480,boxShadow:"0 24px 60px rgba(0,0,0,0.3)",overflow:"hidden"}}>
            {/* Modal header */}
            <div style={{background:B.navy,padding:"20px 24px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
              <div style={{display:"flex",alignItems:"center",gap:10}}>
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke={B.cyan} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>
                <span style={{color:"white",fontWeight:700,fontSize:15}}>Ajustes de la plataforma</span>
              </div>
              <button onClick={()=>setShowSettings(false)} style={{background:"rgba(255,255,255,0.1)",border:"none",borderRadius:6,width:28,height:28,cursor:"pointer",color:"white",fontSize:16,display:"flex",alignItems:"center",justifyContent:"center"}}>✕</button>
            </div>

            <div style={{padding:"24px"}}>
              {/* Logo upload */}
              <div style={{marginBottom:24}}>
                <div style={{fontSize:11,textTransform:"uppercase",letterSpacing:"0.1em",color:B.grisM,fontWeight:600,marginBottom:12}}>Logotipo de la empresa</div>
                <input ref={logoRef} type="file" accept="image/*" style={{display:"none"}} onChange={e=>{
                  const file=e.target.files[0];
                  if(!file) return;
                  const reader=new FileReader();
                  reader.onload=ev=>setCustomLogo(ev.target.result);
                  reader.readAsDataURL(file);
                }}/>
                <div style={{display:"flex",gap:12,alignItems:"center"}}>
                  {/* Preview */}
                  <div style={{width:160,height:64,border:`2px dashed ${customLogo?B.cyan:B.borde}`,borderRadius:12,display:"flex",alignItems:"center",justifyContent:"center",background:customLogo?B.navy:B.cyanLL,overflow:"hidden",flexShrink:0}}>
                    {customLogo
                      ? <img src={customLogo} alt="Logo" style={{maxHeight:52,maxWidth:148,objectFit:"contain"}}/>
                      : <span style={{fontSize:11,color:B.grisM,textAlign:"center",padding:"0 8px"}}>Sin logo<br/>personalizado</span>
                    }
                  </div>
                  <div style={{display:"flex",flexDirection:"column",gap:8,flex:1}}>
                    <button onClick={()=>logoRef.current?.click()} style={{padding:"9px 16px",background:B.cyan,color:B.navy,border:"none",borderRadius:8,fontWeight:700,fontSize:12,cursor:"pointer",letterSpacing:"0.05em"}}>⬆ Subir logo</button>
                    {customLogo && <button onClick={()=>setCustomLogo(null)} style={{padding:"8px 16px",background:"none",color:B.rojo,border:`1px solid ${B.rojo}`,borderRadius:8,fontSize:12,cursor:"pointer"}}>✕ Quitar logo</button>}
                    <span style={{fontSize:10,color:B.grisM}}>PNG, JPG · Recomendado fondo transparente</span>
                  </div>
                </div>
              </div>

              {/* Nombre */}
              <div style={{marginBottom:20}}>
                <label style={{fontSize:11,textTransform:"uppercase",letterSpacing:"0.1em",color:B.grisM,fontWeight:600,display:"block",marginBottom:8}}>Nombre de la empresa</label>
                <input value={brandName} onChange={e=>setBrandName(e.target.value)} style={{width:"100%",padding:"10px 14px",border:`1px solid ${B.borde}`,borderRadius:8,fontSize:13,color:B.navy,outline:"none",fontFamily:"inherit"}} onFocus={e=>e.target.style.borderColor=B.cyan} onBlur={e=>e.target.style.borderColor=B.borde}/>
              </div>

              {/* Subtítulo */}
              <div style={{marginBottom:24}}>
                <label style={{fontSize:11,textTransform:"uppercase",letterSpacing:"0.1em",color:B.grisM,fontWeight:600,display:"block",marginBottom:8}}>Subtítulo / Eslogan</label>
                <input value={brandSub} onChange={e=>setBrandSub(e.target.value)} style={{width:"100%",padding:"10px 14px",border:`1px solid ${B.borde}`,borderRadius:8,fontSize:13,color:B.navy,outline:"none",fontFamily:"inherit"}} onFocus={e=>e.target.style.borderColor=B.cyan} onBlur={e=>e.target.style.borderColor=B.borde}/>
              </div>

              {/* Preview del navbar */}
              <div style={{marginBottom:24}}>
                <div style={{fontSize:11,textTransform:"uppercase",letterSpacing:"0.1em",color:B.grisM,fontWeight:600,marginBottom:8}}>Vista previa en barra de navegación</div>
                <div style={{background:B.navy,borderRadius:10,padding:"10px 16px",display:"flex",alignItems:"center",gap:10}}>
                  {customLogo
                    ? <img src={customLogo} alt={brandName} style={{height:24,width:"auto",maxWidth:140,objectFit:"contain"}}/>
                    : <LogoPM size={20}/>
                  }
                  <div style={{width:1,height:20,background:"rgba(255,255,255,0.15)"}}/>
                  <span style={{fontSize:9,color:"rgba(255,255,255,0.4)",letterSpacing:"0.1em",textTransform:"uppercase"}}>{brandSub}</span>
                </div>
              </div>

              {/* Acciones */}
              <div style={{display:"flex",gap:10}}>
                <button onClick={()=>setShowSettings(false)} style={{flex:1,padding:"11px",background:B.cyan,color:B.navy,border:"none",borderRadius:10,fontWeight:700,fontSize:13,cursor:"pointer",letterSpacing:"0.05em"}}>✓ Guardar ajustes</button>
                <button onClick={()=>{setCustomLogo(null);setBrandName("Hospital Punta Médica");setBrandSub("Alta Especialidad");}} style={{padding:"11px 16px",background:"none",color:B.grisM,border:`1px solid ${B.borde}`,borderRadius:10,fontSize:12,cursor:"pointer"}}>Restaurar</button>
              </div>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}

// ─── SELECTOR DE ROL ─────────────────────────────────────────────────
function SelectorRol({onSelect, customLogo, brandName, brandSub}) {
  return (
    <div style={{minHeight:"100vh",background:B.cyanLL,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:40}}>
      <div style={{maxWidth:580,width:"100%",textAlign:"center"}}>
        {customLogo
          ? <img src={customLogo} alt={brandName} style={{height:48,width:"auto",maxWidth:280,objectFit:"contain",marginBottom:8}}/>
          : <div style={{fontSize:22,fontWeight:700,color:B.navy,letterSpacing:"0.02em",marginBottom:8}}>{brandName}</div>
        }
        <div style={{marginBottom:32}}>
          <ECGLine width={240} height={20} color={B.cyan}/>
        </div>
        <div style={{fontFamily:"Georgia,serif",fontSize:28,fontWeight:700,color:B.navy,marginBottom:6}}>¿Cuál es tu perfil?</div>
        <div style={{fontSize:13,color:B.grisM,marginBottom:32}}>El dashboard se personaliza según tu nivel de acceso</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
          {Object.entries(ROLES).map(([k,v])=>(
            <button key={k} onClick={()=>onSelect(k)}
              style={{background:B.blanco,border:`2px solid ${B.borde}`,borderRadius:16,padding:"24px 20px",cursor:"pointer",textAlign:"left",transition:"all 0.2s"}}
              onMouseEnter={e=>{e.currentTarget.style.border=`2px solid ${B.cyan}`;e.currentTarget.style.boxShadow="0 8px 24px rgba(0,180,216,0.15)";}}
              onMouseLeave={e=>{e.currentTarget.style.border=`2px solid ${B.borde}`;e.currentTarget.style.boxShadow="none";}}>
              <div style={{fontSize:28,marginBottom:10}}>{v.icon}</div>
              <div style={{fontWeight:700,fontSize:15,color:B.navy,marginBottom:4}}>{v.label}</div>
              <div style={{width:32,height:2,background:v.color,borderRadius:1,marginBottom:8}}/>
              <div style={{fontSize:11,color:B.grisM}}>Vistas: {v.views.join(" · ")}</div>
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}

// ─── MOTOR PRINCIPAL ─────────────────────────────────────────────────
export default function PuntaMedicaIntel() {
  const [fuentes,setFuentes]=useState(null);
  const [rol,setRol]=useState(null);
  const [view,setView]=useState("ejecutivo");
  const [filtros,setFiltros]=useState({mes:"Todos",servicio:"Todos",convenio:"Todos",medico:"Todos"});
  const [compMes,setCompMes]=useState(null);
  const [umbral,setUmbral]=useState({margen:75,concentracion:40});
  const [aiQ,setAiQ]=useState("");const [aiH,setAiH]=useState([]);const [aiL,setAiL]=useState(false);
  const [sortPor,setSortPor]=useState("total");
  const [pag,setPag]=useState(0);
  const [showSettings,setShowSettings]=useState(false);
  const [customLogo,setCustomLogo]=useState(null);
  const [brandName,setBrandName]=useState('Hospital Punta Médica');
  const [brandSub,setBrandSub]=useState('Alta Especialidad');
  const logoRef=useRef();

  const allRows=useMemo(()=>{
    if(!fuentes) return [];
    return fuentes.flatMap(f=>f.rows.map(r=>({
      _fuente:f.tipo,
      _cargo:Number(r[f.mapa.cargo])||0,
      _costo:Number(r[f.mapa.costo])||0,
      _medico:f.mapa.medico?String(r[f.mapa.medico]||"Sin asignar").trim():"Sin asignar",
      _paciente:f.mapa.paciente?String(r[f.mapa.paciente]||"").trim():"",
      _servicio:f.mapa.servicio?String(r[f.mapa.servicio]||"General").trim():"General",
      _convenio:f.mapa.convenio?String(r[f.mapa.convenio]||"PARTICULAR").trim():"PARTICULAR",
      _folio:f.mapa.folio?String(r[f.mapa.folio]||""):"",
      _fecha:f.mapa.fecha?parseDate(r[f.mapa.fecha]):null,...r
    }))).map(r=>({...r,_mes:getMes(r._fecha),_q:getQ(r._fecha),_anio:r._fecha?String(r._fecha).substring(0,4):"S/D"}));
  },[fuentes]);

  const opts=useMemo(()=>({
    meses:[...new Set(allRows.map(r=>r._mes).filter(x=>x&&x!=="Sin fecha"))].sort(),
    servicios:[...new Set(allRows.map(r=>r._servicio).filter(Boolean))].sort(),
    convenios:[...new Set(allRows.map(r=>r._convenio).filter(Boolean))].sort(),
    medicos:[...new Set(allRows.map(r=>r._medico).filter(Boolean))].sort(),
  }),[allRows]);

  const rows=useMemo(()=>allRows.filter(r=>
    (filtros.mes==="Todos"||r._mes===filtros.mes)&&
    (filtros.servicio==="Todos"||r._servicio===filtros.servicio)&&
    (filtros.convenio==="Todos"||r._convenio===filtros.convenio)&&
    (filtros.medico==="Todos"||r._medico===filtros.medico)
  ),[allRows,filtros]);

  const rowsC=useMemo(()=>compMes?allRows.filter(r=>r._mes===compMes):[],[allRows,compMes]);

  const K=useMemo(()=>{
    const ing=rows.reduce((s,r)=>s+r._cargo,0);
    const cos=rows.reduce((s,r)=>s+r._costo,0);
    const pac=new Set(rows.map(r=>r._paciente).filter(Boolean)).size;
    const fol=new Set(rows.map(r=>r._folio).filter(Boolean)).size;
    return{ingreso:ing,costo:cos,utilidad:ing-cos,margen:ing>0?(ing-cos)/ing:0,pac,fol,eventos:rows.length,ticket:pac>0?ing/pac:0};
  },[rows]);

  const KC=useMemo(()=>{
    const ing=rowsC.reduce((s,r)=>s+r._cargo,0);
    const cos=rowsC.reduce((s,r)=>s+r._costo,0);
    const pac=new Set(rowsC.map(r=>r._paciente).filter(Boolean)).size;
    return{ingreso:ing,costo:cos,margen:ing>0?(ing-cos)/ing:0,pac};
  },[rowsC]);

  const agg=(campo,val="_cargo")=>{
    const m={};
    for(const r of rows){
      const k=r[campo]||"Sin datos";
      if(!m[k])m[k]={key:k,total:0,costo:0,count:0,pacs:new Set(),fols:new Set()};
      m[k].total+=r[val]||0;m[k].costo+=r._costo||0;m[k].count++;
      if(r._paciente)m[k].pacs.add(r._paciente);
      if(r._folio)m[k].fols.add(r._folio);
    }
    return Object.values(m).map(x=>({...x,pacs:x.pacs.size,fols:x.fols.size,margen:x.total>0?(x.total-x.costo)/x.total:0})).sort((a,b)=>b.total-a.total);
  };

  const byMes=useMemo(()=>{
    const m={};
    for(const r of allRows){
      const k=r._mes||"S/F";
      if(!m[k])m[k]={mes:k,cargo:0,costo:0,ev:0,pacs:new Set(),fols:new Set()};
      m[k].cargo+=r._cargo;m[k].costo+=r._costo;m[k].ev++;
      if(r._paciente)m[k].pacs.add(r._paciente);
      if(r._folio)m[k].fols.add(r._folio);
    }
    return Object.values(m).map(x=>({...x,pacs:x.pacs.size,fols:x.fols.size,margen:x.cargo>0?((x.cargo-x.costo)/x.cargo*100).toFixed(1):0,ticket:x.pacs>0?Math.round(x.cargo/x.pacs):0})).sort((a,b)=>a.mes>b.mes?1:-1);
  },[allRows]);

  const byServ=useMemo(()=>agg("_servicio"),[rows]);
  const byConv=useMemo(()=>agg("_convenio"),[rows]);
  const byMed=useMemo(()=>agg("_medico"),[rows]);

  const forecast=useMemo(()=>{
    if(byMes.length<2)return[];
    const ult=byMes.slice(-3);
    const avg=ult.reduce((s,m)=>s+m.cargo,0)/ult.length;
    const g=byMes.length>1?(byMes[byMes.length-1].cargo/byMes[0].cargo-1)/(byMes.length-1):0.058;
    return["M+1","M+2","M+3","M+4","M+5","M+6"].map((mes,i)=>({mes,conservador:Math.round(avg*Math.pow(1+g*0.85,i+1)),base:Math.round(avg*Math.pow(1+g,i+1)),optimista:Math.round(avg*Math.pow(1+g*1.15,i+1))}));
  },[byMes]);

  const alertas=useMemo(()=>{
    const a=[];
    if(K.margen*100<umbral.margen)a.push({tipo:"critico",titulo:"Margen Bruto bajo el umbral",mensaje:`Margen actual ${fP(K.margen)} — umbral configurado: ${umbral.margen}%`});
    const t3=byMed.slice(0,3).reduce((s,m)=>s+m.total,0);
    if(K.ingreso>0&&t3/K.ingreso*100>umbral.concentracion)a.push({tipo:"alerta",titulo:"Concentración médica alta",mensaje:`Top-3 médicos generan ${fP(t3/K.ingreso)} del ingreso — umbral: ${umbral.concentracion}%`});
    const sinC=rows.filter(r=>r._cargo>0&&r._costo===0).length;
    if(sinC>rows.length*0.3)a.push({tipo:"alerta",titulo:"Registros sin costo capturado",mensaje:`${fN(sinC)} registros (${fP(sinC/rows.length)}) no tienen costo directo`});
    if(a.length===0)a.push({tipo:"ok",titulo:"Indicadores dentro de parámetros",mensaje:"Todos los KPIs cumplen los umbrales definidos"});
    return a;
  },[K,byMed,rows,umbral]);

  const askAI=async()=>{
    if(!aiQ.trim())return;
    setAiL(true);
    const ctx={ingreso:K.ingreso,margen:fP(K.margen),pacientes:K.pac,folios:K.fol,eventos:K.eventos,
      top5med:byMed.slice(0,5).map(m=>({medico:m.key,ingreso:m.total,eventos:m.count})),
      top5serv:byServ.slice(0,5).map(s=>({servicio:s.key,ingreso:s.total})),
      top5conv:byConv.slice(0,5).map(c=>({convenio:c.key,ingreso:c.total})),
      tendencia:byMes.slice(-4).map(m=>({mes:m.mes,ingreso:m.cargo,pacs:m.pacs})),
      filtros};
    const userM={role:"user",content:aiQ};
    const hist=[...aiH,userM];
    try{
      const r=await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json"},
        body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:1000,
          system:`Eres analista financiero senior de Hospital Punta Médica, México. Marca: Alta Especialidad. Contexto: ${JSON.stringify(ctx)}. Responde en español, ejecutivo, con cifras del contexto.`,
          messages:hist})});
      const d=await r.json();
      setAiH([...hist,{role:"assistant",content:d.content?.[0]?.text||"Error"}]);
    }catch(e){setAiH([...hist,{role:"assistant",content:"Error de conexión."}]);}
    setAiL(false);setAiQ("");
  };

  if(!rol) return <SelectorRol onSelect={k=>{setRol(k);setView(ROLES[k].views[0]);}} customLogo={customLogo} brandName={brandName} brandSub={brandSub}/>;
  if(!fuentes) return <Carga onData={setFuentes}/>;

  const VIEWS_MAP={
    ejecutivo:{l:"Dashboard",icon:"📊"},financiero:{l:"Financiero",icon:"💰"},
    operaciones:{l:"Operaciones",icon:"⚙️"},medicos:{l:"Médicos",icon:"👨‍⚕️"},
    servicios:{l:"Servicios",icon:"🏥"},convenios:{l:"Convenios",icon:"📄"},
    tendencias:{l:"Tendencias",icon:"📈"},alertas:{l:"Alertas",icon:"🔔"},ia:{l:"IA",icon:"✨"},
  };

  const TablaDetalle=({data,titulo})=>{
    const total=data.reduce((s,d)=>s+d.total,0);
    const sorted=[...data].sort((a,b)=>sortPor==="count"?b.count-a.count:b.total-a.total);
    const PG=10; const paged=sorted.slice(pag*PG,(pag+1)*PG);
    return <div>
      <div style={{display:"flex",gap:8,marginBottom:12,justifyContent:"flex-end"}}>
        {[["total","Por Ingreso"],["count","Por Eventos"]].map(([k,l])=>(
          <button key={k} onClick={()=>setSortPor(k)} style={{padding:"5px 14px",border:`1px solid ${sortPor===k?B.cyan:B.borde}`,borderRadius:20,background:sortPor===k?B.cyanL:"none",color:sortPor===k?B.cyanD:B.grisM,fontSize:11,cursor:"pointer",fontWeight:sortPor===k?600:400}}>{l}</button>
        ))}
      </div>
      <div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
          <thead><tr style={{background:B.cyanLL}}>
            {["#",titulo,"Ingreso","Margen","Eventos","Pac./Fol.","% Total"].map(h=>(
              <th key={h} style={{padding:"10px 12px",textAlign:h==="#"||h===titulo?"left":"right",fontSize:10,textTransform:"uppercase",letterSpacing:"0.08em",color:B.cyanD,fontWeight:700}}>{h}</th>
            ))}
          </tr></thead>
          <tbody>{paged.map((d,i)=>{
            const pct=total>0?d.total/total:0;
            return <tr key={i} style={{borderBottom:`1px solid ${B.borde}`,background:i===0?B.cyanLL:B.blanco}}>
              <td style={{padding:"10px 12px",color:B.grisM}}>{pag*PG+i+1}</td>
              <td style={{padding:"10px 12px",fontWeight:600,maxWidth:180,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:B.navy}}>{d.key}</td>
              <td style={{padding:"10px 12px",textAlign:"right",fontWeight:700,color:B.cyanD}}>{f$(d.total)}</td>
              <td style={{padding:"10px 12px",textAlign:"right",color:d.margen>0.75?B.verde:d.margen>0.5?B.amber:B.rojo}}>{d.costo>0?fP(d.margen):"—"}</td>
              <td style={{padding:"10px 12px",textAlign:"right"}}>{fN(d.count)}</td>
              <td style={{padding:"10px 12px",textAlign:"right"}}>{fN(d.pacs||d.fols||0)}</td>
              <td style={{padding:"10px 12px",textAlign:"right"}}>
                <div style={{display:"flex",alignItems:"center",gap:6,justifyContent:"flex-end"}}>
                  <div style={{width:44,height:5,background:B.borde,borderRadius:3}}><div style={{width:`${Math.min(pct*100,100)}%`,height:"100%",background:B.cyan,borderRadius:3}}/></div>
                  <span style={{color:B.cyanD,fontWeight:600}}>{fP(pct)}</span>
                </div>
              </td>
            </tr>;
          })}</tbody>
        </table>
      </div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:10,fontSize:11,color:B.grisM}}>
        <span>{pag*PG+1}–{Math.min((pag+1)*PG,data.length)} de {data.length} registros</span>
        <div style={{display:"flex",gap:6}}>
          <button onClick={()=>setPag(p=>Math.max(0,p-1))} disabled={pag===0} style={{padding:"5px 14px",border:`1px solid ${B.borde}`,borderRadius:6,cursor:"pointer",background:"none",color:pag===0?B.grisM:B.cyanD,fontWeight:500,fontSize:11}}>← Anterior</button>
          <button onClick={()=>setPag(p=>(p+1)*PG<data.length?p+1:p)} disabled={(pag+1)*PG>=data.length} style={{padding:"5px 14px",border:`1px solid ${B.borde}`,borderRadius:6,cursor:"pointer",background:"none",color:(pag+1)*PG>=data.length?B.grisM:B.cyanD,fontWeight:500,fontSize:11}}>Siguiente →</button>
        </div>
      </div>
    </div>;
  };

  return (
    <div style={{fontFamily:"system-ui,sans-serif",background:B.grisL,minHeight:"100vh",fontSize:13}}>

      {/* TOP NAV */}
      <div style={{background:B.navy,padding:"0 24px",display:"flex",alignItems:"center",gap:8,position:"sticky",top:0,zIndex:100}}>
        <div style={{marginRight:20,display:"flex",alignItems:"center",gap:10,padding:"8px 0"}}>
          {customLogo ? <img src={customLogo} alt={brandName} style={{height:28,width:"auto",maxWidth:160,objectFit:"contain"}}/> : <LogoPM size={22}/>}
          <div style={{width:1,height:24,background:"rgba(255,255,255,0.15)"}}/>
          <div style={{fontSize:10,color:"rgba(255,255,255,0.4)",letterSpacing:"0.1em",textTransform:"uppercase"}}>{brandSub}</div>
        </div>
        <div style={{display:"flex",gap:0,flex:1,overflowX:"auto"}}>
          {ROLES[rol].views.map(v=>(
            <button key={v} onClick={()=>{setView(v);setPag(0);}} style={{padding:"14px 13px",border:"none",background:"none",cursor:"pointer",fontSize:12,fontWeight:500,whiteSpace:"nowrap",color:view===v?"white":"rgba(255,255,255,0.45)",borderBottom:view===v?`3px solid ${B.cyan}`:"3px solid transparent",transition:"color 0.15s"}}>
              {VIEWS_MAP[v]?.icon} {VIEWS_MAP[v]?.l}
            </button>
          ))}
        </div>
        <div style={{display:"flex",gap:8,alignItems:"center",flexShrink:0}}>
          <select value={filtros.mes} onChange={e=>setFiltros(f=>({...f,mes:e.target.value}))} style={{padding:"5px 8px",borderRadius:6,border:"1px solid rgba(0,180,216,0.3)",background:"rgba(0,180,216,0.1)",color:"white",fontSize:11}}>
            <option value="Todos">Todos los meses</option>
            {opts.meses.map(m=><option key={m} value={m}>{m}</option>)}
          </select>
          <button onClick={()=>{setFuentes(null);}} style={{padding:"5px 12px",background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.1)",borderRadius:6,color:"rgba(255,255,255,0.5)",fontSize:11,cursor:"pointer"}}>+ Datos</button>
          <button onClick={()=>window.print()} style={{padding:"6px 16px",background:B.cyan,border:"none",borderRadius:6,color:B.navy,fontSize:11,fontWeight:700,cursor:"pointer",letterSpacing:"0.05em"}}>⬇ PDF</button>
          <button onClick={()=>setShowSettings(true)} title="Ajustes" style={{width:32,height:32,display:"flex",alignItems:"center",justifyContent:"center",background:"rgba(255,255,255,0.08)",border:"1px solid rgba(0,180,216,0.3)",borderRadius:8,cursor:"pointer",flexShrink:0}} onMouseEnter={e=>e.currentTarget.style.background="rgba(0,180,216,0.2)"} onMouseLeave={e=>e.currentTarget.style.background="rgba(255,255,255,0.08)"}>
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke={B.cyan} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>
          </button>
        </div>
      </div>

      {/* SUBHEADER */}
      <div style={{background:B.cyanLL,borderBottom:`1px solid ${B.borde}`,padding:"8px 24px",display:"flex",gap:16,alignItems:"center",fontSize:11,color:B.grisM,flexWrap:"wrap"}}>
        {fuentes?.map((f,i)=><span key={i} style={{display:"flex",alignItems:"center",gap:6}}><TipoBadge tipo={f.tipo}/> {f.nombre} <span style={{color:B.cyan,fontWeight:600}}>({fN(f.filas)} filas)</span></span>)}
        <span style={{marginLeft:"auto",color:B.gris}}>{fN(rows.length)} registros · <strong style={{color:B.cyanD}}>{fN(K.fol)}</strong> folios · <strong style={{color:B.cyanD}}>{fN(K.pac)}</strong> pacientes</span>
        <span style={{color:B.grisM}}>{ROLES[rol].icon} {ROLES[rol].label}</span>
      </div>

      <div style={{padding:"22px 24px"}}>

        {/* ══ DASHBOARD EJECUTIVO ══ */}
        {view==="ejecutivo"&&<>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:20}}>
            <KpiCard label="Ingreso Total" value={fM(K.ingreso)} sub={`${fN(K.eventos)} eventos`} delta={compMes&&KC.ingreso?fDelta(K.ingreso,KC.ingreso):null} color={B.cyan} icon="💰"/>
            <KpiCard label="Margen Bruto" value={fP(K.margen)} sub={`Costo: ${fM(K.costo)}`} badge={K.margen>0.75?"✓ Sólido":"⚠ Revisar"} badgeOk={K.margen>0.75} color={B.teal} icon="📊"/>
            <KpiCard label="Pacientes Únicos" value={fN(K.pac)} sub={`Ticket: ${f$(K.ticket)}`} delta={compMes&&KC.pac?fDelta(K.pac,KC.pac):null} color={B.cyanD} icon="🧑‍⚕️"/>
            <KpiCard label="Utilidad Bruta" value={fM(K.utilidad)} badge={K.utilidad>0?"Positivo":"⚠ Negativo"} badgeOk={K.utilidad>0} color={B.cyan} icon="📈"/>
          </div>

          <div style={{background:B.blanco,border:`1px solid ${B.borde}`,borderRadius:12,padding:"10px 18px",marginBottom:18,display:"flex",gap:12,alignItems:"center",flexWrap:"wrap"}}>
            <span style={{fontSize:12,fontWeight:600,color:B.navy}}>Comparar período:</span>
            <select value={compMes||""} onChange={e=>setCompMes(e.target.value||null)} style={{padding:"6px 10px",border:`1px solid ${B.borde}`,borderRadius:6,fontSize:12,color:B.navy}}>
              <option value="">— Sin comparación —</option>
              {opts.meses.map(m=><option key={m} value={m}>{m}</option>)}
            </select>
            {compMes&&<span style={{fontSize:12,color:B.grisM}}>Período {compMes}: <strong>{fM(KC.ingreso)}</strong> · {fN(KC.pac)} pac.</span>}
            <div style={{flex:1,display:"flex",gap:8,justifyContent:"flex-end",flexWrap:"wrap"}}>
              {[["mes","Mes"],["servicio","Servicio"],["convenio","Convenio"],["medico","Médico"]].map(([k,l])=>(
                <select key={k} value={filtros[k]} onChange={e=>setFiltros(f=>({...f,[k]:e.target.value}))} style={{padding:"5px 8px",border:`1px solid ${B.borde}`,borderRadius:6,fontSize:11,color:filtros[k]!=="Todos"?B.cyanD:B.grisM}}>
                  <option value="Todos">Todos los {l}s</option>
                  {(k==="mes"?opts.meses:k==="servicio"?opts.servicios:k==="convenio"?opts.convenios:opts.medicos).slice(0,60).map(o=><option key={o} value={o}>{o}</option>)}
                </select>
              ))}
            </div>
          </div>

          <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:16,marginBottom:16}}>
            <Card>
              <div style={{fontWeight:700,marginBottom:14,color:B.navy}}>Ingreso vs Costo — Evolución Mensual</div>
              <ResponsiveContainer width="100%" height={230}>
                <AreaChart data={byMes}>
                  <defs>
                    <linearGradient id="gc" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={B.cyan} stopOpacity={0.2}/><stop offset="95%" stopColor={B.cyan} stopOpacity={0}/></linearGradient>
                    <linearGradient id="gr" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={B.rojo} stopOpacity={0.1}/><stop offset="95%" stopColor={B.rojo} stopOpacity={0}/></linearGradient>
                  </defs>
                  <CartesianGrid strokeDasharray="3 3" stroke={B.borde}/>
                  <XAxis dataKey="mes" tick={{fontSize:10,fill:B.grisM}}/><YAxis tickFormatter={v=>"$"+(v/1e6).toFixed(1)+"M"} tick={{fontSize:10,fill:B.grisM}}/>
                  <Tooltip content={<Tip/>}/><Legend/>
                  <Area type="monotone" dataKey="cargo" stroke={B.cyan} fill="url(#gc)" strokeWidth={2.5} name="Ingreso"/>
                  <Area type="monotone" dataKey="costo" stroke={B.rojo} fill="url(#gr)" strokeWidth={1.5} name="Costo"/>
                </AreaChart>
              </ResponsiveContainer>
            </Card>
            <Card>
              <div style={{fontWeight:700,marginBottom:14,color:B.navy}}>Mix Convenios</div>
              <ResponsiveContainer width="100%" height={230}>
                <PieChart>
                  <Pie data={byConv.slice(0,6)} dataKey="total" nameKey="key" cx="50%" cy="50%" outerRadius={85} innerRadius={30}
                    label={({percent})=>percent>0.05?`${(percent*100).toFixed(0)}%`:""} labelLine={false}>
                    {byConv.slice(0,6).map((_,i)=><Cell key={i} fill={[B.cyan,B.teal,B.cyanD,B.amber,B.violeta,B.verde][i]}/>)}
                  </Pie>
                  <Tooltip formatter={v=>f$(v)}/><Legend iconSize={10} wrapperStyle={{fontSize:10}}/>
                </PieChart>
              </ResponsiveContainer>
            </Card>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
            <Card>
              <div style={{fontWeight:700,marginBottom:12,color:B.navy}}>Top 8 Médicos</div>
              <ResponsiveContainer width="100%" height={210}>
                <BarChart data={byMed.slice(0,8)} layout="vertical">
                  <CartesianGrid strokeDasharray="3 3" stroke={B.borde}/>
                  <XAxis type="number" tickFormatter={v=>"$"+(v/1e6).toFixed(1)+"M"} tick={{fontSize:10}}/><YAxis type="category" dataKey="key" tick={{fontSize:9}} width={110}/>
                  <Tooltip content={<Tip/>}/><Bar dataKey="total" fill={B.cyan} radius={[0,6,6,0]} name="Ingreso"/>
                </BarChart>
              </ResponsiveContainer>
            </Card>
            <Card>
              <div style={{fontWeight:700,marginBottom:12,color:B.navy}}>Top 8 Servicios</div>
              <ResponsiveContainer width="100%" height={210}>
                <BarChart data={byServ.slice(0,8)} layout="vertical">
                  <CartesianGrid strokeDasharray="3 3" stroke={B.borde}/>
                  <XAxis type="number" tickFormatter={v=>"$"+(v/1e6).toFixed(1)+"M"} tick={{fontSize:10}}/><YAxis type="category" dataKey="key" tick={{fontSize:9}} width={110}/>
                  <Tooltip content={<Tip/>}/><Bar dataKey="total" fill={B.teal} radius={[0,6,6,0]} name="Ingreso"/>
                </BarChart>
              </ResponsiveContainer>
            </Card>
          </div>
        </>}

        {/* ══ TENDENCIAS ══ */}
        {view==="tendencias"&&<>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:20}}>
            <KpiCard label="Meses con datos" value={byMes.length} color={B.cyan}/>
            <KpiCard label="Mejor mes" value={byMes.length?fM(Math.max(...byMes.map(m=>m.cargo))):"—"} sub={byMes.length?byMes.reduce((a,b)=>a.cargo>b.cargo?a:b).mes:""} color={B.verde}/>
            <KpiCard label="Promedio mensual" value={byMes.length?fM(byMes.reduce((s,m)=>s+m.cargo,0)/byMes.length):"—"} color={B.teal}/>
            <KpiCard label="Pacs. prom/mes" value={byMes.length?fN(byMes.reduce((s,m)=>s+m.pacs,0)/byMes.length):"—"} color={B.cyanD}/>
          </div>
          <Card style={{marginBottom:16}}>
            <div style={{fontWeight:700,marginBottom:14,color:B.navy}}>Ingreso + Margen % por Mes</div>
            <ResponsiveContainer width="100%" height={260}>
              <ComposedChart data={byMes}>
                <CartesianGrid strokeDasharray="3 3" stroke={B.borde}/>
                <XAxis dataKey="mes" tick={{fontSize:10,fill:B.grisM}}/>
                <YAxis yAxisId="l" tickFormatter={v=>"$"+(v/1e6).toFixed(1)+"M"} tick={{fontSize:10}}/><YAxis yAxisId="r" orientation="right" tickFormatter={v=>v+"%"} tick={{fontSize:10}}/>
                <Tooltip content={<Tip/>}/><Legend/>
                <Bar yAxisId="l" dataKey="cargo" fill={B.cyan} name="Ingreso" radius={[4,4,0,0]} opacity={0.85}/>
                <Line yAxisId="r" type="monotone" dataKey="margen" stroke={B.amber} strokeWidth={2.5} dot={{r:4,fill:B.amber}} name="Margen %"/>
              </ComposedChart>
            </ResponsiveContainer>
          </Card>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>
            <Card>
              <div style={{fontWeight:700,marginBottom:12,color:B.navy}}>Pacientes y Folios / Mes</div>
              <ResponsiveContainer width="100%" height={200}>
                <LineChart data={byMes}>
                  <CartesianGrid strokeDasharray="3 3" stroke={B.borde}/><XAxis dataKey="mes" tick={{fontSize:10}}/><YAxis tick={{fontSize:10}}/>
                  <Tooltip/><Legend/>
                  <Line type="monotone" dataKey="pacs" stroke={B.cyan} strokeWidth={2.5} dot={{r:4}} name="Pacientes"/>
                  <Line type="monotone" dataKey="fols" stroke={B.amber} strokeWidth={2.5} dot={{r:4}} name="Folios"/>
                </LineChart>
              </ResponsiveContainer>
            </Card>
            <Card>
              <div style={{fontWeight:700,marginBottom:12,color:B.navy}}>Ticket Promedio / Mes</div>
              <ResponsiveContainer width="100%" height={200}>
                <AreaChart data={byMes}>
                  <defs><linearGradient id="gt" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={B.teal} stopOpacity={0.2}/><stop offset="95%" stopColor={B.teal} stopOpacity={0}/></linearGradient></defs>
                  <CartesianGrid strokeDasharray="3 3" stroke={B.borde}/><XAxis dataKey="mes" tick={{fontSize:10}}/><YAxis tickFormatter={v=>f$(v)} tick={{fontSize:10}}/>
                  <Tooltip content={<Tip/>}/><Area type="monotone" dataKey="ticket" stroke={B.teal} fill="url(#gt)" strokeWidth={2.5} name="Ticket"/>
                </AreaChart>
              </ResponsiveContainer>
            </Card>
          </div>
          <Card>
            <div style={{fontWeight:700,marginBottom:14,color:B.navy}}>Tabla de Tendencia Mensual</div>
            <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <thead><tr style={{background:B.cyanLL}}>{["Mes","Ingreso","Costo","Utilidad","Margen %","Pacs.","Folios","Ticket","∆ vs Anterior"].map(h=>(
                <th key={h} style={{padding:"10px 12px",textAlign:"left",fontSize:10,textTransform:"uppercase",color:B.cyanD,fontWeight:700}}>{h}</th>
              ))}</tr></thead>
              <tbody>{byMes.map((m,i)=>{
                const prev=byMes[i-1];const delta=prev?fDelta(m.cargo,prev.cargo):"—";
                return <tr key={i} style={{borderBottom:`1px solid ${B.borde}`}}>
                  <td style={{padding:"9px 12px",fontWeight:700,color:B.navy}}>{m.mes}</td>
                  <td style={{padding:"9px 12px",fontWeight:600,color:B.cyanD}}>{f$(m.cargo)}</td>
                  <td style={{padding:"9px 12px",color:B.rojo}}>{m.costo>0?f$(m.costo):"—"}</td>
                  <td style={{padding:"9px 12px",color:B.verde,fontWeight:600}}>{f$(m.cargo-m.costo)}</td>
                  <td style={{padding:"9px 12px"}}><span style={{color:parseFloat(m.margen)>=75?B.verde:B.rojo,fontWeight:600}}>{m.margen}%</span></td>
                  <td style={{padding:"9px 12px"}}>{fN(m.pacs)}</td>
                  <td style={{padding:"9px 12px"}}>{fN(m.fols)}</td>
                  <td style={{padding:"9px 12px"}}>{f$(m.ticket)}</td>
                  <td style={{padding:"9px 12px"}}><span style={{color:dC(delta),fontWeight:700,background:parseFloat(delta)>=0?B.verdeL:B.rojoL,padding:"2px 8px",borderRadius:10}}>{delta}</span></td>
                </tr>;
              })}</tbody>
            </table></div>
          </Card>
        </>}

        {/* ══ MÉDICOS / SERVICIOS / CONVENIOS ══ */}
        {(view==="medicos"||view==="servicios"||view==="convenios")&&(()=>{
          const data=view==="medicos"?byMed:view==="servicios"?byServ:byConv;
          const titulo=view==="medicos"?"Médico":view==="servicios"?"Servicio":"Convenio/Pagador";
          const color=view==="medicos"?B.cyan:view==="servicios"?B.teal:B.cyanD;
          const total=data.reduce((s,d)=>s+d.total,0);
          return <>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:20}}>
              <KpiCard label={`Total ${titulo}s`} value={fN(data.length)} color={color}/>
              <KpiCard label="Ingreso total" value={fM(total)} color={color}/>
              <KpiCard label="Concentración Top-3" value={total?fP(data.slice(0,3).reduce((s,d)=>s+d.total,0)/total):"—"} badge={data.slice(0,3).reduce((s,d)=>s+d.total,0)/Math.max(total,1)>0.5?"⚠ Alta":"Moderada"} badgeOk={data.slice(0,3).reduce((s,d)=>s+d.total,0)/Math.max(total,1)<=0.5} color={color}/>
              <KpiCard label="Ticket / Paciente" value={f$(K.ticket)} color={color}/>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"3fr 2fr",gap:16,marginBottom:16}}>
              <Card>
                <div style={{fontWeight:700,marginBottom:12,color:B.navy}}>Top 10 — {titulo}</div>
                <ResponsiveContainer width="100%" height={280}>
                  <BarChart data={data.slice(0,10)} layout="vertical">
                    <CartesianGrid strokeDasharray="3 3" stroke={B.borde}/>
                    <XAxis type="number" tickFormatter={v=>"$"+(v/1e6).toFixed(1)+"M"} tick={{fontSize:10}}/>
                    <YAxis type="category" dataKey="key" tick={{fontSize:9}} width={120}/>
                    <Tooltip content={<Tip/>}/><Bar dataKey="total" fill={color} radius={[0,6,6,0]} name="Ingreso"/>
                  </BarChart>
                </ResponsiveContainer>
              </Card>
              <Card>
                <div style={{fontWeight:700,marginBottom:12,color:B.navy}}>Distribución %</div>
                <ResponsiveContainer width="100%" height={280}>
                  <PieChart>
                    <Pie data={data.slice(0,6)} dataKey="total" nameKey="key" cx="50%" cy="50%" outerRadius={100} innerRadius={35}
                      label={({percent})=>percent>0.04?`${(percent*100).toFixed(0)}%`:""} labelLine={false}>
                      {data.slice(0,6).map((_,i)=><Cell key={i} fill={[B.cyan,B.teal,B.cyanD,B.amber,B.violeta,B.verde][i]}/>)}
                    </Pie>
                    <Tooltip formatter={v=>f$(v)}/><Legend iconSize={10} wrapperStyle={{fontSize:10}}/>
                  </PieChart>
                </ResponsiveContainer>
              </Card>
            </div>
            <Card><div style={{fontWeight:700,marginBottom:14,color:B.navy}}>Detalle por {titulo}</div><TablaDetalle data={data} titulo={titulo}/></Card>
          </>;
        })()}

        {/* ══ FINANCIERO ══ */}
        {view==="financiero"&&<>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:20}}>
            <KpiCard label="Ingreso Neto" value={fM(K.ingreso)} sub={`${fN(K.fol)} folios`} color={B.cyan}/>
            <KpiCard label="Costo Directo" value={fM(K.costo)} sub={fP(K.costo/Math.max(K.ingreso,1))+" del ingreso"} color={B.rojo}/>
            <KpiCard label="Utilidad Bruta" value={fM(K.utilidad)} badge={K.utilidad>0?"✓ Positivo":"⚠"} badgeOk={K.utilidad>0} color={B.teal}/>
            <KpiCard label="Margen Bruto" value={fP(K.margen)} badge={K.margen>0.75?"Sólido":"Revisar"} badgeOk={K.margen>0.75} color={B.cyanD}/>
          </div>
          <Card style={{marginBottom:16}}>
            <div style={{fontWeight:700,marginBottom:14,color:B.navy}}>Proyección — Próximos 6 Meses</div>
            <ResponsiveContainer width="100%" height={260}>
              <LineChart data={[...byMes.slice(-2).map(m=>({mes:m.mes,real:m.cargo})),...forecast]}>
                <CartesianGrid strokeDasharray="3 3" stroke={B.borde}/><XAxis dataKey="mes" tick={{fontSize:10}}/><YAxis tickFormatter={v=>"$"+(v/1e6).toFixed(1)+"M"} tick={{fontSize:10}}/>
                <Tooltip content={<Tip/>}/><Legend/>
                <Line type="monotone" dataKey="real" stroke={B.cyan} strokeWidth={3} dot={{r:5,fill:B.cyan}} name="Real"/>
                <Line type="monotone" dataKey="optimista" stroke={B.verde} strokeWidth={1.5} strokeDasharray="6 3" dot={{r:3}} name="Optimista"/>
                <Line type="monotone" dataKey="base" stroke={B.amber} strokeWidth={2} strokeDasharray="6 3" dot={{r:4}} name="Base"/>
                <Line type="monotone" dataKey="conservador" stroke={B.rojo} strokeWidth={1.5} strokeDasharray="6 3" dot={{r:3}} name="Conservador"/>
              </LineChart>
            </ResponsiveContainer>
          </Card>
        </>}

        {/* ══ ALERTAS ══ */}
        {view==="alertas"&&<div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:16}}>
          <div>
            <div style={{fontWeight:700,fontSize:17,color:B.navy,marginBottom:16}}>🔔 Alertas Automáticas</div>
            {alertas.map((a,i)=><Alerta key={i} {...a}/>)}
          </div>
          <Card>
            <div style={{fontWeight:700,marginBottom:14,color:B.navy}}>⚙️ Umbrales</div>
            {[{k:"margen",l:"Margen mínimo %"},{k:"concentracion",l:"Concentración máx. top-3 %"}].map(u=>(
              <div key={u.k} style={{marginBottom:18}}>
                <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:8}}>
                  <span style={{color:B.gris}}>{u.l}</span>
                  <strong style={{color:B.cyan}}>{umbral[u.k]}%</strong>
                </div>
                <input type="range" min={0} max={100} value={umbral[u.k]} onChange={e=>setUmbral(p=>({...p,[u.k]:parseInt(e.target.value)}))} style={{width:"100%",accentColor:B.cyan}}/>
              </div>
            ))}
          </Card>
        </div>}

        {/* ══ IA ══ */}
        {view==="ia"&&<div style={{display:"grid",gridTemplateColumns:"1fr 300px",gap:16}}>
          <Card>
            <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:4}}>
              <ECGLine width={80} height={20} color={B.cyan}/>
              <div style={{fontWeight:700,fontSize:16,color:B.navy}}>Análisis con Inteligencia Artificial</div>
            </div>
            <div style={{fontSize:12,color:B.grisM,marginBottom:16}}>Consulta tus datos de Punta Médica en lenguaje natural</div>
            <div style={{height:380,overflowY:"auto",border:`1px solid ${B.borde}`,borderRadius:12,padding:16,marginBottom:10,background:B.cyanLL}}>
              {aiH.length===0?<div style={{textAlign:"center",color:B.grisM,marginTop:80}}>
                <div style={{marginBottom:12}}><ECGLine width={120} height={30} color={B.cyan}/></div>
                <div style={{fontSize:13}}>Pregunta cualquier cosa sobre los datos del hospital</div>
              </div>:aiH.map((m,i)=>(
                <div key={i} style={{marginBottom:14,display:"flex",gap:8,flexDirection:m.role==="user"?"row-reverse":"row"}}>
                  <div style={{width:28,height:28,borderRadius:"50%",background:m.role==="user"?B.navy:B.cyan,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,color:"white",fontWeight:700,flexShrink:0}}>{m.role==="user"?"U":"IA"}</div>
                  <div style={{background:m.role==="user"?B.navy:"white",color:m.role==="user"?"white":B.navy,padding:"10px 14px",borderRadius:m.role==="user"?"12px 12px 4px 12px":"12px 12px 12px 4px",fontSize:12,lineHeight:1.65,maxWidth:"82%",border:m.role==="assistant"?`1px solid ${B.borde}`:"none",whiteSpace:"pre-wrap"}}>{m.content}</div>
                </div>
              ))}
              {aiL&&<div style={{display:"flex",gap:8}}><div style={{width:28,height:28,borderRadius:"50%",background:B.cyan,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,color:B.navy,fontWeight:700}}>IA</div><div style={{background:"white",padding:"10px 14px",borderRadius:"12px 12px 12px 4px",border:`1px solid ${B.borde}`,color:B.grisM,fontSize:12}}>Analizando datos Punta Médica...</div></div>}
            </div>
            <div style={{display:"flex",gap:8}}>
              <input value={aiQ} onChange={e=>setAiQ(e.target.value)} onKeyDown={e=>e.key==="Enter"&&!e.shiftKey&&askAI()} placeholder="Pregunta sobre los datos del hospital..." style={{flex:1,padding:"10px 14px",border:`1px solid ${B.borde}`,borderRadius:8,fontSize:12,outline:"none"}}/>
              <button onClick={askAI} disabled={aiL||!aiQ.trim()} style={{padding:"10px 20px",background:B.cyan,color:B.navy,border:"none",borderRadius:8,fontWeight:700,cursor:"pointer",fontSize:12,letterSpacing:"0.05em"}}>ENVIAR</button>
            </div>
          </Card>
          <div style={{display:"flex",flexDirection:"column",gap:12}}>
            <Card>
              <div style={{fontWeight:700,marginBottom:10,color:B.navy}}>Preguntas sugeridas</div>
              {["¿Cuál es el médico con mayor ingreso y qué % representa?","Resume los resultados para los accionistas","¿Qué convenio tiene mayor riesgo de cartera?","Identifica los 3 principales riesgos del período","Genera el resumen ejecutivo para NAFIN","¿Qué servicios tienen el margen más bajo?","Compara el ticket promedio por convenio"].map((q,i)=>(
                <button key={i} onClick={()=>setAiQ(q)} style={{width:"100%",textAlign:"left",padding:"8px 10px",marginBottom:5,border:`1px solid ${B.borde}`,borderRadius:8,background:"none",cursor:"pointer",fontSize:11,color:B.gris,lineHeight:1.4,transition:"border-color 0.15s",fontFamily:"inherit"}} onMouseEnter={e=>e.currentTarget.style.borderColor=B.cyan} onMouseLeave={e=>e.currentTarget.style.borderColor=B.borde}>{q}</button>
              ))}
            </Card>
          </div>
        </div>}

      </div>
    </div>
  );
}
