import { useState, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, AreaChart, Area,
  PieChart, Pie, Cell, ComposedChart,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer
} from "recharts";

// ── COLORES MARCA ──────────────────────────────────────────
const C = {
  cyan:"#00B4D8",cyanD:"#006D8F",cyanDD:"#003D52",navy:"#003D52",
  cyanL:"#E0F7FA",cyanLL:"#F0FBFD",teal:"#0097A7",
  verde:"#00897B",verdeL:"#E0F2F1",
  amber:"#F59E0B",amberL:"#FEF3C7",
  rojo:"#E53935",rojoL:"#FFEBEE",
  violeta:"#5C6BC0",violetaL:"#E8EAF6",
  gris:"#37474F",grisM:"#78909C",grisL:"#F8FAFB",
  borde:"#E0E8EB",blanco:"#FFFFFF",
  ch:["#00B4D8","#0097A7","#006D8F","#F59E0B","#E53935","#00897B","#5C6BC0","#EA580C","#16A34A","#EC4899"]
};

// ── CANALES (lógica de negocio HPM) ────────────────────────
// Censo histórico solo tiene: Hospitalización y Urgencias
// Pacientes en ECP pero NO en censo = Servicios Externos (IMSS - Dra. Sara Ramos)
const CANALES = {
  hospitalizacion: { label:"Hospitalización", color:"#0097A7", icon:"🏥" },
  urgencias:       { label:"Urgencias",        color:"#E53935", icon:"🚨" },
  externos:        { label:"Servicios Externos (IMSS)", color:"#5C6BC0", icon:"🔬" },
};

// ── UTILS ──────────────────────────────────────────────────
const f$ = n => !n&&n!==0?"—":"$"+Math.round(n).toLocaleString("es-MX");
const fM = n => !n&&n!==0?"—":"$"+(n/1e6).toFixed(2)+"M";
const fP = n => !n&&n!==0?"—":(n*100).toFixed(1)+"%";
const fN = n => !n&&n!==0?"—":Math.round(n).toLocaleString("es-MX");
const fD = (a,b) => !b?"—":`${a>=b?"+":""}${((a/b-1)*100).toFixed(1)}%`;
const dC = v => { const n=parseFloat(v); return isNaN(n)?C.grisM:n>=0?C.verde:C.rojo; };

function detectTipo(headers) {
  const h = headers.map(x=>String(x||"").toLowerCase());
  if (h.some(x=>x.includes("cirug")||x.includes("cirujano"))) return "cirugias";
  if (h.some(x=>x.includes("existencia")||x.includes("inventario"))) return "inventario";
  if (h.some(x=>x.includes("hospitali")||x.includes("censo")||x.includes("egreso"))) return "censo";
  if (h.some(x=>x.includes("costo")||x.includes("cost"))) return "cargos_costo";
  return "cargos";
}

function autoMap(headers) {
  const find=kws=>{
    const h=headers.map(x=>String(x||"").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,""));
    for(const k of kws){const i=h.findIndex(x=>x.includes(k));if(i>=0)return headers[i];}
    return null;
  };
  return {
    fecha:    find(["fecha","date","periodo","mes","fec","ingreso","entrada"]),
    medico:   find(["medico","doctor","cirujano","med_","nombre_med","especialista"]),
    paciente: find(["nombre_pac","paciente","patient","pac_","nombre_p"]),
    servicio: find(["servicio","ubicacion","depto","area","tipo","especialidad"]),
    cargo:    find(["cargo","importe","monto","total_cargo","precio","charge"]),
    costo:    find(["costo","cost","coste","precio_costo"]),
    convenio: find(["convenio","seguro","payer","aseguradora","pagador","tipo_pago"]),
    folio:    find(["folio","ecp","id_ecp","expediente","id_ti","num_exp","id_pac"]),
    tipo_atencion: find(["tipo_atencion","tipo_ingreso","motivo","urgencia","hospitali","tipo_egreso"]),
    id_paciente: find(["id_paciente","id_pac","paciente_id","num_paciente"]),
  };
}

function parseDate(v){
  if(!v)return null;
  const s=String(v);
  if(/^\d{5}$/.test(s)){const d=new Date(Date.UTC(1899,11,30)+parseInt(s)*86400000);return d.toISOString().substring(0,10);}
  if(/^\d{4}-\d{2}-\d{2}/.test(s))return s.substring(0,10);
  if(/^\d{2}\/\d{2}\/\d{4}/.test(s)){const[d,m,y]=s.split("/");return`${y}-${m}-${d}`;}
  return s.substring(0,10);
}
const getMes = f=>f?String(f).substring(0,7):"Sin fecha";

// Detectar canal de atención desde el censo
function detectarCanal(tipoAtencion, servicio) {
  const t = String(tipoAtencion||"").toLowerCase();
  const s = String(servicio||"").toLowerCase();
  if (t.includes("urgencia")||t.includes("emerg")||s.includes("urgencia")||s.includes("emerg")) return "urgencias";
  if (t.includes("hospit")||t.includes("intern")||s.includes("hospit")||s.includes("encama")) return "hospitalizacion";
  return "hospitalizacion"; // default censo = hospitalización
}

// ── COMPONENTES UI ─────────────────────────────────────────
const Card = ({children,style={}}) => (
  <div style={{background:C.blanco,border:`1px solid ${C.borde}`,borderRadius:12,padding:"16px 20px",...style}}>
    {children}
  </div>
);

const KpiCard = ({label,value,sub,delta,color=C.cyan,badge,badgeOk=true,icon}) => (
  <div style={{background:C.blanco,border:`1px solid ${C.borde}`,borderRadius:12,padding:"16px 20px",borderTop:`3px solid ${color}`}}>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:6}}>
      <span style={{fontSize:10,textTransform:"uppercase",letterSpacing:"0.1em",color:C.grisM,fontWeight:600}}>{label}</span>
      {icon&&<span style={{fontSize:18,opacity:0.7}}>{icon}</span>}
    </div>
    <div style={{fontFamily:"Georgia,serif",fontSize:28,fontWeight:700,color:C.navy,lineHeight:1}}>{value}</div>
    <div style={{display:"flex",alignItems:"center",gap:8,marginTop:8}}>
      {sub&&<span style={{fontSize:11,color:C.grisM}}>{sub}</span>}
      {delta&&<span style={{fontSize:11,fontWeight:700,color:dC(delta),background:parseFloat(delta)>=0?C.verdeL:C.rojoL,padding:"1px 7px",borderRadius:10}}>{delta}</span>}
    </div>
    {badge&&<span style={{display:"inline-block",marginTop:8,padding:"3px 10px",borderRadius:20,fontSize:10,fontWeight:600,background:badgeOk?C.cyanL:C.rojoL,color:badgeOk?C.cyanD:C.rojo}}>{badge}</span>}
  </div>
);

const Tip = ({active,payload,label}) => {
  if(!active||!payload||!payload.length)return null;
  return (
    <div style={{background:C.blanco,border:`1px solid ${C.borde}`,borderRadius:8,padding:"10px 14px",fontSize:12,boxShadow:"0 4px 12px rgba(0,0,0,0.1)"}}>
      <div style={{fontWeight:700,marginBottom:6,color:C.navy}}>{label}</div>
      {payload.map((p,i)=>(
        <div key={i} style={{color:p.color,marginBottom:2}}>{p.name}: <strong>{typeof p.value==="number"&&p.value>100?f$(p.value):p.value}</strong></div>
      ))}
    </div>
  );
};

const Logo = ({customLogo,brandName="Hospital Punta Médica",size=24}) => (
  customLogo
    ? <img src={customLogo} alt={brandName} style={{height:size,width:"auto",maxWidth:160,objectFit:"contain"}}/>
    : <div style={{display:"flex",alignItems:"center",gap:6}}>
        <svg width={size} height={size} viewBox="0 0 32 32">
          <rect width="32" height="32" rx="5" fill="rgba(0,180,216,0.2)"/>
          <path d="M4 16 L7 16 L9 9 L12 23 L15 12 L17 16 L28 16" stroke="#00B4D8" strokeWidth="2.2" fill="none" strokeLinecap="round" strokeLinejoin="round"/>
        </svg>
        <span style={{fontFamily:"Georgia,serif",fontWeight:700,fontSize:size*0.55,color:"#00B4D8",letterSpacing:"0.02em"}}>{brandName}</span>
      </div>
);

// ── ROLES ──────────────────────────────────────────────────
const ROLES = {
  accionista: {label:"Accionista",      icon:"👔",views:["dashboard","financiero","canales","tendencias","alertas"]},
  director:   {label:"Director General",icon:"🏛",views:["dashboard","financiero","canales","medicos","servicios","convenios","tendencias","alertas","ia"]},
  gerente:    {label:"Gerente Operativo",icon:"📋",views:["dashboard","canales","medicos","servicios","convenios","tendencias","alertas"]},
  analista:   {label:"Analista",        icon:"📊",views:["financiero","canales","medicos","servicios","convenios","tendencias","ia"]},
};

const VISTAS = {
  dashboard:  {l:"Dashboard", icon:"📊"},
  financiero: {l:"Financiero",icon:"💰"},
  canales:    {l:"Canales",   icon:"🔀"},
  medicos:    {l:"Médicos",   icon:"👨‍⚕️"},
  servicios:  {l:"Servicios", icon:"🏥"},
  convenios:  {l:"Convenios", icon:"📄"},
  tendencias: {l:"Tendencias",icon:"📈"},
  alertas:    {l:"Alertas",   icon:"🔔"},
  ia:         {l:"IA",        icon:"✨"},
};

// ── PANTALLA CARGA ─────────────────────────────────────────
function Carga({onData}) {
  const [drag,setDrag]=useState(false);
  const [loading,setLoading]=useState(false);
  const [prog,setProg]=useState([]);
  const ref=useRef();

  const procesar=async(files)=>{
    if(!files||files.length===0)return;
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
          setProg(p=>p.map(x=>x.nombre===f.name?{...x,estado:`✅ ${rows.length.toLocaleString()} filas · ${tipo}`}:x));
        }
      }catch(e){
        setProg(p=>p.map(x=>x.nombre===f.name?{...x,estado:"❌ Error al leer"}:x));
      }
    }
    setLoading(false);
    if(fuentes.length>0)onData(fuentes);
  };

  return (
    <div style={{minHeight:"100vh",background:C.navy,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:40}}>
      <div style={{maxWidth:660,width:"100%"}}>
        <div style={{textAlign:"center",marginBottom:36}}>
          <div style={{fontSize:30,fontWeight:700,color:C.cyan,fontFamily:"Georgia,serif",marginBottom:8}}>Hospital Punta Médica</div>
          <svg width="280" height="20" viewBox="0 0 280 20" style={{display:"block",margin:"0 auto"}}>
            <path d="M0 10 L56 10 L70 10 L76 2 L84 18 L90 6 L96 10 L140 10 L196 10 L206 2 L214 18 L220 6 L226 10 L280 10"
              stroke={C.cyan} strokeWidth="1.5" fill="none" strokeLinecap="round" opacity="0.6"/>
          </svg>
          <div style={{color:"rgba(255,255,255,0.4)",fontSize:11,letterSpacing:"0.15em",textTransform:"uppercase",marginTop:10}}>
            Centro de Inteligencia Empresarial
          </div>
        </div>
        <div onDrop={e=>{e.preventDefault();setDrag(false);procesar([...e.dataTransfer.files])}}
          onDragOver={e=>{e.preventDefault();setDrag(true)}} onDragLeave={()=>setDrag(false)}
          onClick={()=>!loading&&ref.current&&ref.current.click()}
          style={{border:`2px dashed ${drag?C.cyan:"rgba(0,180,216,0.3)"}`,borderRadius:16,padding:"36px 28px",textAlign:"center",cursor:"pointer",background:drag?"rgba(0,180,216,0.06)":"rgba(255,255,255,0.03)",marginBottom:18}}>
          <input ref={ref} type="file" multiple accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={e=>procesar([...e.target.files])}/>
          {loading?(
            <div>
              <div style={{fontSize:13,color:C.cyan,fontWeight:500,marginBottom:14}}>Procesando archivos ECP...</div>
              {prog.map((p,i)=><div key={i} style={{fontSize:12,color:"rgba(255,255,255,0.5)",marginBottom:5}}><span style={{color:"rgba(255,255,255,0.8)"}}>{p.nombre}</span> → {p.estado}</div>)}
            </div>
          ):(
            <>
              <div style={{fontSize:36,marginBottom:10}}>📂</div>
              <div style={{fontSize:16,fontWeight:600,color:"white",marginBottom:6}}>Arrastra tus reportes ECP aquí</div>
              <div style={{fontSize:13,color:"rgba(255,255,255,0.4)",marginBottom:18}}>.xlsx · .xls · .csv · Múltiples archivos simultáneos</div>
              <div style={{display:"inline-block",background:C.cyan,color:C.navy,padding:"10px 28px",borderRadius:8,fontSize:13,fontWeight:700}}>SELECCIONAR ARCHIVOS</div>
            </>
          )}
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:10,marginBottom:12}}>
          <div style={{background:"rgba(0,180,216,0.08)",border:"1px solid rgba(0,180,216,0.2)",borderRadius:12,padding:"14px 16px"}}>
            <div style={{fontWeight:700,color:C.cyan,fontSize:13,marginBottom:6}}>📋 Archivos de Cargos</div>
            <div style={{fontSize:11,color:"rgba(255,255,255,0.5)",lineHeight:1.6}}>
              Admisión → Reporte General de Cargos con Costo<br/>
              (uno o varios archivos quincenales)
            </div>
          </div>
          <div style={{background:"rgba(92,107,192,0.08)",border:"1px solid rgba(92,107,192,0.2)",borderRadius:12,padding:"14px 16px"}}>
            <div style={{fontWeight:700,color:"#7986CB",fontSize:13,marginBottom:6}}>👥 Censo Histórico</div>
            <div style={{fontSize:11,color:"rgba(255,255,255,0.5)",lineHeight:1.6}}>
              Admisión → Censo Histórico de Pacientes<br/>
              (identifica Hospitalización y Urgencias)
            </div>
          </div>
        </div>
        <div style={{background:"rgba(92,107,192,0.1)",border:"1px solid rgba(92,107,192,0.3)",borderRadius:10,padding:"10px 16px",fontSize:11,color:"rgba(255,255,255,0.6)",lineHeight:1.6}}>
          <strong style={{color:"#7986CB"}}>🔬 Lógica automática:</strong> Pacientes en Cargos pero <strong>NO</strong> en Censo → Servicios Externos (Imagenología, Laboratorio, Radioterapia) → Convenio IMSS → Dra. Sara Ramos
        </div>
      </div>
    </div>
  );
}

// ── SELECTOR ROL ───────────────────────────────────────────
function SelectorRol({onSelect,customLogo,brandName}) {
  return (
    <div style={{minHeight:"100vh",background:C.cyanLL,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:40}}>
      <div style={{maxWidth:520,width:"100%",textAlign:"center"}}>
        <div style={{marginBottom:8,display:"flex",justifyContent:"center"}}><Logo customLogo={customLogo} brandName={brandName} size={32}/></div>
        <svg width="200" height="16" viewBox="0 0 200 16" style={{display:"block",margin:"8px auto 28px"}}>
          <path d="M0 8 L40 8 L46 8 L50 2 L56 14 L60 5 L64 8 L100 8 L140 8 L150 2 L156 14 L160 5 L164 8 L200 8"
            stroke={C.cyan} strokeWidth="1.5" fill="none" strokeLinecap="round" opacity="0.7"/>
        </svg>
        <div style={{fontFamily:"Georgia,serif",fontSize:26,fontWeight:700,color:C.navy,marginBottom:6}}>¿Cuál es tu perfil?</div>
        <div style={{fontSize:13,color:C.grisM,marginBottom:28}}>El dashboard se personaliza según tu nivel de acceso</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
          {Object.entries(ROLES).map(([k,v])=>(
            <button key={k} onClick={()=>onSelect(k)}
              style={{background:C.blanco,border:`2px solid ${C.borde}`,borderRadius:14,padding:"22px 18px",cursor:"pointer",textAlign:"left",transition:"all 0.2s"}}
              onMouseEnter={e=>{e.currentTarget.style.borderColor=C.cyan;e.currentTarget.style.boxShadow="0 6px 20px rgba(0,180,216,0.15)";}}
              onMouseLeave={e=>{e.currentTarget.style.borderColor=C.borde;e.currentTarget.style.boxShadow="none";}}>
              <div style={{fontSize:26,marginBottom:8}}>{v.icon}</div>
              <div style={{fontWeight:700,fontSize:14,color:C.navy,marginBottom:4}}>{v.label}</div>
              <div style={{width:28,height:2,background:C.cyan,borderRadius:1,marginBottom:8}}/>
              <div style={{fontSize:10,color:C.grisM}}>{v.views.length} módulos</div>
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}

// ── MODAL AJUSTES ─────────────────────────────────────────
function ModalAjustes({onClose,customLogo,setCustomLogo,brandName,setBrandName,brandSub,setBrandSub}) {
  const logoRef=useRef();
  return (
    <div onClick={e=>{if(e.target===e.currentTarget)onClose();}}
      style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:999,display:"flex",alignItems:"center",justifyContent:"center",padding:24}}>
      <div style={{background:C.blanco,borderRadius:18,width:"100%",maxWidth:440,overflow:"hidden",boxShadow:"0 24px 60px rgba(0,0,0,0.25)"}}>
        <div style={{background:C.navy,padding:"18px 22px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke={C.cyan} strokeWidth="2" strokeLinecap="round">
              <circle cx="12" cy="12" r="3"/>
              <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/>
            </svg>
            <span style={{color:"white",fontWeight:700,fontSize:15}}>Ajustes</span>
          </div>
          <button onClick={onClose} style={{background:"rgba(255,255,255,0.1)",border:"none",borderRadius:6,width:28,height:28,cursor:"pointer",color:"white",fontSize:16}}>✕</button>
        </div>
        <div style={{padding:24}}>
          <div style={{marginBottom:20}}>
            <div style={{fontSize:11,textTransform:"uppercase",letterSpacing:"0.1em",color:C.grisM,fontWeight:600,marginBottom:10}}>Logotipo</div>
            <input ref={logoRef} type="file" accept="image/*" style={{display:"none"}}
              onChange={e=>{const file=e.target.files[0];if(!file)return;const reader=new FileReader();reader.onload=ev=>setCustomLogo(ev.target.result);reader.readAsDataURL(file);}}/>
            <div style={{display:"flex",gap:12,alignItems:"center"}}>
              <div style={{width:150,height:60,border:`2px dashed ${customLogo?C.cyan:C.borde}`,borderRadius:10,display:"flex",alignItems:"center",justifyContent:"center",background:customLogo?C.navy:C.cyanLL,overflow:"hidden",flexShrink:0}}>
                {customLogo?<img src={customLogo} alt="Logo" style={{maxHeight:50,maxWidth:140,objectFit:"contain"}}/>:<span style={{fontSize:11,color:C.grisM,textAlign:"center"}}>Sin logo</span>}
              </div>
              <div style={{display:"flex",flexDirection:"column",gap:8}}>
                <button onClick={()=>logoRef.current&&logoRef.current.click()} style={{padding:"8px 14px",background:C.cyan,color:C.navy,border:"none",borderRadius:7,fontWeight:700,fontSize:12,cursor:"pointer"}}>⬆ Subir logo</button>
                {customLogo&&<button onClick={()=>setCustomLogo(null)} style={{padding:"7px 14px",background:"none",color:C.rojo,border:`1px solid ${C.rojo}`,borderRadius:7,fontSize:12,cursor:"pointer"}}>✕ Quitar</button>}
                <span style={{fontSize:10,color:C.grisM}}>PNG · JPG · SVG</span>
              </div>
            </div>
          </div>
          <div style={{marginBottom:16}}>
            <label style={{fontSize:11,textTransform:"uppercase",letterSpacing:"0.1em",color:C.grisM,fontWeight:600,display:"block",marginBottom:6}}>Nombre</label>
            <input value={brandName} onChange={e=>setBrandName(e.target.value)} style={{width:"100%",padding:"9px 12px",border:`1px solid ${C.borde}`,borderRadius:7,fontSize:13,color:C.navy,outline:"none",fontFamily:"inherit"}}/>
          </div>
          <div style={{marginBottom:22}}>
            <label style={{fontSize:11,textTransform:"uppercase",letterSpacing:"0.1em",color:C.grisM,fontWeight:600,display:"block",marginBottom:6}}>Subtítulo</label>
            <input value={brandSub} onChange={e=>setBrandSub(e.target.value)} style={{width:"100%",padding:"9px 12px",border:`1px solid ${C.borde}`,borderRadius:7,fontSize:13,color:C.navy,outline:"none",fontFamily:"inherit"}}/>
          </div>
          <div style={{display:"flex",gap:10}}>
            <button onClick={onClose} style={{flex:1,padding:"11px",background:C.cyan,color:C.navy,border:"none",borderRadius:9,fontWeight:700,fontSize:13,cursor:"pointer"}}>✓ Guardar</button>
            <button onClick={()=>{setCustomLogo(null);setBrandName("Hospital Punta Médica");setBrandSub("Alta Especialidad");}} style={{padding:"11px 14px",background:"none",color:C.grisM,border:`1px solid ${C.borde}`,borderRadius:9,fontSize:12,cursor:"pointer"}}>Restaurar</button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ── COMPONENTE PRINCIPAL ───────────────────────────────────
export default function HPM() {
  const [fuentes,setFuentes]         = useState(null);
  const [rol,setRol]                 = useState(null);
  const [view,setView]               = useState("dashboard");
  const [filtros,setFiltros]         = useState({mes:"Todos",canal:"Todos",servicio:"Todos",convenio:"Todos",medico:"Todos"});
  const [showSettings,setShowSettings] = useState(false);
  const [customLogo,setCustomLogo]   = useState(null);
  const [brandName,setBrandName]     = useState("Hospital Punta Médica");
  const [brandSub,setBrandSub]       = useState("Alta Especialidad");
  const [umbral,setUmbral]           = useState({margen:75,concentracion:40});
  const [aiQ,setAiQ]                 = useState("");
  const [aiH,setAiH]                 = useState([]);
  const [aiL,setAiL]                 = useState(false);

  // ── Construir Set de IDs del censo (hospitalizados + urgencias) ──
  const censoMap = useMemo(() => {
    if (!fuentes) return new Map();
    const map = new Map();
    for (const f of fuentes) {
      if (f.tipo !== "censo") continue;
      for (const r of f.rows) {
        // Intentar obtener ID del paciente por múltiples columnas posibles
        const idCols = [f.mapa.id_paciente, f.mapa.folio, f.mapa.paciente].filter(Boolean);
        for (const col of idCols) {
          const id = String(r[col]||"").trim().toUpperCase();
          if (id && id.length > 2) {
            const canal = detectarCanal(r[f.mapa.tipo_atencion], r[f.mapa.servicio]);
            map.set(id, canal);
          }
        }
        // También por nombre de paciente
        if (f.mapa.paciente) {
          const nombre = String(r[f.mapa.paciente]||"").trim().toUpperCase();
          if (nombre && nombre.length > 3) {
            const canal = detectarCanal(r[f.mapa.tipo_atencion], r[f.mapa.servicio]);
            if (!map.has(nombre)) map.set(nombre, canal);
          }
        }
      }
    }
    return map;
  }, [fuentes]);

  const hayCenso = useMemo(() => fuentes?.some(f=>f.tipo==="censo"), [fuentes]);

  // ── Normalizar todas las filas ──
  const allRows = useMemo(() => {
    if (!fuentes) return [];
    const rows = [];
    for (const f of fuentes) {
      if (f.tipo === "censo") continue; // censo solo es referencia
      for (const r of f.rows) {
        const cargo    = Number(r[f.mapa.cargo])    || 0;
        const costo    = Number(r[f.mapa.costo])    || 0;
        const medico   = f.mapa.medico   ? String(r[f.mapa.medico]  ||"Sin asignar").trim() : "Sin asignar";
        const paciente = f.mapa.paciente ? String(r[f.mapa.paciente]||"").trim()            : "";
        const servicio = f.mapa.servicio ? String(r[f.mapa.servicio]||"General").trim()     : "General";
        const convenio = f.mapa.convenio ? String(r[f.mapa.convenio]||"PARTICULAR").trim()  : "PARTICULAR";
        const folio    = f.mapa.folio    ? String(r[f.mapa.folio]   ||"")                   : "";
        const fecha    = f.mapa.fecha    ? parseDate(r[f.mapa.fecha]) : null;

        // ── LÓGICA CANAL ──────────────────────────────────────
        // 1. Si hay censo cargado: buscar por nombre o folio
        // 2. Si está en censo → hospitalizacion o urgencias
        // 3. Si NO está en censo → externos (IMSS - Dra. Sara Ramos)
        let canal = "externos";
        if (hayCenso) {
          const keyNombre = paciente.toUpperCase();
          const keyFolio  = folio.toUpperCase();
          if (censoMap.has(keyFolio))   canal = censoMap.get(keyFolio);
          else if (censoMap.has(keyNombre)) canal = censoMap.get(keyNombre);
          else canal = "externos"; // no está en censo → IMSS / externos
        } else {
          // Sin censo: inferir por servicio/convenio
          const s = servicio.toLowerCase();
          const c = convenio.toLowerCase();
          if (s.includes("urgencia")||s.includes("emerg")) canal = "urgencias";
          else if (s.includes("hospit")||s.includes("encama")||s.includes("intern")) canal = "hospitalizacion";
          else if (c.includes("imss")||s.includes("imagen")||s.includes("lab")||s.includes("radio")||s.includes("tomog")) canal = "externos";
          else canal = "hospitalizacion";
        }

        // Si es externos → convenio = IMSS, médico responsable = Dra. Sara Ramos (si no hay médico)
        const convenioFinal = canal === "externos"
          ? (convenio.includes("IMSS")||convenio.includes("imss") ? convenio : "IMSS - Servicios Externos")
          : convenio;

        rows.push({
          _cargo:cargo, _costo:costo, _medico:medico, _paciente:paciente,
          _servicio:servicio, _convenio:convenioFinal, _folio:folio,
          _fecha:fecha, _mes:getMes(fecha), _canal:canal,
          _canalLabel: CANALES[canal]?.label || canal,
        });
      }
    }
    return rows;
  }, [fuentes, censoMap, hayCenso]);

  // ── Opciones filtros ──
  const opts = useMemo(() => ({
    meses:     [...new Set(allRows.map(r=>r._mes).filter(x=>x&&x!=="Sin fecha"))].sort(),
    canales:   [...new Set(allRows.map(r=>r._canalLabel).filter(Boolean))].sort(),
    servicios: [...new Set(allRows.map(r=>r._servicio).filter(Boolean))].sort(),
    convenios: [...new Set(allRows.map(r=>r._convenio).filter(Boolean))].sort(),
    medicos:   [...new Set(allRows.map(r=>r._medico).filter(Boolean))].sort(),
  }), [allRows]);

  // ── Filtrar ──
  const rows = useMemo(() => allRows.filter(r =>
    (filtros.mes==="Todos"      || r._mes===filtros.mes) &&
    (filtros.canal==="Todos"    || r._canalLabel===filtros.canal) &&
    (filtros.servicio==="Todos" || r._servicio===filtros.servicio) &&
    (filtros.convenio==="Todos" || r._convenio===filtros.convenio) &&
    (filtros.medico==="Todos"   || r._medico===filtros.medico)
  ), [allRows, filtros]);

  // ── KPIs globales ──
  const K = useMemo(() => {
    const ing  = rows.reduce((s,r)=>s+r._cargo,0);
    const cos  = rows.reduce((s,r)=>s+r._costo,0);
    const pacs = new Set(rows.map(r=>r._paciente).filter(Boolean)).size;
    const fols = new Set(rows.map(r=>r._folio).filter(Boolean)).size;
    return {ing,cos,ut:ing-cos,mg:ing>0?(ing-cos)/ing:0,pacs,fols,ev:rows.length,tick:pacs>0?ing/pacs:0};
  }, [rows]);

  // ── KPIs por canal ──
  const porCanal = useMemo(() => {
    const m = {};
    for (const r of rows) {
      const c = r._canal;
      if (!m[c]) m[c] = {canal:c,label:CANALES[c]?.label||c,color:CANALES[c]?.color||C.grisM,icon:CANALES[c]?.icon||"",ing:0,cos:0,ev:0,pacs:new Set()};
      m[c].ing += r._cargo; m[c].cos += r._costo; m[c].ev++;
      if (r._paciente) m[c].pacs.add(r._paciente);
    }
    return Object.values(m).map(x=>({...x,pacs:x.pacs.size,mg:x.ing>0?(x.ing-x.cos)/x.ing:0,tick:x.pacs>0?Math.round(x.ing/x.pacs):0})).sort((a,b)=>b.ing-a.ing);
  }, [rows]);

  // ── Agregaciones ──
  const agg = (campo) => {
    const m={};
    for(const r of rows){
      const k=r[campo]||"Sin datos";
      if(!m[k])m[k]={key:k,total:0,costo:0,count:0,pacs:new Set(),fols:new Set()};
      m[k].total+=r._cargo;m[k].costo+=r._costo;m[k].count++;
      if(r._paciente)m[k].pacs.add(r._paciente);
      if(r._folio)m[k].fols.add(r._folio);
    }
    return Object.values(m).map(x=>({...x,pacs:x.pacs.size,fols:x.fols.size,mg:x.total>0?(x.total-x.costo)/x.total:0})).sort((a,b)=>b.total-a.total);
  };

  const byMed  = useMemo(()=>agg("_medico"),  [rows]);
  const byServ = useMemo(()=>agg("_servicio"), [rows]);
  const byConv = useMemo(()=>agg("_convenio"), [rows]);

  const byMes = useMemo(()=>{
    const m={};
    for(const r of allRows){
      const k=r._mes||"S/F";
      if(!m[k])m[k]={mes:k,cargo:0,costo:0,ev:0,pacs:new Set(),fols:new Set(),hosp:0,urg:0,ext:0};
      m[k].cargo+=r._cargo;m[k].costo+=r._costo;m[k].ev++;
      if(r._paciente)m[k].pacs.add(r._paciente);
      if(r._folio)m[k].fols.add(r._folio);
      if(r._canal==="hospitalizacion")m[k].hosp+=r._cargo;
      else if(r._canal==="urgencias")m[k].urg+=r._cargo;
      else m[k].ext+=r._cargo;
    }
    return Object.values(m).map(x=>({...x,pacs:x.pacs.size,fols:x.fols.size,mg:x.cargo>0?((x.cargo-x.costo)/x.cargo*100).toFixed(1):0,tick:x.pacs>0?Math.round(x.cargo/x.pacs):0})).sort((a,b)=>a.mes>b.mes?1:-1);
  },[allRows]);

  const forecast=useMemo(()=>{
    if(byMes.length<2)return[];
    const ult=byMes.slice(-3);
    const avg=ult.reduce((s,m)=>s+m.cargo,0)/ult.length;
    const g=byMes.length>1?(byMes[byMes.length-1].cargo/byMes[0].cargo-1)/(byMes.length-1):0.058;
    return["M+1","M+2","M+3","M+4","M+5","M+6"].map((mes,i)=>({mes,
      conservador:Math.round(avg*Math.pow(1+g*0.85,i+1)),
      base:Math.round(avg*Math.pow(1+g,i+1)),
      optimista:Math.round(avg*Math.pow(1+g*1.15,i+1)),
    }));
  },[byMes]);

  const alertas=useMemo(()=>{
    const a=[];
    if(K.mg*100<umbral.margen)a.push({tipo:"critico",titulo:"Margen Bruto bajo",msg:`Actual ${fP(K.mg)} — umbral ${umbral.margen}%`});
    const t3=byMed.slice(0,3).reduce((s,m)=>s+m.total,0);
    if(K.ing>0&&t3/K.ing*100>umbral.concentracion)a.push({tipo:"alerta",titulo:"Alta concentración médica",msg:`Top-3 generan ${fP(t3/K.ing)} del ingreso`});
    const ext=porCanal.find(c=>c.canal==="externos");
    if(ext&&K.ing>0&&ext.ing/K.ing>0.4)a.push({tipo:"alerta",titulo:"Servicios Externos >40% del ingreso",msg:`IMSS (Dra. Sara Ramos) representa ${fP(ext.ing/K.ing)} — verificar facturación`});
    if(!hayCenso)a.push({tipo:"alerta",titulo:"Sin Censo Histórico cargado",msg:"Sube el Censo Histórico para clasificar Hospitalización / Urgencias / Externos con precisión"});
    if(a.length===0)a.push({tipo:"ok",titulo:"Todo dentro de parámetros",msg:"Todos los KPIs cumplen los umbrales definidos"});
    return a;
  },[K,byMed,porCanal,hayCenso,umbral]);

  const askAI=async()=>{
    if(!aiQ.trim())return;
    setAiL(true);
    const ctx={
      ingreso:K.ing,margen:fP(K.mg),pacientes:K.pacs,
      canales:porCanal.map(c=>({canal:c.label,ingreso:c.ing,pacientes:c.pacs,margen:fP(c.mg)})),
      top5med:byMed.slice(0,5).map(m=>({medico:m.key,ingreso:m.total,eventos:m.count})),
      imss_externos:porCanal.find(c=>c.canal==="externos"),
      tendencia:byMes.slice(-4).map(m=>({mes:m.mes,ingreso:m.cargo,pacs:m.pacs,hosp:m.hosp,urg:m.urg,ext:m.ext})),
    };
    const userM={role:"user",content:aiQ};
    const hist=[...aiH,userM];
    try{
      const r=await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json"},
        body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:1000,
          system:`Eres analista financiero de Hospital Punta Médica. Contexto: ${JSON.stringify(ctx)}. Los Servicios Externos son 100% IMSS manejados por Dra. Sara Ramos (Imagenología, Lab, Radioterapia). Responde en español ejecutivo.`,
          messages:hist})});
      const d=await r.json();
      setAiH([...hist,{role:"assistant",content:d.content?.[0]?.text||"Sin respuesta"}]);
    }catch(e){setAiH([...hist,{role:"assistant",content:"Error de conexión."}]);}
    setAiL(false);setAiQ("");
  };

  const TablaDetalle=({data,titulo})=>{
    const total=data.reduce((s,d)=>s+d.total,0);
    return(
      <div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
          <thead><tr style={{background:C.cyanLL}}>
            {["#",titulo,"Ingreso","Margen","Eventos","% Total"].map(h=>(
              <th key={h} style={{padding:"9px 12px",textAlign:h==="#"||h===titulo?"left":"right",fontSize:10,textTransform:"uppercase",color:C.cyanD,fontWeight:700}}>{h}</th>
            ))}
          </tr></thead>
          <tbody>
            {data.slice(0,12).map((d,i)=>{
              const pct=total>0?d.total/total:0;
              return(
                <tr key={i} style={{borderBottom:`1px solid ${C.borde}`,background:i===0?"#fffbeb":C.blanco}}>
                  <td style={{padding:"9px 12px",color:C.grisM}}>{i+1}</td>
                  <td style={{padding:"9px 12px",fontWeight:600,maxWidth:180,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:C.navy}}>{d.key}</td>
                  <td style={{padding:"9px 12px",textAlign:"right",fontWeight:700,color:C.cyanD}}>{f$(d.total)}</td>
                  <td style={{padding:"9px 12px",textAlign:"right",color:d.mg>0.75?C.verde:d.mg>0.5?C.amber:C.rojo}}>{d.costo>0?fP(d.mg):"—"}</td>
                  <td style={{padding:"9px 12px",textAlign:"right"}}>{fN(d.count)}</td>
                  <td style={{padding:"9px 12px",textAlign:"right"}}>
                    <div style={{display:"flex",alignItems:"center",gap:6,justifyContent:"flex-end"}}>
                      <div style={{width:40,height:5,background:C.borde,borderRadius:3}}><div style={{width:`${Math.min(pct*100,100)}%`,height:"100%",background:C.cyan,borderRadius:3}}/></div>
                      <span style={{color:C.cyanD,fontWeight:600}}>{fP(pct)}</span>
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    );
  };

  if(!rol)return <SelectorRol onSelect={k=>{setRol(k);setView(ROLES[k].views[0]);}} customLogo={customLogo} brandName={brandName}/>;
  if(!fuentes)return <Carga onData={setFuentes}/>;

  const myViews=ROLES[rol]?.views||[];

  return(
    <div style={{fontFamily:"system-ui,sans-serif",background:C.grisL,minHeight:"100vh",fontSize:13}}>

      {/* NAV */}
      <div style={{background:C.navy,padding:"0 20px",display:"flex",alignItems:"center",gap:6,position:"sticky",top:0,zIndex:100}}>
        <div style={{marginRight:16,padding:"8px 0",flexShrink:0}}><Logo customLogo={customLogo} brandName={brandName} size={22}/></div>
        <div style={{display:"flex",flex:1,overflowX:"auto"}}>
          {myViews.map(v=>(
            <button key={v} onClick={()=>setView(v)} style={{padding:"14px 12px",border:"none",background:"none",cursor:"pointer",fontSize:12,fontWeight:500,whiteSpace:"nowrap",color:view===v?"white":"rgba(255,255,255,0.45)",borderBottom:view===v?`3px solid ${C.cyan}`:"3px solid transparent"}}>
              {VISTAS[v]?.icon} {VISTAS[v]?.l}
            </button>
          ))}
        </div>
        <div style={{display:"flex",gap:8,alignItems:"center",flexShrink:0}}>
          <select value={filtros.mes} onChange={e=>setFiltros(f=>({...f,mes:e.target.value}))} style={{padding:"5px 8px",borderRadius:6,border:"1px solid rgba(0,180,216,0.3)",background:"rgba(0,180,216,0.1)",color:"white",fontSize:11}}>
            <option value="Todos">Todos los meses</option>
            {opts.meses.map(m=><option key={m} value={m}>{m}</option>)}
          </select>
          <select value={filtros.canal} onChange={e=>setFiltros(f=>({...f,canal:e.target.value}))} style={{padding:"5px 8px",borderRadius:6,border:"1px solid rgba(92,107,192,0.4)",background:"rgba(92,107,192,0.15)",color:"white",fontSize:11}}>
            <option value="Todos">Todos los canales</option>
            {opts.canales.map(c=><option key={c} value={c}>{c}</option>)}
          </select>
          <button onClick={()=>setFuentes(null)} style={{padding:"5px 10px",background:"rgba(255,255,255,0.06)",border:"1px solid rgba(255,255,255,0.1)",borderRadius:6,color:"rgba(255,255,255,0.5)",fontSize:11,cursor:"pointer"}}>+ Datos</button>
          <button onClick={()=>setShowSettings(true)} title="Ajustes" style={{width:32,height:32,display:"flex",alignItems:"center",justifyContent:"center",background:"rgba(255,255,255,0.08)",border:`1px solid rgba(0,180,216,0.3)`,borderRadius:8,cursor:"pointer"}}>
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke={C.cyan} strokeWidth="2" strokeLinecap="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>
          </button>
          <button onClick={()=>window.print()} style={{padding:"6px 14px",background:C.cyan,border:"none",borderRadius:6,color:C.navy,fontSize:11,fontWeight:700,cursor:"pointer"}}>⬇ PDF</button>
        </div>
      </div>

      {/* INFO BAR */}
      <div style={{background:C.blanco,borderBottom:`1px solid ${C.borde}`,padding:"7px 20px",display:"flex",gap:16,alignItems:"center",fontSize:11,color:C.grisM,flexWrap:"wrap"}}>
        {fuentes.map((f,i)=>(
          <span key={i} style={{display:"flex",alignItems:"center",gap:5}}>
            <span style={{background:C.cyanL,color:C.cyanD,padding:"1px 8px",borderRadius:10,fontWeight:600,fontSize:10}}>{f.tipo}</span>
            {f.nombre} <strong style={{color:C.cyan}}>({fN(f.filas)} filas)</strong>
          </span>
        ))}
        {hayCenso&&<span style={{background:C.violetaL,color:C.violeta,padding:"1px 8px",borderRadius:10,fontWeight:600,fontSize:10}}>✓ Censo cargado</span>}
        <span style={{marginLeft:"auto"}}>{fN(rows.length)} registros · <strong style={{color:C.cyanD}}>{fN(K.fols)}</strong> folios · <strong style={{color:C.cyanD}}>{fN(K.pacs)}</strong> pacientes</span>
      </div>

      <div style={{padding:"20px"}}>

        {/* ══ DASHBOARD ══ */}
        {view==="dashboard"&&(
          <>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:18}}>
              <KpiCard label="Ingreso Total"    value={fM(K.ing)}  sub={`${fN(K.ev)} eventos`}            color={C.cyan}/>
              <KpiCard label="Margen Bruto"     value={fP(K.mg)}   sub={`Costo: ${fM(K.cos)}`}            color={C.teal}  badge={K.mg>0.75?"✓ Sólido":"⚠ Revisar"} badgeOk={K.mg>0.75}/>
              <KpiCard label="Pacientes Únicos" value={fN(K.pacs)} sub={`Ticket: ${f$(K.tick)}`}          color={C.cyanD}/>
              <KpiCard label="Utilidad Bruta"   value={fM(K.ut)}   badge={K.ut>0?"Positivo":"⚠ Negativo"} badgeOk={K.ut>0} color={C.cyan}/>
            </div>

            {/* KPIs por canal */}
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12,marginBottom:18}}>
              {["hospitalizacion","urgencias","externos"].map(clave=>{
                const d=porCanal.find(c=>c.canal===clave)||{ing:0,pacs:0,ev:0,mg:0,tick:0};
                const cfg=CANALES[clave];
                return(
                  <div key={clave} style={{background:C.blanco,border:`1px solid ${C.borde}`,borderRadius:12,padding:"14px 18px",borderLeft:`4px solid ${cfg.color}`}}>
                    <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                      <span style={{fontSize:20}}>{cfg.icon}</span>
                      <div>
                        <div style={{fontWeight:700,fontSize:13,color:C.navy}}>{cfg.label}</div>
                        {clave==="externos"&&<div style={{fontSize:10,color:C.violeta,fontWeight:600}}>IMSS · Dra. Sara Ramos</div>}
                      </div>
                    </div>
                    <div style={{fontFamily:"Georgia,serif",fontSize:24,fontWeight:700,color:cfg.color,marginBottom:4}}>{fM(d.ing)}</div>
                    <div style={{display:"flex",gap:12,fontSize:11,color:C.grisM}}>
                      <span>{fN(d.pacs)} pacs.</span>
                      <span>{fN(d.ev)} eventos</span>
                      <span>{f$(d.tick)} ticket</span>
                    </div>
                    {K.ing>0&&<div style={{marginTop:8,height:4,background:C.borde,borderRadius:2}}><div style={{width:`${(d.ing/K.ing*100).toFixed(0)}%`,height:"100%",background:cfg.color,borderRadius:2}}/></div>}
                    {K.ing>0&&<div style={{fontSize:10,color:C.grisM,marginTop:3}}>{fP(d.ing/K.ing)} del total</div>}
                  </div>
                );
              })}
            </div>

            <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:16,marginBottom:16}}>
              <Card>
                <div style={{fontWeight:700,marginBottom:12,color:C.navy}}>Ingreso por Canal y Mes</div>
                <ResponsiveContainer width="100%" height={220}>
                  <BarChart data={byMes}>
                    <CartesianGrid strokeDasharray="3 3" stroke={C.borde}/>
                    <XAxis dataKey="mes" tick={{fontSize:10}}/><YAxis tickFormatter={v=>"$"+(v/1e6).toFixed(1)+"M"} tick={{fontSize:10}}/>
                    <Tooltip content={<Tip/>}/><Legend/>
                    <Bar dataKey="hosp" stackId="a" fill={CANALES.hospitalizacion.color} name="Hospitalización" radius={[0,0,0,0]}/>
                    <Bar dataKey="urg"  stackId="a" fill={CANALES.urgencias.color}       name="Urgencias"       radius={[0,0,0,0]}/>
                    <Bar dataKey="ext"  stackId="a" fill={CANALES.externos.color}        name="Externos (IMSS)" radius={[3,3,0,0]}/>
                  </BarChart>
                </ResponsiveContainer>
              </Card>
              <Card>
                <div style={{fontWeight:700,marginBottom:12,color:C.navy}}>Mix por Canal</div>
                <ResponsiveContainer width="100%" height={220}>
                  <PieChart>
                    <Pie data={porCanal} dataKey="ing" nameKey="label" cx="50%" cy="50%" outerRadius={85} innerRadius={30}
                      label={({percent})=>percent>0.04?`${(percent*100).toFixed(0)}%`:""} labelLine={false}>
                      {porCanal.map((c,i)=><Cell key={i} fill={c.color}/>)}
                    </Pie>
                    <Tooltip formatter={v=>f$(v)}/><Legend iconSize={10} wrapperStyle={{fontSize:10}}/>
                  </PieChart>
                </ResponsiveContainer>
              </Card>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
              <Card>
                <div style={{fontWeight:700,marginBottom:12,color:C.navy}}>Top 8 Médicos</div>
                <ResponsiveContainer width="100%" height={210}>
                  <BarChart data={byMed.slice(0,8)} layout="vertical">
                    <CartesianGrid strokeDasharray="3 3" stroke={C.borde}/>
                    <XAxis type="number" tickFormatter={v=>"$"+(v/1e6).toFixed(1)+"M"} tick={{fontSize:10}}/>
                    <YAxis type="category" dataKey="key" tick={{fontSize:9}} width={110}/>
                    <Tooltip content={<Tip/>}/><Bar dataKey="total" fill={C.cyan} radius={[0,5,5,0]} name="Ingreso"/>
                  </BarChart>
                </ResponsiveContainer>
              </Card>
              <Card>
                <div style={{fontWeight:700,marginBottom:12,color:C.navy}}>Top 8 Servicios</div>
                <ResponsiveContainer width="100%" height={210}>
                  <BarChart data={byServ.slice(0,8)} layout="vertical">
                    <CartesianGrid strokeDasharray="3 3" stroke={C.borde}/>
                    <XAxis type="number" tickFormatter={v=>"$"+(v/1e6).toFixed(1)+"M"} tick={{fontSize:10}}/>
                    <YAxis type="category" dataKey="key" tick={{fontSize:9}} width={110}/>
                    <Tooltip content={<Tip/>}/><Bar dataKey="total" fill={C.teal} radius={[0,5,5,0]} name="Ingreso"/>
                  </BarChart>
                </ResponsiveContainer>
              </Card>
            </div>
          </>
        )}

        {/* ══ CANALES ══ */}
        {view==="canales"&&(
          <>
            <div style={{background:C.blanco,border:`2px solid ${C.violeta}`,borderRadius:12,padding:"14px 18px",marginBottom:18,display:"flex",gap:12,alignItems:"center"}}>
              <span style={{fontSize:24}}>ℹ️</span>
              <div style={{fontSize:13,color:C.navy,lineHeight:1.6}}>
                <strong>Lógica de clasificación HPM:</strong> El Censo Histórico identifica pacientes de <strong>Hospitalización</strong> y <strong>Urgencias</strong>.
                Los pacientes en Cargos que <strong>no aparecen en el censo</strong> son <strong>Servicios Externos</strong> (Imagenología, Laboratorio, Radioterapia, etc.) →
                <span style={{color:C.violeta,fontWeight:700}}> 100% Convenio IMSS · Dra. Sara Ramos</span>
                {!hayCenso&&<span style={{display:"block",color:C.amber,fontWeight:600,marginTop:4}}>⚠️ Sube el Censo Histórico para clasificación precisa</span>}
              </div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:16,marginBottom:18}}>
              {["hospitalizacion","urgencias","externos"].map(clave=>{
                const d=porCanal.find(c=>c.canal===clave)||{ing:0,pacs:0,ev:0,mg:0,cos:0};
                const cfg=CANALES[clave];
                return(
                  <div key={clave} style={{background:C.blanco,border:`2px solid ${cfg.color}`,borderRadius:14,padding:"20px"}}>
                    <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14}}>
                      <span style={{fontSize:28}}>{cfg.icon}</span>
                      <div>
                        <div style={{fontWeight:700,fontSize:15,color:C.navy}}>{cfg.label}</div>
                        {clave==="externos"&&<div style={{fontSize:11,color:C.violeta,fontWeight:600}}>IMSS · Dra. Sara Ramos</div>}
                      </div>
                    </div>
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                      {[["Ingreso",fM(d.ing)],["Margen",fP(d.mg)],["Pacientes",fN(d.pacs)],["Eventos",fN(d.ev)],["Costo",fM(d.cos)],["Ticket",f$(d.tick)]].map(([l,v])=>(
                        <div key={l} style={{background:C.grisL,borderRadius:8,padding:"8px 10px"}}>
                          <div style={{fontSize:9,textTransform:"uppercase",color:C.grisM,letterSpacing:"0.08em",marginBottom:3}}>{l}</div>
                          <div style={{fontWeight:700,fontSize:14,color:C.navy}}>{v}</div>
                        </div>
                      ))}
                    </div>
                    {K.ing>0&&(
                      <div style={{marginTop:12}}>
                        <div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginBottom:4}}>
                          <span style={{color:C.grisM}}>% del ingreso total</span>
                          <strong style={{color:cfg.color}}>{fP(d.ing/K.ing)}</strong>
                        </div>
                        <div style={{height:6,background:C.borde,borderRadius:3}}><div style={{width:`${Math.min(d.ing/K.ing*100,100).toFixed(0)}%`,height:"100%",background:cfg.color,borderRadius:3}}/></div>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
            <Card>
              <div style={{fontWeight:700,marginBottom:14,color:C.navy}}>Evolución mensual por Canal</div>
              <ResponsiveContainer width="100%" height={260}>
                <BarChart data={byMes}>
                  <CartesianGrid strokeDasharray="3 3" stroke={C.borde}/>
                  <XAxis dataKey="mes" tick={{fontSize:10}}/><YAxis tickFormatter={v=>"$"+(v/1e6).toFixed(1)+"M"} tick={{fontSize:10}}/>
                  <Tooltip content={<Tip/>}/><Legend/>
                  <Bar dataKey="hosp" stackId="a" fill={CANALES.hospitalizacion.color} name="Hospitalización"/>
                  <Bar dataKey="urg"  stackId="a" fill={CANALES.urgencias.color}       name="Urgencias"/>
                  <Bar dataKey="ext"  stackId="a" fill={CANALES.externos.color}        name="Externos IMSS" radius={[3,3,0,0]}/>
                </BarChart>
              </ResponsiveContainer>
            </Card>
          </>
        )}

        {/* ══ MÉDICOS / SERVICIOS / CONVENIOS ══ */}
        {(view==="medicos"||view==="servicios"||view==="convenios")&&(()=>{
          const data=view==="medicos"?byMed:view==="servicios"?byServ:byConv;
          const tit=view==="medicos"?"Médico":view==="servicios"?"Servicio":"Convenio";
          const color=view==="medicos"?C.cyan:view==="servicios"?C.teal:C.cyanD;
          const total=data.reduce((s,d)=>s+d.total,0);
          return(
            <>
              <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:18}}>
                <KpiCard label={`Total ${tit}s`}    value={fN(data.length)} color={color}/>
                <KpiCard label="Ingreso total"       value={fM(total)}       color={color}/>
                <KpiCard label="Concentración Top-3" value={total?fP(data.slice(0,3).reduce((s,d)=>s+d.total,0)/total):"—"} color={color}
                  badge={data.slice(0,3).reduce((s,d)=>s+d.total,0)/Math.max(total,1)>0.5?"⚠ Alta":"Moderada"}
                  badgeOk={data.slice(0,3).reduce((s,d)=>s+d.total,0)/Math.max(total,1)<=0.5}/>
                <KpiCard label="Ticket / Paciente"   value={f$(K.tick)}      color={color}/>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"3fr 2fr",gap:16,marginBottom:16}}>
                <Card>
                  <div style={{fontWeight:700,marginBottom:12,color:C.navy}}>Top 10 — {tit}</div>
                  <ResponsiveContainer width="100%" height={280}>
                    <BarChart data={data.slice(0,10)} layout="vertical">
                      <CartesianGrid strokeDasharray="3 3" stroke={C.borde}/>
                      <XAxis type="number" tickFormatter={v=>"$"+(v/1e6).toFixed(1)+"M"} tick={{fontSize:10}}/>
                      <YAxis type="category" dataKey="key" tick={{fontSize:9}} width={120}/>
                      <Tooltip content={<Tip/>}/><Bar dataKey="total" fill={color} radius={[0,5,5,0]} name="Ingreso"/>
                    </BarChart>
                  </ResponsiveContainer>
                </Card>
                <Card>
                  <div style={{fontWeight:700,marginBottom:12,color:C.navy}}>Distribución %</div>
                  <ResponsiveContainer width="100%" height={280}>
                    <PieChart>
                      <Pie data={data.slice(0,6)} dataKey="total" nameKey="key" cx="50%" cy="50%" outerRadius={100} innerRadius={30}
                        label={({percent})=>percent>0.04?`${(percent*100).toFixed(0)}%`:""} labelLine={false}>
                        {data.slice(0,6).map((_,i)=><Cell key={i} fill={C.ch[i%C.ch.length]}/>)}
                      </Pie>
                      <Tooltip formatter={v=>f$(v)}/><Legend iconSize={10} wrapperStyle={{fontSize:10}}/>
                    </PieChart>
                  </ResponsiveContainer>
                </Card>
              </div>
              <Card><div style={{fontWeight:700,marginBottom:12,color:C.navy}}>Detalle por {tit}</div><TablaDetalle data={data} titulo={tit}/></Card>
            </>
          );
        })()}

        {/* ══ FINANCIERO ══ */}
        {view==="financiero"&&(
          <>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:18}}>
              <KpiCard label="Ingreso Neto"   value={fM(K.ing)} sub={`${fN(K.fols)} folios`} color={C.cyan}/>
              <KpiCard label="Costo Directo"  value={fM(K.cos)} sub={fP(K.cos/Math.max(K.ing,1))+" del ingreso"} color={C.rojo}/>
              <KpiCard label="Utilidad Bruta" value={fM(K.ut)}  badge={K.ut>0?"✓":"⚠"} badgeOk={K.ut>0} color={C.teal}/>
              <KpiCard label="Margen Bruto"   value={fP(K.mg)}  badge={K.mg>0.75?"Sólido":"Revisar"} badgeOk={K.mg>0.75} color={C.cyanD}/>
            </div>
            <Card style={{marginBottom:16}}>
              <div style={{fontWeight:700,marginBottom:12,color:C.navy}}>Proyección — Próximos 6 Meses</div>
              <ResponsiveContainer width="100%" height={240}>
                <LineChart data={[...byMes.slice(-2).map(m=>({mes:m.mes,real:m.cargo})),...forecast]}>
                  <CartesianGrid strokeDasharray="3 3" stroke={C.borde}/>
                  <XAxis dataKey="mes" tick={{fontSize:10}}/><YAxis tickFormatter={v=>"$"+(v/1e6).toFixed(1)+"M"} tick={{fontSize:10}}/>
                  <Tooltip content={<Tip/>}/><Legend/>
                  <Line type="monotone" dataKey="real"        stroke={C.cyan}  strokeWidth={3}   dot={{r:4}} name="Real"/>
                  <Line type="monotone" dataKey="base"        stroke={C.amber} strokeWidth={2}   strokeDasharray="6 3" dot={{r:3}} name="Base"/>
                  <Line type="monotone" dataKey="optimista"   stroke={C.verde} strokeWidth={1.5} strokeDasharray="6 3" dot={{r:3}} name="Optimista"/>
                  <Line type="monotone" dataKey="conservador" stroke={C.rojo}  strokeWidth={1.5} strokeDasharray="6 3" dot={{r:3}} name="Conservador"/>
                </LineChart>
              </ResponsiveContainer>
            </Card>
          </>
        )}

        {/* ══ TENDENCIAS ══ */}
        {view==="tendencias"&&(
          <>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:14,marginBottom:18}}>
              <KpiCard label="Meses con datos"  value={byMes.length} color={C.cyan}/>
              <KpiCard label="Mejor mes"        value={byMes.length?fM(Math.max(...byMes.map(m=>m.cargo))):"—"} sub={byMes.length?byMes.reduce((a,b)=>a.cargo>b.cargo?a:b).mes:""} color={C.verde}/>
              <KpiCard label="Promedio mensual" value={byMes.length?fM(byMes.reduce((s,m)=>s+m.cargo,0)/byMes.length):"—"} color={C.teal}/>
              <KpiCard label="Pacs. prom/mes"   value={byMes.length?fN(byMes.reduce((s,m)=>s+m.pacs,0)/byMes.length):"—"} color={C.cyanD}/>
            </div>
            <Card style={{marginBottom:16}}>
              <div style={{fontWeight:700,marginBottom:12,color:C.navy}}>Ingreso por Canal Mensual</div>
              <ResponsiveContainer width="100%" height={250}>
                <BarChart data={byMes}>
                  <CartesianGrid strokeDasharray="3 3" stroke={C.borde}/>
                  <XAxis dataKey="mes" tick={{fontSize:10}}/><YAxis tickFormatter={v=>"$"+(v/1e6).toFixed(1)+"M"} tick={{fontSize:10}}/>
                  <Tooltip content={<Tip/>}/><Legend/>
                  <Bar dataKey="hosp" stackId="a" fill={CANALES.hospitalizacion.color} name="Hospitalización"/>
                  <Bar dataKey="urg"  stackId="a" fill={CANALES.urgencias.color}       name="Urgencias"/>
                  <Bar dataKey="ext"  stackId="a" fill={CANALES.externos.color}        name="Externos IMSS" radius={[3,3,0,0]}/>
                </BarChart>
              </ResponsiveContainer>
            </Card>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
              <Card>
                <div style={{fontWeight:700,marginBottom:12,color:C.navy}}>Pacientes y Folios / Mes</div>
                <ResponsiveContainer width="100%" height={200}>
                  <LineChart data={byMes}>
                    <CartesianGrid strokeDasharray="3 3" stroke={C.borde}/>
                    <XAxis dataKey="mes" tick={{fontSize:10}}/><YAxis tick={{fontSize:10}}/>
                    <Tooltip/><Legend/>
                    <Line type="monotone" dataKey="pacs" stroke={C.cyan}  strokeWidth={2.5} dot={{r:3}} name="Pacientes"/>
                    <Line type="monotone" dataKey="fols" stroke={C.amber} strokeWidth={2.5} dot={{r:3}} name="Folios"/>
                  </LineChart>
                </ResponsiveContainer>
              </Card>
              <Card>
                <div style={{fontWeight:700,marginBottom:12,color:C.navy}}>Ticket Promedio / Mes</div>
                <ResponsiveContainer width="100%" height={200}>
                  <AreaChart data={byMes}>
                    <defs><linearGradient id="gt" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={C.teal} stopOpacity={0.2}/><stop offset="95%" stopColor={C.teal} stopOpacity={0}/></linearGradient></defs>
                    <CartesianGrid strokeDasharray="3 3" stroke={C.borde}/>
                    <XAxis dataKey="mes" tick={{fontSize:10}}/><YAxis tickFormatter={v=>f$(v)} tick={{fontSize:10}}/>
                    <Tooltip content={<Tip/>}/><Area type="monotone" dataKey="tick" stroke={C.teal} fill="url(#gt)" strokeWidth={2.5} name="Ticket"/>
                  </AreaChart>
                </ResponsiveContainer>
              </Card>
            </div>
          </>
        )}

        {/* ══ ALERTAS ══ */}
        {view==="alertas"&&(
          <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:16}}>
            <div>
              <div style={{fontWeight:700,fontSize:16,color:C.navy,marginBottom:14}}>🔔 Alertas Automáticas</div>
              {alertas.map((a,i)=>{
                const cfg={critico:{bg:C.rojoL,bl:C.rojo,c:C.rojo,ic:"🔴"},alerta:{bg:C.amberL,bl:C.amber,c:"#92400E",ic:"🟡"},ok:{bg:C.verdeL,bl:C.verde,c:C.verde,ic:"🟢"}};
                const s=cfg[a.tipo]||cfg.alerta;
                return(<div key={i} style={{background:s.bg,borderLeft:`4px solid ${s.bl}`,borderRadius:"0 10px 10px 0",padding:"12px 16px",marginBottom:10}}>
                  <div style={{fontWeight:600,fontSize:13,color:s.c,marginBottom:3}}>{s.ic} {a.titulo}</div>
                  <div style={{fontSize:12,color:C.gris}}>{a.msg}</div>
                </div>);
              })}
            </div>
            <Card>
              <div style={{fontWeight:700,marginBottom:14,color:C.navy}}>⚙️ Umbrales</div>
              {[{k:"margen",l:"Margen mínimo %"},{k:"concentracion",l:"Concentración máx. top-3 %"}].map(u=>(
                <div key={u.k} style={{marginBottom:20}}>
                  <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:8}}><span style={{color:C.gris}}>{u.l}</span><strong style={{color:C.cyan}}>{umbral[u.k]}%</strong></div>
                  <input type="range" min={0} max={100} value={umbral[u.k]} onChange={e=>setUmbral(p=>({...p,[u.k]:parseInt(e.target.value)}))} style={{width:"100%",accentColor:C.cyan}}/>
                </div>
              ))}
            </Card>
          </div>
        )}

        {/* ══ IA ══ */}
        {view==="ia"&&(
          <div style={{display:"grid",gridTemplateColumns:"1fr 280px",gap:16}}>
            <Card>
              <div style={{fontWeight:700,fontSize:16,color:C.navy,marginBottom:4}}>✨ Análisis con IA</div>
              <div style={{fontSize:12,color:C.grisM,marginBottom:14}}>Consulta los datos del hospital en lenguaje natural</div>
              <div style={{height:380,overflowY:"auto",border:`1px solid ${C.borde}`,borderRadius:10,padding:14,marginBottom:10,background:C.cyanLL}}>
                {aiH.length===0
                  ?<div style={{textAlign:"center",color:C.grisM,marginTop:80,fontSize:13}}>Pregunta sobre Hospitalización, Urgencias o Servicios IMSS</div>
                  :aiH.map((m,i)=>(
                    <div key={i} style={{marginBottom:14,display:"flex",gap:8,flexDirection:m.role==="user"?"row-reverse":"row"}}>
                      <div style={{width:26,height:26,borderRadius:"50%",background:m.role==="user"?C.navy:C.cyan,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,color:"white",fontWeight:700,flexShrink:0}}>{m.role==="user"?"U":"IA"}</div>
                      <div style={{background:m.role==="user"?C.navy:"white",color:m.role==="user"?"white":C.navy,padding:"9px 13px",borderRadius:m.role==="user"?"12px 12px 4px 12px":"12px 12px 12px 4px",fontSize:12,lineHeight:1.6,maxWidth:"82%",border:m.role==="assistant"?`1px solid ${C.borde}`:"none",whiteSpace:"pre-wrap"}}>{m.content}</div>
                    </div>
                  ))
                }
                {aiL&&<div style={{display:"flex",gap:8}}><div style={{width:26,height:26,borderRadius:"50%",background:C.cyan,display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,color:C.navy,fontWeight:700}}>IA</div><div style={{background:"white",padding:"9px 13px",borderRadius:"12px 12px 12px 4px",border:`1px solid ${C.borde}`,color:C.grisM,fontSize:12}}>Analizando...</div></div>}
              </div>
              <div style={{display:"flex",gap:8}}>
                <input value={aiQ} onChange={e=>setAiQ(e.target.value)} onKeyDown={e=>e.key==="Enter"&&!e.shiftKey&&askAI()}
                  placeholder="¿Cuánto generan los Servicios Externos IMSS este mes?..."
                  style={{flex:1,padding:"9px 12px",border:`1px solid ${C.borde}`,borderRadius:7,fontSize:12,outline:"none",fontFamily:"inherit"}}/>
                <button onClick={askAI} disabled={aiL||!aiQ.trim()} style={{padding:"9px 18px",background:C.cyan,color:C.navy,border:"none",borderRadius:7,fontWeight:700,cursor:"pointer",fontSize:12}}>Enviar</button>
              </div>
            </Card>
            <div style={{display:"flex",flexDirection:"column",gap:12}}>
              <Card>
                <div style={{fontWeight:700,marginBottom:10,color:C.navy}}>Preguntas rápidas</div>
                {[
                  "¿Cuánto generan los Servicios Externos IMSS?",
                  "Compara Hospitalización vs Urgencias vs IMSS",
                  "¿Cuál es el ticket promedio por canal?",
                  "Resume los resultados por canal para los accionistas",
                  "¿Qué porcentaje del ingreso es IMSS?",
                  "Identifica los 3 principales riesgos financieros",
                ].map((q,i)=>(
                  <button key={i} onClick={()=>setAiQ(q)} style={{width:"100%",textAlign:"left",padding:"7px 10px",marginBottom:5,border:`1px solid ${C.borde}`,borderRadius:7,background:"none",cursor:"pointer",fontSize:11,color:C.gris,lineHeight:1.4,fontFamily:"inherit"}}>{q}</button>
                ))}
              </Card>
            </div>
          </div>
        )}

      </div>

      {showSettings&&(
        <ModalAjustes onClose={()=>setShowSettings(false)}
          customLogo={customLogo} setCustomLogo={setCustomLogo}
          brandName={brandName}   setBrandName={setBrandName}
          brandSub={brandSub}     setBrandSub={setBrandSub}/>
      )}
    </div>
  );
}
