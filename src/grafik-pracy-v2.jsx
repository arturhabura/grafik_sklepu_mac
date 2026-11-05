import React, { useState, useMemo, useEffect, useRef, useCallback } from "react";
import { Plus, X, Trash2, ChevronLeft, ChevronRight, Clock, Users, Calendar, Wallet, Megaphone, Pencil, Check, Printer, Settings, ShieldCheck, Eye, EyeOff, Copy, Database, Download, Upload, LogOut, Lock, BarChart2, FileText } from "lucide-react";
import * as XLSX from "xlsx";

const DAY_START = 5, DAY_END = 22;
const DAY_NAMES = ["Pon","Wt","Śr","Czw","Pt","Sob","Nie"];
const DAY_NAMES_FULL = ["Poniedziałek","Wtorek","Środa","Czwartek","Piątek","Sobota","Niedziela"];
const TYPES = {
  praca:           { label: "Praca",                        bg: "#2F6B4F", text: "#EAF3E2" },
  wolne:           { label: "Dzień wolny",                  bg: "#9C988A", text: "#FBFAF7" },
  urlop:           { label: "Urlop",                        bg: "#3B6FA0", text: "#E7F0FA" },
  urlop_bezplatny: { label: "Urlop bezpłatny",              bg: "#6B6A60", text: "#F4F3EE" },
  l4:              { label: "L4",                           bg: "#C75146", text: "#FBEAE8" },
  l4b:             { label: "L4 ciążowe (kod B)",           bg: "#A6446B", text: "#FBE7EF" },
  opieka_chore:    { label: "Opieka (chore dz./rodzina)",   bg: "#B85C2E", text: "#FBF0EB" },
  opieka_zdrowe:   { label: "Opieka (zdrowe dziecko)",      bg: "#5B8B3E", text: "#EDF4E8" },
};
const PAYMENT_METHODS = [
  { id: "przelew",         label: "Konto bankowe" },
  { id: "przelew_gotowka", label: "Konto bankowe + gotówka" },
  { id: "gotowka",         label: "Tylko gotówka" },
];
const ETAT_OPTIONS = ["1","3/4","2/3","1/2","1/3"];
const parseEtat = (e) => { if (!e||e==="1") return 1; const p=e.split("/"); return p.length===2?Number(p[0])/Number(p[1]):Number(e)||1; };

// ---- Autoryzacja ----
const encodePass = (p) => { try { return btoa(encodeURIComponent(p) + "_gk7x2"); } catch(e) { return btoa(p + "_gk7x2"); } };
const checkPass = (input, stored) => !!(stored && encodePass(input) === stored);
const DEFAULT_ADMIN_HASH = encodePass("admin");
const DEFAULT_MANAGER_HASH = encodePass("sklep");

// ---- Jeden sklep na instalację ----
const INITIAL_STORES = [{ id: "s1", name: "Mój Sklep" }];
const SNAP_MIN = 15, MIN_DURATION = 30;

const uid = () => Math.random().toString(36).slice(2, 10);
// Wyświetla imię w formacie "Nazwisko Imię" (ostatnie słowo = nazwisko)
const displayName = (n) => { if(!n) return n; const p=n.trim().split(/\s+/); return p.length<2 ? n : p[p.length-1]+" "+p.slice(0,-1).join(" "); };
const toMin = (hhmm) => { if (!hhmm) return 0; const [h,m] = hhmm.split(":").map(Number); return h*60+m; };
const toHHMM = (mins) => `${pad2(Math.floor(mins/60))}:${pad2(mins%60)}`;
const pad2 = (n) => String(n).padStart(2,"0");
const dateStr = (d) => `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}`;
const startOfWeek = (d) => { const x=new Date(d); const day=(x.getDay()+6)%7; x.setDate(x.getDate()-day); x.setHours(0,0,0,0); return x; };
const addDays = (d,n) => { const x=new Date(d); x.setDate(x.getDate()+n); return x; };
const monthLabel = (d) => d.toLocaleDateString("pl-PL",{month:"long",year:"numeric"});
const dayLabelFull = (d) => d.toLocaleDateString("pl-PL",{weekday:"long",day:"numeric",month:"long",year:"numeric"});

// ---- Mapowanie DB (snake_case) ↔ React (camelCase) ----
const dbToEmp = (e) => ({ id: e.id, name: e.name, storeId: e.store_id||e.storeId, position: e.position||"", rate: Number(e.rate||28), paymentMethod: e.payment_method||e.paymentMethod||"przelew", defaultTransferAmount: Number(e.default_transfer||e.defaultTransferAmount||0) });
const empToDb = (e) => ({ id: e.id, name: e.name, store_id: e.storeId, position: e.position||"", rate: e.rate, payment_method: e.paymentMethod||"przelew", default_transfer: e.defaultTransferAmount||0 });
const dbToShift = (s) => ({ id: s.id, employeeId: s.employee_id||s.employeeId, date: s.date, type: s.type, start: s.start_time||s.start||null, end: s.end_time||s.end||null, tasks: typeof s.tasks==="string" ? JSON.parse(s.tasks||"[]") : (s.tasks||[]) });
const shiftToDb = (s) => ({ id: s.id, employee_id: s.employeeId, date: s.date, type: s.type, start_time: s.start||null, end_time: s.end||null, tasks: s.tasks||[] });

// ---- Baza danych startowych ----
const seedEmployees = () => {
  const names = ["Anna Kowalska","Piotr Nowak","Marta Wiśniewska","Tomasz Lewandowski"];
  return names.map((name,i) => ({ id:uid(), name, storeId:"s1", position:"Sprzedawca", rate:28+(i%3)*2, etat:"1", paymentMethod:"przelew", defaultTransferAmount:0, contractEndDate:null }));
};
const seedShifts = (employees) => {
  const monday = startOfWeek(new Date()), out = [];
  INITIAL_STORES.forEach((store) => {
    const emp = employees.filter(e => e.storeId===store.id);
    for (let d=0;d<6;d++) {
      const date = dateStr(addDays(monday,d));
      emp.forEach((e,i) => {
        const templates = [["05:00","13:00"],["09:00","17:00"],["13:00","21:00"],["14:00","22:00"]];
        const [st,en] = templates[i%4];
        if ((i+d)%4===3) out.push({ id:uid(), employeeId:e.id, date, type:"wolne", start:null, end:null, tasks:[] });
        else if ((i+d)%9===0) out.push({ id:uid(), employeeId:e.id, date, type:"urlop", start:null, end:null, tasks:[] });
        else out.push({ id:uid(), employeeId:e.id, date, type:"praca", start:st, end:en, tasks:i===0?["Dostawa pieczywa - przyjęcie i sprawdzenie ilości","Uzupełnić regały"]:[] });
      });
    }
  });
  return out;
};

class ErrorBoundary extends React.Component {
  constructor(props) { super(props); this.state = { error:null }; }
  static getDerivedStateFromError(e) { return { error:e }; }
  render() {
    if (this.state.error) return (
      <div style={{ padding:20, background:"#FBEAE8", borderRadius:10, color:"#C75146", fontSize:13 }}>
        <b>Błąd renderowania:</b> {this.state.error.message}
        <pre style={{ fontSize:10, marginTop:8, whiteSpace:"pre-wrap" }}>{this.state.error.stack?.split('\n').slice(0,5).join('\n')}</pre>
        <button onClick={()=>this.setState({error:null})} style={{ marginTop:8, padding:"4px 10px", borderRadius:6, border:"1px solid #C75146", background:"transparent", cursor:"pointer", color:"#C75146" }}>Spróbuj ponownie</button>
      </div>
    );
    return this.props.children;
  }
}

// ============ EKRAN LOGOWANIA ============
function LoginScreen({ onLogin, storeName, adminHash, managerHash }) {
  const [role, setRole] = useState("manager");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const handleSubmit = () => {
    const stored = role==="admin" ? (adminHash||DEFAULT_ADMIN_HASH) : (managerHash||DEFAULT_MANAGER_HASH);
    if (checkPass(password, stored)) { onLogin(role); }
    else { setError("Nieprawidłowe hasło."); }
  };
  return (
    <div style={{ minHeight:"100vh", background:"#162820", display:"flex", alignItems:"center", justifyContent:"center", fontFamily:"ui-sans-serif,system-ui,-apple-system,'Segoe UI',Roboto,sans-serif" }}>
      <div style={{ background:"#fff", borderRadius:18, padding:40, width:400, maxWidth:"90vw", boxShadow:"0 24px 64px rgba(0,0,0,0.4)", border:"1px solid rgba(255,255,255,0.08)" }}>
        <div style={{ textAlign:"center", marginBottom:32 }}>
          <div style={{ width:60, height:60, background:"#E07A3E", borderRadius:16, margin:"0 auto 14px", display:"flex", alignItems:"center", justifyContent:"center", fontSize:28, fontWeight:800, color:"#fff", boxShadow:"0 4px 16px rgba(224,122,62,0.4)" }}>G</div>
          <div style={{ fontWeight:800, fontSize:22, color:"#1A2530", letterSpacing:"-0.02em" }}>Grafik Sklepu</div>
          <div style={{ fontSize:13.5, color:"#8890A0", marginTop:5, fontWeight:500 }}>{storeName||"Sklep"}</div>
        </div>
        <div style={{ marginBottom:18 }}>
          <div style={{ fontSize:11, fontWeight:700, color:"#8890A0", marginBottom:8, textTransform:"uppercase", letterSpacing:"0.06em" }}>Zaloguj się jako</div>
          <div style={{ display:"flex", gap:8, background:"#F7F9FC", borderRadius:10, padding:4 }}>
            {[["manager","Kierownik"],["admin","Administrator"]].map(([r,label])=>(
              <button key={r} onClick={()=>{ setRole(r); setError(""); setPassword(""); }} style={{ flex:1, padding:"9px 8px", borderRadius:7, border:"none", background:role===r?"#fff":"transparent", color:role===r?"#1A2530":"#8890A0", fontWeight:role===r?700:500, fontSize:13, cursor:"pointer", boxShadow:role===r?"0 1px 4px rgba(0,0,0,0.1)":"none", transition:"all 0.15s" }}>{label}</button>
            ))}
          </div>
        </div>
        <div style={{ marginBottom:18 }}>
          <div style={{ fontSize:11, fontWeight:700, color:"#8890A0", marginBottom:6, textTransform:"uppercase", letterSpacing:"0.06em" }}>Hasło</div>
          <input type="password" value={password} onChange={e=>{ setPassword(e.target.value); setError(""); }} onKeyDown={e=>e.key==="Enter"&&handleSubmit()} placeholder="Wpisz hasło..." autoFocus style={{ width:"100%", padding:"11px 13px", borderRadius:9, border:"1px solid #DDE0E8", fontSize:14, fontFamily:"inherit", boxSizing:"border-box", outline:"none", color:"#1A2530" }}/>
        </div>
        {error && <div style={{ background:"#FEF2F2", color:"#C23B3B", fontSize:12.5, padding:"9px 12px", borderRadius:8, marginBottom:14, fontWeight:500, display:"flex", alignItems:"center", gap:6 }}>⚠ {error}</div>}
        <button onClick={handleSubmit} style={{ width:"100%", padding:"13px", borderRadius:9, background:"#1F6B45", color:"#fff", fontWeight:700, fontSize:14, border:"none", cursor:"pointer", letterSpacing:"0.01em", boxShadow:"0 2px 12px rgba(31,107,69,0.3)" }}>Zaloguj się</button>
        <div style={{ marginTop:16, fontSize:11, color:"#B0B6C2", textAlign:"center" }}>Domyślne hasła: Administrator → <b style={{color:"#8890A0"}}>admin</b> · Kierownik → <b style={{color:"#8890A0"}}>sklep</b></div>
      </div>
    </div>
  );
}

// ============ APP ============
export default function App() {
  // ---- Stan aplikacji ----
  const [employees, setEmployees] = useState(seedEmployees());
  const [shifts, setShifts] = useState(()=>seedShifts(seedEmployees()));
  const [stores, setStores] = useState(INITIAL_STORES);
  const [positions, setPositions] = useState(["Sprzedawca","Kasjer","Magazynier","Zastępca kierownika","Kierownik zmiany"]);
  const [settings, setSettingsState] = useState({ l4Percent:80, refHours:8, opiekaChorePercent:80, storeName:"Mój Sklep", adminPasswordHash:"", managerPasswordHash:"" });
  const [adjustments, setAdjustments] = useState({});
  const [currentUser, setCurrentUser] = useState(null);

  const isAdmin = currentUser?.role==="admin";
  const effectiveStoreId = "s1";
  const currentStoreName = settings.storeName||stores[0]?.name||"Sklep";

  const [tab, setTab] = useState("grafik");
  const [viewMode, setViewMode] = useState("week");
  const [weekStart, setWeekStart] = useState(startOfWeek(new Date()));
  const [selectedDate, setSelectedDate] = useState(new Date());
  const [modal, setModal] = useState(null);
  const [bulkModal, setBulkModal] = useState(false);
  const [empModal, setEmpModal] = useState(null);
  const [copyDayModal, setCopyDayModal] = useState(false);
  const [copyWeekModal, setCopyWeekModal] = useState(false);
  const [payMonth, setPayMonth] = useState(new Date());
  const [adjModal, setAdjModal] = useState(null);
  const [dataLoaded, setDataLoaded] = useState(false);
  const [saveState, setSaveState] = useState("idle");
  const saveTimer = useRef(null);

  const electronAvailable = typeof window!=="undefined"&&typeof window.api!=="undefined";
  const storageAvailable = !electronAvailable&&typeof window!=="undefined"&&typeof window.storage!=="undefined";
  const lsAvailable = typeof localStorage!=="undefined";

  // ---- Wczytywanie danych ----
  useEffect(()=>{
    try {
      const saved = lsAvailable ? localStorage.getItem("grafik_data_v2") : null;
      if (saved) {
        const data = JSON.parse(saved);
        if (data.employees?.length) setEmployees(data.employees);
        if (data.shifts?.length) setShifts(data.shifts);
        if (data.stores?.length) setStores(data.stores);
        if (data.positions?.length) setPositions(data.positions);
        if (data.settings) setSettingsState(prev=>({...prev,...data.settings}));
        if (data.adjustments) setAdjustments(data.adjustments);
        setDataLoaded(true); return;
      }
    } catch(e){}
    if (storageAvailable) {
      (async()=>{
        try {
          const res = await window.storage.get("grafik_data_v1",true);
          if (res&&res.value) {
            const data=JSON.parse(res.value);
            if (data.employees?.length) setEmployees(data.employees);
            if (data.shifts?.length) setShifts(data.shifts);
            if (data.stores?.length) setStores(data.stores);
            if (data.positions?.length) setPositions(data.positions);
            if (data.settings) setSettingsState(prev=>({...prev,...data.settings}));
            if (data.adjustments) setAdjustments(data.adjustments);
          }
        } catch(e){}
        finally { setDataLoaded(true); }
      })();
    } else { setDataLoaded(true); }
  },[]); // eslint-disable-line react-hooks/exhaustive-deps

  // ---- Autozapis ----
  useEffect(()=>{
    if (!dataLoaded) return;
    setSaveState("saving");
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(()=>{
      try {
        const payload = JSON.stringify({employees,shifts,stores,positions,settings,adjustments});
        if (lsAvailable) localStorage.setItem("grafik_data_v2",payload);
        if (storageAvailable) window.storage.set("grafik_data_v1",payload,true).catch(()=>{});
        setSaveState("saved");
      } catch(e){ setSaveState("idle"); }
    },800);
    return ()=>clearTimeout(saveTimer.current);
  },[employees,shifts,stores,positions,settings,adjustments,dataLoaded]); // eslint-disable-line react-hooks/exhaustive-deps

  // ---- useMemo PRZED early returns (wymagane przez React Rules of Hooks) ----
  const weekDays = useMemo(()=>Array.from({length:7},(_,i)=>addDays(weekStart,i)),[weekStart]);
  const lastName=(n)=>n.trim().split(/\s+/).slice(-1)[0]||n;
  const storeEmployees = useMemo(()=>employees.filter(e=>e.storeId===effectiveStoreId).sort((a,b)=>lastName(a.name).localeCompare(lastName(b.name),"pl")||a.name.localeCompare(b.name,"pl")),[employees,effectiveStoreId]);

  const adjKey=(empId,d)=>`${empId}|${d.getFullYear()}-${pad2(d.getMonth()+1)}`;
  const defaultAdj=()=>({premia:0,ulga26:0,custIncLabel1:"",custIncAmount1:0,custIncLabel2:"",custIncAmount2:0,ubezpieczenie:0,custDecLabel1:"",custDecAmount1:0,custDecLabel2:"",custDecAmount2:0,transferAmount:0});
  const getAdj=(emp,d)=>({...defaultAdj(),...(adjustments[adjKey(emp.id,d)]||{})});

  const payrollRows = useMemo(()=>{
    const y=payMonth.getFullYear(), m=payMonth.getMonth();
    const lastDayOfMonth = new Date(y, m+1, 0); // ostatni dzień miesiąca
    const firstDayStr = `${y}-${pad2(m+1)}-01`;
    const lastDayStr = dateStr(lastDayOfMonth);
    return storeEmployees
      .filter(e => {
        // Uwzględnij pracownika jeśli jego umowa była aktywna w dowolnym dniu miesiąca
        if (!e.contractEndDate) return true; // brak daty zakończenia = aktywny
        return e.contractEndDate >= firstDayStr; // umowa zakończyła się po pierwszym dniu miesiąca
      })
      .map(e=>{
      const etatF=parseEtat(e.etat||"1");
      const hoursPerDay=settings.refHours*etatF;
      const empShifts=shifts.filter(s=>{
        if (s.employeeId!==e.id) return false;
        const d=new Date(s.date);
        if (d.getFullYear()!==y||d.getMonth()!==m) return false;
        // Nie licz zmian po dacie zakończenia umowy
        if (e.contractEndDate && s.date > e.contractEndDate) return false;
        return true;
      });
      let workMin=0,wolne=0,urlop=0,urlopBezplatny=0,l4=0,l4b=0,opiekaChore=0,opiekaZdrowe=0;
      const absenceDates={urlop:[],urlopBezplatny:[],l4:[],l4b:[],wolne:[],opiekaChore:[],opiekaZdrowe:[]};
      empShifts.forEach(s=>{
        if (s.type==="praca"&&s.start&&s.end) workMin+=Math.max(0,toMin(s.end)-toMin(s.start));
        else if (s.type==="wolne"){ wolne++; absenceDates.wolne.push(s.date); }
        else if (s.type==="urlop"){ urlop++; absenceDates.urlop.push(s.date); }
        else if (s.type==="urlop_bezplatny"){ urlopBezplatny++; absenceDates.urlopBezplatny.push(s.date); }
        else if (s.type==="l4"){ l4++; absenceDates.l4.push(s.date); }
        else if (s.type==="l4b"){ l4b++; absenceDates.l4b.push(s.date); }
        else if (s.type==="opieka_chore"){ opiekaChore++; absenceDates.opiekaChore.push(s.date); }
        else if (s.type==="opieka_zdrowe"){ opiekaZdrowe++; absenceDates.opiekaZdrowe.push(s.date); }
      });
      const hours=workMin/60;
      const urlopHours=urlop*hoursPerDay, l4Hours=l4*hoursPerDay, l4bHours=l4b*hoursPerDay;
      const urlopBezplatnyHours=urlopBezplatny*hoursPerDay;
      const opiekaChoreHours=opiekaChore*hoursPerDay, opiekaZdroweHours=opiekaZdrowe*hoursPerDay;
      const totalHours=hours+urlopHours+l4Hours+l4bHours+opiekaChoreHours+opiekaZdroweHours;
      const workPay=hours*e.rate, urlopPay=urlopHours*e.rate;
      const l4Pay=l4Hours*e.rate*(settings.l4Percent/100), l4bPay=l4bHours*e.rate;
      const opiekaChorePay=opiekaChoreHours*e.rate*((settings.opiekaChorePercent||80)/100);
      const opiekaZdrowePay=opiekaZdroweHours*e.rate;
      const adj=getAdj(e,payMonth);
      const increase=(adj.premia||0)+(adj.ulga26||0)+(adj.custIncAmount1||0)+(adj.custIncAmount2||0);
      const decrease=(adj.ubezpieczenie||0)+(adj.custDecAmount1||0)+(adj.custDecAmount2||0);
      const pay=Math.max(0,workPay+urlopPay+l4Pay+l4bPay+opiekaChorePay+opiekaZdrowePay+increase-decrease);
      const method=e.paymentMethod||"przelew";
      let transfer=0,cash=0;
      if (method==="przelew"){ transfer=pay; cash=0; }
      else if (method==="gotowka"){ transfer=0; cash=pay; }
      else { transfer=adj.transferAmount||0; cash=pay-transfer; }
      return {e,etatF,hoursPerDay,hours,urlopHours,l4Hours,l4bHours,urlopBezplatnyHours,opiekaChoreHours,opiekaZdroweHours,totalHours,wolne,urlop,urlopBezplatny,l4,l4b,opiekaChore,opiekaZdrowe,workPay,urlopPay,l4Pay,l4bPay,opiekaChorePay,opiekaZdrowePay,increase,decrease,pay,transfer,cash,adj,absenceDates,empShifts};
    });
  },[storeEmployees,shifts,payMonth,settings,adjustments]); // eslint-disable-line react-hooks/exhaustive-deps

  const totalPay=payrollRows.reduce((a,r)=>a+r.pay,0);
  const totalHours=payrollRows.reduce((a,r)=>a+r.hours,0);

  // ---- Funkcje operacyjne ----

  const saveShift=(data)=>{
    const newShift=data.id?data:{...data,id:uid()};
    if (data.id) setShifts(prev=>prev.map(s=>s.id===data.id?{...s,...data}:s));
    else setShifts(prev=>[...prev.filter(s=>!(s.employeeId===data.employeeId&&s.date===data.date)),{...data,id:newShift.id}]);
    if (electronAvailable) window.api.shifts.upsertDay(newShift.employeeId,newShift.date,shiftToDb(newShift));
    setModal(null);
  };
  const deleteShift=(id)=>{
    const s=shifts.find(x=>x.id===id);
    setShifts(prev=>prev.filter(s=>s.id!==id));
    if (electronAvailable&&s) window.api.shifts.upsertDay(s.employeeId,s.date,null);
    setModal(null);
  };
  const updateShiftTimes=(id,start,end)=>{
    setShifts(prev=>prev.map(s=>s.id===id?{...s,start,end}:s));
    if (electronAvailable) window.api.shifts.updateTimes(id,start,end);
  };
  const applyBulkTask=({date,template,task})=>{
    if (!task.trim()) return;
    const tStart=toMin(template.start),tEnd=toMin(template.end);
    setShifts(prev=>prev.map(s=>{
      if (s.date!==date||s.type!=="praca"||!s.start) return s;
      const inStore=employees.find(e=>e.id===s.employeeId)?.storeId===effectiveStoreId;
      if (toMin(s.start)<tEnd&&toMin(s.end)>tStart&&inStore){
        const updated={...s,tasks:[...s.tasks,task.trim()]};
        if (electronAvailable) window.api.shifts.upsertDay(updated.employeeId,updated.date,shiftToDb(updated));
        return updated;
      }
      return s;
    }));
    setBulkModal(false);
  };
  const copyDay=(sourceDateStr,targetDateStr)=>{
    setShifts(prev=>{
      const empIds=employees.filter(e=>e.storeId===effectiveStoreId).map(e=>e.id);
      const src=prev.filter(s=>s.date===sourceDateStr&&empIds.includes(s.employeeId));
      const rest=prev.filter(s=>!(s.date===targetDateStr&&empIds.includes(s.employeeId)));
      const copied=src.map(s=>{const n={...s,id:uid(),date:targetDateStr};if(electronAvailable)window.api.shifts.upsertDay(n.employeeId,n.date,shiftToDb(n));return n;});
      return [...rest,...copied];
    });
  };
  const copyWeek=(sourceWeekStartDate,targetWeekStartDate)=>{
    setShifts(prev=>{
      const empIds=employees.filter(e=>e.storeId===effectiveStoreId).map(e=>e.id);
      let result=[...prev];
      for (let i=0;i<7;i++){
        const srcDate=dateStr(addDays(sourceWeekStartDate,i)),tgtDate=dateStr(addDays(targetWeekStartDate,i));
        const src=result.filter(s=>s.date===srcDate&&empIds.includes(s.employeeId));
        result=result.filter(s=>!(s.date===tgtDate&&empIds.includes(s.employeeId)));
        const copied=src.map(s=>{const n={...s,id:uid(),date:tgtDate};if(electronAvailable)window.api.shifts.upsertDay(n.employeeId,n.date,shiftToDb(n));return n;});
        result=[...result,...copied];
      }
      return result;
    });
  };
  const saveEmployee=(data)=>{
    const emp=data.id?data:{storeId:effectiveStoreId,...data,id:uid()};
    if (data.id) setEmployees(prev=>prev.map(e=>e.id===data.id?{...e,...data}:e));
    else setEmployees(prev=>[...prev,emp]);
    if (electronAvailable) window.api.employees.save(empToDb(emp));
    setEmpModal(null);
  };
  const deleteEmployee=(id)=>{
    setEmployees(prev=>prev.filter(e=>e.id!==id));
    setShifts(prev=>prev.filter(s=>s.employeeId!==id));
    if (electronAvailable) window.api.employees.delete(id);
    setEmpModal(null);
  };
  const renameStore=(id,name)=>{ setStores(prev=>prev.map(s=>s.id===id?{...s,name}:s)); if(electronAvailable) window.api.stores.rename(id,name); };
  const addPosition=(name)=>{ if(name.trim()&&!positions.includes(name.trim())){ setPositions(prev=>[...prev,name.trim()]); if(electronAvailable) window.api.positions.add(name.trim()); } };
  const removePosition=(name)=>{ setPositions(prev=>prev.filter(p=>p!==name)); if(electronAvailable) window.api.positions.remove(name); };
  const setSettings=(s)=>{ setSettingsState(s); if(electronAvailable){ window.api.settings.set("l4_percent",String(s.l4Percent)); window.api.settings.set("ref_hours",String(s.refHours)); } };
  const saveAdj=(empId,d,data)=>{
    setAdjustments(prev=>({...prev,[adjKey(empId,d)]:data}));
    if (electronAvailable) window.api.adjustments.save({employee_id:empId,year:d.getFullYear(),month:d.getMonth()+1,premia:data.premia||0,ulga26:data.ulga26||0,cust_inc_label1:data.custIncLabel1||"",cust_inc_amount1:data.custIncAmount1||0,cust_inc_label2:data.custIncLabel2||"",cust_inc_amount2:data.custIncAmount2||0,ubezpieczenie:data.ubezpieczenie||0,cust_dec_label1:data.custDecLabel1||"",cust_dec_amount1:data.custDecAmount1||0,cust_dec_label2:data.custDecLabel2||"",cust_dec_amount2:data.custDecAmount2||0,transfer_amount:data.transferAmount||0});
    setAdjModal(null);
  };

  // ---- Auth & early returns ----
  const handleLogout=()=>setCurrentUser(null);

  if (!dataLoaded) return <div style={{minHeight:"100vh",background:"#162820",display:"flex",alignItems:"center",justifyContent:"center",color:"#C8D8CC",fontFamily:"ui-sans-serif,system-ui,sans-serif",fontSize:16,gap:12}}><div style={{width:8,height:8,borderRadius:"50%",background:"#5EC98A",animation:"pulse 1.2s infinite"}}/>Wczytywanie danych...</div>;
  if (!currentUser) return <LoginScreen onLogin={(r)=>setCurrentUser({role:r})} storeName={currentStoreName} adminHash={settings.adminPasswordHash} managerHash={settings.managerPasswordHash}/>;

  // ---- Render główny ----
  return (
    <div style={{fontFamily:"ui-sans-serif,system-ui,-apple-system,'Segoe UI',Roboto,sans-serif",color:"#1A2530",height:"100vh",overflow:"hidden"}}>
      <style>{`
        *,*::before,*::after{box-sizing:border-box;}
        @media print{
          body *{visibility:hidden;}
          #print-area-wrapper{visibility:visible!important;position:absolute!important;top:0!important;left:0!important;width:100%!important;height:auto!important;overflow:visible!important;z-index:9999!important;background:#fff!important;}
          #print-area-wrapper *{visibility:visible!important;-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important;color-adjust:exact!important;}
          .no-print{display:none!important;visibility:hidden!important;}
          .print-employee-block{page-break-inside:avoid!important;break-inside:avoid!important;}
          @page{size:A4 landscape;margin:10mm;}
        }
      `}</style>

      {/* === SIDEBAR LAYOUT === */}
      <div style={{display:"flex", height:"100vh", overflow:"hidden"}}>

        {/* LEFT SIDEBAR */}
        <div className="no-print" style={{width:220,background:"#162820",color:"#C8D8CC",display:"flex",flexDirection:"column",flexShrink:0,zIndex:10}}>
          {/* Logo */}
          <div style={{padding:"18px 16px 16px",borderBottom:"1px solid rgba(255,255,255,0.07)"}}>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              <div style={{width:38,height:38,borderRadius:10,background:"#E07A3E",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,fontSize:18,color:"#fff",flexShrink:0,boxShadow:"0 2px 8px rgba(224,122,62,0.35)"}}>G</div>
              <div style={{minWidth:0}}>
                <div style={{fontWeight:700,fontSize:13.5,color:"#fff",lineHeight:1.2,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{currentStoreName}</div>
                <div style={{fontSize:10.5,color:"rgba(255,255,255,0.4)",marginTop:1}}>Grafik Sklepu</div>
              </div>
            </div>
          </div>

          {/* Nav items */}
          <nav style={{flex:1,padding:"10px 8px",display:"flex",flexDirection:"column",gap:2,overflowY:"auto"}}>
            <div style={{fontSize:10,fontWeight:700,color:"rgba(255,255,255,0.3)",letterSpacing:"0.08em",padding:"8px 10px 4px",textTransform:"uppercase"}}>Menu</div>
            {[{id:"grafik",label:"Grafik pracy",icon:Calendar},{id:"pracownicy",label:"Pracownicy",icon:Users},{id:"rozliczenia",label:"Rozliczenia",icon:Wallet},{id:"ustawienia",label:"Ustawienia",icon:Settings}].map(t=>{
              const active=tab===t.id;
              return (
                <button key={t.id} onClick={()=>setTab(t.id)} style={{display:"flex",alignItems:"center",gap:10,padding:"9px 12px",border:"none",borderRadius:8,background:active?"rgba(255,255,255,0.1)":"transparent",color:active?"#fff":"rgba(255,255,255,0.55)",fontWeight:active?600:400,fontSize:13.5,cursor:"pointer",textAlign:"left",width:"100%",transition:"background 0.15s",position:"relative"}}>
                  {active&&<span style={{position:"absolute",left:0,top:"50%",transform:"translateY(-50%)",width:3,height:20,borderRadius:"0 3px 3px 0",background:"#5EC98A"}}/>}
                  <t.icon size={17}/>{t.label}
                </button>
              );
            })}
          </nav>

          {/* Bottom: status + user + logout */}
          <div style={{padding:"10px 8px 14px",borderTop:"1px solid rgba(255,255,255,0.07)"}}>
            {dataLoaded&&(
              <div style={{display:"flex",alignItems:"center",gap:5,fontSize:11,color:"rgba(255,255,255,0.38)",padding:"0 10px",marginBottom:8}}>
                <span style={{width:6,height:6,borderRadius:"50%",background:saveState==="saving"?"#E07A3E":"#5EC98A",display:"inline-block",flexShrink:0}}/>
                {saveState==="saving"?"Zapisywanie...":"Dane zapisane"}
              </div>
            )}
            <div style={{display:"flex",alignItems:"center",gap:8,padding:"8px 10px",background:"rgba(255,255,255,0.06)",borderRadius:8,marginBottom:4}}>
              <div style={{width:30,height:30,borderRadius:"50%",background:isAdmin?"#1F6B45":"#3A5068",display:"flex",alignItems:"center",justifyContent:"center",fontSize:12,fontWeight:700,color:"#fff",flexShrink:0}}>
                {isAdmin?"A":"K"}
              </div>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontSize:12,color:"rgba(255,255,255,0.75)",fontWeight:600,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{isAdmin?"Administrator":"Kierownik"}</div>
                <div style={{display:"flex",alignItems:"center",gap:4,marginTop:1}}>{isAdmin?<ShieldCheck size={10} color="#5EC98A"/>:<Lock size={10} color="#9399A6"/>}<span style={{fontSize:10,color:"rgba(255,255,255,0.35)"}}>{isAdmin?"Pełny dostęp":"Ograniczony"}</span></div>
              </div>
            </div>
            <button onClick={handleLogout} style={{display:"flex",alignItems:"center",gap:7,width:"100%",padding:"7px 10px",border:"none",background:"transparent",color:"rgba(255,255,255,0.4)",fontSize:12.5,cursor:"pointer",borderRadius:7,textAlign:"left"}}>
              <LogOut size={14}/> Wyloguj się
            </button>
          </div>
        </div>

        {/* MAIN AREA */}
        <div style={{flex:1,display:"flex",flexDirection:"column",minWidth:0,background:"#F1F4F8"}}>

          {/* TOP BAR */}
          <div className="no-print" style={{background:"#fff",borderBottom:"1px solid #E4E7EE",padding:"0 24px",display:"flex",alignItems:"center",justifyContent:"space-between",height:56,flexShrink:0,boxShadow:"0 1px 3px rgba(0,0,0,0.04)"}}>
            <div style={{fontWeight:700,fontSize:17,color:"#1A2530"}}>
              {tab==="grafik"?"Grafik pracy":tab==="pracownicy"?"Pracownicy":tab==="rozliczenia"?"Rozliczenia":tab==="ustawienia"?"Ustawienia":""}
            </div>
            <div style={{display:"flex",alignItems:"center",gap:6,fontSize:12.5,color:"#8890A0",fontWeight:500}}>
              {isAdmin?<ShieldCheck size={14} color="#1F6B45"/>:<Lock size={14} color="#9399A6"/>}
              <span>{isAdmin?"Administrator":"Kierownik sklepu"}</span>
            </div>
          </div>

          {/* CONTENT */}
          <div style={{flex:1,overflow:"auto",padding:24}}>
            {tab==="grafik"&&(
              <ErrorBoundary>
                <ScheduleView viewMode={viewMode} setViewMode={setViewMode} weekDays={weekDays} setWeekStart={setWeekStart} weekStart={weekStart} selectedDate={selectedDate} setSelectedDate={setSelectedDate} employees={storeEmployees} shifts={shifts} setModal={setModal} setBulkModal={setBulkModal} storeName={currentStoreName} setCopyDayModal={setCopyDayModal} setCopyWeekModal={setCopyWeekModal} onUpdateTimes={updateShiftTimes}/>
              </ErrorBoundary>
            )}
            {tab==="pracownicy"&&<EmployeesView employees={storeEmployees} setEmpModal={setEmpModal} shifts={shifts} isAdmin={isAdmin} positions={positions}/>}
            {tab==="rozliczenia"&&<PayrollView payMonth={payMonth} setPayMonth={setPayMonth} rows={payrollRows} totalPay={totalPay} totalHours={totalHours} isAdmin={isAdmin} settings={settings} onOpenAdj={(emp)=>setAdjModal({emp,date:payMonth})} storeName={currentStoreName} shifts={shifts} electronAvailable={electronAvailable}/>}
            {tab==="ustawienia"&&<SettingsView settings={settings} setSettings={setSettings} positions={positions} addPosition={addPosition} removePosition={removePosition} isAdmin={isAdmin}/>}
          </div>
        </div>
      </div>

      {modal&&<ShiftModal modal={modal} onClose={()=>setModal(null)} onSave={saveShift} onDelete={deleteShift} employeeName={employees.find(e=>e.id===modal.employeeId)?.name}/>}
      {bulkModal&&<BulkTaskModal weekDays={weekDays} onClose={()=>setBulkModal(false)} onApply={applyBulkTask}/>}
      {copyDayModal&&<CopyDayModal sourceDate={selectedDate} onClose={()=>setCopyDayModal(false)} onApply={(src,tgt)=>{copyDay(src,tgt);setSelectedDate(new Date(tgt+"T12:00:00"));setViewMode("day");}}/>}
      {copyWeekModal&&<CopyWeekModal sourceWeekStart={weekStart} onClose={()=>setCopyWeekModal(false)} onApply={copyWeek}/>}
      {empModal&&<EmployeeModal modal={empModal} onClose={()=>setEmpModal(null)} onSave={saveEmployee} onDelete={deleteEmployee} isAdmin={isAdmin} positions={positions}/>}
      {adjModal&&<AdjustmentsModal emp={adjModal.emp} date={adjModal.date} initial={getAdj(adjModal.emp,adjModal.date)} onClose={()=>setAdjModal(null)} onSave={(data)=>saveAdj(adjModal.emp.id,adjModal.date,data)} isAdmin={isAdmin}/>}
    </div>
  );
}

// ============ WIDOK GRAFIKU ============
function ScheduleView({ viewMode, setViewMode, weekDays, setWeekStart, weekStart, selectedDate, setSelectedDate, employees, shifts, setModal, setBulkModal, storeName, setCopyDayModal, setCopyWeekModal, onUpdateTimes }) {
  const [dayPrintPreview, setDayPrintPreview] = useState(false);
  const [showDatePicker, setShowDatePicker] = useState(false);
  const datePickerRef = React.useRef(null);

  const goToDate = (val) => {
    if (!val) return;
    const d = new Date(val + "T12:00:00");
    setSelectedDate(d);
    setWeekStart(startOfWeek(d));
    setShowDatePicker(false);
  };

  return (
    <div>
      <div className="no-print" style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:16, flexWrap:"wrap", gap:10 }}>
        <div style={{ display:"flex", alignItems:"center", gap:10, flexWrap:"wrap" }}>
          <div style={{ display:"flex", border:"1px solid #DDE0E8", borderRadius:8, overflow:"hidden", background:"#fff" }}>
            {["week","day"].map(m=>(
              <button key={m} onClick={()=>setViewMode(m)} style={{ padding:"7px 14px", border:"none", background:viewMode===m?"#1F6B45":"transparent", color:viewMode===m?"#fff":"#5A6070", fontWeight:600, fontSize:12.5, cursor:"pointer" }}>{m==="week"?"Tydzień":"Dzień"}</button>
            ))}
          </div>
          {viewMode==="week" ? (
            <div style={{ display:"flex", alignItems:"center", gap:8 }}>
              <button onClick={()=>setWeekStart(addDays(weekStart,-7))} style={navBtn}><ChevronLeft size={18}/></button>
              <div style={{ fontWeight:700, fontSize:14, minWidth:200, textAlign:"center" }}>{weekDays[0].toLocaleDateString("pl-PL",{day:"numeric",month:"short"})} - {weekDays[6].toLocaleDateString("pl-PL",{day:"numeric",month:"short",year:"numeric"})}</div>
              <button onClick={()=>setWeekStart(addDays(weekStart,7))} style={navBtn}><ChevronRight size={18}/></button>
              <button onClick={()=>setWeekStart(startOfWeek(new Date()))} style={{ ...navBtn, width:"auto", padding:"0 12px", fontSize:12 }}>Dzis</button>
            </div>
          ) : (
            <div style={{ display:"flex", alignItems:"center", gap:8 }}>
              <button onClick={()=>setSelectedDate(addDays(selectedDate,-1))} style={navBtn}><ChevronLeft size={18}/></button>
              <div style={{ fontWeight:700, fontSize:14, minWidth:200, textAlign:"center", textTransform:"capitalize" }}>{dayLabelFull(selectedDate)}</div>
              <button onClick={()=>setSelectedDate(addDays(selectedDate,1))} style={navBtn}><ChevronRight size={18}/></button>
              <button onClick={()=>setSelectedDate(new Date())} style={{ ...navBtn, width:"auto", padding:"0 12px", fontSize:12 }}>Dzis</button>
            </div>
          )}
          {/* Ikona kalendarza - szybkie przejscie do daty */}
          <div style={{ position:"relative" }} ref={datePickerRef}>
            <button onClick={()=>setShowDatePicker(v=>!v)} style={{ ...navBtn, width:36, background:showDatePicker?"#1F6B45":"#fff", color:showDatePicker?"#fff":"#5A6070" }} title="Przejdz do daty">
              <Calendar size={16}/>
            </button>
            {showDatePicker && (
              <div style={{ position:"absolute", top:"calc(100% + 6px)", left:"50%", transform:"translateX(-50%)", background:"#fff", border:"1px solid #DDE0E8", borderRadius:10, boxShadow:"0 6px 24px rgba(0,0,0,0.12)", padding:14, zIndex:200, minWidth:200 }}>
                <div style={{ fontSize:11.5, fontWeight:600, color:"#5A6070", marginBottom:8 }}>Przejdź do dnia:</div>
                <input
                  type="date"
                  defaultValue={dateStr(viewMode==="week"?weekDays[0]:selectedDate)}
                  autoFocus
                  onChange={e=>goToDate(e.target.value)}
                  style={{ ...inputStyle, width:"100%", fontSize:14 }}
                />
                <div style={{ fontSize:11, color:"#9399A6", marginTop:6 }}>Wybierz datę i przejdziesz od razu do tego dnia.</div>
              </div>
            )}
          </div>
        </div>
        <div style={{ display:"flex", gap:8 }}>
          {viewMode==="day" && <>
            <button onClick={()=>setCopyDayModal(true)} style={ghostBtn}><Copy size={14} style={{ marginRight:6 }}/>Kopiuj dzień</button>
            <button onClick={()=>setDayPrintPreview(true)} style={ghostBtn}><Printer size={14} style={{ marginRight:6 }}/>Podgląd wydruku A4</button>
          </>}
          {viewMode==="week" && <button onClick={()=>setCopyWeekModal(true)} style={ghostBtn}><Copy size={14} style={{ marginRight:6 }}/>Kopiuj tydzień</button>}
          <button onClick={()=>setBulkModal(true)} style={primaryBtn}><Megaphone size={15}/> Zadanie dla całej zmiany</button>
        </div>
      </div>

      <div className="no-print" style={{ display:"flex", gap:14, marginBottom:14, fontSize:12, color:"#8890A0", background:"#fff", border:"1px solid #E4E7EE", borderRadius:8, padding:"8px 14px" }}>
        {Object.entries(TYPES).map(([k,v])=>(
          <div key={k} style={{ display:"flex", alignItems:"center", gap:5 }}>
            <span style={{ width:10, height:10, borderRadius:3, background:v.bg, display:"inline-block" }}/>{v.label}
          </div>
        ))}
      </div>

      {viewMode==="week" ? <WeekGrid weekDays={weekDays} employees={employees} shifts={shifts} setModal={setModal}/> : <DayList date={selectedDate} employees={employees} shifts={shifts} setModal={setModal} storeName={storeName} onUpdateTimes={onUpdateTimes}/>}

      {dayPrintPreview && (
        <div id="print-area-wrapper" style={{ position:"fixed", inset:0, background:"#fff", zIndex:100, overflowY:"auto", padding:0 }}>
          <div id="print-area" style={{ background:"#fff", padding:"10mm 14mm" }}>
            <div className="no-print" style={{ position:"sticky", top:0, background:"#fff", borderBottom:"1px solid #E3DFD2", padding:"10px 0", marginBottom:16, display:"flex", gap:10, alignItems:"center", zIndex:101 }}>
              <button onClick={()=>window.print()} style={{ ...primaryBtn, gap:8 }}><Printer size={15}/> Drukuj</button>
              <button onClick={()=>setDayPrintPreview(false)} style={ghostBtn}>✕ Zamknij podgląd</button>
              <span style={{ fontSize:12, color:"#9C988A" }}>Wydruk poziomy (A4 landscape) · cala lista pracownikow · pracownicy nie beda lacieni miedzy stronami</span>
            </div>
            <DayList date={selectedDate} employees={employees} shifts={shifts} setModal={null} storeName={storeName} onUpdateTimes={onUpdateTimes} printMode={true}/>
          </div>
        </div>
      )}
    </div>
  );
}

function WeekGrid({ weekDays, employees, shifts, setModal }) {
  const weekEndDate = dateStr(weekDays[6]);
  const weekStartDate = dateStr(weekDays[0]);
  // Pokaż pracownika jeśli umowa była aktywna przynajmniej w jednym dniu tygodnia
  const visibleEmployees = employees.filter(e =>
    !e.contractEndDate || e.contractEndDate >= weekStartDate
  );
  return (
    <div style={{ background:"#fff", border:"1px solid #E4E7EE", borderRadius:12, overflow:"hidden", boxShadow:"0 1px 4px rgba(0,0,0,0.05)" }}>
      <div style={{ display:"grid", gridTemplateColumns:"160px repeat(7, 1fr)", borderBottom:"1px solid #E4E7EE", background:"#F7F9FC" }}>
        <div style={{ padding:"10px 12px", fontSize:12, fontWeight:700, color:"#8890A0" }}>Pracownik</div>
        {weekDays.map((d,i)=>(
          <div key={i} style={{ padding:"10px 8px", textAlign:"center", borderLeft:"1px solid #EEF0F5" }}>
            <div style={{ fontSize:12, fontWeight:700, color:"#1F6B45" }}>{DAY_NAMES[i]}</div>
            <div style={{ fontSize:11, color:"#9399A6" }}>{d.getDate()}.{pad2(d.getMonth()+1)}</div>
          </div>
        ))}
      </div>
      {visibleEmployees.length===0 && <div style={{ padding:28, textAlign:"center", color:"#9399A6", fontSize:13 }}>Brak pracowników. Dodaj ich w zakładce „Pracownicy".</div>}
      {visibleEmployees.map(emp=>(
        <div key={emp.id} style={{ display:"grid", gridTemplateColumns:"160px repeat(7, 1fr)", borderBottom:"1px solid #F1F3F7" }}>
          <div style={{ padding:"10px 12px", fontSize:13, fontWeight:600, display:"flex", alignItems:"center", gap:6, color:"#1A2530" }}>
            {displayName(emp.name)}
            {emp.contractEndDate && emp.contractEndDate <= weekEndDate && <span style={{ fontSize:9.5, color:"#C23B3B", fontWeight:700, background:"#FEF2F2", padding:"1px 5px", borderRadius:4 }}>do {emp.contractEndDate.slice(8)}.{emp.contractEndDate.slice(5,7)}</span>}
          </div>
          {weekDays.map((d,i)=>{
            const ds=dateStr(d);
            // Dzień po zakończeniu umowy - wyszarz
            const afterContract = emp.contractEndDate && ds > emp.contractEndDate;
            const shift=shifts.find(s=>s.employeeId===emp.id&&s.date===ds);
            if (afterContract) return (
              <div key={i} style={{ borderLeft:"1px solid #F1F3F7", minHeight:56, padding:4, background:"#F7F9FC", display:"flex", alignItems:"center", justifyContent:"center" }}>
                <span style={{ fontSize:10, color:"#C8CDD8" }}>-</span>
              </div>
            );
            return (
              <div key={i} onClick={()=>setModal({employeeId:emp.id,date:ds,shift})} style={{ borderLeft:"1px solid #F1F3F7", minHeight:56, padding:4, cursor:"pointer" }}>
                {!shift && <div style={{ height:48, display:"flex", alignItems:"center", justifyContent:"center", color:"#C8CDD8" }}><Plus size={14}/></div>}
                {shift && (()=>{
                  const t=TYPES[shift.type];
                  if (shift.type==="praca"&&shift.start&&shift.end) return (
                    <div title={`${shift.start}-${shift.end}`} style={{ height:28, borderRadius:6, background:t.bg, color:t.text, fontSize:10.5, fontWeight:600, display:"flex", alignItems:"center", padding:"0 6px", overflow:"hidden", whiteSpace:"nowrap" }}>
                      {shift.start}-{shift.end}
                      {shift.tasks.length>0 && <span style={{ marginLeft:5, background:"rgba(255,255,255,0.25)", borderRadius:4, padding:"0 4px" }}>{shift.tasks.length} zad.</span>}
                    </div>
                  );
                  return <div style={{ height:24, borderRadius:6, background:t.bg, color:t.text, fontSize:10.5, fontWeight:700, display:"flex", alignItems:"center", justifyContent:"center" }}>{t.label}</div>;
                })()}
              </div>
            );
          })}
        </div>
      ))}
      <div style={{ display:"grid", gridTemplateColumns:"160px repeat(7, 1fr)", background:"#FAF8F2" }}>
        <div/>
        {weekDays.map((_,i)=>(
          <div key={i} style={{ borderLeft:"1px solid #F1EEE3", padding:"4px 6px", fontSize:9, color:"#B4B2A9", display:"flex", justifyContent:"space-between" }}>
            <span>5</span><span>13</span><span>22</span>
          </div>
        ))}
      </div>
    </div>
  );
}

// ============ WIDOK DZIENNY ============
function DayList({ date, employees, shifts, setModal, storeName, onUpdateTimes, printMode }) {
  if (!date || !(date instanceof Date) || isNaN(date.getTime())) return <div style={{ padding:20, color:"#C75146" }}>Nieprawidłowa data.</div>;
  const ds = dateStr(date);
  const totalHoursAxis = DAY_END - DAY_START;
  const hourMarks = Array.from({length:totalHoursAxis+1},(_,i)=>DAY_START+i);
  const dayShifts = employees
    .filter(e => !e.contractEndDate || e.contractEndDate >= ds) // ukryj po zakończeniu umowy
    .map(emp=>({ emp, shift:shifts.find(s=>s.employeeId===emp.id&&s.date===ds) }))
    .sort((a,b)=>{const ln=(n)=>n.trim().split(/\s+/).slice(-1)[0]||n; return ln(a.emp.name).localeCompare(ln(b.emp.name),"pl")||a.emp.name.localeCompare(b.emp.name,"pl");});

  return (
    <div id="print-area" style={{ background:"#fff", border:"1px solid #E3DFD2", borderRadius:12, padding:18 }}>
      <div style={{ marginBottom:14, display:"flex", justifyContent:"space-between", alignItems:"flex-start", flexWrap:"wrap", gap:8 }}>
        <div>
          <div style={{ fontWeight:700, fontSize:17 }}>{storeName} -- grafik dnia</div>
          <div style={{ fontSize:13, color:"#6B6A60", textTransform:"capitalize" }}>{dayLabelFull(date)}</div>
        </div>
        <div className="no-print" style={{ fontSize:11, color:"#B4B2A9" }}>Przeciągnij pasek by zmienić godziny · kliknij by edytować</div>
      </div>
      {dayShifts.length===0 && <div style={{ color:"#9C988A", fontSize:13 }}>Brak pracowników.</div>}
      {dayShifts.length>0 && (
        <div>
          <div style={{ display:"grid", gridTemplateColumns:"150px 1fr", marginBottom:4, pageBreakAfter:"avoid", breakAfter:"avoid" }}>
            <div/>
            <div style={{ position:"relative", height:16 }}>
              {hourMarks.map(h=>(
                <div key={h} style={{ position:"absolute", left:`${((h-DAY_START)/totalHoursAxis)*100}%`, transform:h===DAY_START?"translateX(0)":h===DAY_END?"translateX(-100%)":"translateX(-50%)", fontSize:10, color:"#B4B2A9" }}>{h}</div>
              ))}
            </div>
          </div>
          <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
            {dayShifts.map(({emp,shift})=>{
              const t=TYPES[shift?.type||"wolne"];
              const isWork=shift&&shift.type==="praca"&&shift.start&&shift.end;
              return (
                <div key={emp.id} className="print-employee-block" style={{ pageBreakInside:"avoid", breakInside:"avoid" }}>
                  <div style={{ display:"grid", gridTemplateColumns:"150px 1fr", alignItems:"center", gap:10 }}>
                    <div onClick={printMode ? undefined : ()=>setModal({employeeId:emp.id,date:ds,shift})} style={{ cursor:printMode?"default":"pointer" }}>
                      <div style={{ fontWeight:700, fontSize:13 }}>{displayName(emp.name)}</div>
                      {isWork ? <div style={{ fontSize:11.5, color:"#6B6A60", display:"flex", alignItems:"center", gap:4 }}><Clock size={12}/>{shift.start}-{shift.end}</div>
                              : <div style={{ fontSize:11, fontWeight:700, color:t.bg }}>{t.label}</div>}
                    </div>
                    {isWork ? (
                      <PrintableBar shift={shift} type={t} hourMarks={hourMarks} totalHoursAxis={totalHoursAxis} printMode={printMode} onDrag={printMode ? undefined : (s,e)=>onUpdateTimes(shift.id,s,e)} onOpenModal={printMode ? undefined : ()=>setModal({employeeId:emp.id,date:ds,shift})}/>
                    ) : (
                      <div onClick={()=>setModal({employeeId:emp.id,date:ds,shift})} style={{ position:"relative", height:26, background:"#FAF8F2", borderRadius:6, cursor:"pointer" }}>
                        {hourMarks.map(h=><div key={h} style={{ position:"absolute", left:`${((h-DAY_START)/totalHoursAxis)*100}%`, top:0, bottom:0, width:1, background:h%4===0?"#EFEBDE":"transparent" }}/>)}
                      </div>
                    )}
                  </div>
                  {isWork && shift.tasks?.length>0 && <ul style={{ margin:"6px 0 0 160px", paddingLeft:18, fontSize:12, color:"#3a392f", lineHeight:1.6 }}>{shift.tasks.map((task,i)=><li key={i}>{task}</li>)}</ul>}
                  {isWork && (!shift.tasks||shift.tasks.length===0) && <div style={{ margin:"6px 0 0 160px", fontSize:11.5, color:"#B4B2A9" }}>Brak przypisanych zadań.</div>}
                </div>
              );
            })}
          </div>
          <div style={{ marginTop:24, borderTop:"1px solid #EFEBDE", paddingTop:14, display:"flex", justifyContent:"space-between", fontSize:12, color:"#6B6A60" }}>
            <div style={{ borderTop:"1px solid #3a392f", paddingTop:6, width:"40%" }}>Podpis pracownika</div>
            <div style={{ borderTop:"1px solid #3a392f", paddingTop:6, width:"40%", textAlign:"right" }}>Podpis kierownika</div>
          </div>
        </div>
      )}
    </div>
  );
}

// ============ PASEK STATYCZNY (do wydruku) i INTERAKTYWNY (DragBar) ============
function PrintableBar({ shift, type, hourMarks, totalHoursAxis, printMode, onDrag, onOpenModal }) {
  if (printMode) {
    // W trybie podglądu/wydruku - statyczny pasek, zawsze widoczny
    const left = ((toMin(shift.start) - DAY_START*60) / (totalHoursAxis*60)) * 100;
    const width = ((toMin(shift.end) - toMin(shift.start)) / (totalHoursAxis*60)) * 100;
    return (
      <div style={{ position:"relative", height:26, background:"#F0EEE8", borderRadius:6 }}>
        {hourMarks.map(h=><div key={h} style={{ position:"absolute", left:`${((h-DAY_START)/totalHoursAxis)*100}%`, top:0, bottom:0, width:1, background:h%4===0?"#DEDBD0":"transparent" }}/>)}
        <div style={{ position:"absolute", top:2, bottom:2, left:`${left}%`, width:`${width}%`, background:type.bg, borderRadius:4, display:"flex", alignItems:"center", paddingLeft:8, fontSize:10.5, fontWeight:700, color:type.text, overflow:"hidden", whiteSpace:"nowrap" }}>
          {shift.start}-{shift.end}
        </div>
      </div>
    );
  }
  return <DragBar shift={shift} type={type} hourMarks={hourMarks} totalHoursAxis={totalHoursAxis} onChange={onDrag} onOpenModal={onOpenModal}/>;
}

// ============ DRAGBAR (interaktywny) ============
function DragBar({ shift, type, hourMarks, totalHoursAxis, onChange, onOpenModal }) {
  const trackRef = useRef(null);
  const dragRef = useRef(null);
  const liveRef = useRef({ start:shift.start, end:shift.end });
  const [live, setLive] = useState({ start:shift.start, end:shift.end });
  useEffect(()=>{ liveRef.current={start:shift.start,end:shift.end}; setLive({start:shift.start,end:shift.end}); },[shift.start,shift.end]);
  const dayStartMin=DAY_START*60, dayEndMin=DAY_END*60;
  const minToPct=(min)=>((min-dayStartMin)/(totalHoursAxis*60))*100;
  const snap=(min)=>Math.round(min/SNAP_MIN)*SNAP_MIN;
  const onMouseMove=useRef(null), onMouseUp=useRef(null);
  onMouseMove.current=(e)=>{
    const d=dragRef.current; if(!d||!trackRef.current) return;
    if(Math.abs(e.clientX-d.startX)>2) d.moved=true;
    const rect=trackRef.current.getBoundingClientRect();
    const deltaMin=((e.clientX-d.startX)/rect.width)*totalHoursAxis*60;
    let ns=d.origStartMin, ne=d.origEndMin;
    if(d.mode==="move"){const dur=ne-ns;ns=snap(ns+deltaMin);ns=Math.max(dayStartMin,Math.min(dayEndMin-dur,ns));ne=ns+dur;}
    else if(d.mode==="resize-left"){ns=snap(ns+deltaMin);ns=Math.max(dayStartMin,Math.min(ne-MIN_DURATION,ns));}
    else if(d.mode==="resize-right"){ne=snap(ne+deltaMin);ne=Math.min(dayEndMin,Math.max(ns+MIN_DURATION,ne));}
    const updated={start:toHHMM(ns),end:toHHMM(ne)};
    liveRef.current=updated; setLive(updated);
  };
  onMouseUp.current=()=>{
    const d=dragRef.current;
    window.removeEventListener("mousemove",stableMove); window.removeEventListener("mouseup",stableUp);
    if(d&&!d.moved) onOpenModal();
    else if(d) onChange(liveRef.current.start,liveRef.current.end);
    dragRef.current=null;
  };
  const stableMove=useCallback(e=>onMouseMove.current(e),[]);
  const stableUp=useCallback(()=>onMouseUp.current(),[]);
  useEffect(()=>()=>{ window.removeEventListener("mousemove",stableMove); window.removeEventListener("mouseup",stableUp); },[stableMove,stableUp]);
  const beginDrag=(mode)=>(e)=>{
    e.stopPropagation();
    dragRef.current={mode,startX:e.clientX,origStartMin:toMin(liveRef.current.start),origEndMin:toMin(liveRef.current.end),moved:false};
    window.addEventListener("mousemove",stableMove); window.addEventListener("mouseup",stableUp);
  };
  const left=minToPct(toMin(live.start)), width=minToPct(toMin(live.end))-left;
  return (
    <div ref={trackRef} style={{ position:"relative", height:26, background:"#FAF8F2", borderRadius:6 }}>
      {hourMarks.map(h=><div key={h} style={{ position:"absolute", left:`${((h-DAY_START)/totalHoursAxis)*100}%`, top:0, bottom:0, width:1, background:h%4===0?"#EFEBDE":"transparent" }}/>)}
      <div style={{ position:"absolute", top:3, bottom:3, left:`${left}%`, width:`${width}%`, background:type.bg, borderRadius:5, display:"flex", alignItems:"center", justifyContent:"center", fontSize:10.5, fontWeight:700, color:type.text, whiteSpace:"nowrap", overflow:"hidden", cursor:"grab", userSelect:"none" }} onMouseDown={beginDrag("move")}>
        {live.start}-{live.end}
        <div className="no-print" onMouseDown={beginDrag("resize-left")} style={{ position:"absolute", left:0, top:0, bottom:0, width:9, cursor:"ew-resize", borderRadius:"5px 0 0 5px", background:"rgba(0,0,0,0.12)" }}/>
        <div className="no-print" onMouseDown={beginDrag("resize-right")} style={{ position:"absolute", right:0, top:0, bottom:0, width:9, cursor:"ew-resize", borderRadius:"0 5px 5px 0", background:"rgba(0,0,0,0.12)" }}/>
      </div>
    </div>
  );
}

// ============ MODAL ZMIANY ============
function ShiftModal({ modal, onClose, onSave, onDelete, employeeName }) {
  const s=modal.shift;
  const [type, setType] = useState(s?.type||"praca");
  const [start, setStart] = useState(s?.start||"09:00");
  const [end, setEnd] = useState(s?.end||"17:00");
  const [tasks, setTasks] = useState(s?.tasks||[]);
  const [newTask, setNewTask] = useState("");
  const [editIdx, setEditIdx] = useState(null);
  const [editVal, setEditVal] = useState("");

  const startEditTask = (i) => { setEditIdx(i); setEditVal(tasks[i]); };
  const saveEditTask = (i) => { if(editVal.trim()) setTasks(tasks.map((t,j)=>j===i?editVal.trim():t)); setEditIdx(null); setEditVal(""); };
  const cancelEdit = () => { setEditIdx(null); setEditVal(""); };

  return (
    <Overlay onClose={onClose}>
      <ModalHeader title={s?"Edytuj zmianę":"Dodaj zmianę"} subtitle={`${displayName(employeeName)} · ${modal.date}`} onClose={onClose}/>
      <div style={{ padding:18 }}>
        <Field label="Typ dnia">
          <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
            {Object.entries(TYPES).map(([k,v])=>(
              <button key={k} onClick={()=>setType(k)} style={{ padding:"7px 12px", borderRadius:8, border:type===k?`2px solid ${v.bg}`:"1px solid #E3DFD2", background:type===k?v.bg:"#fff", color:type===k?v.text:"#3a392f", fontSize:12.5, fontWeight:600, cursor:"pointer" }}>{v.label}</button>
            ))}
          </div>
        </Field>
        {type==="praca" && (
          <>
            <div style={{ display:"flex", gap:12 }}>
              <Field label="Godzina od" style={{ flex:1 }}><input type="time" min="05:00" max="22:00" value={start} onChange={e=>setStart(e.target.value)} style={inputStyle}/></Field>
              <Field label="Godzina do" style={{ flex:1 }}><input type="time" min="05:00" max="22:00" value={end} onChange={e=>setEnd(e.target.value)} style={inputStyle}/></Field>
            </div>
            <Field label="Zadania przypisane do pracownika">
              <div style={{ display:"flex", flexDirection:"column", gap:6, marginBottom:8 }}>
                {tasks.map((t,i)=>(
                  <div key={i} style={{ display:"flex", alignItems:"center", gap:6, background:"#FAF8F2", padding:"6px 8px", borderRadius:7, fontSize:12.5 }}>
                    {editIdx===i ? (
                      <>
                        <input value={editVal} onChange={e=>setEditVal(e.target.value)} onKeyDown={e=>{ if(e.key==="Enter") saveEditTask(i); if(e.key==="Escape") cancelEdit(); }} style={{ ...inputStyle, flex:1 }} autoFocus/>
                        <button onClick={()=>saveEditTask(i)} style={{ border:"none", background:"transparent", cursor:"pointer", color:"#2F6B4F" }}><Check size={14}/></button>
                        <button onClick={cancelEdit} style={{ border:"none", background:"transparent", cursor:"pointer", color:"#9C988A" }}><X size={14}/></button>
                      </>
                    ) : (
                      <>
                        <span style={{ flex:1 }}>{t}</span>
                        <button onClick={()=>startEditTask(i)} style={{ border:"none", background:"transparent", cursor:"pointer", color:"#6B6A60" }}><Pencil size={13}/></button>
                        <button onClick={()=>setTasks(tasks.filter((_,j)=>j!==i))} style={{ border:"none", background:"transparent", cursor:"pointer", color:"#C75146" }}><X size={14}/></button>
                      </>
                    )}
                  </div>
                ))}
              </div>
              <div style={{ display:"flex", gap:6 }}>
                <input value={newTask} onChange={e=>setNewTask(e.target.value)} onKeyDown={e=>{ if(e.key==="Enter"&&newTask.trim()){ setTasks([...tasks,newTask.trim()]); setNewTask(""); }}} placeholder="Np. Inwentaryzacja chłodni" style={{ ...inputStyle, flex:1 }}/>
                <button onClick={()=>{ if(newTask.trim()){ setTasks([...tasks,newTask.trim()]); setNewTask(""); }}} style={smallBtn}><Plus size={14}/></button>
              </div>
            </Field>
          </>
        )}
      </div>
      <div style={{ padding:"12px 18px", borderTop:"1px solid #EFEBDE", display:"flex", justifyContent:"space-between" }}>
        {s ? <button onClick={()=>onDelete(s.id)} style={dangerBtn}><Trash2 size={14}/> Usuń</button> : <span/>}
        <div style={{ display:"flex", gap:8 }}>
          <button onClick={onClose} style={ghostBtn}>Anuluj</button>
          <button onClick={()=>onSave({ id:s?.id, employeeId:modal.employeeId, date:modal.date, type, start:type==="praca"?start:null, end:type==="praca"?end:null, tasks:type==="praca"?tasks:[] })} style={primaryBtn}><Check size={14}/> Zapisz</button>
        </div>
      </div>
    </Overlay>
  );
}

// ============ MODAL ZADANIA DLA ZMIANY ============
function BulkTaskModal({ weekDays, onClose, onApply }) {
  const templates = [{ name:"Poranna",start:"05:00",end:"13:00" },{ name:"Dzienna",start:"09:00",end:"17:00" },{ name:"Popołudniowa",start:"13:00",end:"21:00" },{ name:"Wieczorna",start:"14:00",end:"22:00" }];
  const [date, setDate] = useState(dateStr(weekDays[0]));
  const [tplIdx, setTplIdx] = useState(0);
  const [task, setTask] = useState("");
  return (
    <Overlay onClose={onClose}>
      <ModalHeader title="Zadanie dla całej zmiany" subtitle="Przypisz zadanie wszystkim pracującym w wybranym przedziale" onClose={onClose}/>
      <div style={{ padding:18 }}>
        <Field label="Data"><select value={date} onChange={e=>setDate(e.target.value)} style={inputStyle}>{weekDays.map((d,i)=><option key={i} value={dateStr(d)}>{DAY_NAMES[i]} {d.getDate()}.{pad2(d.getMonth()+1)}</option>)}</select></Field>
        <Field label="Zmiana">
          <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
            {templates.map((t,i)=>(
              <button key={i} onClick={()=>setTplIdx(i)} style={{ padding:"7px 12px", borderRadius:8, border:tplIdx===i?"2px solid #2F6B4F":"1px solid #E3DFD2", background:tplIdx===i?"#2F6B4F":"#fff", color:tplIdx===i?"#EAF3E2":"#3a392f", fontSize:12, fontWeight:600, cursor:"pointer" }}>{t.name} {t.start}-{t.end}</button>
            ))}
          </div>
        </Field>
        <Field label="Treść zadania"><input value={task} onChange={e=>setTask(e.target.value)} placeholder="Np. Wystawić towary po promocji" style={inputStyle}/></Field>
      </div>
      <div style={{ padding:"12px 18px", borderTop:"1px solid #EFEBDE", display:"flex", justifyContent:"flex-end", gap:8 }}>
        <button onClick={onClose} style={ghostBtn}>Anuluj</button>
        <button onClick={()=>onApply({ date, template:templates[tplIdx], task })} style={primaryBtn}><Megaphone size={14}/> Przypisz do zmiany</button>
      </div>
    </Overlay>
  );
}

// ============ MODAL KOPIOWANIA DNIA ============
function CopyDayModal({ sourceDate, onClose, onApply }) {
  const [target, setTarget] = useState(dateStr(addDays(sourceDate,1)));
  const sameDay=target===dateStr(sourceDate);
  return (
    <Overlay onClose={onClose}>
      <ModalHeader title="Kopiuj grafik dnia" subtitle={`Źródło: ${dayLabelFull(sourceDate)}`} onClose={onClose}/>
      <div style={{ padding:18 }}>
        <Field label="Skopiuj na dzień"><input type="date" value={target} onChange={e=>setTarget(e.target.value)} style={inputStyle}/></Field>
        <div style={{ fontSize:11.5, color:"#9C988A" }}>Istniejący grafik w dniu docelowym zostanie zastąpiony.</div>
      </div>
      <div style={{ padding:"12px 18px", borderTop:"1px solid #EFEBDE", display:"flex", justifyContent:"flex-end", gap:8 }}>
        <button onClick={onClose} style={ghostBtn}>Anuluj</button>
        <button disabled={sameDay} onClick={()=>{ onApply(dateStr(sourceDate),target); onClose(); }} style={{ ...primaryBtn, opacity:sameDay?0.5:1, cursor:sameDay?"default":"pointer" }}><Copy size={14}/> Kopiuj</button>
      </div>
    </Overlay>
  );
}

// ============ MODAL KOPIOWANIA TYGODNIA ============
function CopyWeekModal({ sourceWeekStart, onClose, onApply }) {
  const [targetWeekDate, setTargetWeekDate] = useState(dateStr(addDays(sourceWeekStart,7)));
  const targetStart=startOfWeek(new Date(targetWeekDate));
  const sameWeek=dateStr(targetStart)===dateStr(sourceWeekStart);
  return (
    <Overlay onClose={onClose}>
      <ModalHeader title="Kopiuj grafik tygodnia" subtitle={`Źródło: ${dateStr(sourceWeekStart)} - ${dateStr(addDays(sourceWeekStart,6))}`} onClose={onClose}/>
      <div style={{ padding:18 }}>
        <Field label="Wybierz dowolny dzień docelowego tygodnia"><input type="date" value={targetWeekDate} onChange={e=>setTargetWeekDate(e.target.value)} style={inputStyle}/></Field>
        <div style={{ fontSize:12.5, color:"#3a392f", background:"#FAF8F2", borderRadius:8, padding:"8px 10px", marginBottom:10 }}>
          Docelowy tydzień: <b>{dateStr(targetStart)} - {dateStr(addDays(targetStart,6))}</b>
        </div>
        <div style={{ fontSize:11.5, color:"#9C988A" }}>Istniejący grafik w tygodniu docelowym zostanie zastąpiony.</div>
      </div>
      <div style={{ padding:"12px 18px", borderTop:"1px solid #EFEBDE", display:"flex", justifyContent:"flex-end", gap:8 }}>
        <button onClick={onClose} style={ghostBtn}>Anuluj</button>
        <button disabled={sameWeek} onClick={()=>{ onApply(sourceWeekStart,targetStart); onClose(); }} style={{ ...primaryBtn, opacity:sameWeek?0.5:1, cursor:sameWeek?"default":"pointer" }}><Copy size={14}/> Kopiuj tydzień</button>
      </div>
    </Overlay>
  );
}

// ============ WIDOK PRACOWNICY ============
function EmployeesView({ employees, setEmpModal, shifts, isAdmin, positions }) {
  const today = dateStr(new Date());
  const [analysisEmp, setAnalysisEmp] = useState(null);
  const [analysisYear, setAnalysisYear] = useState(new Date().getFullYear());
  const active = employees.filter(e => !e.contractEndDate || e.contractEndDate >= today);
  const terminated = employees.filter(e => e.contractEndDate && e.contractEndDate < today);
  const yearOptions = [];
  for (let y = new Date().getFullYear(); y >= new Date().getFullYear()-3; y--) yearOptions.push(y);
  return (
    <div>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:20, flexWrap:"wrap", gap:10 }}>
        <div>
          <div style={{ fontWeight:700, fontSize:16, color:"#1A2530" }}>Pracownicy</div>
          <div style={{ fontSize:12.5, color:"#8890A0", marginTop:2 }}>{active.length} aktywnych{terminated.length>0?` · ${terminated.length} po umowie`:""}</div>
        </div>
        <div style={{ display:"flex", gap:8, alignItems:"center", flexWrap:"wrap" }}>
          <div style={{ display:"flex", alignItems:"center", gap:6, fontSize:13, color:"#5A6070", background:"#fff", border:"1px solid #E4E7EE", borderRadius:8, padding:"6px 10px" }}>
            <Calendar size={13} color="#8890A0"/>
            <span style={{color:"#8890A0"}}>Rok:</span>
            <select value={analysisYear} onChange={e=>setAnalysisYear(Number(e.target.value))} style={{ border:"none", background:"transparent", fontSize:13, fontWeight:600, color:"#1A2530", cursor:"pointer", outline:"none", padding:0 }}>
              {yearOptions.map(y=><option key={y} value={y}>{y}</option>)}
            </select>
          </div>
          <button onClick={()=>setEmpModal({ name:"", rate:28, position:positions[0]||"", etat:"1", paymentMethod:"przelew", defaultTransferAmount:0, contractEndDate:null, vacationLimit:20 })} style={primaryBtn}><Plus size={15}/> Dodaj pracownika</button>
        </div>
      </div>
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(290px, 1fr))", gap:14 }}>
        {active.map(e=><EmployeeCard2 key={e.id} e={e} isAdmin={isAdmin} setEmpModal={setEmpModal} shifts={shifts} terminated={false} onAnalysis={()=>setAnalysisEmp(e)} analysisYear={analysisYear}/>)}
      </div>
      {terminated.length>0 && (
        <>
          <div style={{ display:"flex", alignItems:"center", gap:10, margin:"24px 0 14px", color:"#8890A0" }}>
            <div style={{ flex:1, height:1, background:"#E4E7EE" }}/>
            <span style={{ fontWeight:600, fontSize:12, whiteSpace:"nowrap" }}>Zakończone umowy ({terminated.length})</span>
            <div style={{ flex:1, height:1, background:"#E4E7EE" }}/>
          </div>
          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(290px, 1fr))", gap:14 }}>
            {terminated.map(e=><EmployeeCard2 key={e.id} e={e} isAdmin={isAdmin} setEmpModal={setEmpModal} shifts={shifts} terminated={true} onAnalysis={()=>setAnalysisEmp(e)} analysisYear={analysisYear}/>)}
          </div>
        </>
      )}
      {analysisEmp && (
        <EmployeeAnalysisModal
          e={analysisEmp}
          shifts={shifts}
          year={analysisYear}
          isAdmin={isAdmin}
          onClose={()=>setAnalysisEmp(null)}
        />
      )}
    </div>
  );
}

function EmployeeCard2({ e, isAdmin, setEmpModal, shifts, terminated, onAnalysis, analysisYear }) {
  const today = dateStr(new Date());
  const upcoming = shifts.filter(s=>s.employeeId===e.id&&s.type==="praca"&&s.date>=today).length;
  // Urlop w roku analysisYear
  const vacUsedLocal = shifts.filter(s=>s.employeeId===e.id&&s.type==="urlop"&&s.date.startsWith(String(analysisYear))).length;
  // Urlop przeniesiony z poprzedniego sklepu (dla danego roku)
  const vacTransferred = (e.vacationTransferUsed||0) && (e.vacationTransferYear||0)===analysisYear ? (e.vacationTransferUsed||0) : 0;
  const vacUsed = vacUsedLocal + vacTransferred;
  const vacBaseLimit = e.vacationLimit||20;
  // Proporcjonalny urlop wg art. 153/155 KP + zaległy z poprzedniego roku
  const vacLimit = useMemo(()=>{
    let base = vacBaseLimit;
    if (e.contractStartDate) {
      const startYear = Number(e.contractStartDate.slice(0,4));
      if (startYear === analysisYear) {
        const startMonth = Number(e.contractStartDate.slice(5,7));
        base = Math.ceil(((13-startMonth)/12)*vacBaseLimit);
      }
    }
    const carryOver = (e.vacationCarryOver||0) && (e.vacationTransferYear||0)===analysisYear ? (e.vacationCarryOver||0) : 0;
    return base + carryOver;
  },[e.contractStartDate, e.vacationCarryOver, e.vacationTransferYear, analysisYear, vacBaseLimit]);
  const vacLeft = Math.max(0, vacLimit - vacUsed);
  const vacPct = Math.min(100, Math.round((vacUsed/vacLimit)*100));
  const isProrated2 = !!(e.contractStartDate && Number(e.contractStartDate.slice(0,4))===analysisYear);
  // Avatar color from name
  const AVATAR_COLORS=["#1F6B45","#2563A8","#7C3ABB","#B45309","#0F7D7D","#B02D64","#2D5F2D","#6B3A1F"];
  const avatarColor=AVATAR_COLORS[e.name.charCodeAt(0)%AVATAR_COLORS.length];
  const initials=(()=>{ const p=e.name.trim().split(/\s+/); return p.length<2 ? (p[0][0]||"").toUpperCase() : ((p[p.length-1][0]||"")+(p[0][0]||"")).toUpperCase(); })();
  return (
    <div style={{ background:"#fff", border:"1px solid #E4E7EE", borderRadius:14, padding:16, opacity:terminated?0.65:1, boxShadow:"0 1px 4px rgba(0,0,0,0.06)", transition:"box-shadow 0.2s" }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", gap:10 }}>
        {/* Avatar + info */}
        <div style={{ display:"flex", gap:12, flex:1, minWidth:0 }}>
          <div style={{ width:44, height:44, borderRadius:12, background:terminated?"#9399A6":avatarColor, display:"flex", alignItems:"center", justifyContent:"center", color:"#fff", fontWeight:700, fontSize:16, flexShrink:0, letterSpacing:"0.02em" }}>{initials}</div>
          <div style={{ flex:1, minWidth:0 }}>
            <div style={{ fontWeight:700, fontSize:14.5, color:"#1A2530" }}>{displayName(e.name)}</div>
            {e.position && <div style={{ fontSize:12, color:"#1F6B45", fontWeight:600, marginTop:2 }}>{e.position}{e.etat&&e.etat!=="1"?<span style={{color:"#8890A0",fontWeight:400}}> · {e.etat} etatu</span>:""}</div>}
            {terminated
              ? <div style={{ display:"inline-flex",alignItems:"center",gap:4,marginTop:4,background:"#FEF2F2",color:"#C23B3B",borderRadius:6,padding:"2px 7px",fontSize:11,fontWeight:600 }}><span style={{width:5,height:5,borderRadius:"50%",background:"#C23B3B",display:"inline-block"}}/>Umowa zakończona: {e.contractEndDate}</div>
              : isAdmin
                ? <div style={{ fontSize:12, color:"#8890A0", marginTop:3 }}>{e.rate.toFixed(2)} zł/godz · {PAYMENT_METHODS.find(m=>m.id===(e.paymentMethod||"przelew"))?.label}</div>
                : <div style={{ fontSize:12, color:"#B0B6C2", marginTop:3, display:"flex", alignItems:"center", gap:4 }}><EyeOff size={11}/> Stawka ukryta</div>
            }
          </div>
        </div>
        <div style={{ display:"flex", gap:4, flexShrink:0 }}>
          <button title="Analiza pracownika" onClick={onAnalysis} style={iconBtn}><BarChart2 size={14}/></button>
          <button title="Edytuj" onClick={()=>setEmpModal({...e})} style={iconBtn}><Pencil size={14}/></button>
        </div>
      </div>
      {!terminated && (
        <>
          <div style={{ marginTop:13, display:"flex", alignItems:"center", gap:6, fontSize:12, color:"#8890A0" }}>
            <Calendar size={12}/> <span>Zaplanowanych zmian (od dziś): <b style={{color:"#1A2530"}}>{upcoming}</b></span>
          </div>
          <div style={{ marginTop:10, padding:"10px 12px", background:"#F7F9FC", borderRadius:8, border:"1px solid #EEF0F5" }}>
            <div style={{ display:"flex", justifyContent:"space-between", fontSize:12, color:"#5A6070", marginBottom:6 }}>
              <span style={{fontWeight:600}}>Urlop {analysisYear}</span>
              <span style={{ color: vacLeft===0?"#C23B3B":vacLeft<=3?"#D97706":"#1F6B45", fontWeight:600 }}>{vacLeft} dni pozostało</span>
            </div>
            <div style={{ height:6, borderRadius:6, background:"#E4E7EE", overflow:"hidden", marginBottom:5 }}>
              <div style={{ height:"100%", borderRadius:6, width:`${vacPct}%`, background: vacPct>=100?"#C23B3B":vacPct>=75?"#D97706":"#1F6B45", transition:"width 0.4s" }}/>
            </div>
            <div style={{ display:"flex", justifyContent:"space-between", fontSize:11, color:"#9399A6" }}>
              <span>{vacUsed} wykorzystanych</span>
              <span>{vacLimit} dni ogółem{isProrated2&&<span style={{color:"#B0B6C2"}}> (proporcjonalnie)</span>}</span>
            </div>
          </div>
        </>
      )}
    </div>
  );
}

// ============ MODAL ANALIZY PRACOWNIKA ============
function EmployeeAnalysisModal({ e, shifts, year, isAdmin, onClose }) {
  const [activeTab, setActiveTab] = useState("urlop");
  const MONTHS = ["Styczen","Luty","Marzec","Kwiecien","Maj","Czerwiec","Lipiec","Sierpien","Wrzesien","Pazdziernik","Listopad","Grudzien"];
  const MONTHS_SHORT = ["Sty","Lut","Mar","Kwi","Maj","Cze","Lip","Sie","Wrz","Paz","Lis","Gru"];

  const empShifts = useMemo(()=>
    shifts.filter(s=>s.employeeId===e.id && s.date.startsWith(String(year)))
  ,[shifts, e.id, year]);

  // Agreguj per miesiac
  const monthly = useMemo(()=>{
    const res = Array.from({length:12},(_,i)=>({ m:i, work:0, workH:0, urlop:0, urlopB:0, l4:0, l4b:0, opC:0, opZ:0, wolne:0 }));
    empShifts.forEach(s=>{
      const m = new Date(s.date).getMonth();
      if (s.type==="praca"&&s.start&&s.end) { res[m].work++; res[m].workH+=Math.max(0,(toMin(s.end)-toMin(s.start))/60); }
      else if (s.type==="urlop") res[m].urlop++;
      else if (s.type==="urlop_bezplatny") res[m].urlopB++;
      else if (s.type==="l4") res[m].l4++;
      else if (s.type==="l4b") res[m].l4b++;
      else if (s.type==="opieka_chore") res[m].opC++;
      else if (s.type==="opieka_zdrowe") res[m].opZ++;
      else if (s.type==="wolne") res[m].wolne++;
    });
    return res;
  },[empShifts]);

  const totals = useMemo(()=>({
    work:   monthly.reduce((a,m)=>a+m.work,0),
    workH:  monthly.reduce((a,m)=>a+m.workH,0),
    urlop:  monthly.reduce((a,m)=>a+m.urlop,0),
    urlopB: monthly.reduce((a,m)=>a+m.urlopB,0),
    l4:     monthly.reduce((a,m)=>a+m.l4,0),
    l4b:    monthly.reduce((a,m)=>a+m.l4b,0),
    opC:    monthly.reduce((a,m)=>a+m.opC,0),
    opZ:    monthly.reduce((a,m)=>a+m.opZ,0),
    wolne:  monthly.reduce((a,m)=>a+m.wolne,0),
  }),[monthly]);

  const vacLimit = e.vacationLimit||20;
  // Proporcjonalny urlop wg art. 153/155 KP
  // Jesli pracownik zatrudniony w roku analizy: limit = ceil(miesiacePracy/12 * rocznyLimit)
  // miesiacePracy = liczba miesiecy od zatrudnienia do konca roku (wlacznie z miesiacem zatrudnienia)
  const vacEffectiveLimit = useMemo(()=>{
    // Podstawa: proporcja wg art. 153/155 KP jesli nowy pracownik w roku analizy
    let base = vacLimit;
    if (e.contractStartDate) {
      const startYear = Number(e.contractStartDate.slice(0,4));
      if (startYear === year) {
        const startMonth = Number(e.contractStartDate.slice(5,7));
        const monthsWorked = 13 - startMonth;
        base = Math.ceil((monthsWorked / 12) * vacLimit);
      }
    }
    // Urlop zaległy z poprzedniego roku (na plus)
    const carryOver = (e.vacationCarryOver||0) && (e.vacationTransferYear||0)===year ? (e.vacationCarryOver||0) : 0;
    return base + carryOver;
  },[e.contractStartDate, e.vacationCarryOver, e.vacationTransferYear, year, vacLimit]);

  const vacUsedLocal = totals.urlop;
  // Urlop przeniesiony z poprzedniego sklepu (dla roku analizy)
  const vacTransferred = (e.vacationTransferUsed||0) && (e.vacationTransferYear||0)===year ? (e.vacationTransferUsed||0) : 0;
  const vacCarryOver = (e.vacationCarryOver||0) && (e.vacationTransferYear||0)===year ? (e.vacationCarryOver||0) : 0;
  const vacUsed = vacUsedLocal + vacTransferred;
  const vacFuture = empShifts.filter(s=>s.type==="urlop"&&s.date>dateStr(new Date())).length;
  const vacLeft = Math.max(0, vacEffectiveLimit - vacUsed);
  const vacPct = Math.min(100,Math.round((vacUsed/vacEffectiveLimit)*100));
  const isProrated = !!(e.contractStartDate && Number(e.contractStartDate.slice(0,4))===year);

  // Lista nieobecnosci chronologicznie
  const absenceList = useMemo(()=>
    empShifts
      .filter(s=>s.type!=="praca")
      .sort((a,b)=>a.date.localeCompare(b.date))
      .map(s=>({ date:s.date, type:s.type, label:TYPES[s.type]?.label||s.type }))
  ,[empShifts]);

  const handleExportXlsx = () => {
    const wb = XLSX.utils.book_new();
    // Arkusz 1: podsumowanie miesięczne
    const mRows = monthly.map((m,i)=>({
      "Miesiac": MONTHS[i],
      "Dni pracy": m.work,
      "Godziny pracy": Number(m.workH.toFixed(1)),
      "Urlop": m.urlop,
      "Urlop bezpl.": m.urlopB,
      "L4": m.l4,
      "L4 kod B": m.l4b,
      "Opieka chore": m.opC,
      "Opieka zdrowe": m.opZ,
      "Dni wolne": m.wolne,
    }));
    mRows.push({ "Miesiac":"SUMA", "Dni pracy":totals.work, "Godziny pracy":Number(totals.workH.toFixed(1)), "Urlop":totals.urlop, "Urlop bezpl.":totals.urlopB, "L4":totals.l4, "L4 kod B":totals.l4b, "Opieka chore":totals.opC, "Opieka zdrowe":totals.opZ, "Dni wolne":totals.wolne });
    const ws1 = XLSX.utils.json_to_sheet(mRows);
    ws1["!cols"]=[14,10,14,8,12,8,10,13,14,10].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb, ws1, "Podsumowanie miesięczne");
    // Arkusz 2: nieobecności
    const abRows = absenceList.map(a=>({ "Data":a.date, "Typ":a.label }));
    const ws2 = XLSX.utils.json_to_sheet(abRows.length?abRows:[{"Data":"--","Typ":"Brak nieobecnosci"}]);
    ws2["!cols"]=[14,28].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb, ws2, "Nieobecnosci");
    // Arkusz 3: bilans urlopu
    const vacRows=[
      {"Parametr":"Limit urlopu (dni)","Wartosc":vacEffectiveLimit},
      {"Parametr":isProrated?"Limit proporcjonalny (art. 153 KP)":"Limit pelny","Wartosc":isProrated?"TAK":"NIE"},
      ...(vacCarryOver>0?[{"Parametr":"  w tym: urlop zalegly z poprzedniego roku","Wartosc":"+"+vacCarryOver}]:[]),
      {"Parametr":"Wykorzystano lacznie (dni)","Wartosc":vacUsed},
      ...(vacTransferred>0?[{"Parametr":"  w tym: zuzyte w poprzednim sklepie","Wartosc":vacTransferred},{"Parametr":"  w tym: zuzyte w tym sklepie","Wartosc":vacUsedLocal}]:[]),
      {"Parametr":"Pozostalo (dni)","Wartosc":vacLeft},
      {"Parametr":"Zaplanowano w przyszlosci","Wartosc":vacFuture},
    ];
    const ws3 = XLSX.utils.json_to_sheet(vacRows);
    ws3["!cols"]=[30,12].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb, ws3, "Bilans urlopu");
    XLSX.writeFile(wb, `Analiza_${e.name.replace(/\s+/g,"_")}_${year}.xlsx`);
  };

  const printRef = useRef(null);
  const handlePrint = () => {
    const fillColor = vacPct>=100?"#C75146":vacPct>=75?"#E07A3E":"#2F6B4F";
    const bodyHtml = `
      <h1 style="font-size:15px;color:#1F4D38;margin:0 0 4px;">Analiza czasu pracy: ${displayName(e.name)} &middot; ${year}</h1>
      <p style="font-size:10px;color:#555;margin:0 0 4px;">${e.position||""}${e.etat&&e.etat!=="1"?" &middot; etat "+e.etat:""} &middot; Limit urlopu: ${vacEffectiveLimit} dni${isProrated?" (proporcjonalnie art. 153 KP)":""}</p>
      ${vacCarryOver>0?`<p style="font-size:10px;color:#2F6B4F;margin:0 0 2px;">&#10133; Urlop zaległy z poprzedniego roku: +${vacCarryOver} dni</p>`:""}
      ${vacTransferred>0?`<p style="font-size:10px;color:#7A6B30;margin:0 0 8px;">&#10134; Zuzyte w poprzednim sklepie: ${vacTransferred} dni &middot; W tym sklepie: ${vacUsedLocal} dni</p>`:""}
      <h2 style="font-size:12px;color:#1F4D38;border-bottom:1px solid #ccc;padding-bottom:3px;margin:10px 0 6px;">Bilans urlopu ${year}</h2>
      <div style="height:9px;border-radius:4px;background:#E3DFD2;margin:4px 0 6px;overflow:hidden;"><div style="height:100%;border-radius:4px;background:${fillColor};width:${vacPct}%;"></div></div>
      <p style="margin:0 0 10px;font-size:11px;">Limit: <b>${vacEffectiveLimit} dni</b>&nbsp;&nbsp;|&nbsp;&nbsp;Wykorzystano: <b>${vacUsed} dni</b> (${vacPct}%)&nbsp;&nbsp;|&nbsp;&nbsp;Pozostalo: <b style="color:${fillColor}">${vacLeft} dni</b>${vacFuture>0?"&nbsp;&nbsp;|&nbsp;&nbsp;Zaplanowano: <b>"+vacFuture+" dni</b>":""}</p>
      <h2 style="font-size:12px;color:#1F4D38;border-bottom:1px solid #ccc;padding-bottom:3px;margin:10px 0 6px;">Podsumowanie miesięczne</h2>
      <table style="width:100%;border-collapse:collapse;font-size:10px;margin-bottom:12px;">
        <thead><tr style="background:#1F4D38;color:#fff;">
          <th style="padding:4px 5px;text-align:left;">Miesiac</th>
          <th style="padding:4px 3px;text-align:center;">Praca</th><th style="padding:4px 3px;text-align:center;">Godz.</th>
          <th style="padding:4px 3px;text-align:center;background:#3B6FA0;">Urlop</th><th style="padding:4px 3px;text-align:center;">U.bezpl.</th>
          <th style="padding:4px 3px;text-align:center;background:#8B2E2E;">L4</th><th style="padding:4px 3px;text-align:center;">L4 B</th>
          <th style="padding:4px 3px;text-align:center;">Op.ch.</th><th style="padding:4px 3px;text-align:center;">Op.zd.</th><th style="padding:4px 3px;text-align:center;">Wolne</th>
        </tr></thead>
        <tbody>
          ${monthly.map((m,i)=>`<tr style="background:${i%2===0?"#FAFAF7":"#fff"};">
            <td style="padding:3px 5px;font-weight:600;color:#1F4D38;">${MONTHS_SHORT[i]}</td>
            <td style="padding:3px;text-align:center;">${m.work||""}</td><td style="padding:3px;text-align:center;color:#555;">${m.workH>0?m.workH.toFixed(1):""}</td>
            <td style="padding:3px;text-align:center;font-weight:${m.urlop>0?700:400};color:#3B6FA0;">${m.urlop||""}</td><td style="padding:3px;text-align:center;">${m.urlopB||""}</td>
            <td style="padding:3px;text-align:center;font-weight:${m.l4>0?700:400};color:#C75146;">${m.l4||""}</td><td style="padding:3px;text-align:center;">${m.l4b||""}</td>
            <td style="padding:3px;text-align:center;">${m.opC||""}</td><td style="padding:3px;text-align:center;">${m.opZ||""}</td><td style="padding:3px;text-align:center;color:#9C988A;">${m.wolne||""}</td>
          </tr>`).join("")}
          <tr style="background:#EAF3E2;font-weight:700;border-top:2px solid #1F4D38;">
            <td style="padding:4px 5px;color:#1F4D38;">SUMA</td>
            <td style="padding:4px 3px;text-align:center;">${totals.work}</td><td style="padding:4px 3px;text-align:center;color:#555;">${totals.workH.toFixed(1)}</td>
            <td style="padding:4px 3px;text-align:center;color:#3B6FA0;">${totals.urlop||""}</td><td style="padding:4px 3px;text-align:center;">${totals.urlopB||""}</td>
            <td style="padding:4px 3px;text-align:center;color:#C75146;">${totals.l4||""}</td><td style="padding:4px 3px;text-align:center;">${totals.l4b||""}</td>
            <td style="padding:4px 3px;text-align:center;">${totals.opC||""}</td><td style="padding:4px 3px;text-align:center;">${totals.opZ||""}</td><td style="padding:4px 3px;text-align:center;color:#9C988A;">${totals.wolne||""}</td>
          </tr>
        </tbody>
      </table>
      ${absenceList.length>0?`
        <h2 style="font-size:12px;color:#1F4D38;border-bottom:1px solid #ccc;padding-bottom:3px;margin:10px 0 6px;">Nieobecnosci chronologicznie (${absenceList.length})</h2>
        <table style="width:50%;border-collapse:collapse;font-size:10px;">
          <thead><tr style="background:#1F4D38;color:#fff;"><th style="padding:4px 6px;text-align:left;">Data</th><th style="padding:4px 6px;text-align:left;">Typ</th></tr></thead>
          <tbody>${absenceList.map((a,i)=>`<tr style="background:${i%2===0?"#FAFAF7":"#fff"};"><td style="padding:3px 6px;">${a.date}</td><td style="padding:3px 6px;">${a.label}</td></tr>`).join("")}</tbody>
        </table>`:""}
      <p style="margin-top:20px;font-size:9px;color:#aaa;">Wygenerowano: ${new Date().toLocaleDateString("pl-PL")} &middot; Grafik Sklepu</p>
      <p style="margin-top:6px;font-size:9px;color:#bbb;border:1px solid #eee;border-radius:4px;padding:5px;">
        Aby wydrukować lub zapisać jako PDF: Plik &rarr; Drukuj (Ctrl+P) i wybierz "Zapisz jako PDF"
      </p>
    `;
    const fullHtml = `<!DOCTYPE html><html lang="pl"><head><meta charset="utf-8"><title>Analiza: ${displayName(e.name)} ${year}</title>
      <style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:Arial,sans-serif;font-size:11px;color:#222;padding:14mm 16mm}h1,h2{page-break-after:avoid}table{page-break-inside:avoid}@page{size:A4 portrait;margin:0}</style>
    </head><body>${bodyHtml}</body></html>`;
    // Pobierz jako plik HTML — jedyna metoda działająca w środowisku sandbox (Electron/iframe)
    try {
      const blob = new Blob([fullHtml], { type: "text/html;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `analiza_${e.name.replace(/\s+/g,"_")}_${year}.html`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 2000);
    } catch(err) {
      // Fallback: data URI
      const dataUri = "data:text/html;charset=utf-8," + encodeURIComponent(fullHtml);
      const a = document.createElement("a");
      a.href = dataUri;
      a.download = `analiza_${e.name.replace(/\s+/g,"_")}_${year}.html`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    }
  };

  const tabStyle = (t) => ({ padding:"7px 14px", fontSize:12.5, fontWeight:600, cursor:"pointer", border:"none", borderBottom: activeTab===t?"2px solid #1F4D38":"2px solid transparent", background:"transparent", color:activeTab===t?"#1F4D38":"#9C988A" });

  return (
    <Overlay onClose={onClose}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", padding:"14px 18px", borderBottom:"1px solid #EFEBDE" }}>
        <div>
          <div style={{ fontWeight:700, fontSize:15, color:"#1F4D38", display:"flex", alignItems:"center", gap:8 }}><BarChart2 size={16}/> Analiza pracownika</div>
          <div style={{ fontSize:12, color:"#6B6A60", marginTop:2 }}>{displayName(e.name)} · {e.position||""} · rok {year}</div>
        </div>
        <div style={{ display:"flex", gap:6 }}>
          <button title="Eksport Excel" onClick={handleExportXlsx} style={ghostBtn}><Download size={14}/> Excel</button>
          <button title="Pobierz raport jako plik HTML (otwórz w przeglądarce i użyj Ctrl+P → Zapisz jako PDF)" onClick={handlePrint} style={ghostBtn}><Download size={14}/> Pobierz raport</button>
          <button onClick={onClose} style={iconBtn}><X size={16}/></button>
        </div>
      </div>

      {/* Bilans urlopu - zawsze widoczny na gorze */}
      <div style={{ padding:"12px 18px", background:"#F7FAF8", borderBottom:"1px solid #EFEBDE", display:"flex", gap:20, flexWrap:"wrap", alignItems:"center" }}>
        <div style={{ flex:1, minWidth:200 }}>
          <div style={{ fontSize:11.5, color:"#6B6A60", marginBottom:4, fontWeight:600 }}>BILANS URLOPU {year}</div>
          <div style={{ display:"flex", justifyContent:"space-between", fontSize:12 }}>
            <span>Limit: <b>{vacEffectiveLimit} dni</b>{isProrated&&<span style={{fontSize:10,color:"#E07A3E",marginLeft:3}}>(proporcj. art.153 KP)</span>}{vacCarryOver>0&&<span style={{fontSize:10,color:"#2F6B4F",marginLeft:3}}>(+{vacCarryOver} zaległy)</span>}</span>
            <span>Wykorzystano: <b style={{color:vacPct>=100?"#C75146":"#1F4D38"}}>{vacUsed} dni</b></span>
            <span style={{color:vacLeft===0?"#C75146":vacLeft<=3?"#E07A3E":"#2F6B4F", fontWeight:700}}>Pozostalo: {vacLeft} dni</span>
          </div>
          <div style={{ height:8, borderRadius:4, background:"#E3DFD2", margin:"5px 0", overflow:"hidden" }}>
            <div style={{ height:"100%", borderRadius:4, width:`${vacPct}%`, background:vacPct>=100?"#C75146":vacPct>=75?"#E07A3E":"#2F6B4F" }}/>
          </div>
          {vacTransferred>0 && <div style={{ fontSize:11, color:"#7A6B30", background:"#FFF9EC", border:"1px solid #F0D88A", borderRadius:5, padding:"3px 7px", marginTop:4 }}>🔄 Przeniesione z poprzedniego sklepu: <b>{vacTransferred} dni</b> (wliczone w "Wykorzystano")</div>}
          {vacFuture>0 && <div style={{ fontSize:11, color:"#E07A3E" }}>Zaplanowano w przyszlosci: {vacFuture} dni (wliczone w powyzsze)</div>}
        </div>
        <div style={{ display:"flex", gap:12, flexWrap:"wrap" }}>
          {[
            {label:"Dni pracy",val:totals.work,sub:`${totals.workH.toFixed(1)} h`,color:"#2F6B4F"},
            {label:"L4",val:totals.l4+totals.l4b,color:"#C75146"},
            {label:"Opieka",val:totals.opC+totals.opZ,color:"#A6446B"},
            {label:"U.bezpl.",val:totals.urlopB,color:"#6B6A60"},
          ].map(item=>(
            <div key={item.label} style={{ textAlign:"center", minWidth:56 }}>
              <div style={{ fontWeight:700, fontSize:18, color:item.color }}>{item.val}</div>
              <div style={{ fontSize:10.5, color:"#9C988A" }}>{item.label}</div>
              {item.sub && <div style={{ fontSize:10, color:"#B0ADA3" }}>{item.sub}</div>}
            </div>
          ))}
        </div>
      </div>

      {/* Zakladki */}
      <div style={{ display:"flex", borderBottom:"1px solid #EFEBDE", padding:"0 18px" }}>
        <button style={tabStyle("urlop")} onClick={()=>setActiveTab("urlop")}>Bilans urlopu</button>
        <button style={tabStyle("miesiace")} onClick={()=>setActiveTab("miesiace")}>Tabela miesięczna</button>
        <button style={tabStyle("lista")} onClick={()=>setActiveTab("lista")}>Lista nieobecności</button>
      </div>

      <div style={{ padding:"14px 18px", maxHeight:420, overflowY:"auto" }}>

        {/* TAB: BILANS URLOPU */}
        {activeTab==="urlop" && (
          <div>
            <div style={{ fontSize:13, fontWeight:600, color:"#1F4D38", marginBottom:12 }}>Urlop wypoczynkowy {year}</div>
            <div style={{ display:"grid", gridTemplateColumns:"repeat(3,1fr)", gap:10, marginBottom:16 }}>
              {[
                {label:isProrated?`Limit proporcj. (art.153 KP)`:"Limit roczny",val:`${vacEffectiveLimit} dni`,color:"#1F4D38",bg:"#EAF3E2"},
                {label:"Wykorzystano",val:`${vacUsed} dni`,color:vacPct>=100?"#C75146":"#1F4D38",bg:vacPct>=100?"#FBE8E8":"#F7FAF8"},
                {label:"Pozostalo",val:`${vacLeft} dni`,color:vacLeft===0?"#C75146":vacLeft<=3?"#E07A3E":"#2F6B4F",bg:"#FAF8F2"},
              ].map(c=>(
                <div key={c.label} style={{ background:c.bg, borderRadius:10, padding:"12px 14px", textAlign:"center" }}>
                  <div style={{ fontWeight:700, fontSize:22, color:c.color }}>{c.val}</div>
                  <div style={{ fontSize:11.5, color:"#9C988A", marginTop:2 }}>{c.label}</div>
                </div>
              ))}
            </div>
            {vacCarryOver>0 && (
              <div style={{ background:"#F0FAF4", border:"1px solid #B6DFC8", borderRadius:8, padding:"7px 12px", fontSize:11.5, color:"#1F6B3A", marginBottom:8, display:"flex", alignItems:"center", gap:6 }}>
                <span>➕</span>
                <span><b>{vacCarryOver} dni</b> urlopu zaległego z poprzedniego roku — dodane do limitu {year}.</span>
              </div>
            )}
            {vacTransferred>0 && (
              <div style={{ background:"#FFF9EC", border:"1px solid #F0D88A", borderRadius:8, padding:"7px 12px", fontSize:11.5, color:"#7A6B30", marginBottom:8, display:"flex", alignItems:"center", gap:6 }}>
                <span>➖</span>
                <span><b>{vacTransferred} dni</b> zużyte w poprzednim sklepie — wliczone w "Wykorzystano". W tym sklepie: <b>{vacUsedLocal} dni</b>.</span>
              </div>
            )}
            {/* Wykres korkowy miesieczny - urlop */}
            <div style={{ fontSize:12, fontWeight:600, color:"#6B6A60", marginBottom:8 }}>Urlop wg miesięcy (w tym sklepie):</div>
            <div style={{ display:"flex", gap:4, alignItems:"flex-end", height:80 }}>
              {monthly.map((m,i)=>{
                const h = m.urlop>0 ? Math.max(12, (m.urlop/Math.max(...monthly.map(x=>x.urlop),1))*70) : 0;
                return (
                  <div key={i} style={{ flex:1, display:"flex", flexDirection:"column", alignItems:"center", gap:2 }}>
                    {m.urlop>0 && <div style={{ fontSize:9, color:"#3B6FA0", fontWeight:700 }}>{m.urlop}</div>}
                    <div style={{ width:"100%", height:`${h||4}px`, background:m.urlop>0?"#3B6FA0":"#F0EDE6", borderRadius:"3px 3px 0 0", transition:"height 0.3s" }}/>
                    <div style={{ fontSize:9, color:"#9C988A" }}>{MONTHS_SHORT[i]}</div>
                  </div>
                );
              })}
            </div>
            {vacFuture>0 && (
              <div style={{ marginTop:14, background:"#FFF8E7", border:"1px solid #F5D87A", borderRadius:8, padding:"8px 12px", fontSize:12, color:"#7B5800" }}>
                <b>Zaplanowane urlopy (przyszlosc):</b>
                <ul style={{ margin:"4px 0 0 16px", padding:0 }}>
                  {empShifts.filter(s=>s.type==="urlop"&&s.date>dateStr(new Date())).sort((a,b)=>a.date.localeCompare(b.date)).map(s=>(
                    <li key={s.date} style={{ fontSize:11.5 }}>{s.date}</li>
                  ))}
                </ul>
              </div>
            )}
          </div>
        )}

        {/* TAB: TABELA MIESIĘCZNA */}
        {activeTab==="miesiace" && (
          <div>
            <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
              <thead>
                <tr style={{ background:"#1F4D38", color:"#EAF3E2" }}>
                  <th style={{ padding:"6px 8px", textAlign:"left", fontWeight:600 }}>Miesiac</th>
                  <th style={{ padding:"6px 4px", textAlign:"center" }}>Praca</th>
                  <th style={{ padding:"6px 4px", textAlign:"center" }}>Godz.</th>
                  <th style={{ padding:"6px 4px", textAlign:"center", background:"#3B6FA0" }}>Urlop</th>
                  <th style={{ padding:"6px 4px", textAlign:"center" }}>U.bezpl.</th>
                  <th style={{ padding:"6px 4px", textAlign:"center", background:"#A03B3B" }}>L4</th>
                  <th style={{ padding:"6px 4px", textAlign:"center" }}>L4 B</th>
                  <th style={{ padding:"6px 4px", textAlign:"center" }}>Op.chore</th>
                  <th style={{ padding:"6px 4px", textAlign:"center" }}>Op.zdr.</th>
                  <th style={{ padding:"6px 4px", textAlign:"center" }}>Wolne</th>
                </tr>
              </thead>
              <tbody>
                {monthly.map((m,i)=>{
                  const hasData = m.work>0||m.urlop>0||m.urlopB>0||m.l4>0||m.l4b>0||m.opC>0||m.opZ>0||m.wolne>0;
                  return (
                    <tr key={i} style={{ background:i%2===0?"#FAFAF7":"#fff", opacity:hasData?1:0.45 }}>
                      <td style={{ padding:"5px 8px", fontWeight:600, color:"#1F4D38" }}>{MONTHS[i]}</td>
                      <td style={{ padding:"5px 4px", textAlign:"center" }}>{m.work||""}</td>
                      <td style={{ padding:"5px 4px", textAlign:"center", color:"#6B6A60" }}>{m.workH>0?m.workH.toFixed(1):""}</td>
                      <td style={{ padding:"5px 4px", textAlign:"center", background:m.urlop>0?"#E8F0FA":"transparent", fontWeight:m.urlop>0?700:400, color:"#3B6FA0" }}>{m.urlop||""}</td>
                      <td style={{ padding:"5px 4px", textAlign:"center" }}>{m.urlopB||""}</td>
                      <td style={{ padding:"5px 4px", textAlign:"center", background:m.l4>0?"#FBEAE8":"transparent", fontWeight:m.l4>0?700:400, color:"#C75146" }}>{m.l4||""}</td>
                      <td style={{ padding:"5px 4px", textAlign:"center" }}>{m.l4b||""}</td>
                      <td style={{ padding:"5px 4px", textAlign:"center" }}>{m.opC||""}</td>
                      <td style={{ padding:"5px 4px", textAlign:"center" }}>{m.opZ||""}</td>
                      <td style={{ padding:"5px 4px", textAlign:"center", color:"#9C988A" }}>{m.wolne||""}</td>
                    </tr>
                  );
                })}
                <tr style={{ background:"#EAF3E2", borderTop:"2px solid #1F4D38", fontWeight:700 }}>
                  <td style={{ padding:"6px 8px", color:"#1F4D38" }}>SUMA</td>
                  <td style={{ padding:"6px 4px", textAlign:"center" }}>{totals.work}</td>
                  <td style={{ padding:"6px 4px", textAlign:"center", color:"#6B6A60" }}>{totals.workH.toFixed(1)}</td>
                  <td style={{ padding:"6px 4px", textAlign:"center", color:"#3B6FA0" }}>{totals.urlop}</td>
                  <td style={{ padding:"6px 4px", textAlign:"center" }}>{totals.urlopB}</td>
                  <td style={{ padding:"6px 4px", textAlign:"center", color:"#C75146" }}>{totals.l4}</td>
                  <td style={{ padding:"6px 4px", textAlign:"center" }}>{totals.l4b}</td>
                  <td style={{ padding:"6px 4px", textAlign:"center" }}>{totals.opC}</td>
                  <td style={{ padding:"6px 4px", textAlign:"center" }}>{totals.opZ}</td>
                  <td style={{ padding:"6px 4px", textAlign:"center", color:"#9C988A" }}>{totals.wolne}</td>
                </tr>
              </tbody>
            </table>
          </div>
        )}

        {/* TAB: LISTA NIEOBECNOŚCI */}
        {activeTab==="lista" && (
          <div>
            {absenceList.length===0 ? (
              <div style={{ textAlign:"center", color:"#9C988A", fontSize:13, padding:32 }}>Brak nieobecnosci w {year} roku.</div>
            ) : (
              <>
                <div style={{ fontSize:12, color:"#6B6A60", marginBottom:8 }}>Lacznie: <b>{absenceList.length} dni nieobecnosci</b></div>
                <div style={{ display:"flex", flexDirection:"column", gap:4 }}>
                  {absenceList.map((a,i)=>{
                    const ts = TYPES[a.type]||{bg:"#EEE",text:"#333"};
                    return (
                      <div key={i} style={{ display:"flex", alignItems:"center", gap:10, padding:"5px 10px", borderRadius:7, background:i%2===0?"#FAFAF7":"#fff", border:"1px solid #F1EEE3" }}>
                        <div style={{ fontSize:12, fontWeight:600, color:"#1F4D38", minWidth:90 }}>{a.date}</div>
                        <div style={{ fontSize:11, padding:"2px 8px", borderRadius:5, background:ts.bg, color:ts.text, fontWeight:600 }}>{a.label}</div>
                      </div>
                    );
                  })}
                </div>
              </>
            )}
          </div>
        )}

      </div>
    </Overlay>
  );
}

function EmployeeModal({ modal, onClose, onSave, onDelete, isAdmin, positions }) {
  const [name, setName] = useState(modal.name);
  const [rate, setRate] = useState(modal.rate);
  const [position, setPosition] = useState(modal.position||positions[0]||"");
  const [etat, setEtat] = useState(modal.etat||"1");
  const [paymentMethod, setPaymentMethod] = useState(modal.paymentMethod||"przelew");
  const [defaultTransferAmount, setDefaultTransferAmount] = useState(modal.defaultTransferAmount||0);
  const [contractEndDate, setContractEndDate] = useState(modal.contractEndDate||null);
  const [showContractEnd, setShowContractEnd] = useState(!!(modal.contractEndDate));
  const [contractStartDate, setContractStartDate] = useState(modal.contractStartDate||"");
  const [vacationLimit, setVacationLimit] = useState(modal.vacationLimit||20);
  const [vacationTransferUsed, setVacationTransferUsed] = useState(modal.vacationTransferUsed||0);
  const [vacationTransferYear, setVacationTransferYear] = useState(modal.vacationTransferYear||new Date().getFullYear());
  const [vacationCarryOver, setVacationCarryOver] = useState(modal.vacationCarryOver||0);
  return (
    <Overlay onClose={onClose}>
      <ModalHeader title={modal.id?"Edytuj pracownika":"Nowy pracownik"} onClose={onClose}/>
      <div style={{ padding:18 }}>
        <Field label="Imię i nazwisko"><input value={name} onChange={e=>setName(e.target.value)} style={inputStyle}/></Field>
        <Field label="Stanowisko">
          {positions.length>0 ? <select value={position} onChange={e=>setPosition(e.target.value)} style={inputStyle}>{positions.map(p=><option key={p} value={p}>{p}</option>)}</select>
          : <div style={{ fontSize:11.5, color:"#9C988A" }}>Brak stanowisk -- dodaj je w Ustawieniach.</div>}
        </Field>
        <Field label="Wymiar etatu">
          <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
            {ETAT_OPTIONS.map(o=>(
              <button key={o} onClick={()=>setEtat(o)} style={{ padding:"6px 12px", borderRadius:8, border:etat===o?"2px solid #2F6B4F":"1px solid #E3DFD2", background:etat===o?"#2F6B4F":"#fff", color:etat===o?"#EAF3E2":"#3a392f", fontSize:12.5, fontWeight:600, cursor:"pointer" }}>{o==="1"?"Pełny etat (1)":o}</button>
            ))}
          </div>
          <div style={{ fontSize:11, color:"#9C988A", marginTop:4 }}>Wpływa na wyliczenie godzin urlopu i L4 (np. 1/2 etatu = {parseEtat("1/2")*8} godz./dzień).</div>
        </Field>
        {isAdmin ? (
          <>
            <Field label="Stawka godzinowa (zł)"><input type="text" inputMode="decimal" value={rate} onChange={e=>setRate(e.target.value)} onBlur={e=>setRate(parseFloat(String(e.target.value).replace(",","."))||0)} onFocus={e=>e.target.select()} placeholder="0,00" style={inputStyle}/></Field>
            <Field label="Sposób wypłaty wynagrodzenia">
              <div style={{ display:"flex", gap:6, flexWrap:"wrap" }}>
                {PAYMENT_METHODS.map(m=>(
                  <button key={m.id} onClick={()=>setPaymentMethod(m.id)} style={{ padding:"7px 11px", borderRadius:8, border:paymentMethod===m.id?"2px solid #2F6B4F":"1px solid #E3DFD2", background:paymentMethod===m.id?"#2F6B4F":"#fff", color:paymentMethod===m.id?"#EAF3E2":"#3a392f", fontSize:12, fontWeight:600, cursor:"pointer" }}>{m.label}</button>
                ))}
              </div>
            </Field>
            {paymentMethod==="przelew_gotowka" && (
              <div style={{ fontSize:11.5, color:"#9C988A", marginTop:6 }}>Kwotę przelewu ustawiasz co miesiąc przy rozliczeniu (kliknij ołówek przy pracowniku w zakładce Rozliczenia).</div>
            )}
          </>
        ) : (
          <div style={{ fontSize:12, color:"#9C988A", display:"flex", alignItems:"center", gap:6, marginBottom:14 }}><EyeOff size={13}/> Stawkę i sposób wypłaty może zmieniać tylko administrator.</div>
        )}

        <Field label="Data zatrudnienia (poczatek umowy)">
          <input type="date" value={contractStartDate||""} onChange={e=>setContractStartDate(e.target.value||"")} style={inputStyle}/>
          <div style={{ fontSize:11, color:"#9C988A", marginTop:4 }}>Uzywana do proporcjonalnego naliczania urlopu dla nowych pracownikow (art. 153 KP).</div>
        </Field>

        <div style={{ marginTop:6, borderTop:"1px solid #F1EEE3", paddingTop:14, marginBottom:6 }}>
          <div style={{ fontSize:12, fontWeight:600, color:"#6B6A60", marginBottom:8, display:"flex", alignItems:"center", gap:6 }}>
            <span>🔄</span> Korekta urlopu
          </div>
          <Field label="Rok rozliczeniowy korekty">
            <select value={vacationTransferYear} onChange={e=>setVacationTransferYear(Number(e.target.value))} style={inputStyle}>
              {[new Date().getFullYear()-1, new Date().getFullYear(), new Date().getFullYear()+1].map(y=><option key={y} value={y}>{y}</option>)}
            </select>
          </Field>
          <div style={{ display:"flex", gap:10, marginTop:6 }}>
            <div style={{ flex:1, background:"#F0FAF4", border:"1px solid #B6DFC8", borderRadius:8, padding:10 }}>
              <div style={{ fontSize:11.5, fontWeight:600, color:"#1F6B3A", marginBottom:6 }}>➕ Urlop zaległy z poprzedniego roku</div>
              <div style={{ fontSize:11, color:"#4A8A60", marginBottom:6 }}>Nieudzielony urlop przeniesiony na bieżący rok (art. 168 KP — termin do 30.09).</div>
              <input type="number" min="0" max="50" step="1" value={vacationCarryOver} onChange={e=>setVacationCarryOver(Math.max(0,Number(e.target.value)))} onFocus={e=>e.target.select()} style={{...inputStyle, background:"#fff"}} placeholder="0"/>
              {vacationCarryOver>0 && <div style={{ fontSize:10.5, color:"#2F6B4F", marginTop:4 }}>+{vacationCarryOver} dni do limitu {vacationTransferYear}</div>}
            </div>
            <div style={{ flex:1, background:"#FFF9EC", border:"1px solid #F0D88A", borderRadius:8, padding:10 }}>
              <div style={{ fontSize:11.5, fontWeight:600, color:"#7A4F00", marginBottom:6 }}>➖ Urlop zużyty w poprzednim sklepie</div>
              <div style={{ fontSize:11, color:"#9C7030", marginBottom:6 }}>Dni urlopu już wykorzystane w innym sklepie w roku {vacationTransferYear}.</div>
              <input type="number" min="0" max="50" step="1" value={vacationTransferUsed} onChange={e=>setVacationTransferUsed(Math.max(0,Number(e.target.value)))} onFocus={e=>e.target.select()} style={{...inputStyle, background:"#fff"}} placeholder="0"/>
              {vacationTransferUsed>0 && <div style={{ fontSize:10.5, color:"#C75146", marginTop:4 }}>−{vacationTransferUsed} dni od dostępnego urlopu</div>}
            </div>
          </div>
          {(vacationCarryOver>0||vacationTransferUsed>0) && (
            <div style={{ fontSize:11, color:"#2F6B4F", background:"#F0FAF4", border:"1px solid #B6DFC8", borderRadius:6, padding:"5px 10px", marginTop:8 }}>
              Efekt netto w roku {vacationTransferYear}: {vacationCarryOver>0&&<span>+{vacationCarryOver} zaległy </span>}{vacationTransferUsed>0&&<span>−{vacationTransferUsed} zużyty w poprz. sklepie </span>}→ korekta salda: <b>{vacationCarryOver-vacationTransferUsed>=0?"+":""}{vacationCarryOver-vacationTransferUsed} dni</b>
            </div>
          )}
        </div>

        <Field label="Limit urlopu wypoczynkowego (dni/rok)">
          <div style={{ display:"flex", gap:6 }}>
            {[20,26].map(v=>(
              <button key={v} onClick={()=>setVacationLimit(v)} style={{ flex:1, padding:"7px 10px", borderRadius:8, border:vacationLimit===v?"2px solid #2F6B4F":"1px solid #E3DFD2", background:vacationLimit===v?"#2F6B4F":"#fff", color:vacationLimit===v?"#EAF3E2":"#3a392f", fontSize:13, fontWeight:600, cursor:"pointer" }}>{v} dni</button>
            ))}
          </div>
          <div style={{ fontSize:11, color:"#9C988A", marginTop:4 }}>20 dni - do 10 lat stazu pracy, 26 dni - powyzej 10 lat.</div>
        </Field>

        <div style={{ marginTop:6, borderTop:"1px solid #F1EEE3", paddingTop:14 }}>
          <div style={{ fontSize:12, fontWeight:600, color:"#6B6A60", marginBottom:8 }}>Status umowy</div>
          <div style={{ display:"flex", gap:8, marginBottom:showContractEnd?10:0 }}>
            <button onClick={()=>{ setShowContractEnd(false); setContractEndDate(null); }} style={{ flex:1, padding:"7px 10px", borderRadius:8, border:!showContractEnd?"2px solid #2F6B4F":"1px solid #E3DFD2", background:!showContractEnd?"#2F6B4F":"#fff", color:!showContractEnd?"#EAF3E2":"#3a392f", fontSize:12.5, fontWeight:600, cursor:"pointer" }}>Aktywna</button>
            <button onClick={()=>{ setShowContractEnd(true); if(!contractEndDate) setContractEndDate(dateStr(new Date())); }} style={{ flex:1, padding:"7px 10px", borderRadius:8, border:showContractEnd?"2px solid #C75146":"1px solid #E3DFD2", background:showContractEnd?"#C75146":"#fff", color:showContractEnd?"#fff":"#3a392f", fontSize:12.5, fontWeight:600, cursor:"pointer" }}>Zakonczona</button>
          </div>
          {showContractEnd && (
            <div>
              <div style={{ fontSize:12, color:"#6B6A60", marginBottom:6 }}>Data ostatniego dnia pracy</div>
              <input type="date" value={contractEndDate||""} onChange={e=>setContractEndDate(e.target.value||null)} style={inputStyle}/>
              <div style={{ fontSize:11, color:"#9C988A", marginTop:4 }}>Od kolejnego dnia pracownik nie bedzie widoczny w grafiku. W rozliczeniu miesiac, w ktorym umowa sie konczy, zostanie rozliczony tylko do tej daty.</div>
            </div>
          )}
        </div>
      </div>
      <div style={{ padding:"12px 18px", borderTop:"1px solid #EFEBDE", display:"flex", justifyContent:"space-between" }}>
        {modal.id ? <button onClick={()=>onDelete(modal.id)} style={dangerBtn}><Trash2 size={14}/> Usuń</button> : <span/>}
        <div style={{ display:"flex", gap:8 }}>
          <button onClick={onClose} style={ghostBtn}>Anuluj</button>
          <button onClick={()=>name.trim()&&onSave({ id:modal.id, name:name.trim(), rate, position, etat, paymentMethod, defaultTransferAmount, contractEndDate:showContractEnd?contractEndDate:null, contractStartDate:contractStartDate||null, vacationLimit, vacationTransferUsed: vacationTransferUsed||0, vacationTransferYear: vacationTransferYear||new Date().getFullYear(), vacationCarryOver: vacationCarryOver||0 })} style={primaryBtn}><Check size={14}/> Zapisz</button>
        </div>
      </div>
    </Overlay>
  );
}

// ============ WIDOK ROZLICZEŃ ============
function PayrollView({ payMonth, setPayMonth, rows, totalPay, totalHours, isAdmin, settings, onOpenAdj, storeName, shifts, electronAvailable }) {
  const shiftFn = (n) => { const d=new Date(payMonth); d.setMonth(d.getMonth()+n); setPayMonth(d); };
  const [printCard, setPrintCard] = useState(null);
  const [exporting, setExporting] = useState(false);
  // Brak auto-print - użytkownik sam klika "Drukuj" w podglądzie

  const handleExportDB = () => {
    // Eksport całej bazy danych do pliku JSON (dla admina zdalnego)
    const key = "grafik_data_v2";
    const raw = localStorage.getItem(key);
    if (!raw) { alert("Brak danych do eksportu."); return; }
    const blob = new Blob([raw], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    const d = new Date();
    a.download = `grafik_backup_${storeName.replace(/\s+/g,"_")}_${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleImportDB = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const parsed = JSON.parse(ev.target.result);
        // Walidacja podstawowa
        if (!parsed.employees && !parsed.shifts && !parsed.stores) {
          alert("Nieprawidlowy plik - brak danych grafiku.");
          return;
        }
        if (!window.confirm("Czy na pewno chcesz wczytac dane z pliku? Aktualne dane zostana zastapione.")) return;
        localStorage.setItem("grafik_data_v2", ev.target.result);
        window.location.reload();
      } catch {
        alert("Blad odczytu pliku JSON. Sprawdz czy plik jest poprawny.");
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  };

  const handleExportExcel = async () => {
    setExporting(true);
    try {
      const y=payMonth.getFullYear(), m=payMonth.getMonth();
      const toHM = (h) => { const m=Math.round(h*60); return `${Math.floor(m/60)}:${String(m%60).padStart(2,'0')}`; };
      const summaryRows = isAdmin
        ? rows.map(r=>({ "Pracownik":displayName(r.e.name), "Stanowisko":r.e.position||"", "Stawka godz. (zł)":Number(r.e.rate.toFixed(2)), "Godziny pracy":Number(r.hours.toFixed(2)), "Dni wolne":r.wolne, "Urlop":r.urlop, "Urlop bezpłatny":r.urlopBezplatny, "L4":r.l4, "L4 kod B":r.l4b, "Opieka chore (dni)":r.opiekaChore, "Opieka zdrowe (dni)":r.opiekaZdrowe, "Dodatki (zł)":Number(r.increase.toFixed(2)), "Potrącenia (zł)":Number(r.decrease.toFixed(2)), "Wynagrodzenie (zł)":Number(r.pay.toFixed(2)), "Przelew (zł)":Number(r.transfer.toFixed(2)), "Gotówka (zł)":Number(r.cash.toFixed(2)) }))
        : rows.map(r=>({ "Pracownik":displayName(r.e.name), "Stanowisko":r.e.position||"", "Godziny pracy":toHM(r.hours), "Dni wolne":r.wolne, "Urlop":r.urlop, "Urlop bezpłatny":r.urlopBezplatny, "L4":r.l4, "L4 kod B":r.l4b, "Opieka chore (dni)":r.opiekaChore, "Opieka zdrowe (dni)":r.opiekaZdrowe }));
      const empIds=new Set(rows.map(r=>r.e.id));
      const empById=Object.fromEntries(rows.map(r=>[r.e.id,r.e]));
      const dailyRows = shifts.filter(s=>{ if(!empIds.has(s.employeeId)) return false; const d=new Date(s.date); return d.getFullYear()===y&&d.getMonth()===m; }).sort((a,b)=>a.date===b.date?empById[a.employeeId]?.name.localeCompare(empById[b.employeeId]?.name):a.date.localeCompare(b.date)).map(s=>({ "Data":s.date, "Pracownik":displayName(empById[s.employeeId]?.name||""), "Stanowisko":empById[s.employeeId]?.position||"", "Typ dnia":TYPES[s.type]?.label||s.type, "Godzina od":s.type==="praca"?(s.start||""):"", "Godzina do":s.type==="praca"?(s.end||""):"", "Godziny":s.type==="praca"&&s.start&&s.end?Number(((toMin(s.end)-toMin(s.start))/60).toFixed(2)):0, "Zadania":(s.tasks||[]).join("; ") }));

      if (electronAvailable) {
        const res = await window.api.export.excel({ storeName, year:y, month:m+1, summaryRows, dailyRows });
        if (res?.success) { /* dialog sam wybrał ścieżkę */ }
        return;
      }
      const wb=XLSX.utils.book_new();
      // Wiersz sum dla arkusza Rozliczenie
      const sumRow = isAdmin
        ? { "Pracownik":"SUMA", "Stanowisko":"", "Stawka godz. (zł)":"", "Godziny pracy":Number(summaryRows.reduce((a,r)=>a+r["Godziny pracy"],0).toFixed(2)), "Dni wolne":summaryRows.reduce((a,r)=>a+r["Dni wolne"],0), "Urlop":summaryRows.reduce((a,r)=>a+r["Urlop"],0), "Urlop bezpłatny":summaryRows.reduce((a,r)=>a+r["Urlop bezpłatny"],0), "L4":summaryRows.reduce((a,r)=>a+r["L4"],0), "L4 kod B":summaryRows.reduce((a,r)=>a+r["L4 kod B"],0), "Opieka chore (dni)":summaryRows.reduce((a,r)=>a+r["Opieka chore (dni)"],0), "Opieka zdrowe (dni)":summaryRows.reduce((a,r)=>a+r["Opieka zdrowe (dni)"],0), "Dodatki (zł)":Number(summaryRows.reduce((a,r)=>a+r["Dodatki (zł)"],0).toFixed(2)), "Potrącenia (zł)":Number(summaryRows.reduce((a,r)=>a+r["Potrącenia (zł)"],0).toFixed(2)), "Wynagrodzenie (zł)":Number(summaryRows.reduce((a,r)=>a+r["Wynagrodzenie (zł)"],0).toFixed(2)), "Przelew (zł)":Number(summaryRows.reduce((a,r)=>a+r["Przelew (zł)"],0).toFixed(2)), "Gotówka (zł)":Number(summaryRows.reduce((a,r)=>a+r["Gotówka (zł)"],0).toFixed(2)) }
        : { "Pracownik":"SUMA", "Stanowisko":"", "Godziny pracy":"", "Dni wolne":summaryRows.reduce((a,r)=>a+r["Dni wolne"],0), "Urlop":summaryRows.reduce((a,r)=>a+r["Urlop"],0), "Urlop bezpłatny":summaryRows.reduce((a,r)=>a+r["Urlop bezpłatny"],0), "L4":summaryRows.reduce((a,r)=>a+r["L4"],0), "L4 kod B":summaryRows.reduce((a,r)=>a+r["L4 kod B"],0), "Opieka chore (dni)":summaryRows.reduce((a,r)=>a+r["Opieka chore (dni)"],0), "Opieka zdrowe (dni)":summaryRows.reduce((a,r)=>a+r["Opieka zdrowe (dni)"],0) };
      const ws1=XLSX.utils.json_to_sheet([...summaryRows, sumRow]);
      // Pogrubienie wiersza sum
      const sumRowIdx = summaryRows.length + 1; // 1-indexed, +1 nagłówek
      const cols = Object.keys(sumRow);
      cols.forEach((col, ci) => {
        const cellAddr = XLSX.utils.encode_cell({ r: sumRowIdx, c: ci });
        if (ws1[cellAddr]) ws1[cellAddr].s = { font:{ bold:true }, fill:{ fgColor:{ rgb:"E8F5E9" } } };
      });
      ws1["!cols"]=Object.keys(summaryRows[0]||{}).map(k=>({ wch:Math.max(14,k.length+2) }));
      XLSX.utils.book_append_sheet(wb,ws1,`Rozliczenie ${pad2(m+1)}-${y}`.slice(0,31));
      const ws2=XLSX.utils.json_to_sheet(dailyRows);
      ws2["!cols"]=[12,22,18,16,10,10,10,50].map(w=>({ wch:w }));
      ws2["!autofilter"]={ ref:`A1:H${dailyRows.length+1}` };
      ws2["!freeze"]={ xSplit:0, ySplit:1, topLeftCell:"A2", activePane:"bottomLeft" };
      XLSX.utils.book_append_sheet(wb,ws2,"Grafik dzienny");

      // Arkusz 3: Nieobecności z dokładnymi datami
      const absenceRows = [];
      rows.forEach(r => {
        const ab = r.absenceDates||{};
        [["urlop","Urlop"],["l4","L4"],["l4b","L4 ciążowe (kod B)"],["urlopBezplatny","Urlop bezpłatny"],["opiekaChore","Opieka (chore dz./rodzina)"],["opiekaZdrowe","Opieka (zdrowe dziecko)"]].forEach(([key,label])=>{
          (ab[key]||[]).forEach(date=>{
            const hpd = settings.refHours * parseEtat(r.e.etat||"1");
            absenceRows.push({"Data":date,"Pracownik":displayName(r.e.name),"Stanowisko":r.e.position||"","Typ nieobecności":label,"Godziny":Number(hpd.toFixed(2)),"Stawka (zł)":Number(r.e.rate.toFixed(2))});
          });
        });
      });
      absenceRows.sort((a,b)=>a["Data"].localeCompare(b["Data"])||a["Pracownik"].localeCompare(b["Pracownik"]));
      const ws3 = XLSX.utils.json_to_sheet(absenceRows.length ? absenceRows : [{"Data":"--","Pracownik":"Brak płatnych nieobecności","Stanowisko":"","Typ nieobecności":"","Godziny":0,"Stawka (zł)":0}]);
      ws3["!cols"]=[12,22,18,22,10,12].map(w=>({ wch:w }));
      if (absenceRows.length) { ws3["!autofilter"]={ ref:`A1:F${absenceRows.length+1}` }; ws3["!freeze"]={ xSplit:0, ySplit:1, topLeftCell:"A2", activePane:"bottomLeft" }; }
      XLSX.utils.book_append_sheet(wb,ws3,"Nieobecności z datami");
      XLSX.writeFile(wb,`Rozliczenie_${(storeName||"Sklep").replace(/\s+/g,"_")}_${y}-${pad2(m+1)}.xlsx`);
    } finally { setExporting(false); }
  };

  return (
    <div>
      <div style={{ display:"flex", alignItems:"center", justifyContent:"space-between", marginBottom:20, flexWrap:"wrap", gap:10 }}>
        <div style={{ display:"flex", alignItems:"center", gap:6, background:"#fff", border:"1px solid #E4E7EE", borderRadius:10, padding:"4px 4px", boxShadow:"0 1px 3px rgba(0,0,0,0.04)" }}>
          <button onClick={()=>shiftFn(-1)} style={{ ...navBtn, border:"none", background:"transparent" }}><ChevronLeft size={18}/></button>
          <div style={{ fontWeight:700, fontSize:14.5, minWidth:150, textAlign:"center", textTransform:"capitalize", color:"#1A2530" }}>{monthLabel(payMonth)}</div>
          <button onClick={()=>shiftFn(1)} style={{ ...navBtn, border:"none", background:"transparent" }}><ChevronRight size={18}/></button>
        </div>
        <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
            <button onClick={()=>setPrintCard("all")} style={ghostBtn}><Printer size={14}/>Drukuj karty</button>
            <button onClick={handleExportExcel} disabled={exporting||rows.length===0} style={{ ...primaryBtn, opacity:exporting||rows.length===0?0.6:1 }}><Download size={14}/> {exporting?"Generuję...":"Eksport Excel"}</button>
            <label style={{ ...ghostBtn, cursor:"pointer" }} title="Wczytaj plik bazy z innego sklepu">
              <Upload size={14}/> Import bazy
              <input type="file" accept=".json" onChange={handleImportDB} style={{ display:"none" }}/>
            </label>
            <button onClick={handleExportDB} style={ghostBtn} title="Eksportuj bazę danych do pliku JSON"><Database size={14}/> Eksport bazy</button>
          </div>
      </div>
      {!isAdmin && <div style={{ display:"flex", alignItems:"center", gap:6, fontSize:12.5, color:"#8890A0", marginBottom:14, background:"#F7F9FC", border:"1px solid #E4E7EE", borderRadius:8, padding:"8px 12px" }}><EyeOff size={13}/> Stawki i kwoty widoczne tylko dla administratora.</div>}
      {isAdmin && (
        <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit, minmax(180px, 1fr))", gap:12, marginBottom:18 }}>
          <StatCard label="Suma godzin pracy" value={`${(()=>{const m=Math.round(totalHours*60);return `${Math.floor(m/60)}:${String(m%60).padStart(2,'0')}`;})()}`}/>
          <StatCard label="Suma wynagrodzeń" value={`${totalPay.toFixed(2)} zł`} accent/>
          <StatCard label="Liczba pracowników" value={rows.length}/>
        </div>
      )}
      <div style={{ background:"#fff", border:"1px solid #E4E7EE", borderRadius:12, overflow:"auto", boxShadow:"0 1px 4px rgba(0,0,0,0.05)" }}>
        <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13 }}>
          <thead>
            <tr style={{ background:"#F7F9FC", textAlign:"left", borderBottom:"1px solid #E4E7EE" }}>
              {["Pracownik",...(isAdmin?["Stawka/godz"]:[]),"Godziny","Wolne","Urlop","Urlop bezpł.","L4","L4 kod B","Op. chore","Op. zdrowe",...(isAdmin?["Dodatki","Potrącenia","Wynagrodzenie","Przelew","Gotówka"]:[]),""].map((h,i)=>(
                <th key={i} style={{ padding:"10px 14px", fontWeight:700, color:"#8890A0", fontSize:11, textTransform:"uppercase", letterSpacing:"0.04em", whiteSpace:"nowrap" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map(r=>(
              <tr key={r.e.id} style={{ borderTop:"1px solid #F1F3F7" }}>
                <td style={{ padding:"10px 14px", fontWeight:600, whiteSpace:"nowrap", color:"#1A2530" }}>{displayName(r.e.name)}</td>
                {isAdmin && <td style={{ padding:"10px 14px", color:"#5A6070" }}>{r.e.rate.toFixed(2)} zł</td>}
                <td style={{ padding:"10px 14px", color:"#5A6070" }}>{(()=>{const m=Math.round(r.hours*60);return `${Math.floor(m/60)}:${String(m%60).padStart(2,'0')}`;})()}</td>
                <td style={{ padding:"10px 14px", color:"#5A6070" }}>{r.wolne}</td>
                <td style={{ padding:"10px 14px", color:"#5A6070" }}>{r.urlop}</td>
                <td style={{ padding:"10px 14px", color:"#5A6070" }}>{r.urlopBezplatny}</td>
                <td style={{ padding:"10px 14px", color:"#5A6070" }}>{r.l4}</td>
                <td style={{ padding:"10px 14px", color:"#5A6070" }}>{r.l4b}</td>
                <td style={{ padding:"10px 14px", color:"#5A6070" }}>{r.opiekaChore}</td>
                <td style={{ padding:"10px 14px", color:"#5A6070" }}>{r.opiekaZdrowe}</td>
                {isAdmin && <>
                  <td style={{ padding:"10px 14px", color:"#1F6B45", fontWeight:600 }}>{r.increase>0?`+${r.increase.toFixed(2)} zł`:"–"}</td>
                  <td style={{ padding:"10px 14px", color:"#C23B3B", fontWeight:600 }}>{r.decrease>0?`-${r.decrease.toFixed(2)} zł`:"–"}</td>
                  <td style={{ padding:"10px 14px", fontWeight:700, color:"#1A2530", fontSize:13.5 }}>{r.pay.toFixed(2)} zł</td>
                  <td style={{ padding:"10px 14px", color:"#5A6070" }}>{r.transfer>0?`${r.transfer.toFixed(2)} zł`:"–"}</td>
                  <td style={{ padding:"10px 14px", color:r.cash<0?"#C23B3B":"#5A6070" }}>{r.cash<0?`${r.cash.toFixed(2)} zł (!)`:r.cash>0?`${r.cash.toFixed(2)} zł`:"–"}</td>
                </>}
                <td style={{ padding:"10px 14px" }}>
                  <div style={{ display:"flex", gap:4 }}>
                    <button onClick={()=>onOpenAdj(r.e)} style={iconBtn} title="Dodatki i potrącenia"><Pencil size={13}/></button>
                    <button onClick={()=>setPrintCard(r)} style={iconBtn} title="Drukuj kartę"><Printer size={13}/></button>
                  </div>
                </td>
              </tr>
            ))}
            {rows.length===0 && <tr><td colSpan={isAdmin?14:8} style={{ padding:24, textAlign:"center", color:"#9399A6", fontSize:13 }}>Brak danych dla tego miesiąca.</td></tr>}
          </tbody>
        </table>
      </div>
      {isAdmin && <div style={{ fontSize:11.5, color:"#9399A6", marginTop:10 }}>L4 = {settings.l4Percent}% stawki · L4 kod B = 100% · Opieka (chore) = {settings.opiekaChorePercent||80}% · Opieka (zdrowe) = 100% · Urlop = 100% · {settings.refHours} godz./dzień (Ustawienia).</div>}

      {printCard && (
        <div id="print-area-wrapper" style={{ position:"fixed", inset:0, background:"#fff", zIndex:200, overflowY:"auto", padding:0 }}>
          <div id="print-area" style={{ background:"#fff", padding:"10mm 14mm" }}>
            <div className="no-print" style={{ position:"sticky", top:0, background:"#fff", borderBottom:"1px solid #E3DFD2", padding:"10px 0", marginBottom:16, display:"flex", gap:10, alignItems:"center", zIndex:201 }}>
              <button onClick={()=>window.print()} style={{ ...primaryBtn, gap:8 }}><Printer size={15}/> Drukuj</button>
              <button onClick={()=>setPrintCard(null)} style={ghostBtn}>✕ Zamknij podgląd</button>
              <span style={{ fontSize:12, color:"#9C988A" }}>Podglad wydruku kart pracownikow</span>
            </div>
            {(printCard==="all"?rows:[printCard]).map((r,idx)=>(
              <EmployeeCard key={r.e.id} r={r} payMonth={payMonth} settings={settings} storeName={storeName} pageBreak={idx>0} isAdmin={isAdmin}/>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

function EmployeeCard({ r, payMonth, settings, storeName, pageBreak, isAdmin }) {
  const method=PAYMENT_METHODS.find(m=>m.id===(r.e.paymentMethod||"przelew"))?.label||"";
  const etatLabel = r.e.etat && r.e.etat!=="1" ? ` (etat ${r.e.etat})` : "";
  const fmt = (v) => v.toFixed(2);
  const fmtH = (h) => { const totalMin=Math.round(h*60); return `${Math.floor(totalMin/60)}:${String(totalMin%60).padStart(2,'0')}`; };
  return (
    <div style={{ maxWidth:700, margin:"0 auto 40px", pageBreakBefore:pageBreak?"always":"auto", border:"1px solid #E3DFD2", borderRadius:12, padding:28, fontFamily:"inherit" }}>
      <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start", marginBottom:14, borderBottom:"2px solid #1F4D38", paddingBottom:12 }}>
        <div>
          <div style={{ fontWeight:700, fontSize:18, color:"#1F4D38" }}>Karta czasu pracy{isAdmin ? " i wynagrodzenia" : ""}</div>
          <div style={{ fontSize:12, color:"#6B6A60", marginTop:2, textTransform:"capitalize" }}>{storeName} · {monthLabel(payMonth)}</div>
          {r.e.contractEndDate && <div style={{ fontSize:11.5, color:"#C75146", marginTop:2, fontWeight:600 }}>Umowa zakonczona: {r.e.contractEndDate}</div>}
        </div>
        <div style={{ textAlign:"right" }}>
          <div style={{ fontWeight:700, fontSize:15 }}>{displayName(r.e.name)}</div>
          {etatLabel && <div style={{ fontSize:12, color:"#6B6A60" }}>{etatLabel.trim()}</div>}
          {isAdmin && <div style={{ fontSize:12, color:"#6B6A60" }}>Stawka: {r.e.rate.toFixed(2)} zl/godz</div>}
        </div>
      </div>

      {/* Czas pracy */}
      <div style={{ fontSize:11, fontWeight:700, color:"#1F4D38", textTransform:"uppercase", letterSpacing:0.5, marginBottom:6 }}>Czas pracy</div>
      <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12.5, marginBottom:12 }}>
        <thead>
          <tr style={{ background:"#FAF8F2" }}>
            <th style={{ padding:"5px 8px", textAlign:"left", borderBottom:"1px solid #E3DFD2", fontWeight:600, color:"#6B6A60" }}>Skladnik</th>
            <th style={{ padding:"5px 8px", textAlign:"right", borderBottom:"1px solid #E3DFD2", fontWeight:600, color:"#6B6A60" }}>Godziny</th>
            {isAdmin && <th style={{ padding:"5px 8px", textAlign:"right", borderBottom:"1px solid #E3DFD2", fontWeight:600, color:"#6B6A60" }}>Stawka</th>}
            {isAdmin && <th style={{ padding:"5px 8px", textAlign:"right", borderBottom:"1px solid #E3DFD2", fontWeight:600, color:"#6B6A60" }}>Kwota (zl)</th>}
          </tr>
        </thead>
        <tbody>
          <tr style={{ borderBottom:"1px solid #F1EEE3" }}>
            <td style={{ padding:"6px 8px" }}>Przepracowane</td>
            <td style={{ padding:"6px 8px", textAlign:"right" }}>{fmtH(r.hours)}</td>
            {isAdmin && <td style={{ padding:"6px 8px", textAlign:"right" }}>{r.e.rate.toFixed(2)} zl</td>}
            {isAdmin && <td style={{ padding:"6px 8px", textAlign:"right", fontWeight:600 }}>{fmt(r.workPay)}</td>}
          </tr>
          {r.urlop>0 && <tr style={{ borderBottom:"1px solid #F1EEE3" }}>
            <td style={{ padding:"6px 8px" }}>Urlop ({r.urlop} dni)</td>
            <td style={{ padding:"6px 8px", textAlign:"right" }}>{fmtH(r.urlopHours)}</td>
            {isAdmin && <td style={{ padding:"6px 8px", textAlign:"right" }}>{r.e.rate.toFixed(2)} zl (100%)</td>}
            {isAdmin && <td style={{ padding:"6px 8px", textAlign:"right", fontWeight:600 }}>{fmt(r.urlopPay)}</td>}
          </tr>}
          {r.l4>0 && <tr style={{ borderBottom:"1px solid #F1EEE3" }}>
            <td style={{ padding:"6px 8px" }}>L4 ({r.l4} dni)</td>
            <td style={{ padding:"6px 8px", textAlign:"right" }}>{fmtH(r.l4Hours)}</td>
            {isAdmin && <td style={{ padding:"6px 8px", textAlign:"right" }}>{r.e.rate.toFixed(2)} zl ({settings.l4Percent}%)</td>}
            {isAdmin && <td style={{ padding:"6px 8px", textAlign:"right", fontWeight:600 }}>{fmt(r.l4Pay)}</td>}
          </tr>}
          {r.l4b>0 && <tr style={{ borderBottom:"1px solid #F1EEE3" }}>
            <td style={{ padding:"6px 8px" }}>L4 ciazowe kod B ({r.l4b} dni)</td>
            <td style={{ padding:"6px 8px", textAlign:"right" }}>{fmtH(r.l4bHours)}</td>
            {isAdmin && <td style={{ padding:"6px 8px", textAlign:"right" }}>{r.e.rate.toFixed(2)} zl (100%)</td>}
            {isAdmin && <td style={{ padding:"6px 8px", textAlign:"right", fontWeight:600 }}>{fmt(r.l4bPay)}</td>}
          </tr>}
          {r.opiekaChore>0 && <tr style={{ borderBottom:"1px solid #F1EEE3" }}>
            <td style={{ padding:"6px 8px" }}>Opieka (chore dziecko/rodzina) ({r.opiekaChore} dni)</td>
            <td style={{ padding:"6px 8px", textAlign:"right" }}>{fmtH(r.opiekaChoreHours)}</td>
            {isAdmin && <td style={{ padding:"6px 8px", textAlign:"right" }}>{r.e.rate.toFixed(2)} zl ({settings.opiekaChorePercent||80}%)</td>}
            {isAdmin && <td style={{ padding:"6px 8px", textAlign:"right", fontWeight:600 }}>{fmt(r.opiekaChorePay)}</td>}
          </tr>}
          {r.opiekaZdrowe>0 && <tr style={{ borderBottom:"1px solid #F1EEE3" }}>
            <td style={{ padding:"6px 8px" }}>Opieka (zdrowe dziecko) ({r.opiekaZdrowe} dni)</td>
            <td style={{ padding:"6px 8px", textAlign:"right" }}>{fmtH(r.opiekaZdroweHours)}</td>
            {isAdmin && <td style={{ padding:"6px 8px", textAlign:"right" }}>{r.e.rate.toFixed(2)} zl (100%)</td>}
            {isAdmin && <td style={{ padding:"6px 8px", textAlign:"right", fontWeight:600 }}>{fmt(r.opiekaZdrowePay)}</td>}
          </tr>}
          {r.urlopBezplatny>0 && <tr style={{ borderBottom:"1px solid #F1EEE3" }}>
            <td style={{ padding:"6px 8px", color:"#6B6A60" }}>Urlop bezplatny ({r.urlopBezplatny} dni / {fmtH(r.urlopBezplatnyHours)})</td>
            <td style={{ padding:"6px 8px", textAlign:"right", color:"#6B6A60" }}>{fmtH(r.urlopBezplatnyHours)}</td>
            {isAdmin && <td style={{ padding:"6px 8px", textAlign:"right", color:"#6B6A60" }}>--</td>}
            {isAdmin && <td style={{ padding:"6px 8px", textAlign:"right", color:"#6B6A60" }}>0,00</td>}
          </tr>}
          {r.wolne>0 && <tr style={{ borderBottom:"1px solid #F1EEE3" }}>
            <td style={{ padding:"6px 8px", color:"#9C988A" }}>Dni wolne ({r.wolne} dni)</td>
            <td style={{ padding:"6px 8px", textAlign:"right", color:"#9C988A" }}>--</td>
            {isAdmin && <td style={{ padding:"6px 8px", textAlign:"right", color:"#9C988A" }}>--</td>}
            {isAdmin && <td style={{ padding:"6px 8px", textAlign:"right", color:"#9C988A" }}>0,00</td>}
          </tr>}
          <tr style={{ borderTop:"2px solid #1F4D38", background:"#F7FAF8" }}>
            <td colSpan={isAdmin?4:2} style={{ padding:"8px 8px 2px", fontWeight:700, color:"#1F4D38", fontSize:11, textTransform:"uppercase", letterSpacing:0.4 }}>Podsumowanie godzin</td>
          </tr>
          {r.hours>0 && <tr style={{ background:"#F7FAF8" }}>
            <td style={{ padding:"2px 8px 2px 18px", color:"#3a392f" }}>Godziny przepracowane:</td>
            <td style={{ padding:"2px 8px", textAlign:"right", fontWeight:600 }}>{fmtH(r.hours)}</td>
            {isAdmin && <td colSpan={2}/>}
          </tr>}
          {r.urlopHours>0 && <tr style={{ background:"#F7FAF8" }}>
            <td style={{ padding:"2px 8px 2px 18px", color:"#3a392f" }}>Godziny urlopu:</td>
            <td style={{ padding:"2px 8px", textAlign:"right", fontWeight:600 }}>{fmtH(r.urlopHours)}</td>
            {isAdmin && <td colSpan={2}/>}
          </tr>}
          {r.l4Hours>0 && <tr style={{ background:"#F7FAF8" }}>
            <td style={{ padding:"2px 8px 2px 18px", color:"#3a392f" }}>Godziny L4:</td>
            <td style={{ padding:"2px 8px", textAlign:"right", fontWeight:600 }}>{fmtH(r.l4Hours)}</td>
            {isAdmin && <td colSpan={2}/>}
          </tr>}
          {r.l4bHours>0 && <tr style={{ background:"#F7FAF8" }}>
            <td style={{ padding:"2px 8px 2px 18px", color:"#3a392f" }}>Godziny L4 ciazowe (kod B):</td>
            <td style={{ padding:"2px 8px", textAlign:"right", fontWeight:600 }}>{fmtH(r.l4bHours)}</td>
            {isAdmin && <td colSpan={2}/>}
          </tr>}
          {r.opiekaChoreHours>0 && <tr style={{ background:"#F7FAF8" }}>
            <td style={{ padding:"2px 8px 2px 18px", color:"#3a392f" }}>Godziny opieki (chore dz./rodzina):</td>
            <td style={{ padding:"2px 8px", textAlign:"right", fontWeight:600 }}>{fmtH(r.opiekaChoreHours)}</td>
            {isAdmin && <td colSpan={2}/>}
          </tr>}
          {r.opiekaZdroweHours>0 && <tr style={{ background:"#F7FAF8" }}>
            <td style={{ padding:"2px 8px 2px 18px", color:"#3a392f" }}>Godziny opieki (zdrowe dziecko):</td>
            <td style={{ padding:"2px 8px", textAlign:"right", fontWeight:600 }}>{fmtH(r.opiekaZdroweHours)}</td>
            {isAdmin && <td colSpan={2}/>}
          </tr>}
          {r.urlopBezplatnyHours>0 && <tr style={{ background:"#F7FAF8" }}>
            <td style={{ padding:"2px 8px 2px 18px", color:"#6B6A60" }}>Godziny urlopu bezplatnego:</td>
            <td style={{ padding:"2px 8px", textAlign:"right", color:"#6B6A60" }}>{fmtH(r.urlopBezplatnyHours)}</td>
            {isAdmin && <td colSpan={2}/>}
          </tr>}
          <tr style={{ background:"#EAF3E2", borderTop:"1px solid #B8D8C8" }}>
            <td style={{ padding:"6px 8px", fontWeight:700, color:"#1F4D38" }}>Lacznie godzin:</td>
            <td style={{ padding:"6px 8px", textAlign:"right", fontWeight:700, color:"#1F4D38", fontSize:14 }}>{fmtH(r.totalHours)}</td>
            {isAdmin && <td style={{ padding:"6px 8px", textAlign:"right", color:"#6B6A60", fontSize:10.5 }}>weryfikacja: {fmtH(r.hours+r.urlopHours+r.l4Hours+r.l4bHours+r.opiekaChoreHours+r.opiekaZdroweHours)} x {r.e.rate.toFixed(2)} zl =</td>}
            {isAdmin && <td style={{ padding:"6px 8px", textAlign:"right", fontWeight:700, color:"#1F4D38", fontSize:13 }}>{fmt(r.workPay+r.urlopPay+r.l4Pay+r.l4bPay+r.opiekaChorePay+r.opiekaZdrowePay)} zl</td>}
          </tr>
        </tbody>
      </table>

      {/* Korekty - widoczne dla wszystkich (kierownik wypelnia te dane) */}
      {(r.increase>0||r.decrease>0) && (
        <>
          <div style={{ fontSize:11, fontWeight:700, color:"#1F4D38", textTransform:"uppercase", letterSpacing:0.5, marginBottom:6 }}>Korekty</div>
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12.5, marginBottom:12 }}>
            <tbody>
              {r.adj.premia>0 && <tr style={{ borderBottom:"1px solid #F1EEE3" }}><td style={{ padding:"5px 8px" }}>Premia</td><td style={{ padding:"5px 8px", textAlign:"right", color:"#2F6B4F", fontWeight:600 }}>+{fmt(r.adj.premia)}</td></tr>}
              {r.adj.ulga26>0 && <tr style={{ borderBottom:"1px solid #F1EEE3" }}><td style={{ padding:"5px 8px" }}>Ulga dla mlodych (do 26 r.z.)</td><td style={{ padding:"5px 8px", textAlign:"right", color:"#2F6B4F", fontWeight:600 }}>+{fmt(r.adj.ulga26)}</td></tr>}
              {r.adj.custIncAmount1>0 && <tr style={{ borderBottom:"1px solid #F1EEE3" }}><td style={{ padding:"5px 8px" }}>{r.adj.custIncLabel1||"Dodatek 1"}</td><td style={{ padding:"5px 8px", textAlign:"right", color:"#2F6B4F", fontWeight:600 }}>+{fmt(r.adj.custIncAmount1)}</td></tr>}
              {r.adj.custIncAmount2>0 && <tr style={{ borderBottom:"1px solid #F1EEE3" }}><td style={{ padding:"5px 8px" }}>{r.adj.custIncLabel2||"Dodatek 2"}</td><td style={{ padding:"5px 8px", textAlign:"right", color:"#2F6B4F", fontWeight:600 }}>+{fmt(r.adj.custIncAmount2)}</td></tr>}
              {r.adj.ubezpieczenie>0 && <tr style={{ borderBottom:"1px solid #F1EEE3" }}><td style={{ padding:"5px 8px" }}>Ubezpieczenie dodatkowe</td><td style={{ padding:"5px 8px", textAlign:"right", color:"#C75146", fontWeight:600 }}>-{fmt(r.adj.ubezpieczenie)}</td></tr>}
              {r.adj.custDecAmount1>0 && <tr style={{ borderBottom:"1px solid #F1EEE3" }}><td style={{ padding:"5px 8px" }}>{r.adj.custDecLabel1||"Potracenie 1"}</td><td style={{ padding:"5px 8px", textAlign:"right", color:"#C75146", fontWeight:600 }}>-{fmt(r.adj.custDecAmount1)}</td></tr>}
              {r.adj.custDecAmount2>0 && <tr style={{ borderBottom:"1px solid #F1EEE3" }}><td style={{ padding:"5px 8px" }}>{r.adj.custDecLabel2||"Potracenie 2"}</td><td style={{ padding:"5px 8px", textAlign:"right", color:"#C75146", fontWeight:600 }}>-{fmt(r.adj.custDecAmount2)}</td></tr>}
            </tbody>
          </table>
        </>
      )}

      {/* Wynagrodzenie koncowe - tylko admin */}
      {isAdmin && (
        <>
          <div style={{ background:"#1F4D38", color:"#EAF3E2", borderRadius:8, padding:"10px 14px", marginBottom:12 }}>
            <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center" }}>
              <div style={{ fontWeight:700, fontSize:15 }}>Wynagrodzenie do wyplaty</div>
              <div style={{ fontWeight:700, fontSize:18 }}>{fmt(r.pay)} zl</div>
            </div>
            {r.transfer>0&&r.cash!==0 && (
              <div style={{ marginTop:8, fontSize:12.5, display:"flex", justifyContent:"space-between" }}>
                <span>Przelew na konto: <b>{fmt(r.transfer)} zl</b></span>
                <span style={{ color:r.cash<0?"#FFA07A":"inherit" }}>Gotowka: <b>{fmt(r.cash)} zl</b>{r.cash<0?" !":""}</span>
              </div>
            )}
            {r.cash>0&&r.transfer===0 && <div style={{ marginTop:4, fontSize:12.5 }}>Wyplata gotowka: <b>{fmt(r.cash)} zl</b></div>}
          </div>
          {r.cash!==0 && (
            <div style={{ border:`2px dashed ${r.cash<0?"#C75146":"#E07A3E"}`, borderRadius:8, padding:"8px 14px", marginBottom:12, display:"flex", justifyContent:"space-between", alignItems:"center" }}>
              <div style={{ fontWeight:600, fontSize:13 }}>Gotowka:</div>
              <div style={{ fontWeight:700, fontSize:16, color:r.cash<0?"#C75146":"#E07A3E" }}>{fmt(r.cash)} zl{r.cash<0?" (przelew przekracza wynagrodzenie -- do wyjasnienia)":""}</div>
            </div>
          )}
        </>
      )}
      {!isAdmin && (
        <div style={{ background:"#FAF8F2", border:"1px solid #E3DFD2", borderRadius:8, padding:"10px 14px", marginBottom:12, display:"flex", alignItems:"center", gap:8, color:"#9C988A", fontSize:12 }}>
          <EyeOff size={14}/> Stawka i kwoty wynagrodzenia widoczne tylko dla administratora.
        </div>
      )}

      <div style={{ display:"flex", justifyContent:"space-between", marginTop:40, fontSize:12, color:"#6B6A60" }}>
        <div style={{ borderTop:"1px solid #3a392f", paddingTop:6, width:"40%" }}>Podpis pracownika</div>
        <div style={{ borderTop:"1px solid #3a392f", paddingTop:6, width:"40%", textAlign:"right" }}>Podpis kierownika / administratora</div>
      </div>

      {/* ---- KALENDARZ DZIENNY - nowa strona ---- */}
      <div style={{ pageBreakBefore:"always", breakBefore:"always", paddingTop:4 }}>
        <div style={{ borderBottom:"2px solid #1F4D38", paddingBottom:10, marginBottom:14, display:"flex", justifyContent:"space-between", alignItems:"flex-end" }}>
          <div>
            <div style={{ fontWeight:700, fontSize:16, color:"#1F4D38" }}>Kalendarz czasu pracy</div>
            <div style={{ fontSize:12, color:"#6B6A60", marginTop:2, textTransform:"capitalize" }}>{displayName(r.e.name)} · {storeName} · {monthLabel(payMonth)}</div>
          </div>
          <div style={{ fontSize:11, color:"#9C988A", display:"flex", flexWrap:"wrap", gap:"4px 12px", maxWidth:340, justifyContent:"flex-end" }}>
            {[
              {color:"#D4EDDA",label:"Praca"},
              {color:"#C8E6FA",label:"Urlop"},
              {color:"#FDE8C8",label:"Urlop bezpł."},
              {color:"#FCD9D9",label:"L4"},
              {color:"#F5C6FA",label:"L4 kod B"},
              {color:"#E8E0FA",label:"Op. chore"},
              {color:"#D0F0F5",label:"Op. zdrowe"},
              {color:"#F0EDE6",label:"Wolne"},
            ].map(item=>(
              <span key={item.label} style={{ display:"inline-flex", alignItems:"center", gap:4 }}>
                <span style={{ display:"inline-block", width:11, height:11, borderRadius:2, background:item.color, border:"1px solid #ccc" }}/>
                {item.label}
              </span>
            ))}
          </div>
        </div>
        <EmployeeMonthCalendar r={r} payMonth={payMonth} />
      </div>
    </div>
  );
}

function EmployeeMonthCalendar({ r, payMonth }) {
  const y = payMonth.getFullYear();
  const m = payMonth.getMonth();
  const daysInMonth = new Date(y, m+1, 0).getDate();

  // Mapa: "YYYY-MM-DD" -> dane zmiany
  const shiftMap = {};
  (r.empShifts||[]).forEach(s => { shiftMap[s.date] = s; });

  // Kolory i etykiety typow
  const typeStyle = {
    praca:          { bg:"#D4EDDA", color:"#1B5E35", label:"Praca" },
    urlop:          { bg:"#C8E6FA", color:"#0D47A1", label:"Urlop" },
    urlop_bezplatny:{ bg:"#FDE8C8", color:"#7B4800", label:"U.bezpl." },
    l4:             { bg:"#FCD9D9", color:"#B71C1C", label:"L4" },
    l4b:            { bg:"#F5C6FA", color:"#6A1B9A", label:"L4 B" },
    opieka_chore:   { bg:"#E8E0FA", color:"#4527A0", label:"Op.chore" },
    opieka_zdrowe:  { bg:"#D0F0F5", color:"#006064", label:"Op.zdr." },
    wolne:          { bg:"#F0EDE6", color:"#6B6A60", label:"Wolne" },
  };

  const DAY_NAMES = ["Pn","Wt","Sr","Cz","Pt","So","Nd"];

  // Pierwszy dzien miesiaca (0=Nd, 1=Pn, ...) -> przesuniecie do Pn=0
  const firstDow = new Date(y, m, 1).getDay(); // 0=Nd
  const offset = firstDow === 0 ? 6 : firstDow - 1; // offset do poniedz.

  const cells = [];
  for (let i = 0; i < offset; i++) cells.push(null);
  for (let d = 1; d <= daysInMonth; d++) cells.push(d);

  const weeks = [];
  for (let i = 0; i < cells.length; i += 7) weeks.push(cells.slice(i, i+7));

  return (
    <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
      <thead>
        <tr>
          {DAY_NAMES.map(dn => (
            <th key={dn} style={{ padding:"4px 2px", textAlign:"center", fontWeight:600, color:"#1F4D38", fontSize:11, borderBottom:"2px solid #1F4D38", width:"14.28%" }}>
              {dn}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {weeks.map((week, wi) => (
          <tr key={wi}>
            {week.map((day, di) => {
              if (!day) return <td key={di} style={{ border:"1px solid #E3DFD2", background:"#FAFAFA", minHeight:64, padding:4 }}/>;
              const dateStr = `${y}-${pad2(m+1)}-${pad2(day)}`;
              const shift = shiftMap[dateStr];
              const dow = di; // 0=Pn ... 6=Nd
              const isWeekend = dow >= 5;
              const ts = shift ? (typeStyle[shift.type]||{bg:"#EEE",color:"#333",label:shift.type}) : null;
              const hours = shift && shift.type==="praca" && shift.start && shift.end
                ? ((toMin(shift.end)-toMin(shift.start))/60)
                : null;
              return (
                <td key={di} style={{
                  border:"1px solid #E3DFD2",
                  verticalAlign:"top",
                  padding:"4px 5px",
                  background: ts ? ts.bg : (isWeekend ? "#FAF7F0" : "#FFFFFF"),
                  minHeight:64,
                  width:"14.28%",
                }}>
                  <div style={{ fontWeight:600, fontSize:12, color: isWeekend ? "#9C988A" : "#1F4D38", marginBottom:2 }}>{day}</div>
                  {ts && (
                    <>
                      <div style={{ fontSize:10.5, fontWeight:600, color:ts.color, lineHeight:1.3 }}>
                        {ts.label}
                      </div>
                      {hours !== null && (
                        <div style={{ fontSize:11, color:ts.color, marginTop:2 }}>
                          {shift.start}–{shift.end}
                          <br/>
                          <b>{(()=>{const mm=Math.round(hours*60);return `${Math.floor(mm/60)}:${String(mm%60).padStart(2,'0')}`;})()}</b>
                        </div>
                      )}
                    </>
                  )}
                  {!ts && !isWeekend && (
                    <div style={{ fontSize:10, color:"#CCC", marginTop:2 }}>--</div>
                  )}
                </td>
              );
            })}
            {/* dopełnij wiersz do 7 kolumn */}
            {week.length < 7 && Array(7 - week.length).fill(null).map((_,i)=>(
              <td key={"pad"+i} style={{ border:"1px solid #E3DFD2", background:"#FAFAFA" }}/>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
}

// ============ MODAL DODATKÓW I POTRĄCEŃ ============
function AdjustmentsModal({ emp, date, initial, onClose, onSave, isAdmin }) {
  const [form, setForm] = useState(initial);
  // Surowy tekst dla pola przelewu – przechowuje "123,50" bez natychmiastowego parsowania
  const [transferRaw, setTransferRaw] = useState(
    initial.transferAmount ? String(initial.transferAmount).replace(".", ",") : ""
  );
  const set=(key)=>(e)=>setForm({...form,[key]:e.target.value});
  // Akceptuje zarówno kropkę jak i przecinek jako separator dziesiętny (Windows/PL)
  const parseAmt=(v)=>{ const s=String(v).replace(/\s/g,"").replace(",","."); return parseFloat(s)||0; };
  const setNum=(key)=>(e)=>setForm({...form,[key]:parseAmt(e.target.value)});
  // Zatwierdź surową kwotę przelewu przy opuszczeniu pola
  const commitTransfer=()=>{ const v=parseAmt(transferRaw); setForm(f=>({...f,transferAmount:v})); setTransferRaw(v?String(v).replace(".",","):""); };
  const increase=(Number(form.premia)||0)+(Number(form.ulga26)||0)+(Number(form.custIncAmount1)||0)+(Number(form.custIncAmount2)||0);
  const decrease=(Number(form.ubezpieczenie)||0)+(Number(form.custDecAmount1)||0)+(Number(form.custDecAmount2)||0);
  const isSplit=(emp.paymentMethod||"przelew")==="przelew_gotowka";
  return (
    <Overlay onClose={onClose}>
      <ModalHeader title="Dodatki i potrącenia" subtitle={`${displayName(emp.name)} · ${monthLabel(date)}`} onClose={onClose}/>
      <div style={{ padding:18 }}>
        <div style={{ fontSize:12.5, fontWeight:700, color:"#2F6B4F", marginBottom:8 }}>Zwiększające wynagrodzenie</div>
        <div style={{ display:"flex", gap:12 }}>
          <Field label="Premia (zł)" style={{ flex:1 }}><input type="text" inputMode="decimal" value={form.premia||""} onChange={set("premia")} onBlur={e=>setForm(f=>({...f,premia:parseAmt(e.target.value)}))} onFocus={e=>e.target.select()} placeholder="0" style={inputStyle}/></Field>
          <Field label="Ulga do 26 r.ż. (zł)" style={{ flex:1 }}><input type="text" inputMode="decimal" value={form.ulga26||""} onChange={set("ulga26")} onBlur={e=>setForm(f=>({...f,ulga26:parseAmt(e.target.value)}))} onFocus={e=>e.target.select()} placeholder="0" style={inputStyle}/></Field>
        </div>
        <div style={{ display:"flex", gap:12 }}>
          <Field label="Opis 1" style={{ flex:1 }}><input value={form.custIncLabel1} onChange={set("custIncLabel1")} placeholder="np. Dodatek świąteczny" style={inputStyle}/></Field>
          <Field label="Kwota (zł)" style={{ width:110 }}><input type="text" inputMode="decimal" value={form.custIncAmount1||""} onChange={set("custIncAmount1")} onBlur={e=>setForm(f=>({...f,custIncAmount1:parseAmt(e.target.value)}))} onFocus={e=>e.target.select()} placeholder="0" style={inputStyle}/></Field>
        </div>
        <div style={{ display:"flex", gap:12 }}>
          <Field label="Opis 2" style={{ flex:1 }}><input value={form.custIncLabel2} onChange={set("custIncLabel2")} placeholder="np. Premia frekwencyjna" style={inputStyle}/></Field>
          <Field label="Kwota (zł)" style={{ width:110 }}><input type="text" inputMode="decimal" value={form.custIncAmount2||""} onChange={set("custIncAmount2")} onBlur={e=>setForm(f=>({...f,custIncAmount2:parseAmt(e.target.value)}))} onFocus={e=>e.target.select()} placeholder="0" style={inputStyle}/></Field>
        </div>
        <div style={{ fontSize:12.5, fontWeight:700, color:"#C75146", margin:"16px 0 8px" }}>Pomniejszające wynagrodzenie</div>
        <Field label="Dodatkowe ubezpieczenie (zł)"><input type="text" inputMode="decimal" value={form.ubezpieczenie||""} onChange={set("ubezpieczenie")} onBlur={e=>setForm(f=>({...f,ubezpieczenie:parseAmt(e.target.value)}))} onFocus={e=>e.target.select()} placeholder="0" style={inputStyle}/></Field>
        <div style={{ display:"flex", gap:12 }}>
          <Field label="Opis 1" style={{ flex:1 }}><input value={form.custDecLabel1} onChange={set("custDecLabel1")} placeholder="np. Niedobór kasowy" style={inputStyle}/></Field>
          <Field label="Kwota (zł)" style={{ width:110 }}><input type="text" inputMode="decimal" value={form.custDecAmount1||""} onChange={set("custDecAmount1")} onBlur={e=>setForm(f=>({...f,custDecAmount1:parseAmt(e.target.value)}))} onFocus={e=>e.target.select()} placeholder="0" style={inputStyle}/></Field>
        </div>
        <div style={{ display:"flex", gap:12 }}>
          <Field label="Opis 2" style={{ flex:1 }}><input value={form.custDecLabel2} onChange={set("custDecLabel2")} placeholder="np. Zaliczka" style={inputStyle}/></Field>
          <Field label="Kwota (zł)" style={{ width:110 }}><input type="text" inputMode="decimal" value={form.custDecAmount2||""} onChange={set("custDecAmount2")} onBlur={e=>setForm(f=>({...f,custDecAmount2:parseAmt(e.target.value)}))} onFocus={e=>e.target.select()} placeholder="0" style={inputStyle}/></Field>
        </div>
        {isSplit && (
          <>
            <div style={{ fontSize:12.5, fontWeight:700, color:"#1F4D38", margin:"16px 0 8px" }}>Podział wypłaty -- przelew / gotówka</div>
            <Field label="Kwota na przelew w tym miesiącu (zł)">
              <input
                type="text" inputMode="decimal"
                value={transferRaw}
                placeholder="np. 2000,00"
                onChange={e=>setTransferRaw(e.target.value.replace(/[^0-9.,]/g,""))}
                onBlur={commitTransfer}
                onFocus={e=>e.target.select()}
                style={inputStyle}/>
            </Field>
          </>
        )}
        {isAdmin && (
          <div style={{ fontSize:12, color:"#3a392f", background:"#FAF8F2", borderRadius:8, padding:"8px 10px", marginTop:10 }}>
            Stawka: <b>{emp.rate.toFixed(2)} zl/h</b> · Dodatki: <b style={{ color:"#2F6B4F" }}>+{increase.toFixed(2)} zl</b> · Potracenia: <b style={{ color:"#C75146" }}>-{decrease.toFixed(2)} zl</b>
          </div>
        )}
        {!isAdmin && (
          <div style={{ fontSize:12, color:"#9C988A", background:"#FAF8F2", borderRadius:8, padding:"8px 10px", marginTop:10 }}>
            <span style={{ display:"flex", alignItems:"center", gap:6 }}><EyeOff size={12}/> Stawka i kwoty wynagrodzenia widoczne tylko dla administratora.</span>
          </div>
        )}
      </div>
      <div style={{ padding:"12px 18px", borderTop:"1px solid #EFEBDE", display:"flex", justifyContent:"flex-end", gap:8 }}>
        <button onClick={onClose} style={ghostBtn}>Anuluj</button>
        <button onClick={()=>onSave(form)} style={primaryBtn}><Check size={14}/> Zapisz</button>
      </div>
    </Overlay>
  );
}

// ============ USTAWIENIA ============
function SettingsView({ settings, setSettings, positions, addPosition, removePosition, isAdmin }) {
  const [l4, setL4] = useState(settings.l4Percent);
  const [opiekaP, setOpiekaP] = useState(settings.opiekaChorePercent||80);
  const [refHours, setRefHours] = useState(settings.refHours);
  const [storeName, setStoreName] = useState(settings.storeName||"Mój Sklep");
  const dirty=l4!==settings.l4Percent||refHours!==settings.refHours||opiekaP!==(settings.opiekaChorePercent||80)||storeName!==(settings.storeName||"Mój Sklep");
  const [newAdminPass, setNewAdminPass] = useState(""), [newAdminPass2, setNewAdminPass2] = useState("");
  const [newManagerPass, setNewManagerPass] = useState(""), [newManagerPass2, setNewManagerPass2] = useState("");
  const [passMsg, setPassMsg] = useState("");
  const savePasswords = () => {
    const updates = {};
    if (newAdminPass && newAdminPass===newAdminPass2) updates.adminPasswordHash=encodePass(newAdminPass);
    else if (newAdminPass && newAdminPass!==newAdminPass2) { setPassMsg("Hasła administratora nie są zgodne."); return; }
    if (newManagerPass && newManagerPass===newManagerPass2) updates.managerPasswordHash=encodePass(newManagerPass);
    else if (newManagerPass && newManagerPass!==newManagerPass2) { setPassMsg("Hasła kierownika nie są zgodne."); return; }
    if (!Object.keys(updates).length) { setPassMsg("Nie wpisano żadnych haseł."); return; }
    setSettings({...settings,...updates});
    setNewAdminPass(""); setNewAdminPass2(""); setNewManagerPass(""); setNewManagerPass2("");
    setPassMsg("✓ Hasła zostały zmienione."); setTimeout(()=>setPassMsg(""),3000);
  };
  const [newPosition, setNewPosition] = useState("");
  return (
    <div style={{ maxWidth:540, display:"flex", flexDirection:"column", gap:20 }}>
      {isAdmin && <div>
        <div style={{ fontWeight:700, fontSize:12, color:"#8890A0", textTransform:"uppercase", letterSpacing:"0.06em", marginBottom:10 }}>Sklep</div>
        <div style={{ background:"#fff", border:"1px solid #E4E7EE", borderRadius:12, padding:20, boxShadow:"0 1px 4px rgba(0,0,0,0.04)" }}>
          <Field label="Nazwa sklepu (widoczna w nagłówku i na wydrukach)">
            <input value={storeName} onChange={e=>setStoreName(e.target.value)} style={inputStyle}/>
          </Field>
          <div style={{ fontSize:11.5, color:"#9399A6" }}>Zmiana wymaga zapisania ustawień poniżej.</div>
        </div>
      </div>}
      {isAdmin && <div>
        <div style={{ fontWeight:700, fontSize:12, color:"#8890A0", textTransform:"uppercase", letterSpacing:"0.06em", marginBottom:10 }}>Hasła dostępu</div>
        <div style={{ background:"#fff", border:"1px solid #E4E7EE", borderRadius:12, padding:20, boxShadow:"0 1px 4px rgba(0,0,0,0.04)" }}>
          <div style={{ fontSize:12.5, color:"#8890A0", marginBottom:14, background:"#F7F9FC", padding:"8px 12px", borderRadius:7 }}>Pozostaw puste, jeśli nie chcesz zmieniać. Domyślne: Administrator → <b>admin</b> · Kierownik → <b>sklep</b></div>
          <div style={{ fontSize:12, fontWeight:700, color:"#1F6B45", marginBottom:8 }}>Nowe hasło administratora</div>
          <div style={{ display:"flex", gap:8, marginBottom:14 }}>
            <input type="password" value={newAdminPass} onChange={e=>setNewAdminPass(e.target.value)} placeholder="Nowe hasło" style={{ ...inputStyle, flex:1 }}/>
            <input type="password" value={newAdminPass2} onChange={e=>setNewAdminPass2(e.target.value)} placeholder="Powtórz hasło" style={{ ...inputStyle, flex:1 }}/>
          </div>
          <div style={{ fontSize:12, fontWeight:700, color:"#3B6FA0", marginBottom:8 }}>Nowe hasło kierownika</div>
          <div style={{ display:"flex", gap:8, marginBottom:14 }}>
            <input type="password" value={newManagerPass} onChange={e=>setNewManagerPass(e.target.value)} placeholder="Nowe hasło" style={{ ...inputStyle, flex:1 }}/>
            <input type="password" value={newManagerPass2} onChange={e=>setNewManagerPass2(e.target.value)} placeholder="Powtórz hasło" style={{ ...inputStyle, flex:1 }}/>
          </div>
          {passMsg && <div style={{ fontSize:12.5, color:passMsg.startsWith("✓")?"#1F6B45":"#C23B3B", marginBottom:10, fontWeight:500 }}>{passMsg}</div>}
          <button onClick={savePasswords} style={primaryBtn}><Lock size={14}/> Zmień hasła</button>
        </div>
      </div>}
      <div>
        <div style={{ fontWeight:700, fontSize:12, color:"#8890A0", textTransform:"uppercase", letterSpacing:"0.06em", marginBottom:10 }}>Stanowiska</div>
        <div style={{ background:"#fff", border:"1px solid #E4E7EE", borderRadius:12, padding:20, boxShadow:"0 1px 4px rgba(0,0,0,0.04)" }}>
          <div style={{ display:"flex", flexDirection:"column", gap:6, marginBottom:12 }}>
            {positions.map(p=>(
              <div key={p} style={{ display:"flex", alignItems:"center", justifyContent:"space-between", background:"#F7F9FC", padding:"7px 11px", borderRadius:7, fontSize:13, border:"1px solid #EEF0F5" }}>
                <span style={{ color:"#1A2530", fontWeight:500 }}>{p}</span>
                <button onClick={()=>removePosition(p)} style={{ border:"none", background:"transparent", cursor:"pointer", color:"#C23B3B", padding:3, borderRadius:5 }}><X size={14}/></button>
              </div>
            ))}
            {positions.length===0 && <div style={{ fontSize:12, color:"#9399A6" }}>Brak stanowisk.</div>}
          </div>
          <div style={{ display:"flex", gap:6 }}>
            <input value={newPosition} onChange={e=>setNewPosition(e.target.value)} placeholder="Np. Kierownik zmiany" style={{ ...inputStyle, flex:1 }}/>
            <button onClick={()=>{ addPosition(newPosition); setNewPosition(""); }} style={{ ...smallBtn, height:38, padding:"0 14px", borderRadius:8, fontWeight:700, color:"#1F6B45", borderColor:"#B7D9C4" }}><Plus size={14}/></button>
          </div>
        </div>
      </div>
      {isAdmin && <div>
        <div style={{ fontWeight:700, fontSize:12, color:"#8890A0", textTransform:"uppercase", letterSpacing:"0.06em", marginBottom:10 }}>Ustawienia rozliczeń</div>
        <div style={{ background:"#fff", border:"1px solid #E4E7EE", borderRadius:12, padding:20, boxShadow:"0 1px 4px rgba(0,0,0,0.04)" }}>
          <Field label="Stawka za L4 — % stawki podstawowej">
            <div style={{ display:"flex", alignItems:"center", gap:8 }}>
              <input type="number" min="0" max="100" step="1" value={l4} onChange={e=>setL4(Number(e.target.value))} onFocus={e=>e.target.select()} style={{ ...inputStyle, width:100 }}/>
              <span style={{ fontSize:13, color:"#8890A0" }}>%</span>
            </div>
          </Field>
          <Field label="Stawka za opiekę nad chorym dzieckiem/rodziną — % stawki">
            <div style={{ display:"flex", alignItems:"center", gap:8 }}>
              <input type="number" min="0" max="100" step="1" value={opiekaP} onChange={e=>setOpiekaP(Number(e.target.value))} onFocus={e=>e.target.select()} style={{ ...inputStyle, width:100 }}/>
              <span style={{ fontSize:13, color:"#8890A0" }}>% (standardowo 80%)</span>
            </div>
          </Field>
          <Field label="Referencyjne godziny dnia roboczego (do wyceny L4, urlopu, opieki)">
            <div style={{ display:"flex", alignItems:"center", gap:8 }}>
              <input type="number" min="1" max="12" step="0.5" value={refHours} onChange={e=>setRefHours(Number(e.target.value))} onFocus={e=>e.target.select()} style={{ ...inputStyle, width:100 }}/>
              <span style={{ fontSize:13, color:"#8890A0" }}>godz.</span>
            </div>
          </Field>
          <div style={{ fontSize:11.5, color:"#9399A6", marginBottom:14, background:"#F7F9FC", padding:"8px 12px", borderRadius:7 }}>
            Przykład: stawka 30 zł/godz, L4 = {l4}%, {refHours} godz./dzień → dzień L4 = {(30*refHours*l4/100).toFixed(2)} zł
          </div>
          <button disabled={!dirty} onClick={()=>setSettings({...settings,l4Percent:l4,refHours,opiekaChorePercent:opiekaP,storeName})} style={{ ...primaryBtn, opacity:dirty?1:0.5, cursor:dirty?"pointer":"default" }}><Check size={14}/> Zapisz ustawienia</button>
        </div>
      </div>}
    </div>
  );
}


function StatCard({ label, value, accent }) {
  return (
    <div style={{ background:accent?"#1F6B45":"#fff", color:accent?"#fff":"#1A2530", border:accent?"none":"1px solid #E4E7EE", borderRadius:12, padding:"16px 18px", boxShadow:accent?"0 2px 12px rgba(31,107,69,0.25)":"0 1px 4px rgba(0,0,0,0.05)" }}>
      <div style={{ fontSize:11, opacity:accent?0.8:1, fontWeight:700, textTransform:"uppercase", letterSpacing:"0.06em", color:accent?"rgba(255,255,255,0.75)":"#8890A0" }}>{label}</div>
      <div style={{ fontSize:24, fontWeight:700, marginTop:6, letterSpacing:"-0.01em" }}>{value}</div>
    </div>
  );
}

function Overlay({ children, onClose }) {
  return (
    <div className="no-print" onClick={onClose} style={{ position:"fixed", inset:0, background:"rgba(10,20,14,0.5)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:50, padding:16, backdropFilter:"blur(2px)" }}>
      <div onClick={e=>e.stopPropagation()} style={{ background:"#fff", borderRadius:16, width:440, maxWidth:"100%", maxHeight:"90vh", overflowY:"auto", boxShadow:"0 16px 48px rgba(0,0,0,0.22)", border:"1px solid #E4E7EE" }}>{children}</div>
    </div>
  );
}
function ModalHeader({ title, subtitle, onClose }) {
  return (
    <div style={{ padding:"18px 20px", borderBottom:"1px solid #EEF0F5", display:"flex", justifyContent:"space-between", alignItems:"flex-start" }}>
      <div>
        <div style={{ fontWeight:700, fontSize:15.5, color:"#1A2530" }}>{title}</div>
        {subtitle && <div style={{ fontSize:12, color:"#9399A6", marginTop:3 }}>{subtitle}</div>}
      </div>
      <button onClick={onClose} style={{ border:"none", background:"#F7F9FC", cursor:"pointer", color:"#8890A0", borderRadius:8, width:32, height:32, display:"flex", alignItems:"center", justifyContent:"center" }}><X size={16}/></button>
    </div>
  );
}
function Field({ label, children, style }) {
  return <div style={{ marginBottom:14, ...style }}><div style={{ fontSize:11.5, fontWeight:700, color:"#5A6070", marginBottom:5, textTransform:"uppercase", letterSpacing:"0.04em" }}>{label}</div>{children}</div>;
}
const inputStyle={ width:"100%", padding:"8px 11px", borderRadius:8, border:"1px solid #DDE0E8", fontSize:13, fontFamily:"inherit", boxSizing:"border-box", background:"#fff", color:"#1A2530", outline:"none" };
const navBtn={ width:32, height:32, borderRadius:8, border:"1px solid #DDE0E8", background:"#fff", display:"flex", alignItems:"center", justifyContent:"center", cursor:"pointer", color:"#5A6070" };
const primaryBtn={ display:"flex", alignItems:"center", gap:6, padding:"8px 16px", borderRadius:8, border:"none", background:"#1F6B45", color:"#fff", fontWeight:600, fontSize:13, cursor:"pointer", letterSpacing:"0.01em" };
const ghostBtn={ display:"flex", alignItems:"center", gap:6, padding:"8px 14px", borderRadius:8, border:"1px solid #DDE0E8", background:"#fff", color:"#3D4452", fontWeight:600, fontSize:13, cursor:"pointer" };
const dangerBtn={ display:"flex", alignItems:"center", gap:6, padding:"8px 14px", borderRadius:8, border:"1px solid #FACACA", background:"#FEF2F2", color:"#C23B3B", fontWeight:600, fontSize:13, cursor:"pointer" };
const smallBtn={ width:34, borderRadius:8, border:"1px solid #DDE0E8", background:"#fff", cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center", color:"#5A6070" };
const iconBtn={ width:30, height:30, borderRadius:7, border:"1px solid #DDE0E8", background:"#fff", cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center", color:"#5A6070", boxShadow:"0 1px 2px rgba(0,0,0,0.04)" };
