/* ============================================================
   VERPA SUPPORT DESK — app.js
   ============================================================ */

/* ============================================================================
   ██  CONFIG — VUL DEZE WAARDEN IN  ██
   ============================================================================ */
const CONFIG = {
  clientId:    "e82e1484-0864-44a8-a8fe-279915eec8bf",   // Verpa Support Desk app-registratie
  tenantId:    "e65dbe4b-d1e2-4283-b0f5-aa7717e81077",   // Verpa Benelux tenant
  redirectUri: "https://verpa-support.pages.dev",        // je Cloudflare Pages-adres (zie handleiding stap 3)
  siteHostname:"verpabenelux.sharepoint.com",            // jouw SharePoint hostname
  sitePath:    "/sites/OfficeData",                      // pad naar jouw SharePoint-site
  listName:    "Tickets",                                // naam van de SharePoint List
  attachFolder:"Tickets",                                // map in de documentbibliotheek voor bijlagen
  adminRole:   "Admin",                                  // naam van de Azure AD App Role voor beheerders
  userRole:    "User",                                   // naam van de Azure AD App Role voor gewone gebruikers
  adminEmails: ["lukas@verpa.be","sten.huygens@verpa.be","aniel@verpa.be"], // e-mailadressen van beheerders
  mailWorker:  "https://verpa-mail-proxy.lukas-f22.workers.dev", // Cloudflare Worker voor mailverzending
  settingsFile:"Tickets/_instellingen/instellingen.json"   // gedeeld instellingenbestand in de documentbibliotheek
};
/* ============================================================================ */

/* Behandelaars (TEAM) en beheerders-mailadressen worden nu beheerd via Instellingen.
   De standaardwaarden staan in DEFAULT_SETTINGS verderop. */
const ACCOUNT_TYPES = ["Account manager", "Subaccount", "Standaard account"];
const ASSORTMENTS = ["Algemeen assortiment", "Afgeschermd assortiment"];
const WEBSHOP_INFO = `<div class="infobox"><div class="ih"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><path d="M12 16v-4M12 8h.01"/></svg>Webshop account hiërarchie</div><p>Een <b>Account manager</b> is een aankoper die bestellingen van medewerkers (subaccounts) moet goedkeuren. <b>Subaccounts</b> zijn medewerkers die bestellen onder goedkeuring van hun account manager. Een <b>Standaard account</b> is een zelfstandige klant zonder hiërarchie.</p></div>`;

const FORMS = {
  Verkoop: {
    "Prijswijzigingen": { fields:[{k:"onderwerp",label:"Onderwerp",type:"text",req:true,isSubject:true},{k:"omschrijving",label:"Omschrijving vraag",type:"textarea",req:true}]},
    "Afnamerapporten": { title:r=>`Afnamerapport – ${r.klant||"?"}`, fields:[
      {k:"klant",label:"Klant",type:"text",req:true},{k:"klantnummer",label:"Klantnummer",type:"text"},
      {k:"periode",label:"Periode",type:"daterange",req:true},{k:"niveau",label:"Op facturatieniveau of leveradresniveau?",type:"select",options:["Facturatieniveau","Leveradresniveau"],req:true},
      {k:"omschrijving",label:"Omschrijving vraag",type:"textarea"}]},
    "Artikelen aanmaken": { title:r=>`Nieuw artikel – ${r.artikelnaam||"?"}`, fields:[
      {k:"artikelnummer",label:"Artikelnummer",hint:"indien beschikbaar",type:"text"},{k:"artikelnaam",label:"Artikelnaam",type:"text",req:true},
      {k:"leverancier",label:"Leverancier",type:"text",req:true},{k:"inkoopprijs",label:"Inkoopprijs",hint:"indien beschikbaar",type:"text"},
      {k:"verkoopprijs",label:"Verkoopprijs",type:"text",req:true},{k:"opwebshop",label:"Op webshop?",type:"yesno",req:true},
      {k:"migratie",label:"Is het een migratieartikel?",type:"yesno",req:true},{k:"vervangt",label:"Welk artikel wordt vervangen?",type:"text",req:true,showIf:{k:"migratie",val:"Ja"}}]},
    "Klant assortiment": { title:r=>`Klantassortiment – ${r.klantnaam||"?"}`, fields:[
      {k:"klantnaam",label:"Klantnaam",type:"text",req:true},{k:"klantnummer",label:"Klantnummer",type:"text"},
      {k:"klantgroep",label:"Klantgroep",type:"text"},{k:"artikel",label:"Artikel",type:"text",req:true},
      {k:"vervanging",label:"Vervanging voor iets anders?",type:"yesno",req:true},{k:"vervangt",label:"Welk artikel wordt vervangen?",type:"text",req:true,showIf:{k:"vervanging",val:"Ja"}}]},
    "Andere vragen": { fields:[{k:"onderwerp",label:"Onderwerp",type:"text",req:true,isSubject:true},{k:"omschrijving",label:"Omschrijving vraag",type:"textarea",req:true}]},
  },
  Technisch: {
    "Webshop": { fields:[{k:"onderwerp",label:"Onderwerp",type:"text",req:true,isSubject:true},{k:"klantnaam",label:"Klantnaam",type:"text"},{k:"klantnummer",label:"Klantnummer",type:"text"},{k:"email",label:"Email",type:"text"},{k:"omschrijving",label:"Omschrijving vraag",type:"textarea",req:true}]},
    "Webshop login": { infobox:WEBSHOP_INFO, title:r=>`Webshop login – ${r.klantnaam||r.gebruikersnaam||"nieuw account"}`, fields:[
      {k:"gebruikersnaam",label:"Gebruikersnaam webshop account",type:"text",req:true},{k:"email",label:"E-mailadres",type:"text",req:true},
      {k:"klantnummer",label:"Klantennummer",type:"text",req:true},{k:"klantnaam",label:"Klantnaam",type:"text",req:true},
      {k:"assortiment",label:"Afgeschermd assortiment",type:"select",options:ASSORTMENTS,req:true},{k:"accounttype",label:"Type account",type:"select",options:ACCOUNT_TYPES,req:true}]},
    "IT-Probleem": { fields:[{k:"onderwerp",label:"Onderwerp",type:"text",req:true,isSubject:true},{k:"omschrijving",label:"Omschrijving vraag",type:"textarea",req:true}]},
    "Business Central": { fields:[{k:"onderwerp",label:"Onderwerp",type:"text",req:true,isSubject:true},{k:"omschrijving",label:"Omschrijving vraag",type:"textarea",req:true}]},
    "Andere vragen": { fields:[{k:"onderwerp",label:"Onderwerp",type:"text",req:true,isSubject:true},{k:"omschrijving",label:"Omschrijving vraag",type:"textarea",req:true}]},
  }
};
const ALL_SUBS = [...new Set([...Object.keys(FORMS.Verkoop), ...Object.keys(FORMS.Technisch)])];
const CAT = { Verkoop:{label:"Verkoop",color:"var(--sales)",bg:"var(--sales-bg)"}, Technisch:{label:"Technisch",color:"var(--tech)",bg:"var(--tech-bg)"} };
const STATUS = { open:{label:"Open",color:"var(--st-open)",bg:"var(--st-open-bg)"}, progress:{label:"In Behandeling",color:"var(--st-prog)",bg:"var(--st-prog-bg)"}, done:{label:"Opgelost",color:"var(--st-done)",bg:"var(--st-done-bg)"}, closed:{label:"Gesloten",color:"var(--st-closed)",bg:"var(--st-closed-bg)"} };
const PRIO = { high:{label:"Hoog",color:"var(--p-high)",bg:"var(--p-high-bg)"}, mid:{label:"Gemiddeld",color:"var(--p-mid)",bg:"var(--p-mid-bg)"}, low:{label:"Laag",color:"var(--p-low)",bg:"var(--p-low-bg)"} };
const GRAPH="https://graph.microsoft.com/v1.0";
const SCOPES=["User.Read","Sites.ReadWrite.All","Files.ReadWrite.All"];

/* ===================== STATE ===================== */
let msalInstance=null, account=null, currentUser=null;
let adminEmails=[]; // e-mailadressen van beheerders uit CONFIG
let SITE_ID=null, LIST_ID=null, DRIVE_ID=null, COL={};
let tickets=[], view="dashboard", currentId=null, detailTicket=null;
let newCat=null, curSchema=null, newFiles=[], replyFiles=[], replyInternal=false, saving=false;
let filter={ q:"",category:"",subcategory:"",status:"",assignee:"",special:"" };

/* ===================== HELPERS ===================== */
const esc=s=>(s||"").replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const fmtSize=b=>b<1024?b+" B":b<1048576?(b/1024).toFixed(0)+" KB":(b/1048576).toFixed(1)+" MB";
const initials=n=>(n||"?").trim().split(/\s+/).slice(0,2).map(w=>w[0]).join("").toUpperCase();
const firstName=n=>(!n||n==="Niet toegewezen")?"—":n.split(" ")[0];
const val=id=>{ const e=document.getElementById(id); return e?e.value.trim():""; };
function fmtDate(ts){ return new Date(ts).toLocaleString("nl-BE",{day:"2-digit",month:"short",year:"numeric",hour:"2-digit",minute:"2-digit"}); }
function fmtD(iso){ if(!iso)return""; const p=iso.split("-"); return p.length===3?`${p[2]}/${p[1]}/${p[0]}`:iso; }
function ago(ts){ const s=(Date.now()-ts)/1000; if(s<60)return"zojuist"; const m=s/60; if(m<60)return Math.floor(m)+" min geleden"; const h=m/60; if(h<24)return Math.floor(h)+" uur geleden"; const d=h/24; if(d<30)return Math.floor(d)+(Math.floor(d)===1?" dag geleden":" dagen geleden"); return new Date(ts).toLocaleDateString("nl-BE",{day:"2-digit",month:"short"}); }
function toast(m){ const t=document.getElementById("toast"); t.textContent=m; t.classList.add("show"); clearTimeout(t._t); t._t=setTimeout(()=>t.classList.remove("show"),2600); }
const isAdmin=()=>currentUser&&currentUser.isAdmin;
const isUser=()=>currentUser&&currentUser.isUser;
const parseJson=(s,fb)=>{ try{ return s?JSON.parse(s):fb; }catch{ return fb; } };
const clone=o=>JSON.parse(JSON.stringify(o));
const jsq=s=>esc(String(s||"")).replace(/\\/g,"\\\\").replace(/'/g,"\\'");
const EMAIL_RE=/^[^\s@]+@[^\s@]+\.[^\s@]+$/;
function htmlToText(html){ if(!html) return ""; try{ return (new DOMParser().parseFromString(html,"text/html").body.textContent||"").replace(/\s+/g," ").trim(); }catch{ return html.replace(/<[^>]+>/g," "); } }
function previewOf(t){ return (t.description||(t.fields&&t.fields[0]?t.fields[0].value:"")).split("\n")[0].slice(0,90); }
function searchText(t){ return [t.subject,t.ref,t.author,...(t.fields||[]).map(f=>f.value),t.description].join(" ").toLowerCase(); }

/* ===================== MICROSOFT GRAPH ===================== */
async function getToken(){
  try{ const r=await msalInstance.acquireTokenSilent({scopes:SCOPES,account}); return r.accessToken; }
  catch(e){ const r=await msalInstance.acquireTokenPopup({scopes:SCOPES}); return r.accessToken; }
}
async function graph(path, opts={}, raw=false){
  const token=await getToken();
  const headers={ Authorization:`Bearer ${token}`, ...(opts.headers||{}) };
  if(!raw && !(opts.body instanceof Blob)) headers["Content-Type"]="application/json";
  const res=await fetch(path.startsWith("http")?path:GRAPH+path, {...opts, headers});
  if(!res.ok){ const txt=await res.text(); throw new Error(`Graph ${res.status}: ${txt.slice(0,300)}`); }
  if(res.status===204) return null;
  return raw?res:res.json();
}
async function resolveIds(){
  const site=await graph(`/sites/${CONFIG.siteHostname}:${CONFIG.sitePath}`);
  SITE_ID=site.id;
  const drive=await graph(`/sites/${SITE_ID}/drive`); DRIVE_ID=drive.id;
  const lists=await graph(`/sites/${SITE_ID}/lists?$select=id,displayName&$top=200`);
  const list=lists.value.find(l=>l.displayName===CONFIG.listName);
  if(!list) throw new Error(`SharePoint List "${CONFIG.listName}" niet gevonden op de site.`);
  LIST_ID=list.id;
  await resolveColumns();
}
const NEEDED=["Ref","Category","Subcategory","Status","Priority","Assignee","Indiener","OwnerUpn","Description","FieldsJson","MessagesJson","Archived"];
async function resolveColumns(){
  const d=await graph(`/sites/${SITE_ID}/lists/${LIST_ID}/columns?$select=name,displayName&$top=250`);
  const byName={}, byDisplay={};
  d.value.forEach(c=>{ if(c.name) byName[c.name.toLowerCase()]=c.name; if(c.displayName) byDisplay[c.displayName.toLowerCase()]=c.name; });
  COL={ Title:"Title" };
  const missing=[];
  NEEDED.forEach(n=>{ const k=n.toLowerCase(); const internal=byName[k]||byDisplay[k]; if(internal) COL[n]=internal; else missing.push(n); });
  if(missing.length) throw new Error("Ontbrekende kolommen in de lijst '"+CONFIG.listName+"': "+missing.join(", ")+". Maak deze aan als 'Eén regel tekst' (behalve Description, FieldsJson en MessagesJson = 'Meerdere regels tekst', en Archived = 'Ja/Nee').");
  COL.Followers = byName["followers"] || byDisplay["followers"] || null;
}
function itemToTicket(item){
  const f=item.fields||{}; const g=k=>f[COL[k]];
  return { itemId:item.id, ref:g("Ref")||("#"+item.id), category:g("Category")||"Verkoop", subcategory:g("Subcategory")||"",
    subject:f.Title||"(geen onderwerp)", author:g("Indiener")||"", ownerUpn:(g("OwnerUpn")||"").toLowerCase(),
    assignee:g("Assignee")||"Niet toegewezen", priority:g("Priority")||"mid", status:g("Status")||"open",
    description:g("Description")||"", fields:parseJson(g("FieldsJson"),[]), messages:parseJson(g("MessagesJson"),[]),
    archived:!!g("Archived"), followers:(COL.Followers?parseJson(f[COL.Followers],[]):[]), createdAt:new Date(item.createdDateTime||Date.now()).getTime() };
}
async function loadTickets(){
  let url=`/sites/${SITE_ID}/lists/${LIST_ID}/items?expand=fields&$top=200`; const all=[];
  while(url){ const d=await graph(url); all.push(...d.value); url=d["@odata.nextLink"]||null; }
  tickets=all.map(itemToTicket);
}
async function createTicketItem(fieldsObj){
  const d=await graph(`/sites/${SITE_ID}/lists/${LIST_ID}/items`,{method:"POST",body:JSON.stringify({fields:fieldsObj})});
  const full=await graph(`/sites/${SITE_ID}/lists/${LIST_ID}/items/${d.id}?expand=fields`);
  return itemToTicket(full);
}
async function patchTicket(itemId, fieldsObj){ await graph(`/sites/${SITE_ID}/lists/${LIST_ID}/items/${itemId}/fields`,{method:"PATCH",body:JSON.stringify(fieldsObj)}); }
async function deleteTicketItem(itemId){ await graph(`/sites/${SITE_ID}/lists/${LIST_ID}/items/${itemId}`,{method:"DELETE"}); }
async function uploadFile(ref, file){
  const path=`/sites/${SITE_ID}/drive/root:/${encodeURIComponent(CONFIG.attachFolder)}/${encodeURIComponent(ref)}/${encodeURIComponent(file.name)}:/content`;
  const item=await graph(path,{method:"PUT",body:file});
  return { name:file.name, size:file.size, itemId:item.id, webUrl:item.webUrl };
}
async function attachmentDownloadUrl(itemId){ const d=await graph(`/sites/${SITE_ID}/drive/items/${itemId}?$select=@microsoft.graph.downloadUrl`); return d["@microsoft.graph.downloadUrl"]; }

/* ===================== BOOT / AUTH ===================== */
function loadScript(src){ return new Promise((res,rej)=>{ const s=document.createElement("script"); s.src=src; s.onload=res; s.onerror=()=>rej(new Error(src)); document.head.appendChild(s); }); }
async function ensureMsal(){
  if(window.msal) return;
  const sources=[
    "https://alcdn.msauth.net/browser/2.38.4/js/msal-browser.min.js",
    "https://cdn.jsdelivr.net/npm/@azure/msal-browser@2.38.4/lib/msal-browser.min.js",
    "https://unpkg.com/@azure/msal-browser@2.38.4/lib/msal-browser.min.js",
    "https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.10.0/lib/msal-browser.min.js"
  ];
  for(const url of sources){ try{ await loadScript(url); if(window.msal) return; }catch(e){ /* volgende bron proberen */ } }
  throw new Error("De Microsoft-loginbibliotheek (MSAL) kon niet geladen worden.");
}
async function boot(){
  // Sla ticket hash op vóór MSAL redirect (hash gaat verloren tijdens login redirect)
  const hash=window.location.hash;
  if(hash&&hash.startsWith("#ticket-")){ sessionStorage.setItem("openTicketId", hash.replace("#ticket-","")); }
  console.log("Verpa Support Desk — build: v6 met beheerdersinstellingen");
  if(CONFIG.clientId.startsWith("PLAK_HIER")){ return renderConfigError(); }
  try{
    await ensureMsal();
    msalInstance=new msal.PublicClientApplication({
      auth:{ clientId:CONFIG.clientId, authority:`https://login.microsoftonline.com/${CONFIG.tenantId}`, redirectUri:CONFIG.redirectUri },
      cache:{ cacheLocation:"localStorage", storeAuthStateInCookie:false }
    });
    await msalInstance.initialize();
    const resp=await msalInstance.handleRedirectPromise();
    if(resp&&resp.account) account=resp.account;
    if(!account) account=msalInstance.getActiveAccount()||msalInstance.getAllAccounts()[0]||null;
    if(!account){ return renderLogin(); }
    msalInstance.setActiveAccount(account);
    await afterLogin();
  }catch(e){ renderFatal(e.message); }
}
function logoImg(size,radius){ return `<img src="logo.jpg" alt="Verpa" style="width:${size}px;height:${size}px;border-radius:${radius}px;object-fit:cover" onerror="this.style.display='none';this.nextElementSibling.style.display='grid'">`; }
function renderLogin(){
  document.getElementById("root").innerHTML=`
    <div class="auth-wrap"><div class="auth-card">
      <div style="width:72px;height:72px;margin:0 auto 16px;position:relative">
        ${logoImg(72,14)}
        <div style="display:none;width:72px;height:72px;border-radius:14px;background:linear-gradient(135deg,var(--primary),var(--primary-dark));place-items:center;color:#fff;font-weight:800;font-size:28px;box-shadow:0 4px 14px rgba(13,139,128,.32)">V</div>
      </div>
      <h1>Verpa Support</h1><p>Meld je aan met je Verpa Microsoft-account om verder te gaan.</p>
      <button class="ms-btn" onclick="signIn()">
        <svg width="18" height="18" viewBox="0 0 23 23"><path fill="#f25022" d="M1 1h10v10H1z"/><path fill="#7fba00" d="M12 1h10v10H12z"/><path fill="#00a4ef" d="M1 12h10v10H1z"/><path fill="#ffb900" d="M12 12h10v10H12z"/></svg>
        Aanmelden met Microsoft</button>
    </div></div>`;
}
async function signIn(){ try{ await msalInstance.loginRedirect({scopes:SCOPES}); }catch(e){ renderFatal(e.message); } }
async function afterLogin(){
  const claims=account.idTokenClaims||{};
  const roles=claims.roles||[];
  currentUser={ name:account.name||claims.name||account.username, upn:(account.username||"").toLowerCase(), isAdmin:roles.includes(CONFIG.adminRole), isUser:roles.includes(CONFIG.userRole) };
  document.getElementById("root").innerHTML=`<div class="auth-wrap"><div class="auth-card"><div class="spinner"></div><p>Verbinden met SharePoint…</p></div></div>`;
  try{ await resolveIds(); await loadSettings(); await loadTickets(); showApp();
    // Directe link vanuit mail: herstel opgeslagen ticket na MSAL redirect
    const savedId=sessionStorage.getItem("openTicketId");
    if(savedId){ sessionStorage.removeItem("openTicketId"); openDetail(savedId); }
  }
  catch(e){ renderFatal(e.message); }
}
function logout(){ msalInstance.logoutRedirect(); }

function renderConfigError(){
  document.getElementById("root").innerHTML=`<div class="auth-wrap"><div class="auth-card" style="text-align:left">
    <div class="auth-logo-fallback" style="margin:0 0 14px">V</div><h1>Configuratie vereist</h1>
    <p>Vul bovenaan het bestand het <code>CONFIG</code>-blok in met je Azure AD Client ID, Tenant ID, SharePoint-site en lijstnaam.</p>
    <div class="cfg-err">Nog niet ingevuld: <code>clientId</code>.</div>
  </div></div>`;
}
function renderFatal(msg){
  document.getElementById("root").innerHTML=`<div class="auth-wrap"><div class="auth-card" style="text-align:left">
    <div class="auth-logo-fallback" style="margin:0 0 14px;background:#dc2626">!</div><h1>Er ging iets mis</h1>
    <p>De app kon geen verbinding maken. Controleer je configuratie en rechten.</p>
    <div class="cfg-err">${esc(msg)}</div>
    <button class="ms-btn" onclick="location.reload()">Opnieuw proberen</button>
  </div></div>`;
}

/* ===================== SHELL ===================== */
function showApp(){
  document.getElementById("root").innerHTML=`
    <div class="mobilebar">
      <button class="hamb" onclick="toggleSidebar()"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M3 12h18M3 18h18"/></svg></button>
      <div class="mb-brand">
        <div class="mb-logo">${logoImg(28,7)}</div>
        <span>Verpa Support</span>
      </div>
      <button class="hamb" onclick="openNew()"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg></button>
    </div>
    <div class="sb-scrim" id="sbScrim" onclick="toggleSidebar(false)"></div>
    <div class="shell">
      <aside class="sidebar" id="sidebar">
        <div class="sb-brand"><div class="row">
          <div class="sb-logo">${logoImg(38,10)}</div>
          <div><h1>Verpa Support</h1><p>Ticketbeheer · v6</p></div>
        </div></div>
        <nav class="sb-nav">
          <div class="nav-item" data-nav="dashboard" onclick="go('dashboard')"><svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="3" width="7" height="7" rx="1.5"/><rect x="3" y="14" width="7" height="7" rx="1.5"/><rect x="14" y="14" width="7" height="7" rx="1.5"/></svg>Dashboard</div>
          <div class="nav-item" onclick="openNew()"><svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>Nieuw Ticket</div>
          <div class="nav-item" data-nav="list" onclick="go('list')"><svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M3 12h18M3 18h18"/></svg>Alle Tickets</div>
          <div class="nav-item" data-nav="archive" onclick="go('archive')"><svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="4" rx="1"/><path d="M5 8v11a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V8"/><path d="M10 12h4"/></svg>Archief</div>
          ${isAdmin()?`<div class="nav-item" data-nav="settings" onclick="go('settings')"><svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12.22 2h-.44a2 2 0 0 0-2 2v.18a2 2 0 0 1-1 1.73l-.43.25a2 2 0 0 1-2 0l-.15-.08a2 2 0 0 0-2.73.73l-.22.38a2 2 0 0 0 .73 2.73l.15.1a2 2 0 0 1 1 1.72v.51a2 2 0 0 1-1 1.74l-.15.09a2 2 0 0 0-.73 2.73l.22.38a2 2 0 0 0 2.73.73l.15-.08a2 2 0 0 1 2 0l.43.25a2 2 0 0 1 1 1.73V20a2 2 0 0 0 2 2h.44a2 2 0 0 0 2-2v-.18a2 2 0 0 1 1-1.73l.43-.25a2 2 0 0 1 2 0l.15.08a2 2 0 0 0 2.73-.73l.22-.39a2 2 0 0 0-.73-2.73l-.15-.08a2 2 0 0 1-1-1.74v-.5a2 2 0 0 1 1-1.74l.15-.09a2 2 0 0 0 .73-2.73l-.22-.38a2 2 0 0 0-2.73-.73l-.15.08a2 2 0 0 1-2 0l-.43-.25a2 2 0 0 1-1-1.73V4a2 2 0 0 0-2-2z"/><circle cx="12" cy="12" r="3"/></svg>Instellingen</div>`:""}
          <div class="nav-item" onclick="refresh()"><svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 12a9 9 0 0 1 15-6.7L21 8"/><path d="M21 3v5h-5"/><path d="M21 12a9 9 0 0 1-15 6.7L3 16"/><path d="M3 21v-5h5"/></svg>Vernieuwen</div>
        </nav>
        <div class="sb-section">Per status</div><div class="sb-list" id="sbStatus"></div>
        <div class="sb-section">Snelfilters</div><div class="sb-list" id="sbQuick"></div>
        <div class="sb-foot">
          <div class="userchip"><div class="av">${initials(currentUser.name)}</div><div><div class="nm">${esc(currentUser.name)}</div><div class="rl">${isAdmin()?'<span class="rolepill">Beheerder</span>':'Gebruiker'}</div></div></div>
          <button class="logout" onclick="logout()"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><path d="m16 17 5-5-5-5M21 12H9"/></svg>Uitloggen</button>
        </div>
      </aside>
      <main class="main"><div class="main-inner" id="view"></div></main>
    </div>`;
  view="dashboard"; currentId=null; render(); startPolling();
}
function toggleSidebar(force){ const sb=document.getElementById("sidebar"), sc=document.getElementById("sbScrim"); if(!sb)return; const open=force===undefined?!sb.classList.contains("open"):force; sb.classList.toggle("open",open); if(sc)sc.classList.toggle("show",open); }
async function refresh(){
  toast("Vernieuwen…");
  const keepDraft=view==="settings"&&settingsDirty();
  try{
    await Promise.all([loadTickets(), keepDraft?null:loadSettings()]);
    if(view==="settings"&&!keepDraft) settingsDraft=clone(SETTINGS);
    render(); toast(keepDraft?"Tickets bijgewerkt — je niet-opgeslagen instellingen zijn behouden":"Bijgewerkt");
  }catch(e){ toast("Kon niet vernieuwen"); }
}

/* ===================== VISIBILITY ===================== */
function isFollower(t){ return (t.followers||[]).some(x=>(x.upn||"").toLowerCase()===currentUser.upn); }
function allVisible(){ return isAdmin()?tickets.slice():tickets.filter(t=>t.ownerUpn===currentUser.upn||isFollower(t)); }
function visibleTickets(){ return allVisible().filter(t=>!t.archived); }
function archivedTickets(){ return allVisible().filter(t=>t.archived); }
function counts(){ const c={open:0,progress:0,done:0,closed:0}; visibleTickets().forEach(t=>c[t.status]!==undefined&&c[t.status]++); return c; }
function setNav(){ document.querySelectorAll(".nav-item[data-nav]").forEach(e=>e.classList.toggle("active", e.dataset.nav===view && !currentId)); }
function confirmLeave(next){
  if(view!=="settings"||next==="settings"||!settingsDirty()) return true;
  return confirm("Je hebt niet-opgeslagen instellingen. Pagina verlaten zonder op te slaan?");
}
function go(v){
  if(document.getElementById("sbScrim"))toggleSidebar(false);
  if(v==="settings"&&view==="settings") return;
  if(!confirmLeave(v)) return;
  view=v; currentId=null;
  if(v==="settings") settingsDraft=clone(SETTINGS);
  render();
}
function render(){ if(!currentUser)return; renderSidebar(); setNav(); if(view==="dashboard")renderDashboard(); else if(view==="list")renderListView(); else if(view==="archive")renderArchive(); else if(view==="detail")renderDetail(); else if(view==="settings")renderSettings(); }
function renderSidebar(){
  const c=counts();
  document.getElementById("sbStatus").innerHTML=Object.keys(STATUS).map(k=>`<div class="sb-link" onclick="quick('status','${k}')"><span class="dot" style="background:${STATUS[k].color}"></span><span class="lbl">${STATUS[k].label}</span><span class="cnt">${c[k]}</span></div>`).join("");
  const vt=visibleTickets();
  const cW=vt.filter(t=>t.subcategory==="Webshop login").length, cA=vt.filter(t=>t.subcategory==="Artikelen aanmaken").length;
  const cH=vt.filter(t=>t.priority==="high"&&t.status!=="done"&&t.status!=="closed").length, cU=vt.filter(t=>(!t.assignee||t.assignee==="Niet toegewezen")&&t.status!=="closed").length;
  let q="";
  q+=`<div class="sb-link" onclick="quick('special','webshoplogin')"><span class="dot" style="background:var(--sub)"></span><span class="lbl">Webshop logins</span><span class="cnt">${cW}</span></div>`;
  q+=`<div class="sb-link" onclick="quick('subcategory','Artikelen aanmaken')"><span class="dot" style="background:var(--sales)"></span><span class="lbl">Artikelen aanmaken</span><span class="cnt">${cA}</span></div>`;
  q+=`<div class="sb-link" onclick="quick('special','highprio')"><span class="dot" style="background:var(--p-high)"></span><span class="lbl">Hoge prioriteit</span><span class="cnt">${cH}</span></div>`;
  q+=`<div class="sb-link" onclick="quick('special','unassigned')"><span class="dot" style="background:var(--faint)"></span><span class="lbl">Niet toegewezen</span><span class="cnt">${cU}</span></div>`;
  const me=myTeamName();
  if(me){ const mine=vt.filter(t=>t.assignee===me&&t.status!=="closed").length; q+=`<div class="sb-link" onclick="quick('assignee','${jsq(me)}')"><span class="dot" style="background:var(--primary)"></span><span class="lbl">Aan mij toegewezen</span><span class="cnt">${mine}</span></div>`; }
  document.getElementById("sbQuick").innerHTML=q;
}
function quick(type,v){ if(document.getElementById("sbScrim"))toggleSidebar(false); if(!confirmLeave("list")) return; filter={q:"",category:"",subcategory:"",status:"",assignee:"",special:""}; if(type==="status")filter.status=v; else if(type==="special")filter.special=v; else if(type==="assignee")filter.assignee=v; else if(type==="subcategory")filter.subcategory=v; view="list"; render(); }

/* ===================== DASHBOARD / LIST / ARCHIVE ===================== */
function renderDashboard(){
  const c=counts(), total=visibleTickets().length;
  const active=visibleTickets().filter(t=>t.status==="open"||t.status==="progress").sort((a,b)=>b.createdAt-a.createdAt).slice(0,6);
  document.getElementById("view").innerHTML=`
    <div class="page-head"><div><h2>Dashboard</h2><div class="sub">${isAdmin()?`${total} ticket${total===1?"":"s"} totaal`:`Je hebt ${total} ticket${total===1?"":"s"} ingediend`}</div></div>
      <button class="btn btn-dark" onclick="openNew()"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>Nieuw Ticket</button></div>
    <div class="stat-grid">${statCard("inbox",c.open,"Open","var(--st-open)","var(--st-open-bg)")}${statCard("loop",c.progress,"In Behandeling","var(--st-prog)","var(--st-prog-bg)")}${statCard("check",c.done,"Opgelost","var(--st-done)","var(--st-done-bg)")}${statCard("x",c.closed,"Gesloten","var(--st-closed)","var(--st-closed-bg)")}</div>
    <div class="section-head"><h3>Actieve tickets</h3>${total?`<span class="link" onclick="go('list')">Bekijk alle →</span>`:""}</div>
    ${active.length?`<div class="card-grid">${active.map(tcard).join("")}</div>`:emptyBox(total,!isAdmin())}`;
}
function statCard(icon,n,label,color,bg){ const ic={inbox:'<path d="M22 12h-6l-2 3h-4l-2-3H2"/><path d="M5.5 5.5 2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.5-6.5A2 2 0 0 0 16.8 4H7.2a2 2 0 0 0-1.7 1.5Z"/>',loop:'<path d="M3 12a9 9 0 0 1 15-6.7L21 8"/><path d="M21 3v5h-5"/><path d="M21 12a9 9 0 0 1-15 6.7L3 16"/><path d="M3 21v-5h5"/>',check:'<path d="M21.8 10A10 10 0 1 1 17 3.3"/><path d="m9 11 3 3L22 4"/>',x:'<circle cx="12" cy="12" r="9"/><path d="m15 9-6 6m0-6 6 6"/>'}[icon]; return `<div class="stat"><div class="ic" style="background:${bg};color:${color}"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">${ic}</svg></div><div><div class="n">${n}</div><div class="l">${label}</div></div></div>`; }
function tcard(t){
  const cat=CAT[t.category]||CAT.Verkoop, st=STATUS[t.status]||STATUS.open, pr=PRIO[t.priority]||PRIO.mid;
  const subTag=t.subcategory?`<span class="tag" style="color:var(--sub);background:var(--sub-bg)">${esc(t.subcategory)}</span>`:"";
  const pv=previewOf(t);
  return `<div class="tcard ${t.archived?"arch":""}" onclick="openDetail('${t.itemId}')">
    <div class="ch"><h4>${esc(t.subject)}</h4><span class="badge" style="color:${st.color};background:${st.bg}">${st.label}</span></div>
    <div class="desc ${pv?"":"none"}">${pv?esc(pv):"Geen omschrijving"}</div>
    <div class="tags"><span class="tag" style="color:${cat.color};background:${cat.bg}">${cat.label}</span>${subTag}<span class="tag" style="color:${pr.color};background:${pr.bg}">${pr.label}</span></div>
    <div class="who"><span class="it"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/></svg>Door <b>${esc(firstName(t.author))}</b></span>
      <span class="it"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 21v-2a4 4 0 0 0-3-3.9"/><path d="M16 3.1a4 4 0 0 1 0 7.8"/><circle cx="9" cy="7" r="4"/><path d="M2 21v-2a4 4 0 0 1 4-4h4a4 4 0 0 1 4 4v2"/></svg>Behandelaar <b>${esc(firstName(t.assignee))}</b></span>
      <span class="ti">${ago(t.createdAt)}</span></div></div>`;
}
function emptyBox(total,isUser){ return `<div class="empty"><svg width="44" height="44" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M9 13h6M9 17h4"/></svg><h3>${total?"Geen actieve tickets":(isUser?"Je hebt nog geen tickets":"Nog geen tickets")}</h3><p>${total?"Alle tickets zijn opgelost of gesloten.":"Maak een ticket aan om te starten."}</p></div>`; }
function renderListView(){
  const subOpts=["Alle onderdelen",...ALL_SUBS];
  document.getElementById("view").innerHTML=`
    <div class="page-head"><div><h2>Alle tickets</h2><div class="sub" id="listSub"></div></div><button class="btn btn-dark" onclick="openNew()"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>Nieuw Ticket</button></div>
    <div class="toolbar">
      <div class="search"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3"/></svg><input id="q" placeholder="Zoek op onderwerp, klant of referentie…" oninput="filter.q=this.value;renderList()" /></div>
      <select class="filter" id="fCat" onchange="filter.category=this.value;renderList()"><option value="">Alle categorieën</option><option>Verkoop</option><option>Technisch</option></select>
      <select class="filter" id="fSub" onchange="filter.subcategory=this.value==='Alle onderdelen'?'':this.value;renderList()">${subOpts.map(o=>`<option>${esc(o)}</option>`).join("")}</select>
      <select class="filter" id="fStatus" onchange="filter.status=this.value;renderList()"><option value="">Alle statussen</option>${Object.keys(STATUS).map(k=>`<option value="${k}">${STATUS[k].label}</option>`).join("")}</select>
      <select class="filter" id="fAssignee" onchange="filter.assignee=this.value;renderList()"><option value="">Alle behandelaars</option>${[...new Set([...teamNames(),...visibleTickets().map(t=>t.assignee||"Niet toegewezen")])].map(n=>`<option>${esc(n)}</option>`).join("")}</select>
    </div><div id="listResults"></div>`;
  document.getElementById("q").value=filter.q; document.getElementById("fCat").value=filter.category; document.getElementById("fSub").value=filter.subcategory||"Alle onderdelen"; document.getElementById("fStatus").value=filter.status; document.getElementById("fAssignee").value=filter.assignee;
  renderList();
}
function applyFilter(items){
  let r=items.slice();
  if(filter.category)r=r.filter(t=>t.category===filter.category);
  if(filter.subcategory)r=r.filter(t=>t.subcategory===filter.subcategory);
  if(filter.status)r=r.filter(t=>t.status===filter.status);
  if(filter.assignee)r=r.filter(t=>(t.assignee||"Niet toegewezen")===filter.assignee);
  if(filter.special==="webshoplogin")r=r.filter(t=>t.subcategory==="Webshop login");
  if(filter.special==="highprio")r=r.filter(t=>t.priority==="high"&&t.status!=="done"&&t.status!=="closed");
  if(filter.special==="unassigned")r=r.filter(t=>(!t.assignee||t.assignee==="Niet toegewezen")&&t.status!=="closed");
  if(filter.q){ const q=filter.q.toLowerCase(); r=r.filter(t=>searchText(t).includes(q)); }
  return r.sort((a,b)=>b.createdAt-a.createdAt);
}
function renderList(){
  const items=applyFilter(visibleTickets());
  const sp={webshoplogin:"Webshop logins",highprio:"Hoge prioriteit",unassigned:"Niet toegewezen"}[filter.special];
  const chips=[]; if(sp)chips.push(`<span class="activefilter">${sp}<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" onclick="filter.special='';renderList()"><path d="M18 6 6 18M6 6l12 12"/></svg></span>`);
  if(filter.subcategory)chips.push(`<span class="activefilter">${esc(filter.subcategory)}<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" onclick="filter.subcategory='';document.getElementById('fSub').value='Alle onderdelen';renderList()"><path d="M18 6 6 18M6 6l12 12"/></svg></span>`);
  const chipRow=chips.length?`<div style="margin-bottom:14px;display:flex;gap:8px;flex-wrap:wrap">${chips.join("")}</div>`:"";
  const sub=document.getElementById("listSub"); if(sub)sub.textContent=`${items.length} ticket${items.length===1?"":"s"}`;
  document.getElementById("listResults").innerHTML=chipRow+(items.length?`<div class="card-grid">${items.map(tcard).join("")}</div>`:`<div class="empty"><svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3"/></svg><h3>Geen tickets gevonden</h3><p>Pas je filters of zoekopdracht aan.</p></div>`);
}
function renderArchive(){
  const items=archivedTickets().sort((a,b)=>b.createdAt-a.createdAt);
  document.getElementById("view").innerHTML=`<div class="page-head"><div><h2>Archief</h2><div class="sub">${items.length} gearchiveerd ticket${items.length===1?"":"s"} · data blijft volledig bewaard</div></div></div>
    ${items.length?`<div class="card-grid">${items.map(tcard).join("")}</div>`:`<div class="empty"><svg width="44" height="44" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><rect x="3" y="4" width="18" height="4" rx="1"/><path d="M5 8v11a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V8"/><path d="M10 12h4"/></svg><h3>Archief is leeg</h3><p>Gearchiveerde tickets verschijnen hier.</p></div>`}`;
}

/* ===================== NEW TICKET ===================== */
function openNew(){
  newCat=null; curSchema=null; newFiles=[]; renderNewFiles();
  document.querySelectorAll("#catSeg button").forEach(b=>b.classList.remove("active"));
  document.getElementById("subField").classList.add("hidden"); document.getElementById("dynArea").classList.add("hidden");
  document.getElementById("stepHint").classList.remove("hidden"); document.getElementById("f_prio").value="mid";
  document.getElementById("f_assignee").innerHTML=teamNames().map(n=>`<option>${esc(n)}</option>`).join("");
  document.getElementById("f_assignee").value="Niet toegewezen"; document.getElementById("assigneeField").classList.toggle("hidden",!isAdmin());
  document.getElementById("dynFields").innerHTML=""; document.getElementById("saveBtn").disabled=false; document.getElementById("saveBtn").textContent="Ticket aanmaken";
  document.getElementById("overlay").classList.add("show");
}
function closeModal(){ document.getElementById("overlay").classList.remove("show"); }
function setCat(c){ newCat=c; document.querySelectorAll("#catSeg button").forEach(b=>b.classList.toggle("active",b.dataset.v===c));
  const subs=Object.keys(FORMS[c]); document.getElementById("f_subcat").innerHTML=`<option value="">Maak een keuze</option>`+subs.map(s=>`<option>${esc(s)}</option>`).join("");
  document.getElementById("subField").classList.remove("hidden"); document.getElementById("dynArea").classList.add("hidden"); document.getElementById("stepHint").classList.remove("hidden"); document.getElementById("dynFields").innerHTML=""; curSchema=null;
}
function onSubcatChange(){
  const sub=document.getElementById("f_subcat").value;
  if(!newCat||!sub){ document.getElementById("dynArea").classList.add("hidden"); document.getElementById("stepHint").classList.remove("hidden"); curSchema=null; return; }
  curSchema=FORMS[newCat][sub];
  document.getElementById("dynFields").innerHTML=(curSchema.infobox||"")+curSchema.fields.map(fieldHTML).join("");
  document.getElementById("dynArea").classList.remove("hidden"); document.getElementById("stepHint").classList.add("hidden"); updateConditionals();
  const first=curSchema.fields.find(f=>!f.showIf); const el=first&&document.getElementById("dyn_"+first.k); if(el)setTimeout(()=>el.focus(),40);
}
function fieldHTML(f){
  const req=f.req?' <span class="req">*</span>':'', hint=f.hint?` <span class="hint">(${f.hint})</span>`:'', hid=f.showIf?' hidden':'';
  let inner;
  if(f.type==="textarea") inner=`<textarea id="dyn_${f.k}" placeholder="Beschrijf de vraag…"></textarea>`;
  else if(f.type==="yesno") inner=`<select id="dyn_${f.k}" onchange="updateConditionals()"><option value="">Maak een keuze</option><option>Ja</option><option>Nee</option></select>`;
  else if(f.type==="select") inner=`<select id="dyn_${f.k}"><option value="">Maak een keuze</option>${f.options.map(o=>`<option>${esc(o)}</option>`).join("")}</select>`;
  else if(f.type==="daterange") inner=`<div class="row2"><input type="date" id="dyn_${f.k}_from"><input type="date" id="dyn_${f.k}_to"></div>`;
  else inner=`<input id="dyn_${f.k}" />`;
  return `<div class="field${hid}" id="wrap_${f.k}"><label>${esc(f.label)}${req}${hint}</label>${inner}</div>`;
}
function updateConditionals(){ if(!curSchema)return; curSchema.fields.forEach(f=>{ if(!f.showIf)return; const ctrl=document.getElementById("dyn_"+f.showIf.k); const show=ctrl&&ctrl.value===f.showIf.val; const w=document.getElementById("wrap_"+f.k); if(w)w.classList.toggle("hidden",!show); }); }
function readField(f){ if(f.type==="daterange"){ const a=fmtD(val("dyn_"+f.k+"_from")),b=fmtD(val("dyn_"+f.k+"_to")); if(!a&&!b)return""; return `${a||"?"} – ${b||"?"}`; } return val("dyn_"+f.k); }
function addNewFiles(files){ for(const f of files){ newFiles.push({id:"n"+Date.now()+Math.random().toString(36).slice(2,6),file:f,name:f.name,size:f.size}); } renderNewFiles(); document.getElementById("newFileInput").value=""; }
function renderNewFiles(){ const el=document.getElementById("newFiles"); if(el)el.innerHTML=newFiles.map(f=>fileChip(f,`removeNewFile('${f.id}')`)).join(""); }
function removeNewFile(id){ newFiles=newFiles.filter(f=>f.id!==id); renderNewFiles(); }
function fileChip(f,rm){ return `<div class="filechip"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" style="color:var(--muted);flex:none"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/></svg><span class="fn">${esc(f.name)}</span><span class="fs">${fmtSize(f.size)}</span>${rm?`<span class="rm" onclick="${rm}"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 6 6 18M6 6l12 12"/></svg></span>`:""}</div>`; }

async function saveTicket(){
  if(saving)return;
  if(!newCat){ toast("Kies een categorie"); return; }
  const sub=document.getElementById("f_subcat").value; if(!sub){ toast("Kies een onderdeel"); return; }
  const schema=FORMS[newCat][sub];
  const raw={}; schema.fields.forEach(f=>raw[f.k]=readField(f));
  const visible=f=>!f.showIf||raw[f.showIf.k]===f.showIf.val;
  for(const f of schema.fields){ if(f.req&&visible(f)&&!raw[f.k]){ toast(`Vul "${f.label}" in`); const el=document.getElementById("dyn_"+f.k)||document.getElementById("dyn_"+f.k+"_from"); if(el)el.focus(); return; } }
  const subjField=schema.fields.find(f=>f.isSubject);
  const subject=subjField?raw[subjField.k]:(schema.title?schema.title(raw):sub);
  const description=raw.omschrijving||"";
  const fields=schema.fields.filter(f=>!f.isSubject&&f.k!=="omschrijving"&&visible(f)&&raw[f.k]).map(f=>({label:f.label,value:raw[f.k]}));
  const year=new Date().getFullYear();
  const nums=tickets.filter(t=>t.ref&&t.ref.startsWith("VRP-"+year)).map(t=>parseInt((t.ref.split("-")[2]||"0"),10)||0);
  const ref=`VRP-${year}-${String((nums.length?Math.max(...nums):0)+1).padStart(4,"0")}`;
  const assignee=isAdmin()?document.getElementById("f_assignee").value:"Niet toegewezen";

  saving=true; const btn=document.getElementById("saveBtn"); btn.disabled=true; btn.textContent="Opslaan…";
  try{
    const att=[];
    for(const nf of newFiles){ btn.textContent=`Uploaden ${nf.name}…`; att.push(await uploadFile(ref,nf.file)); }
    const messages=[]; if(description||att.length) messages.push({author:currentUser.name,internal:false,ts:Date.now(),text:description||(att.length?"Bijlage(n) toegevoegd":""),attachments:att});
    const fieldsObj={};
    fieldsObj[COL.Title]=subject; fieldsObj[COL.Ref]=ref; fieldsObj[COL.Category]=newCat; fieldsObj[COL.Subcategory]=sub;
    fieldsObj[COL.Status]="open"; fieldsObj[COL.Priority]=document.getElementById("f_prio").value;
    fieldsObj[COL.Assignee]=assignee; fieldsObj[COL.Indiener]=currentUser.name; fieldsObj[COL.OwnerUpn]=currentUser.upn;
    fieldsObj[COL.Description]=description; fieldsObj[COL.FieldsJson]=JSON.stringify(fields);
    fieldsObj[COL.MessagesJson]=JSON.stringify(messages); fieldsObj[COL.Archived]=false;
    if(COL.Followers) fieldsObj[COL.Followers]=JSON.stringify([]);
    const t=await createTicketItem(fieldsObj);
    tickets.push(t);
    console.log("[saveTicket] Ticket aangemaakt, notifyNewTicket wordt aangeroepen. adminEmails:",adminEmails);
    notifyNewTicket(t);
    console.log("[saveTicket] notifyNewTicket klaar");
    closeModal(); openDetail(t.itemId); toast(`Ticket ${ref} aangemaakt`);
  }catch(e){ console.error("[saveTicket] FOUT:",e.message); toast("Opslaan mislukt — controleer je rechten"); }
  finally{ saving=false; btn.disabled=false; btn.textContent="Ticket aanmaken"; }
}

/* ===================== DETAIL ===================== */
function findTicket(itemId){ return tickets.find(t=>t.itemId===itemId); }
async function openDetail(itemId){
  if(!confirmLeave("detail")) return;
  currentId=itemId; view="detail";
  document.querySelectorAll(".nav-item[data-nav]").forEach(e=>e.classList.remove("active"));
  detailTicket=findTicket(itemId); replyFiles=[]; replyInternal=false;
  if(!detailTicket){ document.getElementById("view").innerHTML=`<div class="empty"><h3>Ticket niet gevonden</h3></div>`; return; }
  if(!isAdmin()&&detailTicket.ownerUpn!==currentUser.upn&&!isFollower(detailTicket)){ detailTicket=null; document.getElementById("view").innerHTML=`<div class="empty"><svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg><h3>Geen toegang</h3><p>Je kunt enkel je eigen of gedeelde tickets bekijken.</p></div>`; return; }
  renderDetail();
}

/* Render message text: stored as HTML (rich) or plain text (legacy) */
function renderMsgText(text){
  if(!text) return "";
  // If the text contains any HTML tags anywhere, render as HTML
  if(/<[a-zA-Z][^>]*>/.test(text)) return `<div class="mc">${text}</div>`;
  // Otherwise treat as plain text
  return `<div class="mc">${esc(text).replace(/\n/g,"<br>")}</div>`;
}

function renderDetail(){
  const t=detailTicket; if(!t){ go("list"); return; }
  const cat=CAT[t.category]||CAT.Verkoop, st=STATUS[t.status]||STATUS.open, pr=PRIO[t.priority]||PRIO.mid;
  const subTag=t.subcategory?`<span class="tag" style="color:var(--sub);background:var(--sub-bg)">${esc(t.subcategory)}</span>`:"";
  const convo=t.messages.length?t.messages.map(m=>`<div class="msg ${m.internal?"internal":""}"><div class="av">${initials(m.author)}</div><div class="mb">
      <div class="mh"><span class="nm">${esc(m.author)}</span>${m.internal?`<span class="int-badge">Interne notitie</span>`:""}<span class="tm">${ago(m.ts)}</span></div>
      ${renderMsgText(m.text)}
      ${m.attachments&&m.attachments.length?`<div class="msg-att">${m.attachments.map(a=>`<div class="filechip"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" style="color:var(--muted);flex:none"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/></svg><span class="fn">${esc(a.name)}</span><span class="fs">${fmtSize(a.size)}</span><span class="pv" title="Voorbeeld" onclick="previewAtt('${a.itemId}','${esc(a.name).replace(/'/g,"\\'")}')"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7-10-7-10-7Z"/><circle cx="12" cy="12" r="3"/></svg></span><span class="dl" title="Downloaden" onclick="downloadAtt('${a.itemId}')"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 3v12m0 0 4-4m-4 4-4-4"/><path d="M5 21h14"/></svg></span></div>`).join("")}</div>`:""}
    </div></div>`).join(""):`<div style="color:var(--faint);font-size:13px;padding:14px 0">Nog geen berichten in dit gesprek.</div>`;
  const fieldsCard=(t.fields&&t.fields.length)?`<div class="panel-card pc-pad" style="margin-bottom:22px"><div style="font-size:15px;font-weight:700;margin-bottom:14px">Aanvraaggegevens</div><div class="dl-list">${t.fields.map(f=>`<div class="dl-row"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 12h6M9 8h6M9 16h4"/><rect x="3" y="3" width="18" height="18" rx="2"/></svg><div><span class="k">${esc(f.label)}: </span><span class="v">${esc(f.value)}</span></div></div>`).join("")}</div></div>`:"";
  const canAssign=isAdmin();
  document.getElementById("view").innerHTML=`
    <div class="back" onclick="go('${t.archived?"archive":"list"}')"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M19 12H5m0 0 7 7m-7-7 7-7"/></svg>${t.archived?"Archief":"Alle tickets"}</div>
    ${t.archived?`<div class="arch-banner"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="4" rx="1"/><path d="M5 8v11a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V8"/></svg>Dit ticket is gearchiveerd. Alle gegevens blijven bewaard.</div>`:""}
    <div class="detail-layout"><div>
      <div class="panel-card pc-pad" style="margin-bottom:22px"><div class="d-titlerow"><div><div class="d-title">${esc(t.subject)}</div>
        <div class="tags" style="margin-top:10px"><span class="tag" style="color:${cat.color};background:${cat.bg}">${cat.label}</span>${subTag}<span class="tag" style="color:${pr.color};background:${pr.bg}">${pr.label}</span></div></div>
        <span class="badge" style="color:${st.color};background:${st.bg}">${st.label}</span></div></div>
      <div class="panel-card"><div class="pc-head">Gesprek</div><div class="pc-pad" style="padding-top:4px;padding-bottom:4px">${convo}</div>
        <div class="reply">
          <div class="editor-toolbar" id="editorToolbar">
            <button title="Vetgedrukt" onclick="editorCmd('bold')"><b>B</b></button>
            <button title="Cursief" onclick="editorCmd('italic')"><i>I</i></button>
            <button title="Onderstrepen" onclick="editorCmd('underline')"><u>U</u></button>
            <div class="sep"></div>
            <button title="Ongeordende lijst" onclick="editorCmd('insertUnorderedList')"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01"/></svg></button>
            <button title="Geordende lijst" onclick="editorCmd('insertOrderedList')"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10 6h11M10 12h11M10 18h11M4 6h1v4M4 10H6M4 14v1a1 1 0 0 0 1 1h1a1 1 0 0 0 1-1v-1a1 1 0 0 0-1-1H5a1 1 0 0 1-1-1v-1a1 1 0 0 1 1-1h2"/></svg></button>
            <div class="sep"></div>
            <button title="Tabel invoegen" onclick="insertTable()"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18M15 3v18"/></svg></button>
            <div class="sep"></div>
            <button class="img-upload-btn" title="Afbeelding invoegen" onclick="document.getElementById('inlineImgInput').click()"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><path d="m21 15-5-5L5 21"/></svg> Afbeelding</button>
            <input type="file" id="inlineImgInput" accept="image/*" multiple style="display:none" onchange="insertInlineImages(this.files)">
          </div>
          <div class="reply-editor" id="replyEditor" contenteditable="true" data-placeholder="Typ je antwoord…"></div>
          <div id="replyFiles"></div>
          <div class="reply-bar">
            <button class="attach-btn" onclick="document.getElementById('replyFileInput').click()"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="m21.4 11.1-9.2 9.2a5 5 0 0 1-7-7l9.1-9.2a3.3 3.3 0 0 1 4.7 4.7l-9.2 9.1a1.7 1.7 0 0 1-2.3-2.3l8.5-8.5"/></svg>Bijlage</button>
            <input type="file" id="replyFileInput" multiple style="display:none" onchange="addReplyFiles(this.files)" />
            <div class="toggle ${replyInternal?"on":""}" id="intToggle" onclick="toggleInternal()"><span class="switch"></span>Interne notitie</div>
            <div class="spacer"></div>
            <button class="btn btn-primary" id="replyBtn" onclick="sendReply()"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="m22 2-7 20-4-9-9-4Z"/><path d="M22 2 11 13"/></svg>Versturen</button>
          </div></div></div>
    </div><div>
      <div class="panel-card pc-pad" style="margin-bottom:22px"><div style="font-size:15px;font-weight:700;margin-bottom:14px">Details</div><div class="dl-list">
        ${dlRow('tag',"Type",cat.label)}${t.subcategory?dlRow('grid',"Onderdeel",t.subcategory):""}${dlRow('handle',"Behandelaar",t.assignee||"—")}${dlRow('user',"Ingediend door",t.author||"—")}${dlRow('clock',"Aangemaakt",fmtDate(t.createdAt))}${dlRow('ref',"Referentie",t.ref)}</div></div>
      ${fieldsCard}
      ${shareCard(t)}
      <div class="panel-card pc-pad"><div style="font-size:15px;font-weight:700;margin-bottom:14px">Bijwerken</div>
        ${canAssign?`<div class="upd-field"><label>Toegewezen aan (behandelaar)</label><select onchange="updateField('assignee',this.value)">${assigneeOptions(t.assignee).map(n=>`<option ${t.assignee===n?"selected":""}>${esc(n)}</option>`).join("")}</select></div>`:`<div class="upd-field"><label>Behandelaar</label></div><div class="lockrow"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>${esc(t.assignee||"Niet toegewezen")} — enkel een beheerder kan dit wijzigen</div>`}
        <div class="upd-field"><label>Status</label><select onchange="updateField('status',this.value)">${Object.keys(STATUS).map(k=>`<option value="${k}" ${t.status===k?"selected":""}>${STATUS[k].label}</option>`).join("")}</select></div>
        <div class="upd-field"><label>Prioriteit</label><select onchange="updateField('priority',this.value)">${Object.keys(PRIO).map(k=>`<option value="${k}" ${t.priority===k?"selected":""}>${PRIO[k].label}</option>`).join("")}</select></div>
        <button class="btn btn-ghost btn-block" onclick="toggleArchive()" style="margin-bottom:10px"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="4" rx="1"/><path d="M5 8v11a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V8"/><path d="M10 12h4"/></svg>${t.archived?"Herstellen uit archief":"Archiveren"}</button>
        ${isAdmin()?`<button class="btn btn-danger btn-block" onclick="deleteTicket()"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2m3 0v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6"/></svg>Ticket verwijderen</button>`:""}
      </div>
    </div></div>`;
}
function dlRow(icon,k,v){ const ic={tag:'<path d="M7 7h.01"/><path d="M3 5v6.6a2 2 0 0 0 .6 1.4l8.4 8.4a2 2 0 0 0 2.8 0l5.6-5.6a2 2 0 0 0 0-2.8L12.6 4.6A2 2 0 0 0 11.2 4H5a2 2 0 0 0-2 2Z"/>',user:'<circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0"/>',handle:'<path d="M22 21v-2a4 4 0 0 0-3-3.9"/><path d="M16 3.1a4 4 0 0 1 0 7.8"/><circle cx="9" cy="7" r="4"/><path d="M2 21v-2a4 4 0 0 1 4-4h4a4 4 0 0 1 4 4v2"/>',clock:'<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',grid:'<rect x="3" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="3" width="7" height="7" rx="1.5"/><rect x="3" y="14" width="7" height="7" rx="1.5"/><rect x="14" y="14" width="7" height="7" rx="1.5"/>',ref:'<path d="M4 7V4h16v3M9 20h6M12 4v16"/>'}[icon]; return `<div class="dl-row"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">${ic}</svg><div><span class="k">${k}: </span><span class="v">${esc(v)}</span></div></div>`; }
function toggleInternal(){ replyInternal=!replyInternal; document.getElementById("intToggle").classList.toggle("on",replyInternal); }
function addReplyFiles(files){ for(const f of files){ replyFiles.push({id:"r"+Date.now()+Math.random().toString(36).slice(2,6),file:f,name:f.name,size:f.size}); } renderReplyFiles(); document.getElementById("replyFileInput").value=""; }
function renderReplyFiles(){ const el=document.getElementById("replyFiles"); if(el)el.innerHTML=replyFiles.map(f=>fileChip(f,`removeReplyFile('${f.id}')`)).join(""); }
function removeReplyFile(id){ replyFiles=replyFiles.filter(f=>f.id!==id); renderReplyFiles(); }

/* ===================== RICH TEXT EDITOR ===================== */
function editorCmd(cmd){ document.getElementById("replyEditor").focus(); document.execCommand(cmd,false,null); }

function insertTable(){
  const cols=parseInt(prompt("Aantal kolommen:",3)||3);
  const rows=parseInt(prompt("Aantal rijen (inclusief header):",3)||3);
  if(!cols||!rows||cols<1||rows<1) return;
  let html="<table><thead><tr>";
  for(let c=0;c<cols;c++) html+=`<th>Kolom ${c+1}</th>`;
  html+="</tr></thead><tbody>";
  for(let r=0;r<rows-1;r++){
    html+="<tr>";
    for(let c=0;c<cols;c++) html+="<td>&nbsp;</td>";
    html+="</tr>";
  }
  html+="</tbody></table><p><br></p>";
  document.getElementById("replyEditor").focus();
  document.execCommand("insertHTML",false,html);
}

function insertInlineImages(files){
  const editor=document.getElementById("replyEditor");
  if(!editor) return;
  for(const file of files){
    if(!file.type.startsWith("image/")) continue;
    const reader=new FileReader();
    reader.onload=e=>{
      editor.focus();
      document.execCommand("insertHTML",false,`<img src="${e.target.result}" alt="${esc(file.name)}" style="max-width:100%;border-radius:6px;margin:4px 0">`);
    };
    reader.readAsDataURL(file);
  }
  document.getElementById("inlineImgInput").value="";
}

function getEditorContent(){
  const el=document.getElementById("replyEditor");
  if(!el) return "";
  // Return inner HTML; empty = just <br> or whitespace
  const html=el.innerHTML.trim();
  if(html===""||html==="<br>"||html==="<br/>") return "";
  return html;
}

async function sendReply(){
  const txt=getEditorContent(); if(!txt&&!replyFiles.length){ toast("Schrijf een bericht of voeg een bestand toe"); return; }
  const btn=document.getElementById("replyBtn"); btn.disabled=true;
  try{
    const att=[]; for(const rf of replyFiles) att.push(await uploadFile(detailTicket.ref,rf.file));
    detailTicket.messages.push({author:currentUser.name,internal:replyInternal,ts:Date.now(),text:txt,attachments:att});
    await patchTicket(detailTicket.itemId,{[COL.MessagesJson]:JSON.stringify(detailTicket.messages)});
    const wasInt=replyInternal;
    const snip=txt?htmlToText(txt).slice(0,160):(att.length?att.length+" bijlage(n) toegevoegd":"");
    notifyUpdate(wasInt?"internal":"message", detailTicket, [(wasInt?"Interne notitie van ":"Nieuw bericht van ")+currentUser.name+": "+snip], att);
    replyFiles=[]; replyInternal=false; renderDetail(); toast(wasInt?"Interne notitie toegevoegd":"Bericht verstuurd");
  }catch(e){ toast("Versturen mislukt"); btn.disabled=false; }
}

async function updateField(field,value){
  if(field==="assignee"&&!isAdmin()){ toast("Enkel beheerders kunnen de behandelaar wijzigen"); return; }
  const map={assignee:COL.Assignee,status:COL.Status,priority:COL.Priority};
  const prev=detailTicket[field]; detailTicket[field]=value;
  try{
    await patchTicket(detailTicket.itemId,{[map[field]]:value}); renderDetail(); renderSidebar(); toast("Ticket bijgewerkt");
    const lbl={status:"Status",priority:"Prioriteit",assignee:"Behandelaar"}[field];
    const disp=v=>field==="status"?(STATUS[v]?STATUS[v].label:v):field==="priority"?(PRIO[v]?PRIO[v].label:v):v;
    if(prev!==value){
      const line=lbl+" gewijzigd van "+disp(prev)+" naar "+disp(value)+" door "+currentUser.name;
      if(field==="status") notifyStatus(detailTicket,[line]); else notifyUpdate(field, detailTicket, [line]);
    }
  }
  catch(e){ detailTicket[field]=prev; renderDetail(); toast("Bijwerken mislukt"); }
}
async function toggleArchive(){
  detailTicket.archived=!detailTicket.archived;
  try{ await patchTicket(detailTicket.itemId,{[COL.Archived]:detailTicket.archived}); const a=detailTicket.archived; renderDetail(); renderSidebar(); toast(a?"Ticket gearchiveerd":"Ticket hersteld uit archief");
    notifyUpdate("archive", detailTicket, [a?"Ticket gearchiveerd door "+currentUser.name:"Ticket hersteld uit archief door "+currentUser.name]); }
  catch(e){ detailTicket.archived=!detailTicket.archived; toast("Archiveren mislukt"); }
}
async function downloadAtt(itemId){ try{ const url=await attachmentDownloadUrl(itemId); window.open(url,"_blank"); }catch(e){ toast("Bijlage kon niet geopend worden"); } }
async function deleteTicket(){
  if(!isAdmin())return;
  if(!confirm("Dit ticket definitief verwijderen? Overweeg 'Archiveren' om de gegevens te bewaren.")) return;
  try{ await deleteTicketItem(detailTicket.itemId); tickets=tickets.filter(t=>t.itemId!==detailTicket.itemId); toast("Ticket verwijderd"); go("list"); }
  catch(e){ toast("Verwijderen mislukt"); }
}

/* ===================== DELEN MET COLLEGA ===================== */
function shareCard(t){
  const followers=t.followers||[];
  const canShare = isAdmin() || t.ownerUpn===currentUser.upn || isFollower(t);
  if(!COL.Followers){ return canShare?`<div class="panel-card pc-pad" style="margin-bottom:22px"><div style="font-size:15px;font-weight:700;margin-bottom:6px">Delen met collega</div><div class="lockrow" style="margin:0"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><path d="M12 16v-4M12 8h.01"/></svg>Voeg een kolom <b>Followers</b> (meerdere regels tekst) toe aan de lijst om delen te activeren.</div></div>`:""; }
  const chips = followers.length ? followers.map(u=>`<span class="fchip">${esc(u.name||u.upn)}${canShare?`<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" onclick="removeFollower('${esc(u.upn)}')"><path d="M18 6 6 18M6 6l12 12"/></svg>`:""}</span>`).join("")
    : `<span style="font-size:12.5px;color:var(--faint)">Nog niet gedeeld met collega's</span>`;
  return `<div class="panel-card pc-pad" style="margin-bottom:22px">
    <div style="font-size:15px;font-weight:700;margin-bottom:6px;display:flex;align-items:center;gap:7px"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="color:var(--primary)"><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75"/></svg>Delen met collega</div>
    <div style="font-size:12.5px;color:var(--muted);margin-bottom:11px">Voer het e-mailadres van een collega in om het ticket met hem/haar te delen.</div>
    <div class="fchips">${chips}</div>
    ${canShare?`<div class="share-row"><input id="shareSearch" type="email" placeholder="e-mailadres collega…" onkeydown="if(event.key==='Enter')addFollowerFromInput()"><button class="btn btn-primary btn-sm" onclick="addFollowerFromInput()">Delen</button></div>`:""}
  </div>`;
}
function addFollowerFromInput(){ const el=document.getElementById("shareSearch"); if(!el)return; const email=(el.value||"").trim().toLowerCase(); if(!email||!email.includes("@")){ toast("Vul een geldig e-mailadres in"); return; } addFollower(email, email); el.value=""; }
async function persistFollowers(){ await patchTicket(detailTicket.itemId,{[COL.Followers]:JSON.stringify(detailTicket.followers)}); const ix=findTicket(detailTicket.itemId); if(ix) ix.followers=detailTicket.followers; }
async function addFollower(upn,name){
  upn=(upn||"").toLowerCase(); if(!upn)return;
  detailTicket.followers=detailTicket.followers||[];
  if(detailTicket.followers.some(x=>(x.upn||"").toLowerCase()===upn)){ toast("Al gedeeld met deze collega"); return; }
  detailTicket.followers.push({upn, name:name||upn});
  try{ await persistFollowers(); renderDetail();
    notifyShared(detailTicket, upn, name||upn, ["Ticket gedeeld met "+(name||upn)+" door "+currentUser.name]);
    toast(`Gedeeld met ${name||upn}`);
  }catch(e){ detailTicket.followers=detailTicket.followers.filter(x=>(x.upn||"").toLowerCase()!==upn); toast("Delen mislukt"); }
}
async function removeFollower(upn){
  upn=(upn||"").toLowerCase();
  detailTicket.followers=(detailTicket.followers||[]).filter(x=>(x.upn||"").toLowerCase()!==upn);
  try{ await persistFollowers(); renderDetail(); toast("Collega verwijderd"); }catch(e){ toast("Verwijderen mislukt"); }
}

/* ===================== E-MAIL ===================== */
/**
 * Verstuurt e-mail via de Cloudflare Worker (verpa-mail-proxy).
 * De Worker gebruikt een application-level Graph permissie (Mail.Send)
 * die éénmalig door de tenant-beheerder goedgekeurd is.
 * Er is geen delegated toestemming per gebruiker nodig.
 *
 * @param {string[]} toEmails  - Lijst van ontvangers (UPN / e-mailadres)
 * @param {string}   subject   - Onderwerpregel
 * @param {string}   html      - HTML-body van de mail
 */
async function sendMail(toEmails, subject, html){
  if(!CONFIG.mailWorker){ console.warn("sendMail: mailWorker niet geconfigureerd in CONFIG"); return false; }
  const recipients=(Array.isArray(toEmails)?toEmails:[toEmails]).filter(e=>e&&e.includes("@"));
  if(!recipients.length) return false;
  try{
    const res=await fetch(CONFIG.mailWorker,{
      method:"POST",
      headers:{"Content-Type":"application/json"},
      body:JSON.stringify({ to:recipients, subject, html })
    });
    if(!res.ok){
      const txt=await res.text();
      console.warn("sendMail: Worker antwoordde",res.status,txt.slice(0,200));
      return false;
    }
    return true;
  }catch(e){
    // Mail-fouten mogen de app niet blokkeren — stil loggen
    console.warn("sendMail mislukt:",e.message);
    return false;
  }
}
const statusLabel=k=>STATUS[k]?STATUS[k].label:k, prioLabel=k=>PRIO[k]?PRIO[k].label:k;
function emailShell(title, intro, rows, ticket, extraTable=""){
  const rowsHtml=rows.map(r=>`<tr><td style="padding:7px 0;color:#5b6677;font-size:13px;width:150px;vertical-align:top">${r[0]}</td><td style="padding:7px 0;color:#111826;font-size:13px;font-weight:600">${r[1]}</td></tr>`).join("");
  const base=CONFIG.redirectUri||"#";
  const link=ticket&&ticket.itemId?`${base}#ticket-${ticket.itemId}`:base;
  const head=ticket?`<div style="font-size:15px;font-weight:700;color:#111826;margin-bottom:4px">${esc(ticket.subject)}</div>
        <div style="font-size:12px;color:#98a2b3;margin-bottom:10px">${esc(ticket.ref)}</div>`:"";
  const logoUrl=`${base}/logo.jpg`;
  return `<div style="margin:0;padding:24px;background:#eef1f4;font-family:Inter,Segoe UI,Arial,sans-serif">
  <div style="max-width:560px;margin:0 auto;background:#fff;border:1px solid #e7ebf0;border-radius:14px;overflow:hidden">
    <table width="100%" cellpadding="0" cellspacing="0" bgcolor="#f37a2b" style="background:#f37a2b"><tr><td style="padding:20px 24px;background:#f37a2b">
      <table cellpadding="0" cellspacing="0"><tr><td style="width:42px;height:42px;border-radius:9px;overflow:hidden;vertical-align:middle"><img src="${logoUrl}" alt="Verpa" width="42" height="42" style="width:42px;height:42px;border-radius:9px;object-fit:cover;display:block" /></td><td style="padding-left:12px;vertical-align:middle"><div style="font-size:16px;font-weight:700;color:#ffffff;mso-color-alt:#ffffff">Verpa Support</div><div style="font-size:12px;color:#ffffff;mso-color-alt:#ffffff;opacity:.85">Ticketbeheer</div></td></tr></table>
    </td></tr></table>
    <div style="padding:24px">
      <div style="font-size:18px;font-weight:750;color:#111826;letter-spacing:-.3px;margin-bottom:6px">${title}</div>
      <div style="font-size:13.5px;color:#5b6677;line-height:1.55;margin-bottom:18px;mso-color-alt:#5b6677">${intro}</div>
      <div style="background:#f8fafc;border:1px solid #eef1f5;border-radius:10px;padding:14px 16px;margin-bottom:20px">
        ${head}
        ${rows.length?`<table style="width:100%;border-collapse:collapse">${rowsHtml}</table>`:""}
        ${extraTable}
      </div>
      <a href="${link}" style="display:inline-block;background:#f37a2b;color:#fff;text-decoration:none;font-size:13.5px;font-weight:600;padding:11px 20px;border-radius:9px">${ticket?"Ticket openen":"Support Desk openen"}</a>
      <div style="font-size:11.5px;color:#98a2b3;margin-top:22px;border-top:1px solid #eef1f5;padding-top:14px">Automatisch verstuurd door Verpa Support · Sales Support.</div>
    </div>
  </div></div>`;
}
function buildNewTicketEmail(t){ return emailShell("Nieuw ticket ingediend", `Er is een nieuw ticket aangemaakt door <b>${esc(t.author)}</b>. Als beheerder kun je het oppakken en toewijzen.`,
  [["Categorie", esc(t.category)+(t.subcategory?" · "+esc(t.subcategory):"")],["Prioriteit", prioLabel(t.priority)],["Status", statusLabel(t.status)],["Ingediend door", esc(t.author)],["Omschrijving", esc((t.description||"—").slice(0,200))]], t); }
function buildUpdateEmail(t, lines, attachments=[], internal=false){
  const linesHtml=lines.map(l=>`<tr><td style="padding:7px 0;color:#5b6677;font-size:13px;width:150px;vertical-align:top">Update</td><td style="padding:7px 0;color:#111826;font-size:13px;font-weight:600">${esc(l)}</td></tr>`).join("");
  const attHtml=attachments&&attachments.length?`<tr><td style="padding:7px 0;color:#5b6677;font-size:13px;width:150px;vertical-align:top">Bijlagen</td><td style="padding:7px 0;font-size:13px">${attachments.map(a=>`<div style="margin-bottom:4px"><a href="${a.webUrl}" style="color:#f37a2b;text-decoration:none;font-weight:500">📎 ${esc(a.name)}</a></div>`).join("")}</td></tr>`:"";
  const extraTable=`<table style="width:100%;border-collapse:collapse">
    ${linesHtml}
    <tr><td style="padding:7px 0;color:#5b6677;font-size:13px;width:150px;vertical-align:top">Huidige status</td><td style="padding:7px 0;color:#111826;font-size:13px;font-weight:600">${statusLabel(t.status)}</td></tr>
    <tr><td style="padding:7px 0;color:#5b6677;font-size:13px;width:150px;vertical-align:top">Prioriteit</td><td style="padding:7px 0;color:#111826;font-size:13px;font-weight:600">${prioLabel(t.priority)}</td></tr>
    <tr><td style="padding:7px 0;color:#5b6677;font-size:13px;width:150px;vertical-align:top">Behandelaar</td><td style="padding:7px 0;color:#111826;font-size:13px;font-weight:600">${esc(t.assignee||"—")}</td></tr>
    ${attHtml}
  </table>`;
  return internal
    ? emailShell("Nieuwe interne notitie", "Er is een interne notitie toegevoegd aan dit ticket. Deze melding gaat enkel naar het team.", [], t, extraTable)
    : emailShell("Er is een update op een ticket", "Een ticket dat je hebt ingediend, behandelt of volgt, is bijgewerkt.", [], t, extraTable); }
function buildConfirmEmail(t){ return emailShell("We hebben je ticket ontvangen", `Bedankt, <b>${esc(t.author)}</b>. Je ticket is geregistreerd. Via de knop hieronder volg je het gesprek op.`,
  [["Categorie", esc(t.category)+(t.subcategory?" · "+esc(t.subcategory):"")],["Prioriteit", prioLabel(t.priority)],["Status", statusLabel(t.status)]], t); }
function buildTestEmail(){ const n=SETTINGS.notifications; return emailShell("Testmail geslaagd", `Als je deze mail leest, werkt de mailverzending van Verpa Support. Verstuurd door <b>${esc(currentUser.name)}</b> vanuit de instellingen.`,
  [["Meldingen", n.enabled?"Ingeschakeld":"Uitgeschakeld"],["Verstuurd op", fmtDate(Date.now())]], null); }
function buildShareEmail(t, name){ return emailShell("Een ticket is met je gedeeld", `<b>${esc(name)}</b>, dit ticket is met je gedeeld zodat je mee kunt opvolgen. Je ontvangt voortaan ook updates.`,
  [["Categorie", esc(t.category)+(t.subcategory?" · "+esc(t.subcategory):"")],["Prioriteit", prioLabel(t.priority)],["Status", statusLabel(t.status)],["Ingediend door", esc(t.author)]], t); }
/* ---- Meldingen: wie krijgt welke mail (volgens Instellingen) ---- */
function assigneeEmail(name){
  if(!name||name==="Niet toegewezen") return "";
  const m=SETTINGS.team.find(x=>x.name===name);
  return m&&m.email?m.email.toLowerCase():"";
}
function eventRecipients(evKey, t, opts={}){
  const n=SETTINGS.notifications; if(!n.enabled) return [];
  const e=n.events[evKey]; if(!e) return [];
  const on=k=>e[k]&&!(opts.omit||[]).includes(k);
  const set=new Set(); const add=x=>{ x=(x||"").trim().toLowerCase(); if(x&&x.includes("@")) set.add(x); };
  if(on("beheerders")) adminEmails.forEach(add);
  if(on("indiener")) add(t.ownerUpn);
  if(on("behandelaar")) add(assigneeEmail(t.assignee));
  if(on("volgers")) (t.followers||[]).forEach(f=>add(f.upn));
  if(n.excludeActor) set.delete((currentUser.upn||"").toLowerCase());
  (opts.skip||[]).forEach(x=>set.delete((x||"").toLowerCase()));
  return [...set];
}
function notifyNewTicket(t){
  const n=SETTINGS.notifications; if(!n.enabled) return;
  const e=n.events.created;
  // Ontvangstbevestiging: gaat naar de indiener, ook al deed die zelf de actie
  const confirm=e.indiener&&t.ownerUpn;
  if(confirm) sendMail([t.ownerUpn], `Ticket ${t.ref} ontvangen: ${t.subject}`, buildConfirmEmail(t));
  const to=eventRecipients("created", t, {omit:["indiener"], skip:confirm?[t.ownerUpn]:[]});
  if(to.length) sendMail(to, `Nieuw ticket ${t.ref}: ${t.subject}`, buildNewTicketEmail(t));
}
function notifyUpdate(evKey, t, lines, attachments=[]){
  const to=eventRecipients(evKey, t); if(!to.length) return;
  const internal=evKey==="internal";
  sendMail(to, `${internal?"Interne notitie":"Update"} ticket ${t.ref}: ${t.subject}`, buildUpdateEmail(t, lines, attachments, internal));
}
function notifyStatus(t, lines){
  const e=SETTINGS.notifications.events.status;
  if(!(e.statuses||[]).includes(t.status)) return;
  notifyUpdate("status", t, lines);
}
function notifyShared(t, upn, name, lines){
  const n=SETTINGS.notifications; if(!n.enabled) return;
  upn=(upn||"").toLowerCase();
  // Kolom Volgers = enkel de nieuwe collega, met een uitnodigingsmail
  if(n.events.shared.volgers && !(n.excludeActor && upn===currentUser.upn)) sendMail([upn], `Ticket ${t.ref} met je gedeeld`, buildShareEmail(t, name));
  const others=eventRecipients("shared", t, {omit:["volgers"], skip:[upn]});
  if(others.length) sendMail(others, `Update ticket ${t.ref}: ${t.subject}`, buildUpdateEmail(t, lines));
}

/* ===================== INSTELLINGEN (enkel beheerders) ===================== */
/* Opgeslagen als JSON in de documentbibliotheek (CONFIG.settingsFile), zodat
   de instellingen gedeeld zijn en iedere gebruiker ze kan lezen: de mails
   worden verstuurd vanuit de browser van wie de wijziging doet.            */
const RECIPIENTS=[
  {k:"indiener",    label:"Indiener",    hint:"maakte het ticket aan"},
  {k:"behandelaar", label:"Behandelaar", hint:"toegewezen teamlid"},
  {k:"volgers",     label:"Volgers",     hint:"collega's met wie gedeeld"},
  {k:"beheerders",  label:"Beheerders",  hint:"lijst hieronder"}
];
const ALL_R=RECIPIENTS.map(r=>r.k);
const NOTIFY_EVENTS=[
  {k:"created",  label:"Nieuw ticket",             desc:"Een ticket wordt aangemaakt.",
   allow:["indiener","behandelaar","beheerders"], notes:{indiener:"ontvangstbevestiging", behandelaar:"indien meteen toegewezen", volgers:"nog geen volgers bij aanmaak"}},
  {k:"status",   label:"Status gewijzigd",         desc:"Enkel wanneer de nieuwe status hieronder aangeduid is.", allow:ALL_R, statusFilter:true},
  {k:"priority", label:"Prioriteit gewijzigd",     desc:"Laag, Gemiddeld of Hoog.", allow:ALL_R},
  {k:"assignee", label:"Behandelaar gewijzigd",    desc:"Een ticket krijgt een (andere) behandelaar.", allow:ALL_R, notes:{behandelaar:"de nieuwe behandelaar"}},
  {k:"message",  label:"Nieuw bericht",            desc:"Een zichtbaar antwoord in het gesprek, met eventuele bijlagen.", allow:ALL_R},
  {k:"internal", label:"Interne notitie",          desc:"Een notitie die enkel voor het team bedoeld is.",
   allow:["behandelaar","beheerders"], notes:{indiener:"nooit bij interne notities", volgers:"nooit bij interne notities"}},
  {k:"archive",  label:"Gearchiveerd of hersteld", desc:"Het ticket gaat naar of uit het archief.", allow:ALL_R},
  {k:"shared",   label:"Gedeeld met collega",      desc:"Iemand deelt het ticket met een collega.", allow:ALL_R, notes:{volgers:"enkel de nieuwe collega"}}
];
/* Standaardwaarden = het gedrag van v5, zodat er na de update niets verandert.
   Enige toevoeging: de nieuwe behandelaar krijgt een mail bij toewijzing. */
const DEFAULT_SETTINGS={
  notifications:{
    enabled:true, excludeActor:true,
    events:{
      created: {indiener:false,behandelaar:false,volgers:false,beheerders:true},
      status:  {indiener:true, behandelaar:false,volgers:true, beheerders:true, statuses:["open","progress","done","closed"]},
      priority:{indiener:true, behandelaar:false,volgers:true, beheerders:true},
      assignee:{indiener:true, behandelaar:true, volgers:true, beheerders:true},
      message: {indiener:true, behandelaar:false,volgers:true, beheerders:true},
      internal:{indiener:false,behandelaar:false,volgers:false,beheerders:false},
      archive: {indiener:true, behandelaar:false,volgers:true, beheerders:true},
      shared:  {indiener:false,behandelaar:false,volgers:true, beheerders:false}
    }
  },
  team:[
    {name:"Lukas Vanderheyden", email:"lukas@verpa.be"},
    {name:"Aniel Haeyaert",     email:"aniel@verpa.be"},
    {name:"Sten Huygens",       email:"sten.huygens@verpa.be"},
    {name:"Yana Verspreet",     email:""}
  ]
};
let SETTINGS=null, settingsETag=null, settingsLoadedAt=0, settingsLoadError=false, settingsDraft=null, settingsSaving=false;

function mergeSettings(raw){
  const d=DEFAULT_SETTINGS, r=raw||{}, n=r.notifications||{};
  const out={
    version:1,
    notifications:{
      enabled:      n.enabled!==undefined?!!n.enabled:d.notifications.enabled,
      excludeActor: n.excludeActor!==undefined?!!n.excludeActor:d.notifications.excludeActor,
      events:{}
    },
    team: Array.isArray(r.team)
      ? r.team.filter(m=>m&&String(m.name||"").trim()).map(m=>({name:String(m.name).trim(), email:String(m.email||"").trim().toLowerCase()}))
      : clone(d.team),
    adminEmails: Array.isArray(r.adminEmails)
      ? [...new Set(r.adminEmails.map(e=>String(e).trim().toLowerCase()).filter(Boolean))]
      : (CONFIG.adminEmails||[]).map(e=>e.toLowerCase()),
    updatedBy: r.updatedBy||"", updatedAt: r.updatedAt||0
  };
  NOTIFY_EVENTS.forEach(ev=>{
    const base=d.notifications.events[ev.k], got=(n.events||{})[ev.k]||{}, e={};
    ALL_R.forEach(k=>{ e[k]=ev.allow.includes(k) && (got[k]!==undefined?!!got[k]:base[k]); });
    if(ev.statusFilter){ const src=Array.isArray(got.statuses)?got.statuses:base.statuses; e.statuses=Object.keys(STATUS).filter(s=>src.includes(s)); }
    out.notifications.events[ev.k]=e;
  });
  return out;
}
SETTINGS=mergeSettings(null);

function applySettings(){ adminEmails=SETTINGS.adminEmails.slice(); }
function teamNames(){ return ["Niet toegewezen", ...SETTINGS.team.map(m=>m.name)]; }
function assigneeOptions(current){ const names=teamNames(); if(current&&!names.includes(current)) names.push(current); return names; }
function myTeamName(){ const m=SETTINGS.team.find(x=>(x.email&&x.email===currentUser.upn)||x.name===currentUser.name); return m?m.name:null; }
function settingsPath(){ return (CONFIG.settingsFile||`${CONFIG.attachFolder}/_instellingen/instellingen.json`).split("/").map(encodeURIComponent).join("/"); }

async function loadSettings(){
  try{
    const meta=await graph(`/sites/${SITE_ID}/drive/root:/${settingsPath()}?$select=id,eTag,@microsoft.graph.downloadUrl`);
    const res=await fetch(meta["@microsoft.graph.downloadUrl"],{cache:"no-store"});
    if(!res.ok) throw new Error("Download instellingen: "+res.status);
    SETTINGS=mergeSettings(JSON.parse(await res.text()));
    settingsETag=meta.eTag||null; settingsLoadError=false;
  }catch(e){
    if(String(e.message).startsWith("Graph 404")){ SETTINGS=mergeSettings(null); settingsETag=null; settingsLoadError=false; }
    else { settingsLoadError=true; console.warn("Instellingen konden niet geladen worden — vorige/standaardwaarden blijven actief:", e.message); }
  }
  settingsLoadedAt=Date.now(); applySettings();
}

const settingsKey=s=>JSON.stringify({n:s.notifications, t:s.team, a:s.adminEmails});
function settingsDirty(){ return !!settingsDraft && settingsKey(settingsDraft)!==settingsKey(SETTINGS); }
function validateSettings(s){
  const seen=new Set();
  for(const m of s.team){
    const n=(m.name||"").trim();
    if(!n) return "Elke behandelaar heeft een naam nodig";
    if(n.toLowerCase()==="niet toegewezen") return "\"Niet toegewezen\" kan geen naam van een behandelaar zijn";
    if(seen.has(n.toLowerCase())) return `Behandelaar "${n}" staat er twee keer in`;
    seen.add(n.toLowerCase());
    if(m.email&&!EMAIL_RE.test(m.email.trim())) return `Het e-mailadres van ${n} is ongeldig`;
  }
  const st=s.notifications.events.status;
  if(!st.statuses.length && ALL_R.some(k=>st[k])) return "Duid bij 'Status gewijzigd' minstens één status aan, of vink de ontvangers uit";
  return "";
}

async function saveSettings(){
  if(!isAdmin()||settingsSaving||!settingsDirty()) return;
  if(settingsLoadError){ toast("De instellingen konden niet geladen worden. Klik eerst op Vernieuwen."); return; }
  const err=validateSettings(settingsDraft); if(err){ toast(err); return; }
  const payload=mergeSettings(settingsDraft);
  payload.updatedBy=currentUser.name; payload.updatedAt=Date.now();
  settingsSaving=true; updateSaveBar();
  try{
    const headers=settingsETag?{"If-Match":settingsETag}:{};
    const blob=new Blob([JSON.stringify(payload,null,2)],{type:"application/json"});
    const item=await graph(`/sites/${SITE_ID}/drive/root:/${settingsPath()}:/content`,{method:"PUT",body:blob,headers});
    SETTINGS=payload; settingsETag=(item&&item.eTag)||null; applySettings();
    settingsDraft=clone(SETTINGS);
    settingsSaving=false; renderSettings(); renderSidebar(); toast("Instellingen opgeslagen");
  }catch(e){
    settingsSaving=false; updateSaveBar();
    console.error("[saveSettings]", e.message);
    if(String(e.message).startsWith("Graph 412")){
      if(confirm("Een andere beheerder heeft de instellingen intussen gewijzigd. De nieuwste versie laden? Je eigen wijzigingen gaan dan verloren.")){
        await loadSettings(); settingsDraft=clone(SETTINGS); renderSettings(); renderSidebar();
      }
    } else toast("Opslaan mislukt — controleer je rechten op de documentbibliotheek");
  }
}
function discardSettings(){ settingsDraft=clone(SETTINGS); renderSettings(); toast("Wijzigingen ongedaan gemaakt"); }

/* ---- Weergave ---- */
function renderSettings(){
  const el=document.getElementById("view");
  if(!isAdmin()){ el.innerHTML=`<div class="empty"><svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg><h3>Geen toegang</h3><p>Enkel beheerders kunnen de instellingen bekijken en wijzigen.</p></div>`; return; }
  if(!settingsDraft) settingsDraft=clone(SETTINGS);
  const meta=SETTINGS.updatedAt
    ? `Laatst gewijzigd door ${esc(SETTINGS.updatedBy||"onbekend")} op ${fmtDate(SETTINGS.updatedAt)}. Wijzigingen gelden voor iedereen.`
    : "Nog niet aangepast: de standaardwaarden zijn actief. Wijzigingen gelden voor iedereen.";
  el.innerHTML=`
    <div class="page-head"><div><h2>Instellingen</h2><div class="sub">${meta}</div></div></div>
    ${settingsLoadError?`<div class="set-alert"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><path d="M12 8v4M12 16h.01"/></svg><div>De opgeslagen instellingen konden niet geladen worden. Opslaan is uitgeschakeld tot je op <b>Vernieuwen</b> klikt, zodat je niets overschrijft.</div></div>`:""}
    <div class="set-stack">
      <section class="panel-card" id="setMail"></section>
      <section class="panel-card nm-card" id="setMatrix"></section>
      <div class="set-grid">
        <section class="panel-card" id="setTeam"></section>
        <section class="panel-card" id="setAdmins"></section>
      </div>
    </div>
    <div class="set-savebar" id="setSaveBar">
      <span class="msg" id="setSaveMsg"></span>
      <button class="btn btn-ghost btn-sm" id="setDiscardBtn" onclick="discardSettings()">Ongedaan maken</button>
      <button class="btn btn-primary btn-sm" id="setSaveBtn" onclick="saveSettings()">Instellingen opslaan</button>
    </div>`;
  renderSetMail(); renderSetMatrix(); renderSetTeam(); renderSetAdmins(); updateSaveBar();
}
function updateSaveBar(){
  const bar=document.getElementById("setSaveBar"); if(!bar) return;
  const dirty=settingsDirty();
  bar.classList.toggle("dirty",dirty);
  document.getElementById("setSaveMsg").textContent=settingsSaving?"Opslaan…":dirty?"Je hebt niet-opgeslagen wijzigingen":"Alle wijzigingen zijn opgeslagen";
  const sb=document.getElementById("setSaveBtn");
  sb.disabled=!dirty||settingsSaving||settingsLoadError; sb.textContent=settingsSaving?"Opslaan…":"Instellingen opslaan";
  document.getElementById("setDiscardBtn").disabled=!dirty||settingsSaving;
}
function cardHead(title,desc){ return `<div class="set-head"><h3>${title}</h3>${desc?`<p>${desc}</p>`:""}</div>`; }
function switchBtn(on,label,onclick){ return `<button type="button" role="switch" aria-checked="${on}" aria-label="${esc(label)}" class="sw ${on?"on":""}" onclick="${onclick}"><span class="switch"></span></button>`; }

function renderSetMail(){
  const n=settingsDraft.notifications;
  document.getElementById("setMail").innerHTML=`
    ${cardHead("E-mailmeldingen")}
    <div class="set-row"><div><div class="sr-t">Meldingen versturen</div><div class="sr-d">Hoofdschakelaar voor alle automatische mails. Zet uit tijdens onderhoud of tests.</div></div>
      ${switchBtn(n.enabled,"Meldingen versturen","setNotifyFlag('enabled',this)")}</div>
    <div class="set-row"><div><div class="sr-t">Geen mail over je eigen wijzigingen</div><div class="sr-d">Wie een wijziging doet, krijgt daar zelf geen mail over. De ontvangstbevestiging voor de indiener wordt wel verstuurd.</div></div>
      ${switchBtn(n.excludeActor,"Geen mail over je eigen wijzigingen","setNotifyFlag('excludeActor',this)")}</div>
    <div class="set-row"><div><div class="sr-t">Testmail</div><div class="sr-d">Controleer of de mailverzending werkt. De mail gaat naar ${esc(currentUser.upn)}.</div></div>
      <button class="btn btn-ghost btn-sm" onclick="sendTestMail(this)"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="m22 2-7 20-4-9-9-4Z"/><path d="M22 2 11 13"/></svg>Testmail sturen</button></div>`;
}
function renderSetMatrix(){
  const n=settingsDraft.notifications;
  const cell=(ev,r)=>{
    const note=(ev.notes||{})[r.k], lbl=`${ev.label}: ${r.label}`;
    if(!ev.allow.includes(r.k)) return `<td class="na" data-label="${r.label}"><span aria-label="${esc(lbl)} niet van toepassing">—</span>${note?`<small>${note}</small>`:""}</td>`;
    const on=n.events[ev.k][r.k];
    return `<td data-label="${r.label}"><label class="ck" data-label="${r.label}"><input type="checkbox" ${on?"checked":""} aria-label="${esc(lbl)}" onchange="setNotify('${ev.k}','${r.k}',this.checked)"><span class="box"></span></label>${note?`<small>${note}</small>`:""}</td>`;
  };
  const statusChips=()=>{ const sel=n.events.status.statuses; return `<div class="st-chips" role="group" aria-label="Statussen die een mail versturen">${Object.keys(STATUS).map(k=>{ const on=sel.includes(k); return `<button type="button" class="st-chip ${on?"on":""}" style="--c:${STATUS[k].color};--cb:${STATUS[k].bg}" aria-pressed="${on}" onclick="toggleNotifyStatus('${k}',this)">${STATUS[k].label}</button>`; }).join("")}</div>`; };
  const box=document.getElementById("setMatrix");
  box.classList.toggle("off",!n.enabled);
  box.innerHTML=`
    ${cardHead("Wie krijgt een mail bij welke wijziging?","Vink per wijziging aan welke personen die aan het ticket gekoppeld zijn een mail ontvangen. Iemand die in meerdere kolommen valt, krijgt één mail.")}
    <div class="nm-off-note"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><path d="M12 8v4M12 16h.01"/></svg>Meldingen staan uit. Er worden geen mails verstuurd, ongeacht de keuzes hieronder.</div>
    <div class="nm-wrap"><table class="nmatrix">
      <thead><tr><th class="ev" scope="col">Wijziging</th>${RECIPIENTS.map(r=>`<th scope="col">${r.label}<small>${r.hint}</small></th>`).join("")}</tr></thead>
      <tbody>${NOTIFY_EVENTS.map(ev=>`<tr><th class="ev" scope="row"><div class="evl">${ev.label}</div><div class="evd">${ev.desc}</div>${ev.statusFilter?statusChips():""}</th>${RECIPIENTS.map(r=>cell(ev,r)).join("")}</tr>`).join("")}</tbody>
    </table></div>`;
}
function renderSetTeam(){
  const team=settingsDraft.team;
  const rows=team.map((m,i)=>`
    <div class="team-row ${m.email?"":"noemail"}" id="tm_${i}">
      <input id="tmn_${i}" value="${esc(m.name)}" placeholder="Voor- en achternaam" aria-label="Naam behandelaar ${i+1}" oninput="setTeam(${i},'name',this.value)">
      <input id="tme_${i}" type="email" value="${esc(m.email)}" placeholder="naam@verpa.be" aria-label="E-mailadres behandelaar ${i+1}" oninput="setTeam(${i},'email',this.value)">
      <button class="icon-btn" title="Verwijderen" aria-label="Behandelaar ${esc(m.name||String(i+1))} verwijderen" onclick="removeTeam(${i})"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2m3 0v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6"/></svg></button>
      <div class="warn">Zonder e-mailadres krijgt deze persoon geen mails als behandelaar.</div>
    </div>`).join("");
  document.getElementById("setTeam").innerHTML=`
    ${cardHead("Behandelaars","Teamleden aan wie je tickets kunt toewijzen. Het e-mailadres wordt gebruikt voor de kolom Behandelaar. Een naam wijzigen past bestaande tickets niet aan.")}
    <div class="team-list">
      ${team.length?`<div class="team-row team-head"><span>Naam</span><span>E-mailadres</span><span></span></div>${rows}`:`<div class="set-empty">Nog geen behandelaars. Tickets kunnen dan enkel op "Niet toegewezen" staan.</div>`}
    </div>
    <div class="set-foot"><button class="btn btn-ghost btn-sm" onclick="addTeam()"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>Behandelaar toevoegen</button></div>`;
}
function renderSetAdmins(){
  const list=settingsDraft.adminEmails;
  document.getElementById("setAdmins").innerHTML=`
    ${cardHead("Beheerders","Ontvangers voor de kolom Beheerders. Wie beheerdersrechten heeft in de app, stel je in via de Azure AD-rol Admin.")}
    <div class="adm-list">${list.length
      ? `<div class="fchips">${list.map((e,i)=>`<span class="fchip">${esc(e)}<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" role="button" tabindex="0" aria-label="${esc(e)} verwijderen" onclick="removeAdminEmail(${i})" onkeydown="if(event.key==='Enter')removeAdminEmail(${i})"><path d="M18 6 6 18M6 6l12 12"/></svg></span>`).join("")}</div>`
      : `<div class="set-empty">Geen adressen. De kolom Beheerders verstuurt dan geen mails.</div>`}</div>
    <div class="adm-add"><input id="admAdd" type="email" placeholder="naam@verpa.be" aria-label="E-mailadres beheerder toevoegen" onkeydown="if(event.key==='Enter')addAdminEmail()"><button class="btn btn-ghost btn-sm" onclick="addAdminEmail()">Toevoegen</button></div>`;
}

/* ---- Interactie ---- */
function setNotify(ev,r,on){ settingsDraft.notifications.events[ev][r]=on; updateSaveBar(); }
function setNotifyFlag(flag,btn){
  const n=settingsDraft.notifications; n[flag]=!n[flag];
  btn.classList.toggle("on",n[flag]); btn.setAttribute("aria-checked",n[flag]);
  if(flag==="enabled"){ const m=document.getElementById("setMatrix"); if(m) m.classList.toggle("off",!n.enabled); }
  updateSaveBar();
}
function toggleNotifyStatus(k,btn){
  const e=settingsDraft.notifications.events.status, on=!e.statuses.includes(k);
  e.statuses=Object.keys(STATUS).filter(s=>s===k?on:e.statuses.includes(s));
  btn.classList.toggle("on",on); btn.setAttribute("aria-pressed",on); updateSaveBar();
}
function setTeam(i,key,v){
  settingsDraft.team[i][key]=key==="email"?v.trim().toLowerCase():v;
  if(key==="email"){ const r=document.getElementById("tm_"+i); if(r) r.classList.toggle("noemail",!v.trim()); }
  updateSaveBar();
}
function addTeam(){
  settingsDraft.team.push({name:"",email:""}); renderSetTeam(); updateSaveBar();
  const el=document.getElementById("tmn_"+(settingsDraft.team.length-1)); if(el) el.focus();
}
function removeTeam(i){
  const m=settingsDraft.team[i], name=(m.name||"").trim();
  const open=name?tickets.filter(t=>t.assignee===name&&t.status!=="closed"&&!t.archived).length:0;
  if(open&&!confirm(`${name} heeft nog ${open} actieve ticket${open===1?"":"s"}. Die blijven op deze naam staan tot je ze opnieuw toewijst. Toch verwijderen?`)) return;
  settingsDraft.team.splice(i,1); renderSetTeam(); updateSaveBar();
}
function addAdminEmail(){
  const el=document.getElementById("admAdd"), v=(el.value||"").trim().toLowerCase();
  if(!EMAIL_RE.test(v)){ toast("Vul een geldig e-mailadres in"); el.focus(); return; }
  if(settingsDraft.adminEmails.includes(v)){ toast("Dit adres staat al in de lijst"); return; }
  settingsDraft.adminEmails.push(v); renderSetAdmins(); updateSaveBar();
  document.getElementById("admAdd").focus();
}
function removeAdminEmail(i){ settingsDraft.adminEmails.splice(i,1); renderSetAdmins(); updateSaveBar(); }
async function sendTestMail(btn){
  btn.disabled=true;
  const ok=await sendMail([currentUser.upn], "Testmail Verpa Support", buildTestEmail());
  btn.disabled=false;
  toast(ok?`Testmail verstuurd naar ${currentUser.upn}`:"Testmail mislukt — controleer de mailWorker in CONFIG (details in de console)");
}
window.addEventListener("beforeunload", e=>{ if(view==="settings"&&settingsDirty()){ e.preventDefault(); e.returnValue=""; } });

/* ===================== DOCUMENTVIEWER ===================== */
function openViewer(name){
  document.getElementById("viewerTitle").textContent=name||"Document";
  document.getElementById("viewerDownload").href="#";
  document.getElementById("viewerBody").innerHTML=`<div class="viewer-loading">Voorbeeld laden…</div>`;
  document.getElementById("viewerOverlay").classList.add("show");
}
function closeViewer(){ document.getElementById("viewerOverlay").classList.remove("show"); document.getElementById("viewerBody").innerHTML=""; }
const IMG_EXT=["png","jpg","jpeg","gif","webp","bmp","svg"];
async function previewAtt(itemId, name){
  openViewer(name);
  const body=document.getElementById("viewerBody");
  try{
    const dl=await attachmentDownloadUrl(itemId);
    const dlBtn=document.getElementById("viewerDownload"); if(dl){ dlBtn.href=dl; }
    const ext=(name.split(".").pop()||"").toLowerCase();
    if(IMG_EXT.includes(ext)){ body.innerHTML=`<img src="${dl}" alt="${esc(name)}" />`; return; }
    const r=await graph(`/sites/${SITE_ID}/drive/items/${itemId}/preview`,{method:"POST",body:"{}"});
    const url=r&&r.getUrl;
    if(url){ body.innerHTML=`<iframe src="${url}" allow="fullscreen"></iframe>`; }
    else throw new Error("geen preview-url");
  }catch(e){
    body.innerHTML=`<div class="viewer-fallback">Voor dit bestandstype is geen ingebouwd voorbeeld beschikbaar.<br><a class="link" href="#" onclick="(async()=>{const u=await attachmentDownloadUrl('${itemId}');window.open(u,'_blank');})();return false;">Bestand downloaden</a></div>`;
  }
}

/* ===================== POLLING ===================== */
let pollTimer=null;
function startPolling(){
  if(pollTimer) clearInterval(pollTimer);
  pollTimer=setInterval(pollTickets, 15000);
  document.addEventListener("visibilitychange", ()=>{ if(!document.hidden) pollTickets(); });
}
function ticketSignature(){ return tickets.map(t=>t.itemId+"|"+t.status+"|"+t.assignee+"|"+t.archived+"|"+(t.messages?t.messages.length:0)).sort().join(","); }
async function pollTickets(){
  if(document.hidden || !currentUser || !LIST_ID) return;
  const modalOpen=document.getElementById("overlay").classList.contains("show");
  const viewerOpen=document.getElementById("viewerOverlay").classList.contains("show");
  const replyEditor=document.getElementById("replyEditor");
  const replyText=replyEditor?replyEditor.innerText.trim():"";
  try{
    if(view!=="settings" && Date.now()-settingsLoadedAt>5*60*1000) await loadSettings();
    const before=ticketSignature();
    await loadTickets();
    if(ticketSignature()===before) return;
    if(modalOpen || viewerOpen) return;
    if(view==="detail" && replyText) return;
    if(view==="dashboard"||view==="list"||view==="archive"){ render(); }
    else if(view==="detail" && detailTicket){ const fresh=findTicket(detailTicket.itemId); if(fresh){ detailTicket=fresh; renderDetail(); } }
  }catch(e){ /* stil negeren */ }
}

/* ===================== DRAG & DROP ===================== */
const drop=document.getElementById("drop");
["dragenter","dragover"].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add("over");}));
["dragleave","drop"].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove("over");}));
drop.addEventListener("drop",e=>{ if(e.dataTransfer.files.length) addNewFiles(e.dataTransfer.files); });

boot();
