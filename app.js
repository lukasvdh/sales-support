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
  adminEmails:   ["lukas@verpa.be","sten.huygens@verpa.be","aniel@verpa.be"], // UPNs van alle beheerders
  mailWorker:  "https://verpa-mail-proxy.lukas-f22.workers.dev" // Cloudflare Worker voor mailverzending
};
/* ============================================================================ */

const TEAM = ["Niet toegewezen", "Lukas Vanderheyden", "Aniel Haeyaert", "Sten Huygens", "Yana Verspreet"];
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
let adminEmails=[]; // UPNs van alle beheerders — geladen uit CONFIG.adminEmails
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
const parseJson=(s,fb)=>{ try{ return s?JSON.parse(s):fb; }catch{ return fb; } };
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
async function loadAdminUpns(){
  adminEmails=(CONFIG.adminEmails||[]).map(e=>e.toLowerCase());
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
  console.log("Verpa Support Desk — build: v5 met logo + rijke teksteditor");
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
  currentUser={ name:account.name||claims.name||account.username, upn:(account.username||"").toLowerCase(), isAdmin:roles.includes(CONFIG.adminRole) };
  document.getElementById("root").innerHTML=`<div class="auth-wrap"><div class="auth-card"><div class="spinner"></div><p>Verbinden met SharePoint…</p></div></div>`;
  try{ await resolveIds(); loadAdminUpns(); await loadTickets(); showApp(); }
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
          <div><h1>Verpa Support</h1><p>Ticketbeheer · v5</p></div>
        </div></div>
        <nav class="sb-nav">
          <div class="nav-item" data-nav="dashboard" onclick="go('dashboard')"><svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="3" width="7" height="7" rx="1.5"/><rect x="3" y="14" width="7" height="7" rx="1.5"/><rect x="14" y="14" width="7" height="7" rx="1.5"/></svg>Dashboard</div>
          <div class="nav-item" onclick="openNew()"><svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>Nieuw Ticket</div>
          <div class="nav-item" data-nav="list" onclick="go('list')"><svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M3 12h18M3 18h18"/></svg>Alle Tickets</div>
          <div class="nav-item" data-nav="archive" onclick="go('archive')"><svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="4" width="18" height="4" rx="1"/><path d="M5 8v11a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V8"/><path d="M10 12h4"/></svg>Archief</div>
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
  document.getElementById("f_assignee").innerHTML=TEAM.map(n=>`<option>${esc(n)}</option>`).join("");
  view="dashboard"; currentId=null; render(); startPolling();
}
function toggleSidebar(force){ const sb=document.getElementById("sidebar"), sc=document.getElementById("sbScrim"); if(!sb)return; const open=force===undefined?!sb.classList.contains("open"):force; sb.classList.toggle("open",open); if(sc)sc.classList.toggle("show",open); }
async function refresh(){ toast("Vernieuwen…"); try{ await loadTickets(); render(); toast("Bijgewerkt"); }catch(e){ toast("Kon niet vernieuwen"); } }

/* ===================== VISIBILITY ===================== */
function isFollower(t){ return (t.followers||[]).some(x=>(x.upn||"").toLowerCase()===currentUser.upn); }
function allVisible(){ return isAdmin()?tickets.slice():tickets.filter(t=>t.ownerUpn===currentUser.upn||isFollower(t)); }
function visibleTickets(){ return allVisible().filter(t=>!t.archived); }
function archivedTickets(){ return allVisible().filter(t=>t.archived); }
function counts(){ const c={open:0,progress:0,done:0,closed:0}; visibleTickets().forEach(t=>c[t.status]!==undefined&&c[t.status]++); return c; }
function setNav(){ document.querySelectorAll(".nav-item[data-nav]").forEach(e=>e.classList.toggle("active", e.dataset.nav===view && !currentId)); }
function go(v){ view=v; currentId=null; if(document.getElementById("sbScrim"))toggleSidebar(false); render(); }
function render(){ if(!currentUser)return; renderSidebar(); setNav(); if(view==="dashboard")renderDashboard(); else if(view==="list")renderListView(); else if(view==="archive")renderArchive(); else if(view==="detail")renderDetail(); }
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
  if(TEAM.includes(currentUser.name)){ const mine=vt.filter(t=>t.assignee===currentUser.name&&t.status!=="closed").length; q+=`<div class="sb-link" onclick="quick('assignee','${esc(currentUser.name)}')"><span class="dot" style="background:var(--primary)"></span><span class="lbl">Aan mij toegewezen</span><span class="cnt">${mine}</span></div>`; }
  document.getElementById("sbQuick").innerHTML=q;
}
function quick(type,v){ filter={q:"",category:"",subcategory:"",status:"",assignee:"",special:""}; if(type==="status")filter.status=v; else if(type==="special")filter.special=v; else if(type==="assignee")filter.assignee=v; else if(type==="subcategory")filter.subcategory=v; view="list"; render(); }

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
      <select class="filter" id="fAssignee" onchange="filter.assignee=this.value;renderList()"><option value="">Alle behandelaars</option>${TEAM.map(n=>`<option>${esc(n)}</option>`).join("")}</select>
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
    notifyNewTicket(t);
    closeModal(); openDetail(t.itemId); toast(`Ticket ${ref} aangemaakt`);
  }catch(e){ toast("Opslaan mislukt — controleer je rechten"); console.error(e.message); }
  finally{ saving=false; btn.disabled=false; btn.textContent="Ticket aanmaken"; }
}

/* ===================== DETAIL ===================== */
function findTicket(itemId){ return tickets.find(t=>t.itemId===itemId); }
async function openDetail(itemId){
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
        ${canAssign?`<div class="upd-field"><label>Toegewezen aan (behandelaar)</label><select onchange="updateField('assignee',this.value)">${TEAM.map(n=>`<option ${t.assignee===n?"selected":""}>${esc(n)}</option>`).join("")}</select></div>`:`<div class="upd-field"><label>Behandelaar</label></div><div class="lockrow"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>${esc(t.assignee||"Niet toegewezen")} — enkel een beheerder kan dit wijzigen</div>`}
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
    if(!wasInt){ const snip=txt?(txt.replace(/<[^>]+>/g," ").slice(0,120)):(att.length?att.length+" bijlage(n) toegevoegd":""); notifyUpdate(detailTicket, ["Nieuw bericht van "+currentUser.name+": "+snip]); }
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
    if(prev!==value) notifyUpdate(detailTicket, [lbl+" gewijzigd van "+disp(prev)+" naar "+disp(value)+" door "+currentUser.name]);
  }
  catch(e){ detailTicket[field]=prev; renderDetail(); toast("Bijwerken mislukt"); }
}
async function toggleArchive(){
  detailTicket.archived=!detailTicket.archived;
  try{ await patchTicket(detailTicket.itemId,{[COL.Archived]:detailTicket.archived}); const a=detailTicket.archived; renderDetail(); renderSidebar(); toast(a?"Ticket gearchiveerd":"Ticket hersteld uit archief");
    notifyUpdate(detailTicket, [a?"Ticket gearchiveerd door "+currentUser.name:"Ticket hersteld uit archief door "+currentUser.name]); }
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
    sendMail([upn], `Ticket ${detailTicket.ref} met je gedeeld`, buildShareEmail(detailTicket, name||upn));
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
  if(!CONFIG.mailWorker){ console.warn("sendMail: mailWorker niet geconfigureerd in CONFIG"); return; }
  const recipients=(Array.isArray(toEmails)?toEmails:[toEmails]).filter(e=>e&&e.includes("@"));
  if(!recipients.length) return;
  try{
    const res=await fetch(CONFIG.mailWorker,{
      method:"POST",
      headers:{"Content-Type":"application/json"},
      body:JSON.stringify({ to:recipients, subject, html })
    });
    if(!res.ok){
      const txt=await res.text();
      console.warn("sendMail: Worker antwoordde",res.status,txt.slice(0,200));
    }
  }catch(e){
    // Mail-fouten mogen de app niet blokkeren — stil loggen
    console.warn("sendMail mislukt:",e.message);
  }
}
const statusLabel=k=>STATUS[k]?STATUS[k].label:k, prioLabel=k=>PRIO[k]?PRIO[k].label:k;
function emailShell(title, intro, rows, ticket){
  const rowsHtml=rows.map(r=>`<tr><td style="padding:7px 0;color:#5b6677;font-size:13px;width:150px;vertical-align:top">${r[0]}</td><td style="padding:7px 0;color:#111826;font-size:13px;font-weight:600">${r[1]}</td></tr>`).join("");
  const link=CONFIG.redirectUri||"#";
  return `<div style="margin:0;padding:24px;background:#eef1f4;font-family:Inter,Segoe UI,Arial,sans-serif">
  <div style="max-width:560px;margin:0 auto;background:#fff;border:1px solid #e7ebf0;border-radius:14px;overflow:hidden">
    <div style="background:linear-gradient(135deg,#0d8b80,#0a6f66);padding:20px 24px;color:#fff">
      <table><tr><td style="width:36px;height:36px;border-radius:9px;background:rgba(255,255,255,.18);text-align:center;font-weight:800;font-size:17px;color:#fff">V</td><td style="padding-left:12px"><div style="font-size:16px;font-weight:700">Verpa Support</div><div style="font-size:12px;opacity:.85">Ticketbeheer</div></td></tr></table>
    </div>
    <div style="padding:24px">
      <div style="font-size:18px;font-weight:750;color:#111826;letter-spacing:-.3px;margin-bottom:6px">${title}</div>
      <div style="font-size:13.5px;color:#5b6677;line-height:1.55;margin-bottom:18px">${intro}</div>
      <div style="background:#f8fafc;border:1px solid #eef1f5;border-radius:10px;padding:14px 16px;margin-bottom:20px">
        <div style="font-size:15px;font-weight:700;color:#111826;margin-bottom:4px">${esc(ticket.subject)}</div>
        <div style="font-size:12px;color:#98a2b3;margin-bottom:10px">${esc(ticket.ref)}</div>
        <table style="width:100%;border-collapse:collapse">${rowsHtml}</table>
      </div>
      <a href="${link}" style="display:inline-block;background:#0d8b80;color:#fff;text-decoration:none;font-size:13.5px;font-weight:600;padding:11px 20px;border-radius:9px">Ticket openen</a>
      <div style="font-size:11.5px;color:#98a2b3;margin-top:22px;border-top:1px solid #eef1f5;padding-top:14px">Automatisch verstuurd door Verpa Support · Sales Support.</div>
    </div>
  </div></div>`;
}
function buildNewTicketEmail(t){ return emailShell("Nieuw ticket ingediend", `Er is een nieuw ticket aangemaakt door <b>${esc(t.author)}</b>. Als beheerder kun je het oppakken en toewijzen.`,
  [["Categorie", esc(t.category)+(t.subcategory?" · "+esc(t.subcategory):"")],["Prioriteit", prioLabel(t.priority)],["Status", statusLabel(t.status)],["Ingediend door", esc(t.author)],["Omschrijving", esc((t.description||"—").slice(0,200))]], t); }
function buildUpdateEmail(t, lines){ const ch=lines.map(l=>`<div style="font-size:13px;color:#111826;padding:5px 0;border-bottom:1px solid #f0f2f5">• ${esc(l)}</div>`).join("");
  return emailShell("Er is een update op je ticket", `Het ticket dat je hebt ingediend of volgt, is bijgewerkt:<div style="margin-top:12px">${ch}</div>`,
  [["Huidige status", statusLabel(t.status)],["Prioriteit", prioLabel(t.priority)],["Behandelaar", esc(t.assignee||"—")]], t); }
function buildShareEmail(t, name){ return emailShell("Een ticket is met je gedeeld", `<b>${esc(name)}</b>, dit ticket is met je gedeeld zodat je mee kunt opvolgen. Je ontvangt voortaan ook updates.`,
  [["Categorie", esc(t.category)+(t.subcategory?" · "+esc(t.subcategory):"")],["Prioriteit", prioLabel(t.priority)],["Status", statusLabel(t.status)],["Ingediend door", esc(t.author)]], t); }
function allTicketRecipients(t){
  const set=new Set();
  adminEmails.forEach(e=>{ if(e) set.add(e); });
  if(t.ownerUpn) set.add(t.ownerUpn.toLowerCase());
  (t.followers||[]).forEach(f=>{ if(f.upn) set.add(f.upn.toLowerCase()); });
  set.delete((currentUser.upn||"").toLowerCase());
  return [...set];
});
  if(t.ownerUpn) set.add(t.ownerUpn.toLowerCase());
  (t.followers||[]).forEach(f=>{ if(f.upn) set.add(f.upn.toLowerCase()); });
  // Huidige gebruiker ontvangt geen mail over zijn eigen actie
  set.delete((currentUser.upn||"").toLowerCase());
  return [...set];
}
function notifyNewTicket(t){
  const toAdmins=adminEmails.filter(e=>e!==(currentUser.upn||"").toLowerCase());
  if(toAdmins.length) sendMail(toAdmins, `Nieuw ticket ${t.ref}: ${t.subject}`, buildNewTicketEmail(t));
}: ${t.subject}`, buildNewTicketEmail(t));
}
function notifyUpdate(t, lines){
  const to=allTicketRecipients(t);
  if(!to.length) return;
  sendMail(to, `Update ticket ${t.ref}: ${t.subject}`, buildUpdateEmail(t, lines));
}: ${t.subject}`, buildUpdateEmail(t, lines));
}

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
