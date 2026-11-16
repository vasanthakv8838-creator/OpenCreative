(() => {
"use strict";
const el = id => document.getElementById(id);
const announce = message => { const box=el("portalStatus"); if(box) box.textContent=message; };
const draftStatus = message => { const box=el("draftStatus"); if(box) box.textContent=message; };
const makeButton = (label, handler) => {
  const b=document.createElement("button"); b.type="button"; b.textContent=label; b.addEventListener("click",handler); return b;
};
const download = (name, content, type) => {
  const blob=new Blob([content],{type:type});
  const url=URL.createObjectURL(blob), a=document.createElement("a");
  a.href=url; a.download=name; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),1200);
};
const cleanHeader = value => String(value||"").replace(/[\r\n]+/g," ").trim();
const safeFile = value => String(value||"draft").replace(/[^a-z0-9_-]+/gi,"-").replace(/^-+|-+$/g,"").slice(0,48)||"draft";
state.drafts=Array.isArray(state.drafts)?state.drafts.filter(x=>x&&typeof x==="object").map((d,i)=>({
  id:String(d.id||("draft-"+i)),to:String(d.to||""),subject:String(d.subject||""),body:String(d.body||""),updatedAt:String(d.updatedAt||"")
})):[];
state.contacts=Array.isArray(state.contacts)?state.contacts.filter(x=>x&&typeof x==="object").map((c,i)=>({
  id:String(c.id||("contact-"+i)),name:String(c.name||""),email:String(c.email||"")
})).filter(c=>c.email):[];
save();
let editingDraftId=null;
const draftForm=el("draftForm");
const renderDrafts=()=>{
  const box=el("drafts"); if(!box)return;
  box.replaceChildren();
  if(!state.drafts.length){const empty=document.createElement("p");empty.className="muted small";empty.textContent="No saved drafts yet.";box.append(empty);return;}
  state.drafts.slice().sort((a,b)=>String(b.updatedAt).localeCompare(String(a.updatedAt))).forEach(d=>{
    const card=document.createElement("article");card.className="draft-item";
    const heading=document.createElement("b");heading.textContent=d.subject||"(No subject)";
    const recipient=document.createElement("div");recipient.className="small muted";recipient.textContent=d.to?"To: "+d.to:"No recipient";
    const preview=document.createElement("p");preview.className="small";preview.textContent=d.body.slice(0,180)+(d.body.length>180?"…":"");
    const actions=document.createElement("div");actions.className="draft-actions";
    actions.append(makeButton("Edit",()=>{editingDraftId=d.id;el("draftTo").value=d.to;el("draftSubject").value=d.subject;el("draftBody").value=d.body;draftStatus("Editing saved draft.");el("draftTo").focus()}));
    actions.append(makeButton("Download .eml",()=>downloadEml(d)));
    actions.append(makeButton("Delete",()=>{if(!confirm("Delete this saved draft?"))return;state.drafts=state.drafts.filter(x=>x.id!==d.id);if(editingDraftId===d.id){editingDraftId=null;draftForm.reset()}save();renderDrafts();draftStatus("Draft deleted.")}));
    card.append(heading,recipient,preview,actions);box.append(card);
  });
};
const readDraft=()=>({id:editingDraftId||("draft-"+Date.now()+"-"+state.drafts.length),to:cleanHeader(el("draftTo").value),subject:cleanHeader(el("draftSubject").value),body:String(el("draftBody").value||""),updatedAt:new Date().toISOString()});
const storeDraft=()=>{
  const draft=readDraft();
  if(!draft.subject&&!draft.body.trim()){draftStatus("Add a subject or message before saving.");return null;}
  const index=state.drafts.findIndex(x=>x.id===draft.id);
  if(index>=0)state.drafts[index]=draft;else state.drafts.push(draft);
  editingDraftId=draft.id;save();renderDrafts();draftStatus("Draft saved on this device.");return draft;
};
const downloadEml=d=>{
  const bytes=new TextEncoder().encode(d.subject||"");
  let binary="";bytes.forEach(n=>binary+=String.fromCharCode(n));
  const subject=/^[\x20-\x7e]*$/.test(d.subject||"")?(d.subject||""):"=?UTF-8?B?"+btoa(binary)+"?=";
  const body=String(d.body||"").replace(/\r?\n/g,"\r\n");
  const content="To: "+cleanHeader(d.to)+"\r\nSubject: "+subject+"\r\nMIME-Version: 1.0\r\nContent-Type: text/plain; charset=UTF-8\r\nContent-Transfer-Encoding: 8bit\r\nX-Unsent: 1\r\n\r\n"+body+"\r\n";
  download(safeFile(d.subject)+".eml",content,"message/rfc822;charset=utf-8");
};
if(draftForm)draftForm.addEventListener("submit",e=>{e.preventDefault();storeDraft()});
if(el("sendDraft"))el("sendDraft").addEventListener("click",()=>{
  const d=storeDraft();if(!d)return;
  if(!d.to){draftStatus("Add a recipient before opening your email app.");return;}
  const url="mailto:"+encodeURIComponent(d.to)+"?subject="+encodeURIComponent(d.subject)+"&body="+encodeURIComponent(d.body);
  window.location.href=url;draftStatus("Your email app should open with this message. OpenCreative has not sent it.");
});
if(el("newDraft"))el("newDraft").addEventListener("click",()=>{editingDraftId=null;draftForm.reset();draftStatus("New draft ready.")});
const contactForm=el("contactForm");
const renderContacts=()=>{
  const box=el("contacts");if(!box)return;box.replaceChildren();
  if(!state.contacts.length){const empty=document.createElement("p");empty.className="muted small";empty.textContent="No contacts saved.";box.append(empty);return;}
  state.contacts.forEach(c=>{
    const row=document.createElement("div");row.className="contact-item";
    const name=document.createElement("b");name.textContent=c.name||c.email;
    const email=document.createElement("div");email.className="small muted";email.textContent=c.email;
    const actions=document.createElement("div");actions.className="draft-actions";
    actions.append(makeButton("Use in draft",()=>{const field=el("draftTo");field.value=field.value?field.value+", "+c.email:c.email;el("mailDesk")?.scrollIntoView({behavior:"smooth",block:"start"});field.focus()}));
    actions.append(makeButton("Delete",()=>{if(!confirm("Delete this contact?"))return;state.contacts=state.contacts.filter(x=>x.id!==c.id);save();renderContacts()}));
    row.append(name,email,actions);box.append(row);
  });
};
if(contactForm)contactForm.addEventListener("submit",e=>{
  e.preventDefault();const name=el("contactName").value.trim(),email=el("contactEmail").value.trim();
  const existing=state.contacts.find(c=>c.email.toLowerCase()===email.toLowerCase());
  if(existing){existing.name=name;}else state.contacts.push({id:"contact-"+Date.now()+"-"+state.contacts.length,name,email});
  save();contactForm.reset();renderContacts();announce("Contact saved on this device.");
});
const icsEscape=value=>String(value||"").replace(/\\/g,"\\\\").replace(/\r?\n/g,"\\n").replace(/,/g,"\\,").replace(/;/g,"\\;");
const icsDate=value=>{const date=new Date(value);if(Number.isNaN(date.getTime()))return null;return date.toISOString().replace(/[-:]/g,"").replace(/\.\d{3}Z$/,"Z")};
if(el("exportCalendar"))el("exportCalendar").addEventListener("click",()=>{
  const lines=["BEGIN:VCALENDAR","VERSION:2.0","PRODID:-//OpenCreative//Portal Calendar//EN","CALSCALE:GREGORIAN"];
  state.events.forEach((event,i)=>{
    const start=icsDate(event.date);if(!start)return;
    const endDate=new Date(event.date);endDate.setMinutes(endDate.getMinutes()+30);
    const end=icsDate(endDate);if(!end)return;
    lines.push("BEGIN:VEVENT","UID:opencreative-"+i+"-"+Date.now()+"@local","DTSTAMP:"+icsDate(new Date()),"DTSTART:"+start,"DTEND:"+end,"SUMMARY:"+icsEscape(event.title),"END:VEVENT");
  });
  lines.push("END:VCALENDAR");
  download("opencreative-calendar.ics",lines.join("\r\n")+"\r\n","text/calendar;charset=utf-8");
  announce("Calendar exported as an .ics file.");
});
const backupButton=el("exportPortalBackup"),importButton=el("importPortalBackup"),backupFile=el("portalBackupFile");
if(backupButton)backupButton.addEventListener("click",()=>download("opencreative-portal-backup.json",JSON.stringify({...state,schemaVersion:3,exportedAt:new Date().toISOString()},null,2),"application/json"));
const mergeBy=(current,incoming,key)=>{
  const result=current.slice(),seen=new Set(current.map(key));
  incoming.forEach(item=>{const k=key(item);if(!seen.has(k)){seen.add(k);result.push(item)}});
  return result;
};
if(importButton&&backupFile)importButton.addEventListener("click",()=>backupFile.click());
if(backupFile)backupFile.addEventListener("change",async()=>{
  const file=backupFile.files&&backupFile.files[0];if(!file)return;
  try{
    const incoming=JSON.parse(await file.text());
    if(!incoming||typeof incoming!=="object"||!Array.isArray(incoming.events)||!Array.isArray(incoming.notes)||!Array.isArray(incoming.watchlist))throw new Error("This is not an OpenCreative portal backup.");
    if(!confirm("Merge this backup with the data already saved on this device? Existing items will be kept."))return;
    const events=incoming.events.filter(x=>x&&typeof x==="object"),notes=incoming.notes.filter(x=>x&&typeof x==="object");
    state.events=mergeBy(state.events,events,x=>String(x.title||"")+"|"+String(x.date||""));
    state.notes=mergeBy(state.notes,notes,x=>String(x.title||"")+"|"+String(x.body||""));
    state.watchlist=mergeBy(state.watchlist,incoming.watchlist.filter(x=>typeof x==="string"),x=>String(x).toUpperCase());
    if(Array.isArray(incoming.drafts))state.drafts=mergeBy(state.drafts,incoming.drafts.filter(x=>x&&typeof x==="object"),x=>String(x.id||"")+String(x.to||"")+String(x.subject||"")+String(x.body||""));
    if(Array.isArray(incoming.contacts))state.contacts=mergeBy(state.contacts,incoming.contacts.filter(x=>x&&typeof x==="object"),x=>String(x.email||"").toLowerCase());
    if(Array.isArray(incoming.newsFeeds))state.newsFeeds=mergeBy(Array.isArray(state.newsFeeds)?state.newsFeeds:[],incoming.newsFeeds.filter(x=>x&&typeof x.url==="string"),x=>String(x.url));
    if(Array.isArray(incoming.newsCache))state.newsCache=mergeBy(Array.isArray(state.newsCache)?state.newsCache:[],incoming.newsCache.filter(x=>x&&typeof x.feedUrl==="string"),x=>String(x.feedUrl));
    if(Array.isArray(incoming.savedArticles))state.savedArticles=mergeBy(Array.isArray(state.savedArticles)?state.savedArticles:[],incoming.savedArticles.filter(x=>x&&typeof x.id==="string"),x=>String(x.id));
    save();renderEvents();renderNotes();renderWatch();renderDrafts();renderContacts();window.dispatchEvent(new Event("opencreative:portal-data-imported"));announce("Backup merged. Existing local items were preserved.");
  }catch(error){announce("Import failed: "+error.message)}
  finally{backupFile.value=""}
});
renderDrafts();renderContacts();
})();