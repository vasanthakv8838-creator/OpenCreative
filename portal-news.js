(() => {
"use strict";
const byId = id => document.getElementById(id);
const status = message => { const node=byId("feedStatus"); if(node) node.textContent=message; };
const maxFeedBytes = 1500000, maxFeeds = 20, maxItemsPerFeed = 40;
const safeText = value => String(value || "").replace(/\s+/g," ").trim();
const safeFeedUrl = value => {
  let url;
  try { url=new URL(String(value||"").trim()); } catch { throw new Error("Enter a valid feed URL."); }
  if(url.protocol!=="https:") throw new Error("For safety, feeds must use HTTPS.");
  if(url.username || url.password) throw new Error("Feed URLs cannot contain embedded credentials.");
  url.hash=""; return url.href;
};
const hash = value => { let h=2166136261; for(let i=0;i<value.length;i++){h^=value.charCodeAt(i);h=Math.imul(h,16777619);} return (h>>>0).toString(36); };
const escapeXml = value => String(value||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&apos;");
const download = (name, content, type) => { const blob=new Blob([content],{type:type}), url=URL.createObjectURL(blob), a=document.createElement("a"); a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1200); };
state.newsFeeds=Array.isArray(state.newsFeeds)?state.newsFeeds.filter(x=>{if(!x||typeof x.url!=="string")return false;try{return safeFeedUrl(x.url)===x.url;}catch{return false;}}).slice(0,maxFeeds):[];
state.newsCache=Array.isArray(state.newsCache)?state.newsCache.filter(x=>x&&typeof x.feedUrl==="string"&&Array.isArray(x.items)&&state.newsFeeds.some(feed=>feed.url===x.feedUrl)).slice(0,maxFeeds):[];
state.savedArticles=Array.isArray(state.savedArticles)?state.savedArticles.filter(x=>x&&typeof x.id==="string"&&typeof x.title==="string").slice(0,200):[];
const persist=()=>{try{save();return true;}catch(error){status("Could not save reader data. Export your portal backup or free browser storage.");return false;}};
const makeButton=(label,handler,disabled=false)=>{const b=document.createElement("button");b.type="button";b.textContent=label;b.disabled=disabled;b.addEventListener("click",handler);return b;};
const makeEmpty=message=>{const p=document.createElement("p");p.className="news-empty";p.textContent=message;return p;};
const safeLink=(value,base)=>{try{const url=new URL(value,base);return (url.protocol==="https:"||url.protocol==="http:")?url.href:"";}catch{return "";}};
const stripMarkup=value=>{const doc=new DOMParser().parseFromString(String(value||""),"text/html");doc.querySelectorAll("script,style,iframe,object,embed").forEach(node=>node.remove());return safeText(doc.body.textContent||"").slice(0,700);};
const firstText=(node,selectors)=>{for(const selector of selectors){const found=node.querySelector(selector);if(found&&safeText(found.textContent))return safeText(found.textContent);}return "";};
const parseFeed=(xmlText,feedUrl)=>{
  const xml=new DOMParser().parseFromString(xmlText,"application/xml");
  if(xml.querySelector("parsererror"))throw new Error("The response is not valid RSS or Atom XML.");
  const root=xml.documentElement, rssItems=[...xml.getElementsByTagName("item")], atomItems=[...xml.getElementsByTagName("entry")], nodes=rssItems.length?rssItems:atomItems;
  if(!nodes.length)throw new Error("No RSS items or Atom entries were found.");
  const channel=xml.querySelector("channel")||root, feedTitle=firstText(channel,["title"])||new URL(feedUrl).hostname;
  const items=nodes.slice(0,maxItemsPerFeed).map(node=>{
    const title=firstText(node,["title"])||"Untitled story";
    let link="";
    const atomLink=[...node.children].find(child=>child.localName==="link"&&(!child.getAttribute("rel")||child.getAttribute("rel")==="alternate"));
    if(atomLink)link=atomLink.getAttribute("href")||atomLink.textContent;
    if(!link)link=firstText(node,["link"]);
    const resolved=safeLink(link,feedUrl), guidNode=node.querySelector("guid")||node.querySelector("id"), guid=safeText(guidNode&&guidNode.textContent)||resolved||title;
    const date=firstText(node,["pubDate","published","updated","dc\\:date"]);
    const descriptionNode=node.querySelector("description,summary,content");
    let description=descriptionNode?descriptionNode.textContent:"";
    const encoded=node.getElementsByTagNameNS("http://purl.org/rss/1.0/modules/content/","encoded")[0];
    if(encoded)description=encoded.textContent;
    return {id:hash(feedUrl+"|"+guid+"|"+title),feedUrl,feedTitle,title,link:resolved,date:date||"",excerpt:stripMarkup(description)};
  }).filter(item=>item.link);
  return {title:feedTitle,items};
};
const readResponse=async response=>{
  const length=Number(response.headers.get("content-length")||0);
  if(length>maxFeedBytes)throw new Error("Feed is larger than 1.5 MB; use a smaller feed.");
  if(response.body&&response.body.getReader){
    const reader=response.body.getReader(),decoder=new TextDecoder("utf-8"),chunks=[];let total=0;
    while(true){const part=await reader.read();if(part.done)break;total+=part.value.byteLength;if(total>maxFeedBytes){try{await reader.cancel();}catch{}throw new Error("Feed exceeded the 1.5 MB size limit.");}chunks.push(decoder.decode(part.value,{stream:true}));}
    chunks.push(decoder.decode());return chunks.join("");
  }
  const text=await response.text();if(text.length>maxFeedBytes)throw new Error("Feed exceeded the 1.5 MB size limit.");return text;
};
const refreshFeed=async feed=>{
  const controller=new AbortController(),timeout=setTimeout(()=>controller.abort(),12000);
  try{
    const response=await fetch(feed.url,{method:"GET",mode:"cors",credentials:"omit",redirect:"follow",referrerPolicy:"no-referrer",cache:"no-store",headers:{Accept:"application/atom+xml, application/rss+xml, application/xml, text/xml;q=0.9, */*;q=0.5"},signal:controller.signal});
    if(!response.ok)throw new Error("Publisher returned HTTP "+response.status+".");
    if(response.url&&new URL(response.url).protocol!=="https:")throw new Error("Feed redirected to a non-HTTPS URL.");
    const parsed=parseFeed(await readResponse(response),feed.url),entry={feedUrl:feed.url,feedTitle:parsed.title,fetchedAt:new Date().toISOString(),items:parsed.items};
    const index=state.newsCache.findIndex(x=>x.feedUrl===feed.url);if(index>=0)state.newsCache[index]=entry;else state.newsCache.push(entry);
    feed.title=parsed.title;feed.lastFetched=entry.fetchedAt;feed.error="";persist();return parsed.items.length;
  }catch(error){
    feed.error=error.name==="AbortError"?"Request timed out.":(error.message==="Failed to fetch"?"This publisher may block browser access with CORS, or the network may be unavailable.":error.message);
    persist();throw new Error(feed.error);
  }finally{clearTimeout(timeout);}
};
const renderFeeds=()=>{
  const box=byId("feedList");if(!box)return;box.replaceChildren();
  if(!state.newsFeeds.length){box.append(makeEmpty("No feeds yet. Add an HTTPS RSS or Atom feed URL above."));return;}
  state.newsFeeds.forEach(feed=>{
    const row=document.createElement("div");row.className="feed-item";
    const info=document.createElement("div");info.style.minWidth="0";
    const title=document.createElement("b");title.textContent=feed.title||new URL(feed.url).hostname;
    const meta=document.createElement("div");meta.className="feed-meta";meta.textContent=feed.error||feed.url;info.append(title,meta);
    const actions=document.createElement("div");actions.className="article-actions";
    actions.append(makeButton("Refresh",async()=>{status("Refreshing "+(feed.title||feed.url)+"…");try{const count=await refreshFeed(feed);status("Loaded "+count+" stories.");}catch(error){status(error.message);}renderFeeds();renderArticles();}));
    actions.append(makeButton("Remove",()=>{if(!confirm("Remove this feed? Saved stories will remain in your reading list."))return;state.newsFeeds=state.newsFeeds.filter(x=>x.url!==feed.url);state.newsCache=state.newsCache.filter(x=>x.feedUrl!==feed.url);persist();renderFeeds();renderArticles();status("Feed removed. Saved stories were kept.");}));
    row.append(info,actions);box.append(row);
  });
};
const storyCard=(item,saved)=>{
  const card=document.createElement("article");card.className="article-card";
  const meta=document.createElement("div");meta.className="article-meta";meta.textContent=(item.feedTitle||"Feed")+(item.date?" · "+item.date:"");
  const heading=document.createElement("h3");
  const articleUrl=safeLink(item.link,"https://invalid.example/");
  if(articleUrl){const link=document.createElement("a");link.href=articleUrl;link.target="_blank";link.rel="noopener noreferrer";link.textContent=item.title;heading.append(link);}else heading.textContent=item.title;
  card.append(meta,heading);
  if(item.excerpt){const p=document.createElement("p");p.textContent=item.excerpt;card.append(p);}
  const actions=document.createElement("div");actions.className="article-actions";
  if(saved)actions.append(makeButton("Remove saved",()=>{state.savedArticles=state.savedArticles.filter(x=>x.id!==item.id);persist();renderArticles();}));
  else actions.append(makeButton(state.savedArticles.some(x=>x.id===item.id)?"Saved":"Save for later",()=>{if(!state.savedArticles.some(x=>x.id===item.id)){state.savedArticles.unshift(item);state.savedArticles=state.savedArticles.slice(0,200);persist();renderArticles();}},state.savedArticles.some(x=>x.id===item.id)));
  actions.append(makeButton("Copy link",()=>{if(navigator.clipboard&&navigator.clipboard.writeText)navigator.clipboard.writeText(item.link).then(()=>status("Story link copied."),()=>status("Could not copy link; open the story instead."));else status("Clipboard access is unavailable in this browser.");}));
  card.append(actions);return card;
};
const renderArticles=()=>{
  const all=state.newsCache.flatMap(cache=>cache.items||[]).sort((a,b)=>String(b.date).localeCompare(String(a.date)));
  const query=safeText(byId("articleSearch")?.value).toLowerCase(),savedOnly=!!byId("savedOnly")?.checked;
  const filtered=all.filter(item=>!query||(item.title+" "+item.excerpt+" "+item.feedTitle).toLowerCase().includes(query));
  const list=byId("articleList");if(list){list.replaceChildren();const visible=savedOnly?filtered.filter(item=>state.savedArticles.some(saved=>saved.id===item.id)):filtered.slice(0,100);if(!visible.length)list.append(makeEmpty(state.newsFeeds.length?"No matching stories yet. Refresh your feeds or change the search.":"Add a feed and refresh it to see stories."));else visible.forEach(item=>list.append(storyCard(item,false)));}
  const saved=byId("savedArticleList");if(saved){saved.replaceChildren();const savedItems=state.savedArticles.filter(item=>!query||(item.title+" "+item.excerpt+" "+item.feedTitle).toLowerCase().includes(query));if(!savedItems.length)saved.append(makeEmpty("Stories you save will appear here."));else savedItems.forEach(item=>saved.append(storyCard(item,true)));}
};
const form=byId("feedForm");
if(form)form.addEventListener("submit",async event=>{
  event.preventDefault();
  try{
    const url=safeFeedUrl(byId("feedUrl").value);
    if(state.newsFeeds.some(feed=>feed.url===url)){status("That feed is already in your list.");return;}
    if(state.newsFeeds.length>=maxFeeds){status("You can save up to "+maxFeeds+" feeds.");return;}
    const feed={url,title:new URL(url).hostname,addedAt:new Date().toISOString(),lastFetched:"",error:""};
    state.newsFeeds.push(feed);persist();byId("feedUrl").value="";renderFeeds();status("Feed added. Fetching stories…");
    const count=await refreshFeed(feed);status("Feed added; loaded "+count+" stories.");
  }catch(error){status(error.message);}
  renderFeeds();renderArticles();
});
if(byId("refreshFeeds"))byId("refreshFeeds").addEventListener("click",async()=>{
  if(!state.newsFeeds.length){status("Add at least one feed first.");return;}
  const button=byId("refreshFeeds");button.disabled=true;status("Refreshing feeds…");
  const results=await Promise.allSettled(state.newsFeeds.map(refreshFeed)),ok=results.filter(x=>x.status==="fulfilled").length,failed=results.length-ok;
  button.disabled=false;renderFeeds();renderArticles();status("Refresh complete: "+ok+" succeeded, "+failed+" failed. "+(failed?"Check each feed for its error.":""));
});
if(byId("articleSearch"))byId("articleSearch").addEventListener("input",renderArticles);
if(byId("savedOnly"))byId("savedOnly").addEventListener("change",renderArticles);
if(byId("exportOpml"))byId("exportOpml").addEventListener("click",()=>{
  const outlines=state.newsFeeds.map(feed=>'<outline text="'+escapeXml(feed.title||feed.url)+'" title="'+escapeXml(feed.title||feed.url)+'" type="rss" xmlUrl="'+escapeXml(feed.url)+'"/>').join("");
  const opml='<?xml version="1.0" encoding="UTF-8"?><opml version="2.0"><head><title>OpenCreative feeds</title></head><body>'+outlines+'</body></opml>';
  download("opencreative-feeds.opml",opml,"text/x-opml;charset=utf-8");status("Feed list exported as OPML.");
});
if(byId("importOpml"))byId("importOpml").addEventListener("click",()=>byId("opmlFile").click());
if(byId("opmlFile"))byId("opmlFile").addEventListener("change",async()=>{
  const file=byId("opmlFile").files&&byId("opmlFile").files[0];if(!file)return;
  try{
    if(file.size>500000)throw new Error("OPML file is larger than 500 KB.");
    const xml=new DOMParser().parseFromString(await file.text(),"application/xml");
    if(xml.querySelector("parsererror"))throw new Error("That file is not valid OPML/XML.");
    const urls=[...xml.getElementsByTagName("outline")].map(node=>node.getAttribute("xmlUrl")).filter(Boolean);
    if(!urls.length)throw new Error("No feed URLs were found in the OPML file.");
    let added=0;
    for(const value of urls){
      if(state.newsFeeds.length>=maxFeeds)break;
      try{const url=safeFeedUrl(value);if(!state.newsFeeds.some(feed=>feed.url===url)){state.newsFeeds.push({url,title:new URL(url).hostname,addedAt:new Date().toISOString(),lastFetched:"",error:""});added++;}}catch{}
    }
    persist();renderFeeds();status("Imported "+added+" new feeds. Refresh to load their stories.");
  }catch(error){status("OPML import failed: "+error.message);}
  finally{byId("opmlFile").value="";}
});
window.addEventListener("opencreative:portal-data-imported",()=>{renderFeeds();renderArticles();});
renderFeeds();renderArticles();
})();