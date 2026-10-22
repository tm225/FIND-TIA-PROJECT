'use strict';
const $=s=>document.querySelector(s),$$=s=>[...document.querySelectorAll(s)];
const esc=s=>s.replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
const norm=s=>s.replace(/[\u2018\u2019\u00a0]/g,m=>m=='\u00a0'?' ':"'");
const fold=s=>norm(s).replace(/[^\x00-\x7f]/g,c=>{const d=c.normalize('NFD');return d.length>1&&/[\u0300-\u036f]/.test(d[1])?d[0]:c});
const col=n=>(n>25?col(Math.floor(n/26)-1):'')+String.fromCharCode(65+n%26);
const S={docs:[],res:[],n:150,i:0,rx:null};

/* ---------- Stockage hors ligne (IndexedDB) ---------- */
const DB={o:null,
 init:()=>new Promise((ok,ko)=>{const r=indexedDB.open('findtia',1);r.onupgradeneeded=()=>r.result.createObjectStore('d',{keyPath:'id'});r.onsuccess=()=>{DB.o=r.result;ok()};r.onerror=ko}),
 tx:(m,f)=>new Promise((ok,ko)=>{const t=DB.o.transaction('d',m),q=f(t.objectStore('d'));t.oncomplete=()=>ok(q&&q.result);t.onerror=t.onabort=()=>ko(t.error)}),
 all:()=>DB.tx('readonly',s=>s.getAll()),put:d=>DB.tx('readwrite',s=>s.put(d)),del:id=>DB.tx('readwrite',s=>s.delete(id)),clr:()=>DB.tx('readwrite',s=>s.clear())};

/* ---------- Lecture des fichiers Office ---------- */
const ln=n=>n.localName,kids=(n,l)=>[...n.childNodes].filter(c=>c.nodeType==1&&(!l||ln(c)==l)),all=n=>[...n.getElementsByTagName('*')];
const xml=async(z,p)=>z.file(p)?new DOMParser().parseFromString(await z.file(p).async('string'),'application/xml'):null;
const ptx=el=>{let s='';(function w(n){for(const c of n.childNodes){if(c.nodeType!=1)continue;const l=ln(c);if(l=='t')s+=c.textContent;else if(l=='tab')s+='\t';else if(l=='br'||l=='cr')s+='\n';else if(!['p','txbxContent','Fallback','delText','instrText'].includes(l))w(c)}})(el);return s.trim()};

function sheets(wb,u,h,m){
 const X=XLSX.utils;
 for(const n of wb.SheetNames){const ws=wb.Sheets[n];let a=1e9,b=-1,c=1e9,d=-1;
  for(const k in ws){if(k[0]=='!'||(ws[k].t=='z'&&!ws[k].f))continue;const x=X.decode_cell(k);a=Math.min(a,x.r);b=Math.max(b,x.r);c=Math.min(c,x.c);d=Math.max(d,x.c)}
  if(b<0)continue;
  m[n]={af:ws['!autofilter']&&ws['!autofilter'].ref,hc:(ws['!cols']||[]).map((x,i)=>x&&x.hidden?i:-1).filter(i=>i>=0)};
  for(let r=a;r<=b;r++){const cells=[],fx=[],nt=[];let any=0;
   for(let j=c;j<=d;j++){const z=ws[X.encode_cell({r,c:j})];let v='',f='',t='';
    if(z){v=String(z.w!=null?z.w:z.v!=null?z.v:'').trim();if(z.f)f='='+z.f;t=[(z.c||[]).map(o=>o.t).join(' '),z.l&&z.l.Target||''].filter(Boolean).join(' ').trim()}
    cells.push(v);fx.push(f);nt.push(t);if(v||f||t)any=1}
   if(!any)continue;
   while(cells.length&&!cells[cells.length-1]&&!fx[cells.length-1]&&!nt[cells.length-1]){cells.pop();fx.pop();nt.pop()}
   if(!h[n])h[n]=cells;
   const hid=!!(ws['!rows']&&ws['!rows'][r]&&ws['!rows'][r].hidden);
   u.push({k:'x',loc:n+' · Ligne '+(r+1),sh:n,c0:c,cells,fx,nt,hid,t:cells.flatMap((v,i)=>[v,fx[i],nt[i]]).filter(Boolean).join(' | ')})}}}

async function parse(f){
 const buf=await f.arrayBuffer(),e=f.name.split('.').pop().toLowerCase(),u=[],h={};let ty,lab,tabs,meta={};
 if(/^(xlsx|xlsm|xlsb|xls|ods)$/.test(e)){ty='Excel';lab='lignes';const wb=XLSX.read(buf,{type:'array',cellStyles:true,sheetStubs:true});tabs=wb.SheetNames;sheets(wb,u,h,meta);
  if(/^xls[xm]$/.test(e)){try{const z=await JSZip.loadAsync(buf);for(const p of Object.keys(z.files).filter(n=>/^xl\/drawings\/drawing\d+\.xml$/.test(n))){const t=all(await xml(z,p)).filter(x=>ln(x)=='p').map(ptx).filter(Boolean);if(t.length)u.push({k:'p',loc:'Dessin / zone de texte ('+p.match(/\d+/)[0]+')',t:t.join('\n')})}}catch(_){}}}
 else if(e=='csv'){ty='Excel';lab='lignes';const wb=XLSX.read(new TextDecoder().decode(buf),{type:'string',sheetStubs:true});tabs=wb.SheetNames;sheets(wb,u,h,meta)}
 else if(/^(docx|docm|dotx)$/.test(e)){ty='Word';lab='éléments';const z=await JSZip.loadAsync(buf),
  L={header:'En-tête',footer:'Pied de page',footnotes:'Notes',endnotes:'Notes de fin',comments:'Commentaires'};
  for(const p of Object.keys(z.files).filter(n=>/^word\/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$/.test(n))){
   const d=await xml(z,p),k=p.slice(5,-4).replace(/\d+$/,''),pre=k=='document'?'':L[k]+' · ';let pn=0,tn=0;
   const walk=n=>{for(const c of kids(n)){const l=ln(c);
    if(l=='p'){const t=ptx(c);if(t)u.push({k:'p',loc:pre+'Paragraphe '+(++pn),t});all(c).filter(x=>ln(x)=='txbxContent').forEach(walk)}
    else if(l=='tbl'){let r=0;tn++;for(const tr of kids(c,'tr')){r++;const cells=kids(tr,'tc').map(tc=>all(tc).filter(x=>ln(x)=='p').map(ptx).filter(Boolean).join(' '));
     if(cells.some(Boolean))u.push({k:'t',loc:pre+'Tableau '+tn+' · Ligne '+r,cells,t:cells.filter(Boolean).join(' | ')})}}
    else if(l!='Fallback')walk(c)}};
   walk(d.documentElement)}}
 else if(/^(pptx|pptm|ppsx|potx)$/.test(e)){ty='PowerPoint';lab='diapositives';const z=await JSZip.loadAsync(buf),map={};
  all(await xml(z,'ppt/_rels/presentation.xml.rels')).forEach(r=>map[r.getAttribute('Id')]=r.getAttribute('Target'));
  const ord=all(await xml(z,'ppt/presentation.xml')).filter(x=>ln(x)=='sldId').map(x=>'ppt/'+map[x.getAttribute('r:id')].replace(/^\/?(ppt\/)?/,''));
  const P=async(path,ok)=>{const d=await xml(z,path);return d?all(d).filter(x=>ln(x)=='p').map(ptx).filter(t=>t&&ok(t)):[]};
  for(let i=0;i<ord.length;i++){const t=await P(ord[i],()=>1);if(t.length)u.push({k:'s',loc:'Diapositive '+(i+1),t:t.join('\n')});
   const rels=await xml(z,ord[i].replace('slides/','slides/_rels/')+'.rels'),nr=rels&&all(rels).find(r=>/notesSlide/.test(r.getAttribute('Target')));
   if(nr){const nt=await P('ppt/notesSlides/'+nr.getAttribute('Target').split('/').pop(),x=>!/^(\d+|‹#›)$/.test(x));if(nt.length)u.push({k:'s',loc:'Diapositive '+(i+1)+' · Notes',t:nt.join('\n')})}}}
 else if(/^(txt|md|log)$/.test(e)){ty='Texte';lab='lignes';new TextDecoder().decode(buf).split(/\r?\n/).forEach((t,i)=>{if(t.trim())u.push({k:'p',loc:'Ligne '+(i+1),t})})}
 else throw new Error('format non pris en charge (les anciens .doc et .ppt doivent être enregistrés en .docx / .pptx)');
 if(!u.length)throw new Error('aucun texte trouvé');
 return{id:Date.now()+Math.random(),name:f.name,size:f.size,type:ty,lab,units:u,heads:h,tabs,meta,added:Date.now()}}

/* ---------- Recherche ---------- */
const txt=(u,a)=>a?(u.n??=norm(u.t)):(u.f??=fold(u.t));
function build(q){
 const c=$('#oc').checked,w=$('#ow').checked,a=$('#oa').checked,m=$('#om').value,E=s=>s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');q=(a?norm(q):fold(q)).trim();
 let p=m=='regex'?[q]:m=='words'?q.split(/\s+/).filter(Boolean).map(E):m=='wild'?[[...q].map(x=>x=='*'?'\\S*':x=='?'?'\\S':x==' '?'\\s+':E(x)).join('')]:[E(q).replace(/ +/g,'\\s+')];
 if(w){const W='[A-Za-z0-9_\\u00C0-\\u024F]';p=p.map(x=>`(?<!${W})(?:${x})(?!${W})`)}
 const f=c?'':'i';return{t:p.map(x=>new RegExp(x,f)),h:new RegExp(p.join('|'),'g'+f),a}}
function rng(s){const z=S.rx.a?norm(s):fold(s),o=[],r=S.rx.h;r.lastIndex=0;let m;while((m=r.exec(z))){if(!m[0]){r.lastIndex++;continue}o.push([m.index,m.index+m[0].length])}return o}
function mk(s,from=0,to=s.length){let h='',p=from;for(const[a,b]of rng(s)){if(b<=from||a>=to)continue;const x=Math.max(a,from),y=Math.min(b,to);h+=esc(s.slice(p,x))+'<mark>'+esc(s.slice(x,y))+'</mark>';p=y}return h+esc(s.slice(p,to))}
function snip(u){const r=rng(u.t)[0]||[0,0],a=Math.max(0,r[0]-50),b=Math.min(u.t.length,r[1]+110);return(a?'…':'')+mk(u.t,a,b)+(b<u.t.length?'…':'')}
function run(){
 S.res=[];S.rx=null;S.n=150;S.scan=0;const q=$('#q').value.trim();
 if(q){try{S.rx=build(q)}catch(e){$('#info').textContent='Expression invalide';$('#list').innerHTML='';return}
  const sc=$('#os').value;for(const d of S.docs){if(sc!='all'&&d.id!=sc)continue;const ot=sc!='all'?$('#ot').value:'all';d.units.forEach((u,i)=>{if(ot!='all'&&u.sh!=ot)return;if($('#oh').checked&&u.hid)return;S.scan++;const z=txt(u,S.rx.a);if(S.rx.t.every(r=>r.test(z)))S.res.push([d,i])})}}
 draw()}
const locOf=u=>u.loc+(u.hid?' · ligne masquée':'');
function draw(){
 const L=$('#list'),I=$('#info'),q=$('#q').value.trim();
 if(!S.docs.length){I.textContent='';L.innerHTML='<div class="empty">Aucun document en mémoire.<br>Ajoutez d’abord un fichier dans « Documents ».</div>';return}
 if(!q){I.textContent='';L.innerHTML='<div class="empty">Tapez ce que vous cherchez.<br>La recherche se fait dans tous vos documents.</div>';return}
 if(!S.res.length){I.textContent='';const o=[];if($('#oc').checked)o.push('Respecter la casse');if($('#ow').checked)o.push('Mot entier');if($('#oh').checked)o.push('Ignorer les lignes masquées');if($('#oa').checked)o.push('Respecter les accents');if($('#om').value!='phrase')o.push($('#om').selectedOptions[0].text.split(' :')[0]);if($('#os').value!='all')o.push('un seul document'+(!$('#ot').hidden&&$('#ot').value!='all'?' (onglet « '+$('#ot').value+' » seulement)':''));
 L.innerHTML='<div class="empty">Aucun résultat pour « '+esc(q)+' » parmi '+S.scan+' éléments lus.'+(o.length?'<br><br>Options actives : '+esc(o.join(', '))+'<button class="fab" id="rst" style="margin-top:12px">Réinitialiser les options</button>':'')+'</div>';return}
 I.textContent=S.res.length+' résultat'+(S.res.length>1?'s':'')+' dans '+new Set(S.res.map(r=>r[0].id)).size+' document(s)';
 let h='',last;S.res.slice(0,S.n).forEach(([d,i],j)=>{if(d!==last){h+='<h3>'+esc(d.name)+'</h3>';last=d}h+='<button class="res" data-j="'+j+'"><b>'+esc(locOf(d.units[i]))+'</b><span>'+snip(d.units[i])+'</span></button>'});
 if(S.res.length>S.n)h+='<button class="fab" id="more">Afficher plus</button>';
 L.innerHTML=h}
function show(j){
 S.i=j;const[d,i]=S.res[j],u=d.units[i],hd=d.heads[u.sh];let h;
 $('#dt').textContent=d.name;$('#dsub').textContent=locOf(u)+'  ·  '+(j+1)+' / '+S.res.length;
 const mt=(d.meta&&d.meta[u.sh])||{};
 h=u.cells?(u.hid?'<div class="note">Ligne masquée (par un filtre ou manuellement)</div>':'')+(mt.af?'<div class="note">Filtre automatique sur la plage '+esc(mt.af)+'</div>':'')+u.cells.map((v,x)=>{const f=u.fx&&u.fx[x],n=u.nt&&u.nt[x];if(!(v||f||n))return'';
  const lb=u.k=='x'?col(u.c0+x)+(hd&&hd!==u.cells&&hd[x]?' · '+esc(hd[x]):'')+((mt.hc||[]).includes(u.c0+x)?' · colonne masquée':''):'Cellule '+(x+1);
  return'<div class="cell"><small>'+lb+'</small>'+(v?mk(v):'<i>(vide)</i>')+(f?'<em class="fx">'+mk(f)+'</em>':'')+(n?'<em class="nt">'+mk(n)+'</em>':'')+'</div>'}).join(''):'<p class="big">'+mk(u.t)+'</p>';
 const b=$('#db');b.innerHTML=h;b.scrollTop=0;b.classList.remove('f');void b.offsetWidth;b.classList.add('f');$('#det').classList.add('on')}

/* ---------- Écrans, dialogues, quitter ---------- */
const ORD=['home','search','docs'];let cur='home',cancel=null;
function go(n){if(!ORD.includes(n))return;cur=n;document.body.dataset.s=n;ORD.forEach((s,i)=>{const e=$('#'+s);e.className='scr'+(s==n?' on':i<ORD.indexOf(n)?' l':'')});
 $$('nav button[data-go]').forEach(b=>b.classList.toggle('on',b.dataset.go==n));
 if(n=='home'){const h=$('#home');h.classList.remove('play');void h.offsetWidth;h.classList.add('play')}}
function ask(t,m,ok,bad){return new Promise(r=>{$('#dtt').textContent=t;$('#dm').textContent=m;const b=$('#dok');b.textContent=ok;b.className=bad?'bad':'';$('#dlg').classList.add('on');
 const f=v=>{$('#dlg').classList.remove('on');cancel=null;r(v)};b.onclick=()=>f(true);$('#dno').onclick=()=>f(false);cancel=()=>f(false)})}
async function quit(){if(await ask('Quitter FIND TIA ?','Vos documents restent en mémoire jusqu’à ce que vous les supprimiez.','Quitter',1)){document.body.classList.add('bye');setTimeout(()=>{const A=window.Capacitor&&Capacitor.Plugins&&Capacitor.Plugins.App;A?A.exitApp():window.close()},450)}}
function toast(m){const t=$('#toast');t.textContent=m;t.classList.add('on');setTimeout(()=>t.classList.remove('on'),4000)}
const wait=ms=>new Promise(r=>setTimeout(r,ms));

/* ---------- Documents ---------- */
function fill(){
 const kb=n=>n>1048576?(n/1048576).toFixed(1)+' Mo':Math.ceil(n/1024)+' Ko';
 $('#dl').innerHTML=S.docs.length?S.docs.map(d=>'<div class="doc" data-id="'+d.id+'"><i class="ty '+d.type+'">'+d.type+'</i><div><b>'+esc(d.name)+'</b><small>'+d.units.length+' '+d.lab+(d.tabs?' · '+d.tabs.length+' onglet'+(d.tabs.length>1?'s':''):'')+' · '+kb(d.size)+' · '+new Date(d.added).toLocaleDateString('fr-FR')+'</small></div><button class="del" data-del="'+d.id+'">Supprimer</button></div>').join('')
  +'<button class="lnk" id="delall">Tout supprimer</button>':'<div class="empty">Aucun document.<br>Ajoutez des fichiers Word, Excel ou PowerPoint : ils restent ici, même hors ligne, jusqu’à ce que vous les supprimiez.</div>';
 $('#hc').textContent=S.docs.length?S.docs.length+' document'+(S.docs.length>1?'s':'')+' en mémoire':'Aucun document en mémoire';
 const s=$('#os'),v=s.value;s.innerHTML='<option value="all">Tous les documents</option>'+S.docs.map(d=>'<option value="'+d.id+'">'+esc(d.name)+'</option>').join('');s.value=S.docs.some(d=>d.id==v)?v:'all';tabSel()}
function tabSel(){const d=S.docs.find(x=>x.id==$('#os').value),t=$('#ot');t.hidden=!(d&&d.tabs&&d.tabs.length>1);t.innerHTML='<option value="all">Tous les onglets</option>'+(d&&d.tabs?d.tabs.map(n=>'<option>'+esc(n)+'</option>').join(''):'')}
async function add(files){
 const B=$('#busy');B.classList.add('on');
 for(const f of files){$('#bt').textContent='Lecture de '+f.name+'…';await wait(60);
  try{const d=await parse(f),o=S.docs.find(x=>x.name==d.name&&x.size==d.size);if(o){await DB.del(o.id);S.docs=S.docs.filter(x=>x!=o)}await DB.put(d);S.docs.push(d)}catch(e){toast(f.name+' : '+(e.message||e))}}
 B.classList.remove('on');fill();if($('#q').value.trim())run()}

/* ---------- Démarrage ---------- */
document.addEventListener('click',async e=>{
 const t=e.target.closest('button');if(!t)return;
 if(t.dataset.go)go(t.dataset.go);
 else if('q' in t.dataset)quit();
 else if(t.id=='add')$('#fi').click();
 else if(t.id=='cl'){$('#q').value='';run();$('#q').focus()}
 else if(t.id=='rst'){$$('#oc,#ow,#oa,#oh').forEach(x=>x.checked=false);$('#om').value='phrase';$('#os').value='all';tabSel();run()}
 else if(t.id=='more'){S.n+=150;draw()}
 else if(t.dataset.j)show(+t.dataset.j);
 else if(t.id=='dx')$('#det').classList.remove('on');
 else if(t.id=='dp'&&S.i>0)show(S.i-1);
 else if(t.id=='dn'&&S.i<S.res.length-1)show(S.i+1);
 else if(t.dataset.del){const d=S.docs.find(x=>x.id==t.dataset.del);if(await ask('Supprimer ce document ?',d.name,'Supprimer',1)){t.closest('.doc').classList.add('out');await wait(300);await DB.del(d.id);S.docs=S.docs.filter(x=>x!=d);fill();run()}}
 else if(t.id=='delall'&&await ask('Tout supprimer ?','Les '+S.docs.length+' documents seront retirés de la mémoire de l’application.','Tout supprimer',1)){await DB.clr();S.docs=[];fill();run()}});
$('#fi').onchange=e=>{const f=[...e.target.files];e.target.value='';if(f.length)add(f)};
let tm;$('#q').oninput=()=>{clearTimeout(tm);tm=setTimeout(run,250)};
$$('#oc,#ow,#oa,#oh,#om,#ot').forEach(x=>x.onchange=run);$('#os').onchange=()=>{tabSel();run()};
let sx=0;$('#det').addEventListener('touchstart',e=>sx=e.touches[0].clientX,{passive:true});
$('#det').addEventListener('touchend',e=>{const d=e.changedTouches[0].clientX-sx;if(d<-70&&S.i<S.res.length-1)show(S.i+1);else if(d>70&&S.i>0)show(S.i-1)},{passive:true});
function back(){if(cancel)cancel();else if($('#det').classList.contains('on'))$('#det').classList.remove('on');else if(cur!='home')go('home');else quit()}
document.addEventListener('keydown',e=>{if(e.key=='Escape')back()});
(async()=>{
 try{await DB.init();S.docs=await DB.all();S.docs.sort((a,b)=>a.added-b.added)}catch(e){toast('Stockage indisponible : '+e)}
 if(navigator.storage&&navigator.storage.persist)navigator.storage.persist();
 fill();draw();go('home');
 const A=window.Capacitor&&Capacitor.Plugins&&Capacitor.Plugins.App;if(A)A.addListener('backButton',back)})();
