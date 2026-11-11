/* JHCSC Manual Review - the website logic. Talks to the Apps Script API through js/api.js. */
const BUILD='24';   // must equal <meta name="build"> in index.html and the ?v= on every file; change all of them together on each release
let STALE='';   // set when index.html and app.js are from different releases; the banner then stays on screen
(function(){
  const m=document.querySelector('meta[name=build]');
  if(!m||m.content!==BUILD){STALE='Your browser is showing an old copy of the website files (page '+(m?m.content:'unknown')+', script '+BUILD+'). Press Ctrl+Shift+R to load the new files.';showDiag(STALE)}
})();
/* ---------- visible diagnostics: any load-time problem shows in a red banner instead of failing silently ---------- */
function showDiag(msg){const d=document.getElementById('diag');if(d){d.textContent=msg;d.classList.remove('hidden')}}
function hideDiag(){if(STALE){showDiag(STALE);return}const d=document.getElementById('diag');if(d)d.classList.add('hidden')}
window.addEventListener('error',e=>showDiag('Page error: '+(e.message||'unknown')+(e.lineno?' (line '+e.lineno+')':'')));
window.addEventListener('unhandledrejection',e=>showDiag('Page error: '+((e.reason&&e.reason.message)||e.reason)));
/* ---------- storage that never throws (Safari / blocked third-party storage) ---------- */
const mem={};
const store={get(k){try{return localStorage.getItem(k)||''}catch(e){return mem[k]||''}},set(k,v){try{localStorage.setItem(k,v)}catch(e){mem[k]=v}},del(k){try{localStorage.removeItem(k)}catch(e){delete mem[k]}}};

let state=null,pending=0,refreshing=false,currentView='dashboard',sessionToken=store.get('mrmsSession');
/* ui keeps drafts, open forms and list sizes alive across redraws */
const ui={drafts:{},open:{},limit:{issueList:25,issueList2:25,archiveList:25}};
const $=id=>document.getElementById(id);
const EMAIL_RE=/^[A-Za-z0-9._%+\-]+@[A-Za-z0-9\-]+(\.[A-Za-z0-9\-]+)*\.[A-Za-z]{2,}$/;

function call(name,...args){
  const limit=(name==='mrmsSyncDriveComments'||name==='mrmsGenerateReport')?400000:90000;
  return API.call(name,args,{token:sessionToken,timeout:limit}).catch(e=>{
    const msg=e.message||String(e);
    if(/session expired/i.test(msg)){sessionToken='';store.del('mrmsSession');if(name!=='mrmsGetAppData')setTimeout(()=>refresh(),1500)}
    else if(/access is not active/i.test(msg)){setTimeout(()=>refresh(),1500)}
    throw e;
  });
}
function toast(text,isError){const m=$('message');m.textContent=text;m.classList.toggle('error',!!isError);m.classList.remove('hidden');clearTimeout(toast.timer);toast.timer=setTimeout(()=>m.classList.add('hidden'),isError?9000:4500)}
$('message').addEventListener('click',()=>$('message').classList.add('hidden'));
const STATUS_TEXT={Open:'\u26A0 Needs work',Resolved:'\u23F3 Fix written',Archived:'\u2714 Finished'};
const VOTE_TEXT={Agree:'\uD83D\uDC4D I agree',Disagree:'\uD83D\uDC4E I disagree','Needs discussion':'\uD83D\uDCAC Let\u2019s talk'};
const CONSENSUS_TEXT={'Agreement threshold met':'Most people agree','Disagreement threshold met':'Most people disagree','No majority \u2014 chair review':'No clear answer yet. The chair will decide.','Needs discussion':'People want to talk about it','Awaiting votes':'Waiting for votes','No eligible reviewers':'No reviewers yet'};
const DECISION_TEXT={'Approved':'Approved','Returned for revision':'Sent back'};
const statusText=v=>STATUS_TEXT[v]||STATUS_TEXT.Open, consensusText=v=>CONSENSUS_TEXT[v]||v, decisionText=v=>DECISION_TEXT[v]||v;
function esc(s){return String(s==null?'':s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function safeUrl(u){try{const x=new URL(String(u||''));return x.protocol==='https:'?x.href:''}catch(e){return ''}}
function date(s){return s?new Date(s).toLocaleDateString():'—'}

/* ---------- busy handling: global bar + the clicked button is disabled until the call returns ---------- */
async function run(btn,fn){
  if(btn){if(btn.disabled)return;btn.disabled=true;btn.classList.add('busy')}
  pending++;$('bar').classList.remove('hidden');
  try{return await fn()}catch(e){toast(e.message||String(e),true)}
  finally{pending--;if(!pending)$('bar').classList.add('hidden');if(btn&&btn.isConnected){btn.disabled=false;btn.classList.remove('busy')}}
}

let lastRefresh=0;
async function refresh(){
  if(refreshing)return;
  refreshing=true;pending++;$('bar').classList.remove('hidden');
  const watchdog=setTimeout(()=>showDiag('Still waiting for the server... If this does not finish, check API_URL in js/config.js and that the web app is deployed as a New version.'),15000);
  const load=async()=>{
    const s=await call('mrmsGetAppData');
    if(!s||typeof s!=='object')throw new Error('The server returned no data. Check SPREADSHEET_ID in the Apps Script project properties.');
    state=s;draw();lastRefresh=Date.now();
  };
  try{
    try{await load();if(viewerManualId)loadDriveComments(viewerToken,true);hideDiag()}
    catch(e){
      if(/session expired/i.test(e.message||'')){sessionToken='';store.del('mrmsSession');await load();$('codeBox').classList.add('hidden');$('regCode').value='';hideDiag();toast('Your session expired. Please verify your email again to sign in.',true)}
      else throw e;
    }
  }catch(e){const m=e.message||String(e);showDiag('Could not load the app: '+m);toast(m,true)}
  finally{clearTimeout(watchdog);refreshing=false;pending=Math.max(0,pending-1);if(!pending)$('bar').classList.add('hidden')}
}
/* Automatic refresh: when the page opens, every 60 seconds while the tab is visible, and when you come back to the tab.
   It waits while you are typing or another request is running, so nothing you are writing is interrupted. */
function typing(){const e=document.activeElement;return !!(e&&e.matches&&e.matches('textarea,input:not([type=checkbox]):not([type=file]),select'))}
function autoRefresh(){if(document.hidden||pending||refreshing||typing()||!state||!state.email)return;refresh()}
setInterval(autoRefresh,60000);
document.addEventListener('visibilitychange',()=>{if(!document.hidden&&Date.now()-lastRefresh>20000)autoRefresh()});

/* ---------- drawing ---------- */
function titleOf(id){const m=(state.manuals||[]).find(x=>x['File ID']===id);return m?m['Manual title']:'Unknown manual'}
function assignable(){const s=state.scope||[];return s.length?state.manuals.filter(m=>s.includes(m['File ID'])):state.manuals}
function checkList(manuals,checkedIds){return manuals.map(m=>`<label class="check"><input type="checkbox" data-mid="${esc(m['File ID'])}" ${checkedIds.includes(m['File ID'])?'checked':''}> ${esc(m['Manual title'])}</label>`).join('')}
function checkedIds(root){return Array.from(root.querySelectorAll('input[data-mid]:checked')).map(o=>o.dataset.mid)}

function draw(){
  $('boot').classList.add('hidden');
  $('school').textContent=state.school;$('user').textContent=state.email||'';$('signout').classList.toggle('hidden',!state.email);
  if(!state.chair&&state.accessStatus!=='Approved'){drawGate();return}
  $('accessGate').classList.add('hidden');$('mainNav').classList.remove('hidden');
  $('chairTab').classList.toggle('hidden',!state.chair);$('sync').classList.toggle('hidden',!state.chair);$('report').classList.toggle('hidden',!state.chair);
  if(state.logoUrl&&safeUrl(state.logoUrl)){$('logo').src=safeUrl(state.logoUrl);$('logo').classList.remove('hidden');$('crest').classList.add('hidden')}
  const labels=[['Manuals',state.cards.manuals],['All comments',state.cards.total],['Not finished',state.cards.open],['Fix written',state.cards.resolved],['Finished',state.cards.archived]];
  $('cards').innerHTML=labels.map(x=>`<div class="card"><span>${x[0]}</span><strong>${x[1]}</strong></div>`).join('');
  ['manualFilter','manualFilter2'].forEach(id=>{const s=$(id),v=s.value;s.innerHTML='<option value="">All manuals</option>'+state.manuals.map(m=>`<option value="${esc(m['File ID'])}">${esc(m['Manual title'])}</option>`).join('');s.value=v});
  $('manualList').innerHTML=state.manuals.length?`<table><thead><tr><th>Manual</th><th>What do you want to do?</th><th>Updated</th><th>Comments</th></tr></thead><tbody>${state.manuals.map(m=>{const count=state.issues.filter(i=>i.manualIds.includes(m['File ID'])).length;const download='https://drive.google.com/uc?export=download&id='+encodeURIComponent(m['File ID']);return `<tr><td>${esc(m['Manual title'])}</td><td><button class="btn small primary" data-act="view-manual" data-id="${esc(m['File ID'])}">\uD83D\uDCD6 Read it here</button> &nbsp; <a href="${esc(download)}" target="_blank" rel="noopener">Download</a></td><td>${date(m['Last synced'])}</td><td>${count}</td></tr>`}).join('')}</tbody></table>`:'<div class="empty">No manuals yet. The chair can get them from Google Drive.</div>';
  const mine=assignable(),primary=$('entryManual'),related=$('relatedOptions'),annotated=$('annotatedManual'),priorPrimary=primary.value,priorPdf=annotated.value,priorRelated=checkedIds(related);
  const options=mine.map(m=>`<option value="${esc(m['File ID'])}">${esc(m['Manual title'])}</option>`).join('');
  primary.innerHTML='<option value="">Choose a manual</option>'+options;annotated.innerHTML='<option value="">Choose a manual</option>'+options;
  related.innerHTML=checkList(state.manuals,priorRelated);
  if(mine.some(m=>m['File ID']===priorPrimary))primary.value=priorPrimary;if(mine.some(m=>m['File ID']===priorPdf))annotated.value=priorPdf;
  const wordSel=$('wordManual'),priorWord=wordSel.value;wordSel.innerHTML='<option value="">Choose a manual</option>'+options;if(mine.some(m=>m['File ID']===priorWord))wordSel.value=priorWord;
  renderIssues('issueList');renderIssues('issueList2');renderArchive();renderChair();
  showView(currentView);
}
function drawGate(){
  $('boot').classList.add('hidden');
  $('mainNav').classList.add('hidden');document.querySelectorAll('main>section').forEach(s=>s.classList.add('hidden'));$('accessGate').classList.remove('hidden');
  const st=state.accessStatus;
  $('accessMessage').textContent=st==='Pending'?'Thank you! Your email is checked. Now we wait for the chair to say yes. We will send you an email. You can also come back to this page to check.'
    :st==='Rejected'?'Sorry, the chair did not say yes this time. You can fix your details and try again.'
    :st==='Deactivated'?'Your access is stopped. Please ask the chair to turn it on again.'
    :'Anyone can ask for an account with any email. We send a code to prove the email is yours. Then the chair says yes before you can come in.';
  $('registrationForm').classList.toggle('hidden',st==='Pending'||st==='Deactivated');
  $('gateHelp').classList.toggle('hidden',st==='Pending'||st==='Deactivated');
  $('switchEmail').classList.toggle('hidden',!state.email);
  if(state.email)$('regEmail').value=state.email;
  if(state.displayName&&!$('regName').value)$('regName').value=state.displayName;
  const keep=checkedIds($('regManuals'));$('regManuals').innerHTML=checkList(state.manuals,keep);
}
function showView(view){
  currentView=view;
  document.querySelectorAll('main>section').forEach(s=>s.classList.add('hidden'));
  $('accessGate').classList.add('hidden');
  $(view).classList.remove('hidden');
  document.querySelectorAll('.nav button').forEach(b=>{const on=b.dataset.view===view;b.classList.toggle('active',on);if(on)b.setAttribute('aria-current','page');else b.removeAttribute('aria-current')});
}
function goView(v){showView(v);window.scrollTo({top:0,behavior:'smooth'})}
function filterChanged(target){ui.limit[target]=25;renderIssues(target);if(target==='archiveList')renderArchive()}

/* ---------- issue cards ---------- */
function filterRows(rows,target){
  const q=({issueList:$('search1'),issueList2:$('search2'),archiveList:$('search3')}[target].value||'').trim().toLowerCase();
  const manual=target==='issueList'?$('manualFilter').value:target==='issueList2'?$('manualFilter2').value:'';
  const status=target==='issueList'?$('statusFilter').value:'';
  return rows.filter(i=>(!manual||i.manualIds.includes(manual))&&(!status||i.status===status)&&(!q||(i.comment+' '+i.manualTitles.join(' ')+' '+i.author+' '+i.resolution).toLowerCase().includes(q)))
    .sort((a,b)=>String(b.created).localeCompare(String(a.created)));
}
function listHtml(rows,target,cardFn,emptyText){
  if(!rows.length)return `<div class="empty">${emptyText}</div>`;
  const lim=ui.limit[target];
  return rows.slice(0,lim).map(cardFn).join('')+(rows.length>lim?`<button class="btn" data-act="more" data-target="${target}">Show more (${rows.length-lim} more)</button>`:'');
}
function renderIssues(target='issueList'){
  if(!state||!state.issues)return;
  const rows=filterRows(state.issues.filter(i=>i.status!=='Archived'),target);
  $(target).innerHTML=listHtml(rows,target,issueCard,'No comments to show here.');
}
function renderArchive(){
  if(!state||!state.issues)return;
  $('archiveList').innerHTML=listHtml(filterRows(state.issues.filter(i=>i.status==='Archived'),'archiveList'),'archiveList',archiveCard,'Nothing is finished yet. Approved comments show up here.');
}
function resolutionBox(i){
  if(!i.resolution&&!i.decision)return '';
  const returned=i.decision==='Returned for revision';
  return `<div class="box${returned?' returned':''}">${i.resolution?'<b>The fix that was written</b><br>'+esc(i.resolution):''}${i.decision?`${i.resolution?'<br><br>':''}<b>The chair says:</b> ${esc(decisionText(i.decision))}${i.chairDate?' · '+date(i.chairDate):''}${returned?'<br>Read the messages below to see what to change.':''}`:''}</div>`;
}
function discussionHtml(i,canPost){
  const k=i.id+'|';
  const comments=i.comments.map(c=>`<p><b>${esc(c.email)}</b> · ${date(c.time)}<br>${esc(c.message)}</p>`).join('');
  const composer=canPost?`<div class="row" style="margin-top:10px"><textarea class="wide" data-field="message" maxlength="2000" placeholder="Write a message for the team" aria-label="Message for the team">${esc(ui.drafts[k+'message']||'')}</textarea><button class="btn small primary" data-act="post-comment" data-id="${esc(i.id)}">Send</button></div>`:'';
  return `<details class="discussion" data-disc ${ui.open[k+'disc']?'open':''}><summary>Talk about this (${i.comments.length})</summary>${comments}${composer}</details>`;
}
function issueCard(i){
  const id=esc(i.id),k=i.id+'|',vis=f=>ui.open[k+f]?'':' hidden';
  const badge=i.status==='Resolved'?'warn':'open';
  const manualText=i.manualTitles.filter(Boolean).join(' · ')||'Manual title pending';
  const why=state.chair&&!i.canVote?'The chair does not vote unless listed as a reviewer':'You are not assigned to review this manual';
  const votes=['Agree','Disagree','Needs discussion'].map(v=>`<button class="vote${i.myVote===v?' selected':''}" aria-pressed="${i.myVote===v}" ${i.canVote?'':`disabled title="${esc(why)}"`} data-act="vote" data-id="${id}" data-value="${v}">${VOTE_TEXT[v]} <span>(${i.tally[v]||0})</span></button>`).join('');
  const tools=[];
  if(i.canVote&&i.assignee!==state.email)tools.push(`<button class="btn small" data-act="assign-me" data-id="${id}">I will work on this</button>`);
  if(i.canAct)tools.push(`<button class="btn small" data-act="toggle-form" data-target-form="resolve">${i.status==='Resolved'?'Change the fix':'Write the fix'}</button>`);
  if(state.chair)tools.push(`<button class="btn small" data-act="toggle-form" data-target-form="link">Other manuals it affects</button>`);
  const src=safeUrl(i.url);
  return `<article class="issue" data-card="${id}"><div class="issue-head"><div><h3>${esc(manualText)}</h3><p class="meta">${esc(i.author||'Drive reviewer')} · ${date(i.created)} · ${i.eligible} reviewers${i.assignee?' · Working on it: '+esc(i.assignee):''}${i.due?' · Due '+date(i.due):''}</p></div><span class="pill ${badge}">${esc(statusText(i.status))}</span></div><p class="comment">${esc(i.comment)}</p>${resolutionBox(i)}<div class="row"><span class="pill ${/chair review|Tie/.test(i.consensus)?'warn':''}">${esc(consensusText(i.consensus))} · \uD83D\uDC4D ${i.tally.Agree} · \uD83D\uDC4E ${i.tally.Disagree}</span>${src?`<a href="${esc(src)}" target="_blank" rel="noopener">Open the PDF</a>`:''}</div><div class="votes">${votes}</div>${i.canVote?'':'<p class="note" style="margin:0 0 .6rem">Only the people assigned to this manual can vote on this comment.</p>'}<div class="toolbar">${tools.join('')}</div>
  <div class="inline-form${vis('resolve')}" data-form="resolve"><label>What should change?</label><textarea data-field="resolution" maxlength="3000" placeholder="Write what should change. The chair will check it.">${esc(ui.drafts[k+'resolution']||'')}</textarea><div class="toolbar"><button class="btn small primary" data-act="save-resolution" data-id="${id}">Save my fix</button><button class="btn small" data-act="toggle-form" data-target-form="resolve">Cancel</button></div></div>
  ${state.chair?`<div class="inline-form${vis('link')}" data-form="link"><label>Which manuals does this affect?</label><div class="checks">${checkList(state.manuals,i.manualIds)}</div><div class="toolbar"><button class="btn small primary" data-act="save-link" data-id="${id}">Save</button><button class="btn small" data-act="toggle-form" data-target-form="link">Cancel</button></div></div>`:''}
  ${discussionHtml(i,i.canAct)}</article>`;
}
function archiveCard(i){
  const manualText=i.manualTitles.filter(Boolean).join(' · ')||'Manual issue',src=safeUrl(i.url);
  return `<article class="issue" data-card="${esc(i.id)}"><div class="issue-head"><div><h3>${esc(manualText)}</h3><p class="meta">${esc(i.author||'Drive reviewer')} · ${date(i.created)} · approved on ${date(i.chairDate)} by ${esc(i.chairEmail||'chair')}</p></div><span class="pill">${esc(statusText('Archived'))}</span></div><p class="comment">${esc(i.comment)}</p>${resolutionBox(i)}<div class="row"><span class="pill">${esc(consensusText(i.consensus))} · \uD83D\uDC4D ${i.tally.Agree} · \uD83D\uDC4E ${i.tally.Disagree} · \uD83D\uDCAC ${i.tally['Needs discussion']} (out of ${i.eligible} reviewers${i.frozen?' when it was approved':''})</span>${src?`<a href="${esc(src)}" target="_blank" rel="noopener">Open the PDF</a>`:''}</div>${discussionHtml(i,false)}</article>`;
}
function renderChair(){
  const el=$('chairList');if(!state.chair){el.innerHTML='';return}
  el.innerHTML=state.chairQueue.length?state.chairQueue.map(i=>`<article class="issue" data-card="${esc(i.id)}"><h3>${esc(i.manualTitles.filter(Boolean).join(' · ')||'Manual issue')}</h3><p class="comment">${esc(i.comment)}</p><div class="box"><b>The fix</b><br>${esc(i.resolution||'—')}<br><span class="note">${esc(consensusText(i.consensus))} · \uD83D\uDC4D ${i.tally.Agree} · \uD83D\uDC4E ${i.tally.Disagree}</span></div><label>Why are you sending it back?</label><textarea class="field wide" data-field="returnNote" rows="2" maxlength="1000" placeholder="Tell reviewers what needs to change">${esc(ui.drafts[i.id+'|returnNote']||'')}</textarea><div class="toolbar"><button class="btn primary" data-act="chair-approve" data-id="${esc(i.id)}">Approve (finish this comment)</button><button class="btn" data-act="chair-return" data-id="${esc(i.id)}">Send back</button></div></article>`).join(''):'<div class="empty">Nothing is waiting for you.</div>';
  $('registrationList').innerHTML=state.registrations.length?state.registrations.map(r=>`<article class="issue" data-card="${esc(r.email)}"><h3>${esc(r.name)}</h3><p class="meta">${esc(r.email)} · Requested ${date(r.submitted)} · \u2714 Email checked with a code</p><label>Which manuals can they check?</label><div class="checks">${checkList(state.manuals,r.requested)}</div><p class="note">${r.requested.length?'They asked for: '+esc(r.requested.map(titleOf).join(', ')):'They did not choose any manuals.'} If you tick none, they can see all manuals.</p><div class="toolbar"><button class="btn primary" data-act="reg-approve">Let them in</button><button class="btn" data-act="reg-reject">Say no</button></div></article>`).join(''):'<div class="empty">No one is waiting to join.</div>';
  $('memberList').innerHTML=state.members.length?state.members.map(m=>`<article class="issue" data-card="${esc(m.email)}"><div class="issue-head"><div><h3>${esc(m.name)}</h3><p class="meta">${esc(m.email)} · ${m.manualIds.length?esc(m.manualIds.map(titleOf).join(', ')):'All manuals'}</p></div><span class="pill ${m.active?'':'open'}">${m.active?'\u2714 Can sign in':'\u26D4 Stopped'}</span></div><details><summary>Change their manuals</summary><div class="checks" style="margin-top:8px">${checkList(state.manuals,m.manualIds)}</div><div class="toolbar"><button class="btn small primary" data-act="member-save" data-active="${m.active}">Save</button></div></details><div class="toolbar"><button class="btn small ${m.active?'danger':''}" data-act="member-toggle" data-active="${m.active}">${m.active?'Stop access':'Allow access'}</button></div></article>`).join(''):'<div class="empty">No members yet.</div>';
}

/* ---------- one delegated handler for every dynamic button (no inline JS built from data) ---------- */
document.addEventListener('input',e=>{const f=e.target.closest&&e.target.closest('[data-field]'),c=f&&f.closest('[data-card]');if(f&&c)ui.drafts[c.dataset.card+'|'+f.dataset.field]=f.value});
document.addEventListener('toggle',e=>{const d=e.target;if(d&&d.matches&&d.matches('[data-disc]')){const c=d.closest('[data-card]');if(c)ui.open[c.dataset.card+'|disc']=d.open}},true);
document.addEventListener('click',e=>{
  const b=e.target.closest('[data-act]');if(!b||b.disabled)return;
  const act=b.dataset.act,card=b.closest('[data-card]'),id=b.dataset.id||(card&&card.dataset.card)||'';
  const field=n=>card.querySelector(`[data-field="${n}"]`);
  const ok=(msg,after)=>({ok:msg,after:after});
  switch(act){
    case 'vote':return doAct(b,id,'vote',b.dataset.value,ok('Your vote was saved: '+VOTE_TEXT[b.dataset.value]));
    case 'assign-me':return doAct(b,id,'assign',state.email,ok('Done. You are working on this.'));
    case 'toggle-form':{const f=card.querySelector(`[data-form="${b.dataset.targetForm}"]`);f.classList.toggle('hidden');ui.open[id+'|'+b.dataset.targetForm]=!f.classList.contains('hidden');return}
    case 'save-resolution':{const v=field('resolution').value.trim();if(!v){toast('Please write what should change.',true);return}
      return doAct(b,id,'resolve',v,{ok:'Saved. The chair will check your fix.',after:()=>{delete ui.drafts[id+'|resolution'];ui.open[id+'|resolve']=false}})}
    case 'save-link':{const ids=checkedIds(card);if(!ids.length){toast('Please tick at least one manual.',true);return}
      return doAct(b,id,'link',ids.join(','),{ok:'Saved.',after:()=>{ui.open[id+'|link']=false}})}
    case 'post-comment':{const v=field('message').value.trim();if(!v){toast('Please write a message first.',true);return}
      return doAct(b,id,'comment',v,{ok:'Message sent.',after:()=>{delete ui.drafts[id+'|message']}})}
    case 'chair-approve':return doAct(b,id,'chair','Approved',ok('Approved. This comment is finished.'));
    case 'chair-return':{const note=field('returnNote').value.trim();if(!note){toast('Please say what needs to change.',true);return}
      return doAct(b,id,'chair','Returned for revision',{note:note,ok:'Sent back. People can read your reason.',after:()=>{delete ui.drafts[id+'|returnNote']}})}
    case 'reg-approve':return run(b,async()=>{state=await call('mrmsReviewRegistration',id,'Approved',checkedIds(card));draw();toast('They are in. We sent them an email.')});
    case 'reg-reject':return run(b,async()=>{state=await call('mrmsReviewRegistration',id,'Rejected',[]);draw();toast('You said no.')});
    case 'member-save':return run(b,async()=>{state=await call('mrmsSaveMember',id,checkedIds(card),b.dataset.active==='true');draw();toast('Saved.')});
    case 'member-toggle':{const nowActive=b.dataset.active==='true';
      if(nowActive&&!confirm('Stop this person\u2019s access? They cannot sign in until you allow it again.'))return;
      return run(b,async()=>{state=await call('mrmsSaveMember',id,checkedIds(card),!nowActive);draw();toast(nowActive?'Their access is stopped.':'They can sign in again.')})}
    case 'more':ui.limit[b.dataset.target]+=25;if(b.dataset.target==='archiveList')renderArchive();else renderIssues(b.dataset.target);return;
    case 'view-manual':return showManual(id);
    case 'close-manual':return closeManualViewer();
  }
});
function doAct(btn,id,action,value,opt){
  opt=opt||{};
  return run(btn,async()=>{state=await call('mrmsAction',{id:id,action:action,value:value,note:opt.note||''});if(opt.after)opt.after();draw();toast(opt.ok||'Saved.')});
}

/* ---------- manual viewer: built-in PDF reader (highlight text -> page + text go to the form) + that PDF's Drive comments ---------- */
let viewerManualId='',viewerToken=0,viewerDoc=null,viewerScale=1,viewerPage=1,viewerObserver=null,viewerRaf=0;

/* The PDF engine (PDF.js) ships inside this project (js/vendor/pdfjs/), so no CDN is needed. The two CDN copies are only a backup. */
const PDFJS_SOURCES=[
  {lib:'js/vendor/pdfjs/pdf.min.js',worker:'js/vendor/pdfjs/pdf.worker.min.js',local:true},
  {lib:'https://cdn.jsdelivr.net/npm/pdfjs-dist@6.4.299/build/pdf.min.mjs',worker:'https://cdn.jsdelivr.net/npm/pdfjs-dist@6.4.299/build/pdf.worker.min.mjs'},
  {lib:'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/6.4.299/pdf.min.mjs',worker:'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/6.4.299/pdf.worker.min.mjs'}
];
async function loadPdfJs(){
  if(window.pdfjsLib&&window.pdfjsLib.GlobalWorkerOptions&&window.pdfjsLib.GlobalWorkerOptions.workerSrc)return window.pdfjsLib;
  const tried=[];
  for(const s of PDFJS_SOURCES){
    try{
      const abs=u=>s.local?new URL(u,location.href).href:u;
      const lib=await import(abs(s.lib));
      lib.GlobalWorkerOptions.workerSrc=abs(s.worker);   // the reader cannot start without its worker
      window.pdfjsLib=lib;return lib;
    }catch(e){tried.push((s.local?'local copy':new URL(s.lib).host)+': '+(e&&e.message||e))}
  }
  throw new Error('The PDF engine could not be loaded ('+tried.join('; ')+'). Check that the folder frontend/js/vendor/pdfjs/ was uploaded to GitHub.');
}

/* Downloads the PDF in 3 MB pieces (works for large manuals) and shows the progress. */
async function fetchPdf(id,tok){
  const first=await call('mrmsGetManualPdf',id,0);
  if(tok!==viewerToken)throw new Error('cancelled');
  const total=first.total,piece=first.piece||3145728,n=Math.max(1,Math.ceil(total/piece)),parts=new Array(n);
  parts[0]=b64ToBytes(first.data);let done=parts[0].length,next=1;
  const show=()=>{$('pdfScroll').innerHTML='<div class="empty">Getting the manual ready… '+Math.min(100,Math.round(done*100/Math.max(1,total)))+'%</div>'};show();
  async function worker(){                                   // three pieces are fetched at the same time
    while(next<n){
      const i=next++,r=await call('mrmsGetManualPdf',id,i*piece);
      if(tok!==viewerToken)throw new Error('cancelled');
      parts[i]=b64ToBytes(r.data);done+=parts[i].length;show();
    }
  }
  await Promise.all([worker(),worker(),worker()]);
  const all=new Uint8Array(parts.reduce((s,x)=>s+x.length,0));let p=0;parts.forEach(x=>{all.set(x,p);p+=x.length});return all;
}
function b64ToBytes(s){const bin=atob(s),out=new Uint8Array(bin.length);for(let i=0;i<bin.length;i++)out[i]=bin.charCodeAt(i);return out}

async function showManual(id){
  const manual=state.manuals.find(m=>m['File ID']===id);if(!manual)return;
  if(!assignable().some(m=>m['File ID']===id))toast('You can read this manual, but it is not in your assigned list for commenting.',true);
  closeManualViewer(true);
  $('readerMsg').classList.add('hidden');
  viewerManualId=id;const tok=++viewerToken;
  $('manualViewerTitle').textContent=manual['Manual title'];
  $('manualFrame').classList.add('hidden');$('manualFrame').src='about:blank';
  $('pdfScroll').classList.remove('hidden');$('pdfScroll').innerHTML='<div class="empty">Opening the manual…</div>';$('pdfPageInfo').textContent='Loading…';
  $('manualViewerPanel').classList.remove('hidden');
  $('manualViewerPanel').scrollIntoView({behavior:'smooth',block:'start'});
  loadDriveComments(tok);
  try{
    const [lib,bytes]=await Promise.all([loadPdfJs(),fetchPdf(id,tok)]);
    if(tok!==viewerToken)return;
    $('pdfScroll').innerHTML='<div class="empty">Almost ready…</div>';
    const doc=await lib.getDocument({data:bytes}).promise;
    if(tok!==viewerToken){doc.destroy();return}
    viewerDoc=doc;await buildPages(doc,tok);
  }catch(e){
    if(tok!==viewerToken||/cancelled/.test(String(e&&e.message)))return;
    console.error('Reader failed',e);showDrivePreview(id);
    const m=String(e&&e.message||e);
    $('readerMsg').textContent='The built-in reader could not open this PDF, so the plain Drive preview is shown below (highlighting will not fill the form; type the page and text yourself).\n\nReason: '+m
      +(/Unknown action/i.test(m)?'\nWhat to do: the server code is older than the website. In Apps Script open Deploy > Manage deployments > pencil > Version: New version > Deploy.'
        :/TOO_LARGE/.test(m)?'\nWhat to do: this PDF is over 80 MB, which is too large for the built-in reader.'
        :/session expired|access is not active/i.test(m)?'\nWhat to do: sign in again.'
        :/Cannot reach|Unexpected reply|taking too long/i.test(m)?'\nWhat to do: check API_URL in js/config.js and that the web app is deployed with access "Anyone".':'')
      +'\nIf the preview below is blank, sign in to Google in this browser (with an account that can open the PDF) or use Download PDF in the Manuals list.';
    $('readerMsg').classList.remove('hidden');
  }
}
async function buildPages(doc,tok){
  const box=$('pdfScroll');box.innerHTML='';
  const first=await doc.getPage(1);if(tok!==viewerToken)return;
  viewerScale=Math.min(2,Math.max(280,box.clientWidth-26)/first.getViewport({scale:1}).width);
  const vp=first.getViewport({scale:viewerScale});
  viewerObserver=new IntersectionObserver(es=>es.forEach(en=>{if(en.isIntersecting)renderPage(en.target)}),{root:box,rootMargin:'700px 0px'});
  for(let n=1;n<=doc.numPages;n++){
    const d=document.createElement('div');d.className='pdfpage';d.dataset.page=n;d.style.width=vp.width+'px';d.style.height=vp.height+'px';
    box.appendChild(d);viewerObserver.observe(d);
  }
  $('pdfGo').max=doc.numPages;updateCurrentPage();
}
async function renderPage(div){
  if(div.dataset.done)return;div.dataset.done='1';
  const tok=viewerToken,doc=viewerDoc,n=Number(div.dataset.page);
  try{
    const page=await doc.getPage(n);if(tok!==viewerToken)return;
    const vp=page.getViewport({scale:viewerScale}),ratio=Math.min(2,window.devicePixelRatio||1);
    div.style.width=vp.width+'px';div.style.height=vp.height+'px';div.style.setProperty('--total-scale-factor',vp.scale);div.style.setProperty('--scale-factor',vp.scale);
    const canvas=document.createElement('canvas');
    canvas.width=Math.floor(vp.width*ratio);canvas.height=Math.floor(vp.height*ratio);canvas.style.width=vp.width+'px';canvas.style.height=vp.height+'px';
    div.appendChild(canvas);
    await page.render({canvasContext:canvas.getContext('2d'),viewport:vp,transform:ratio!==1?[ratio,0,0,ratio,0,0]:null}).promise;
    if(tok!==viewerToken)return;
    try{
      const tl=document.createElement('div');tl.className='textLayer';div.appendChild(tl);
      await new window.pdfjsLib.TextLayer({textContentSource:page.streamTextContent(),container:tl,viewport:vp}).render();
    }catch(e){/* the page stays readable; only text selection is unavailable */}
  }catch(e){console.error('Page '+n+' could not be drawn',e);div.dataset.done='';if(!div.querySelector('canvas'))div.innerHTML='<div class="empty" style="margin:20px">Page '+n+' could not be drawn. Scroll away and back to try again.</div>'}
}

/* Highlight text in the PDF -> the page number and the text are put in the comment form. */
let selTimer=0;
function pageOfNode(node){const el=node&&(node.nodeType===1?node:node.parentElement),pg=el&&el.closest&&el.closest('.pdfpage');return pg&&$('pdfScroll').contains(pg)?Number(pg.dataset.page):0}
function readSelection(){
  if(!viewerManualId)return;
  const sel=window.getSelection();if(!sel||!sel.rangeCount)return;
  const text=sel.toString().replace(/\s+/g,' ').trim();if(!text)return;
  const p1=pageOfNode(sel.anchorNode),p2=pageOfNode(sel.focusNode);if(!p1&&!p2)return;
  const a=Math.min(p1||p2,p2||p1),b=Math.max(p1||p2,p2||p1);
  $('vQuote').value=text.slice(0,2000);
  $('vLocation').value=a===b?'Page '+a:'Pages '+a+'-'+b;
  $('selNote').textContent='Got it: '+$('vLocation').value+'. Now write what you think.';
}
document.addEventListener('selectionchange',()=>{if(!viewerManualId)return;clearTimeout(selTimer);selTimer=setTimeout(readSelection,300)});

$('pdfScroll').addEventListener('scroll',()=>{if(viewerRaf)return;viewerRaf=requestAnimationFrame(()=>{viewerRaf=0;updateCurrentPage()})});
$('pdfGo').addEventListener('change',()=>{const d=$('pdfScroll').querySelector('[data-page="'+Number($('pdfGo').value)+'"]');if(d)$('pdfScroll').scrollTo({top:d.offsetTop-8,behavior:'smooth'})});
function updateCurrentPage(){
  if(!viewerDoc)return;
  const box=$('pdfScroll'),line=box.getBoundingClientRect().top+box.clientHeight/3;let cur=1;
  for(const d of box.children){if(!d.dataset.page)continue;if(d.getBoundingClientRect().top<=line)cur=Number(d.dataset.page);else break}
  viewerPage=cur;$('pdfPageInfo').textContent='Page '+cur+' of '+viewerDoc.numPages;
}
function useCurrentPage(){$('vLocation').value='Page '+viewerPage;$('vComment').focus()}
function showDrivePreview(id){
  $('pdfScroll').classList.add('hidden');$('manualFrame').classList.remove('hidden');
  $('manualFrame').src='https://drive.google.com/file/d/'+encodeURIComponent(id)+'/preview';$('pdfPageInfo').textContent='Drive preview';
}

async function loadDriveComments(tok,quiet){
  if(!viewerManualId)return;
  if(!quiet)$('driveComments').innerHTML='<div class="note">Loading comments…</div>';
  try{const list=await call('mrmsListDriveComments',viewerManualId);if(tok!==viewerToken)return;renderDriveComments(list)}
  catch(e){if(tok!==viewerToken||quiet)return;$('driveComments').innerHTML='<div class="note">Could not load the Drive comments: '+esc(e.message||e)+'</div>'}
}
function renderDriveComments(list){
  const open=list.filter(c=>!c.resolved).length;
  $('driveCommentCount').textContent=list.length?'('+open+' open, '+(list.length-open)+' resolved)':'';
  $('driveComments').innerHTML=list.length?list.map(c=>`<div class="dc${c.resolved?' done':''}"><b>${esc(c.author)}</b> · ${date(c.time)}${c.resolved?' · resolved':''}<br>${esc(c.text)}${c.replies.map(r=>`<div class="rep"><b>${esc(r.author)}</b>: ${esc(r.text)}</div>`).join('')}</div>`).join(''):'<div class="note">No comments on this PDF yet.</div>';
}
/* While a PDF is open, its Drive comments reload every 20 seconds, so comments made in Drive appear here too. */
setInterval(()=>{if(viewerManualId&&!document.hidden&&!pending)loadDriveComments(viewerToken,true)},20000);
function closeManualViewer(keepPanel){
  viewerToken++;viewerManualId='';
  if(viewerObserver){viewerObserver.disconnect();viewerObserver=null}
  if(viewerDoc){try{viewerDoc.destroy()}catch(e){}viewerDoc=null}
  $('pdfScroll').innerHTML='';$('manualFrame').src='about:blank';
  if(!keepPanel)$('manualViewerPanel').classList.add('hidden');
}
function submitViewerComment(btn){
  const payload={manualId:viewerManualId,location:$('vLocation').value,quote:$('vQuote').value,comment:$('vComment').value,suggestion:$('vSuggestion').value};
  if(!payload.manualId){toast('Open a manual first.',true);return}
  if(!payload.location.trim()){toast('Highlight words in the PDF, or type the page.',true);return}
  if(payload.comment.trim().length<5){toast('Please write a little more.',true);return}
  return run(btn,async()=>{
    const list=await call('mrmsPostDriveComment',payload);
    if(payload.manualId!==viewerManualId)return;
    renderDriveComments(list);
    ['vLocation','vQuote','vComment','vSuggestion'].forEach(x=>$(x).value='');
    $('selNote').textContent='Highlight words in the PDF. The page and the words appear here by themselves.';
    toast('Thank you! Your comment was posted.');
  });
}

/* ---------- registration / session ---------- */
function sendCode(btn){
  const email=$('regEmail').value.trim(),name=$('regName').value.trim(),ids=checkedIds($('regManuals'));
  if(!EMAIL_RE.test(email)){toast('Please type your full email address.',true);return}
  if(name.length<2){toast('Please type your name.',true);return}
  return run(btn,async()=>{await call('mrmsSendRegistrationCode',email,name,ids);$('codeBox').classList.remove('hidden');$('regCode').focus();toast('We sent a code to your email. Open your email and type the 6 numbers here.')});
}
function verifyCode(btn){
  const email=$('regEmail').value.trim(),code=$('regCode').value.trim();
  if(!/^\d{6}$/.test(code)){toast('Please type the 6 numbers from your email.',true);return}
  return run(btn,async()=>{const r=await call('mrmsVerifyRegistration',email,code);sessionToken=r.token;store.set('mrmsSession',sessionToken);await refresh();toast(r.status==='Approved'?'You are in. Welcome!':r.status==='Deactivated'?'Your email is checked, but your access is stopped. Please ask the chair.':'Your email is checked. Now wait for the chair to say yes.')});
}
function logout(btn){
  return run(btn,async()=>{try{await call('mrmsLogout')}catch(e){}
    sessionToken='';store.del('mrmsSession');$('codeBox').classList.add('hidden');$('regCode').value='';$('regEmail').value='';$('regName').value='';currentView='dashboard';ui.drafts={};ui.open={};
    state=await call('mrmsGetAppData');draw();toast('You are signed out. See you again!')});
}

/* ---------- chair tools ---------- */
function syncDrive(btn){
  if(!confirm('Scan the configured Drive folder (and its subfolders) for PDF comments and add new comments to the issue register?'))return;
  return run(btn,async()=>{const r=await call('mrmsSyncDriveComments');await refresh();
    toast(`${r.count} new comment(s) imported`+(r.updated?`, ${r.updated} updated`:'')+(r.manualsAdded?`, ${r.manualsAdded} new manual(s)`:'')+(r.deactivated?`, ${r.deactivated} manual(s) no longer in the folder were hidden`:'')+'.'+(r.partial?' Time limit reached; run the sync again to continue.':''),!!r.partial)});
}
function report(btn){
  if(!state.chair)return;const sel=$('manualFilter').value;
  return run(btn,async()=>{const r=await call('mrmsGenerateReport',sel),box=$('reportBox'),pdf=safeUrl(r.url),doc=safeUrl(r.docUrl);
    box.innerHTML='<b>Report ready: '+esc(r.title)+'</b><br>'+(pdf?`<a href="${esc(pdf)}" target="_blank" rel="noopener">Open PDF</a>`:'')+(doc?` &nbsp;·&nbsp; <a href="${esc(doc)}" target="_blank" rel="noopener">Open Google Doc</a>`:'')+(r.shared?'':'<br><span class="note">Could not share automatically. Ask the coordinator to share the files with you.</span>');
    box.classList.remove('hidden');toast('Your report is ready.')});
}

/* ---------- add / import comments ---------- */
function submitPortalComment(btn){
  const payload={manualId:$('entryManual').value,relatedManualIds:checkedIds($('relatedOptions')),location:$('entryLocation').value,quote:$('entryQuote').value,comment:$('entryComment').value,suggestion:$('entrySuggestion').value};
  if(!payload.manualId){toast('Please choose a manual first.',true);return}
  if(payload.comment.trim().length<5){toast('Please write a little more.',true);return}
  return run(btn,async()=>{state=await call('mrmsSubmitPortalComment',payload);['entryLocation','entryQuote','entryComment','entrySuggestion'].forEach(x=>$(x).value='');$('relatedOptions').querySelectorAll('input').forEach(o=>o.checked=false);currentView='issues';draw();toast('Thank you! Your comment was sent.')});
}
function capturePdfNotes(btn){
  const file=$('annotatedPdf').files[0],manualId=$('annotatedManual').value;
  if(!file||!manualId){toast('Choose a manual and the annotated PDF first.',true);return}
    return run(btn,async()=>{
    await loadPdfJs();
    const pdf=await window.pdfjsLib.getDocument({data:new Uint8Array(await file.arrayBuffer())}).promise,annotations=[];
    for(let pageNo=1;pageNo<=pdf.numPages;pageNo++){const page=await pdf.getPage(pageNo),items=await page.getAnnotations({intent:'display'});items.forEach(a=>{const text=typeof a.contents==='string'?a.contents:(a.contentsObj&&a.contentsObj.str)||'';if(text.trim())annotations.push({page:pageNo,type:a.subtype||'',text:text.trim()})})}
    if(!annotations.length){toast('No text comments were found. Add a sticky note or comment in the PDF, or enter it with Add a comment.',true);return}
    const r=await call('mrmsImportPdfAnnotations',manualId,file.name,annotations);state=r.data;$('annotatedPdf').value='';currentView='issues';draw();
    toast(`${r.count} PDF comment(s) added to Issues.`+(r.skipped?` ${r.skipped} already in the list were skipped.`:''));
  });
}
/* ---------- Word review import: read in the browser, preview, then import the ticked comments ---------- */
let wordItems=[];
function readWordReview(btn){
  const file=$('wordFile').files[0],manualId=$('wordManual').value;
  if(!manualId||!file){toast('Choose the manual and the Word file first.',true);return}
  return run(btn,async()=>{
    const res=await WordReview.parse(await file.arrayBuffer());
    wordItems=res.items;
    if(!wordItems.length){$('wordPreview').classList.add('hidden');toast('No comments were found in that Word file. It needs headings with numbered or bulleted points, real Word comments, or a table with a Comment column.',true);return}
    $('wordSummary').textContent='Found '+wordItems.length+' comment(s)'+(res.source?' in "'+res.source+'"':'')+'. Untick any you do not want, then import.';
    $('wordList').innerHTML=wordItems.map((it,i)=>`<label class="check" style="align-items:flex-start"><input type="checkbox" data-wi="${i}" checked><span><b>${esc(it.location)}</b>${it.group?' <span class="note">('+esc(it.group)+')</span>':''}<br>${esc(it.comment.slice(0,260))}${it.comment.length>260?'…':''}${it.suggestion?'<br><span class="note"><b>Suggested:</b> '+esc(it.suggestion.slice(0,200))+(it.suggestion.length>200?'…':'')+'</span>':''}</span></label>`).join('');
    $('wordPreview').classList.remove('hidden');
  });
}
function toggleWordAll(on){$('wordList').querySelectorAll('input[data-wi]').forEach(c=>{c.checked=on})}
function importWordReview(btn){
  const manualId=$('wordManual').value,file=$('wordFile').files[0];
  const chosen=Array.from($('wordList').querySelectorAll('input[data-wi]:checked')).map(c=>wordItems[Number(c.dataset.wi)]);
  if(!manualId){toast('Choose the manual first.',true);return}
  if(!chosen.length){toast('Tick at least one comment.',true);return}
  return run(btn,async()=>{
    const r=await call('mrmsImportReviewItems',manualId,file?file.name:'review.docx',chosen.map(x=>({location:x.location,quote:x.quote,comment:x.comment,suggestion:x.suggestion})));
    state=r.data;$('wordFile').value='';$('wordPreview').classList.add('hidden');wordItems=[];currentView='issues';draw();
    toast(r.count+' comment(s) from the Word file added to Issues.'+(r.skipped?' '+r.skipped+' already in the list were skipped.':''));
  });
}
function csvCell(value){return '"'+String(value||'').replace(/"/g,'""')+'"'}
function downloadCommentTemplate(){const headers=['Manual title','Page or section','Text from manual','Reviewer comment','Suggested change','Also affects these manual titles (use |)'];const blob=new Blob(['\uFEFF'+headers.map(csvCell).join(',')+'\r\n'],{type:'text/csv;charset=utf-8'});const link=document.createElement('a');link.href=URL.createObjectURL(blob);link.download='JHCSC_Manual_Review_Comment_Template.csv';document.body.append(link);link.click();link.remove();URL.revokeObjectURL(link.href)}
function importCommentTemplate(btn){
  const file=$('commentCsv').files[0];if(!file){toast('Choose the completed CSV template first.',true);return}
  return run(btn,async()=>{const r=await call('mrmsImportCommentCsv',await file.text());state=r.data;$('commentCsv').value='';currentView='issues';draw();toast(`${r.count} comment(s) imported into Issues.`+(r.skipped?` ${r.skipped} already in the list were skipped.`:''))});
}
/* ---------- comfort settings: text size and high contrast (remembered on this device) ---------- */
function applyComfort(){
  const size=store.get('mrmsTextSize')||'100',hc=store.get('mrmsContrast')==='high';
  document.documentElement.style.fontSize=size+'%';
  document.documentElement.dataset.contrast=hc?'high':'normal';
  document.querySelectorAll('[data-size]').forEach(b=>{const on=b.dataset.size===size;b.setAttribute('aria-pressed',String(on));b.classList.toggle('on',on)});
  const c=$('contrastBtn');if(c){c.setAttribute('aria-pressed',String(hc));c.textContent=hc?'High contrast: on':'High contrast: off'}
}
document.addEventListener('click',e=>{const b=e.target.closest&&e.target.closest('[data-size]');if(b){store.set('mrmsTextSize',b.dataset.size);applyComfort()}});
function toggleContrast(){store.set('mrmsContrast',store.get('mrmsContrast')==='high'?'normal':'high');applyComfort()}
applyComfort();
refresh();
