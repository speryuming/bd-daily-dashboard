const $ = s => document.querySelector(s);
let issues = [], selectedId = null, uploadToken = null, currentView = 'work';
let auth = {user:null,manager:null,setup_needed:false}, users=[];
const e = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const val = id => $(id).value.trim();
const short = value => value ? e(value) : '—';
const safeLink = value => { try { const u = new URL(value); return ['http:','https:'].includes(u.protocol) ? u.href : ''; } catch { return ''; } };
const localToday = () => { const d=new Date(); return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`; };
function setManualDefaults(){ $('#manual-date').value=localToday(); $('#manual-due-date').value=localToday(); }

async function request(path, data, form=false) {
  const options = data === undefined ? {} : {method:'POST',headers:{'X-Rectification-App':'1'},body:form ? data : JSON.stringify(data)};
  if (data !== undefined && !form) options.headers['Content-Type']='application/json';
  const res = await fetch(path,options);
  const body = await res.json();
  if (!res.ok) throw new Error(body.error || '操作失败');
  return body;
}

function toast(message) {
  const node=$('#toast'); node.textContent=message; node.classList.add('show');
  clearTimeout(toast.timer); toast.timer=setTimeout(()=>node.classList.remove('show'),3200);
}

async function refreshAuth() {
  auth=await request('/api/auth');
  $('#manager-button').textContent=auth.user?`${auth.user.display_name} · ${auth.user.role==='manager'?'管理':'市场'}`:'账号登录';
  $('#manager-auth-form').hidden=!!auth.user;
  $('#account-tools').hidden=!auth.user;
  $('#add-manager-form').hidden=!auth.manager;
  $('#manager-auth-submit').textContent=auth.setup_needed?'创建首位管理人员':'登录';
  $('#manager-hint').textContent=auth.user?(auth.manager?'管理人员可导入、分派和复核。':'市场人员可反馈分配给自己的问题。'):auth.setup_needed?'首次使用请设置管理人员账号和密码。':'请输入账号和密码。';
  $('#signed-manager').textContent=auth.user?`当前登录：${auth.user.display_name}（${auth.user.role==='manager'?'管理人员':'市场人员'}）`:'';
  document.querySelector('[data-view="import"]').hidden=!auth.manager;
  $('#go-import').hidden=!auth.manager;
  $('#close-manager').hidden=!auth.user;
  users=auth.manager?(await request('/api/users')).items:[];
  const manualOwner=$('#manual-owner');
  if(manualOwner) manualOwner.innerHTML='<option value="">暂不分派</option>'+users.filter(x=>x.active).map(x=>`<option value="${e(x.display_name)}">${e(x.display_name)}（${x.role==='manager'?'管理':'市场'}）</option>`).join('');
  const accountList=$('#account-list');
  if(accountList) {
    accountList.innerHTML=auth.manager?`<h3>现有账号</h3>${users.map(x=>`<div class="account-row"><div><strong>${e(x.display_name)}</strong><span>${e(x.username)} · ${x.role==='manager'?'管理人员':'市场人员'}</span></div>${x.username!==auth.user.username?`<button class="secondary reset-password" data-id="${x.id}" data-name="${e(x.display_name)}">重置密码</button>`:''}</div>`).join('')}`:'';
    document.querySelectorAll('.reset-password').forEach(button=>button.onclick=()=>{
      const dialog=$('#reset-password-dialog'),form=$('#reset-password-form');
      form.reset();form.elements.user_id.value=button.dataset.id;
      $('#reset-password-name').textContent=button.dataset.name;
      dialog.showModal();
    });
  }
  if(!auth.user && !$('#manager-dialog').open) $('#manager-dialog').showModal();
}

function showView(view) {
  currentView=view;
  $('#work-view').hidden=view!=='work'; $('#import-view').hidden=view!=='import';
  document.querySelectorAll('.nav').forEach(x=>x.classList.toggle('active',x.dataset.view===view));
  if(view==='work') return refresh();
}

async function refresh() {
  if(!auth.user) return;
  try {
    const [list,summary]=await Promise.all([request('/api/issues'),request('/api/summary')]);
    issues=list.items;
    $('#today').textContent=summary.today;
    const counts=summary.counts;
    const cards=[['待反馈',counts['待反馈']||0,'warn'],['待复核',counts['待复核']||0,''],['待核查',counts['待核查']||0,''],['已通过',counts['通过']||0,'good'],['待映射',summary.unmapped,'']];
    $('#stats').innerHTML=cards.map(([label,n,css])=>`<div class="stat ${css}"><span>${label}</span><b>${n}</b></div>`).join('');
    renderList();
    if(selectedId) await loadDetail(selectedId);
  } catch(err) {toast(err.message);}
}

function renderList() {
  const status=val('#status-filter'), q=val('#search').toLowerCase();
  const filtered=issues.filter(i=>(status==='全部'||i.status===status)&&(!q||[i.raw_store_name,i.issue_text,i.owner].some(x=>String(x||'').toLowerCase().includes(q))));
  $('#list-count').textContent=`${filtered.length} 条`;
  $('#issue-list').innerHTML=filtered.length ? filtered.map(i=>`<button class="issue ${i.id===selectedId?'active':''}" data-id="${i.id}"><div class="issue-top"><span class="issue-title">${e(i.raw_store_name)}</span><span class="badge" data-status="${e(i.status)}">${e(i.status)}</span></div><div class="issue-desc">${e(i.issue_text)}</div><div class="issue-bottom"><span>${e(i.issue_type)} · ${e(i.owner||'待分派')}</span><span>${e(i.discovered_date)} ${i.mapping_id?'· 已映射':'· 未映射'}</span></div></button>`).join('') : '<div class="no-data">暂无符合条件的问题</div>';
  document.querySelectorAll('.issue').forEach(x=>x.onclick=()=>loadDetail(Number(x.dataset.id)));
}

async function loadDetail(id) {
  selectedId=id; renderList();
  try { const i=await request(`/api/issues/${id}`); renderDetail(i); }
  catch(err){toast(err.message);}
}

function eventHtml(item) {
  const url=safeLink(item.evidence);
  return `<li><div class="event-head">${e(item.kind)}${item.action?' · '+e(item.action):''}<span class="event-date">${e(item.created_at.replace('T',' ').slice(0,16))}</span></div><p>${e(item.actor||'系统')}${item.note?'：'+e(item.note):''}</p>${url?`<a href="${e(url)}" target="_blank" rel="noopener">查看证据</a>`:item.evidence?`<p>证据：${e(item.evidence)}</p>`:''}</li>`;
}

function renderDetail(i) {
  const mapped=i.mapping_id ? `<strong>${short(i.own_store_name||i.raw_store_name)} ↔ ${short(i.competitor_store_name)}</strong><span class="hint">我方商户ID：${short(i.own_merchant_id)}</span>` : '<span class="badge">尚未映射，不影响整改</span>';
  const sourceParts=[i.product_name&&`商品：${e(i.product_name)}`,i.competitor_sales&&`竞对销量：${e(i.competitor_sales)}`,i.own_sales&&`我方销量：${e(i.own_sales)}`,i.competitor_price&&`竞对价格：${e(i.competitor_price)}`,i.own_price&&`我方价格：${e(i.own_price)}`].filter(Boolean);
  const staffUsers=users.filter(x=>x.active).map(x=>`<option value="${e(x.display_name)}" ${i.owner===x.display_name?'selected':''}>${e(x.display_name)}（${x.role==='manager'?'管理':'市场'}）</option>`).join('');
  const canFeedback=auth.manager||[auth.user.username,auth.user.display_name].includes(i.owner);
  $('#detail-panel').innerHTML=`<div class="detail-content"><div class="detail-head"><span class="badge" data-status="${e(i.status)}">${e(i.status)}</span><h2>${e(i.raw_store_name)}</h2><div class="detail-sub">问题 #${i.id} · ${e(i.issue_type)} · 发现于 ${e(i.discovered_date)}</div></div><div class="detail-text">${e(i.issue_text)}</div>${sourceParts.length?`<div class="hint">${sourceParts.join('　·　')}</div>`:''}<div class="meta"><div><span>负责人</span><strong>${short(i.owner||'待分派')}</strong></div><div><span>整改截止</span><strong>${short(i.due_date)}</strong></div><div><span>计划复核</span><strong>${short(i.review_due_date)}</strong></div><div><span>原表位置</span><strong>${e(i.source_sheet||'人工新增')} ${i.source_row?'第'+i.source_row+'行':''}</strong></div></div>
  <h3 class="section-title">分派负责人</h3>${auth.manager?`<form id="assign-form" class="detail-form"><div class="two"><label>责任人<select name="owner" required><option value="">请选择账号</option>${i.owner&&!users.some(x=>x.display_name===i.owner)?`<option selected>${e(i.owner)}</option>`:''}${staffUsers}</select></label><label>截止日期<input name="due_date" type="date" value="${e(i.due_date)}"></label></div><button class="secondary">保存分派</button></form>`:`<div class="review-locked">负责人由管理人员分派。</div>`}
  <h3 class="section-title">晚间整改反馈</h3>${canFeedback?`<form id="feedback-form" class="detail-form"><div class="two"><label>反馈人<input value="${e(auth.user.display_name)}" readonly></label><label>处理动作<select name="action"><option>已上架商品</option><option>已报名活动</option><option>已调整价格</option><option>已调整配送费</option><option>已调整起送价</option><option>持续沟通</option><option>暂无法整改</option><option>其他</option></select></label></div><label>处理说明<textarea name="note" placeholder="写明做了什么，以及未完成的原因" required></textarea></label><label>证据链接或文件名<input name="evidence" placeholder="截图链接或文件名"></label><button class="primary">提交反馈</button></form>`:`<div class="review-locked">此问题尚未分配给你，不能提交反馈。</div>`}
  <h3 class="section-title">次日运营复核</h3><form id="review-form" class="detail-form"><div class="two"><label>复核人<input name="actor" placeholder="姓名" required></label><label>复核结论<select name="result"><option>通过</option><option>未通过</option><option>待核查</option><option>例外待决策</option></select></label></div><label>实测结果与原因<textarea name="note" placeholder="写明次日看到的商品、价格或配送费，以及判断依据" required></textarea></label><label>复核证据链接或文件名<input name="evidence" placeholder="截图链接或文件名"></label><button class="primary">提交复核</button><span class="hint">通过结论须在计划复核日及之后提交。</span></form>
  ${auth.manager?`<details class="mapping"><summary>人工映射店铺（可后补）</summary><p class="hint">${mapped}</p><form id="mapping-form" class="detail-form"><div class="two"><label>我方商户ID<input name="own_merchant_id" placeholder="手工确认" required></label><label>我方店铺名称<input name="own_store_name" value="${e(i.raw_store_name)}"></label></div><label>美团店铺名称<input name="competitor_store_name" required></label><label>美团店铺链接或标识<input name="competitor_url"></label><div class="two"><label>配对证据<input name="evidence" placeholder="截图或备注"></label><label>确认人<input name="confirmed_by" value="${e(auth.user.display_name)}" readonly required></label></div><button class="secondary">保存人工配对</button></form><div id="existing-mappings"></div></details>`:''}
  <h3 class="section-title">跟进记录</h3><ol class="timeline">${i.events.length?i.events.map(eventHtml).join(''):'<li>暂无记录</li>'}</ol></div>`;
  if(auth.manager) bindIssueForm('assign-form',i.id,'assign',['owner','due_date']);
  if(canFeedback) bindIssueForm('feedback-form',i.id,'feedback',['action','note','evidence']);
  bindIssueForm('review-form',i.id,'review',['actor','result','note','evidence']);
  if(auth.manager) bindIssueForm('mapping-form',i.id,'map',['own_merchant_id','own_store_name','competitor_store_name','competitor_url','evidence','confirmed_by']);
  if(auth.manager){
    const actor=$('#review-form [name="actor"]');
    actor.value=auth.manager.display_name;actor.readOnly=true;
  } else {
    $('#review-form').outerHTML='<div class="review-locked">仅管理人员可以提交复核。市场人员可继续填写上方的整改反馈。</div>';
  }
  const mappingDetails=document.querySelector('details.mapping');
  if(mappingDetails) mappingDetails.ontoggle=event=>{if(event.target.open) loadMappings(i);};
}

function bindIssueForm(formId,id,operation,fields) {
  const formNode=$(`#${formId}`); if(!formNode)return;
  formNode.onsubmit=async event=>{
    event.preventDefault();
    const form=new FormData(event.target),data={}; fields.forEach(k=>data[k]=String(form.get(k)||'').trim());
    try {await request(`/api/issues/${id}/${operation}`,data); toast('已保存'); await refresh();}
    catch(err){toast(err.message);}
  };
}

async function loadMappings(i) {
  try {
    const result=await request('/api/mappings');
    const matches=result.items.filter(x=>x.raw_store_name===i.raw_store_name);
    $('#existing-mappings').innerHTML=matches.length?`<p class="hint">已有人工配对（选择后仍需确认）：</p>${matches.map(m=>`<button class="secondary mapping-choice" data-id="${m.id}">#${m.id} ${e(m.own_store_name)} ↔ ${e(m.competitor_store_name)}</button>`).join(' ')}`:'';
    document.querySelectorAll('.mapping-choice').forEach(button=>button.onclick=async()=>{
      const actor=prompt('请填写本次确认人姓名'); if(!actor) return;
      try {await request(`/api/issues/${i.id}/link-mapping`,{mapping_id:Number(button.dataset.id),actor});toast('已关联人工映射');await refresh();}
      catch(err){toast(err.message);}
    });
  } catch(err){toast(err.message);}
}

async function upload() {
  const file=$('#file').files[0]; if(!file)return;
  const form=new FormData();form.append('file',file);
  $('#upload-result').textContent='正在读取工作簿…';
  $('#preview').disabled=true;$('#commit').disabled=true;uploadToken=null;
  try {
    const result=await request('/api/upload',form,true);uploadToken=result.token;
    $('#upload-result').textContent=`已读取：${result.filename}`;
    $('#sheet').innerHTML=result.sheets.map(s=>`<option value="${e(s)}">${e(s)}</option>`).join('');
    const best=result.sheets.find(s=>s==='9月')||result.sheets.find(s=>s.includes('9月'));
    if(best) $('#sheet').value=best;
    $('#sheet').disabled=false;$('#preview').disabled=false;
  } catch(err){$('#upload-result').textContent=err.message;toast(err.message);}
}

async function preview() {
  if(!uploadToken)return;
  try {
    const body={token:uploadToken,sheet:val('#sheet'),year:Number(val('#year')),fallback_date:val('#fallback-date')};
    const result=await request('/api/import-preview',body);
    $('#preview-count').textContent=`${result.total} 条问题${result.needs_date?' · '+result.needs_date+' 条缺日期':''}`;
    $('#preview-list').innerHTML=result.rows.length?result.rows.map(r=>`<div class="preview-item"><strong>${e(r.raw_store_name)}</strong> <span class="badge">${e(r.issue_type)}</span><p>${e(r.issue_text)}</p><small>第 ${r.source_row} 行 · ${e(r.discovered_date||'日期待补')}${r.date_assumed?'（补填）':''} · ${e(r.owner||'待分派')}</small></div>`).join(''):'<div class="preview-empty">该表没有可导入的问题</div>';
    $('#commit').disabled=result.total===0 || result.needs_date>0;
    message(result.needs_date?'有问题缺少有效日期，请在源表中补齐后重新上传。':'预览完成，请确认工作表和问题拆分结果。',!!result.needs_date);
  } catch(err){message(err.message,true);toast(err.message);}
}

function message(text,error=false){const node=$('#import-message');node.hidden=false;node.textContent=text;node.classList.toggle('error',error);}

async function commit() {
  if(!uploadToken)return;
  if(!confirm(`确认导入“${val('#sheet')}”中的问题？`))return;
  try {
    const result=await request('/api/import-commit',{token:uploadToken,sheet:val('#sheet'),year:Number(val('#year')),fallback_date:val('#fallback-date')});
    const x=result.result;message(`导入完成：新建 ${x.created} 条，接续旧问题 ${x.reused} 条，重复导入 ${x.already_imported} 条。`);
    toast('导入完成'); await refresh();
  } catch(err){message(err.message,true);toast(err.message);}
}

document.querySelectorAll('.nav').forEach(x=>x.onclick=()=>showView(x.dataset.view));
$('#go-import').onclick=()=>showView('import');
$('#status-filter').onchange=renderList;
$('#search').oninput=renderList;
$('#file').onchange=upload;
for(const field of ['#sheet','#year','#fallback-date']) $(field).onchange=()=>{$('#commit').disabled=true;};
$('#preview').onclick=preview;
$('#commit').onclick=commit;
$('#manual-issue-form').onsubmit=async event=>{
  event.preventDefault();
  const form=new FormData(event.target),data={};
  ['discovered_date','raw_store_name','issue_type','owner','due_date','product_name','issue_text','competitor_price','own_price','competitor_sales','own_sales'].forEach(k=>data[k]=String(form.get(k)||'').trim());
  try {
    const result=await request('/api/issues/manual',data);
    event.target.reset();setManualDefaults();toast('问题已添加');
    selectedId=result.id;await showView('work');await loadDetail(result.id);
  } catch(err){toast(err.message);}
};
$('#manager-button').onclick=()=>$('#manager-dialog').showModal();
$('#close-manager').onclick=()=>$('#manager-dialog').close();
$('#manager-auth-form').onsubmit=async event=>{
  event.preventDefault();
  const form=new FormData(event.target);
  const username=String(form.get('username')||'').trim(),password=String(form.get('password')||'');
  try {
    if(auth.setup_needed) await request('/api/setup-manager',{username,password,display_name:username});
    await request('/api/login',{username,password});
    event.target.reset();await refreshAuth();$('#manager-dialog').close();
    if(selectedId) await loadDetail(selectedId);
    toast('已登录');
  } catch(err){toast(err.message);}
};
$('#add-manager-form').onsubmit=async event=>{
  event.preventDefault();const form=new FormData(event.target);
  try {await request('/api/users',{username:String(form.get('username')||'').trim(),display_name:String(form.get('display_name')||'').trim(),role:String(form.get('role')||'staff'),password:String(form.get('password')||'')});event.target.reset();await refreshAuth();toast('已添加账号');}
  catch(err){toast(err.message);}
};
$('#close-reset-password').onclick=()=>$('#reset-password-dialog').close();
$('#reset-password-form').onsubmit=async event=>{
  event.preventDefault();const form=new FormData(event.target);
  const password=String(form.get('password')||''),confirmation=String(form.get('password_confirm')||'');
  if(password!==confirmation){toast('两次输入的密码不一致');return;}
  try {
    await request(`/api/users/${Number(form.get('user_id'))}/reset-password`,{password});
    event.target.reset();$('#reset-password-dialog').close();toast('密码已重置，请将新密码告知该人员');
  } catch(err){toast(err.message);}
};
$('#logout-manager').onclick=async()=>{try{await request('/api/logout',{});await refreshAuth();$('#manager-dialog').close();if(selectedId)await loadDetail(selectedId);toast('已退出');}catch(err){toast(err.message);}};
setManualDefaults();
refreshAuth().then(refresh).catch(err=>toast(err.message));
