const defaults=[[2,1,3,0,0,0],[3,1,3,0,0,0],[3,2,3,0,0,0],[3,3,3,0,0,0]];
const tbody=document.querySelector('#mapping tbody');defaults.forEach((row,r)=>{const tr=document.createElement('tr');tr.innerHTML=`<th>CO${r+1}</th>`+row.map((v,c)=>`<td><input type="number" min="0" max="3" value="${v}" data-r="${r}" data-c="${c}"></td>`).join('');tbody.appendChild(tr)});
const outcomes=(prefix,count)=>Array.from({length:count},(_,i)=>`${prefix}${i+1}`);
function addTargetRows(tableId,names,initial){
 const body=document.querySelector(`#${tableId} tbody`);
 [['Previous-cycle target','previous'],['Previous-cycle attainment','attained'],['Suggested current target','current']].forEach(([label,type])=>{
  const tr=document.createElement('tr'); tr.innerHTML=`<th>${label}</th>`+names.map((_,i)=>`<td><input type="number" min="0" max="3" step="0.01" value="${type==='previous'?initial:''}" data-type="${type}" data-i="${i}" ${type==='current'?'class="suggested"':''}></td>`).join(''); body.appendChild(tr);
 });
 body.addEventListener('input',()=>refreshTargets(tableId));
}
function recommended(previous,attained){
 if(!Number.isFinite(previous)||!Number.isFinite(attained))return '';
 if(attained<previous)return previous.toFixed(2).replace(/0+$/,'').replace(/\.$/,'');
 return Math.min(3,Math.round((attained+0.05)*10)/10).toFixed(2).replace(/0+$/,'').replace(/\.$/,'');
}
function refreshTargets(tableId){
 const table=document.querySelector(`#${tableId}`), prev=[...table.querySelectorAll('[data-type="previous"]')], attained=[...table.querySelectorAll('[data-type="attained"]')];
 table.querySelectorAll('[data-type="current"]').forEach((field,i)=>{if(document.activeElement!==field)field.value=recommended(Number(prev[i].value),Number(attained[i].value));});
}
function targetData(tableId){const table=document.querySelector(`#${tableId}`);return ['previous','attained','current'].reduce((o,type)=>(o[type]=[...table.querySelectorAll(`[data-type="${type}"]`)].map(x=>x.value===''?null:+x.value),o),{});}
addTargetRows('coTargets',outcomes('CO',4),2.4); addTargetRows('poTargets',outcomes('PO',6),2.0);
const form=document.querySelector('#form'),status=document.querySelector('#status'),button=form.querySelector('button'),result=document.querySelector('#result');
form.addEventListener('submit',async e=>{e.preventDefault();const mapping=defaults.map(row=>[...row]);document.querySelectorAll('#mapping input').forEach(x=>mapping[+x.dataset.r][+x.dataset.c]=+x.value);document.querySelector('#mappingValue').value=JSON.stringify(mapping);document.querySelector('#coTargetsValue').value=JSON.stringify(targetData('coTargets'));document.querySelector('#poTargetsValue').value=JSON.stringify(targetData('poTargets'));button.disabled=true;status.textContent='Validating files and calculating attainment…';result.classList.add('hidden');try{const response=await fetch('/api/generate',{method:'POST',body:new FormData(form)});const data=await response.json();if(!response.ok)throw new Error(data.error||'Generation failed');document.querySelector('#metrics').innerHTML=`<div class="metrics"><div class="metric">Students<b>${data.students}</b></div><div class="metric">Survey responses<b>${data.survey_responses}</b></div><div class="metric">Overall COs<b>${data.overall.join(', ')}</b></div><div class="metric">Status<b>${data.co_status.join(', ')}</b></div></div>`;document.querySelector('#warnings').innerHTML=(data.warnings||[]).map(x=>`<div class="warning">${x}</div>`).join('');const link=document.querySelector('#download');link.href=data.download_url;result.classList.remove('hidden');status.textContent='Completed.';result.scrollIntoView({behavior:'smooth'});}catch(err){status.textContent='Error: '+err.message;}finally{button.disabled=false;}});
