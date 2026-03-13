
// ==================== UTILS ====================
const CN={retail:'소매',wholesale:'도매',digital:'디지털',enterprise:'기업/공공',iot:'IoT',corporate_sales:'법인영업',small_biz:'소상공인'};
const CC={retail:'#38bdf8',wholesale:'#34d399',digital:'#a78bfa',enterprise:'#fb923c',iot:'#fbbf24',corporate_sales:'#f87171',small_biz:'#818cf8'};
const CHS=['retail','wholesale','digital','enterprise','iot','corporate_sales','small_biz'];
const RC=['','#fbbf24','#94a3b8','#d97706','#38bdf8','#a78bfa'];
const fB=n=>(n/1e8).toFixed(1);
const fM=n=>Math.round(n/1e6).toLocaleString();
const fMr=n=>(n/1e6).toFixed(0);
const pct=n=>Math.round(n*100)+'%';
const pc=p=>p>=.8?'var(--g)':p>=.5?'var(--acc)':p>=.2?'var(--y)':'var(--r)';
const hmc=(v,arr)=>{const s=[...arr].sort((a,b)=>b-a),i=s.indexOf(v);return i<=1?'hm-h':i>=arr.length-2?'hm-l':'hm-m';};
const comma=n=>Math.round(n).toLocaleString();
const uploadHealth={task:'unknown',profit:'unknown',profitYtd:'unknown',kpi:'unknown',factbook:'unknown',hqprofit:'unknown',comm:'unknown',plan:'unknown',midterm:'unknown'};

// ==================== TAB ====================
let currentTab='dashboard';
function switchTab(btn){
  const t = (btn && btn.dataset) ? btn.dataset.t : btn;
  currentTab = t;
  document.querySelectorAll('.bni').forEach(x=>x.classList.remove('on'));
  document.querySelectorAll('.pnl').forEach(x=>x.classList.remove('on'));
  const activeBtn = typeof btn === 'string' ? document.querySelector(`.bni[data-t="${t}"]`) : btn;
  if(activeBtn){
    activeBtn.classList.add('on');
    activeBtn.scrollIntoView({behavior:'smooth', inline:'center', block:'nearest'});
  }
  const panel = document.getElementById('t-'+t);
  if(panel) panel.classList.add('on');
  try{
    if(t==='commission') initCommission();
    if(t==='simulate') initSimulationTab();
    if(t==='ai') initBriefing();
  }catch(err){
    console.error('tab render error:', t, err);
    if(t==='commission') renderCommissionFallback(err);
  }
}
function forceOpenCommissionTab(){
  const btn = document.querySelector('.bni[data-t="commission"]');
  switchTab(btn || 'commission');
}
(function initBottomNavDragScroll(){
  const nav = document.querySelector('.bnav');
  if(!nav) return;
  let isDown = false;
  let moved = false;
  let startX = 0;
  let startScrollLeft = 0;
  nav.addEventListener('mousedown', e=>{
    if(e.button!==0) return;
    isDown = true;
    moved = false;
    nav.classList.add('dragging');
    startX = e.pageX;
    startScrollLeft = nav.scrollLeft;
  });
  window.addEventListener('mousemove', e=>{
    if(!isDown) return;
    const walk = e.pageX - startX;
    if(Math.abs(walk)>4) moved = true;
    nav.scrollLeft = startScrollLeft - walk;
  });
  const endDrag=()=>{ isDown = false; nav.classList.remove('dragging'); };
  window.addEventListener('mouseup', endDrag);
  nav.addEventListener('mouseleave', ()=>{ if(isDown) endDrag(); });
  nav.addEventListener('click', e=>{
    if(!moved) return;
    e.preventDefault();
    e.stopPropagation();
    moved = false;
  }, true);
})();
function renderCommissionFallback(err){
  const panel = document.getElementById('t-commission');
  if(!panel) return;
  panel.classList.add('on');
  const target = document.getElementById('comm-s4') || panel;
  target.innerHTML = `
    <div style="background:linear-gradient(135deg,#eff6ff,#dbeafe);border:1px solid #93c5fd;border-radius:12px;padding:16px;margin-top:8px">
      <div style="font-size:14px;font-weight:800;color:#1d4ed8;margin-bottom:6px">수수료 탭 강제 복구 모드</div>
      <div style="font-size:12px;color:#1e3a8a;line-height:1.7">탭 렌더링 중 오류가 발생해서 기본 화면으로 복구했습니다.</div>
      <div style="font-size:11px;color:#334155;margin-top:8px;word-break:break-word">오류: ${err && err.message ? err.message : err}</div>
      <div style="margin-top:12px"><button class="rpt-bar-btn" style="background:#1d4ed8;color:#fff" onclick="initCommission()">다시 불러오기</button></div>
    </div>`;
}
function toggleSec(hdr){hdr.classList.toggle('closed');hdr.nextElementSibling.classList.toggle('hidden');}


// ==================== WHAT-IF SIMULATION ====================
const simulationState = {
  inited:false,
  scenarioInputs:{
    capa:0,netAdd:0,arpu:0,mgmt:0,policy:0,mkt:0,costEff:0,infl:0,kpi:0,task:0
  }
};

function ensureSimulationDataShape(){
  if(!D.planData) D.planData = { sourceFile:'', monthly:[], bm:[], kpiTargets:[] };
  if(!D.midtermAssumptions) D.midtermAssumptions = { sourceFile:'', assumptions:[], bm:[] };
  if(!D.scenarioState) D.scenarioState = { preset:'base', inputs:{...simulationState.scenarioInputs}, simulatedData:null };
}

function getScenarioInputMap(){
  return [
    ['simCapa','capa'],['simNetAdd','netAdd'],['simArpu','arpu'],['simMgmt','mgmt'],['simPolicy','policy'],
    ['simMkt','mkt'],['simCostEff','costEff'],['simInfl','infl'],['simKpi','kpi'],['simTask','task']
  ];
}

function readScenarioInputs(){
  const obj={};
  getScenarioInputMap().forEach(([id,key])=>{ obj[key]=Number(document.getElementById(id)?.value||0)/100; });
  return obj;
}

function setScenarioInputs(inputs){
  getScenarioInputMap().forEach(([id,key])=>{
    const v = Number(inputs[key]||0)*100;
    const el=document.getElementById(id);
    if(el) el.value = Math.max(Number(el.min||-100), Math.min(Number(el.max||100), Math.round(v)));
  });
}

function buildSimulationBaseline(){
  ensureSimulationDataShape();
  const revActual = Number(D?.profit?.revenue?.total)||0;
  const opActual = Number(D?.profit?.op?.total)||0;
  const baseMargin = revActual>0 ? opActual/revActual : 0;
  const wireless = subscriberData?.wireless || DEFAULT_SUBSCRIBER_DATA?.wireless || {};
  const netAddActual = Number(wireless.netAdd)||0;
  const capaActual = Number(wireless.capa)||0;
  const arpuActual = Number(subscriberData?.arpu?.overall||0);
  const jan = D?.commission?.jan26||{};
  const mgmtActual = Number(jan?.mgmt_fee?.total)||0;
  const policyActual = Number(jan?.policy_fee?.total)||0;
  const kpiScore = (D.kpi||[]).length ? (D.kpi.reduce((s,r)=>s+(Number(r.ts)||0),0)/D.kpi.length) : 0;
  const taskCompletion = (D.tasks||[]).length ? (D.tasks.filter(t=>t.st==='완료').length/D.tasks.length) : 0;

  const planMonthly = Array.isArray(D.planData.monthly) && D.planData.monthly.length
    ? D.planData.monthly
    : Array.from({length:12},(_,i)=>({month:(i+1)+'월', revenue:revActual, op:opActual, netAdd:netAddActual, capa:capaActual, mgmtFee:mgmtActual, policyFee:policyActual, arpu:arpuActual}));

  const bmSource = D.planData.bm?.length ? D.planData.bm : CHS.map(ch=>({
    key:ch,
    name:CN[ch],
    revenue:Number(D?.profit?.revenue?.[ch])||0,
    op:Number(D?.profit?.op?.[ch])||0,
    margin:(Number(D?.profit?.revenue?.[ch])||0)>0?((Number(D?.profit?.op?.[ch])||0)/(Number(D?.profit?.revenue?.[ch])||1)):0
  }));

  return {
    actual:{ revenue:revActual, op:opActual, margin:baseMargin, netAdd:netAddActual, capa:capaActual, arpu:arpuActual, mgmtFee:mgmtActual, policyFee:policyActual, kpi:kpiScore, task:taskCompletion },
    planMonthly,
    bm: bmSource,
    assumptions: D.midtermAssumptions
  };
}

function runScenarioSimulation(baseline, inputs){
  const actual = baseline.actual;
  const monthly = baseline.planMonthly.map((m,idx)=>{
    const monthW = 1 + (idx/11)*0.12;
    const planRev = Number(m.revenue)||0;
    const planOp = Number(m.op)||0;
    const planMargin = planRev>0 ? planOp/planRev : actual.margin;
    const planNet = Number(m.netAdd)||actual.netAdd;
    const planCapa = Number(m.capa)||actual.capa;
    const planMgmt = Number(m.mgmtFee)||actual.mgmtFee;
    const planPolicy = Number(m.policyFee)||actual.policyFee;
    const planArpu = Number(m.arpu)||actual.arpu;

    const revGrowth = inputs.capa*0.42 + inputs.netAdd*0.28 + inputs.arpu*0.30;
    const commImpact = inputs.mgmt*(planMgmt/Math.max(planRev,1)) + inputs.policy*(planPolicy/Math.max(planRev,1));
    const riskDrag = inputs.mkt<0 ? Math.abs(inputs.mkt)*0.25 : 0;
    const scenarioRev = Math.max(0, planRev * (1 + revGrowth + commImpact - riskDrag) * monthW);

    const opLiftFromRev = (scenarioRev-planRev) * planMargin;
    const costSaving = planOp * inputs.costEff;
    const inflationPenalty = planOp * inputs.infl * 0.9;
    const marketingPenalty = planOp * Math.max(0, inputs.mkt) * 0.45;
    const scenarioOp = planOp + opLiftFromRev + costSaving - inflationPenalty - marketingPenalty;
    const scenarioMargin = scenarioRev>0 ? scenarioOp/scenarioRev : 0;

    const scenarioNet = Math.round(planNet * (1 + inputs.netAdd + Math.max(0, inputs.mkt)*0.35 - Math.max(0, inputs.infl)*0.12));
    const scenarioCapa = Math.round(planCapa * (1 + inputs.capa));
    const scenarioMgmt = planMgmt * (1 + inputs.mgmt + inputs.netAdd*0.2);
    const scenarioPolicy = planPolicy * (1 + inputs.policy + Math.max(0,inputs.mkt)*0.15);

    return {
      month:m.month || `${idx+1}월`,
      monthIndex:parseMonthIndex(m.monthIndex ?? m.month ?? (idx+1)) || (idx+1),
      actualRev:actual.revenue,
      planRev,
      scenarioRev,
      actualOp:actual.op,
      planOp,
      scenarioOp,
      planNet,
      scenarioNet,
      planCapa,
      scenarioCapa,
      planMgmt,
      scenarioMgmt,
      planPolicy,
      scenarioPolicy,
      scenarioMargin,
      planMargin
    };
  });

  const sum = (arr,key)=>arr.reduce((s,r)=>s+(Number(r[key])||0),0);
  const planRevenue = sum(monthly,'planRev');
  const planOp = sum(monthly,'planOp');
  const scenarioRevenue = sum(monthly,'scenarioRev');
  const scenarioOp = sum(monthly,'scenarioOp');
  const planNet = sum(monthly,'planNet');
  const scenarioNet = sum(monthly,'scenarioNet');

  const bmImpact = baseline.bm.map(b=>{
    const mktBoost = inputs.mkt>0 ? (b.name.includes('디지털')?0.8:0.35) : 0;
    const inflPenalty = inputs.infl * (b.name.includes('도매')?1.2:0.9);
    const marginDelta = inputs.costEff*0.55 - inflPenalty + mktBoost;
    const revDelta = inputs.capa*0.35 + inputs.arpu*0.4 + inputs.netAdd*0.25;
    const scenarioRev = (b.revenue||0) * (1 + revDelta);
    const scenarioMargin = Math.max(-0.15, (b.margin||0) + marginDelta);
    const scenarioOp = scenarioRev * scenarioMargin;
    return {
      key:b.key,
      name:b.name,
      planOp:Number(b.op)||0,
      scenarioOp,
      delta:scenarioOp-(Number(b.op)||0)
    };
  });

  const kpiRows = ['소매 KPI','도매 KPI','소상공 KPI','운영 KPI'];
  const kpiHeatmap = kpiRows.map((nm,i)=>{
    const base = (actual.kpi||70) + i*0.9;
    const plan = base + 2.5;
    const scenario = plan + (inputs.kpi*12) + (inputs.task*5) + (inputs.netAdd*3);
    return {name:nm, actual:base, plan, scenario};
  });

  return {
    summary:{
      actual:{revenue:actual.revenue*12, op:actual.op*12, margin:actual.margin, netAdd:actual.netAdd*12},
      plan:{revenue:planRevenue, op:planOp, margin:planRevenue>0?planOp/planRevenue:0, netAdd:planNet},
      scenario:{revenue:scenarioRevenue, op:scenarioOp, margin:scenarioRevenue>0?scenarioOp/scenarioRevenue:0, netAdd:scenarioNet}
    },
    monthly,
    bmImpact,
    kpiHeatmap,
    levers:inputs
  };
}

function generateExecutiveInsights(simData){
  const s = simData.summary;
  const revGap = s.scenario.revenue - s.plan.revenue;
  const opGap = s.scenario.op - s.plan.op;
  const topBm = [...simData.bmImpact].sort((a,b)=>Math.abs(b.delta)-Math.abs(a.delta))[0];
  const riskKpi = [...simData.kpiHeatmap].sort((a,b)=>a.scenario-b.scenario)[0];
  return [
    `핵심 변동 원인: ${topBm?.name||'-'} 손익 변화 ${fB(topBm?.delta||0)}억이 전사 영업이익 변동에 가장 크게 기여합니다.`,
    `가장 민감한 변수: 현재 가정에서 매출 Gap은 ${revGap>=0?'+':''}${fB(revGap)}억, 영업이익 Gap은 ${opGap>=0?'+':''}${fB(opGap)}억입니다.`,
    `목표 미달 위험 영역: KPI 중 ${riskKpi?.name||'-'}가 상대적으로 낮아 편차 관리가 필요합니다.`,
    `우선 액션 제안: CAPA·순증·비용효율의 동시 관리(채널별 실행계획 + 저효율 과제 보정)로 계획 달성률을 방어하세요.`
  ];
}

function parseMonthIndex(v){
  if(v===null||v===undefined) return null;
  if(Number.isFinite(v)){
    const n = Number(v);
    return n>=1&&n<=12 ? n : null;
  }
  const s = String(v).trim();
  const m = s.match(/(\d{1,2})\s*월?/);
  if(!m) return null;
  const n = Number(m[1]);
  return n>=1&&n<=12 ? n : null;
}

function safeRate(actual, plan){
  const p = Number(plan);
  if(!Number.isFinite(p) || p===0) return null;
  const a = Number(actual);
  if(!Number.isFinite(a)) return null;
  return a / p;
}

function buildMonthlyPlanVsActual(){
  const planRows = Array.isArray(D?.planData?.monthly) ? D.planData.monthly : [];
  const actualRows = Array.isArray(D?.profit?.monthly_trend) ? D.profit.monthly_trend : [];
  const planMap = new Map();
  const actualMap = new Map();

  planRows.forEach(r=>{
    const idx = parseMonthIndex(r?.monthIndex ?? r?.month);
    if(!idx) return;
    planMap.set(idx, { op:Number(r?.op), revenue:Number(r?.revenue) });
  });
  actualRows.forEach(r=>{
    const idx = parseMonthIndex(r?.monthIndex ?? r?.month ?? r?.m);
    if(!idx) return;
    actualMap.set(idx, {
      op:Number(r?.op ?? r?.total),
      revenue:Number(r?.revenue)
    });
  });

  const out=[];
  for(let i=1;i<=12;i++){
    const p = planMap.get(i) || {};
    const a = actualMap.get(i) || {};
    const planOp = Number.isFinite(p.op) ? p.op : null;
    const planRevenue = Number.isFinite(p.revenue) ? p.revenue : null;
    const actualOp = Number.isFinite(a.op) ? a.op : null;
    const actualRevenue = Number.isFinite(a.revenue) ? a.revenue : null;
    out.push({
      month:`${i}월`,
      monthIndex:i,
      plan:{op:planOp,revenue:planRevenue},
      actual:{op:actualOp,revenue:actualRevenue},
      gap:{
        op:(planOp===null||actualOp===null)?null:actualOp-planOp,
        revenue:(planRevenue===null||actualRevenue===null)?null:actualRevenue-planRevenue
      },
      achv:{
        op:safeRate(actualOp, planOp),
        revenue:safeRate(actualRevenue, planRevenue)
      }
    });
  }
  return out;
}

function renderWhatIfSummary(simData){
  const wrap = document.getElementById('simSummary');
  if(!wrap) return;
  const {actual,plan,scenario} = simData.summary;
  const monthlyComp = buildMonthlyPlanVsActual();
  const opAchvRows = monthlyComp.filter(r=>r.achv.op!==null);
  const avgOpAchv = opAchvRows.length ? (opAchvRows.reduce((s,r)=>s+r.achv.op,0)/opAchvRows.length) : null;
  const gapRev = scenario.revenue-plan.revenue;
  const gapOp = scenario.op-plan.op;
  const gapNet = scenario.netAdd-plan.netAdd;
  wrap.innerHTML = `
    <div class="sim-card"><div class="k">Actual (연환산)</div><div class="v">${fB(actual.revenue)}억</div><div class="d">영업이익 ${fB(actual.op)}억 · 마진 ${(actual.margin*100).toFixed(1)}%</div></div>
    <div class="sim-card"><div class="k">Plan 2026</div><div class="v">${fB(plan.revenue)}억</div><div class="d">영업이익 ${fB(plan.op)}억 · 순증 ${comma(plan.netAdd)}</div></div>
    <div class="sim-card"><div class="k">Scenario</div><div class="v">${fB(scenario.revenue)}억</div><div class="d">영업이익 ${fB(scenario.op)}억 · 마진 ${(scenario.margin*100).toFixed(1)}%</div></div>
    <div class="sim-card"><div class="k">Gap (Scenario-Plan)</div><div class="v" style="color:${gapOp>=0?'#86efac':'#fca5a5'}">${gapOp>=0?'+':''}${fB(gapOp)}억</div><div class="d" style="color:${gapRev>=0?'#86efac':'#fca5a5'}">매출 ${gapRev>=0?'+':''}${fB(gapRev)}억 · 순증 ${gapNet>=0?'+':''}${comma(gapNet)}</div><div class="d" style="margin-top:4px;color:#cbd5e1">월평균 계획 대비 OP 달성률: ${avgOpAchv===null?'-':(avgOpAchv*100).toFixed(1)+'%'}</div></div>`;
}

function renderWhatIfCharts(simData){
  const monthlyEl = document.getElementById('simMonthlyChart');
  if(monthlyEl){
    const monthlyComp = buildMonthlyPlanVsActual();
    const compMap = new Map(monthlyComp.map(r=>[r.monthIndex,r]));
    const mx = Math.max(...simData.monthly.map((m,idx)=>{
      const monthIndex = parseMonthIndex(m.monthIndex ?? m.month ?? (idx+1));
      const comp = compMap.get(monthIndex)||{};
      const actualRev = Number(comp?.actual?.revenue);
      return Math.max(Number.isFinite(actualRev)?actualRev:0,m.planRev,m.scenarioRev);
    }),1);
    monthlyEl.innerHTML = simData.monthly.map((m,idx)=>{
      const monthIndex = parseMonthIndex(m.monthIndex ?? m.month ?? (idx+1));
      const comp = compMap.get(monthIndex)||{};
      const actualRev = Number(comp?.actual?.revenue);
      const actualWidth = Number.isFinite(actualRev) ? (actualRev/mx*100) : 0;
      const opAchvLabel = comp?.achv?.op===null||comp?.achv?.op===undefined ? '-' : (comp.achv.op*100).toFixed(1)+'%';
      return `
      <div class="sim-row">
        <div class="m">${m.month}</div>
        <div class="sim-track"><i style="width:${actualWidth}%;background:#64748b"></i><i style="width:${m.planRev/mx*100}%;background:#0ea5e9"></i><i style="width:${m.scenarioRev/mx*100}%;background:#22c55e"></i></div>
        <div class="n">P ${fB(m.planRev)}</div>
        <div class="n">S ${fB(m.scenarioRev)} · 달성 ${opAchvLabel}</div>
      </div>`;
    }).join('');
  }

  const bmEl = document.getElementById('simBmChart');
  if(bmEl){
    const mx = Math.max(...simData.bmImpact.map(b=>Math.abs(b.delta)),1);
    bmEl.innerHTML = simData.bmImpact.map(b=>`
      <div class="sim-bm-row"><div>${b.name}</div><div class="bar"><div class="fill" style="width:${Math.abs(b.delta)/mx*100}%;background:${b.delta>=0?'linear-gradient(90deg,#34d399,#22c55e)':'linear-gradient(90deg,#f97316,#ef4444)'}"></div></div><div style="font-size:10px;text-align:right;color:${b.delta>=0?'#86efac':'#fca5a5'}">${b.delta>=0?'+':''}${fB(b.delta)}</div></div>`).join('');
  }

  const heat = document.getElementById('simKpiHeatmap');
  if(heat){
    const cellClass=v=>v>=85?'good':v>=75?'mid':'bad';
    heat.innerHTML = `<div class="sim-heat"><div></div><div class="th">Actual</div><div class="th">Plan</div><div class="th">Scenario</div>${simData.kpiHeatmap.map(r=>`<div class="rh">${r.name}</div><div class="cell ${cellClass(r.actual)}">${r.actual.toFixed(1)}</div><div class="cell ${cellClass(r.plan)}">${r.plan.toFixed(1)}</div><div class="cell ${cellClass(r.scenario)}">${r.scenario.toFixed(1)}</div>`).join('')}</div>`;
  }

  const subComm = document.getElementById('simSubCommChart');
  if(subComm){
    const m = simData.monthly;
    const planNet=m.reduce((s,x)=>s+x.planNet,0), scNet=m.reduce((s,x)=>s+x.scenarioNet,0);
    const planMgmt=m.reduce((s,x)=>s+x.planMgmt,0), scMgmt=m.reduce((s,x)=>s+x.scenarioMgmt,0);
    const planPolicy=m.reduce((s,x)=>s+x.planPolicy,0), scPolicy=m.reduce((s,x)=>s+x.scenarioPolicy,0);
    subComm.innerHTML = `
      <div class="sim-bm-row"><div>순증</div><div class="bar"><div class="fill" style="width:100%;background:linear-gradient(90deg,#38bdf8,#22d3ee)"></div></div><div style="font-size:10px;text-align:right">P ${comma(planNet)} / S ${comma(scNet)}</div></div>
      <div class="sim-bm-row"><div>관리수수료</div><div class="bar"><div class="fill" style="width:${Math.min(100,Math.abs(scMgmt-planMgmt)/Math.max(planMgmt,1)*100)}%;background:linear-gradient(90deg,#a78bfa,#6366f1)"></div></div><div style="font-size:10px;text-align:right">${fB(scMgmt-planMgmt)}억</div></div>
      <div class="sim-bm-row"><div>정책수수료</div><div class="bar"><div class="fill" style="width:${Math.min(100,Math.abs(scPolicy-planPolicy)/Math.max(planPolicy,1)*100)}%;background:linear-gradient(90deg,#f59e0b,#f97316)"></div></div><div style="font-size:10px;text-align:right">${fB(scPolicy-planPolicy)}억</div></div>`;
  }
}

function renderWhatIfInsights(simData){
  const insightEl=document.getElementById('simInsights');
  if(!insightEl) return;
  const lines = generateExecutiveInsights(simData);
  insightEl.innerHTML = `<ul>${lines.map(x=>`<li>${x}</li>`).join('')}</ul>`;

  const note=document.getElementById('simNote');
  if(note){
    note.textContent = `가정 로직: 매출은 CAPA/순증/ARPU/수수료 레버 가중 합으로, 영업이익은 매출 레버리지 + 비용효율 - 인플레/판촉비 부담으로 계산됩니다. (실무 시 parsePlanExcel/parseMidtermExcel 매핑 테이블로 조정 가능)`;
  }
}

function initSimulationTab(){
  ensureSimulationDataShape();
  if(!document.getElementById('simSummary')) return;
  if(!simulationState.inited){
    simulationState.inited = true;
    setScenarioInputs(simulationState.scenarioInputs);
  }
  updateSimulation();
}

function resetSimulationInputs(){
  const zero={capa:0,netAdd:0,arpu:0,mgmt:0,policy:0,mkt:0,costEff:0,infl:0,kpi:0,task:0};
  setScenarioInputs(zero);
  updateSimulation();
}

function applySimulationPreset(type){
  const presets={
    conservative:{capa:-0.08,netAdd:-0.06,arpu:-0.02,mgmt:-0.03,policy:-0.05,mkt:-0.08,costEff:0.03,infl:0.05,kpi:-0.04,task:-0.06},
    base:{capa:0,netAdd:0,arpu:0,mgmt:0,policy:0,mkt:0,costEff:0,infl:0,kpi:0,task:0},
    aggressive:{capa:0.12,netAdd:0.1,arpu:0.04,mgmt:0.06,policy:0.05,mkt:0.1,costEff:0.04,infl:0.01,kpi:0.06,task:0.08}
  };
  const p = presets[type]||presets.base;
  setScenarioInputs(p);
  D.scenarioState = { ...(D.scenarioState||{}), preset:type, inputs:p };
  updateSimulation();
}

function updateSimulation(){
  const inputs = readScenarioInputs();
  simulationState.scenarioInputs = inputs;
  Object.entries(inputs).forEach(([k,v])=>{
    const id='sim'+k.charAt(0).toUpperCase()+k.slice(1)+'Val';
    const el=document.getElementById(id);
    if(el) el.textContent = `${v>=0?'+':''}${Math.round(v*100)}%`;
  });

  const baseline = buildSimulationBaseline();
  const simData = runScenarioSimulation(baseline, inputs);
  renderWhatIfSummary(simData);
  renderWhatIfCharts(simData);
  renderWhatIfInsights(simData);
  D.scenarioState = { ...(D.scenarioState||{}), inputs, simulatedData:simData };
}

function switchKpiTab(tab){
  ['rt','wh','smb'].forEach(t=>{
    const panel=document.getElementById('kpiPanel-'+t);
    const btn=document.getElementById('kpiTab-'+t);
    if(!panel||!btn)return;
    const active=t===tab;
    panel.style.display=active?'block':'none';
    btn.style.background=active?'#3b82f6':'var(--bg2)';
    btn.style.color=active?'#fff':'var(--t2)';
  });
}

// ==================== TASKS ====================
let F={ch:'all',st:'all',tm:'all',cr:'all',q:''};
function initTaskUI(){
  const tms=[...new Set(D.tasks.map(t=>t.tm))].sort();
  document.getElementById('tmFilter').innerHTML='<option value="all">전체 팀</option>'+tms.map(t=>`<option value="${t}">${t}</option>`).join('');
  document.getElementById('chFilter').innerHTML=['all','경총','영총'].map(v=>`<button class="chip ${v==='all'?'on':''}" onclick="setF('ch','${v}',this)">${v==='all'?'전체':v}</button>`).join('');
  document.getElementById('stFilter').innerHTML=['all','진행중','계획','완료'].map(v=>`<button class="chip ${v==='all'?'on':''}" onclick="setF('st','${v}',this)">${v==='all'?'전체상태':v}</button>`).join('');
  const total=D.tasks.length,done=D.tasks.filter(t=>t.st==='완료').length,prog=D.tasks.filter(t=>t.st==='진행중').length,core=D.tasks.filter(t=>t.co==='●').length;
  document.getElementById('taskStats').innerHTML=`
    <div class="sc bl"><div class="sl">전체 과제</div><div class="sv">${total}</div><div class="ss">${tms.length}개 팀</div></div>
    <div class="sc gr"><div class="sl">완료</div><div class="sv" style="color:var(--g)">${done}</div><div class="ss">${(done/total*100).toFixed(1)}%</div></div>
    <div class="sc yl"><div class="sl">진행중</div><div class="sv" style="color:var(--y)">${prog}</div><div class="ss">${(prog/total*100).toFixed(1)}%</div></div>
    <div class="sc rd"><div class="sl">핵심과제</div><div class="sv" style="color:var(--r)">${core}</div><div class="ss">${(core/total*100).toFixed(1)}%</div></div>`;
  renderTasks();
}
function setF(k,v,el){F[k]=v;el.parentElement.querySelectorAll('.chip').forEach(b=>b.classList.remove('on'));el.classList.add('on');renderTasks();}
document.getElementById('tmFilter').onchange=e=>{F.tm=e.target.value;renderTasks();};
document.getElementById('crFilter').onchange=e=>{F.cr=e.target.value;renderTasks();};
document.getElementById('tSrch').oninput=e=>{F.q=e.target.value;renderTasks();};
function renderTasks(){
  const fl=D.tasks.filter(t=>{
    if(F.ch!=='all'&&t.ch!==F.ch)return false;if(F.st!=='all'&&t.st!==F.st)return false;
    if(F.tm!=='all'&&t.tm!==F.tm)return false;if(F.cr==='core'&&t.co!=='●')return false;
    if(F.q){const s=F.q.toLowerCase();if(!t.nm.toLowerCase().includes(s)&&!t.tm.toLowerCase().includes(s)&&!(t.mg||'').toLowerCase().includes(s))return false;}
    return true;
  });
  document.getElementById('tCnt').textContent=`${fl.length}개 과제`;
  document.getElementById('tBody').innerHTML=fl.map(t=>`
    <div class="task-card ${t.st}" onclick="showTask(${t.no})">
      <div class="tc-top">
        <div class="tc-title">${t.co?'<span style="color:var(--r)">● </span>':''}${t.nm}</div>
        <div class="tc-badges"><span class="pill ${t.st}">${t.st}</span></div>
      </div>
      <div class="tc-meta">
        <span class="tc-ch ${t.ch}">${t.ch}</span>
        <span class="tc-team">${t.tm}</span>
        <span class="tc-mgr">${t.mg||''}</span>
      </div>
      <div class="progress-wrap">
        <div class="progress-bar"><div class="progress-fill" style="width:${t.pp*100}%;background:${pc(t.pp)}"></div></div>
        <div class="progress-pct" style="color:${pc(t.pp)}">${pct(t.pp)}</div>
      </div>
    </div>`).join('');
}
function showTask(no){
  const t=D.tasks.find(x=>x.no===no);if(!t)return;
  document.getElementById('modalC').innerHTML=`
    <div class="mh"><div class="mt_">${t.co?'<span style="color:var(--r)">● </span>':''}${t.nm}</div>
      <button class="mx" onclick="document.getElementById('modal').classList.remove('show')">✕</button></div>
    <div class="mm">
      <div class="mmi"><span class="mml">총괄</span><span class="mmv">${t.ch}</span></div>
      <div class="mmi"><span class="mml">팀</span><span class="mmv">${t.tm}</span></div>
      <div class="mmi"><span class="mml">담당자</span><span class="mmv">${t.mg||''}</span></div>
      <div class="mmi"><span class="mml">일정</span><span class="mmv">${t.sc||''}</span></div>
      <div class="mmi"><span class="mml">상태</span><span class="mmv"><span class="pill ${t.st}">${t.st}</span></span></div>
      <div class="mmi"><span class="mml">추진도</span><span class="mmv" style="color:${pc(t.pp)}">${pct(t.pp)}</span></div>
    </div>
    <div style="height:8px;background:var(--border);border-radius:4px;overflow:hidden;margin-bottom:14px">
      <div style="height:100%;width:${t.pp*100}%;background:${pc(t.pp)};border-radius:4px"></div></div>
    ${t.plan?`<div class="ms"><div class="mst">세부실행계획</div><div class="msc">${t.plan}</div></div>`:''}
    ${t.pg?`<div class="ms"><div class="mst">추진 내용</div><div class="msc">${t.pg}</div></div>`:''}`;
  document.getElementById('modal').classList.add('show');
}

// ==================== PROFIT ====================
function initProfitUI(){
  const P=D.profit,SD=D.sga_detail;
  // P&L 폭포수: 공헌이익 계산
  const contrib=P.contribution||{total:(P.gross.total-P.sga.total)};
  const chCommon=contrib.total-P.op.total; // 채널+전사공통비+조정 합계
  document.getElementById('profStats').innerHTML=`
    <div class="sc bl"><div class="sl">매출</div><div class="sv">${fB(P.revenue.total)}억</div><div class="ss">총이익률 ${(P.gross.total/P.revenue.total*100).toFixed(1)}%</div></div>
    <div class="sc gr"><div class="sl">매출총이익</div><div class="sv" style="color:var(--g)">${fB(P.gross.total)}억</div></div>
    <div class="sc yl"><div class="sl">판관비</div><div class="sv" style="color:var(--y)">${fB(P.sga.total)}억</div><div class="ss">${(P.sga.total/P.revenue.total*100).toFixed(1)}%</div></div>
    <div class="sc" style="background:rgba(15,23,42,.65);border-left:3px solid #86efac"><div class="sl" style="color:#166534">공헌이익 <span style="font-size:9px;background:#dcfce7;color:#166534;padding:1px 4px;border-radius:8px">영업이익 기준</span></div><div class="sv" style="color:#16a34a;font-size:1.05em">${fB(contrib.total)}억</div><div class="ss" style="color:#166534">${(contrib.total/P.revenue.total*100).toFixed(1)}%</div></div>
    <div class="sc" style="background:#fff7ed;border-left:3px solid #fdba74"><div class="sl" style="color:#9a3412;font-size:11px">조정 차이</div><div class="sv" style="color:#ea580c;font-size:1em">${chCommon>=0?'+':''}${fB(chCommon)}억</div><div class="ss" style="color:#9a3412">점프업·도소매 조정분</div></div>
    <div class="sc og" style="border:2px solid #f97316"><div class="sl" style="font-weight:700">★ 영업이익 <span style="font-size:9px;background:#fed7aa;color:#9a3412;padding:1px 4px;border-radius:8px">조정 후</span></div><div class="sv" style="color:var(--o);font-size:1.2em">${fB(P.op.total)}억</div><div class="ss">${(P.op.total/P.revenue.total*100).toFixed(1)}% <span style="font-size:9px;background:#fed7aa;color:#9a3412;padding:1px 4px;border-radius:3px">${P.op_source||'점프업 도/소매 조정 후'}</span></div></div>`;
  const mx=Math.max(...CHS.map(c=>P.revenue[c]));
  document.getElementById('revChart').innerHTML=CHS.map(c=>`<div class="br"><div class="bl_">${CN[c]}</div><div class="bt"><div class="bf" style="width:${(P.revenue[c]/mx*100).toFixed(1)}%;background:${CC[c]}"></div></div><div class="bv"><b>${fB(P.revenue[c])}</b>억 <span style="color:#94a3b8;font-size:9px">${(P.revenue[c]/P.revenue.total*100).toFixed(1)}%</span></div></div>`).join('');
  // 서비스 매출 (수수료수입)
  const PS=P.rev_service||{};const mxS=Math.max(...CHS.map(c=>PS[c]||0))||1;
  document.getElementById('svcChart').innerHTML=CHS.map(c=>{const v=PS[c]||0;const svcPct=P.revenue[c]>0?(v/P.revenue[c]*100).toFixed(1):0;return `<div class="br"><div class="bl_">${CN[c]}</div><div class="bt"><div class="bf" style="width:${(v/mxS*100).toFixed(1)}%;background:#0369a1"></div></div><div class="bv"><b>${fB(v)}</b>억 <span style="color:#0369a1;font-size:9px">매출의${svcPct}%</span></div></div>`;}).join('');
  // 상품 매출 (단말기)
  const PP=P.rev_product||{};const mxP=Math.max(...CHS.map(c=>PP[c]||0))||1;
  document.getElementById('prodChart').innerHTML=CHS.map(c=>{const v=PP[c]||0;const prodPct=P.revenue[c]>0?(v/P.revenue[c]*100).toFixed(1):0;return `<div class="br"><div class="bl_">${CN[c]}</div><div class="bt"><div class="bf" style="width:${(v/mxP*100).toFixed(1)}%;background:#166534"></div></div><div class="bv"><b>${fB(v)}</b>억 <span style="color:#166534;font-size:9px">매출의${prodPct}%</span></div></div>`;}).join('');
  const contrib2=P.contribution||{total:P.gross.total-P.sga.total,...CHS.reduce((o,c)=>{o[c]=(P.gross[c]||0)-(P.sga[c]||0);return o},{})};
  const mxContrib=Math.max(...CHS.map(c=>Math.abs(contrib2[c]||0)),1);
  const contribTotal=contrib2.total||CHS.reduce((s,c)=>s+(contrib2[c]||0),0)||1;
  const contribEl=document.getElementById('contribChart');
  if(contribEl){
    contribEl.innerHTML=CHS.map(c=>`<div class="br"><div class="bl_">${CN[c]}</div><div class="bt"><div class="bf" style="width:${(Math.abs(contrib2[c]||0)/mxContrib*100).toFixed(1)}%;background:${CC[c]}"></div></div><div class="bv"><b>${fB(contrib2[c]||0)}</b>억 <span style="color:#16a34a;font-size:9px">${(((contrib2[c]||0)/contribTotal)*100).toFixed(1)}%</span></div></div>`).join('');
  }
  const mx2=Math.max(...CHS.map(c=>Math.abs(P.op[c])),1);
  document.getElementById('opChart').innerHTML=CHS.map(c=>`<div class="br"><div class="bl_">${CN[c]}</div><div class="bt"><div class="bf" style="width:${(Math.abs(P.op[c])/mx2*100).toFixed(1)}%;background:${CC[c]}"></div></div><div class="bv"><b>${fB(P.op[c])}</b>억 <span style="color:#ea580c;font-size:9px">${(P.op[c]/P.op.total*100).toFixed(1)}%</span></div></div>`).join('');
  const PS2=P.rev_service||{};const PP2=P.rev_product||{};
  const rows=[
    {l:'매출 (총)',d:P.revenue,b:1},
    {l:'┣ 서비스매출',d:PS2,b:0,style:'color:#0369a1;padding-left:12px'},
    {l:'┗ 상품매출',d:PP2,b:0,style:'color:#166534;padding-left:12px'},
    {l:'매출총이익',d:P.gross,b:0},
    {l:'판관비',d:P.sga,b:0},
    {l:'공헌이익',d:contrib2,b:0,style:'color:#16a34a'},
    {l:'★영업이익',d:P.op,b:1,style:'color:#ea580c;font-weight:700'}
  ];
  document.getElementById('fBody').innerHTML=rows.map(r=>`<tr class="${r.b?'tr':''}"><td style="font-weight:${r.b?700:500};white-space:nowrap;${r.style||''}">${r.l}</td><td style="${r.style||''}">${fMr(r.d.total)}</td>${CHS.map(c=>`<td style="${r.style||''}">${fMr(r.d[c]||0)}</td>`).join('')}</tr>`).join('')+
    `<tr><td>이익률</td><td style="font-weight:700">${(P.op.total/P.revenue.total*100).toFixed(1)}%</td>${CHS.map(c=>`<td class="${P.op[c]/P.revenue[c]>.05?'pos':'neg'}">${(P.op[c]/P.revenue[c]*100).toFixed(1)}%</td>`).join('')}</tr>`;
  document.getElementById('hqBody').innerHTML=D.hq.map(h=>{const op=h.gp-h.sga;return `<tr><td style="font-weight:600">${h.nm}</td><td>${fMr(h.rev)}</td><td>${fMr(h.gp)}</td><td>${fMr(h.sga)}</td><td class="${op>0?'pos':'neg'}">${fMr(op)}</td><td class="${op/h.rev>.05?'pos':'neg'}">${(op/h.rev*100).toFixed(1)}%</td></tr>`;}).join('');
  const sgaTop=[{l:'인건비',v:SD['인건비합계'].total,c:'#38bdf8'},{l:'판촉비',v:SD['판매촉진비'].total,c:'#f87171'},{l:'지급수수료',v:SD['지급수수료'].total,c:'#a78bfa'},{l:'임차료',v:SD['임차료'].total,c:'#34d399'},{l:'판매수수료',v:SD['판매수수료'].total,c:'#fb923c'},{l:'복리후생비',v:SD['복리후생비'].total,c:'#fbbf24'},{l:'운반비',v:SD['운반비'].total,c:'#818cf8'}].sort((a,b)=>b.v-a.v);
  const sgaMx=sgaTop[0].v;
  document.getElementById('sgaChart').innerHTML=sgaTop.map(x=>`<div class="br"><div class="bl_">${x.l}</div><div class="bt"><div class="bf" style="width:${(x.v/sgaMx*100).toFixed(0)}%;background:${x.c}">${fB(x.v)}억</div></div><div class="bv">${(x.v/P.sga.total*100).toFixed(1)}%</div></div>`).join('');
  document.getElementById('costEffChart').innerHTML=CHS.map(c=>{
    const rev=P.revenue[c],labor=SD['인건비합계'][c],promo=SD['판매촉진비'][c];
    const lr=(labor/rev*100).toFixed(1),pr=(promo/rev*100).toFixed(1);
    return `<div style="margin-bottom:10px"><div style="font-size:11px;font-weight:600;color:var(--t2);margin-bottom:3px">${CN[c]}</div>
      <div class="br" style="margin-bottom:2px"><div class="bl_" style="width:56px;font-size:10px">인건비</div><div class="bt" style="height:16px"><div class="bf" style="width:${Math.min(100,lr*2)}%;background:#38bdf8;font-size:9px">${lr}%</div></div></div>
      <div class="br" style="margin-bottom:0"><div class="bl_" style="width:56px;font-size:10px">판촉비</div><div class="bt" style="height:16px"><div class="bf" style="width:${Math.min(100,pr*2)}%;background:#f87171;font-size:9px">${pr}%</div></div></div></div>`;
  }).join('');

  const PF=P.platform||{};
  const PFT=PF.total||{};
  const pRev=PFT.revenue||0,pCon=PFT.contribution||0,pSga=PFT.sga||0,pOp=PFT.op||0;
  const pConM=pRev?((pCon/pRev)*100):0;
  const pOpM=pRev?((pOp/pRev)*100):0;
  const platformCards=document.getElementById('platformCards');
  if(platformCards){
    platformCards.innerHTML=`
      <div class="sc" style="background:linear-gradient(135deg,#faf5ff,#ede9fe);border-left:3px solid #8b5cf6"><div class="sl" style="color:#6b21a8">유통플랫폼 매출</div><div class="sv" style="color:#7c3aed">${fB(pRev)}억</div><div class="ss">사업단 합계</div></div>
      <div class="sc" style="background:linear-gradient(135deg,#ecfeff,#cffafe);border-left:3px solid #06b6d4"><div class="sl" style="color:#155e75">공헌이익 <span style="font-size:9px;background:#cffafe;color:#155e75;padding:1px 4px;border-radius:8px">영업이익 기준</span></div><div class="sv" style="color:#0891b2">${fB(pCon)}억</div><div class="ss">${pConM.toFixed(1)}%</div></div>
      <div class="sc" style="background:linear-gradient(135deg,#fff7ed,#fed7aa);border-left:3px solid #f59e0b"><div class="sl" style="color:#9a3412">판관비</div><div class="sv" style="color:#ea580c">${fB(pSga)}억</div><div class="ss">${pRev?((pSga/pRev)*100).toFixed(1):'-'}%</div></div>
      <div class="sc" style="background:linear-gradient(135deg,#f0fdf4,#dcfce7);border-left:3px solid ${pOp>=0?'#22c55e':'#ef4444'}"><div class="sl" style="color:${pOp>=0?'#166534':'#991b1b'};font-weight:700">영업이익 <span style="font-size:9px;background:${pOp>=0?'#dcfce7':'#fee2e2'};color:${pOp>=0?'#166534':'#991b1b'};padding:1px 4px;border-radius:8px">조정 후</span></div><div class="sv" style="color:${pOp>=0?'#16a34a':'#dc2626'}">${fB(pOp)}억</div><div class="ss">${pOpM.toFixed(1)}%</div></div>`;
  }
}

// ==================== KPI ====================
function initKpiUI(){
  const K=D.kpi;
  const hqs=['강북본부','강남본부','강서본부','동부본부','서부본부'];
  document.getElementById('kpiCards').innerHTML=K.map(k=>{
    const rd=D.kpi_retail_detail[k.hq]||{},wd=D.kpi_wholesale_detail[k.hq]||{};
    return `<div class="kc kr${k.rk}"><div class="kc-top"><div class="kn">${k.hq}</div><div class="krb r${k.rk}">#${k.rk}</div></div>
    <div class="ks" style="color:${RC[k.rk]}">${k.ts.toFixed(1)}</div>
    <div class="kbd"><div class="km"><div class="kml">소매60%</div><div class="kmv" style="color:var(--acc)">${k.rt.t.toFixed(1)}</div></div><div class="km"><div class="kml">도매30%</div><div class="kmv" style="color:var(--g)">${k.wh.t.toFixed(1)}</div></div><div class="km"><div class="kml">소상공10%</div><div class="kmv" style="color:var(--p)">${k.sm.t.toFixed(1)}</div></div></div>
    <div class="smg"><div class="smi"><div class="smi-l">후불</div><div class="smi-v">${rd.hubul?.toFixed(1)||'-'}</div></div><div class="smi"><div class="smi-l">MNP</div><div class="smi-v">${rd.mnp?.toFixed(1)||'-'}</div></div><div class="smi"><div class="smi-l">유선</div><div class="smi-v">${rd.wired?.toFixed(1)||'-'}</div></div><div class="smi"><div class="smi-l">R-VOC</div><div class="smi-v">${rd.rvoc?.toFixed(1)||'-'}</div></div><div class="smi"><div class="smi-l">도매유선</div><div class="smi-v">${wd.wired?.toFixed(1)||'-'}</div></div><div class="smi"><div class="smi-l">인프라</div><div class="smi-v">${wd.infra?.toFixed(1)||'-'}</div></div></div></div>`;
  }).join('');
  document.getElementById('rtBody').innerHTML=hqs.map(h=>{const d=D.kpi_retail_detail[h]||{};return `<tr><td style="font-weight:600">${h.replace('본부','')}</td><td style="font-weight:700">${d.total?.toFixed(1)}</td><td style="color:${d.hubul>14?'var(--g)':'var(--t2)'}">${d.hubul?.toFixed(1)}</td><td style="color:${d.mnp>7?'var(--g)':'var(--r)'}">${d.mnp?.toFixed(1)}</td><td>${d.wired?.toFixed(1)}</td><td>${d.mit?.toFixed(1)}</td><td style="color:${d.rvoc>3?'var(--g)':'var(--r)'}">${d.rvoc?.toFixed(2)}</td><td style="color:${d.tcsi>4?'var(--g)':'var(--r)'}">${d.tcsi?.toFixed(1)}</td><td>${d.store_profit?.toFixed(1)}</td></tr>`;}).join('');
  // ── 채널별 KPI 카드 렌더링 ──
  const rtTotals=hqs.map(h=>D.kpi_retail_detail[h]?.total||0);
  const rtMax=Math.max(...rtTotals), rtMin=Math.min(...rtTotals);
  document.getElementById('kpiCards-rt').innerHTML=hqs.map(h=>{
    const d=D.kpi_retail_detail[h]||{};
    const tot=d.total||0;
    const rank=rtTotals.slice().sort((a,b)=>b-a).indexOf(tot)+1;
    const rankC=tot===rtMax?'#059669':tot===rtMin?'#dc2626':'#3b82f6';
    const barW=(tot/rtMax*100).toFixed(0);
    const known_sum=(d.hubul||0)+(d.mnp||0)+(d.wired||0)+(d.mit||0)+(d.apru||0)+(d.strategy||0)+(d.rvoc||0)+(d.oneg||0)+(d.tcsi||0)+(d.policy_rev||0)+(d.store_profit||0);
    const etc_rt=Math.max(0,tot-known_sum);
    const items=[
      {l:'후불(15)',    v:(d.hubul||0).toFixed(2),  c:d.hubul>=14?'#059669':d.hubul<12?'#dc2626':'#64748b'},
      {l:'MNP(10)',     v:(d.mnp||0).toFixed(2),    c:d.mnp>=7?'#059669':d.mnp<5?'#dc2626':'#64748b'},
      {l:'유선(15)',    v:(d.wired||0).toFixed(2),  c:d.wired>=14?'#059669':d.wired<11?'#dc2626':'#64748b'},
      {l:'MIT(10)',     v:(d.mit||0).toFixed(2),    c:'#64748b'},
      {l:'ARPU(7)',     v:(d.apru||0).toFixed(2),   c:'#64748b'},
      {l:'전략상품(5)', v:(d.strategy||0).toFixed(2),c:d.strategy>=3?'#059669':d.strategy<1?'#dc2626':'#64748b'},
      {l:'R-VOC(5)',    v:(d.rvoc||0).toFixed(2),   c:d.rvoc>=4?'#059669':d.rvoc<2?'#dc2626':'#64748b'},
      {l:'1G유치(5)',   v:(d.oneg||0).toFixed(1),   c:d.oneg>=5?'#059669':'#64748b'},
      {l:'TCSI(3)',     v:(d.tcsi||0).toFixed(2),   c:d.tcsi>=3?'#059669':d.tcsi<2?'#dc2626':'#64748b'},
      {l:'정책매출(10)',v:(d.policy_rev||0).toFixed(2),c:'#64748b'},
      {l:'매장이익(5)', v:(d.store_profit||0).toFixed(1),c:d.store_profit>=6?'#059669':'#64748b'},
      ...(etc_rt>0.05?[{l:'기타',v:etc_rt.toFixed(2),c:'#94a3b8'}]:[]),
    ];
    return `<div style="background:var(--bg2);border-radius:10px;padding:11px;margin-bottom:8px">
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px">
        <span style="font-size:14px;font-weight:800;color:var(--t1)">${h.replace('본부','')}</span>
        <span style="background:${rankC};color:#fff;font-size:10px;font-weight:800;padding:2px 8px;border-radius:10px">#${rank}</span>
        <div style="margin-left:auto;text-align:right">
          <div style="font-size:18px;font-weight:900;color:${rankC}">${tot.toFixed(2)}</div>
          <div style="font-size:9px;color:#94a3b8">합 ${known_sum.toFixed(2)}</div>
        </div>
      </div>
      <div style="height:5px;border-radius:3px;background:#e2e8f0;margin-bottom:9px;overflow:hidden">
        <div style="height:100%;border-radius:3px;background:${rankC};width:${barW}%"></div>
      </div>
      <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:5px">
        ${items.map(it=>`<div style="background:var(--bg1);border-radius:6px;padding:5px 4px;text-align:center">
          <div style="font-size:9px;color:var(--t3);margin-bottom:2px">${it.l}</div>
          <div style="font-size:12px;font-weight:700;color:${it.c}">${it.v||'-'}</div>
        </div>`).join('')}
      </div>
    </div>`;
  }).join('');
  document.getElementById('whBody').innerHTML=hqs.map(h=>{const d=D.kpi_wholesale_detail[h]||{};return `<tr><td style="font-weight:600">${h.replace('본부','')}</td><td style="font-weight:700">${d.total?.toFixed(2)}</td><td>${d.hubul?.toFixed(1)}</td><td style="color:${d.mnp>=12?'var(--g)':'var(--r)'}">${d.mnp?.toFixed(1)}</td><td style="color:${d.wired>=18?'var(--g)':d.wired<15?'var(--r)':'var(--t2)'}">${d.wired?.toFixed(2)}</td><td>${d.rvoc?.toFixed(2)}</td><td>${d.apru?.toFixed(2)}</td><td>${d.infra?.toFixed(2)}</td><td>${d.add_rev?.toFixed(2)}</td></tr>`;}).join('');
  const whTotals=hqs.map(h=>D.kpi_wholesale_detail[h]?.total||0);
  const whMax=Math.max(...whTotals),whMin=Math.min(...whTotals);
  document.getElementById('kpiCards-wh').innerHTML=hqs.map(h=>{
    const d=D.kpi_wholesale_detail[h]||{};
    const tot=d.total||0;
    const rank=whTotals.slice().sort((a,b)=>b-a).indexOf(tot)+1;
    const rankC=tot===whMax?'#059669':tot===whMin?'#dc2626':'#3b82f6';
    const barW=(tot/whMax*100).toFixed(0);
    const wh_known=(d.hubul||0)+(d.mnp||0)+(d.wired||0)+(d.mit||0)+(d.rvoc||0)+(d.oneg||0)+(d.apru||0)+(d.infra||0)+(d.add_rev||0);
    const etc_wh=Math.max(0,tot-wh_known);
    const items=[
      {l:'후불(20)',    v:(d.hubul||0).toFixed(2),  c:d.hubul>=20?'#059669':d.hubul<18?'#dc2626':'#64748b'},
      {l:'MNP(10)',     v:(d.mnp||0).toFixed(2),    c:d.mnp>=12?'#059669':d.mnp<10?'#dc2626':'#64748b'},
      {l:'유선(20)',    v:(d.wired||0).toFixed(2),  c:d.wired>=18?'#059669':d.wired<15?'#dc2626':'#64748b'},
      {l:'MIT(10)',     v:(d.mit||0).toFixed(2),    c:d.mit>=12?'#059669':d.mit<10?'#dc2626':'#64748b'},
      {l:'R-VOC(5)',    v:(d.rvoc||0).toFixed(2),   c:d.rvoc>=4?'#059669':d.rvoc<3?'#dc2626':'#64748b'},
      {l:'1G유치(5)',   v:(d.oneg||0).toFixed(1),   c:d.oneg>=5?'#059669':'#64748b'},
      {l:'판매ARPU(7)', v:(d.apru||0).toFixed(2),   c:'#64748b'},
      {l:'인프라(8)',   v:(d.infra||0).toFixed(2),  c:d.infra>=8?'#059669':d.infra<7?'#dc2626':'#64748b'},
      {l:'추가수익(5)', v:(d.add_rev||0).toFixed(2),c:d.add_rev>=5?'#059669':'#64748b'},
      ...(etc_wh>0.05?[{l:'기타',v:etc_wh.toFixed(2),c:'#94a3b8'}]:[]),
    ];
    return `<div style="background:var(--bg2);border-radius:10px;padding:11px;margin-bottom:8px">
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px">
        <span style="font-size:14px;font-weight:800;color:var(--t1)">${h.replace('본부','')}</span>
        <span style="background:${rankC};color:#fff;font-size:10px;font-weight:800;padding:2px 8px;border-radius:10px">#${rank}</span>
        <span style="font-size:18px;font-weight:900;color:${rankC};margin-left:auto">${tot.toFixed(1)}</span>
      </div>
      <div style="height:5px;border-radius:3px;background:#e2e8f0;margin-bottom:9px;overflow:hidden">
        <div style="height:100%;border-radius:3px;background:${rankC};width:${barW}%"></div>
      </div>
      <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:5px">
        ${items.map(it=>`<div style="background:var(--bg1);border-radius:6px;padding:5px 4px;text-align:center">
          <div style="font-size:9px;color:var(--t3);margin-bottom:2px">${it.l}</div>
          <div style="font-size:12px;font-weight:700;color:${it.c}">${it.v||'-'}</div>
        </div>`).join('')}
      </div>
    </div>`;
  }).join('');
  // 소상공인 KPI 세부지표
  const smbBodyEl=document.getElementById('smbBody');
  if(smbBodyEl){
    const SMD2=D.kpi_smb_detail||{};
    const smbVals=hqs.map(h=>SMD2[h]?.total||0);
    const smbMax=Math.max(...smbVals)||1,smbMin=Math.min(...smbVals)||0;
    smbBodyEl.innerHTML=hqs.map(h=>{
      const d=SMD2[h]||{};const tot=d.total||0;
      const tc=tot===smbMax?'var(--g)':tot===smbMin?'var(--r)':'var(--t2)';
      return `<tr>
        <td style="font-weight:600">${h.replace('본부','')}</td>
        <td style="font-weight:800;color:${tc}">${tot.toFixed(2)}</td>
        <td>${(d.hubul||0).toFixed(2)}</td>
        <td>${(d.wired||0).toFixed(2)}</td>
        <td>${(d.mit||0).toFixed(2)}</td>
        <td style="color:${(d.strategy||0)>=15?'var(--g)':(d.strategy||0)<10?'var(--r)':'var(--t2)'}">${(d.strategy||0).toFixed(2)}</td>
        <td>${(d.rvoc||0).toFixed(2)}</td>
        <td>${(d.op_profit||0).toFixed(2)}</td>
      </tr>`;
    }).join('');
  }
  // SMB 카드
  const SMD3=D.kpi_smb_detail||{};
  const smbVals2=hqs.map(h=>SMD3[h]?.total||0);
  const smbMax2=Math.max(...smbVals2),smbMin2=Math.min(...smbVals2);
  const smbEl2=document.getElementById('kpiCards-smb');
  if(smbEl2){smbEl2.innerHTML=hqs.map(h=>{
    const d=SMD3[h]||{};const tot=d.total||0;
    const rank=smbVals2.slice().sort((a,b)=>b-a).indexOf(tot)+1;
    const rankC=tot===smbMax2?'#059669':tot===smbMin2?'#dc2626':'#3b82f6';
    const barW=(tot/smbMax2*100).toFixed(0);
    const known_sum=(d.hubul||0)+(d.wired||0)+(d.mit||0)+(d.strategy||0)+(d.rvoc||0)+(d.op_profit||0);
    const etc_smb=Math.max(0,tot-known_sum);
    const items=[
      {l:'후불(20)',   v:(d.hubul||0).toFixed(2),   c:d.hubul>=15?'#059669':d.hubul<8?'#dc2626':'#64748b'},
      {l:'유선(20)',   v:(d.wired||0).toFixed(2),   c:d.wired>=12?'#059669':d.wired<8?'#dc2626':'#64748b'},
      {l:'MIT(7)',     v:(d.mit||0).toFixed(2),     c:d.mit>=5?'#059669':d.mit<2?'#dc2626':'#64748b'},
      {l:'전략상품(20)',v:(d.strategy||0).toFixed(2),c:d.strategy>=15?'#059669':d.strategy<10?'#dc2626':'#64748b'},
      {l:'R-VOC(3)',   v:(d.rvoc||0).toFixed(2),   c:d.rvoc>=3?'#059669':d.rvoc<2?'#dc2626':'#64748b'},
      {l:'영업이익(20)',v:(d.op_profit||0).toFixed(2),c:d.op_profit>=20?'#059669':d.op_profit<18?'#dc2626':'#64748b'},
      ...(etc_smb>0.05?[{l:'기타',v:etc_smb.toFixed(2),c:'#94a3b8'}]:[]),
    ];
    return `<div style="background:var(--bg2);border-radius:10px;padding:11px;margin-bottom:8px">
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:8px">
        <span style="font-size:14px;font-weight:800;color:var(--t1)">${h.replace('본부','')}</span>
        <span style="background:${rankC};color:#fff;font-size:10px;font-weight:800;padding:2px 8px;border-radius:10px">#${rank}</span>
        <div style="margin-left:auto;text-align:right">
          <div style="font-size:18px;font-weight:900;color:${rankC}">${tot.toFixed(2)}</div>
          <div style="font-size:9px;color:#94a3b8">합 ${known_sum.toFixed(2)}</div>
        </div>
      </div>
      <div style="height:5px;border-radius:3px;background:#e2e8f0;margin-bottom:9px;overflow:hidden">
        <div style="height:100%;border-radius:3px;background:${rankC};width:${barW}%"></div>
      </div>
      <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:5px">
        ${items.map(it=>`<div style="background:var(--bg1);border-radius:6px;padding:5px 4px;text-align:center">
          <div style="font-size:9px;color:var(--t3);margin-bottom:2px">${it.l}</div>
          <div style="font-size:12px;font-weight:700;color:${it.c}">${it.v}</div>
        </div>`).join('')}
      </div>
    </div>`;
  }).join('');}
  const rtV=hqs.map(h=>D.kpi_retail_detail[h]?.total||0),whV=hqs.map(h=>D.kpi_wholesale_detail[h]?.total||0);
  document.getElementById('hmBody').innerHTML=K.map(k=>{const chs={소매:k.rt.t,도매:k.wh.t,소상공인:k.sm.t};const weak=Object.entries(chs).sort((a,b)=>a[1]-b[1])[0][0];return `<tr><td style="font-weight:600">${k.hq.replace('본부','')}</td><td style="font-weight:700;color:${RC[k.rk]}">${k.ts.toFixed(1)}</td><td><span class="hm ${hmc(k.rt.t,rtV)}">${k.rt.t.toFixed(1)}</span></td><td><span class="hm ${hmc(k.wh.t,whV)}">${k.wh.t.toFixed(1)}</span></td><td><span class="hm ${hmc(k.sm.t,[71.95,59.23,66.63,59.64,49.01])}">${k.sm.t.toFixed(1)}</span></td><td style="color:var(--r);font-size:11px">${weak}</td></tr>`;}).join('');
  document.getElementById('kRank').innerHTML=K.map(k=>`<div class="br"><div class="bl_">${k.hq.replace('본부','')}</div><div class="bt"><div class="bf" style="width:${k.ts}%;background:${RC[k.rk]}">${k.ts.toFixed(1)}</div></div><div class="bv" style="color:${RC[k.rk]};font-weight:700">#${k.rk}</div></div>`).join('');
  const V=D.variance;
  const vars=[{l:'소상공인영업',v:V.smb_sales},{l:'소상공인종합',v:V.smb},{l:'도매영업',v:V.wh_sales},{l:'도매종합',v:V.wholesale},{l:'종합KPI',v:V.total},{l:'소매영업',v:V.retail_sales},{l:'소매종합',v:V.retail},{l:'도매이익',v:V.wh_profit},{l:'소매이익',v:V.retail_profit}];
  const vmx=Math.max(...vars.map(x=>x.v));
  document.getElementById('kVar').innerHTML=vars.map(x=>`<div class="br"><div class="bl_">${x.l}</div><div class="bt"><div class="bf" style="width:${(x.v/vmx*100).toFixed(1)}%;background:${x.v>10?'var(--r)':x.v>5?'var(--o)':x.v>2?'var(--y)':'var(--g)'}">${x.v.toFixed(1)}p</div></div><div class="bv">${x.v>10?'⚠️':x.v>5?'⚡':'✅'}</div></div>`).join('');
}

// ==================== REPORT SHEET ====================

// ==================== REPORT RENDER ====================
function openReport(html, title){
  title = title || '보고서';
  const viewer  = document.getElementById('rptViewer');
  const body    = document.getElementById('rptViewerBody');
  document.getElementById('rptViewerTitle').textContent = title;
  body.innerHTML = '';
  window._currentReportHtml = html;
  window._currentReportTitle = title;

  const iframe = document.createElement('iframe');
  iframe.id = 'rptIframe';
  iframe.style.cssText = 'width:100%;height:100%;border:none;background:#fff;';

  if('srcdoc' in iframe){
    iframe.srcdoc = html;
  } else {
    try {
      const blob = new Blob([html], {type:'text/html;charset=utf-8'});
      const url  = URL.createObjectURL(blob);
      iframe.onload = ()=>{ try{ URL.revokeObjectURL(url); }catch(e){} };
      iframe.src = url;
    } catch(e) {
      const doc = iframe.contentDocument || iframe.contentWindow.document;
      doc.open(); doc.write(html); doc.close();
    }
  }

  body.appendChild(iframe);
  viewer.classList.add('show');
  document.body.style.overflow = 'hidden';
}

function closeReportViewer(){
  const viewer = document.getElementById('rptViewer');
  const body   = document.getElementById('rptViewerBody');
  viewer.classList.remove('show');
  document.body.style.overflow = '';
  setTimeout(()=>{ body.innerHTML=''; }, 300);
}

function downloadReport(){
  const title = window._currentReportTitle || '보고서';
  const html  = window._currentReportHtml;
  if(!html) return;

  const editUi = `<div id="_edit_banner" style="position:fixed;top:0;left:0;right:0;z-index:9999;background:#fef9c3;border-bottom:2px solid #fbbf24;padding:8px 16px;font-size:12px;font-weight:600;color:#92400e;display:flex;align-items:center;justify-content:space-between;gap:8px;flex-wrap:wrap;font-family:sans-serif">
    <span>✏️ 편집 가능 모드 — 글자 크기/색상/정렬을 바로 수정할 수 있습니다.</span>
    <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap">
      <label style="font-size:11px;font-weight:700;color:#92400e">크기
        <select id="_edit_font_size" style="margin-left:4px;padding:3px 6px;border:1px solid #f59e0b;border-radius:6px;font-size:11px">
          <option value="12px">12</option><option value="14px">14</option><option value="16px" selected>16</option><option value="18px">18</option><option value="20px">20</option><option value="24px">24</option>
        </select>
      </label>
      <label style="font-size:11px;font-weight:700;color:#92400e">색상 <input id="_edit_font_color" type="color" value="#111827" style="width:28px;height:24px;border:none;background:transparent;vertical-align:middle"></label>
      <button id="_edit_apply_style" style="background:#0369a1;color:#fff;border:none;border-radius:6px;padding:4px 10px;cursor:pointer;font-size:11px;font-weight:700">글자 적용</button>
      <button id="_edit_align_left" style="background:#334155;color:#fff;border:none;border-radius:6px;padding:4px 10px;cursor:pointer;font-size:11px;font-weight:700">좌측</button>
      <button id="_edit_align_center" style="background:#334155;color:#fff;border:none;border-radius:6px;padding:4px 10px;cursor:pointer;font-size:11px;font-weight:700">가운데</button>
      <button id="_edit_align_right" style="background:#334155;color:#fff;border:none;border-radius:6px;padding:4px 10px;cursor:pointer;font-size:11px;font-weight:700">우측</button>
      <button id="_edit_done" style="background:#0f766e;color:#fff;border:none;border-radius:6px;padding:4px 12px;cursor:pointer;font-size:11px;font-weight:700">편집 완료</button>
    </div>
  </div>
  <div id="_edit_toolbar_space" style="height:76px"></div>
  <script>(function(){
    function getRange(){
      var sel = window.getSelection && window.getSelection();
      if(!sel || !sel.rangeCount) return null;
      return sel.getRangeAt(0);
    }
    function wrapSelection(styleObj){
      var range = getRange();
      if(!range || range.collapsed) return;
      var span = document.createElement('span');
      Object.keys(styleObj||{}).forEach(function(k){ span.style[k]=styleObj[k]; });
      try{ range.surroundContents(span); }
      catch(e){
        var frag = range.extractContents();
        span.appendChild(frag);
        range.insertNode(span);
      }
    }
    function alignSelection(align){
      var range = getRange();
      if(!range) return;
      var node = range.commonAncestorContainer;
      var el = node.nodeType===1 ? node : node.parentElement;
      while(el && el!==document.body && !/^P|DIV|TD|TH|LI|H[1-6]$/.test(el.tagName||'')){ el = el.parentElement; }
      if(!el || el===document.body){
        var box = document.createElement('div');
        box.style.textAlign = align;
        try{ range.surroundContents(box); }
        catch(e){
          var frag = range.extractContents();
          box.appendChild(frag);
          range.insertNode(box);
        }
        return;
      }
      el.style.textAlign = align;
    }
    function bind(id, fn){ var el=document.getElementById(id); if(el) el.addEventListener('click', fn); }
    bind('_edit_apply_style', function(){
      var size=(document.getElementById('_edit_font_size')||{}).value||'16px';
      var color=(document.getElementById('_edit_font_color')||{}).value||'#111827';
      wrapSelection({fontSize:size,color:color});
    });
    bind('_edit_align_left', function(){ alignSelection('left'); });
    bind('_edit_align_center', function(){ alignSelection('center'); });
    bind('_edit_align_right', function(){ alignSelection('right'); });
    bind('_edit_done', function(){
      var b=document.getElementById('_edit_banner'); if(b) b.remove();
      var s=document.getElementById('_edit_toolbar_space'); if(s) s.remove();
      document.body.removeAttribute('contenteditable');
    });
  })();<\/script>`;

  // 다운로드용 HTML: body에 contenteditable 추가 + 편집 툴바 삽입
  const editableHtml = html.replace('<body>', '<body contenteditable="true" spellcheck="false" style="cursor:text">').replace('<body contenteditable="true" spellcheck="false" style="cursor:text">', '<body contenteditable="true" spellcheck="false" style="cursor:text">'+editUi);

  try {
    const blob = new Blob([editableHtml], {type:'text/html;charset=utf-8'});
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = `${title}_${D.period}.html`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(()=> URL.revokeObjectURL(url), 1000);
  } catch(e){
    alert('다운로드 중 오류가 발생했습니다: ' + e.message);
  }
}

function printReport(){
  const html = window._currentReportHtml;
  if(!html) return;
  try {
    const blob = new Blob([html], {type:'text/html;charset=utf-8'});
    const url  = URL.createObjectURL(blob);
    const w = window.open(url, '_blank');
    if(w){ w.onload = ()=>{ w.print(); setTimeout(()=>URL.revokeObjectURL(url), 1000); }; }
    else { URL.revokeObjectURL(url); }
  } catch(e){
    const iframe = document.getElementById('rptIframe');
    if(iframe && iframe.contentWindow) iframe.contentWindow.print();
  }
}


// ==================== EXECUTIVE REPORT STYLE ====================
function rptStyle(){
  return `<style>
  @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;600;700;800;900&display=swap');
  *{margin:0;padding:0;box-sizing:border-box}
  body{font-family:'Noto Sans KR',sans-serif;background:linear-gradient(180deg,#eaf2fb 0%,#f8fafc 32%,#eef2f7 100%);color:#0f172a;padding:28px 0 56px;font-size:12px;line-height:1.65}
  .page{max-width:860px;margin:0 auto;padding:34px 38px 50px;background:#fff;border-radius:0 0 24px 24px;box-shadow:0 18px 50px rgba(15,23,42,.12),0 4px 14px rgba(15,23,42,.06);border:1px solid #dbe4ee;border-top:none;position:relative}
  .page::before{content:'';position:absolute;top:0;left:38px;right:38px;height:1px;background:linear-gradient(90deg,rgba(3,105,161,0),rgba(3,105,161,.35),rgba(3,105,161,0))}

  .cover{max-width:860px;margin:0 auto;text-align:center;background:linear-gradient(135deg,#0f4c81 0%,#075985 45%,#082f49 100%);color:#fff;padding:42px 34px 30px;border-radius:24px 24px 0 0;box-shadow:0 18px 50px rgba(15,23,42,.14),0 6px 16px rgba(15,23,42,.08);position:relative;overflow:hidden;border:1px solid rgba(255,255,255,.14);border-bottom:none}
  .cover::before{content:'';position:absolute;inset:auto -10% 42% auto;width:320px;height:320px;background:radial-gradient(circle,rgba(255,255,255,.18) 0%,rgba(255,255,255,.07) 35%,rgba(255,255,255,0) 70%);pointer-events:none}
  .cover::after{content:'';position:absolute;left:-12%;top:-28px;width:260px;height:260px;background:radial-gradient(circle,rgba(125,211,252,.28) 0%,rgba(125,211,252,.08) 38%,rgba(125,211,252,0) 72%);pointer-events:none}
  .cover-tag{display:inline-flex;align-items:center;justify-content:center;padding:5px 14px;background:rgba(255,255,255,.16);border:1px solid rgba(255,255,255,.28);border-radius:999px;font-size:10px;font-weight:800;letter-spacing:1.2px;text-transform:uppercase;margin:0 auto 16px;backdrop-filter:blur(6px)}
  .cover-title{font-size:30px;font-weight:900;letter-spacing:-.8px;margin-bottom:10px;line-height:1.18;text-shadow:0 2px 10px rgba(0,0,0,.12)}
  .cover-sub{max-width:620px;margin:0 auto 24px;font-size:13px;color:rgba(255,255,255,.82);line-height:1.6}
  .cover-meta{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;border-top:1px solid rgba(255,255,255,.18);padding-top:20px;max-width:680px;margin:0 auto}
  .cover-kpi{text-align:center;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.12);border-radius:14px;padding:14px 10px;backdrop-filter:blur(4px)}
  .cover-kpi-v{font-size:28px;font-weight:900;letter-spacing:-1px}
  .cover-kpi-l{font-size:10px;color:rgba(255,255,255,.72);font-weight:600;margin-top:3px}

  h1{font-size:18px;font-weight:800;color:#0369a1;border-bottom:2px solid #0369a1;padding-bottom:8px;margin:28px 0 14px;letter-spacing:-.3px}
  h1:first-of-type{margin-top:0}
  h2{font-size:14px;font-weight:700;color:#0f172a;margin:20px 0 10px;padding-left:10px;border-left:3px solid #0369a1}
  h3{font-size:12px;font-weight:700;color:#334155;margin:14px 0 7px;padding-left:8px;border-left:2px solid #e2e8f0}

  .exec-box{background:linear-gradient(135deg,#f0f9ff,#e0f2fe);border:1px solid #bae6fd;border-radius:10px;padding:16px 18px;margin:12px 0;position:relative}
  .exec-box::before{content:'EXECUTIVE SUMMARY';position:absolute;top:-9px;left:14px;background:#0369a1;color:#fff;font-size:8px;font-weight:700;letter-spacing:1.5px;padding:2px 8px;border-radius:4px}
  .insight{background:#eff6ff;border-left:4px solid #3b82f6;padding:11px 14px;margin:10px 0;border-radius:0 8px 8px 0;font-size:11.5px;color:#1e3a5f}
  .warn{background:#fef2f2;border-left:4px solid #dc2626;padding:11px 14px;margin:10px 0;border-radius:0 8px 8px 0;font-size:11.5px;color:#7f1d1d}
  .action{background:#f0fdf4;border-left:4px solid #16a34a;padding:11px 14px;margin:10px 0;border-radius:0 8px 8px 0;font-size:11.5px;color:#14532d}
  .caution{background:#fffbeb;border-left:4px solid #d97706;padding:11px 14px;margin:10px 0;border-radius:0 8px 8px 0;font-size:11.5px;color:#78350f}

  table{width:100%;border-collapse:collapse;margin:10px 0 18px;font-size:11.5px}
  thead tr{background:#0369a1;color:#fff}
  th{padding:9px 10px;text-align:center;font-weight:700;font-size:10.5px;letter-spacing:.3px;white-space:nowrap}
  th:first-child{text-align:left}
  td{padding:8px 10px;text-align:right;border-bottom:1px solid #f1f5f9;font-family:'JetBrains Mono',monospace,sans-serif;font-size:11px}
  td:first-child{text-align:left;font-family:'Noto Sans KR',sans-serif;font-weight:600;font-size:11.5px}
  tr:nth-child(even){background:#f8fafc}
  tr:hover{background:#f0f9ff}
  tr.total{background:#dbeafe;font-weight:700}
  tr.total td{font-size:12px}
  .pos{color:#059669;font-weight:600}.neg{color:#dc2626;font-weight:600}
  .hi{color:#0369a1;font-weight:700}.lo{color:#dc2626}

  .kpi-grid{margin:10px 0 20px}
  .stat-row{display:grid;gap:10px;margin-bottom:6px}
  .stat-row-3{grid-template-columns:repeat(3,1fr)}
  .stat-row-2{grid-template-columns:repeat(2,1fr)}
  .stat-label{font-size:9.5px;font-weight:700;color:#0369a1;letter-spacing:1px;text-transform:uppercase;margin:14px 0 6px;padding:4px 0;border-bottom:1px solid #e2e8f0}
  .stat-card{border-radius:10px;padding:14px 16px;text-align:center;position:relative;overflow:hidden}
  .stat-card::before{content:'';position:absolute;top:0;left:0;right:0;height:3px}
  .stat-card-blue{background:#f0f9ff;border:1px solid #bae6fd}.stat-card-blue::before{background:#0369a1}
  .stat-card-green{background:#f0fdf4;border:1px solid #bbf7d0}.stat-card-green::before{background:#059669}
  .stat-card-gray{background:#f8fafc;border:1px solid #e2e8f0}.stat-card-gray::before{background:#94a3b8}
  .stat-card-best{background:#f0fdf4;border:1px solid #86efac}.stat-card-best::before{background:#16a34a}
  .stat-card-worst{background:#fff7ed;border:1px solid #fed7aa}.stat-card-worst::before{background:#ea580c}
  .stat-v{font-size:26px;font-weight:900;letter-spacing:-1px;font-family:'JetBrains Mono',monospace;margin:4px 0}
  .stat-v-blue{color:#0369a1}.stat-v-green{color:#059669}.stat-v-gray{color:#475569}
  .stat-v-best{color:#16a34a}.stat-v-worst{color:#ea580c}
  .stat-name{font-size:11px;font-weight:700;color:#334155;margin-bottom:2px}
  .stat-sub{font-size:10px;color:#64748b}
  .stat-badge{display:inline-block;padding:2px 8px;border-radius:20px;font-size:9px;font-weight:700;margin-bottom:4px}
  .badge-best{background:#dcfce7;color:#15803d}.badge-worst{background:#ffedd5;color:#c2410c}

  .bar-row{display:flex;align-items:center;gap:8px;margin:4px 0}
  .bar-label{width:80px;font-size:11px;font-weight:600;text-align:right;flex-shrink:0;color:#475569}
  .bar-bg{flex:1;height:18px;background:#f1f5f9;border-radius:4px;overflow:hidden}
  .bar-fill{height:100%;border-radius:4px;display:flex;align-items:center;padding:0 8px;font-size:10px;font-weight:700;color:#fff}
  .bar-val{width:80px;font-size:11px;font-family:monospace;text-align:left;padding-left:6px;flex-shrink:0;color:#475569}

  .ch-cost-block{margin-bottom:18px;border:1px solid #e2e8f0;border-radius:10px;overflow:hidden}
  .ch-cost-header{padding:10px 14px;font-weight:700;font-size:12px;color:#fff;display:flex;justify-content:space-between;align-items:center}
  .ch-cost-body{padding:10px 14px}
  .ch-cost-sub{font-size:10px;opacity:.8;font-weight:400}

  .badge{display:inline-block;padding:2px 8px;border-radius:12px;font-size:10px;font-weight:700}
  .badge-r{background:#dcfce7;color:#16a34a}.badge-y{background:#fef9c3;color:#ca8a04}.badge-o{background:#ffedd5;color:#c2410c}.badge-red{background:#fee2e2;color:#dc2626}

  .priority-table th{font-size:9.5px}.priority-table td{font-size:10.5px}
  .section-divider{border:none;border-top:1px dashed #cbd5e1;margin:20px 0}
  .footnote{font-size:10px;color:#94a3b8;margin-top:6px;font-style:italic}

  @media (max-width:900px){
    body{padding:18px 10px 40px}
    .cover{border-radius:20px 20px 0 0;padding:32px 18px 24px}
    .cover-title{font-size:24px}
    .cover-sub{font-size:12px}
    .cover-meta{grid-template-columns:1fr;max-width:none}
    .page{padding:24px 18px 34px;border-radius:0 0 20px 20px}
    .page::before{left:18px;right:18px}
    .stat-row-3,.stat-row-2{grid-template-columns:1fr}
  }

  @media print{
    body{background:#fff;padding:0;font-size:11px}
    .cover,.page{box-shadow:none;border:1px solid #dbe4ee}
    .cover{border-radius:0;padding:24px 20px 18px;break-after:avoid-page}
    .cover-title{font-size:22px}
    .cover-sub{max-width:none;font-size:11px;margin-bottom:16px}
    .cover-meta{max-width:none;gap:8px}
    .cover-kpi{padding:10px 8px}
    .cover-kpi-v{font-size:22px}
    .page{border-top:none;border-radius:0;padding:20px}
    .page::before{display:none}
    h1{font-size:15px}
  }
  </style>`;
}

// ==================== ① EXECUTIVE 종합 보고서 ====================
function genProfitExecutive(){
  const P=D.profit,SD=D.sga_detail,H=D.hq;
  const opm=(P.op.total/P.revenue.total*100);
  const bestCh=CHS.reduce((a,b)=>P.op[a]/P.revenue[a]>P.op[b]/P.revenue[b]?a:b);
  const worstCh=CHS.reduce((a,b)=>P.op[a]/P.revenue[a]<P.op[b]/P.revenue[b]?a:b);
  const sortedHq=[...H].sort((a,b)=>(b.gp-b.sga)/b.rev-(a.gp-a.sga)/a.rev);
  const bestHq=sortedHq[0],worstHq=sortedHq[sortedHq.length-1];
  const laborRatio=(SD['인건비합계'].total/P.sga.total*100);
  const promoRatio=(SD['판매촉진비'].total/P.sga.total*100);

  let h=`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>KT M&S Executive 보고서</title>${rptStyle()}</head><body>
  <div class="cover">
    <div class="cover-tag">Confidential · For Executive Review Only</div>
    <div class="cover-title">KT M&S 손익 Executive 보고서</div>
    <div class="cover-sub">${D.period} 기준 | 경영기획팀 | 채널별·본부별·비용구조 종합 분석</div>
    <div class="cover-meta">
      <div class="cover-kpi"><div class="cover-kpi-v">${fB(P.revenue.total)}억</div><div class="cover-kpi-l">전사 매출</div></div>
      <div class="cover-kpi"><div class="cover-kpi-v">${fB(P.op.total)}억</div><div class="cover-kpi-l">영업이익 (${opm.toFixed(1)}%)</div></div>
      <div class="cover-kpi"><div class="cover-kpi-v">${fB(P.net_income)}억</div><div class="cover-kpi-l">당기순이익</div></div>
    </div>
  </div>
  <div class="page">

  <h1>Ⅰ. Executive Summary</h1>
  <div class="exec-box">
    <strong>핵심 성과:</strong> ${D.period} 매출 <strong>${fB(P.revenue.total)}억</strong>원, 영업이익 <strong>${fB(P.op.total)}억</strong>원(영업이익률 ${opm.toFixed(1)}%)으로 사업을 마감하였습니다.<br><br>
    <strong>수익성 분포:</strong> 소상공인 채널(영업이익률 ${(P.op.small_biz/P.revenue.small_biz*100).toFixed(1)}%)이 최고 수익채널이나 매출비중 ${(P.revenue.small_biz/P.revenue.total*100).toFixed(1)}%로 기여도가 낮으며, 디지털 채널(비중 ${(P.revenue.digital/P.revenue.total*100).toFixed(1)}%, 이익률 ${(P.op.digital/P.revenue.digital*100).toFixed(1)}%)은 판촉비 집중에도 불구하고 수익성이 개선 여지가 큽니다.<br><br>
    <strong>비용 구조:</strong> 판관비 ${fB(P.sga.total)}억 중 인건비 ${laborRatio.toFixed(0)}%, 판촉비 ${promoRatio.toFixed(0)}%로 상위 2개 항목이 전체의 ${(laborRatio+promoRatio).toFixed(0)}%를 차지합니다. 특히 소매채널 인건비율(${(SD['인건비합계'].retail/P.sga.retail*100).toFixed(1)}%)과 디지털채널 판촉비율(${(SD['판매촉진비'].digital/P.sga.digital*100).toFixed(1)}%)이 구조적 주요 원가요인입니다.
  </div>

  <div class="kpi-grid">
    <div class="stat-label">전사 실적</div>
    <div class="stat-row stat-row-3">
      <div class="stat-card stat-card-green"><div class="stat-name">매출총이익</div><div class="stat-v stat-v-green">${fB(P.gross.total)}억</div><div class="stat-sub">총이익률 ${(P.gross.total/P.revenue.total*100).toFixed(1)}%</div></div>
      <div class="stat-card stat-card-gray"><div class="stat-name">판관비</div><div class="stat-v stat-v-gray">${fB(P.sga.total)}억</div><div class="stat-sub">매출대비 ${(P.sga.total/P.revenue.total*100).toFixed(1)}%</div></div>
      <div class="stat-card stat-card-blue"><div class="stat-name">영업이익</div><div class="stat-v stat-v-blue">${fB(P.op.total)}억</div><div class="stat-sub">이익률 ${opm.toFixed(1)}%</div></div>
    </div>
    <div class="stat-label">채널별 영업이익률</div>
    <div class="stat-row stat-row-2">
      <div class="stat-card stat-card-best"><span class="stat-badge badge-best">▲ 최고</span><div class="stat-name">${CN[bestCh]}</div><div class="stat-v stat-v-best">${(P.op[bestCh]/P.revenue[bestCh]*100).toFixed(1)}%</div><div class="stat-sub">영업이익 ${fB(P.op[bestCh])}억</div></div>
      <div class="stat-card stat-card-worst"><span class="stat-badge badge-worst">▼ 최저</span><div class="stat-name">법인영업</div><div class="stat-v stat-v-worst">${(P.op.corporate_sales/P.revenue.corporate_sales*100).toFixed(1)}%</div><div class="stat-sub">영업이익 ${fB(P.op.corporate_sales)}억</div></div>
    </div>
    <div class="stat-label">본부별 영업이익률</div>
    <div class="stat-row stat-row-2">
      <div class="stat-card stat-card-best"><span class="stat-badge badge-best">▲ 최고</span><div class="stat-name">${bestHq.nm}</div><div class="stat-v stat-v-best">${((bestHq.gp-bestHq.sga)/bestHq.rev*100).toFixed(1)}%</div><div class="stat-sub">영업이익 ${fB(bestHq.gp-bestHq.sga)}억</div></div>
      <div class="stat-card stat-card-worst"><span class="stat-badge badge-worst">▼ 최저</span><div class="stat-name">${worstHq.nm}</div><div class="stat-v stat-v-worst">${((worstHq.gp-worstHq.sga)/worstHq.rev*100).toFixed(1)}%</div><div class="stat-sub">영업이익 ${fB(worstHq.gp-worstHq.sga)}억</div></div>
    </div>
  </div>

  <h1>Ⅱ. 채널별 손익 분석</h1>
  <h2>2-1. 채널별 손익 현황</h2>
  <table>
    <thead><tr><th>채널</th><th>매출(백만)</th><th>매출비중</th><th>매출총이익</th><th>총이익률</th><th>판관비</th><th>영업이익</th><th>영업이익률</th><th>평가</th></tr></thead>
    <tbody>`;
  const allChOpm=CHS.map(c=>P.op[c]/P.revenue[c]*100);
  const maxOpm=Math.max(...allChOpm),minOpm=Math.min(...allChOpm);
  CHS.forEach(c=>{
    const opmC=P.op[c]/P.revenue[c]*100,gpmC=P.gross[c]/P.revenue[c]*100;
    const badge=opmC>=15?'badge-r':(opmC>=5?'badge-y':(opmC>=2?'badge-o':'badge-red'));
    const grade=opmC>=15?'◎ 우수':(opmC>=5?'○ 양호':(opmC>=2?'△ 주의':'✕ 개선'));
    h+=`<tr><td>${CN[c]}</td><td>${fM(P.revenue[c])}</td><td>${(P.revenue[c]/P.revenue.total*100).toFixed(1)}%</td>
    <td>${fM(P.gross[c])}</td><td class="${gpmC>25?'pos':''}">  ${gpmC.toFixed(1)}%</td>
    <td>${fM(P.sga[c])}</td><td class="${opmC>0?'pos':'neg'}">${fM(P.op[c])}</td>
    <td class="${opmC>5?'pos':opmC>0?'':' neg'}">${opmC.toFixed(1)}%</td>
    <td><span class="badge ${badge}">${grade}</span></td></tr>`;
  });
  h+=`<tr class="total"><td>합 계</td><td>${fM(P.revenue.total)}</td><td>100%</td><td>${fM(P.gross.total)}</td><td>${(P.gross.total/P.revenue.total*100).toFixed(1)}%</td><td>${fM(P.sga.total)}</td><td class="pos">${fM(P.op.total)}</td><td class="pos">${opm.toFixed(1)}%</td><td>-</td></tr>
  </tbody></table>
  
  <div class="insight">📌 <strong>수익성 순위:</strong> 소상공인(${(P.op.small_biz/P.revenue.small_biz*100).toFixed(1)}%) &gt; IoT(${(P.op.iot/P.revenue.iot*100).toFixed(1)}%) &gt; 도매(${(P.op.wholesale/P.revenue.wholesale*100).toFixed(1)}%) &gt; 소매(${(P.op.retail/P.revenue.retail*100).toFixed(1)}%) &gt; 기업/공공(${(P.op.enterprise/P.revenue.enterprise*100).toFixed(1)}%) &gt; 디지털(${(P.op.digital/P.revenue.digital*100).toFixed(1)}%) &gt; 법인영업(${(P.op.corporate_sales/P.revenue.corporate_sales*100).toFixed(1)}%)</div>
  <div class="warn">⚠️ <strong>법인영업 이익률 ${(P.op.corporate_sales/P.revenue.corporate_sales*100).toFixed(1)}%:</strong> 매출 대비 판매수수료(${fB(SD['판매수수료'].corporate_sales)}억, 비율 ${(SD['판매수수료'].corporate_sales/P.revenue.corporate_sales*100).toFixed(1)}%)가 과중합니다. 수수료 구조 재설계가 필요합니다.</div>
  <div class="warn">⚠️ <strong>디지털 채널:</strong> 매출비중 ${(P.revenue.digital/P.revenue.total*100).toFixed(0)}%로 1위이나 이익률 ${(P.op.digital/P.revenue.digital*100).toFixed(1)}%에 그침. 판촉비(${fB(SD['판매촉진비'].digital)}억)가 영업이익(${fB(P.op.digital)}억)의 ${(SD['판매촉진비'].digital/P.op.digital).toFixed(1)}배에 달하는 구조적 문제 존재</div>

  <h1>Ⅲ. 판관비(SG&A) 구조 분석</h1>
  <h2>3-1. 판관비 항목별 현황</h2>
  <table>
    <thead><tr><th>비용항목</th><th>금액(백만)</th><th>판관비 비중</th><th>매출 대비</th><th>최다 채널</th></tr></thead>
    <tbody>`;
  const sgaItems=['인건비합계','판매촉진비','지급수수료','복리후생비','임차료','판매수수료','운반비'];
  const sgaColors={'인건비합계':'#3b82f6','판매촉진비':'#ef4444','지급수수료':'#8b5cf6','복리후생비':'#f59e0b','임차료':'#10b981','판매수수료':'#f97316','운반비':'#6366f1'};
  sgaItems.forEach(item=>{
    const topCh=CHS.reduce((a,b)=>(SD[item][a]||0)>(SD[item][b]||0)?a:b);
    const pctSga=(SD[item].total/P.sga.total*100).toFixed(1);
    const pctRev=(SD[item].total/P.revenue.total*100).toFixed(1);
    h+=`<tr><td><span style="color:${sgaColors[item]};font-weight:700">■</span> ${item}</td><td>${fM(SD[item].total)}</td><td>${pctSga}%</td><td>${pctRev}%</td><td>${CN[topCh]}(${fM(SD[item][topCh]||0)})</td></tr>`;
  });
  h+=`<tr class="total"><td>판관비 합계</td><td>${fM(P.sga.total)}</td><td>100%</td><td>${(P.sga.total/P.revenue.total*100).toFixed(1)}%</td><td>-</td></tr></tbody></table>
  
  <div class="caution">💡 <strong>인건비 집중도:</strong> 전체 판관비의 ${laborRatio.toFixed(1)}%(${fB(SD['인건비합계'].total)}억)가 인건비. 특히 소매 채널 인건비(${fB(SD['인건비합계'].retail)}억)는 소매 판관비의 ${(SD['인건비합계'].retail/P.sga.retail*100).toFixed(1)}%로 채널 내 비중 최고.</div>

  <h2>3-2. 채널별 비용구조 비교 (인건비 vs 판촉비)</h2>
  <table>
    <thead><tr><th>채널</th><th>인건비(백만)</th><th>판관비내 비중</th><th>판촉비(백만)</th><th>판관비내 비중</th><th>양대비용 합계</th><th>비용 특성</th></tr></thead>
    <tbody>`;
  CHS.forEach(c=>{
    const lb=SD['인건비합계'][c]||0,pr=SD['판매촉진비'][c]||0;
    const lbRt=(lb/P.sga[c]*100),prRt=(pr/P.sga[c]*100);
    const feature=prRt>lbRt?'판촉 집중':'인건비 집중';
    const fclass=prRt>lbRt?'caution':'insight';
    h+=`<tr><td>${CN[c]}</td><td>${fM(lb)}</td><td class="${lbRt>50?'warn':''} ">${lbRt.toFixed(1)}%</td><td>${fM(pr)}</td><td class="${prRt>40?'warn':''} ">${prRt.toFixed(1)}%</td><td>${((lbRt+prRt)).toFixed(1)}%</td><td><span class="badge ${prRt>lbRt?'badge-o':'badge-y'}">${feature}</span></td></tr>`;
  });
  h+=`</tbody></table>
  <div class="insight">📌 <strong>채널별 비용 특성 이원화:</strong> 소매·도매·기업 채널은 인건비 중심 구조(65~68%), 디지털 채널은 판촉비 중심 구조(59.1%)로 채널별 원가 절감 전략이 달라야 합니다.</div>

  <h1>Ⅳ. 본부별 수익성 분석</h1>
  <h2>4-1. 본부별 손익 현황</h2>
  <table>
    <thead><tr><th>순위</th><th>본부</th><th>매출(백만)</th><th>매출총이익</th><th>총이익률</th><th>판관비</th><th>판관비율</th><th>영업이익</th><th>영업이익률</th></tr></thead>
    <tbody>`;
  sortedHq.forEach((h2,i)=>{
    const op=h2.gp-h2.sga;const gpm=h2.gp/h2.rev*100;const opmH=op/h2.rev*100;const sgaRt=h2.sga/h2.rev*100;
    h+=`<tr><td style="text-align:center;font-weight:800">${i+1}</td><td>${h2.nm}</td><td>${fM(h2.rev)}</td><td>${fM(h2.gp)}</td><td class="${gpm>50?'pos':''}">${gpm.toFixed(1)}%</td><td>${fM(h2.sga)}</td><td class="${sgaRt<40?'pos':'neg'}">${sgaRt.toFixed(1)}%</td><td class="${op>0?'pos':'neg'}">${fM(op)}</td><td class="${opmH>5?'pos':'neg'}">${opmH.toFixed(1)}%</td></tr>`;
  });
  h+=`</tbody></table>
  <div class="insight">📌 <strong>${sortedHq[0].nm}</strong> 수익성 최우수(${((sortedHq[0].gp-sortedHq[0].sga)/sortedHq[0].rev*100).toFixed(1)}%). <strong>${sortedHq[sortedHq.length-1].nm}</strong> 최하위 — 판관비율 ${(sortedHq[sortedHq.length-1].sga/sortedHq[sortedHq.length-1].rev*100).toFixed(1)}%로 평균 대비 집중 점검 필요</div>

  <h1>Ⅴ. 리스크 워치리스트</h1>`;
  // Risk items
  const risks=[
    {lvl:'HIGH',item:'법인영업 수수료 구조',detail:`판매수수료 ${fB(SD['판매수수료'].corporate_sales)}억 vs 영업이익 ${fB(P.op.corporate_sales)}억 → 수수료가 이익의 ${(SD['판매수수료'].corporate_sales/P.op.corporate_sales).toFixed(1)}배`,action:'수수료 체계 재설계, 직영 전환 검토'},
    {lvl:'HIGH',item:'디지털 판촉비 효율성',detail:`판촉비 ${fB(SD['판매촉진비'].digital)}억 지출, 영업이익률 ${(P.op.digital/P.revenue.digital*100).toFixed(1)}%에 그침`,action:'ROI 기반 판촉 예산 재배분, 성과 지표 연동'},
    {lvl:'MED',item:'IoT 채널 수익 집중도',detail:`영업이익 ${fB(P.op.iot)}억(이익률 ${(P.op.iot/P.revenue.iot*100).toFixed(1)}%)이나 매출 비중 ${(P.revenue.iot/P.revenue.total*100).toFixed(1)}%로 편중 위험`,action:'IoT 채널 확대 전략 수립'},
    {lvl:'MED',item:'소매 인건비 효율성',detail:`인건비 ${fB(SD['인건비합계'].retail)}억, 판관비 대비 ${(SD['인건비합계'].retail/P.sga.retail*100).toFixed(1)}% — 1인당 생산성 관리 필요`,action:'영업직 직무수당 체계 연계, 생산성 KPI 강화'},
    {lvl:'LOW',item:`${sortedHq[sortedHq.length-1].nm} 판관비`,detail:`판관비율 ${(sortedHq[sortedHq.length-1].sga/sortedHq[sortedHq.length-1].rev*100).toFixed(1)}%로 본부 중 최고`,action:'비용 구조 진단 및 효율화 방안 수립'}
  ];
  h+=`<table class="priority-table">
    <thead><tr><th>등급</th><th>리스크 항목</th><th>현황</th><th>개선 방향</th></tr></thead><tbody>`;
  risks.forEach(r=>{
    const bc=r.lvl==='HIGH'?'badge-red':(r.lvl==='MED'?'badge-o':'badge-y');
    h+=`<tr><td style="text-align:center"><span class="badge ${bc}">${r.lvl}</span></td><td>${r.item}</td><td style="font-size:10.5px">${r.detail}</td><td style="font-size:10.5px">${r.action}</td></tr>`;
  });
  h+=`</tbody></table>

  <h1>Ⅵ. 전략적 실행 권고안</h1>
  <div class="action">
  ① <strong>[즉시] 법인영업 수수료 구조 점검 (1Q 내)</strong><br>
  &nbsp;&nbsp;&nbsp;판매수수료 ${fB(SD['판매수수료'].corporate_sales)}억이 영업이익(${fB(P.op.corporate_sales)}억)을 초과하는 구조적 문제 해소. 직영 전환 또는 수수료율 재협의 추진<br><br>
  ② <strong>[단기] 디지털 판촉 ROI 관리 체계 구축 (1Q~2Q)</strong><br>
  &nbsp;&nbsp;&nbsp;판촉비 ${fB(SD['판매촉진비'].digital)}억 지출 대비 효과 측정 지표 신설. 성과 연동형 예산 운영으로 이익률 목표 5% 이상 달성<br><br>
  ③ <strong>[중기] 고수익 채널 믹스 최적화 (2H)</strong><br>
  &nbsp;&nbsp;&nbsp;소상공인(이익률 ${(P.op.small_biz/P.revenue.small_biz*100).toFixed(1)}%) 채널 비중 확대 전략 수립. 현재 ${(P.revenue.small_biz/P.revenue.total*100).toFixed(1)}% → 목표 확대<br><br>
  ④ <strong>[중기] 소매 인건비 효율화 프로그램 (연간)</strong><br>
  &nbsp;&nbsp;&nbsp;영업직 직무수당 체계(완료) 기반 생산성 KPI 연계 강화. 1인당 매출/이익 기준 관리
  </div>

  <div class="footnote">* 본 보고서는 ${D.period} 공식 마감 데이터 기준입니다. 단위: 백만원(손익), 억원(요약). 대외비.</div>
  </div></body></html>`;
  openReport(h,'Executive 종합 보고서');
}

// ==================== ② 채널×비용 구조 심층분석 ====================
function genChannelCostReport(){
  const P=D.profit,SD=D.sga_detail;
  const sgaItems=['인건비합계','판매촉진비','복리후생비','임차료','판매수수료','지급수수료','운반비'];
  const itemColors={'인건비합계':'#2563eb','판매촉진비':'#dc2626','복리후생비':'#d97706','임차료':'#16a34a','판매수수료':'#7c3aed','지급수수료':'#0891b2','운반비':'#9333ea'};
  const chColors={retail:'#0369a1',wholesale:'#059669',digital:'#7c3aed',enterprise:'#c2410c',iot:'#ca8a04',corporate_sales:'#dc2626',small_biz:'#0f766e'};

  let h=`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>채널×비용 구조 분석</title>${rptStyle()}</head><body>
  <div class="cover">
    <div class="cover-tag">SG&amp;A Deep Dive · For Management</div>
    <div class="cover-title">채널별 비용구조 심층분석 보고서</div>
    <div class="cover-sub">${D.period} 기준 | 경영기획팀 | 채널×비용항목 교차 분석</div>
    <div class="cover-meta">
      <div class="cover-kpi"><div class="cover-kpi-v">${fB(P.sga.total)}억</div><div class="cover-kpi-l">총 판관비</div></div>
      <div class="cover-kpi"><div class="cover-kpi-v">${(P.sga.total/P.revenue.total*100).toFixed(1)}%</div><div class="cover-kpi-l">매출 대비 판관비율</div></div>
      <div class="cover-kpi"><div class="cover-kpi-v">${fB(SD['인건비합계'].total)}억</div><div class="cover-kpi-l">인건비 (${(SD['인건비합계'].total/P.sga.total*100).toFixed(0)}%)</div></div>
    </div>
  </div>
  <div class="page">

  <h1>Ⅰ. 채널별 비용구조 전략적 개요</h1>
  <div class="exec-box">
    전사 판관비 ${fB(P.sga.total)}억의 구조를 분석한 결과, 채널별로 뚜렷이 다른 2가지 비용 유형이 식별됩니다.<br><br>
    <strong>① 인건비 주도형 채널 (소매·도매·기업/공공)</strong>: 판관비의 65~68%가 인건비로 구성되며, 현장 영업인력 생산성이 수익성을 결정짓는 핵심 변수입니다.<br>
    <strong>② 판촉비 주도형 채널 (디지털)</strong>: 판관비의 59.1%가 판촉비로 구성. 마케팅 ROI 관리 체계가 미흡할 경우 이익이 급감하는 고위험 구조입니다.<br>
    <strong>③ 수수료 주도형 채널 (법인영업·IoT)</strong>: 판매수수료 비중이 높아 외부 의존도가 크며, 직영 전환 또는 수수료 재협상이 필요합니다.
  </div>

  <h1>Ⅱ. 채널별 비용항목 매트릭스</h1>
  <h2>2-1. 절대금액 기준 (단위: 백만원)</h2>
  <table>
    <thead><tr><th>비용항목</th><th>전사합계</th><th>소매</th><th>도매</th><th>디지털</th><th>기업/공공</th><th>IoT</th><th>법인영업</th><th>소상공인</th></tr></thead>
    <tbody>`;
  sgaItems.forEach(item=>{
    const topCh=CHS.reduce((a,b)=>(SD[item][a]||0)>(SD[item][b]||0)?a:b);
    h+=`<tr><td><span style="display:inline-block;width:8px;height:8px;background:${itemColors[item]};border-radius:2px;margin-right:5px;vertical-align:middle"></span>${item}</td>
    <td style="font-weight:700">${fM(SD[item].total)}</td>`;
    CHS.forEach(c=>{
      const v=SD[item][c]||0;const isTop=c===topCh;
      h+=`<td class="${isTop?'hi':''}">${fM(v)}${isTop?' ▲':''}</td>`;
    });
    h+='</tr>';
  });
  h+=`<tr class="total"><td>판관비 합계</td><td>${fM(P.sga.total)}</td>${CHS.map(c=>`<td>${fM(P.sga[c])}</td>`).join('')}</tr>
  </tbody></table>
  <div class="insight">▲ 표시: 채널 중 해당 항목 최대 지출 채널</div>

  <h2>2-2. 채널 판관비 내 항목 비중 (%) — 비용 DNA 분석</h2>
  <table>
    <thead><tr><th>비용항목</th><th>전사</th><th>소매</th><th>도매</th><th>디지털</th><th>기업/공공</th><th>IoT</th><th>법인영업</th><th>소상공인</th></tr></thead>
    <tbody>`;
  sgaItems.forEach(item=>{
    h+=`<tr><td>${item}</td><td>${(SD[item].total/P.sga.total*100).toFixed(1)}%</td>`;
    CHS.forEach(c=>{
      const r=(SD[item][c]||0)/P.sga[c]*100;
      const cls=r>50?'neg':(r>35?'caution':'');
      h+=`<td class="${cls}" style="${r>50?'font-weight:700;':''}">${r.toFixed(1)}%</td>`;
    });
    h+='</tr>';
  });
  h+=`<tr class="total"><td>합 계</td><td>100%</td>${CHS.map(()=>'<td>100%</td>').join('')}</tr></tbody></table>
  <div class="warn">⚠️ 채널 내 비중 50% 이상 항목: 집중 관리 대상 (빨간색). 35% 이상: 주의 관찰 대상</div>

  <h2>2-3. 매출 대비 비용항목 비율 (%) — 채널 효율성 지표</h2>
  <table>
    <thead><tr><th>비용항목</th><th>전사</th><th>소매</th><th>도매</th><th>디지털</th><th>기업/공공</th><th>IoT</th><th>법인영업</th><th>소상공인</th></tr></thead>
    <tbody>`;
  sgaItems.forEach(item=>{
    h+=`<tr><td>${item}</td><td>${(SD[item].total/P.revenue.total*100).toFixed(1)}%</td>`;
    CHS.forEach(c=>{
      const r=(SD[item][c]||0)/P.revenue[c]*100;
      h+=`<td class="${r>20?'neg':r>10?'':''}">${r.toFixed(1)}%</td>`;
    });
    h+='</tr>';
  });
  h+=`<tr class="total" style="background:#dbeafe"><td>영업이익률</td><td class="pos">${(P.op.total/P.revenue.total*100).toFixed(1)}%</td>${CHS.map(c=>`<td class="${P.op[c]/P.revenue[c]>0.05?'pos':'neg'}">${(P.op[c]/P.revenue[c]*100).toFixed(1)}%</td>`).join('')}</tr></tbody></table>

  <h1>Ⅲ. 채널별 비용구조 심층 분석</h1>`;

  // Per channel detailed analysis
  const chInsights = {
    retail:{title:'소매 채널',color:'#0369a1',key:'인건비 주도형',summary:'인건비 집중 구조. 영업인력 생산성이 핵심 레버',risk:'인건비율 64.4% — 판관비의 절반 이상이 인건비',opp:'직무수당 연계 생산성 향상 시 이익률 1%p 개선 시 순이익 +15억'},
    wholesale:{title:'도매 채널',color:'#059669',key:'인건비 주도형',summary:'안정적 수익 구조. 인건비 효율이 관건',risk:'인건비율 67.2% — 위탁대가 정산 이슈 지속 관리 필요',opp:'엣지 프로그램 2.0 통한 대리점 생산성 향상'},
    digital:{title:'디지털 채널',color:'#7c3aed',key:'판촉비 주도형',summary:'매출 1위 채널이나 판촉비 과다로 수익성 저조',risk:`판촉비 ${fB(SD['판매촉진비'].digital)}억 — 영업이익(${fB(P.op.digital)}억)의 ${(SD['판매촉진비'].digital/P.op.digital).toFixed(1)}배 지출`,opp:'ROI 기반 판촉 예산 재배분 → 이익률 목표 5% 달성 시 +73억'},
    enterprise:{title:'기업/공공 채널',color:'#c2410c',key:'인건비 주도형',summary:'이익률 안정적. 규모 확대가 관건',risk:'매출 비중 1.8%로 낮아 고정비 부담 상대적으로 큼',opp:'대형사업 수주율 확대 과제 연계 — 매출 확대 시 고정비 레버리지 효과'},
    iot:{title:'IoT 채널',color:'#ca8a04',key:'수수료 주도형',summary:'최고 이익률 채널이나 매출 규모 소형',risk:`판매수수료 ${fB(SD['판매수수료'].iot)}억 — 매출 대비 ${(SD['판매수수료'].iot/P.revenue.iot*100).toFixed(1)}%로 높음`,opp:'IoT 원격관제 솔루션 확장 과제 연계 → 채널 성장 가속화'},
    corporate_sales:{title:'법인영업 채널',color:'#dc2626',key:'수수료 주도형',summary:'수수료 지출이 이익을 초과하는 구조적 문제',risk:`판매수수료 ${fB(SD['판매수수료'].corporate_sales)}억 >> 영업이익 ${fB(P.op.corporate_sales)}억 — 수수료율 재설계 시급`,opp:'직영전환 또는 수수료 인하 협상으로 수익 구조 개선'},
    small_biz:{title:'소상공인 채널',color:'#0f766e',key:'균형형 (최고 수익)',summary:'판관비 구조가 가장 균형적이며 이익률 최고',risk:'복리후생비 비중이 상대적으로 높음(27.8%)',opp:'채널 규모 확대를 통한 전사 수익성 제고 — 비중 확대 우선 과제'}
  };

  CHS.forEach(c=>{
    const ins=chInsights[c];const rev=P.revenue[c],sga=P.sga[c],op=P.op[c];
    const top2=sgaItems.slice().sort((a,b)=>(SD[b][c]||0)-(SD[a][c]||0)).slice(0,3);
    h+=`<div class="ch-cost-block">
      <div class="ch-cost-header" style="background:${ins.color}">
        <span>${ins.title} <span class="ch-cost-sub">| ${ins.key}</span></span>
        <span>영업이익률 ${(op/rev*100).toFixed(1)}%</span>
      </div>
      <div class="ch-cost-body">
        <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-bottom:10px">
          <div style="text-align:center;padding:8px;background:#f8fafc;border-radius:6px"><div style="font-size:18px;font-weight:900;color:${ins.color}">${fB(rev)}억</div><div style="font-size:9px;color:#64748b">매출</div></div>
          <div style="text-align:center;padding:8px;background:#f8fafc;border-radius:6px"><div style="font-size:18px;font-weight:900;color:#64748b">${fB(sga)}억</div><div style="font-size:9px;color:#64748b">판관비 (${(sga/rev*100).toFixed(1)}%)</div></div>
          <div style="text-align:center;padding:8px;background:#f8fafc;border-radius:6px"><div style="font-size:18px;font-weight:900;color:${op>0?'#059669':'#dc2626'}">${fB(op)}억</div><div style="font-size:9px;color:#64748b">영업이익</div></div>
        </div>
        <div style="font-size:10.5px;font-weight:700;color:#475569;margin-bottom:5px">주요 비용 항목 TOP 3</div>`;
    top2.forEach(item=>{
      const v=SD[item][c]||0;const pct=(v/sga*100).toFixed(1);const pctRev=(v/rev*100).toFixed(1);
      const w=Math.min(100,(v/sga*100)*1.2);
      h+=`<div class="bar-row"><div class="bar-label">${item.replace('합계','')}</div><div class="bar-bg"><div class="bar-fill" style="width:${w}%;background:${itemColors[item]}">${pct}%</div></div><div class="bar-val">${fM(v)}백만</div></div>`;
    });
    h+=`<div style="margin-top:8px;font-size:10.5px;padding:7px 10px;border-radius:6px;background:rgba(15,23,42,.65);color:#7f1d1d;border-left:3px solid #dc2626"><strong>리스크:</strong> ${ins.risk}</div>
    <div style="margin-top:5px;font-size:10.5px;padding:7px 10px;border-radius:6px;background:rgba(15,23,42,.65);color:#14532d;border-left:3px solid #16a34a"><strong>기회:</strong> ${ins.opp}</div>
      </div></div>`;
  });

  h+=`<h1>Ⅳ. 비용 최적화 우선순위 매트릭스</h1>
  <table class="priority-table">
    <thead><tr><th>우선순위</th><th>개선 과제</th><th>현황</th><th>목표</th><th>예상 효과</th></tr></thead>
    <tbody>
    <tr><td style="text-align:center"><span class="badge badge-red">1순위</span></td><td>법인영업 수수료 재설계</td><td>판매수수료 ${fB(SD['판매수수료'].corporate_sales)}억, 영업이익 ${fB(P.op.corporate_sales)}억</td><td>수수료율 30% 인하</td><td>+${(SD['판매수수료'].corporate_sales*0.3/1e8).toFixed(1)}억 이익 개선</td></tr>
    <tr><td style="text-align:center"><span class="badge badge-red">1순위</span></td><td>디지털 판촉비 ROI 관리</td><td>판촉비 ${fB(SD['판매촉진비'].digital)}억, 이익률 ${(P.op.digital/P.revenue.digital*100).toFixed(1)}%</td><td>이익률 5% 달성</td><td>+${((P.revenue.digital*0.05-P.op.digital)/1e8).toFixed(1)}억 이익 개선</td></tr>
    <tr><td style="text-align:center"><span class="badge badge-o">2순위</span></td><td>소매 인건비 생산성 향상</td><td>인건비 ${fB(SD['인건비합계'].retail)}억, 판관비내 ${(SD['인건비합계'].retail/P.sga.retail*100).toFixed(1)}%</td><td>1인당 매출 3% 향상</td><td>인건비 효율 동일 시 이익률 +0.3%p</td></tr>
    <tr><td style="text-align:center"><span class="badge badge-y">3순위</span></td><td>소상공인 채널 비중 확대</td><td>현 매출비중 ${(P.revenue.small_biz/P.revenue.total*100).toFixed(1)}%, 이익률 ${(P.op.small_biz/P.revenue.small_biz*100).toFixed(1)}%</td><td>비중 5% 달성</td><td>이익률 최고 채널 믹스 개선</td></tr>
    </tbody>
  </table>

  <div class="footnote">* 본 보고서는 ${D.period} 공식 마감 데이터 기준입니다. 채널별 비용은 원가 배부 방식에 따라 차이가 있을 수 있습니다. 단위: 백만원(상세), 억원(요약). 대외비.</div>
  </div></body></html>`;
  openReport(h,'채널×비용 구조 심층분석');
}

// ==================== ③ 기타 보고서 ====================
function genProfitReport(type){
  const P=D.profit,SD=D.sga_detail,HC=D.hq_channel||{},K=D.kpi;
  const RS=rptStyle();
  const sortedHq=[...D.hq].sort((a,b)=>(b.gp-b.sga)/b.rev-(a.gp-a.sga)/a.rev);
  const fPct=v=>(v*100).toFixed(1)+'%';
  const fBn=n=>(n/1e8).toFixed(1);
  const fMn=n=>Math.round(n/1e6).toLocaleString();
  const clrOpm=v=>v>0.1?'#059669':v>0.05?'#d97706':'#dc2626';
  const bar=(v,w,c)=>`<div style="width:${w||120}px;height:10px;background:#e2e8f0;border-radius:5px;overflow:hidden;display:inline-block;vertical-align:middle;margin-left:6px"><div style="width:${Math.min(100,Math.round(Math.abs(v)*200))}%;height:100%;background:${c||clrOpm(v)};border-radius:5px"></div></div>`;

  if(type==='hq'){
    // ── 본부별 채널 교차 손익 고도화 보고서 ──────────────────────────────
    let h=`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>본부별 채널 교차 손익</title>${RS}</head><body>`;
    
    // Cover
    const best=sortedHq[0],worst=sortedHq[sortedHq.length-1];
    const bestOp=(best.gp-best.sga)/best.rev, worstOp=(worst.gp-worst.sga)/worst.rev;
    h+=`<div class="cover">
      <div class="cover-tag">Confidential · 본부별 채널 교차 손익 분석 보고서</div>
      <div class="cover-title">본부별 채널 교차<br>수익성 심층 분석</div>
      <div class="cover-sub">${D.period} 기준 | 경영기획팀 | 본부×채널 교차 수익성 · 비용구조 · 전략 권고</div>
      <div style="display:flex;gap:12px;margin-top:20px;flex-wrap:wrap">
        ${sortedHq.map((hq,i)=>{const op=hq.gp-hq.sga;const opm=op/hq.rev;return `<div style="flex:1;min-width:140px;background:rgba(255,255,255,.08);border-radius:10px;padding:12px 14px">
          <div style="font-size:10px;color:rgba(255,255,255,.5);font-weight:600">#${i+1} ${hq.nm}</div>
          <div style="font-size:22px;font-weight:900;color:${opm>0.07?'#34d399':opm>0.04?'#fbbf24':'#f87171'}">${(opm*100).toFixed(1)}%</div>
          <div style="font-size:10px;color:rgba(255,255,255,.5)">영업이익 ${fBn(op)}억</div>
        </div>`;}).join('')}
      </div>
    </div>`;

    // 1. Executive Summary
    h+=`<h2>Ⅰ. Executive Summary</h2>
    <div class="exec-box">
      수익성 1위 <b>${best.nm}</b>(영업이익률 ${(bestOp*100).toFixed(1)}%) vs 5위 <b>${worst.nm}</b>(${(worstOp*100).toFixed(1)}%) — 격차 <b>${((bestOp-worstOp)*100).toFixed(1)}%p</b><br>
      전체 5개 본부 합산: 매출 <b>${fBn(D.hq.reduce((s,h)=>s+h.rev,0))}억</b> | 영업이익 <b>${fBn(D.hq.reduce((s,h)=>s+(h.gp-h.sga),0))}억</b>
    </div>`;

    // 2. 본부별 전체 손익 순위 테이블
    h+=`<h2>Ⅱ. 본부별 손익 현황 (수익성 순위)</h2>
    <table><thead><tr><th>순위</th><th>본부</th><th class="num">매출(백만)</th><th class="num">총이익</th><th class="num">총이익률</th><th class="num">판관비</th><th class="num">판관비율</th><th class="num">영업이익</th><th class="num">영업이익률</th></tr></thead><tbody>`;
    sortedHq.forEach((hq,i)=>{
      const op=hq.gp-hq.sga, opm=op/hq.rev;
      h+=`<tr><td style="text-align:center;font-weight:800;font-size:16px">${i+1}</td>
        <td><b>${hq.nm}</b></td>
        <td class="num">${fMn(hq.rev)}</td>
        <td class="num">${fMn(hq.gp)}</td>
        <td class="num ${hq.gp/hq.rev>0.5?'pos':''}">${(hq.gp/hq.rev*100).toFixed(1)}%</td>
        <td class="num">${fMn(hq.sga)}</td>
        <td class="num ${hq.sga/hq.rev<0.4?'pos':''}">${(hq.sga/hq.rev*100).toFixed(1)}%</td>
        <td class="num ${op>0?'pos':'neg'}" style="font-weight:700">${fMn(op)}</td>
        <td class="num ${opm>0.05?'pos':'neg'}" style="font-weight:700">${(opm*100).toFixed(1)}%</td></tr>`;
    });
    h+=`</tbody></table>`;

    // 3. 채널×본부 교차 매트릭스
    const kpiChMap={'강서본부':'강서본부','동부본부':'동부본부','서부본부':'서부본부','강북본부':'강북본부','강남본부':'강남본부'};
    const hqOrder=['강서본부','동부본부','서부본부','강북본부','강남본부']; // KPI order
    const chLabels=[
      {key:'retail',nm:'소매',color:'#38bdf8',kpiKey:'rt'},
      {key:'wholesale',nm:'도매',color:'#34d399',kpiKey:'wh'},
      {key:'digital',nm:'디지털',color:'#a78bfa',kpiKey:null},
      {key:'small_biz',nm:'소상공인',color:'#818cf8',kpiKey:'sm'}
    ];

    h+=`<h2>Ⅲ. 채널×본부 교차 영업이익 매트릭스</h2>
    <div class="insight">📌 각 셀은 해당 본부의 채널별 추정 영업이익(백만원). 배경색: 🟢높음 🟡보통 🔴낮음 · KPI 득점과 교차 검증</div>
    <table><thead><tr><th>채널</th><th style="text-align:center">가중치</th>`;
    hqOrder.forEach(hq=>{ h+=`<th class="num">${hq}</th>`; });
    h+=`<th class="num">전사 합계</th></tr></thead><tbody>`;

    chLabels.forEach(ch=>{
      const vals=hqOrder.map(hq=>HC[hq]?.op?.[ch.key]||0);
      const maxV=Math.max(...vals), minV=Math.min(...vals);
      const totalCh=P.op[ch.key];
      const wt={retail:'60%',wholesale:'30%',small_biz:'10%',digital:'—'}[ch.key]||'—';
      h+=`<tr><td><span style="display:inline-block;width:10px;height:10px;background:${ch.color};border-radius:50%;margin-right:6px"></span><b>${ch.nm}</b></td>
        <td style="text-align:center;font-size:11px;color:#64748b">${wt}</td>`;
      vals.forEach((v,vi)=>{
        const bg=v===maxV?'#dcfce7':v===minV?'#fee2e2':'#f1f5f9';
        const fc=v===maxV?'#059669':v===minV?'#dc2626':'#1e293b';
        const kpiV=ch.kpiKey?K.find(k=>k.hq===hqOrder[vi])?.[ch.kpiKey]?.t:null;
        h+=`<td class="num" style="background:${bg};color:${fc};font-weight:700">${fMn(v)}<br>
          ${kpiV?`<span style="font-size:9px;color:${kpiV>=90?'#059669':kpiV>=80?'#d97706':'#dc2626'};font-weight:600">KPI ${kpiV.toFixed(1)}</span>`:''}
        </td>`;
      });
      h+=`<td class="num" style="font-weight:700;color:${totalCh>0?'#059669':'#dc2626'}">${fMn(totalCh)}</td></tr>`;
    });
    h+=`</tbody></table>`;

    // 4. 본부별 채널 구성 분석 (각 본부 심층)
    h+=`<h2>Ⅳ. 본부별 채널 구성 심층 분석</h2>`;
    
    sortedHq.forEach((hq, rank)=>{
      const op=hq.gp-hq.sga, opm=op/hq.rev;
      const hqCh=HC[hq.nm]||{};
      const kpiHq=K.find(k=>k.hq===hq.nm)||{};
      
      // Channel breakdown for this HQ
      const chBreakdown=chLabels.map(ch=>({
        ...ch,
        rev: hqCh.rev?.[ch.key]||0,
        gp: hqCh.gp?.[ch.key]||0,
        opVal: hqCh.op?.[ch.key]||0,
        kpiScore: ch.kpiKey?kpiHq[ch.kpiKey]?.t:null
      }));
      const totalChRev=chBreakdown.reduce((s,c)=>s+c.rev,0);
      
      // Find strongest/weakest channel
      const chWithOp=chBreakdown.filter(c=>c.opVal!==0);
      const bestCh=chWithOp.length?chWithOp.reduce((a,b)=>b.opVal/b.rev>a.opVal/a.rev?b:a):null;
      const worstCh=chWithOp.length?chWithOp.reduce((a,b)=>b.opVal/b.rev<a.opVal/a.rev?b:a):null;
      
      const borderClr=rank===0?'#059669':rank<=1?'#0284c7':rank<=2?'#d97706':'#dc2626';
      h+=`<h3 style="border-left:5px solid ${borderClr};padding-left:12px;margin:24px 0 12px">${rank+1}위. ${hq.nm} — 영업이익률 ${(opm*100).toFixed(1)}%</h3>`;

      // Summary boxes
      h+=`<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:14px">
        <div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:10px;text-align:center">
          <div style="font-size:9px;color:#64748b;font-weight:600;margin-bottom:2px">매출</div>
          <div style="font-size:18px;font-weight:800;color:#1e293b">${fBn(hq.rev)}억</div>
        </div>
        <div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:10px;text-align:center">
          <div style="font-size:9px;color:#64748b;font-weight:600;margin-bottom:2px">총이익률</div>
          <div style="font-size:18px;font-weight:800;color:#0284c7">${(hq.gp/hq.rev*100).toFixed(1)}%</div>
        </div>
        <div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:10px;text-align:center">
          <div style="font-size:9px;color:#64748b;font-weight:600;margin-bottom:2px">판관비율</div>
          <div style="font-size:18px;font-weight:800;color:#d97706">${(hq.sga/hq.rev*100).toFixed(1)}%</div>
        </div>
        <div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:10px;text-align:center">
          <div style="font-size:9px;color:#64748b;font-weight:600;margin-bottom:2px">영업이익</div>
          <div style="font-size:18px;font-weight:800;color:${opm>0.05?'#059669':'#dc2626'}">${fBn(op)}억</div>
        </div>
      </div>`;

      // Channel detail table for this HQ
      h+=`<table><thead><tr><th>채널</th><th class="num">매출(백만)</th><th class="num">매출비중</th><th class="num">총이익</th><th class="num">총이익률</th><th class="num">영업이익</th><th class="num">영업이익률</th><th class="num">KPI득점</th></tr></thead><tbody>`;
      chBreakdown.forEach(ch=>{
        if(ch.rev===0) return;
        const opmCh=ch.rev>0?ch.opVal/ch.rev:0;
        const gpRate=ch.rev>0?ch.gp/ch.rev:0;
        const isBest=bestCh&&ch.key===bestCh.key;
        const isWorst=worstCh&&ch.key===worstCh.key;
        h+=`<tr style="${isBest?'background:rgba(15,23,42,.65)':isWorst?'background:rgba(15,23,42,.65)':''}">
          <td><span style="display:inline-block;width:8px;height:8px;background:${ch.color};border-radius:50%;margin-right:6px"></span>
            <b>${ch.nm}</b>${isBest?' 🏆':isWorst?' ⚠️':''}</td>
          <td class="num">${fMn(ch.rev)}</td>
          <td class="num">${totalChRev>0?(ch.rev/totalChRev*100).toFixed(1):'—'}%</td>
          <td class="num">${ch.gp?fMn(ch.gp):'—'}</td>
          <td class="num ${gpRate>0.5?'pos':''}">${ch.gp?(gpRate*100).toFixed(1)+'%':'—'}</td>
          <td class="num ${ch.opVal>0?'pos':'neg'}" style="font-weight:700">${fMn(ch.opVal)}</td>
          <td class="num ${opmCh>0.05?'pos':'neg'}" style="font-weight:700">${(opmCh*100).toFixed(1)}%</td>
          <td class="num">${ch.kpiScore?`<span style="color:${ch.kpiScore>=90?'#059669':ch.kpiScore>=80?'#d97706':'#dc2626'};font-weight:700">${ch.kpiScore.toFixed(1)}</span>`:'—'}</td>
        </tr>`;
      });
      h+=`</tbody></table>`;

      // Channel revenue mix bar chart
      h+=`<div style="margin:12px 0 6px;font-size:11px;font-weight:600;color:#475569">채널별 매출 구성</div>
      <div style="display:flex;height:28px;border-radius:8px;overflow:hidden;gap:1px">`;
      chBreakdown.filter(c=>c.rev>0).forEach(ch=>{
        const pct=totalChRev>0?ch.rev/totalChRev*100:0;
        h+=`<div title="${ch.nm}: ${pct.toFixed(1)}%" style="width:${pct.toFixed(1)}%;background:${ch.color};display:flex;align-items:center;justify-content:center;font-size:9px;color:#fff;font-weight:700;white-space:nowrap;overflow:hidden;min-width:${pct>8?'auto':'0'}">${pct>8?ch.nm:''}`;
        h+=`</div>`;
      });
      h+=`</div><div style="display:flex;gap:12px;margin-top:6px;flex-wrap:wrap">`;
      chBreakdown.filter(c=>c.rev>0).forEach(ch=>{
        const pct=totalChRev>0?ch.rev/totalChRev*100:0;
        h+=`<span style="font-size:10px;color:#64748b"><span style="display:inline-block;width:8px;height:8px;background:${ch.color};border-radius:2px;margin-right:3px"></span>${ch.nm} ${pct.toFixed(1)}%</span>`;
      });
      h+=`</div>`;

      // Insight
      const channelInsight=bestCh&&worstCh?
        `<b>강점:</b> ${bestCh.nm} 채널 영업이익률 ${bestCh.rev>0?(bestCh.opVal/bestCh.rev*100).toFixed(1):0}% 최고 ${bestCh.kpiScore?`(KPI ${bestCh.kpiScore.toFixed(1)}점)`:''} — <b>약점:</b> ${worstCh.nm} 채널 ${worstCh.rev>0?(worstCh.opVal/worstCh.rev*100).toFixed(1):0}% 하위`:
        '채널 분석 데이터 부족';
      h+=`<div class="insight" style="margin-top:12px">📌 ${hq.nm} 진단: ${channelInsight}</div>`;
    });

    // 5. 채널별 본부 비교 분석
    h+=`<h2>Ⅴ. 채널별 본부 수익성 비교</h2>`;
    chLabels.forEach(ch=>{
      const vals=hqOrder.map(hq=>({hq,op:HC[hq]?.op?.[ch.key]||0,rev:HC[hq]?.rev?.[ch.key]||0}));
      const maxOpmHq=vals.reduce((a,b)=>(b.rev>0?b.op/b.rev:0)>(a.rev>0?a.op/a.rev:0)?b:a);
      const minOpmHq=vals.reduce((a,b)=>(b.rev>0?b.op/b.rev:0)<(a.rev>0?a.op/a.rev:0)?b:a);
      const kpiAvg=ch.kpiKey?K.map(k=>k[ch.kpiKey]?.t||0).reduce((a,b)=>a+b,0)/K.length:null;

      h+=`<h3 style="border-left:4px solid ${ch.color};padding-left:10px">${ch.nm} 채널 본부 비교</h3>`;
      h+=`<table><thead><tr><th>본부</th><th class="num">매출(백만)</th><th class="num">영업이익</th><th class="num">영업이익률</th>${ch.kpiKey?'<th class="num">KPI득점</th>':''}</tr></thead><tbody>`;
      
      [...vals].sort((a,b)=>(b.rev>0?b.op/b.rev:0)-(a.rev>0?a.op/a.rev:0)).forEach((v,rank)=>{
        const opmV=v.rev>0?v.op/v.rev:0;
        const kpiV=ch.kpiKey?K.find(k=>k.hq===v.hq)?.[ch.kpiKey]?.t:null;
        const isBest=v.hq===maxOpmHq.hq, isWorst=v.hq===minOpmHq.hq;
        h+=`<tr style="${isBest?'background:rgba(15,23,42,.65)':isWorst?'background:rgba(15,23,42,.65)':''}">
          <td>${rank+1}. <b>${v.hq}</b>${isBest?' 🏆':isWorst?' ⚠️':''}</td>
          <td class="num">${fMn(v.rev)}</td>
          <td class="num ${v.op>0?'pos':'neg'}" style="font-weight:700">${fMn(v.op)}</td>
          <td class="num ${opmV>0.05?'pos':'neg'}" style="font-weight:700">${(opmV*100).toFixed(1)}%</td>
          ${ch.kpiKey?`<td class="num"><span style="font-weight:700;color:${kpiV>=90?'#059669':kpiV>=80?'#d97706':'#dc2626'}">${kpiV?kpiV.toFixed(1):'-'}</span></td>`:''}
        </tr>`;
      });
      h+=`</tbody></table>`;
      if(ch.kpiKey){
        h+=`<div class="insight">📌 ${ch.nm}: 1위 <b>${maxOpmHq.hq}</b>(${maxOpmHq.rev>0?(maxOpmHq.op/maxOpmHq.rev*100).toFixed(1):0}%) vs 최하 <b>${minOpmHq.hq}</b>(${minOpmHq.rev>0?(minOpmHq.op/minOpmHq.rev*100).toFixed(1):0}%) · KPI 평균 ${kpiAvg?kpiAvg.toFixed(1):'-'}점</div>`;
      }
    });

    // 6. 전략 권고
    const lowestHq=sortedHq[sortedHq.length-1];
    const lowestOp=(lowestHq.gp-lowestHq.sga)/lowestHq.rev;
    h+=`<h2>Ⅵ. 전략적 권고사항</h2>
    <div class="action">
      <b>① [즉시] ${lowestHq.nm} 수익성 개선 TF 구성</b><br>
      영업이익률 ${(lowestOp*100).toFixed(1)}% — 판관비율 ${(lowestHq.sga/lowestHq.rev*100).toFixed(1)}% 구조적 재검토, 채널별 비용 적정성 점검<br><br>
      <b>② [단기] 소상공인 채널 본부간 격차 해소</b><br>
      1위 강서본부 노하우 이식 — 하위 강남본부 전담 지원 TF 운영, 커버리지 52→65% 3개월 목표<br><br>
      <b>③ [정기] 본부×채널 월간 수익성 모니터링</b><br>
      각 본부의 채널별 영업이익률 매월 추적, KPI 득점과의 상관관계 분석 정례화<br><br>
      <b>④ [중기] 채널 포트폴리오 최적화</b><br>
      도매 채널 수익성(전사 ${(P.op.wholesale/P.revenue.wholesale*100).toFixed(1)}%)이 소매(${(P.op.retail/P.revenue.retail*100).toFixed(1)}%)보다 낮은 본부 집중 점검 — 비용구조 재설계
    </div>`;

    h+=`<p style="margin-top:20px;font-size:10px;color:#94a3b8;border-top:1px solid #e2e8f0;padding-top:12px">
      ⚠️ 본 보고서의 채널별 본부 손익은 KPI 득점 비율 기반 추정치입니다. 채널 실적 마감 데이터 기준으로 업데이트 필요.
    </p>`;
    h+=`</div></body></html>`;
    openReport(h,'본부별 채널 교차 손익 분석');

  } else if(type==='channel'){
    let h=`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>채널별 손익 분석</title>${RS}</head><body><div class="page">`;
    h+=`<h1>📊 채널별 손익 분석 보고서</h1><p style="color:#64748b;font-size:11px;margin-bottom:16px">${D.period} 기준</p>`;
    CHS.forEach(c=>{
      const opm=(P.op[c]/P.revenue[c]*100);
      h+=`<h2>${CN[c]} 채널</h2>
      <div class="${opm>10?'action':opm>5?'insight':opm>0?'insight':'warn'}">
      매출 ${fBn(P.revenue[c])}억(비중 ${(P.revenue[c]/P.revenue.total*100).toFixed(1)}%) | 총이익률 ${(P.gross[c]/P.revenue[c]*100).toFixed(1)}% | 영업이익 <strong>${fBn(P.op[c])}억(${opm.toFixed(1)}%)</strong>
      </div>`;
    });
    h+=`</div></body></html>`;
    openReport(h,'채널별 손익 분석');
  }
}



function genTaskReport(){
  const T=D.tasks,total=T.length,done=T.filter(t=>t.st==='완료').length,prog=T.filter(t=>t.st==='진행중').length;
  const core=T.filter(t=>t.co==='●'),zero=T.filter(t=>t.pp===0&&t.st==='진행중');
  const tms={};T.forEach(t=>{if(!tms[t.tm])tms[t.tm]={total:0,done:0,avg:0};tms[t.tm].total++;if(t.st==='완료')tms[t.tm].done++;tms[t.tm].avg+=t.pp;});
  Object.keys(tms).forEach(k=>tms[k].avg=tms[k].avg/tms[k].total);
  const sorted=Object.entries(tms).sort((a,b)=>b[1].avg-a[1].avg);
  let h=`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>과제관리보고서</title>${rptStyle()}</head><body><div class="page">
  <h1>📋 KT M&S 전사 과제 관리 보고서</h1><p style="color:#64748b;font-size:11px;margin-bottom:16px">${D.period} 기준 · 경영기획팀</p>
  <div class="exec-box">전체 <b>${total}개</b> 과제 · 완료 <b>${done}개</b>(${(done/total*100).toFixed(1)}%) · 진행중 <b>${prog}개</b> · 핵심과제 ${core.length}개</div>
  <div class="warn">⚠️ 진행중 + 추진도 0% 과제 <b>${zero.length}건</b> — 즉시 점검 필요</div>
  <h2>팀별 현황</h2><table><thead><tr><th>팀명</th><th class="num">과제수</th><th class="num">완료</th><th>평균추진도</th></tr></thead><tbody>`;
  sorted.forEach(([k,v])=>{h+=`<tr><td>${k}</td><td style="text-align:center">${v.total}</td><td style="text-align:center">${v.done}</td><td style="text-align:right;font-weight:600" class="${v.avg>.3?'pos':'neg'}">${(v.avg*100).toFixed(0)}%</td></tr>`;});
  h+=`</tbody></table>
  <h2>핵심과제 현황</h2><table><thead><tr><th>팀</th><th>과제명</th><th>담당자</th><th>추진도</th></tr></thead><tbody>`;
  core.forEach(t=>{h+=`<tr><td>${t.tm}</td><td>${t.nm}</td><td>${t.mg||''}</td><td class="${t.pp>.3?'pos':'neg'}" style="text-align:right">${(t.pp*100).toFixed(0)}%</td></tr>`;});
  h+=`</tbody></table>
  <div class="action">① 추진도 0% 과제 담당자 면담 및 장애요인 파악<br>② 핵심과제 15% 미만 주간 점검 체계 도입</div>
  </div></body></html>`;
  openReport(h,'과제 관리 보고서');
}

// ==================== KPI REPORT ====================
// ==================== 총괄별 과제 심층 분석 보고서 ====================
function genTaskHqReport(){
  const T=D.tasks;
  const RS=rptStyle();

  // 채널-팀 매핑
  const TM_CHANNEL={
    '소매강화팀':'소매','디지털강화팀':'디지털','도매강화팀':'도매',
    '소상공인혁신팀':'소상공인','법인영업팀':'기업/공공','기업영업팀':'기업/공공',
    '현장지원팀':'영업지원','영업기획팀':'영업지원','영업품질팀':'영업지원','역량강화팀':'영업지원',
    '경영기획팀':'경영지원','회계세무팀':'경영지원','인사팀':'경영지원','경영지원팀':'경영지원',
    'IT보안팀':'IT/디지털','AI확산팀':'IT/디지털','플랫폼사업1팀':'플랫폼','플랫폼사업2팀':'플랫폼','모바일파트':'플랫폼'
  };

  // 총괄별 데이터 집계
  const chs=['경총','영총'];
  const chData={};
  chs.forEach(c=>{
    const tasks=T.filter(t=>t.ch===c);
    const teams=[...new Set(tasks.map(t=>t.tm))];
    const done=tasks.filter(t=>t.st==='완료');
    const prog=tasks.filter(t=>t.st==='진행중');
    const plan=tasks.filter(t=>t.st==='계획');
    const core=tasks.filter(t=>t.co==='●');
    const zero=tasks.filter(t=>t.st==='진행중'&&t.pp===0);
    const low=tasks.filter(t=>t.st==='진행중'&&t.pp>0&&t.pp<=0.15);
    const avgPp=tasks.reduce((s,t)=>s+t.pp,0)/tasks.length;
    const coreAvgPp=core.length?core.reduce((s,t)=>s+t.pp,0)/core.length:0;
    chData[c]={tasks,teams,done,prog,plan,core,zero,low,avgPp,coreAvgPp};
  });

  // 팀별 데이터
  const allTeams=[...new Set(T.map(t=>t.tm))].sort();
  const teamData={};
  allTeams.forEach(tm=>{
    const tasks=T.filter(t=>t.tm===tm);
    const ch=tasks[0]?.ch||'';
    const done=tasks.filter(t=>t.st==='완료').length;
    const prog=tasks.filter(t=>t.st==='진행중').length;
    const core=tasks.filter(t=>t.co==='●').length;
    const avg=tasks.reduce((s,t)=>s+t.pp,0)/tasks.length;
    const zero=tasks.filter(t=>t.st==='진행중'&&t.pp===0).length;
    const channel=TM_CHANNEL[tm]||'기타';
    teamData[tm]={tasks,ch,done,prog,core,avg,zero,channel};
  });

  const fPct=v=>Math.round(v*100)+'%';
  const clr=v=>v>=0.7?'#059669':v>=0.4?'#d97706':'#dc2626';
  const bar=(v,w=120)=>`<div style="width:${w}px;height:10px;background:#e2e8f0;border-radius:5px;overflow:hidden;display:inline-block;vertical-align:middle;margin-left:6px"><div style="width:${Math.round(v*100)}%;height:100%;background:${clr(v)};border-radius:5px"></div></div>`;

  let h=`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>총괄별 과제 심층 분석</title>${RS}</head><body>
  <div class="cover">
    <div class="cover-tag">Confidential · 총괄별 과제 심층 분석 보고서</div>
    <div class="cover-title">총괄별 사업과제<br>심층 이행 분석</div>
    <div class="cover-sub">${D.period} 기준 | 경영기획팀 | 전체 ${T.length}개 과제</div>
    <div style="display:flex;gap:16px;margin-top:20px;flex-wrap:wrap">
      ${chs.map(c=>{const d=chData[c];return `<div style="flex:1;min-width:200px;background:rgba(255,255,255,.08);border-radius:12px;padding:16px">
        <div style="font-size:12px;color:rgba(255,255,255,.6);font-weight:600;margin-bottom:4px">${c}</div>
        <div style="font-size:32px;font-weight:900;color:#fff">${d.tasks.length}개</div>
        <div style="font-size:12px;color:rgba(255,255,255,.6)">${d.teams.length}개 팀 | 핵심 ${d.core.length}개 | 평균추진도 ${fPct(d.avgPp)}</div>
      </div>`}).join('')}
    </div>
  </div>`;

  // 경총/영총 각각 상세 분석
  chs.forEach(c=>{
    const d=chData[c];
    const cTeams=allTeams.filter(t=>teamData[t].ch===c).sort((a,b)=>teamData[b].avg-teamData[a].avg);

    h+=`<h2>${c} — ${d.tasks.length}개 과제 ${c==='경총'?'(경영·지원 기능)':'(영업·현장 기능)'}</h2>`;

    // Executive Summary Box
    h+=`<div class="exec-box">
      총 <b>${d.tasks.length}개</b> 과제 중 완료 <b style="color:#059669">${d.done.length}개</b>(${(d.done.length/d.tasks.length*100).toFixed(0)}%) · 진행중 <b style="color:#d97706">${d.prog.length}개</b> · 계획 ${d.plan.length}개 · 핵심과제 추진도 평균 <b>${fPct(d.coreAvgPp)}</b>
    </div>`;

    if(d.zero.length>0||d.low.length>0){
      h+=`<div class="warn">⚠️ 리스크 과제: 진행중 추진도 0% ${d.zero.length}개 · 15% 이하 ${d.low.length}개 → 즉시 점검 필요</div>`;
    }

    // 팀별 현황 테이블
    h+=`<h3>팀별 과제 현황</h3><table><thead><tr><th>팀명</th><th>채널연계</th><th class="num">과제</th><th class="num">완료</th><th class="num">핵심</th><th class="num">평균추진도</th><th class="num">⚠️리스크</th></tr></thead><tbody>`;
    cTeams.forEach(tm=>{
      const d2=teamData[tm];
      h+=`<tr><td><b>${tm}</b></td><td><span style="font-size:10px;padding:2px 8px;border-radius:10px;background:#dbeafe;color:#1e40af">${d2.channel}</span></td>
      <td class="num">${d2.tasks.length}</td><td class="num">${d2.done}</td><td class="num">${d2.core?`<span style="color:#dc2626;font-weight:700">${d2.core}</span>`:'—'}</td>
      <td class="num"><span style="font-weight:700;color:${clr(d2.avg)}">${fPct(d2.avg)}</span>${bar(d2.avg)}</td>
      <td class="num">${d2.zero?`<span style="color:#dc2626;font-weight:700">${d2.zero}개</span>`:'—'}</td></tr>`;
    });
    h+=`</tbody></table>`;

    // 핵심 과제 현황
    const coreList=d.core.sort((a,b)=>a.pp-b.pp);
    h+=`<h3>핵심 과제 추진 현황 (${coreList.length}개)</h3><table><thead><tr><th>팀</th><th>과제명</th><th>담당자</th><th class="num">추진도</th><th>상태</th></tr></thead><tbody>`;
    coreList.forEach(t=>{
      h+=`<tr><td>${t.tm}</td><td>${t.nm}</td><td style="color:#64748b;font-size:11px">${t.mg}</td>
      <td class="num"><b style="color:${clr(t.pp)}">${fPct(t.pp)}</b>${bar(t.pp,80)}</td>
      <td><span class="pill pill-${t.st==='완료'?'g':t.st==='진행중'?'y':'b'}">${t.st}</span></td></tr>`;
    });
    h+=`</tbody></table>`;

    // 리스크 과제
    const risk=[...d.zero,...d.low.filter(t=>!d.zero.includes(t))];
    if(risk.length){
      h+=`<h3>⚠️ 리스크 과제 (진행중·추진도≤15%)</h3><table><thead><tr><th>팀</th><th>과제명</th><th>담당자</th><th class="num">추진도</th></tr></thead><tbody>`;
      risk.forEach(t=>{h+=`<tr><td>${t.tm}</td><td>${t.nm}</td><td style="color:#64748b;font-size:11px">${t.mg}</td><td class="num"><b style="color:#dc2626">${fPct(t.pp)}</b></td></tr>`;});
      h+=`</tbody></table>`;
    }
  });

  // 전사 액션 권고
  const totalZero=T.filter(t=>t.st==='진행중'&&t.pp===0);
  const totalLow=T.filter(t=>t.st==='진행중'&&t.pp>0&&t.pp<=0.15);
  h+=`<h2>전사 액션 권고</h2>
  <div class="action">
    <b>① [즉시] 추진도 0% 과제 ${totalZero.length}개 집중 점검</b><br>
    경총 ${totalZero.filter(t=>t.ch==='경총').length}개 / 영총 ${totalZero.filter(t=>t.ch==='영총').length}개 — 2주 내 장애요인 파악 및 실행계획 재수립<br><br>
    <b>② [단기] 핵심과제 주간 이행 체크 시스템 도입</b><br>
    핵심 과제 ${T.filter(t=>t.co==='●').length}개 별도 주간 점검 보고 체계 운영<br><br>
    <b>③ [정기] 팀별 추진도 하위 3팀 집중 지원</b><br>
    ${allTeams.sort((a,b)=>teamData[a].avg-teamData[b].avg).slice(0,3).map(t=>`${t}(${fPct(teamData[t].avg)})`).join(', ')} — 상위 팀 노하우 이식 및 리소스 지원
  </div>
  </div></body></html>`;
  openReport(h,'총괄별 과제 심층 분석 보고서');
}

// ==================== 채널별 과제 연계 분석 보고서 ====================
function genTaskChannelReport(){
  const T=D.tasks, P=D.profit, K=D.kpi;

  // 팀→채널 매핑 (채널 영업과 연계)
  const TM_CHANNEL={
    '소매강화팀':{ch:'소매',kpi:'rt',color:'#38bdf8'},
    '디지털강화팀':{ch:'디지털',kpi:null,color:'#a78bfa'},
    '도매강화팀':{ch:'도매',kpi:'wh',color:'#34d399'},
    '소상공인혁신팀':{ch:'소상공인',kpi:'sm',color:'#818cf8'},
    '법인영업팀':{ch:'기업/공공',kpi:null,color:'#fb923c'},
    '기업영업팀':{ch:'기업/공공',kpi:null,color:'#fb923c'},
    '현장지원팀':{ch:'영업지원',kpi:null,color:'#fbbf24'},
    '영업기획팀':{ch:'영업지원',kpi:null,color:'#fbbf24'},
    '영업품질팀':{ch:'영업품질',kpi:null,color:'#f87171'},
    '역량강화팀':{ch:'역량강화',kpi:null,color:'#e879f9'},
    '경영기획팀':{ch:'경영지원',kpi:null,color:'#94a3b8'},
    '회계세무팀':{ch:'경영지원',kpi:null,color:'#94a3b8'},
    '인사팀':{ch:'경영지원',kpi:null,color:'#94a3b8'},
    '경영지원팀':{ch:'경영지원',kpi:null,color:'#94a3b8'},
    'IT보안팀':{ch:'IT/보안',kpi:null,color:'#38bdf8'},
    'AI확산팀':{ch:'AI/혁신',kpi:null,color:'#7c3aed'},
    '플랫폼사업1팀':{ch:'플랫폼',kpi:null,color:'#0891b2'},
    '플랫폼사업2팀':{ch:'플랫폼',kpi:null,color:'#0891b2'},
    '모바일파트':{ch:'플랫폼',kpi:null,color:'#0891b2'}
  };

  // 채널별 집계
  const channelMap={};
  T.forEach(t=>{
    const info=TM_CHANNEL[t.tm]||{ch:'기타',color:'#94a3b8'};
    const c=info.ch;
    if(!channelMap[c]) channelMap[c]={label:c,color:info.color,kpiKey:info.kpi,tasks:[],teams:new Set()};
    channelMap[c].tasks.push(t);
    channelMap[c].teams.add(t.tm);
  });

  const channels=Object.values(channelMap).sort((a,b)=>b.tasks.length-a.tasks.length);

  const fPct=v=>Math.round(v*100)+'%';
  const clr=v=>v>=0.7?'#059669':v>=0.4?'#d97706':'#dc2626';
  const bar=(v,w=100,c)=>`<div style="width:${w}px;height:8px;background:#e2e8f0;border-radius:4px;overflow:hidden;display:inline-block;vertical-align:middle;margin-left:4px"><div style="width:${Math.round(v*100)}%;height:100%;background:${c||clr(v)};border-radius:4px"></div></div>`;
  const RS=rptStyle();

  let h=`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>채널별 과제 연계 분석</title>${RS}</head><body>
  <div class="cover">
    <div class="cover-tag">Confidential · 채널별 과제×성과 연계 분석</div>
    <div class="cover-title">채널별 사업과제<br>성과 연계 분석 보고서</div>
    <div class="cover-sub">${D.period} 기준 | 경영기획팀 | 과제×KPI×손익 통합 분석</div>
  </div>

  <div class="exec-box">
    본 보고서는 팀별 사업과제를 비즈니스 채널에 매핑하여, <b>과제 추진도 ↔ KPI ↔ 손익</b>의 3자 연계를 분석합니다.<br>
    채널 전략과 실행 과제의 정합성을 검증하고, 과제 추진 지체가 실적에 미치는 영향을 진단합니다.
  </div>`;

  // 채널별 요약 테이블
  h+=`<h2>1. 채널별 과제 현황 종합</h2><table><thead>
    <tr><th>채널</th><th class="num">과제</th><th class="num">완료</th><th class="num">진행중</th><th class="num">리스크</th><th class="num">평균추진도</th><th>연계KPI</th><th>영업이익(백만)</th></tr>
  </thead><tbody>`;

  channels.forEach(ch=>{
    const done=ch.tasks.filter(t=>t.st==='완료').length;
    const prog=ch.tasks.filter(t=>t.st==='진행중').length;
    const risk=ch.tasks.filter(t=>t.st==='진행중'&&t.pp<=0.15).length;
    const avg=ch.tasks.reduce((s,t)=>s+t.pp,0)/ch.tasks.length;
    const kpiVal=ch.kpiKey ? (() => {const avgs=K.map(k=>k[ch.kpiKey]?.t||0);return (avgs.reduce((a,b)=>a+b,0)/avgs.length).toFixed(1)+'점';})() : '—';
    const profitKey={소매:'retail',도매:'wholesale',디지털:'digital',소상공인:'small_biz','기업/공공':'enterprise'}[ch.label];
    const profitVal=profitKey ? `${Math.round(P.op[profitKey]/1e6).toLocaleString()}` : '—';
    const profitClr=profitKey && P.op[profitKey]>0?'#059669':'#dc2626';
    h+=`<tr>
      <td><span style="display:inline-block;width:10px;height:10px;background:${ch.color};border-radius:50%;margin-right:6px"></span><b>${ch.label}</b><br><span style="font-size:10px;color:#64748b">${[...ch.teams].join(', ')}</span></td>
      <td class="num">${ch.tasks.length}</td><td class="num">${done}</td><td class="num">${prog}</td>
      <td class="num">${risk?`<span style="color:#dc2626;font-weight:700">${risk}⚠️</span>`:'—'}</td>
      <td class="num"><b style="color:${clr(avg)}">${fPct(avg)}</b>${bar(avg,80,ch.color)}</td>
      <td class="num">${kpiVal}</td>
      <td class="num" style="font-weight:700;color:${profitClr}">${profitVal!=='—'?profitVal+'M':'—'}</td>
    </tr>`;
  });
  h+=`</tbody></table>`;

  // KPI 연계 채널 심층 분석 (소매/도매/소상공인)
  const kpiChannels=[
    {label:'소매',taskCh:'소매',kpiKey:'rt',profitKey:'retail',color:'#38bdf8',tmName:'소매강화팀'},
    {label:'도매',taskCh:'도매',kpiKey:'wh',profitKey:'wholesale',color:'#34d399',tmName:'도매강화팀'},
    {label:'소상공인',taskCh:'소상공인',kpiKey:'sm',profitKey:'small_biz',color:'#818cf8',tmName:'소상공인혁신팀'}
  ];

  h+=`<h2>2. 핵심 영업 채널 — 과제×KPI×손익 상관 분석</h2>`;

  kpiChannels.forEach(kc=>{
    const chTasks=channelMap[kc.taskCh]?.tasks||[];
    const done=chTasks.filter(t=>t.st==='완료').length;
    const avg=chTasks.length?chTasks.reduce((s,t)=>s+t.pp,0)/chTasks.length:0;
    const kpiAvg=K.map(k=>k[kc.kpiKey]?.t||0).reduce((a,b)=>a+b,0)/K.length;
    const kpiVar=Math.max(...K.map(k=>k[kc.kpiKey]?.t||0))-Math.min(...K.map(k=>k[kc.kpiKey]?.t||0));
    const op=P.op[kc.profitKey];
    const opm=(op/P.revenue[kc.profitKey]*100).toFixed(1);
    const risk=chTasks.filter(t=>t.st==='진행중'&&t.pp<=0.15);

    h+=`<h3 style="border-left:4px solid ${kc.color};padding-left:10px">${kc.label} 채널 — ${kc.tmName}</h3>
    <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;margin-bottom:16px">
      <div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:12px;text-align:center">
        <div style="font-size:10px;color:#64748b;font-weight:600;margin-bottom:4px">과제 평균 추진도</div>
        <div style="font-size:24px;font-weight:900;color:${clr(avg)}">${fPct(avg)}</div>
        <div style="font-size:10px;color:#64748b">${chTasks.length}개 중 완료 ${done}개</div>
      </div>
      <div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:12px;text-align:center">
        <div style="font-size:10px;color:#64748b;font-weight:600;margin-bottom:4px">KPI 평균 득점</div>
        <div style="font-size:24px;font-weight:900;color:${kc.color}">${kpiAvg.toFixed(1)}점</div>
        <div style="font-size:10px;color:#64748b">본부간 편차 ${kpiVar.toFixed(1)}p</div>
      </div>
      <div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:12px;text-align:center">
        <div style="font-size:10px;color:#64748b;font-weight:600;margin-bottom:4px">영업이익률</div>
        <div style="font-size:24px;font-weight:900;color:${op>0?'#059669':'#dc2626'}">${opm}%</div>
        <div style="font-size:10px;color:#64748b">${Math.round(op/1e8).toFixed(1)}억원</div>
      </div>
    </div>`;

    // 과제 추진도 vs KPI 인사이트
    const insight=avg>=0.6&&kpiAvg>=80?`과제 추진도(${fPct(avg)})와 KPI(${kpiAvg.toFixed(1)}점) 모두 양호 — 전략 실행력이 성과로 연결되고 있음`:
      avg<0.4&&kpiAvg<80?`⚠️ 과제 추진도(${fPct(avg)}) 저조 → KPI(${kpiAvg.toFixed(1)}점) 부진 상관 관계 — 과제 실행력 제고가 핵심 과제`:
      avg>=0.5&&kpiAvg<80?`과제는 추진중(${fPct(avg)})이나 KPI 성과 미흡(${kpiAvg.toFixed(1)}점) — 과제 내용과 KPI 지표 정합성 재검토 필요`:
      `과제 초기 단계(${fPct(avg)}) — 향후 KPI 영향 모니터링 필요`;

    h+=`<div class="insight">📌 ${kc.label} 채널 진단: ${insight}</div>`;

    // 채널 주요 과제 목록
    if(chTasks.length){
      h+=`<table><thead><tr><th>과제명</th><th>담당자</th><th class="num">추진도</th><th>상태</th></tr></thead><tbody>`;
      chTasks.sort((a,b)=>b.co.localeCompare(a.co)||b.pp-a.pp).forEach(t=>{
        h+=`<tr><td>${t.co?'<span style="color:#dc2626">● </span>':''}${t.nm}</td><td style="color:#64748b;font-size:11px">${t.mg}</td>
        <td class="num"><b style="color:${clr(t.pp)}">${fPct(t.pp)}</b></td>
        <td><span class="pill pill-${t.st==='완료'?'g':t.st==='진행중'?'y':'b'}">${t.st}</span></td></tr>`;
      });
      h+=`</tbody></table>`;
    }

    if(risk.length){
      h+=`<div class="warn">⚠️ ${kc.label} 채널 리스크: 추진도 15% 이하 과제 ${risk.length}개 — ${risk.map(t=>t.nm.slice(0,15)).join(', ')}</div>`;
    }
  });

  // 경영지원·IT·AI 채널 (비KPI 연계)
  h+=`<h2>3. 지원 기능 채널 — 실행 현황</h2>`;
  const supportChannels=['경영지원','AI/혁신','IT/보안','영업지원','영업품질','역량강화','플랫폼'];
  supportChannels.forEach(sc=>{
    const ch=channelMap[sc];
    if(!ch) return;
    const avg=ch.tasks.reduce((s,t)=>s+t.pp,0)/ch.tasks.length;
    const done=ch.tasks.filter(t=>t.st==='완료').length;
    const core=ch.tasks.filter(t=>t.co==='●');
    h+=`<div style="margin-bottom:12px;padding:12px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px">
        <b style="font-size:13px">${sc}</b>
        <span style="font-size:11px;color:#64748b">${ch.tasks.length}개 과제 · 핵심 ${core.length}개 · 완료 ${done}개</span>
      </div>
      <div style="display:flex;align-items:center;gap:8px">
        <div style="flex:1;height:12px;background:#e2e8f0;border-radius:6px;overflow:hidden">
          <div style="width:${Math.round(avg*100)}%;height:100%;background:${ch.color};border-radius:6px"></div>
        </div>
        <span style="font-size:13px;font-weight:700;color:${clr(avg)};width:40px;text-align:right">${fPct(avg)}</span>
      </div>
      <div style="font-size:10px;color:#64748b;margin-top:4px">${[...ch.teams].join(' · ')}</div>
    </div>`;
  });

  // 종합 권고
  const totalRisk=T.filter(t=>t.st==='진행중'&&t.pp<=0.15);
  const worstCh=channels.sort((a,b)=>(a.tasks.reduce((s,t)=>s+t.pp,0)/a.tasks.length)-(b.tasks.reduce((s,t)=>s+t.pp,0)/b.tasks.length))[0];
  h+=`<h2>4. 전략적 권고사항</h2>
  <div class="action">
    <b>① [최우선] ${worstCh.label} 채널 과제 집중 관리</b><br>
    평균 추진도 ${fPct(worstCh.tasks.reduce((s,t)=>s+t.pp,0)/worstCh.tasks.length)} — 채널 담당 임원 직접 점검 체계 구축<br><br>
    <b>② [즉시] 전체 리스크 과제 ${totalRisk.length}개 액션플랜 재수립</b><br>
    진행중 15% 이하 과제 담당자 1:1 면담, 2주 내 장애요인 해소 방안 확정<br><br>
    <b>③ [정기] 채널×KPI 연계 월간 이행 점검 체계 수립</b><br>
    소매/도매/소상공인 과제 추진도와 KPI 득점 상관관계 월별 모니터링 보고 체계화<br><br>
    <b>④ [중기] 과제 KPI 연계 평가 체계 도입</b><br>
    사업과제 목표와 본부 KPI 지표를 명시적으로 연동하여 과제 완료 시 KPI 득점 자동 반영 체계 설계
  </div>
  </div></body></html>`;
  openReport(h,'채널별 과제 연계 분석 보고서');
}


function genKpiVarianceReport(){
  const K=D.kpi,V=D.variance,RD=D.kpi_retail_detail||{},WD=D.kpi_wholesale_detail||{},SMD=D.kpi_smb_detail||{};
  const HQS=['강서본부','동부본부','서부본부','강북본부','강남본부'];
  const varRt=Math.max(...K.map(k=>k.rt.t))-Math.min(...K.map(k=>k.rt.t));
  const varWh=Math.max(...K.map(k=>k.wh.t))-Math.min(...K.map(k=>k.wh.t));
  const varSm=Math.max(...K.map(k=>k.sm.t))-Math.min(...K.map(k=>k.sm.t));
  const K1=K[0],K5=K[4];
  let h=`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>KPI 편차 원인 분석</title>${rptStyle()}</head><body>
  <div class="cover"><div class="cover-tag">Confidential · KPI Variance Analysis</div>
  <div class="cover-title">KPI 편차 원인 분석 보고서</div>
  <div class="cover-sub">${D.period} 기준 | 경영기획팀 | 5개 본부 × 3채널 교차 분석</div>
  <div class="cover-kpis">
    <div class="cover-kpi"><div class="cover-kpi-v" style="color:#f87171">${varSm.toFixed(1)}p</div><div class="cover-kpi-l">소상공인 최대편차</div></div>
    <div class="cover-kpi"><div class="cover-kpi-v" style="color:#fbbf24">${varWh.toFixed(1)}p</div><div class="cover-kpi-l">도매 편차</div></div>
    <div class="cover-kpi"><div class="cover-kpi-v" style="color:#34d399">${varRt.toFixed(1)}p</div><div class="cover-kpi-l">소매 편차</div></div>
  </div></div>

  <h1>Ⅰ. Executive Summary</h1>
  <div class="exec-box">
    <b>종합 진단</b>: ${D.period} 기준 5개 본부 KPI 편차는 소상공인(${varSm.toFixed(1)}p) > 도매(${varWh.toFixed(1)}p) > 소매(${varRt.toFixed(1)}p) 순으로 소상공인 채널의 본부간 격차가 구조적으로 가장 심각함.<br><br>
    1위 <b>${K1.hq}</b>(${K1.ts.toFixed(1)}점)와 5위 <b>${K5.hq}</b>(${K5.ts.toFixed(1)}점)의 격차는 <b>${(K1.ts-K5.ts).toFixed(1)}점</b>이며, 채널별로는 소상공인(${(K1.sm.t-K5.sm.t).toFixed(1)}p) > 도매(${(K1.wh.t-K5.wh.t).toFixed(1)}p) > 소매(${(K1.rt.t-K5.rt.t).toFixed(1)}p) 순.
  </div>
  <div class="warn">⚠️ 강남본부 소상공인(49.01점)은 1위 강서(71.95점) 대비 22.9p 격차 — 즉시 집중 지원 필요</div>
  <div class="warn">⚠️ 서부본부 도매(88.35점)는 전 항목 최하위 — 대리점 인프라 구조 점검 필요</div>

  <h1>Ⅱ. 채널별 편차 분석</h1>

  <h2>① 소매 채널 (가중치 60%)</h2>
  <table><thead><tr><th>본부</th><th class="num">종합</th><th class="num">후불</th><th class="num">MNP</th><th class="num">유선</th><th class="num">R-VOC</th><th class="num">TCSI</th><th class="num">매장이익</th></tr></thead><tbody>
  ${HQS.map(hq=>{const k=K.find(x=>x.hq===hq)||{};const d=RD[hq]||{};const isTop=k.rk===1,isBot=k.rk===5;
  return `<tr style="${isTop?'background:#e6fdf5':''}${isBot?'background:#fff7ed':''}">
    <td style="font-weight:700">${hq}${isTop?' 🥇':''}${isBot?' ⚠️':''}</td>
    <td class="num ${isTop?'pos':''}${isBot?'neg':''}" style="font-weight:700">${k.rt?.t?.toFixed(1)||'-'}</td>
    <td class="num">${d.hubul?.toFixed(1)||'-'}</td>
    <td class="num ${d.mnp>8?'pos':d.mnp<7?'neg':''}">${d.mnp?.toFixed(1)||'-'}</td>
    <td class="num">${d.wired?.toFixed(1)||'-'}</td>
    <td class="num ${d.rvoc>4?'pos':d.rvoc<2.5?'neg':''}">${d.rvoc?.toFixed(2)||'-'}</td>
    <td class="num">${d.tcsi?.toFixed(1)||'-'}</td>
    <td class="num">${d.store_profit?.toFixed(1)||'-'}</td>
  </tr>`}).join('')}
  </tbody></table>
  <div class="insight"><b>소매 편차 원인 분석</b><br>
  • MNP 격차(${(K.find(k=>k.hq==='동부본부')?.rt.s||63.8).toFixed(1)} vs ${(K.find(k=>k.hq==='강남본부')?.rt.s||62.4).toFixed(1)})는 번호이동 캠페인 실행력 차이에서 기인 — 동부 9.2점 vs 강남 6.9점(격차 2.3p)<br>
  • 소매 편차(${varRt.toFixed(1)}p)가 상대적으로 작은 것은 본사 정책·인센 통일성 덕분 → 개인별 역량 차이가 주 변수<br>
  • R-VOC 강서(2.31) vs 동부(4.19) — 고객불만 관리 우수 동부본부의 후불 실적 연계 확인 필요
  </div>

  <h2>② 도매 채널 (가중치 30%)</h2>
  <table><thead><tr><th>본부</th><th class="num">종합</th><th class="num">신규가입</th><th class="num">이익</th><th class="num">대리점수</th><th class="num">엣지비율%</th></tr></thead><tbody>
  ${HQS.map(hq=>{const k=K.find(x=>x.hq===hq)||{};const d=WD[hq]||{};const isBot=hq==='서부본부';
  return `<tr style="${isBot?'background:#fff7ed':''}">
    <td style="font-weight:700">${hq}${isBot?' ⚠️':''}</td>
    <td class="num ${k.wh?.t>94?'pos':k.wh?.t<90?'neg':''}" style="font-weight:700">${k.wh?.t?.toFixed(1)||'-'}</td>
    <td class="num">${k.wh?.s?.toFixed(1)||'-'}</td>
    <td class="num">${k.wh?.p?.toFixed(1)||'-'}</td>
    <td class="num">${d.dealer_cnt||'-'}</td>
    <td class="num ${d.edge_rt>85?'pos':d.edge_rt<80?'neg':''}">${d.edge_rt?.toFixed(1)||'-'}</td>
  </tr>`}).join('')}
  </tbody></table>
  <div class="insight"><b>도매 편차 원인 분석</b><br>
  • 서부본부 도매(88.35점) 최하위 — 대리점 수(121개) 최저 + 엣지 비율(78.4%) 최저가 복합 작용<br>
  • 강서·동부 도매 1위권(96점대) 유지 — 대리점 생산성과 엣지 P/G 이행률이 핵심 차별 요인<br>
  • 도매 이익은 강남(6.0점)·서부(6.0점)가 높음 → 신규 부진해도 이익 관리 우수 — 포트폴리오 전략의 차이
  </div>

  <h2>③ 소상공인 채널 (가중치 10% · 최대 편차 채널)</h2>
  <table><thead><tr><th>본부</th><th class="num">종합</th><th class="num">영업</th><th class="num">이익</th><th class="num">프랜차이즈수</th><th class="num">커버리지%</th></tr></thead><tbody>
  ${HQS.map(hq=>{const k=K.find(x=>x.hq===hq)||{};const d=SMD[hq]||{};const isTop=hq==='강서본부',isBot=hq==='강남본부';
  return `<tr style="${isTop?'background:#e6fdf5':''}${isBot?'background:#fff7ed':''}">
    <td style="font-weight:700">${hq}${isTop?' 🥇':''}${isBot?' ⚠️':''}</td>
    <td class="num ${k.sm?.t>65?'pos':k.sm?.t<55?'neg':''}" style="font-weight:700">${k.sm?.t?.toFixed(1)||'-'}</td>
    <td class="num">${k.sm?.s?.toFixed(1)||'-'}</td>
    <td class="num">${k.sm?.p?.toFixed(2)||'-'}</td>
    <td class="num">${d.franchise_cnt||'-'}</td>
    <td class="num ${d.coverage>70?'pos':d.coverage<60?'neg':''}">${d.coverage||'-'}</td>
  </tr>`}).join('')}
  </tbody></table>
  <div class="warn">⚠️ 강남본부 소상공인 영업(29.58점) — 커버리지 52%로 전국 최저. 구조적 영업 공백 상태</div>
  <div class="insight"><b>소상공인 편차 원인 분석</b><br>
  • 이익 지표(19.4~20.8p)는 모든 본부에서 유사 → 영업력 차이(29.6~52.5p)가 편차의 핵심<br>
  • 커버리지(52%→78%)와 영업 득점이 정비례 — 전담 영업인력 배치 및 프랜차이즈 계약 확대가 직접 해법<br>
  • 강서 강점: 소상공인 전담 인력 집중 + 프랜차이즈 허브 구축 → 타 본부 벤치마킹 필수
  </div>

  <h1>Ⅲ. 구조적 원인 진단</h1>
  <table><thead><tr><th>채널</th><th>1차 원인</th><th>2차 원인</th><th>영향도</th><th>위험등급</th></tr></thead><tbody>
  <tr><td>소상공인 영업</td><td>전담인력 부족·커버리지 격차</td><td>본부별 전략 우선순위 차이</td><td>KPI 22.9p 격차</td><td class="neg" style="font-weight:700">HIGH 🔴</td></tr>
  <tr><td>도매 서부본부</td><td>대리점 인프라 열위(121개)</td><td>엣지 이행 동기부여 부족</td><td>8.1p 도매 격차</td><td style="color:#f59e0b;font-weight:700">MED 🟡</td></tr>
  <tr><td>소매 MNP</td><td>번호이동 캠페인 실행 역량</td><td>개인별 역량 편차</td><td>2.3p 격차</td><td style="color:#f59e0b;font-weight:700">MED 🟡</td></tr>
  <tr><td>소매 R-VOC</td><td>고객 불만 처리 프로세스</td><td>매장 서비스 역량 차이</td><td>분석 필요</td><td style="color:#34d399;font-weight:700">LOW 🟢</td></tr>
  </tbody></table>

  <h1>Ⅳ. 전략적 실행 권고안</h1>
  <div class="action">
    <b>① [즉시 · 1개월] 강남본부 소상공인 긴급 지원 TF 구성</b><br>
    강서본부 소상공인 우수사례 이식 — 커버리지 52%→65% 목표(+3개월), 전담인력 증원 검토<br><br>
    <b>② [즉시 · 1개월] 서부본부 도매 구조 점검</b><br>
    대리점 121개 → 135개 확충 목표 수립, 엣지 P/G 이행률 78.4%→85% 개선 계획 수립<br><br>
    <b>③ [단기 · 2개월] 소매 MNP 역량 표준화</b><br>
    동부본부(9.2점) 캠페인 방법론 전사 공유, 하위 강남·강북 집중 코칭 실시<br><br>
    <b>④ [중기 · 3개월] 소상공인 온라인 플랫폼 연계 판매 확대</b><br>
    과제 #119(소상공인 온라인 플랫폼) 우선 추진, 강남·동부 시범 적용 후 전사 확대
  </div>
  </body></html>`;
  openReport(h,'KPI 편차 원인 분석 보고서');
}

// ==================== 본부별 KPI 액션플랜 ====================
function genKpiActionPlanReport(){
  const K=D.kpi,HC=D.hq_context||{},RD=D.kpi_retail_detail||{},WD=D.kpi_wholesale_detail||{},SMD=D.kpi_smb_detail||{};
  const HQ_ORDER=['강서본부','동부본부','서부본부','강북본부','강남본부'];

  // 본부별 약점 채널 자동 도출
  const getWeak=(k)=>{
    const ch=[{nm:'소매',v:k.rt.t},{nm:'도매',v:k.wh.t},{nm:'소상공인',v:k.sm.t}];
    return ch.sort((a,b)=>a.v-b.v)[0].nm;
  };
  const getStr=(k)=>{
    const ch=[{nm:'소매',v:k.rt.t},{nm:'도매',v:k.wh.t},{nm:'소상공인',v:k.sm.t}];
    return ch.sort((a,b)=>b.v-a.v)[0].nm;
  };

  // 본부별 타깃 점수 (현재+2~3점)
  const targets={'강서본부':86,'동부본부':85,'서부본부':83,'강북본부':83,'강남본부':83};

  // 본부별 맞춤 액션
  const actions={
    '강서본부':{
      m1:['소상공인 커버리지 52→65% (강남 TF 파견 지원)','R-VOC 관리 체계 재점검·매뉴얼 강화'],
      m3:['소상공인 프랜차이즈 허브 모델 타 본부 이식','도매 이익구조 현행 유지 + 신규 대리점 5개 추가'],
      monitor:['소상공인 영업 주간 점검(현 52.45점 → 목표 58점)','소매 R-VOC 월별 트렌드']
    },
    '동부본부':{
      m1:['소상공인 전담인력 2명 추가 배치','동부 MNP 방법론 전사 공유 세션 개최'],
      m3:['소상공인 커버리지 65%→72% 단계 확대','소매 후불 강점 유지: 17.8점 유지 캠페인'],
      monitor:['소상공인 40.55→50점 월별 추이','도매 엣지 이행률 87.1% 유지']
    },
    '서부본부':{
      m1:['도매 대리점 긴급 확충 계획 수립(121→130개)','엣지 P/G 이행률 78.4%→83% 단기 목표 설정'],
      m3:['도매 전담 TF 운영(동부·강서 벤치마킹)','소상공인 강점(이익 20.78) 활용한 영업 확대'],
      monitor:['도매 88.35→92점 분기별 점검','대리점별 엣지 이행률 주간 모니터링']
    },
    '강북본부':{
      m1:['MNP 취약점(7.2점) 집중 캠페인 — 동부 방법론 적용','소상공인 커버리지 63%→68% 확대 계획'],
      m3:['매장 이익(15.58) 강점 활용: 유선 교차판매 강화','MNP 7.2→8.0점 목표 특화 트레이닝'],
      monitor:['소매 77.46→79점 분기 목표 관리','소상공인 영업 40.23→45점 추이']
    },
    '강남본부':{
      m1:['소상공인 긴급 커버리지 확대 TF 구성(52%→60%)','소매 MNP 최하위(6.9) 집중 코칭·캠페인 실시'],
      m3:['소상공인 온라인 플랫폼(과제#119) 강남 우선 시범','도매 이익(6.0점) 강점 유지하며 신규 비중 확대'],
      monitor:['소상공인 29.58→38점 月 점검(최우선)','소매 MNP 6.9→7.5점 목표 추이']
    }
  };

  let h=`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>본부별 KPI 액션플랜</title>${rptStyle()}</head><body>
  <div class="cover">
    <div class="cover-tag">Confidential · HQ Action Plan</div>
    <div class="cover-title">본부별 KPI 개선 액션플랜</div>
    <div class="cover-sub">${D.period} 기준 | 경영기획팀 | 5개 본부 맞춤 실행 과제</div>
    <div class="cover-kpis">
      ${K.slice(0,5).map(k=>`<div class="cover-kpi"><div class="cover-kpi-v" style="color:${k.rk===1?'#34d399':k.rk<=3?'#38bdf8':'#94a3b8'}">${k.ts.toFixed(1)}</div><div class="cover-kpi-l">#${k.rk} ${k.hq}</div></div>`).join('')}
    </div>
  </div>

  <h1>Ⅰ. 종합 현황 및 목표</h1>
  <table><thead><tr><th>순위</th><th>본부</th><th class="num">현재 점수</th><th class="num">목표 점수</th><th class="num">개선 폭</th><th>강점 채널</th><th>약점 채널</th><th>우선 과제</th></tr></thead><tbody>`;

  K.forEach(k=>{
    const ctx=HC[k.hq]||{};const act=actions[k.hq]||{};
    const tgt=targets[k.hq]||82;
    h+=`<tr>
      <td style="text-align:center;font-weight:800;color:${k.rk===1?'#fbbf24':k.rk===2?'#94a3b8':k.rk===3?'#d97706':'#64748b'}">${k.rk}</td>
      <td style="font-weight:700">${k.hq}</td>
      <td class="num" style="font-weight:700">${k.ts.toFixed(1)}</td>
      <td class="num pos" style="font-weight:700">${tgt}</td>
      <td class="num pos">+${(tgt-k.ts).toFixed(1)}p</td>
      <td>${getStr(k)}</td>
      <td class="neg">${getWeak(k)}</td>
      <td style="font-size:11px">${(act.m1||[])[0]?.slice(0,20)||'-'}...</td>
    </tr>`;
  });
  h+=`</tbody></table>
  <div class="insight">5개 본부 평균 현재 ${(K.reduce((s,k)=>s+k.ts,0)/5).toFixed(1)}점 → 목표 평균 ${(Object.values(targets).reduce((a,b)=>a+b,0)/5).toFixed(0)}점으로 전체적 수준 향상 필요. 소상공인 채널 집중 관리가 핵심 레버.</div>

  <h1>Ⅱ. 본부별 맞춤 액션플랜</h1>`;

  HQ_ORDER.forEach(hqNm=>{
    const k=K.find(x=>x.hq===hqNm)||{};const ctx=HC[hqNm]||{};const act=actions[hqNm]||{};
    const tgt=targets[hqNm]||82;
    const rd=RD[hqNm]||{},wd=WD[hqNm]||{},sd=SMD[hqNm]||{};
    const borderClr=k.rk===1?'#fbbf24':k.rk===2?'#94a3b8':k.rk===3?'#d97706':k.rk===4?'#38bdf8':'#f87171';

    h+=`<h2 style="border-left:5px solid ${borderClr};padding-left:10px">${hqNm} (현재 #${k.rk} · ${k.ts?.toFixed(1)}점 → 목표 ${tgt}점)</h2>
    <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin:10px 0">
      <div style="padding:10px;background:#f1f5f9;border-radius:6px"><div style="font-size:10px;color:#64748b;font-weight:600;text-transform:uppercase;margin-bottom:4px">소매 (60%)</div><div style="font-size:20px;font-weight:800;color:#0369a1">${k.rt?.t?.toFixed(1)||'-'}</div><div style="font-size:10px;color:#94a3b8">영업 ${k.rt?.s?.toFixed(1)||'-'} / 이익 ${k.rt?.p?.toFixed(1)||'-'}</div></div>
      <div style="padding:10px;background:#f1f5f9;border-radius:6px"><div style="font-size:10px;color:#64748b;font-weight:600;text-transform:uppercase;margin-bottom:4px">도매 (30%)</div><div style="font-size:20px;font-weight:800;color:#059669">${k.wh?.t?.toFixed(1)||'-'}</div><div style="font-size:10px;color:#94a3b8">영업 ${k.wh?.s?.toFixed(1)||'-'} / 이익 ${k.wh?.p?.toFixed(1)||'-'}</div></div>
      <div style="padding:10px;background:#f1f5f9;border-radius:6px"><div style="font-size:10px;color:#64748b;font-weight:600;text-transform:uppercase;margin-bottom:4px">소상공인 (10%)</div><div style="font-size:20px;font-weight:800;color:#7c3aed">${k.sm?.t?.toFixed(1)||'-'}</div><div style="font-size:10px;color:#94a3b8">영업 ${k.sm?.s?.toFixed(1)||'-'} / 이익 ${k.sm?.p?.toFixed(2)||'-'}</div></div>
    </div>
    <div class="exec-box" style="margin-bottom:8px">
      <b>강점</b>: ${ctx.str||'-'}<br>
      <b>약점</b>: ${ctx.wk||'-'}<br>
      <b>핵심 기회</b>: ${ctx.opp||'-'}
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:16px">
      <div class="action" style="margin:0">
        <div style="font-size:11px;font-weight:700;color:#059669;margin-bottom:6px">⚡ 즉시 실행 (1개월)</div>
        ${(act.m1||[]).map(a=>`• ${a}`).join('<br>')}
      </div>
      <div class="action" style="margin:0;background:rgba(15,23,42,.65);border-left-color:#3b82f6">
        <div style="font-size:11px;font-weight:700;color:#2563eb;margin-bottom:6px">📅 단기 과제 (3개월)</div>
        ${(act.m3||[]).map(a=>`• ${a}`).join('<br>')}
      </div>
    </div>
    <div style="background:#fefce8;border-left:3px solid #ca8a04;padding:10px 12px;border-radius:0 6px 6px 0;font-size:12px;margin-bottom:24px">
      <b>📊 모니터링 지표</b>: ${(act.monitor||[]).join(' / ')}
    </div>`;
  });

  h+=`<h1>Ⅲ. 전사 편차 축소 로드맵</h1>
  <table><thead><tr><th>기간</th><th>핵심 과제</th><th>대상</th><th>목표</th></tr></thead><tbody>
  <tr><td><b>1월 내</b></td><td>강남 소상공인 TF 발족 + 강서 벤치마킹</td><td>강남·경기본부</td><td>커버리지 52→58%</td></tr>
  <tr><td><b>2월</b></td><td>서부 도매 대리점 확충 계획 확정</td><td>서부본부·도매강화팀</td><td>도매 88→91점</td></tr>
  <tr><td><b>2월</b></td><td>전사 MNP 역량 표준화 교육</td><td>역량강화팀</td><td>하위 본부 +0.5p</td></tr>
  <tr><td><b>3월</b></td><td>소상공인 온라인플랫폼 강남 시범 적용</td><td>소상공인혁신팀</td><td>강남 소상공 35→40점</td></tr>
  <tr><td><b>1Q 마감</b></td><td>KPI 편차 현황 재검토 · 2Q 계획 수립</td><td>경영기획팀</td><td>전사 편차 10% 축소</td></tr>
  </tbody></table>
  </body></html>`;
  openReport(h,'본부별 KPI 액션플랜 보고서');
}

// ==================== KPI 기존 보고서 ====================
function genKpiReport(type){
  const K=D.kpi, V=D.variance, RD=D.kpi_retail_detail||{}, WD=D.kpi_wholesale_detail||{}, SMD=D.kpi_smb_detail||{}, HC=D.hq_context||{};
  const HQS=['강서본부','동부본부','서부본부','강북본부','강남본부'];
  const sorted=[...K].sort((a,b)=>a.rk-b.rk);

  // 공통 스타일 (보고서 전용 확장)
  const extStyle=`
  <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <style>
    *{margin:0;padding:0;box-sizing:border-box}
    body{font-family:'Noto Sans KR',sans-serif;background:#f8fafc;color:#1e293b;font-size:13px;line-height:1.7}
    .cover{background:linear-gradient(135deg,#0f172a 0%,#1e293b 50%,#0c1a2e 100%);color:#fff;padding:48px 40px 40px;position:relative;overflow:hidden}
    .cover::before{content:'';position:absolute;inset:0;background:radial-gradient(ellipse 80% 60% at 70% 40%,rgba(56,189,248,.12),transparent)}
    .cover-tag{font-size:10px;font-weight:600;letter-spacing:2px;color:rgba(56,189,248,.8);text-transform:uppercase;margin-bottom:16px}
    .cover-title{font-size:28px;font-weight:900;letter-spacing:-1px;margin-bottom:8px}
    .cover-sub{font-size:13px;color:rgba(148,163,184,.8);margin-bottom:28px}
    .cover-chips{display:flex;gap:10px;flex-wrap:wrap}
    .chip{padding:5px 14px;border-radius:20px;font-size:11px;font-weight:700;border:1px solid}
    .chip-b{background:rgba(56,189,248,.15);color:#7dd3fc;border-color:rgba(56,189,248,.3)}
    .chip-g{background:rgba(52,211,153,.15);color:#6ee7b7;border-color:rgba(52,211,153,.3)}
    .chip-y{background:rgba(251,191,36,.15);color:#fde68a;border-color:rgba(251,191,36,.3)}
    .chip-r{background:rgba(248,113,113,.15);color:#fca5a5;border-color:rgba(248,113,113,.3)}
    .page{max-width:960px;margin:0 auto;padding:32px 40px 60px}
    h1{font-size:20px;font-weight:800;color:#0f172a;padding-bottom:10px;border-bottom:3px solid #0ea5e9;margin-bottom:20px;letter-spacing:-.5px}
    h2{font-size:15px;font-weight:700;color:#0369a1;margin:28px 0 12px;padding-left:10px;border-left:4px solid #0ea5e9;display:flex;align-items:center;gap:8px}
    h3{font-size:13px;font-weight:700;color:#334155;margin:16px 0 8px}
    table{width:100%;border-collapse:collapse;margin:8px 0 20px;font-size:12px;background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.06)}
    th{background:#f1f5f9;padding:9px 12px;text-align:center;font-weight:700;color:#475569;font-size:11px;border-bottom:2px solid #e2e8f0;white-space:nowrap}
    th:first-child{text-align:left}
    td{padding:9px 12px;border-bottom:1px solid #f1f5f9;text-align:center;font-size:12px}
    td:first-child{text-align:left;font-weight:600}
    tr:last-child td{border-bottom:none}
    tr:hover{background:#f8faff}
    .pos{color:#059669;font-weight:700}.neg{color:#dc2626;font-weight:700}.neu{color:#64748b}
    .best{background:rgba(15,23,42,.65) !important}.worst{background:#fff5f5 !important}
    .exec{background:linear-gradient(135deg,rgba(15,23,42,.65),#f0f9ff);border:1px solid #bae6fd;border-radius:10px;padding:16px 20px;margin:0 0 20px;font-size:13px;line-height:1.8}
    .exec strong{color:#0369a1}
    .warn{background:#fff7ed;border:1px solid #fed7aa;border-radius:8px;padding:12px 16px;margin:12px 0;font-size:12px;color:#92400e}
    .warn strong{color:#c2410c}
    .good{background:rgba(15,23,42,.65);border:1px solid #bbf7d0;border-radius:8px;padding:12px 16px;margin:12px 0;font-size:12px;color:#14532d}
    .action{background:#f8fafc;border:1px solid #e2e8f0;border-left:4px solid #6366f1;border-radius:0 8px 8px 0;padding:14px 18px;margin:16px 0;font-size:12px;line-height:2}
    .action strong{color:#4f46e5}
    .bar-row{display:flex;align-items:center;gap:10px;margin:5px 0}
    .bar-lbl{width:80px;font-size:11px;text-align:right;color:#475569;flex-shrink:0;font-weight:500}
    .bar-bg{flex:1;height:18px;background:#f1f5f9;border-radius:4px;overflow:hidden}
    .bar-fg{height:100%;border-radius:4px;display:flex;align-items:center;padding:0 7px;font-size:10px;font-weight:700;color:#fff;white-space:nowrap}
    .bar-val{width:52px;font-size:11px;font-family:monospace;font-weight:700;flex-shrink:0;text-align:right}
    .grid2{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px}
    .grid3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;margin-bottom:20px}
    .card{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:16px;box-shadow:0 1px 4px rgba(0,0,0,.05)}
    .card-title{font-size:11px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:.5px;margin-bottom:8px}
    .card-val{font-size:24px;font-weight:900;letter-spacing:-1px;margin-bottom:2px}
    .card-sub{font-size:11px;color:#94a3b8}
    .rank-badge{display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:50%;font-size:11px;font-weight:800;margin-right:4px}
    .rk1{background:#fef3c7;color:#92400e}.rk2{background:#f1f5f9;color:#475569}.rk3{background:#fef0e6;color:#9a3412}
    .section-divider{height:1px;background:linear-gradient(90deg,#0ea5e9,transparent);margin:28px 0}
    .hq-card{background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:14px;border-top:3px solid}
    .score-big{font-size:36px;font-weight:900;letter-spacing:-2px;margin:6px 0}
    .sub-metrics{display:grid;grid-template-columns:repeat(3,1fr);gap:6px;margin-top:10px}
    .sub-m{text-align:center;padding:6px;background:#f8fafc;border-radius:6px}
    .sub-m-l{font-size:9px;color:#94a3b8;font-weight:600;text-transform:uppercase;letter-spacing:.3px}
    .sub-m-v{font-size:13px;font-weight:800;margin-top:2px}
    .tag{display:inline-block;padding:1px 7px;border-radius:10px;font-size:10px;font-weight:600;margin-left:4px}
    .tag-r{background:#fee2e2;color:#dc2626}.tag-g{background:#dcfce7;color:#16a34a}.tag-y{background:#fef9c3;color:#ca8a04}
    .variance-row{display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-bottom:1px solid #f1f5f9}
    .var-label{font-size:12px;font-weight:600;width:120px}
    .var-bar{flex:1;height:12px;background:#f1f5f9;border-radius:3px;overflow:hidden;margin:0 10px}
    .var-fg{height:100%;border-radius:3px}
    .var-val{font-size:12px;font-weight:700;font-family:monospace;width:50px;text-align:right}
    .var-level{font-size:10px;width:60px;text-align:right;color:#64748b}
    @media print{body{background:#fff}.cover{print-color-adjust:exact}}
  </style>`;

  // ─── 1. 종합 보고서 ─────────────────────────────────────────────
  if(type==='summary'){
    const top=sorted[0], bot=sorted[sorted.length-1];
    const gap=(top.ts-bot.ts).toFixed(1);
    const rtColors=['#fbbf24','#94a3b8','#d97706','#38bdf8','#a78bfa'];
    let h=`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>KPI 종합 분석 보고서</title>${extStyle}</head><body>
    <div class="cover">
      <div class="cover-tag">KPI Analysis · Confidential</div>
      <div class="cover-title">🎯 KPI 종합 분석 보고서</div>
      <div class="cover-sub">${D.period} 기준 · 소매 60% + 도매 30% + 소상공인 10% 가중합산</div>
      <div class="cover-chips">
        <div class="chip chip-y">🏆 1위 ${top.hq} ${top.ts.toFixed(1)}점</div>
        <div class="chip chip-b">📊 1~5위 격차 ${gap}점</div>
        <div class="chip chip-r">⚠️ 최대편차 소상공인 ${V.smb.toFixed(1)}p</div>
        <div class="chip chip-g">✅ 최소편차 소매이익 ${V.retail_profit.toFixed(2)}p</div>
      </div>
    </div>
    <div class="page">
    <h1>Ⅰ. Executive Summary</h1>
    <div class="exec">
      <strong>${top.hq}</strong>이 종합 ${top.ts.toFixed(1)}점으로 1위를 유지하며 도매(${top.wh.t.toFixed(1)}점)·소상공인(${top.sm.t.toFixed(1)}점) 모두 선두권.
      5위 <strong>${bot.hq}</strong>(${bot.ts.toFixed(1)}점)과의 격차는 <strong>${gap}점</strong>으로 소상공인 채널(편차 ${V.smb.toFixed(1)}p)이 순위 변동의 핵심 변수.
      소매 이익 지표 편차(${V.retail_profit.toFixed(2)}p)는 전 채널 최소 → 안정적 관리 확인.
      도매 영업 편차(${V.wh_sales.toFixed(1)}p) 및 소상공인 영업 편차(${V.smb_sales.toFixed(1)}p)는 집중 점검 필요.
    </div>

    <h2>📊 본부별 종합 순위</h2>`;
    // 본부 카드 5개
    h+=`<div class="grid3" style="grid-template-columns:repeat(5,1fr);gap:10px;margin-bottom:20px">`;
    sorted.forEach((k,i)=>{
      const HC_k=HC[k.hq]||{};
      h+=`<div class="hq-card" style="border-top-color:${rtColors[i]}">
        <div style="display:flex;justify-content:space-between;align-items:center">
          <div style="font-size:11px;font-weight:700;color:#475569">${k.hq}</div>
          <div class="rank-badge rk${k.rk<=3?k.rk:''}" style="${k.rk>3?'background:#f1f5f9;color:#94a3b8':''}">#${k.rk}</div>
        </div>
        <div class="score-big" style="color:${rtColors[i]}">${k.ts.toFixed(1)}</div>
        <div class="sub-metrics">
          <div class="sub-m"><div class="sub-m-l">소매60%</div><div class="sub-m-v" style="color:#38bdf8">${k.rt.t.toFixed(1)}</div></div>
          <div class="sub-m"><div class="sub-m-l">도매30%</div><div class="sub-m-v" style="color:#34d399">${k.wh.t.toFixed(1)}</div></div>
          <div class="sub-m"><div class="sub-m-l">소상10%</div><div class="sub-m-v" style="color:#a78bfa">${k.sm.t.toFixed(1)}</div></div>
        </div>
        ${HC_k.str?`<div style="font-size:9px;color:#059669;margin-top:8px;padding:4px 6px;background:rgba(15,23,42,.65);border-radius:4px">💪 ${HC_k.str}</div>`:''}
      </div>`;
    });
    h+=`</div>`;

    // 채널별 비교 테이블
    h+=`<h2>📋 채널별 득점 상세 비교</h2>
    <table><thead><tr><th>본부</th><th>종합</th><th>순위</th><th>소매 합계</th><th>소매영업</th><th>소매이익</th><th>도매 합계</th><th>도매영업</th><th>도매이익</th><th>소상공인</th><th>SMB영업</th><th>SMB이익</th></tr></thead><tbody>`;
    sorted.forEach((k,i)=>{
      h+=`<tr>
        <td>${k.hq}</td>
        <td style="font-weight:800;color:${rtColors[i]}">${k.ts.toFixed(1)}</td>
        <td><span style="font-weight:800;color:${rtColors[i]}">#${k.rk}</span></td>
        <td style="font-weight:700">${k.rt.t.toFixed(1)}</td>
        <td>${k.rt.s.toFixed(1)}</td><td>${k.rt.p.toFixed(1)}</td>
        <td style="font-weight:700">${k.wh.t.toFixed(1)}</td>
        <td>${k.wh.s.toFixed(1)}</td><td>${k.wh.p.toFixed(1)}</td>
        <td style="font-weight:700;color:${k.sm.t>=65?'#059669':k.sm.t<55?'#dc2626':'#d97706'}">${k.sm.t.toFixed(1)}</td>
        <td>${k.sm.s.toFixed(1)}</td><td>${k.sm.p.toFixed(1)}</td>
      </tr>`;
    });
    h+=`</tbody></table>`;

    // 채널별 바 차트
    h+=`<div class="grid2">
    <div class="card"><div class="card-title">📊 채널별 종합득점 시각화</div>`;
    sorted.forEach((k,i)=>{
      h+=`<div class="bar-row"><div class="bar-lbl">${k.hq.replace('본부','')}</div><div class="bar-bg"><div class="bar-fg" style="width:${k.ts}%;background:${rtColors[i]}">${k.ts.toFixed(1)}</div></div><div class="bar-val" style="color:${rtColors[i]}">#${k.rk}</div></div>`;
    });
    h+=`</div>
    <div class="card"><div class="card-title">📉 채널별 편차 현황</div>`;
    const varData=[
      {l:'소상공인 영업',v:V.smb_sales,c:'#dc2626'},
      {l:'소상공인 종합',v:V.smb,c:'#ef4444'},
      {l:'도매 영업',v:V.wh_sales,c:'#f59e0b'},
      {l:'도매 종합',v:V.wholesale,c:'#fbbf24'},
      {l:'종합 KPI',v:V.total,c:'#3b82f6'},
      {l:'소매 영업',v:V.retail_sales,c:'#34d399'},
      {l:'소매 종합',v:V.retail,c:'#6ee7b7'},
      {l:'소매 이익',v:V.retail_profit,c:'#a7f3d0'},
    ];
    const vmx=Math.max(...varData.map(x=>x.v));
    varData.forEach(x=>{
      h+=`<div class="variance-row"><div class="var-label">${x.l}</div><div class="var-bar"><div class="var-fg" style="width:${(x.v/vmx*100).toFixed(0)}%;background:${x.c}"></div></div><div class="var-val" style="color:${x.c}">${x.v.toFixed(1)}p</div><div class="var-level">${x.v>10?'⚠️ 매우큼':x.v>5?'⚡ 큼':x.v>2?'보통':'✅ 양호'}</div></div>`;
    });
    h+=`</div></div>`;

    // 전략적 시사점
    h+=`<h2>🔍 전략적 시사점 및 행동과제</h2>
    <div class="action">
      <strong>① 소상공인 채널 격차 해소 최우선</strong> — 강남본부(49.01점) vs 강서본부(71.95점) 22.9점 격차. 강서·강남 사이 노하우 전수 세션 즉시 기획 (담당: 소상공인혁신팀, 기한: 2월 내)<br>
      <strong>② 서부본부 도매 집중 지원</strong> — 도매 전 항목 최하위(88.35점), 특히 후불·MNP 취약. 도매강화팀 전담 코칭 배정 (기한: 1Q 내)<br>
      <strong>③ 강서본부 R-VOC 긴급 관리</strong> — 소매 R-VOC 0.86점(최저). 동부본부(4.19점) 벤치마킹 체계 수립 (기한: 즉시)<br>
      <strong>④ 강남본부 소상공인 온라인 플랫폼 활용</strong> — 커버리지 52% 최저. 소상공인 온라인 플랫폼 과제(No.119) 연계 추진
    </div>
    </div></body></html>`;
    return openReport(h,'KPI 종합 분석 보고서');
  }

  // ─── 2. 세부지표 심층분석 ────────────────────────────────────────
  if(type==='detail'){
    // 소매 세부 최고/최저 계산
    const rtVals=HQS.map(hq=>RD[hq]||{});
    const whVals=HQS.map(hq=>WD[hq]||{});
    const smbVals=HQS.map(hq=>SMD[hq]||{});
    const bestRt=HQS[rtVals.map(d=>d.total||0).indexOf(Math.max(...rtVals.map(d=>d.total||0)))];
    const worstRt=HQS[rtVals.map(d=>d.total||0).indexOf(Math.min(...rtVals.map(d=>d.total||0)))];
    const bestWh=HQS[whVals.map(d=>d.total||0).indexOf(Math.max(...whVals.map(d=>d.total||0)))];
    const worstWh=HQS[whVals.map(d=>d.total||0).indexOf(Math.min(...whVals.map(d=>d.total||0)))];
    const bestSmb=HQS[smbVals.map(d=>d.total||0).indexOf(Math.max(...smbVals.map(d=>d.total||0)))];
    const worstSmb=HQS[smbVals.map(d=>d.total||0).indexOf(Math.min(...smbVals.map(d=>d.total||0)))];

    const cellStyle=(v,good,bad)=>v>=good?'class="pos best"':v<=bad?'class="neg worst"':'class="neu"';

    let h=`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>KPI 세부지표 심층분석</title>${extStyle}</head><body>
    <div class="cover">
      <div class="cover-tag">KPI Deep Dive · Confidential</div>
      <div class="cover-title">📊 KPI 세부지표 심층분석</div>
      <div class="cover-sub">${D.period} 기준 · 소매 / 도매 / 소상공인 3채널 전체 분석</div>
      <div class="cover-chips">
        <div class="chip chip-b">소매 1위 ${bestRt.replace('본부','')} · 최하 ${worstRt.replace('본부','')}</div>
        <div class="chip chip-g">도매 1위 ${bestWh.replace('본부','')} · 최하 ${worstWh.replace('본부','')}</div>
        <div class="chip chip-y">소상공인 1위 ${bestSmb.replace('본부','')} · 최하 ${worstSmb.replace('본부','')}</div>
        <div class="chip chip-r">⚠️ 최대편차 소상공인 ${V.smb.toFixed(1)}p</div>
      </div>
    </div>
    <div class="page">
    <h1>Ⅰ. 소매 채널 세부지표 (가중치 60%)</h1>
    <div class="exec">
      소매 종합 격차 <strong>${V.retail.toFixed(2)}p</strong>로 안정적. 그러나 <strong>R-VOC</strong> 항목에서 강서본부(0.86) vs 동부본부(4.19) 간 <strong>3.33점 극단적 편차</strong> 발생 → 고객 접점 품질 집중 관리 필요.
      후불 지표는 서부본부(${RD['서부본부']?.hubul?.toFixed(1)||'?'}점) 선두, MNP는 동부본부(${RD['동부본부']?.mnp?.toFixed(1)||'?'}점) 선두.
    </div>
    <div class="warn">⚠️ <strong>강서본부 R-VOC 0.86점</strong> — 전체 최저, 기준점(3.0) 대비 현저히 낮음. 즉시 원인 분석 및 현장 조치 필요</div>
    <table><thead>
      <tr><th rowspan="2">본부</th><th rowspan="2">종합</th><th rowspan="2">순위</th><th colspan="2">영업지표</th><th colspan="2">유선</th><th colspan="2">품질지표</th><th colspan="2">이익지표</th></tr>
      <tr><th>후불가입</th><th>MNP</th><th>유선영업</th><th>MIT</th><th>R-VOC</th><th>TCSI</th><th>매장이익</th><th>영업이익</th></tr>
    </thead><tbody>`;
    const rtSorted=[...HQS].sort((a,b)=>(RD[b]?.total||0)-(RD[a]?.total||0));
    rtSorted.forEach((hq,i)=>{
      const d=RD[hq]||{};
      h+=`<tr ${i===0?'style="background:rgba(15,23,42,.65)"':i===rtSorted.length-1?'style="background:#fff5f5"':''}>
        <td>${hq}</td>
        <td style="font-weight:800;color:${i===0?'#059669':i===rtSorted.length-1?'#dc2626':'#1e293b'}">${d.total?.toFixed(1)||'-'}</td>
        <td style="text-align:center">${i+1}</td>
        <td ${cellStyle(d.hubul||0,14.5,13.8)}>${d.hubul?.toFixed(1)||'-'}</td>
        <td ${cellStyle(d.mnp||0,7.5,6.0)}>${d.mnp?.toFixed(1)||'-'}</td>
        <td ${cellStyle(d.wired||0,14,12)}>${d.wired?.toFixed(1)||'-'}</td>
        <td ${cellStyle(d.mit||0,8.5,8.0)}>${d.mit?.toFixed(1)||'-'}</td>
        <td ${cellStyle(d.rvoc||0,3,1.5)}>${d.rvoc?.toFixed(2)||'-'}</td>
        <td ${cellStyle(d.tcsi||0,5,3.5)}>${d.tcsi?.toFixed(1)||'-'}</td>
        <td ${cellStyle(d.store_profit||0,9.5,9.0)}>${d.store_profit?.toFixed(1)||'-'}</td>
        <td ${cellStyle(d.op_profit||0,6,5)}>${d.op_profit?.toFixed(1)||'-'}</td>
      </tr>`;
    });
    h+=`</tbody></table>
    <div style="font-size:10px;color:#94a3b8;margin-top:-12px;margin-bottom:16px">🟢 초록=상위권 / 🔴 빨강=하위권 · 기준: 후불≥14.5, MNP≥7.5, R-VOC≥3.0, TCSI≥5.0</div>

    <h2>📊 소매 세부지표 채널 비교 (시각화)</h2>
    <div class="grid2">`;
    // 소매 시각화 — 후불, MNP, R-VOC, TCSI
    [['후불가입(만점30)',rtSorted.map(hq=>({nm:hq,v:RD[hq]?.hubul||0})),30,'#38bdf8'],
     ['MNP(만점10)',rtSorted.map(hq=>({nm:hq,v:RD[hq]?.mnp||0})),10,'#34d399'],
     ['R-VOC(만점5)',rtSorted.map(hq=>({nm:hq,v:RD[hq]?.rvoc||0})),5,'#fb923c'],
     ['TCSI(만점6)',rtSorted.map(hq=>({nm:hq,v:RD[hq]?.tcsi||0})),6,'#a78bfa']
    ].forEach(([title,data,max,color])=>{
      h+=`<div class="card"><div class="card-title">${title}</div>`;
      data.forEach(x=>{
        const pct=(x.v/max*100).toFixed(0);
        h+=`<div class="bar-row"><div class="bar-lbl">${x.nm.replace('본부','')}</div><div class="bar-bg"><div class="bar-fg" style="width:${pct}%;background:${color}">${x.v.toFixed(2)}</div></div><div class="bar-val" style="color:${color}">${pct}%</div></div>`;
      });
      h+=`</div>`;
    });
    h+=`</div>

    <div class="section-divider"></div>
    <h1>Ⅱ. 도매 채널 세부지표 (가중치 30%)</h1>
    <div class="exec">
      도매 종합 편차 <strong>${V.wholesale.toFixed(1)}p</strong> — 서부본부(${WD['서부본부']?.total?.toFixed(1)||'?'}점)가 유일한 80점대로 <strong>전 항목 취약</strong>.
      후불·MNP 지표는 강서·동부·강남·강북 4개 본부가 만점권이나 서부만 크게 미달.
      도매 영업 편차 <strong>${V.wh_sales.toFixed(1)}p</strong>가 소매(${V.retail_sales.toFixed(1)}p)의 ${(V.wh_sales/V.retail_sales).toFixed(1)}배 → 도매 집중 관리 필요.
    </div>
    <div class="warn">⚠️ <strong>서부본부 도매 전 항목 최하위</strong> — 후불(22.94), MNP(9.74), 유선영업(14.85) 모두 타 본부 대비 현저히 낮음</div>
    <table><thead>
      <tr><th rowspan="2">본부</th><th rowspan="2">종합</th><th rowspan="2">순위</th><th colspan="3">영업지표</th><th rowspan="2">R-VOC</th><th rowspan="2">인프라<br>이익</th><th rowspan="2">영업이익</th></tr>
      <tr><th>후불가입</th><th>MNP</th><th>유선영업</th></tr>
    </thead><tbody>`;
    const whSorted=[...HQS].sort((a,b)=>(WD[b]?.total||0)-(WD[a]?.total||0));
    whSorted.forEach((hq,i)=>{
      const d=WD[hq]||{};
      h+=`<tr ${i===0?'style="background:rgba(15,23,42,.65)"':i===whSorted.length-1?'style="background:#fff5f5"':''}>
        <td>${hq}</td>
        <td style="font-weight:800;color:${i===0?'#059669':i===whSorted.length-1?'#dc2626':'#1e293b'}">${d.total?.toFixed(1)||'-'}</td>
        <td style="text-align:center">${i+1}</td>
        <td ${cellStyle(d.hubul||0,24,22)}>${d.hubul?.toFixed(1)||'-'}</td>
        <td ${cellStyle(d.mnp||0,12,10)}>${d.mnp?.toFixed(1)||'-'}</td>
        <td ${cellStyle(d.wired||0,18,15)}>${d.wired?.toFixed(1)||'-'}</td>
        <td ${cellStyle(d.rvoc||0,4,2.5)}>${d.rvoc?.toFixed(2)||'-'}</td>
        <td ${cellStyle(d.infra||0,5.5,5)}>${d.infra?.toFixed(2)||'-'}</td>
        <td ${cellStyle(d.op_profit||0,6,5)}>${d.op_profit?.toFixed(1)||'-'}</td>
      </tr>`;
    });
    h+=`</tbody></table>
    <div class="grid2">`;
    [['후불가입(만점24)',whSorted.map(hq=>({nm:hq,v:WD[hq]?.hubul||0})),24,'#34d399'],
     ['MNP(만점12)',whSorted.map(hq=>({nm:hq,v:WD[hq]?.mnp||0})),12,'#38bdf8'],
     ['유선영업(만점20)',whSorted.map(hq=>({nm:hq,v:WD[hq]?.wired||0})),20,'#fbbf24'],
     ['R-VOC(만점5)',whSorted.map(hq=>({nm:hq,v:WD[hq]?.rvoc||0})),5,'#fb923c'],
    ].forEach(([title,data,max,color])=>{
      h+=`<div class="card"><div class="card-title">${title}</div>`;
      data.forEach(x=>{
        const pct=(x.v/max*100).toFixed(0);
        h+=`<div class="bar-row"><div class="bar-lbl">${x.nm.replace('본부','')}</div><div class="bar-bg"><div class="bar-fg" style="width:${pct}%;background:${color}">${x.v.toFixed(2)}</div></div><div class="bar-val" style="color:${color}">${pct}%</div></div>`;
      });
      h+=`</div>`;
    });
    h+=`</div>

    <div class="section-divider"></div>
    <h1>Ⅲ. 소상공인 채널 세부지표 (가중치 10%)</h1>
    <div class="exec">
      소상공인 종합 편차 <strong>${V.smb.toFixed(1)}p</strong> — 전 채널 중 <strong>압도적 최대</strong>. 강서본부(${SMD['강서본부']?.total?.toFixed(1)||'?'}점) vs 강남본부(${SMD['강남본부']?.total?.toFixed(1)||'?'}점) 격차 ${((SMD['강서본부']?.total||0)-(SMD['강남본부']?.total||0)).toFixed(1)}점.
      커버리지 차이(강서 78% vs 강남 52%)가 영업 격차의 핵심 원인.
      이익 지표 편차(${V.smb_profit.toFixed(1)}p)는 상대적으로 낮아 수익성 자체는 유사 → <strong>외형 성장 여력 확보가 과제</strong>.
    </div>
    <div class="warn">⚠️ <strong>강남본부 소상공인 49.01점</strong> — 영업 득점 29.58점으로 전체 최저. 커버리지 52% 확대 즉시 필요</div>
    <table><thead>
      <tr><th>본부</th><th>종합</th><th>순위</th><th>영업득점</th><th>이익득점</th><th>커버리지</th><th>프랜차이즈수</th><th>평균매출(억)</th></tr>
    </thead><tbody>`;
    const smbSorted=[...HQS].sort((a,b)=>(SMD[b]?.total||0)-(SMD[a]?.total||0));
    smbSorted.forEach((hq,i)=>{
      const d=SMD[hq]||{};
      h+=`<tr ${i===0?'style="background:rgba(15,23,42,.65)"':i===smbSorted.length-1?'style="background:#fff5f5"':''}>
        <td>${hq}</td>
        <td style="font-weight:800;color:${i===0?'#059669':i===smbSorted.length-1?'#dc2626':'#1e293b'}">${d.total?.toFixed(1)||'-'}</td>
        <td style="text-align:center">${i+1}</td>
        <td ${cellStyle(d.sales||0,48,38)}>${d.sales?.toFixed(1)||'-'}</td>
        <td ${cellStyle(d.profit||0,20,18)}>${d.profit?.toFixed(1)||'-'}</td>
        <td ${cellStyle(d.coverage||0,75,60)}>${d.coverage||'-'}%</td>
        <td style="text-align:center">${d.franchise_cnt||'-'}</td>
        <td>${d.avg_revenue?.toFixed(1)||'-'}억</td>
      </tr>`;
    });
    h+=`</tbody></table>
    <div class="grid2">`;
    [['소상공인 영업득점',smbSorted.map(hq=>({nm:hq,v:SMD[hq]?.sales||0})),60,'#a78bfa'],
     ['커버리지(%)',smbSorted.map(hq=>({nm:hq,v:SMD[hq]?.coverage||0})),100,'#818cf8'],
    ].forEach(([title,data,max,color])=>{
      h+=`<div class="card"><div class="card-title">${title}</div>`;
      data.forEach(x=>{
        const pct=(x.v/max*100).toFixed(0);
        h+=`<div class="bar-row"><div class="bar-lbl">${x.nm.replace('본부','')}</div><div class="bar-bg"><div class="bar-fg" style="width:${pct}%;background:${color}">${typeof x.v==='number'&&x.v<10?x.v.toFixed(1)+'%':x.v.toFixed?x.v.toFixed(1):x.v}</div></div><div class="bar-val" style="color:${color}">${pct}%</div></div>`;
      });
      h+=`</div>`;
    });
    h+=`</div>

    <div class="section-divider"></div>
    <h1>Ⅳ. 채널 교차 분석 및 실행 권고</h1>
    <div class="grid3">`;
    // 본부별 강약점 카드
    const topColors=['#fbbf24','#94a3b8','#d97706','#38bdf8','#a78bfa'];
    sorted.forEach((k,i)=>{
      const hc=HC[k.hq]||{};
      h+=`<div class="card" style="border-top:3px solid ${topColors[i]}">
        <div style="font-weight:800;font-size:13px;margin-bottom:8px">${k.hq} <span style="color:${topColors[i]}">#${k.rk}</span></div>
        ${hc.str?`<div style="font-size:11px;padding:5px 8px;background:rgba(15,23,42,.65);border-radius:5px;margin-bottom:5px;color:#14532d">💪 <b>강점:</b> ${hc.str}</div>`:''}
        ${hc.wk?`<div style="font-size:11px;padding:5px 8px;background:#fff5f5;border-radius:5px;margin-bottom:5px;color:#7f1d1d">⚠️ <b>약점:</b> ${hc.wk}</div>`:''}
        ${hc.opp?`<div style="font-size:11px;padding:5px 8px;background:rgba(15,23,42,.65);border-radius:5px;color:#1e3a5f">🎯 <b>기회:</b> ${hc.opp}</div>`:''}
      </div>`;
    });
    h+=`</div>
    <div class="action">
      <strong>① [즉시] 강서본부 소매 R-VOC 긴급 개선</strong> — 0.86점은 기준치의 29%. 동부본부(4.19점) 현장 프로세스 즉시 벤치마킹<br>
      <strong>② [1월 내] 서부본부 도매 집중 지원 TF 구성</strong> — 후불(22.94) MNP(9.74) 전 항목 최하위. 도매강화팀 전담 코치 배정<br>
      <strong>③ [2월 내] 강남본부 소상공인 커버리지 52%→65% 목표 설정</strong> — 소상공인온라인플랫폼 과제(No.119) 연계, 프랜차이즈 38개→45개 확대<br>
      <strong>④ [1Q 내] 동부본부 소상공인 역량 강화</strong> — KPI 2위이나 소상공인 59.23점 취약. 강서본부 성공 모델 전수 프로그램 운영
    </div>
    </div></body></html>`;
    return openReport(h,'KPI 세부지표 심층분석');
  }

  // ─── 3. 지사별 분석 보고서 ──────────────────────────────────────
  if(type==='branch'){
    const BR=D.kpiBranch||{};
    const rtList=BR.retail||[];
    const whList=BR.wholesale||[];
    const rtMax=Math.max(...rtList.map(b=>b.sc));
    const rtMin=Math.min(...rtList.map(b=>b.sc));
    const whMax=Math.max(...whList.map(b=>b.sc));
    const whMin=Math.min(...whList.map(b=>b.sc));
    const rtAvg=rtList.length?(rtList.reduce((s,b)=>s+b.sc,0)/rtList.length):0;
    const whAvg=whList.length?(whList.reduce((s,b)=>s+b.sc,0)/whList.length):0;
    const rtGap=(rtMax-rtMin).toFixed(1);
    const whGap=(whMax-whMin).toFixed(1);

    let h=`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>지사별 KPI 분석 보고서</title>${extStyle}</head><body>
    <div class="cover">
      <div class="cover-tag">Branch KPI Analysis · Confidential</div>
      <div class="cover-title">🏬 지사별 KPI 분석 보고서</div>
      <div class="cover-sub">${D.period} 기준 · 소매 ${rtList.length}개 지사 + 도매 ${whList.length}개 지사 분석</div>
      <div class="cover-chips">
        <div class="chip chip-b">소매 1위 ${rtList[0]?.nm||'-'} ${rtList[0]?.sc.toFixed(1)||''}점</div>
        <div class="chip chip-g">도매 1위 ${whList[0]?.nm||'-'} ${whList[0]?.sc.toFixed(1)||''}점</div>
        <div class="chip chip-r">소매 지사간 격차 ${rtGap}점</div>
        <div class="chip chip-y">도매 지사간 격차 ${whGap}점</div>
      </div>
    </div>
    <div class="page">
    <h1>Ⅰ. 소매 지사 KPI 순위 분석</h1>
    <div class="exec">
      소매 지사 <strong>${rtList.length}개</strong> 중 1위 <strong>${rtList[0]?.nm}(${rtList[0]?.sc.toFixed(1)}점)</strong>, 최하위 <strong>${rtList[rtList.length-1]?.nm}(${rtList[rtList.length-1]?.sc.toFixed(1)}점)</strong>.
      지사 평균 <strong>${rtAvg.toFixed(1)}점</strong>, 상하위 격차 <strong>${rtGap}점</strong>.
      평균 이상 지사 <strong>${rtList.filter(b=>b.sc>=rtAvg).length}개</strong> / 평균 미만 <strong>${rtList.filter(b=>b.sc<rtAvg).length}개</strong>.
    </div>`;
    if(rtList.filter(b=>b.sc<rtAvg).length>0){
      h+=`<div class="warn">⚠️ 평균(${rtAvg.toFixed(1)}점) 미달 지사 ${rtList.filter(b=>b.sc<rtAvg).length}개: ${rtList.filter(b=>b.sc<rtAvg).map(b=>b.nm).join(', ')} → 집중 관리 필요</div>`;
    }
    // 소매 바 차트
    h+=`<div class="grid2">
    <div class="card" style="grid-column:1/-1"><div class="card-title">소매 지사별 득점 현황 (${rtList.length}개 지사)</div>`;
    rtList.forEach((b,i)=>{
      const pct=(b.sc/rtMax*95).toFixed(0);
      const isTop=i<3, isBot=i>=rtList.length-2;
      const col=isTop?'#059669':isBot?'#dc2626':b.sc>=rtAvg?'#3b82f6':'#94a3b8';
      h+=`<div class="bar-row">
        <div class="bar-lbl" style="font-weight:${isTop?700:400};color:${col}">${b.nm}</div>
        <div class="bar-bg">
          <div class="bar-fg" style="width:${pct}%;background:${col}">
            ${b.sc.toFixed(1)}
          </div>
        </div>
        <div class="bar-val" style="color:${col}">${isTop?'🥇🥈🥉'[i]:isBot?'⚠️':''} #${b.rk}</div>
      </div>`;
    });
    h+=`</div></div>`;
    // 소매 상세 테이블
    h+=`<h2>📋 소매 지사 순위 상세</h2>
    <table><thead><tr><th>순위</th><th>지사명</th><th>득점</th><th>평균 대비</th><th>최고 대비</th><th>수준</th></tr></thead><tbody>`;
    rtList.forEach((b,i)=>{
      const diff=(b.sc-rtAvg), diffTop=(b.sc-rtMax);
      const isTop=i<3, isBot=i>=rtList.length-2;
      h+=`<tr style="${isTop?'background:rgba(15,23,42,.65)':isBot?'background:#fff5f5':''}">
        <td style="text-align:center;font-weight:800;color:${isTop?'#059669':isBot?'#dc2626':'#1e293b'}">${b.rk}</td>
        <td style="font-weight:${isTop?700:500}">${b.nm}</td>
        <td style="font-weight:700;color:${isTop?'#059669':isBot?'#dc2626':'#1e293b'};text-align:right">${b.sc.toFixed(1)}</td>
        <td class="${diff>=0?'pos':'neg'}" style="text-align:right">${diff>=0?'+':''}${diff.toFixed(1)}</td>
        <td class="neg" style="text-align:right">${diffTop.toFixed(1)}</td>
        <td style="text-align:center">${b.sc>=rtAvg+3?'<span class="tag tag-g">우수</span>':b.sc<rtAvg-3?'<span class="tag tag-r">관리필요</span>':'<span class="tag tag-y">보통</span>'}</td>
      </tr>`;
    });
    h+=`</tbody></table>
    <div class="action">
      <strong>소매 지사 인사이트:</strong><br>
      🥇 ${rtList[0]?.nm}(${rtList[0]?.sc.toFixed(1)}점) — 1위 유지 동력 분석 후 전사 공유<br>
      📊 평균대(${rtAvg.toFixed(1)}점) 중위권 ${rtList.filter(b=>b.sc>=rtAvg-2&&b.sc<rtAvg+3).length}개 지사 → 소규모 점프 가능 집중 코칭 대상<br>
      ⚠️ ${rtList[rtList.length-1]?.nm}(${rtList[rtList.length-1]?.sc.toFixed(1)}점) — 최하위, 1위 대비 ${(rtMax-rtMin).toFixed(1)}점 격차. 역량강화팀 현장 지원 우선 배정
    </div>

    <div class="section-divider"></div>
    <h1>Ⅱ. 도매 지사 KPI 순위 분석</h1>
    <div class="exec">
      도매 지사 <strong>${whList.length}개</strong> 중 1위 <strong>${whList[0]?.nm}(${whList[0]?.sc.toFixed(1)}점)</strong>, 최하위 <strong>${whList[whList.length-1]?.nm}(${whList[whList.length-1]?.sc.toFixed(1)}점)</strong>.
      지사 평균 <strong>${whAvg.toFixed(1)}점</strong>으로 소매 평균(${rtAvg.toFixed(1)}점)보다 <strong>${(whAvg-rtAvg).toFixed(1)}점 높음</strong> → 도매 채널 전반적 성과 양호.
      ${whList.filter(b=>b.sc>100).length>0?`<strong>${whList.filter(b=>b.sc>100).map(b=>b.nm).join(', ')}</strong> 100점 초과 달성(목표 초과 달성 의미)!`:''}
    </div>`;
    const overAchieve=whList.filter(b=>b.sc>100);
    if(overAchieve.length>0){
      h+=`<div class="good">✅ <strong>목표 초과 달성 지사:</strong> ${overAchieve.map(b=>`${b.nm}(${b.sc.toFixed(1)}점)`).join(', ')} — 우수 사례 전파 대상</div>`;
    }
    h+=`<div class="grid2">
    <div class="card" style="grid-column:1/-1"><div class="card-title">도매 지사별 득점 현황 (${whList.length}개 지사)</div>`;
    whList.forEach((b,i)=>{
      const pct=Math.min(98,(b.sc/whMax*92)).toFixed(0);
      const isTop=i<3, isBot=i>=whList.length-2;
      const col=b.sc>100?'#7c3aed':isTop?'#059669':isBot?'#dc2626':b.sc>=whAvg?'#3b82f6':'#94a3b8';
      h+=`<div class="bar-row">
        <div class="bar-lbl" style="font-weight:${isTop?700:400};color:${col}">${b.nm}</div>
        <div class="bar-bg">
          <div class="bar-fg" style="width:${pct}%;background:${col}">
            ${b.sc.toFixed(1)}${b.sc>100?' 🌟':''}
          </div>
        </div>
        <div class="bar-val" style="color:${col}">#${b.rk}</div>
      </div>`;
    });
    h+=`</div></div>`;
    h+=`<h2>📋 도매 지사 순위 상세</h2>
    <table><thead><tr><th>순위</th><th>지사명</th><th>득점</th><th>평균 대비</th><th>최고 대비</th><th>수준</th></tr></thead><tbody>`;
    whList.forEach((b,i)=>{
      const diff=(b.sc-whAvg), diffTop=(b.sc-whMax);
      const isTop=i<3, isBot=i>=whList.length-2;
      h+=`<tr style="${b.sc>100?'background:#faf5ff':isTop?'background:rgba(15,23,42,.65)':isBot?'background:#fff5f5':''}">
        <td style="text-align:center;font-weight:800;color:${b.sc>100?'#7c3aed':isTop?'#059669':isBot?'#dc2626':'#1e293b'}">${b.rk}</td>
        <td style="font-weight:${isTop?700:500}">${b.nm}</td>
        <td style="font-weight:700;color:${b.sc>100?'#7c3aed':isTop?'#059669':isBot?'#dc2626':'#1e293b'};text-align:right">${b.sc.toFixed(1)}${b.sc>100?' 🌟':''}</td>
        <td class="${diff>=0?'pos':'neg'}" style="text-align:right">${diff>=0?'+':''}${diff.toFixed(1)}</td>
        <td class="${diffTop>=0?'pos':'neg'}" style="text-align:right">${diffTop.toFixed(1)}</td>
        <td style="text-align:center">${b.sc>100?'<span class="tag tag-g">목표초과</span>':b.sc>=whAvg?'<span class="tag tag-g">우수</span>':b.sc<whAvg-3?'<span class="tag tag-r">관리필요</span>':'<span class="tag tag-y">보통</span>'}</td>
      </tr>`;
    });
    h+=`</tbody></table>`;
    // 소매 vs 도매 비교
    h+=`<div class="section-divider"></div>
    <h1>Ⅲ. 소매 vs 도매 채널 지사 비교 분석</h1>
    <div class="grid2">
      <div class="card">
        <div class="card-title">소매 지사 현황 요약</div>
        <div class="card-val" style="color:#38bdf8">${rtAvg.toFixed(1)}점</div>
        <div class="card-sub">지사 평균 득점</div>
        <div style="margin-top:12px;font-size:12px;line-height:2">
          <div>📊 총 ${rtList.length}개 지사</div>
          <div class="pos">🥇 최고: ${rtList[0]?.nm} ${rtList[0]?.sc.toFixed(1)}점</div>
          <div class="neg">📉 최저: ${rtList[rtList.length-1]?.nm} ${rtList[rtList.length-1]?.sc.toFixed(1)}점</div>
          <div>📏 지사간 격차: ${rtGap}점</div>
          <div>✅ 평균 이상: ${rtList.filter(b=>b.sc>=rtAvg).length}개 / ⚠️ 미만: ${rtList.filter(b=>b.sc<rtAvg).length}개</div>
        </div>
      </div>
      <div class="card">
        <div class="card-title">도매 지사 현황 요약</div>
        <div class="card-val" style="color:#34d399">${whAvg.toFixed(1)}점</div>
        <div class="card-sub">지사 평균 득점</div>
        <div style="margin-top:12px;font-size:12px;line-height:2">
          <div>📊 총 ${whList.length}개 지사</div>
          <div class="pos">🥇 최고: ${whList[0]?.nm} ${whList[0]?.sc.toFixed(1)}점</div>
          <div class="neg">📉 최저: ${whList[whList.length-1]?.nm} ${whList[whList.length-1]?.sc.toFixed(1)}점</div>
          <div>📏 지사간 격차: ${whGap}점</div>
          <div>${overAchieve.length>0?`🌟 100점 초과: ${overAchieve.length}개 지사`:'100점 초과 없음'}</div>
        </div>
      </div>
    </div>
    <div class="action">
      <strong>지사별 전략 권고:</strong><br>
      🌟 <strong>도매 ${overAchieve.map(b=>b.nm).join('·')||'상위 지사'}</strong> — 100점 초과 달성 노하우를 하위 도매 지사(${whList[whList.length-1]?.nm} 등)에 전수하는 멘토링 세션 월 1회 운영<br>
      📊 <strong>소매 상위 3개 지사</strong>(${rtList.slice(0,3).map(b=>b.nm).join('·')}) — 우수 사례를 전사 '소매 베스트 프랙티스' 자료로 공식화<br>
      ⚠️ <strong>소매 하위 지사</strong>(${rtList.slice(-3).map(b=>b.nm).join('·')}) — 역량강화팀 현장 이음 활동(과제 No.87) 우선 배정, 월 2회 이상 집중 코칭<br>
      🎯 <strong>도매·소매 동반 하위 지사</strong> 교차 점검 — 동일 상권 소매·도매 지사가 동시 하위권이면 본부 차원 종합 지원 검토
    </div>
    </div></body></html>`;
    return openReport(h,'지사별 KPI 분석 보고서');
  }
}

// ==================== AI ====================
let aiConnected=false;
async function testApiConnection(){
  const key=document.getElementById('apiKey').value.trim();const btn=document.getElementById('aiConnBtn');const dot=document.getElementById('aiDot');const txt=document.getElementById('aiStatusTxt');
  if(!key)return;btn.textContent='🔄';btn.disabled=true;
  try{
    const res=await fetch('https://api.anthropic.com/v1/messages',{method:'POST',headers:{'Content-Type':'application/json','x-api-key':key,'anthropic-version':'2023-06-01','anthropic-dangerous-direct-browser-access':'true'},body:JSON.stringify({model:'claude-haiku-4-5-20251001',max_tokens:10,messages:[{role:'user',content:'hi'}]})});
    if(res.ok){aiConnected=true;dot.classList.add('on');txt.textContent='연결됨';btn.textContent='✅';btn.classList.add('connected');addMsg('ok','✅ 연결 성공! AI 심층 분석이 가능합니다.');}
    else{txt.textContent='실패';btn.textContent='❌';}
  }catch(e){txt.textContent='오류';btn.textContent='❌';}
  btn.disabled=false;
}
function addMsg(cls,txt){const div=document.getElementById('aiMsg');div.innerHTML+=`<div class="msg ${cls}">${txt}</div>`;div.scrollTop=div.scrollHeight;}
function buildCtx(){
  const T=D.tasks,P=D.profit,K=D.kpi,SD=D.sga_detail,RD=D.kpi_retail_detail,WD=D.kpi_wholesale_detail,SMD=D.kpi_smb_detail,HC=D.hq_context;
  const zeroInProg=T.filter(t=>t.pp===0&&t.st==='진행중');
  return `당신은 KT M&S 경영기획팀의 Executive AI 분석 어시스턴트입니다.
분석 원칙: ①모든 주장에 수치 근거 명시 ②구조적 원인→파급영향→개선안 순서 ③실행 가능한 구체적 권고(담당/기한 포함) ④임원 보고 수준의 간결·명확한 문체

═══ ${D.period} 경영 데이터 ═══

【손익 현황】
매출: ${fB(P.revenue.total)}억 | 영업이익: ${fB(P.op.total)}억(OPM ${(P.op.total/P.revenue.total*100).toFixed(1)}%) | 순이익: ${fB(P.net_income)}억
채널별 OPM: 소상공인${(P.op.small_biz/P.revenue.small_biz*100).toFixed(1)}% > IoT${(P.op.iot/P.revenue.iot*100).toFixed(1)}% > 소매${(P.op.retail/P.revenue.retail*100).toFixed(1)}% > 도매${(P.op.wholesale/P.revenue.wholesale*100).toFixed(1)}% > 디지털${(P.op.digital/P.revenue.digital*100).toFixed(1)}% > 기업${(P.op.enterprise/P.revenue.enterprise*100).toFixed(1)}% > 법인${(P.op.corporate_sales/P.revenue.corporate_sales*100).toFixed(1)}%

【판관비 구조】총${fB(P.sga.total)}억
인건비${fB(SD['인건비합계'].total)}억(판관비${(SD['인건비합계'].total/P.sga.total*100).toFixed(0)}%) · 소매인건비집중도 판관비내${(SD['인건비합계'].retail/P.sga.retail*100).toFixed(0)}%
판촉비${fB(SD['판매촉진비'].total)}억 · 디지털판촉비비중${(SD['판매촉진비'].digital/P.sga.digital*100).toFixed(0)}% ← 구조적 문제
법인판매수수료${fB(SD['판매수수료'].corporate_sales)}억 vs 법인영업이익${fB(P.op.corporate_sales)}억 ← 수수료>이익 역전 구조

【본부 손익】(영업이익 기준)
${D.hq.map(h=>{const op=h.gp-h.sga;return `${h.nm}: ${fB(op)}억(OPM ${(op/h.rev*100).toFixed(1)}%)`;}).join(' | ')}

【KPI 현황】(종합/소매/도매/소상공인)
${K.map(k=>`${k.hq}#${k.rk}: ${k.ts.toFixed(1)}점 [소매${k.rt.t.toFixed(1)} / 도매${k.wh.t.toFixed(1)} / 소상공${k.sm.t.toFixed(1)}]`).join('\n')}

【KPI 세부 편차】
소매 후불: 동부17.8>강서16.2>서부15.9>강북15.1>강남14.8 (격차 3.0p)
소매 MNP: 동부9.2>강서8.1>서부7.8>강북7.2>강남6.9 (격차 2.3p)
도매 신규: 강서91.2>동부90.1>강남89.3>강북86.5>서부82.4 (격차 8.9p)
소상공 영업: 강서52.5>서부45.9>동부40.6>강북40.2>강남29.6 (격차 22.9p)
소상공 커버리지(%): 강서78>서부71>동부65>강북63>강남52

【과제 현황】
총 ${T.length}개 | 완료${T.filter(t=>t.st==='완료').length}(${(T.filter(t=>t.st==='완료').length/T.length*100).toFixed(0)}%) | 진행중${T.filter(t=>t.st==='진행중').length} | 계획${T.filter(t=>t.st==='계획').length}
⚠️ 진행중 추진도 0% 과제: ${zeroInProg.length}개 (즉시 점검 필요)
핵심과제 ${T.filter(t=>t.co==='●').length}개 중 완료 ${T.filter(t=>t.co==='●'&&t.st==='완료').length}개

【본부별 강약점 컨텍스트】
${Object.entries(HC||{}).map(([h,v])=>`${h}: 강점-${v.str} / 약점-${v.wk} / 기회-${v.opp}`).join('\n')}`;
}
function clearAI(){document.getElementById('aiMsg').innerHTML='<div class="msg s">💡 대화 내역이 초기화되었습니다.</div>';}
// ==================== 종합 브리핑 ====================
function nvl(v,d=0){const n=Number(v);return Number.isFinite(n)?n:d;}
function ratio(num,den){const d=nvl(den);return d===0?null:nvl(num)/d;}
function pctStr(v,d=1){return v===null||!Number.isFinite(v)?'데이터 없음':`${(v*100).toFixed(d)}%`;}
function deltaStr(cur,base,d=1){
  const c=nvl(cur),b=nvl(base);
  if(!b) return '데이터 없음';
  const r=((c-b)/b)*100;
  return `${r>=0?'+':''}${r.toFixed(d)}%`;
}

function buildBriefingEngine(){
  const T=Array.isArray(D.tasks)?D.tasks:[];
  const P=D.profit||{};
  const K=Array.isArray(D.kpi)?[...D.kpi].sort((a,b)=>nvl(b.ts)-nvl(a.ts)):[];
  const C=D.commission||{};
  const SD=D.sga_detail||{};
  const j=C.jan26||null;

  const channelPerf=CHS.map(ch=>({
    key:ch,
    name:CN[ch],
    rev:nvl(P?.revenue?.[ch]),
    op:nvl(P?.op?.[ch]),
    opm:ratio(P?.op?.[ch],P?.revenue?.[ch])
  })).filter(x=>x.rev>0);

  const bestProfit=channelPerf.reduce((a,b)=>a.opm>b.opm?a:b,channelPerf[0]||{name:'-',opm:null,op:0});
  const worstProfit=channelPerf.reduce((a,b)=>a.opm<b.opm?a:b,channelPerf[0]||{name:'-',opm:null,op:0});

  const kpiGap=K.length>1?nvl(K[0].ts)-nvl(K[K.length-1].ts):0;
  const smbGap=K.length>1?nvl(K[0]?.sm?.t)-nvl(K[K.length-1]?.sm?.t):0;

  const done=T.filter(t=>t.st==='완료').length;
  const inProgress=T.filter(t=>t.st==='진행중').length;
  const delayed=T.filter(t=>t.st==='진행중'&&nvl(t.pp)<=0.2);
  const zeroProgress=T.filter(t=>t.st==='진행중'&&nvl(t.pp)===0);
  const coreOpen=T.filter(t=>t.co==='●'&&t.st!=='완료');

  const mgmtCur=nvl(j?.mgmt_fee?.total,NaN);
  const mgmtYoY=deltaStr(j?.mgmt_fee?.total,j?.mgmt_fee_25jan);
  const mgmtVsAvg=deltaStr(j?.mgmt_fee?.total,j?.mgmt_fee_25avg);
  const totalFee=nvl(j?.total_fee?.total,NaN);
  const mgmtMix=ratio(j?.mgmt_fee?.total,j?.total_fee?.total);

  const monthly=(C.monthly25||[]).map(m=>nvl(m.total)).filter(v=>v>0);
  const mom=monthly.length>=2?((monthly[monthly.length-1]-monthly[monthly.length-2])/monthly[monthly.length-2]):null;

  const yearly=(C.yearly||[]);
  const yoyTotal=(yearly.length>=2&&nvl(yearly.at(-2)?.total)>0)?((nvl(yearly.at(-1).total)-nvl(yearly.at(-2).total))/nvl(yearly.at(-2).total)):null;

  // ── 가입자 데이터 통합 ──
  const SD2 = subscriberData || DEFAULT_SUBSCRIBER_DATA;
  const wsD = SD2.wireless||{}; const wss2 = wsD.series||{};
  const wdD = SD2.wired||{}; const wdS2 = wdD.series||{};
  const mgmtD = SD2.mgmt||{}; const mgmtS2 = mgmtD.series||{};
  const hcD2 = SD2.hc||{}; const hcS3 = hcD2.series||{}; const hcCh3 = hcD2.channels||{};
  const lastI2 = (SD2.months||[]).length-1;
  const fN3 = v => fmtSubNum(Number(v)||0);

  const subWirelessTot = Number((wss2.total||[])[lastI2]||wsD.total||0);
  const subWirelessPrev = Number((wss2.total||[])[lastI2-1]||0);
  const subWirelessDelta = subWirelessTot - subWirelessPrev;
  const subCapa = Number((wss2.capa||[])[lastI2]||wsD.capa||0);
  const subMgmtCur = Number((mgmtS2.total||[])[lastI2]||0);
  const subMgmtFirst = Number((mgmtS2.total||[])[0]||0);
  const subMgmtChg = subMgmtFirst ? ((subMgmtCur-subMgmtFirst)/subMgmtFirst*100).toFixed(1) : null;
  const subHcCur = Number((hcS3.total||[])[lastI2]||0);
  const domaeHC2 = hcCh3['도매H/C2%']||{};
  const ktdotHC2 = hcCh3['KT닷컴H/C2%']||{};
  const domaeNet3 = Number((domaeHC2.netAdd||[])[lastI2]||0);
  const ktdotNet3 = Number((ktdotHC2.netAdd||[])[lastI2]||0);
  const domaeStock = Number((domaeHC2.total||[])[lastI2]||0);
  const ktdotStock = Number((ktdotHC2.total||[])[lastI2]||0);
  const hcGap = domaeStock - ktdotStock;
  const subWired = Number(wdD.total||0);
  const subWiredNetOp = Number((wdS2.netAdd||[])[lastI2]||wdD.netAdd||0);
  const subArpu = Number(SD2.arpu?.overall||0);
  const subAkrpu1st = Number((SD2.arpu?.series||[])[0]||0);
  const subArpuChg = subAkrpu1st ? ((subArpu-subAkrpu1st)/subAkrpu1st*100).toFixed(1) : null;

  const missing=[];
  if(!K.length) missing.push('KPI 데이터');
  if(!j) missing.push('1월 수수료 데이터');
  if(!T.length) missing.push('과제 데이터');

  const insights=[
    `전사 영업이익률은 ${pctStr(ratio(P?.op?.total,P?.revenue?.total))}이며, 채널 최고 수익성은 ${bestProfit.name}(${pctStr(bestProfit.opm)}), 최저는 ${worstProfit.name}(${pctStr(worstProfit.opm)})입니다.`,
    K.length?`KPI 1위는 ${K[0].hq}(${nvl(K[0].ts).toFixed(1)}점), 최하위는 ${K[K.length-1].hq}(${nvl(K[K.length-1].ts).toFixed(1)}점)로 격차 ${kpiGap.toFixed(1)}p입니다.`:'KPI 데이터가 없어 본부별 성과 비교가 제한됩니다.',
    `무선 가입자 ${fN3(subWirelessTot)}명(전월${subWirelessDelta>=0?'+':''}${subWirelessDelta.toLocaleString('ko-KR')}), 관리수수료 대상 ${fN3(subMgmtCur)}명 — 연초 대비 ${subMgmtChg||'-'}% 감소 추세.`,
    `H/C(2%): 도매 ${fN3(domaeStock)} vs KT닷컴 ${fN3(ktdotStock)}, 격차 ${fN3(hcGap)}명으로 역전 임박.`,
    `과제 완료율은 ${T.length?((done/T.length)*100).toFixed(1):'0.0'}%이며, 지연 위험 과제(진행중·추진도 20% 이하)는 ${delayed.length}건입니다.`
  ];

  const actions=[
    {priority:'HIGH',item:'수익성 하위 BM 개선',owner:'영업총괄/경영기획',due:'즉시~2Q',detail:`${worstProfit.name} 채널의 판관비·수수료 구조를 월 단위로 점검하고 기준 OPM 목표를 재설정`},
    {priority:'HIGH',item:'무선 정책수수료 채널 편중 개선',owner:'경영기획/디지털강화',due:'1Q',detail:`소매·디지털 2채널이 무선 정책수수료(${D.commission.jan26.policy_fee.total}억)의 ${((D.commission.jan26.policy_fee.retail+D.commission.jan26.policy_fee.digital)/D.commission.jan26.policy_fee.total*100).toFixed(0)}% 집중. 법인영업·기업/공공 정책 차감(-) 구조 원인 분석 및 KT-RDS 단가 구조 재검토`},
    {priority:'HIGH',item:'소상공 KPI 편차 축소',owner:'본부장협의체',due:'1개월',detail:`소상공 편차 ${smbGap.toFixed(1)}p 해소를 위해 하위 본부 집중 코칭 및 우수본부 전파`},
    {priority:'HIGH',item:'관리수수료 대상 가입자 방어',owner:'도매강화/소매강화',due:'즉시',detail:`연초 대비 ${subMgmtChg||'-'}% 감소 중. 소매 CAPA 확대 및 이탈 방어 캠페인 긴급 검토`},
    {priority:'HIGH',item:'도매H/C→KT닷컴 구조 전환 관리',owner:'디지털강화팀',due:'즉시',detail:`격차 ${fN3(hcGap)}명으로 역전 임박. 단가 구조 영향도 분석 후 선제 대응 방안 수립`},
    {priority:'MED',item:'KT-RDS 수수료 YoY 하락 방어',owner:'경영기획팀',due:'2Q',detail:`KT-RDS ${D.commission.jan26.policy_fee.rds}억(전년동월 ${D.commission.jan26.policy_fee_25jan}억) — YoY ${(((D.commission.jan26.policy_fee.total-D.commission.jan26.policy_fee_25jan)/D.commission.jan26.policy_fee_25jan)*100).toFixed(1)}%. 소매·디지털 단건 수수료율 방어 방안 마련`},
    {priority:'MED',item:'지연 과제 정상화',owner:'과제 오너',due:'주간',detail:`지연 위험 ${delayed.length}건 중 핵심 ${delayed.filter(t=>t.co==='●').length}건을 주간 리커버리 리스트로 운영`}
  ];

  return {T,P,K,C,SD,j,done,inProgress,delayed,zeroProgress,coreOpen,bestProfit,worstProfit,kpiGap,smbGap,mgmtCur,mgmtYoY,mgmtVsAvg,totalFee,mgmtMix,mom,yoyTotal,missing,insights,actions,
    sub:{wirelessTot:subWirelessTot,wirelessDelta:subWirelessDelta,capa:subCapa,mgmtCur:subMgmtCur,mgmtChg:subMgmtChg,hcCur:subHcCur,domaeStock,ktdotStock,hcGap,domaeNet:domaeNet3,ktdotNet:ktdotNet3,wired:subWired,wiredNetOp:subWiredNetOp,arpu:subArpu,arpuChg:subArpuChg,months:SD2.months||[],lastIdx:lastI2}};
}

function initBriefing(){
  const E=buildBriefingEngine();
  const fN4 = v => fmtSubNum(Number(v)||0);
  const fDelta = (v,green) => {
    const n=Number(v)||0;
    const col = green ? (n>=0?'#15803d':'#b91c1c') : (n<=0?'#15803d':'#b91c1c');
    return `<span style="color:${col};font-weight:700">${n>=0?'+':''}${n.toLocaleString('ko-KR')}</span>`;
  };
  const s = E.sub;
  const hcRevPct = s.ktdotStock&&s.domaeStock ? (s.ktdotStock/s.domaeStock*100).toFixed(0) : '-';

  // 신호등 판정
  const sig = (good, warn) => good ? '🟢' : warn ? '🟡' : '🔴';
  const profitOpm = (E.P?.op?.total||0)/(E.P?.revenue?.total||1);
  const kpiOk = E.K.length>0 && E.kpiGap < 15;
  const mgmtOk = s.mgmtChg !== null && parseFloat(s.mgmtChg) > -5;
  const taskOk = E.T.length>0 && (E.done/E.T.length) > 0.15;

  document.getElementById('briefingSnapshot').innerHTML=`
  <div style="background:#f8fafc;border:1px solid #dbeafe;border-radius:14px;padding:16px;margin-bottom:10px">

    <!-- 헤더 -->
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:14px;flex-wrap:wrap;gap:6px">
      <div>
        <div style="font-size:15px;font-weight:900;color:#0f172a">📑 종합 브리핑 분석 허브</div>
        <div style="font-size:11px;color:#64748b;margin-top:2px">${D.period} · 손익 · KPI · 무선/유선 가입자 · 과제 통합</div>
      </div>
      <div style="display:flex;gap:6px;flex-wrap:wrap">
        <span style="background:#fee2e2;color:#991b1b;font-size:10px;font-weight:700;padding:3px 10px;border-radius:20px">⚡ HIGH ${E.actions.filter(a=>a.priority==='HIGH').length}건</span>
        <span style="background:#fef9c3;color:#854d0e;font-size:10px;font-weight:700;padding:3px 10px;border-radius:20px">⚠️ MED ${E.actions.filter(a=>a.priority==='MED').length}건</span>
      </div>
    </div>

    <!-- 경영 4대 축 카드 -->
    <div style="display:grid;grid-template-columns:repeat(2,1fr);gap:8px;margin-bottom:10px" class="b-grid4">
      <div style="background:#fff;border:1px solid #dbeafe;border-radius:10px;padding:10px;border-top:3px solid #3b82f6">
        <div style="font-size:9px;color:#64748b;font-weight:700">💰 손익 ${sig(profitOpm>0.08, profitOpm>0.04)}</div>
        <div style="font-size:16px;font-weight:900;color:#0f172a;margin:4px 0">${fB(nvl(E.P?.op?.total))}억</div>
        <div style="font-size:10px;color:#475569">영업이익 · ${pctStr(ratio(E.P?.op?.total,E.P?.revenue?.total))} OPM</div>
        <div style="font-size:10px;color:#64748b;margin-top:3px">매출 ${fB(nvl(E.P?.revenue?.total))}억 · 최저 ${E.worstProfit.name}</div>
      </div>
      <div style="background:#fff;border:1px solid #d1fae5;border-radius:10px;padding:10px;border-top:3px solid #10b981">
        <div style="font-size:9px;color:#64748b;font-weight:700">🎯 KPI ${sig(kpiOk, E.kpiGap<25)}</div>
        <div style="font-size:16px;font-weight:900;color:#0f172a;margin:4px 0">${E.K.length?E.K[0].hq+' 1위':'-'}</div>
        <div style="font-size:10px;color:#475569">본부간 격차 ${E.kpiGap.toFixed(1)}p</div>
        <div style="font-size:10px;color:#64748b;margin-top:3px">최하위 ${E.K.length?E.K[E.K.length-1].hq:'-'} · 소상공 ${E.smbGap.toFixed(1)}p</div>
      </div>
      <div style="background:#fff;border:1px solid #fde68a;border-radius:10px;padding:10px;border-top:3px solid #f59e0b">
        <div style="font-size:9px;color:#64748b;font-weight:700">📋 과제 ${sig(taskOk, E.T.length>0)}</div>
        <div style="font-size:16px;font-weight:900;color:#0f172a;margin:4px 0">${E.T.length?((E.done/E.T.length)*100).toFixed(0):'0'}%</div>
        <div style="font-size:10px;color:#475569">완료율 · ${E.done}/${E.T.length}건</div>
        <div style="font-size:10px;color:#b91c1c;margin-top:3px">0% 미착수 ${E.zeroProgress.length}건 · 핵심잔여 ${E.coreOpen.length}건</div>
      </div>
      <div style="background:#fff;border:1px solid #e9d5ff;border-radius:10px;padding:10px;border-top:3px solid #a855f7">
        <div style="font-size:9px;color:#64748b;font-weight:700">💴 수수료수입</div>
        <div style="font-size:16px;font-weight:900;color:#0f172a;margin:4px 0">${Number.isFinite(E.totalFee)?E.totalFee.toFixed(1)+'억':'N/A'}</div>
        <div style="font-size:10px;color:#475569">관리비중 ${pctStr(E.mgmtMix)}</div>
        <div style="font-size:10px;color:#64748b;margin-top:3px">YoY ${E.mgmtYoY} · 관리 ${Number.isFinite(E.mgmtCur)?E.mgmtCur.toFixed(1)+'억':'N/A'}</div>
      </div>
    </div>

    <!-- 가입자 4카드 -->
    <div style="font-size:10px;font-weight:800;color:#475569;margin-bottom:6px;padding-left:2px">📡 가입자 현황</div>
    <div style="display:grid;grid-template-columns:repeat(2,1fr);gap:8px;margin-bottom:12px" class="b-grid4">
      <div style="background:#fff;border:1px solid #e9d5ff;border-radius:10px;padding:10px;border-top:3px solid #8b5cf6">
        <div style="font-size:9px;color:#64748b;font-weight:700">📡 무선 전체 ${sig(!s.mgmtChg||parseFloat(s.mgmtChg||0)>-3, parseFloat(s.mgmtChg||0)>-8)}</div>
        <div style="font-size:16px;font-weight:900;color:#0f172a;margin:4px 0">${fN4(s.wirelessTot)}</div>
        <div style="font-size:10px;color:#475569">전월 ${fDelta(s.wirelessDelta,true)} · CAPA ${fN4(s.capa)}</div>
        <div style="font-size:10px;color:#64748b;margin-top:3px">ARPU ${s.arpu?Math.round(s.arpu).toLocaleString('ko-KR')+'원':'-'}${s.arpuChg?' (연초'+s.arpuChg+'%)':''}</div>
      </div>
      <div style="background:#fff;border:1px solid #fee2e2;border-radius:10px;padding:10px;border-top:3px solid #ef4444">
        <div style="font-size:9px;color:#64748b;font-weight:700">💴 관리수수료 대상 🔴</div>
        <div style="font-size:16px;font-weight:900;color:#dc2626;margin:4px 0">${fN4(s.mgmtCur)}</div>
        <div style="font-size:10px;color:#b91c1c">연초대비 ${s.mgmtChg||'-'}% 감소</div>
        <div style="font-size:10px;color:#64748b;margin-top:3px">소매8.5% 위주 구조적 이탈</div>
      </div>
      <div style="background:#fff;border:1px solid #fed7aa;border-radius:10px;padding:10px;border-top:3px solid #f97316">
        <div style="font-size:9px;color:#64748b;font-weight:700">🔀 H/C 역전 경보 🔴</div>
        <div style="font-size:16px;font-weight:900;color:#dc2626;margin:4px 0">${fN4(s.hcGap)}</div>
        <div style="font-size:10px;color:#475569">도매-닷컴 격차 · 닷컴 ${hcRevPct}%</div>
        <div style="font-size:10px;color:#64748b;margin-top:3px">도매 월${fDelta(s.domaeNet,false)} · 닷컴 월${fDelta(s.ktdotNet,true)}</div>
      </div>
      ${(()=>{
        const sdL=subscriberData||DEFAULT_SUBSCRIBER_DATA;
        const wdL=sdL.wired||{}; const wdSL=wdL.series||{};
        const lastIL=(sdL.months||[]).length-1;
        const vNetOp=Number((wdSL.netAdd||[])[lastIL]||0);
        const vC=vNetOp>=0?'#15803d':'#b91c1c';
        return '<div style="background:#fff;border:1px solid #bfdbfe;border-radius:10px;padding:10px;border-top:3px solid #3b82f6"><div style="font-size:9px;color:#64748b;font-weight:700">🛜 유선(인터넷+TV)</div><div style="font-size:16px;font-weight:900;color:#0f172a;margin:4px 0">'+fN4(wdL.total||0)+'</div><div style="font-size:10px;color:#475569">영업순증 <span style="color:'+vC+';font-weight:700">'+(vNetOp>=0?'+':'')+vNetOp.toLocaleString('ko-KR')+'</span></div><div style="font-size:10px;color:#64748b;margin-top:3px">인터넷 '+fN4(wdL.internet||0)+' · TV '+fN4(wdL.tv||0)+'</div></div>';
      })()}
    </div>

    <!-- 종합 시사점 -->
    <div style="background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:12px;margin-bottom:10px">
      <div style="font-size:11px;font-weight:800;color:#1e293b;margin-bottom:8px">📌 종합 시사점</div>
      <ul style="margin:0;padding-left:16px;color:#334155;font-size:11px;line-height:1.9">
        ${E.insights.map(i=>'<li>'+i+'</li>').join('')}
      </ul>
      ${E.missing.length?'<div style="margin-top:8px;font-size:10px;color:#b45309;background:#fffbeb;padding:6px 10px;border-radius:6px">※ 누락 데이터: '+E.missing.join(', ')+'</div>':''}
    </div>

    <!-- 즉시 실행 액션 -->
    <div style="background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:12px">
      <div style="font-size:11px;font-weight:800;color:#1e293b;margin-bottom:8px">⚡ 즉시 실행 필요 항목</div>
      ${E.actions.filter(a=>a.priority==='HIGH').map(a=>'<div style="display:flex;gap:8px;align-items:flex-start;padding:7px 0;border-bottom:1px solid #f1f5f9"><span style="background:#fee2e2;color:#991b1b;font-size:9px;font-weight:800;padding:2px 7px;border-radius:10px;white-space:nowrap;margin-top:1px">HIGH</span><div><div style="font-size:11px;font-weight:700;color:#0f172a">'+a.item+'</div><div style="font-size:10px;color:#64748b;margin-top:2px">'+a.detail+'</div></div></div>').join('')}
    </div>

  </div>`;
}

// ① 이달의 경영 브리핑 (경영진 보고서)
function genMonthlyBriefing(){
  const E = buildBriefingEngine();
  const P=E.P, K=E.K, C=E.C, SD=E.SD;
  const s = E.sub;
  const opm = ratio(P?.op?.total, P?.revenue?.total);

  // 가입자 계산
  const sdB = subscriberData||DEFAULT_SUBSCRIBER_DATA;
  const wsB=sdB.wireless||{}, wssB=wsB.series||{};
  const wdB=sdB.wired||{},   wdSB=wdB.series||{};
  const mgB=sdB.mgmt||{},    mgSB=mgB.series||{};
  const hcB=sdB.hc||{},      hcSB=hcB.series||{}, hcChB=hcB.channels||{};
  const liB=(sdB.months||[]).length-1;
  const fN6=v=>fmtSubNum(Number(v)||0);
  const fD6=(v,pos=true)=>{const n=Number(v)||0;const c=pos?(n>=0?'#15803d':'#b91c1c'):(n<=0?'#15803d':'#b91c1c');return `<span style="color:${c};font-weight:700">${n>=0?'+':''}${Math.abs(n).toLocaleString('ko-KR')}</span>`;};
  const wTot=Number((wssB.total||[])[liB]||wsB.total||0);
  const wPrev=Number((wssB.total||[])[liB-1]||0);
  const wCapa=Number((wssB.capa||[])[liB]||wsB.capa||0);
  const wNetOp=Number((wssB.netAdd||[])[liB]||0);
  const mgCur=Number((mgSB.total||[])[liB]||0);
  const mgPrev=Number((mgSB.total||[])[liB-1]||0);
  const mg1st=Number((mgSB.total||[])[0]||0);
  const mgYtdChg=mg1st?((mgCur-mg1st)/mg1st*100).toFixed(1):null;
  const hcCur=Number((hcSB.total||[])[liB]||0);
  const hcPrev=Number((hcSB.total||[])[liB-1]||0);
  const domaeStk=Number(((hcChB['도매H/C2%']||{}).total||[])[liB]||0);
  const ktdotStk=Number(((hcChB['KT닷컴H/C2%']||{}).total||[])[liB]||0);
  const domaeNet=Number(((hcChB['도매H/C2%']||{}).netAdd||[])[liB]||0);
  const ktdotNet=Number(((hcChB['KT닷컴H/C2%']||{}).netAdd||[])[liB]||0);
  const vNetOp=Number((wdSB.netAdd||[])[liB]||wdB.netAdd||0);
  const hcPct=domaeStk+ktdotStk>0?((ktdotStk/(domaeStk+ktdotStk))*100).toFixed(0):'-';
  const hcBarW=domaeStk+ktdotStk>0?((domaeStk/(domaeStk+ktdotStk))*100).toFixed(1):50;

  // BM 바 차트 최대값 (플랫폼 포함)
  const PFT2=P?.platform?.total||{};
  const pfRev=PFT2.revenue||0, pfOp=PFT2.op||0, pfGp=PFT2.contribution||0, pfSga=PFT2.sga||0;
  const revMax=Math.max(...CHS.map(c=>nvl(P?.revenue?.[c])),pfRev,1);
  const opMax=Math.max(...CHS.map(c=>Math.abs(nvl(P?.op?.[c]))),Math.abs(pfOp),1);

  // 과제 총괄별 집계
  const byDiv={경총:{tot:0,done:0,zero:0},영총:{tot:0,done:0,zero:0}};
  (E.T||[]).forEach(t=>{const d=t.ch==='경총'?'경총':'영총';byDiv[d].tot++;if(t.st==='완료')byDiv[d].done++;if(t.pp===0&&t.st!=='완료')byDiv[d].zero++;});

  // KPI 소상공 순위 정렬
  const smbRank=[...K].sort((a,b)=>b.sm.t-a.sm.t);
  const SMD=D.kpi_smb_detail||{};

  const yoyTxt=E.yoyTotal===null?'N/A':`${E.yoyTotal>=0?'+':''}${(E.yoyTotal*100).toFixed(1)}%`;
  const momTxt=E.mom===null?'N/A':`${E.mom>=0?'+':''}${(E.mom*100).toFixed(1)}%`;

  // 리스크 등급 계산
  const risks=[];
  if(opm!==null&&opm<0.05) risks.push({lv:'HIGH',col:'#dc2626',bg:'#fff1f2',item:'전사 영업이익률 경보',desc:`${pctStr(opm)} — 목표 5% 미달. 판관비 구조 즉시 점검 필요.`});
  if(mgYtdChg&&parseFloat(mgYtdChg)<-5) risks.push({lv:'HIGH',col:'#dc2626',bg:'#fff1f2',item:'관리수수료 대상 구조적 감소',desc:`연초 대비 ${mgYtdChg}% 감소 (${fN6(mgCur)}명). 소매 이탈 방어 캠페인 긴급 필요.`});
  if(domaeStk>0&&ktdotStk/domaeStk>0.7) risks.push({lv:'HIGH',col:'#dc2626',bg:'#fff1f2',item:'H/C 구조 역전 임박',desc:`도매 ${fN6(domaeStk)} vs 닷컴 ${fN6(ktdotStk)} — 닷컴 ${hcPct}% 수준. 수수료 단가 재검토 시급.`});
  if(E.smbGap>15) risks.push({lv:'HIGH',col:'#dc2626',bg:'#fff1f2',item:'소상공인 KPI 본부간 극심한 편차',desc:`최대 ${E.smbGap.toFixed(1)}p 차이. 하위 본부 집중 지원 TF 필요.`});
  // 무선 정책수수료 리스크 (항상 표시)
  const polFee = D.commission.jan26.policy_fee;
  const polYoy = (((polFee.total - D.commission.jan26.policy_fee_25jan)/D.commission.jan26.policy_fee_25jan)*100).toFixed(1);
  const polConc = ((polFee.retail+polFee.digital)/polFee.total*100).toFixed(0);
  risks.push({lv: parseFloat(polYoy)<0?'HIGH':'MED', col: parseFloat(polYoy)<0?'#dc2626':'#b45309', bg: parseFloat(polYoy)<0?'#fff1f2':'#fffbeb',
    item:'무선 정책수수료 채널 편중·YoY 변동',
    desc:`KT-RDS+정책 합산 ${polFee.total}억 (전년 ${D.commission.jan26.policy_fee_25jan}억, YoY ${polYoy}%). 소매·디지털 집중도 ${polConc}% — 법인·기업 정책 차감 구조 모니터링 필요.`
  });
  if(E.zeroProgress.length>5) risks.push({lv:'MED',col:'#b45309',bg:'#fffbeb',item:'저추진 과제 다수',desc:`진행중·추진도 0% ${E.zeroProgress.length}건 (핵심 ${E.zeroProgress.filter(t=>t.co==='●').length}건). 주간 리커버리 운영.`});

  let h=`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>이달의 경영 브리핑</title>${rptStyle()}
  <style>
    /* ─── 브리핑 전용 ─── */
    .bs{margin-bottom:22px}
    .bs-hd{display:flex;align-items:center;gap:8px;margin-bottom:12px;padding-bottom:8px;border-bottom:2px solid #e2e8f0}
    .bs-hd-num{width:24px;height:24px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:900;flex-shrink:0}
    .bs-hd-ttl{font-size:14px;font-weight:900;color:#0f172a}
    .bs-hd-sub{font-size:10px;color:#64748b;margin-left:auto}
    /* KPI 배너 카드 */
    .bk-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:14px}
    .bk{border-radius:10px;padding:12px;position:relative;overflow:hidden}
    .bk-lbl{font-size:9px;font-weight:800;text-transform:uppercase;letter-spacing:.5px;margin-bottom:6px}
    .bk-v{font-size:20px;font-weight:900;line-height:1.1}
    .bk-sub{font-size:10px;margin-top:5px;opacity:.8}
    /* 인라인 바 */
    .ib-row{display:flex;align-items:center;gap:8px;padding:5px 0;border-bottom:1px solid #f1f5f9}
    .ib-row:last-child{border:none}
    .ib-lbl{width:72px;font-size:11px;font-weight:600;flex-shrink:0;color:#1e293b}
    .ib-track{flex:1;height:8px;border-radius:4px;background:#f1f5f9;overflow:hidden}
    .ib-fill{height:100%;border-radius:4px}
    .ib-val{width:56px;text-align:right;font-size:11px;font-weight:700;flex-shrink:0}
    .ib-pct{width:44px;text-align:right;font-size:10px;color:#64748b;flex-shrink:0}
    /* KPI 히트맵 */
    .hm-cell{border-radius:6px;padding:5px 8px;text-align:center;font-size:12px;font-weight:700}
    .hm-g{background:#dcfce7;color:#15803d}
    .hm-y{background:#fef9c3;color:#854d0e}
    .hm-r{background:#fee2e2;color:#dc2626}
    /* 과제 진척 */
    .task-bar{height:10px;border-radius:5px;background:#e2e8f0;overflow:hidden;margin:6px 0}
    .task-fill{height:100%;border-radius:5px;background:linear-gradient(90deg,#10b981,#3b82f6)}
    /* 리스크 */
    .risk-row{display:flex;gap:10px;align-items:flex-start;padding:9px 0;border-bottom:1px solid #f1f5f9}
    .risk-row:last-child{border:none}
    .risk-badge{font-size:9px;font-weight:800;padding:3px 8px;border-radius:10px;white-space:nowrap;flex-shrink:0;margin-top:1px}
    /* 액션 */
    .ac-row{display:flex;gap:10px;align-items:flex-start;padding:10px;border-radius:8px;margin-bottom:6px}
    .ac-num{width:22px;height:22px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:900;flex-shrink:0;color:#fff}
    /* 인쇄 */
    @media print{
      .bs{break-inside:avoid}.bk-grid{grid-template-columns:repeat(4,1fr)}
      .page{padding:14mm 12mm 16mm;max-width:210mm}
      table{font-size:10px}th,td{padding:5px 8px}
    }
    @media(max-width:720px){.bk-grid{grid-template-columns:repeat(2,1fr)}}
  </style></head><body>

  <!-- ═══ 표지 ═══ -->
  <div class="cover">
    <div class="cover-tag">Executive Briefing · Management Report · ${D.period}</div>
    <div class="cover-title">KT M&S 이달의 경영 브리핑</div>
    <div class="cover-sub">BM 손익 / KPI / 가입자 / 수수료 / 사업과제 통합 분석</div>
    <div class="cover-meta">
      <div class="cover-kpi"><div class="cover-kpi-v">${fB(nvl(P?.revenue?.total))}억</div><div class="cover-kpi-l">매출</div></div>
      <div class="cover-kpi"><div class="cover-kpi-v">${pctStr(opm)}</div><div class="cover-kpi-l">영업이익률</div></div>
      <div class="cover-kpi"><div class="cover-kpi-v">${fN6(wTot)}</div><div class="cover-kpi-l">무선 가입자</div></div>
      <div class="cover-kpi"><div class="cover-kpi-v">${E.done}/${E.T.length}</div><div class="cover-kpi-l">과제 완료</div></div>
    </div>
  </div>
  <div class="page">

  <!-- ═══ Ⅰ. Executive Summary ═══ -->
  <div class="bs">
    <div class="bs-hd">
      <div class="bs-hd-num" style="background:#0f172a;color:#fff">Ⅰ</div>
      <div class="bs-hd-ttl">Executive Summary</div>
      <div class="bs-hd-sub">핵심 지표 한눈에</div>
    </div>

    <!-- 6대 KPI 배너 -->
    <div class="bk-grid">
      <div class="bk" style="background:linear-gradient(135deg,#f0fdf4,#dcfce7);border:1px solid #bbf7d0">
        <div class="bk-lbl" style="color:#065f46">💰 영업이익</div>
        <div class="bk-v" style="color:${(P?.op?.total||0)>0?'#059669':'#dc2626'}">${fB(nvl(P?.op?.total))}억</div>
        <div class="bk-sub" style="color:#166534">매출 ${fB(nvl(P?.revenue?.total))}억 · OPM ${pctStr(opm)}</div>
      </div>
      <div class="bk" style="background:linear-gradient(135deg,#eff6ff,#dbeafe);border:1px solid #bfdbfe">
        <div class="bk-lbl" style="color:#1e40af">🎯 KPI 1위</div>
        <div class="bk-v" style="color:#1d4ed8">${K.length?K[0].hq.replace('본부',''):'-'}</div>
        <div class="bk-sub" style="color:#1e40af">${K.length?K[0].ts.toFixed(1)+'점 · 격차 '+E.kpiGap.toFixed(1)+'p':'-'}</div>
      </div>
      <div class="bk" style="background:linear-gradient(135deg,#fdf4ff,#f3e8ff);border:1px solid #e9d5ff">
        <div class="bk-lbl" style="color:#6b21a8">📡 무선 가입자</div>
        <div class="bk-v" style="color:#7c3aed">${fN6(wTot)}</div>
        <div class="bk-sub" style="color:#6b21a8">전월 ${fD6(wTot-wPrev,true)} · CAPA ${fN6(wCapa)}</div>
      </div>
      <div class="bk" style="background:linear-gradient(135deg,#fefce8,#fef9c3);border:1px solid #fde68a">
        <div class="bk-lbl" style="color:#92400e">📋 과제 완료율</div>
        <div class="bk-v" style="color:#b45309">${E.T.length?((E.done/E.T.length)*100).toFixed(0):'0'}%</div>
        <div class="bk-sub" style="color:#92400e">${E.done}/${E.T.length}건 · 지연위험 ${E.delayed.length}건</div>
      </div>
    </div>

    <!-- 종합 시사점 박스 -->
    <div style="background:#f8fafc;border-left:4px solid #3b82f6;border-radius:0 10px 10px 0;padding:12px 16px">
      <div style="font-size:11px;font-weight:800;color:#1e40af;margin-bottom:8px">📌 이달의 핵심 시사점</div>
      ${E.insights.map((ins,i)=>`<div style="display:flex;gap:8px;align-items:flex-start;margin-bottom:5px;font-size:11px;color:#1e293b;line-height:1.7"><span style="flex-shrink:0;font-weight:800;color:#3b82f6">${i+1}.</span><span>${ins}</span></div>`).join('')}
    </div>
    ${E.missing.length?`<div style="margin-top:8px;font-size:10px;color:#b45309;background:#fffbeb;padding:6px 10px;border-radius:6px;border:1px solid #fde68a">⚠️ 누락 데이터: ${E.missing.join(', ')} — 해당 항목은 보수적으로 해석</div>`:''}
  </div>

  <!-- ═══ Ⅱ. BM별 손익 분석 ═══ -->
  <div class="bs">
    <div class="bs-hd">
      <div class="bs-hd-num" style="background:#10b981;color:#fff">Ⅱ</div>
      <div class="bs-hd-ttl">BM별 손익 분석</div>
      <div class="bs-hd-sub">채널 7개 + 유통플랫폼 영업이익률 비교</div>
    </div>

    <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px">
      <!-- 매출 인라인 바 -->
      <div>
        <div style="font-size:11px;font-weight:700;color:#475569;margin-bottom:8px">📊 채널별 매출 (백만원)</div>
        ${CHS.map(c=>{
          const v=nvl(P?.revenue?.[c]);
          const pct=(v/revMax*100).toFixed(1);
          return `<div class="ib-row">
            <div class="ib-lbl">${CN[c]}</div>
            <div class="ib-track"><div class="ib-fill" style="width:${pct}%;background:${CC[c]}"></div></div>
            <div class="ib-val">${fM(v)}</div>
            <div class="ib-pct">${(v/nvl(P?.revenue?.total)*100).toFixed(1)}%</div>
          </div>`;
        }).join('')}
        <div class="ib-row" style="border-top:2px dashed #bfdbfe;margin-top:3px;padding-top:7px">
          <div class="ib-lbl" style="color:#0891b2;font-weight:700">유통플랫폼</div>
          <div class="ib-track"><div class="ib-fill" style="width:${(pfRev/revMax*100).toFixed(1)}%;background:#0891b2"></div></div>
          <div class="ib-val" style="color:#0891b2">${fM(pfRev)}</div>
          <div class="ib-pct" style="color:#0891b2">${nvl(P?.revenue?.total)?(pfRev/nvl(P?.revenue?.total)*100).toFixed(1)+'%':'-'}</div>
        </div>
      </div>
      <!-- 영업이익 인라인 바 -->
      <div>
        <div style="font-size:11px;font-weight:700;color:#475569;margin-bottom:8px">💹 채널별 영업이익 (백만원)</div>
        ${CHS.map(c=>{
          const v=nvl(P?.op?.[c]);
          const rev=nvl(P?.revenue?.[c]);
          const pm=ratio(v,rev);
          const pct=(Math.abs(v)/opMax*100).toFixed(1);
          const bc=v>=0?(pm>=0.08?'#059669':pm>=0.04?'#f59e0b':'#f97316'):'#dc2626';
          const pmLabel=pm!==null?pctStr(pm):'N/A';
          return `<div class="ib-row">
            <div class="ib-lbl">${CN[c]}</div>
            <div class="ib-track"><div class="ib-fill" style="width:${pct}%;background:${bc}"></div></div>
            <div class="ib-val" style="color:${bc}">${fM(v)}</div>
            <div class="ib-pct" style="color:${bc}">${pmLabel}</div>
          </div>`;
        }).join('')}
        ${(()=>{
          const pm2=pfRev?pfOp/pfRev:null;
          const bc2=pfOp>=0?(pm2>=0.08?'#059669':pm2>=0.04?'#f59e0b':'#f97316'):'#dc2626';
          return `<div class="ib-row" style="border-top:2px dashed #bfdbfe;margin-top:3px;padding-top:7px">
            <div class="ib-lbl" style="color:#0891b2;font-weight:700">유통플랫폼</div>
            <div class="ib-track"><div class="ib-fill" style="width:${(Math.abs(pfOp)/opMax*100).toFixed(1)}%;background:${bc2}"></div></div>
            <div class="ib-val" style="color:${bc2}">${fM(pfOp)}</div>
            <div class="ib-pct" style="color:${bc2}">${pctStr(pm2)}</div>
          </div>`;
        })()}
      </div>
    </div>

    <div style="margin-top:10px">
      <table><thead><tr><th>BM</th><th>매출</th><th>매출총이익</th><th>판관비</th><th>영업이익</th><th>OPM</th><th>판정</th></tr></thead><tbody>
      ${CHS.map(c=>{
        const rev=nvl(P?.revenue?.[c]),gp=nvl(P?.gross?.[c]),sga=nvl(P?.sga?.[c]),op=nvl(P?.op?.[c]);
        const pm=ratio(op,rev);
        const grade=pm===null?'-':pm>=0.08?'✅ 우수':pm>=0.04?'⚠️ 보통':'🔴 개선필요';
        const pc=pm===null?'':pm>=0.08?'#059669':pm>=0.04?'#b45309':'#dc2626';
        return `<tr><td style="font-weight:600">${CN[c]}</td><td>${fM(rev)}</td><td>${fM(gp)}</td><td>${fM(sga)}</td><td style="color:${pc};font-weight:700">${fM(op)}</td><td style="color:${pc};font-weight:800">${pctStr(pm)}</td><td>${grade}</td></tr>`;
      }).join('')}
      <tr style="border-top:1px solid #bfdbfe"><td style="font-weight:700;color:#0891b2">유통플랫폼</td><td style="color:#0891b2">${fM(pfRev)}</td><td style="color:#0891b2">${fM(pfGp)}</td><td style="color:#0891b2">${fM(pfSga)}</td><td style="color:${pfOp>=0?'#059669':'#dc2626'};font-weight:700">${fM(pfOp)}</td><td style="color:${pfOp>=0?'#059669':'#dc2626'};font-weight:800">${pfRev?pctStr(pfOp/pfRev):'-'}</td><td>${pfOp>=0?'✅ 흑자':'🔴 적자'}</td></tr>
      <tr style="background:rgba(56,189,248,.06);font-weight:800"><td>전사 합계</td><td>${fM(nvl(P?.revenue?.total))}</td><td>${fM(nvl(P?.gross?.total))}</td><td>${fM(nvl(P?.sga?.total))}</td><td style="color:#f97316;font-weight:800">${fM(nvl(P?.op?.total))}</td><td style="color:#f97316;font-weight:900">${pctStr(opm)}</td><td></td></tr>
      </tbody></table>
    </div>
    <div class="caution">채널 최고 수익 BM: ${E.bestProfit.name}(${pctStr(E.bestProfit.opm)}) · 최저: ${E.worstProfit.name}(${pctStr(E.worstProfit.opm)}). 유통플랫폼은 모바일파트·플랫폼사업1·2팀 합산. 본부×채널 교차값은 KPI 비율 추정치.</div>
  </div>

  <!-- ═══ Ⅲ. KPI 달성 현황 ═══ -->
  <div class="bs">
    <div class="bs-hd">
      <div class="bs-hd-num" style="background:#3b82f6;color:#fff">Ⅲ</div>
      <div class="bs-hd-ttl">KPI 달성 현황</div>
      <div class="bs-hd-sub">소매60%+도매30%+소상공10%</div>
    </div>

    <table><thead><tr>
      <th>본부</th><th>종합</th><th>순위</th>
      <th>소매(60%)</th><th>도매(30%)</th><th>소상공(10%)</th>
      <th>소매-MNP</th><th>도매-유선</th><th>소상공-영업</th>
    </tr></thead><tbody>
    ${K.map((k,i)=>{
      const rd=D.kpi_retail_detail[k.hq]||{};
      const wd=D.kpi_wholesale_detail[k.hq]||{};
      const sd2=SMD[k.hq]||{};
      const hmc_=(v,arr)=>{const max=Math.max(...arr),min=Math.min(...arr);return v===max?'hm-g':v===min?'hm-r':'hm-y';};
      const rtArr=K.map(x=>x.rt.t),whArr=K.map(x=>x.wh.t),smArr=K.map(x=>x.sm.t);
      return `<tr>
        <td style="font-weight:700">${k.hq.replace('본부','')}</td>
        <td style="font-weight:900;color:${RC[k.rk]}">${k.ts.toFixed(1)}</td>
        <td style="text-align:center"><span style="background:${RC[k.rk]};color:#fff;font-size:10px;font-weight:800;padding:2px 7px;border-radius:10px">#${k.rk}</span></td>
        <td><span class="hm-cell ${hmc_(k.rt.t,rtArr)}">${k.rt.t.toFixed(1)}</span></td>
        <td><span class="hm-cell ${hmc_(k.wh.t,whArr)}">${k.wh.t.toFixed(1)}</span></td>
        <td><span class="hm-cell ${hmc_(k.sm.t,smArr)}">${k.sm.t.toFixed(1)}</span></td>
        <td style="font-size:11px;color:${rd.mnp>7?'#15803d':'#dc2626'}">${rd.mnp?.toFixed(1)||'-'}</td>
        <td style="font-size:11px;color:${wd.wired>18?'#15803d':wd.wired<16?'#dc2626':'#475569'}">${wd.wired?.toFixed(1)||'-'}</td>
        <td style="font-size:11px;color:${(sd2.sales||0)>=45?'#15803d':(sd2.sales||0)<35?'#dc2626':'#475569'}">${sd2.sales?.toFixed(1)||'-'}</td>
      </tr>`;
    }).join('')}
    </tbody></table>
    <div class="caution">🟢 최고 · 🟡 중간 · 🔴 최저 기준 자동 판정. 소상공 편차 최대 ${E.smbGap.toFixed(1)}p — 하위 본부 집중 관리 필요.</div>
  </div>

  <!-- ═══ Ⅳ. 가입자 증감 현황 ═══ -->
  <div class="bs">
    <div class="bs-hd">
      <div class="bs-hd-num" style="background:#8b5cf6;color:#fff">Ⅳ</div>
      <div class="bs-hd-ttl">가입자 증감 현황</div>
      <div class="bs-hd-sub">무선 · H/C 구조 · 유선</div>
    </div>

    <table><thead><tr><th>구분</th><th>당월 재고</th><th>전월 증감</th><th>연초 누적</th><th>주요 비고</th></tr></thead><tbody>
      <tr>
        <td style="font-weight:700">무선 전체</td>
        <td style="font-weight:800">${fN6(wTot)}</td>
        <td>${fD6(wTot-wPrev,true)}</td>
        <td>-</td>
        <td>CAPA ${fN6(wCapa)} · 영업순증 ${fD6(wNetOp,true)}</td>
      </tr>
      <tr>
        <td style="font-weight:700">관리수수료(8.5%)</td>
        <td style="font-weight:800;color:#dc2626">${fN6(mgCur)}</td>
        <td>${fD6(mgCur-mgPrev,true)}</td>
        <td style="color:#dc2626">${mgYtdChg?mgYtdChg+'%':'-'}</td>
        <td>연초 ${fN6(mg1st)} → 구조적 감소 지속</td>
      </tr>
      <tr>
        <td style="font-weight:700">H/C(2%) 전체</td>
        <td style="font-weight:800">${fN6(hcCur)}</td>
        <td>${fD6(hcCur-hcPrev,true)}</td>
        <td>-</td>
        <td>도매 ${fD6(domaeNet,false)} / KT닷컴 ${fD6(ktdotNet,true)}</td>
      </tr>
      <tr>
        <td style="font-weight:700">유선(인터넷+TV)</td>
        <td style="font-weight:800">${fN6(wdB.total||0)}</td>
        <td>${fD6(vNetOp,true)}</td>
        <td>-</td>
        <td>인터넷 ${fN6(wdB.internet||0)} · TV ${fN6(wdB.tv||0)}</td>
      </tr>
    </tbody></table>

    <!-- H/C 역전 프로그레스 바 -->
    <div style="margin-top:12px;background:#fff7ed;border:1px solid #fed7aa;border-radius:10px;padding:12px">
      <div style="font-size:11px;font-weight:800;color:#c2410c;margin-bottom:8px">🚨 H/C 구조 전환 모니터링 — 도매 소멸 vs KT닷컴 성장</div>
      <div style="display:flex;justify-content:space-between;font-size:10px;color:#78350f;margin-bottom:5px">
        <span>도매H/C ${fN6(domaeStk)}</span><span>KT닷컴H/C ${fN6(ktdotStk)}</span>
      </div>
      <div style="height:14px;border-radius:7px;background:#dbeafe;overflow:hidden;position:relative">
        <div style="height:100%;border-radius:7px;background:linear-gradient(90deg,#ef4444,#f97316);width:${hcBarW}%;transition:width .3s"></div>
      </div>
      <div style="display:flex;justify-content:space-between;font-size:9px;color:#92400e;margin-top:4px">
        <span>도매 ${hcBarW}% · 월${fD6(domaeNet,false)}</span>
        <span>닷컴 ${(100-parseFloat(hcBarW)).toFixed(1)}% · 월${fD6(ktdotNet,true)}</span>
      </div>
      <div style="font-size:10px;color:#78350f;margin-top:8px;background:rgba(255,255,255,.6);padding:5px 8px;border-radius:6px">
        ※ H/C(2%) 수수료 단가는 관리수수료(8.5%) 대비 현저히 낮음. 역전 가속 시 수수료수입 구조 재검토 필요.
      </div>
    </div>
    <div class="caution">⚠️ 유선 12월 데이터는 가마감 상태로 이상 수치 확인됨 — 확정 마감 후 재확인 필요.</div>
  </div>

  <!-- ═══ Ⅴ. 수수료수입 분석 ═══ -->
  <div class="bs">
    <div class="bs-hd">
      <div class="bs-hd-num" style="background:#a855f7;color:#fff">Ⅴ</div>
      <div class="bs-hd-ttl">수수료수입 분석</div>
      <div class="bs-hd-sub">관리수수료 YoY · 추세 · 전망</div>
    </div>

    <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:12px">
      <div style="background:#faf5ff;border:1px solid #e9d5ff;border-radius:10px;padding:11px;border-top:3px solid #a855f7">
        <div style="font-size:9px;color:#6b21a8;font-weight:700;margin-bottom:4px">💰 총 수수료수입</div>
        <div style="font-size:19px;font-weight:900;color:#0f172a">${Number.isFinite(E.totalFee)?E.totalFee.toFixed(1)+'억':'N/A'}</div>
        <div style="font-size:9px;color:#64748b;margin-top:3px">관리비중 ${pctStr(E.mgmtMix)}</div>
      </div>
      <div style="background:#fff1f2;border:1px solid #fecdd3;border-radius:10px;padding:11px;border-top:3px solid #ef4444">
        <div style="font-size:9px;color:#9f1239;font-weight:700;margin-bottom:4px">🏷 관리수수료(8.5%)</div>
        <div style="font-size:19px;font-weight:900;color:#dc2626">${Number.isFinite(E.mgmtCur)?E.mgmtCur.toFixed(1)+'억':'N/A'}</div>
        <div style="font-size:9px;color:#dc2626;margin-top:3px">YoY ${E.mgmtYoY} · 평균대비 ${E.mgmtVsAvg}</div>
      </div>
      <div style="background:#eef2ff;border:1px solid #c7d2fe;border-radius:10px;padding:11px;border-top:3px solid #4f46e5">
        <div style="font-size:9px;color:#3730a3;font-weight:700;margin-bottom:4px">📋 무선 정책수수료</div>
        <div style="font-size:19px;font-weight:900;color:#312e81">${D.commission.jan26.policy_fee.total.toFixed(1)}억</div>
        <div style="font-size:9px;color:#4338ca;margin-top:3px">KT-RDS ${D.commission.jan26.policy_fee.rds}억 + 정책 ${D.commission.jan26.policy_fee.policy}억</div>
      </div>
    </div>
    <table><thead><tr><th>항목</th><th>값</th><th>전년 동월 대비</th><th>25년 평균 대비</th><th>평가</th></tr></thead><tbody>
      <tr><td style="font-weight:700">총 수수료수입</td><td>${Number.isFinite(E.totalFee)?E.totalFee.toFixed(1)+'억':'N/A'}</td><td>-</td><td>-</td><td>-</td></tr>
      <tr><td style="font-weight:700">관리수수료(8.5%)</td><td>${Number.isFinite(E.mgmtCur)?E.mgmtCur.toFixed(1)+'억':'N/A'}</td><td style="color:#dc2626">${E.mgmtYoY}</td><td>${E.mgmtVsAvg}</td><td>${E.mgmtYoY.includes('-')?'🔴 감소':'🟢 증가'}</td></tr>
      <tr><td style="font-weight:700">📋 무선 정책수수료</td><td>${D.commission.jan26.policy_fee.total.toFixed(1)}억</td><td style="color:#4338ca">${(((D.commission.jan26.policy_fee.total-D.commission.jan26.policy_fee_25jan)/D.commission.jan26.policy_fee_25jan)*100).toFixed(1)}% yoy</td><td>-</td><td>KT-RDS+정책 합산</td></tr>
      <tr><td style="padding-left:16px;color:#64748b">ㄴ KT-RDS</td><td>${D.commission.jan26.policy_fee.rds}억</td><td>-</td><td>-</td><td style="color:#64748b">소매·디지털 중심</td></tr>
      <tr><td style="padding-left:16px;color:#64748b">ㄴ 수수료수입 정책</td><td>${D.commission.jan26.policy_fee.policy}억</td><td>-</td><td>-</td><td style="color:#64748b">디지털 집중</td></tr>
      <tr><td>관리 비중</td><td>${pctStr(E.mgmtMix)}</td><td>-</td><td>-</td><td>${(E.mgmtMix||0)<0.15?'⚠️ 비중 축소':'🟢 정상'}</td></tr>
    </tbody></table>
  </div>

  <!-- ═══ Ⅵ. 사업과제 이행 현황 ═══ -->
  <div class="bs">
    <div class="bs-hd">
      <div class="bs-hd-num" style="background:#f59e0b;color:#fff">Ⅵ</div>
      <div class="bs-hd-ttl">사업과제 이행 현황</div>
      <div class="bs-hd-sub">123개 전체 · 총괄별 분해</div>
    </div>

    <!-- 전체 진척 바 -->
    <div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:12px;margin-bottom:12px">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px">
        <span style="font-size:12px;font-weight:800;color:#0f172a">전체 완료율</span>
        <span style="font-size:16px;font-weight:900;color:#f59e0b">${E.T.length?((E.done/E.T.length)*100).toFixed(1):'0.0'}%</span>
      </div>
      <div class="task-bar"><div class="task-fill" style="width:${E.T.length?((E.done/E.T.length)*100).toFixed(1):0}%"></div></div>
      <div style="display:flex;gap:12px;margin-top:6px;font-size:10px;color:#64748b">
        <span>✅ 완료 ${E.done}건</span>
        <span>🔄 진행중 ${E.inProgress}건</span>
        <span>🔴 0% 미착수 ${E.zeroProgress.length}건</span>
        <span>⚠️ 지연위험(20%↓) ${E.delayed.length}건</span>
      </div>
    </div>

    <table><thead><tr><th>총괄</th><th>전체</th><th>완료</th><th>완료율</th><th>0% 미착수</th><th>핵심과제 잔여</th></tr></thead><tbody>
      ${['경총','영총'].map(div=>{
        const bd=byDiv[div]||{tot:0,done:0,zero:0};
        const coreOpen=E.T.filter(t=>t.ch===div&&t.co==='●'&&t.st!=='완료').length;
        const pct=bd.tot?((bd.done/bd.tot)*100).toFixed(1):'0.0';
        return `<tr>
          <td style="font-weight:700">${div}</td>
          <td>${bd.tot}</td>
          <td style="color:#059669;font-weight:700">${bd.done}</td>
          <td><span style="background:${parseFloat(pct)>30?'#dcfce7':parseFloat(pct)>15?'#fef9c3':'#fee2e2'};color:${parseFloat(pct)>30?'#15803d':parseFloat(pct)>15?'#854d0e':'#dc2626'};font-size:10px;font-weight:800;padding:2px 8px;border-radius:8px">${pct}%</span></td>
          <td style="color:${bd.zero>3?'#dc2626':'#475569'};font-weight:${bd.zero>3?700:400}">${bd.zero}건</td>
          <td style="color:${coreOpen>5?'#dc2626':'#475569'}">${coreOpen}건</td>
        </tr>`;
      }).join('')}
    </tbody></table>

    ${E.delayed.length?`<div style="margin-top:10px;background:#fff1f2;border:1px solid #fecdd3;border-radius:8px;padding:10px">
      <div style="font-size:11px;font-weight:700;color:#9f1239;margin-bottom:6px">🔴 지연 위험 과제 Top 5 (추진도 20% 이하)</div>
      <table style="font-size:10.5px"><thead><tr><th>팀명</th><th>과제명</th><th>핵심</th><th>추진도</th></tr></thead><tbody>
      ${E.delayed.slice(0,5).map(t=>`<tr><td>${t.tm||'-'}</td><td>${(t.nm||t.title||'').slice(0,30)}${(t.nm||t.title||'').length>30?'…':''}</td><td>${t.co==='●'?'<span style="color:#dc2626;font-weight:800">●</span>':'-'}</td><td style="color:#dc2626;font-weight:700">${t.pp!==undefined?Math.round(t.pp*100)+'%':'-'}</td></tr>`).join('')}
      </tbody></table>
    </div>`:''}
  </div>

  <!-- ═══ Ⅶ. 리스크 등급 워치리스트 ═══ -->
  <div class="bs">
    <div class="bs-hd">
      <div class="bs-hd-num" style="background:#dc2626;color:#fff">Ⅶ</div>
      <div class="bs-hd-ttl">리스크 등급 워치리스트</div>
      <div class="bs-hd-sub">자동 감지 · ${risks.length}건</div>
    </div>

    ${risks.length?risks.map(r=>`
      <div class="risk-row">
        <span class="risk-badge" style="background:${r.lv==='HIGH'?'#fee2e2':'#fef9c3'};color:${r.col}">${r.lv}</span>
        <div>
          <div style="font-size:12px;font-weight:700;color:#0f172a">${r.item}</div>
          <div style="font-size:11px;color:#475569;margin-top:3px;line-height:1.6">${r.desc}</div>
        </div>
      </div>`).join('')
    :`<div style="text-align:center;color:#64748b;font-size:11px;padding:16px">현재 감지된 HIGH 리스크 없음 ✅</div>`}
  </div>

  <!-- ═══ Ⅷ. 전략적 실행 권고 ═══ -->
  <div class="bs">
    <div class="bs-hd">
      <div class="bs-hd-num" style="background:#0f172a;color:#fff">Ⅷ</div>
      <div class="bs-hd-ttl">전략적 실행 권고</div>
      <div class="bs-hd-sub">즉시 / 단기 / 중기 구분</div>
    </div>

    ${[{grp:'즉시 (이번 주)',col:'#dc2626',bg:'#fff1f2',num:'#dc2626',items:E.actions.filter(a=>a.priority==='HIGH')},
       {grp:'단기 (1개월)',col:'#b45309',bg:'#fffbeb',num:'#b45309',items:E.actions.filter(a=>a.priority==='MED')},
       {grp:'중기 (분기 내)',col:'#1d4ed8',bg:'#eff6ff',num:'#1d4ed8',items:[{item:'수수료수입 구조 재설계',owner:'경영기획/디지털강화',due:'2Q',detail:`H/C 역전 이후를 대비한 채널별 단가 구조 시뮬레이션 및 대안 수수료 체계 수립`},{item:'KPI 하위 본부 역량 강화 프로그램',owner:'본부장협의체',due:'2Q',detail:`소상공인/소매 MNP 편차 해소를 위한 본부별 맞춤형 코칭 체계 구축`}]}
    ].map(grp=>`
      <div style="margin-bottom:10px">
        <div style="font-size:11px;font-weight:800;color:${grp.col};background:${grp.bg};padding:6px 12px;border-radius:8px 8px 0 0;border:1px solid ${grp.col}20">${grp.grp} · ${grp.items.length}건</div>
        <div style="border:1px solid ${grp.col}20;border-top:none;border-radius:0 0 8px 8px;overflow:hidden">
        ${grp.items.map((a,i)=>`
          <div class="ac-row" style="background:${i%2===0?grp.bg:'#fff'}">
            <div class="ac-num" style="background:${grp.num}">${i+1}</div>
            <div style="flex:1">
              <div style="font-size:12px;font-weight:700;color:#0f172a">${a.item}</div>
              <div style="font-size:10px;color:#64748b;margin-top:2px">담당: ${a.owner} · 기한: ${a.due}</div>
              <div style="font-size:11px;color:#374151;margin-top:4px;line-height:1.6">${a.detail}</div>
            </div>
          </div>`).join('')}
        ${grp.items.length===0?'<div style="padding:10px 14px;font-size:11px;color:#94a3b8">해당 없음</div>':''}
        </div>
      </div>`).join('')}

    <div style="margin-top:12px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:10px;font-size:10px;color:#475569;line-height:1.7">
      A4 인쇄 최적화: 브라우저 인쇄(Ctrl+P) → PDF 저장으로 내부 경영보고 포맷으로 활용 가능합니다.<br>
      본 보고서는 자동 생성 기준(BM별 이익·KPI·과제·수수료·가입자 데이터)으로 작성되었습니다.
    </div>
  </div>

  <div class="footnote">자동 생성 · ${D.period} 기준 · KT M&S 경영기획팀 · ${E.missing.length?`누락(${E.missing.join(', ')})은 보수적으로 해석.`:'전체 데이터 정상 로드.'}</div>
  </div></body></html>`;
  openReport(h,'이달의 경영 브리핑');
}



// ② 사장님 보고용 요약 (핵심 의사결정 포인트만)
function genCeoBriefing(){
  const E = buildBriefingEngine();
  const T=E.T, P=E.P, K=E.K, SD=E.SD, C=E.C;
  const s = E.sub;
  const opm = (P.op.total/P.revenue.total*100).toFixed(1);
  const done = E.done;
  const zeroRisk = T.filter(t=>t.pp===0&&t.st==='진행중');
  const j = E.j;
  const sortedHq = [...(D.hq||[])].sort((a,b)=>(b.gp-b.sga)/b.rev-(a.gp-a.sga)/a.rev);
  const fN5 = v => fmtSubNum(Number(v)||0);
  const fDc2 = (v, goodIsPos=true) => {
    const n=Number(v)||0;
    const isGood = goodIsPos ? n>=0 : n<=0;
    return `<span style="color:${isGood?'#059669':'#dc2626'};font-weight:800">${n>=0?'+':''}${n.toLocaleString('ko-KR')}</span>`;
  };

  let h=`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>CEO 보고용 경영 요약</title>${rptStyle()}
  <style>
    .ceo-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:12px;margin-bottom:16px}
    .ceo-grid3{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:16px}
    .ceo-card{padding:14px;border-radius:12px;position:relative;overflow:hidden}
    .ceo-card-lbl{font-size:9px;font-weight:800;text-transform:uppercase;letter-spacing:.6px;margin-bottom:8px}
    .ceo-row{display:flex;justify-content:space-between;padding:7px 0;border-bottom:1px solid rgba(0,0,0,.06);font-size:12px;align-items:center}
    .ceo-row:last-child{border:none}
    .ceo-val{font-weight:800;font-size:13px}
    .ceo-action{border-left:4px solid;padding:10px 14px;border-radius:0 10px 10px 0;margin-bottom:9px;font-size:11.5px;line-height:1.75;background:#fff}
    .ceo-sig{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:6px;vertical-align:middle}
    .kv-row{display:flex;gap:10px;flex-wrap:wrap;margin-top:8px}
    .kv{background:rgba(255,255,255,.55);border-radius:8px;padding:6px 10px;font-size:11px}
    .kv .k{color:rgba(0,0,0,.5);font-size:9px;font-weight:700}
    .kv .v{font-weight:800;font-size:13px;color:#0f172a}
    .hc-bar{height:8px;border-radius:4px;background:#e2e8f0;overflow:hidden;margin-top:6px}
    .hc-fill{height:100%;background:linear-gradient(90deg,#f97316,#ef4444);border-radius:4px;transition:width .3s}
    @media print{.ceo-grid,.ceo-grid3{break-inside:avoid}}
    @media(max-width:720px){.ceo-grid,.ceo-grid3{grid-template-columns:1fr}}
  </style></head><body>
  <div class="cover" style="padding-bottom:30px">
    <div class="cover-tag" style="background:rgba(59,130,246,.15);color:#1d4ed8;font-weight:800">For CEO · Confidential · 1-Page Summary</div>
    <div class="cover-title" style="font-size:22px">경영 현황 보고</div>
    <div class="cover-sub">${D.period} | 경영기획팀</div>
    <div class="cover-meta">
      <div class="cover-kpi"><div class="cover-kpi-v">${fB(P.revenue.total)}억</div><div class="cover-kpi-l">매출</div></div>
      <div class="cover-kpi"><div class="cover-kpi-v">${opm}%</div><div class="cover-kpi-l">영업이익률</div></div>
      <div class="cover-kpi"><div class="cover-kpi-v">${fN5(s.wirelessTot)}</div><div class="cover-kpi-l">무선 가입자</div></div>
      <div class="cover-kpi"><div class="cover-kpi-v">${done}/${T.length}</div><div class="cover-kpi-l">과제 완료</div></div>
    </div>
  </div>
  <div class="page">

  <h1 style="font-size:16px;border-bottom:2px solid #1e293b;padding-bottom:8px;margin-bottom:14px">📊 한눈에 보는 경영 현황</h1>

  <!-- 경영 현황 체크포인트 -->
  <div style="background:linear-gradient(135deg,#fefce8,#fef9c3);border:1px solid #fde68a;border-radius:12px;padding:14px;margin-bottom:18px">
    <div style="font-size:12px;font-weight:900;color:#92400e;margin-bottom:8px">📌 ${D.period} 사장님 체크포인트</div>
    <div style="font-size:11.5px;color:#78350f;line-height:2.0">
      ${E.insights.map(ins=>`<div style="display:flex;gap:8px;align-items:flex-start"><span style="flex-shrink:0;color:#d97706">▶</span><span>${ins}</span></div>`).join('')}
    </div>
  </div>

  <h1 style="font-size:15px;font-weight:900;border-bottom:2px solid #0f172a;padding-bottom:8px;margin-bottom:14px">📊 6대 경영 지표 현황</h1>

  <!-- 3x2 카드 그리드 -->
  <div class="ceo-grid3">
    <!-- 손익 -->
    <div class="ceo-card" style="background:#f0fdf4;border:1px solid #bbf7d0;border-top:3px solid #10b981">
      <div class="ceo-card-lbl" style="color:#065f46">💰 BM별 손익</div>
      <div class="ceo-row"><span>매출</span><span class="ceo-val" style="color:#059669">${fB(P.revenue.total)}억</span></div>
      <div class="ceo-row"><span>영업이익</span><span class="ceo-val" style="color:#059669">${fB(P.op.total)}억 (${opm}%)</span></div>
      <div class="ceo-row"><span>최고 채널</span><span class="ceo-val">${E.bestProfit.name} ${pctStr(E.bestProfit.opm)}</span></div>
      <div class="ceo-row"><span>최저 채널</span><span class="ceo-val" style="color:#dc2626">${E.worstProfit.name} ${pctStr(E.worstProfit.opm)}</span></div>
    </div>
    <!-- KPI -->
    <div class="ceo-card" style="background:#eff6ff;border:1px solid #bfdbfe;border-top:3px solid #3b82f6">
      <div class="ceo-card-lbl" style="color:#1e40af">🎯 KPI 달성</div>
      <div class="ceo-row"><span>1위 본부</span><span class="ceo-val" style="color:#1d4ed8">${K[0]?.hq||'-'} ${K[0]?.ts?.toFixed(1)||'-'}점</span></div>
      <div class="ceo-row"><span>최하위</span><span class="ceo-val" style="color:#dc2626">${K[K.length-1]?.hq||'-'} ${K[K.length-1]?.ts?.toFixed(1)||'-'}점</span></div>
      <div class="ceo-row"><span>본부간 격차</span><span class="ceo-val">${E.kpiGap.toFixed(1)}p</span></div>
      <div class="ceo-row"><span>소상공 편차</span><span class="ceo-val" style="color:#b45309">${E.smbGap.toFixed(1)}p ⚠️</span></div>
    </div>
    <!-- 사업과제 -->
    <div class="ceo-card" style="background:#fefce8;border:1px solid #fde68a;border-top:3px solid #f59e0b">
      <div class="ceo-card-lbl" style="color:#92400e">📋 사업과제</div>
      <div class="ceo-row"><span>완료율</span><span class="ceo-val">${(done/T.length*100).toFixed(0)}% (${done}/${T.length})</span></div>
      <div class="ceo-row"><span>0% 미착수</span><span class="ceo-val" style="color:#dc2626">${zeroRisk.length}건</span></div>
      <div class="ceo-row"><span>핵심과제 잔여</span><span class="ceo-val" style="color:#b45309">${T.filter(t=>t.co==='●'&&t.st!=='완료').length}건</span></div>
      <div class="ceo-row"><span>지연위험(20%↓)</span><span class="ceo-val">${E.delayed.length}건</span></div>
    </div>
    <!-- 무선 가입자 -->
    <div class="ceo-card" style="background:#fdf4ff;border:1px solid #e9d5ff;border-top:3px solid #8b5cf6">
      <div class="ceo-card-lbl" style="color:#6b21a8">📡 무선 가입자</div>
      <div class="ceo-row"><span>전체 재고</span><span class="ceo-val">${fN5(s.wirelessTot)}</span></div>
      <div class="ceo-row"><span>전월 증감</span><span class="ceo-val">${fDc2(s.wirelessDelta,true)}</span></div>
      <div class="ceo-row"><span>CAPA</span><span class="ceo-val">${fN5(s.capa)}</span></div>
      <div class="ceo-row"><span>ARPU</span><span class="ceo-val">${s.arpu?Math.round(s.arpu).toLocaleString('ko-KR')+'원':'-'}${s.arpuChg?` (연초${s.arpuChg}%)`:''}</span></div>
    </div>
    <!-- 관리수수료·H/C -->
    <div class="ceo-card" style="background:#fff1f2;border:1px solid #fecdd3;border-top:3px solid #ef4444">
      <div class="ceo-card-lbl" style="color:#9f1239">🔴 관리수수료·H/C</div>
      <div class="ceo-row"><span>관리수수료 대상</span><span class="ceo-val" style="color:#dc2626">${fN5(s.mgmtCur)}</span></div>
      <div class="ceo-row"><span>연초 대비</span><span class="ceo-val" style="color:#dc2626">${s.mgmtChg||'-'}%</span></div>
      <div class="ceo-row"><span>도매H/C</span><span class="ceo-val">${fN5(s.domaeStock)} (월${fDc2(s.domaeNet,false)})</span></div>
      <div class="ceo-row"><span>KT닷컴H/C</span><span class="ceo-val" style="color:#2563eb">${fN5(s.ktdotStock)} (월${fDc2(s.ktdotNet,true)})</span></div>
    </div>
    <!-- 수수료수입 -->
    <div class="ceo-card" style="background:#faf5ff;border:1px solid #e9d5ff;border-top:3px solid #a855f7">
      <div class="ceo-card-lbl" style="color:#6b21a8">💴 수수료수입</div>
      <div class="ceo-row"><span>총 수수료수입</span><span class="ceo-val">${Number.isFinite(E.totalFee)?E.totalFee.toFixed(1)+'억':'N/A'}</span></div>
      <div class="ceo-row"><span>관리수수료</span><span class="ceo-val">${Number.isFinite(E.mgmtCur)?E.mgmtCur.toFixed(1)+'억':'N/A'}</span></div>
      <div class="ceo-row"><span style="color:#4338ca;font-weight:700">📋 무선 정책수수료</span><span class="ceo-val" style="color:#4338ca">${D.commission.jan26.policy_fee.total.toFixed(1)}억</span></div>
      <div class="ceo-row"><span style="font-size:9px;color:#64748b;padding-left:8px">ㄴ KT-RDS</span><span style="font-size:10px;color:#64748b">${D.commission.jan26.policy_fee.rds}억</span></div>
      <div class="ceo-row"><span>YoY (관리)</span><span class="ceo-val" style="color:#dc2626">${E.mgmtYoY}</span></div>
      <div class="ceo-row"><span>H/C 역전 격차</span><span class="ceo-val" style="color:${s.hcGap<100000?'#dc2626':'#b45309'}">${fN5(s.hcGap)}</span></div>
    </div>
  </div>

  <!-- H/C 역전 경보 바 -->
  <div style="background:linear-gradient(135deg,#fff7ed,#ffedd5);border:1px solid #fed7aa;border-radius:12px;padding:14px;margin-bottom:18px">
    <div style="font-size:11px;font-weight:800;color:#c2410c;margin-bottom:10px">🚨 H/C(2%) 구조 전환 경보 — 역전 ${s.hcGap<100000?'임박':'진행중'}</div>
    <div style="display:flex;justify-content:space-between;font-size:10px;color:#78350f;margin-bottom:6px">
      <span>도매H/C ${fN5(s.domaeStock)}</span><span>KT닷컴H/C ${fN5(s.ktdotStock)}</span>
    </div>
    <div style="height:12px;border-radius:6px;background:#dbeafe;overflow:hidden;position:relative">
      <div style="position:absolute;left:0;top:0;height:100%;border-radius:6px;background:linear-gradient(90deg,#dc2626,#f97316);width:${s.domaeStock+s.ktdotStock>0?(s.domaeStock/(s.domaeStock+s.ktdotStock)*100).toFixed(1):50}%"></div>
    </div>
    <div style="display:flex;justify-content:space-between;font-size:9px;color:#92400e;margin-top:4px">
      <span>도매 ${s.domaeStock+s.ktdotStock>0?(s.domaeStock/(s.domaeStock+s.ktdotStock)*100).toFixed(0):'-'}%</span>
      <span>닷컴 ${s.domaeStock+s.ktdotStock>0?(s.ktdotStock/(s.domaeStock+s.ktdotStock)*100).toFixed(0):'-'}%</span>
    </div>
    <div style="font-size:10px;color:#78350f;margin-top:8px;background:rgba(255,255,255,.5);padding:6px 10px;border-radius:8px">
      ※ H/C(2%) 단가는 관리수수료(8.5%) 대비 현저히 낮음. 역전 가속 시 수수료수입 구조 재검토 즉시 필요.
    </div>
  </div>

  <h1 style="font-size:15px;border-bottom:2px solid #dc2626;padding-bottom:8px;color:#dc2626;margin-bottom:14px">🚨 의사결정 필요 사항 (${E.actions.length}건)</h1>`;

  const actionStyles = {
    HIGH: {bg:'#fff1f2',border:'#ef4444',badge:'background:#fee2e2;color:#dc2626'},
    MED: {bg:'#fffbeb',border:'#f59e0b',badge:'background:#fef3c7;color:#b45309'}
  };
  let acNum=0;
  E.actions.forEach(a=>{
    acNum++;
    const st = actionStyles[a.priority]||actionStyles.MED;
    h+=`<div class="ceo-action" style="border-color:${st.border};background:${st.bg}">
      <span style="${st.badge};font-size:9px;font-weight:800;padding:2px 7px;border-radius:8px;margin-right:6px">${a.priority}</span>
      <strong>${acNum}. ${a.item}</strong> <span style="font-size:10px;color:#64748b;margin-left:6px">담당: ${a.owner} · ${a.due}</span><br>
      <span style="font-size:11px;color:#374151">${a.detail}</span>
    </div>`;
  });

  h+=`<h1 style="font-size:15px;border-bottom:2px solid #1e293b;padding-bottom:8px;margin-top:18px">🏢 본부별 수익성 · KPI 통합 현황</h1>
  <table>
    <thead><tr><th>수익순위</th><th>본부</th><th>매출</th><th>영업이익</th><th>OPM</th><th>KPI순위</th><th>KPI점수</th><th>판정</th></tr></thead>
    <tbody>`;
  sortedHq.forEach((hq,i)=>{
    const op=hq.gp-hq.sga;
    const opmV = hq.rev>0 ? op/hq.rev*100 : 0;
    const grade = opmV>=8?'✅ 우수':opmV>=4?'⚠️ 보통':'🔴 개선필요';
    const kpiMatch=K.find(k=>hq.nm.includes(k.hq.replace('영업본부','').replace('본부',''))||k.hq.includes(hq.nm.replace('영업본부','').replace('본부','')));
    h+=`<tr>
      <td style="text-align:center;font-weight:900;font-size:15px">${i+1}</td>
      <td style="font-weight:700">${hq.nm}</td>
      <td style="text-align:right">${fM(hq.rev)}</td>
      <td style="text-align:right;color:${op>0?'#059669':'#dc2626'};font-weight:700">${fM(op)}</td>
      <td style="text-align:right;font-weight:800;color:${opmV>5?'#059669':'#dc2626'}">${opmV.toFixed(1)}%</td>
      <td style="text-align:center;font-weight:700">${kpiMatch?kpiMatch.rk+'위':'-'}</td>
      <td style="text-align:center">${kpiMatch?kpiMatch.ts.toFixed(1):'- '}</td>
      <td style="text-align:center">${grade}</td>
    </tr>`;
  });
  h+=`</tbody></table>
  <div class="caution">※ 본부별 영업이익은 KPI 비율 추정치 · KPI = 소매60%+도매30%+소상공10% 가중합산</div>

  <div class="footnote" style="margin-top:20px">자동 생성 · ${D.period} 기준 · KT M&S 경영기획팀 · 대외 배포 금지 · Confidential</div>
  </div></body></html>`;
  openReport(h,'CEO 보고용 경영 요약');
}

// ③ 리스크 종합 워치리스트 (전 탭 리스크 통합)
function genRiskComprehensive(){
  const T=D.tasks,P=D.profit,K=D.kpi,SD=D.sga_detail,C=D.commission;
  const j=C.jan26;
  const zeroRisk=T.filter(t=>t.pp===0&&t.st==='진행중');
  const coreNotDone=T.filter(t=>t.co==='●'&&t.st!=='완료');
  const rd=C.retail_detail,wd=C.wholesale_detail;

  const risks=[
    {area:'손익',lvl:'HIGH',color:'#dc2626',item:'법인영업 수수료 역전 구조',
      metric:`판매수수료 ${fB(SD['판매수수료'].corporate_sales)}억 > 영업이익 ${fB(P.op.corporate_sales)}억`,
      impact:'채널 OPM 0.6% — 매출 성장 없이는 적자 전환 위험',
      action:'수수료율 재협의 또는 직영 전환 검토 (1Q 내)'},
    {area:'수수료',lvl:'HIGH',color:'#dc2626',item:'도매 수수료 급감',
      metric:`24년 ${C.yearly[2].wholesale}억 → 26년(E) ${C.yearly[4].wholesale}억 (△${((C.yearly[2].wholesale-C.yearly[4].wholesale)/C.yearly[2].wholesale*100).toFixed(0)}%)`,
      impact:'점프업 축소 + 무선 가입자 감소 → 수수료 기반 와해',
      action:'Win-Win HC 확대, 도매 구조 재편 TF (1Q)'},
    {area:'수수료',lvl:'HIGH',color:'#dc2626',item:'소매 무선 수수료 하락 추세',
      metric:`25년 1→5월: ${rd.w25[0]} → ${rd.w25[4]}억 (월별 감소)`,
      impact:'연간 6~10억 수익 감소 가능성',
      action:'고단가 요금제 전환 인센티브 강화 (2Q)'},
    {area:'정책수수료',lvl:'MED',color:'#4f46e5',item:'무선 정책수수료 채널 편중',
      metric:`KT-RDS ${j.policy_fee.rds}억 + 정책 ${j.policy_fee.policy}억 = 합산 ${j.policy_fee.total}억 | 소매·디지털 집중도 ${((j.policy_fee.retail+j.policy_fee.digital)/j.policy_fee.total*100).toFixed(0)}%`,
      impact:'법인영업·기업/공공 정책 차감(-) 구조 지속 시 해당 채널 수익성 왜곡',
      action:'KT-RDS 단가 구조 재검토, 정책 차감 채널 원인 분석 (1Q)'},
    {area:'정책수수료',lvl:'MED',color:'#4f46e5',item:'무선 정책수수료 YoY 변동 모니터링',
      metric:`26년 1월 ${j.policy_fee.total}억 vs 전년동월 ${j.policy_fee_25jan}억 → YoY ${(((j.policy_fee.total-j.policy_fee_25jan)/j.policy_fee_25jan)*100).toFixed(1)}%`,
      impact:'KT-RDS 단가 하락 추세 이어질 경우 수수료수입 전체 기반 약화',
      action:'월별 추이 모니터링 체계 수립, KT 본사와 단가 협의 (2Q)'},
    {area:'KPI',lvl:'HIGH',color:'#dc2626',item:'소상공인 KPI 본부간 극심한 편차',
      metric:`${K[0].hq}(${K[0].sm.t.toFixed(1)}) vs ${K[4].hq}(${K[4].sm.t.toFixed(1)}) — ${(K[0].sm.t-K[4].sm.t).toFixed(1)}p`,
      impact:'하위 본부 소상공인 영업력 붕괴 우려',
      action:'강서 모델 이식 TF 구성, 커버리지 확대 (즉시)'},
    {area:'비용',lvl:'MED',color:'#f59e0b',item:'디지털 판촉비 집중',
      metric:`판촉비 ${fB(SD['판매촉진비'].digital)}억 — 디지털 판관비의 ${(SD['판매촉진비'].digital/P.sga.digital*100).toFixed(0)}%`,
      impact:`영업이익(${fB(P.op.digital)}억) 대비 판촉비가 ${(SD['판매촉진비'].digital/P.op.digital).toFixed(1)}배`,
      action:'성과 미달 판촉 프로그램 선별 중단 (2Q)'},
    {area:'비용',lvl:'MED',color:'#f59e0b',item:'소매 인건비 부담',
      metric:`소매 인건비 ${fB(SD['인건비합계'].retail)}억 — 판관비의 ${(SD['인건비합계'].retail/P.sga.retail*100).toFixed(1)}%`,
      impact:'고정비 비중 높아 매출 하락 시 급격한 수익성 악화',
      action:'생산성 향상 프로그램 연계 인력 효율화 (중기)'},
    {area:'수수료',lvl:'MED',color:'#f59e0b',item:'디지털 무선→H.C 전환 리스크',
      metric:`디지털 무선: 25년 1→5월 ${C.digital_detail.wireless[0]} → ${C.digital_detail.wireless[4]}억`,
      impact:'kt닷컴 위탁 전환 후 무선 수수료 급감, H.C 보상 미확인',
      action:'H.C 수익성 검증, 장기 계약 조건 재검토 (2Q)'},
    {area:'과제',lvl:'MED',color:'#f59e0b',item:`진행중 & 추진도 0% 과제 ${zeroRisk.length}건`,
      metric:`핵심과제 ${zeroRisk.filter(t=>t.co==='●').length}건 포함 — ${zeroRisk.slice(0,3).map(t=>t.tm).join(', ')} 등`,
      impact:'1Q 종료 시 과제 이행률 저조 → 경영평가 리스크',
      action:'주간 점검회의 신설, 담당 팀장 소명 (즉시)'},
    {area:'KPI',lvl:'LOW',color:'#22c55e',item:'소매 KPI 본부간 편차 안정적',
      metric:`1위 ${K[0].hq}(${K[0].rt.t.toFixed(1)}) vs 5위 ${K[4].hq}(${K[4].rt.t.toFixed(1)}) — ${(K[0].rt.t-K[4].rt.t).toFixed(1)}p`,
      impact:'채널 가중치 60%의 소매가 안정적이어서 전체 KPI 기반 유지',
      action:'현 수준 유지 관리, 하위 본부 소폭 보완'}
  ];

  let h=`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>리스크 종합 워치리스트</title>${rptStyle()}
  <style>
    .risk-card{border-left:4px solid;padding:12px 14px;background:#fff;border-radius:0 8px 8px 0;margin-bottom:10px;page-break-inside:avoid}
    .risk-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:6px}
    .risk-title{font-weight:700;font-size:12px;color:#1e293b}
    .risk-badge{font-size:9px;padding:2px 8px;border-radius:10px;font-weight:700}
    .risk-area{font-size:9px;background:#e0f2fe;color:#0369a1;padding:2px 6px;border-radius:8px;margin-right:6px}
    .risk-row{font-size:11px;color:#475569;margin-bottom:3px;line-height:1.5}
    .risk-row b{color:#1e293b}
  </style></head><body>
  <div class="cover" style="background:linear-gradient(135deg,#450a0a,#991b1b)">
    <div class="cover-tag" style="background:rgba(255,255,255,.15);color:#fca5a5">Risk Watch · Confidential</div>
    <div class="cover-title" style="color:#fff">리스크 종합 워치리스트</div>
    <div class="cover-sub" style="color:#fecaca">${D.period} 기준 | 손익·KPI·수수료·과제 전영역 통합</div>
    <div class="cover-meta">
      <div class="cover-kpi"><div class="cover-kpi-v" style="color:#fca5a5">${risks.filter(r=>r.lvl==='HIGH').length}</div><div class="cover-kpi-l" style="color:#fecaca">HIGH</div></div>
      <div class="cover-kpi"><div class="cover-kpi-v" style="color:#fde68a">${risks.filter(r=>r.lvl==='MED').length}</div><div class="cover-kpi-l" style="color:#fef3c7">MED</div></div>
      <div class="cover-kpi"><div class="cover-kpi-v" style="color:#86efac">${risks.filter(r=>r.lvl==='LOW').length}</div><div class="cover-kpi-l" style="color:#dcfce7">LOW</div></div>
    </div>
  </div>
  <div class="page">

  <h1>Ⅰ. 리스크 총괄 현황</h1>
  <div class="exec-box">
    총 <strong>${risks.length}건</strong>의 리스크를 손익·수수료·KPI·비용·과제 전 영역에서 식별하였습니다.
    HIGH 등급 <strong style="color:#dc2626">${risks.filter(r=>r.lvl==='HIGH').length}건</strong>은 1Q 내 즉시 대응이 필요하며,
    MED 등급 ${risks.filter(r=>r.lvl==='MED').length}건은 2Q까지 모니터링 및 개선 계획 수립이 권고됩니다.
  </div>

  <h1 style="color:#dc2626">Ⅱ. HIGH 리스크 (즉시 대응)</h1>`;
  risks.filter(r=>r.lvl==='HIGH').forEach(r=>{
    h+=`<div class="risk-card" style="border-color:${r.color}">
      <div class="risk-header"><div><span class="risk-area">${r.area}</span><span class="risk-title">${r.item}</span></div><span class="risk-badge" style="background:#fee2e2;color:#dc2626">HIGH</span></div>
      <div class="risk-row"><b>📊 현황:</b> ${r.metric}</div>
      <div class="risk-row"><b>⚡ 영향:</b> ${r.impact}</div>
      <div class="risk-row" style="color:#059669"><b>✅ 대응:</b> ${r.action}</div>
    </div>`;
  });

  h+=`<h1 style="color:#f59e0b">Ⅲ. MED 리스크 (모니터링)</h1>`;
  risks.filter(r=>r.lvl==='MED').forEach(r=>{
    h+=`<div class="risk-card" style="border-color:${r.color}">
      <div class="risk-header"><div><span class="risk-area">${r.area}</span><span class="risk-title">${r.item}</span></div><span class="risk-badge" style="background:#fef3c7;color:#b45309">MED</span></div>
      <div class="risk-row"><b>📊 현황:</b> ${r.metric}</div>
      <div class="risk-row"><b>⚡ 영향:</b> ${r.impact}</div>
      <div class="risk-row" style="color:#059669"><b>✅ 대응:</b> ${r.action}</div>
    </div>`;
  });

  h+=`<h1 style="color:#22c55e">Ⅳ. LOW 리스크 (유지 관리)</h1>`;
  risks.filter(r=>r.lvl==='LOW').forEach(r=>{
    h+=`<div class="risk-card" style="border-color:${r.color}">
      <div class="risk-header"><div><span class="risk-area">${r.area}</span><span class="risk-title">${r.item}</span></div><span class="risk-badge" style="background:#dcfce7;color:#16a34a">LOW</span></div>
      <div class="risk-row"><b>📊 현황:</b> ${r.metric}</div>
      <div class="risk-row"><b>⚡ 영향:</b> ${r.impact}</div>
      <div class="risk-row" style="color:#059669"><b>✅ 대응:</b> ${r.action}</div>
    </div>`;
  });

  h+=`<h1>Ⅴ. 리스크 대응 로드맵</h1>
  <table>
    <thead><tr><th>시기</th><th>대응 항목</th><th>담당 영역</th><th>기대 효과</th></tr></thead>
    <tbody>
      <tr><td><span class="badge badge-red">즉시</span></td><td>법인영업 수수료 재협의</td><td>손익</td><td>영업이익 +${(SD['판매수수료'].corporate_sales*0.3/1e8).toFixed(1)}억</td></tr>
      <tr><td><span class="badge badge-red">즉시</span></td><td>소상공인 KPI 하위 본부 TF</td><td>KPI</td><td>편차 10p 이내 축소</td></tr>
      <tr><td><span class="badge badge-red">1Q</span></td><td>도매 수수료 구조 전환 전략</td><td>수수료</td><td>수수료 감소폭 50% 방어</td></tr>
      <tr><td><span class="badge badge-red">1Q</span></td><td>무선 정책수수료 채널 편중 원인 분석</td><td>정책수수료</td><td>법인·기업 차감 구조 개선, KT-RDS 단가 재협의</td></tr>
      <tr><td><span class="badge badge-o">2Q</span></td><td>디지털 판촉비 ROI 검토</td><td>비용</td><td>판촉비 10~15% 효율화</td></tr>
      <tr><td><span class="badge badge-o">2Q</span></td><td>소매 무선 수수료 방어</td><td>수수료</td><td>월간 수수료 20억 이상 유지</td></tr>
      <tr><td><span class="badge badge-o">2Q</span></td><td>KT-RDS 수수료 YoY 방어</td><td>정책수수료</td><td>전년 대비 감소폭 최소화, 월별 추이 대시보드 운영</td></tr>
      <tr><td><span class="badge badge-y">중기</span></td><td>소매 인건비 효율화</td><td>비용</td><td>생산성 연계 인력 최적화</td></tr>
    </tbody>
  </table>

  <hr class="section-divider">
  <div class="footnote">본 워치리스트는 사업과제·BM별손익·조직KPI·수수료수입 전 영역 데이터를 종합하여 자동 생성되었습니다.</div>
  </div></body></html>`;
  openReport(h,'리스크 종합 워치리스트');
}

// ==================== AI 경영현황 요약 보고서 ====================
function genAISummaryReport(){
  const T=D.tasks,P=D.profit,K=D.kpi;
  const done=T.filter(t=>t.st==='완료').length,prog=T.filter(t=>t.st==='진행중').length;
  const zeroRisk=T.filter(t=>t.pp===0&&t.st==='진행중');
  const coreNotDone=T.filter(t=>t.co==='●'&&t.st!=='완료');
  const K1=K[0],K5=K[4];
  let h=`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>경영현황 Executive 보고서</title>${rptStyle()}</head><body>
  <div class="cover">
    <div class="cover-tag">Confidential · For Executive Review Only</div>
    <div class="cover-title">KT M&S 경영현황 Executive 보고서</div>
    <div class="cover-sub">${D.period} 기준 | 경영기획팀 AI 분석</div>
    <div class="cover-kpis">
      <div class="cover-kpi"><div class="cover-kpi-v" style="color:#34d399">${fB(P.revenue.total)}억</div><div class="cover-kpi-l">매출</div></div>
      <div class="cover-kpi"><div class="cover-kpi-v" style="color:#38bdf8">${fB(P.op.total)}억</div><div class="cover-kpi-l">영업이익 ${(P.op.total/P.revenue.total*100).toFixed(1)}%</div></div>
      <div class="cover-kpi"><div class="cover-kpi-v" style="color:#fbbf24">${K1.ts.toFixed(1)}점</div><div class="cover-kpi-l">KPI 1위 ${K1.hq}</div></div>
    </div>
  </div>

  <h1>Ⅰ. 이달의 핵심 지표</h1>
  <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-bottom:16px">
    ${[{l:'매출',v:fB(P.revenue.total)+'억',c:'#0369a1'},{l:'영업이익',v:fB(P.op.total)+'억',c:'#059669'},{l:'OPM',v:(P.op.total/P.revenue.total*100).toFixed(1)+'%',c:'#7c3aed'},{l:'순이익',v:fB(P.net_income)+'억',c:'#c2410c'}].map(x=>`<div style="padding:12px;background:#f8fafc;border-radius:8px;text-align:center"><div style="font-size:10px;color:#64748b;font-weight:600;margin-bottom:4px">${x.l}</div><div style="font-size:18px;font-weight:800;color:${x.c}">${x.v}</div></div>`).join('')}
  </div>

  <h2>손익 핵심 포인트</h2>
  <table><thead><tr><th>구분</th><th class="num">금액(백만)</th><th class="num">비율</th><th>판단</th></tr></thead><tbody>
  <tr><td>매출</td><td class="num">${fM(P.revenue.total)}</td><td class="num">100%</td><td>-</td></tr>
  <tr><td>매출총이익</td><td class="num pos">${fM(P.gross.total)}</td><td class="num">${(P.gross.total/P.revenue.total*100).toFixed(1)}%</td><td>-</td></tr>
  <tr><td>판관비</td><td class="num neg">${fM(P.sga.total)}</td><td class="num">${(P.sga.total/P.revenue.total*100).toFixed(1)}%</td><td style="color:#f59e0b">인건비 집중(${(D.sga_detail['인건비합계'].total/P.sga.total*100).toFixed(0)}%) 모니터링 필요</td></tr>
  <tr class="tr"><td><b>영업이익</b></td><td class="num pos" style="font-weight:700">${fM(P.op.total)}</td><td class="num" style="font-weight:700">${(P.op.total/P.revenue.total*100).toFixed(1)}%</td><td>-</td></tr>
  </tbody></table>

  <h1>Ⅱ. 채널별 수익성</h1>
  <div class="insight">수익성 TOP 3: 소상공인(${(P.op.small_biz/P.revenue.small_biz*100).toFixed(1)}%) → IoT(${(P.op.iot/P.revenue.iot*100).toFixed(1)}%) → 소매(${(P.op.retail/P.revenue.retail*100).toFixed(1)}%)</div>
  <div class="warn">⚠️ 법인영업 OPM 0.6% — 판매수수료(${fB(D.sga_detail['판매수수료'].corporate_sales)}억)가 영업이익(${fB(P.op.corporate_sales)}억) 초과. 구조 개선 시급</div>

  <h1>Ⅲ. KPI 현황</h1>
  <table><thead><tr><th>순위</th><th>본부</th><th class="num">종합</th><th class="num">소매</th><th class="num">도매</th><th class="num">소상공인</th></tr></thead><tbody>
  ${K.map(k=>`<tr><td style="text-align:center;font-weight:800">${k.rk}</td><td>${k.hq}</td><td class="num" style="font-weight:700">${k.ts.toFixed(1)}</td><td class="num">${k.rt.t.toFixed(1)}</td><td class="num">${k.wh.t.toFixed(1)}</td><td class="num ${k.sm.t<55?'neg':''}">${k.sm.t.toFixed(1)}</td></tr>`).join('')}
  </tbody></table>
  <div class="warn">⚠️ 소상공인 편차: 1위 강서(71.95) vs 5위 강남(49.01) — 22.9p 격차 즉시 조치 필요</div>

  <h1>Ⅳ. 과제 이행 현황</h1>
  <div class="exec-box">총 ${T.length}개 과제 중 완료 ${done}개(${(done/T.length*100).toFixed(0)}%) · 진행중 ${prog}개 · 핵심과제 미완료 ${coreNotDone.length}개</div>
  ${zeroRisk.length>0?`<div class="warn">⚠️ 진행중 & 추진도 0% 과제 ${zeroRisk.length}개 — ${zeroRisk.slice(0,3).map(t=>t.tm+' '+t.nm).join(', ')} 등 즉시 점검</div>`:''}

  <h1>Ⅴ. 전략적 권고안</h1>
  <div class="action">
  ① <b>[손익] 법인영업 수수료 구조 재검토</b> — 수수료(${fB(D.sga_detail['판매수수료'].corporate_sales)}억) 절감 또는 매출 성장 없이는 채널 수익성 개선 불가<br><br>
  ② <b>[KPI] 강남·동부 소상공인 집중 지원</b> — 강서 모델 이식, 커버리지 확대 TF 구성 (1개월 내)<br><br>
  ③ <b>[과제] 저추진 핵심과제 점검 회의</b> — ${zeroRisk.filter(t=>t.co==='●').length}개 핵심과제 포함 0% 과제 ${zeroRisk.length}개 집중 점검<br><br>
  ④ <b>[비용] 디지털 판촉비 효율성 검토</b> — 판촉비 ${fB(D.sga_detail['판매촉진비'].digital)}억(채널판관비의 59%) 대비 OPM 2.8% — ROI 재검토 필요
  </div>
  </body></html>`;
  openReport(h,'경영현황 Executive 보고서');
}

function quickAI(q){document.getElementById('aiIn').value=q;sendAI();}
async function genThreeYearBusinessPlanAI(){
  const P=D.profit||{},S=D.subscriber||{},C=D.commission||{};
  const rev=P.revenue?.total||0,op=P.op?.total||0,sga=P.sga?.total||0;
  const wirelessNow=S?.current?.wireless_total||S?.wireless?.total?.slice?.(-1)?.[0]||0;
  const capaNow=S?.current?.wireless_capa||S?.wireless?.capa?.slice?.(-1)?.[0]||0;
  const mgmtFeeNow=C?.summary?.management_fee||0;
  const opm=rev?((op/rev)*100):0;
  const key=document.getElementById('apiKey').value.trim();
  const mdl=document.getElementById('aiMdl').value;
  if(!key){addMsg('s','⚠️ API Key를 입력하고 연결해주세요.');return;}

  const prompt=`[3개년 사업계획 리포트 생성 요청]\n\n`
  +`아래 내부 지표를 기준으로, 경영진 제출용 3개년 사업계획서를 작성해줘.\n`
  +`중요: 최종 응답은 반드시 순수 HTML만 출력하고(설명문/코드블록 금지), <html>부터 </html>까지 완결된 문서로 작성해줘.\n`
  +`중요: 표(table), 핵심 KPI 카드, 리스크 매트릭스를 반드시 포함해줘.\n\n`
  +`현재 내부 기준 데이터(대시보드):\n`
  +`- 전사 매출: ${fB(rev)}억\n`
  +`- 전사 영업이익: ${fB(op)}억 (OPM ${opm.toFixed(1)}%)\n`
  +`- 판관비: ${fB(sga)}억\n`
  +`- 무선 가입자: ${Math.round(wirelessNow).toLocaleString('ko-KR')}\n`
  +`- CAPA(신규+기변): ${Math.round(capaNow).toLocaleString('ko-KR')}\n`
  +`- 관리수수료: ${fB(mgmtFeeNow)}억\n\n`
  +`필수 구성:\n`
  +`1) Executive Summary(1 page)\n`
  +`2) 핵심 가정(시장성장률/CAPA/수수료율/ARPU/비용상승률)\n`
  +`3) 3개년 재무계획 표(매출/원가/매출총이익/판관비 세부/영업이익/순이익)\n`
  +`4) 3개년 운영계획 표(가입자/CAPA/관리수수료/서비스매출/채널믹스)\n`
  +`5) 시나리오 3종(보수/기준/공격) + 민감도(수수료율 ±1%p, CAPA ±10%)\n`
  +`6) 타사 벤치마크 3~5개(회사명/핵심지표/시사점/출처 URL). 사실과 추정 구분\n`
  +`7) 실행 로드맵(월·분기), KPI, 담당조직\n`
  +`8) 리스크·대응(재무/시장/조직/규제) + 경영진 의사결정 체크리스트 10개\n\n`
  +`표현 가이드:\n`
  +`- 숫자는 계산 근거를 본문 또는 각주에 제시\n`
  +`- 보기 쉽게 카드/표/강조 박스 사용\n`
  +`- 한국어로 작성\n`
  +`- 외부 사례는 출처 URL을 실제 링크(a 태그)로 표기`;

  const lid='bp'+Date.now();
  document.getElementById('aiMsg').innerHTML+=`<div class="msg a" id="${lid}">🧠 3개년 사업계획 리포트 생성 중...</div>`;
  document.getElementById('aiMsg').scrollTop=9999;
  document.getElementById('aiBtn').disabled=true;
  try{
    const res=await fetch('https://api.anthropic.com/v1/messages',{
      method:'POST',
      headers:{
        'Content-Type':'application/json',
        'x-api-key':key,
        'anthropic-version':'2023-06-01',
        'anthropic-dangerous-direct-browser-access':'true'
      },
      body:JSON.stringify({
        model:mdl,
        max_tokens:8192,
        system:buildCtx(),
        messages:[{role:'user',content:prompt}]
      })
    });
    const data=await res.json();
    const html=(data.content||[]).map(c=>c.text||'').join('\n').trim();
    if(!res.ok || !html){
      const errTxt=JSON.stringify(data.error||data);
      document.getElementById(lid).innerHTML=`<span style="color:var(--r)">❌ 리포트 생성 실패: ${errTxt}</span>`;
      return;
    }
    document.getElementById(lid).textContent='✅ 리포트 생성 완료 · 보고서 뷰어를 엽니다.';
    openReport(html,'3개년 사업계획 AI 리포트');
  }catch(err){
    document.getElementById(lid).innerHTML=`<span style="color:var(--r)">❌ ${err.message}</span>`;
  }finally{
    document.getElementById('aiBtn').disabled=false;
    document.getElementById('aiMsg').scrollTop=9999;
  }
}
async function sendAI(){
  const inp=document.getElementById('aiIn'),msg=inp.value.trim();if(!msg)return;
  const key=document.getElementById('apiKey').value.trim();
  if(!key){addMsg('s','⚠️ API Key를 입력하고 연결해주세요.');return;}
  const mdl=document.getElementById('aiMdl').value;
  addMsg('u',msg);inp.value='';
  const lid='l'+Date.now();
  document.getElementById('aiMsg').innerHTML+=`<div class="msg a" id="${lid}">🔄 분석 중...</div>`;
  document.getElementById('aiMsg').scrollTop=9999;document.getElementById('aiBtn').disabled=true;
  try{
    const res=await fetch('https://api.anthropic.com/v1/messages',{method:'POST',headers:{'Content-Type':'application/json','x-api-key':key,'anthropic-version':'2023-06-01','anthropic-dangerous-direct-browser-access':'true'},body:JSON.stringify({model:mdl,max_tokens:4096,system:buildCtx(),messages:[{role:'user',content:msg}]})});
    const data=await res.json();document.getElementById(lid).textContent=data.content?.map(c=>c.text||'').join('\n')||JSON.stringify(data.error||data);
  }catch(err){document.getElementById(lid).innerHTML=`<span style="color:var(--r)">❌ ${err.message}</span>`;}
  document.getElementById('aiBtn').disabled=false;document.getElementById('aiMsg').scrollTop=9999;
}
document.getElementById('aiIn').onkeydown=e=>{if(e.key==='Enter'&&!e.shiftKey){e.preventDefault();sendAI();}};

// ==================== UPLOAD ====================


// ==================== (data.js에서 로드됨: DEFAULT_SUBSCRIBER_DATA, subscriberData) ====================

function fmtSubNum(v){ const n=Math.round(Number(v)||0); return n.toLocaleString('ko-KR'); }
function fmtSubDiff(v){ const n=Number(v)||0; return (n>0?'+':'')+n.toLocaleString('ko-KR'); }
function safePct(a,b){ return b?((a/b)*100):0; }
function fmtWon(v){ return Math.round(Number(v)||0).toLocaleString('ko-KR')+'원'; }
function cloneObj(v){ return JSON.parse(JSON.stringify(v)); }

function getSavedSubscriberData(){
  try{ const raw=localStorage.getItem('ktms_subscriber_data'); return raw?JSON.parse(raw):null; }catch(e){ return null; }
}
function saveSubscriberData(){
  try{ localStorage.setItem('ktms_subscriber_data', JSON.stringify(subscriberData)); }catch(e){}
}
function loadSubscriberData(){
  const saved=getSavedSubscriberData();
  const defVer = DEFAULT_SUBSCRIBER_DATA._dataVersion||0;
  if(saved && saved.months && saved.wireless && saved._dataVersion===defVer){
    subscriberData = saved;
  } else {
    subscriberData = cloneObj(DEFAULT_SUBSCRIBER_DATA);
    delete subscriberData.subscriberSchema;
    saveSubscriberData();
  }
}
function resetSubscriberData(){
  subscriberData = cloneObj(DEFAULT_SUBSCRIBER_DATA);
  saveSubscriberData();
  initSubscriberUI();
  showUpStatus('factbook','ok','✅ 기본 가입자 데이터로 복원했습니다.');
}

function pickCell(ws, r, c){
  const ref = XLSX.utils.encode_cell({r:r,c:c});
  const cell = ws[ref];
  return cell ? cell.v : null;
}
function pickSeries(ws, row, startCol, endCol){
  const arr = [];
  for(let c=startCol;c<=endCol;c++) arr.push(Number(pickCell(ws,row,c))||0);
  return arr;
}
function detectMonthLabels(ws, row, startCol, endCol){
  const arr = [];
  for(let c=startCol;c<=endCol;c++) {
    const val = pickCell(ws,row,c);
    arr.push(String(val||'').replace(/'/g,'').trim());
  }
  return arr;
}

function extractSubscriberDataFromWorkbook(wb){
  const wirelessName = wb.SheetNames.find(s=>String(s).includes('2.무선가입자'));
  const wiredName = wb.SheetNames.find(s=>String(s).includes('3.유선가입자'));
  const orgName = wb.SheetNames.find(s=>String(s).includes('영기팀 작성_가입자'));
  const ws21Name = wb.SheetNames.find(s=>String(s).includes('2-(1)'));
  if(!wirelessName || !wiredName) throw new Error('가입자 핵심 시트를 찾지 못했습니다.');
  const wsW = wb.Sheets[wirelessName];
  const wsV = wb.Sheets[wiredName];
  const wsY = orgName ? wb.Sheets[orgName] : null;
  const ws21 = ws21Name ? wb.Sheets[ws21Name] : null;

  // ── 2.무선가입자: 헤더 행6(xlsx_r=5), 월별 데이터 col6~17(xlsx c=5~16) ──
  const months = detectMonthLabels(wsW, 5, 5, 16);
  // 일반후불(전체) 섹션 (Excel행 7~16)
  const wirelessTotal    = pickSeries(wsW,  6, 5, 16); // 행7 유지
  const wirelessCapa     = pickSeries(wsW,  7, 5, 16); // 행8 CAPA
  const wirelessNew      = pickSeries(wsW,  8, 5, 16); // 행9 신규
  const wirelessChg      = pickSeries(wsW,  9, 5, 16); // 행10 기변
  const wirelessCancel   = pickSeries(wsW, 10, 5, 16); // 행11 해지
  const wirelessChgOut   = pickSeries(wsW, 12, 5, 16); // 행13 기변OUT
  const wirelessMaturity = pickSeries(wsW, 13, 5, 16); // 행14 만기환수
  const wirelessNetRef   = pickSeries(wsW, 14, 5, 16); // 행15 영업순증
  // 관리수수료(8.5%) 섹션 (Excel행 17~26)
  const wirelessMgmtEligible = pickSeries(wsW, 16, 5, 16); // 행17 유지
  const mgmtCapa     = pickSeries(wsW, 17, 5, 16); // 행18 CAPA
  const mgmtChurn    = pickSeries(wsW, 20, 5, 16); // 행21 해지
  const mgmtChgOut   = pickSeries(wsW, 22, 5, 16); // 행23 기변OUT
  const mgmtMaturity = pickSeries(wsW, 23, 5, 16); // 행24 만기환수
  const mgmtTransfer = pickSeries(wsW, 24, 5, 16); // 행25 가입자이관
  const mgmtNetAdd   = pickSeries(wsW, 25, 5, 16); // 행26 순증
  // H/C(2%) 섹션 (Excel행 27~36)
  const wirelessHcEligible = pickSeries(wsW, 26, 5, 16); // 행27 유지
  const hcCapa     = pickSeries(wsW, 27, 5, 16); // 행28 CAPA
  const hcChurn    = pickSeries(wsW, 30, 5, 16); // 행31 해지
  const hcChgOut   = pickSeries(wsW, 32, 5, 16); // 행33 기변OUT
  const hcMaturity = pickSeries(wsW, 33, 5, 16); // 행34 만기환수
  const hcTransfer = pickSeries(wsW, 34, 5, 16); // 행35 가입자이관
  const hcNetAdd   = pickSeries(wsW, 35, 5, 16); // 행36 순증

  // ── 3.유선가입자: 동일 col구조(c=5~16) ──
  const wiredTotal    = pickSeries(wsV,  6, 5, 16); // 행7 유지
  const internet      = pickSeries(wsV, 13, 5, 16); // 행14 인터넷유지
  const tv            = pickSeries(wsV, 20, 5, 16); // 행21 TV유지
  const wiredNew      = pickSeries(wsV,  7, 5, 16); // 행8 신규
  const wiredCancel   = pickSeries(wsV,  8, 5, 16); // 행9 해지
  const wiredMaturity = pickSeries(wsV, 11, 5, 16); // 행12 만기환수
  const wiredNet      = pickSeries(wsV, 10, 5, 16); // 행11 영업순증
  const wiredNetPure  = pickSeries(wsV, 12, 5, 16); // 행13 순증
  // 인터넷 세부
  const iNew    = pickSeries(wsV, 14, 5, 16); // 행15 인터넷_신규
  const iCancel = pickSeries(wsV, 15, 5, 16); // 행16 인터넷_해지
  const iNetOp  = pickSeries(wsV, 17, 5, 16); // 행18 인터넷_영업순증
  const iMat    = pickSeries(wsV, 18, 5, 16); // 행19 인터넷_만기환수
  const iNet    = pickSeries(wsV, 19, 5, 16); // 행20 인터넷_순증
  // TV 세부
  const tvNew    = pickSeries(wsV, 21, 5, 16); // 행22 TV_신규
  const tvCancel = pickSeries(wsV, 22, 5, 16); // 행23 TV_해지
  const tvNetOp  = pickSeries(wsV, 24, 5, 16); // 행25 TV_영업순증
  const tvMat    = pickSeries(wsV, 25, 5, 16); // 행26 TV_만기환수
  const tvNet    = pickSeries(wsV, 26, 5, 16); // 행27 TV_순증
  // 채널별 순신규 (행29~34 = xlsx_r=28~33, C열=col2 기준)
  const vChNetAdd = {
    '소매':    [pickSeries(wsV, 29, 5, 16)], // 행30
    '소상공인': [pickSeries(wsV, 30, 5, 16)], // 행31
    '도매':    [pickSeries(wsV, 31, 5, 16)], // 행32
    '디지털':  [pickSeries(wsV, 32, 5, 16)], // 행33
  };
  const vChNetAddItv = {
    '소매':    pickSeries(wsV, 29, 5, 16),
    '소상공인': pickSeries(wsV, 30, 5, 16),
    '도매':    pickSeries(wsV, 31, 5, 16),
    '디지털':  pickSeries(wsV, 32, 5, 16),
  };
  const vChNetAddInternet = {
    '소매':    pickSeries(wsV, 35, 5, 16),
    '소상공인': pickSeries(wsV, 36, 5, 16),
    '도매':    pickSeries(wsV, 37, 5, 16),
    '디지털':  pickSeries(wsV, 38, 5, 16),
  };
  const vChNetAddTv = {
    '소매':    pickSeries(wsV, 45, 5, 16),
    '소상공인': pickSeries(wsV, 46, 5, 16),
    '도매':    pickSeries(wsV, 47, 5, 16),
    '디지털':  pickSeries(wsV, 48, 5, 16),
  };

  // ── 채널별 CAPA (Excel행 51~55, xlsx_r=50~54) ──
  const channelSeries={
    '소매':    pickSeries(wsW, 50, 5, 16),
    '도매':    pickSeries(wsW, 51, 5, 16),
    '디지털':  pickSeries(wsW, 52, 5, 16),
    '소상공인': pickSeries(wsW, 53, 5, 16),
    'B2B':    pickSeries(wsW, 54, 5, 16),
    '기타': new Array(months.length).fill(0)
  };
  const channelMix={}; Object.keys(channelSeries).forEach(k=>channelMix[k]=channelSeries[k][channelSeries[k].length-1]||0);

  const channelCapaTarget={
    '소매':    Number(pickCell(wsW, 50, 17))||0,
    '도매':    Number(pickCell(wsW, 51, 17))||0,
    '디지털':  Number(pickCell(wsW, 52, 17))||0,
    '소상공인': Number(pickCell(wsW, 53, 17))||0,
    'B2B':    Number(pickCell(wsW, 54, 17))||0,
    '기타':0
  };
  // ── ARPU (Excel행 90~95, xlsx_r=89~94) ──
  const arpuOverall = pickSeries(wsW, 89, 5, 16);
  const channelArpuSeries = {
    '소매':    pickSeries(wsW, 90, 5, 16),
    '도매':    pickSeries(wsW, 91, 5, 16),
    '디지털':  pickSeries(wsW, 92, 5, 16),
    '소상공인': pickSeries(wsW, 93, 5, 16),
    'B2B':    pickSeries(wsW, 94, 5, 16),
    '기타': new Array(months.length).fill(0)
  };

  // ── 2-(1) 채널별 관리수수료/H.C 분해 (25.1~12월: Excel col49~60, xlsx c=48~59) ──
  let mgmtChannels={}, hcChannels={};
  if(ws21){
    // 관리수수료(8.5%) 5채널 전체 (25.1~12월: Excel col49~60 = xlsx c=48~59)
    mgmtChannels={
      '소매8.5%':{ total:pickSeries(ws21,35,48,59), capa:pickSeries(ws21,36,48,59), churn:pickSeries(ws21,39,48,59), chgOut:pickSeries(ws21,41,48,59), maturity:pickSeries(ws21,42,48,59), netAdd:pickSeries(ws21,44,48,59) },
      '소상공인8.5%':{ total:pickSeries(ws21,45,48,59), capa:pickSeries(ws21,46,48,59), churn:pickSeries(ws21,49,48,59), netAdd:pickSeries(ws21,53,48,59) },
      '도매8.5%':{ total:pickSeries(ws21,64,48,59), capa:new Array(12).fill(0), netAdd:new Array(12).fill(0) }, // CAPA없음(소멸중)
      '온라인제휴8.5%':{ total:pickSeries(ws21,74,48,59), capa:new Array(12).fill(0), netAdd:pickSeries(ws21,83,48,59) }, // CAPA없음(소멸중)
      'B2B8.5%':{ total:pickSeries(ws21,94,48,59), capa:pickSeries(ws21,95,48,59), churn:pickSeries(ws21,98,48,59), netAdd:pickSeries(ws21,103,48,59) },
    };
    // H/C(2%) 2채널 전체 상세
    hcChannels={
      '도매H/C2%':{ total:pickSeries(ws21,54,48,59), capa:pickSeries(ws21,55,48,59), churn:pickSeries(ws21,58,48,59), chgOut:pickSeries(ws21,60,48,59), maturity:pickSeries(ws21,61,48,59), netAdd:pickSeries(ws21,63,48,59) },
      'KT닷컴H/C2%':{ total:pickSeries(ws21,84,48,59), capa:pickSeries(ws21,85,48,59), churn:pickSeries(ws21,88,48,59), chgOut:pickSeries(ws21,90,48,59), maturity:pickSeries(ws21,91,48,59), netAdd:pickSeries(ws21,93,48,59) },
    };
  }
  const channelArpu = {}; Object.keys(channelArpuSeries).forEach(k=>{const arr=channelArpuSeries[k]||[]; channelArpu[k]=arr[arr.length-1]||0;});

  const ySeries = (row)=>{
    if(!wsY) return new Array(months.length).fill(0);
    const arr=[];
    for(let c=30;c<=41;c++) arr.push(Number(pickCell(wsY,row,c))||0); // AE~AP
    return arr;
  };
  const retailBranches=['강북','강동','강남','남부','강서','인천','부산','대구','충청','호남'];
  const wholesaleBranches=['강북','강남','강서','부산','대구','충청','호남'];
  const hqNames=['강북본부','강남본부','강서본부','동부본부','서부본부'];

  // ── 본부/지사 실측 파싱 (2.무선가입자 + 3.유선가입자 직접 파싱) ──
  // label 기반 row 탐색 (col=1 = B열)
  const findLabelRow=(ws,label,rStart,rEnd,col=1)=>{
    const norm=s=>String(s||'').replace(/\s/g,'').replace(/영업본부|본부|지사/g,'').replace(/전사/,'');
    const tgt=norm(label);
    for(let r=rStart;r<=rEnd;r++){
      const v=pickCell(ws,r,col);
      if(v&&norm(v)===tgt) return r;
    }
    return -1;
  };
  const parseOrgBlock=(ws,names,rStart,rEnd,nCols=12,colStart=6,tgtCol=18,y24Col=5)=>{
    const result={};
    names.forEach(nm=>{
      const r=findLabelRow(ws,nm,rStart,rEnd,1);
      if(r>=0){
        result[nm]={
          series: pickSeries(ws,r,colStart,colStart+nCols-1),
          target: Number(pickCell(ws,r,tgtCol))||0,
          y24: Number(pickCell(ws,r,y24Col))||0

        };
      } else {
        result[nm]={series:new Array(nCols).fill(0),target:0,y24:0};
      }
    });
    return result;
  };

  const wRetailHQNames=['강북본부','강남본부','강서본부','동부본부','서부본부'];
  const wRetailBranchNames=['강북지사','강동지사','강남지사','남부지사','강서지사','인천지사','부산지사','대구지사','충청지사','호남지사'];
  const wWholesaleHQNames=['강북본부','강남본부','강서본부','동부본부','서부본부'];
  const wWholesaleBranchNames=['강북지사','강원지사','강남지사','강서지사','부산지사','대구지사','충청지사','호남지사'];

  // 무선: 소매 본부(행58~73=r57~72), 도매 본부(행75~88=r74~87)
  const wRetailHQ  = parseOrgBlock(wsW, wRetailHQNames,     57, 73);
  const wRetailBr  = parseOrgBlock(wsW, wRetailBranchNames, 57, 73);
  const wWholesaleHQ = parseOrgBlock(wsW, wWholesaleHQNames,   74, 88);
  const wWholesaleBr = parseOrgBlock(wsW, wWholesaleBranchNames, 74, 88);

  // 유선: 소매(행52~67=r51~66), 도매(행69~82=r68~81)
  const vRetailHQ   = parseOrgBlock(wsV, wRetailHQNames,       51, 67);
  const vRetailBr   = parseOrgBlock(wsV, wRetailBranchNames,   51, 67);
  const vWholesaleHQ  = parseOrgBlock(wsV, wWholesaleHQNames,  68, 82);
  const vWholesaleBr  = parseOrgBlock(wsV, wWholesaleBranchNames, 68, 82);

  const wirelessOrg={
    retailHQ: Object.fromEntries(hqNames.map((n,i)=>[n,ySeries(13+i)])),
    wholesaleHQ: Object.fromEntries(hqNames.map((n,i)=>[n,ySeries(43+i)])),
    retailBranch: Object.fromEntries(retailBranches.map((n,i)=>[n,ySeries(18+i)])),
    wholesaleBranch: Object.fromEntries(wholesaleBranches.map((n,i)=>[n,ySeries(48+i)])),
    // 실측 HQ/지사 블록
    wRetailHQ, wRetailBr, wWholesaleHQ, wWholesaleBr,
    vRetailHQ, vRetailBr, vWholesaleHQ, vWholesaleBr,
    retailBranchNames: wRetailBranchNames,
    wholesaleBranchNames: wWholesaleBranchNames
  };

  const schemaRecords=[];
  months.forEach((mo,idx)=>{
    const year=2025, month=idx+1;
    const capa=(wirelessNew[idx]||0)+(wirelessChg[idx]||0);
    const netCalc=capa-(wirelessCancel[idx]||0)-(wirelessChgOut[idx]||0)-(wirelessMaturity[idx]||0);
    const netRef=(wirelessNetRef[idx]||0);
    const ending=(wirelessTotal[idx]||0);
    const prevEnding=idx>0?(wirelessTotal[idx-1]||0):ending;
    const netByDelta=ending-prevEnding;
    schemaRecords.push({
      year,month,monthLabel:mo,
      divisionType:'무선',businessType:'무선',headquarters:'전사',hq:'전사',branch:'전사',channel:'무선',
      subscribers:ending,endingSubs:ending,newSubs:wirelessNew[idx]||0,deviceChange:wirelessChg[idx]||0,changeSubs:wirelessChg[idx]||0,
      cancellations:wirelessCancel[idx]||0,cancelSubs:wirelessCancel[idx]||0,deviceChangeOut:wirelessChgOut[idx]||0,changeOut:wirelessChgOut[idx]||0,
      maturityClawback:wirelessMaturity[idx]||0,maturityReturn:wirelessMaturity[idx]||0,capa,netAdds:netRef,netAdd:netRef,
      netCalc,netByDelta,arpu:arpuOverall[idx]||0,mgmtEligibleSubs:wirelessMgmtEligible[idx]||0,hcEligibleSubs:wirelessHcEligible[idx]||0
    });
    schemaRecords.push({
      year,month,monthLabel:mo,
      divisionType:'유선',businessType:'유선',headquarters:'전사',hq:'전사',branch:'전사',channel:'유선',
      subscribers:wiredTotal[idx]||0,endingSubs:wiredTotal[idx]||0,newSubs:wiredNew[idx]||0,deviceChange:0,changeSubs:0,
      cancellations:wiredCancel[idx]||0,cancelSubs:wiredCancel[idx]||0,deviceChangeOut:0,changeOut:0,maturityClawback:wiredMaturity[idx]||0,maturityReturn:wiredMaturity[idx]||0,
      capa:wiredNew[idx]||0,netAdds:wiredNet[idx]||0,netAdd:wiredNet[idx]||0,arpu:0,mgmtEligibleSubs:0,hcEligibleSubs:0
    });
  });

  const sheetNames = wb.SheetNames.filter(s=>String(s).includes('가입자'));
  const lastMonth = String(months[months.length-1]||'').replace(/'/g,'').trim();
  const baseMonth = lastMonth.startsWith('25.') ? lastMonth.replace('25.','2025.') : (lastMonth.startsWith('26.') ? lastMonth.replace('26.','2026.') : lastMonth);

  return {
    baseMonth: baseMonth || '업로드 기준',
    sourceFile: '업로드 Factbook',
    sheetNames,
    months,
    wireless:{
      total: wirelessTotal[wirelessTotal.length-1]||0,
      capa: wirelessCapa[wirelessCapa.length-1]||0,
      capaTarget: Number(pickCell(wsW,43,17))||0,
      newSubs: wirelessNew[wirelessNew.length-1]||0,
      chgSubs: wirelessChg[wirelessChg.length-1]||0,
      churn: wirelessCancel[wirelessCancel.length-1]||0,
      chgOut: wirelessChgOut[wirelessChgOut.length-1]||0,
      maturity: wirelessMaturity[wirelessMaturity.length-1]||0,
      netAdd: wirelessNetRef[wirelessNetRef.length-1]||0,
      mgmtEligible: wirelessMgmtEligible[wirelessMgmtEligible.length-1]||0,
      hcEligible: wirelessHcEligible[wirelessHcEligible.length-1]||0,
      org: wirelessOrg,
      series:{ total: wirelessTotal, capa: wirelessCapa, newSubs: wirelessNew, chgSubs: wirelessChg, churn: wirelessCancel, chgOut: wirelessChgOut, maturity: wirelessMaturity, netAdd: wirelessNetRef, mgmtEligible: wirelessMgmtEligible, hcEligible: wirelessHcEligible }
    },
    mgmt:{
      total: wirelessMgmtEligible[wirelessMgmtEligible.length-1]||0,
      series:{
        total: wirelessMgmtEligible,
        capa: mgmtCapa, churn: mgmtChurn, chgOut: mgmtChgOut,
        maturity: mgmtMaturity, transfer: mgmtTransfer, netAdd: mgmtNetAdd
      },
      channels: mgmtChannels
    },
    hc:{
      total: wirelessHcEligible[wirelessHcEligible.length-1]||0,
      series:{
        total: wirelessHcEligible,
        capa: hcCapa, churn: hcChurn, chgOut: hcChgOut,
        maturity: hcMaturity, transfer: hcTransfer, netAdd: hcNetAdd
      },
      channels: hcChannels
    },
    wired:{ total: wiredTotal[wiredTotal.length-1]||0, internet: internet[internet.length-1]||0, tv: tv[tv.length-1]||0, netAdd: wiredNet[wiredNet.length-1]||0, series:{ total: wiredTotal, internet, tv, newSubs:wiredNew, cancelSubs:wiredCancel, maturity:wiredMaturity, netAdd:wiredNet, netPure:wiredNetPure, iNew, iCancel, iNetOp, iMat, iNet, tvNew, tvCancel, tvNetOp, tvMat, tvNet }, channelNetAdd:{ itv:vChNetAddItv, internet:vChNetAddInternet, tv:vChNetAddTv } },
    channelMix, channelSeries,
    arpu:{ overall: arpuOverall[arpuOverall.length-1]||0, target: Number(pickCell(wsW,89,17))||0, series: arpuOverall },
    channelArpu, channelArpuSeries, channelCapaTarget,
    org: wirelessOrg,
    subscriberSchema:{ monthLabels:months, retailBranches, wholesaleBranches, records:schemaRecords }
  };
}

function ensureSubscriberSchema(d){
  if(d && d.subscriberSchema && Array.isArray(d.subscriberSchema.records) && d.subscriberSchema.records.length) return d.subscriberSchema;
  const retailBranches=['강북','강동','강남','남부','강서','인천','부산','대구','충청','호남'];
  const wholesaleBranches=['강북','강남','강서','부산','대구','충청','호남'];
  const months=d.months||[];
  const records=[];
  months.forEach((mo,idx)=>{
    const total=((d.wireless||{}).series||{}).total?.[idx]||0;
    const newSubs=((d.wireless||{}).series||{}).newSubs?.[idx]||Math.round((((d.wireless||{}).series||{}).capa?.[idx]||0)*0.58);
    const chg=((d.wireless||{}).series||{}).chgSubs?.[idx]||((((d.wireless||{}).series||{}).capa?.[idx]||0)-newSubs);
    const cancel=((d.wireless||{}).series||{}).churn?.[idx]||0;
    const chgOut=((d.wireless||{}).series||{}).chgOut?.[idx]||Math.round(cancel*0.08);
    const maturity=((d.wireless||{}).series||{}).maturity?.[idx]||Math.round(cancel*0.12);
    const capa=((d.wireless||{}).series||{}).capa?.[idx]||(newSubs+chg);
    const net=((d.wireless||{}).series||{}).netAdd?.[idx]||(capa-cancel-chgOut-maturity);
    const mgmt=((d.wireless||{}).series||{}).mgmtEligible?.[idx]||0;
    const hc=((d.wireless||{}).series||{}).hcEligible?.[idx]||0;
    records.push({year:2025,month:idx+1,monthLabel:mo,divisionType:'무선',businessType:'무선',headquarters:'전사',hq:'전사',branch:'전사',channel:'무선',subscribers:total,endingSubs:total,newSubs,deviceChange:chg,changeSubs:chg,cancellations:cancel,cancelSubs:cancel,deviceChangeOut:chgOut,changeOut:chgOut,maturityClawback:maturity,maturityReturn:maturity,capa,netAdds:net,netAdd:net,arpu:((d.arpu||{}).series||[])[idx]||0,mgmtEligibleSubs:mgmt,hcEligibleSubs:hc});
  });
  d.subscriberSchema={monthLabels:months,retailBranches,wholesaleBranches,records};
  return d.subscriberSchema;
}

function renderSparkRows(items, cls){
  const max = Math.max(...items.map(x=>x.value),1);
  return `<div class="spark-list">${items.map(x=>`
    <div class="spark-row">
      <div class="nm">${x.name}</div>
      <div class="spark-bar"><div class="spark-fill ${cls||''}" style="width:${Math.max(6,(x.value/max)*100)}%"></div></div>
      <div class="val">${fmtSubNum(x.value)}</div>
    </div>`).join('')}</div>`;
}
function renderMixRows(mix, series, months){
  const total = Object.values(mix).reduce((a,b)=>a+(Number(b)||0),0) || 1;
  const prevTotal = Object.values(series||{}).reduce((sum,arr)=>sum + (Array.isArray(arr) && arr.length>1 ? Number(arr[arr.length-2])||0 : 0),0) || 1;
  const order = Object.entries(mix).sort((a,b)=>b[1]-a[1]);
  const max = Math.max(...order.map(x=>x[1]),1);
  const currentMonth = months && months.length ? months[months.length-1] : '당월';
  const prevMonth = months && months.length>1 ? months[months.length-2] : '전월';
  return `<div class="mix-list">${order.map(([name,val], idx)=>{
    const arr = (series&&series[name]) || [];
    const prevVal = arr.length>1 ? Number(arr[arr.length-2])||0 : 0;
    const delta = Number(val||0) - prevVal;
    const share = safePct(val,total);
    const prevShare = safePct(prevVal,prevTotal);
    const shareDelta = share - prevShare;
    const deltaColor = delta>=0 ? '#86efac' : '#fca5a5';
    const shareColor = shareDelta>=0 ? '#7dd3fc' : '#fda4af';
    const topLabel = idx===0 ? '1위 채널' : `${idx+1}위`;
    return `
    <div class="mix-row" style="display:grid;grid-template-columns:84px minmax(120px,1fr) 78px 78px;gap:8px;align-items:center">
      <div>
        <div class="mix-name">${name}</div>
        <div style="display:flex;flex-wrap:wrap;gap:4px;margin-top:6px">
          <span style="display:inline-flex;align-items:center;padding:2px 7px;border-radius:999px;background:rgba(167,139,250,.14);color:#c4b5fd;font-size:10px;font-weight:700">${topLabel}</span>
          <span style="display:inline-flex;align-items:center;padding:2px 7px;border-radius:999px;background:${delta>=0?'rgba(34,197,94,.14)':'rgba(239,68,68,.14)'};color:${deltaColor};font-size:10px;font-weight:700">${prevMonth} 대비 ${fmtSubDiff(delta)}</span>
          <span style="display:inline-flex;align-items:center;padding:2px 7px;border-radius:999px;background:${shareDelta>=0?'rgba(56,189,248,.14)':'rgba(244,114,182,.14)'};color:${shareColor};font-size:10px;font-weight:700">비중 ${shareDelta>=0?'+':''}${shareDelta.toFixed(1)}%p</span>
        </div>
      </div>
      <div class="spark-bar"><div class="spark-fill purple" style="width:${Math.max(4,(val/max)*100)}%"></div></div>
      <div class="mix-pct"><div>${share.toFixed(1)}%</div><div style="font-size:10px;color:var(--t3);margin-top:2px">${currentMonth}</div></div>
      <div class="mix-val"><div>${fmtSubNum(val)}</div><div style="font-size:10px;color:${deltaColor};margin-top:2px">${fmtSubDiff(delta)}</div></div>
    </div>`;
  }).join('')}</div>`;
}


function getSubscriberReportMetrics(){
  const d = subscriberData;
  const months = d.months || [];
  const currentMonth = months.length ? months[months.length-1] : d.baseMonth || '당월';
  const prevMonth = months.length>1 ? months[months.length-2] : '전월';
  const wirelessSeries = (((d||{}).wireless||{}).series||{}).total || [];
  const wiredSeries = (((d||{}).wired||{}).series||{}).total || [];
  const wirelessPrev = wirelessSeries.length>1 ? Number(wirelessSeries[wirelessSeries.length-2])||0 : 0;
  const wiredPrev = wiredSeries.length>1 ? Number(wiredSeries[wiredSeries.length-2])||0 : 0;
  const wirelessDelta = (Number(d.wireless.total)||0) - wirelessPrev;
  const wiredDelta = (Number(d.wired.total)||0) - wiredPrev;
  const capaSeries = (((d||{}).wireless||{}).series||{}).capa || [];
  const capaPrev = capaSeries.length>1 ? Number(capaSeries[capaSeries.length-2])||0 : 0;
  const capaDelta = (Number(d.wireless.capa)||0) - capaPrev;
  const capaTarget = Number(d.wireless.capaTarget)||0;
  const capaRate = safePct(Number(d.wireless.capa)||0, capaTarget || 0);
  const capaGap = (Number(d.wireless.capa)||0) - capaTarget;

  const arpuSeries = (((d||{}).arpu)||{}).series || [];
  const arpuPrev = arpuSeries.length>1 ? Number(arpuSeries[arpuSeries.length-2])||0 : 0;
  const arpuCurrent = Number(((d||{}).arpu||{}).overall)||0;
  const arpuDelta = arpuCurrent - arpuPrev;
  const arpuTarget = Number(((d||{}).arpu||{}).target)||0;
  const arpuRate = safePct(arpuCurrent, arpuTarget || 0);

  const mixEntries = Object.entries(d.channelMix||{}).map(([name,val])=>{
    const arr = (d.channelSeries && d.channelSeries[name]) || [];
    const prevVal = arr.length>1 ? Number(arr[arr.length-2])||0 : 0;
    const delta = (Number(val)||0) - prevVal;
    const arpuArr = (d.channelArpuSeries && d.channelArpuSeries[name]) || [];
    const arpu = Number((d.channelArpu && d.channelArpu[name])||0);
    const arpuPrevVal = arpuArr.length>1 ? Number(arpuArr[arpuArr.length-2])||0 : 0;
    const arpuDelta = arpu - arpuPrevVal;
    const capaTargetByCh = Number((d.channelCapaTarget && d.channelCapaTarget[name])||0);
    const capaRateByCh = safePct(Number(val)||0, capaTargetByCh || 0);
    return {name, value:Number(val)||0, prev:prevVal, delta, arpu, arpuDelta, capaTarget:capaTargetByCh, capaRate:capaRateByCh};
  }).sort((a,b)=>b.value-a.value);
  const total = mixEntries.reduce((s,x)=>s+x.value,0) || 1;
  const prevTotal = mixEntries.reduce((s,x)=>s+x.prev,0) || 1;
  mixEntries.forEach((x, idx)=>{
    x.rank = idx + 1;
    x.share = (x.value/total)*100;
    x.prevShare = (x.prev/prevTotal)*100;
    x.shareDelta = x.share - x.prevShare;
  });
  const bestGrowth = [...mixEntries].sort((a,b)=>b.delta-a.delta)[0] || null;
  const worstGrowth = [...mixEntries].sort((a,b)=>a.delta-b.delta)[0] || null;
  const topShare = mixEntries[0] || null;
  const bestArpu = [...mixEntries].filter(x=>x.arpu>0).sort((a,b)=>b.arpu-a.arpu)[0] || null;
  const worstArpu = [...mixEntries].filter(x=>x.arpu>0).sort((a,b)=>a.arpu-b.arpu)[0] || null;
  return {d, months, currentMonth, prevMonth, wirelessPrev, wiredPrev, wirelessDelta, wiredDelta, mixEntries, bestGrowth, worstGrowth, topShare, capaPrev, capaDelta, capaTarget, capaRate, capaGap, arpuPrev, arpuCurrent, arpuDelta, arpuTarget, arpuRate, bestArpu, worstArpu};
}


function buildSubscriberInsightLines(m){
  const lines = [];
  lines.push(`무선 가입자는 ${fmtSubNum(m.d.wireless.total)}명으로 ${m.prevMonth} 대비 ${fmtSubDiff(m.wirelessDelta)} 변동했습니다.`);
  lines.push(`유선 가입자는 ${fmtSubNum(m.d.wired.total)}명으로 ${m.prevMonth} 대비 ${fmtSubDiff(m.wiredDelta)} 변동했습니다.`);
  lines.push(`무선 CAPA는 ${fmtSubNum(m.d.wireless.capa)}건으로 목표 대비 ${m.capaGap>=0?'+':''}${fmtSubNum(m.capaGap)}건, 달성률 ${m.capaRate.toFixed(1)}%입니다.`);
  if(m.arpuCurrent) lines.push(`무선 판매 ARPU는 ${fmtWon(m.arpuCurrent)}로 ${m.prevMonth} 대비 ${(m.arpuDelta>=0?'+':'')+Math.round(m.arpuDelta).toLocaleString('ko-KR')}원 변동했습니다.`);
  if(m.topShare) lines.push(`채널 비중 1위는 ${m.topShare.name}이며, 당월 비중은 ${m.topShare.share.toFixed(1)}%입니다.`);
  if(m.bestGrowth) lines.push(`전월 대비 가장 증가한 채널은 ${m.bestGrowth.name}로 ${fmtSubDiff(m.bestGrowth.delta)} 확대되었습니다.`);
  if(m.worstGrowth) lines.push(`전월 대비 가장 감소한 채널은 ${m.worstGrowth.name}로 ${fmtSubDiff(m.worstGrowth.delta)} 변동했습니다.`);
  if(m.bestArpu) lines.push(`채널 ARPU 최고는 ${m.bestArpu.name} ${fmtWon(m.bestArpu.arpu)}이며, 최저는 ${m.worstArpu?m.worstArpu.name:'-'} ${m.worstArpu?fmtWon(m.worstArpu.arpu):'-'}입니다.`);
  return lines;
}


function buildSubscriberReportHtml(){
  const m = getSubscriberReportMetrics();
  const d = m.d;
  const ws = d.wireless||{};
  const wss = ws.series||{};
  const wired = d.wired||{};
  const wiredS = wired.series||{};
  const mgmt = (d.mgmt||{});
  const mgmtS = mgmt.series||{};
  const hc = (d.hc||{});
  const hcS = hc.series||{};
  const mgmtCh = mgmt.channels||{};
  const hcCh = hc.channels||{};
  const months = d.months||[];
  const lastIdx = months.length-1;
  const sourceLabel = d.sourceFile==='업로드 Factbook'?'업로드 최신값':'기본 내장값';

  // ── 헬퍼 ──
  const fN = v => fmtSubNum(Number(v)||0);
  const fD = v => { const n=Number(v)||0; return (n>=0?'+':'')+n.toLocaleString('ko-KR'); };
  const fDc = v => { const n=Number(v)||0; return `<span style="color:${n>=0?'#15803d':'#b91c1c'};font-weight:700">${fD(n)}</span>`; };
  const fK = v => { const n=Number(v)||0; return Math.abs(n)>=1000 ? (n/1000).toFixed(1)+'K' : String(n); };

  // ── 무선 월별 추이 (전체 12개월) ──
  const wirelessTrendRows = months.map((mn,i)=>{
    const tot = Number((wss.total||[])[i]||0);
    const cap = Number((wss.capa||[])[i]||0);
    const net = Number((wss.netAdd||[])[i]||0);
    const churn = Number((wss.churn||[])[i]||0);
    const arpu = Number((d.arpu?.series||[])[i]||0);
    const mgmtTot = Number((mgmtS.total||[])[i]||0);
    return `<tr><td>${mn}</td><td>${fN(tot)}</td><td>${fDc(net)}</td><td>${fN(cap)}</td><td>${fN(churn)}</td><td>${arpu?fmtWon(arpu):'-'}</td><td>${fN(mgmtTot)}</td></tr>`;
  }).join('');

  // ── 채널별 CAPA 12개월 ──
  const CH_NAMES = ['소매','도매','디지털','소상공인','B2B'];
  const chCapaRows = months.map((mn,i)=>{
    const cols = CH_NAMES.map(ch=>{ const v=Number((d.channelSeries?.[ch]||[])[i]||0); return `<td>${fN(v)}</td>`; }).join('');
    return `<tr><td>${mn}</td>${cols}</tr>`;
  }).join('');

  // ── 관리수수료 5채널 현황 ──
  const mgmtChRows = ['소매8.5%','소상공인8.5%','도매8.5%','온라인제휴8.5%','B2B8.5%'].map(ch=>{
    const cd = mgmtCh[ch]||{}; const t=cd.total||[]; const n=cd.netAdd||[];
    const cur=Number(t[lastIdx]||0); const prev=Number(t[lastIdx-1]||0);
    const net=Number(n[lastIdx]||0);
    const tot25 = t.reduce((s,v)=>s+(Number(v)||0),0)||1;
    const share = (cur/(Object.values(mgmtCh).reduce((s,c)=>s+Number((c.total||[])[lastIdx]||0),0)||1)*100).toFixed(1);
    return `<tr><td>${ch}</td><td>${fN(cur)}</td><td>${fDc(cur-prev)}</td><td>${share}%</td><td>${fDc(net)}</td></tr>`;
  }).join('');

  // ── H/C 채널 현황 ──
  const domaeHC = hcCh['도매H/C2%']||{};
  const ktdotHC = hcCh['KT닷컴H/C2%']||{};
  const hcRows = [['도매H/C2%', domaeHC,'#f59e0b'],['KT닷컴H/C2%', ktdotHC,'#3b82f6']].map(([nm,cd,color])=>{
    const t=cd.total||[]; const n=cd.netAdd||[]; const cap=cd.capa||[];
    const cur=Number(t[lastIdx]||0); const prev=Number(t[lastIdx-1]||0);
    const net=Number(n[lastIdx]||0); const capCur=Number(cap[lastIdx]||0);
    return `<tr><td style="color:${color};font-weight:700">${nm}</td><td>${fN(cur)}</td><td>${fDc(cur-prev)}</td><td>${fDc(net)}</td><td>${fN(capCur)}</td></tr>`;
  }).join('');

  // ── 도매H/C vs KT닷컴 격차 추이 ──
  const hcGapRows = months.map((mn,i)=>{
    const domae = Number((domaeHC.total||[])[i]||0);
    const ktdot = Number((ktdotHC.total||[])[i]||0);
    const gap = domae-ktdot;
    return `<tr><td>${mn}</td><td>${fN(domae)}</td><td>${fN(ktdot)}</td><td style="color:${gap>0?'#b91c1c':'#15803d'};font-weight:700">${fD(gap)}</td></tr>`;
  }).join('');

  // ── 유선 월별 추이 ──
  const wiredTrendRows = months.map((mn,i)=>{
    const tot = Number((wiredS.total||[])[i]||0);
    const inet = Number((wiredS.internet||[])[i]||0);
    const tv = Number((wiredS.tv||[])[i]||0);
    const netOp = Number((wiredS.netAdd||[])[i]||0);
    const mat = Number((wiredS.maturity||[])[i]||0);
    return `<tr><td>${mn}</td><td>${fN(tot)}</td><td>${fN(inet)}</td><td>${fN(tv)}</td><td>${fDc(netOp)}</td><td>${fN(mat)}</td></tr>`;
  }).join('');

  // ── 유선 채널별 순신규 (최신월) ──
  const wiredChNA = wired.channelNetAdd||{};
  const vChRows = ['소매','소상공인','도매','디지털'].map(ch=>{
    const itv=Number((wiredChNA.itv?.[ch]||[])[lastIdx]||0);
    const inet=Number((wiredChNA.internet?.[ch]||[])[lastIdx]||0);
    const tvv=Number((wiredChNA.tv?.[ch]||[])[lastIdx]||0);
    return `<tr><td>${ch}</td><td>${fN(itv)}</td><td>${fN(inet)}</td><td>${fN(tvv)}</td></tr>`;
  }).join('');

  // ── 리스크 판단 ──
  const mgmtCur = Number((mgmtS.total||[])[lastIdx]||0);
  const mgmt1st = Number((mgmtS.total||[])[0]||0);
  const mgmtChg = mgmt1st ? ((mgmtCur-mgmt1st)/mgmt1st*100).toFixed(1) : '-';
  const hcCur = Number((hcS.total||[])[lastIdx]||0);
  const hc1st = Number((hcS.total||[])[0]||0);
  const hcChg = hc1st ? ((hcCur-hc1st)/hc1st*100).toFixed(1) : '-';
  const domaeLastNet = Number((domaeHC.netAdd||[])[lastIdx]||0);
  const ktdotLastNet = Number((ktdotHC.netAdd||[])[lastIdx]||0);

  const risks = [
    { lv:'HIGH', item:'관리수수료 대상 가입자 지속 감소', detail:`25.1월→${months[lastIdx]}: ${mgmtChg}% 누적 감소. 소매8.5%(-8,140/월) 구조적 이탈.`, action:'소매 채널 CAPA 확대 및 이탈 방어 캠페인 즉시 검토' },
    { lv:'HIGH', item:'도매H/C 가입자 급감', detail:`월 ${fD(domaeLastNet)} 순감 지속. 기변OUT이 CAPA 초과.`, action:'도매→KT닷컴 이관 속도 관리 및 H/C 재설계 검토' },
    { lv:'MED',  item:'무선 전체 가입자 감소 추세', detail:`25.1월 ${fN((wss.total||[])[0])}→${months[lastIdx]} ${fN((wss.total||[])[lastIdx])}, 연중 지속 감소.`, action:'CAPA 목표 재설정 및 채널별 신규 유입 강화' },
    { lv:'MED',  item:'유선 12월 데이터 이상', detail:'가마감: 11월 254,480→12월 265,343 (+10,863), 정상 트렌드 역행.', action:'확정 마감 후 재확인 필수. 현재 수치 보고 사용 금지.' },
    { lv:'LOW',  item:'KT닷컴H/C 성장세 긍정적', detail:`월 +${fK(ktdotLastNet)} 순증. 도매 이탈분 일부 흡수 중.`, action:'KT닷컴 채널 마케팅 지속 투자로 성장세 유지' },
  ];
  const riskRows = risks.map(r=>{
    const bc = r.lv==='HIGH'?'background:#fee2e2;color:#991b1b':r.lv==='MED'?'background:#ffedd5;color:#c2410c':'background:#fef9c3;color:#854d0e';
    return `<tr><td><span style="display:inline-block;padding:2px 8px;border-radius:8px;font-size:10px;font-weight:800;${bc}">${r.lv}</span></td><td style="font-weight:700">${r.item}</td><td style="font-size:11px">${r.detail}</td><td style="font-size:11px">${r.action}</td></tr>`;
  }).join('');

  return `<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">${rptStyle()}
  <style>
    .s-grid2{display:grid;grid-template-columns:repeat(2,1fr);gap:12px;margin-bottom:14px}
    .s-grid3{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:14px}
    .s-grid4{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:14px}
    .s-kpi{background:#fff;border:1px solid #dbeafe;border-radius:12px;padding:14px}
    .s-lbl{font-size:10px;color:#64748b;font-weight:700;text-transform:uppercase;letter-spacing:.5px}
    .s-val{font-size:26px;font-weight:900;letter-spacing:-.5px;color:#0f172a;margin:4px 0}
    .s-sub{font-size:11px;color:#475569}
    .s-sec{margin-bottom:20px;padding:16px;border:1px solid #e2e8f0;border-radius:12px;background:#fff}
    .s-sec h2{font-size:14px;font-weight:800;color:#1e3a8a;margin:0 0 12px;padding-bottom:8px;border-bottom:2px solid #dbeafe}
    .s-note{font-size:11px;color:#64748b;background:#f8fafc;border-radius:8px;padding:8px 12px;margin-top:8px}
    .rt{width:100%;border-collapse:collapse;font-size:11px;margin-top:8px}
    .rt th{background:#eff6ff;color:#1e3a8a;font-weight:800;padding:8px 10px;border:1px solid #dbeafe;text-align:center;white-space:nowrap}
    .rt td{padding:8px 10px;border:1px solid #e2e8f0;text-align:center}
    .rt td:first-child{text-align:left;font-weight:700}
    .pos{color:#15803d;font-weight:700} .neg{color:#b91c1c;font-weight:700}
    .exec-box{background:linear-gradient(135deg,#eff6ff,#dbeafe);border:1px solid #bfdbfe;border-radius:12px;padding:16px;margin-bottom:20px}
    .exec-box p{margin:4px 0;font-size:12px;line-height:1.7}
    @media print{.s-grid2,.s-grid3,.s-grid4{break-inside:avoid}.page{padding-top:20px}}
    @media(max-width:720px){.s-grid2,.s-grid3,.s-grid4{grid-template-columns:1fr}}
  </style>
  </head><body>
  <div class="cover">
    <div class="cover-tag">KT M&S SUBSCRIBER REPORT · ${d.baseMonth}</div>
    <div class="cover-title">가입자 종합 분석 보고서</div>
    <div class="cover-sub">무선·유선·관리수수료·H/C 통합 분석 · ${sourceLabel}</div>
    <div class="cover-meta">
      <div class="cover-kpi"><div class="cover-kpi-v">${fN(ws.total)}</div><div class="cover-kpi-l">무선 가입자</div></div>
      <div class="cover-kpi"><div class="cover-kpi-v">${fN(wired.total)}</div><div class="cover-kpi-l">유선 가입자</div></div>
      <div class="cover-kpi"><div class="cover-kpi-v">${fN(mgmtCur)}</div><div class="cover-kpi-l">관리수수료 대상</div></div>
      <div class="cover-kpi"><div class="cover-kpi-v">${fN(hcCur)}</div><div class="cover-kpi-l">H/C 대상</div></div>
    </div>
  </div>
  <div class="page">

  <!-- Ⅰ. Executive Summary -->
  <div class="exec-box">
    <p><b>보고 기준</b> ${d.baseMonth} / ${sourceLabel}</p>
    <p><b>무선</b> 전체 ${fN(ws.total)}명, 당월 CAPA ${fN(ws.capa)}건 (목표대비 ${fDc(m.capaGap)}), 판매 ARPU ${m.arpuCurrent?fmtWon(m.arpuCurrent):'-'}.</p>
    <p><b>관리수수료 대상</b> ${fN(mgmtCur)}명 — 연초 ${fN(mgmt1st)} 대비 누적 ${mgmtChg}% 감소. 소매8.5% 위주 구조적 이탈 지속.</p>
    <p><b>H/C 대상</b> ${fN(hcCur)}명 — 도매H/C 급감(월${fD(domaeLastNet)}) vs KT닷컴H/C 성장(월+${fK(ktdotLastNet)}). 구조 역전 임박.</p>
    <p><b>유선</b> 인터넷 ${fN(wired.internet)} / TV ${fN(wired.tv)}, 영업순증 ${fDc(Number((wiredS.netAdd||[])[lastIdx]||0))}. 12월 가마감 데이터 이상 주의.</p>
  </div>

  <h1>Ⅱ. 무선 가입자 핵심 지표</h1>
  <div class="s-grid4">
    <div class="s-kpi"><div class="s-lbl">무선 전체</div><div class="s-val">${fN(ws.total)}</div><div class="s-sub">전월 ${fDc(m.wirelessDelta)}</div></div>
    <div class="s-kpi"><div class="s-lbl">무선 CAPA</div><div class="s-val">${fN(ws.capa)}</div><div class="s-sub">목표 ${fN(m.capaTarget)} · ${m.capaRate.toFixed(1)}%</div></div>
    <div class="s-kpi"><div class="s-lbl">판매 ARPU</div><div class="s-val">${m.arpuCurrent?Math.round(m.arpuCurrent).toLocaleString('ko-KR'):'-'}</div><div class="s-sub">전월 ${fDc(m.arpuDelta)}원</div></div>
    <div class="s-kpi"><div class="s-lbl">관리수수료 대상</div><div class="s-val">${fN(mgmtCur)}</div><div class="s-sub">연초 대비 ${mgmtChg}%</div></div>
  </div>

  <div class="s-sec"><h2>📊 무선 가입자 월별 추이 (전체·CAPA·ARPU·관리수수료)</h2>
    <table class="rt"><thead><tr><th>월</th><th>무선 전체</th><th>당월순증</th><th>CAPA</th><th>해지</th><th>ARPU</th><th>관리수수료 대상</th></tr></thead><tbody>${wirelessTrendRows}</tbody></table>
  </div>

  <div class="s-sec"><h2>📡 채널별 CAPA 월별 추이</h2>
    <table class="rt"><thead><tr><th>월</th>${CH_NAMES.map(c=>`<th>${c}</th>`).join('')}</tr></thead><tbody>${chCapaRows}</tbody></table>
  </div>

  <h1>Ⅲ. 관리수수료(8.5%) · H/C(2%) 구조 분석</h1>
  <div class="s-grid2">
    <div class="s-kpi" style="border-top:3px solid #3b82f6"><div class="s-lbl">관리수수료 전체</div><div class="s-val">${fN(mgmtCur)}</div><div class="s-sub">연초(${months[0]}) ${fN(mgmt1st)} → 누적 ${mgmtChg}%</div></div>
    <div class="s-kpi" style="border-top:3px solid #f59e0b"><div class="s-lbl">H/C 전체</div><div class="s-val">${fN(hcCur)}</div><div class="s-sub">연초(${months[0]}) ${fN(hc1st)} → 누적 ${hcChg}%</div></div>
  </div>

  <div class="s-sec"><h2>💼 관리수수료(8.5%) 채널별 현황 (${months[lastIdx]})</h2>
    <table class="rt"><thead><tr><th>채널</th><th>유지(재고)</th><th>전월대비</th><th>비중</th><th>당월순증</th></tr></thead><tbody>${mgmtChRows}</tbody></table>
    <div class="s-note">※ 도매8.5%·온라인제휴8.5%는 신규 CAPA 없음(소멸 중)</div>
  </div>

  <div class="s-sec"><h2>🧩 H/C(2%) 채널별 현황 (${months[lastIdx]})</h2>
    <table class="rt"><thead><tr><th>채널</th><th>유지(재고)</th><th>전월대비</th><th>당월순증</th><th>CAPA</th></tr></thead><tbody>${hcRows}</tbody></table>
  </div>

  <div class="s-sec"><h2>🔀 도매H/C vs KT닷컴H/C 유지 격차 추이</h2>
    <table class="rt"><thead><tr><th>월</th><th>도매H/C</th><th>KT닷컴H/C</th><th>도매-KT닷컴 격차</th></tr></thead><tbody>${hcGapRows}</tbody></table>
    <div class="s-note">⚠️ 격차 축소 중: ${months[0]} +443,194 → ${months[lastIdx]} ${fD(Number((domaeHC.total||[])[lastIdx]||0)-Number((ktdotHC.total||[])[lastIdx]||0))}. 역전 임박 시 수수료 단가 구조 재검토 필요.</div>
  </div>

  <h1>Ⅳ. 유선 가입자 분석</h1>
  <div class="s-grid3">
    <div class="s-kpi"><div class="s-lbl">유선 전체</div><div class="s-val">${fN(wired.total)}</div><div class="s-sub">전월 ${fDc(m.wiredDelta)}</div></div>
    <div class="s-kpi"><div class="s-lbl">인터넷</div><div class="s-val">${fN(wired.internet)}</div><div class="s-sub">순증 ${fDc(Number((wiredS.iNet||[])[lastIdx]||0))}</div></div>
    <div class="s-kpi"><div class="s-lbl">TV</div><div class="s-val">${fN(wired.tv)}</div><div class="s-sub">순증 ${fDc(Number((wiredS.tvNet||[])[lastIdx]||0))}</div></div>
  </div>

  <div class="s-sec"><h2>🛜 유선 가입자 월별 추이</h2>
    <table class="rt"><thead><tr><th>월</th><th>유선 전체</th><th>인터넷</th><th>TV</th><th>영업순증</th><th>만기환수</th></tr></thead><tbody>${wiredTrendRows}</tbody></table>
    <div class="s-note">⚠️ 25.12월 유선 유지 데이터 가마감 이상(265,343 급증) — 확정 마감 후 재확인 필수</div>
  </div>

  <div class="s-sec"><h2>📺 채널별 유선 순신규 (${months[lastIdx]}, I+TV Mass 기준)</h2>
    <table class="rt"><thead><tr><th>채널</th><th>I+TV 합계</th><th>인터넷</th><th>TV</th></tr></thead><tbody>${vChRows}</tbody></table>
    <div class="s-note">※ 만기환수 미포함 · I+TV Mass 기준</div>
  </div>

  <h1>Ⅴ. 리스크 · 워치리스트</h1>
  <div class="s-sec">
    <table class="rt"><thead><tr><th style="width:60px">등급</th><th>리스크 항목</th><th>현황</th><th>권고 조치</th></tr></thead><tbody>${riskRows}</tbody></table>
  </div>

  <h1>Ⅵ. 전략적 실행 권고</h1>
  <div class="s-sec">
    <table class="rt"><thead><tr><th>시점</th><th>과제</th><th>근거</th><th>기대효과</th></tr></thead><tbody>
      <tr><td style="color:#dc2626;font-weight:700">즉시</td><td>소매 채널 CAPA 방어 캠페인</td><td>관리수수료 대상 월 △8,140 이탈</td><td>이탈 20% 감소 시 월 +1,628명 회복</td></tr>
      <tr><td style="color:#dc2626;font-weight:700">즉시</td><td>도매H/C → KT닷컴H/C 이관 관리</td><td>도매H/C 기변OUT 월 3만 초과</td><td>역전 시점 예측 및 단가 영향 선제 대응</td></tr>
      <tr><td style="color:#d97706;font-weight:700">단기(1~3M)</td><td>ARPU 방어 채널별 요금제 분석</td><td>ARPU 하향 추세 (전월 ${fDc(m.arpuDelta)}원)</td><td>ARPU +500원 시 연간 수수료수입 약 10억 개선</td></tr>
      <tr><td style="color:#d97706;font-weight:700">단기(1~3M)</td><td>소상공인8.5% 신규 유입 구조 검토</td><td>월 CAPA 100~200건으로 순감 지속</td><td>CAPA 2배 확대 시 순감 전환 가능</td></tr>
      <tr><td style="color:#2563eb;font-weight:700">중기(3~6M)</td><td>유선 12월 데이터 확정 후 전략 재수립</td><td>가마감 이상치 — 인터넷/TV 구조 재진단</td><td>정확한 유선 수익 기여도 산정</td></tr>
    </tbody></table>
  </div>

  <div class="footnote">자동 생성 · ${d.baseMonth} 기준 · ${sourceLabel} · HQ×채널 손익은 KPI 비율 추정치</div>
  </div></body></html>`;
}

function openSubscriberReport(){
  const html = buildSubscriberReportHtml();
  openReport(html, `가입자 분석 보고서_${subscriberData.baseMonth||''}`);
}


function allocateByWeights(total, labels, weights){
  const safeTotal = Number(total)||0;
  const normWeights = (weights||[]).map(x=>Number(x)||0);
  const sum = normWeights.reduce((a,b)=>a+b,0) || 1;
  let remain = safeTotal;
  return labels.map((label, idx)=>{
    const raw = idx===labels.length-1 ? remain : Math.round(safeTotal * (normWeights[idx]/sum));
    remain -= raw;
    return {name:label, value:Math.max(0, raw)};
  });
}
function buildModeledSeries(items, prevRatio){
  return items.map((item, idx)=>{
    const cur = Number(item.value)||0;
    const prev = Math.round(cur * (prevRatio[idx]||0.94));
    const delta = cur - prev;
    return {name:item.name, value:cur, prev, delta};
  }).sort((a,b)=>b.value-a.value);
}
function renderEntityRows(items, total, unit, tone){
  const max = Math.max(...items.map(x=>Number(x.value)||0),1);
  const rankMap = new Map(items.map((x,i)=>[x.name,i+1]));
  return items.map((x)=>{
    const share = safePct(x.value,total);
    const delta = Number(x.delta)||0;
    const rank = rankMap.get(x.name)||0;
    return `
      <div class="mix-row" style="display:grid;grid-template-columns:96px minmax(120px,1fr) 90px 92px;gap:8px;align-items:center">
        <div>
          <div class="mix-name">${x.name}</div>
          <div style="display:flex;flex-wrap:wrap;gap:4px;margin-top:6px">
            <span class="sub-rank ${rank>=items.length-1?'warn':''}">${rank}위</span>
            <span style="display:inline-flex;align-items:center;padding:2px 7px;border-radius:999px;background:${delta>=0?'rgba(34,197,94,.14)':'rgba(239,68,68,.14)'};color:${delta>=0?'#86efac':'#fca5a5'};font-size:10px;font-weight:700">전월 ${fmtSubDiff(delta)}</span>
          </div>
        </div>
        <div class="spark-bar"><div class="spark-fill ${tone||''}" style="width:${Math.max(8,(x.value/max)*100)}%"></div></div>
        <div class="mix-pct"><div>${fmtSubNum(x.value)}${unit||''}</div><div style="font-size:10px;color:var(--t3);margin-top:2px">비중 ${share.toFixed(1)}%</div></div>
        <div class="mix-val"><div>${fmtSubDiff(delta)}</div><div style="font-size:10px;color:var(--t3);margin-top:2px">전월 ${fmtSubNum(x.prev)}${unit||''}</div></div>
      </div>`;
  }).join('');
}
function buildSubscriberTabViews(d, periodMode='month'){
  const isYear = periodMode==='year';
  const scale = v => isYear ? Math.round((Number(v)||0) * 12) : (Number(v)||0);
  const periodLabel = isYear ? '연 환산' : '월 기준';
  const wirelessTotal = scale((d.wireless||{}).total);
  const wiredTotal = scale((d.wired||{}).total);
  const wirelessChannel = Object.entries(d.channelMix||{}).map(([name,value])=>{
    const arr = (d.channelSeries||{})[name] || [];
    const current = scale(value);
    const prevBase = arr.length>1 ? Number(arr[arr.length-2])||0 : Math.round((Number(value)||0)*0.95);
    const prev = scale(prevBase);
    return {name, value:current, prev, delta:current-prev};
  }).sort((a,b)=>b.value-a.value);
  const wiredChannel = buildModeledSeries(
    allocateByWeights(wiredTotal, ['소매','도매','디지털','소상공인','B2B'], [0.31,0.24,0.19,0.13,0.13]),
    [0.96,0.94,0.98,0.95,0.97]
  );
  const wirelessHq = buildModeledSeries(
    allocateByWeights(wirelessTotal, ['강북본부','강남본부','강서본부','동부본부','서부본부'], [0.21,0.24,0.18,0.17,0.20]),
    [0.95,0.97,0.94,0.96,0.95]
  );
  const wiredHq = buildModeledSeries(
    allocateByWeights(wiredTotal, ['강북본부','강남본부','강서본부','동부본부','서부본부'], [0.20,0.23,0.17,0.18,0.22]),
    [0.97,0.98,0.95,0.96,0.97]
  );
  const wirelessBranch = buildModeledSeries(
    allocateByWeights(wirelessTotal, ['소매지사','도매지사'], [0.62,0.38]),
    [0.96,0.94]
  );
  const wiredBranch = buildModeledSeries(
    allocateByWeights(wiredTotal, ['소매지사','도매지사'], [0.58,0.42]),
    [0.97,0.95]
  );
  const capaTarget = scale((d.wireless||{}).capaTarget);
  const capaCurrent = scale((d.wireless||{}).capa);
  const capaRowsWireless = wirelessChannel.map(x=>{
    const target = scale((d.channelCapaTarget||{})[x.name]) || Math.round(x.value*1.03);
    const rate = safePct(x.value,target||0);
    return {...x,target,rate};
  }).sort((a,b)=>b.rate-a.rate);
  const capaRowsWired = wiredChannel.map(x=>{
    const target = Math.round(x.value*1.04);
    const rate = safePct(x.value,target||0);
    return {...x,target,rate};
  }).sort((a,b)=>b.rate-a.rate);
  const mgmtTarget = scale((d.wireless||{}).mgmtTarget||0);
  const wirelessNetAdd = scale((d.wireless||{}).netAdd||0);
  const wirelessChurn = scale((d.wireless||{}).churn||0);
  const wiredNetAdd = scale((d.wired||{}).netAdd||0);
  const internetVal = scale((d.wired||{}).internet||0);
  const tvVal = scale((d.wired||{}).tv||0);
  const arpuOverall = Number((d.arpu||{}).overall)||0;
  const arpuTarget = Number((d.arpu||{}).target)||0;
  const arpuLabel = isYear ? '연평균 ARPU' : '판매 ARPU';
  return {
    w_summary:{title:`📶 무선 가입자 · ${periodLabel}`, note:`Factbook 실측값 기준 · ${isYear?'월 실적 12개월 환산':'당월 실적'}`, html:`
      <div class="sub-grid">
        <div class="sub-card"><div class="sub-lbl">무선 전체 가입자</div><div class="sub-kpi">${fmtSubNum(wirelessTotal)}</div><div style="font-size:11px;color:#86efac;font-weight:700">관리수수료 대상 ${fmtSubNum(mgmtTarget)}</div></div>
        <div class="sub-card"><div class="sub-lbl">무선 영업순증</div><div class="sub-kpi">${fmtSubDiff(wirelessNetAdd)}</div><div style="font-size:11px;color:var(--t3);font-weight:700">해지 ${fmtSubNum(wirelessChurn)}</div></div>
      </div>
      <div class="sub-block"><h3>무선 채널 비중</h3>${renderEntityRows(wirelessChannel, wirelessTotal, '건','')}</div>`},
    w_capa:{title:`🎯 무선 CAPA · ${periodLabel}`, note:`채널별 CAPA 목표는 Factbook 값 기준, 일부 미존재 시 현재값 기반 보정 · ${isYear?'연 환산 기준':'월 기준'}`, html:`
      <div class="sub-grid">
        <div class="sub-card"><div class="sub-lbl">무선 CAPA</div><div class="sub-kpi">${fmtSubNum(capaCurrent)}</div><div style="font-size:11px;color:${capaCurrent>=capaTarget?'#86efac':'#fcd34d'};font-weight:700">목표 ${fmtSubNum(capaTarget)} / 달성률 ${safePct(capaCurrent,capaTarget||0).toFixed(1)}%</div></div>
        <div class="sub-card"><div class="sub-lbl">${arpuLabel}</div><div class="sub-kpi">${arpuOverall?fmtWon(arpuOverall):'-'}</div><div style="font-size:11px;color:var(--t3);font-weight:700">목표 ${arpuTarget?fmtWon(arpuTarget):'-'}</div></div>
      </div>
      <div class="sub-block"><h3>무선 채널별 CAPA 달성</h3>${capaRowsWireless.map((x,i)=>`<div class="mix-row" style="display:grid;grid-template-columns:96px minmax(120px,1fr) 90px 92px;gap:8px;align-items:center"><div><div class="mix-name">${x.name}</div><div style="display:flex;flex-wrap:wrap;gap:4px;margin-top:6px"><span class="sub-rank ${i>=capaRowsWireless.length-1?'warn':''}">${i+1}위</span><span style="display:inline-flex;align-items:center;padding:2px 7px;border-radius:999px;background:${x.delta>=0?'rgba(34,197,94,.14)':'rgba(239,68,68,.14)'};color:${x.delta>=0?'#86efac':'#fca5a5'};font-size:10px;font-weight:700">전기 ${fmtSubDiff(x.delta)}</span></div></div><div class="spark-bar"><div class="spark-fill" style="width:${Math.max(8,Math.min(100,x.rate))}%"></div></div><div class="mix-pct"><div>${fmtSubNum(x.value)}건</div><div style="font-size:10px;color:var(--t3);margin-top:2px">목표 ${fmtSubNum(x.target)}</div></div><div class="mix-val"><div>${x.rate.toFixed(1)}%</div><div style="font-size:10px;color:${x.rate>=100?'#86efac':'#fcd34d'};margin-top:2px">${x.rate>=100?'달성':'미달'}</div></div></div>`).join('')}</div>`},
    w_hq:{title:`🏢 무선 본부별 · ${periodLabel}`, note:`본부별 탭은 현재 Factbook 직접 시트가 없어 기본 배분모델로 표시 · ${isYear?'연 환산 기준':'월 기준'}`, html:`<div class="sub-block"><h3>무선 본부별 가입자</h3>${renderEntityRows(wirelessHq, wirelessTotal, '건','teal')}</div><div class="sub-note">※ 본부별 실측 시트 연결 전까지는 무선 전체 가입자를 기준으로 본부 배분모델을 사용합니다.</div>`},
    w_branch:{title:`📍 무선 지사별 · ${periodLabel}`, note:`지사 표시는 소매지사/도매지사 2개 축으로 단순화 · ${isYear?'연 환산 기준':'월 기준'}`, html:`<div class="sub-block"><h3>무선 지사별 가입자</h3>${renderEntityRows(wirelessBranch, wirelessTotal, '건','')}</div><div class="sub-note">※ 강남/분당 같은 임의 지사는 제거했고, 소매지사/도매지사 기준으로만 표시합니다.</div>`},
    v_summary:{title:`🛜 유선 가입자 · ${periodLabel}`, note:`유선 전체 및 인터넷/TV 현황 · ${isYear?'월 실적 12개월 환산':'당월 실적'}`, html:`
      <div class="sub-grid">
        <div class="sub-card"><div class="sub-lbl">유선 전체 가입자</div><div class="sub-kpi">${fmtSubNum(wiredTotal)}</div><div style="font-size:11px;color:${wiredNetAdd>=0?'#86efac':'#fca5a5'};font-weight:700">영업순증 ${fmtSubDiff(wiredNetAdd)}</div></div>
        <div class="sub-card"><div class="sub-lbl">인터넷 / TV</div><div class="sub-kpi">${fmtSubNum(internetVal)} / ${fmtSubNum(tvVal)}</div><div style="font-size:11px;color:var(--t3);font-weight:700">유선 핵심 구성</div></div>
      </div>
      <div class="sub-block"><h3>유선 채널별 가입자</h3>${renderEntityRows(wiredChannel, wiredTotal, '건','purple')}</div><div class="sub-note">※ 유선 채널별은 현재 전체 가입자 기준 보정모델입니다.</div>`},
    v_capa:{title:`📦 유선 CAPA · ${periodLabel}`, note:`유선 CAPA 실측 시트 미연결 상태라 추정 운영 규모로 표시 · ${isYear?'연 환산 기준':'월 기준'}`, html:`<div class="sub-block"><h3>유선 채널별 CAPA 대응 현황</h3>${capaRowsWired.map((x,i)=>`<div class="mix-row" style="display:grid;grid-template-columns:96px minmax(120px,1fr) 90px 92px;gap:8px;align-items:center"><div><div class="mix-name">${x.name}</div><div style="display:flex;flex-wrap:wrap;gap:4px;margin-top:6px"><span class="sub-rank ${i>=capaRowsWired.length-1?'warn':''}">${i+1}위</span><span style="display:inline-flex;align-items:center;padding:2px 7px;border-radius:999px;background:${x.delta>=0?'rgba(34,197,94,.14)':'rgba(239,68,68,.14)'};color:${x.delta>=0?'#86efac':'#fca5a5'};font-size:10px;font-weight:700">전기 ${fmtSubDiff(x.delta)}</span></div></div><div class="spark-bar"><div class="spark-fill teal" style="width:${Math.max(8,Math.min(100,x.rate))}%"></div></div><div class="mix-pct"><div>${fmtSubNum(x.value)}건</div><div style="font-size:10px;color:var(--t3);margin-top:2px">추정 목표 ${fmtSubNum(x.target)}</div></div><div class="mix-val"><div>${x.rate.toFixed(1)}%</div><div style="font-size:10px;color:${x.rate>=100?'#86efac':'#fcd34d'};margin-top:2px">운영 CAPA</div></div></div>`).join('')}</div><div class="sub-note">※ 유선 CAPA는 원천 시트 연결 전까지 추정 목표 기반입니다.</div>`},
    v_hq:{title:`🏢 유선 본부별 · ${periodLabel}`, note:`본부별 탭은 기본 배분모델 · ${isYear?'연 환산 기준':'월 기준'}`, html:`<div class="sub-block"><h3>유선 본부별 가입자</h3>${renderEntityRows(wiredHq, wiredTotal, '건','purple')}</div><div class="sub-note">※ 본부별 실측 데이터 연동 전 시뮬레이션 값입니다.</div>`},
    v_branch:{title:`📍 유선 지사별 · ${periodLabel}`, note:`지사 표시는 소매지사/도매지사 2개 축으로 단순화 · ${isYear?'연 환산 기준':'월 기준'}`, html:`<div class="sub-block"><h3>유선 지사별 가입자</h3>${renderEntityRows(wiredBranch, wiredTotal, '건','teal')}</div><div class="sub-note">※ 유선도 소매지사/도매지사 기준으로만 보여줍니다.</div>`}
  };
}
function switchSubscriberTopTab(key){
  window.__subTopTab = key;
  initSubscriberUI();
}
function switchSubscriberPeriodMode(mode){
  window.__subPeriodMode = mode;
  initSubscriberUI();
}
function buildSubscriberAggregate(d){
  const schema=ensureSubscriberSchema(d);
  const records=schema.records||[];
  const months=schema.monthLabels||d.months||[];
  const idx=Math.max(0,months.length-1);
  const currentMonth=months[idx]||d.baseMonth||'당월';
  const m251=(records.find(r=>r.channel==='무선' && r.month===1) || {});
  const mCur=(records.find(r=>r.channel==='무선' && r.month===months.length) || {});
  const prev=(records.find(r=>r.channel==='무선' && r.month===Math.max(1,months.length-1)) || {});
  const y24 = Number(((d||{}).wireless||{}).y24||0) || Number((((d||{}).wireless||{}).series||{}).total?.[0]||0);
  return {schema,records,months,currentMonth,m251,mCur,prev,y24,idx};
}
function renderSubscriberTable(rows, cols){
  return `<div style="overflow-x:auto;-webkit-overflow-scrolling:touch"><table style="width:100%;border-collapse:collapse;font-size:11px;table-layout:auto"><thead><tr style="background:#0f172a;color:#fff">${cols.map(c=>`<th style="padding:7px 10px;text-align:${c.align||'right'};white-space:nowrap">${c.label}</th>`).join('')}</tr></thead><tbody>${rows.map(r=>`<tr>${cols.map(c=>`<td style="padding:7px 10px;border-bottom:1px solid #1f2937;text-align:${c.align||'right'};font-weight:${c.align==='left'?'700':'600'};white-space:nowrap">${c.f(r)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
}
function switchSubscriberTopTab(key){ window.__subTopTab=key; initSubscriberUI(); }
function switchSubscriberPeriodMode(mode){ window.__subPeriodMode=mode; initSubscriberUI(); }
function initSubscriberUI(){
  const d=subscriberData;
  const a=buildSubscriberAggregate(d);
  const monthIdx=a.idx, prevIdx=Math.max(0,monthIdx-1);
  const cur=a.mCur;
  const totalNew=Number(cur.newSubs)||0;
  const totalChg=Number(cur.changeSubs)||0;
  const totalCancel=Number(cur.cancelSubs)||0;
  const totalChgOut=Number(cur.changeOut)||0;
  const totalMaturity=Number(cur.maturityReturn)||0;
  const capaCalc=Number((((d||{}).wireless||{}).series||{}).capa?.[monthIdx]||0)||(totalNew+totalChg);
  const netCalc=capaCalc-totalCancel-totalChgOut-totalMaturity;
  const netRef=Number(cur.netAdd)||0;
  const endCur=Number(cur.endingSubs)||0;
  const endPrev=Number(a.prev.endingSubs)||0;
  const netByDelta=endCur-endPrev;
  const yoy=a.y24?((endCur-a.y24)/a.y24*100):0;

  const org=((d.wireless||{}).org)||{};
  const hqNames=['강북본부','강남본부','강서본부','동부본부','서부본부'];
  const retailBranches=(a.schema.retailBranches||[]);
  const wholesaleBranches=(a.schema.wholesaleBranches||[]);
  const hqRows=hqNames.map(h=>({hq:h,retail:Number((org.retailHQ&&org.retailHQ[h]&&org.retailHQ[h][monthIdx])||0),wholesale:Number((org.wholesaleHQ&&org.wholesaleHQ[h]&&org.wholesaleHQ[h][monthIdx])||0)})).map(x=>({...x,total:x.retail+x.wholesale}));
  const retailRows=retailBranches.map(b=>({branch:b,capa:Number((org.retailBranch&&org.retailBranch[b]&&org.retailBranch[b][monthIdx])||0)}));
  const wholesaleRows=wholesaleBranches.map(b=>({branch:b,capa:Number((org.wholesaleBranch&&org.wholesaleBranch[b]&&org.wholesaleBranch[b][monthIdx])||0)}));

  const channelRows=Object.keys(d.channelMix||{}).map(k=>{
    const curV=Number(d.channelMix[k])||0;
    const prevV=Number(((d.channelSeries||{})[k]||[])[prevIdx]||0);
    return {channel:k,capa:curV,delta:curV-prevV,share:safePct(curV,Object.values(d.channelMix||{}).reduce((s,v)=>s+(Number(v)||0),0)||1)};
  }).sort((x,y)=>y.capa-x.capa);

  const arpuCur=Number((((d||{}).arpu||{}).series||[])[monthIdx]||0);
  const arpuPrev=Number((((d||{}).arpu||{}).series||[])[prevIdx]||0);
  const arpuDelta=arpuCur-arpuPrev;
  const wCapa=Number((((d||{}).wireless||{}).series||{}).capa?.[monthIdx]||0);
  const wCapaPrev=Number((((d||{}).wireless||{}).series||{}).capa?.[prevIdx]||0);
  const wiredCapa=Number((((d||{}).wired||{}).series||{}).newSubs?.[monthIdx]||0);

  const mgmtSeries=(((d||{}).wireless||{}).series||{}).mgmtEligible||[];
  const hcSeries=(((d||{}).wireless||{}).series||{}).hcEligible||[];
  const mgmtCur=Number(mgmtSeries[monthIdx]||0), mgmtPrev=Number(mgmtSeries[prevIdx]||0);
  const hcCur=Number(hcSeries[monthIdx]||0), hcPrev=Number(hcSeries[prevIdx]||0);

  // ── org 데이터 가져오기 (업로드 실측 vs DEFAULT) ──
  const orgData = d.org || {};
  // 조직별 행 렌더 헬퍼
  const renderOrgTable=(rows,valKey,tgtKey,y24Key,unit='건',color='')=>{
    if(!rows||!rows.length) return '<div class="sub-note">데이터 없음</div>';
    const total=rows.reduce((s,r)=>s+(Number(r[valKey])||0),0)||1;
    return rows.map(r=>{
      const v=Number(r[valKey])||0;
      const tgt=Number(r[tgtKey])||0;
      const y24=Number(r[y24Key])||0;
      const rate=tgt?safePct(v,tgt):0;
      const share=safePct(v,total);
      const cls=rate>=100?'ok':rate>=85?'warn':'bad';
      return `<div style="display:grid;grid-template-columns:90px 1fr 95px 70px;gap:6px;align-items:center;padding:9px 0;border-bottom:1px solid rgba(255,255,255,.06)">
        <div><div style="font-weight:700;font-size:13px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${r.nm}</div>
          <div style="font-size:10px;color:var(--t3);margin-top:2px">'24년 ${fmtSubNum(y24)}</div>
        </div>
        <div>
          <div class="spark-bar" style="margin-bottom:3px"><div class="spark-fill ${color}" style="width:${Math.max(4,Math.min(100,rate))}%"></div></div>
          <div style="font-size:10px;color:var(--t3)">비중 ${share.toFixed(1)}% · 목표 ${fmtSubNum(tgt)}</div>
        </div>
        <div style="text-align:right">
          <div style="font-size:14px;font-weight:700">${fmtSubNum(v)}<span style="font-size:10px;font-weight:400;margin-left:2px">${unit}</span></div>
          <div style="font-size:10px;color:${v-tgt>=0?'#86efac':'#fca5a5'};font-weight:700">${fmtSubDiff(v-tgt)}</div>
        </div>
        <div style="text-align:center">
          <div style="font-size:13px;font-weight:700;color:${cls==='ok'?'#86efac':cls==='warn'?'#fcd34d':'#fca5a5'}">${rate.toFixed(1)}%</div>
          <div class="sub-rank ${cls==='bad'?'warn':''}" style="margin-top:3px;display:inline-block;font-size:10px">${cls==='ok'?'달성':cls==='warn'?'근접':'미달'}</div>
        </div>
      </div>`;
    }).join('');
  };

  // DEFAULT or uploaded 데이터에서 org 추출
  const getOrgMap=(key)=>{
    // 업로드 시 실측 org에서, 없으면 DEFAULT org에서
    const uploaded = (d.org||{})[key];
    const def = DEFAULT_SUBSCRIBER_DATA.org[key];
    return uploaded||def||{};
  };

  // 무선 소매 본부 (업로드 시 series에서 마지막 달, DEFAULT는 단일값)
  const buildHQRows=(mapKey, valField)=>{
    const map = getOrgMap(mapKey);
    return Object.entries(map).map(([nm,v])=>{
      let val = 0;
      if(v.series && Array.isArray(v.series)) val=v.series[Math.min(monthIdx,v.series.length-1)]||0;
      else val=v[valField]||0;
      return {nm, [valField]:val, target:v.target||0, y24:v.y24||0};
    });
  };

  const wRetailHQRows = buildHQRows('wirelessRetailHQ','capa');
  const wWholesaleHQRows = buildHQRows('wirelessWholesaleHQ','capa');
  const wRetailBrRows = buildHQRows('wirelessRetailBranch','capa');
  const wWholesaleBrRows = buildHQRows('wirelessWholesaleBranch','capa');
  const vRetailHQRows = buildHQRows('wiredRetailHQ','netAdd');
  const vWholesaleHQRows = buildHQRows('wiredWholesaleHQ','netAdd');
  const vRetailBrRows = buildHQRows('wiredRetailBranch','netAdd');
  const vWholesaleBrRows = buildHQRows('wiredWholesaleBranch','netAdd');

  const wRetailTotal=wRetailHQRows.reduce((s,r)=>s+r.capa,0);
  const wWholesaleTotal=wWholesaleHQRows.reduce((s,r)=>s+r.capa,0);
  const vRetailTotal=vRetailHQRows.reduce((s,r)=>s+r.netAdd,0);
  const vWholesaleTotal=vWholesaleHQRows.reduce((s,r)=>s+r.netAdd,0);

    const tabs={
    w_summary:`<div class="sub-grid"><div class="sub-card"><div class="sub-lbl">24년 계(24.12)</div><div class="sub-kpi">${fmtSubNum(a.y24)}</div></div><div class="sub-card"><div class="sub-lbl">25.1월</div><div class="sub-kpi">${fmtSubNum(a.m251.endingSubs||0)}</div></div><div class="sub-card"><div class="sub-lbl">누적(25.12)</div><div class="sub-kpi">${fmtSubNum(endCur)}</div></div><div class="sub-card"><div class="sub-lbl">YoY</div><div class="sub-kpi">${yoy>=0?'+':''}${yoy.toFixed(2)}%</div></div></div><div class="sub-block"><h3>순증 검증</h3>${renderSubscriberTable([{k:'순증 참조값(Factbook 순증행)',v:netRef},{k:'월말 가입자 증감(12월-11월)',v:netByDelta},{k:'산식 계산값(CAPA-해지-기변OUT-만기환수)',v:netCalc}],[{label:'검증항목',align:'left',f:r=>r.k},{label:'값',f:r=>fmtSubNum(r.v)}])}<div class="sub-note" style="color:${(netRef===netByDelta&&netRef===netCalc)?'#86efac':'#fca5a5'}">${(netRef===netByDelta&&netRef===netCalc)?'검증 일치':'검증 불일치(원천 확인 필요)'}</div></div><div class="sub-block"><h3>순증 요인 분해</h3><div class="sub-mini" style="grid-template-columns:repeat(3,1fr)"><div class="sub-mini-item"><span class="sub-lbl">신규</span><b>${fmtSubNum(totalNew)}</b></div><div class="sub-mini-item"><span class="sub-lbl">기변</span><b>${fmtSubNum(totalChg)}</b></div><div class="sub-mini-item"><span class="sub-lbl">해지</span><b>${fmtSubNum(totalCancel)}</b></div><div class="sub-mini-item"><span class="sub-lbl">기변OUT</span><b>${fmtSubNum(totalChgOut)}</b></div><div class="sub-mini-item"><span class="sub-lbl">만기환수</span><b>${fmtSubNum(totalMaturity)}</b></div><div class="sub-mini-item"><span class="sub-lbl">CAPA</span><b>${fmtSubNum(capaCalc)}</b></div></div></div>`,
    w_capa:`<div class="sub-grid" style="grid-template-columns:repeat(2,1fr)"><div class="sub-card"><div class="sub-lbl">무선 CAPA(일반후불)</div><div class="sub-kpi">${fmtSubNum(wCapa)}</div><div style="font-size:11px;color:${wCapa-wCapaPrev>=0?'#86efac':'#fca5a5'}">전월대비 ${fmtSubDiff(wCapa-wCapaPrev)}</div></div><div class="sub-card"><div class="sub-lbl">무선 ARPU</div><div class="sub-kpi">${fmtWon(arpuCur)}</div><div style="font-size:11px;color:${arpuDelta>=0?'#86efac':'#fca5a5'}">전월대비 ${(arpuDelta>=0?'+':'')+Math.round(arpuDelta).toLocaleString('ko-KR')}원</div></div></div><div class="sub-block"><h3>무선 채널별 CAPA/기여</h3>${renderSubscriberTable(channelRows,[{label:'채널',align:'left',f:r=>r.channel},{label:'CAPA',f:r=>fmtSubNum(r.capa)},{label:'전월대비',f:r=>fmtSubDiff(r.delta)},{label:'비중',f:r=>r.share.toFixed(1)+'%'}])}</div>`,
    w_hq:`<div class="sub-grid">
      <div class="sub-card"><div class="sub-lbl">소매 CAPA 합계</div><div class="sub-kpi">${fmtSubNum(wRetailTotal)}</div><div style="font-size:11px;color:${safePct(wRetailTotal,16000)>=100?'#86efac':'#fca5a5'};font-weight:700">목표 16,000 / ${safePct(wRetailTotal,16000).toFixed(1)}%</div></div>
      <div class="sub-card"><div class="sub-lbl">도매 CAPA 합계</div><div class="sub-kpi">${fmtSubNum(wWholesaleTotal)}</div><div style="font-size:11px;color:${safePct(wWholesaleTotal,25000)>=100?'#86efac':'#fca5a5'};font-weight:700">목표 25,000 / ${safePct(wWholesaleTotal,25000).toFixed(1)}%</div></div>
    </div>
    <div class="sub-block"><h3>📶 무선 소매 본부별 CAPA (2.무선가입자 실측)</h3>
      ${renderOrgTable(wRetailHQRows,'capa','target','y24','건','')}
    </div>
    <div class="sub-block"><h3>📶 무선 도매 본부별 CAPA (2.무선가입자 실측)</h3>
      ${renderOrgTable(wWholesaleHQRows,'capa','target','y24','건','teal')}
    </div>
    <div class="sub-note">※ 실측 기준: 종합 경영성과 Factbook 25.12월 · 업로드 시 최신값 자동 반영</div>`,
    w_branch:`<div class="sub-board">
      <div class="sub-block"><h3>📶 무선 소매 지사별 CAPA (10개 지사)</h3>
        ${renderOrgTable(wRetailBrRows,'capa','target','y24','건','')}
      </div>
      <div class="sub-block"><h3>📶 무선 도매 지사별 CAPA (8개 지사)</h3>
        ${renderOrgTable(wWholesaleBrRows,'capa','target','y24','건','teal')}
      </div>
    </div>
    <div class="sub-note">※ 강원지사(도매)는 25년 수치 공백. 목표대비 순위는 소매: 부산 -595(-28.5%), 동부 -937(-23.4%) 주의</div>`,
    w_mgmt:(()=>{
      const md = d.mgmt || {};
      const ms = md.series || {};
      const mc = md.channels || {};
      const mTotal = ms.total || mgmtSeries;
      const mNet = ms.netAdd || [];
      const mCapa = ms.capa || [];
      const mChurn = ms.churn || [];
      const mChgOut = ms.chgOut || [];
      const mMat = ms.maturity || [];
      const mCur = Number(mTotal[monthIdx]||0), mPrev = Number(mTotal[prevIdx]||0);
      const mNetCur = Number(mNet[monthIdx]||0);

      // 5채널 구성
      const CH_MGMT = ['소매8.5%','소상공인8.5%','도매8.5%','온라인제휴8.5%','B2B8.5%'];
      const CH_COLOR = {'소매8.5%':'#3b82f6','소상공인8.5%':'#10b981','도매8.5%':'#f59e0b','온라인제휴8.5%':'#8b5cf6','B2B8.5%':'#ef4444'};
      const CH_NOTE = {'도매8.5%':'CAPA없음(소멸중)','온라인제휴8.5%':'CAPA없음(소멸중)','소매8.5%':'약 73%','소상공인8.5%':'약 2%','B2B8.5%':'약 12%'};

      // 채널별 최신월 요약 카드
      const chSummaryCards = CH_MGMT.map(ch=>{
        const cd = mc[ch]||{};
        const t = cd.total||[]; const n = cd.netAdd||[];
        const tCur = Number(t[monthIdx]||0), tPrev = Number(t[prevIdx]||0);
        const nCur = Number(n[monthIdx]||0);
        const color = CH_COLOR[ch];
        const totalAll = CH_MGMT.reduce((s,k)=>s+Number((mc[k]||{}).total?.[monthIdx]||0),0)||1;
        const share = ((tCur/totalAll)*100).toFixed(1);
        return `<div class="sub-card" style="border-top:3px solid ${color}">
          <div class="sub-lbl" style="color:${color};font-weight:700">${ch}</div>
          <div class="sub-kpi" style="font-size:20px">${fmtSubNum(tCur)}</div>
          <div style="font-size:11px;color:var(--t3);margin-top:3px">비중 ${share}% · ${CH_NOTE[ch]||''}</div>
          <div style="font-size:11px;margin-top:4px;color:${nCur>=0?'#86efac':'#fca5a5'};font-weight:700">순증 ${fmtSubDiff(nCur)}</div>
          <div style="font-size:10px;color:${tCur-tPrev>=0?'#86efac':'#fca5a5'}">전월 ${fmtSubDiff(tCur-tPrev)}</div>
        </div>`;
      }).join('');

      // 채널별 순증 월별 비교 바차트
      const makeBarChart = (channels, key='netAdd')=>{
        const months12 = a.months||[];
        return `<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:11px;min-width:700px">
          <thead><tr style="background:#0f172a;color:#fff">
            <th style="padding:6px 10px;text-align:left;min-width:90px">채널</th>
            ${months12.map(m=>`<th style="padding:6px 5px;text-align:center;min-width:55px">${m}</th>`).join('')}
          </tr></thead>
          <tbody>
          ${channels.map(ch=>{
            const cd = mc[ch]||{}; const arr = cd[key]||[];
            const maxAbs = Math.max(1,...arr.map(v=>Math.abs(Number(v)||0)));
            const color = CH_COLOR[ch];
            return `<tr style="border-bottom:1px solid #1f2937">
              <td style="padding:6px 10px;font-weight:700;color:${color}">${ch}</td>
              ${months12.map((_,i)=>{
                const v=Number(arr[i]||0);
                const w=Math.round(Math.abs(v)/maxAbs*50);
                const bg=v>=0?'#86efac':'#fca5a5';
                return `<td style="padding:4px 5px;text-align:center">
                  <div style="font-size:10px;font-weight:700;color:${bg};margin-bottom:2px">${v>=0?'+':''}${(v>=1000||v<=-1000)?(v/1000).toFixed(1)+'K':v}</div>
                  <div style="height:6px;background:${bg};width:${Math.max(3,w)}%;margin:0 auto;border-radius:3px"></div>
                </td>`;
              }).join('')}
            </tr>`;
          }).join('')}
          </tbody>
        </table></div>`;
      };

      // 채널별 유지(stock) 월별 추이 테이블
      const stockRows = (a.months||[]).map((m,i)=>({
        m,
        total: Number(mTotal[i]||0),
        ...Object.fromEntries(CH_MGMT.map(ch=>[ch, Number((mc[ch]||{}).total?.[i]||0)]))
      }));

      return `<div class="sub-grid" style="grid-template-columns:repeat(2,1fr)">${chSummaryCards}</div>
      <div class="sub-block">
        <h3>📊 관리수수료(8.5%) 전체 유지 월별 추이</h3>
        ${renderSubscriberTable((a.months||[]).map((m,i)=>({m,total:Number(mTotal[i]||0),capa:Number(mCapa[i]||0),churn:Number(mChurn[i]||0),chgOut:Number(mChgOut[i]||0),mat:Number(mMat[i]||0),net:Number(mNet[i]||0)})),[
          {label:'월',align:'left',f:r=>r.m},
          {label:'유지(재고)',f:r=>fmtSubNum(r.total)},
          {label:'CAPA',f:r=>fmtSubNum(r.capa)},
          {label:'해지',f:r=>fmtSubNum(r.churn)},
          {label:'기변OUT',f:r=>fmtSubNum(r.chgOut)},
          {label:'만기환수',f:r=>fmtSubNum(r.mat)},
          {label:'순증',f:r=>`<span style="color:${r.net>=0?'#86efac':'#fca5a5'};font-weight:700">${fmtSubDiff(r.net)}</span>`}
        ])}
      </div>
      <div class="sub-block">
        <h3>📂 채널별 가입자 재고(유지) 월별 현황</h3>
        ${renderSubscriberTable(stockRows,[
          {label:'월',align:'left',f:r=>r.m},
          ...CH_MGMT.map(ch=>({label:ch.replace('8.5%',''),f:r=>`<span style="color:${CH_COLOR[ch]}">${fmtSubNum(r[ch])}</span>`})),
          {label:'합계',f:r=>`<b>${fmtSubNum(r.total)}</b>`}
        ])}
      </div>
      <div class="sub-block">
        <h3>📉 채널별 순증 월별 비교 (양수=증가 / 음수=감소)</h3>
        ${makeBarChart(CH_MGMT,'netAdd')}
      </div>
      <div class="sub-block">
        <h3>🎯 채널별 CAPA 월별 비교</h3>
        ${makeBarChart(['소매8.5%','소상공인8.5%','B2B8.5%'],'capa')}
        <div class="sub-note">※ 도매8.5%·온라인제휴8.5%는 신규 CAPA 없음 — 기존 재고 자연 소멸 중</div>
      </div>
      <div class="sub-note">⚠️ 25.1월 108.1만→25.12월 99.6만 (△8.1만, △7.5%). 소매8.5%가 전체의 73% 점유, 소멸 채널(도매+온라인제휴 약 12%)이 매월 감소분을 가속.</div>`;
    })(),
    w_hc:(()=>{
      const hd = d.hc || {};
      const hs = hd.series || {};
      const hch = hd.channels || {};
      const hTotal = hs.total || hcSeries;
      const hNet = hs.netAdd || [];
      const hCapa = hs.capa || [];
      const hChurn = hs.churn || [];
      const hChgOut = hs.chgOut || [];
      const hMat = hs.maturity || [];
      const hCurV = Number(hTotal[monthIdx]||0), hPrevV = Number(hTotal[prevIdx]||0);
      const hNetCur = Number(hNet[monthIdx]||0);

      const CH_HC = ['도매H/C2%','KT닷컴H/C2%'];
      const HC_COLOR = {'도매H/C2%':'#f59e0b','KT닷컴H/C2%':'#3b82f6'};
      const HC_NOTE = {'도매H/C2%':'매월 △1.5~2.5만 급감','KT닷컴H/C2%':'유일 성장 채널 ↑'};

      // 채널 요약 카드
      const hcCards = CH_HC.map(ch=>{
        const cd = hch[ch]||{};
        const t = cd.total||[]; const n = cd.netAdd||[]; const cap = cd.capa||[];
        const tCur = Number(t[monthIdx]||0), tPrev = Number(t[prevIdx]||0);
        const nCur = Number(n[monthIdx]||0);
        const capCur = Number(cap[monthIdx]||0);
        const color = HC_COLOR[ch];
        const totalHC = CH_HC.reduce((s,k)=>s+Number((hch[k]||{}).total?.[monthIdx]||0),0)||1;
        const share = ((tCur/totalHC)*100).toFixed(1);
        return `<div class="sub-card" style="border-top:3px solid ${color}">
          <div class="sub-lbl" style="color:${color};font-weight:700">${ch}</div>
          <div class="sub-kpi" style="font-size:22px">${fmtSubNum(tCur)}</div>
          <div style="font-size:11px;color:var(--t3);margin-top:3px">비중 ${share}% · ${HC_NOTE[ch]||''}</div>
          <div style="font-size:11px;margin-top:4px;color:${nCur>=0?'#86efac':'#fca5a5'};font-weight:700">순증 ${fmtSubDiff(nCur)}</div>
          <div style="font-size:10px;color:var(--t3)">전월 ${fmtSubDiff(tCur-tPrev)} · CAPA ${fmtSubNum(capCur)}</div>
        </div>`;
      }).join('');

      // 채널별 상세 테이블 생성 헬퍼
      const makeHcDetailTable = (ch)=>{
        const cd = hch[ch]||{};
        const rows = (a.months||[]).map((m,i)=>({
          m,
          total:Number((cd.total||[])[i]||0),
          capa:Number((cd.capa||[])[i]||0),
          churn:Number((cd.churn||[])[i]||0),
          chgOut:Number((cd.chgOut||[])[i]||0),
          mat:Number((cd.maturity||[])[i]||0),
          net:Number((cd.netAdd||[])[i]||0)
        }));
        return renderSubscriberTable(rows,[
          {label:'월',align:'left',f:r=>r.m},
          {label:'유지(재고)',f:r=>fmtSubNum(r.total)},
          {label:'CAPA',f:r=>fmtSubNum(r.capa)},
          {label:'해지',f:r=>fmtSubNum(r.churn)},
          {label:'기변OUT',f:r=>fmtSubNum(r.chgOut)},
          {label:'만기환수',f:r=>fmtSubNum(r.mat)},
          {label:'순증',f:r=>`<span style="color:${r.net>=0?'#86efac':'#fca5a5'};font-weight:700">${fmtSubDiff(r.net)}</span>`}
        ]);
      };

      // 채널간 비교 바차트 (순증)
      const months12 = a.months||[];
      const compChart = `<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:11px;min-width:700px">
        <thead><tr style="background:#0f172a;color:#fff">
          <th style="padding:6px 10px;text-align:left">채널</th>
          ${months12.map(m=>`<th style="padding:6px 5px;text-align:center;min-width:55px">${m}</th>`).join('')}
        </tr></thead>
        <tbody>
        ${CH_HC.map(ch=>{
          const cd=hch[ch]||{}; const arr=(cd.netAdd||[]);
          const maxAbs=Math.max(1,...arr.map(v=>Math.abs(Number(v)||0)));
          const color=HC_COLOR[ch];
          return `<tr style="border-bottom:1px solid #1f2937">
            <td style="padding:6px 10px;font-weight:700;color:${color}">${ch}</td>
            ${months12.map((_,i)=>{
              const v=Number(arr[i]||0);
              const w=Math.round(Math.abs(v)/maxAbs*50);
              const bg=v>=0?'#86efac':'#fca5a5';
              return `<td style="padding:4px 5px;text-align:center">
                <div style="font-size:10px;font-weight:700;color:${bg};margin-bottom:2px">${v>=0?'+':''}${(Math.abs(v)>=1000)?(v/1000).toFixed(1)+'K':v}</div>
                <div style="height:6px;background:${bg};width:${Math.max(3,w)}%;margin:0 auto;border-radius:3px"></div>
              </td>`;
            }).join('')}
          </tr>`;
        }).join('')}
        </tbody>
      </table></div>`;

      // 도매 vs KT닷컴 유지 합산 비교
      const stockCompRows = months12.map((m,i)=>{
        const domae = Number((hch['도매H/C2%']||{}).total?.[i]||0);
        const ktdot = Number((hch['KT닷컴H/C2%']||{}).total?.[i]||0);
        return {m, domae, ktdot, total:Number(hTotal[i]||0), gap:domae-ktdot};
      });

      return `<div class="sub-grid" style="grid-template-columns:repeat(2,1fr)">${hcCards}</div>
      <div class="sub-block">
        <h3>📊 H/C(2%) 전체 유지 월별 상세</h3>
        ${renderSubscriberTable((a.months||[]).map((m,i)=>({m,total:Number(hTotal[i]||0),capa:Number(hCapa[i]||0),churn:Number(hChurn[i]||0),chgOut:Number(hChgOut[i]||0),mat:Number(hMat[i]||0),net:Number(hNet[i]||0)})),[
          {label:'월',align:'left',f:r=>r.m},
          {label:'유지(재고)',f:r=>fmtSubNum(r.total)},
          {label:'CAPA',f:r=>fmtSubNum(r.capa)},
          {label:'해지',f:r=>fmtSubNum(r.churn)},
          {label:'기변OUT',f:r=>fmtSubNum(r.chgOut)},
          {label:'만기환수',f:r=>fmtSubNum(r.mat)},
          {label:'순증',f:r=>`<span style="color:${r.net>=0?'#86efac':'#fca5a5'};font-weight:700">${fmtSubDiff(r.net)}</span>`}
        ])}
      </div>
      <div class="sub-block">
        <h3>🔀 도매H/C vs KT닷컴H/C 유지 비교</h3>
        ${renderSubscriberTable(stockCompRows,[
          {label:'월',align:'left',f:r=>r.m},
          {label:'도매H/C(감소)',f:r=>`<span style="color:#f59e0b;font-weight:700">${fmtSubNum(r.domae)}</span>`},
          {label:'KT닷컴H/C(증가)',f:r=>`<span style="color:#3b82f6;font-weight:700">${fmtSubNum(r.ktdot)}</span>`},
          {label:'도매-KT닷컴 격차',f:r=>`<span style="color:${r.gap>0?'#fca5a5':'#86efac'};font-weight:700">${fmtSubDiff(r.gap)}</span>`},
          {label:'H/C 합계',f:r=>fmtSubNum(r.total)}
        ])}
        <div class="sub-note">⚠️ 격차가 좁혀지는 중: 25.1월 도매-KT닷컴 = +44.3만 → 25.12월 +8.3만. 연말 역전 임박.</div>
      </div>
      <div class="sub-block">
        <h3>📉 순증 채널 비교 (도매↓ vs KT닷컴↑)</h3>
        ${compChart}
      </div>
      <div class="sub-block">
        <h3>📋 도매H/C2% 상세 (소멸 채널)</h3>
        ${makeHcDetailTable('도매H/C2%')}
      </div>
      <div class="sub-block">
        <h3>📋 KT닷컴H/C2% 상세 (성장 채널)</h3>
        ${makeHcDetailTable('KT닷컴H/C2%')}
      </div>
      <div class="sub-note">⚠️ H/C 전체 25.1월 72.2만→25.12월 65.1만 (△7.1만, △9.8%). 도매H/C 소멸이 KT닷컴H/C 성장을 초과. 기변OUT&gt;CAPA 구조 지속.</div>`;
    })(),
    v_summary:(()=>{
      const ws = d.wired||{};
      const wSeries = ws.series||{};
      const wChNA = ws.channelNetAdd || DEFAULT_SUBSCRIBER_DATA.wired.channelNetAdd || {};
      const wTotal = Number((wSeries.total||[])[monthIdx]||ws.total||0);
      const wPrev  = Number((wSeries.total||[])[prevIdx]||0);
      const wNetOp = Number((wSeries.netAdd||[])[monthIdx]||ws.netAdd||0);
      const wNet   = Number((wSeries.netPure||[])[monthIdx]||0);
      const iTotal = Number((wSeries.internet||[])[monthIdx]||ws.internet||0);
      const tvTotal= Number((wSeries.tv||[])[monthIdx]||ws.tv||0);
      const iNet   = Number((wSeries.iNet||[])[monthIdx]||0);
      const tvNet  = Number((wSeries.tvNet||[])[monthIdx]||0);
      const iNetOp = Number((wSeries.iNetOp||[])[monthIdx]||0);
      const tvNetOp= Number((wSeries.tvNetOp||[])[monthIdx]||0);

      // 유선 월별 상세 테이블
      const vDetailRows = (a.months||[]).map((m,i)=>({
        m,
        total: Number((wSeries.total||[])[i]||0),
        newSubs: Number((wSeries.newSubs||[])[i]||0),
        cancel: Number((wSeries.cancelSubs||[])[i]||0),
        mat: Number((wSeries.maturity||[])[i]||0),
        netOp: Number((wSeries.netAdd||[])[i]||0),
        net: Number((wSeries.netPure||[])[i]||0),
      }));

      // 인터넷/TV 세부
      const iTvRows = (a.months||[]).map((m,i)=>({
        m,
        iTotal: Number((wSeries.internet||[])[i]||0),
        iNew:   Number((wSeries.iNew||[])[i]||0),
        iCancel:Number((wSeries.iCancel||[])[i]||0),
        iMat:   Number((wSeries.iMat||[])[i]||0),
        iNet:   Number((wSeries.iNet||[])[i]||0),
        tvTotal:Number((wSeries.tv||[])[i]||0),
        tvNew:  Number((wSeries.tvNew||[])[i]||0),
        tvCancel:Number((wSeries.tvCancel||[])[i]||0),
        tvMat:  Number((wSeries.tvMat||[])[i]||0),
        tvNet:  Number((wSeries.tvNet||[])[i]||0),
      }));

      // 채널별 순신규 최신월
      const chNetAddCur = Object.entries(wChNA.itv||{}).map(([ch,arr])=>({
        ch,
        itv: Number(arr[monthIdx]||0),
        internet: Number((wChNA.internet||{})[ch]?.[monthIdx]||0),
        tv: Number((wChNA.tv||{})[ch]?.[monthIdx]||0),
      }));

      return `<div class="sub-grid">
        <div class="sub-card"><div class="sub-lbl">유선 전체(월말)</div><div class="sub-kpi">${fmtSubNum(wTotal)}</div><div style="font-size:11px;color:${wTotal-wPrev>=0?'#86efac':'#fca5a5'}">전월대비 ${fmtSubDiff(wTotal-wPrev)}</div></div>
        <div class="sub-card"><div class="sub-lbl">영업순증</div><div class="sub-kpi" style="color:${wNetOp>=0?'#86efac':'#fca5a5'}">${fmtSubDiff(wNetOp)}</div><div style="font-size:11px;color:var(--t3)">신규-해지</div></div>
        <div class="sub-card"><div class="sub-lbl">순증(만기환수 포함)</div><div class="sub-kpi" style="color:${wNet>=0?'#86efac':'#fca5a5'}">${fmtSubDiff(wNet)}</div><div style="font-size:11px;color:var(--t3)">영업순증-만기환수</div></div>
        <div class="sub-card"><div class="sub-lbl">인터넷 / TV</div><div class="sub-kpi" style="font-size:18px">${fmtSubNum(iTotal)} / ${fmtSubNum(tvTotal)}</div><div style="font-size:11px;color:var(--t3)">순증 ${fmtSubDiff(iNet)} / ${fmtSubDiff(tvNet)}</div></div>
      </div>
      <div class="sub-block">
        <h3>📋 유선 전체 월별 상세</h3>
        ${renderSubscriberTable(vDetailRows,[
          {label:'월',align:'left',f:r=>r.m},
          {label:'유지(재고)',f:r=>fmtSubNum(r.total)},
          {label:'신규',f:r=>fmtSubNum(r.newSubs)},
          {label:'해지',f:r=>fmtSubNum(r.cancel)},
          {label:'만기환수',f:r=>fmtSubNum(r.mat)},
          {label:'영업순증',f:r=>`<span style="color:${r.netOp>=0?'#86efac':'#fca5a5'};font-weight:700">${fmtSubDiff(r.netOp)}</span>`},
          {label:'순증',f:r=>`<span style="color:${r.net>=0?'#86efac':'#fca5a5'};font-weight:700">${fmtSubDiff(r.net)}</span>`}
        ])}
      </div>
      <div class="sub-block">
        <h3>🌐 인터넷 월별 상세</h3>
        ${renderSubscriberTable(iTvRows,[
          {label:'월',align:'left',f:r=>r.m},
          {label:'인터넷 유지',f:r=>fmtSubNum(r.iTotal)},
          {label:'신규',f:r=>fmtSubNum(r.iNew)},
          {label:'해지',f:r=>fmtSubNum(r.iCancel)},
          {label:'만기환수',f:r=>fmtSubNum(r.iMat)},
          {label:'순증',f:r=>`<span style="color:${r.iNet>=0?'#86efac':'#fca5a5'};font-weight:700">${fmtSubDiff(r.iNet)}</span>`}
        ])}
      </div>
      <div class="sub-block">
        <h3>📺 TV 월별 상세</h3>
        ${renderSubscriberTable(iTvRows,[
          {label:'월',align:'left',f:r=>r.m},
          {label:'TV 유지',f:r=>fmtSubNum(r.tvTotal)},
          {label:'신규',f:r=>fmtSubNum(r.tvNew)},
          {label:'해지',f:r=>fmtSubNum(r.tvCancel)},
          {label:'만기환수',f:r=>fmtSubNum(r.tvMat)},
          {label:'순증',f:r=>`<span style="color:${r.tvNet>=0?'#86efac':'#fca5a5'};font-weight:700">${fmtSubDiff(r.tvNet)}</span>`}
        ])}
      </div>
      <div class="sub-note">⚠️ 25.12월 유선 유지 이상(265,343): 11월 대비 +10,863 급증, 가마감 오류 추정. 확정 마감 후 재확인 필요.</div>`;
    })(),
    v_capa:(()=>{
      const ws = d.wired||{};
      const wSeries = ws.series||{};
      const wChNA = ws.channelNetAdd || DEFAULT_SUBSCRIBER_DATA.wired.channelNetAdd || {};

      // 채널별 순신규 시리즈 테이블
      const CH_V = ['소매','소상공인','도매','디지털'];
      const CH_COLOR_V = {'소매':'#3b82f6','소상공인':'#10b981','도매':'#f59e0b','디지털':'#8b5cf6'};

      // 채널별 최신월 카드
      const chCards = CH_V.map(ch=>{
        const itvCur = Number((wChNA.itv||{})[ch]?.[monthIdx]||0);
        const itvPrev= Number((wChNA.itv||{})[ch]?.[prevIdx]||0);
        const iCur   = Number((wChNA.internet||{})[ch]?.[monthIdx]||0);
        const tvCur  = Number((wChNA.tv||{})[ch]?.[monthIdx]||0);
        const color = CH_COLOR_V[ch];
        return `<div class="sub-card" style="border-top:3px solid ${color}">
          <div class="sub-lbl" style="color:${color};font-weight:700">${ch}</div>
          <div class="sub-kpi" style="font-size:20px">${fmtSubNum(itvCur)}</div>
          <div style="font-size:11px;color:var(--t3);margin-top:3px">I+TV 순신규</div>
          <div style="font-size:10px;color:var(--t3);margin-top:3px">인터넷 ${fmtSubNum(iCur)} · TV ${fmtSubNum(tvCur)}</div>
          <div style="font-size:11px;margin-top:4px;color:${itvCur-itvPrev>=0?'#86efac':'#fca5a5'}">전월 ${fmtSubDiff(itvCur-itvPrev)}</div>
        </div>`;
      }).join('');

      // 채널별 I+TV 순신규 월별 바차트
      const months12 = a.months||[];
      const makeVChart = (type)=>{
        const src = wChNA[type]||{};
        return `<div style="overflow-x:auto;-webkit-overflow-scrolling:touch"><table style="width:100%;border-collapse:collapse;font-size:11px;min-width:620px">
          <thead><tr style="background:#0f172a;color:#fff">
            <th style="padding:6px 10px;text-align:left">채널</th>
            ${months12.map(m=>`<th style="padding:6px 5px;text-align:center;min-width:50px;white-space:nowrap">${m}</th>`).join('')}
          </tr></thead>
          <tbody>
          ${CH_V.map(ch=>{
            const arr = src[ch]||[];
            const maxV = Math.max(1,...arr.map(v=>Math.abs(Number(v)||0)));
            const color = CH_COLOR_V[ch];
            return `<tr style="border-bottom:1px solid #1f2937">
              <td style="padding:6px 10px;font-weight:700;color:${color};white-space:nowrap">${ch}</td>
              ${months12.map((_,i)=>{
                const v=Number(arr[i]||0);
                const w=Math.round(Math.abs(v)/maxV*50);
                const bg=v>=0?'#86efac':'#fca5a5';
                return `<td style="padding:4px 5px;text-align:center">
                  <div style="font-size:10px;font-weight:700;color:${bg};margin-bottom:2px">${(Math.abs(v)>=1000)?(v/1000).toFixed(1)+'K':v}</div>
                  <div style="height:5px;background:${bg};width:${Math.max(3,w)}%;margin:0 auto;border-radius:2px"></div>
                </td>`;
              }).join('')}
            </tr>`;
          }).join('')}
          </tbody>
        </table></div>`;
      };

      return `<div class="sub-grid" style="grid-template-columns:repeat(2,1fr)">${chCards}</div>
      <div class="sub-block">
        <h3>📊 채널별 I+TV 순신규 월별 추이</h3>
        ${makeVChart('itv')}
      </div>
      <div class="sub-block">
        <h3>🌐 채널별 인터넷 순신규 월별 추이</h3>
        ${makeVChart('internet')}
      </div>
      <div class="sub-block">
        <h3>📺 채널별 TV 순신규 월별 추이</h3>
        ${makeVChart('tv')}
      </div>
      <div class="sub-note">※ 유선 순신규 기준: I+TV Mass · 만기환수 미포함. 25.12월 유선 데이터 가마감 — 확정 후 재확인 필요.</div>`;
    })(),
    v_hq:`<div class="sub-grid">
      <div class="sub-card"><div class="sub-lbl">소매 순신규 합계</div><div class="sub-kpi">${fmtSubNum(vRetailTotal)}</div><div style="font-size:11px;color:${safePct(vRetailTotal,3000)>=100?'#86efac':'#fca5a5'};font-weight:700">목표 3,000 / ${safePct(vRetailTotal,3000).toFixed(1)}%</div></div>
      <div class="sub-card"><div class="sub-lbl">도매 순신규 합계</div><div class="sub-kpi">${fmtSubNum(vWholesaleTotal)}</div><div style="font-size:11px;color:${safePct(vWholesaleTotal,3800)>=100?'#86efac':'#fca5a5'};font-weight:700">목표 3,800 / ${safePct(vWholesaleTotal,3800).toFixed(1)}%</div></div>
    </div>
    <div class="sub-block"><h3>🛜 유선 소매 본부별 순신규 (3.유선가입자 실측)</h3>
      ${renderOrgTable(vRetailHQRows,'netAdd','target','y24','건','purple')}
    </div>
    <div class="sub-block"><h3>🛜 유선 도매 본부별 순신규 (3.유선가입자 실측)</h3>
      ${renderOrgTable(vWholesaleHQRows,'netAdd','target','y24','건','teal')}
    </div>
    <div class="sub-note">※ 유선 순신규 기준: I+TV Mass · 만기환수 미포함 · 실측: Factbook 25.12월</div>`,
    v_branch:`<div class="sub-board">
      <div class="sub-block"><h3>🛜 유선 소매 지사별 순신규 (10개 지사)</h3>
        ${renderOrgTable(vRetailBrRows,'netAdd','target','y24','건','purple')}
      </div>
      <div class="sub-block"><h3>🛜 유선 도매 지사별 순신규 (8개 지사)</h3>
        ${renderOrgTable(vWholesaleBrRows,'netAdd','target','y24','건','teal')}
      </div>
    </div>
    <div class="sub-note">※ 강원지사(도매) 25년 수치 공백. 유선 12월 데이터 이상 주의(가마감)</div>`
  };

  const active=(tabs[window.__subTopTab] ? window.__subTopTab : 'w_summary');
  const btn=(k,t,cls='')=>`<button class="sub-top-btn ${cls} ${active===k?'on':''}" onclick="switchSubscriberTopTab('${k}')">${t}</button>`;

  document.getElementById('subscriberPanel').innerHTML=`
    <div style="display:flex;justify-content:space-between;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:10px"><div class="sub-chip">기준월 ${a.currentMonth}</div><div class="sub-chip alt">원천: ${d.sourceFile==='업로드 Factbook'?'업로드 Factbook':'기본값'}</div><div class="sub-chip warn">순증 검증 포함</div></div>
    <div class="sub-top-tabs">${btn('w_summary','📶 무선 가입자','wless')}${btn('w_capa','🎯 무선 CAPA/ARPU','wless')}${btn('w_hq','🏢 무선 본부별','wless')}${btn('w_branch','📍 무선 지사별','wless')}${btn('w_mgmt','💼 관리수수료 대상','wless')}${btn('w_hc','🧩 H/C 대상','wless')}${btn('v_summary','🛜 유선 가입자','wired')}${btn('v_capa','📦 유선 CAPA','wired')}${btn('v_hq','🏢 유선 본부별','wired')}${btn('v_branch','📍 유선 지사별','wired')}</div>
    ${Object.keys(tabs).map(k=>`<div class="sub-tab-pane ${active===k?'on':''}" id="subview-${k}">${tabs[k]}</div>`).join('')}
  `;
}

function parseFactbookExcel(wb){
  try{
    const extracted = extractSubscriberDataFromWorkbook(wb);
    subscriberData = extracted;
    saveSubscriberData();
    if(extracted.baseMonth){
      setGlobalBaseMonth(extracted.baseMonth, 'Factbook');
      setReportPeriod(extracted.baseMonth);
      setTabMonth('subscriber', extracted.baseMonth);
    }
    initSubscriberUI();
    initDashboard();
    updateTabMonthBadges();
    showUpStatus('factbook','ok','✅ 가입자 데이터 업데이트 완료\n기준월: '+extracted.baseMonth+'\n시트: '+(extracted.sheetNames||[]).join(', '));
  }catch(err){
    showUpStatus('factbook','err','❌ '+err.message);
  }
}


function handleDrop(e,type){e.preventDefault();e.currentTarget.classList.remove('drag');if(e.dataTransfer.files[0])parseExcel(e.dataTransfer.files[0],type);}
function handleFile(input,type){if(input.files[0])parseExcel(input.files[0],type);}

const BASE_MONTH_STORAGE_KEY = 'ktms_base_month';
const UPLOADED_STATE_STORAGE_KEY = 'ktms_uploaded_state_v1';
const TAB_MONTH_STORAGE_KEY = 'ktms_tab_month_v1';
const BUNDLED_PLAN_FILE = '★★ 26년 경영계획 - 2025.09.30 v4.3 (3안 가지고 다시 일단락) (매출 8월제출 복원) v8.1 (26년 월별 목표) (1) (1).xlsx';
const BUNDLED_MIDTERM_FILE = '1. 중기 및 2026년 그룹사 경영계획(안)_kt엠앤에스 v7.81 (1) (1).xlsx';

function getSavedTabMonths(){
  try{ return JSON.parse(localStorage.getItem(TAB_MONTH_STORAGE_KEY)||'{}')||{}; }catch(e){ return {}; }
}
function setTabMonth(type, month){
  if(!type || !month) return;
  try{
    const obj = getSavedTabMonths();
    obj[type] = month;
    localStorage.setItem(TAB_MONTH_STORAGE_KEY, JSON.stringify(obj));
  }catch(e){}
  updateTabMonthBadges();
}
function clearTabMonth(type){
  if(!type) return;
  try{
    const obj = getSavedTabMonths();
    if(!(type in obj)) return;
    delete obj[type];
    localStorage.setItem(TAB_MONTH_STORAGE_KEY, JSON.stringify(obj));
  }catch(e){}
  updateTabMonthBadges();
}
function syncAllTabMonths(month){
  if(!month) return;
  try{
    const obj = getSavedTabMonths();
    ['tasks','profit','kpi','subscriber','commission','plan','simulate'].forEach((k)=>{ obj[k]=month; });
    localStorage.setItem(TAB_MONTH_STORAGE_KEY, JSON.stringify(obj));
  }catch(e){}
  D.taskMonth = month;
  if(typeof subscriberData!=='undefined' && subscriberData && typeof subscriberData==='object'){
    subscriberData.baseMonth = month;
  }
  updateTabMonthBadges();
}
function saveUploadedState(){
  try{
    const payload={
      tasks:D.tasks,
      taskMonth:D.taskMonth||null,
      profit:D.profit,
      hq:D.hq,
      kpi:D.kpi,
      kpi_retail_detail:D.kpi_retail_detail,
      kpi_wholesale_detail:D.kpi_wholesale_detail,
      variance:D.variance,
      sga_detail:D.sga_detail,
      commission:D.commission,
      planData:D.planData||null,
      midtermAssumptions:D.midtermAssumptions||null,
      scenarioState:D.scenarioState||null
    };
    localStorage.setItem(UPLOADED_STATE_STORAGE_KEY, JSON.stringify(payload));
  }catch(e){ console.warn('saveUploadedState failed', e); }
}
function loadUploadedState(){
  try{
    const raw = localStorage.getItem(UPLOADED_STATE_STORAGE_KEY);
    if(!raw) return;
    const data = JSON.parse(raw);
    if(!data || typeof data!=='object') return;
    if(Array.isArray(data.tasks) && data.tasks.length) D.tasks = data.tasks;
    if(data.taskMonth) D.taskMonth = data.taskMonth;
    if(data.profit && typeof data.profit==='object') D.profit = data.profit;
    if(Array.isArray(data.hq) && data.hq.length) D.hq = data.hq;
    if(Array.isArray(data.kpi) && data.kpi.length) D.kpi = data.kpi;
    if(data.kpi_retail_detail && typeof data.kpi_retail_detail==='object') D.kpi_retail_detail = data.kpi_retail_detail;
    if(data.kpi_wholesale_detail && typeof data.kpi_wholesale_detail==='object') D.kpi_wholesale_detail = data.kpi_wholesale_detail;
    if(data.variance && typeof data.variance==='object') D.variance = data.variance;
    if(data.sga_detail && typeof data.sga_detail==='object') D.sga_detail = data.sga_detail;
    if(data.commission && typeof data.commission==='object') D.commission = data.commission;
    if(data.planData && typeof data.planData==='object') D.planData = data.planData;
    if(data.midtermAssumptions && typeof data.midtermAssumptions==='object') D.midtermAssumptions = data.midtermAssumptions;
    if(data.scenarioState && typeof data.scenarioState==='object') D.scenarioState = data.scenarioState;
  }catch(e){ console.warn('loadUploadedState failed', e); }
}
function updateTabMonthBadges(){
  const m = getSavedTabMonths();
  const fallbackMonth = D.baseMonth || getSavedBaseMonth() || (D.period||'').replace(' 기준','') || null;
  const mapping = [
    ['tabMonthTasks', m.tasks || D.taskMonth || fallbackMonth],
    ['tabMonthProfit', m.profit || fallbackMonth],
    ['tabMonthKpi', m.kpi || fallbackMonth],
    ['tabMonthSubscriber', (typeof subscriberData!=='undefined' && subscriberData && subscriberData.baseMonth) ? subscriberData.baseMonth : (m.subscriber || fallbackMonth)],
    ['tabMonthCommission', m.commission || fallbackMonth],
    ['tabMonthSim', m.simulate || m.plan || m.profit || fallbackMonth]
  ];
  mapping.forEach(([id,val])=>{
    const el=document.getElementById(id);
    if(el) el.textContent='기준월 '+(val||'-');
  });
}
function monthToNumber(month){
  if(!month || !/^\d{4}\.\d{2}$/.test(month)) return null;
  const parts = month.split('.');
  return Number(parts[0])*100 + Number(parts[1]);
}
function getSavedBaseMonth(){
  try{ return localStorage.getItem(BASE_MONTH_STORAGE_KEY); }catch(e){ return null; }
}
function saveBaseMonth(month){
  if(!month) return;
  try{
    // 최근 업로드 기준월을 그대로 저장한다.
    // (이전 값이 더 크다는 이유로 저장을 막으면 탭 기준월이 과거/현재와 어긋날 수 있음)
    localStorage.setItem(BASE_MONTH_STORAGE_KEY, month);
  }catch(e){}
}
function applyBaseMonthToUi(month, sourceLabel){
  if(!month) return;
  D.baseMonth = month;
  D.period = month + ' 기준';
  const badge = document.querySelector('#hdrMonthBadge, .hdr-badge');
  if(badge) badge.textContent = month;
  const dashMonth = document.getElementById('dashMonth2');
  if(dashMonth) dashMonth.textContent = month;
  const upBadge = document.getElementById('upBasemonthBadge');
  if(upBadge) upBadge.textContent = '기준월: ' + month + (sourceLabel ? ' · ' + sourceLabel : '');
  updateTabMonthBadges();
}

// ── 파일명에서 연/월 감지 ──
function extractMonthFromFilename(filename){
  if(!filename) return null;
  const raw = String(filename);
  const fn = raw.replace(/[_\s\-\.]/g,' ');

  const toYm = (yy, mm)=>{
    const m = Number(mm);
    if(!(m>=1 && m<=12)) return null;
    const y = Number(yy);
    const yyyy = y < 100 ? (y < 50 ? 2000+y : 1900+y) : y;
    return `${yyyy}.${String(m).padStart(2,'0')}`;
  };

  // 패턴1: "26년 1월", "2026년 01월", "26년 1Q(1월)" 등
  const m1 = fn.match(/((?:19|20)?\d{2})년[^\d]{0,10}(?:[1-4][Qq][^\d]{0,10})?(\d{1,2})월/);
  if(m1){
    const ym = toYm(m1[1], m1[2]);
    if(ym) return ym;
  }

  // 패턴2: "202601", "2026_01" 형태
  const m2 = fn.match(/((?:19|20)\d{2})[^\d]?(0?[1-9]|1[0-2])(?!\d)/);
  if(m2){
    const ym = toYm(m2[1], m2[2]);
    if(ym) return ym;
  }

  // 패턴3: "12월 가마감"처럼 월만 있는 파일명은 기준 연도(D.baseMonth 또는 D.period)로 보완
  const m3 = fn.match(/(^|[^\d])(1[0-2]|0?[1-9])월/);
  if(m3){
    const baseYear = Number((D.baseMonth||'').slice(0,4)) || Number((D.period||'').match(/(19|20)\d{2}/)?.[0]) || new Date().getFullYear();
    const ym = toYm(baseYear, m3[2]);
    if(ym) return ym;
  }

  // 주의: YYMMDD(예: 260205)는 파일 생성일/배포일인 경우가 많아 기준월로 오인될 수 있음.
  // 따라서 날짜 6자리는 기준월 추출 대상으로 사용하지 않는다.

  return null;
}

// ── 워크북 내부에서 연/월 추정 (파일명 실패 시 보조) ──
function extractMonthFromWorkbook(wb){
  if(!wb || !Array.isArray(wb.SheetNames) || !wb.SheetNames.length) return null;
  const toYm = (yy, mm)=>{
    const m = Number(mm);
    if(!(m>=1 && m<=12)) return null;
    const y = Number(yy);
    const yyyy = y < 100 ? (y < 50 ? 2000+y : 1900+y) : y;
    if(yyyy < 2000 || yyyy > 2099) return null;
    return `${yyyy}.${String(m).padStart(2,'0')}`;
  };
  const candidates = [];
  const collect = (text)=>{
    const s = String(text||'');
    if(!s) return;
    let m;
    const re1 = /((?:19|20)?\d{2})년[^\d]{0,10}(\d{1,2})월/g;
    while((m=re1.exec(s))!==null){
      const ym = toYm(m[1], m[2]);
      if(ym) candidates.push(ym);
    }
    const re2 = /((?:19|20)\d{2})[^\d]?(0?[1-9]|1[0-2])(?!\d)/g;
    while((m=re2.exec(s))!==null){
      const ym = toYm(m[1], m[2]);
      if(ym) candidates.push(ym);
    }
    const re3 = /(^|[^\d])([2-3]\d)(0[1-9]|1[0-2])(?!\d)/g; // 2401, 2601
    while((m=re3.exec(s))!==null){
      const ym = toYm(m[2], m[3]);
      if(ym) candidates.push(ym);
    }
  };

  wb.SheetNames.forEach(collect);
  const maxSheets = Math.min(3, wb.SheetNames.length);
  for(let i=0;i<maxSheets;i++){
    const ws = wb.Sheets[wb.SheetNames[i]];
    if(!ws || !ws['!ref']) continue;
    const rows = XLSX.utils.sheet_to_json(ws,{header:1,defval:'',range:0});
    for(let r=0;r<Math.min(40, rows.length);r++){
      const row = rows[r] || [];
      for(let c=0;c<Math.min(20, row.length);c++) collect(row[c]);
    }
  }

  if(!candidates.length) return null;
  return candidates.sort((a,b)=>monthToNumber(b)-monthToNumber(a))[0] || null;
}

// ── 헤더 배지 + D.baseMonth 업데이트 ──
function setGlobalBaseMonth(month, sourceLabel){
  if(!month) return;
  applyBaseMonthToUi(month, sourceLabel);
  saveBaseMonth(month);
  syncAllTabMonths(month);
  console.log('기준월 변경:', month, sourceLabel||'');
}
function setReportPeriod(month){
  if(!month) return;
  D.period = month + ' 기준';
}

function parseExcel(file,type){
  const filename = file.name;
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const wb=XLSX.read(e.target.result,{type:'array'});
      if(type==='task') parseTaskExcel(wb, filename);
      else if(type==='profit') parseProfitExcel(wb, filename);
      else if(type==='profitYtd') parseProfitYtdExcel(wb, filename);
      else if(type==='factbook') parseFactbookExcel(wb);
      else if(type==='comm') parseCommExcel(wb, filename);
      else if(type==='hqprofit') parseHqProfitExcel(wb, filename);
      else if(type==='plan') parsePlanExcel(wb, filename);
      else if(type==='midterm') parseMidtermExcel(wb, filename);
      else parseKpiExcel(wb, filename);
    }catch(err){showUpStatus(type,'err','❌ '+err.message);}
  };
  reader.readAsArrayBuffer(file);
}

async function parseBundledExcelFile(path, type){
  const res = await fetch(encodeURI(path), {cache:'no-store'});
  if(!res.ok) throw new Error(`${path} 파일을 찾을 수 없습니다. (${res.status})`);
  const buf = await res.arrayBuffer();
  const wb = XLSX.read(buf, {type:'array'});
  if(type==='plan') parsePlanExcel(wb, path);
  else if(type==='midterm') parseMidtermExcel(wb, path);
}

async function loadBundledManagementPlans(){
  ensureSimulationDataShape();
  const hasPlan = Array.isArray(D.planData?.monthly) && D.planData.monthly.length>0;
  const hasMidterm = Array.isArray(D.midtermAssumptions?.assumptions) && D.midtermAssumptions.assumptions.length>0;
  if(hasPlan && hasMidterm) return;

  if(!hasPlan){
    try{
      await parseBundledExcelFile(BUNDLED_PLAN_FILE, 'plan');
    }catch(err){
      showUpStatus('plan','warn','⚠️ 내장 경영계획 자동 로드 실패: '+err.message+'\n필요 시 업로드 탭에서 직접 업로드해 주세요.');
    }
  }

  if(!hasMidterm){
    try{
      await parseBundledExcelFile(BUNDLED_MIDTERM_FILE, 'midterm');
    }catch(err){
      showUpStatus('midterm','warn','⚠️ 내장 중기/그룹계획 자동 로드 실패: '+err.message+'\n필요 시 업로드 탭에서 직접 업로드해 주세요.');
    }
  }
}

function parsePlanExcel(wb, filename){
  try{
    ensureSimulationDataShape();
    const month = extractMonthFromFilename(filename);
    const monthLabels = ['1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월'];
    const norm=v=>String(v||'').replace(/\s+/g,'').trim();
    const toNum=v=>{
      if(v===null||v===undefined||v==='') return null;
      const n = Number(String(v).replace(/,/g,''));
      return Number.isFinite(n)?n:null;
    };
    const pickSheet=name=>wb.Sheets[wb.SheetNames.find(s=>String(s).includes(name))||name];

    const monthly = monthLabels.map((m,i)=>({month:m,monthIndex:i+1,revenue:null,op:null,netAdd:null,capa:null,mgmtFee:null,policyFee:null,arpu:null}));

    const isWs = pickSheet('전사 손익계산서');
    if(isWs){
      const rows = XLSX.utils.sheet_to_json(isWs,{header:1,defval:''});
      const hdr = rows.find(r=>r.some(c=>monthLabels.includes(String(c).trim()))) || [];
      const monthCols={};
      hdr.forEach((c,i)=>{ const m=String(c||'').trim(); if(monthLabels.includes(m)) monthCols[m]=i; });
      const findRow=(name)=>rows.find(r=>norm(r[0]).includes(norm(name)));
      const revRow = findRow('매출');
      const opRow = findRow('영업이익');
      monthLabels.forEach((m,i)=>{
        const c = monthCols[m];
        if(c!==undefined){
          if(revRow) monthly[i].revenue = toNum(revRow[c]);
          if(opRow) monthly[i].op = toNum(opRow[c]);
        }
      });
    }

    const capaWs = pickSheet('무선 CAPA') || pickSheet('무선');
    if(capaWs){
      const rows = XLSX.utils.sheet_to_json(capaWs,{header:1,defval:''});
      const hdr = rows.find(r=>r.some(c=>monthLabels.includes(String(c).trim()))) || [];
      const monthCols={};
      hdr.forEach((c,i)=>{ const m=String(c||'').trim(); if(monthLabels.includes(m)) monthCols[m]=i; });
      const findRow=(kw)=>rows.find(r=>norm(r[0]).includes(norm(kw)));
      const capaRow=findRow('CAPA'), netRow=findRow('순증'), mgRow=findRow('관리수수료'), poRow=findRow('정책수수료'), arRow=findRow('ARPU');
      monthLabels.forEach((m,i)=>{
        const c=monthCols[m];
        if(c===undefined) return;
        if(capaRow) monthly[i].capa = toNum(capaRow[c]);
        if(netRow) monthly[i].netAdd = toNum(netRow[c]);
        if(mgRow) monthly[i].mgmtFee = toNum(mgRow[c]);
        if(poRow) monthly[i].policyFee = toNum(poRow[c]);
        if(arRow) monthly[i].arpu = toNum(arRow[c]);
      });
    }

    const fallback = buildSimulationBaseline().actual;
    monthly.forEach(m=>{
      if(m.revenue===null) m.revenue=fallback.revenue;
      if(m.op===null) m.op=fallback.op;
      if(m.netAdd===null) m.netAdd=fallback.netAdd;
      if(m.capa===null) m.capa=fallback.capa;
      if(m.mgmtFee===null) m.mgmtFee=fallback.mgmtFee;
      if(m.policyFee===null) m.policyFee=fallback.policyFee;
      if(m.arpu===null) m.arpu=fallback.arpu;
    });

    const bm = CHS.map(ch=>({
      key:ch,
      name:CN[ch],
      revenue:Number(D?.profit?.revenue?.[ch])||0,
      op:Number(D?.profit?.op?.[ch])||0,
      margin:(Number(D?.profit?.revenue?.[ch])||0)>0?((Number(D?.profit?.op?.[ch])||0)/(Number(D?.profit?.revenue?.[ch])||1)):0
    }));

    D.planData = {
      sourceFile: filename,
      monthly,
      bm,
      kpiTargets:['소매 KPI','도매 KPI','소상공 KPI','운영 KPI'].map((k,i)=>({name:k,target:82+i*1.5}))
    };

    if(month){ setGlobalBaseMonth(month, '경영계획'); setReportPeriod(month); }
    setTabMonth('plan', month || (D.period||'').replace(' 기준',''));
    setTabMonth('simulate', month || (D.period||'').replace(' 기준',''));
    saveUploadedState();
    updateTabMonthBadges();
    initSimulationTab();
    showUpStatus('plan','ok',`✅ 2026 경영계획 로드 완료 · 월별 ${D.planData.monthly.length}건`);
  }catch(err){
    showUpStatus('plan','err','❌ 경영계획 파싱 오류: '+err.message);
  }
}

function parseMidtermExcel(wb, filename){
  try{
    ensureSimulationDataShape();
    const norm=v=>String(v||'').replace(/\s+/g,'').trim();
    const toNum=v=>{
      if(v===null||v===undefined||v==='') return null;
      const n = Number(String(v).replace(/,/g,''));
      return Number.isFinite(n)?n:null;
    };
    const assumptions=[];

    const keyWs = wb.Sheets[wb.SheetNames.find(s=>String(s).toLowerCase().includes('key index'))||'key Index'];
    if(keyWs){
      const rows = XLSX.utils.sheet_to_json(keyWs,{header:1,defval:''});
      rows.forEach(r=>{
        const nm = String(r[0]||'').trim();
        if(!nm) return;
        const v = toNum(r[1]);
        if(v===null) return;
        assumptions.push({driver:nm,baseValue:v,unit:String(r[2]||''),sourceSheet:'key Index'});
      });
    }

    const bmWs = wb.Sheets[wb.SheetNames.find(s=>String(s).includes('BM별손익'))||'BM별손익'];
    const bm=[];
    if(bmWs){
      const rows=XLSX.utils.sheet_to_json(bmWs,{header:1,defval:''});
      rows.forEach(r=>{
        const nm=String(r[2]||r[0]||'').trim();
        if(!nm||/BM|구분|합계/.test(nm)) return;
        const rev=toNum(r[3])||toNum(r[4])||0;
        const op=toNum(r[7])||toNum(r[8])||0;
        if(rev===0&&op===0) return;
        bm.push({name:nm,revenue:rev,op,margin:rev?op/rev:0,sourceSheet:'BM별손익'});
      });
    }

    D.midtermAssumptions = {
      sourceFile: filename,
      assumptions: assumptions.slice(0,60),
      bm
    };

    saveUploadedState();
    initSimulationTab();
    showUpStatus('midterm','ok',`✅ 중기/그룹 가정 로드 완료 · Key ${D.midtermAssumptions.assumptions.length}건`);
  }catch(err){
    showUpStatus('midterm','err','❌ 중기계획 파싱 오류: '+err.message);
  }
}

function extractCommissionDataFromWorkbook(wb){
  const ws = wb.Sheets[wb.SheetNames[0]];
  if(!ws) throw new Error('수수료 시트를 찾을 수 없습니다.');
  const rows = XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
  const parsed = cloneObj(D.commission||{});
  if(!parsed.jan26 || typeof parsed.jan26!=='object') parsed.jan26 = {};

  const norm=v=>String(v||'').replace(/\s+/g,'').trim();
  const toNum=v=>{
    if(v===null || v===undefined || v==='') return null;
    const n = Number(String(v).replace(/,/g,''));
    return Number.isFinite(n) ? n : null;
  };
  const findRow=(keyword)=>rows.find(r=>r.some(c=>norm(c).includes(keyword)));

  const setIfNumber=(obj,key,val)=>{ if(val!==null) obj[key]=val; };

  const mgmtRow = findRow('관리수수료') || findRow('관리수수료합계');
  const policyRow = findRow('정책수수료') || findRow('수수료수입정책');
  const totalRow = findRow('전사합계') || findRow('합계');

  if(mgmtRow){
    if(!parsed.jan26.mgmt_fee) parsed.jan26.mgmt_fee={};
    const mg = parsed.jan26.mgmt_fee;
    setIfNumber(mg,'retail',toNum(mgmtRow[2]));
    setIfNumber(mg,'wholesale',toNum(mgmtRow[3]));
    setIfNumber(mg,'digital',toNum(mgmtRow[4]));
    setIfNumber(mg,'b2b',toNum(mgmtRow[5]));
    setIfNumber(mg,'small_biz',toNum(mgmtRow[6]));
    setIfNumber(mg,'common',toNum(mgmtRow[7]));
    const sum = ['retail','wholesale','digital','b2b','small_biz','common'].reduce((a,k)=>a+(Number(mg[k])||0),0);
    if(sum>0) mg.total = Number(sum.toFixed(1));
  }

  if(policyRow){
    if(!parsed.jan26.policy_fee) parsed.jan26.policy_fee={};
    const pf = parsed.jan26.policy_fee;
    setIfNumber(pf,'retail',toNum(policyRow[2]));
    setIfNumber(pf,'wholesale',toNum(policyRow[3]));
    setIfNumber(pf,'digital',toNum(policyRow[4]));
    setIfNumber(pf,'enterprise',toNum(policyRow[5]));
    setIfNumber(pf,'corporate_sales',toNum(policyRow[6]));
    setIfNumber(pf,'rds',toNum(policyRow[7]));
    setIfNumber(pf,'policy',toNum(policyRow[8]));
    const sum = ['retail','wholesale','digital','enterprise','corporate_sales','rds','policy'].reduce((a,k)=>a+(Number(pf[k])||0),0);
    if(sum>0) pf.total = Number(sum.toFixed(1));
  }

  if(totalRow){
    if(!parsed.jan26.total) parsed.jan26.total={};
    const tt = parsed.jan26.total;
    setIfNumber(tt,'retail',toNum(totalRow[2]));
    setIfNumber(tt,'wholesale',toNum(totalRow[3]));
    setIfNumber(tt,'digital',toNum(totalRow[4]));
    setIfNumber(tt,'b2b',toNum(totalRow[5]));
    setIfNumber(tt,'small_biz',toNum(totalRow[6]));
    setIfNumber(tt,'common_corp',toNum(totalRow[7]));
    setIfNumber(tt,'common_ch',toNum(totalRow[8]));
    const sum = ['retail','wholesale','digital','b2b','small_biz','common_corp','common_ch'].reduce((a,k)=>a+(Number(tt[k])||0),0);
    if(sum>0) tt.total = Number(sum.toFixed(1));
  }

  if(parsed.jan26 && parsed.jan26.total && parsed.jan26.total.total){
    const y26 = (parsed.yearly||[]).find(y=>y.yr==='26년(E)');
    if(y26){
      y26.total = Number((parsed.jan26.total.total*12).toFixed(1));
      ['retail','wholesale','digital','b2b','small_biz'].forEach(k=>{
        y26[k] = Number(((parsed.jan26.total[k]||0)*12).toFixed(1));
      });
    }
  }

  return parsed;
}

// ── 수수료 파일: 파일명 기준월만 감지 ──
function parseCommExcel(wb, filename){
  const month = extractMonthFromFilename(filename) || extractMonthFromWorkbook(wb) || D.baseMonth || getSavedBaseMonth() || null;
  if(month){
    setGlobalBaseMonth(month, '수수료');
    setReportPeriod(month);
  }

  // 업로드한 원본을 로컬에 보존해 새로고침 후에도 최신 수치가 유지되도록 동기화
  try{
    const parsed = extractCommissionDataFromWorkbook(wb);
    if(parsed && typeof parsed==='object') D.commission = parsed;
  }catch(err){
    console.warn('commission parse fallback:', err && err.message ? err.message : err);
  }

  initCommission();
  if(month) setTabMonth('commission', month);
  else clearTabMonth('commission');
  saveUploadedState();
  updateTabMonthBadges();
  showUpStatus('comm', month?'ok':'warn', month ? `✅ 수수료 데이터 업데이트 · 기준월 ${month}` : '✅ 수수료 데이터 업데이트 (기준월 감지 실패)');
}
function parseTaskExcel(wb, filename){
  const month = extractMonthFromFilename(filename);
  if(month){ D.taskMonth = month; setTabMonth('tasks', month); }
  const ws=wb.Sheets['data']||(function(){var m=wb.SheetNames.filter(function(s){return/마감/.test(s)&&!/핵심/.test(s);}).sort(function(a,b){var sa=wb.Sheets[a],sb=wb.Sheets[b];var ra=sa['!ref']?parseInt(sa['!ref'].replace(/.*:/,'').replace(/[A-Z]+/g,'')):0;var rb=sb['!ref']?parseInt(sb['!ref'].replace(/.*:/,'').replace(/[A-Z]+/g,'')):0;return rb-ra;});return m.length?wb.Sheets[m[0]]:null;})()||wb.Sheets[wb.SheetNames[0]];if(!ws){showUpStatus('task','err','❌ 시트를 찾을 수 없습니다');return;}const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});const tasks=[];for(let i=5;i<rows.length;i++){const r=rows[i];if(!r[5]||!r[1])continue;tasks.push({no:r[1],ch:r[2]||'',tm:r[3]||'',co:r[4]==='●'?'●':'',nm:String(r[5]).replace(/\n/g,' ').trim(),mg:String(r[6]||'').trim(),sc:String(r[7]||'').trim(),plan:String(r[8]||'').trim(),pg:String(r[9]||'').trim(),st:r[10]||'계획',cp:r[11]||'미완료',pp:parseFloat(r[12])||0});}
  if(tasks.length>0){D.tasks=tasks;saveUploadedState();initTaskUI();initDashboard();updateTabMonthBadges();showUpStatus('task','ok','✅ '+tasks.length+'개 과제 로드!'+(month?` · 과제 기준월 ${month} (전사 보고 기준월은 유지)`:''));}
  else showUpStatus('task','err','❌ 데이터 없음');
}


function parseProfitExcel(wb, filename){try{
  const month = extractMonthFromFilename(filename);
  if(month){ setGlobalBaseMonth(month, 'BM별손익'); setReportPeriod(month); }
  var ws=wb.Sheets['종합']||wb.Sheets[wb.SheetNames.find(function(s){return String(s).indexOf('종합')>=0;})];
  if(!ws){showUpStatus('profit','err','❌ 종합 시트 없음');return;}
  var rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:0});
  var norm=function(x){return String(x||'').trim().replace(/\s+/g,'');};
  var CH_NAME={'소매':'retail','도매':'wholesale','디지털':'digital','기업/공공':'enterprise','IoT':'iot','법인영업':'corporate_sales','소상공인':'small_biz'};
  var colIdx={retail:5,wholesale:6,digital:7,enterprise:8,iot:9,corporate_sales:10,small_biz:11};
  var warnings=[];
  if(rows[1]){
    var det={};
    rows[1].forEach(function(v,i){var n=String(v||'').trim();if(CH_NAME[n]&&!det[CH_NAME[n]])det[CH_NAME[n]]=i;});
    if(Object.keys(det).length>=5){
      for(var ch in det){if(colIdx[ch]!==undefined&&colIdx[ch]!==det[ch])warnings.push('📌 '+ch+' 컬럼위치: 기존 '+(colIdx[ch]+1)+'열 → '+(det[ch]+1)+'열');}
      for(var ch2 in det){colIdx[ch2]=det[ch2];}
    } else {
      warnings.push('⚠️ 채널 헤더 자동감지 실패('+Object.keys(det).length+'개만 감지) — 기본 컬럼 사용');
    }
  }
  var find=function(l){return rows.find(function(r){return r[0]&&norm(r[0])===norm(l);});};
  var findInc=function(l){return rows.find(function(r){return r[0]&&norm(r[0]).indexOf(norm(l))>=0;});};
  var mk=function(r){return {total:r[1]||0,retail:r[colIdx.retail]||0,wholesale:r[colIdx.wholesale]||0,digital:r[colIdx.digital]||0,enterprise:r[colIdx.enterprise]||0,iot:r[colIdx.iot]||0,corporate_sales:r[colIdx.corporate_sales]||0,small_biz:r[colIdx.small_biz]||0};};
  var keyCheck={'매출':null,'매출총이익':null,'판매비와일반관리비':null};
  var keyInc={'상품매출':'상품매출','수수료수입':'수수료수입','점프업':'점프업도/소매조정후영업이익'};
  for(var k in keyCheck){var r=find(k);keyCheck[k]=r;if(!r)warnings.push('❌ "'+k+'" 행 없음');}
  for(var k2 in keyInc){if(!findInc(k2))warnings.push('⚠️ "'+keyInc[k2]+'" 행 없음');}
  var rev=keyCheck['매출'],gp=keyCheck['매출총이익'],sga=keyCheck['판매비와일반관리비'];
  var revProd=findInc('상품매출'),revSvc=findInc('수수료수입');
  if(rev)D.profit.revenue=mk(rev);
  if(gp)D.profit.gross=mk(gp);
  if(sga)D.profit.sga=mk(sga);
  if(revProd)D.profit.rev_product=mk(revProd);
  if(revSvc)D.profit.rev_service=mk(revSvc);
  var contrib=findInc('공헌이익');
  if(contrib){
    D.profit.contribution=mk(contrib);
    // 점프업 조정: 소매→도매 이동 (점프업이 소매로 잡혀 들어오므로 수작업 조정)
    var jumpupAdj=25805899;
    D.profit.contribution.retail=(D.profit.contribution.retail||0)-jumpupAdj;
    D.profit.contribution.wholesale=(D.profit.contribution.wholesale||0)+jumpupAdj;
  }
  var opReal=findInc('점프업');
  if(opReal){
    D.profit.op=mk(opReal);D.profit.op_source='점프업 도/소매 조정 후';
  } else {
    CHS.forEach(function(c){D.profit.op[c]=(D.profit.gross[c]||0)-(D.profit.sga[c]||0);});
    D.profit.op.total=(D.profit.gross.total||0)-(D.profit.sga.total||0);
    D.profit.op_source='공헌이익(gross-sga)';
  }
  var sgaNames={'판매촉진비':'판매촉진비','복리후생비':'복리후생비','임차료':'지급임차료','판매수수료':'판매수수료','지급수수료':'지급수수료','운반비':'운반비'};
  var sgaMissing=[];
  for(var sn in sgaNames){
    var sr=find(sgaNames[sn]);
    if(sr){D.sga_detail[sn]=mk(sr);}
    else{sgaMissing.push(sn);}
  }
  var laborParts=['임원급여','직원급여','상여금','퇴직급여충당금전입'];
  var laborSum={total:0,retail:0,wholesale:0,digital:0,enterprise:0,iot:0,corporate_sales:0,small_biz:0};
  var laborMiss=[];
  laborParts.forEach(function(lp){
    var lr=find(lp);
    if(lr){for(var lk in laborSum){laborSum[lk]+=(lk==='total'?lr[1]:lr[colIdx[lk]])||0;}}
    else{laborMiss.push(lp);}
  });
  if(laborMiss.length===0){D.sga_detail['인건비합계']=laborSum;}
  else{sgaMissing.push('인건비(누락:'+laborMiss.join(',')+')');warnings.push('⚠️ 인건비 구성항목 누락: '+laborMiss.join(', '));}
  if(sgaMissing.length>0)warnings.push('⚠️ SGA 세부 미파싱: '+sgaMissing.join(', '));
  var hqMap={'강북본부':'강북영업본부','강남본부':'강남영업본부','강서본부':'강서영업본부','동부본부':'동부영업본부','서부본부':'서부영업본부'};
  var hqResult=[];var hqMissing=[];
  for(var hqNm in hqMap){
    var iuSheet=wb.SheetNames.find(function(s){return s.indexOf(hqMap[hqNm])>=0&&s.indexOf('IU')>=0;});
    if(!iuSheet){hqMissing.push(hqNm);continue;}
    var hws=wb.Sheets[iuSheet];
    var hrows=XLSX.utils.sheet_to_json(hws,{header:1,defval:0});
    var hfind=function(l){return hrows.find(function(r2){return r2[0]&&norm(r2[0])===norm(l);});};
    var hr=hfind('매출'),hg=hfind('매출총이익'),hs=hfind('판매비와일반관리비');
    hqResult.push({nm:hqNm,rev:hr?Math.round(hr[1]||0):0,gp:hg?Math.round(hg[1]||0):0,sga:hs?Math.round(hs[1]||0):0});
  }
  if(hqResult.length>0)D.hq=hqResult;
  if(hqMissing.length>0)warnings.push('⚠️ 본부 IU시트 없음: '+hqMissing.join(', '));
  var niSheet=wb.Sheets['전사손익계산서'];
  if(niSheet){
    var niRows=XLSX.utils.sheet_to_json(niSheet,{header:1,defval:0});
    var niRow=niRows.find(function(r3){return r3[0]&&norm(r3[0]).indexOf('당기순이익')>=0;});
    if(!niRow)niRow=niRows.find(function(r4){return r4[0]&&norm(r4[0])==='손익';});
    if(niRow)D.profit.net_income=Math.round(niRow[1]||0);
  }
  if(month) setTabMonth('profit', month);
  else clearTabMonth('profit');
  saveUploadedState();
  initProfitUI();renderTrendChart();initDashboard();updateTabMonthBadges();
  if(warnings.length>0){
    showUpStatus('profit','warn','✅ 손익 업데이트 완료 ('+D.profit.op_source+')'+(month?` · 기준월 ${month}`:'')+'\n\n🔍 <b>구조 검증 결과 ('+warnings.length+'건)</b>\n'+warnings.join('\n'));
  } else {
    showUpStatus('profit','ok','✅ 손익 전체 업데이트!'+(month?` · 기준월 ${month}`:'')+' (매출/총이익/판관비/SGA세부/본부별/영업이익)\n영업이익: '+D.profit.op_source);
  }
}catch(err){showUpStatus('profit','err','❌ 파싱 오류: '+err.message);}}

function parseProfitYtdExcel(wb, filename){try{
  const month = extractMonthFromFilename(filename);
  if(month){ setGlobalBaseMonth(month, 'BM누적'); setReportPeriod(month); setTabMonth('profit', month); }
  else clearTabMonth('profit');
  var MONTHS=['1\uc6d4','2\uc6d4','3\uc6d4','4\uc6d4','5\uc6d4','6\uc6d4','7\uc6d4','8\uc6d4','9\uc6d4','10\uc6d4','11\uc6d4','12\uc6d4'];
  var CH_HDR={'\uc18c\ub9e4':'retail','\ub3c4\ub9e4':'wholesale','\ub514\uc9c0\ud138':'digital','\ub514\uc9c0\ud138(KT\uc0f5)':'digital','\uae30\uc5c5/\uacf5\uacf5':'enterprise','IoT':'iot','\ubc95\uc778\uc601\uc5c5':'corporate_sales','\uc18c\uc0c1\uacf5\uc778':'small_biz'};
  var norm=function(x){return String(x||'').trim().replace(/\s+/g,'').replace(/\n/g,'');};
  var trend=[];var loaded=0;
  for(var mi=0;mi<MONTHS.length;mi++){
    var m=MONTHS[mi];
    var ws=wb.Sheets[m];if(!ws)continue;
    var rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:0});
    var hdrRow=null;
    for(var hi=0;hi<5;hi++){if(rows[hi]&&rows[hi].some(function(v){return String(v||'').indexOf('\uc18c\ub9e4')>=0;})){hdrRow=rows[hi];break;}}
    var colMap={};
    if(hdrRow)for(var ci=0;ci<hdrRow.length;ci++){var k=CH_HDR[String(hdrRow[ci]||'').trim()];if(k&&!colMap[k])colMap[k]=ci;}
    var opRow=null;
    for(var ri=260;ri<Math.min(rows.length,285);ri++){if(rows[ri]&&rows[ri][0]&&norm(rows[ri][0]).indexOf('\uc810\ud504\uc5c5')>=0&&norm(rows[ri][0]).indexOf('\uc870\uc815')>=0){opRow=rows[ri];break;}}
    if(opRow){
      var entry={m:m,monthIndex:mi+1};entry.total=Math.round((opRow[1]||0)/1e6)/100;
      for(var ck in colMap)entry[ck]=Math.round((opRow[colMap[ck]]||0)/1e6)/100;
      trend.push(entry);loaded++;
    }
  }
  var wsZ=wb.Sheets['\uc885\ud569']||wb.Sheets[wb.SheetNames.filter(function(s){return String(s).indexOf('\uc885\ud569')>=0;})[0]];
  if(wsZ){
    var rows2=XLSX.utils.sheet_to_json(wsZ,{header:1,defval:0});
    var find2=function(l){return rows2.find(function(r){return r[0]&&norm(r[0])===norm(l);});};
    var findInc2=function(l){return rows2.find(function(r){return r[0]&&norm(r[0]).indexOf(norm(l))>=0;});};
    var mk2=function(r){return {total:r[1]||0,retail:r[5]||0,wholesale:r[6]||0,digital:r[7]||0,enterprise:r[8]||0,iot:r[9]||0,corporate_sales:r[10]||0,small_biz:r[11]||0};};
    var rv=find2('\ub9e4\ucd9c'),gp=find2('\ub9e4\ucd9c\uc758\ucd1d\uc774\uc775'),sg=find2('\ud310\ub9e4\ube44\uc640\uc77c\ubc18\uad00\ub9ac\ube44'),op2=findInc2('\uc810\ud504\uc5c5');
    if(rv)D.profit.revenue=mk2(rv);if(gp)D.profit.gross=mk2(gp);if(sg)D.profit.sga=mk2(sg);
    if(op2){D.profit.op=mk2(op2);D.profit.op_source='\uc810\ud504\uc5c5 \ub3c4/\uc18c\ub9e4 \uc870\uc815 \ud6c4(\uc5f0\uac04)';}
    saveUploadedState();
    initProfitUI();initDashboard();updateTabMonthBadges();
  }
  if(loaded>0){D.profit.monthly_trend=trend;renderTrendChart();showUpStatus('profitYtd','ok','✅ '+loaded+'개월 트렌드 로드!'+(month?` · 기준월 ${month}`:'')+' (월별 영업이익 차트 업데이트)');}
  else showUpStatus('profitYtd','err','❌ 월별 시트 데이터 없음');
}catch(err){showUpStatus('profitYtd','err','\u274c '+err.message);}}
function renderTrendChart(){
  var trend=D.profit.monthly_trend||[];
  if(!trend.length||!document.getElementById('trendChart'))return;
  var sel=document.getElementById('trendChSel');
  var ch=sel?sel.value:'total';
  var vals=trend.map(function(t){return t[ch]||0;});
  var absMax=Math.max.apply(null,vals.map(Math.abs).concat([1]));
  var posC='#16a34a',negC='#dc2626';
  var rows=[];
  for(var i=0;i<trend.length;i++){
    var t=trend[i];var v=t[ch]||0;var pct=Math.abs(v)/absMax*45;
    var col=v>=0?posC:negC;
    var nf=v===0?'0':(v>0?'+'+v.toFixed(1):v.toFixed(1));
    var bar=v>=0
      ?'<div style="width:50%"></div><div style="width:2px;height:18px;background:#cbd5e1;flex-shrink:0"></div><div style="width:'+pct+'%;height:14px;background:'+col+';border-radius:0 3px 3px 0;min-width:2px"></div>'
      :'<div style="width:'+(50-pct)+'%"></div><div style="width:'+pct+'%;height:14px;background:'+col+';border-radius:3px 0 0 3px"></div><div style="width:2px;height:18px;background:#cbd5e1;flex-shrink:0"></div>';
    rows.push('<div class="br" style="margin-bottom:4px"><div class="bl_" style="width:28px;font-size:10px">'+t.m+'</div><div style="flex:1;display:flex;align-items:center;height:18px">'+bar+'</div><div style="width:48px;font-size:10px;font-weight:700;color:'+col+';text-align:right">'+nf+'</div></div>');
  }
  document.getElementById('trendChart').innerHTML=rows.join('');
  var posM=0,negM=0,tv=0;
  for(var j=0;j<vals.length;j++){if(vals[j]>0)posM++;else if(vals[j]<0)negM++;tv+=vals[j];}
  var bi=0,wi=0;
  for(var k=1;k<vals.length;k++){if(vals[k]>vals[bi])bi=k;if(vals[k]<vals[wi])wi=k;}
  var bm=trend[bi]?trend[bi].m:'';var wm=trend[wi]?trend[wi].m:'';
  var bv=trend[bi]?(trend[bi][ch]||0):0;var wv=trend[wi]?(trend[wi][ch]||0):0;
  var sEl=document.getElementById('trendSummary');
  if(sEl)sEl.innerHTML='<div style="font-size:11px;padding:4px 10px;border-radius:8px;background:#dcfce7;color:#166534">\uc218\uc775 '+posM+'\uac1c\uc6d4</div><div style="font-size:11px;padding:4px 10px;border-radius:8px;background:#fee2e2;color:#991b1b">\uc801\uc790 '+negM+'\uac1c\uc6d4</div><div style="font-size:11px;padding:4px 10px;border-radius:8px;background:#f1f5f9;color:#334155">\uc5f0\uac04\ud569\uacc4 '+tv.toFixed(1)+'\uc5b5</div><div style="font-size:11px;padding:4px 10px;border-radius:8px;background:#dbeafe;color:#1e40af">\ucd5c\uace0 '+bm+' '+bv.toFixed(1)+'\uc5b5</div><div style="font-size:11px;padding:4px 10px;border-radius:8px;background:#fef3c7;color:#92400e">\ucd5c\uc800 '+wm+' '+wv.toFixed(1)+'\uc5b5</div>';
  var CHS2=['retail','wholesale','digital','enterprise','iot','corporate_sales','small_biz'];
  var tb=document.getElementById('trendTable');if(!tb)return;
  var trs=[];
  for(var ti=0;ti<trend.length;ti++){
    var td=trend[ti];var tc=td.total||0;
    var cl=tc>=0?'color:#16a34a;font-weight:700':'color:#dc2626;font-weight:700';
    var cells='<td style="font-weight:600">'+td.m+'</td><td style="'+cl+'">'+(tc>0?'+':'')+tc.toFixed(1)+'</td>';
    for(var x=0;x<CHS2.length;x++){var cv=td[CHS2[x]]||0;cells+='<td style="'+(cv<0?'color:#dc2626':cv>0?'color:#16a34a':'')+'">'+(cv>0?'+':'')+cv.toFixed(1)+'</td>';}
    trs.push('<tr>'+cells+'</tr>');
  }
  tb.innerHTML=trs.join('');
}
function parseKpiExcel(wb, filename){try{
  const month = extractMonthFromFilename(filename) || extractMonthFromWorkbook(wb);
  if(month){ setGlobalBaseMonth(month, 'KPI'); setReportPeriod(month); }
  const ws=wb.Sheets['1Q 종합']||wb.Sheets[wb.SheetNames.find(s=>String(s).includes('종합'))];if(!ws){showUpStatus('kpi','err','❌ 종합 시트 없음');return;}const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:0});const kpi=[];
  const hqSet = new Set(['강북본부','강남본부','강서본부','동부본부','서부본부']);
  for(let i=0;i<rows.length;i++){
    const r=rows[i];
    const hq = String(r?.[1]||'').trim();
    if(!hqSet.has(hq)) continue;
    kpi.push({hq,ts:Number(r[2])||0,rk:Number(r[3])||0,rt:{t:Number(r[4])||0,s:Number(r[5])||0,p:Number(r[6])||0},wh:{t:Number(r[7])||0,s:Number(r[8])||0,p:Number(r[9])||0},sm:{t:Number(r[10])||0,s:Number(r[11])||0,p:Number(r[12])||0}});
  }
  if(kpi.length>=3){kpi.sort((a,b)=>b.ts-a.ts);kpi.forEach((r,i)=>r.rk=i+1);D.kpi=kpi;if(month) setTabMonth('kpi', month);else clearTabMonth('kpi');saveUploadedState();initKpiUI();initDashboard();updateTabMonthBadges();showUpStatus('kpi','ok','✅ KPI 업데이트!'+(month?` · 기준월 ${month}`:''));}
  else showUpStatus('kpi','err','❌ KPI 데이터 부족');
}catch(err){showUpStatus('kpi','err','❌ '+err.message);}}
// ── 본부별 손익 파싱 (sum 시트 기준, 원화 단위) ──────────────────
function parseHqProfitExcel(wb, filename){
  try{
    const month = extractMonthFromFilename(filename);

    // sum 시트 우선 (요약 시트는 수식만 있어 XLSX.js가 값 읽기 불가)
    const wsName = wb.SheetNames.find(s=>s==='sum') || wb.SheetNames[0];
    const ws = wb.Sheets[wsName];
    if(!ws){ showUpStatus('hqprofit','err','❌ sum 시트를 찾을 수 없습니다.'); return; }
    const rows = XLSX.utils.sheet_to_json(ws,{header:1, defval:0});

    // ── 열 탐지: 헤더행(보통 행1=idx0)에서 본부명 위치 자동 감지 ──
    const hqOrder = ['강북본부','강남본부','강서본부','동부본부','서부본부'];
    let colMap = {};
    const headerRow = rows[0] || [];
    headerRow.forEach((cell,ci)=>{
      const s = String(cell||'');
      hqOrder.forEach(nm=>{ if(s.includes(nm.replace('본부',''))) colMap[nm]=ci; });
    });
    // fallback: 강북=1,강남=2,강서=3,동부=4,서부=5 (B~F열)
    if(Object.keys(colMap).length<5){
      colMap={강북본부:1,강남본부:2,강서본부:3,동부본부:4,서부본부:5};
    }

    // ── 행 탐지: 계정명 공백제거 후 키워드 정확 매칭 ──
    let revRow=-1, gpRow=-1, sgaRow=-1;
    rows.forEach((r,ri)=>{
      const nm = String(r[0]||'').replace(/\s/g,'');
      if(nm==='매출'&&revRow<0) revRow=ri;
      else if(nm==='매출총이익'&&gpRow<0) gpRow=ri;
      else if(nm==='판매비와일반관리비'&&sgaRow<0) sgaRow=ri;
    });
    // fallback (sum 시트 기준: 매출=행3→idx2, 총이익=행27→idx26, 판관비=행28→idx27)
    if(revRow<0)  revRow=2;
    if(gpRow<0)   gpRow=26;
    if(sgaRow<0)  sgaRow=27;

    // ── D.hq 갱신 ──
    const newHq = D.hq.map(h=>{
      const ci = colMap[h.nm];
      if(ci===undefined) return h;
      return {
        ...h,
        rev: Math.round(Number(rows[revRow]?.[ci])||0),
        gp:  Math.round(Number(rows[gpRow]?.[ci])||0),
        sga: Math.round(Number(rows[sgaRow]?.[ci])||0)
      };
    });
    D.hq = newHq;
    if(month){ setGlobalBaseMonth(month, '본부손익'); setReportPeriod(month); setTabMonth('profit', month); }
    else clearTabMonth('profit');
    saveUploadedState();
    initProfitUI();
    updateTabMonthBadges();

    const lines = newHq.map(h=>{
      const op = h.gp - h.sga;
      return `  ${h.nm}: 매출 ${(h.rev/1e8).toFixed(0)}억 / 영업이익 ${(op/1e8).toFixed(1)}억 (${h.rev>0?(op/h.rev*100).toFixed(1):0}%)`;
    });
    showUpStatus('hqprofit','ok',`✅ 본부별 손익 업데이트!${month?` · ${month}`:''}\n${lines.join('\n')}`);
  }catch(err){ showUpStatus('hqprofit','err','❌ 파싱 오류: '+err.message); }
}

function showUpStatus(type,cls,msg){const el=document.getElementById('s'+type.charAt(0).toUpperCase()+type.slice(1));if(el){el.className='up-status '+cls;el.innerHTML=msg.replace(/\n/g,'<br>');}uploadHealth[type]=cls;try{if(typeof initDashboard==='function')initDashboard();}catch(e){}}

function runLaunchpadReport(type){
  const actions={
    ceo:()=>genCeoBriefing(),
    exec:()=>genExecutiveManagementReport(),
    risk:()=>genRiskComprehensive(),
    task:()=>genTaskReport()
  };
  if(actions[type]) actions[type]();
}

// ==================== INIT ====================
function normalizeProfitData(){
  if(!D || !D.profit || !D.profit.contribution) return;
  const contrib=D.profit.contribution;
  const channelSum=CHS.reduce((s,c)=>s+(Number(contrib[c])||0),0);
  // total = 채널합 + 유통플랫폼 영업이익(op) — 유통플랫폼은 op 기준 합산
  const pfOp=(D.profit.platform&&D.profit.platform.total)?Number(D.profit.platform.total.op||0):0;
  if(channelSum>0) contrib.total=channelSum+pfOp;
}
normalizeProfitData();
ensureSimulationDataShape();
loadUploadedState();
ensureSimulationDataShape();
const savedBaseMonth = getSavedBaseMonth();
if(savedBaseMonth) applyBaseMonthToUi(savedBaseMonth, '최근 업로드');
loadBundledManagementPlans();
loadSubscriberData();
initTaskUI();initProfitUI();initKpiUI();initSubscriberUI();initDashboard();
updateTabMonthBadges();
document.title='KT M&S 경영관리';

// 수수료 탭 강제 바인딩 및 초기 렌더 준비
window.addEventListener('load', ()=>{
  const commBtn = document.querySelector('.bni[data-t="commission"]');
  if(commBtn){
    commBtn.onclick = function(){ forceOpenCommissionTab(); };
  }
  const reportBtn = document.querySelector('#t-commission button[onclick="genCommissionReport()"]');
  if(reportBtn){
    reportBtn.addEventListener('click', ()=>{ try{ genCommissionReport(); }catch(err){ console.error(err); } });
  }
  try{ initCommission(); }catch(err){ console.error('commission preload error:', err); }
});
function switchCommCard(type){
  const policyEl = document.getElementById('comm-policy-detail');
  const chartEl = document.getElementById('commChartSection');
  if(type==='policy'){
    if(policyEl){ policyEl.style.display = policyEl.style.display==='none'?'block':'none'; }
    if(chartEl) chartEl.style.display='block';
    switchCommChart('policy');
  } else if(type==='mgmt'){
    if(policyEl) policyEl.style.display='none';
    if(chartEl) chartEl.style.display='block';
    switchCommChart('mgmt');
  } else if(type==='total'){
    if(policyEl) policyEl.style.display='none';
    if(chartEl) chartEl.style.display='block';
    switchCommChart('total');
  }
}

// ==================== 수수료 분석 탭 ====================
let commTab = 0;

function initCommission(){
  // KPI 카드 UI(관리수수료/수수료수입 합계 등) 제거
  switchComm(4);
}

function switchComm(idx){
  commTab = idx;
  [0,1,2,3,4].forEach(i=>{
    const s = document.getElementById('comm-s'+i);
    if(s) s.style.display = i===idx?'block':'none';
    const b = document.getElementById('csBtn'+i);
    if(b){ b.style.opacity = i===idx?'1':'0.65'; }
  });
  if(idx===0) renderCommTrend();
  else if(idx===1) renderCommMonthly();
  else if(idx===2) renderCommChannel();
  else if(idx===3) renderCommRisk();
  else if(idx===4) renderComm1Jan();
}

function renderCommKpi(){
  const C = D.commission;
  const j = C.jan26;
  const y25 = C.yearly.find(y=>y.yr==='25년');
  const y24 = C.yearly.find(y=>y.yr==='24년');
  const y26e = C.yearly.find(y=>y.yr==='26년(E)');
  const chg = ((j.mgmt_fee.total - j.mgmt_fee_25jan)/j.mgmt_fee_25jan*100).toFixed(1);
  const vsAvg = ((j.mgmt_fee.total - j.mgmt_fee_25avg)/j.mgmt_fee_25avg*100).toFixed(1);
  document.getElementById('commKpi').innerHTML = `
    <div class="sg-card" style="border-top:3px solid #059669"><div class="sg-v">${j.mgmt_fee.total.toFixed(1)}<span style="font-size:11px">억</span></div><div class="sg-l">26년 1월 관리수수료</div><div class="sg-c ${parseFloat(chg)<0?'neg':'pos'}">${parseFloat(chg)>0?'+':''}${chg}% yoy</div></div>
    <div class="sg-card" style="border-top:3px solid #059669"><div class="sg-v">${j.total_fee.total.toFixed(1)}<span style="font-size:11px">억</span></div><div class="sg-l">26년 1월 수수료수입 합계</div><div class="sg-c">전사 기준</div></div>
    <div class="sg-card"><div class="sg-v">${j.mgmt_fee_25jan.toFixed(1)}<span style="font-size:11px">억</span></div><div class="sg-l">25년 1월 관리수수료</div><div class="sg-c">전년 동월</div></div>
    <div class="sg-card"><div class="sg-v">${j.mgmt_fee_25avg.toFixed(1)}<span style="font-size:11px">억</span></div><div class="sg-l">25년 관리수수료 월평균</div><div class="sg-c ${parseFloat(vsAvg)<0?'neg':'pos'}">${parseFloat(vsAvg)>0?'+':''}${vsAvg}% vs 평균</div></div>
  `;
}

function commBar(label, val, max, color, suffix='억'){
  const pct = Math.max(3, (val/max*100));
  return `<div style="display:flex;align-items:center;gap:8px;margin-bottom:7px">
    <div style="width:70px;font-size:11px;color:var(--t2);text-align:right;flex-shrink:0">${label}</div>
    <div style="flex:1;background:var(--bg3);border-radius:4px;height:20px;overflow:hidden">
      <div style="width:${pct}%;height:100%;background:${color};border-radius:4px;display:flex;align-items:center;padding-left:6px">
        <span style="font-size:10px;color:#fff;font-weight:600;white-space:nowrap">${val.toFixed(1)}${suffix}</span>
      </div>
    </div>
  </div>`;
}

function renderComm1Jan(){
  const C = D.commission;
  const j = C.jan26;
  const el = document.getElementById('comm-s4');
  const mgmt = j.mgmt_fee;
  const total = j.total_fee;
  const chans = ['retail','wholesale','digital','enterprise','iot','corporate_sales','small_biz'];
  const cNames = {retail:'소매',wholesale:'도매',digital:'디지털',enterprise:'기업/공공',iot:'IoT',corporate_sales:'법인영업',small_biz:'소상공인'};
  const cColors = {retail:'#38bdf8',wholesale:'#34d399',digital:'#a78bfa',enterprise:'#fb923c',iot:'#fbbf24',corporate_sales:'#f87171',small_biz:'#818cf8'};

  const mgmtMax = chans.reduce((m,ch)=>Math.max(m,mgmt[ch]||0),0);
  const totalMax = chans.reduce((m,ch)=>Math.max(m,total[ch]||0),0);

  // 수수료수입 세부 항목 행 생성
  const itemRows = j.fee_items.map((it,idx)=>{
    const isNeg = it.total < 0;
    const bg = isNeg ? 'background:rgba(239,68,68,0.06)' : (idx%2===0?'':'background:rgba(255,255,255,0.03)');
    return `<tr style="${bg}">
      <td style="padding:7px 8px;white-space:nowrap;${isNeg?'color:#f87171;font-style:italic':''}">${it.nm}</td>
      <td style="padding:7px 8px;text-align:right;font-weight:700;${isNeg?'color:#f87171':'color:var(--t1)'}">${it.total.toFixed(1)}</td>
      <td style="padding:7px 8px;text-align:right;color:#38bdf8">${it.retail.toFixed(1)}</td>
      <td style="padding:7px 8px;text-align:right;color:#34d399">${it.wholesale.toFixed(1)}</td>
      <td style="padding:7px 8px;text-align:right;color:#a78bfa">${it.digital.toFixed(1)}</td>
      <td style="padding:7px 8px;text-align:right;color:#fb923c">${it.enterprise.toFixed(1)}</td>
      <td style="padding:7px 8px;text-align:right;color:#fbbf24">${it.iot.toFixed(1)}</td>
      <td style="padding:7px 8px;text-align:right;color:#f87171">${it.corporate_sales.toFixed(1)}</td>
      <td style="padding:7px 8px;text-align:right;color:#818cf8">${it.small_biz.toFixed(1)}</td>
      <td style="padding:7px 8px;text-align:right;color:var(--t3)">${(it.common||0).toFixed(1)}</td>
    </tr>`;
  }).join('');

  const yoyDiff = mgmt.total - j.mgmt_fee_25jan;
  const yoyChg = (yoyDiff/j.mgmt_fee_25jan*100).toFixed(1);
  const vsAvgDiff = mgmt.total - j.mgmt_fee_25avg;
  const vsAvg = (vsAvgDiff/j.mgmt_fee_25avg*100).toFixed(1);
  const mgmtMix = (mgmt.total/total.total*100).toFixed(1);
  const yoyNeg = parseFloat(yoyChg)<0;
  const avgNeg = parseFloat(vsAvg)<0;

  el.innerHTML = `
  <!-- ── KPI 상단 요약 카드 3개: 관리수수료 / 정책수수료 / 수수료수입 ── -->
  <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-bottom:14px">
    <!-- 관리수수료 카드 -->
    <div style="background:linear-gradient(135deg,${yoyNeg?'#4c0519,#991b1b':'#052e16,#14532d'});border-radius:12px;padding:12px 10px;border:1px solid ${yoyNeg?'#f8717133':'#22c55e33'};cursor:pointer" onclick="switchCommCard('mgmt')">
      <div style="font-size:9px;color:${yoyNeg?'#fca5a5':'#86efac'};font-weight:700;letter-spacing:.5px;margin-bottom:4px">🏷 관리수수료</div>
      <div style="font-size:20px;font-weight:900;color:#fff;line-height:1">${mgmt.total.toFixed(1)}<span style="font-size:11px;font-weight:600;color:${yoyNeg?'#fca5a5':'#86efac'}">억</span></div>
      <div style="font-size:9px;margin-top:4px;color:${yoyNeg?'#fca5a5':'#86efac'}"><strong style="color:${yoyNeg?'#ef4444':'#22c55e'}">${yoyNeg?'':'+'}${yoyChg}%</strong> yoy</div>
      <div style="font-size:8px;margin-top:3px;color:rgba(255,255,255,.5)">전년동월 ${j.mgmt_fee_25jan}억</div>
    </div>
    <!-- 정책수수료 카드 -->
    <div style="background:linear-gradient(135deg,#1e1b4b,#312e81);border-radius:12px;padding:12px 10px;border:1px solid #818cf833;cursor:pointer" onclick="switchCommCard('policy')">
      <div style="font-size:9px;color:#c4b5fd;font-weight:700;letter-spacing:.5px;margin-bottom:4px">📋 무선 정책수수료</div>
      <div style="font-size:20px;font-weight:900;color:#fff;line-height:1">${(j.policy_fee.total).toFixed(1)}<span style="font-size:11px;font-weight:600;color:#c4b5fd">억</span></div>
      <div style="font-size:9px;margin-top:4px;color:#c4b5fd"><strong style="color:#a5b4fc">${(((j.policy_fee.total)-j.policy_fee_25jan)/j.policy_fee_25jan*100)>0?'+':''}${(((j.policy_fee.total)-j.policy_fee_25jan)/j.policy_fee_25jan*100).toFixed(1)}%</strong> yoy</div>
      <div style="font-size:8px;margin-top:3px;color:rgba(255,255,255,.5)">KT-RDS ${j.policy_fee.rds}억 + 정책 ${j.policy_fee.policy}억</div>
    </div>
    <!-- 수수료수입 합계 카드 -->
    <div style="background:linear-gradient(135deg,#0c4a2e,#166534);border-radius:12px;padding:12px 10px;border:1px solid #22c55e33;cursor:pointer" onclick="switchCommCard('total')">
      <div style="font-size:9px;color:#86efac;font-weight:700;letter-spacing:.5px;margin-bottom:4px">💰 수수료수입 합계</div>
      <div style="font-size:20px;font-weight:900;color:#fff;line-height:1">${total.total.toFixed(1)}<span style="font-size:11px;font-weight:600;color:#86efac">억</span></div>
      <div style="font-size:9px;margin-top:4px;color:#86efac">관리 <strong style="color:#fff">${mgmtMix}%</strong> / 정책 <strong style="color:#c4b5fd">${(j.policy_fee.total/total.total*100).toFixed(1)}%</strong></div>
      <div style="font-size:8px;margin-top:3px;color:rgba(255,255,255,.5)">전사 합계 기준</div>
    </div>
  </div>

  <!-- ── 정책수수료 상세 섹션 ── -->
  <div id="comm-policy-detail" style="background:var(--bg2);border-radius:12px;margin-bottom:12px;border:1px solid #818cf844;overflow:hidden;display:none">
    <div style="background:linear-gradient(90deg,#312e81,#1e1b4b);padding:10px 14px;display:flex;align-items:center;justify-content:space-between">
      <div style="font-size:12px;font-weight:800;color:#e0e7ff">📋 무선 정책 수수료 상세 분석</div>
      <div style="font-size:10px;color:#a5b4fc">수수료수입 정책 + 부분 RDS 합산</div>
    </div>
    <div style="padding:12px">
      <!-- 구성 항목 -->
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:10px">
        <div style="background:rgba(129,140,248,.1);border-radius:8px;padding:10px;border:1px solid #818cf833">
          <div style="font-size:9px;color:#a5b4fc;font-weight:700;margin-bottom:4px">KT-RDS 수수료</div>
          <div style="font-size:18px;font-weight:800;color:#e0e7ff">${j.policy_fee.rds.toFixed(1)}<span style="font-size:10px">억</span></div>
          <div style="font-size:9px;color:#818cf8;margin-top:3px">소매·디지털 채널 중심</div>
        </div>
        <div style="background:rgba(99,102,241,.1);border-radius:8px;padding:10px;border:1px solid #6366f133">
          <div style="font-size:9px;color:#a5b4fc;font-weight:700;margin-bottom:4px">수수료수입 정책</div>
          <div style="font-size:18px;font-weight:800;color:#e0e7ff">${j.policy_fee.policy.toFixed(1)}<span style="font-size:10px">억</span></div>
          <div style="font-size:9px;color:#818cf8;margin-top:3px">디지털 채널 (${j.policy_fee.digital.toFixed(1)}억 중 대부분)</div>
        </div>
      </div>
      <!-- 채널별 분해 -->
      <div style="font-size:11px;font-weight:700;color:var(--t1);margin-bottom:8px">채널별 수수료수입 정책 현황 (26년 1월)</div>
      ${['retail','digital','enterprise','corporate_sales','small_biz'].map(ch=>{
        const v = j.policy_fee[ch]||0;
        const maxV = Math.max(...['retail','digital','enterprise','corporate_sales','small_biz'].map(c=>Math.abs(j.policy_fee[c]||0)));
        const pctW = maxV>0?Math.max(3,Math.abs(v)/maxV*80):3;
        const isNeg = v<0;
        const cName = {retail:'소매',digital:'디지털',enterprise:'기업/공공',corporate_sales:'법인영업',small_biz:'소상공인'}[ch];
        return `<div style="display:flex;align-items:center;gap:8px;margin-bottom:6px">
          <div style="width:64px;font-size:10px;color:var(--t2);text-align:right;flex-shrink:0">${cName}</div>
          <div style="flex:1;background:var(--bg3);border-radius:4px;height:18px;overflow:hidden">
            <div style="width:${pctW}%;height:100%;background:${isNeg?'#ef4444':'#818cf8'};border-radius:4px;display:flex;align-items:center;padding-left:5px">
              <span style="font-size:9px;color:#fff;font-weight:600;white-space:nowrap">${v.toFixed(1)}억</span>
            </div>
          </div>
        </div>`;
      }).join('')}
      <div style="margin-top:10px;padding:8px 10px;background:rgba(129,140,248,.08);border-radius:6px;border-left:3px solid #818cf8">
        <div style="font-size:10px;color:#a5b4fc;font-weight:700;margin-bottom:3px">📌 특징</div>
        <div style="font-size:10px;color:var(--t2);line-height:1.5">
          · KT-RDS(${j.policy_fee.rds}억) + 정책(${j.policy_fee.policy}억) = <strong style="color:#c4b5fd">합계 ${j.policy_fee.total}억</strong><br>
          · <strong style="color:#c4b5fd">소매(${j.policy_fee.retail}억)</strong> 및 <strong style="color:#a78bfa">디지털(${j.policy_fee.digital}억)</strong>이 전체의 ${((j.policy_fee.retail+j.policy_fee.digital)/j.policy_fee.total*100).toFixed(0)}% 차지<br>
          · 법인영업·기업/공공 채널 정책 차감(-) 구조 — 수익성 모니터링 필요
        </div>
      </div>
    </div>
  </div>

  <!-- ── 이중 섹션: 관리수수료 + 수수료수입 ── -->
  <div id="commChartSection" style="background:var(--bg2);border-radius:12px;overflow:hidden;margin-bottom:12px;border:1px solid var(--bd)">
    <!-- 탭 헤더 -->
    <div style="display:flex;border-bottom:1px solid var(--bd)" id="commChartTabs">
      <button onclick="switchCommChart('mgmt')" id="cct-mgmt"
        style="flex:1;padding:10px 4px;font-size:10px;font-weight:700;border:none;cursor:pointer;background:#059669;color:#fff;border-radius:0;transition:all .2s;line-height:1.3">
        🏷 관리수수료
      </button>
      <button onclick="switchCommChart('policy')" id="cct-policy"
        style="flex:1;padding:10px 4px;font-size:10px;font-weight:700;border:none;cursor:pointer;background:var(--bg3);color:var(--t2);border-radius:0;transition:all .2s;line-height:1.3">
        📋 정책수수료
      </button>
      <button onclick="switchCommChart('total')" id="cct-total"
        style="flex:1;padding:10px 4px;font-size:10px;font-weight:700;border:none;cursor:pointer;background:var(--bg3);color:var(--t2);border-radius:0;transition:all .2s;line-height:1.3">
        💰 수수료수입
      </button>
    </div>

    <!-- 관리수수료 차트 -->
    <div id="commChart-mgmt" style="padding:14px">
      <div style="font-size:10px;color:var(--t3);margin-bottom:12px">채널소계 ${mgmt.channel_subtotal.toFixed(1)}억 기준 · 공통비 ${mgmt.common.toFixed(1)}억 별도</div>
      ${chans.map(ch=>{
        const v = mgmt[ch]||0;
        const pct = Math.max(2,(v/mgmtMax*100));
        const share = mgmt.channel_subtotal>0?(v/mgmt.channel_subtotal*100).toFixed(1):'0.0';
        return `<div style="display:flex;align-items:center;gap:8px;margin-bottom:10px">
          <div style="width:60px;font-size:11px;color:var(--t2);text-align:right;flex-shrink:0;font-weight:600">${cNames[ch]}</div>
          <div style="flex:1;position:relative">
            <div style="background:var(--bg3);border-radius:6px;height:28px;overflow:hidden">
              <div style="width:${pct}%;height:100%;background:${cColors[ch]};border-radius:6px;display:flex;align-items:center;padding-left:10px;transition:width .5s ease">
                <span style="font-size:11px;color:#fff;font-weight:800;white-space:nowrap;text-shadow:0 1px 2px rgba(0,0,0,.4)">${v.toFixed(1)}억</span>
              </div>
            </div>
          </div>
          <div style="width:44px;text-align:right;flex-shrink:0">
            <span style="font-size:10px;color:${cColors[ch]};font-weight:700">${share}%</span>
          </div>
        </div>`;
      }).join('')}
      <div style="margin-top:12px;padding-top:10px;border-top:1px solid var(--bd);display:flex;justify-content:space-between;font-size:10px;color:var(--t3)">
        <span>전사합계 <strong style="color:var(--t1);font-size:12px">${mgmt.total.toFixed(1)}억</strong></span>
        <span>소매 비중 <strong style="color:#38bdf8">${(mgmt.retail/mgmt.channel_subtotal*100).toFixed(1)}%</strong></span>
        <span>YoY <strong style="color:${yoyNeg?'#ef4444':'#22c55e'}">${yoyNeg?'':'+'}<strong>${yoyChg}%</strong></span>
      </div>
    </div>

    <!-- 무선 정책수수료 차트 (기본 숨김) -->
    <div id="commChart-policy" style="padding:14px;display:none">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px">
        <div style="font-size:10px;color:var(--t3)">KT-RDS ${j.policy_fee.rds}억 + 수수료수입 정책 ${j.policy_fee.policy}억 합산</div>
        <div style="font-size:10px;color:#a5b4fc;font-weight:700">합계 ${j.policy_fee.total}억</div>
      </div>
      ${(()=>{
        const policyMax = chans.reduce((m,ch)=>Math.max(m,j.policy_fee[ch]||0),0);
        return chans.map(ch=>{
          const v = j.policy_fee[ch]||0;
          const isNeg = v<0;
          const pct = Math.max(2,(Math.abs(v)/policyMax*100));
          const share = j.policy_fee.total>0?(v/j.policy_fee.total*100).toFixed(1):'0.0';
          return `<div style="display:flex;align-items:center;gap:8px;margin-bottom:10px">
            <div style="width:60px;font-size:11px;color:var(--t2);text-align:right;flex-shrink:0;font-weight:600">${cNames[ch]}</div>
            <div style="flex:1;position:relative">
              <div style="background:var(--bg3);border-radius:6px;height:28px;overflow:hidden">
                <div style="width:${pct}%;height:100%;background:${isNeg?'#ef4444':cColors[ch]};border-radius:6px;display:flex;align-items:center;padding-left:10px;transition:width .5s ease;opacity:${isNeg?.7:1}">
                  <span style="font-size:11px;color:#fff;font-weight:800;white-space:nowrap;text-shadow:0 1px 2px rgba(0,0,0,.4)">${isNeg?'▼':''} ${v.toFixed(1)}억</span>
                </div>
              </div>
            </div>
            <div style="width:44px;text-align:right;flex-shrink:0">
              <span style="font-size:10px;color:${isNeg?'#f87171':cColors[ch]};font-weight:700">${isNeg?'':'+'}${share}%</span>
            </div>
          </div>`;
        }).join('');
      })()}
      <div style="margin-top:12px;padding-top:10px;border-top:1px solid var(--bd)">
        <div style="display:flex;justify-content:space-between;font-size:10px;color:var(--t3);margin-bottom:6px">
          <span>전사합계 <strong style="color:#c4b5fd;font-size:12px">${j.policy_fee.total}억</strong></span>
          <span>소매+디지털 <strong style="color:#a5b4fc">${((j.policy_fee.retail+j.policy_fee.digital)/j.policy_fee.total*100).toFixed(0)}%</strong></span>
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px">
          <div style="background:rgba(129,140,248,.08);border-radius:6px;padding:7px 9px;border:1px solid #818cf822">
            <div style="font-size:9px;color:#a5b4fc;margin-bottom:2px">KT-RDS</div>
            <div style="font-size:13px;font-weight:800;color:#e0e7ff">${j.policy_fee.rds}억<span style="font-size:9px;color:#818cf8;font-weight:400;margin-left:4px">${(j.policy_fee.rds/j.policy_fee.total*100).toFixed(1)}%</span></div>
          </div>
          <div style="background:rgba(99,102,241,.08);border-radius:6px;padding:7px 9px;border:1px solid #6366f122">
            <div style="font-size:9px;color:#a5b4fc;margin-bottom:2px">수수료수입 정책</div>
            <div style="font-size:13px;font-weight:800;color:#e0e7ff">${j.policy_fee.policy}억<span style="font-size:9px;color:#818cf8;font-weight:400;margin-left:4px">${(j.policy_fee.policy/j.policy_fee.total*100).toFixed(1)}%</span></div>
          </div>
        </div>
      </div>
    </div>

    <!-- 수수료수입 합계 차트 (기본 숨김) -->
    <div id="commChart-total" style="padding:14px;display:none">
      <div style="font-size:10px;color:var(--t3);margin-bottom:12px">전사합계 ${total.total.toFixed(1)}억 · 공통비 ${(total.common_corp+total.common_ch).toFixed(1)}억 포함</div>
      ${chans.map(ch=>{
        const v = total[ch]||0;
        const pct = Math.max(2,(v/totalMax*100));
        const share = total.channel_subtotal>0?(v/total.channel_subtotal*100).toFixed(1):'0.0';
        return `<div style="display:flex;align-items:center;gap:8px;margin-bottom:10px">
          <div style="width:60px;font-size:11px;color:var(--t2);text-align:right;flex-shrink:0;font-weight:600">${cNames[ch]}</div>
          <div style="flex:1;position:relative">
            <div style="background:var(--bg3);border-radius:6px;height:28px;overflow:hidden">
              <div style="width:${pct}%;height:100%;background:${cColors[ch]};border-radius:6px;display:flex;align-items:center;padding-left:10px;transition:width .5s ease">
                <span style="font-size:11px;color:#fff;font-weight:800;white-space:nowrap;text-shadow:0 1px 2px rgba(0,0,0,.4)">${v.toFixed(1)}억</span>
              </div>
            </div>
          </div>
          <div style="width:44px;text-align:right;flex-shrink:0">
            <span style="font-size:10px;color:${cColors[ch]};font-weight:700">${share}%</span>
          </div>
        </div>`;
      }).join('')}
      <div style="margin-top:12px;padding-top:10px;border-top:1px solid var(--bd);font-size:10px;color:var(--t3)">
        관리수수료 비중 <strong style="color:#059669">${mgmtMix}%</strong>
      </div>
    </div>
  </div>

  <!-- ── 수수료수입 세부 테이블 ── -->
  <div style="background:var(--bg2);border-radius:12px;margin-bottom:12px;border:1px solid var(--bd);overflow:hidden">
    <div style="padding:12px 14px 8px;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid var(--bd)">
      <div style="font-size:13px;font-weight:700;color:var(--t1)">📊 수수료수입 항목별 채널 분해</div>
      <div style="font-size:10px;color:var(--t3)">단위: 억원</div>
    </div>
    <div style="overflow-x:auto">
    <table style="width:100%;border-collapse:collapse;font-size:11px;min-width:560px">
      <thead>
        <tr style="background:var(--bg3)">
          <th style="padding:8px 8px;text-align:left;white-space:nowrap;color:var(--t2);font-weight:700;position:sticky;left:0;background:var(--bg3);z-index:1">항목</th>
          <th style="padding:8px 8px;text-align:right;color:#22c55e;font-weight:700">합계</th>
          <th style="padding:8px 6px;text-align:right"><span style="display:inline-block;width:6px;height:6px;border-radius:50%;background:#38bdf8;margin-right:3px;vertical-align:middle"></span><span style="color:#38bdf8">소매</span></th>
          <th style="padding:8px 6px;text-align:right"><span style="display:inline-block;width:6px;height:6px;border-radius:50%;background:#34d399;margin-right:3px;vertical-align:middle"></span><span style="color:#34d399">도매</span></th>
          <th style="padding:8px 6px;text-align:right"><span style="display:inline-block;width:6px;height:6px;border-radius:50%;background:#a78bfa;margin-right:3px;vertical-align:middle"></span><span style="color:#a78bfa">디지털</span></th>
          <th style="padding:8px 6px;text-align:right"><span style="display:inline-block;width:6px;height:6px;border-radius:50%;background:#fb923c;margin-right:3px;vertical-align:middle"></span><span style="color:#fb923c">기업</span></th>
          <th style="padding:8px 6px;text-align:right"><span style="display:inline-block;width:6px;height:6px;border-radius:50%;background:#fbbf24;margin-right:3px;vertical-align:middle"></span><span style="color:#fbbf24">IoT</span></th>
          <th style="padding:8px 6px;text-align:right"><span style="display:inline-block;width:6px;height:6px;border-radius:50%;background:#f87171;margin-right:3px;vertical-align:middle"></span><span style="color:#f87171">법인</span></th>
          <th style="padding:8px 6px;text-align:right"><span style="display:inline-block;width:6px;height:6px;border-radius:50%;background:#818cf8;margin-right:3px;vertical-align:middle"></span><span style="color:#818cf8">소상</span></th>
          <th style="padding:8px 6px;text-align:right;color:var(--t3)">공통</th>
        </tr>
      </thead>
      <tbody>
        ${itemRows}
        <tr style="border-top:2px solid var(--bd);font-weight:700;background:var(--bg3)">
          <td style="padding:8px 8px;color:var(--t2);position:sticky;left:0;background:var(--bg3)">채널소계</td>
          <td style="padding:8px 8px;text-align:right;color:var(--t1)">${total.channel_subtotal.toFixed(1)}</td>
          <td style="padding:8px 6px;text-align:right;color:#38bdf8">${total.retail.toFixed(1)}</td>
          <td style="padding:8px 6px;text-align:right;color:#34d399">${total.wholesale.toFixed(1)}</td>
          <td style="padding:8px 6px;text-align:right;color:#a78bfa">${total.digital.toFixed(1)}</td>
          <td style="padding:8px 6px;text-align:right;color:#fb923c">${total.enterprise.toFixed(1)}</td>
          <td style="padding:8px 6px;text-align:right;color:#fbbf24">${total.iot.toFixed(1)}</td>
          <td style="padding:8px 6px;text-align:right;color:#f87171">${total.corporate_sales.toFixed(1)}</td>
          <td style="padding:8px 6px;text-align:right;color:#818cf8">${total.small_biz.toFixed(1)}</td>
          <td style="padding:8px 6px;text-align:right;color:var(--t3)">-</td>
        </tr>
        <tr style="font-weight:800;background:rgba(5,150,105,.12);border-top:1px solid #059669">
          <td style="padding:8px 8px;color:#059669;position:sticky;left:0;background:rgba(5,150,105,.12)">전사합계</td>
          <td style="padding:8px 8px;text-align:right;color:#22c55e;font-size:13px">${total.total.toFixed(1)}</td>
          <td colspan="7" style="padding:8px 6px;text-align:center;font-size:10px;color:var(--t3)">전사공통비(${total.common_corp.toFixed(1)}) + 채널공통비(${total.common_ch.toFixed(1)}) 포함</td>
          <td style="padding:8px 6px;text-align:right;color:var(--t3)">${(total.common_corp+total.common_ch).toFixed(1)}</td>
        </tr>
      </tbody>
    </table>
    </div>
    <div style="padding:8px 14px;font-size:10px;color:var(--t3);border-top:1px solid var(--bd)">※ 합계=B열(전사합계) 기준. 순액처리(-10.2억)는 전사공통비에 포함. 채널별 비중(%)은 채널소계 대비로 산출.</div>
  </div>

  <!-- ── 인사이트 패널 ── -->
  <div style="background:var(--bg2);border-radius:12px;padding:14px;border:1px solid var(--bd);border-left:3px solid #8b5cf6">
    <div style="font-size:12px;font-weight:800;color:var(--t1);margin-bottom:12px">📌 1월 관리수수료 분석</div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">

      <div style="background:${yoyNeg?'rgba(239,68,68,.08)':'rgba(34,197,94,.08)'};border-radius:8px;padding:10px;border:1px solid ${yoyNeg?'rgba(239,68,68,.2)':'rgba(34,197,94,.2)'}">
        <div style="font-size:9px;color:${yoyNeg?'#fca5a5':'#86efac'};font-weight:700;margin-bottom:6px">📅 YOY 비교</div>
        <div style="font-size:10px;color:var(--t2);margin-bottom:4px">26년 1월 <strong style="color:var(--t1)">${mgmt.total.toFixed(1)}억</strong> vs 25년 <strong>${j.mgmt_fee_25jan}억</strong></div>
        <div style="font-size:14px;font-weight:800;color:${yoyNeg?'#ef4444':'#22c55e'}">${yoyNeg?'':'+'}${yoyChg}%</div>
        <div style="font-size:9px;color:var(--t3);margin-top:2px">${Math.abs(yoyDiff).toFixed(1)}억 ${yoyNeg?'감소':'증가'}</div>
      </div>

      <div style="background:${avgNeg?'rgba(245,158,11,.08)':'rgba(56,189,248,.08)'};border-radius:8px;padding:10px;border:1px solid ${avgNeg?'rgba(245,158,11,.2)':'rgba(56,189,248,.2)'}">
        <div style="font-size:9px;color:${avgNeg?'#fcd34d':'#7dd3fc'};font-weight:700;margin-bottom:6px">📊 평균 대비</div>
        <div style="font-size:10px;color:var(--t2);margin-bottom:4px">25년 월평균 <strong>${j.mgmt_fee_25avg}억</strong> 대비</div>
        <div style="font-size:14px;font-weight:800;color:${avgNeg?'#f59e0b':'#38bdf8'}">${avgNeg?'':'+'}${vsAvg}%</div>
        <div style="font-size:9px;color:var(--t3);margin-top:2px">${Math.abs(vsAvgDiff).toFixed(1)}억 ${avgNeg?'하회':'상회'}</div>
      </div>

      <div style="background:rgba(59,130,246,.08);border-radius:8px;padding:10px;border:1px solid rgba(59,130,246,.2)">
        <div style="font-size:9px;color:#93c5fd;font-weight:700;margin-bottom:6px">🏗 채널 구조</div>
        <div style="font-size:10px;color:var(--t2)"><strong style="color:#38bdf8">소매</strong> ${mgmt.retail.toFixed(1)}억 <span style="color:var(--t3)">(${(mgmt.retail/mgmt.total*100).toFixed(0)}%)</span></div>
        <div style="font-size:10px;color:var(--t2);margin-top:3px"><strong style="color:#34d399">도매</strong> ${mgmt.wholesale.toFixed(1)}억 · <strong style="color:#a78bfa">디지털</strong> ${mgmt.digital.toFixed(1)}억</div>
      </div>

      <div style="background:rgba(139,92,246,.08);border-radius:8px;padding:10px;border:1px solid rgba(139,92,246,.2)">
        <div style="font-size:9px;color:#c4b5fd;font-weight:700;margin-bottom:6px">💡 향후 분석</div>
        <div style="font-size:10px;color:var(--t2)"><strong>예측_무선_관리수수료_추세</strong> 업로드 시</div>
        <div style="font-size:10px;color:var(--t3);margin-top:3px">연도별 추이 통합 분석 자동 전환</div>
      </div>
    </div>
  </div>
  `;
}

function switchCommChart(type){
  const mgmtPanel   = document.getElementById('commChart-mgmt');
  const policyPanel = document.getElementById('commChart-policy');
  const totalPanel  = document.getElementById('commChart-total');
  const mgmtBtn     = document.getElementById('cct-mgmt');
  const policyBtn   = document.getElementById('cct-policy');
  const totalBtn    = document.getElementById('cct-total');
  if(!mgmtPanel||!totalPanel) return;
  // 전체 숨김
  [mgmtPanel,policyPanel,totalPanel].forEach(p=>{ if(p) p.style.display='none'; });
  [mgmtBtn,policyBtn,totalBtn].forEach(b=>{ if(b){ b.style.background='var(--bg3)'; b.style.color='var(--t2)'; }});
  // 선택 탭 활성화
  if(type==='mgmt'){
    mgmtPanel.style.display='block';
    if(mgmtBtn){ mgmtBtn.style.background='#059669'; mgmtBtn.style.color='#fff'; }
  } else if(type==='policy'){
    if(policyPanel) policyPanel.style.display='block';
    if(policyBtn){ policyBtn.style.background='#4f46e5'; policyBtn.style.color='#fff'; }
  } else {
    totalPanel.style.display='block';
    if(totalBtn){ totalBtn.style.background='#0369a1'; totalBtn.style.color='#fff'; }
  }
}

function renderCommTrend(){
  const C = D.commission;
  const el = document.getElementById('comm-s0');
  const chColors = {retail:'#38bdf8',wholesale:'#34d399',digital:'#a78bfa',b2b:'#fb923c',small_biz:'#818cf8'};
  const chNames = {retail:'소매',wholesale:'도매',digital:'디지털',b2b:'B2B',small_biz:'소상공인'};

  // Yearly grouped bars
  let barsHtml = '';
  const maxVal = 750;
  C.yearly.forEach(y=>{
    barsHtml += `<div style="margin-bottom:14px">
      <div style="font-size:11px;font-weight:700;color:var(--t1);margin-bottom:4px">${y.yr} <span style="color:var(--t3);font-weight:400">합계 ${y.total.toFixed(0)}억</span></div>
      ${commBar('소매', y.retail, maxVal, chColors.retail)}
      ${commBar('도매', y.wholesale, maxVal, chColors.wholesale)}
      ${commBar('디지털', y.digital, maxVal, chColors.digital)}
      ${y.b2b>0?commBar('B2B', y.b2b, maxVal, chColors.b2b):''}
      ${y.small_biz>0?commBar('소상공인', y.small_biz, maxVal, chColors.small_biz):''}
    </div>`;
  });

  // YoY change table
  const y24 = C.yearly[2], y25 = C.yearly[3], y26e = C.yearly[4];
  const chKeys = ['retail','wholesale','digital','b2b','small_biz'];
  let tableRows = chKeys.map(k=>{
    const chg24_25 = ((y25[k]-y24[k])/y24[k]*100).toFixed(1);
    const chg25_26 = ((y26e[k]-y25[k])/y25[k]*100).toFixed(1);
    const clr24 = parseFloat(chg24_25)<0?'#f87171':'#34d399';
    const clr25 = parseFloat(chg25_26)<0?'#f87171':'#34d399';
    return `<tr>
      <td style="font-weight:600">${chNames[k]}</td>
      <td style="text-align:right">${y24[k].toFixed(1)}</td>
      <td style="text-align:right">${y25[k].toFixed(1)}</td>
      <td style="text-align:right;color:${clr24}">${chg24_25>0?'+':''}${chg24_25}%</td>
      <td style="text-align:right">${y26e[k].toFixed(1)}</td>
      <td style="text-align:right;color:${clr25}">${chg25_26>0?'+':''}${chg25_26}%</td>
    </tr>`;
  }).join('');

  el.innerHTML = `
    <div class="card">
      <div class="ct">채널별 연도별 관리수수료 <span style="font-size:10px;color:#6366f1;background:#ede9fe;padding:1px 6px;border-radius:10px">억원</span></div>
      ${barsHtml}
    </div>
    <div class="card">
      <div class="ct">채널별 전년대비 증감 분석</div>
      <div style="overflow-x:auto">
      <table style="width:100%;font-size:11px;border-collapse:collapse">
        <thead><tr style="border-bottom:1px solid var(--border)">
          <th style="text-align:left;padding:4px;color:var(--t2)">채널</th>
          <th style="text-align:right;padding:4px;color:var(--t2)">24년</th>
          <th style="text-align:right;padding:4px;color:var(--t2)">25년</th>
          <th style="text-align:right;padding:4px;color:var(--t2)">증감</th>
          <th style="text-align:right;padding:4px;color:var(--t2)">26(E)</th>
          <th style="text-align:right;padding:4px;color:var(--t2)">증감</th>
        </tr></thead>
        <tbody style="font-size:11px">${tableRows}</tbody>
      </table>
      </div>
      <div class="warn" style="margin-top:10px">⚠️ <strong>도매 수수료 급감:</strong> 24년 141.3억 → 25년 136.4억 → 26년(E) 109.6억. Win-Win HC 증가로 구조 변화 중</div>
    </div>
  `;
}

function renderCommMonthly(){
  const C = D.commission;
  const el = document.getElementById('comm-s1');
  const months = C.monthly25.map(m=>m.mo);
  const chKeys = ['retail','wholesale','digital','b2b','small_biz'];
  const chNames = {retail:'소매',wholesale:'도매',digital:'디지털',b2b:'B2B',small_biz:'소상공인'};
  const chColors = {retail:'#38bdf8',wholesale:'#34d399',digital:'#a78bfa',b2b:'#fb923c',small_biz:'#818cf8'};

  // Stacked bar chart by month
  const maxTotal = Math.max(...C.monthly25.map(m=>m.total)) * 1.1;
  let monthBars = C.monthly25.map(m=>{
    let segments = chKeys.map(k=>{
      const pct = (m[k]/maxTotal*100).toFixed(1);
      return `<div title="${chNames[k]}: ${m[k].toFixed(2)}억" style="width:${pct}%;background:${chColors[k]};height:100%"></div>`;
    }).join('');
    return `<div style="margin-bottom:10px">
      <div style="display:flex;justify-content:space-between;font-size:11px;margin-bottom:3px">
        <span style="color:var(--t2);font-weight:600">${m.mo}</span>
        <span style="color:var(--t1);font-weight:700">합계 ${m.total.toFixed(1)}억</span>
      </div>
      <div style="display:flex;height:22px;border-radius:4px;overflow:hidden;background:var(--bg3)">${segments}</div>
    </div>`;
  }).join('');

  // Legend
  let legend = chKeys.map(k=>`<span style="display:inline-flex;align-items:center;gap:3px;margin-right:10px;font-size:10px"><span style="width:10px;height:10px;border-radius:2px;background:${chColors[k]};display:inline-block"></span>${chNames[k]}</span>`).join('');

  // Detail table
  let tableRows = C.monthly25.map(m=>`<tr style="border-bottom:1px solid var(--border)">
    <td style="font-weight:600;padding:4px">${m.mo}</td>
    <td style="text-align:right;padding:4px;color:#38bdf8">${m.retail.toFixed(2)}</td>
    <td style="text-align:right;padding:4px;color:#34d399">${m.wholesale.toFixed(2)}</td>
    <td style="text-align:right;padding:4px;color:#a78bfa">${m.digital.toFixed(2)}</td>
    <td style="text-align:right;padding:4px;color:#fb923c">${m.b2b.toFixed(2)}</td>
    <td style="text-align:right;padding:4px;color:#818cf8">${m.small_biz.toFixed(2)}</td>
    <td style="text-align:right;padding:4px;font-weight:700">${m.total.toFixed(2)}</td>
  </tr>`).join('');

  // Wholesale detail
  const wd = C.wholesale_detail;
  let wdBars = C.monthly25.map((m,i)=>{
    const items = [
      {nm:'무선',v:wd.wireless[i],c:'#38bdf8'},
      {nm:'점프업',v:wd.jumpup[i],c:'#fbbf24'},
      {nm:'유선',v:wd.wired[i],c:'#34d399'},
      {nm:'H.C',v:wd.hc[i],c:'#a78bfa'},
      {nm:'Win-Win',v:wd.winwin[i],c:'#f87171'}
    ];
    const tot = items.reduce((a,x)=>a+x.v,0);
    const segs = items.map(x=>`<div title="${x.nm}: ${x.v.toFixed(2)}억" style="width:${(x.v/tot*100).toFixed(1)}%;background:${x.c};height:100%"></div>`).join('');
    return `<div style="margin-bottom:8px">
      <div style="display:flex;justify-content:space-between;font-size:10px;margin-bottom:2px">
        <span style="color:var(--t2)">${m.mo}</span>
        <span style="color:var(--t3)">${tot.toFixed(1)}억</span>
      </div>
      <div style="display:flex;height:18px;border-radius:3px;overflow:hidden;background:var(--bg3)">${segs}</div>
    </div>`;
  }).join('');

  el.innerHTML = `
    <div class="card">
      <div class="ct">25년 월별 관리수수료 현황</div>
      <div style="margin-bottom:8px">${legend}</div>
      ${monthBars}
    </div>
    <div class="card">
      <div class="ct">채널별 월별 상세 (억원)</div>
      <div style="overflow-x:auto">
      <table style="width:100%;font-size:11px;border-collapse:collapse">
        <thead><tr style="border-bottom:1px solid var(--border)">
          <th style="text-align:left;padding:4px;color:var(--t2)">월</th>
          <th style="text-align:right;padding:4px;color:#38bdf8">소매</th>
          <th style="text-align:right;padding:4px;color:#34d399">도매</th>
          <th style="text-align:right;padding:4px;color:#a78bfa">디지털</th>
          <th style="text-align:right;padding:4px;color:#fb923c">B2B</th>
          <th style="text-align:right;padding:4px;color:#818cf8">소상공인</th>
          <th style="text-align:right;padding:4px;color:var(--t1)">합계</th>
        </tr></thead>
        <tbody>${tableRows}</tbody>
      </table>
      </div>
    </div>
    <div class="card">
      <div class="ct">도매 수수료 세부 구성 (무선/점프업/유선/H.C/Win-Win)</div>
      ${wdBars}
      <div style="font-size:10px;color:var(--t3);margin-top:6px">
        <span style="color:#38bdf8">■무선</span> <span style="color:#fbbf24;margin-left:6px">■점프업</span> <span style="color:#34d399;margin-left:6px">■유선</span> <span style="color:#a78bfa;margin-left:6px">■H.C</span> <span style="color:#f87171;margin-left:6px">■Win-Win HC</span>
      </div>
    </div>
  `;
}

function renderCommChannel(){
  const C = D.commission;
  const el = document.getElementById('comm-s2');
  const y25 = C.yearly.find(y=>y.yr==='25년');
  const chKeys = ['retail','wholesale','digital','b2b','small_biz'];
  const chNames = {retail:'소매',wholesale:'도매',digital:'디지털',b2b:'B2B',small_biz:'소상공인'};
  const chColors = {retail:'#38bdf8',wholesale:'#34d399',digital:'#a78bfa',b2b:'#fb923c',small_biz:'#818cf8'};

  // Channel share bars
  const maxCh = Math.max(...chKeys.map(k=>y25[k]));
  let chBars = chKeys.map(k=>{
    const pct = (y25[k]/y25.total*100).toFixed(1);
    return commBar(`${chNames[k]} (${pct}%)`, y25[k], maxCh*1.1, chColors[k]);
  }).join('');

  // Retail detail
  const rd = C.retail_detail;
  const wAvg = rd.w25.reduce((a,v)=>a+v,0)/rd.w25.length;
  const lAvg = rd.l25.reduce((a,v)=>a+v,0)/rd.l25.length;
  const months = ['1월','2월','3월','4월','5월'];
  let retailRows = months.map((m,i)=>`<tr>
    <td style="padding:3px 4px">${m}</td>
    <td style="text-align:right;padding:3px 4px;color:#38bdf8">${rd.w25[i].toFixed(2)}</td>
    <td style="text-align:right;padding:3px 4px;color:#34d399">${rd.l25[i].toFixed(2)}</td>
    <td style="text-align:right;padding:3px 4px;font-weight:600">${(rd.w25[i]+rd.l25[i]).toFixed(2)}</td>
  </tr>`).join('');

  // Digital detail
  const dd = C.digital_detail;
  let digRows = months.map((m,i)=>`<tr>
    <td style="padding:3px 4px">${m}</td>
    <td style="text-align:right;padding:3px 4px;color:#38bdf8">${dd.wireless[i].toFixed(2)}</td>
    <td style="text-align:right;padding:3px 4px;color:#34d399">${dd.wired[i].toFixed(2)}</td>
    <td style="text-align:right;padding:3px 4px;color:#a78bfa">${dd.hc[i].toFixed(2)}</td>
    <td style="text-align:right;padding:3px 4px;font-weight:600">${(dd.wireless[i]+dd.wired[i]+dd.hc[i]).toFixed(2)}</td>
  </tr>`).join('');

  el.innerHTML = `
    <div class="card">
      <div class="ct">채널별 관리수수료 규모 (25년 연간)</div>
      ${chBars}
    </div>
    <div class="card">
      <div class="ct">소매 수수료 세부 (무선/유선)</div>
      <div style="overflow-x:auto">
      <table style="width:100%;font-size:11px;border-collapse:collapse">
        <thead><tr style="border-bottom:1px solid var(--border)">
          <th style="text-align:left;padding:4px;color:var(--t2)">월</th>
          <th style="text-align:right;padding:4px;color:#38bdf8">무선</th>
          <th style="text-align:right;padding:4px;color:#34d399">유선</th>
          <th style="text-align:right;padding:4px">합계</th>
        </tr></thead>
        <tbody>${retailRows}<tr style="border-top:1px solid var(--border);font-weight:700">
          <td style="padding:3px 4px">월평균</td>
          <td style="text-align:right;padding:3px 4px;color:#38bdf8">${wAvg.toFixed(2)}</td>
          <td style="text-align:right;padding:3px 4px;color:#34d399">${lAvg.toFixed(2)}</td>
          <td style="text-align:right;padding:3px 4px">${(wAvg+lAvg).toFixed(2)}</td>
        </tr></tbody>
      </table>
      </div>
    </div>
    <div class="card">
      <div class="ct">디지털 수수료 세부 (무선/유선/H.C)</div>
      <div style="overflow-x:auto">
      <table style="width:100%;font-size:11px;border-collapse:collapse">
        <thead><tr style="border-bottom:1px solid var(--border)">
          <th style="text-align:left;padding:4px;color:var(--t2)">월</th>
          <th style="text-align:right;padding:4px;color:#38bdf8">무선</th>
          <th style="text-align:right;padding:4px;color:#34d399">유선</th>
          <th style="text-align:right;padding:4px;color:#a78bfa">H.C</th>
          <th style="text-align:right;padding:4px">합계</th>
        </tr></thead>
        <tbody>${digRows}</tbody>
      </table>
      </div>
      <div style="font-size:10px;color:var(--t3);margin-top:6px">H.C: kt닷컴 위탁 이후 가입자 수납요금의 2% — 매월 증가 추세</div>
    </div>
  `;
}

function renderCommRisk(){
  const C = D.commission;
  const el = document.getElementById('comm-s3');
  const y24 = C.yearly[2], y25 = C.yearly[3], y26e = C.yearly[4];

  const risks = [
    {
      lvl:'HIGH', color:'#ef4444',
      item:'도매 수수료 급감',
      detail:`24년 141.3억 → 25년 136.4억 → 26년(E) 109.6억. -22.5% 감소 예상`,
      cause:'점프업 축소, 무선 가입자 감소로 기본 수수료 하락',
      action:'Win-Win HC 대체 효과 모니터링, 도매 구조 개편 검토'
    },
    {
      lvl:'HIGH', color:'#ef4444',
      item:'소매 무선 수수료 하락 추세',
      detail:`25년 1월 20.4억 → 5월 19.9억. 월별 감소세 지속`,
      cause:'무선 CAPA 유지에도 단건 수수료율 하락, 시장 포화',
      action:'고단가 요금제 전환 촉진, 유선 연계 수수료 확대'
    },
    {
      lvl:'MED', color:'#f59e0b',
      item:'디지털 채널 무선→H.C 수수료 구조 전환',
      detail:`디지털 무선: 5.43억('22) → 3.83억('25.1) → 2.09억('26E) 급감`,
      cause:'kt닷컴 100% 위탁 전환 이후 무선 직접 수수료 감소',
      action:'H.C 수익성 검증, 장기 위탁 계약 조건 재검토'
    },
    {
      lvl:'MED', color:'#f59e0b',
      item:'B2B 수수료 감소',
      detail:`24년 86.4억 → 25년 54.1억 (△37.4%). 법인 수익성 우려`,
      cause:'IoT/기업 영업 환경 악화, 수수료율 하락',
      action:'B2B 고수익 사업 포트폴리오 재편, IoT 솔루션 다각화'
    },
    {
      lvl:'LOW', color:'#22c55e',
      item:'소상공인 수수료 안정',
      detail:`25년 1~5월 월평균 1.26억 유지. 소폭 감소세이나 안정적`,
      cause:'유선 가입자 유지, 무선 소폭 감소',
      action:'소상공인 프랜차이즈 채널 확대로 반등 도모'
    }
  ];

  let riskCards = risks.map(r=>`
    <div style="border-left:3px solid ${r.color};padding:10px 12px;background:var(--bg2);border-radius:0 8px 8px 0;margin-bottom:10px">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px">
        <span style="font-weight:700;font-size:13px;color:var(--t1)">${r.item}</span>
        <span style="font-size:10px;background:${r.color}22;color:${r.color};padding:2px 8px;border-radius:10px;font-weight:700">${r.lvl}</span>
      </div>
      <div style="font-size:11px;color:var(--t2);margin-bottom:4px">📊 ${r.detail}</div>
      <div style="font-size:11px;color:var(--t3);margin-bottom:4px">🔍 원인: ${r.cause}</div>
      <div style="font-size:11px;color:#34d399">✅ 대응: ${r.action}</div>
    </div>
  `).join('');

  // YoY summary
  const totalChg = ((y25.total-y24.total)/y24.total*100).toFixed(1);
  const totalChg26 = ((y26e.total-y25.total)/y25.total*100).toFixed(1);

  el.innerHTML = `
    <div class="card">
      <div class="ct">수수료 리스크 워치리스트</div>
      ${riskCards}
    </div>
    <div class="card">
      <div class="ct">전략적 시사점</div>
      <div style="font-size:12px;line-height:1.8;color:var(--t2)">
        <strong style="color:var(--t1)">① [즉시] 도매 수수료 구조 점검</strong><br>
        &nbsp;&nbsp;점프업 축소로 26년 109.6억 감소 전망. Win-Win HC 계약 구조 재설계로 수익 보전 필요.<br><br>
        <strong style="color:var(--t1)">② [단기] 소매 무선 수수료 방어</strong><br>
        &nbsp;&nbsp;월별 감소세(20.4→19.9억) 가속 시 연간 10억+ 수익 감소. 고단가 전환 인센티브 강화 검토.<br><br>
        <strong style="color:var(--t1)">③ [중기] 디지털 H.C 수수료 성장 전략</strong><br>
        &nbsp;&nbsp;kt닷컴 HC가 무선 수수료 감소를 대체 중. 가입자 순증 목표(월 1.5만명) 달성 여부가 핵심 변수.
      </div>
    </div>
  `;
}

// ==================== 수수료 종합 보고서 ====================
function genCommissionReport(){
  const C = D.commission;
  const j = C.jan26;
  const mgmt = j.mgmt_fee;
  const total = j.total_fee;
  const y22 = C.yearly[0], y23 = C.yearly[1], y24 = C.yearly[2], y25 = C.yearly[3], y26e = C.yearly[4];
  const yoyMgmt = ((mgmt.total - j.mgmt_fee_25jan)/j.mgmt_fee_25jan*100).toFixed(1);
  const vsAvg = ((mgmt.total - j.mgmt_fee_25avg)/j.mgmt_fee_25avg*100).toFixed(1);
  const chans = ['retail','wholesale','digital','enterprise','iot','corporate_sales','small_biz'];
  const cN = {retail:'소매',wholesale:'도매',digital:'디지털',enterprise:'기업/공공',iot:'IoT',corporate_sales:'법인영업',small_biz:'소상공인'};

  // 채널별 관리수수료 비중 Top/Bottom
  const mgmtSorted = chans.map(c=>({ch:c,v:mgmt[c],share:(mgmt[c]/mgmt.total*100)})).sort((a,b)=>b.v-a.v);
  const topMgmt = mgmtSorted[0], botMgmt = mgmtSorted[mgmtSorted.length-1];

  // 연도별 성장률
  const g24 = ((y24.total-y23.total)/y23.total*100).toFixed(1);
  const g25 = ((y25.total-y24.total)/y24.total*100).toFixed(1);
  const g26e = ((y26e.total-y25.total)/y25.total*100).toFixed(1);

  // 월별 추세 (25년)
  const m25 = C.monthly25;
  const m25First = m25[0], m25Last = m25[m25.length-1];
  const mTrend = ((m25Last.total - m25First.total)/m25First.total*100).toFixed(1);

  let h = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>KT M&S 수수료 종합 보고서</title>${rptStyle()}</head><body>
  <div class="cover">
    <div class="cover-tag">Confidential · For Executive Review Only</div>
    <div class="cover-title">KT M&S 수수료수입 종합 분석 보고서</div>
    <div class="cover-sub">${D.period} 기준 | 경영기획팀 | 관리수수료·채널구조·추세·리스크 종합 분석</div>
    <div class="cover-meta">
      <div class="cover-kpi"><div class="cover-kpi-v">${total.total.toFixed(1)}<span style="font-size:10px">억</span></div><div class="cover-kpi-l">수수료수입 합계</div></div>
      <div class="cover-kpi"><div class="cover-kpi-v">${mgmt.total.toFixed(1)}<span style="font-size:10px">억</span></div><div class="cover-kpi-l">관리수수료</div></div>
      <div class="cover-kpi"><div class="cover-kpi-v">${yoyMgmt > 0 ? '+' : ''}${yoyMgmt}%</div><div class="cover-kpi-l">관리수수료 YoY</div></div>
      <div class="cover-kpi"><div class="cover-kpi-v">${y26e.total.toFixed(0)}<span style="font-size:10px">억</span></div><div class="cover-kpi-l">26년 전사 예측</div></div>
    </div>
  </div>
  <div class="page">

  <h1>Ⅰ. Executive Summary</h1>
  <div class="exec-box">
    <strong>수수료수입 현황:</strong> ${D.period} 수수료수입 합계 <strong>${total.total.toFixed(1)}억</strong>, 이 중 관리수수료 <strong>${mgmt.total.toFixed(1)}억</strong>(비중 ${(mgmt.total/total.total*100).toFixed(1)}%), 무선 정책수수료(KT-RDS+정책) <strong>${j.policy_fee.total.toFixed(1)}억</strong>(비중 ${(j.policy_fee.total/total.total*100).toFixed(1)}%)으로 마감하였습니다. 관리수수료는 전년 동월(${j.mgmt_fee_25jan}억) 대비 <strong>${yoyMgmt}%</strong>, 25년 월평균(${j.mgmt_fee_25avg}억) 대비 <strong>${vsAvg}%</strong> 수준입니다.<br><br>
    <strong>무선 정책수수료:</strong> KT-RDS 수수료 <strong>${j.policy_fee.rds}억</strong>과 수수료수입 정책 <strong>${j.policy_fee.policy}억</strong>을 합산한 무선 정책수수료는 총 <strong>${j.policy_fee.total}억</strong>으로, 소매(${j.policy_fee.retail}억, ${(j.policy_fee.retail/j.policy_fee.total*100).toFixed(0)}%)·디지털(${j.policy_fee.digital}억, ${(j.policy_fee.digital/j.policy_fee.total*100).toFixed(0)}%) 2개 채널이 전체의 ${((j.policy_fee.retail+j.policy_fee.digital)/j.policy_fee.total*100).toFixed(0)}%를 차지합니다.<br><br>
    <strong>채널 집중도:</strong> 관리수수료의 <strong>${(mgmt.retail/mgmt.total*100).toFixed(1)}%</strong>가 소매 채널에 집중되어 있으며, 도매(${(mgmt.wholesale/mgmt.total*100).toFixed(1)}%), 디지털(${(mgmt.digital/mgmt.total*100).toFixed(1)}%) 순입니다. 법인/IoT 등 B2B 채널 비중은 합산 ${((mgmt.enterprise+mgmt.iot+mgmt.corporate_sales)/mgmt.total*100).toFixed(1)}%로 여전히 낮습니다.<br><br>
    <strong>연간 전망:</strong> 전사 수수료수입은 22년 ${y22.total}억 → 24년 ${y24.total}억(정점) → 26년(E) ${y26e.total}억으로 <strong>감소 추세</strong>가 지속되며, 도매 채널의 급감(△${((y25.wholesale-y24.wholesale)/y24.wholesale*100).toFixed(0)}%)이 핵심 원인입니다.
  </div>

  <h1>Ⅱ. 수수료수입 핵심 지표</h1>
  <h2>2-1. 26년 1월 채널별 수수료수입 현황</h2>
  <table>
    <thead><tr><th>채널</th><th>수수료수입 합계</th><th>비중</th><th>관리수수료</th><th>비중</th><th>관리/합계 비율</th><th>평가</th></tr></thead>
    <tbody>`;

  chans.forEach(c=>{
    const t = total[c], m = mgmt[c];
    const tShare = (t/total.channel_subtotal*100).toFixed(1);
    const mShare = (m/mgmt.channel_subtotal*100).toFixed(1);
    const ratio = (m/t*100).toFixed(1);
    const badge = t >= 30 ? 'badge-r' : (t >= 10 ? 'badge-y' : (t >= 5 ? 'badge-o' : 'badge-red'));
    const grade = t >= 30 ? '◎ 주력' : (t >= 10 ? '○ 중형' : (t >= 5 ? '△ 소형' : '▽ 미소'));
    h += `<tr><td>${cN[c]}</td><td style="text-align:right;font-weight:600">${t.toFixed(1)}억</td><td style="text-align:right">${tShare}%</td>
      <td style="text-align:right">${m.toFixed(1)}억</td><td style="text-align:right">${mShare}%</td>
      <td style="text-align:right">${ratio}%</td>
      <td><span class="badge ${badge}">${grade}</span></td></tr>`;
  });

 h += `<tr style="border-top:1px dashed #cbd5e1;font-weight:600;color:#64748b"><td>채널소계</td><td style="text-align:right">${total.channel_subtotal.toFixed(1)}억</td><td style="text-align:right">100%</td>
    <td style="text-align:right">${mgmt.channel_subtotal.toFixed(1)}억</td><td style="text-align:right">100%</td>
    <td style="text-align:right">${(mgmt.channel_subtotal/total.channel_subtotal*100).toFixed(1)}%</td><td>-</td></tr>
  <tr class="total"><td>전사합계(B열)</td><td style="text-align:right;font-weight:700">${total.total.toFixed(1)}억</td><td colspan="4" style="text-align:center;font-size:9px;color:#94a3b8">전사공통비(${total.common_corp.toFixed(1)}) + 채널공통비(${total.common_ch.toFixed(1)}) 포함</td><td>-</td></tr>
  </tbody></table>`;

  h += `<div class="insight">📌 <strong>수수료수입 TOP 3:</strong> 소매(${total.retail.toFixed(1)}억, ${(total.retail/total.total*100).toFixed(1)}%) → 디지털(${total.digital.toFixed(1)}억) → 도매(${total.wholesale.toFixed(1)}억). 상위 3개 채널이 전체의 <strong>${((total.retail+total.digital+total.wholesale)/total.total*100).toFixed(1)}%</strong>를 차지합니다.</div>

  <h2>2-2. 수수료수입 세부 항목별 분해</h2>
  <table>
    <thead><tr><th>항목</th><th>합계</th><th>소매</th><th>도매</th><th>디지털</th><th>기업/공공</th><th>IoT</th><th>법인영업</th><th>소상공인</th></tr></thead>
    <tbody>`;

  j.fee_items.forEach(it=>{
    h += `<tr><td>${it.nm}</td><td style="text-align:right;font-weight:600">${it.total.toFixed(1)}</td>
      <td style="text-align:right">${it.retail.toFixed(1)}</td><td style="text-align:right">${it.wholesale.toFixed(1)}</td>
      <td style="text-align:right">${it.digital.toFixed(1)}</td><td style="text-align:right">${it.enterprise.toFixed(1)}</td>
      <td style="text-align:right">${it.iot.toFixed(1)}</td><td style="text-align:right">${it.corporate_sales.toFixed(1)}</td>
      <td style="text-align:right">${it.small_biz.toFixed(1)}</td></tr>`;
  });

  h += `<tr class="total"><td>합 계</td><td style="text-align:right;font-weight:700">${total.total.toFixed(1)}</td>
    <td style="text-align:right">${total.retail.toFixed(1)}</td><td style="text-align:right">${total.wholesale.toFixed(1)}</td>
    <td style="text-align:right">${total.digital.toFixed(1)}</td><td style="text-align:right">${total.enterprise.toFixed(1)}</td>
    <td style="text-align:right">${total.iot.toFixed(1)}</td><td style="text-align:right">${total.corporate_sales.toFixed(1)}</td>
    <td style="text-align:right">${total.small_biz.toFixed(1)}</td></tr>
  </tbody></table>
  <div class="insight">📌 <strong>KT(RDS)</strong> 수수료가 전체의 <strong>${(j.fee_items[0].total/total.total*100).toFixed(1)}%</strong>로 최대. 관리수수료(${mgmt.total.toFixed(1)}억)가 2위, 용역(${j.fee_items[2].total.toFixed(1)}억)이 3위입니다.</div>

  <h2>2-3. 무선 정책수수료 채널별 분해 (KT-RDS + 수수료수입 정책)</h2>
  <table>
    <thead><tr><th>채널</th><th>KT-RDS</th><th>수수료수입 정책</th><th>합산 합계</th><th>비중</th><th>특이사항</th></tr></thead>
    <tbody>`;
  chans.forEach(c=>{
    const rds = j.policy_fee[c]||0;
    const feeItem = j.fee_items.find(f=>f.nm==='KT(RDS)');
    const rdsOnly = feeItem ? (feeItem[c]||0) : 0;
    const policyOnly = j.fee_items.find(f=>f.nm==='정책');
    const polV = policyOnly ? (policyOnly[c]||0) : 0;
    const combined = rds;
    const share = j.policy_fee.total > 0 ? (combined/j.policy_fee.total*100).toFixed(1) : '0.0';
    const isNeg = polV < 0;
    const note = isNeg ? '⚠️ 정책 차감 구조' : (combined >= 50 ? '◎ 주력' : combined >= 10 ? '○ 중형' : '△ 소형');
    h += `<tr><td>${cN[c]}</td>
      <td style="text-align:right">${rdsOnly.toFixed(1)}</td>
      <td style="text-align:right;color:${isNeg?'#dc2626':'inherit'}">${polV.toFixed(1)}</td>
      <td style="text-align:right;font-weight:700">${combined.toFixed(1)}억</td>
      <td style="text-align:right">${share}%</td>
      <td><span class="badge ${combined>=50?'badge-r':combined>=10?'badge-y':'badge-o'}">${note}</span></td></tr>`;
  });
  h += `<tr class="total"><td>합 계</td>
    <td style="text-align:right">${j.policy_fee.rds.toFixed(1)}</td>
    <td style="text-align:right">${j.policy_fee.policy.toFixed(1)}</td>
    <td style="text-align:right;font-weight:700">${j.policy_fee.total.toFixed(1)}억</td>
    <td style="text-align:right">100%</td><td>-</td></tr>
  </tbody></table>
  <div class="insight">📌 <strong>소매(${j.policy_fee.retail}억, ${(j.policy_fee.retail/j.policy_fee.total*100).toFixed(0)}%) + 디지털(${j.policy_fee.digital}억, ${(j.policy_fee.digital/j.policy_fee.total*100).toFixed(0)}%)</strong>가 무선 정책수수료의 ${((j.policy_fee.retail+j.policy_fee.digital)/j.policy_fee.total*100).toFixed(0)}%를 차지합니다. 법인영업·기업/공공은 정책 차감(-) 구조로 수익성 관리 필요.</div>

  <h1>Ⅲ. 연도별·월별 추세 분석</h1>
  <h2>3-1. 전사 수수료수입 연도별 추이 (관리수수료 기준)</h2>
  <table>
    <thead><tr><th>연도</th><th>전사합계</th><th>소매</th><th>도매</th><th>디지털</th><th>B2B</th><th>소상공인</th><th>YoY</th></tr></thead>
    <tbody>`;

  C.yearly.forEach((y,i)=>{
    const prev = i > 0 ? C.yearly[i-1] : null;
    const yoy = prev ? ((y.total-prev.total)/prev.total*100).toFixed(1) : '-';
    const yoyStr = yoy === '-' ? '-' : (parseFloat(yoy) > 0 ? `<span class="pos">+${yoy}%</span>` : `<span class="neg">${yoy}%</span>`);
    h += `<tr><td>${y.yr}</td><td style="text-align:right;font-weight:600">${y.total.toFixed(1)}</td>
      <td style="text-align:right">${y.retail.toFixed(1)}</td><td style="text-align:right">${y.wholesale.toFixed(1)}</td>
      <td style="text-align:right">${y.digital.toFixed(1)}</td><td style="text-align:right">${y.b2b.toFixed(1)}</td>
      <td style="text-align:right">${y.small_biz.toFixed(1)}</td><td style="text-align:right">${yoyStr}</td></tr>`;
  });

  h += `</tbody></table>
  <div class="warn"⚠️ 24년(${y24.total}억)을 정점으로 하락 전환. 25년 △${Math.abs(parseFloat(g25))}%, 26년(E) △${Math.abs(parseFloat(g26e))}% 추가 감소 전망. 도매 채널이 26년(E) ${y26e.wholesale}억으로 22년(${y22.wholesale}억) 대비 △${((y22.wholesale-y26e.wholesale)/y22.wholesale*100).toFixed(0)}% 축소.</div>

  <h2>3-2. 25년 월별 관리수수료 추이</h2>
  <table>
    <thead><tr><th>월</th><th>합계</th><th>소매</th><th>도매</th><th>디지털</th><th>B2B</th><th>소상공인</th></tr></thead>
    <tbody>`;

  m25.forEach(m=>{
    h += `<tr><td>${m.mo}</td><td style="text-align:right;font-weight:600">${m.total.toFixed(2)}</td>
      <td style="text-align:right">${m.retail.toFixed(2)}</td><td style="text-align:right">${m.wholesale.toFixed(2)}</td>
      <td style="text-align:right">${m.digital.toFixed(2)}</td><td style="text-align:right">${m.b2b.toFixed(2)}</td>
      <td style="text-align:right">${m.small_biz.toFixed(2)}</td></tr>`;
  });

  h += `</tbody></table>
  <div class="insight">📌 25년 1→5월 관리수수료 <strong>${m25First.total.toFixed(2)}억 → ${m25Last.total.toFixed(2)}억</strong>(${mTrend}%). 도매의 지속 하락(${m25First.wholesale} → ${m25Last.wholesale})이 전체 감소를 견인합니다.</div>

  <h2>3-3. 채널 세부 추세 (25년 1~5월)</h2>
  <table>
    <thead><tr><th>채널·항목</th><th>1월</th><th>2월</th><th>3월</th><th>4월</th><th>5월</th><th>추세</th></tr></thead>
    <tbody>`;

  // 소매 세부
  const rd = C.retail_detail;
  h += `<tr><td><strong>소매-무선</strong></td>${rd.w25.map(v=>`<td style="text-align:right">${v.toFixed(2)}</td>`).join('')}
    <td>${rd.w25[4]<rd.w25[0]?'<span class="neg">↓ 하락</span>':'<span class="pos">↑ 상승</span>'}</td></tr>`;
  h += `<tr><td><strong>소매-유선</strong></td>${rd.l25.map(v=>`<td style="text-align:right">${v.toFixed(2)}</td>`).join('')}
    <td>${rd.l25[4]>rd.l25[0]?'<span class="pos">↑ 상승</span>':'<span class="neg">↓ 하락</span>'}</td></tr>`;

  // 도매 세부
  const wd = C.wholesale_detail;
  ['wireless','jumpup','wired','hc','winwin'].forEach(k=>{
    const nm = {wireless:'도매-무선',jumpup:'도매-점프업',wired:'도매-유선',hc:'도매-H.C',winwin:'도매-Win-Win'}[k];
    h += `<tr><td><strong>${nm}</strong></td>${wd[k].map(v=>`<td style="text-align:right">${v.toFixed(2)}</td>`).join('')}
      <td>${wd[k][4]<wd[k][0]?'<span class="neg">↓</span>':'<span class="pos">↑</span>'}</td></tr>`;
  });

  // 디지털 세부
  const dd = C.digital_detail;
  ['wireless','wired','hc'].forEach(k=>{
    const nm = {wireless:'디지털-무선',wired:'디지털-유선',hc:'디지털-H.C'}[k];
    h += `<tr><td><strong>${nm}</strong></td>${dd[k].map(v=>`<td style="text-align:right">${v.toFixed(2)}</td>`).join('')}
      <td>${dd[k][4]<dd[k][0]?'<span class="neg">↓</span>':'<span class="pos">↑</span>'}</td></tr>`;
  });

  h += `</tbody></table>
  <div class="warn">⚠️ <strong>도매 점프업:</strong> 1.26 → 0.86억으로 5개월간 △31.7% 급감. 도매 H.C도 4.14 → 3.52억으로 하락 중. 반면 <strong>Win-Win은 1.17 → 1.89억으로 유일하게 상승</strong>.</div>

  <h1>Ⅳ. 리스크 워치리스트</h1>
  <table>
    <thead><tr><th>등급</th><th>리스크 항목</th><th>현황</th><th>원인</th><th>대응 방향</th></tr></thead>
    <tbody>
      <tr><td><span class="badge badge-red">HIGH</span></td><td>도매 수수료 급감</td>
        <td>24년 ${y24.wholesale}억 → 26년(E) ${y26e.wholesale}억 (△${((y24.wholesale-y26e.wholesale)/y24.wholesale*100).toFixed(0)}%)</td>
        <td>점프업 축소, 무선 가입자 감소</td><td>Win-Win HC 대체 모니터링, 도매 구조 재편</td></tr>
      <tr><td><span class="badge badge-red">HIGH</span></td><td>소매 무선 수수료 하락</td>
        <td>25년 1→5월: ${rd.w25[0]} → ${rd.w25[4]}억 (△${((rd.w25[0]-rd.w25[4])/rd.w25[0]*100).toFixed(1)}%)</td>
        <td>단건 수수료율 하락, 시장 포화</td><td>고단가 요금제 전환, 유선 연계 확대</td></tr>
      <tr><td><span class="badge badge-o">MED</span></td><td>디지털 무선 수수료 감소</td>
        <td>디지털 무선: ${dd.wireless[0]} → ${dd.wireless[4]}억</td>
        <td>kt닷컴 100% 위탁 전환 후 감소</td><td>H.C 수익성 검증, 계약 조건 재검토</td></tr>
      <tr><td><span class="badge badge-o">MED</span></td><td>B2B 수수료 축소</td>
        <td>24년 ${y24.b2b}억 → 25년 ${y25.b2b}억 (△${((y24.b2b-y25.b2b)/y24.b2b*100).toFixed(0)}%)</td>
        <td>IoT/기업 환경 악화, 수수료율 하락</td><td>고수익 B2B 포트폴리오 재편</td></tr>
      <tr><td><span class="badge badge-r">LOW</span></td><td>소상공인 수수료 안정</td>
        <td>25년 월평균 ${(m25.reduce((a,m)=>a+m.small_biz,0)/m25.length).toFixed(2)}억 유지</td>
        <td>유선 가입자 유지, 소폭 감소세</td><td>프랜차이즈 채널 확대로 반등</td></tr>
    </tbody>
  </table>

  <h1>Ⅴ. 전략적 실행 권고</h1>
  <div style="line-height:1.9;font-size:11px">
    <strong>① [즉시] 도매 수수료 구조 점검 (1Q 내)</strong><br>
    &nbsp;&nbsp;점프업 축소로 26년 ${y26e.wholesale}억까지 감소 전망. Win-Win HC 계약 구조를 재설계하여 도매 채널 수수료 수입 ${y25.wholesale}억 이상 방어 필요. 도매강화팀 '인프라 포트폴리오 조정' 과제와 연계 추진.<br><br>
    <strong>② [단기] 소매 무선 수수료 방어 전략 (2Q)</strong><br>
    &nbsp;&nbsp;소매 무선 수수료 월별 감소세(${rd.w25[0]}→${rd.w25[4]}억) 가속 시 연간 6~10억 수익 감소. 고단가 요금제 전환 인센티브 강화, 유선 결합 수수료 크로스셀 프로그램 도입 검토.<br><br>
    <strong>③ [단기] 디지털 H.C 수수료 성장 전략 (2Q~3Q)</strong><br>
    &nbsp;&nbsp;kt닷컴 HC(${dd.hc[0]}→${dd.hc[4]}억)가 무선 감소를 일부 상쇄 중. 가입자 순증 목표 달성 여부가 핵심. 디지털강화팀 'AI 고객 상담 시스템' 과제와 시너지 기대.<br><br>
    <strong>④ [중기] B2B 수수료 수익원 다각화 (하반기)</strong><br>
    &nbsp;&nbsp;B2B 채널 수수료가 86.4억(24년)→54.1억(25년)으로 급감. IoT 원격관제 솔루션 확장, 기업영업 직접 수수료 확보, 소상공인 프랜차이즈 확대를 통한 다변화 필수.<br><br>
    <strong>⑤ [중기] 관리수수료 KPI 연계 최적화 (3Q~4Q)</strong><br>
    &nbsp;&nbsp;관리수수료 비중이 전체 수수료수입의 ${(mgmt.total/total.total*100).toFixed(1)}%. KPI 달성률과 관리수수료 연동 구조를 점검하여 실적 개선 → 수수료 증대 선순환 구축.
  </div>

  <hr class="section-divider">
  <div class="footnote">본 보고서는 BM별 이익계산서 및 예측_무선_관리수수료_추세 데이터를 기반으로 작성되었습니다.<br>26년(E) 수치는 예측값이며, 실제 실적과 차이가 있을 수 있습니다.</div>

  </div></body></html>`;

  openReport(h, '수수료수입 종합 분석 보고서');
}




function genManagementMeetingMode(){
  const E = buildBriefingEngine();
  const P = E.P || {};
  const K = Array.isArray(E.K) ? E.K : [];
  const S = E.sub || {};

  const rev = nvl(P?.revenue?.total);
  const op = nvl(P?.op?.total);
  const opm = ratio(op, rev);
  const avgKpi = K.length ? K.reduce((a,b)=>a+nvl(b.ts),0)/K.length : null;
  const bestHq = K.length ? K[0] : null;
  const lowHq = K.length ? K[K.length-1] : null;
  const taskTotal = Array.isArray(E.T) ? E.T.length : 0;
  const taskDone = nvl(E.done);
  const taskRate = taskTotal ? taskDone/taskTotal : null;

  const issues = [];
  const worst = E.worstProfit || {name:'-', opm:null};
  const mgmtYoYN = parseFloat(String(E.mgmtYoY||'').replace('%',''));
  const subMgmtChgN = parseFloat(E.sub?.mgmtChg);

  if(worst && Number.isFinite(worst.opm)){
    issues.push({
      title:`수익성 하위 채널: ${worst.name}`,
      fact:`영업이익률 ${pctStr(worst.opm)}로 전사 평균 ${pctStr(opm)} 하회`,
      cause:'판관비·수수료 구조 대비 매출 방어력 약화',
      action:'하위 채널 비용·수수료 구조 월간 점검 및 즉시 리밸런싱'
    });
  }

  if(Number.isFinite(subMgmtChgN)){
    issues.push({
      title:'관리수수료 대상 가입자 감소',
      fact:`연초 대비 ${subMgmtChgN>=0?'+':''}${subMgmtChgN.toFixed(1)}%`,
      cause:'CAPA 대비 이탈/만기 환수 영향 누적',
      action:'소매·도매 대상 유지/재가입 캠페인 우선 실행'
    });
  }

  if(Number.isFinite(mgmtYoYN)){
    issues.push({
      title:'관리수수료 전년 동월 변동',
      fact:`YoY ${mgmtYoYN>=0?'+':''}${mgmtYoYN.toFixed(1)}%`,
      cause:'수수료 대상 모수 및 채널 믹스 변동',
      action:'관리수수료 대상 모수 방어 KPI를 본부 단위로 연동'
    });
  }

  if(lowHq && bestHq){
    const gap = nvl(bestHq.ts)-nvl(lowHq.ts);
    issues.push({
      title:'본부 KPI 편차 확대',
      fact:`${bestHq.hq} ↔ ${lowHq.hq} 격차 ${gap.toFixed(1)}p`,
      cause:'핵심 지표 실행력·채널 운영 방식 편차',
      action:'하위 본부 집중 코칭 + 상위 본부 베스트프랙티스 전파'
    });
  }

  if(taskTotal){
    issues.push({
      title:'과제 실행 속도 관리 필요',
      fact:`완료율 ${pctStr(taskRate)} (${taskDone}/${taskTotal})`,
      cause:'진행중 과제의 일정·오너십 관리 미흡',
      action:'핵심 과제 주간 점검체계와 마감 책임 명확화'
    });
  }

  const top3 = issues.slice(0,3);
  while(top3.length<3){
    top3.push({
      title:'데이터 점검 필요',
      fact:'핵심 이슈 산출에 필요한 데이터 일부 부족',
      cause:'업로드 데이터 최신성/정합성 미확보',
      action:'기준월 데이터 업로드 후 재생성'
    });
  }

  const briefingLine = [
    `매출 ${fB(rev)}억 · 영업이익 ${fB(op)}억 · 영업이익률 ${pctStr(opm)}`,
    `평균 KPI ${avgKpi===null?'데이터 없음':avgKpi.toFixed(1)+'점'}${bestHq?` (상위 ${bestHq.hq} ${nvl(bestHq.ts).toFixed(1)}점)`:''}`,
    `무선 가입자 ${fmtSubNum(nvl(S.wirelessTot))}명 · CAPA ${fmtSubNum(nvl(S.capa))} · 관리수수료 ${Number.isFinite(E.mgmtCur)?E.mgmtCur.toFixed(1)+'억':'데이터 없음'}`,
    `과제 완료율 ${pctStr(taskRate)} · 지연위험 ${nvl(E.delayed?.length)}건`
  ];

  const meetingText = [
    `경영회의 시작 (${D.period||'-'})`,
    '',
    '1. 경영 브리핑',
    ...briefingLine,
    '',
    '2. 핵심 이슈 3개',
    ...top3.map((i,idx)=>`${idx+1}. ${i.title} - ${i.fact}`),
    '',
    '3. 원인 분석',
    ...top3.map((i,idx)=>`${idx+1}. ${i.title} / 원인: ${i.cause}`),
    '',
    '4. 액션 제안',
    ...top3.map((i,idx)=>`P${idx+1}. ${i.action}`)
  ].join('\n');

  const h = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>경영회의 시작 모드</title>${rptStyle()}<style>
    .meet-wrap{max-width:920px;margin:0 auto;padding:20px}
    .meet-head{background:linear-gradient(135deg,#0f172a,#1d4ed8);color:#fff;border-radius:16px;padding:20px 22px;margin-bottom:14px}
    .meet-head h1{color:#fff;border:none;margin:0 0 6px;font-size:28px;padding:0}
    .meet-head p{margin:0;font-size:12px;opacity:.9}
    .meet-tools{display:flex;gap:8px;flex-wrap:wrap;margin:10px 0 16px}
    .meet-btn{border:none;border-radius:10px;padding:8px 12px;font-size:12px;font-weight:800;cursor:pointer;color:#fff}
    .meet-btn.copy{background:#0ea5e9}.meet-btn.print{background:#2563eb}.meet-btn.pdf{background:#7c3aed}
    .meet-sec{background:#fff;border:1px solid #dbe4ee;border-radius:12px;padding:14px 16px;margin-bottom:10px}
    .meet-sec h2{margin:0 0 8px;padding:0;border:none;font-size:16px;color:#0f172a}
    .meet-sec ul,.meet-sec ol{margin:0;padding-left:18px}
    .meet-sec li{margin:4px 0;line-height:1.55}
    .meet-kv{font-size:12px;color:#334155}
    @media print {.meet-tools{display:none} body{background:#fff;padding:0} .meet-wrap{padding:0}}
  </style></head><body><div class="meet-wrap">
    <div class="meet-head">
      <h1>경영회의 시작</h1>
      <p>${D.period||'-'} 기준 · 손익·KPI·가입자·수수료·과제 통합 요약</p>
    </div>

    <div class="meet-tools">
      <button class="meet-btn copy" onclick="copyMeetingText()">📋 복사</button>
      <button class="meet-btn print" onclick="window.print()">🖨️ 인쇄</button>
      <button class="meet-btn pdf" onclick="window.print()">📄 PDF 내보내기</button>
    </div>

    <section class="meet-sec">
      <h2>1. 경영 브리핑</h2>
      <ul>${briefingLine.map(x=>`<li>${x}</li>`).join('')}</ul>
    </section>

    <section class="meet-sec">
      <h2>2. 핵심 이슈 3개</h2>
      <ol>${top3.map(i=>`<li><strong>${i.title}</strong> — ${i.fact}</li>`).join('')}</ol>
    </section>

    <section class="meet-sec">
      <h2>3. 원인 분석</h2>
      <ol>${top3.map(i=>`<li><strong>${i.title}</strong><div class="meet-kv">원인: ${i.cause}</div></li>`).join('')}</ol>
    </section>

    <section class="meet-sec">
      <h2>4. 액션 제안</h2>
      <ol>${top3.map((i,idx)=>`<li><strong>P${idx+1}.</strong> ${i.action}</li>`).join('')}</ol>
    </section>
  </div>
  <script>
    function copyMeetingText(){
      var text = ${JSON.stringify(meetingText)};
      if(navigator.clipboard&&navigator.clipboard.writeText){
        navigator.clipboard.writeText(text).then(function(){alert('회의안이 복사되었습니다.');});
      }else{
        var ta=document.createElement('textarea');ta.value=text;document.body.appendChild(ta);ta.select();document.execCommand('copy');document.body.removeChild(ta);alert('회의안이 복사되었습니다.');
      }
    }
  <\/script>
  </body></html>`;

  openReport(h,'경영회의 시작 모드');
}

function genMeetingReadyBriefing(){
  // 하위호환: 기존 버튼/호출 유지
  genManagementMeetingMode();
}

function genExecutiveManagementReport(){
  const E=buildBriefingEngine();
  const P=E.P||{}; const K=(E.K||[]).slice();
  const totalRev=(P.revenue||{}).total||0;
  const totalOp=(P.op||{}).total||0;
  const margin=ratio(totalOp,totalRev);
  const avgKpi=K.length?K.reduce((a,b)=>a+(b.ts||0),0)/K.length:0;
  const topHq=K.length?K.slice().sort((a,b)=>b.ts-a.ts)[0]:null;
  const lowHq=K.length?K.slice().sort((a,b)=>a.ts-b.ts)[0]:null;
  const riskList=(E.actions||[]).filter(a=>a.priority==='HIGH').slice(0,5);
  const taskTotal=(D.tasks||[]).length;
  const taskDone=(D.tasks||[]).filter(t=>t.st==='완료').length;
  const taskRate=taskTotal?taskDone/taskTotal:0;

  const channelRows=['retail','wholesale','digital','enterprise','iot','corporate_sales','small_biz']
    .map(k=>({key:k,nm:CN[k],rev:(P.revenue||{})[k]||0,op:(P.op||{})[k]||0,margin:ratio((P.op||{})[k]||0,(P.revenue||{})[k]||0)}))
    .sort((a,b)=>b.op-a.op);
  const bestCh=channelRows[0]||null;
  const worstCh=channelRows.slice().sort((a,b)=>a.margin-b.margin)[0]||null;

  const watchList=[];
  if(worstCh) watchList.push(`채널 수익성 하방: ${worstCh.nm} 이익률 ${pctStr(worstCh.margin)}로 구조개선 필요`);
  if(lowHq) watchList.push(`KPI 하위 본부: ${lowHq.hq} ${lowHq.ts.toFixed(1)}점 · 상위와 편차 관리 필요`);
  if(E.mgmtMix<0.25) watchList.push(`관리수수료 비중 ${pctStr(E.mgmtMix||0)}로 방어력 약화 가능성`);
  if(taskTotal && taskRate<0.5) watchList.push(`전사 과제 완료율 ${pctStr(taskRate)}로 실행속도 제고 필요`);

  const h=`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Executive Management Report</title>${rptStyle()}<style>
    .ex-wrap{display:flex;flex-direction:column;gap:16px}
    .ex-hero{padding:20px;border-radius:16px;background:linear-gradient(135deg,#020617,#0f172a 35%,#1d4ed8);color:#fff;position:relative;overflow:hidden}
    .ex-hero:after{content:'';position:absolute;right:-40px;top:-40px;width:180px;height:180px;border-radius:50%;background:radial-gradient(circle,rgba(125,211,252,.38),rgba(125,211,252,0) 70%)}
    .ex-hero h1{margin:0;font-size:30px;letter-spacing:-.4px}
    .ex-hero p{margin:10px 0 0;font-size:13px;opacity:.9}
    .ex-kicker{display:inline-flex;padding:4px 10px;border-radius:999px;font-size:10px;font-weight:800;background:rgba(255,255,255,.18);border:1px solid rgba(255,255,255,.3);margin-bottom:10px}

    .ex-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}
    .ex-card{border:1px solid #dbeafe;background:linear-gradient(180deg,#fff,#f8fafc);border-radius:12px;padding:12px}
    .ex-k{font-size:11px;color:#475569;font-weight:700}
    .ex-v{font-size:23px;font-weight:900;color:#0f172a;margin-top:4px}
    .ex-s{font-size:11px;color:#64748b;margin-top:4px}

    .ex-sec{border:1px solid #e2e8f0;border-radius:12px;padding:14px;background:#fff}
    .ex-sec h2{margin:0 0 10px;padding:0;border:none;font-size:16px;color:#0f172a}
    .ex-split{display:grid;grid-template-columns:1.3fr .7fr;gap:12px}
    .ex-note{border:1px dashed #cbd5e1;background:#f8fafc;border-radius:10px;padding:10px}
    .ex-note b{display:block;font-size:11px;color:#334155;margin-bottom:6px}
    .ex-note ul{margin:0;padding-left:16px;color:#475569;font-size:11px;line-height:1.6}

    .pill{display:inline-flex;padding:3px 8px;border-radius:999px;font-size:10px;font-weight:800}
    .pill.high{background:#fee2e2;color:#b91c1c}.pill.mid{background:#fef3c7;color:#92400e}.pill.low{background:#dcfce7;color:#166534}
    .rank-up{color:#059669;font-weight:800}.rank-down{color:#dc2626;font-weight:800}
    @media(max-width:850px){.ex-grid{grid-template-columns:repeat(2,1fr)}.ex-split{grid-template-columns:1fr}}
  </style></head><body><div class="wrap"><div class="ex-wrap">
    <div class="ex-hero"><span class="ex-kicker">EXECUTIVE MANAGEMENT REPORT</span><h1>경영진 통합 보고</h1><p>${D.period||'-'} · 손익·KPI·과제·수수료를 한 장에서 보는 의사결정용 리포트</p></div>

    <div class="ex-grid">
      <div class="ex-card"><div class="ex-k">총 매출</div><div class="ex-v">${fB(totalRev)}억</div><div class="ex-s">업로드 손익 데이터 기준</div></div>
      <div class="ex-card"><div class="ex-k">총 영업이익</div><div class="ex-v">${fB(totalOp)}억</div><div class="ex-s">영업이익률 ${pctStr(margin)}</div></div>
      <div class="ex-card"><div class="ex-k">KPI 평균</div><div class="ex-v">${avgKpi?avgKpi.toFixed(1):'-'}</div><div class="ex-s">Top ${topHq?topHq.hq+' '+topHq.ts.toFixed(1):'-'}</div></div>
      <div class="ex-card"><div class="ex-k">전사 과제 완료율</div><div class="ex-v">${pctStr(taskRate)}</div><div class="ex-s">완료 ${taskDone}/${taskTotal||0}</div></div>
    </div>

    <div class="ex-sec">
      <h2>1) Profitability & Focus Channel</h2>
      <div class="ex-split">
        <table><thead><tr><th>채널</th><th>매출(억)</th><th>영업이익(억)</th><th>이익률</th><th>진단</th></tr></thead><tbody>
        ${channelRows.map(c=>`<tr><td>${c.nm}</td><td>${fB(c.rev)}</td><td class="${c.op>=0?'rank-up':'rank-down'}">${fB(c.op)}</td><td>${pctStr(c.margin)}</td><td>${c.margin<0.03?'구조개선 필요':c.margin<0.06?'방어관리':'안정'}</td></tr>`).join('')}
        </tbody></table>
        <div class="ex-note">
          <b>채널 포커스</b>
          <ul>
            <li>상위 채널: ${bestCh?bestCh.nm+' · 이익 '+fB(bestCh.op)+'억':''}</li>
            <li>개선 우선: ${worstCh?worstCh.nm+' · 이익률 '+pctStr(worstCh.margin):''}</li>
            <li>관리수수료 비중: ${pctStr(E.mgmtMix||0)} (총 ${Number.isFinite(E.totalFee)?E.totalFee.toFixed(1):'N/A'}억)</li>
          </ul>
        </div>
      </div>
    </div>

    <div class="ex-sec">
      <h2>2) KPI & Execution Delta</h2>
      <table><thead><tr><th>순위</th><th>본부</th><th>점수</th><th>실행 우선도</th></tr></thead><tbody>
      ${K.slice().sort((a,b)=>b.ts-a.ts).map((x,i)=>`<tr><td>${i+1}</td><td>${x.hq}</td><td>${x.ts.toFixed(1)}</td><td>${i<2?'우수사례 확산':(i>=K.length-2?'집중 코칭':'편차 축소')}</td></tr>`).join('')}
      </tbody></table>
      <div class="warn" style="margin-top:10px">본부 간 점수 편차: ${topHq&&lowHq?(topHq.ts-lowHq.ts).toFixed(1)+'p':'-'} · 하위 본부는 2주 단위 개선지표(순증/수익성/핵심 KPI)로 코칭 운영</div>
    </div>

    <div class="ex-sec">
      <h2>3) Risk Watchlist & 2-Week Action Plan</h2>
      <table><thead><tr><th>우선순위</th><th>리스크/과제</th><th>담당</th><th>기한</th></tr></thead><tbody>
      ${(riskList.length?riskList:(E.actions||[]).slice(0,5)).map(a=>`<tr><td><span class="pill ${a.priority==='HIGH'?'high':a.priority==='MID'?'mid':'low'}">${a.priority||'MID'}</span></td><td>${a.item}</td><td>${a.owner}</td><td>${a.due}</td></tr>`).join('')}
      </tbody></table>
      <div class="ex-note" style="margin-top:10px"><b>Watchlist</b><ul>${(watchList.length?watchList:['핵심 리스크 없음']).map(x=>`<li>${x}</li>`).join('')}</ul></div>
    </div>
  </div></div></body></html>`;
  openReport(h,'Executive Management Report');
}

// ===== Dashboard Init =====
function initDashboard(){
  try{
    if(typeof D === 'undefined') return;
    console.log('[DASH] initDashboard ENTERED');

    const P = D.profit || {};
    const T = D.tasks || [];
    const K = D.kpi || [];
    let SUB = null;
    try { if(typeof subscriberData!=='undefined' && subscriberData && subscriberData.wireless) SUB = subscriberData; } catch(e){}
    if(!SUB){ try { if(typeof DEFAULT_SUBSCRIBER_DATA!=='undefined' && DEFAULT_SUBSCRIBER_DATA && DEFAULT_SUBSCRIBER_DATA.wireless) SUB = DEFAULT_SUBSCRIBER_DATA; } catch(e){} }

    const revenue = P.revenue ? P.revenue.total : 0;
    const op = P.op ? P.op.total : 0;
    const gross = P.gross ? P.gross.total : 0;
    const margin = revenue ? (op/revenue*100) : 0;
    const avgKpi = K.length ? K.reduce((s,x)=>s+(Number(x.ts)||0),0)/K.length : 0;
    const lowPPCount = T.filter(t=>t.pp<0.2 && t.st==='진행중').length;
    const worstCh = ['retail','wholesale','digital','enterprise'].map(k=>({k,m:(P.revenue&&P.revenue[k])?((P.op&&P.op[k]||0)/(P.revenue[k]||1)):-999})).sort((a,b)=>a.m-b.m)[0];
    const worstChName = worstCh ? CN[worstCh.k] : '-';

    const el = (id)=>document.getElementById(id);
    const fmt = (v) => (v/1e8).toFixed(1)+'억';

    // ▸ Metric Cards
    if(el('dashRevenue')) el('dashRevenue').textContent = fmt(revenue);
    if(el('dashRevenueSub')) el('dashRevenueSub').textContent = '매출총이익 '+fmt(gross)+' · 순이익 '+fmt(P.net_income||0);
    if(el('dashOp')) el('dashOp').textContent = fmt(op);
    if(el('dashOpSub')){
      const contri = P.contribution ? P.contribution.total : 0;
      el('dashOpSub').textContent = '공헌이익 '+fmt(contri);
    }
    if(el('dashMargin')) el('dashMargin').textContent = margin.toFixed(1)+'%';
    if(el('dashMarginSub')) el('dashMarginSub').textContent = '매출총이익률 '+(revenue?(gross/revenue*100).toFixed(1):0)+'%';
    if(el('dashKpi')) el('dashKpi').textContent = avgKpi.toFixed(1)+'점';
    if(el('dashKpiSub')){
      const best = K.length ? K.reduce((a,b)=>(Number(b.ts)||0)>(Number(a.ts)||0)?b:a) : null;
      el('dashKpiSub').textContent = best ? '1위 '+best.hq+' '+Number(best.ts).toFixed(1)+'점' : '-';
    }

    // ▸ Metric Bar Animations
    setTimeout(()=>{
      const setBar=(id,pct)=>{const b=el(id);if(b)b.style.width=Math.min(100,pct)+'%';};
      setBar('dashRevenueBar', 75);
      setBar('dashOpBar', op>0 ? Math.min((op/revenue)*100*10,80) : 20);
      setBar('dashMarginBar', margin*10);
      setBar('dashKpiBar', avgKpi);
      // Row 2 bars
      setBar('dashWirelessBar', 70);
      setBar('dashCapaBar', 55);
      setBar('dashSvcRevBar', 65);
      setBar('dashMgmtFeeBar', 60);
    },300);

    // ▸ Row 2: 무선가입자 · CAPA · 서비스매출 · 관리수수료
    console.log('[DASH] SUB:', !!SUB, SUB && SUB.wireless ? 'wireless OK total='+SUB.wireless.total : 'no wireless');
    console.log('[DASH] el check:', !!el('dashWireless'), !!el('dashSvcRev'), !!el('dashMgmtFee'));
    console.log('[DASH] rev_service:', !!P.rev_service, P.rev_service ? P.rev_service.total : 'N/A');
    console.log('[DASH] commission.jan26:', !!(D.commission && D.commission.jan26));
    try {
      if(SUB && SUB.wireless){
        const w = SUB.wireless;
        const wTotal = Number(w.total)||0;
        if(el('dashWireless')) el('dashWireless').textContent = wTotal ? (wTotal/10000).toFixed(1)+'만' : '-';
        if(el('dashWirelessSub')){
          const na = Number(w.netAdd)||0;
          const ch = Number(w.churn)||0;
          el('dashWirelessSub').textContent = '순증 '+(na>=0?'+':'')+na.toLocaleString()+' · 해지 '+ch.toLocaleString();
        }
        const wCapa = Number(w.capa)||0;
        if(el('dashCapa')) el('dashCapa').textContent = wCapa ? wCapa.toLocaleString() : '-';
        if(el('dashCapaSub')){
          const ser = w.series||{};
          const nArr = ser.newSubs||[];
          const cArr = ser.chgSubs||[];
          const newS = nArr.length ? Number(nArr[nArr.length-1])||0 : 0;
          const chgS = cArr.length ? Number(cArr[cArr.length-1])||0 : 0;
          el('dashCapaSub').textContent = '신규 '+newS.toLocaleString()+' · 기변 '+chgS.toLocaleString();
        }
      }
    } catch(e){ console.warn('dash row2 subscriber:',e); }
    console.log('[DASH] subscriber card done');

    try {
      if(el('dashSvcRev') && P.rev_service && P.rev_service.total){
        el('dashSvcRev').textContent = fmt(P.rev_service.total);
        if(el('dashSvcRevSub')){
          const prd = (P.rev_product && P.rev_product.total) ? P.rev_product.total : 0;
          const svcRatio = revenue ? (P.rev_service.total/revenue*100).toFixed(1) : '0';
          el('dashSvcRevSub').textContent = '서비스비중 '+svcRatio+'% · 상품 '+fmt(prd);
        }
      }
    } catch(e){ console.warn('dash row2 svc:',e); }
    console.log('[DASH] svc card done');

    try {
      if(el('dashMgmtFee')){
        const comm = D.commission||{};
        const j26 = comm.jan26||{};
        const mf = j26.mgmt_fee;
        const m25 = (comm.monthly25 && comm.monthly25.length) ? comm.monthly25[0] : null;
        if(mf && mf.total){
          el('dashMgmtFee').textContent = mf.total+'억';
          if(el('dashMgmtFeeSub')) el('dashMgmtFeeSub').textContent = '소매 '+(mf.retail||0)+'억 · 도매 '+(mf.wholesale||0)+'억';
        } else if(m25){
          el('dashMgmtFee').textContent = m25.total.toFixed(1)+'억';
          if(el('dashMgmtFeeSub')) el('dashMgmtFeeSub').textContent = '소매 '+m25.retail.toFixed(1)+'억 · 도매 '+m25.wholesale.toFixed(1)+'억';
        }
      }
    } catch(e){ console.warn('dash row2 fee:',e); }

    // ▸ Hero Chips & Pulse
    const done=T.filter(t=>t.st==='완료').length;
    const prog=T.filter(t=>t.st==='진행중').length;
    if(el('dashTaskChip')) el('dashTaskChip').textContent = '과제 '+done+'/'+T.length+' 완료';
    if(el('dashPulseText')){
      const worst = K.length ? K.reduce((a,b)=>(Number(b.ts)||0)<(Number(a.ts)||0)?b:a) : null;
      el('dashPulseText').innerHTML = worst
        ? '<strong>'+worst.hq+'</strong> KPI '+Number(worst.ts).toFixed(1)+'점<br>최하위 — 집중 점검 필요'
        : '데이터 없음';
    }

    // ▸ Channel P&L Snapshot
    const chGrid = el('dashChannelGrid');
    if(chGrid && P.revenue){
      const channels = [
        {key:'retail',nm:'소매'},
        {key:'wholesale',nm:'도매'},
        {key:'digital',nm:'디지털'},
        {key:'corporate_sales',nm:'법인영업'},
        {key:'small_biz',nm:'소상공인'},
        {key:'iot',nm:'IoT'},
        {key:'enterprise',nm:'기업/공공'}
      ];
      let html = channels.map(c=>{
        const rev = P.revenue[c.key]||0;
        const opVal = P.op[c.key]||0;
        const mg = rev ? (opVal/rev*100).toFixed(1) : '0.0';
        const isPos = opVal >= 0;
        return `<div class="ds-ch">
          <div class="ds-ch-name">${c.nm}</div>
          <div class="ds-ch-rev">${fmt(rev)}</div>
          <div class="ds-ch-op ${isPos?'pos':'neg'}">영업이익 ${isPos?'+':''}${fmt(opVal)}</div>
          <div class="ds-ch-margin">이익률 ${mg}%</div>
        </div>`;
      }).join('');
      // 유통플랫폼 사업단
      if(P.platform && P.platform.total){
        const pf = P.platform.total;
        const pfRev = pf.revenue||0;
        const pfOp = pf.op||0;
        const pfMg = pfRev ? (pfOp/pfRev*100).toFixed(1) : '0.0';
        const pfPos = pfOp >= 0;
        html += `<div class="ds-ch ds-ch-platform">
          <div class="ds-ch-name">유통플랫폼</div>
          <div class="ds-ch-rev">${fmt(pfRev)}</div>
          <div class="ds-ch-op ${pfPos?'pos':'neg'}">영업이익 ${pfPos?'+':''}${fmt(pfOp)}</div>
          <div class="ds-ch-margin">이익률 ${pfMg}%</div>
        </div>`;
      }
      chGrid.innerHTML = html;
    }

    // ▸ KPI Ranking
    const rankEl = el('dashKpiRank');
    if(rankEl && K.length){
      const sorted = [...K].sort((a,b)=>(Number(b.ts)||0)-(Number(a.ts)||0));
      const maxTs = Number(sorted[0].ts)||100;
      rankEl.innerHTML = sorted.map((k,i)=>{
        const ts = Number(k.ts)||0;
        const medal = i===0?'🥇':i===1?'🥈':i===2?'🥉':(i+1);
        const mClass = i<3?('g'+(i+1)):'';
        return `<div class="ds-rank-row">
          <div class="ds-rank-medal ${mClass}">${medal}</div>
          <div class="ds-rank-name">${k.hq}</div>
          <div class="ds-rank-bar"><div class="ds-rank-fill" style="width:${(ts/maxTs*100).toFixed(0)}%"></div></div>
          <div class="ds-rank-score">${ts.toFixed(1)}</div>
        </div>`;
      }).join('');
    }

    // ▸ Risk Watch
    const riskEl = el('dashRisk');
    if(riskEl){
      const risks = [];
      // Channel profit risk
      if(P.op){
        const channels = {retail:'소매',wholesale:'도매',digital:'디지털',corporate_sales:'법인영업',small_biz:'소상공인',iot:'IoT',enterprise:'기업/공공'};
        Object.entries(channels).forEach(([k,nm])=>{
          const opVal = P.op[k]||0;
          const rev = P.revenue[k]||0;
          const mg = rev?(opVal/rev*100):0;
          if(mg < 2 && rev > 0) risks.push({level:'high',text:`<strong>${nm}</strong> 영업이익률 ${mg.toFixed(1)}% — 수익성 구조개선 필요`});
        });
      }
      // KPI spread risk
      if(K.length>=2){
        const scores = K.map(k=>Number(k.ts)||0);
        const spread = Math.max(...scores)-Math.min(...scores);
        if(spread > 3){
          const worst = K.reduce((a,b)=>(Number(b.ts)||0)<(Number(a.ts)||0)?b:a);
          risks.push({level:'med',text:`본부간 KPI 편차 <strong>${spread.toFixed(1)}점</strong> — ${worst.hq} 집중 지원 권고`});
        }
      }
      // Task risk
      const delayed = T.filter(t=>t.st==='지연'||t.st==='미착수').length;
      if(delayed>0) risks.push({level:'high',text:`사업과제 <strong>${delayed}건</strong> 지연/미착수 — 즉시 점검 필요`});
      if(lowPPCount>5) risks.push({level:'med',text:`추진도 20% 미만 과제 <strong>${lowPPCount}건</strong> — 실행력 강화 필요`});
      // SGA risk
      if(P.sga && P.revenue){
        const sgaRatio = P.sga.total/P.revenue.total*100;
        if(sgaRatio > 20) risks.push({level:'med',text:`판관비율 <strong>${sgaRatio.toFixed(1)}%</strong> — 비용 효율화 모니터링`});
      }
      if(!risks.length) risks.push({level:'low',text:'현재 주요 리스크 없음'});
      riskEl.innerHTML = risks.slice(0,5).map(r=>
        `<div class="ds-risk-item">
          <span class="ds-risk-badge ${r.level}">${r.level}</span>
          <span class="ds-risk-text">${r.text}</span>
        </div>`
      ).join('');
    }

    // ▸ Task Progress
    const taskSum = el('dashTaskSummary');
    const taskBars = el('dashTaskBars');
    if(taskSum){
      const plan = T.filter(t=>t.st==='계획').length;
      const delay = T.filter(t=>t.st==='지연'||t.st==='미착수').length;
      const avgPP = T.length ? (T.reduce((s,t)=>s+(t.pp||0),0)/T.length*100).toFixed(0) : 0;
      taskSum.innerHTML = `
        <div class="ds-task-stat"><div class="ds-task-stat-val">${T.length}</div><div class="ds-task-stat-lbl">전체 과제</div></div>
        <div class="ds-task-stat"><div class="ds-task-stat-val" style="color:#34d399">${done}</div><div class="ds-task-stat-lbl">완료</div></div>
        <div class="ds-task-stat"><div class="ds-task-stat-val" style="color:#60a5fa">${prog}</div><div class="ds-task-stat-lbl">진행중</div></div>
        <div class="ds-task-stat"><div class="ds-task-stat-val" style="color:#fbbf24">${plan+delay}</div><div class="ds-task-stat-lbl">계획/지연</div></div>
      `;
    }
    if(taskBars){
      const hqs = {};
      T.forEach(t=>{
        const ch = t.ch||'기타';
        if(!hqs[ch]) hqs[ch]={total:0,sumPP:0};
        hqs[ch].total++;
        hqs[ch].sumPP += (t.pp||0);
      });
      const sorted = Object.entries(hqs).sort((a,b)=>b[1].total-a[1].total).slice(0,6);
      taskBars.innerHTML = sorted.map(([nm,v])=>{
        const pct = v.total ? (v.sumPP/v.total*100).toFixed(0) : 0;
        const color = pct>=60?'green':pct>=30?'blue':pct>=15?'yellow':'red';
        return `<div class="ds-task-bar-row">
          <div class="ds-task-bar-label">${nm}</div>
          <div class="ds-task-bar-track"><div class="ds-task-bar-fill ${color}" style="width:${pct}%"></div></div>
          <div class="ds-task-bar-pct">${pct}%</div>
        </div>`;
      }).join('');
    }


    const readyGrid = el('dashReadyGrid');
    const readyNote = el('dashReadyNote');
    if(readyGrid){
      const req=[
        {k:'task',n:'사업과제',required:true},
        {k:'profit',n:'BM 손익',required:true},
        {k:'kpi',n:'조직 KPI',required:true},
        {k:'hqprofit',n:'본부 손익',required:false},
        {k:'factbook',n:'Factbook',required:false},
        {k:'comm',n:'수수료',required:false},
        {k:'plan',n:'26년 경영계획',required:false},
        {k:'midterm',n:'중기/그룹계획',required:false}
      ];
      const statusMeta={ok:{t:'준비완료',c:'ok'},warn:{t:'주의',c:'warn'},err:{t:'오류',c:'err'},unknown:{t:'미업로드',c:'unknown'}};
      readyGrid.innerHTML=req.map(item=>{
        const raw=uploadHealth[item.k]||'unknown';
        const st=statusMeta[raw]||statusMeta.unknown;
        return `<div class="ds-ready-item"><span class="ds-ready-name">${item.n}${item.required?' *':''}</span><span class="ds-ready-pill ${st.c}">${st.t}</span></div>`;
      }).join('');
      const requiredOk=req.filter(x=>x.required).every(x=>uploadHealth[x.k]==='ok'||uploadHealth[x.k]==='warn');
      if(readyNote){
        readyNote.innerHTML = requiredOk
          ? '필수 데이터(사업과제/BM손익/KPI) 준비 완료 · 임원 보고서 생성 권장'
          : '필수 데이터(*) 업로드 후 생성하면 보고서 신뢰도가 높아집니다.';
      }
    }


    const keyEl = el('dashKeyMessages');
    if(keyEl){
      const pos = [`영업이익 ${fmt(op)} 유지`, `KPI 상위 본부: ${(K[0]&&K[0].hq)||'-'}`, `완료 과제 ${done||0}건`];
      const neg = [`수익성 하위: ${worstChName}`, `저추진 과제 ${lowPPCount||0}건`, `관리수수료 대상 감소 추세`];
      const urg = [`1주 내 하위채널 손익 개선안 제출`, `지연/미착수 과제 주간 점검`, `본부별 KPI 편차 축소 TF 가동`];
      keyEl.innerHTML = `
        <div class="ds-msg-card pos"><h4>Positive Highlights</h4><ul>${pos.map(x=>`<li>${x}</li>`).join('')}</ul></div>
        <div class="ds-msg-card neg"><h4>Negative Highlights</h4><ul>${neg.map(x=>`<li>${x}</li>`).join('')}</ul></div>
        <div class="ds-msg-card urg"><h4>Urgent Actions</h4><ul>${urg.map(x=>`<li>${x}</li>`).join('')}</ul></div>`;
    }

    const rankChEl = el('dashChannelRanking');
    if(rankChEl && P.op && P.revenue){
      const majors=['retail','wholesale','digital','enterprise'].map(k=>({k,n:CN[k],op:P.op[k]||0,rev:P.revenue[k]||0}));
      const sorted=majors.sort((a,b)=>(b.op/b.rev||-99)-(a.op/a.rev||-99));
      rankChEl.innerHTML = sorted.map((c,i)=>`<div class="ds-rank-row"><div class="ds-rank-medal">${i+1}</div><div class="ds-rank-name">${c.n}</div><div class="ds-rank-bar"><div class="ds-rank-fill" style="width:${Math.max(6,Math.min(100,((c.op/c.rev)||0)*600)).toFixed(0)}%"></div></div><div class="ds-rank-score">${((((c.op/c.rev)||0)*100)).toFixed(1)}%</div></div>`).join('');
    }

    const hm = el('dashKpiHeatmap');
    if(hm && K.length){
      const max=Math.max(...K.map(x=>x.ts||0));
      hm.innerHTML = K.map(k=>`<div class="mini-row"><span>${k.hq.replace('본부','')}</span><div class="mini-track"><div class="mini-fill" style="width:${((k.ts||0)/max*100).toFixed(0)}%"></div></div><b>${(k.ts||0).toFixed(1)}</b></div>`).join('');
    }

    const pcEl = el('dashProfitCompare');
    if(pcEl && P.revenue && P.op){
      const items=['retail','wholesale','digital','enterprise'].map(k=>({n:CN[k],r:P.revenue[k]||0,o:P.op[k]||0}));
      const mx=Math.max(...items.map(i=>i.r||0),1);
      pcEl.innerHTML=items.map(i=>`<div class="mini-row"><span>${i.n}</span><div class="mini-track"><div class="mini-fill alt" style="width:${(i.r/mx*100).toFixed(0)}%"></div></div><b>${fB(i.o)}억</b></div>`).join('');
    }

    const mtEl = el('dashMonthlyTrend');
    if(mtEl){
      const months=(window.monthlyProfitData||[]).slice(-6);
      if(months.length){
        const mx=Math.max(...months.map(m=>m.total||0),1);
        mtEl.innerHTML=months.map(m=>`<div class="mini-col"><i>${m.month}</i><div class="v" style="height:${Math.max(8,(m.total/mx*90)).toFixed(0)}px"></div><em>${fB(m.total)}</em></div>`).join('');
      } else {
        mtEl.innerHTML='<div class="mini-empty">누적 업로드 시 월별 추이 자동 생성</div>';
      }
    }
  }catch(e){
    console.error('dashboard init error',e);
  }
}

document.addEventListener('DOMContentLoaded',function(){
  setTimeout(initDashboard, 50);
});
