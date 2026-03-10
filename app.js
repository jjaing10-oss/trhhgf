
function initDashboard(){

const revenue = DATA.revenue
const op = DATA.op
const margin = ((op/revenue)*100).toFixed(1)

document.getElementById("rev").innerText = revenue + "억"
document.getElementById("op").innerText = op + "억"
document.getElementById("margin").innerText = margin + "%"
document.getElementById("kpi").innerText = DATA.kpi + "%"

document.getElementById("aiBrief").innerHTML = `
• 소매 채널 매출 증가<br>
• 디지털 채널 수익성 개선<br>
• 판관비 증가 모니터링 필요
`

document.getElementById("topIssue").innerHTML = `
<li>소매 KPI 편차 확대</li>
<li>도매 수익성 하락</li>
<li>판촉비 증가</li>
`
}

document.addEventListener("DOMContentLoaded", initDashboard)
