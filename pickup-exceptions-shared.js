// Pure shared policy: used by Apps Script, admin and LIFF. No network or storage.
function _rxVersion(){return 'pickup-exceptions-v1';}
function _rxRouteColor(route){
  var s=String(route||'').replace(/〔週五14:45加班〕$/,'').trim();
  if(s==='A車加')return '#BE185D';
  if(s==='小巴(加)'||s==='小巴（加）')return '#475569';
  if(/走路|步行/.test(s))return '#2E7D32';
  if(/^交通車A\b|^A車$/.test(s))return '#1D4ED8';
  if(/^交通車B\b|^B車$/.test(s))return '#7E22CE';
  if(/^交通車C\b|^C車$/.test(s))return '#B45309';
  if(s==='小巴')return '#0E7490';
  if(s==='半巴')return '#795548';
  return '#334155';
}
function _rxDate(date){
  if(!/^\d{4}-\d{2}-\d{2}$/.test(String(date)))throw new Error('日期格式須為 YYYY-MM-DD');
  var d=new Date(date+'T00:00:00Z');
  if(isNaN(d.getTime())||d.toISOString().slice(0,10)!==date)throw new Error('日期不存在');
  return d;
}
function _rxMatch(rule,date,session,studentId){
  _rxDate(date);
  if(rule.studentId!==studentId||date<rule.start||date>rule.end)return false;
  if(rule.createdAt&&date<rule.createdAt.slice(0,10))return false;
  if(rule.status==='cancelled'&&(!rule.updatedAt||date>=rule.updatedAt.slice(0,10)))return false;
  if(!['active','cancelled'].includes(rule.status))return false;
  return (rule.session==='all'||rule.session===session)&&rule.weekdays.includes(_rxDate(date).getUTCDay());
}
function _rxOverlap(a,b){
  if(a.studentId!==b.studentId||a.status!=='active'||b.status!=='active')return false;
  if(a.session!==b.session&&a.session!=='all'&&b.session!=='all')return false;
  var start=a.start>b.start?a.start:b.start,end=a.end<b.end?a.end:b.end;
  if(start>end)return false;
  var common=a.weekdays.filter(function(d){return b.weekdays.includes(d);});
  for(var d=_rxDate(start),i=0;i<7&&d.toISOString().slice(0,10)<=end;i++,d.setUTCDate(d.getUTCDate()+1)){
    if(common.includes(d.getUTCDay()))return true;
  }
  return false;
}
function _rxReason(rules,date,session,id){
  return rules.filter(function(r){return _rxMatch(r,date,session,id);}).map(function(r){return r.reason;}).join('；');
}
