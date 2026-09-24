// Shared pure rules. No I/O, no activation side effects.
// Official calendars: https://www.dgpa.gov.tw/information?uid=41&pid=12573
// https://www.dgpa.gov.tw/information?uid=30&pid=12982
function _rpSpec() {
  return {
    version:'fixed-pickup-v1',effective:'2026-09-29',through:'2027-01-20',
    extraSuffix:'〔週五14:45加班〕',
    routes:[
      {id:'A',name:'交通車A (RDW-5365)',capacity:6,vehicle:'A6'},
      {id:'B',name:'交通車B (RGE-2523)',capacity:4,vehicle:'B4'},
      {id:'C',name:'交通車C (RGE-2522)',capacity:4,vehicle:'C4'},
      {id:'BUS',name:'小巴',capacity:8,vehicle:'小巴8'},
      {id:'WALK',name:'走路接送',capacity:null,vehicle:'步行'},
      {id:'HALF',name:'半巴',capacity:4,vehicle:'半巴4'}
    ],
    holidays:{
      '2026-09-25':'中秋節','2026-09-28':'教師節',
      '2026-10-09':'國慶日補假','2026-10-10':'國慶日',
      '2026-10-25':'臺灣光復暨金門古寧頭大捷紀念日','2026-10-26':'光復節補假',
      '2026-12-25':'行憲紀念日','2027-01-01':'元旦'
    }
  };
}
function _rpDate(date) {
  date=String(date||'');
  if(!/^\d{4}-\d{2}-\d{2}$/.test(date))throw new Error('日期格式須為 YYYY-MM-DD');
  var d=new Date(date+'T00:00:00Z');
  if(isNaN(d.getTime())||d.toISOString().slice(0,10)!==date)throw new Error('日期不存在');
  return d;
}
function _rpApplies(date) { _rpDate(date);return date>=_rpSpec().effective; }
function _rpDay(date) {
  var d=_rpDate(date),p=_rpSpec(),dow=d.getUTCDay();
  if(date<p.effective)return {enforced:false,date:date,dow:dow};
  if(date>p.through)return {enforced:true,closed:true,reason:'超出已核定學期／假日日曆範圍'};
  if(p.holidays[date])return {enforced:true,closed:true,reason:p.holidays[date],dow:dow};
  if(dow===0||dow===6)return {enforced:true,closed:true,reason:'週末不排車',dow:dow};
  return {enforced:true,closed:false,dow:dow,noon:dow!==2,pm:dow!==3,extra:dow===5};
}
function _rpRule(date,route) {
  var day=_rpDay(date),p=_rpSpec();
  if(!day.enforced)return {enforced:false};
  if(day.closed)return {enforced:true,closed:true,reason:day.reason};
  var extra=String(route).endsWith(p.extraSuffix),base=extra?String(route).slice(0,-p.extraSuffix.length):route;
  var vehicle=p.routes.filter(function(r){return r.name===base;})[0];
  if(!vehicle)return {enforced:true,closed:true,reason:'非固定六條路線；A車加不可當成半巴'};
  if(extra&&!day.extra)return {enforced:true,closed:true,reason:'14:45 加班只允許週五'};
  var noon=!extra&&day.noon,pm=extra||day.pm;
  return {enforced:true,closed:false,capacity:vehicle.capacity,id:vehicle.id,extra:extra,
    noon:noon,pm:pm,noon_time:noon?'12:25':'',pm_time:pm?(extra?'14:45':'15:25'):'',
    noon_vehicle:noon?vehicle.vehicle:'',pm_vehicle:pm?vehicle.vehicle:'',weekday:['週日','週一','週二','週三','週四','週五','週六'][day.dow]};
}
function _rpTemplates(date) {
  var day=_rpDay(date),p=_rpSpec();
  if(!day.enforced)throw new Error('固定範本自 '+p.effective+' 起生效');
  if(day.closed)return [];
  var names=p.routes.map(function(r){return r.name;});
  if(day.extra)names=names.concat(p.routes.map(function(r){return r.name+p.extraSuffix;}));
  return names.map(function(name){
    var r=_rpRule(date,name);
    return {date:date,weekday:r.weekday,status:'上',route:name,capacity:r.capacity==null?'':String(r.capacity),
      noon_vehicle:r.noon_vehicle,noon_driver:'',noon_driver2:'',noon_school:'',noon_count:'',noon_time:r.noon_time,
      pm_vehicle:r.pm_vehicle,pm_driver:'',pm_driver2:'',pm_school:'',pm_count:'',pm_time:r.pm_time,remark:'',
      noon_returned_at:'',pm_returned_at:'',noon_transferred_from:'',pm_transferred_from:''};
  });
}
function _rpRowErrors(row,ignoreCounts) {
  var rule=_rpRule(String(row.date||''),String(row.route||''));
  if(!rule.enforced)return [];
  if(row.status==='休')return []; // inactive rows and names are retained, not operated.
  if(rule.closed)return [rule.reason];
  var errors=[];
  if(String(row.capacity||'')!==(rule.capacity==null?'':String(rule.capacity)))errors.push('固定乘車上限不符');
  ['noon_time','pm_time','noon_vehicle','pm_vehicle','weekday'].forEach(function(k){
    if(String(row[k]||'')!==rule[k])errors.push(k+' 必須為 '+(rule[k]||'空白'));
  });
  ['noon','pm'].forEach(function(session){
    var count=String(row[session+'_count']||'');
    if(count && (!/^\d+$/.test(count)||(!ignoreCounts&&rule.capacity!=null&&Number(count)>rule.capacity)))errors.push(session+' 人數超過上限或格式不正確');
    if(!rule[session])['driver','driver2','school','count','returned_at','transferred_from'].forEach(function(k){
      var value=String(row[session+'_'+k]||'');
      if(value && !(k==='count'&&value==='0'))errors.push(session+' 為不開行班次，不可排定');
    });
  });
  if(!['上','休'].includes(String(row.status)))errors.push('狀態須為上或休');
  return errors;
}
