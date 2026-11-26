const ONLINE_FORM_ID = '1HDcvrT_4QUU-xDZdl_PSEzzRhR_-XTqqL4NsxSPf9Bw';
const ONLINE_HEADERS = ['ID','Token digest','Submitted at','Transaction','Name','Profile JSON','Stage','Appointment','Sequence','Queue number'];

function initializeOnlineQueue() {
  const props = PropertiesService.getScriptProperties();
  const queueId=props.getProperty('QUEUE_SPREADSHEET_ID');
  if(!queueId||queueId===CFG.spreadsheetId) throw Error('Configure a separate QUEUE_SPREADSHEET_ID before initializing the online queue.');
  if (!props.getProperty('ONLINE_STORE')) {
    const store = SpreadsheetApp.create('TIPON TAWAG Online Recruitment');
    store.getSheets()[0].setName('Applications').appendRow(ONLINE_HEADERS);
    store.insertSheet('Messages').appendRow(['Application','Author','Message','Sent at']);
    store.insertSheet('Documents').appendRow(['Application','File ID','Filename']);
    props.setProperty('ONLINE_STORE', store.getId());
  }
  if (!props.getProperty('ONLINE_FOLDER')) props.setProperty('ONLINE_FOLDER', DriveApp.createFolder('TIPON TAWAG Private Pre-Requirements').getId());
  if (!props.getProperty('ONLINE_ONSITE_KEY')) props.setProperty('ONLINE_ONSITE_KEY', onlineToken_());
  onlineFields_();
}

function onlineToken_() {
  return (Utilities.getUuid()+Utilities.getUuid()).replace(/-/g,'');
}

function onlineDigest_(token) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token).map(byte=>(byte&255).toString(16).padStart(2,'0')).join('');
}

function onlineStaff_() {
  const email = Session.getActiveUser().getEmail().toLowerCase().trim();
  const roster = String(PropertiesService.getScriptProperties().getProperty('ONLINE_RECRUITERS')||'').split(',').map(entry=>entry.trim().toLowerCase());
  if (!email || !roster.includes(email)) throw Error('Sign in with an authorized recruiter Google account.');
  return email;
}

function onlineFields_() {
  const data = FormApp.openById(ONLINE_FORM_ID).getItems();
  const converters = {
    TEXT:'asTextItem',PARAGRAPH_TEXT:'asParagraphTextItem',DATE:'asDateItem',DATETIME:'asDateTimeItem',TIME:'asTimeItem',
    LIST:'asListItem',MULTIPLE_CHOICE:'asMultipleChoiceItem',CHECKBOX:'asCheckboxItem'
  };
  return data.map(item=>{
    const type=String(item.getType());
    if (['SECTION_HEADER','PAGE_BREAK'].includes(type)) return {type,title:item.getTitle()};
    if (['IMAGE','VIDEO'].includes(type)) return null;
    if (!converters[type]) throw Error('Unsupported AP CONS question type: '+type);
    const question=item[converters[type]]();
    return {id:String(item.getId()),type,title:item.getTitle(),help:question.getHelpText?question.getHelpText():'',
      required:question.isRequired(),choices:question.getChoices?question.getChoices().map(choice=>choice.getValue()):[],
      other:question.hasOtherOption?question.hasOtherOption():false};
  }).filter(Boolean);
}

function onlineRead_(sheet) {
  return sheet.getDataRange().getDisplayValues().slice(1);
}

function onlineText_(value,limit) {
  return typeof value==='string'?value.trim().slice(0,limit):'';
}

function onlineWrite_(sheet, values) {
  sheet.appendRow(values.map(value=>"'"+String(value ?? '')));
}

function onlineQueueView_(params) {
  const view=String(params.online||'applicant');
  if(view==='bridge'||view==='staff-bridge') return onlineBridge_(String(params.channel||''),view==='staff-bridge');
  if (!['applicant','staff','onsite'].includes(view)) throw Error('Unknown online queue page.');
  const ticket=String(params.case||'');
  const onsite=String(params.site||'');
  const page=HtmlService.createTemplateFromFile('Online');
  page.initial=JSON.stringify({view,token:/^[a-f0-9]{64}$/.test(ticket)?ticket:'',site:/^[a-f0-9]{64}$/.test(onsite)?onsite:''});
  return page.evaluate().setTitle('TIPON TAWAG · Online Queue');
}

function onlineQueue(action, input) {
  const d=input||{};
  if (action==='fields') return onlineFields_();
  const staffActions=['list','complete','staffChat','readStaff','onsiteLink','download'];
  const staff=staffActions.includes(action)?onlineStaff_():'';
  const mutations=['submit','upload','chat','complete','staffChat','book','arrive'];
  const lock=mutations.includes(action)?LockService.getScriptLock():null;
  if(lock) lock.waitLock(20000);
  try { return onlineOperate_(action,d,staff); }
  finally { if(lock) { try { SpreadsheetApp.flush(); } finally { lock.releaseLock(); } } }
}

function onlineOperate_(action,d,staff) {
  const props=PropertiesService.getScriptProperties(), storeId=props.getProperty('ONLINE_STORE');
  if(!storeId) throw Error('Online Queue has not been configured.');
  const data=SpreadsheetApp.openById(storeId), apps=data.getSheetByName('Applications');
  const messages=data.getSheetByName('Messages'), files=data.getSheetByName('Documents');
  const all=onlineRead_(apps), url=props.getProperty('ONLINE_PUBLIC_URL')||ScriptApp.getService().getUrl();
  const now=new Date(), today=day_();
  if(action==='onsiteLink') return url+'?online=onsite&site='+props.getProperty('ONLINE_ONSITE_KEY');
  const fileList=id=>onlineRead_(files).filter(file=>file[0]===id);
  if(action==='list') {
    const docs=onlineRead_(files);
    return all.slice().reverse().map(row=>({id:row[0],date:row[2],transaction:row[3],name:row[4],stage:row[6],
      appointment:row[7],number:row[9],files:docs.filter(file=>file[0]===row[0]).map(file=>({name:file[2],id:file[1]}))}));
  }
  if(action==='submit'){
    if(!['NH','RH'].includes(d.transaction)) throw Error('Choose New Hire or Rehire.');
    if(!/^[a-f0-9]{64}$/.test(String(d.token||''))) throw Error('Invalid application session.');
    const digest=onlineDigest_(d.token);
    if(all.some(row=>row[1]===digest)) return {url:url+'?online=applicant&case='+d.token};
    const profile={};
    const copiedFields=onlineFields_().filter(field=>field.id);
    const firstAnswer=title=>onlineText_(d.answers?.[copiedFields.find(field=>field.title===title)?.id],250);
    const name=[firstAnswer('First Name'),firstAnswer('Last Name')].filter(Boolean).join(' ');
    if(!firstAnswer('First Name')||!firstAnswer('Last Name')) throw Error('First and last name are required.');
    const same=copiedFields.find(field=>field.title.startsWith('Is your permanent address'));
    const skipPermanent=same && /^yes\b/i.test(String(d.answers?.[same.id]||''));
    const email=onlineText_(d.email,254);
    if(!/^\S+@\S+\.\S+$/.test(email)) throw Error('Enter a valid email address.');
    profile.Email=email;
    for(const q of copiedFields){
      if(skipPermanent&&q.title.startsWith('Permanent '))continue;
      const raw=d.answers?.[q.id];
      const value=Array.isArray(raw)?raw.map(answer=>onlineText_(answer,500)).filter(Boolean):onlineText_(raw,1500);
      const selections=Array.isArray(value)?value:value?[value]:[];
      if(q.required&&!selections.length) throw Error('Required: '+q.title);
      if(q.choices.length&&!q.other&&selections.some(answer=>!q.choices.includes(answer))) throw Error('Invalid answer: '+q.title);
      profile[q.title]=value;
    }
    const record=Utilities.getUuid();
    if(!Array.isArray(d.files)||d.files.length<1||d.files.length>5) throw Error('Upload one to five PDF or JPEG documents.');
    const saved=onlineFiles_(props,d.files);
    try {
      for(const file of saved) onlineWrite_(files,[record,file.getId(),file.getName()]);
      onlineWrite_(apps,[record,digest,now.toISOString(),d.transaction,name,JSON.stringify(profile),'Review','','','']);
    } catch(error) { for(const file of saved) file.setTrashed(true); throw error; }
    return {url:url+'?online=applicant&case='+d.token};
  }
  const idx=staff?all.findIndex(row=>row[0]===d.id):(/^[a-f0-9]{64}$/.test(String(d.token||''))?all.findIndex(row=>row[1]===onlineDigest_(d.token)):-1);
  if(idx<0) throw Error('Application link unavailable.');
  const row=all[idx], id=row[0], stage=row[6];
  const update=(next,appointment=row[7],sequence=row[8],number=row[9])=>
    apps.getRange(idx+2,7,1,4).setValues([[next,appointment,sequence,number].map(value=>"'"+String(value))]);
  if(action==='download'){
    const file=fileList(id).find(doc=>doc[1]===d.file);
    if(!file)throw Error('Document unavailable.');
    const blob=DriveApp.getFileById(file[1]).getBlob();
    return {name:file[2],mime:blob.getContentType(),data:Utilities.base64Encode(blob.getBytes())};
  }
  if(action==='readStaff'||action==='read'){
    return {name:row[4],stage,appointment:row[7],number:row[9],url, today,
      answers:staff?JSON.parse(row[5]):null,files:fileList(id).map(f=>({name:f[2],id:staff?f[1]:''})),
      messages:onlineRead_(messages).filter(m=>m[0]===id).map(m=>({author:m[1],text:m[2],at:m[3]}))};
  }
  if(action==='chat'||action==='staffChat'){
    const message=onlineText_(d.message,2000);
    if(!message) throw Error('Enter a message.');
    onlineWrite_(messages,[id,staff?'Recruiter':'Applicant',message,now.toISOString()]);
  } else if(action==='upload'){
    if(stage!=='Review') throw Error('Documents can be added while requirements are under review.');
    const incoming=Array.isArray(d.files)?d.files:[];
    if(!incoming.length||fileList(id).length+incoming.length>10) throw Error('Upload up to 10 documents total.');
    for(const file of onlineFiles_(props,incoming)) onlineWrite_(files,[id,file.getId(),file.getName()]);
  } else if(action==='complete'){
    if(stage==='Review')update('Ready to book');
  } else if(action==='book'){
    if(stage==='Booked'||stage==='Arrived')return {number:row[9]};
    if(stage!=='Ready to book')throw Error('Wait for the recruiter to complete document review.');
    const appointment=onlineText_(d.appointment,16);
    const date=new Date(appointment+':00+08:00');
    if(!/^\d{4}-\d\d-\d\dT\d\d:\d\d$/.test(appointment)||!Number.isFinite(date.getTime())||date<=now||
        Utilities.formatDate(date,'Asia/Manila',"yyyy-MM-dd'T'HH:mm")!==appointment) throw Error('Choose a future appointment in Philippine time.');
    const sequence=all.reduce((max,app)=>Math.max(max,Number(app[8])||0),0)+1;
    const number='OL'+row[3]+'-'+String(sequence).padStart(6,'0');
    update('Booked',appointment,sequence,number);
    return {number};
  } else if(action==='arrive'){
    if(stage==='Arrived')return {number:row[9]};
    if(stage!=='Booked'||row[7].slice(0,10)!==today||!props.getProperty('ONLINE_ONSITE_KEY')||d.site!==props.getProperty('ONLINE_ONSITE_KEY'))
      throw Error('Scan the office QR on your scheduled date.');
    const queueId=props.getProperty('QUEUE_SPREADSHEET_ID');
    if(!queueId||queueId===CFG.spreadsheetId)throw Error('The duplicate office queue is unavailable.');
    const ss=ss_(), active=ss.getSheetByName(CFG.queue), tx=findTx_(ss,row[3]);
    if(!tx)throw Error('This transaction is not active in the office queue.');
    if(!active.getRange(1,1,active.getLastRow(),1).createTextFinder(id).matchEntireCell(true).findNext()){
      const entry=[id,today,now,row[9],row[3],tx.name,"'"+row[4],"Online recruitment",'WAITING','','','','','','','','','','',''];
      active.appendRow(entry);
    }
    update('Arrived');
    CacheService.getScriptCache().removeAll(['TIPON_STATE_display','TIPON_STATE_staff','TIPON_STATE_setup']);
    return {number:row[9]};
  } else throw Error('Unknown online queue action.');
  return {ok:true};
}

function onlineFiles_(props, input) {
  let total=0;
  const blobs=input.map(file=>{
    const name=onlineText_(file?.name,150), encoded=String(file?.data||'');
    if(!name||encoded.length>7000000)throw Error('Each file must be at most 5 MB.');
    const bytes=Utilities.base64Decode(encoded);
    total+=bytes.length;
    const signature=bytes.slice(0,5).map(b=>(b&255).toString(16).padStart(2,'0')).join('');
    const mime=signature.startsWith('255044462d')?'application/pdf':signature.startsWith('ffd8ff')?'image/jpeg':'';
    if(!mime||bytes.length>5*1024*1024||total>10*1024*1024)throw Error('Upload PDF or JPEG files, up to 5 MB each and 10 MB total.');
    return Utilities.newBlob(bytes,mime,name);
  });
  const folder=DriveApp.getFolderById(props.getProperty('ONLINE_FOLDER')), saved=[];
  try { for(const blob of blobs)saved.push(folder.createFile(blob)); return saved; }
  catch(error){ for(const file of saved)file.setTrashed(true); throw error; }
}

function onlineBridge_(channel, staff) {
  if(!/^[a-f0-9-]{36}$/.test(channel))throw Error('Invalid Online Queue connection.');
  const html=`<!doctype html><meta charset="utf-8"><title>TIPON Online Queue connection</title><script>
    const channel=${JSON.stringify(channel)}, origin='https://jericknoelsabang-hue.github.io';
    const allowed=${JSON.stringify(staff?['list','complete','staffChat','readStaff','onsiteLink','download']:['fields','submit','read','upload','chat','book','arrive'])};
    if(window.top!==window){
      window.addEventListener('message',event=>{
        const data=event.data;
        if(event.origin!==origin||event.source!==window.top||data?.type!=='tipon-online-request'||data.channel!==channel||!allowed.includes(data.action))return;
        const reply=result=>window.top.postMessage({type:'tipon-online-result',channel,id:data.id,...result},origin);
        google.script.run.withSuccessHandler(result=>reply({result})).withFailureHandler(error=>reply({error:error.message})).onlineQueue(data.action,data.input);
      });
      window.top.postMessage({type:'tipon-online-ready',channel},origin);
    }
  </script>`;
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
