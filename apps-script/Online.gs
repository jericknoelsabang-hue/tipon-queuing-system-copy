const ONLINE_FORM_ID = '1HDcvrT_4QUU-xDZdl_PSEzzRhR_-XTqqL4NsxSPf9Bw';
const ONLINE_HEADERS = ['ID','Token digest','Submitted at','Transaction','Name','Profile JSON','Stage','Appointment','Sequence','Queue number'];
const ONLINE_RECORDS_SHEET = 'ONLINE_RECORDS';
const ONLINE_DOCUMENTS_SHEET = 'ONLINE_DOCUMENTS';
const ONLINE_MESSAGES_SHEET = 'ONLINE_MESSAGES';

function initializeOnlineQueue() {
  const props = PropertiesService.getScriptProperties();
  const queue = ss_();
  if(queue.getId()!==CFG.spreadsheetId) throw Error('The TIPON TAWAG Copy database is unavailable.');

  ensureOnlineSheet_(queue,ONLINE_RECORDS_SHEET,ONLINE_HEADERS);
  ensureOnlineSheet_(queue,ONLINE_DOCUMENTS_SHEET,['Application','File ID','Filename']);
  ensureOnlineSheet_(queue,ONLINE_MESSAGES_SHEET,['Application','Author','Message','Sent at']);

  const answers=onlineAnswersSheet_(props);
  if (!props.getProperty('ONLINE_FOLDER')) props.setProperty('ONLINE_FOLDER', DriveApp.createFolder('TIPON TAWAG Private Pre-Requirements').getId());
  const folder=DriveApp.getFolderById(props.getProperty('ONLINE_FOLDER'));
  const access=folder.getSharingAccess();
  if(access===DriveApp.Access.ANYONE||access===DriveApp.Access.ANYONE_WITH_LINK)
    throw Error('The applicant upload folder must use restricted sharing.');
  if (!props.getProperty('ONLINE_ONSITE_KEY')) props.setProperty('ONLINE_ONSITE_KEY', onlineToken_());

  const migrated=migrateLegacyOnlineStore_(props,queue);
  const fields=onlineFields_();
  SpreadsheetApp.flush();
  return {ok:true,spreadsheetId:queue.getId(),records:queue.getSheetByName(ONLINE_RECORDS_SHEET).getLastRow()-1,
    answersSheet:answers.getName(),formFields:fields.filter(field=>field.id).length,legacyRecordsMigrated:migrated};
}

function ensureOnlineSheet_(book,name,headers) {
  let sheet=book.getSheetByName(name);
  if(!sheet) sheet=book.insertSheet(name);
  if(sheet.getLastRow()===0) {
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
  } else {
    const width=Math.max(sheet.getLastColumn(),headers.length);
    const current=sheet.getRange(1,1,1,width).getDisplayValues()[0];
    const matches=headers.every((header,index)=>String(current[index]||'').trim()===header);
    if(!matches) throw Error('Unexpected header row in '+name+'. Existing data was not changed.');
  }
  sheet.setFrozenRows(1);
  return sheet;
}

function migrateLegacyOnlineStore_(props,book) {
  const legacyId=String(props.getProperty('ONLINE_STORE')||'').trim();
  if(!legacyId||legacyId===CFG.spreadsheetId||props.getProperty('ONLINE_STORE_MIGRATED')==='YES') return 0;
  const legacy=SpreadsheetApp.openById(legacyId);
  const mappings=[
    ['Applications',ONLINE_RECORDS_SHEET],
    ['Documents',ONLINE_DOCUMENTS_SHEET],
    ['Messages',ONLINE_MESSAGES_SHEET]
  ];
  let migrated=0;
  for(const [sourceName,targetName] of mappings) {
    const source=legacy.getSheetByName(sourceName),target=book.getSheetByName(targetName);
    if(!source||!target||source.getLastRow()<2) continue;
    const existing=new Set(onlineRead_(target).map(row=>String(row[0]||'')));
    for(const row of onlineRead_(source)) {
      const id=String(row[0]||'');
      if(!id||existing.has(id)) continue;
      target.appendRow(row.map(value=>String(value??'')));
      existing.add(id);
      migrated++;
    }
  }
  props.setProperty('ONLINE_STORE_MIGRATED','YES');
  props.setProperty('ONLINE_STORE_MIGRATED_AT',new Date().toISOString());
  return migrated;
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
  if(view==='staff') throw Error('The recruiter Online Queue interface is hosted on GitHub Pages.');
  if (!['applicant','onsite'].includes(view)) throw Error('Unknown online queue page.');
  const ticket=String(params.case||'');
  const onsite=String(params.site||'');
  const page=HtmlService.createTemplateFromFile('Online');
  page.initial=JSON.stringify({view,token:/^[a-f0-9]{64}$/.test(ticket)?ticket:'',site:/^[a-f0-9]{64}$/.test(onsite)?onsite:''});
  return page.evaluate().setTitle('TIPON TAWAG · Online Queue');
}

function onlineQueue(action, input) {
  const d=input||{};
  if (action==='fields') return onlineFields_();
  if(action==='chat'||action==='staffChat') throw Error('Applicant chat and replies are disabled.');
  const staffActions=['list','complete','readStaff','onsiteLink','download','sendMessage'];
  const staff=staffActions.includes(action)?onlineStaff_():'';
  const mutations=['submit','upload','complete','book','arrive','confirmArrival','sendMessage'];
  const lock=mutations.includes(action)?LockService.getScriptLock():null;
  if(lock) lock.waitLock(20000);
  try { return onlineOperate_(action,d,staff); }
  finally { if(lock) { try { SpreadsheetApp.flush(); } finally { lock.releaseLock(); } } }
}

function onlineOperate_(action,d,staff) {
  const props=PropertiesService.getScriptProperties(), storeId=props.getProperty('ONLINE_STORE');
  if(!storeId) throw Error('Online Queue has not been configured.');
  const data=ss_();
  const apps=data.getSheetByName(ONLINE_RECORDS_SHEET);
  const files=data.getSheetByName(ONLINE_DOCUMENTS_SHEET);
  const messages=data.getSheetByName(ONLINE_MESSAGES_SHEET);
  if(!apps||!files||!messages) throw Error('Run initializeOnlineQueue() before using Online Recruitment.');
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
    const existing=all.find(row=>row[1]===digest);
    if(existing){
      saveOnlineApplicantAnswers_(props,existing);
      return {url:url+'?online=applicant&case='+d.token};
    }
    const profile={};
    const copiedFields=onlineFields_().filter(field=>field.id);
    const firstAnswer=title=>onlineText_(d.answers?.[copiedFields.find(field=>field.title===title)?.id],250);
    const name=[firstAnswer('First Name'),firstAnswer('Last Name')].filter(Boolean).join(' ');
    if(!firstAnswer('First Name')||!firstAnswer('Last Name')) throw Error('First and last name are required.');
    const same=copiedFields.find(field=>field.title.startsWith('Is your permanent address'));
    const skipPermanent=same && /^yes\b/i.test(String(d.answers?.[same.id]||''));
    const emailField=copiedFields.find(field=>field.title==='Email / Gmail Address');
    const email=onlineText_(d.answers?.[emailField?.id],254);
    if(!/^\S+@\S+\.\S+$/.test(email)) throw Error('Enter a valid Email / Gmail Address.');
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
    saveOnlineApplicantAnswers_(props,[record,digest,now.toISOString(),d.transaction,name,JSON.stringify(profile)]);
    return {url:url+'?online=applicant&case='+d.token};
  }
  const idx=staff?all.findIndex(row=>row[0]===d.id):(/^[a-f0-9]{64}$/.test(String(d.token||''))?all.findIndex(row=>row[1]===onlineDigest_(d.token)):-1);
  if(idx<0) throw Error('Application link unavailable.');
  const row=all[idx], id=row[0], stage=row[6];
  const update=(next,appointment=row[7],sequence=row[8],number=row[9])=>{
    apps.getRange(idx+2,7,1,4).setValues([[next,appointment,sequence,number].map(value=>"'"+String(value))]);
    const updated=apps.getRange(idx+2,1,1,ONLINE_HEADERS.length).getValues()[0];
    saveOnlineApplicantAnswers_(props,updated);
    return updated;
  };
  if(action==='sendMessage'){
    if(!staff)throw Error('Only authorized recruiters can send applicant messages.');
    const message=onlineText_(d.message,2000);
    if(!message)throw Error('Enter a message before sending.');
    onlineWrite_(messages,[id,staff,message,now.toISOString()]);
    return {ok:true};
  }
  if(action==='download'){
    const file=fileList(id).find(doc=>doc[1]===d.file);
    if(!file)throw Error('Document unavailable.');
    const blob=DriveApp.getFileById(file[1]).getBlob();
    return {name:file[2],mime:blob.getContentType(),data:Utilities.base64Encode(blob.getBytes())};
  }
  if(action==='readStaff'||action==='read'){
    const notices=onlineRead_(messages).filter(message=>String(message[0])===String(id)).map(message=>({author:'Recruiter',text:message[2],at:message[3]}));
    return {name:row[4],stage,appointment:row[7],number:row[9],url,today,
      answers:staff?JSON.parse(row[5]):null,files:fileList(id).map(f=>({name:f[2],id:staff?f[1]:''})),messages:notices};
  }
  if(action==='upload'){
    if(stage!=='Review') throw Error('Documents can be added while requirements are under review.');
    const incoming=Array.isArray(d.files)?d.files:[];
    if(!incoming.length||fileList(id).length+incoming.length>10) throw Error('Upload up to 10 documents total.');
    for(const file of onlineFiles_(props,incoming)) onlineWrite_(files,[id,file.getId(),file.getName()]);
    saveOnlineApplicantAnswers_(props,apps.getRange(idx+2,1,1,ONLINE_HEADERS.length).getValues()[0]);
  } else if(action==='complete'){
    if(stage==='Review')update('Ready to book');
  } else if(action==='book'){
    if(stage==='Booked'||stage==='Arrived')return {number:row[9],appointment:row[7]};
    if(stage!=='Ready to book')throw Error('Wait for the recruiter to complete document review.');
    const appointment=onlineText_(d.appointment,16);
    const date=new Date(appointment+':00+08:00');
    if(!/^\d{4}-\d\d-\d\dT\d\d:\d\d$/.test(appointment)||!Number.isFinite(date.getTime())||date<=now||
        Utilities.formatDate(date,'Asia/Manila',"yyyy-MM-dd'T'HH:mm")!==appointment) throw Error('Choose a future appointment in Philippine time.');
    const sequence=all.reduce((max,app)=>Math.max(max,Number(app[8])||0),0)+1;
    const number='OL'+row[3]+'-'+String(sequence).padStart(6,'0');
    update('Booked',appointment,sequence,number);
    return {number,appointment};
  } else if(action==='arrive'||action==='confirmArrival'){
    if(stage==='Arrived')return {number:row[9]};
    if(stage!=='Booked'||row[7].slice(0,10)!==today)throw Error('Check in on your scheduled appointment date.');
    if(action==='arrive'&&(!props.getProperty('ONLINE_ONSITE_KEY')||d.site!==props.getProperty('ONLINE_ONSITE_KEY')))throw Error('Scan the office QR to confirm arrival.');
    const ss=ss_(), active=ss.getSheetByName(CFG.queue), tx=findTx_(ss,row[3]);
    if(!active)throw Error('The TIPON TAWAG Copy queue sheet is unavailable.');
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
  const html=`<!doctype html><meta charset="utf-8"><title>TIPON Online Queue data connection</title><script>
    const channel=${JSON.stringify(channel)}, origin='https://jericknoelsabang-hue.github.io';
    const allowed=${JSON.stringify(staff?['list','complete','readStaff','onsiteLink','download','sendMessage']:['fields','submit','read','upload','book','arrive','confirmArrival'])};
    const pages=window.top!==window?window.top:window.opener;
    if(pages){
      window.addEventListener('message',event=>{
        const data=event.data;
        if(event.origin!==origin||event.source!==pages||data?.type!=='tipon-online-request'||data.channel!==channel||!allowed.includes(data.action))return;
        const reply=result=>pages.postMessage({type:'tipon-online-result',channel,id:data.id,...result},origin);
        google.script.run.withSuccessHandler(result=>reply({result})).withFailureHandler(error=>reply({error:error.message})).onlineQueue(data.action,data.input);
      });
      pages.postMessage({type:'tipon-online-ready',channel},origin);
    }
  </script>`;
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onlineAnswersSheet_(props) {
  const id=CFG.spreadsheetId;
  const access=DriveApp.getFileById(id).getSharingAccess();
  if(access===DriveApp.Access.ANYONE||access===DriveApp.Access.ANYONE_WITH_LINK)
    throw Error('The applicant database must use restricted sharing.');
  const book=ss_();
  return book.getSheetByName('ONLINE_APPLICATIONS')||book.insertSheet('ONLINE_APPLICATIONS');
}

function saveOnlineApplicantAnswers_(props, row, destination) {
  const sheet=destination||onlineAnswersSheet_(props);
  const profile=JSON.parse(row[5]||'{}');
  const documentsSheet=ss_().getSheetByName(ONLINE_DOCUMENTS_SHEET);
  const documents=documentsSheet?onlineRead_(documentsSheet).filter(file=>String(file[0])===String(row[0])):[];
  const metadata={
    'Application ID':row[0],
    'Submitted at (UTC)':row[2],
    'Recruitment Transaction':row[3]==='NH'?'New Hire':'Rehire',
    'Applicant Name':row[4],
    'Application Status':row[6]||'',
    'Appointment (PHT)':row[7]||'',
    'Queue Number':row[9]||'',
    'Uploaded File Names':documents.map(file=>file[2]).join('\\n'),
    'Uploaded File Links':documents.map(file=>'https://drive.google.com/file/d/'+file[1]+'/view').join('\\n')
  };
  const answers={...metadata,...profile};
  const previous=sheet.getLastRow()?sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0]:[];
  const headers=[...new Set([...previous,...Object.keys(answers)])];
  if(headers.length>sheet.getMaxColumns())sheet.insertColumnsAfter(sheet.getMaxColumns(),headers.length-sheet.getMaxColumns());
  if(headers.length!==previous.length){
    sheet.getRange(1,1,1,headers.length).setValues([headers.map(value=>"'"+value)]).setFontWeight('bold').setWrap(true);
    sheet.setFrozenRows(1);
  }
  const match=sheet.getLastRow()>1?sheet.getRange(2,1,sheet.getLastRow()-1,1).createTextFinder(row[0]).matchEntireCell(true).findNext():null;
  const values=headers.map(key=>{
    const value=answers[key];
    return value==null||value===''?'':"'"+(Array.isArray(value)?value.join('\n'):String(value));
  });
  if(match)sheet.getRange(match.getRow(),1,1,values.length).setValues([values]);
  else sheet.appendRow(values);
}

function backfillOnlineApplicantAnswers() {
  const lock=LockService.getScriptLock();
  lock.waitLock(20000);
  try{
    onlineStaff_();
    const props=PropertiesService.getScriptProperties();
    const source=ss_().getSheetByName(ONLINE_RECORDS_SHEET);
    if(!source) throw Error('Run initializeOnlineQueue() before backfilling applicant answers.');
    const destination=onlineAnswersSheet_(props);
    const applications=onlineRead_(source);
    for(const row of applications)saveOnlineApplicantAnswers_(props,row,destination);
    return {applications:applications.length};
  }finally{try{SpreadsheetApp.flush();}finally{lock.releaseLock();}}
}
