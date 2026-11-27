const ONLINE_FORM_ID = '1HDcvrT_4QUU-xDZdl_PSEzzRhR_-XTqqL4NsxSPf9Bw';
const ONLINE_HEADERS = ['ID','Token digest','Submitted at','Transaction','Name','Profile JSON','Stage','Appointment','Sequence','Queue number'];
const ONLINE_RECORDS_SHEET = 'ONLINE_RECORDS';
const ONLINE_DOCUMENTS_SHEET = 'ONLINE_DOCUMENTS';
const ONLINE_MESSAGES_SHEET = 'ONLINE_MESSAGES';
const ONLINE_PDF_TEMPLATES_SHEET = 'PDF_TEMPLATES';
const ONLINE_PDF_LOG_SHEET = 'PDF_GENERATION';
const ONLINE_PDF_ROOT_FOLDER_ID = '1Gma29KL9HnJFScY7nf9u5O4eoSP9EPH3';
const ONLINE_PDF_TEMPLATE_HEADERS = ['Code','Template Type','Template File ID','Template Name','Destination Folder ID','Active','Created At','Naming Convention'];
const ONLINE_PDF_LOG_HEADERS = ['Ticket ID','Queue Number','Applicant Reference','Template Code','Template Type','PDF File ID','PDF URL','Generated At','Status'];

function initializeOnlineQueue() {
  const props = PropertiesService.getScriptProperties();
  const queue = ss_();
  if(queue.getId()!==CFG.spreadsheetId) throw Error('The TIPON TAWAG Copy database is unavailable.');

  ensureOnlineSheet_(queue,ONLINE_RECORDS_SHEET,ONLINE_HEADERS);
  ensureOnlineSheet_(queue,ONLINE_DOCUMENTS_SHEET,['Application','File ID','Filename']);
  ensureOnlineSheet_(queue,ONLINE_MESSAGES_SHEET,['Application','Author','Message','Sent at']);
  ensureOnlineSheet_(queue,ONLINE_PDF_TEMPLATES_SHEET,ONLINE_PDF_TEMPLATE_HEADERS);
  ensureOnlineSheet_(queue,ONLINE_PDF_LOG_SHEET,ONLINE_PDF_LOG_HEADERS);

  const answers=onlineAnswersSheet_(props);
  if (!props.getProperty('ONLINE_FOLDER')) props.setProperty('ONLINE_FOLDER', DriveApp.createFolder('TIPON TAWAG Private Pre-Requirements').getId());
  const folder=DriveApp.getFolderById(props.getProperty('ONLINE_FOLDER'));
  const access=folder.getSharingAccess();
  if(access===DriveApp.Access.ANYONE||access===DriveApp.Access.ANYONE_WITH_LINK)
    throw Error('The applicant upload folder must use restricted sharing.');
  const pdfFolder=DriveApp.getFolderById(ONLINE_PDF_ROOT_FOLDER_ID);
  if([DriveApp.Access.ANYONE,DriveApp.Access.ANYONE_WITH_LINK].includes(pdfFolder.getSharingAccess()))throw Error('The generated PDF folder must use restricted sharing.');
  if (!props.getProperty('ONLINE_ONSITE_KEY')) props.setProperty('ONLINE_ONSITE_KEY', onlineToken_());

  const migrated=migrateLegacyOnlineStore_(props,queue);
  const fields=onlineFields_();
  // Rebuild the applicant-facing answer table from the authoritative records.
  // This makes first-time setup and recovery from an empty answer sheet idempotent.
  const source=queue.getSheetByName(ONLINE_RECORDS_SHEET);
  const records=onlineRead_(source);
  for(const row of records)saveOnlineApplicantAnswers_(props,row,answers);
  SpreadsheetApp.flush();
  return {ok:true,spreadsheetId:queue.getId(),records:records.length,answersRows:records.length,
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
    const existing=current.slice(0,Math.max(0,sheet.getLastColumn())).map(value=>String(value||'').trim());
    const missing=headers.filter(header=>!existing.includes(header));
    const compatible=existing.every((header,index)=>!header||headers[index]===header);
    if(!compatible) throw Error('Unexpected header row in '+name+'. Existing data was not changed.');
    if(missing.length){
      const start=sheet.getLastColumn()+1;
      if(sheet.getMaxColumns()<start+missing.length-1)sheet.insertColumnsAfter(sheet.getMaxColumns(),start+missing.length-1-sheet.getMaxColumns());
      sheet.getRange(1,start,1,missing.length).setValues([missing]);
    }
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
    const existing=new Map();
    for(const row of onlineRead_(target)){
      const key=sourceName==='Applications'?String(row[0]||''):JSON.stringify(row.map(String));
      existing.set(key,(existing.get(key)||0)+1);
    }
    const occurrences=new Map();
    for(const row of onlineRead_(source)) {
      const id=String(row[0]||'');
      if(!id)continue;
      const key=sourceName==='Applications'?id:JSON.stringify(row.map(String));
      const occurrence=(occurrences.get(key)||0)+1;
      occurrences.set(key,occurrence);
      if(occurrence<=(existing.get(key)||0))continue;
      onlineWrite_(target,row);
      migrated++;
    }
  }
  props.setProperty('ONLINE_STORE_MIGRATED','YES');
  props.setProperty('ONLINE_STORE_MIGRATED_AT',new Date().toISOString());
  return migrated;
}

function onlineActiveTiponTicket_(d) {
  const id=String(d.id||'').trim(),window=Number(d.window);
  if(!/^[a-f0-9-]{36}$/i.test(id)||!Number.isInteger(window)||window<1)throw Error('Select the active queue ticket first.');
  const book=ss_(),queue=book.getSheetByName(CFG.queue);
  if(!queue||queue.getLastRow()<2)throw Error('The active queue ticket is unavailable.');
  const match=queue.getRange(2,1,queue.getLastRow()-1,1).createTextFinder(id).matchEntireCell(true).findNext();
  const row=match?queue.getRange(match.getRow(),1,1,QH.length).getDisplayValues()[0]:null;
  if(!row||String(row[8]).toUpperCase()!=='SERVING'||Number(row[9])!==window)throw Error('Tipon Form is available only for the ticket currently serving at this window.');
  const transaction=String(row[4]||'').toUpperCase();
  if(!['NH','RH','TR'].includes(transaction))throw Error('Tipon Form is only available for New Hire, Rehire, and Transfer.');
  const assignment=windowTx_(book,window);
  if(!assignmentAllows_(assignment,transaction))throw Error('This window is not assigned to the ticket transaction.');
  return {id,window,row,transaction,queueNo:String(row[3]||'')};
}

function onlineApplicationSheetRow_(sheet,id) {
  if(!sheet||sheet.getLastRow()<2)return null;
  const headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0].map(value=>String(value||'').trim());
  const idColumn=headers.indexOf('Application ID');
  if(idColumn<0)return null;
  const match=sheet.getRange(2,idColumn+1,sheet.getLastRow()-1,1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return !match?null:{headers,row:sheet.getRange(match.getRow(),1,1,sheet.getLastColumn()).getDisplayValues()[0],sheetRow:match.getRow()};
}

function onlineTiponProfile_(ticket,apps,answersSheet) {
  const found=apps&&apps.getLastRow()>1?apps.getRange(2,1,apps.getLastRow()-1,1).createTextFinder(ticket.id).matchEntireCell(true).findNext():null;
  const recordRow=found?found.getRow():0,record=recordRow?apps.getRange(recordRow,1,1,ONLINE_HEADERS.length).getDisplayValues()[0]:null;
  const existingIndex=record&&String(record[6])!=='Deleted'?recordRow-2:-1;
  if(existingIndex>=0){
    const profile=JSON.parse(record[5]||'{}'),saved=onlineApplicationSheetRow_(answersSheet,ticket.id);
    const metadata=new Set(['Application ID','Submitted at (UTC)','Recruitment Transaction','Applicant Name','Application Status','Appointment (PHT)','Queue Number','Uploaded File Names','Uploaded File Links','PDF Generation Status','Generated PDF URL','Generated PDF At','Generated PDF Template Code']);
    if(saved)for(let i=0;i<saved.headers.length;i++)if(!metadata.has(saved.headers[i]))profile[saved.headers[i]]=String(saved.row[i]||'');
    return {profile,recordIndex:existingIndex,recordRow,record};
  }
  const saved=onlineApplicationSheetRow_(answersSheet,ticket.id);
  const profile={};
  const metadata=new Set(['Application ID','Submitted at (UTC)','Recruitment Transaction','Applicant Name','Application Status','Appointment (PHT)','Queue Number','Uploaded File Names','Uploaded File Links','PDF Generation Status','Generated PDF URL','Generated PDF At','Generated PDF Template Code']);
  if(saved)for(let i=0;i<saved.headers.length;i++){
    const key=saved.headers[i];
    if(!metadata.has(key))profile[key]=String(saved.row[i]||'');
  }
  const bits=String(ticket.row[6]||'').trim().split(/\s+/);
  if(!profile['First Name'])profile['First Name']=bits.length>1?bits.slice(0,-1).join(' '):bits[0]||'';
  if(!profile['Last Name'])profile['Last Name']=bits.length>1?bits.at(-1):'';
  return {profile,recordIndex:-1,recordRow:0,record:null};
}

function onlineGetTiponForm_(d,apps,answersSheet) {
  onlineStaff_();
  const ticket=onlineActiveTiponTicket_(d);
  const stored=onlineTiponProfile_(ticket,apps,answersSheet),profile=stored.profile;
  const fields=onlineFields_();
  return {id:ticket.id,window:ticket.window,transaction:ticket.transaction,queueNo:ticket.queueNo,
    name:String(ticket.row[6]||''),purpose:String(ticket.row[7]||''),answers:profile,fields,
    recordExists:stored.recordIndex>=0};
}

function onlineTiponFormStatus_(d,apps,answersSheet) {
  onlineStaff_();
  const ticket=onlineActiveTiponTicket_(d),stored=onlineTiponProfile_(ticket,apps,answersSheet);
  return {id:ticket.id,saved:Boolean(stored.profile['Photo URL']&&stored.profile['Signature URL']),recordExists:stored.recordIndex>=0};
}

function onlineSaveCapture_(props,ticketId,kind,data) {
  const value=String(data||'');
  if(!value)return '';
  const match=value.match(/^data:(image\/(?:jpeg|png));base64,([A-Za-z0-9+/=]+)$/);
  if(!match)throw Error('Photo and signature must be JPEG or PNG images.');
  const bytes=Utilities.base64Decode(match[2]),limit=kind==='photo'?5*1024*1024:1*1024*1024;
  if(!bytes.length||bytes.length>limit)throw Error(kind==='photo'?'Photo must be 5 MB or smaller.':'Signature image must be 1 MB or smaller.');
  const folderId=props.getProperty('ONLINE_FOLDER');
  if(!folderId)throw Error('Online Queue storage has not been initialized.');
  const folder=DriveApp.getFolderById(folderId);
  const extension=match[1]==='image/png'?'png':'jpg';
  const filename='TIPON-'+kind.toUpperCase()+'-'+ticketId+'.'+extension;
  const file=folder.createFile(Utilities.newBlob(bytes,match[1],filename));
  return 'https://drive.google.com/file/d/'+file.getId()+'/view';
}

function onlineUpsertTiponAnswers_(sheet,ticket,profile,name,stage) {
  const saved=onlineApplicationSheetRow_(sheet,ticket.id);
  const metadata={
    'Application ID':ticket.id,
    'Submitted at (UTC)':saved?.row[saved.headers.indexOf('Submitted at (UTC)')]||new Date().toISOString(),
    'Recruitment Transaction':ticket.transaction==='NH'?'New Hire':ticket.transaction==='RH'?'Rehire':'Transfer',
    'Applicant Name':name,
    'Application Status':stage||'Tipon Form Saved',
    'Appointment (PHT)':saved?.row[saved.headers.indexOf('Appointment (PHT)')]||'',
    'Queue Number':ticket.queueNo,
    'Uploaded File Names':saved?.row[saved.headers.indexOf('Uploaded File Names')]||'',
    'Uploaded File Links':saved?.row[saved.headers.indexOf('Uploaded File Links')]||'',
    'PDF Generation Status':saved?.row[saved.headers.indexOf('PDF Generation Status')]||'',
    'Generated PDF URL':saved?.row[saved.headers.indexOf('Generated PDF URL')]||'',
    'Generated PDF At':saved?.row[saved.headers.indexOf('Generated PDF At')]||'',
    'Generated PDF Template Code':saved?.row[saved.headers.indexOf('Generated PDF Template Code')]||''
  };
  const values={...metadata,...profile},previous=sheet.getLastRow()?sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0].map(value=>String(value||'').trim()):[];
  const headers=[...new Set([...previous,...Object.keys(values)])];
  if(headers.length>sheet.getMaxColumns())sheet.insertColumnsAfter(sheet.getMaxColumns(),headers.length-sheet.getMaxColumns());
  if(headers.length!==previous.length)sheet.getRange(1,1,1,headers.length).setValues([headers.map(value=>"'"+value)]).setFontWeight('bold').setWrap(true).setFrozenRows(1);
  const rowValues=headers.map(key=>{const value=values[key];return value==null||value===''?'':"'"+(Array.isArray(value)?value.join('\\n'):String(value));});
  const current=onlineApplicationSheetRow_(sheet,ticket.id);
  if(current)sheet.getRange(current.sheetRow,1,1,headers.length).setValues([rowValues]);
  else sheet.appendRow(rowValues);
}

function onlineSaveTiponForm_(props,d,apps,answersSheet) {
  onlineStaff_();
  const ticket=onlineActiveTiponTicket_(d),fields=onlineFields_().filter(field=>field.id);
  if(!d.answers||typeof d.answers!=='object'||Array.isArray(d.answers))throw Error('Complete the Tipon Form fields.');
  const stored=onlineTiponProfile_(ticket,apps,answersSheet),profile=stored.profile;
  for(const field of fields){
    const raw=d.answers[field.title];
    const value=Array.isArray(raw)?raw.map(item=>onlineText_(item,500)).filter(Boolean):onlineText_(raw,1500);
    const selected=Array.isArray(value)?value:value?[value]:[];
    if(field.required&&!selected.length)throw Error('Required: '+field.title);
    if(field.choices.length&&!field.other&&selected.some(item=>!field.choices.includes(item)))throw Error('Invalid answer: '+field.title);
    profile[field.title]=value;
  }
  const photoUrl=onlineSaveCapture_(props,ticket.id,'photo',d.photoData)||String(profile['Photo URL']||'');
  const signatureUrl=onlineSaveCapture_(props,ticket.id,'signature',d.signatureData)||String(profile['Signature URL']||'');
  if(!photoUrl||!signatureUrl)throw Error('Capture the applicant photo and signature before saving.');
  profile['Photo URL']=photoUrl;profile['Signature URL']=signatureUrl;
  const first=String(profile['First Name']||'').trim(),last=String(profile['Last Name']||'').trim();
  const name=[first,last].filter(Boolean).join(' ')||String(ticket.row[6]||'');
  const existingIndex=stored.recordIndex;
  let stage='Tipon Form Saved';
  if(existingIndex>=0){
    const original=stored.record;stage=String(original[6]||'Review');
    apps.getRange(stored.recordRow,4,1,3).setValues([[ticket.transaction,name,JSON.stringify(profile)].map(value=>"'"+String(value))]);
    const updated=apps.getRange(stored.recordRow,1,1,ONLINE_HEADERS.length).getValues()[0];
    saveOnlineApplicantAnswers_(props,updated,answersSheet);
  }else onlineUpsertTiponAnswers_(answersSheet,ticket,profile,name,stage);
  SpreadsheetApp.flush();
  return {ok:true,id:ticket.id,name,photoUrl,signatureUrl,stage};
}

function onlineMergeSheetAnswers_(records, sheet) {
  if(!sheet)return records.map(row=>{row._deleted=true;return row;});
  const lastRow=sheet.getLastRow(), lastColumn=sheet.getLastColumn();
  if(lastRow<2||lastColumn<1)return records.map(row=>{row._deleted=true;return row;});
  const values=sheet.getRange(1,1,lastRow,lastColumn).getDisplayValues();
  const headers=values[0].map(value=>String(value||'').trim());
  const idColumn=headers.indexOf('Application ID');
  if(idColumn<0)return records.map(row=>{row._deleted=true;return row;});
  const answerById=new Map(values.slice(1).filter(row=>row[idColumn]).map(row=>[String(row[idColumn]),row]));
  const metadata=new Set(['Application ID','Submitted at (UTC)','Recruitment Transaction','Applicant Name','Application Status','Appointment (PHT)','Queue Number','Uploaded File Names','Uploaded File Links','PDF Generation Status','Generated PDF URL','Generated PDF At','Generated PDF Template Code']);
  return records.map(row=>{
    const answer=answerById.get(String(row[0]));
    if(!answer){row._deleted=true;return row;}
    const cell=key=>{const column=headers.indexOf(key);return column<0?'':String(answer[column]||'');};
    const profile=JSON.parse(row[5]||'{}');
    for(const key of headers){
      if(metadata.has(key))continue;
      profile[key]=String(answer[headers.indexOf(key)]||'');
    }
    const transaction=cell('Recruitment Transaction');
    if(transaction==='New Hire')row[3]='NH';
    else if(transaction==='Rehire')row[3]='RH';
    const name=cell('Applicant Name');
    if(name)row[4]=name;
    else {
      const first=String(profile['First Name']||'').trim(),last=String(profile['Last Name']||'').trim();
      if(first||last)row[4]=[first,last].filter(Boolean).join(' ');
    }
    const status=cell('Application Status'),appointment=cell('Appointment (PHT)'),number=cell('Queue Number');
    if(status)row[6]=status;
    row[7]=appointment;row[9]=number;
    row[5]=JSON.stringify(profile);
    if(row[6]==='Deleted')row._deleted=true;
    return row;
  });
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
  const cache=CacheService.getScriptCache(),cacheKey='ONLINE_FORM_FIELDS_V2',cached=cache.get(cacheKey);
  if(cached){try{return JSON.parse(cached)}catch(e){cache.remove(cacheKey)}}
  const persisted=PropertiesService.getScriptProperties().getProperty('ONLINE_FORM_FIELDS_JSON');
  if(persisted){try{const parsed=JSON.parse(persisted);cache.put(cacheKey,persisted,21600);return parsed}catch(e){PropertiesService.getScriptProperties().deleteProperty('ONLINE_FORM_FIELDS_JSON')}}
  const data = FormApp.openById(ONLINE_FORM_ID).getItems();
  const converters = {
    TEXT:'asTextItem',PARAGRAPH_TEXT:'asParagraphTextItem',DATE:'asDateItem',DATETIME:'asDateTimeItem',TIME:'asTimeItem',
    LIST:'asListItem',MULTIPLE_CHOICE:'asMultipleChoiceItem',CHECKBOX:'asCheckboxItem'
  };
  const fields=data.map(item=>{
    const type=String(item.getType());
    if (['SECTION_HEADER','PAGE_BREAK'].includes(type)) return {type,title:item.getTitle()};
    if (['IMAGE','VIDEO'].includes(type)) return null;
    if (!converters[type]) throw Error('Unsupported AP CONS question type: '+type);
    const question=item[converters[type]]();
    return {id:String(item.getId()),type,title:item.getTitle(),help:question.getHelpText?question.getHelpText():'',
      required:question.isRequired(),choices:question.getChoices?question.getChoices().map(choice=>choice.getValue()):[],
      other:question.hasOtherOption?question.hasOtherOption():false};
  }).filter(Boolean);
  const json=JSON.stringify(fields);
  cache.put(cacheKey,json,21600);
  // Script Properties limit each value to about 9 KB; persist only a safe-sized schema.
  if(json.length<8000)PropertiesService.getScriptProperties().setProperty('ONLINE_FORM_FIELDS_JSON',json);
  return fields;
}

function onlineFindHeaderRow_(sheet,headerName,value) {
  if(!sheet||sheet.getLastRow()<2)return null;
  const headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0].map(item=>String(item||'').trim());
  const index=headers.indexOf(headerName);
  if(index<0)return null;
  const found=sheet.getRange(2,index+1,sheet.getLastRow()-1,1).createTextFinder(String(value)).matchEntireCell(true).findNext();
  return found?{headers,row:sheet.getRange(found.getRow(),1,1,sheet.getLastColumn()).getDisplayValues()[0],sheetRow:found.getRow()}:null;
}

function onlinePdfTemplateRows_(sheet) {
  return onlineRead_(sheet).filter(row=>String(row[5]).toUpperCase()!=='FALSE'&&String(row[5]).toUpperCase()!=='NO')
    .map(row=>({code:String(row[0]||''),type:String(row[1]||''),fileId:String(row[2]||''),name:String(row[3]||''),folderId:'',namingConvention:String(row[7]||'')}));
}

function onlineEnsurePdfInfrastructure_(props,data) {
  ensureOnlineSheet_(data,ONLINE_PDF_TEMPLATES_SHEET,ONLINE_PDF_TEMPLATE_HEADERS);
  ensureOnlineSheet_(data,ONLINE_PDF_LOG_SHEET,ONLINE_PDF_LOG_HEADERS);
}

function onlinePdfTemplateFor_(templates,transaction) {
  const code=String(transaction||'').toUpperCase();
  const exact=templates.filter(item=>item.type===code),both=templates.filter(item=>item.type==='BOTH');
  const matches=exact.length?exact:both;
  if(!matches.length)return null;
  if(matches.length>1)throw Error('More than one active PDF template matches '+code+'. Keep only one active template for each transaction type.');
  return matches[0];
}

function onlineSavePdfTemplate_(d,props,sheet) {
  const label=onlineText_(d.code||'',80),code=label.toUpperCase(),type='BOTH';
  const naming=onlineText_(d.namingConvention||'',160)||'{{Applicant Name}} - {{Queue Number}} - Pre-employment Requirements';
  const name=onlineText_(d.filename||'TIPON recruitment PDF template',150);
  if(!label||!naming)throw Error('Enter a pre-employment requirements label and naming convention.');
  let fileId=String(d.templateId||'').trim();
  const urlMatch=fileId.match(/\/document\/d\/([A-Za-z0-9_-]+)/)||fileId.match(/[?&]id=([A-Za-z0-9_-]+)/);
  if(urlMatch)fileId=urlMatch[1];
  let file;
  if(d.fileData){
    const mime=String(d.fileMime||''),encoded=String(d.fileData||'');
    if(!['application/vnd.openxmlformats-officedocument.wordprocessingml.document','application/msword'].includes(mime)||encoded.length>7000000)throw Error('Upload a Word template (.doc or .docx) up to 5 MB.');
    const bytes=Utilities.base64Decode(encoded);
    if(!bytes.length||bytes.length>5*1024*1024)throw Error('Template file must be 5 MB or smaller.');
    const parent=props.getProperty('ONLINE_FOLDER');
    if(!parent)throw Error('Online Queue storage has not been initialized.');
    const converted=Drive.Files.create({name:'TIPON-PDF-TEMPLATE-'+code,mimeType:'application/vnd.google-apps.document',parents:[parent]},Utilities.newBlob(bytes,mime,name),{fields:'id,mimeType'});
    fileId=String(converted.id||'');
  }
  if(!fileId)throw Error('Upload a Word template (.doc or .docx).');
  try{file=DriveApp.getFileById(fileId)}catch(e){throw Error('The template file could not be opened. Check its Drive ID and permissions.')}
  if(file.getMimeType()!=='application/vnd.google-apps.document')throw Error('Use a Google Docs template or upload a Word document for conversion.');
  const activeRows=onlineRead_(sheet);
  activeRows.forEach((row,index)=>{if(String(row[5]).toUpperCase()!=='FALSE')sheet.getRange(index+2,6).setValue(false)});
  sheet.appendRow([code,type,fileId,file.getName(),' ',true,new Date().toISOString(),naming]);
  SpreadsheetApp.flush();
  return {ok:true,template:{code,type,name:file.getName(),label,namingConvention:naming,fileId}};
}

function onlinePdfStatuses_(d,data) {
  const tickets=Array.isArray(d.tickets)?d.tickets.slice(0,100):[];
  const answers=data.getSheetByName('ONLINE_APPLICATIONS'),log=data.getSheetByName(ONLINE_PDF_LOG_SHEET),templates=onlinePdfTemplateRows_(data.getSheetByName(ONLINE_PDF_TEMPLATES_SHEET));
  const result={};
  for(const ticket of tickets){
    const id=String(ticket?.ticketId||''),queueNo=String(ticket?.queueNo||''),transaction=String(ticket?.transaction||'').toUpperCase();
    if(!/^[a-f0-9-]{36}$/i.test(id)||!new RegExp('^(?:OL)?'+transaction+'[-_]').test(queueNo)||!['NH','RH'].includes(transaction))continue;
    const logged=onlineFindHeaderRow_(log,'Ticket ID',id),generated=logged?.row;
    if(generated&&String(generated[8]).toLowerCase()==='generated'){
      result[id]={status:'Generated',url:String(generated[6]||''),templateCode:String(generated[3]||''),available:false};continue;
    }
    const answer=onlineFindHeaderRow_(answers,'Queue Number',queueNo),template=onlinePdfTemplateFor_(templates,transaction);
    const answerStatus=answer&&answer.headers.includes('PDF Generation Status')?String(answer.row[answer.headers.indexOf('PDF Generation Status')]||''):'';
    const answerUrl=answer&&answer.headers.includes('Generated PDF URL')?String(answer.row[answer.headers.indexOf('Generated PDF URL')]||''):'';
    if(answerStatus.toLowerCase()==='generated'&&answerUrl){result[id]={status:'Generated',url:answerUrl,templateCode:answer.row[answer.headers.indexOf('Generated PDF Template Code')]||'',available:false};continue}
    result[id]={status:!answer?'Tipon Form missing':!template?'Template not configured':'Not generated',url:'',templateCode:template?.code||'',available:Boolean(answer&&template)};
  }
  return result;
}

function onlineReplaceTemplate_(fileId,profile) {
  const doc=DocumentApp.openById(fileId),containers=[doc.getBody(),doc.getHeader(),doc.getFooter()].filter(Boolean);
  const values={...profile};
  Object.keys(values).forEach(key=>{
    const escaped=String(key).replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
    const pattern='\\{\\{'+escaped+'\\}\\}';
    const replacement=Array.isArray(values[key])?values[key].join(', '):String(values[key]??'');
    containers.forEach(container=>container.replaceText(pattern,replacement.replace(/\$/g,'\\$')));
  });
  doc.saveAndClose();
}

function onlineGeneratePdfs_(d,props,data) {
  onlineStaff_();
  const tickets=Array.isArray(d.tickets)?d.tickets.slice(0,10):[];
  if(!tickets.length)throw Error('Select one to ten completed New Hire or Rehire records.');
  const queue=data.getSheetByName(CFG.queue),answers=data.getSheetByName('ONLINE_APPLICATIONS'),templates=onlinePdfTemplateRows_(data.getSheetByName(ONLINE_PDF_TEMPLATES_SHEET)),log=data.getSheetByName(ONLINE_PDF_LOG_SHEET);
  const answerHeaders=answers.getRange(1,1,1,answers.getLastColumn()).getDisplayValues()[0].map(value=>String(value||'').trim());
  const appRefColumn=answerHeaders.indexOf('Applicant Reference No.');
  const generated=[];
  for(const ticket of tickets){
    const id=String(ticket?.ticketId||''),queueNo=String(ticket?.queueNo||''),transaction=String(ticket?.transaction||'').toUpperCase();
    if(!/^[a-f0-9-]{36}$/i.test(id)||!['NH','RH'].includes(transaction)||!new RegExp('^(?:OL)?'+transaction+'[-_]').test(queueNo))throw Error('Only completed New Hire and Rehire queue numbers can be generated.');
    const queueMatch=queue.getLastRow()>1?queue.getRange(2,1,queue.getLastRow()-1,1).createTextFinder(id).matchEntireCell(true).findNext():null;
    if(!queueMatch||String(queue.getRange(queueMatch.getRow(),9).getDisplayValue()).toUpperCase()!=='COMPLETED')throw Error(queueNo+' is not in Completed Queue.');
    const existing=onlineRead_(log).find(row=>String(row[0])===id&&String(row[8]).toLowerCase()==='generated');
    if(existing)throw Error(queueNo+' already has a generated PDF.');
    const answer=onlineFindHeaderRow_(answers,'Queue Number',queueNo);
    if(!answer)throw Error('No ONLINE_APPLICATIONS form data was found for '+queueNo+'.');
    const template=onlinePdfTemplateFor_(templates,transaction);
    if(!template)throw Error('No active PDF template is configured for '+transaction+'.');
    const root=DriveApp.getFolderById(ONLINE_PDF_ROOT_FOLDER_ID);
    const name=String(answer.row[answerHeaders.indexOf('Applicant Name')]||queueNo).replace(/[\\/:*?"<>|#]+/g,' ').replace(/\s+/g,' ').trim().slice(0,90)||queueNo;
    let applicantFolderIterator=root.getFoldersByName(name),folder=applicantFolderIterator.hasNext()?applicantFolderIterator.next():root.createFolder(name);
    if([DriveApp.Access.ANYONE,DriveApp.Access.ANYONE_WITH_LINK].includes(folder.getSharingAccess()))throw Error('Applicant PDF folders must use restricted sharing.');
    const privateFolderId=props.getProperty('ONLINE_FOLDER');
    const working=DriveApp.getFileById(template.fileId).makeCopy('TIPON-TEMP-'+queueNo+'-'+Date.now(),privateFolderId?DriveApp.getFolderById(privateFolderId):root);
    let pdfFile;
    try{
      const profile={};answerHeaders.forEach((header,i)=>{profile[header]=String(answer.row[i]||'')});
      profile['Queue Number']=queueNo;profile['Applicant Reference No.']=appRefColumn>=0?String(answer.row[appRefColumn]||''):'';
      onlineReplaceTemplate_(working.getId(),profile);
      const convention=String(template.namingConvention||'{{Applicant Name}} - {{Queue Number}} - Pre-employment Requirements');
      let pdfName=convention.replace(/\{\{\s*([^{}]+?)\s*\}\}/g,(_,key)=>String(profile[key]||''));
      pdfName=pdfName.replace(/[\\/:*?"<>|#]+/g,' ').replace(/\s+/g,' ').trim().replace(/\.pdf$/i,'').slice(0,140)||name+' - '+queueNo+' - Pre-employment Requirements';
      pdfFile=folder.createFile(working.getAs(MimeType.PDF).setName(pdfName+'.pdf'));
    }catch(error){working.setTrashed(true);throw error}
    working.setTrashed(true);
    const stamp=new Date().toISOString(),ref=appRefColumn>=0?String(answer.row[appRefColumn]||''):'';
    log.appendRow([id,queueNo,ref,template.code,transaction,pdfFile.getId(),pdfFile.getUrl(),stamp,'Generated']);
    const extra={'PDF Generation Status':'Generated','Generated PDF URL':pdfFile.getUrl(),'Generated PDF At':stamp,'Generated PDF Template Code':template.code};
    let lastCol=answers.getLastColumn(),headers=answerHeaders;
    const added=Object.keys(extra).filter(key=>!headers.includes(key));
    if(added.length){answers.getRange(1,lastCol+1,1,added.length).setValues([added]);headers=headers.concat(added);lastCol+=added.length}
    const updates=Object.keys(extra).map(key=>[headers.indexOf(key)+1,extra[key]]);
    updates.forEach(([col,value])=>answers.getRange(answer.sheetRow,col).setValue(value));
    generated.push({ticketId:id,queueNo,url:pdfFile.getUrl(),templateCode:template.code,status:'Generated'});
  }
  SpreadsheetApp.flush();
  return {ok:true,generated};
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
  const page=HtmlService.createTemplateFromFile('OnlinePage');
  page.initial=JSON.stringify({view,token:/^[a-f0-9]{64}$/.test(ticket)?ticket:'',site:/^[a-f0-9]{64}$/.test(onsite)?onsite:''});
  return page.evaluate().setTitle('TIPON TAWAG · Online Queue');
}

function onlineQueue(action, input) {
  const d=input||{};
  if (action==='fields') return onlineFields_();
  const staffActions=['list','complete','deleteApplication','getTiponForm','tiponFormStatus','saveTiponForm','readStaff','onsiteLink','download','sendMessage','listPdfTemplates','savePdfTemplate','deletePdfTemplate','getPdfStatuses','generatePdfs'];
  const staff=staffActions.includes(action)?onlineStaff_():'';
  const mutations=['submit','upload','complete','book','arrive','confirmArrival','saveTiponForm','sendMessage','sendApplicantMessage','savePdfTemplate','deletePdfTemplate','generatePdfs'];
  const lock=mutations.includes(action)?LockService.getScriptLock():null;
  if(lock) lock.waitLock(20000);
  try { return onlineOperate_(action,d,staff); }
  finally { if(lock) { try { SpreadsheetApp.flush(); } finally { lock.releaseLock(); } } }
}

function onlineOperate_(action,d,staff) {
  const props=PropertiesService.getScriptProperties();
  const data=ss_();
  const apps=data.getSheetByName(ONLINE_RECORDS_SHEET);
  const files=data.getSheetByName(ONLINE_DOCUMENTS_SHEET);
  const messages=data.getSheetByName(ONLINE_MESSAGES_SHEET);
  if(!apps||!files||!messages) throw Error('Run initializeOnlineQueue() before using Online Recruitment.');
  const answersSheet=data.getSheetByName('ONLINE_APPLICATIONS');
  if(action==='tiponFormStatus')return onlineTiponFormStatus_(d,apps,answersSheet);
  if(action==='getTiponForm')return onlineGetTiponForm_(d,apps,answersSheet);
  if(['listPdfTemplates','savePdfTemplate','deletePdfTemplate','getPdfStatuses','generatePdfs'].includes(action))onlineEnsurePdfInfrastructure_(props,data);
  if(action==='listPdfTemplates')return onlinePdfTemplateRows_(data.getSheetByName(ONLINE_PDF_TEMPLATES_SHEET)).map(({code,type,name,fileId,folderId})=>({code,type,name,fileId,folderId}));
  if(action==='savePdfTemplate')return onlineSavePdfTemplate_(d,props,data.getSheetByName(ONLINE_PDF_TEMPLATES_SHEET));
  if(action==='deletePdfTemplate'){
    const code=String(d.code||'').trim().toUpperCase(),sheet=data.getSheetByName(ONLINE_PDF_TEMPLATES_SHEET),rows=onlineRead_(sheet),index=rows.findIndex(row=>String(row[0]).toUpperCase()===code&&String(row[5]).toUpperCase()!=='FALSE');
    if(index<0)throw Error('Active PDF template not found.');
    sheet.getRange(index+2,6).setValue(false);return {ok:true,code};
  }
  if(action==='getPdfStatuses')return onlinePdfStatuses_(d,data);
  if(action==='generatePdfs')return onlineGeneratePdfs_(d,props,data);
  const rawRecords=onlineRead_(apps);
  const all=onlineMergeSheetAnswers_(rawRecords,answersSheet), url=props.getProperty('ONLINE_PUBLIC_URL')||ScriptApp.getService().getUrl();
  const now=new Date(), today=day_();
  if(action==='onsiteLink') return url+'?online=onsite&site='+props.getProperty('ONLINE_ONSITE_KEY');
  const fileList=id=>onlineRead_(files).filter(file=>file[0]===id);
  if(action==='getTiponForm')return onlineGetTiponForm_(d,all,answersSheet);
  if(action==='saveTiponForm')return onlineSaveTiponForm_(props,d,apps,answersSheet);
  if(action==='list') {
    const docs=onlineRead_(files);
    return all.filter(row=>!row._deleted&&row[6]!=='Deleted').slice().reverse().map(row=>({id:row[0],date:row[2],transaction:row[3],name:row[4],stage:row[6],
      appointment:row[7],number:row[9],files:docs.filter(file=>file[0]===row[0]).map(file=>({name:file[2],id:file[1]}))}));
  }
  if(action==='submit'){
    if(!['NH','RH'].includes(d.transaction)) throw Error('Choose New Hire or Rehire.');
    if(!/^[a-f0-9]{64}$/.test(String(d.token||''))) throw Error('Invalid application session.');
    const digest=onlineDigest_(d.token);
    const existing=all.find(row=>row[1]===digest&&!row._deleted&&row[6]!=='Deleted');
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
      const applicationRow=[record,digest,now.toISOString(),d.transaction,name,JSON.stringify(profile),'Review','','',''];
      onlineWrite_(apps,applicationRow);
      saveOnlineApplicantAnswers_(props,applicationRow);
    } catch(error) { for(const file of saved) file.setTrashed(true); throw error; }
    return {url:url+'?online=applicant&case='+d.token};
  }
  const idx=staff?all.findIndex(row=>row[0]===d.id&&!row._deleted&&row[6]!=='Deleted'):(/^[a-f0-9]{64}$/.test(String(d.token||''))?all.findIndex(row=>row[1]===onlineDigest_(d.token)&&!row._deleted&&row[6]!=='Deleted'):-1);
  if(idx<0) throw Error('Application link unavailable.');
  const row=all[idx], id=row[0], stage=row[6];
  const update=(next,appointment=row[7],sequence=row[8],number=row[9])=>{
    apps.getRange(idx+2,4,1,3).setValues([[row[3],row[4],row[5]].map(value=>"'"+String(value))]);
    apps.getRange(idx+2,7,1,4).setValues([[next,appointment,sequence,number].map(value=>"'"+String(value))]);
    try {
      const updated=apps.getRange(idx+2,1,1,ONLINE_HEADERS.length).getValues()[0];
      saveOnlineApplicantAnswers_(props,updated);
      return updated;
    } catch(error) {
      // Keep the authoritative application record and its mirrored sheet in sync on failed writes.
      apps.getRange(idx+2,7,1,4).setValues([[row[6],row[7],row[8],row[9]].map(value=>"'"+String(value))]);
      SpreadsheetApp.flush();
      throw error;
    }
  };
  if(action==='deleteApplication'){
    if(!staff)throw Error('Only authorized recruiters can remove applications.');
    update('Deleted');
    return {ok:true};
  }
  if(action==='sendMessage'||action==='sendApplicantMessage'){
    const applicantReply=action==='sendApplicantMessage';
    if(applicantReply&&staff)throw Error('Applicant replies must use the private application link.');
    if(!applicantReply&&!staff)throw Error('Only authorized recruiters can send messages.');
    const message=onlineText_(d.message,2000);
    if(!message)throw Error('Enter a message before sending.');
    onlineWrite_(messages,[id,applicantReply?'Applicant':staff,message,now.toISOString()]);
    return {ok:true};
  }
  if(action==='download'){
    const file=fileList(id).find(doc=>doc[1]===d.file);
    if(!file)throw Error('Document unavailable.');
    const blob=DriveApp.getFileById(file[1]).getBlob();
    return {name:file[2],mime:blob.getContentType(),data:Utilities.base64Encode(blob.getBytes())};
  }
  if(action==='readStaff'||action==='read'){
    const notices=onlineRead_(messages).filter(message=>String(message[0])===String(id)).map(message=>({author:String(message[1])==='Applicant'?'Applicant':'Recruiter',text:message[2],at:message[3]}));
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
    const appointment=onlineText_(d.appointment,10);
    const date=new Date(appointment+'T00:00:00+08:00');
    const validDate=Number.isFinite(date.getTime())&&Utilities.formatDate(date,'Asia/Manila','yyyy-MM-dd')===appointment;
    if(!/^\d{4}-\d\d-\d\d$/.test(appointment)||!validDate||appointment<today) throw Error('Choose a valid office visit date that is today or later.');
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
    const allowed=${JSON.stringify(staff?['list','complete','deleteApplication','getTiponForm','tiponFormStatus','saveTiponForm','readStaff','onsiteLink','download','sendMessage','listPdfTemplates','savePdfTemplate','deletePdfTemplate','getPdfStatuses','generatePdfs']:['fields','submit','read','upload','book','arrive','confirmArrival','sendApplicantMessage'])};
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
  const prior=onlineApplicationSheetRow_(sheet,row[0]);
  const priorValue=key=>prior?String(prior.row[prior.headers.indexOf(key)]||''):'';
  const metadata={
    'Application ID':row[0],
    'Submitted at (UTC)':row[2],
    'Recruitment Transaction':row[3]==='NH'?'New Hire':'Rehire',
    'Applicant Name':row[4],
    'Application Status':row[6]||'',
    'Appointment (PHT)':row[7]||'',
    'Queue Number':row[9]||'',
    'Uploaded File Names':documents.map(file=>file[2]).join('\\n'),
    'Uploaded File Links':documents.map(file=>'https://drive.google.com/file/d/'+file[1]+'/view').join('\\n'),
    'PDF Generation Status':priorValue('PDF Generation Status'),
    'Generated PDF URL':priorValue('Generated PDF URL'),
    'Generated PDF At':priorValue('Generated PDF At'),
    'Generated PDF Template Code':priorValue('Generated PDF Template Code')
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
