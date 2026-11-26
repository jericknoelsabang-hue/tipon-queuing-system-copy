'use strict';
const bridgeOrigin=window.location.origin;
const bridgeChannel=crypto.randomUUID();
const staffBridge=document.currentScript?.dataset.mode==='staff';
const onlineBackend=new URL(staffBridge?window.TIPON_CONFIG.onlineStaffBackendUrl:window.TIPON_CONFIG.onlineBackendUrl);
const pendingOnline=new Map();
let bridgeWindow, bridgeSourceOrigin, acceptBridge, rejectBridge;
const bridgeReady=new Promise((resolve,reject)=>{acceptBridge=resolve;rejectBridge=reject;});
// Rendering the form no longer waits for this connection. A later submission still needs it.
bridgeReady.catch(()=>{});
const bridgeDeadline=setTimeout(()=>rejectBridge(Error('Online Queue connection timed out. Reload this page to retry.')),30000);
window.addEventListener('message',event=>{
  const data=event.data;
  if(!data||data.channel!==bridgeChannel||!/^https:\/\/(?:script\.google\.com|[a-z0-9-]+-script\.googleusercontent\.com|script\.googleusercontent\.com)$/.test(event.origin))return;
  if(data.type==='tipon-online-ready'&&!bridgeWindow){
    bridgeWindow=event.source;bridgeSourceOrigin=event.origin;clearTimeout(bridgeDeadline);acceptBridge();return;
  }
  if(event.source!==bridgeWindow||event.origin!==bridgeSourceOrigin||data.type!=='tipon-online-result')return;
  const request=pendingOnline.get(data.id);
  if(!request)return;
  pendingOnline.delete(data.id);clearTimeout(request.timer);
  if(data.error)request.reject(Error(data.error));else request.resolve(data.result);
});
const onlineFrame=document.createElement('iframe');
onlineFrame.hidden=true;onlineFrame.title='TIPON Online Queue connection';
onlineBackend.searchParams.set('online',staffBridge?'staff-bridge':'bridge');
onlineBackend.searchParams.set('channel',bridgeChannel);
onlineFrame.src=onlineBackend.href;
document.body.append(onlineFrame);
async function onlineRequest(action,input){
  await bridgeReady;
  return new Promise((resolve,reject)=>{
    const id=crypto.randomUUID();
    const timer=setTimeout(()=>{pendingOnline.delete(id);reject(Error('The request timed out. Please retry.'));},90000);
    pendingOnline.set(id,{resolve,reject,timer});
    bridgeWindow.postMessage({type:'tipon-online-request',channel:bridgeChannel,id,action,input},bridgeSourceOrigin);
  });
}
