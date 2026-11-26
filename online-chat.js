'use strict';
const onlineChat=(()=>{
  const escape=value=>String(value??'').replace(/[&<>"']/g,character=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[character]));
  const timeFormat=new Intl.DateTimeFormat('en-PH',{timeZone:'Asia/Manila',month:'short',day:'numeric',hour:'numeric',minute:'2-digit'});
  function markup(messages,ownRole,action,draft=''){
    const thread=(messages||[]).map(message=>{
      const own=message.author===ownRole;
      const date=new Date(message.at);
      const time=message.at&&!Number.isNaN(date.getTime())?'<time datetime="'+escape(date.toISOString())+'">'+escape(timeFormat.format(date))+'</time>':'';
      return '<article class="chat-bubble'+(own?' chat-bubble--own':'')+'"><div class="chat-meta"><span>'+escape(own?'You':message.author)+'</span>'+time+'</div><p class="chat-text">'+escape(message.text)+'</p></article>';
    }).join('');
    return '<section class="conversation"><div class="chat-heading"><h2>Conversation</h2><button type="button" data-refresh>Refresh messages</button></div>'+
      '<div class="chat-thread" role="log" aria-label="Messages" tabindex="0">'+(thread||'<p class="chat-empty">No messages yet.</p>')+'</div>'+
      '<form class="chat-composer" data-action="'+escape(action)+'"><label><span>Your message</span><textarea name="message" rows="2" maxlength="2000" placeholder="Type a message…" required>'+escape(draft)+'</textarea></label><button>Send message</button></form></section>';
  }
  function scroll(root){
    const thread=root.querySelector('.chat-thread');
    if(thread)thread.scrollTop=thread.scrollHeight;
  }
  function draft(root){return root.querySelector('.chat-composer textarea')?.value||'';}
  return {markup,scroll,draft};
})();
