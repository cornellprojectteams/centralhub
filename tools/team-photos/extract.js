() => {
  // whole words only, so "Barrow" or "Ismail" in a file name is not taken for an arrow or mail icon
  const bad=/(^|[^a-z])(logos?|icons?|sponsors?|banners?|favicon|sprites?|arrows?|social|linkedin|instagram|facebook|github|twitter|e?mail|full-?team|fullteam|groups?|candids?)([^a-z]|$)/i;
  const cands=[];
  for (const im of document.querySelectorAll('img')) {
    const src=im.currentSrc||im.src||im.getAttribute('data-src')||'';
    cands.push({el:im,src});
  }
  for (const el of document.querySelectorAll('div,span,a,figure,li')) {
    const bg=getComputedStyle(el).backgroundImage;
    if (bg && bg.startsWith('url(')) cands.push({el,src:bg.slice(4,-1).replace(/["']/g,'')});
  }
  const ok=cands.filter(c=>{
    if(!c.src) return false;
    // a generated initials image (an inline SVG) stands in for a missing photo: keep the person
    if(c.src.startsWith('data:')){ if(!/^data:image\/svg/i.test(c.src)||!(c.el.alt||'').trim()) return false; c.src='placeholder:svg'; }
    else if(/\.svg(\?|$)/i.test(c.src)&&!/placeholder|profile|headshot/i.test(c.src)) return false;
    let path=c.src; try{ path=decodeURIComponent(new URL(c.src,location.href).pathname+new URL(c.src,location.href).search); }catch(e){}
    if (bad.test(path)||bad.test(c.el.alt||'')) return false;
    const r=c.el.getBoundingClientRect();
    if(!r.width||!r.height){
      // a card hidden when the page was read (collapsed tab, flip card): judge by the image itself
      const nw=c.el.naturalWidth||0, nh=c.el.naturalHeight||0;
      c.w=nw; c.h=nh;
      return nw>=70 && nh>=70 && nh/nw>0.6 && nh/nw<2.2 && !!(c.el.alt||'').trim();
    }
    const w=r.width, h=r.height;
    c.w=Math.round(w); c.h=Math.round(h);
    return w>=70 && h>=70 && w<=650 && h/w>0.6 && h/w<2.2;
  });
  const set=new Set(ok.map(c=>c.el));
  const countIn=(n)=>{let k=0; for(const e of set) if(n.contains(e)) k++; return k;};
  const texts=(n)=>{const out=[];const tw=document.createTreeWalker(n,NodeFilter.SHOW_TEXT);let t;while(t=tw.nextNode()){const p=t.parentElement; if(p&&/SCRIPT|STYLE|NOSCRIPT/.test(p.tagName))continue; const s=t.textContent.replace(/\s+/g,' ').trim(); if(s)out.push(s);}return out;};
  const res=[];
  for (const c of ok) {
    let card=c.el;
    while (card.parentElement && card.parentElement!==document.body && countIn(card.parentElement)===1) card=card.parentElement;
    let lines=texts(card).filter(s=>s.length<120);
    // fallback: text that follows the photo in document order, up to the next photo
    let after=[];
    { const tw=document.createTreeWalker(document.body,NodeFilter.SHOW_ELEMENT|NodeFilter.SHOW_TEXT); tw.currentNode=card; let n, k=0;
      while((n=tw.nextNode()) && k<6){ if(n.nodeType===1){ if(set.has(n)&&!card.contains(n)) break; continue;} if(card.contains(n))continue; const p=n.parentElement; if(p&&/SCRIPT|STYLE|NOSCRIPT/.test(p.tagName))continue; const s=n.textContent.replace(/\s+/g,' ').trim(); if(s&&s.length<120){after.push(s);k++;} } }
    // heading context: nearest preceding heading
    let head=''; let n=card;
    for (let i=0;i<8&&n;i++){ let s=n.previousElementSibling; while(s){ const h=s.matches('h1,h2,h3,h4')?s:s.querySelector&&[...s.querySelectorAll('h1,h2,h3,h4')].pop(); if(h&&h.innerText.trim()){head=h.innerText.trim();break;} s=s.previousElementSibling;} if(head)break; n=n.parentElement;}
    res.push({src:c.src,alt:c.el.alt||'',w:c.w,h:c.h,lines:lines.slice(0,12),after,head:head.slice(0,80)});
  }
  return res;
}
