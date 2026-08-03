const {launch}=require('./browser');
const fs=require('fs');
const BASE={A:'https://articulateusercontent.com/review/uploads/ii4eAmTLR7Uy6fYzgF6uD30Iw_2_-k0i/Cy6Sg2km/index.html',
            B:'https://articulateusercontent.com/review/uploads/LPABCAopJ3zRMchyQcSTbSBjQ3kbSZfJ/5sfAJDFK/index.html'};
const LESSONS=require('./data/lessons.json');
const MOD=process.argv[2];

const SCRAPE=()=>{
  const root=document.querySelector('.blocks-lesson')||document.body;
  const walk=(el)=>{
    const lines=[];
    const rec=(n,depth)=>{
      if(n.nodeType===3){ const t=n.textContent.replace(/\s+/g,' ').trim(); if(t) lines.push(t); return; }
      if(n.nodeType!==1) return;
      const tag=n.tagName;
      if(tag==='SCRIPT'||tag==='STYLE') return;
      if(tag==='IMG'){ lines.push(`[IMG alt="${n.getAttribute('alt')||''}"]`); return; }
      const c=String(n.className||'');
      let marker=null;
      if(/blocks-accordion__header/.test(c)) marker='[ACCORDION HEADER]';
      else if(/blocks-accordion__content|accordion__body/.test(c)) marker='[ACCORDION BODY]';
      else if(/flashcard-side--front/.test(c)) marker='[CARD FRONT]';
      else if(/flashcard-side--back/.test(c)) marker='[CARD BACK]';
      else if(/carousel-slide/.test(c)) marker='[SLIDE]';
      else if(/tabs__label|tabs-label/.test(c)) marker='[TAB LABEL]';
      else if(/tabs__panel|tabs-panel/.test(c)) marker='[TAB PANEL]';
      else if(/labeled-graphic-marker/.test(c)) marker='[MARKER]';
      else if(/block-gallery__caption/.test(c)) marker='[GALLERY CAPTION]';
      else if(/blocks-button__button/.test(c)) marker='[BUTTON]';
      else if(/block-text|block-image|block-quote|block-gallery|block-embed|block-list|block-divider|block-multimedia|block-statement/.test(c) && /^block-[a-z]+ /.test(c+' ')) marker='[BLOCK '+(c.match(/block-([a-z]+)/)||[])[1]+']';
      if(/^(H1|H2|H3|H4|H5|H6)$/.test(tag)) marker='['+tag+']';
      if(tag==='LI') marker='  •';
      if(marker) lines.push('\n'+marker);
      for(const ch of n.childNodes) rec(ch,depth+1);
      if(tag==='P'||tag==='LI'||/^H[1-6]$/.test(tag)) lines.push('\n');
    };
    rec(el,0);
    return lines.join(' ').replace(/[ \t]+/g,' ').replace(/\n /g,'\n').replace(/\n{3,}/g,'\n\n');
  };
  return {title:document.title, text:walk(root)};
};

(async()=>{
  const b=await launch();
  const p=await b.newPage({viewport:{width:1600,height:1200}});
  const res=[];
  for(const l of LESSONS[MOD].slice(1)){
    process.stderr.write('> '+l.text+'\n');
    await p.goto('about:blank');
    await p.goto(BASE[MOD]+l.href,{waitUntil:'domcontentloaded',timeout:120000});
    await p.waitForTimeout(6000);
    const target = l.href;
    const guard = async () => {
      const h = await p.evaluate(()=>location.hash);
      if (h !== target) { await p.goto(BASE[MOD]+target,{waitUntil:'domcontentloaded',timeout:120000}); await p.waitForTimeout(5000); return true; }
      return false;
    };
    for(let pass=0;pass<3;pass++){
      await p.evaluate(async()=>{for(let y=0;y<document.body.scrollHeight+2000;y+=500){window.scrollTo(0,y);await new Promise(r=>setTimeout(r,80));}});
      // only block-level continue buttons, never the lesson-footer continue
      for(const c of await p.$$('.blocks-continue__button')){
        try{ await c.click({timeout:1500}); await p.waitForTimeout(500); if(await guard()) break; }catch(e){}
      }
      if(await guard()) {}
    }
    // expand all accordions & tabs
    for(const h of await p.$$('.blocks-accordion__header')){ try{ const exp=await h.getAttribute('aria-expanded'); if(exp!=='true'){await h.click({timeout:2000}); await p.waitForTimeout(350);} }catch(e){} }
    await p.waitForTimeout(1200);
    const data=await p.evaluate(SCRAPE);
    data.lesson=l.text; data.href=l.href;
    res.push(data);
    process.stderr.write('  chars='+data.text.length+'\n');
  }
  fs.writeFileSync(`data/${MOD}_deep.json`, JSON.stringify(res,null,1));
  await b.close();
})();
