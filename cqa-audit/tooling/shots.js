const {launch}=require('./browser'); const fs=require('fs');
const BASE={A:'https://articulateusercontent.com/review/uploads/ii4eAmTLR7Uy6fYzgF6uD30Iw_2_-k0i/Cy6Sg2km/index.html',
            B:'https://articulateusercontent.com/review/uploads/LPABCAopJ3zRMchyQcSTbSBjQ3kbSZfJ/5sfAJDFK/index.html'};
const LESSONS=require('./data/lessons.json');
const MOD=process.argv[2];
const SLICE=940, W=1400;
(async()=>{
  const b=await launch(); const p=await b.newPage({viewport:{width:W,height:SLICE}});
  let idx=0;
  for(const l of LESSONS[MOD].slice(1)){
    idx++;
    await p.goto('about:blank');
    await p.goto(BASE[MOD]+l.href,{waitUntil:'domcontentloaded',timeout:120000});
    await p.waitForTimeout(7000);
    const target=l.href;
    for(let pass=0;pass<2;pass++){
      await p.evaluate(async()=>{const c=document.querySelector('.page-wrap')||document.scrollingElement;for(let y=0;y<c.scrollHeight+1500;y+=400){c.scrollTop=y;await new Promise(r=>setTimeout(r,70));}});
      for(const c of await p.$$('.blocks-continue__button')){ try{await c.click({timeout:1200}); await p.waitForTimeout(400);}catch(e){} }
      const h=await p.evaluate(()=>location.hash);
      if(h!==target){ await p.goto(BASE[MOD]+target,{waitUntil:'domcontentloaded',timeout:120000}); await p.waitForTimeout(5000); }
    }
    // hide the left nav sidebar for clean content shots
    // collapse the nav sidebar so the reading view matches a learner's default
    try{ const t=await p.$('button[aria-label="Close navigation menu"]'); if(t) await t.click({timeout:2000}); }catch(e){}
    await p.waitForTimeout(1200);
    await p.evaluate(()=>{const c=document.querySelector('.page-wrap');if(c)c.scrollTop=0;}); await p.waitForTimeout(1200);
    const H=await p.evaluate(()=>{const c=document.querySelector('.page-wrap');return c?c.scrollHeight:document.body.scrollHeight;});
    const n=Math.min(Math.ceil(H/SLICE), 14);
    const name=String(idx).padStart(2,'0')+'_'+l.text.replace(/[^a-z0-9]+/gi,'_');
    for(let i=0;i<n;i++){
      await p.evaluate(y=>{const c=document.querySelector('.page-wrap');if(c)c.scrollTop=y;else window.scrollTo(0,y);}, i*SLICE);
      await p.waitForTimeout(700);
      await p.screenshot({path:`shots/${MOD}_${name}_s${i}.png`});
    }
    console.log(MOD, name, 'height='+H, 'slices='+n);
  }
  await b.close();
})();
