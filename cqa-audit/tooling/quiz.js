const {launch}=require('./browser'); const fs=require('fs');
const BASE={A:'https://articulateusercontent.com/review/uploads/ii4eAmTLR7Uy6fYzgF6uD30Iw_2_-k0i/Cy6Sg2km/index.html',
            B:'https://articulateusercontent.com/review/uploads/LPABCAopJ3zRMchyQcSTbSBjQ3kbSZfJ/5sfAJDFK/index.html'};
const QH={A:'#/lessons/W0p7W8u1zReSyvXwYJRL2RbchSJnemSs', B:'#/lessons/n_nbYpGhYbm4ksjVJeU8UV10ubUNH5sG'};
const MOD=process.argv[2]; const PICK=parseInt(process.argv[3]||'0',10);
(async()=>{
  const b=await launch(); const p=await b.newPage({viewport:{width:1400,height:1000}});
  await p.goto(BASE[MOD]+QH[MOD],{waitUntil:'domcontentloaded',timeout:120000});
  await p.waitForTimeout(8000);
  const out=[];
  try{ await p.click('.quiz-start__start-button',{timeout:8000}); }catch(e){ console.log('start fail',e.message.slice(0,60)); }
  await p.waitForTimeout(2500);
  for(let q=0;q<12;q++){
    const card = await p.$('.quiz__card--active');
    if(!card) break;
    const snap = async()=> p.evaluate(()=>{
      const c=document.querySelector('.quiz__card--active'); if(!c) return null;
      return {
        head:(c.querySelector('h1,.quiz-card__header')||{}).innerText||'',
        stem:(c.querySelector('.quiz-item__content, .quiz-card__main')||{}).innerText||'',
        options:Array.from(c.querySelectorAll('.quiz-multiple-choice-option__text, .quiz-matching-item, label')).map(e=>e.innerText.trim()).filter(Boolean),
        feedback:(c.querySelector('.quiz-card__feedback')||{}).innerText||'',
        all:c.innerText
      };
    });
    const before = await snap();
    if(!before || !before.all.trim()) break;
    const opts = await p.$$('.quiz__card--active .quiz-multiple-choice-option__input, .quiz__card--active .quiz-multiple-choice-option');
    if(opts.length){ try{ await opts[Math.min(PICK,opts.length-1)].click({timeout:3000}); }catch(e){} }
    await p.waitForTimeout(700);
    try{ await p.click('.quiz__card--active .quiz-card__submit, .quiz__card--active button:has-text("SUBMIT")',{timeout:5000}); }catch(e){}
    await p.waitForTimeout(1800);
    const after = await snap();
    out.push({q:q+1, picked:PICK, before, after});
    console.log('--- Q'+(q+1)+' ---'); console.log(before.all.replace(/\n+/g,' | ').slice(0,400));
    console.log('FEEDBACK:', (after&&after.feedback||'').replace(/\n+/g,' | ').slice(0,250));
    let advanced=false;
    for(const sel of ['.quiz__card--active .quiz-card__button--next','.quiz__card--active button:has-text("NEXT")','.quiz-card__next button']){
      try{ await p.click(sel,{timeout:3000}); advanced=true; break; }catch(e){}
    }
    if(!advanced) break;
    await p.waitForTimeout(2000);
  }
  const results = await p.evaluate(()=>{const r=document.querySelector('.quiz-results');return r?r.innerText:'';});
  console.log('RESULTS:', results.replace(/\n+/g,' | ').slice(0,300));
  fs.writeFileSync(`data/${MOD}_quiz_pick${PICK}.json`, JSON.stringify(out,null,1));
  await b.close();
})();
