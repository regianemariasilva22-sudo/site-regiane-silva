/* ============================================================
   REGIANE SILVA — efeitos cinematográficos compartilhados
   ============================================================ */

/* ── LOADER ── */
(function(){
  const bar=document.getElementById('loader-bar'), ldr=document.getElementById('loader');
  if(!ldr) return;
  let p=0;
  const t=setInterval(()=>{
    p+=Math.random()*16;
    if(p>=100){ p=100; clearInterval(t); setTimeout(()=>{ ldr.classList.add('out'); setTimeout(()=>ldr.style.display='none',650); },200); }
    if(bar) bar.style.width=p+'%';
  },60);
})();

/* ── PETALS / PARTICLE BG ── */
(function(){
  const c=document.getElementById('bg-canvas');
  if(!c) return;
  const x=c.getContext('2d');
  function rsz(){ c.width=innerWidth; c.height=innerHeight; } rsz();
  window.addEventListener('resize',rsz);
  const pts=Array.from({length:70},()=>({
    px:Math.random(),py:Math.random(),r:Math.random()*1.1+.3,
    a:Math.random()*.28+.06,vx:(Math.random()-.5)*.00012,vy:(Math.random()-.5)*.00012,
    ph:Math.random()*Math.PI*2
  }));
  (function draw(){
    x.clearRect(0,0,c.width,c.height);
    const W=c.width,H=c.height;
    pts.forEach(p=>{ p.px+=p.vx; p.py+=p.vy; p.ph+=.011;
      if(p.px<0)p.px=1;if(p.px>1)p.px=0;if(p.py<0)p.py=1;if(p.py>1)p.py=0;
      const op=p.a*(.65+.35*Math.sin(p.ph));
      x.beginPath();x.arc(p.px*W,p.py*H,p.r,0,Math.PI*2);x.fillStyle=`rgba(122,42,59,${op})`;x.fill();
    });
    for(let i=0;i<pts.length;i++) for(let j=i+1;j<pts.length;j++){
      const dx=(pts[i].px-pts[j].px)*W,dy=(pts[i].py-pts[j].py)*H,d=Math.sqrt(dx*dx+dy*dy);
      if(d<130){ x.beginPath();x.moveTo(pts[i].px*W,pts[i].py*H);x.lineTo(pts[j].px*W,pts[j].py*H);
        x.strokeStyle=`rgba(122,42,59,${.05*(1-d/130)})`;x.lineWidth=.4;x.stroke(); }
    }
    requestAnimationFrame(draw);
  })();
})();

/* ── GRAIN ── */
(function(){
  const c=document.getElementById('grain-canvas');
  if(!c) return;
  const g=c.getContext('2d');
  function rsz(){ c.width=innerWidth; c.height=innerHeight; } rsz();
  window.addEventListener('resize',rsz);
  (function grain(){ const d=g.createImageData(c.width,c.height),b=d.data;
    for(let i=0;i<b.length;i+=4){const v=Math.random()*255;b[i]=b[i+1]=b[i+2]=v;b[i+3]=255;}
    g.putImageData(d,0,0); setTimeout(()=>requestAnimationFrame(grain),100); })();
})();

/* ── CURSOR ── */
(function(){
  const dot=document.getElementById('cdot'),ring=document.getElementById('cring');
  if(!dot||!ring) return;
  let mx=0,my=0,rx=0,ry=0;
  document.addEventListener('mousemove',e=>{
    mx=e.clientX;my=e.clientY;
    dot.style.left=mx+'px';dot.style.top=my+'px';
    document.body.style.setProperty('--smx',mx+'px');
    document.body.style.setProperty('--smy',my+'px');
  });
  (function raf(){ rx+=(mx-rx)*.11;ry+=(my-ry)*.11;ring.style.left=rx+'px';ring.style.top=ry+'px';requestAnimationFrame(raf); })();
  function bindHover(){
    document.querySelectorAll('a,button,.glass-card,.method-card,.ladder-card,.tried-pill,.pillar-chip,.comm-card,input,.mem-list-item,.belief-col,.timeline-card,.id-pill,.quote-pill,.check-item,.quiz-opt').forEach(el=>{
      el.addEventListener('mouseenter',()=>ring.classList.add('hov'));
      el.addEventListener('mouseleave',()=>ring.classList.remove('hov'));
    });
  }
  bindHover();
  window.__rebindCursor = bindHover;
  document.addEventListener('mousedown',()=>ring.classList.add('clk'));
  document.addEventListener('mouseup',()=>ring.classList.remove('clk'));
})();

/* ── SCROLL REVEALS ── */
(function(){
  const io=new IntersectionObserver(entries=>entries.forEach(e=>{if(e.isIntersecting)e.target.classList.add('in');}),{threshold:.08,rootMargin:'0px 0px -24px 0px'});
  document.querySelectorAll('.reveal,.slide-l,.slide-r').forEach(el=>io.observe(el));
})();

/* ── CARD SPOTLIGHT + TILT ── */
document.querySelectorAll('.glass-card,.method-card,.ladder-card,.belief-col').forEach(card=>{
  card.style.transition='border-color .4s,box-shadow .4s,transform .12s ease';
  card.addEventListener('mousemove',e=>{
    const r=card.getBoundingClientRect();
    card.style.setProperty('--bx',(e.clientX-r.left)+'px');
    card.style.setProperty('--by',(e.clientY-r.top)+'px');
    const x=(e.clientX-r.left)/r.width-.5,y=(e.clientY-r.top)/r.height-.5;
    card.style.transform=`perspective(900px) rotateY(${x*5}deg) rotateX(${-y*5}deg) translateZ(4px)`;
  });
  card.addEventListener('mouseleave',()=>card.style.transform='perspective(900px) rotateY(0) rotateX(0) translateZ(0)');
});

/* ── MAGNETIC BUTTONS ── */
document.querySelectorAll('.btn,.nav-cta').forEach(el=>{
  el.addEventListener('mousemove',e=>{ const r=el.getBoundingClientRect();el.style.transform=`translate(${(e.clientX-r.left-r.width/2)*.2}px,${(e.clientY-r.top-r.height/2)*.2}px)`; });
  el.addEventListener('mouseleave',()=>el.style.transform='');
});

/* ── RIPPLE ── */
document.addEventListener('click',e=>{
  const r=document.createElement('div');r.className='ripple-fx';
  Object.assign(r.style,{left:e.clientX+'px',top:e.clientY+'px',width:'6px',height:'6px',background:'rgba(122,42,59,0.22)',transform:'translate(-50%,-50%)'});
  document.body.appendChild(r);
  r.animate([{transform:'translate(-50%,-50%) scale(0)',opacity:.8},{transform:'translate(-50%,-50%) scale(20)',opacity:0}],{duration:700,easing:'ease-out'}).onfinish=()=>r.remove();
});

/* ── HEADING SHIMMER ── */
(function(){
  const s=document.createElement('style');
  s.textContent='@keyframes h-shimmer{0%{background-position:-220% center}100%{background-position:220% center}}.h-shimmer{background:linear-gradient(105deg,#7A2A3B 30%,#C6899A 44%,#F1DCDF 50%,#C6899A 56%,#7A2A3B 70%);background-size:260% auto;-webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent;animation:h-shimmer 1.1s cubic-bezier(.4,0,.2,1) forwards;}';
  document.head.appendChild(s);
  document.querySelectorAll('h1,h2,.card-title,.method-card h3,.ladder-card h3').forEach(el=>{
    el.addEventListener('mouseenter',()=>el.classList.add('h-shimmer'));
    el.addEventListener('mouseleave',()=>el.classList.remove('h-shimmer'));
    el.addEventListener('animationend',()=>el.classList.remove('h-shimmer'));
  });
})();

/* ── PHOTO TILT (sobre) ── */
(function(){
  const frame=document.querySelector('.photo-frame'), border=document.querySelector('.photo-border');
  if(!frame||!border) return;
  frame.addEventListener('mousemove',e=>{
    const r=frame.getBoundingClientRect();
    const x=(e.clientX-r.left)/r.width-.5, y=(e.clientY-r.top)/r.height-.5;
    border.style.transform=`rotateY(${x*10}deg) rotateX(${-y*10}deg)`;
  });
  frame.addEventListener('mouseleave',()=>border.style.transform='rotateY(0) rotateX(0)');
})();

/* ── HERO PARALLAX ── */
(function(){
  const h=document.querySelector('.hero .h1');
  if(h) window.addEventListener('scroll',()=>{ h.style.transform=`translateY(${scrollY*.05}px)`; },{passive:true});
})();

/* ── NAV SCROLL + MOBILE TOGGLE ── */
(function(){
  const nav=document.querySelector('nav');
  if(nav) window.addEventListener('scroll',()=>{ nav.style.borderBottomColor=scrollY>50?'rgba(122,42,59,.14)':'rgba(122,42,59,.07)'; },{passive:true});
  const toggle=document.querySelector('.nav-toggle'), links=document.querySelector('.nav-links');
  if(toggle&&links) toggle.addEventListener('click',()=>links.classList.toggle('open'));
})();

/* ── LADDER DRAG-SCROLL (desktop) ── */
(function(){
  const track=document.querySelector('.ladder-track');
  if(!track) return;
  let down=false,startX,scrollLeft;
  track.addEventListener('mousedown',e=>{ down=true; startX=e.pageX-track.offsetLeft; scrollLeft=track.scrollLeft; });
  track.addEventListener('mouseleave',()=>down=false);
  track.addEventListener('mouseup',()=>down=false);
  track.addEventListener('mousemove',e=>{
    if(!down) return; e.preventDefault();
    const x=e.pageX-track.offsetLeft, walk=(x-startX)*1.2;
    track.scrollLeft=scrollLeft-walk;
  });

  /* auto-advance, one card at a time, pauses on hover/drag/touch */
  let paused=false, autoTimer;
  function cardStep(){
    const card=track.querySelector('.ladder-card');
    return card ? card.getBoundingClientRect().width + 20 : 300;
  }
  function autoNext(){
    if(paused) return;
    const max=track.scrollWidth-track.clientWidth;
    if(track.scrollLeft>=max-4){ track.scrollTo({left:0,behavior:'smooth'}); }
    else{ track.scrollBy({left:cardStep(),behavior:'smooth'}); }
  }
  autoTimer=setInterval(autoNext,3200);
  track.addEventListener('mouseenter',()=>paused=true);
  track.addEventListener('mouseleave',()=>{ paused=false; down=false; });
  track.addEventListener('touchstart',()=>paused=true,{passive:true});
  track.addEventListener('touchend',()=>setTimeout(()=>paused=false,2000));
})();

/* ── TESTIMONIAL CAROUSEL ── */
(function(){
  const wrap=document.querySelector('.testi-carousel');
  if(!wrap) return;
  const slides=wrap.querySelectorAll('.testi-slide');
  const dots=wrap.parentElement.querySelectorAll('.testi-dot');
  let i=0, timer;
  function show(n){
    slides.forEach(s=>s.classList.remove('active'));
    dots.forEach(d=>d.classList.remove('active'));
    i=(n+slides.length)%slides.length;
    slides[i].classList.add('active');
    if(dots[i]) dots[i].classList.add('active');
  }
  function next(){ show(i+1); }
  function start(){ timer=setInterval(next,2200); }
  function stop(){ clearInterval(timer); }
  dots.forEach((d,idx)=>d.addEventListener('click',()=>{ show(idx); stop(); start(); }));
  wrap.addEventListener('mouseenter',stop);
  wrap.addEventListener('mouseleave',start);
  show(0);
  start();
})();

/* ── ACTIVE SECTION NUMBER (index page) ── */
(function(){
  const sections=document.querySelectorAll('section[data-step]');
  if(!sections.length) return;
  const io=new IntersectionObserver(entries=>entries.forEach(e=>{
    if(e.isIntersecting){ document.querySelectorAll('.sec-divider .num').forEach(n=>{}); }
  }),{threshold:.4});
  sections.forEach(s=>io.observe(s));
})();
