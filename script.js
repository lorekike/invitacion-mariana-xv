const $=s=>document.querySelector(s), $$=s=>[...document.querySelectorAll(s)];
const intro=$('#intro'), site=$('#site'), openBtn=$('#openBtn'), song=$('#song'), musicToggle=$('#musicToggle'), toast=$('#toast');
let musicAvailable=true;
function showToast(msg){if(!toast)return;toast.textContent=msg;toast.classList.add('show');setTimeout(()=>toast.classList.remove('show'),2800)}

function removeIntro(){
  if(!intro)return;
  intro.classList.add('opened');
  intro.style.pointerEvents='none';
  intro.setAttribute('aria-hidden','true');
  setTimeout(()=>{
    intro.style.display='none';
    intro.remove();
  },150);
}

if(openBtn){
  openBtn.addEventListener('click',async e=>{
    e.preventDefault();
    e.stopPropagation();
    openBtn.disabled=true;
    intro.classList.add('opening');
    site.setAttribute('aria-hidden','false');
    document.body.classList.remove('locked');
    musicToggle.classList.add('show');
    try{
      await song.play();
      musicToggle.classList.add('playing');
    }catch(err){
      musicAvailable=false;
    }
    setTimeout(removeIntro,1200);
    setTimeout(()=>window.scrollTo({top:0,left:0,behavior:'auto'}),1250);
  },{passive:false});
}

// Respaldo para navegadores móviles que interrumpen la transición
if(intro){
  intro.addEventListener('transitionend',e=>{
    if(intro.classList.contains('opening') && (e.target===intro || e.target.classList.contains('curtain'))){
      removeIntro();
    }
  });
}

if(musicToggle){
  musicToggle.addEventListener('click',async()=>{
    if(!musicAvailable && !song.currentSrc){showToast('Falta el archivo de música');return}
    if(song.paused){try{await song.play();musicToggle.classList.add('playing')}catch(e){showToast('No se encontró el archivo de música')}}else{song.pause();musicToggle.classList.remove('playing')}
  });
}

// Fondo estrellado con parallax suave
const canvas=$('#sky'),ctx=canvas.getContext('2d');let stars=[],w=0,h=0,dpr=1,mx=.5,my=.5;
function resize(){dpr=Math.min(devicePixelRatio||1,2);w=innerWidth;h=innerHeight;canvas.width=w*dpr;canvas.height=h*dpr;canvas.style.width=w+'px';canvas.style.height=h+'px';ctx.setTransform(dpr,0,0,dpr,0,0);stars=Array.from({length:Math.min(230,Math.floor(w*h/4200))},()=>({x:Math.random()*w,y:Math.random()*h,r:Math.random()*1.45+.2,a:Math.random()*.75+.18,s:Math.random()*.012+.003,p:Math.random()*6.28}))}
addEventListener('resize',resize);addEventListener('pointermove',e=>{mx=e.clientX/w;my=e.clientY/h});resize();
function draw(t){ctx.clearRect(0,0,w,h);for(const s of stars){const tw=.55+.45*Math.sin(t*s.s+s.p);const x=s.x+(mx-.5)*s.r*10,y=s.y+(my-.5)*s.r*10;ctx.beginPath();ctx.fillStyle=`rgba(255,235,182,${s.a*tw})`;ctx.arc(x,y,s.r*tw,0,Math.PI*2);ctx.fill()}requestAnimationFrame(draw)}requestAnimationFrame(draw);

// Revelado al hacer scroll
const io=new IntersectionObserver(entries=>entries.forEach(e=>{if(e.isIntersecting){e.target.classList.add('visible');io.unobserve(e.target)}}),{threshold:.14});$$('.reveal').forEach(el=>io.observe(el));

// Cuenta regresiva
const target=new Date('2026-11-15T20:00:00-05:00').getTime();
function tick(){const d=Math.max(0,target-Date.now());$('#days').textContent=String(Math.floor(d/86400000)).padStart(3,'0');$('#hours').textContent=String(Math.floor(d/3600000)%24).padStart(2,'0');$('#minutes').textContent=String(Math.floor(d/60000)%60).padStart(2,'0');$('#seconds').textContent=String(Math.floor(d/1000)%60).padStart(2,'0')}tick();setInterval(tick,1000);

// Carrusel
const slides=$$('.slide'), dots=$('.dots');let current=0,timer;
if(slides.length&&dots){slides.forEach((_,i)=>{const b=document.createElement('button');b.setAttribute('aria-label',`Ver foto ${i+1}`);b.addEventListener('click',()=>go(i));dots.appendChild(b)});
function go(n){slides[current].classList.remove('active');dots.children[current].classList.remove('active');current=(n+slides.length)%slides.length;slides[current].classList.add('active');dots.children[current].classList.add('active');clearInterval(timer);timer=setInterval(()=>go(current+1),5200)}
const prev=$('.prev'),next=$('.next');if(prev)prev.onclick=()=>go(current-1);if(next)next.onclick=()=>go(current+1);go(0)}

// Calendario ICS
const calendarBtn=$('#calendarBtn');if(calendarBtn)calendarBtn.addEventListener('click',()=>{const ics=`BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//XV Mariana//ES\r\nBEGIN:VEVENT\r\nUID:mariana-xv-20261115@example.com\r\nDTSTAMP:20260714T190000Z\r\nDTSTART:20261116T010000Z\r\nDTEND:20261116T060000Z\r\nSUMMARY:XV años de Mariana Rojas Sierra\r\nLOCATION:Orquideorama, Av. 2 Norte # 48-10\r\nDESCRIPTION:Celebración de los XV años de Mariana. Código de vestuario: semi formal.\r\nEND:VEVENT\r\nEND:VCALENDAR`;const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([ics],{type:'text/calendar'}));a.download='XV-Mariana-Rojas-Sierra.ics';a.click();URL.revokeObjectURL(a.href)});

// RSVP WhatsApp
const dialog=$('#rsvpDialog'),form=$('#rsvpForm');
$$('.guest-card').forEach(b=>b.onclick=()=>{const type=b.dataset.type;$('#guestType').value=type;$('#formTitle').textContent=type==='school'?'Confirmación · Colegio':'Confirmación · Invitado adulto';$('#schoolFields').classList.toggle('hidden',type!=='school');$('#adultFields').classList.toggle('hidden',type!=='adult');dialog.showModal()});
const adultCount=$('#adultCount'),withCompanion=$('#withCompanion');if(adultCount)adultCount.onchange=e=>$('#adult2Label').classList.toggle('hidden',e.target.value!=='2');if(withCompanion)withCompanion.onchange=e=>$('#companionLabel').classList.toggle('hidden',e.target.value!=='Sí');
if(form)form.addEventListener('submit',e=>{e.preventDefault();const type=$('#guestType').value,name=$('#guestName').value.trim(),att=$('#attendance').value;if(!name||!att){showToast('Completa tu nombre y asistencia');return}let lines=[`✨ *Confirmación XV de Mariana*`,``,`Nombre: ${name}`,`Tipo de invitado: ${type==='school'?'Compañero(a) del colegio':'Invitado adulto'}`,`Asistencia: ${att}`];if(att==='Sí'&&type==='school'){lines.push(`Adultos acompañantes: ${$('#adultCount').value}`,`Adulto 1: ${$('#adult1').value.trim()||'Sin registrar'}`);if($('#adultCount').value==='2')lines.push(`Adulto 2: ${$('#adult2').value.trim()||'Sin registrar'}`)}if(att==='Sí'&&type==='adult'){lines.push(`Con acompañante: ${$('#withCompanion').value}`);if($('#withCompanion').value==='Sí')lines.push(`Acompañante: ${$('#companionName').value.trim()||'Sin registrar'}`)}const notes=$('#notes').value.trim();if(notes)lines.push(`Observaciones: ${notes}`);window.open(`https://wa.me/573016578609?text=${encodeURIComponent(lines.join('\n'))}`,'_blank','noopener');dialog.close()});

// Mariposas doradas volando suavemente por la invitación
function createButterflies(){
  if(document.querySelector('.butterfly-field'))return;
  const field=document.createElement('div');
  field.className='butterfly-field';
  field.setAttribute('aria-hidden','true');
  const total=innerWidth<=600?7:12;
  for(let i=0;i<total;i++){
    const butterfly=document.createElement('span');
    butterfly.className='butterfly';
    butterfly.style.setProperty('--size',`${22+Math.random()*24}px`);
    butterfly.style.setProperty('--duration',`${15+Math.random()*12}s`);
    butterfly.style.setProperty('--delay',`${-Math.random()*24}s`);
    butterfly.style.setProperty('--start-y',`${12+Math.random()*76}vh`);
    butterfly.style.setProperty('--end-y',`${8+Math.random()*70}vh`);
    butterfly.innerHTML=`<span class="butterfly-inner"><svg class="butterfly-svg" viewBox="0 0 120 90" role="presentation"><g class="svg-wing svg-wing-left"><path class="wing-main" d="M58 44C45 8 9 2 8 25c-1 18 18 24 35 25-13 5-25 18-17 29 9 12 27-7 34-24z"/><path class="wing-mark" d="M45 39C35 17 17 14 15 27c-2 9 13 13 30 12M43 57c-10 4-17 12-11 17 6 4 14-7 18-17"/><circle cx="25" cy="28" r="4"/></g><g class="svg-wing svg-wing-right"><path class="wing-main" d="M62 44c13-36 49-42 50-19 1 18-18 24-35 25 13 5 25 18 17 29-9 12-27-7-34-24z"/><path class="wing-mark" d="M75 39c10-22 28-25 30-12 2 9-13 13-30 12M77 57c10 4 17 12 11 17-6 4-14-7-18-17"/><circle cx="95" cy="28" r="4"/></g><path class="antenna" d="M58 21C48 8 41 11 39 4M62 21C72 8 79 11 81 4"/><ellipse class="svg-body" cx="60" cy="48" rx="4.8" ry="27"/><circle class="svg-head" cx="60" cy="20" r="6"/></svg></span>`;
    field.appendChild(butterfly);
  }
  document.body.appendChild(field);
}
createButterflies();
