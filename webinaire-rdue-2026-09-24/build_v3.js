const pptxgen = require('pptxgenjs');
const React = require('react'); const RDS = require('react-dom/server'); const sharp = require('sharp');
const lu = require('react-icons/lu'); const gi = require('react-icons/gi'); const fs = require('fs');

// ---------- Charte ----------
const C = { BLUE:'00439A', NAVY:'0A2A5E', DEEP:'071E45', PINK:'E6005F', INK:'1F2A3A', GREY:'5B6B7F', MUTE:'8A98A8',
            LIGHT:'F3F6FB', LINE:'D9E1EC', WHITE:'FFFFFF', GREEN:'1E8E5A', GREEN_L:'E6F5EE', RED:'D0313B', RED_L:'FBE9EA',
            AMBER:'B7791F', AMBER_L:'FFF6E0', BLUE_L:'E8EFFA', TEAL:'0E8A8A' };
const FONT='Calibri'; const W=13.333, H=7.5; const TOTAL=34;
const pres = new pptxgen(); pres.layout='LAYOUT_WIDE'; pres.author='CCI Lyon Métropole · CF² Douane'; pres.title='Maîtriser le RDUE';
const iconCache={};
async function icon(name, color='FFFFFF', size=256){
  const key=name+color; if(iconCache[key]) return iconCache[key];
  const Comp = lu[name]||gi[name]; if(!Comp) throw new Error('icon '+name);
  const svg = RDS.renderToStaticMarkup(React.createElement(Comp,{color:'#'+color,size}));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return iconCache[key]='image/png;base64,'+buf.toString('base64');
}
const logoCCI='image/png;base64,'+fs.readFileSync('logo_cci.png').toString('base64');
const logoCF2='image/png;base64,'+fs.readFileSync('logo_cf2.png').toString('base64');
const sh=()=>({type:'outer',color:'8AA0B4',blur:4,offset:1,angle:90,opacity:0.22});
const R=(t,o={})=>Object.assign({text:t},{options:o});

// ---------- Helpers ----------
let n=0;
function slide(dark=false){ n++; const s=pres.addSlide(); s.background={color: dark?C.NAVY:C.WHITE}; s._n=n; return s; }
function chrome(s,{dark=false}={}){
  // logo top-right, identical on every slide
  if(dark){ s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:10.55,y:0.32,w:2.35,h:0.62,fill:{color:C.WHITE},line:{color:C.WHITE},rectRadius:0.06}); }
  s.addImage({data:logoCCI,x:10.65,y:0.38,w:2.15,h:0.48});
  s.addText('RDUE / EUDR  ·  Webinaire CCI Lyon Métropole  ·  24 septembre 2026',{x:0.6,y:7.02,w:8,h:0.3,fontFace:FONT,fontSize:9,color:dark?'9FB3D1':C.MUTE,isTextBox:true,margin:0});
  s.addText(`${s._n} / ${TOTAL}`,{x:11.7,y:7.02,w:1.1,h:0.3,fontFace:FONT,fontSize:9,color:dark?'9FB3D1':C.MUTE,align:'right',isTextBox:true,margin:0});
}
function title(s,t,sub){
  s.addText(t,{x:0.6,y:0.42,w:9.7,h:0.7,fontFace:FONT,fontSize:28,bold:true,color:C.NAVY,isTextBox:true,margin:0,valign:'middle'});
  if(sub) s.addText(sub,{x:0.6,y:1.1,w:11.5,h:0.4,fontFace:FONT,fontSize:14,color:C.GREY,isTextBox:true,margin:0,valign:'middle'});
}
function card(s,x,y,w,h,o={}){ s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x,y,w,h,fill:{color:o.fill||C.WHITE},line:{color:o.line||C.LINE,width:o.lw||0.75},rectRadius:o.r||0.08,shadow:o.noShadow?undefined:sh()}); }
function pill(s,x,y,w,h,txt,fill,o={}){ s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x,y,w,h,fill:{color:fill},line:{color:fill},rectRadius:0.5});
  s.addText(txt,{x,y,w,h,fontFace:FONT,fontSize:o.size||10,bold:true,color:o.color||C.WHITE,align:'center',valign:'middle',isTextBox:true,margin:0}); }
async function iconCircle(s,x,y,d,name,fill=C.BLUE,fg='FFFFFF'){ s.addShape(pres.shapes.OVAL,{x,y,w:d,h:d,fill:{color:fill},line:{color:fill}}); s.addImage({data:await icon(name,fg),x:x+d*0.25,y:y+d*0.25,w:d*0.5,h:d*0.5}); }
function txt(s,t,x,y,w,h,o={}){ s.addText(t,Object.assign({x,y,w,h,fontFace:FONT,fontSize:13,color:C.INK,isTextBox:true,margin:0,valign:'top'},o)); }
function bullets(s,items,x,y,w,h,o={}){ const arr=items.map((t,i)=>({text:t,options:{bullet:{indent:12},breakLine:i<items.length-1,paraSpaceAfter:o.gap||4}}));
  s.addText(arr,Object.assign({x,y,w,h,fontFace:FONT,fontSize:12.5,color:C.INK,isTextBox:true,valign:'top',margin:0},o)); }
const PARTS=[['IN','Contexte & enjeux'],['I','Êtes-vous concernés ?'],['II','La diligence raisonnée'],['III','Mise en œuvre & formalités douanières'],['IV','Bonnes pratiques & conseils']];
function divider(idx,label,time,bul){
  const s=slide(true); chrome(s,{dark:true});
  s.addShape(pres.shapes.OVAL,{x:-2.2,y:3.9,w:7.5,h:7.5,fill:{color:C.BLUE,transparency:78},line:{color:C.BLUE,transparency:78}});
  txt(s,`${label==='IN'?'INTRODUCTION':'PARTIE '+label}   ·   ${time}`,0.8,1.5,6,0.4,{fontSize:13,bold:true,color:C.PINK,charSpacing:2});
  txt(s,PARTS[idx][1],0.8,1.95,7.2,1.2,{fontSize:34,bold:true,color:C.WHITE,valign:'middle'});
  bullets(s,bul,0.85,3.35,6.8,2.6,{fontSize:15,color:'DCE6F5',gap:8});
  // nav
  PARTS.forEach(([k,t],i)=>{ const y=1.55+i*0.95; const cur=i===idx;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:8.6,y,w:4.1,h:0.72,fill:{color:cur?C.PINK:'123A78'},line:{color:cur?C.PINK:'123A78'},rectRadius:0.08});
    txt(s,k,8.75,y,0.55,0.72,{fontSize:14,bold:true,color:C.WHITE,valign:'middle',align:'center'});
    txt(s,t,9.4,y,3.2,0.72,{fontSize:cur?12.5:11.5,bold:cur,color:cur?C.WHITE:'C7D4E8',valign:'middle'}); });
  return s;
}

(async()=>{
// ============ 1. TITRE ============
{ const s=slide(true);
  s.addShape(pres.shapes.OVAL,{x:8.6,y:-2.6,w:8,h:8,fill:{color:C.BLUE,transparency:70},line:{color:C.BLUE,transparency:70}});
  s.addShape(pres.shapes.OVAL,{x:10.8,y:4.2,w:5,h:5,fill:{color:C.PINK,transparency:82},line:{color:C.PINK,transparency:82}});
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:0.8,y:0.7,w:3.3,h:0.9,fill:{color:C.WHITE},line:{color:C.WHITE},rectRadius:0.06});
  s.addImage({data:logoCCI,x:0.95,y:0.8,w:3.0,h:0.67});
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:4.3,y:0.7,w:0.9,h:0.9,fill:{color:C.WHITE},line:{color:C.WHITE},rectRadius:0.06});
  s.addImage({data:logoCF2,x:4.35,y:0.75,w:0.8,h:0.8});
  txt(s,'WEBINAIRE  ·  DOUANE & RSE',0.8,2.15,8,0.4,{fontSize:13,bold:true,color:C.PINK,charSpacing:2});
  txt(s,'Maîtriser le RDUE et sécuriser vos opérations douanières',0.8,2.55,8.6,1.7,{fontSize:40,bold:true,color:C.WHITE,valign:'middle'});
  txt(s,'Règlement européen contre la déforestation (UE) 2023/1115, modifié par le règlement (UE) 2025/2650 — de la parcelle au passage en douane',0.8,4.3,8.4,0.7,{fontSize:15,color:'C7D4E8'});
  for(const [i,[nm,role]] of [['Céline Fontana','CF² Douane — Consultante & formatrice en douane'],['Houda Zebar','CCI Lyon Métropole — Co-animatrice du webinaire']].entries()){
    const x=0.8+i*4.05; s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x,y:5.35,w:3.85,h:0.95,fill:{color:'123A78'},line:{color:'123A78'},rectRadius:0.08});
    txt(s,nm,x+0.2,5.42,3.5,0.4,{fontSize:15,bold:true,color:C.WHITE}); txt(s,role,x+0.2,5.8,3.5,0.45,{fontSize:11,color:'C7D4E8'}); }
  txt(s,'En distanciel  ·  jeudi 24 septembre 2026  ·  1 h 30',0.8,6.6,8,0.4,{fontSize:13,color:'9FB3D1'});
}
// ============ 2. SOMMAIRE ============
{ const s=slide(); chrome(s); title(s,'Sommaire','Le déroulé du webinaire — 1 h 30, une introduction et quatre parties.');
  const rows=[['IN','Introduction — Contexte & enjeux','Urgence écologique · objectifs du règlement · le texte 2023/1115 modifié par 2025/2650','10 min'],
    ['I','Êtes-vous concernés ?','7 matières & dérivés (annexe I) · ce qui change en 2026-2027 · calendrier · définitions · acteurs & responsabilités','20 min'],
    ['II','La diligence raisonnée','Le système de diligence (SDR) · les 3 étapes · la diligence simplifiée · le réexamen annuel','30 min'],
    ['III','Mise en œuvre & formalités douanières','Système d\'information & EU Login · déclarer en douane (Delta) : C716 / C717 / Y · transmission amont-aval','20 min'],
    ['IV','Bonnes pratiques & conseils','Risques & sanctions · leviers · feuille de route · quiz','10 min']];
  rows.forEach(([k,t,d,m],i)=>{ const y=1.75+i*1.02; card(s,0.6,y,12.1,0.88);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:0.78,y:y+0.14,w:0.9,h:0.6,fill:{color:C.BLUE},line:{color:C.BLUE},rectRadius:0.08});
    txt(s,k,0.78,y+0.14,0.9,0.6,{fontSize:16,bold:true,color:C.WHITE,align:'center',valign:'middle'});
    txt(s,t,1.9,y+0.1,8.5,0.38,{fontSize:16,bold:true,color:C.NAVY}); txt(s,d,1.9,y+0.48,8.6,0.35,{fontSize:11.5,color:C.GREY});
    pill(s,10.9,y+0.25,1.55,0.38,m,C.BLUE_L,{color:C.BLUE,size:11}); });
}
// ============ 3. OBJECTIFS ============
{ const s=slide(); chrome(s); title(s,'Objectifs du webinaire','Ce que vous saurez faire en sortant.');
  const items=[['LuMapPin','Vous situer','Déterminer si votre entreprise est concernée, et à quel titre : opérateur, opérateur en aval ou commerçant.'],
    ['LuSearch','Identifier vos produits','Les retrouver dans l\'annexe I et connaître les bons codes douaniers.'],
    ['LuShieldCheck','Maîtriser la diligence','Le système de diligence, les trois étapes, les preuves à réunir.'],
    ['LuRocket','Passer à l\'action','Système d\'information, déclaration en douane, échéances, contacts utiles.']];
  for(const [i,[ic,t,d]] of items.entries()){ const x=0.6+(i%2)*6.15, y=1.85+Math.floor(i/2)*2.3; card(s,x,y,5.95,2.0);
    await iconCircle(s,x+0.3,y+0.35,0.8,ic); txt(s,t,x+1.35,y+0.35,4.4,0.45,{fontSize:18,bold:true,color:C.NAVY}); txt(s,d,x+1.35,y+0.85,4.4,1.0,{fontSize:13,color:C.GREY}); }
  txt(s,'Co-animation Céline Fontana & Houda Zebar  ·  cas pratiques et quiz tout au long du webinaire',0.6,6.5,12,0.4,{fontSize:12,italic:true,color:C.PINK,align:'center'});
}
// ============ 4. DIVIDER INTRO ============
divider(0,'IN','10 min',['L\'urgence écologique : le constat en chiffres','Les objectifs du RDUE','Le texte : (UE) 2023/1115 modifié par (UE) 2025/2650, et le report d\'un an']);
// ============ 5. COÛT DE L'INACTION ============
{ const s=slide(); chrome(s); title(s,'Le coût de l\'inaction','Pourquoi ce règlement ? L\'urgence écologique en chiffres.');
  const st=[['420 M','hectares de forêts disparus entre 1990 et 2020','LuTrees'],['10 M','hectares supplémentaires perdus chaque année','LuTreePine'],['11 %','des émissions mondiales de gaz à effet de serre liées à la déforestation','LuDroplet']];
  for(const [i,[v,l,ic]] of st.entries()){ const x=0.6+i*4.1; card(s,x,1.85,3.9,3.1); await iconCircle(s,x+0.35,2.15,0.7,ic,C.BLUE_L,C.BLUE);
    txt(s,v,x+0.35,2.95,3.3,0.9,{fontSize:44,bold:true,color:C.BLUE,valign:'middle'}); txt(s,l,x+0.35,3.9,3.3,0.9,{fontSize:13,color:C.GREY}); }
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:0.6,y:5.3,w:12.1,h:1.1,fill:{color:C.NAVY},line:{color:C.NAVY},rectRadius:0.08});
  txt(s,'Le RDUE vise à réduire la contribution de la consommation européenne à la déforestation importée et à promouvoir des chaînes d\'approvisionnement durables.',0.9,5.3,11.5,1.1,{fontSize:15,color:C.WHITE,valign:'middle'});
  txt(s,'Source : brochure DGDDI / Douane — Règlement européen contre la déforestation ; FAO, Global Forest Resources Assessment 2020.',0.6,6.55,12,0.3,{fontSize:10,color:C.MUTE});
}
// ============ 6. RDUE EN BREF ============
{ const s=slide(); chrome(s); title(s,'Le RDUE en bref','Un texte, une interdiction, une logique d\'accès au marché.');
  card(s,0.6,1.85,7.6,4.75);
  const blocks=[['Le texte','Règlement (UE) 2023/1115 du 31 mai 2023, modifié par le règlement (UE) 2025/2650 du 19 décembre 2025 (JOUE du 23/12/2025, en vigueur le 26/12/2025). Il abroge le règlement bois RBUE (995/2010) à la date d\'application.'],
    ['L\'interdiction','Interdire la mise sur le marché de l\'UE — et l\'exportation depuis l\'UE — de produits ayant contribué à la déforestation ou à la dégradation des forêts après le 31/12/2020.'],
    ['La logique','Ce n\'est pas un label « vert » : c\'est une condition d\'accès au marché européen. Reporté d\'un an, le calendrier est désormais confirmé par la Commission (revue du 4 mai 2026).']];
  blocks.forEach(([h,t],i)=>{ const y=2.1+i*1.5; txt(s,h,0.95,y,6.9,0.35,{fontSize:15,bold:true,color:C.BLUE}); txt(s,t,0.95,y+0.38,6.9,1.05,{fontSize:12.5,color:C.INK}); });
  const tiles=[['31/12/2020','DATE BUTOIR','Aucune déforestation postérieure n\'est admise.',C.NAVY],['30/12/2026','APPLICATION','Grandes & moyennes entreprises. Micro & petites : 30/06/2027.',C.BLUE]];
  tiles.forEach(([d,l,t,col],i)=>{ const y=1.85+i*2.45; s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:8.5,y,w:4.2,h:2.3,fill:{color:col},line:{color:col},rectRadius:0.08});
    txt(s,d,8.5,y+0.3,4.2,0.7,{fontSize:32,bold:true,color:C.WHITE,align:'center'}); txt(s,l,8.5,y+1.0,4.2,0.35,{fontSize:11,bold:true,color:C.PINK,align:'center',charSpacing:2}); txt(s,t,8.8,y+1.4,3.6,0.8,{fontSize:12,color:'DCE6F5',align:'center'}); });
}
// ============ 7. DIVIDER I ============
divider(1,'I','20 min',['Les produits visés : 7 matières + dérivés (annexe I) et ce qui change','Le calendrier consolidé','Définitions clés : zéro déforestation, légalité, DDR','Acteurs & responsabilités : opérateur, aval, commerçant']);
// ============ 8. 7 MATIÈRES ============
{ const s=slide(); chrome(s); title(s,'Les 7 matières premières — annexe I','Un produit est concerné s\'il figure à l\'annexe I ET contient l\'une de ces matières.');
  const prods=[['Café','NC 0901','GiCoffeeBeans'],['Cacao','NC 1801 → 1806','GiChocolateBar'],['Caoutchouc','NC 4001 → 4017','GiTyre'],['Soja','NC 1201 · 1507','GiPlantSeed'],['Huile de palme','NC 1511 · 1513','GiPalmTree'],['Bœuf','NC 0102 · 0201 · 0202','GiCow'],['Bois & papier','NC 4401 → 4421 · ch. 47-48','GiWoodPile']];
  for(const [i,[nm,nc,ic]] of prods.entries()){ const x=0.6+(i%4)*3.05, y=1.8+Math.floor(i/4)*2.25; card(s,x,y,2.85,2.05);
    await iconCircle(s,x+0.95,y+0.25,0.95,ic,C.BLUE_L,C.BLUE); txt(s,nm,x,y+1.28,2.85,0.38,{fontSize:15,bold:true,color:C.NAVY,align:'center'}); txt(s,nc,x,y+1.62,2.85,0.3,{fontSize:10.5,color:C.GREY,align:'center'}); }
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:9.75,y:4.05,w:2.95,h:2.05,fill:{color:C.AMBER_L},line:{color:'F0D9A0'},rectRadius:0.08});
  txt(s,[R('Attention aux dérivés : ',{bold:true,color:C.AMBER}),R('papier, charbon de bois, pneus, chocolat, meubles… vous font entrer dans le champ. Le cuir en sort (acte délégué du 13/07/2026).',{color:C.INK})],9.95,4.15,2.6,1.9,{fontSize:11.5});
  txt(s,'Codes NC indicatifs — la liste exhaustive figure à l\'annexe I du règlement.',0.6,6.45,12,0.3,{fontSize:10,color:C.MUTE});
}
// ============ 9. ANNEXE I ============
{ const s=slide(); chrome(s); title(s,'Annexe I : ce qui change en 2026-2027','Acte délégué adopté le 13 juillet 2026 — exclusions dès l\'entrée en vigueur (publication au JOUE attendue mi-septembre 2026), ajouts au 30/12/2027.');
  const cols=[['Désormais soumis (au 30/12/2027)',C.GREEN,C.GREEN_L,'LuCircleCheck',['Café soluble (SH 2101 11 00)','Certains dérivés d\'huile de palme, acides oléiques, savons','Langues de bovins congelées']],
    ['Désormais hors champ (dès l\'entrée en vigueur)',C.RED,C.RED_L,'LuCircleX',['Cuirs et peaux de bovins (SH 4101, 4104, 4107)','Pneus rechapés, courroies, certains articles en caoutchouc','Soja de semence','Sièges d\'avion et de véhicules','Produits imprimés — déjà sortis par le règlement 2025/2650']]];
  for(const [i,[h,col,bg,ic,items]] of cols.entries()){ const x=0.6+i*6.15; card(s,x,1.85,5.95,4.3);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x,y:1.85,w:5.95,h:0.85,fill:{color:col},line:{color:col},rectRadius:0.08});
    s.addShape(pres.shapes.RECTANGLE,{x,y:2.4,w:5.95,h:0.3,fill:{color:col},line:{color:col}});
    s.addImage({data:await icon(ic),x:x+0.3,y:2.05,w:0.45,h:0.45}); txt(s,h,x+0.95,1.85,4.9,0.85,{fontSize:16,bold:true,color:C.WHITE,valign:'middle'});
    bullets(s,items,x+0.35,2.95,5.3,3.0,{fontSize:13.5,gap:8}); }
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:0.6,y:6.3,w:12.1,h:0.5,fill:{color:C.AMBER_L},line:{color:'F0D9A0'},rectRadius:0.08});
  txt(s,[R('Conséquence : ',{bold:true,color:C.AMBER}),R('l\'audit de vos nomenclatures n\'est pas un exercice ponctuel — rejouez-le après publication de l\'acte délégué.',{color:C.INK})],0.85,6.3,11.6,0.5,{fontSize:12,valign:'middle'});
}
// ============ 10. CALENDRIER ============
{ const s=slide(); chrome(s); title(s,'Le calendrier consolidé','Reporté d\'un an par le règlement (UE) 2025/2650 — et confirmé par la revue de simplification du 4 mai 2026.');
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:0.6,y:1.75,w:12.1,h:0.95,fill:{color:C.NAVY},line:{color:C.NAVY},rectRadius:0.08});
  await iconCircle(s,0.85,1.9,0.65,'LuCalendar',C.PINK);
  txt(s,[R('Report d\'un an, pas suppression. ',{bold:true,color:C.WHITE}),R('L\'échéance 2025 est décalée à 2026/2027 ; la date butoir du 31/12/2020 est inchangée et la Commission ne rouvrira pas le texte.',{color:'DCE6F5'})],1.7,1.75,10.8,0.95,{fontSize:13.5,valign:'middle'});
  const jal=[['26/12/2025','Modificatif en vigueur','Règl. 2025/2650 : report + simplifications',C.MUTE,true],['fin juin 2026','Système d\'information rouvert','Dépôt des DDR possible (ex-TRACES)',C.MUTE,true],['13/07/2026','Acte délégué annexe I','+ règl. d\'exécution 2026/1565 (SI)',C.MUTE,true],
    ['30/12/2026','Application — socle','Grandes & moyennes entreprises',C.PINK,false],['30/06/2027','Micro & petites entreprises','Fin du report TPE/PE',C.BLUE,false],['30/12/2027','Nouveaux produits','Ajouts de l\'annexe I',C.BLUE,false]];
  s.addShape(pres.shapes.LINE,{x:0.9,y:3.55,w:11.5,h:0,line:{color:C.LINE,width:2,dashType:'dash'}});
  jal.forEach(([d,t,l,col,past],i)=>{ const x=0.6+i*2.03; s.addShape(pres.shapes.OVAL,{x:x+0.75,y:3.38,w:0.36,h:0.36,fill:{color:col},line:{color:C.WHITE,width:2}});
    card(s,x,3.95,1.88,2.45,{line:past?C.LINE:col,lw:past?0.75:1.5}); txt(s,d,x,4.1,1.88,0.45,{fontSize:15,bold:true,color:past?C.GREY:C.NAVY,align:'center'});
    txt(s,t,x+0.1,4.6,1.68,0.7,{fontSize:12.5,bold:true,color:past?C.GREY:col,align:'center'}); txt(s,l,x+0.1,5.35,1.68,0.95,{fontSize:11.5,color:C.GREY,align:'center'}); });
  txt(s,'Jalons passés en gris, échéances à venir en couleur. Acte délégué : publication au JOUE à vérifier la veille du webinaire.',0.6,6.6,12,0.3,{fontSize:10.5,color:C.MUTE,italic:true});
}
// ============ 11. TROIS CONDITIONS ============
{ const s=slide(); chrome(s); title(s,'Trois conditions cumulatives (art. 3)','Il en manque une : le produit ne peut être ni mis sur le marché ni exporté.');
  const cs=[['LuLeaf','Zéro déforestation','Production sur des terres non déboisées après le 31/12/2020. Pour le bois : pas de dégradation forestière après cette date.','art. 3 a)'],['LuScale','Légalité','Conformité à la législation du pays de production : foncier, environnement, travail, droits des peuples autochtones, fiscalité, douane.','art. 3 b)'],['LuFileText','Déclaration (DDR)','Couvert par une déclaration de diligence raisonnée déposée dans le système d\'information.','art. 3 c)']];
  for(const [i,[ic,h,t,a]] of cs.entries()){ const x=0.6+i*4.1; card(s,x,1.85,3.9,3.9); await iconCircle(s,x+1.5,2.15,0.9,ic); txt(s,h,x+0.3,3.2,3.3,0.45,{fontSize:18,bold:true,color:C.NAVY,align:'center'}); txt(s,t,x+0.35,3.7,3.2,1.5,{fontSize:12.5,color:C.GREY,align:'center'}); pill(s,x+1.35,5.2,1.2,0.32,a,C.BLUE,{size:10}); }
  txt(s,'Un produit légal mais issu d\'une parcelle déforestée après 2020 reste interdit — et le volet export est le plus souvent oublié.',0.6,6.05,12.1,0.5,{fontSize:13,italic:true,color:C.PINK,align:'center'});
}
// ============ 12. ENTRANT / SORTANT ============
{ const s=slide(); chrome(s); title(s,'Qui est concerné ? Produit entrant, produit sortant','Une fois le produit dans le champ (annexe I + matière) : que se passe-t-il selon le flux ?');
  const LX=0.6,LW=5.6,RX=7.1,RW=5.6,AX=6.35;
  pill(s,LX+1.3,1.62,3,0.38,'PRODUIT ENTRANT',C.NAVY,{size:11.5}); pill(s,RX+1.3,1.62,3,0.38,'PRODUIT SORTANT',C.NAVY,{size:11.5});
  const rows=[[['Importation sur le marché de l\'UE','Vous = opérateur · DDR + n° en douane (C716)',C.GREEN,'DDR requise'],['Exportation depuis le marché de l\'UE','Vous = opérateur · DDR avant l\'export',C.GREEN,'DDR requise'],true],
    [['Acheté sur le marché UE et utilisé pour un produit RDUE sortant','Opérateur en aval · pas de DDR, conserve le n° du fournisseur',C.BLUE,'N° de DDR à conserver'],['Vendu au sein du marché UE (« mise à disposition »)','Commerçant · références fournisseur conservées',C.BLUE,'N° transmis par le fournisseur'],true],
    [['Acheté sur le marché UE et PAS utilisé pour un produit RDUE sortant','Produit non relevant',C.RED,'Pas de DDR'],['Exporté ou vendu sur le marché de l\'UE','Produit non relevant',C.RED,'Pas de DDR'],true],
    [['Acheté sur le marché UE — consommation interne','Produit non relevant',C.RED,'Pas de DDR'],['Hors obligation','Aucune DDR, aucun n° à conserver',C.RED,'Pas de DDR'],false]];
  let y=2.12; const RH=1.0;
  for(const [L,Rr,arr] of rows){ for(const [X,Wd,[t,d,col,lab]] of [[LX,LW,L],[RX,RW,Rr]]){ card(s,X,y,Wd,RH,{line:col===C.GREEN?C.GREEN:C.LINE,lw:col===C.GREEN?1.5:0.75});
      txt(s,t,X+0.2,y+0.1,Wd-0.4,0.4,{fontSize:12.5,bold:true,color:C.NAVY}); txt(s,d,X+0.2,y+0.5,Wd-2.4,0.4,{fontSize:10.5,color:col===C.RED?C.RED:C.GREY,bold:col===C.RED}); pill(s,X+Wd-2.25,y+0.58,2.05,0.28,lab,col,{size:9}); }
    if(arr) s.addShape(pres.shapes.RIGHT_ARROW,{x:AX,y:y+0.33,w:0.4,h:0.34,fill:{color:C.NAVY},line:{color:C.NAVY}}); else txt(s,'✕',AX,y+0.25,0.4,0.5,{fontSize:18,bold:true,color:C.MUTE,align:'center'});
    y+=RH+0.1; }
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:0.6,y:6.48,w:12.1,h:0.42,fill:{color:C.AMBER_L},line:{color:'F0D9A0'},rectRadius:0.08});
  txt(s,[R('Régimes douaniers particuliers : ',{bold:true,color:C.AMBER}),R('entrepôt douanier, perfectionnement actif, admission temporaire… ne sont pas soumis au RDUE (FAQ de la Commission).',{color:C.INK})],0.85,6.48,11.6,0.42,{fontSize:11,valign:'middle'});
}
// ============ 13. ACTEURS ============
{ const s=slide(); chrome(s); title(s,'Acteurs & responsabilités','Amont ou aval, le rôle — flux par flux — fixe le niveau d\'obligation (définitions art. 2).');
  const cols=[['LuShip','Opérateur (amont)','art. 2 · art. 4 à 11','CHARGE MAXIMALE',C.NAVY,['Met sur le marché UE en premier, ou exporte.','Diligence raisonnée complète (art. 8 à 11).','Dépose la DDR dans le système d\'information.','N° de référence en douane (C716).']],
    ['LuFactory','Opérateur en aval','art. 2 · règl. 2025/2650','CHARGE ALLÉGÉE',C.BLUE,['Transforme un produit déjà mis sur le marché UE.','Pas de diligence, pas de DDR à déposer.','1er aval : collecte et conserve le n° de DDR amont (5 ans).','Avals suivants : informations fournisseur seulement.']],
    ['LuStore','Commerçant','art. 2 · art. 5','CHARGE MINIMALE',C.TEAL,['Met à disposition sans transformer.','1er commerçant aval : conserve le n° de DDR.','Suivants : nom, adresse, site du fournisseur.','Enregistrement au système d\'information si non-PME.']]];
  for(const [i,[ic,h,art,ch,col,items]] of cols.entries()){ const x=0.6+i*4.1; card(s,x,1.85,3.9,4.55);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x,y:1.85,w:3.9,h:0.95,fill:{color:col},line:{color:col},rectRadius:0.08}); s.addShape(pres.shapes.RECTANGLE,{x,y:2.5,w:3.9,h:0.3,fill:{color:col},line:{color:col}});
    s.addImage({data:await icon(ic),x:x+0.3,y:2.08,w:0.5,h:0.5}); txt(s,h,x+1.0,1.85,2.8,0.95,{fontSize:17,bold:true,color:C.WHITE,valign:'middle'});
    pill(s,x+0.9,3.0,2.1,0.3,art,C.BLUE_L,{color:C.BLUE,size:9.5}); txt(s,ch,x,3.4,3.9,0.3,{fontSize:10,bold:true,color:col,align:'center',charSpacing:1.5});
    bullets(s,items,x+0.3,3.85,3.35,2.5,{fontSize:12,gap:5}); }
  txt(s,'Le mandataire (art. 6), établi dans l\'UE, peut déposer la DDR au nom de l\'opérateur — qui reste responsable.',0.6,6.55,12.1,0.35,{fontSize:12,italic:true,color:C.GREY,align:'center'});
}
// ============ 14. CAS PRATIQUES RÔLES ============
{ const s=slide(); chrome(s); title(s,'Cas pratiques : quel rôle pour qui ?','Sondage — opérateur, opérateur en aval ou commerçant ?');
  const cs=[['LuShip','Un négociant importe des fèves de cacao du Ghana et les revend à un chocolatier français.','Opérateur',C.NAVY],['LuFactory','Un fabricant de meubles achète du contreplaqué déjà déclaré (DDR déposée) et fabrique des meubles.','Opérateur en aval',C.BLUE],['LuStore','Une enseigne achète en France des meubles déjà couverts et les revend en magasin.','Commerçant',C.TEAL]];
  for(const [i,[ic,t,a,col]] of cs.entries()){ const y=1.9+i*1.45; card(s,0.6,y,12.1,1.25); await iconCircle(s,0.9,y+0.28,0.7,ic,C.BLUE_L,C.BLUE); txt(s,`Cas ${i+1}`,0.85,y+1.0,0.8,0.22,{fontSize:9,bold:true,color:C.MUTE,align:'center'});
    txt(s,t,1.9,y,7.6,1.25,{fontSize:14,color:C.INK,valign:'middle'}); pill(s,9.9,y+0.38,2.5,0.5,a,col,{size:12}); }
  txt(s,'Réponses à révéler après le vote (animation ou masquage).',0.6,6.35,12,0.3,{fontSize:10.5,italic:true,color:C.MUTE});
}
// ============ 15. FAQ ============
{ const s=slide(); chrome(s); title(s,'Précisions clés — FAQ de la Commission (v5, avril 2026)','Ce que la 5ᵉ édition de la FAQ confirme ou clarifie. Numéros de question à contrôler sur la version publiée.');
  const it=[['LuShoppingCart','Le e-commerce est couvert','Ventes en ligne B2B et B2C, places de marché : opérateur, aval ou commerçant selon le rôle. L\'usage privé entre particuliers est exclu.','FAQ 3.17 à 3.19'],['LuBan','Pas de dilution possible','Un lot mélangé avec une part non conforme, non isolable, devient non conforme en entier.','FAQ 1.4 & 1.5'],['LuLayers','Produits composés','Produit multi-matières : diligence sur chaque matière de base en cause.','FAQ 1.3'],['LuUsers','Un compte, plusieurs rôles','Dans le système d\'information, un même compte porte opérateur, aval, commerçant et mandataire.','FAQ 7.29']];
  for(const [i,[ic,h,t,f]] of it.entries()){ const x=0.6+(i%2)*6.15, y=1.85+Math.floor(i/2)*2.35; card(s,x,y,5.95,2.15); await iconCircle(s,x+0.3,y+0.3,0.7,ic); txt(s,h,x+1.2,y+0.32,4.5,0.45,{fontSize:16,bold:true,color:C.NAVY}); txt(s,t,x+1.2,y+0.8,4.5,1.0,{fontSize:12.5,color:C.GREY}); pill(s,x+1.2,y+1.7,1.6,0.28,f,C.BLUE_L,{color:C.BLUE,size:9}); }
}
// ============ 16. DIVIDER II ============
divider(2,'II','30 min',['Le système de diligence raisonnée (SDR)','La démarche en 3 étapes : recueil, évaluation, atténuation','La diligence simplifiée (pays à risque faible)','Le réexamen annuel du SDR']);
// ============ 17. SDR ============
{ const s=slide(); chrome(s); title(s,'Le système de diligence raisonnée (SDR)','Une obligation permanente, pas un acte ponctuel (art. 12).');
  card(s,0.6,1.85,7.6,4.5);
  txt(s,'Ce que c\'est',0.95,2.1,6.9,0.4,{fontSize:16,bold:true,color:C.BLUE}); txt(s,'Un ensemble de procédures et de mesures internes qui prouvent la légalité et l\'absence de déforestation de vos produits, et qui restent auditables.',0.95,2.5,6.9,0.9,{fontSize:13,color:C.INK});
  txt(s,'Ce qu\'il impose',0.95,3.55,6.9,0.4,{fontSize:16,bold:true,color:C.BLUE});
  bullets(s,['Documenter la démarche et conserver les preuves 5 ans.','Réexaminer le SDR au moins une fois par an.','Publier un rapport annuel pour les entreprises non-PME (art. 12 § 3).','Intégrer des clauses RDUE dans les contrats fournisseurs.'],0.95,3.95,6.9,2.3,{fontSize:13,gap:6});
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:8.5,y:1.85,w:4.2,h:4.5,fill:{color:C.NAVY},line:{color:C.NAVY},rectRadius:0.08});
  await iconCircle(s,10.1,2.3,1.0,'LuRefreshCw',C.PINK); txt(s,'Réexamen\nau moins 1 fois par an',8.7,3.5,3.8,0.9,{fontSize:18,bold:true,color:C.WHITE,align:'center'});
  txt(s,'Recueil → Évaluation → Atténuation,\nrevus régulièrement.',8.7,4.5,3.8,0.8,{fontSize:12.5,color:'DCE6F5',align:'center'}); pill(s,9.9,5.55,1.4,0.32,'ART. 12',C.PINK,{size:10});
}
// ============ 18. 3 ÉTAPES ============
{ const s=slide(); chrome(s); title(s,'La diligence raisonnée en 3 étapes','Articles 8 à 11 du règlement.');
  const st=[['LuSearch','Recueil d\'informations','Produit, quantité, pays, fournisseurs, géolocalisation des parcelles, dates de production, preuves de légalité.','art. 9'],['LuChartColumn','Évaluation du risque','Selon la zone de production, la complexité de la chaîne, la présence de forêts et de peuples autochtones.','art. 10'],['LuShieldCheck','Atténuation du risque','Mesures correctives, audits ou enquêtes indépendantes si un risque non négligeable est identifié.','art. 11']];
  for(const [i,[ic,h,t,a]] of st.entries()){ const x=0.6+i*4.1; card(s,x,1.85,3.9,3.3); await iconCircle(s,x+0.3,2.15,0.8,ic); txt(s,String(i+1),x+3.0,2.05,0.7,0.8,{fontSize:40,bold:true,color:C.LINE,align:'right'});
    txt(s,h,x+0.3,3.1,3.3,0.45,{fontSize:17,bold:true,color:C.NAVY}); txt(s,t,x+0.3,3.55,3.3,1.2,{fontSize:12.5,color:C.GREY}); pill(s,x+0.3,4.7,0.9,0.3,a,C.BLUE,{size:9.5});
    if(i<2) s.addShape(pres.shapes.RIGHT_ARROW,{x:x+3.95,y:3.3,w:0.15,h:0.3,fill:{color:C.LINE},line:{color:C.LINE}}); }
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:0.6,y:5.45,w:12.1,h:1.1,fill:{color:C.NAVY},line:{color:C.NAVY},rectRadius:0.08}); await iconCircle(s,0.85,5.62,0.75,'LuMapPin',C.PINK);
  txt(s,[R('Le vrai chantier : la donnée amont. ',{bold:true,color:C.WHITE}),R('La géolocalisation des parcelles et les preuves dépendent d\'un fournisseur hors UE qui n\'a aucune obligation directe — c\'est l\'effort le plus lourd, à lancer dès maintenant.',{color:'DCE6F5'})],1.8,5.45,10.7,1.1,{fontSize:13,valign:'middle'});
}
// ============ 19. PAYS ============
{ const s=slide(); chrome(s); title(s,'La diligence simplifiée : le rôle du pays','Le pays de production module l\'intensité de la diligence (benchmarking, art. 29) — liste du 22/05/2025, révision annoncée pour 2026.');
  const rows=[['140','RISQUE FAIBLE','pays, dont toute l\'UE, le Royaume-Uni, les États-Unis, la Chine…','Diligence simplifiée : seule la collecte d\'informations (étape 1) est exigée — l\'avantage stratégique.','1 %',C.GREEN],['~50','RISQUE STANDARD','pays — Brésil, Indonésie, Malaisie, Côte d\'Ivoire…','Diligence raisonnée complète : les trois étapes sont obligatoires.','3 %',C.AMBER],['4','RISQUE ÉLEVÉ','Biélorussie · Corée du Nord · Birmanie · Russie','Vigilance renforcée et contrôles accrus.','9 %',C.RED]];
  rows.forEach(([nb,lab,p,d,ctrl,col],i)=>{ const y=1.85+i*1.5; card(s,0.6,y,12.1,1.35); s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:0.6,y,w:2.3,h:1.35,fill:{color:col},line:{color:col},rectRadius:0.08}); s.addShape(pres.shapes.RECTANGLE,{x:2.6,y,w:0.3,h:1.35,fill:{color:col},line:{color:col}});
    txt(s,nb,0.6,y+0.12,2.3,0.7,{fontSize:34,bold:true,color:C.WHITE,align:'center'}); txt(s,lab,0.6,y+0.85,2.3,0.35,{fontSize:10,bold:true,color:C.WHITE,align:'center',charSpacing:1.5});
    txt(s,p,3.2,y+0.15,6.6,0.4,{fontSize:13.5,bold:true,color:C.NAVY}); txt(s,d,3.2,y+0.6,6.6,0.7,{fontSize:12,color:C.GREY});
    txt(s,'Contrôles',10.3,y+0.2,2.2,0.3,{fontSize:10,color:C.MUTE,align:'center',charSpacing:1.5}); txt(s,ctrl,10.3,y+0.45,2.2,0.7,{fontSize:26,bold:true,color:col,align:'center'}); });
  txt(s,'Taux de contrôle minimaux des autorités (art. 16) : part des opérateurs — et des quantités pour le risque élevé. Le mélange en cours de chaîne avec des produits d\'origine inconnue fait retomber dans le régime complet.',0.6,6.4,12.1,0.5,{fontSize:10.5,color:C.MUTE,italic:true});
}
// ============ 20. LOGIGRAMME ============
{ const s=slide(); chrome(s); title(s,'Le parcours de diligence raisonnée','Le parcours obligatoire, article par article (art. 8 à 13).');
  card(s,0.6,1.65,12.1,1.2); pill(s,0.8,1.78,0.4,0.4,'1',C.NAVY,{size:13}); txt(s,'Collecte des informations',1.35,1.75,5,0.45,{fontSize:15,bold:true,color:C.NAVY}); pill(s,11.4,1.82,1.1,0.3,'ARTICLE 9',C.RED,{size:9});
  bullets(s,['Description, nom commercial, type de produit','Quantités et pays (zones) de production','Fournisseurs et clients (nom, adresse)'],0.9,2.22,5.6,0.6,{fontSize:11,gap:1});
  bullets(s,['Géolocalisation de toutes les parcelles','Date ou période de production','Preuves de légalité et de « zéro déforestation »'],6.7,2.22,5.8,0.6,{fontSize:11,gap:1});
  s.addShape(pres.shapes.DOWN_ARROW,{x:8.75,y:2.92,w:0.35,h:0.3,fill:{color:C.NAVY},line:{color:C.NAVY}});
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:5.4,y:3.28,w:7.3,h:0.75,fill:{color:C.GREEN},line:{color:C.GREEN},rectRadius:0.15});
  txt(s,[R('Pays de production à risque standard ou élevé ?\n',{bold:true,fontSize:13,color:C.WHITE}),R('classification des pays — benchmarking, art. 29',{fontSize:10,color:'D7EBDD'})],5.5,3.28,7.1,0.75,{align:'center',valign:'middle'});
  txt(s,'NON',4.85,3.33,0.55,0.3,{fontSize:10,bold:true,color:C.GREEN,align:'center'}); s.addShape(pres.shapes.LEFT_ARROW,{x:4.95,y:3.58,w:0.4,h:0.3,fill:{color:C.GREEN},line:{color:C.GREEN}});
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:0.6,y:3.28,w:4.2,h:2.35,fill:{color:C.BLUE_L},line:{color:C.BLUE,dashType:'dash',width:1},rectRadius:0.08});
  pill(s,0.8,3.42,1.9,0.28,'pays à risque faible',C.WHITE,{color:C.BLUE,size:9}); txt(s,'Diligence raisonnée simplifiée',0.8,3.8,3.8,0.4,{fontSize:14,bold:true,color:C.NAVY});
  txt(s,[R('Seule l\'étape 1 est exigée. ',{bold:true}),R('L\'opérateur n\'est pas tenu à l\'évaluation ni à l\'atténuation du risque (art. 10 et 11).',{})],0.8,4.2,3.8,1.0,{fontSize:11.5,color:C.INK}); pill(s,3.5,5.2,1.1,0.3,'ARTICLE 13',C.RED,{size:9});
  txt(s,'OUI',5.4,4.08,0.6,0.3,{fontSize:10,bold:true,color:C.GREEN,align:'center'}); s.addShape(pres.shapes.DOWN_ARROW,{x:8.75,y:4.08,w:0.35,h:0.3,fill:{color:C.GREEN},line:{color:C.GREEN}});
  for(const [j,[num,h,a,col,items]] of [['2','Évaluation du risque','ARTICLE 10',C.BLUE,['Vérifier et analyser les informations (critères de l\'art. 10)','Réexamen au moins une fois par an','Pas de mise sur le marché si le risque n\'est pas nul ou négligeable']],['3','Atténuation du risque','ARTICLE 11',C.AMBER,['Informations, données ou documents complémentaires','Audits ou enquêtes indépendants','Procédures et contrôles proportionnés']]].entries()){
    const x=5.4+j*3.7; card(s,x,4.45,3.6,1.9,{line:col,lw:1.5}); pill(s,x+0.15,4.55,0.34,0.34,num,col,{size:11}); txt(s,h,x+0.6,4.52,2.9,0.4,{fontSize:13,bold:true,color:C.NAVY}); bullets(s,items,x+0.15,4.95,3.3,1.1,{fontSize:10,gap:1}); pill(s,x+2.45,6.0,1.0,0.26,a,C.RED,{size:8.5}); }
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:5.4,y:6.5,w:3.6,h:0.42,fill:{color:C.GREEN_L},line:{color:C.GREEN},rectRadius:0.3}); txt(s,'Risque nul ou négligeable → DDR',5.4,6.5,3.6,0.42,{fontSize:11,bold:true,color:C.GREEN,align:'center',valign:'middle'});
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:9.1,y:6.5,w:3.6,h:0.42,fill:{color:C.RED_L},line:{color:C.RED},rectRadius:0.3}); txt(s,'Risque non négligeable → STOP (art. 3)',9.1,6.5,3.6,0.42,{fontSize:11,bold:true,color:C.RED,align:'center',valign:'middle'});
  s.addShape(pres.shapes.DOWN_ARROW,{x:2.5,y:5.72,w:0.35,h:0.3,fill:{color:C.BLUE},line:{color:C.BLUE}});
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:0.6,y:6.1,w:4.2,h:0.82,fill:{color:C.NAVY},line:{color:C.NAVY},rectRadius:0.08});
  txt(s,[R('Déclaration de diligence raisonnée\n',{bold:true,fontSize:12,color:C.WHITE}),R('déposée dans le SI avant mise sur le marché ou export → n° de DDR (C716) · art. 4 & 12',{fontSize:9.5,color:'DCE6F5'})],0.75,6.1,3.95,0.82,{valign:'middle'});
}
// ============ 21. DIVIDER III ============
divider(3,'III','20 min',['Le système d\'information RDUE & le compte EU Login','La déclaration de diligence raisonnée (DDR) → n° de référence','Déclarer en douane (Delta) : codes C716 / C717 / Y','La transmission amont / aval du n° de DDR']);
// ============ 22. SYSTÈME D'INFORMATION ============
{ const s=slide(); chrome(s); title(s,'Le système d\'information RDUE & EU Login','Enregistrer la DDR pour obtenir un numéro de référence — rouvert fin juin 2026, formations de la Commission en septembre.');
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:0.6,y:1.85,w:4.3,h:4.7,fill:{color:C.NAVY},line:{color:C.NAVY},rectRadius:0.08}); await iconCircle(s,2.25,2.2,1.0,'LuKey',C.PINK);
  txt(s,'Prérequis',0.8,3.35,3.9,0.45,{fontSize:18,bold:true,color:C.WHITE,align:'center'});
  bullets(s,['Un compte EU Login par personne.','Un numéro EORI actif.','L\'enregistrement de l\'entreprise dans le système d\'information (obligatoire pour les opérateurs et les non-PME aval).'],1.0,3.85,3.6,2.2,{fontSize:12,color:'DCE6F5',gap:5});
  pill(s,1.9,6.05,1.7,0.32,'dès maintenant',C.PINK,{size:10});
  card(s,5.2,1.85,7.5,4.7); txt(s,'La DDR dans le système d\'information — à saisir',5.5,2.05,7,0.4,{fontSize:15,bold:true,color:C.NAVY});
  const f=['N° EORI de l\'opérateur','Code du système harmonisé (SH)','Description libre (nom commercial ou scientifique)','Quantité (mise sur le marché / export)','Pays de production','Géolocalisation de toutes les parcelles','Attestation de risque nul ou négligeable','Signature de la personne habilitée'];
  for(const [i,t] of f.entries()){ const x=5.5+(i%2)*3.6, y=2.6+Math.floor(i/2)*0.62; s.addImage({data:await icon('LuCircleCheck',C.GREEN),x,y:y+0.05,w:0.28,h:0.28}); txt(s,t,x+0.4,y,3.1,0.4,{fontSize:12,color:C.INK,valign:'middle'}); }
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:5.5,y:5.25,w:6.9,h:0.95,fill:{color:C.GREEN},line:{color:C.GREEN},rectRadius:0.08});
  txt(s,[R('→ Génération du numéro de référence DDR et du code de vérification. ',{bold:true,color:C.WHITE}),R('Une DDR citée en douane ne peut plus être modifiée ni retirée.',{color:'E6F5EE'})],5.7,5.25,6.5,0.95,{fontSize:12.5,valign:'middle'});
  txt(s,'Nouveautés du règlement d\'exécution (UE) 2026/1565 : déclaration simplifiée pour les micro/petits producteurs, groupage de déclarations, API pour les dépôts en volume.',0.6,6.7,12.1,0.3,{fontSize:10.5,color:C.MUTE,italic:true});
}
// ============ 23. DÉCLARER EN DOUANE ============
{ const s=slide(); chrome(s); title(s,'Déclarer en douane (Delta)','Le numéro de DDR se reporte dans la déclaration douanière, avant mainlevée.');
  card(s,0.6,1.85,6.0,2.55); txt(s,'Cas n° 1 — produit concerné',0.85,2.0,5.5,0.4,{fontSize:14,bold:true,color:C.NAVY});
  pill(s,0.85,2.5,1.1,0.42,'C716',C.NAVY,{size:13}); txt(s,'Opérateur : code C716 + son numéro de DDR.',2.1,2.5,4.3,0.42,{fontSize:12.5,valign:'middle'});
  pill(s,0.85,3.1,1.1,0.42,'C717',C.BLUE,{size:13}); txt(s,'PME aval (DDR amont existante) : C717 + n° de DDR du fournisseur — à confirmer sur la fiche douane 2026.',2.1,3.05,4.3,0.75,{fontSize:11.5,valign:'middle'});
  card(s,6.8,1.85,5.9,2.55); txt(s,'Cas n° 2 — exemptions (codes Y)',7.05,2.0,5.5,0.4,{fontSize:14,bold:true,color:C.NAVY});
  const ys=[['Y129','hors champ'],['Y132','fabriqué avant le 29/06/2023'],['Y133','déchets / recyclé'],['Y141','micro & petite ent. (report)'],['Y142','activité non commerciale']];
  ys.forEach(([c,l],i)=>{ const x=7.05+(i%2)*2.85, y=2.5+Math.floor(i/2)*0.6; pill(s,x,y,0.8,0.36,c,C.LIGHT,{color:C.NAVY,size:10.5}); txt(s,l,x+0.9,y,1.9,0.36,{fontSize:10.5,color:C.GREY,valign:'middle'}); });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:0.6,y:4.6,w:12.1,h:1.1,fill:{color:C.NAVY},line:{color:C.NAVY},rectRadius:0.08}); await iconCircle(s,0.85,4.78,0.75,'LuLink',C.PINK);
  txt(s,[R('Transmission amont / aval. ',{bold:true,color:C.WHITE}),R('L\'opérateur transmet le n° de DDR et le code de vérification à son seul client aval immédiat ; les maillons suivants ne conservent que les informations fournisseur.',{color:'DCE6F5'})],1.8,4.6,10.7,1.1,{fontSize:13,valign:'middle'});
  txt(s,'Autres cas : n° conventionnel 99EU9999999999 pour l\'export ou la réimportation de produits fabriqués entre le 29/06/2023 et la date d\'application. Delta = téléservice de dédouanement de la DGDDI.',0.6,5.9,12.1,0.6,{fontSize:11,color:C.GREY});
}
// ============ 24. HUILE DE PALME ============
{ const s=slide(); chrome(s); title(s,'Cas pratique : l\'huile de palme','Le même produit peut être soumis… ou hors champ, selon l\'usage final déclaré.');
  card(s,0.6,1.85,5.95,4.5); s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:0.6,y:1.85,w:5.95,h:0.85,fill:{color:C.GREEN},line:{color:C.GREEN},rectRadius:0.08}); s.addShape(pres.shapes.RECTANGLE,{x:0.6,y:2.4,w:5.95,h:0.3,fill:{color:C.GREEN},line:{color:C.GREEN}});
  txt(s,'A · Scénario standard — soumis au RDUE',0.9,1.85,5.5,0.85,{fontSize:15,bold:true,color:C.WHITE,valign:'middle'});
  txt(s,[R('Le flux\n',{bold:true,color:C.NAVY}),R('Importation d\'huile de palme brute ou de dérivés pour la fabrication de savons ou d\'additifs alimentaires.\n\n',{}),R('La conséquence\n',{bold:true,color:C.NAVY}),R('Diligence raisonnée complète, dépôt de la DDR dans le système d\'information, puis code C716 + n° de DDR en douane.',{})],0.9,2.95,5.4,3.2,{fontSize:12.5,color:C.INK});
  card(s,6.75,1.85,5.95,4.5); s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:6.75,y:1.85,w:5.95,h:0.85,fill:{color:C.GREY},line:{color:C.GREY},rectRadius:0.08}); s.addShape(pres.shapes.RECTANGLE,{x:6.75,y:2.4,w:5.95,h:0.3,fill:{color:C.GREY},line:{color:C.GREY}});
  txt(s,'B · Exclusions — hors champ',7.05,1.85,5.5,0.85,{fontSize:15,bold:true,color:C.WHITE,valign:'middle'});
  txt(s,[R('Flux 1 · pharmaceutique\n',{bold:true,color:C.NAVY}),R('Dérivés oléochimiques utilisés exclusivement pour la fabrication de médicaments → code Y129.\n\n',{}),R('Flux 2 · biocarburants\n',{bold:true,color:C.NAVY}),R('Huiles usagées et déchets récupérés (matière en fin de vie) → code Y133.\n\n',{}),R('Aucune DDR requise.',{bold:true,color:C.RED})],7.05,2.95,5.4,3.2,{fontSize:12.5,color:C.INK});
  txt(s,'Deux flux physiquement identiques peuvent recevoir un traitement opposé : l\'usage réel et le code déterminent le champ.',0.6,6.55,12.1,0.35,{fontSize:12.5,italic:true,color:C.PINK,align:'center'});
}
// ============ 25. DIVIDER IV ============
divider(4,'IV','10 min',['Cartographier les risques et connaître les sanctions','Les leviers pour réduire la charge','La feuille de route en quatre marches','Quiz et synthèse']);
// ============ 26. RISQUES ============
{ const s=slide(); chrome(s); title(s,'Cartographier les risques','Le RDUE est pluriel : ce n\'est pas qu\'un sujet d\'image.');
  const it=[['LuScale','Juridique','Diligence insuffisante, rôle mal qualifié, DDR incomplète.'],['LuShip','Douanier','Refus de mainlevée, mauvais code, lots bloqués à la frontière.'],['LuEuro','Financier','Amende (plafond ≥ 4 % du CA UE), confiscation, exclusion des marchés publics.'],['LuTruck','Opérationnel','Fournisseurs sans données, ruptures d\'approvisionnement.'],['LuEye','RSE / réputation','Exigences des clients et des financeurs, exposition médiatique.']];
  for(const [i,[ic,h,t]] of it.entries()){ const x=0.6+(i%3)*4.1, y=1.85+Math.floor(i/3)*2.3; card(s,x,y,3.9,2.05); await iconCircle(s,x+0.3,y+0.3,0.7,ic); txt(s,h,x+1.2,y+0.35,2.5,0.45,{fontSize:16,bold:true,color:C.NAVY}); txt(s,t,x+0.3,y+1.05,3.3,0.95,{fontSize:12.5,color:C.GREY}); }
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:8.8,y:4.15,w:3.9,h:2.05,fill:{color:C.NAVY},line:{color:C.NAVY},rectRadius:0.08});
  txt(s,'Une déclaration en douane engage la responsabilité juridique de l\'opérateur sur le risque zéro déforestation (art. 25).',9.1,4.15,3.3,2.05,{fontSize:13,color:C.WHITE,valign:'middle'});
}
// ============ 27. SANCTIONS ============
{ const s=slide(); chrome(s); title(s,'Le régime des sanctions','Fixé par chaque État membre dans le cadre de l\'article 25 du règlement.');
  const it=[['LuEuro','Financière','Plafond d\'amende d\'au moins 4 % du chiffre d\'affaires annuel réalisé dans l\'UE.',C.RED],['LuPackageX','Logistique','Immobilisation et confiscation des produits concernés et des revenus tirés de la transaction.',C.AMBER],['LuBan','Commerciale','Exclusion temporaire (12 mois maximum) des marchés publics et des financements publics.',C.BLUE],['LuGavel','Pénale','Poursuites possibles dans le cadre de la directive (UE) 2024/1203 sur la criminalité environnementale.',C.NAVY]];
  for(const [i,[ic,h,t,col]] of it.entries()){ const x=0.6+i*3.05; card(s,x,1.85,2.85,3.5); await iconCircle(s,x+0.95,2.15,0.95,ic,col); txt(s,h,x,3.25,2.85,0.45,{fontSize:17,bold:true,color:C.NAVY,align:'center'}); txt(s,t,x+0.25,3.75,2.35,1.5,{fontSize:12,color:C.GREY,align:'center'}); }
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:0.6,y:5.6,w:12.1,h:0.9,fill:{color:C.NAVY},line:{color:C.NAVY},rectRadius:0.08});
  txt(s,[R('Taux de contrôle minimaux : ',{bold:true,color:C.WHITE}),R('1 % des opérateurs (risque faible)  ·  3 % (risque standard)  ·  9 % des opérateurs et des quantités (risque élevé).',{color:'DCE6F5'})],0.9,5.6,11.5,0.9,{fontSize:13.5,valign:'middle',align:'center'});
}
// ============ 28. LEVIERS ============
{ const s=slide(); chrome(s); title(s,'Les leviers pour réduire la charge','Bien joué, le RDUE se pilote.');
  const it=[['LuFactory','Statut d\'opérateur en aval','S\'appuyer sur la DDR de l\'amont plutôt que reconstituer la traçabilité.'],['LuGlobe','Sourcing pays à risque faible','Bascule vers la diligence simplifiée (art. 13).'],['LuHandshake','Mandataire / mutualisation','Déléguer la diligence (art. 6) ou mutualiser en filière.'],['LuCalendar','Exploiter le calendrier','Délai micro/petites entreprises (30/06/2027) ; nouveaux produits (30/12/2027).'],['LuFileCheck','Déclaration simplifiée','Déclaration unique pour les micro/petits producteurs de pays à risque faible.'],['LuScrollText','Transfert par contrat','Garanties, indemnisation, résiliation portées par les fournisseurs.']];
  for(const [i,[ic,h,t]] of it.entries()){ const x=0.6+(i%3)*4.1, y=1.85+Math.floor(i/3)*2.3; card(s,x,y,3.9,2.05); await iconCircle(s,x+0.3,y+0.3,0.7,ic); txt(s,h,x+1.2,y+0.3,2.6,0.75,{fontSize:14,bold:true,color:C.NAVY,valign:'middle'}); txt(s,t,x+0.3,y+1.1,3.3,0.9,{fontSize:12,color:C.GREY}); }
}
// ============ 29. DEUX PROFILS ============
{ const s=slide(); chrome(s); title(s,'Deux profils, deux stratégies','Des exemples d\'entreprises ayant structuré leur SDR (profils génériques).');
  const p2=[['LuStore','Le négociant en invendus','déstockage · catalogue multi-produits',C.AMBER,'Impact fort et compliqué','Sourcing opportuniste, traçabilité amont quasi impossible à reconstituer.',['Statut d\'aval : s\'appuyer sur les DDR existantes, exiger le n° à l\'achat.','Mobiliser les exclusions (Y132, Y133, Y129).','Filtrer le sourcing en amont.']],
    ['LuFactory','L\'industriel à diligence mature','filière caoutchouc · chaîne longue',C.NAVY,'Impact maîtrisable — enjeu d\'échelle','Diligence déjà structurée. Point dur : les petits planteurs (géolocalisation à grande échelle).',['Dossier de preuve opposable.','Sécuriser la donnée des petits producteurs.','Valoriser (ESG / CSRD) et cascader par contrat.']]];
  for(const [i,[ic,h,sub,col,imp,impd,strat]] of p2.entries()){ const x=0.6+i*6.15; card(s,x,1.85,5.95,4.7); s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x,y:1.85,w:5.95,h:1.0,fill:{color:col},line:{color:col},rectRadius:0.08}); s.addShape(pres.shapes.RECTANGLE,{x,y:2.55,w:5.95,h:0.3,fill:{color:col},line:{color:col}});
    s.addImage({data:await icon(ic),x:x+0.3,y:2.1,w:0.5,h:0.5}); txt(s,h,x+1.0,1.9,4.8,0.5,{fontSize:17,bold:true,color:C.WHITE}); txt(s,sub,x+1.0,2.38,4.8,0.35,{fontSize:11,color:'E6ECF5'});
    txt(s,imp,x+0.3,3.05,5.4,0.35,{fontSize:13,bold:true,color:col}); txt(s,impd,x+0.3,3.4,5.4,0.7,{fontSize:12,color:C.GREY}); txt(s,'Stratégie',x+0.3,4.15,5.4,0.35,{fontSize:13,bold:true,color:C.NAVY}); bullets(s,strat,x+0.3,4.5,5.4,1.9,{fontSize:12,gap:4}); }
}
// ============ 30. ESCALIER ============
{ const s=slide(); chrome(s); title(s,'L\'escalier de conformité','Quatre marches, dans cet ordre : on ne connecte pas les outils avant d\'avoir cartographié les flux.');
  const st=[['LuMap','Cartographier','Audit flash de vos flux : quels produits (codes SH), quels pays, quel statut (amont, aval, commerçant) ?'],['LuBuilding2','Structurer','Bâtir le SDR : procédures, référent, gouvernance, clauses fournisseurs, réexamen annuel.'],['LuArchive','Documenter','La traçabilité : géolocalisation, preuves, archivage 5 ans des évaluations et mesures d\'atténuation.'],['LuLogIn','Connecter','Les accès : EORI, comptes EU Login, enregistrement dans le système d\'information, codes douaniers.']];
  const BASE=6.55;
  for(const [i,[ic,h,t]] of st.entries()){ const x=0.6+i*3.05, bh=0.5+i*0.5, by=BASE-bh, ch=by-1.85-0.12;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x,y:by,w:2.85,h:bh,fill:{color:C.BLUE},line:{color:C.BLUE},rectRadius:0.05});
    txt(s,String(i+1),x+1.9,by+bh-0.6,0.8,0.55,{fontSize:26,bold:true,color:C.WHITE,align:'right',valign:'bottom'});
    card(s,x,1.85,2.85,ch); await iconCircle(s,x+0.25,2.1,0.7,ic);
    txt(s,h,x+0.25,2.95,2.4,0.45,{fontSize:17,bold:true,color:C.NAVY}); txt(s,t,x+0.25,3.4,2.4,ch-1.65,{fontSize:12.5,color:C.GREY}); }
}
// ============ 31. CONSEILS ============
{ const s=slide(); chrome(s); title(s,'Les conseils de l\'experte','Anticiper, archiver, se faire accompagner.');
  const it=[['LuClock','Anticiper','Tester la déclaration en ligne dès maintenant sur le système d\'information ; faire un dépôt à blanc sur une référence simple.'],['LuArchive','Archiver','Conserver registres, évaluations de risque et preuves de géolocalisation pendant 5 ans.'],['LuSearch','Auditer vos nomenclatures','Extraire vos codes SH importés et exportés sur 12 mois et les confronter à l\'annexe I.'],['LuKey','Préparer EU Login','Créer vos comptes et enregistrer l\'entreprise pour accéder au système d\'information.'],['LuMapPin','Sécuriser la géolocalisation','Obtenir les données de vos fournisseurs étrangers — indispensables à la DDR.'],['LuUsers','Se faire accompagner','Ministères (fond, analyse de risque) · Pôle d\'action économique des douanes (formalités) · CCI.']];
  for(const [i,[ic,h,t]] of it.entries()){ const x=0.6+(i%3)*4.1, y=1.85+Math.floor(i/3)*2.3; card(s,x,y,3.9,2.05); await iconCircle(s,x+0.3,y+0.3,0.7,ic); txt(s,h,x+1.2,y+0.3,2.6,0.75,{fontSize:14,bold:true,color:C.NAVY,valign:'middle'}); txt(s,t,x+0.3,y+1.1,3.3,0.9,{fontSize:11.5,color:C.GREY}); }
}
// ============ 32. SYNTHÈSE ============
{ const s=slide(); chrome(s); title(s,'La conformité en 3 niveaux','La synthèse : du produit à la déclaration parfaite. Chaque niveau s\'appuie sur le précédent.');
  const lv=[['Le produit conforme','Risque nul ou négligeable : zéro déforestation après le 31/12/2020 et légalité du pays de production.',C.GREEN,9.0],['Le système documenté','Preuves géolocalisées, DDR enregistrée dans le système d\'information, évaluations et mesures archivées 5 ans.',C.BLUE,7.0],['La déclaration parfaite','Liaison sans faille entre le n° de DDR et le code douanier : C716, C717 ou code Y.',C.NAVY,5.0]];
  lv.forEach(([h,t,col,w],i)=>{ const y=1.9+i*1.45, x=(W-w)/2; s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x,y,w,h:1.25,fill:{color:col},line:{color:col},rectRadius:0.08}); txt(s,h,x+0.3,y+0.15,w-0.6,0.4,{fontSize:16,bold:true,color:C.WHITE,align:'center'}); txt(s,t,x+0.4,y+0.55,w-0.8,0.65,{fontSize:12,color:'F0F4FA',align:'center'}); });
  txt(s,'Tout ce qui précède ne vaut que par la justesse de cette dernière ligne — la base, c\'est la donnée amont.',0.6,6.3,12.1,0.4,{fontSize:13,italic:true,color:C.PINK,align:'center'});
}
// ============ 33. QUIZ ============
{ const s=slide(); chrome(s); title(s,'Vérifions ensemble : 3 questions','Sondage dans le chat, puis correction.');
  const q=[['Un produit de l\'annexe I fabriqué le 12 mars 2023 : quel code portez-vous en douane ?','A. C716 + n° de DDR    B. Y132    C. Y129','Y132 — fabriqué avant le 29/06/2023'],['Le RDUE s\'applique-t-il aux exportations depuis l\'Union européenne ?','A. Oui, marché ET exportation    B. Non, importation seulement    C. Uniquement pour le bois','Oui — mise sur le marché ET exportation'],['« Zéro déforestation » : aucune déforestation de la parcelle après quelle date ?','A. 31 décembre 2020    B. 29 juin 2023    C. jamais déforestée','Le 31 décembre 2020']];
  for(const [i,[qq,opt,a]] of q.entries()){ const y=1.85+i*1.55; card(s,0.6,y,12.1,1.35); pill(s,0.85,y+0.42,0.55,0.5,`Q${i+1}`,C.NAVY,{size:13}); txt(s,qq,1.6,y+0.15,7.8,0.5,{fontSize:14,bold:true,color:C.NAVY}); txt(s,opt,1.6,y+0.7,7.8,0.5,{fontSize:12,color:C.GREY}); pill(s,9.7,y+0.42,2.75,0.5,a,C.GREEN,{size:11}); }
  txt(s,'Réponses à masquer avant le vote.',0.6,6.55,12,0.3,{fontSize:10.5,italic:true,color:C.MUTE});
}
// ============ 34. MERCI ============
{ const s=slide(true); chrome(s,{dark:true});
  s.addShape(pres.shapes.OVAL,{x:9.5,y:3.5,w:6.5,h:6.5,fill:{color:C.BLUE,transparency:75},line:{color:C.BLUE,transparency:75}});
  txt(s,'Merci — place à vos questions.',0.8,1.3,11,0.9,{fontSize:36,bold:true,color:C.WHITE});
  txt(s,'Vos 3 actions immédiates : auditer vos nomenclatures  ·  créer votre compte EU Login  ·  sécuriser vos géolocalisations.',0.8,2.25,11.5,0.5,{fontSize:14,italic:true,color:'C7D4E8'});
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:0.8,y:3.1,w:5.8,h:3.1,fill:{color:'123A78'},line:{color:'123A78'},rectRadius:0.08}); txt(s,'Vos interlocuteurs',1.1,3.25,5.2,0.4,{fontSize:15,bold:true,color:C.WHITE});
  bullets(s,['Fond, analyse de risque, DDR : ministères Transition écologique & Agriculture — deforestation@developpement-durable.gouv.fr','Formalités : Pôle d\'action économique (PAE) des douanes · Infos Douane Service 0 800 94 40 40','Commission européenne : FAQ RDUE (v5) et système d\'information','deforestationimportee.ecologie.gouv.fr'],1.1,3.7,5.3,2.4,{fontSize:11.5,color:'DCE6F5',gap:5});
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:6.9,y:3.1,w:5.8,h:3.1,fill:{color:'123A78'},line:{color:'123A78'},rectRadius:0.08}); txt(s,'Vos intervenantes',7.2,3.25,5.2,0.4,{fontSize:15,bold:true,color:C.WHITE});
  s.addShape(pres.shapes.ROUNDED_RECTANGLE,{x:7.2,y:3.8,w:0.9,h:0.9,fill:{color:C.WHITE},line:{color:C.WHITE},rectRadius:0.06}); s.addImage({data:logoCF2,x:7.25,y:3.85,w:0.8,h:0.8});
  txt(s,[R('Céline Fontana\n',{bold:true,fontSize:13,color:C.WHITE}),R('CF² Douane — Conseil & Formation\n06 86 93 95 72 · celine.fontana@gmail.com',{fontSize:11,color:'DCE6F5'})],8.3,3.75,4.3,1.0);
  txt(s,[R('Houda Zebar\n',{bold:true,fontSize:13,color:C.WHITE}),R('CCI Lyon Métropole Saint-Étienne Roanne — co-animatrice',{fontSize:11,color:'DCE6F5'})],7.2,4.95,5.3,0.9);
  txt(s,'Le support et l\'enregistrement vous seront adressés par courriel.',0.8,6.45,11,0.4,{fontSize:12,color:'9FB3D1'});
}
if(n!==TOTAL) throw new Error('slides '+n);
await pres.writeFile({fileName:'RDUE_Webinaire_2026-09-24_v3.pptx'}); console.log('written',n);
})().catch(e=>{console.error(e);process.exit(1)});
