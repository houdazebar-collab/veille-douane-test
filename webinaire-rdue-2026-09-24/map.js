const topo=require('world-atlas/countries-110m.json'); const tc=require('topojson-client'); const d3=require('d3-geo'); const sharp=require('sharp'); const fs=require('fs');
const HIGH=new Set(['Belarus','Myanmar','North Korea','Russia']);
const STD=new Set(JSON.parse(fs.readFileSync('standard.json','utf8')));
const g=tc.feature(topo,topo.objects.countries);
const W=2600,H=1300;
const proj=d3.geoNaturalEarth1().fitExtent([[10,10],[W-10,H-10]],{type:'Sphere'});
const path=d3.geoPath(proj);
const col=n=>HIGH.has(n)?'#D0313B':STD.has(n)?'#D9A13B':'#3FA66B';
let svg=`<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}"><rect width="${W}" height="${H}" fill="#FFFFFF"/>`;
svg+=`<path d="${path({type:'Sphere'})}" fill="#EAF1F8" stroke="none"/>`;
const missing=[];
for(const f of g.features){ const n=f.properties.name; if(n==='Antarctica'||n==='Fr. S. Antarctic Lands') continue; svg+=`<path d="${path(f)}" fill="${col(n)}" stroke="#FFFFFF" stroke-width="1.2"/>`; }
for(const n of [...STD,...HIGH]) if(!g.features.some(f=>f.properties.name===n)) missing.push(n);
svg+='</svg>';
fs.writeFileSync('map.svg',svg);
sharp(Buffer.from(svg)).png().toFile('map.png').then(()=>console.log('map ok; unmatched names:',missing.join(', ')));
