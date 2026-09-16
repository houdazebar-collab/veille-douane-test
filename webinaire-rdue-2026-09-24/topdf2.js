const { chromium } = require('playwright-core'); const path=require('path');
(async()=>{ const b=await chromium.launch({executablePath:'/opt/pw-browsers/chromium/chrome-linux/chrome'}).catch(async e=>{ const fs=require('fs'); const dirs=fs.readdirSync('/opt/pw-browsers').filter(d=>d.startsWith('chromium-')); const p='/opt/pw-browsers/'+dirs[0]+'/chrome-linux/chrome'; return chromium.launch({executablePath:p}); });
 const pg=await b.newPage(); await pg.goto('file://'+path.resolve('guide2_print.html'),{waitUntil:'load'});
 await pg.pdf({path:'GUIDE_RDUE_approfondi.pdf',format:'A4',printBackground:true,displayHeaderFooter:true,
   headerTemplate:'<div style="font-size:7.5pt;color:#8A98A8;width:100%;padding:0 18mm;font-family:Carlito,Calibri,Arial">Guide approfondi RDUE · CCI Lyon Métropole Saint-Étienne Roanne · CF² Douane</div>',
   footerTemplate:'<div style="font-size:7.5pt;color:#8A98A8;width:100%;padding:0 18mm;display:flex;justify-content:space-between;font-family:Carlito,Calibri,Arial"><span>État du droit au 16/09/2026 — document d\'information, ne vaut pas conseil juridique</span><span><span class="pageNumber"></span> / <span class="totalPages"></span></span></div>',
   margin:{top:'22mm',bottom:'20mm',left:'18mm',right:'18mm'}});
 await b.close(); console.log('pdf ok'); })().catch(e=>{console.error(e);process.exit(1)});
