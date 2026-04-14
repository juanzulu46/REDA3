const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');
const url = require('url');

const PORT = process.env.PORT || 8080;
const GAS_URL = 'https://script.google.com/macros/s/AKfycbwvk6g9qhm_fqLbyjMspkvTf4MitW0gd0K-kvAU0KSmpYngxquq0XWV5mWS8EU8ATwk/exec';

const MIME = {'.html':'text/html','.js':'application/javascript','.css':'text/css','.json':'application/json','.xlsx':'application/octet-stream'};

// Sigue redirects de Google Apps Script (importante: 302 debe seguirse con GET)
function httpsRequest(reqUrl, options){
  return new Promise((resolve, reject)=>{
    const req = https.request(reqUrl, options, res=>{
      if(res.statusCode>=300 && res.statusCode<400 && res.headers.location){
        // En redirects, seguir con GET sin body (estándar HTTP para 302/303)
        httpsRequest(res.headers.location, {method:'GET'}).then(resolve).catch(reject);
      } else {
        let body='';
        res.on('data', c=>body+=c);
        res.on('end', ()=>resolve({status:res.statusCode, body}));
      }
    });
    req.on('error', reject);
    if(options.body) req.write(options.body);
    req.end();
  });
}

http.createServer(async (req, res)=>{
  const parsed = url.parse(req.url, true);

  // PROXY: /api?action=login  -->  Google Apps Script
  if(parsed.pathname==='/api'){
    try{
      if(req.method==='GET'){
        const gasUrl = GAS_URL+'?'+new url.URL(req.url,'http://localhost').searchParams.toString();
        const r = await httpsRequest(gasUrl, {method:'GET'});

        // Si GAS respondió con error HTTP, propagar al frontend
        if (r.status >= 400) {
          console.error('[proxy GET] GAS devolvió HTTP', r.status, r.body.slice(0, 500));
          res.writeHead(502,{'Content-Type':'application/json','Access-Control-Allow-Origin':'*'});
          res.end(JSON.stringify({ ok:false, error:'Backend GAS HTTP '+r.status, detalle:r.body.slice(0,500) }));
          return;
        }

        // Intentar parsear; si no es JSON, propagar 502 en vez de tragar
        let data;
        try { data = JSON.parse(r.body); }
        catch (e) {
          console.error('[proxy GET] Respuesta no-JSON de GAS:', r.body.slice(0, 500));
          res.writeHead(502,{'Content-Type':'application/json','Access-Control-Allow-Origin':'*'});
          res.end(JSON.stringify({ ok:false, error:'Respuesta inválida del backend (no-JSON)', detalle:r.body.slice(0,500) }));
          return;
        }

        // Enriquecer respuesta de mis_negocios: calcular flag 'efectuado' en cada pago
        if (data && data.ok && data.pagos) {
          const hoy = new Date();
          const warnings = [];
          data.pagos.forEach(function(p) {
            if (!p.fecha_pago || p.fecha_pago === '') {
              p.efectuado = true;
              return;
            }
            const fp = new Date(p.fecha_pago);
            if (isNaN(fp.getTime())) {
              const ap = Number(p['año_pago']) || Number(p['ano_pago']) || 0;
              const mp = Number(p.mes_pago) || 0;
              if (ap && mp) {
                p.efectuado = (ap * 12 + mp) <= (hoy.getFullYear() * 12 + hoy.getMonth() + 1);
              } else {
                p.efectuado = true;
                warnings.push('Pago '+p.id_pago+': fecha_pago inválida y sin año_pago/mes_pago');
              }
            } else {
              p.efectuado = fp <= hoy;
            }
          });
          if (warnings.length) {
            console.warn('[proxy GET] Warnings mis_negocios:', warnings);
            data.warnings = warnings;
          }
        }
        res.writeHead(200,{'Content-Type':'application/json','Access-Control-Allow-Origin':'*'});
        res.end(JSON.stringify(data));
      } else if(req.method==='POST'){
        let body='';
        req.on('data',c=>body+=c);
        req.on('end', async()=>{
          try {
            const r = await httpsRequest(GAS_URL, {method:'POST',headers:{'Content-Type':'text/plain'},body});
            if (r.status >= 400) {
              console.error('[proxy POST] GAS devolvió HTTP', r.status, r.body.slice(0, 500));
              res.writeHead(502,{'Content-Type':'application/json','Access-Control-Allow-Origin':'*'});
              res.end(JSON.stringify({ ok:false, error:'Backend GAS HTTP '+r.status, detalle:r.body.slice(0,500) }));
              return;
            }
            // Validar que sea JSON antes de devolver
            try { JSON.parse(r.body); }
            catch (e) {
              console.error('[proxy POST] Respuesta no-JSON de GAS:', r.body.slice(0, 500));
              res.writeHead(502,{'Content-Type':'application/json','Access-Control-Allow-Origin':'*'});
              res.end(JSON.stringify({ ok:false, error:'Respuesta inválida del backend (no-JSON)', detalle:r.body.slice(0,500) }));
              return;
            }
            res.writeHead(200,{'Content-Type':'application/json','Access-Control-Allow-Origin':'*'});
            res.end(r.body);
          } catch (err) {
            console.error('[proxy POST] Error de red:', err);
            res.writeHead(502,{'Content-Type':'application/json','Access-Control-Allow-Origin':'*'});
            res.end(JSON.stringify({ ok:false, error:'Error de red al contactar backend: '+err.message }));
          }
        });
      } else if(req.method==='OPTIONS'){
        res.writeHead(204,{'Access-Control-Allow-Origin':'*','Access-Control-Allow-Methods':'GET,POST','Access-Control-Allow-Headers':'Content-Type'});
        res.end();
      }
    }catch(e){
      res.writeHead(500,{'Content-Type':'application/json'});
      res.end(JSON.stringify({ok:false,error:e.message}));
    }
    return;
  }

  // STATIC FILES
  if(parsed.pathname==='/'){res.writeHead(302,{'Location':'/asesor.html'});res.end();return}
  let filePath = path.join(__dirname, parsed.pathname);
  const ext = path.extname(filePath);
  fs.readFile(filePath, (err, data)=>{
    if(err){res.writeHead(404);res.end('Not found');return}
    res.writeHead(200,{'Content-Type':MIME[ext]||'text/plain'});
    res.end(data);
  });
}).listen(PORT, ()=>console.log('REDA3 corriendo en http://localhost:'+PORT));
