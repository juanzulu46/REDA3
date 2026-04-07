const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');
const url = require('url');

const PORT = process.env.PORT || 8080;
const GAS_URL = 'https://script.google.com/macros/s/AKfycbyeBVMDSD9-giQaUTCIhvnLu6RlQLqDC3t8dtGMfUPi-cULruHc-_QrISZqPUXXfog_/exec';

const MIME = {'.html':'text/html','.js':'application/javascript','.css':'text/css','.json':'application/json','.xlsx':'application/octet-stream'};

// Sigue redirects de Google Apps Script
function httpsRequest(reqUrl, options){
  return new Promise((resolve, reject)=>{
    const req = https.request(reqUrl, options, res=>{
      if(res.statusCode>=300 && res.statusCode<400 && res.headers.location){
        httpsRequest(res.headers.location, options).then(resolve).catch(reject);
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
        res.writeHead(200,{'Content-Type':'application/json','Access-Control-Allow-Origin':'*'});
        res.end(r.body);
      } else if(req.method==='POST'){
        let body='';
        req.on('data',c=>body+=c);
        req.on('end', async()=>{
          const r = await httpsRequest(GAS_URL, {method:'POST',headers:{'Content-Type':'text/plain'},body});
          res.writeHead(200,{'Content-Type':'application/json','Access-Control-Allow-Origin':'*'});
          res.end(r.body);
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
