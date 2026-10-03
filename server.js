const http=require('http'),fs=require('fs'),path=require('path');
const root=__dirname;
const types={'.html':'text/html; charset=utf-8','.js':'text/javascript; charset=utf-8','.json':'application/json; charset=utf-8','.webmanifest':'application/manifest+json; charset=utf-8','.png':'image/png','.css':'text/css; charset=utf-8'};
http.createServer((req,res)=>{
  let u=(req.url||'/').split('?')[0];
  if(u==='/') u='/index.html';
  let f=path.join(root,decodeURIComponent(u));
  if(!f.startsWith(root)){res.writeHead(403);return res.end('Forbidden');}
  fs.stat(f,(err,st)=>{
    if(err||!st.isFile()) f=path.join(root,'index.html');
    fs.readFile(f,(e,data)=>{
      if(e){res.writeHead(500);return res.end('Server error');}
      res.setHeader('Content-Type',types[path.extname(f)]||'application/octet-stream');
      res.setHeader('Cache-Control',path.basename(f)==='service-worker.js'?'no-cache':'public, max-age=300');
      res.end(data);
    });
  });
}).listen(process.env.PORT||3000,'0.0.0.0');