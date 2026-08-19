// 로컬 미리보기용 정적 서버 (개발 확인용)
const http = require('http');
const fs = require('fs');
const path = require('path');
const root = __dirname;
const types = {'.html':'text/html; charset=utf-8','.js':'text/javascript; charset=utf-8','.css':'text/css','.png':'image/png','.svg':'image/svg+xml','.ico':'image/x-icon'};
http.createServer((req,res)=>{
  let p = decodeURIComponent(req.url.split('?')[0]);
  if(p==='/')p='/index.html';
  const fp = path.join(root,p);
  if(!fp.startsWith(root)||!fs.existsSync(fp)||fs.statSync(fp).isDirectory()){res.writeHead(404);return res.end('404');}
  res.writeHead(200,{'Content-Type':types[path.extname(fp)]||'application/octet-stream'});
  fs.createReadStream(fp).pipe(res);
}).listen(4599,()=>console.log('preview on http://localhost:4599'));
