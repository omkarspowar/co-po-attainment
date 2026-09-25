from __future__ import annotations

import json
import mimetypes
import os
import shutil
import tempfile
import threading
import time
import uuid
import webbrowser
from email.parser import BytesParser
from email.policy import default
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote

from engine import run_job

ROOT=Path(__file__).resolve().parent
STATIC=ROOT/'static'
JOBS=Path(tempfile.gettempdir())/'co_po_attainment_downloads'
MAX_UPLOAD=30*1024*1024
DOWNLOAD_TTL=15*60
DOWNLOADS={}
DOWNLOAD_LOCK=threading.Lock()


def cleanup_downloads():
    now=time.time()
    with DOWNLOAD_LOCK:
        expired=[token for token,item in DOWNLOADS.items() if item['expires']<now or not item['path'].exists()]
        for token in expired:
            item=DOWNLOADS.pop(token,None)
            if item and item['path'].exists(): item['path'].unlink(missing_ok=True)


class Handler(SimpleHTTPRequestHandler):
    def _json(self,status,payload):
        body=json.dumps(payload).encode(); self.send_response(status); self.send_header('Content-Type','application/json'); self.send_header('Content-Length',str(len(body))); self.end_headers(); self.wfile.write(body)
    def do_GET(self):
        cleanup_downloads()
        if self.path in {'/','/index.html'}:
            self.path='/index.html'; return self._serve_static()
        if self.path.startswith('/static/'):
            return self._serve_static(self.path.removeprefix('/static/'))
        if self.path.startswith('/download/'):
            token=Path(unquote(self.path.removeprefix('/download/'))).name
            with DOWNLOAD_LOCK: item=DOWNLOADS.pop(token,None)
            if not item or not item['path'].exists(): return self.send_error(404,'This one-time download has expired or was already used.')
            try:
                data=item['path'].read_bytes(); self.send_response(200); self.send_header('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); self.send_header('Content-Disposition',f'attachment; filename="{item["name"]}"'); self.send_header('Content-Length',str(len(data))); self.end_headers(); self.wfile.write(data)
            finally: item['path'].unlink(missing_ok=True)
            return
        if self.path=='/api/health': return self._json(200,{'status':'ok'})
        self.send_error(404)
    def _serve_static(self,name=None):
        name=name or self.path.lstrip('/'); path=(STATIC/name).resolve()
        if STATIC.resolve() not in path.parents and path!=STATIC.resolve(): return self.send_error(403)
        if not path.exists(): return self.send_error(404)
        data=path.read_bytes(); self.send_response(200); self.send_header('Content-Type',mimetypes.guess_type(path.name)[0] or 'application/octet-stream'); self.send_header('Content-Length',str(len(data))); self.end_headers(); self.wfile.write(data)
    def do_POST(self):
        if self.path!='/api/generate': return self.send_error(404)
        length=int(self.headers.get('Content-Length','0'))
        if length>MAX_UPLOAD: return self._json(413,{'error':'Uploads exceed the 30 MB limit.'})
        temp=Path(tempfile.mkdtemp(prefix='co_po_'))
        try:
            content_type=self.headers.get('Content-Type','')
            if 'multipart/form-data' not in content_type: return self._json(400,{'error':'Expected a multipart form upload.'})
            body=self.rfile.read(length)
            message=BytesParser(policy=default).parsebytes(f'Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n'.encode()+body)
            form_fields={}; form_files={}
            for part in message.iter_parts():
                name=part.get_param('name',header='content-disposition'); filename=part.get_filename(); payload=part.get_payload(decode=True) or b''
                if not name: continue
                if filename: form_files[name]=(filename,payload)
                else: form_fields[name]=payload.decode(part.get_content_charset() or 'utf-8',errors='replace')
            files={}
            for key in ['template','marks','grades','survey','course_plan']:
                item=form_files.get(key)
                if item:
                    suffix=Path(item[0]).suffix.lower(); path=temp/(key+suffix); path.write_bytes(item[1]); files[key]=path
                else: files[key]=None
            if not files['template'] or not files['marks']: return self._json(400,{'error':'Template and CO-wise marks files are required.'})
            fields={key:form_fields.get(key,'') for key in ['course_code','course_name','faculty','school','program','semester','year','odd_even','section','course_type','cie_weight','see_weight','direct_weight','indirect_weight','mapping','co_targets','po_targets','co1','co2','co3','co4']}
            output,summary=run_job(files,fields,JOBS)
            public_name=output.name.split('_',1)[1] if '_' in output.name else output.name; token=uuid.uuid4().hex
            with DOWNLOAD_LOCK: DOWNLOADS[token]={'path':output,'name':public_name,'expires':time.time()+DOWNLOAD_TTL}
            summary['file']=public_name; summary['download_url']='/download/'+token; summary['download_expires_minutes']=15; self._json(200,summary)
        except Exception as exc:
            self._json(400,{'error':str(exc)})
        finally: shutil.rmtree(temp,ignore_errors=True)


def main():
    host=os.environ.get('COPO_HOST','127.0.0.1'); port=int(os.environ.get('PORT',os.environ.get('COPO_PORT','8765'))); JOBS.mkdir(parents=True,exist_ok=True)
    url=f'http://{host}:{port}'; print(f'CO-PO Automation is running at {url}'); print('Press Ctrl+C to stop.')
    if os.environ.get('COPO_NO_BROWSER')!='1': webbrowser.open(url)
    ThreadingHTTPServer((host,port),Handler).serve_forever()


if __name__=='__main__': main()
