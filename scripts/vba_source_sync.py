"""Read module streams from an OLE compound VBA project, no third-party packages.
This is a source-extraction aid; it never executes macros or modifies workbooks.
"""
import struct,zipfile,math,re,sys,hashlib
FREE=0xFFFFFFFF; END=0xFFFFFFFE; FAT=0xFFFFFFFD; DIF=0xFFFFFFFC
u16=lambda b,o: struct.unpack_from('<H',b,o)[0]
u32=lambda b,o: struct.unpack_from('<I',b,o)[0]
u64=lambda b,o: struct.unpack_from('<Q',b,o)[0]
class CFB:
 def __init__(self,b):
  assert b[:8]==bytes.fromhex('D0CF11E0A1B11AE1')
  self.b=b; self.ss=1<<u16(b,30);self.ms=1<<u16(b,32)
  self.mini_cut=u32(b,56)
  self.dif=[u32(b,76+i*4) for i in range(109) if u32(b,76+i*4)<DIF]
  difsec=u32(b,68)
  for _ in range(u32(b,72)):
   if difsec>=DIF:break
   raw=self.sector(difsec); entries=[u32(raw,i) for i in range(0,self.ss,4)]
   self.dif += [x for x in entries[:-1] if x<DIF]; difsec=entries[-1]
  self.fat=[]
  for i in self.dif[:u32(b,44)]:
   raw=self.sector(i); self.fat += [u32(raw,j) for j in range(0,self.ss,4)]
  self.directory=self.chain(u32(b,48))
  self.entries=[]
  for i in range(0,len(self.directory),128):
   e=self.directory[i:i+128];
   if len(e)<128: break
   nlen=u16(e,64); typ=e[66]
   if nlen>=2 and nlen<=64:
    name=e[:nlen-2].decode('utf-16le',errors='replace')
   else:name=''
   self.entries.append(dict(index=i//128,name=name,type=typ,start=u32(e,116),size=u64(e,120),child=u32(e,76),left=u32(e,68),right=u32(e,72)))
  self.root=self.entries[0]
  self.ministream=self.chain(self.root['start'],self.root['size'])
  msid=u32(b,60); self.minifat=[]
  if msid<END:
   dat=self.chain(msid,u32(b,64)*self.ss)
   self.minifat=[u32(dat,j) for j in range(0,len(dat),4)]
 def sector(self,sid): return self.b[512+sid*self.ss:512+(sid+1)*self.ss]
 def chain(self,start,size=None):
  out=bytearray(); sid=start; seen=set()
  while sid<END and sid<len(self.fat) and sid not in seen:
   seen.add(sid);out.extend(self.sector(sid));sid=self.fat[sid]
   if size and len(out)>=size:break
  return bytes(out[:size] if size is not None else out)
 def minichain(self,start,size):
  out=bytearray(); sid=start; seen=set()
  while sid<END and sid<len(self.minifat) and sid not in seen:
   seen.add(sid); off=sid*self.ms;out.extend(self.ministream[off:off+self.ms]);sid=self.minifat[sid]
   if len(out)>=size:break
  return bytes(out[:size])
 def stream(self,e):
  if e['size']<self.mini_cut and e['type']==2: return self.minichain(e['start'],e['size'])
  return self.chain(e['start'],e['size'])

def decompress(buf):
 if not buf or buf[0]!=1: raise ValueError('not compressed')
 pos=1;out=bytearray()
 while pos+2<=len(buf):
  head=u16(buf,pos);pos+=2
  if (head&0x7000)!=0x3000: raise ValueError('bad signature at '+str(pos))
  chunk_size=(head & 0x0FFF)+3
  end=pos+chunk_size-2
  if end>len(buf):raise ValueError('truncated chunk')
  chunk=buf[pos:end];pos=end
  if not (head & 0x8000):
   out.extend(chunk)
   continue
  cpos=0;dec=bytearray()
  while cpos<len(chunk):
   flag=chunk[cpos];cpos+=1
   for bit in range(8):
    if cpos>=len(chunk):break
    if flag & (1<<bit):
     if cpos+2>len(chunk):raise ValueError('bad copy token')
     token=u16(chunk,cpos);cpos+=2
     bits=max(4,(max(0,len(dec)-1)).bit_length())
     bits=min(bits,12)
     length_bits=16-bits
     length=(token & ((1<<length_bits)-1))+3
     offset=(token>>length_bits)+1
     if offset>len(dec):raise ValueError('bad offset')
     for _ in range(length):dec.append(dec[-offset])
    else:
     dec.append(chunk[cpos]);cpos+=1
    if len(dec)>4096:raise ValueError('chunk overflow')
  out.extend(dec)
  if len(dec)>4096:raise ValueError('bad chunk')
 return bytes(out)



def workbook_modules(xlsm_path):
    """Return {module_name: unmodified decoded source bytes} from the XLSM VBA project."""
    with zipfile.ZipFile(xlsm_path) as z:
        b = z.read('xl/vbaProject.bin')
    c = CFB(b)
    result = {}
    for entry in c.entries:
        name = entry['name']
        if entry['type'] != 2 or not re.fullmatch(r'mod[A-Za-z0-9_]+', name):
            continue
        data = c.stream(entry)
        valid = []
        for off in range(len(data)-3):
            if data[off] != 1 or (data[off+2]&0x70)!=0x30:
                continue
            try:
                decoded = decompress(data[off:])
                if (b'Attribute VB_Name = "'+name.encode('ascii')+b'"') in decoded and b'Option Explicit' in decoded:
                    valid.append(decoded)
            except (ValueError, IndexError, struct.error):
                continue
        if not valid:
            raise ValueError('Cannot decode validated VBA source: '+name)
        result[name] = max(valid,key=len)
    return result

if __name__ == '__main__':
    from pathlib import Path
    root=Path(__file__).resolve().parents[1]
    workbook=root/'excel/LongiSafeLab.xlsm'
    modules=workbook_modules(workbook)
    fail=[]
    for name,raw in sorted(modules.items()):
        source=root/'excel/vba'/f'{name}.bas'
        if not source.is_file() or source.read_bytes()!=raw:
            fail.append(name)
            print(name,'MISMATCH')
        else:
            print(name,'MATCH',hashlib.sha256(raw).hexdigest())
    extra=set(p.stem for p in (root/'excel/vba').glob('*.bas'))-set(modules)
    if extra:fail.extend(sorted(extra))
    print('Workbook/VBA source byte match:', 'PASS' if not fail else 'FAIL')
    sys.exit(bool(fail))
