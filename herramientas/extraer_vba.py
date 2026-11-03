"""Extrae modulos VBA de .xlsm (solo libreria estandar: CFB + MS-OVBA)."""
import sys, os, zipfile, struct, re

def decompress(data):
    if not data or data[0] != 1:
        raise ValueError("firma invalida")
    out = bytearray(); pos = 1
    while pos < len(data):
        hdr = struct.unpack_from('<H', data, pos)[0]; pos += 2
        size = (hdr & 0x0FFF) + 3
        flag = (hdr >> 15) & 1
        end = min(pos - 2 + size, len(data))
        start = len(out)
        if not flag:
            out += data[pos:pos + 4096]; pos += 4096; continue
        while pos < end:
            fb = data[pos]; pos += 1
            for bit in range(8):
                if pos >= end: break
                if not (fb >> bit) & 1:
                    out.append(data[pos]); pos += 1
                else:
                    tok = struct.unpack_from('<H', data, pos)[0]; pos += 2
                    d = len(out) - start
                    bc = max((d - 1).bit_length(), 4)
                    lmask = 0xFFFF >> bc
                    off = (tok >> (16 - bc)) + 1
                    ln = (tok & lmask) + 3
                    for _ in range(ln):
                        out.append(out[-off])
    return bytes(out)

class CFB:
    def __init__(self, b):
        self.b = b
        self.ss = 1 << struct.unpack_from('<H', b, 0x1E)[0]
        self.mss = 1 << struct.unpack_from('<H', b, 0x20)[0]
        nfat = struct.unpack_from('<I', b, 0x2C)[0]
        dir0 = struct.unpack_from('<I', b, 0x30)[0]
        self.cutoff = struct.unpack_from('<I', b, 0x38)[0]
        mf0 = struct.unpack_from('<I', b, 0x3C)[0]
        difat0 = struct.unpack_from('<I', b, 0x44)[0]
        difat = list(struct.unpack_from('<109I', b, 0x4C))
        s = difat0
        while s < 0xFFFFFFFA:
            sec = self.sector(s)
            vals = struct.unpack('<%dI' % (self.ss // 4), sec)
            difat += vals[:-1]; s = vals[-1]
        difat = [x for x in difat if x < 0xFFFFFFFA][:nfat]
        self.fat = []
        for s in difat:
            self.fat += struct.unpack('<%dI' % (self.ss // 4), self.sector(s))
        self.minifat = []
        if mf0 < 0xFFFFFFFA:
            d = self.chain(mf0)
            self.minifat = list(struct.unpack('<%dI' % (len(d) // 4), d))
        dd = self.chain(dir0)
        self.entries = []
        for i in range(len(dd) // 128):
            e = dd[i * 128:(i + 1) * 128]
            nl = struct.unpack_from('<H', e, 64)[0]
            name = e[:max(nl - 2, 0)].decode('utf-16-le', 'ignore')
            typ = e[66]
            l, r, c = struct.unpack_from('<III', e, 68)
            st, sz = struct.unpack_from('<II', e, 116)
            self.entries.append(dict(name=name, type=typ, l=l, r=r, c=c, st=st, sz=sz))
        root = self.entries[0]
        self.ministream = self.chain(root['st'])[:root['sz']] if root['st'] < 0xFFFFFFFA else b''

    def sector(self, s):
        o = (s + 1) * self.ss
        return self.b[o:o + self.ss]

    def chain(self, s):
        out = bytearray(); seen = 0
        while s < 0xFFFFFFFA and seen < 1000000:
            out += self.sector(s); s = self.fat[s]; seen += 1
        return bytes(out)

    def read(self, e):
        if e['sz'] < self.cutoff:
            out = bytearray(); s = e['st']
            while s < 0xFFFFFFFA:
                o = s * self.mss; out += self.ministream[o:o + self.mss]; s = self.minifat[s]
            return bytes(out[:e['sz']])
        return self.chain(e['st'])[:e['sz']]

    def children(self, idx):
        res = []
        def walk(i):
            if i == 0xFFFFFFFF or i >= len(self.entries): return
            e = self.entries[i]; walk(e['l']); res.append(i); walk(e['r'])
        walk(self.entries[idx]['c'])
        return res

    def find(self, path):
        idx = 0
        for part in path.split('/'):
            for c in self.children(idx):
                if self.entries[c]['name'].lower() == part.lower():
                    idx = c; break
            else:
                return None
        return self.entries[idx]

def modules(vbabin):
    cf = CFB(vbabin)
    # localizar storage que contiene "dir"
    prefix = None
    for p in ('VBA', '_VBA_PROJECT_CUR/VBA'):
        if cf.find(p + '/dir'): prefix = p; break
    d = decompress(cf.read(cf.find(prefix + '/dir')))
    pos = 0; mods = []; cur = None; codepage = 1252
    while pos + 6 <= len(d):
        rid, size = struct.unpack_from('<HI', d, pos); pos += 6
        if rid == 0x0009: size = 6
        data = d[pos:pos + size]; pos += size
        if rid == 0x0003: codepage = struct.unpack_from('<H', data)[0]
        elif rid == 0x0019: cur = dict(name=data.decode('cp%d' % codepage, 'replace'), type='bas'); mods.append(cur)
        elif rid == 0x0047 and cur: cur['name'] = data.decode('utf-16-le', 'replace')
        elif rid == 0x001A and cur: cur['stream'] = data.decode('cp%d' % codepage, 'replace')
        elif rid == 0x0032 and cur: cur['stream'] = data.decode('utf-16-le', 'replace')
        elif rid == 0x0031 and cur: cur['off'] = struct.unpack_from('<I', data)[0]
        elif rid == 0x0022 and cur: cur['type'] = 'cls'
    out = []
    for m in mods:
        e = cf.find(prefix + '/' + m.get('stream', m['name']))
        if not e: continue
        raw = cf.read(e)
        code = decompress(raw[m.get('off', 0):]).decode('cp%d' % codepage, 'replace')
        out.append((m['name'], m['type'], code))
    return out

if __name__ == '__main__':
    src, dst = sys.argv[1], sys.argv[2]
    os.makedirs(dst, exist_ok=True)
    with zipfile.ZipFile(src) as z:
        names = [n for n in z.namelist() if n.lower().endswith('vbaproject.bin')]
        if not names: print('sin VBA'); sys.exit(0)
        vb = z.read(names[0])
    for name, typ, code in modules(vb):
        code = code.replace('\r\n', '\n')
        body = [l for l in code.split('\n') if l.strip() and not l.startswith('Attribute VB_')]
        if not body: continue
        ext = '.cls' if typ == 'cls' else '.bas'
        if re.search(r'^Attribute VB_Name', code, re.M) and 'VB_Base = "0{' in code: ext = '.frm'
        with open(os.path.join(dst, name + ext), 'w', encoding='utf-8', newline='\n') as f:
            f.write(code)
        print('%-28s %5d lineas' % (name + ext, len(code.split('\n'))))
