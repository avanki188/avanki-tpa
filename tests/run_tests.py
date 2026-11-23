#!/usr/bin/env python3
"""Bộ kiểm tra tự động cho Avanki – TPA.

Chạy:  python3 tests/run_tests.py
Cần:   pip install playwright openpyxl  (+ Chrome/Chromium; đặt CHROME=/đường/dẫn nếu không dùng bản của Playwright)

Mở index.html trong trình duyệt không giao diện, giả lập kho dữ liệu trong bộ nhớ (KHÔNG đụng OneDrive),
dùng file Excel tạo ngay lúc chạy với tên và số liệu BỊA, rồi kiểm tra từng chức năng.
"""
import base64, io, os, re, sys, zipfile, datetime
import openpyxl
from playwright.sync_api import sync_playwright

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
APP = 'file://' + os.path.join(ROOT, 'index.html')

# ---------------------------------------------------------------- tiện ích Excel
def xlsx(rows, name='Sheet1', fmt=None):
    """Tạo .xlsx (bytes). fmt: {(row, col): number_format} – row/col tính từ 1."""
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = name
    for r in rows: ws.append(r)
    for (r, c), f in (fmt or {}).items(): ws.cell(r, c).number_format = f
    b = io.BytesIO(); wb.save(b); return b.getvalue()

def with_ns_prefix(data):
    """Ghi lại file theo kiểu hệ thống xuất: mọi thẻ XML có tiền tố x: (<x:sheet>, <x:row>…)."""
    src = zipfile.ZipFile(io.BytesIO(data)); out = io.BytesIO(); dst = zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED)
    NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    for it in src.infolist():
        b = src.read(it.filename)
        if it.filename.startswith('xl/') and it.filename.endswith('.xml') and NS.encode() in b:
            t = b.decode('utf-8')
            t = re.sub(r'<(/?)(?!\?)([A-Za-z_][\w.-]*)(?=[\s/>])', lambda m: f'<{m.group(1)}x:{m.group(2)}' if ':' not in m.group(2) else m.group(0), t)
            t = t.replace(f'xmlns="{NS}"', f'xmlns:x="{NS}"')
            b = t.encode('utf-8')
        dst.writestr(it, b)
    dst.close(); return out.getvalue()

def read_xlsx(b64):
    wb = openpyxl.load_workbook(io.BytesIO(base64.b64decode(b64)))
    return {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb}

def b64(b): return base64.b64encode(b).decode()

# ---------------------------------------------------------------- dữ liệu bịa
POL = 'P-99/TST/P01/0000/000001'
CONTRACT = {'id': 'd1', 'unitCode': '7608000005', 'policy': POL, 'customerCode': 'KH-TEST', 'contractName': 'CONG TY THU NGHIEM ABC', 'effectivePeriod': '01/01/2026 - 31/12/2026', 'tpa': 'FHVI', 'group': ''}
FHVI_SRC = [['BÁO CÁO DPBT (DỮ LIỆU THỬ)'], [],
            ['STT', 'CLAIM NO', 'POLICY NUMBER', 'MEMBER NAME', 'NATION ID', 'DOB', 'PAYABLE AMOUNT', 'PROVIDER NAME', 'BENEFIT SERVICE', 'DIAGNOSIS REMARK', 'TREATMENT DATE'],
            [1, 'TST-0001', POL, 'NGUYEN VAN THU', '001090000001', '01/01/1990', 1500000, 'BENH VIEN THU A', 'Ngoại trú', 'Cảm', '05/11/2026'],
            [2, 'TST-0002', POL, 'TRAN THI NGHIEM', '001090000002', '02/02/1991', 250000, 'BENH VIEN THU B', 'Ngoại trú', 'Ho', '06/11/2026'],
            [3, 'TST-0003', POL, 'LE VAN MAU', '001090000003', '03/03/1992', 333000, 'BENH VIEN THU C', 'Nội trú', 'Sốt', '07/11/2026']]
GD = '01.11 - 04.11.26'
NOTE = 'GD ' + GD

# ---------------------------------------------------------------- khung chạy
RESULTS = []
def check(name, cond, detail=''):
    RESULTS.append((name, bool(cond), detail))
    print(('  ✓ ' if cond else '  ✗ ') + name + ('' if cond or not detail else f'  →  {detail}'))

def main():
    with sync_playwright() as p:
        kw = {'args': ['--no-sandbox']}
        if os.environ.get('CHROME'): kw['executable_path'] = os.environ['CHROME']
        elif os.path.exists('/usr/bin/google-chrome'): kw['executable_path'] = '/usr/bin/google-chrome'
        br = p.chromium.launch(**kw); pg = br.new_page(viewport={'width': 1400, 'height': 1000})
        errs = []; pg.on('pageerror', lambda e: errs.append(str(e)))
        pg.goto(APP); pg.wait_for_timeout(1200)

        def reset():
            pg.evaluate("""(c)=>{for(const k in STORE_DEFS)DB[k]=STORE_DEFS[k].init();DB.data.rows.push(Object.assign({},c));rebuildClaimIdx();rebuildPolicyIdx();
              dirH={};started=true;window._saved=[];touch=n=>{window._saved.push(n);if(n==='claims')CLAIMS_VER++;if(n==='tracker')TRK_VER++};
              window._al=[];window.alert=m=>{window._al.push(String(m))};window._cf=[];window._cfAns=true;confirmBox=async o=>{window._cf.push(o.title+' | '+(o.msg||''));return window._cfAns};
              window.mkf=(n,b)=>{const s=atob(b),u=new Uint8Array(s.length);for(let i=0;i<s.length;i++)u[i]=s.charCodeAt(i);return new File([u],n)};
              window.u64=u=>{let s='';for(let i=0;i<u.length;i++)s+=String.fromCharCode(u[i]);return btoa(s)};
              S.tab='xl';S.conv.result=null;S.bk.result=null;S.dn.result=null;S.tool.fx=null;S.tool.dc=null;renderApp();}""", CONTRACT)
        ev = pg.evaluate

        def conv(src=FHVI_SRC, tpc=False):
            ev("""async(a)=>{S.tab='xl';Object.assign(S.conv,{profile:'FHVI',date:'20/11/2026',gd:'GD 01.11 - 04.11.26',tpc:a[1],noteMan:null,file:mkf('nguon.xlsx',a[0]),result:null,review:null});await ACT.convGo();}""", [b64(xlsx(src)), tpc])
            return ev("()=>{const r=S.conv.result;return r&&{msgs:r.msgs.map(m=>m.t),files:Object.fromEntries(Object.entries(r.files||{}).map(([k,v])=>[k,u64(v.u8)]))}}")

        # ============ 1. Đọc Excel
        print('\n1. Đọc Excel')
        reset()
        plain = xlsx([['Số HSBT TPA', 'Số tiền'], ['TST-1', 100], ['TST-2', 200]])
        for label, data in [('file thường', plain), ('file hệ thống có tiền tố x:', with_ns_prefix(plain))]:
            r = ev("async(b)=>{const sh=await readExcelFile(mkf('t.xlsx',b));return sh.map(s=>[s.name,s.rows.length,(s.rows[2]||[]).map(String)])}", b64(data))
            check(f'Đọc được {label}', r and r[0][1] == 3 and r[0][2] == ['TST-2', '200'], r)

        # ============ 2. Chuyển đổi DPBT
        print('\n2. Chuyển đổi DPBT')
        reset()
        ev("()=>{DB.data.rows=[];rebuildPolicyIdx();}")
        r = conv()
        check('Thiếu DATA hợp đồng → dừng, có file lỗi thiếu trường', r and 'missing' in r['files'] and 'full' not in r['files'], r and r['msgs'])
        check('Thiếu DATA → không lưu hồ sơ', ev('()=>DB.claims.rows.length') == 0)
        reset()
        r = conv()
        ok = r and 'full' in r['files']
        check('Tạo được file FULL', ok, r and r['msgs'])
        full_b64 = r['files']['full'] if ok else None
        if ok:
            rows = read_xlsx(full_b64); sh = list(rows.values())[0]; hd = sh[0]
            ix = {h: i for i, h in enumerate(hd)}
            body = [x for x in sh[1:] if x[ix['Số HSBT TPA']]]
            check('FULL có đủ 3 hồ sơ', len(body) == 3, len(body))
            check('FULL: Số tiền DPBT ntệ đúng', [x[ix['Số tiền DPBT ntệ']] for x in body] == [1500000, 250000, 333000])
            check('FULL: Đơn vị + Đối tượng BH lấy từ DATA', body[0][ix['Đơn vị']] == '7608000005' and body[0][ix['Đối tượng bảo hiểm']] == 'CONG TY THU NGHIEM ABC')
            check('FULL: Ghi chú đúng mẫu FHVI', body[0][ix['Ghi chú']] == NOTE, body[0][ix['Ghi chú']])
            check('FULL: Ngày lập dự phòng là ngày thật', isinstance(body[0][ix['Ngày lập dự phòng']], datetime.datetime))
        check('Lưu 3 hồ sơ vào kho', ev('()=>DB.claims.rows.length') == 3)
        r2 = conv()
        check('Upload lại cùng file → phát hiện trùng, không lưu thêm', ev('()=>DB.claims.rows.length') == 3 and ev('()=>!!S.conv.review || (S.conv.result&&S.conv.result.msgs.some(m=>/trùng/i.test(m.t)))'), r2 and r2['msgs'])

        # ============ 3. Lập BK
        print('\n3. Lập BK')
        bk_b64 = None
        if full_b64:
            ev("""async(b)=>{Object.assign(S.bk,{profile:'FHVI',date:'20/11/2026',gd:'01.11 - 04.11.26',infoMan:null,file:mkf('full.xlsx',b),result:null});await ACT.bkGo();}""", full_b64)
            r = ev("()=>{const x=S.bk.result;return x&&{msgs:x.msgs.map(m=>m.t),u8:x.u8?u64(x.u8):null}}")
            check('Tạo được bảng kê', r and r['u8'], r and r['msgs'])
            if r and r['u8']:
                bk_b64 = r['u8']; rows = list(read_xlsx(bk_b64).values())[0]
                flat = [c for row in rows for c in row if c is not None]
                check('BK có đủ 3 Số HSBT TPA', all(x in flat for x in ['TST-0001', 'TST-0002', 'TST-0003']))
                check('BK: tổng cộng 2.083.000', 2083000 in flat)
                check('Thông báo tổng tiền đúng', any('2.083.000' in m for m in r['msgs']), r['msgs'])

        # ============ 4. Giấy ĐNTT
        print('\n4. Giấy ĐNTT')
        if bk_b64:
            ev("()=>{_al.length=0;DB.settings.dntt.FHVI={name:'',account:'',bank:''};}")
            ev("""async(b)=>{Object.assign(S.dn,{profile:'FHVI',gd:'01.11 - 04.11.26',contentMan:null,file:mkf('bk.xlsx',b),result:null});await ACT.dnGo();}""", bk_b64)
            al = ev('()=>_al'); check('Cài đặt thiếu người thụ hưởng → báo lỗi, không tạo file', ev('()=>!S.dn.result') and any('còn thiếu' in a for a in al), al)
            ev("()=>{_al.length=0;DB.settings.dntt.FHVI={name:'CONG TY THU HUONG THU',account:'000.111.222',bank:'NGAN HANG THU'};}")
            ev("""async(b)=>{Object.assign(S.dn,{file:mkf('bk.xlsx',b),result:null});await ACT.dnGo();}""", bk_b64)
            r = ev("()=>{const x=S.dn.result;return x&&{total:x.total,name:x.name,u8:x.u8?u64(x.u8):null}}")
            check('Tạo được Giấy ĐNTT', r and r['u8'], ev('()=>_al'))
            if r and r['u8']:
                check('ĐNTT: tổng tiền 2.083.000', r['total'] == 2083000, r['total'])
                doc = zipfile.ZipFile(io.BytesIO(base64.b64decode(r['u8']))).read('word/document.xml').decode('utf-8')
                txt = re.sub(r'<[^>]+>', '', doc)
                check('ĐNTT: Word có người thụ hưởng + STK từ Cài đặt', 'CONG TY THU HUONG THU' in txt and '000.111.222' in txt)
                check('ĐNTT: Word có số tiền 2,083,000 và bằng chữ', '2,083,000' in txt and 'Hai triệu' in txt)
            check('Màn ③ không còn dòng Người thụ hưởng', ev("()=>{S.tab='xl';renderApp();const e=document.getElementById('xc-dntt');return !!e&&!/Người thụ hưởng:/.test(e.innerText)}"))

        # ============ 5. Fix lỗi
        print('\n5. Fix lỗi')
        if full_b64:
            H = ['Đơn vị', 'Ngày lập dự phòng', 'Ngày tổn thất', 'Số HSBT TPA', 'Số đơn BH', 'Tỷ giá', 'Mã lý do', 'Số tiền DPBT ntệ', 'Ghi chú', 'Lỗi']
            err = xlsx([H,
                        ['7608000005', 46345, 46331, 'TST-0001', POL, '1.000000000000000', '500105', '1500000.000000000000000', NOTE, 'Hồ sơ TPA có 2 số đơn BH'],
                        ['7608000005', 46345, 46332, 'TST-0002', POL, '1.000000000000000', '500105', '250000.000000000000000', NOTE, 'Lỗi khác ABC']],
                       fmt={(2, 2): '0', (3, 2): '0'})
            ev("""async(a)=>{S.tab='tool';S.tool.fn='fix';Object.assign(S.tool,{profile:'FHVI',file:mkf('full.xlsx',a[0]),fixFile:mkf('baoloi.xlsx',a[1]),fx:null});_al.length=0;await ACT.fixGo();}""", [full_b64, b64(with_ns_prefix(err))])
            r = ev("()=>{const x=S.tool.fx;return x&&{msgs:x.msgs.map(m=>m.t),fix:x.files.fix&&u64(x.files.fix.u8),rest:x.files.rest&&u64(x.files.rest.u8),names:[x.files.fix&&x.files.fix.name,x.files.rest&&x.files.rest.name]}}")
            check('Tên file: FIX LOI - TPA - … (bỏ chữ DPBT)', r and r['names'] == ['FIX LOI - FHVI - full.xlsx', 'LOI CON LAI - FHVI - full.xlsx'], r and r['names'])
            check('Fix lỗi chạy được với file báo lỗi kiểu hệ thống', r and r['fix'], ev('()=>_al'))
            if r and r['fix']:
                fx = list(read_xlsx(r['fix']).values())[0]; ci = fx[0].index('Số HSBT TPA')
                check('FIX LỖI: chỉ có hồ sơ lỗi 2 số đơn, đã thêm A', [x[ci] for x in fx[1:]] == ['TST-0001A'], [x[ci] for x in fx[1:]])
            if r and r['rest']:
                rs = list(read_xlsx(r['rest']).values())[0]; hd = rs[0]; row = rs[1]
                check('LỖI CÒN LẠI: có đúng 1 dòng lỗi khác', len(rs) == 2 and row[hd.index('Số HSBT TPA')] == 'TST-0002', rs)
                check('LỖI CÒN LẠI: cột Ngày ra ngày thật', isinstance(row[hd.index('Ngày lập dự phòng')], datetime.datetime) and row[hd.index('Ngày tổn thất')].date() == datetime.date(2026, 11, 6), row)
                check('LỖI CÒN LẠI: số dạng chữ → số thật', row[hd.index('Tỷ giá')] == 1 and row[hd.index('Số tiền DPBT ntệ')] == 250000, row)
                check('LỖI CÒN LẠI: Mã lý do giữ dạng chữ', row[hd.index('Mã lý do')] == '500105', row[hd.index('Mã lý do')])
            else:
                check('Có file LỖI CÒN LẠI', False, r and r['msgs'])

        # ============ 6. Đối chiếu số tiền + dấu ✓
        print('\n6. Đối chiếu số tiền')
        if full_b64:
            def tracker(n=3):
                ev("(n)=>{DB.tracker.rows=[cleanTracker({id:'t1',receivedDate:'2026-11-05',tpa:'FHVI',officialLetter:'01.11 - 04.11.26 - CV 9999',amount:2083000,caseCount:n,status:'Mới tiếp nhận',enteredDate:'',accountingDate:'',paymentNote:''})];}", n)
            HB = ['STT', 'Trạng thái', 'Số HSBT', 'Số HSBT TPA', 'NĐBH', 'Tổng tiền']
            def src(rows): return b64(with_ns_prefix(xlsx([HB] + rows)))
            good = [[1, 'TRANSFER_ACCOUNTING', 'C-1', 'TST-0001', 'A', 1500000], [2, 'TRANSFER_ACCOUNTING', 'C-2', 'TST-0002', 'B', 200000], [3, 'TRANSFER_ACCOUNTING', 'C-2b', 'TST-0002', 'B', 50000], [4, 'TRANSFER_ACCOUNTING', 'C-3', 'TST-0003', 'C', 333000]]
            def dc(srcb):
                ev("""async(a)=>{S.tab='tool';S.tool.fn='dc';Object.assign(S.tool,{file:mkf('DPBT.xlsx',a[0]),listFile:mkf('nguon.xlsx',a[1]),dc:null});_al.length=0;await ACT.dcGo();}""", [full_b64, srcb])
                return ev("()=>{const r=S.tool.dc;return r&&{err:r.rows.filter(x=>x.kind!=='ok').map(x=>[x.kind,x.no]),drafts:r.drafts.map(x=>x.no),trk:r.trk&&r.trk.t,mark:DB.tracker.rows[0].dcCheck||null}}")
            tracker(); r = dc(src(good))
            check('Khớp hết (HS trùng dòng được cộng tiền) → không lỗi', r and not r['err'] and not r['drafts'], r)
            check('Khớp hết + đúng số HS → gắn dấu ✓ cho đợt', r and r['mark'] and r['mark']['n'] == 3, r and r['trk'])
            check('Không lỗi → không hiện bảng danh sách', ev("()=>{renderApp();return !document.querySelector('.tw.dc')}"))
            tracker(4); r = dc(src(good))
            check('Lệch số lượng HS với đợt → không gắn dấu', r and not r['mark'] and 'lệch' in (r['trk'] or ''), r and r['trk'])
            tracker(); bad = [[1, 'TRANSFER_ACCOUNTING', 'C-1', 'TST-0001', 'A', 1499500], [2, 'DRAFT', 'C-2', 'TST-0002', 'B', 250000], [3, 'TRANSFER_ACCOUNTING', 'C-3', 'TST-0003', 'C', 333000], [4, 'TRANSFER_ACCOUNTING', 'C-9', 'TST-0009', 'Z', 1000]]
            r = dc(src(bad))
            check('Bắt lệch tiền, DRAFT, thiếu ở file cần dò', r and sorted(r['err']) == [['b', 'TST-0009'], ['lech', 'TST-0001']] and r['drafts'] == ['TST-0002'], r)
            check('HS DRAFT chỉ báo 1 lần (không báo thêm Thiếu ở nguồn)', r and not any(k == 'a' for k, _ in r['err']))
            check('Có lỗi → không gắn dấu', r and not r['mark'])
            check('Có lỗi → hiện bảng lỗi + nút tải', ev("()=>{renderApp();return !!document.querySelector('.tw.dc') && !!document.querySelector('[data-act=dcDl]')}"))

        # ============ 7. Theo dõi HS TPA
        print('\n7. Theo dõi HS TPA')
        tests = [('01.11 - 04.11.26 - CV 2700', 'FHVI'), ('684-2026 (T11 DOT 8)', 'PPY'), ('CV 123 PJICO/2026', 'INS'), ('CV 1820/26/FPT IS', 'FPTIS'), ('Công văn lạ', '')]
        got = ev("(a)=>a.map(x=>cvTpa(x[0]))", tests)
        check('Nhận đúng TPA theo mẫu Công văn', got == [t for _, t in tests], got)
        if full_b64:
            ev("()=>{DB.tracker.rows=[cleanTracker({id:'t1',receivedDate:'2026-11-05',tpa:'FHVI',officialLetter:'01.11 - 04.11.26 - CV 9999',amount:1,caseCount:3,status:'Mới tiếp nhận',enteredDate:'',accountingDate:'',paymentNote:''})];}")
            k = ev("()=>trkKhoMatch('FHVI','01.11 - 04.11.26 - CV 9999').n")
            check('Cột Kho đếm đúng hồ sơ đã lưu của đợt', k == 3, k)
        def save(tpa, cv, ans=True, rid=None):
            ev("""(a)=>{_cf.length=0;_cfAns=a[3];S.tab='trk';S.trk.edit=a[2]?Object.assign({},DB.tracker.rows.find(r=>r.id===a[2]),{entL:[],accL:[]}):{receivedDate:'18/11/2026',tpa:a[0],officialLetter:a[1],amount:1000,caseCount:5,status:'Mới tiếp nhận',tpc:false,entL:[],accL:[],paymentNote:''};if(a[2])Object.assign(S.trk.edit,{tpa:a[0],officialLetter:a[1]});}""", [tpa, cv, rid, ans])
            ev("async()=>{await ACT.trkSave();}")
            return ev("()=>({cf:_cf.slice(),n:DB.tracker.rows.length})")
        ev("()=>{DB.tracker.rows=[];}")
        r = save('FHVI', '684-2026 (T11 DOT 8)', ans=False)
        check('Công văn sai mẫu TPA → hỏi, bấm Sửa lại thì không lưu', r['n'] == 0 and any('giống mẫu PPY' in c for c in r['cf']), r)
        r = save('FHVI', '684-2026 (T11 DOT 8)', ans=True)
        check('Bấm Vẫn lưu → lưu được', r['n'] == 1, r)
        r = save('PPY', '685-2026 (T11 DOT 9)')
        check('Công văn đúng mẫu → không hỏi', r['n'] == 2 and not any('mẫu' in c for c in r['cf']), r)
        rid = ev("()=>{const r=DB.tracker.rows[1];r.dcCheck={at:new Date().toISOString(),n:5,sum:1000};return r.id}")
        save('PPY', '685-2026 (T11 DOT 9)', rid=rid)
        check('Sửa đợt vẫn giữ dấu ✓ đối chiếu', ev("(id)=>!!DB.tracker.rows.find(r=>r.id===id).dcCheck", rid))

        # ============ 8. TPC và bắt trùng đợt
        print('\n8. TPC và bắt trùng đợt')
        reset()
        conv()
        tpc_src = FHVI_SRC[:3] + [[1, 'TST-0101', POL, 'PHAM THI TPC', '001090000101', '04/04/1993', 777000, 'BENH VIEN THU D', 'Ngoại trú', 'Đau', '08/11/2026']]
        r = conv(tpc_src, tpc=True)
        check('Chuyển đổi có tick TPC → Ghi chú có đuôi " - TPC"', ev("()=>DB.claims.rows.filter(c=>/- TPC$/.test(c.note)).length") == 1, r and r['msgs'])
        ev("""()=>{S.conv.tpc=false;DB.tracker.rows=[cleanTracker({id:'rg',receivedDate:'2026-11-05',tpa:'FHVI',officialLetter:'01.11 - 04.11.26 - CV 9999',amount:2083000,caseCount:3,status:'Mới tiếp nhận',enteredDate:'',accountingDate:'',paymentNote:''}),
          cleanTracker({id:'tp',receivedDate:'2026-11-06',tpa:'FHVI',officialLetter:'01.11 - 04.11.26 - CV 9999',amount:777000,caseCount:1,tpc:true,status:'Mới tiếp nhận',enteredDate:'',accountingDate:'',paymentNote:''})];}""")
        k = ev("()=>[trkKhoMatch('FHVI','01.11 - 04.11.26 - CV 9999',false).n,trkKhoMatch('FHVI','01.11 - 04.11.26 - CV 9999',true).n]")
        check('Kho: đợt thường chỉ đếm hồ sơ thường, đợt TPC chỉ đếm hồ sơ TPC', k == [3, 1], k)
        m = ev("()=>[fxMatch('FHVI','GD 01.11 - 04.11.26'),fxMatch('FHVI','GD 01.11 - 04.11.26 - TPC')].map(x=>x.row?x.row.id:x.err)")
        check('Đối chiếu ✓ / Fix lỗi: Ghi chú TPC dò ra đợt TPC, thường dò ra đợt thường', m == ['rg', 'tp'], m)
        f = ev("()=>[true,false].map(t=>{const x=trkFindMatch({tpa:'FHVI',tpc:t,linkDate:'2026-11-20',texts:['GD 01.11 - 04.11.26'+(t?' - TPC':'')],caseCount:0,amount:0});return x.row?x.row.id:x.res})")
        check('Tự chuyển trạng thái: tick TPC dò đợt TPC, không tick dò đợt thường', f == ['tp', 'rg'], f)
        ev("""()=>{S.tab='trk';S.trk.year='2026';S.trk.applied={year:'2026',f:{...S.trk.f},all:true};renderApp();}""")
        cells = ev("()=>[...document.querySelectorAll('tr[data-act=trkEdit]')].map(tr=>tr.innerText.replace(/\\s+/g,' '))")
        check('Bảng Theo dõi: cột Kho hiện 3/3 và 1/1', any('Kho 3/3' in c for c in cells) and any('Kho 1/1' in c for c in cells), cells)
        def add(tpc, date, amt, ans):
            ev("""(a)=>{_cf.length=0;_al.length=0;_cfAns=a[3];S.tab='trk';S.trk.edit={receivedDate:a[1],tpa:'FHVI',officialLetter:'GD 01.11-04.11.26 CV9999',amount:a[2],caseCount:3,status:'Mới tiếp nhận',tpc:a[0],entL:[],accL:[],paymentNote:''};}""", [tpc, date, amt, ans])
            ev("async()=>{await ACT.trkSave();}")
            return ev("()=>({cf:_cf.slice(),al:_al.slice(),n:DB.tracker.rows.length})")
        r = add(False, '21/11/2026', 2083001, False)
        check('Cùng TPA + Công văn (gõ khác, khác ngày, khác tiền) → hỏi trùng đợt, Sửa lại thì không lưu', r['n'] == 2 and any('trùng đợt' in c for c in r['cf']), r)
        r = add(False, '21/11/2026', 2083001, True)
        check('Bấm Vẫn lưu → lưu được', r['n'] == 3, r)
        r = add(True, '22/11/2026', 5000, True)
        check('Thêm đợt TPC khi đã có đợt TPC cùng Công văn → hỏi trùng (chỉ nêu đợt TPC)', len(r['cf']) == 1 and 'trùng đợt' in r['cf'][0] and '· TPC' in r['cf'][0] and '2.083.000' not in r['cf'][0], r)
        ev("()=>{DB.tracker.rows=DB.tracker.rows.filter(x=>x.id==='rg');}")
        r = add(True, '22/11/2026', 5000, True)
        check('Đợt TPC cùng Công văn với đợt thường → không coi là trùng', r['n'] == 2 and not any('trùng đợt' in c for c in r['cf']), r)
        ev("()=>{_al.length=0;S.trk.edit={receivedDate:'05/11/2026',tpa:'FHVI',officialLetter:'01.11 - 04.11.26 - CV 9999',amount:2083000,caseCount:3,status:'Mới tiếp nhận',tpc:false,entL:[],accL:[],paymentNote:''};}")
        ev("async()=>{await ACT.trkSave();}")
        check('Trùng y hệt (ngày, CV, tiền, TPA, TPC) → chặn', any('đã có trong danh sách' in a for a in ev('()=>_al')) and ev('()=>DB.tracker.rows.length') == 2, ev('()=>_al'))

        # ============ 9. Giao diện tab Công cụ
        print('\n9. Giao diện tab Công cụ')
        tabs = ev("()=>{S.tab='tool';S.tool.fn='cccd';renderApp();return [...document.querySelectorAll('[data-act=toolFn]')].map(b=>b.innerText.trim())}")
        check('Thứ tự + tên mục mới', tabs == ['Bổ sung CCCD', 'Fix lỗi FTS', 'Check endors', 'Check nhập liệu', 'Update dữ liệu trống', 'DPBT quý'], tabs)
        def labels(fn): return ev("(fn)=>{Object.assign(S.tool,{fn,file:null,fixFile:null,listFile:null});renderApp();return [...document.querySelectorAll('input[type=file][data-k^=\"tool.\"]')].map(i=>{const w=i.closest('.fl')||i.closest('label').parentElement;const sp=w.querySelector(':scope > span');return (sp?sp.innerText:w.innerText).trim()})}", fn)
        for fn in ['cccd', 'fix', 'ps', 'dc']:
            lb = labels(fn)
            check(f'{fn}: File cần dò ở trên, File nguồn ở dưới', len(lb) == 2 and lb[0].startswith('File cần dò') and lb[1].startswith('File nguồn'), lb)

        # ============ 10. Bổ sung CCCD
        print('\n10. Bổ sung CCCD')
        reset()
        H = FHVI_SRC[2]
        def row(i, name, cid, dob): return [i, f'TST-CC{i}', POL, name, cid, dob, 1000, 'BV', 'Ngoại trú', 'X', '05/11/2026']
        goc = [H, row(1, 'NGUYEN CC MOT', '79192016939', '01/01/1990'), row(2, 'TRAN CC HAI', '', '02/02/1991'), row(3, 'LE CC BA', '', '03/03/1992'),
               row(4, 'PHAM CC BON', '001090000004', '04/04/1993'), row(5, 'VO CC NAM', '01234567890', '05/05/1994'), row(6, 'DO CC SAU', 'B1234567', '06/06/1995')]
        ds = [['Tên NĐBH', 'Ngày sinh', 'CCCD'], ['TRAN CC HAI', '02/02/1991', '12345678901'], ['LE CC BA', '03/03/1992', '183168684']]
        ev("""async(a)=>{S.tab='tool';S.tool.fn='cccd';Object.assign(S.tool,{profile:'FHVI',file:mkf('goc.xlsx',a[0]),listFile:mkf('ds.xlsx',a[1]),cc:null});_al.length=0;await ACT.ccGo();}""", [b64(xlsx(goc)), b64(xlsx(ds))])
        r = ev("()=>{const x=S.tool.cc;return x&&{msgs:x.msgs.map(m=>m.t),u8:x.file?u64(x.file.u8):null}}")
        check('Bổ sung CCCD chạy được', r and r['u8'], (r and r['msgs'], ev('()=>_al')))
        if r and r['u8']:
            wb = openpyxl.load_workbook(io.BytesIO(base64.b64decode(r['u8']))); ws = wb.active
            got = {ws.cell(i, 1).value: (ws.cell(i, 4).value, (ws.cell(i, 4).font.color.rgb if ws.cell(i, 4).font.color else None)) for i in range(2, ws.max_row + 1)}
            check('File gốc CCCD 11 số → thêm 0, chữ màu cam', got.get('TST-CC1', ('',))[0] == '079192016939' and 'C2410C' in str(got['TST-CC1'][1]), got.get('TST-CC1'))
            check('CCCD 11 số trong file danh sách → thêm 0 rồi dùng', got.get('TST-CC2', ('',))[0] == '012345678901', got.get('TST-CC2'))
            check('CMND 9 số giữ nguyên', got.get('TST-CC3', ('',))[0] == '183168684', got.get('TST-CC3'))
            check('CCCD đủ 12 số / hộ chiếu → không cần bổ sung', 'TST-CC4' not in got and 'TST-CC6' not in got, list(got))
            check('11 số bắt đầu bằng 0 → vẫn báo lỗi như cũ', 'LỖI' in str(got.get('TST-CC5', ('',))[0]), got.get('TST-CC5'))
            check('Thông báo có số hồ sơ tự thêm 0', any('tự thêm 0' in m for m in r['msgs']), r['msgs'])

        check('Không có lỗi JavaScript trong suốt quá trình', not errs, errs[:3])
        br.close()

    bad = [r for r in RESULTS if not r[1]]
    print(f"\n{'=' * 50}\n{len(RESULTS) - len(bad)}/{len(RESULTS)} mục đạt" + ('' if not bad else f" · {len(bad)} mục LỖI"))
    sys.exit(1 if bad else 0)

if __name__ == '__main__':
    main()
