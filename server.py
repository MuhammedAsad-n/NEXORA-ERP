import json,sqlite3,os,datetime,io,zipfile
from http.server import BaseHTTPRequestHandler,ThreadingHTTPServer
from urllib.parse import urlparse,parse_qs
from xml.sax.saxutils import escape
ROOT=os.path.dirname(os.path.abspath(__file__)); DB=os.path.join(ROOT,'nexora.db')
def db():
 c=sqlite3.connect(DB);c.row_factory=sqlite3.Row;return c
def init():
 c=db();c.executescript('''PRAGMA foreign_keys=ON; CREATE TABLE IF NOT EXISTS companies(id INTEGER PRIMARY KEY,name TEXT UNIQUE NOT NULL); CREATE TABLE IF NOT EXISTS units(id INTEGER PRIMARY KEY,company_id INTEGER NOT NULL,name TEXT NOT NULL,FOREIGN KEY(company_id) REFERENCES companies(id)); CREATE TABLE IF NOT EXISTS items(id INTEGER PRIMARY KEY,company_id INTEGER,unit_id INTEGER,name TEXT NOT NULL,code TEXT,category TEXT,uom TEXT NOT NULL DEFAULT 'pcs',stock REAL NOT NULL DEFAULT 0,min_stock REAL NOT NULL DEFAULT 0,FOREIGN KEY(company_id) REFERENCES companies(id)); CREATE TABLE IF NOT EXISTS movements(id INTEGER PRIMARY KEY,item_id INTEGER NOT NULL,kind TEXT NOT NULL,qty REAL NOT NULL,note TEXT,created TEXT NOT NULL,FOREIGN KEY(item_id) REFERENCES items(id)); CREATE TABLE IF NOT EXISTS workers(id INTEGER PRIMARY KEY,company_id INTEGER NOT NULL,unit_id INTEGER,name TEXT NOT NULL,FOREIGN KEY(company_id) REFERENCES companies(id)); CREATE TABLE IF NOT EXISTS rates(id INTEGER PRIMARY KEY,worker_id INTEGER NOT NULL,item_id INTEGER NOT NULL,rate REAL NOT NULL,FOREIGN KEY(worker_id) REFERENCES workers(id),FOREIGN KEY(item_id) REFERENCES items(id)); CREATE TABLE IF NOT EXISTS production(id INTEGER PRIMARY KEY,company_id INTEGER NOT NULL,unit_id INTEGER,worker_id INTEGER,item_id INTEGER NOT NULL,qty REAL NOT NULL,rejected REAL NOT NULL DEFAULT 0,status TEXT NOT NULL DEFAULT 'Approved',date TEXT NOT NULL,note TEXT,FOREIGN KEY(company_id) REFERENCES companies(id),FOREIGN KEY(worker_id) REFERENCES workers(id),FOREIGN KEY(item_id) REFERENCES items(id)); CREATE TABLE IF NOT EXISTS sales(id INTEGER PRIMARY KEY,company_id INTEGER NOT NULL,unit_id INTEGER,customer TEXT NOT NULL,item_id INTEGER NOT NULL,qty REAL NOT NULL,unit_price REAL NOT NULL,total REAL NOT NULL,date TEXT NOT NULL,note TEXT,FOREIGN KEY(company_id) REFERENCES companies(id),FOREIGN KEY(item_id) REFERENCES items(id)); CREATE TABLE IF NOT EXISTS shops(id INTEGER PRIMARY KEY,name TEXT UNIQUE NOT NULL,contact TEXT,address TEXT,note TEXT);''');# Upgrade older databases so catalog items can be created before assigning a company.
 # Add shop registry and multi-line sale grouping metadata to existing databases.
 for col,definition in [('shop_id','INTEGER'),('sale_group_id','INTEGER')]:
  if col not in [r['name'] for r in c.execute('PRAGMA table_info(sales)')]: c.execute(f'ALTER TABLE sales ADD COLUMN {col} {definition}')
 c.execute('PRAGMA foreign_keys=OFF')
 info=[r for r in c.execute('PRAGMA table_info(items)') if r['name']=='company_id']
 if info and info[0]['notnull']:
  c.execute('CREATE TABLE items_new(id INTEGER PRIMARY KEY,company_id INTEGER,unit_id INTEGER,name TEXT NOT NULL,code TEXT,category TEXT,uom TEXT NOT NULL DEFAULT \'pcs\',stock REAL NOT NULL DEFAULT 0,min_stock REAL NOT NULL DEFAULT 0)')
  c.execute('INSERT INTO items_new SELECT id,company_id,unit_id,name,code,category,uom,stock,min_stock FROM items')
  c.execute('DROP TABLE items');c.execute('ALTER TABLE items_new RENAME TO items')
 c.execute('PRAGMA foreign_keys=ON');c.commit();c.close()
def all(sql,args=()):
 c=db();r=[dict(x) for x in c.execute(sql,args)];c.close();return r
def execsql(sql,args=()):
 c=db();c.execute(sql,args);c.commit();c.close()
def safe(v):return escape(str(v if v is not None else ''))
def xlsx(sheets):
 b=io.BytesIO();names=list(sheets)
 with zipfile.ZipFile(b,'w',zipfile.ZIP_DEFLATED) as z:
  z.writestr('[Content_Types].xml','<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'+''.join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1,len(names)+1))+'</Types>')
  z.writestr('_rels/.rels','<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
  z.writestr('xl/workbook.xml','<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'+''.join(f'<sheet name="{safe(n[:31])}" sheetId="{i}" r:id="rId{i}"/>' for i,n in enumerate(names,1))+'</sheets></workbook>')
  z.writestr('xl/_rels/workbook.xml.rels','<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+''.join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1,len(names)+1))+f'<Relationship Id="rId{len(names)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
  z.writestr('xl/styles.xml','<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="3"><xf fontId="0" fillId="0" borderId="0" xfId="0"/><xf fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf></cellXfs></styleSheet>')
  for si,(name,rows) in enumerate(sheets.items(),1):
   maxcols=max((len(r) for r in rows),default=1);widths=[]
   for ci in range(maxcols):
    vals=[str(r[ci]) for r in rows if ci<len(r) and r[ci] is not None];widths.append(min(48,max(12,max((len(v) for v in vals),default=8)+3)))
   cols=''.join(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>' for i,w in enumerate(widths,1))
   xml=['<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView showGridLines="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>'+cols+'</cols><sheetData>']
   for ri,row in enumerate(rows,1):
    xml.append(f'<row r="{ri}" ht="{24 if ri==1 else 20}" customHeight="1">')
    for ci,val in enumerate(row,1):
     col='';n=ci
     while n:n,rem=divmod(n-1,26);col=chr(65+rem)+col
     cell=f'{col}{ri}';style=' s="1"' if ri==1 else ' s="2"'
     if val is None:continue
     if isinstance(val,(int,float)):xml.append(f'<c r="{cell}"{style}><v>{val}</v></c>')
     else:xml.append(f'<c r="{cell}" t="inlineStr"{style}><is><t>{safe(val)}</t></is></c>')
    xml.append('</row>')
   xml.append('</sheetData>')
   if len(rows)>1:
    lastcol='';n=maxcols
    while n:n,rem=divmod(n-1,26);lastcol=chr(65+rem)+lastcol
    xml.append(f'<autoFilter ref="A1:{lastcol}{len(rows)}"/>')
   xml.append('</worksheet>');z.writestr(f'xl/worksheets/sheet{si}.xml',''.join(xml))
 return b.getvalue()
class H(BaseHTTPRequestHandler):
 def send(self,status,body,typ='application/json'):
  self.send_response(status);self.send_header('Content-Type',typ);self.send_header('Cache-Control','no-store');self.end_headers();self.wfile.write(body if isinstance(body,bytes) else body.encode())
 def body(self):return json.loads(self.rfile.read(int(self.headers.get('Content-Length',0)) or 0))
 def do_GET(self):
  p=urlparse(self.path).path
  if p=='/':self.send(200,open(os.path.join(ROOT,'index.html'),'rb').read(),'text/html; charset=utf-8');return
  if p=='/api/data':
   d={'companies':all('SELECT * FROM companies'),'units':all('SELECT units.*,companies.name company FROM units JOIN companies ON companies.id=units.company_id'),'items':all('SELECT items.*,companies.name company,units.name unit FROM items LEFT JOIN companies ON companies.id=items.company_id LEFT JOIN units ON units.id=items.unit_id'),'movements':all('SELECT movements.*,items.name item,items.uom FROM movements JOIN items ON items.id=movements.item_id ORDER BY movements.id DESC'),'workers':all('SELECT workers.*,companies.name company,units.name unit FROM workers JOIN companies ON companies.id=workers.company_id LEFT JOIN units ON units.id=workers.unit_id'),'rates':all('SELECT rates.*,workers.name worker,items.name item,items.uom FROM rates JOIN workers ON workers.id=rates.worker_id JOIN items ON items.id=rates.item_id'),'production':all('SELECT production.*,companies.name company,units.name unit,workers.name worker,items.name item,items.uom FROM production JOIN companies ON companies.id=production.company_id LEFT JOIN units ON units.id=production.unit_id LEFT JOIN workers ON workers.id=production.worker_id JOIN items ON items.id=production.item_id ORDER BY production.id DESC'),'shops':all('SELECT * FROM shops ORDER BY name'),'sales':all('SELECT sales.*,companies.name company,units.name unit,items.name item,items.uom FROM sales JOIN companies ON companies.id=sales.company_id LEFT JOIN units ON units.id=sales.unit_id JOIN items ON items.id=sales.item_id ORDER BY sales.id DESC')};self.send(200,json.dumps(d));return
  if p=='/api/export':
   qs=parse_qs(urlparse(self.path).query); company_id=int(qs.get('company_id',[0])[0] or 0); sheets={}
   for title,query,headers in [('Companies','SELECT id,name FROM companies',['Company ID','Company Name']),('Factory Units','SELECT id,company_id,name FROM units',['Unit ID','Company ID','Unit Name']),('Inventory','SELECT id,company_id,unit_id,name,code,category,uom,stock,min_stock FROM items',['Item ID','Company ID','Unit ID','Item','Code','Category','Unit','Current Stock','Low Stock Level']),('Stock Ledger','SELECT movements.id,items.name,movements.kind,movements.qty,items.uom,movements.note,movements.created FROM movements JOIN items ON items.id=movements.item_id ORDER BY movements.id',['Entry ID','Item','Movement','Quantity','Unit','Reason','Date & Time']),('Workers','SELECT id,company_id,unit_id,name FROM workers',['Worker ID','Company ID','Unit ID','Worker Name']),('Piecework Rates','SELECT workers.name,items.name,rates.rate,items.uom FROM rates JOIN workers ON workers.id=rates.worker_id JOIN items ON items.id=rates.item_id',['Worker','Product','Rate per Unit','Unit']),('Production','SELECT production.date,companies.name,units.name,workers.name,items.name,production.qty,production.rejected,production.qty-production.rejected,production.status,production.note FROM production JOIN companies ON companies.id=production.company_id LEFT JOIN units ON units.id=production.unit_id LEFT JOIN workers ON workers.id=production.worker_id JOIN items ON items.id=production.item_id ORDER BY production.date',['Date','Company','Factory Unit','Worker','Product','Produced','Rejected','Accepted','Status','Notes'])]:
    data=all(query);sheets[title]=[headers]+[list(r.values()) for r in data]
   if company_id:
    co=all('SELECT id,name FROM companies WHERE id=?',(company_id,))
    if not co: raise ValueError('Company not found')
    cname=str(co[0]['name'])
    sheets={}
    sheets['Company']=[['Company'],[cname]]
    sheets['Factory Units']=[['Unit'],*[[r['name']] for r in all('SELECT name FROM units WHERE company_id=?',(company_id,))]]
    sheets['Inventory']=[['Item','Code','Category','Unit','Current Stock','Low Stock Level','Factory Unit']]+[[r['name'],r['code'],r['category'],r['uom'],r['stock'],r['min_stock'],r['unit']] for r in all('SELECT items.*,units.name unit FROM items LEFT JOIN units ON units.id=items.unit_id WHERE items.company_id=?',(company_id,))]
    sheets['Stock Ledger']=[['Date','Item','Movement','Quantity','Unit','Reason']]+[[r['created'],r['item'],r['kind'],r['qty'],r['uom'],r['note']] for r in all('SELECT movements.*,items.name item,items.uom FROM movements JOIN items ON items.id=movements.item_id WHERE items.company_id=? ORDER BY movements.id',(company_id,))]
    sheets['Workers']=[['Worker','Factory Unit']]+[[r['name'],r['unit']] for r in all('SELECT workers.name,units.name unit FROM workers LEFT JOIN units ON units.id=workers.unit_id WHERE workers.company_id=?',(company_id,))]
    sheets['Production']=[['Date','Unit','Worker','Product','Produced','Rejected','Accepted','Status','Notes']]+[[r['date'],r['unit'],r['worker'],r['item'],r['qty'],r['rejected'],r['qty']-r['rejected'],r['status'],r['note']] for r in all('SELECT production.*,units.name unit,workers.name worker,items.name item FROM production LEFT JOIN units ON units.id=production.unit_id LEFT JOIN workers ON workers.id=production.worker_id JOIN items ON items.id=production.item_id WHERE production.company_id=? ORDER BY production.date',(company_id,))]
    sheets['Piecework Rates']=[['Worker','Product','Rate','Unit']]+[[r['worker'],r['item'],r['rate'],r['uom']] for r in all('SELECT workers.name worker,items.name item,rates.rate,items.uom FROM rates JOIN workers ON workers.id=rates.worker_id JOIN items ON items.id=rates.item_id WHERE workers.company_id=?',(company_id,))]
    sheets['Sales']=[['Date','Company','Shop / Customer','Item','Quantity','Unit Price','Total','Notes']]+[[r['date'],r['company'],r['customer'],r['item'],r['qty'],r['unit_price'],r['total'],r['note']] for r in all('SELECT sales.*,companies.name company,units.name unit,items.name item FROM sales JOIN companies ON companies.id=sales.company_id LEFT JOIN units ON units.id=sales.unit_id JOIN items ON items.id=sales.item_id WHERE sales.company_id=? ORDER BY sales.date',(company_id,))]
    sheets['Wage Summary']=[['Worker','Product','Rate','Unit','Accepted Quantity','Gross Earnings']]
    for r in all('SELECT workers.name worker,items.name item,rates.rate,items.uom,workers.id wid,items.id iid FROM rates JOIN workers ON workers.id=rates.worker_id JOIN items ON items.id=rates.item_id WHERE workers.company_id=?',(company_id,)):
     ps=all("SELECT qty,rejected FROM production WHERE worker_id=? AND item_id=? AND status='Approved'",(r['wid'],r['iid']));q=sum(max(0,pp['qty']-pp['rejected']) for pp in ps);sheets['Wage Summary'].append([r['worker'],r['item'],r['rate'],r['uom'],q,round(q*r['rate'],2)])
    raw=xlsx(sheets);self.send(200,raw,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');return
   sheets['Shops']=[['Shop Name','Contact','Address','Notes']]+[[r['name'],r['contact'],r['address'],r['note']] for r in all('SELECT * FROM shops ORDER BY name')]
   # Earnings summary
   rates=all('SELECT workers.name worker,items.name item,rates.rate,items.uom,workers.id wid,items.id iid FROM rates JOIN workers ON workers.id=rates.worker_id JOIN items ON items.id=rates.item_id');earn=[['Worker','Product','Rate','Unit','Accepted Quantity','Gross Earnings']]
   for r in rates:
    ps=all("SELECT qty,rejected FROM production WHERE worker_id=? AND item_id=? AND status='Approved'",(r['wid'],r['iid']));q=sum(max(0,p['qty']-p['rejected']) for p in ps);earn.append([r['worker'],r['item'],r['rate'],r['uom'],q,round(q*r['rate'],2)])
   # Company-specific worksheets, easier to review without filtering across companies
   for co in all('SELECT id,name FROM companies'):
    cid=co['id']; cname=str(co['name'])[:20]
    sheets[(cname+' Inventory')[:31]]=[['Item','Code','Category','Unit','Stock','Low Stock']]+[[r['name'],r['code'],r['category'],r['uom'],r['stock'],r['min_stock']] for r in all('SELECT * FROM items WHERE company_id=?',(cid,))]
    sheets[(cname+' Production')[:31]]=[['Date','Unit','Worker','Product','Produced','Rejected','Accepted','Status','Notes']]+[[r['date'],r['unit'],r['worker'],r['item'],r['qty'],r['rejected'],r['qty']-r['rejected'],r['status'],r['note']] for r in all('SELECT production.*,units.name unit,workers.name worker,items.name item FROM production LEFT JOIN units ON units.id=production.unit_id LEFT JOIN workers ON workers.id=production.worker_id JOIN items ON items.id=production.item_id WHERE production.company_id=? ORDER BY production.date',(cid,))]
   sheets['Sales']=[['Date','Company','Shop / Customer','Item','Quantity','Unit Price','Total','Notes']]+[[r['date'],r['company'],r['customer'],r['item'],r['qty'],r['unit_price'],r['total'],r['note']] for r in all('SELECT sales.*,companies.name company,units.name unit,items.name item FROM sales JOIN companies ON companies.id=sales.company_id LEFT JOIN units ON units.id=sales.unit_id JOIN items ON items.id=sales.item_id ORDER BY sales.date')];sheets['Wage Summary']=earn;raw= xlsx(sheets);self.send(200,raw,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');self.send_response if False else None;return
  self.send(404,'Not found','text/plain')
 def do_POST(self):
  p=urlparse(self.path).path;d=self.body()
  try:
   if p=='/api/companies':
    n=d.get('name','').strip();assert n,'Enter a company name';execsql('INSERT INTO companies(name) VALUES(?)',(n,))
   elif p=='/api/units':
    n=d.get('name','').strip();assert n and d.get('company_id'),'Choose company and enter unit name';execsql('INSERT INTO units(company_id,name) VALUES(?,?)',(int(d['company_id']),n))
   elif p=='/api/items':
    n=d.get('name','').strip();assert n,'Enter item name';stock=float(d.get('stock',0));assert stock>=0,'Opening stock cannot be negative';execsql('INSERT INTO items(company_id,unit_id,name,code,category,uom,stock,min_stock) VALUES(?,?,?,?,?,?,?,?)',(int(d['company_id']) if d.get('company_id') else None,d.get('unit_id') or None,n,d.get('code',''),d.get('category',''),d.get('uom','pcs'),stock,float(d.get('min_stock',0))));item=all('SELECT id FROM items WHERE name=? ORDER BY id DESC',(n,))[0];
    if stock>0:execsql('INSERT INTO movements(item_id,kind,qty,note,created) VALUES(?,?,?,?,?)',(item['id'],'Opening Stock',stock,'Opening balance',datetime.datetime.now().isoformat(timespec='seconds')))
   elif p=='/api/assign-catalog':
    source=all('SELECT * FROM items WHERE id=? AND company_id IS NULL',(int(d['item_id']),));assert source,'Choose a shared catalog item';assert d.get('company_id'),'Choose a company';it=source[0];execsql('INSERT INTO items(company_id,unit_id,name,code,category,uom,stock,min_stock) VALUES(?,?,?,?,?,?,0,?)',(int(d['company_id']),d.get('unit_id') or None,it['name'],it['code'],it['category'],it['uom'],it['min_stock']))
   elif p=='/api/movement':
    it=all('SELECT * FROM items WHERE id=?',(int(d['item_id']),))[0];kind=d['kind'];q=float(d['qty']);assert q>0,'Quantity must be greater than zero';delta=q if kind=='Receipt' else -q if kind=='Issue' else q-it['stock'];assert it['stock']+delta>=0,'This issue would make stock negative. Check the quantity.';execsql('UPDATE items SET stock=stock+? WHERE id=?',(delta,it['id']));execsql('INSERT INTO movements(item_id,kind,qty,note,created) VALUES(?,?,?,?,?)',(it['id'],kind,q,d.get('note',''),datetime.datetime.now().isoformat(timespec='seconds')))
   elif p=='/api/workers':
    n=d.get('name','').strip();assert n and d.get('company_id'),'Choose company and enter worker name';execsql('INSERT INTO workers(company_id,unit_id,name) VALUES(?,?,?)',(int(d['company_id']),d.get('unit_id') or None,n))
   elif p=='/api/rates':
    rate=float(d['rate']);assert rate>=0,'Rate cannot be negative';wid=int(d['worker_id']);iid=int(d['item_id']);existing=all('SELECT id FROM rates WHERE worker_id=? AND item_id=?',(wid,iid));
    if existing: execsql('UPDATE rates SET rate=? WHERE id=?',(rate,existing[0]['id']))
    else: execsql('INSERT INTO rates(worker_id,item_id,rate) VALUES(?,?,?)',(wid,iid,rate))
   elif p=='/api/bulk-rates':
    company_id=int(d['company_id']);item_id=int(d['item_id']);rows=d.get('rows') or [];assert rows,'No worker rates to save';
    c=db();c.execute('BEGIN')
    try:
     for row in rows:
      wid=int(row['worker_id']);rate=float(row['rate']);assert rate>=0,'Rates cannot be negative';worker=c.execute('SELECT company_id,unit_id FROM workers WHERE id=?',(wid,)).fetchone();assert worker and worker['company_id']==company_id,'Worker does not belong to selected company';assert worker['unit_id'] is None or worker['unit_id'] in [int(x) for x in d.get('unit_ids',[])],'Worker is outside the selected units';existing=c.execute('SELECT id FROM rates WHERE worker_id=? AND item_id=?',(wid,item_id)).fetchone()
      if existing:c.execute('UPDATE rates SET rate=? WHERE id=?',(rate,existing['id']))
      else:c.execute('INSERT INTO rates(worker_id,item_id,rate) VALUES(?,?,?)',(wid,item_id,rate))
     c.commit()
    except: c.rollback();raise
    finally:c.close()
   elif p=='/api/shops':
    name=d.get('name','').strip();assert name,'Enter shop name';execsql('INSERT INTO shops(name,contact,address,note) VALUES(?,?,?,?)',(name,d.get('contact','').strip(),d.get('address','').strip(),d.get('note','').strip()))
   elif p=='/api/sales':
    assert d.get('company_id') and d.get('shop_id'),'Choose company and shop';lines=d.get('lines') or ([{'item_id':d.get('item_id'),'qty':d.get('qty'),'unit_price':d.get('unit_price')}] if d.get('item_id') else []);assert lines,'Add at least one product line';shop=all('SELECT * FROM shops WHERE id=?',(int(d['shop_id']),));assert shop,'Choose a registered shop';gid=int(all('SELECT COALESCE(MAX(sale_group_id),0)+1 g FROM sales')[0]['g']);date=d.get('date') or datetime.date.today().isoformat();note_text=d.get('note','')
    c=db();c.execute('BEGIN')
    try:
     for line in lines:
      qty=float(line['qty']);price=float(line['unit_price']);iid=int(line['item_id']);assert qty>0 and price>=0,'Each quantity must be positive and price cannot be negative';item=c.execute('SELECT id,company_id FROM items WHERE id=?',(iid,)).fetchone();assert item,'Invalid item';assert item['company_id'] is None or item['company_id']==int(d['company_id']),'Item belongs to another company';total=round(qty*price,2);c.execute('INSERT INTO sales(company_id,customer,item_id,qty,unit_price,total,date,note,shop_id,sale_group_id) VALUES(?,?,?,?,?,?,?,?,?,?)',(int(d['company_id']),shop[0]['name'],iid,qty,price,total,date,note_text,int(d['shop_id']),gid))
     c.commit()
    except: c.rollback();raise
    finally: c.close()
   elif p=='/api/production':
    qty=float(d['qty']);rej=float(d.get('rejected',0));assert qty>0 and 0<=rej<=qty,'Produced quantity must be positive and rejected quantity cannot exceed produced quantity';assert d.get('company_id') and d.get('item_id'),'Choose company and product';coid=int(d['company_id']);uid=int(d['unit_id']) if d.get('unit_id') else None;iid=int(d['item_id']);item=all('SELECT * FROM items WHERE id=?',(iid,));assert item,'Choose a valid product';status=d.get('status','Approved');accepted=qty-rej;date=d.get('date') or datetime.date.today().isoformat();c=db();c.execute('BEGIN')
    try:
     # A shared catalog item becomes a company inventory item on first production entry.
     if item[0]['company_id'] is None:
      c.execute('INSERT INTO items(company_id,unit_id,name,code,category,uom,stock,min_stock) VALUES(?,?,?,?,?,?,0,?)',(coid,uid,item[0]['name'],item[0]['code'],item[0]['category'],item[0]['uom'],item[0]['min_stock']));iid=c.execute('SELECT last_insert_rowid()').fetchone()[0];c.execute('INSERT INTO rates(worker_id,item_id,rate) SELECT worker_id,?,rate FROM rates JOIN workers ON workers.id=rates.worker_id WHERE workers.company_id=? AND rates.item_id=?',(iid,coid,int(d['item_id'])))
     else: assert item[0]['company_id']==coid,'Selected item belongs to a different company'
     c.execute('INSERT INTO production(company_id,unit_id,worker_id,item_id,qty,rejected,status,date,note) VALUES(?,?,?,?,?,?,?,?,?)',(coid,uid,int(d['worker_id']) if d.get('worker_id') else None,iid,qty,rej,status,date,d.get('note','')))
     if status=='Approved' and accepted>0:
      c.execute('UPDATE items SET stock=stock+? WHERE id=?',(accepted,iid));c.execute('INSERT INTO movements(item_id,kind,qty,note,created) VALUES(?,?,?,?,?)',(iid,'Production Output',accepted,'Accepted production output',datetime.datetime.now().isoformat(timespec='seconds')))
     c.commit()
    except: c.rollback();raise
    finally: c.close()
   else:self.send(404,'Not found');return
   self.send(200,json.dumps({'ok':True}))
  except Exception as e:self.send(400,json.dumps({'error':str(e)}))
 def log_message(self,*args):pass
init();print('NEXORA ERP running at http://127.0.0.1:8765');ThreadingHTTPServer(('127.0.0.1',8765),H).serve_forever()
