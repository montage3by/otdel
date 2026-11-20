# Converts deck slides (HTML in «Отдел» style) into an editable PPTX.
import sys,json,os,re,math
from bs4 import BeautifulSoup,NavigableString
from pptx import Presentation
from pptx.util import Emu,Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN,MSO_ANCHOR
ROOT,OUT,ASSETS=sys.argv[1],sys.argv[2],sys.argv[3]
EXTRA=json.loads(sys.argv[4]) if len(sys.argv)>4 else {}
PX=6350
def E(px): return Emu(int(px*PX))
NAVY='1E3250';LIGHT='F7F9FC';OR='FF7928';DOR='C2540C';CARD='E9EEF6';GREY='4A5D78';DCARD='2A4166';LGREY='C9D4E3'
HF='Raleway';BF='Manrope'
def rgb(h): return RGBColor.from_string(h.lstrip('#').upper())
prs=Presentation(); prs.slide_width=E(1920); prs.slide_height=E(1080)
blank=prs.slide_layouts[6]
deck=json.load(open(os.path.join(ROOT,'project/deck.json'),encoding='utf-8'))
def rect(sl,x,y,w,h,fill,shape=MSO_SHAPE.RECTANGLE,radius=None):
    s=sl.shapes.add_shape(shape,E(x),E(y),E(w),E(h)); s.fill.solid(); s.fill.fore_color.rgb=rgb(fill); s.line.fill.background(); s.shadow.inherit=False
    if shape==MSO_SHAPE.ROUNDED_RECTANGLE and radius is not None: s.adjustments[0]=min(0.5,radius/max(1,min(w,h)))
    return s
def runs_of(node,bold=False,color=None):
    out=[]
    for c in node.children:
        if isinstance(c,NavigableString):
            t=str(c)
            if t: out.append((t,bold,color))
        elif c.name=='br': out.append(('\n',bold,color))
        else:
            st=c.get('style','') or ''
            m=re.search(r'color:\s*#([0-9A-Fa-f]{6})',st)
            out+=runs_of(c,bold or c.name in('b','strong'),m.group(1) if m else color)
    # collapse whitespace
    res=[]
    for t,b,co in out:
        t=re.sub(r'\s+',' ',t)
        res.append((t,b,co))
    if res: res[0]=(res[0][0].lstrip(),)+res[0][1:]; res[-1]=(res[-1][0].rstrip(),)+res[-1][1:]
    return res
def add_par(tf,runs,size,color,font=BF,bold=False,first=False,space_after=0,align=None,bullet=False,line=1.3):
    p=tf.paragraphs[0] if first else tf.add_paragraph()
    if align: p.alignment=align
    p.space_after=Pt(space_after*0.5); p.line_spacing=line
    if bullet: runs=[('•  ',False,None)]+runs
    for t,b,co in runs:
        r=p.add_run(); r.text=t; f=r.font; f.size=Pt(size*0.5); f.name=font; f.bold=bold or b; f.color.rgb=rgb(co or color)
    return p
def tb(sl,x,y,w,h,anchor=MSO_ANCHOR.TOP):
    t=sl.shapes.add_textbox(E(x),E(y),E(w),E(h)); tf=t.text_frame; tf.word_wrap=True
    tf.margin_left=tf.margin_right=tf.margin_top=tf.margin_bottom=0; tf.vertical_anchor=anchor; return tf
def lines(text,size,width,cw=0.53):
    if not text: return 0
    per=max(1,int(width/(size*cw)))
    return sum(max(1,math.ceil(len(seg)/per)) for seg in text.split('\n'))
def txt(n): return re.sub(r'\s+',' ',n.get_text(' ',strip=True))
# ---- element measuring/drawing
def card_items(c):
    items=[]
    for ch in c.find_all(recursive=False):
        if ch.name=='p' and 'num' in (ch.get('class') or []): items.append(('num',ch))
        elif ch.name=='h3': items.append(('h3',ch))
        elif ch.name=='p': items.append(('txt',ch))
        elif ch.name=='ul': items.append(('ul',ch))
    return items
def ul_size(u):
    m=re.search(r'font-size:\s*(\d+)px',u.get('style','') or ''); return int(m.group(1)) if m else 26
def card_h(c,w):
    iw=w-64; h=64
    it=card_items(c)
    for k,n in it:
        if k=='num': h+=lines(txt(n),44,iw,0.62)*44*1.15
        elif k=='h3': h+=lines(txt(n),32,iw,0.6)*32*1.3
        elif k=='txt': h+=lines(txt(n),24,iw,0.58)*24*1.45
        elif k=='ul':
            s=ul_size(n); h+=sum(lines(txt(li),s,iw-30)*s*1.4+10 for li in n.find_all('li'))
    h+=14*max(0,len(it)-1)+16
    return h
def draw_card(sl,c,x,y,w,h,dark):
    navy='navy' in (c.get('class') or [])
    fill=NAVY if navy else (DCARD if dark else CARD)
    rect(sl,x,y,w,h,fill,MSO_SHAPE.ROUNDED_RECTANGLE,16)
    rect(sl,x+8,y,w-16,6,OR)
    light=navy or dark
    tf=tb(sl,x+32,y+32,w-64,h-64); first=True
    for k,n in card_items(c):
        if k=='num': add_par(tf,[(txt(n),True,None)],44,OR if light else DOR,HF,True,first,12,line=1.05)
        elif k=='h3': add_par(tf,runs_of(n),32,'F7F9FC' if light else NAVY,HF,True,first,12,line=1.15)
        elif k=='txt': add_par(tf,runs_of(n),24,LGREY if light else GREY,BF,False,first,12)
        elif k=='ul':
            s=ul_size(n)
            for li in n.find_all('li'):
                add_par(tf,runs_of(li),s,LGREY if light else GREY,BF,False,first,10,bullet=True); first=False
        first=False
def table_rows(t):
    rows=[]
    for tr in t.find_all('tr'):
        rows.append([(c.name=='th',c) for c in tr.find_all(['th','td'])])
    return rows
def tbl_size(t):
    m=re.search(r'font-size:\s*(\d+)px',t.get('style','') or ''); return int(m.group(1)) if m else 26
def col_widths(t,w):
    ws=[]
    for c in t.find('tr').find_all(['th','td']):
        m=re.search(r'width:\s*([\d.]+)%',c.get('style','') or ''); ws.append(float(m.group(1)) if m else None)
    n=len(ws); known=sum(x for x in ws if x); unk=[i for i,x in enumerate(ws) if not x]
    rest=(100-known)/len(unk) if unk else 0
    ws=[x if x else rest for x in ws]; s=sum(ws)
    return [w*x/s for x in ws]
def table_h(t,w):
    s=tbl_size(t); cw=col_widths(t,w); h=0
    for row in table_rows(t):
        h+=max(lines(txt(c),s,cw[i]-40)*s*1.35 for i,(th,c) in enumerate(row))+28
    return h
def draw_table(sl,t,x,y,w,h,dark):
    s=tbl_size(t); rows=table_rows(t); cw=col_widths(t,w)
    gt=sl.shapes.add_table(len(rows),len(rows[0]),E(x),E(y),E(w),E(h)).table
    for i,wid in enumerate(cw): gt.columns[i].width=E(wid)
    for r,row in enumerate(rows):
        rh=max(lines(txt(c),s,cw[i]-40)*s*1.35 for i,(th,c) in enumerate(row))+28
        gt.rows[r].height=E(rh)
        for ci,(th,c) in enumerate(row):
            cell=gt.cell(r,ci); cell.margin_left=cell.margin_right=E(20); cell.margin_top=cell.margin_bottom=E(12)
            cell.fill.solid()
            if th: cell.fill.fore_color.rgb=rgb(NAVY)
            else: cell.fill.fore_color.rgb=rgb('FFFFFF' if r%2 else 'EEF2F8')
            tf=cell.text_frame; tf.word_wrap=True
            add_par(tf,runs_of(c),s,'FFFFFF' if th else NAVY,BF,th,True,line=1.2)
def child_h(n,w):
    cls=n.get('class') or []
    if n.name=='table': return table_h(n,w)
    if 'card' in cls: return card_h(n,w)
    return 0
def grid_cols(g):
    cls=g.get('class') or []; st=g.get('style','') or ''
    m=re.search(r'grid-template-columns:\s*([^;]+)',st)
    if m:
        parts=m.group(1).split(); fr=[float(p.replace('fr','')) for p in parts]; return fr
    n=int(re.search(r'grid(\d)',' '.join(cls)).group(1)); return [1]*n
def draw_grid(sl,g,x,y,w,dark,do=True):
    fr=grid_cols(g); n=len(fr); gap=24; avail=w-gap*(n-1); cws=[avail*f/sum(fr) for f in fr]
    kids=[k for k in g.find_all(recursive=False)]
    rows=[kids[i:i+n] for i in range(0,len(kids),n)]; cy=y
    rhs=[max(child_h(k,cws[i]) for i,k in enumerate(row)) for row in rows]
    tot=sum(rhs)+gap*(len(rows)-1)
    if y+tot>950: f=(950-y-gap*(len(rows)-1))/sum(rhs); rhs=[r*f for r in rhs]
    for row,rh in zip(rows,rhs):
        if do:
            cx=x
            for i,k in enumerate(row):
                if k.name=='table': draw_table(sl,k,cx,cy,cws[i],rh,dark)
                elif 'card' in (k.get('class') or []): draw_card(sl,k,cx,cy,cws[i],rh,dark)
                elif k.name=='p': draw_para(sl,k,cx,cy,cws[i],dark)
                cx+=cws[i]+gap
        cy+=rh+gap
    return cy-gap-y
def para_h(p,w):
    cls=p.get('class') or []; s=32 if 'lead' in cls else 24; st=p.get('style','') or ''
    pad=40 if 'background' in st else 0
    return lines(txt(p),s,w-(56 if pad else 0))*s*1.4+pad
def draw_para(sl,p,x,y,w,dark):
    cls=p.get('class') or []; st=p.get('style','') or ''; s=32 if 'lead' in cls else 24
    h=para_h(p,w)
    if 'background' in st:
        rect(sl,x,y,w,h,OR,MSO_SHAPE.ROUNDED_RECTANGLE,16); tf=tb(sl,x+28,y+20,w-56,h-40); col=NAVY
    else:
        tf=tb(sl,x,y,w,h); col=('F7F9FC' if dark else NAVY) if 'lead' in cls else (LGREY if dark else GREY)
    add_par(tf,runs_of(p),s,col,BF,False,True)
    return h
def footer(sl,dark,page):
    sl.shapes.add_picture(os.path.join(ASSETS,'otdel-white.png' if dark else 'otdel-navy.png'),E(128),E(1080-64-24),height=E(24))
    if page:
        tf=tb(sl,1920-128-300,1080-64-30,300,32); add_par(tf,[(page,False,None)],24,'A9B8CD' if dark else GREY,BF,False,True,align=PP_ALIGN.RIGHT)
def bg(sl,dark): 
    f=sl.background.fill; f.solid(); f.fore_color.rgb=rgb(NAVY if dark else LIGHT)
for sid in deck['order']:
    h=open(os.path.join(ROOT,'project/slides',sid+'.html'),encoding='utf-8').read()
    soup=BeautifulSoup(h,'html.parser'); sec=soup.find('section')
    dark='dark' in (sec.get('class') or [])
    sl=prs.slides.add_slide(blank); bg(sl,dark)
    page=sec.find('p',class_='page'); page=page.get_text(strip=True) if page else ''
    if sid=='cover':
        x=128
        sl.shapes.add_picture(os.path.join(ASSETS,'otdel-white.png'),E(128),E(96),height=E(40))
        tf=tb(sl,470,88,60,56); add_par(tf,[('×',False,None)],48,OR,BF,False,True)
        cl=EXTRA.get('client_logo')
        if cl: sl.shapes.add_picture(cl,E(540),E(96-8),height=E(EXTRA.get('client_logo_h',56)))
        ps=[p for p in sec.find_all(['p','h1'],recursive=False) if 'bottom' not in (p.get('style') or '')]
        hs=[lines(txt(p),104 if p.name=='h1' else (32 if 'uppercase' in (p.get('style') or '') else 36),1600,0.6 if p.name=='h1' else 0.55)*(104*1.05 if p.name=='h1' else (36*1.35)) for p in ps]
        total=6+40+sum(hs)+40*len(ps); y=(1080-total)/2
        rect(sl,128,y,120,6,OR); y+=46
        for p,hh in zip(ps,hs):
            st=p.get('style') or ''; tf=tb(sl,128,y,1640,hh+10)
            if p.name=='h1': add_par(tf,runs_of(p),104,'F7F9FC',HF,True,True,line=1.0)
            elif 'uppercase' in st: add_par(tf,[(txt(p).upper(),True,None)],32,LGREY,BF,True,True)
            else: add_par(tf,runs_of(p),36,'D6DEEA',BF,False,True,line=1.3)
            y+=hh+40
        foot=[p for p in sec.find_all('p',recursive=False) if 'bottom' in (p.get('style') or '')]
        if foot: tf=tb(sl,128,1080-64-30,1200,32); add_par(tf,runs_of(foot[0]),24,'A9B8CD',BF,False,True)
        continue
    if sec.find('div',class_='head') is None:  # separator
        ps=[p for p in sec.find_all(['p','h1','h2'],recursive=False) if 'page' not in (p.get('class') or [])]
        hs=[lines(txt(p),104,1600,0.6)*110 if p.name=='h1' else lines(txt(p),36,1450)*50 for p in ps]
        y=(1080-(46+sum(hs)+36*len(ps)))/2; rect(sl,128,y,120,6,OR); y+=46
        for p,hh in zip(ps,hs):
            tf=tb(sl,128,y,1640,hh+10); st=p.get('style') or ''
            if p.name in('h1','h2'): add_par(tf,runs_of(p),104 if p.name=='h1' else 72,'F7F9FC',HF,True,True,line=1.0)
            elif 'uppercase' in st: add_par(tf,[(txt(p).upper(),True,None)],32,LGREY,BF,True,True)
            else: add_par(tf,runs_of(p),36,'D6DEEA',BF,False,True)
            y+=hh+36
        footer(sl,True,page); continue
    head=sec.find('div',class_='head'); eb=head.find('p',class_='eyebrow'); h2=head.find('h2')
    ebt=txt(eb).upper(); ew=len(ebt)*24*0.68+48
    r=rect(sl,128,128,ew,48,OR if dark else NAVY,MSO_SHAPE.ROUNDED_RECTANGLE,24)
    tf=r.text_frame; tf.margin_left=tf.margin_right=E(20); tf.margin_top=tf.margin_bottom=0; tf.vertical_anchor=MSO_ANCHOR.MIDDLE; tf.word_wrap=False
    add_par(tf,[(ebt,True,None)],24,NAVY if dark else 'FFFFFF',BF,True,True)
    hl=lines(txt(h2),72,1664,0.64); tf=tb(sl,128,196,1664,hl*82)
    add_par(tf,runs_of(h2),72,'F7F9FC' if dark else NAVY,HF,True,True,line=1.05)
    y=196+hl*82+44
    for n in sec.find_all(recursive=False):
        cls=n.get('class') or []
        if n is head or 'logo' in cls or 'page' in cls or n.name=='img': continue
        if n.name=='div' and any(c.startswith('grid') for c in cls): y+=draw_grid(sl,n,128,y,1664,dark)+44
        elif n.name=='table': hh=table_h(n,1664); draw_table(sl,n,128,y,1664,hh,dark); y+=hh+44
        elif n.name=='p': y+=draw_para(sl,n,128,y,1664,dark)+44
        elif n.name=='div' and 'card' in cls: hh=card_h(n,1664); draw_card(sl,n,128,y,1664,hh,dark); y+=hh+44
    footer(sl,dark,page)
prs.save(OUT); print('slides',len(prs.slides))
