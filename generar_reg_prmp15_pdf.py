import sys, os, re, io
import fitz
import docx
from docx.table import _Cell
from docx.shared import Pt, Cm, RGBColor

MARK='✓'
SOURCE_RE=re.compile(r'^\d{3,4}-DIT-\d{3,5}$',re.I)

def clean(s): return (s or '').strip()
def norm(s): return re.sub(r'[\s:]+$','',re.sub(r'\s+',' ',clean(s))).casefold()
def get_row_cells(table,row_idx):
    cells=[]; seen=set()
    for tc in table.rows[row_idx]._tr.tc_lst:
        if id(tc) not in seen: seen.add(id(tc)); cells.append(_Cell(tc,table))
    return cells
def find_cell_index(cells,text,exact=True):
    target=norm(text)
    for i,c in enumerate(cells):
        t=norm(c.text)
        if (t==target) if exact else t.startswith(target): return i
    return None
def append_text(cell,value,bold=False):
    if clean(value):
        r=cell.paragraphs[-1].add_run(' '+str(value)); r.bold=bold
def fill_label_field(table,row,label,value,exact=False):
    cells=get_row_cells(table,row); idx=find_cell_index(cells,label,exact)
    if idx is None: return False
    append_text(cells[idx],value); return True
def mark_option(table,row,option,mark=MARK,color=None):
    cells=get_row_cells(table,row); idx=find_cell_index(cells,option,False)
    if idx is None or idx+1>=len(cells): return False
    r=cells[idx+1].paragraphs[0].add_run(mark); r.bold=True; r.font.size=Pt(12)
    if color: r.font.color.rgb=RGBColor(*color)
    return True

def selected_symbol(ans):
    if ans=='Sí': return MARK,None
    if ans=='No': return 'X',(255,0,0)
    return 'N',None

ROW=dict(accesos_opt=9,accesos_eval=10,estab_opt=11,estab_eval=12,limpieza_opt=13,limpieza_eval=14,
         segfis_opt=15,segfis_eval=16,senal_opt=17,senal_eval=18,amort_opt=21,amort_eval=22,
         obtura_opt=23,obtura_eval=24)
SECTION_MAP={'accesos':'accesos','estabilidad mecánica':'estab','limpieza':'limpieza','seguridad física':'segfis',
             'señalización':'senal','amortiguador de vibraciones':'amort','prueba de obturador':'obtura'}
GOOD_BAD={'accesos':('Apto','No Apto'),'estab':('Funcional','No Funcional'),'limpieza':('Limpio','Sucio'),
          'segfis':('Funcional','No Funcional'),'senal':('Funcional','Deteriorado'),'amort':('Funcional','Deteriorado'),
          'obtura':('Funcional','Deteriorado')}

def lines(page):
    out=[]
    for b in page.get_text('dict')['blocks']:
        if 'lines' not in b: continue
        for l in b['lines']:
            txt=''.join(s['text'] for s in l['spans']).strip()
            if txt: out.append((l['bbox'][1],l['bbox'][3],l['bbox'][0],txt))
    return sorted(out)

def blue_selected(page, y0,y1):
    clip=fitz.Rect(25,y0-2,47,y1+2); pix=page.get_pixmap(matrix=fitz.Matrix(2,2),clip=clip,alpha=False)
    import numpy as np
    a=np.frombuffer(pix.samples,np.uint8).reshape(pix.height,pix.width,pix.n)
    return int(((a[:,:,2]>150)&(a[:,:,0]<120)&(a[:,:,1]>70)).sum())>20

def green_selected(page,y0,y1):
    # Evaluación: el icono seleccionado es de color (verde=Aprobado, naranja=Alerta, rojo=Falla); los no seleccionados no tienen icono.
    clip=fitz.Rect(25,y0-2,47,y1+2); pix=page.get_pixmap(matrix=fitz.Matrix(2,2),clip=clip,alpha=False)
    import numpy as np
    a=np.frombuffer(pix.samples,np.uint8).reshape(pix.height,pix.width,pix.n)[:,:,:3].astype(int)
    return int(((a.max(axis=2)-a.min(axis=2))>60).sum())>8

def value_after_label(ls,label):
    n=norm(label)
    for i,(_,_,_,t) in enumerate(ls):
        if norm(t)==n:
            vals=[]
            for j in range(i+1,min(i+8,len(ls))):
                tt=ls[j][3]
                if tt.startswith('Rellenado por') or tt.startswith('Firmado por'): break
                vals.append(tt)
            return '\n'.join(vals)
    return ''

def parse_pdf(path):
    pdf=fitz.open(path); all_lines=[lines(p) for p in pdf]
    full='\n'.join(p.get_text() for p in pdf)
    m=re.search(r'#(\d+)',full); wid=m.group(1) if m else 'sin_id'
    common={'nombre del poe':'','global id del poe':'','unidad operativa del poe':''}
    # Common values from first pages. Names may be multiple lines.
    for pi in range(min(3,len(pdf))):
        ls=all_lines[pi]
        for key,label in [('nombre del poe','Nombre del POE'),('global id del poe','Global ID del POE'),('unidad operativa del poe','Unidad Operativa del POE')]:
            if not common[key]:
                v=value_after_label(ls,label)
                # For radio fields, keep only selected option text/value appearing immediately after label if possible.
                common[key]=v
    # Clean global/unit: in this export the chosen numeric value is first line.
    for k in ('global id del poe','unidad operativa del poe'):
        if common[k]: common[k]=common[k].split('\n')[0]
    # signature crop from page containing Firma del POE
    signature=None
    for pi,ls in enumerate(all_lines[:3]):
        for i,(y0,y1,x,t) in enumerate(ls):
            if norm(t)=='firma del poe':
                end=next((yy0 for yy0,_,_,tt in ls[i+1:] if tt.startswith('Firmado por')),y1+80)
                clip=fitz.Rect(35,y1+4,180,end-2)
                signature=pdf[pi].get_pixmap(matrix=fitz.Matrix(2.5,2.5),clip=clip,alpha=False).tobytes('png')
                break
    sources=[]; current=None; section=None
    # page source context and photos
    page_source=[]
    for pi,(p,ls) in enumerate(zip(pdf,all_lines)):
        # source headings with y positions
        headings=[(y0,t.replace(' ','')) for y0,_,_,t in ls if SOURCE_RE.match(t.replace(' ',''))]
        for idx,(y0,y1,x,t) in enumerate(ls):
            tt=t.replace(' ','')
            if SOURCE_RE.match(tt):
                current={'source':tt.upper(),'items':[],'dose':{},'photos':[]}; sources.append(current); section=None; continue
            if current is None: continue
            nt=norm(t)
            if nt in SECTION_MAP: section=SECTION_MAP[nt]; continue
            if nt=='vigilancia radiológica': section=None; continue
            if nt in ('tasa de dosis superficial','tasa de dosis a 1 metro','comentario'):
                vals=[]
                for yy0,yy1,xx,tx in ls[idx+1:]:
                    if tx.startswith('Rellenado por'): break
                    vals.append(tx)
                current['dose'][nt]=' '.join(vals).strip(); continue
            # Question labels: look ahead for Sí/No/N/D within 70pt.
            if section and t.endswith(':') and nt not in ('evaluación','evaluación:'):
                opts=[]
                for yy0,yy1,xx,tx in ls[idx+1:idx+7]:
                    if yy0-y1>75: break
                    if tx in ('Sí','No','N/D'): opts.append((yy0,yy1,tx))
                if len(opts)>=3:
                    ans=next((tx for yy0,yy1,tx in opts if blue_selected(p,yy0,yy1)),None)
                    if ans: current['items'].append((section,t.rstrip(' :'),ans))
            if section and nt.startswith('evaluación'):
                choices=[]
                for yy0,yy1,xx,tx in ls[idx+1:idx+7]:
                    if tx in ('Aprobado','Alerta','Falla'): choices.append((yy0,yy1,tx))
                sel=next((tx for yy0,yy1,tx in choices if green_selected(p,yy0,yy1)),None)
                if sel: current['items'].append((section,'__eval__',sel))
        # embedded photos > 250px each side; assign based on source current at image vertical position
        for info in p.get_images(full=True):
            xref=info[0]
            try:
                rects=p.get_image_rects(xref)
                data=pdf.extract_image(xref); w,h=data.get('width',0),data.get('height',0)
                if w<250 or h<250: continue
                for r in rects:
                    # Ignore header/company art by requiring an active source above image.
                    src=None
                    for sy,st in headings:
                        if sy<r.y0: src=next((s for s in reversed(sources) if s['source']==st.upper()),None)
                    if src is None: src=current
                    if src is not None:
                        img=data['image']
                        if img not in src['photos']: src['photos'].append(img)
            except Exception: pass
    pdf.close(); return wid,common,signature,sources

def fill_one(template,wid,common,signature,src,out_path):
    doc=docx.Document(template); table=doc.tables[1]
    fill_label_field(table,0,'Número de Registro:',wid)
    fill_label_field(table,2,'Nombre:',common.get('nombre del poe',''))
    fill_label_field(table,2,'Global ID:',common.get('global id del poe',''))
    fill_label_field(table,3,'Unidad Operativa:',common.get('unidad operativa del poe',''))
    if signature:
        cells=get_row_cells(table,3); idx=find_cell_index(cells,'Firma:',False)
        if idx is not None: cells[idx].add_paragraph().add_run().add_picture(io.BytesIO(signature),width=Cm(2.8))
    fill_label_field(table,5,'Código Contenedor:',src['source'])
    for sec,q,a in src['items']:
        if q=='__eval__':
            good,bad=GOOD_BAD[sec]; opt=good if a=='Aprobado' else bad
            mark_option(table,ROW[sec+'_eval'],opt)
        else:
            sym,col=selected_symbol(a); mark_option(table,ROW[sec+'_opt'],q,sym,col)
    dose=src['dose']
    fill_label_field(table,26,'Tasa de Dosis Superficial:',dose.get('tasa de dosis superficial',''))
    fill_label_field(table,26,'Tasa de Dosis a 1 metro:',dose.get('tasa de dosis a 1 metro',''))
    fill_label_field(table,26,'Comentario:',dose.get('comentario',''))
    if src['photos']:
        doc.add_heading('Evidencias fotográficas',level=2)
        for i,img in enumerate(src['photos'],1):
            p=doc.add_paragraph(); p.alignment=1
            p.add_run().add_picture(io.BytesIO(img),width=Cm(8.5))
            c=doc.add_paragraph(f'Evidencia {i} - {src["source"]}'); c.alignment=1
    doc.save(out_path)

def main():
    if len(sys.argv)!=4:
        print('Uso: python generar_reg_prmp15_pdf.py <pdf> <plantilla.docx> <carpeta_salida>'); sys.exit(1)
    pdf,template,outdir=sys.argv[1:4]; os.makedirs(outdir,exist_ok=True)
    wid,common,signature,sources=parse_pdf(pdf)
    print('Orden:',wid,'Fuentes:',[s['source'] for s in sources])
    for s in sources:
        out=os.path.join(outdir,f'REG-PRMP-15_{wid}_{s["source"]}.docx')
        fill_one(template,wid,common,signature,s,out); print('Generado:',out,'fotos:',len(s['photos']))
if __name__=='__main__': main()
