import sys, csv, io, os, re, urllib.request
import docx
from docx.table import _Cell
from docx.shared import Pt, Cm, RGBColor

MARK='✓'
SOURCE_RE=re.compile(r'^\d{3,4}\s*-?\s*DIT\s*-?\s*\d{3,5}$', re.I)

def clean(v): return (v or '').strip()
def norm(v): return re.sub(r'\s+',' ',clean(v)).strip().rstrip(':').casefold()

def get_row_cells(table,row_idx):
    cells=[]; seen=set()
    for tc in table.rows[row_idx]._tr.tc_lst:
        if id(tc) not in seen: seen.add(id(tc)); cells.append(_Cell(tc,table))
    return cells

def find_cell_index(cells,text,exact=True):
    target=norm(text)
    for i,c in enumerate(cells):
        t=norm(c.text)
        if (t==target) if exact else (t.startswith(target)):
            return i
    return None

def append_text(cell,value,bold=False):
    if not clean(value): return
    r=cell.paragraphs[-1].add_run(' '+str(value)); r.bold=bold

def mark_option(table,row_idx,option_text,mark=MARK,color=None):
    cells=get_row_cells(table,row_idx); idx=find_cell_index(cells,option_text,False)
    if idx is None or idx+1>=len(cells):
        print(f"  [aviso] No se encontró opción '{option_text}' fila {row_idx}"); return False
    r=cells[idx+1].paragraphs[0].add_run(mark); r.bold=True; r.font.size=Pt(12)
    if color: r.font.color.rgb=RGBColor(*color)
    return True

def maintainx_mark(answer):
    a=norm(answer)
    if a in ('yes','sí','si','true','1','checked'): return (MARK,None)
    if a in ('no','false','0','unchecked'): return ('X',(255,0,0))
    if a in ('n/a','na','n.a.','no aplica','not applicable'): return ('N',None)
    return None

def fill_label_field(table,row_idx,label,value,exact=False):
    cells=get_row_cells(table,row_idx); idx=find_cell_index(cells,label,exact)
    if idx is None: print(f"  [aviso] No se encontró etiqueta '{label}' fila {row_idx}"); return False
    append_text(cells[idx],value); return True

def insert_signature_image(table,row_idx,label,image_bytes):
    cells=get_row_cells(table,row_idx); idx=find_cell_index(cells,label,False)
    if idx is None: return False
    cells[idx].add_paragraph().add_run().add_picture(io.BytesIO(image_bytes),width=Cm(2.8)); return True

def download_image(url):
    try:
        req=urllib.request.Request(url,headers={'User-Agent':'Mozilla/5.0'})
        with urllib.request.urlopen(req,timeout=10) as resp: return resp.read()
    except Exception as e:
        print(f'  [aviso] Firma no descargada: {e}'); return None

ROW=dict(header=0,nombre=2,global_id=2,unidad_operativa=3,firma=3,
         accesos_opt=9,accesos_eval=10,estab_opt=11,estab_eval=12,
         limpieza_opt=13,limpieza_eval=14,segfis_opt=15,segfis_eval=16,
         senal_opt=17,senal_eval=18,letreros_opt=19,letreros_eval=20,
         amort_opt=21,amort_eval=22,obtura_opt=23,obtura_eval=24,vigilancia=26)

SECTION_MAP={
 'accesos':('accesos_opt','accesos_eval'),
 'estabilidad mecánica':('estab_opt','estab_eval'),
 'limpieza':('limpieza_opt','limpieza_eval'),
 'seguridad física':('segfis_opt','segfis_eval'),
 'señalización':('senal_opt','senal_eval'),
 'letreros':('letreros_opt','letreros_eval'),
 'amortiguador de vibraciones':('amort_opt','amort_eval'),
 'prueba de obturador':('obtura_opt','obtura_eval'),
 'prueba obturador':('obtura_opt','obtura_eval'),
}
GOOD_BAD={
 'accesos_eval':('Apto','No Apto'), 'estab_eval':('Funcional','No Funcional'),
 'limpieza_eval':('Limpio','Sucio'), 'segfis_eval':('Funcional','No Funcional'),
 'senal_eval':('Funcional','Deteriorado'), 'letreros_eval':('Funcional','Deteriorado'),
 'amort_eval':('Funcional','Deteriorado'), 'obtura_eval':('Funcional','Deteriorado')}

def pairs(row):
    out=[]; i=1
    while f'Question {i}' in row:
        out.append((clean(row.get(f'Question {i}')),clean(row.get(f'Answer {i}')))); i+=1
    return out

def is_source(q): return bool(SOURCE_RE.match(re.sub(r'\s+','',q)))
def canonical_source(q): return re.sub(r'\s+','',q).upper()

def split_sources(row):
    ps=pairs(row); common={}; starts=[]
    for i,(q,a) in enumerate(ps):
        if is_source(q): starts.append(i)
        elif not starts and q: common[norm(q)]=a
    blocks=[]
    for n,s in enumerate(starts):
        e=starts[n+1] if n+1<len(starts) else len(ps)
        blocks.append((canonical_source(ps[s][0]),ps[s+1:e]))
    return common,blocks

def fill_one(template,row,common,source,block,out_path):
    doc=docx.Document(template); table=doc.tables[1]
    fill_label_field(table,0,'Número de Registro:',clean(row.get('ID')))
    fill_label_field(table,0,'Área/Ubicación:',clean(row.get('Location')))
    dt=clean(row.get('Last updated') or row.get('Created on'))
    if dt:
        p=dt.split(); fill_label_field(table,0,'Fecha:',p[0] if p else '')
        fill_label_field(table,0,'Hora:',p[1] if len(p)>1 else '')
    fill_label_field(table,2,'Nombre:',common.get('nombre del poe',''))
    fill_label_field(table,2,'Global ID:',common.get('global id del poe',''))
    fill_label_field(table,3,'Unidad Operativa:',common.get('unidad operativa del poe',''))
    firma=common.get('firma del poe','')
    url=next((x for x in firma.split(',') if x.startswith('http')),None)
    if url:
        img=download_image(url)
        if img: insert_signature_image(table,3,'Firma:',img)
        else: fill_label_field(table,3,'Firma:','(firma digital registrada en MaintainX)')
    # El tag de MaintainX identifica la fuente/equipo inspeccionado.
    fill_label_field(table,5,'Código Contenedor:',source)

    current=None; dose={}
    for q,a in block:
        nq=norm(q)
        if nq in SECTION_MAP and not a:
            current=SECTION_MAP[nq]; continue
        if nq=='vigilancia radiológica': current=None; continue
        if nq in ('tasa de dosis superficial','tasa de dosis a 1 metro','comentario'):
            dose[nq]=a; continue
        if current:
            optkey,evalkey=current
            if nq=='evaluación':
                if a:
                    good,bad=GOOD_BAD[evalkey]
                    mark_option(table,ROW[evalkey], good if a.upper()=='PASS' else bad if a.upper()=='FLAG' else a)
            else:
                mx=maintainx_mark(a)
                if mx:
                    symbol,color=mx
                    mark_option(table,ROW[optkey],q,symbol,color)
    fill_label_field(table,26,'Tasa de Dosis Superficial:',dose.get('tasa de dosis superficial',''))
    fill_label_field(table,26,'Tasa de Dosis a 1 metro:',dose.get('tasa de dosis a 1 metro',''))
    fill_label_field(table,26,'Comentario:',dose.get('comentario',''))
    doc.save(out_path); print('Generado:',out_path)

def main():
    if len(sys.argv)!=4:
        print('Uso: python generar_reg_prmp15.py <csv> <plantilla.docx> <carpeta_salida>'); sys.exit(1)
    csv_path,template,outdir=sys.argv[1:4]; os.makedirs(outdir,exist_ok=True)
    generated=[]; ignored=[]; errors=[]
    with open(csv_path,encoding='utf-8-sig',newline='') as f:
        for row in csv.DictReader(f):
            wid=clean(row.get('ID')) or 'sin_id'
            if clean(row.get('Status')).upper()!='DONE': continue
            common,blocks=split_sources(row)
            if 'nombre del poe' not in common or not blocks:
                ignored.append(wid); continue
            for source,block in blocks:
                safe=re.sub(r'[^A-Za-z0-9_-]+','-',source)
                path=os.path.join(outdir,f'REG-PRMP-15_{wid}_{safe}.docx')
                try: fill_one(template,row,common,source,block,path); generated.append(f'{wid}/{source}')
                except Exception as e: errors.append((wid,source,str(e))); print(f'[ERROR] {wid}/{source}: {e}')
    print('\n----- RESUMEN -----')
    print(f'Words generados: {len(generated)} -> {generated}')
    print(f'Órdenes ignoradas: {len(ignored)} -> {ignored}')
    print(f'Con error: {len(errors)} -> {errors}')

if __name__=='__main__': main()
