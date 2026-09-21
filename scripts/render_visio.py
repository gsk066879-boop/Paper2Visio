"""Render article-specific layout JSON as native editable Visio shapes.

Windows: python render_visio.py layout.json --output-dir outputs --name diagram
Use --check-only for layout validation without COM or file writes.
"""
import argparse
import json
import math
import pathlib
import shutil
import sys
from PIL import Image, ImageFont

DEFAULT_FONT = 'Microsoft YaHei UI'

def rgb(value):
    value = value.lstrip('#')
    if len(value) != 6:
        raise ValueError('Expected #RRGGBB color')
    return 'RGB(%d,%d,%d)' % tuple(int(value[i:i+2],16) for i in (0,2,4))

def fontfile(bold=False):
    import os
    root = pathlib.Path(os.environ.get('WINDIR', r'C:\Windows')) / 'Fonts'
    path = root / ('msyhbd.ttc' if bold else 'msyh.ttc')
    if not path.exists():
        raise RuntimeError('Microsoft YaHei font file unavailable; install/select and measure a matching Chinese font before rendering')
    return path

def validate(spec):
    errors=[]
    width,height=spec.get('width',0),spec.get('height',0)
    if width<=0 or height<=0 or spec.get('units_per_inch',80)<=0:
        errors.append('Positive width, height and units_per_inch required')
    ids={}
    for e in spec.get('elements',[]):
        ident=e.get('id');kind=e.get('kind')
        if not ident or ident in ids:errors.append(f'Duplicate/missing id: {ident}')
        ids[ident]=e
        if kind not in {'rect','text','poly','line','connector'}:
            errors.append(f'{ident}: unsupported kind {kind}');continue
        if kind in {'rect','text'}:
            if any(n not in e for n in ('x','y','w','h')):
                errors.append(f'{ident}: missing rectangle geometry');continue
            if min(e['w'],e['h'])<=0 or e['x']<0 or e['y']<0 or e['x']+e['w']>width+.01 or e['y']+e['h']>height+.01:
                errors.append(f'{ident}: outside canvas or nonpositive size')
        if kind in {'line','poly'}:
            pts=e.get('points',[])
            if len(pts)<2:errors.append(f'{ident}: requires at least two points')
            if any(not(0<=x<=width and 0<=y<=height) for x,y in pts):errors.append(f'{ident}: points outside canvas')
        if kind=='text':
            text=e.get('text','');size=e.get('size',14)
            if size<spec.get('min_font_size',10):errors.append(f'{ident}: below minimum font size')
            font=ImageFont.truetype(str(fontfile(e.get('bold',False))),100)
            lines=text.split('\n')
            if max((font.getlength(t) for t in lines),default=0)*size/100>e['w']-.5:
                errors.append(f'{ident}: text wider than box; wrap or widen')
            if len(lines)*size*1.25>e['h']+.5:
                errors.append(f'{ident}: text taller than box; increase height')
            for run in e.get('runs',[]):
                if not(0<=run['start']<run['end']<=len(text)):errors.append(f'{ident}: invalid text run')
    for e in spec.get('elements',[]):
        if e.get('kind')=='connector':
            for end in ['from','to']:
                if e.get(end) not in ids or ids[e[end]].get('kind')!='rect':errors.append(f'{e.get("id")}: {end} must reference a rect')
                port=e.get(end+'_port',[.5,0] if end=='from' else [.5,1])
                if len(port)!=2 or any(not 0<=p<=1 for p in port):errors.append(f'{e.get("id")}: invalid port')
    if not spec.get('elements'):errors.append('Empty layout')
    if errors:raise ValueError('\n'.join(errors))
    return ids

def cell(shape,name,value):shape.CellsU(name).FormulaU=str(value)

def render(spec,out,name,reference=None,overwrite=False,test_connections=False):
    import pythoncom
    import win32com.client
    validate(spec)
    out=pathlib.Path(out).resolve();out.mkdir(parents=True,exist_ok=True)
    paths={k:out/(name+s) for k,s in {'vsdx':'.vsdx','preview':'-preview.png','layout':'-layout.json','report':'-validation.json','script':'-render.py','replay':'-replay.ps1','reference':'-reference.png'}.items()}
    for k,p in paths.items():
        if p.exists() and not overwrite:raise FileExistsError(f'Refusing to overwrite {p}; choose another --name')
    app=win32com.client.Dispatch('Visio.Application');app.Visible=True
    old_update=app.ScreenUpdating
    doc=None
    try:
        doc=app.Documents.Add('');page=doc.Pages.Item(1);page.Name='原生可编辑图'
        width,height,scale=spec['width'],spec['height'],spec.get('units_per_inch',80)
        for key,value in [('PageWidth',width/scale),('PageHeight',height/scale)]:cell(page.PageSheet,key,value)
        # Choose a font registered in this actual Visio document.
        available={f.Name.casefold():f.ID for f in doc.Fonts}
        requested=spec.get('font',DEFAULT_FONT)
        candidates=[requested,DEFAULT_FONT,'Microsoft YaHei','微软雅黑','微软雅黑 UI']
        selected=next((f for f in candidates if f.casefold() in available),None)
        if not selected:raise RuntimeError('No registered Microsoft YaHei font in Visio; refusing silent font fallback')
        font_id=available[selected.casefold()]
        shapes={};connectors=[]
        app.ScreenUpdating=False
        background=page.DrawRectangle(0,0,width/scale,height/scale)
        background.NameU='CanvasBackground'
        cell(background,'FillForegnd',rgb(spec.get('background','#FAF8E8')));cell(background,'LinePattern',0);cell(background,'ShdwPattern',0)
        def make_points(points):return [v for x,y in points for v in (x/scale,(height-y)/scale)]
        for index,e in enumerate(spec['elements']):
            kind=e['kind']
            if kind=='connector':continue
            if kind in ('rect','text'):
                x,y,w,h=(e[k] for k in ('x','y','w','h'))
                sh=page.DrawRectangle(x/scale,(height-y-h)/scale,(x+w)/scale,(height-y)/scale)
            else:sh=page.DrawPolyline(make_points(e['points']),0)
            sh.NameU=e['id'];shapes[e['id']]=sh
            cell(sh,'ShdwPattern',0)
            if kind=='text':
                cell(sh,'FillPattern',0);cell(sh,'LinePattern',0);sh.Text=e.get('text','')
                for key in ['Char.Font','Char.AsianFont','Char.ComplexScriptFont']:
                    if sh.CellExistsU(key,0):cell(sh,key,font_id)
                cell(sh,'Char.Size',f'{e.get("size",14)*72/scale} pt')
                cell(sh,'Char.Color',rgb(e.get('color','#252525')))
                cell(sh,'Char.Style',1 if e.get('bold') else 0)
                cell(sh,'Para.HorzAlign',e.get('align',0));cell(sh,'VerticalAlign',1)
                for key in ['LeftMargin','RightMargin','TopMargin','BottomMargin']:cell(sh,key,0)
                for run in e.get('runs',[]):
                    chars=sh.Characters
                    # COM character offsets count UTF-16 code units.
                    chars.Begin=len(sh.Text[:run['start']].encode('utf-16-le'))//2
                    chars.End=len(sh.Text[:run['end']].encode('utf-16-le'))//2
                    chars._oleobj_.Invoke(chars._oleobj_.GetIDsOfNames('CharProps'),0,pythoncom.DISPATCH_PROPERTYPUT,0,1,2)
                    row=chars.CharPropsRow(0)
                    sh.CellsSRC(3,row,1).FormulaU=rgb(run.get('color','#A3262C'))
            elif kind in ('rect','poly'):
                cell(sh,'FillPattern',1 if e.get('fill') else 0)
                if e.get('fill'):cell(sh,'FillForegnd',rgb(e['fill']))
                cell(sh,'LinePattern',1 if e.get('line') else 0)
                if e.get('line'):cell(sh,'LineColor',rgb(e['line']))
                cell(sh,'LineWeight',f'{e.get("weight",.6)} pt')
                if kind=='rect':cell(sh,'Rounding',e.get('r',2)/scale)
            else:
                cell(sh,'FillPattern',0);cell(sh,'LineColor',rgb(e.get('color','#155675')))
                cell(sh,'LineWeight',f'{e.get("weight",1)} pt');cell(sh,'Rounding',.04)
                cell(sh,'EndArrow',4 if e.get('arrow',False) else 0);cell(sh,'EndArrowSize',1)
            if index%70==0:print(f'Drawing {index}/{len(spec["elements"])}',flush=True)
        for e in spec['elements']:
            if e['kind']!='connector':continue
            sh=page.Drop(app.ConnectorToolDataObject,0,0);sh.NameU=e['id'];shapes[e['id']]=sh
            cell(sh,'ShapeRouteStyle',1);cell(sh,'Rounding',.05)
            cell(sh,'LineColor',rgb(e.get('color','#155675')));cell(sh,'LineWeight',f'{e.get("weight",1)} pt')
            cell(sh,'BeginArrow',0);cell(sh,'EndArrow',4 if e.get('arrow',True) else 0);cell(sh,'EndArrowSize',1)
            for field,end in [('BeginX','from'),('EndX','to')]:
                fx,fy=e.get(end+'_port',[.5,0] if end=='from' else [.5,1])
                sh.CellsU(field).GlueToPos(shapes[e[end]],fx,fy)
            connectors.append((e,sh))
        app.ScreenUpdating=True
        pythoncom.PumpWaitingMessages()
        move_results=[]
        if test_connections:
            for e,sh in connectors:
                node=shapes[e['from']];before=sh.CellsU('BeginX').ResultIU
                old_x=node.CellsU('PinX').FormulaU
                node.CellsU('PinX').ResultIU+=.125
                delta=sh.CellsU('BeginX').ResultIU-before
                node.CellsU('PinX').FormulaU=old_x
                ok=abs(delta-.125)<1e-5
                move_results.append({'id':e['id'],'follows_node':ok})
                if not ok:raise RuntimeError(f'Connector {e["id"]} did not follow moved node')
        main_count=page.Shapes.Count
        if reference:
            shutil.copyfile(reference,paths['reference'])
            ref=doc.Pages.Add();ref.Name='参考图对照（图片）'
            cell(ref.PageSheet,'PageWidth',width/scale);cell(ref.PageSheet,'PageHeight',height/scale)
            pic=ref.Import(str(paths['reference']))
            for key,val in [('Width',width/scale),('Height',height/scale),('PinX',width/scale/2),('PinY',height/scale/2)]:cell(pic,key,val)
        doc.SaveAs(str(paths['vsdx']));doc.Close();doc=None
        doc=app.Documents.Open(str(paths['vsdx']));page=doc.Pages.Item(1)
        app.ActiveWindow.Page=page
        try:app.ActiveWindow.ViewFit=1
        except Exception:pass
        app.ScreenUpdating=True
        page.Export(str(paths['preview']))
        actual_text={s.NameU:s.Text for s in page.Shapes if s.Type!=4}
        expected_text={e['id']:e['text'] for e in spec['elements'] if e['kind']=='text'}
        if any(actual_text.get(k)!=v for k,v in expected_text.items()):raise RuntimeError('Text changed after save/reopen')
        glue=[]
        for e,_ in connectors:
            sh=page.Shapes.ItemU(e['id'])
            records=[{'cell':c.FromCell.Name,'target':c.ToSheet.NameU} for c in sh.Connects]
            targets={r['target'] for r in records}
            if not {e['from'],e['to']}.issubset(targets):raise RuntimeError(f'Lost endpoint glue: {e["id"]}')
            glue.append({'id':e['id'],'connections':records})
        report={'pages':doc.Pages.Count,'native_shapes':main_count,'main_foreign_images':sum(s.Type==4 for s in page.Shapes),'text_shapes':len(expected_text),'connectors':glue,'move_tests':move_results,'reopened':True,'text_preserved':True,'font':selected,'visible':bool(app.Visible),'vsdx_bytes':paths['vsdx'].stat().st_size,'preview_size':list(Image.open(paths['preview']).size),'visual_review':'Required by agent; not established by script'}
        paths['layout'].write_text(json.dumps(spec,ensure_ascii=False,indent=2),encoding='utf-8')
        paths['report'].write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
        shutil.copyfile(__file__,paths['script'])
        quote=lambda s:"'"+s.replace("'","''")+"'"
        command='python (Join-Path $PSScriptRoot '+quote(paths['script'].name)+') (Join-Path $PSScriptRoot '+quote(paths['layout'].name)+') --output-dir $PSScriptRoot --name '+quote(name+'-replay')
        if reference:command+=' --reference (Join-Path $PSScriptRoot '+quote(paths['reference'].name)+')'
        paths['replay'].write_text("$ErrorActionPreference = 'Stop'\n"+command+"\nif ($LASTEXITCODE -ne 0) { throw 'Visio rendering failed' }\n",encoding='utf-8-sig')
        print(json.dumps({'output':str(paths['vsdx']),'shapes':main_count,'connectors':len(glue),'move_tests':move_results},ensure_ascii=True),flush=True)
        return report
    finally:
        app.ScreenUpdating=old_update

def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('layout',type=pathlib.Path);p.add_argument('--output-dir',type=pathlib.Path)
    p.add_argument('--name',default='paper-roadmap');p.add_argument('--reference',type=pathlib.Path)
    p.add_argument('--check-only',action='store_true');p.add_argument('--overwrite',action='store_true')
    p.add_argument('--test-connections',action='store_true')
    args=p.parse_args()
    if pathlib.Path(args.name).name!=args.name or any(c in args.name for c in '<>:"/\\|?*'):
        p.error('--name must be a filename stem')
    spec=json.loads(args.layout.read_text(encoding='utf-8-sig'));validate(spec)
    if args.check_only:print('Layout valid');return
    if not args.output_dir:p.error('--output-dir required unless --check-only')
    if args.reference and not args.reference.is_file():p.error('--reference file does not exist')
    render(spec,args.output_dir,args.name,args.reference,args.overwrite,args.test_connections)

if __name__=='__main__':main()
