"""Raster-only color contour tracing to native editable VSDX geometry.

This is appearance tracing, not a chemical/flowchart semantic recognizer.
Optional reviewed OCR labels become text; unreviewed glyphs remain outlines.
No source VSDX, SVG, or previous native scene is read by this pipeline.
"""
import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import time
import xml.etree.ElementTree as E
import cv2
import numpy as np
from PIL import Image,ImageFont
import vsdx_scene as native


V=native.V


def tag(name):return '{'+V+'}'+name


def put(parent,name,value,formula=None):
    attrs={'N':name,'V':str(value)}
    if formula is not None:attrs['F']=formula
    return E.SubElement(parent,tag('Cell'),attrs)


def shape(parent,index,x,y,w,h,color):
    s=E.SubElement(parent,tag('Shape'),ID=str(index),NameU=f'trace_{index}',Type='Shape',LineStyle='0',FillStyle='0',TextStyle='0')
    for name,value in [('PinX',x+w/2),('PinY',y+h/2),('Width',w),('Height',h),('LocPinX',w/2),('LocPinY',h/2),
                       ('FillPattern',1),('LinePattern',0),('FillForegnd',color),('ShdwPattern',0),
                       ('LeftMargin',0),('RightMargin',0),('TopMargin',0),('BottomMargin',0)]:put(s,name,value)
    return s


def trace(source,colors=32,min_area=2,epsilon=0.22,labels=None):
    started=time.perf_counter();image=Image.open(source).convert('RGB');w,h=image.size;scale=1/96
    working=np.array(image);text_items=[]
    if labels:
        for line in labels:
            x,y,bw,bh=line['x'],line['y'],line['width'],line['height']
            # Reviewed labels in this mode must be on plain white backgrounds.
            working[max(0,int(y)-2):min(h,int(y+bh)+3),max(0,int(x)-2):min(w,int(x+bw)+3)]=255
            text_items.append(line)
    indexed=Image.fromarray(working).quantize(colors=colors,method=Image.Quantize.FASTOCTREE)
    upsample=3
    pixels=cv2.resize(np.array(indexed),None,fx=upsample,fy=upsample,interpolation=cv2.INTER_NEAREST)
    palette=np.array(indexed.getpalette(),dtype=np.uint8).reshape(-1,3)
    content=E.Element(tag('PageContents'),{'{http://www.w3.org/XML/1998/namespace}space':'preserve'})
    shapes=E.SubElement(content,tag('Shapes'));index=1;vertices=0;components=0
    background=shape(shapes,index,0,0,w*scale,h*scale,'#FFFFFF')
    geometry=E.SubElement(background,tag('Section'),N='Geometry',IX='0')
    for i,(x,y) in enumerate([(0,0),(w*scale,0),(w*scale,h*scale),(0,h*scale),(0,0)],1):
        row=E.SubElement(geometry,tag('Row'),T='MoveTo' if i==1 else 'LineTo',IX=str(i));put(row,'X',x);put(row,'Y',y)
    # Light colors first; nested holes keep their winding. Every retained color
    # component is its own native shape, not an embedded SVG or bitmap.
    for color_index in sorted(np.unique(pixels),key=lambda i:-float(palette[i].mean())):
        color=palette[color_index]
        if np.min(color)>248:continue
        mask=np.uint8(pixels==color_index)*255
        # OpenCV uses pixel centers for boundaries. A subpixel expansion closes
        # the gaps between adjacent palette regions when Visio rasterizes again.
        mask=cv2.dilate(mask,np.ones((3,3),np.uint8))
        contours,hierarchy=cv2.findContours(mask,cv2.RETR_CCOMP,cv2.CHAIN_APPROX_SIMPLE)
        if hierarchy is None:continue
        hierarchy=hierarchy[0]
        for ci,outer in enumerate(contours):
            if hierarchy[ci][3]>=0 or abs(cv2.contourArea(outer))<min_area*upsample**2:continue
            contour_set=[outer];child=hierarchy[ci][2]
            while child>=0:
                contour_set.append(contours[child]);child=hierarchy[child][0]
            bx,by,bw,bh=cv2.boundingRect(outer);hex_color='#'+''.join(f'{int(c):02X}' for c in color)
            index+=1;components+=1;s=shape(shapes,index,bx*scale/upsample,(h-(by+bh)/upsample)*scale,bw*scale/upsample,bh*scale/upsample,hex_color)
            geometry=E.SubElement(s,tag('Section'),N='Geometry',IX='0');put(geometry,'NoLine',1);put(geometry,'NoFill',0)
            row=0
            anchor=None
            for contour in contour_set:
                points=cv2.approxPolyDP(contour,epsilon*upsample,True).reshape(-1,2)
                if len(points)<3:continue
                if anchor is None:anchor=points[0]
                for j,(px,py) in enumerate(list(points)+[points[0]]):
                    row+=1;vertices+=1;r=E.SubElement(geometry,tag('Row'),T='MoveTo' if row==1 else 'LineTo',IX=str(row))
                    put(r,'X',(float(px)-bx)*scale/upsample);put(r,'Y',(by+bh-float(py))*scale/upsample)
                # Visio treats interior MoveTo gaps as filled straight segments.
                # Return explicitly over the same bridge to preserve hole winding.
                if contour is not outer:
                    row+=1;r=E.SubElement(geometry,tag('Row'),T='LineTo',IX=str(row))
                    put(r,'X',(float(anchor[0])-bx)*scale/upsample);put(r,'Y',(by+bh-float(anchor[1]))*scale/upsample)
    for label in text_items:
        x,y,bw,bh=(label[k] for k in ('x','y','width','height'));index+=1
        s=shape(shapes,index,(x-2)*scale,(h-y-bh*1.3)*scale,(bw+4)*scale,bh*1.6*scale,'#FFFFFF')
        s.set('NameU',f'label_{index}');put(s,'FillPattern',0)
        chars=E.SubElement(s,tag('Section'),N='Character');r=E.SubElement(chars,tag('Row'),IX='0')
        put(r,'Font','Times New Roman','FONT("Times New Roman")');put(r,'Color','#000000');put(r,'Style',1)
        font=ImageFont.truetype('timesbd.ttf',size=1000)
        font_pixels=min(bh*1.4,bw/max(1,font.getlength(label['text']))*1000)
        put(r,'Size',font_pixels*scale)
        para=E.SubElement(s,tag('Section'),N='Paragraph');pr=E.SubElement(para,tag('Row'),IX='0');put(pr,'HorzAlign',1)
        E.SubElement(s,tag('Text')).text=label['text']
    sheet=E.Element(tag('PageSheet'));put(sheet,'PageWidth',w*scale);put(sheet,'PageHeight',h*scale)
    support=[f'<DocumentSettings xmlns="{V}" TopPage="0" DefaultTextStyle="0" DefaultLineStyle="0" DefaultFillStyle="0"/>',
             f'<FaceNames xmlns="{V}"><FaceName ID="0" NameU="Times New Roman"/></FaceNames>',
             f'<StyleSheets xmlns="{V}"><StyleSheet ID="0" NameU="No Style" IsCustomNameU="1"/></StyleSheets>']
    model={'format':'visio-native-scene-v1','source_sha256':hashlib.sha256(Path(source).read_bytes()).hexdigest(),
           'support':support,'pages':[{'name':'Raster reconstruction','sheet':E.tostring(sheet,encoding='unicode'),
                                     'content':E.tostring(content,encoding='unicode')}]}
    return model,{'raster_width':w,'raster_height':h,'palette':colors,'min_area':min_area,'epsilon':epsilon,
                  'vector_components':components,'editable_text_labels':len(text_items),'vertices':vertices,
                  'trace_seconds':time.perf_counter()-started,'source_kind':'raster only'}


def reviewed_labels(ocr,allowed):
    result=[]
    for line in ocr['lines']:
        if line['text'] not in allowed:continue
        words=line['words']
        if not words:continue
        x=min(w['x'] for w in words);y=min(w['y'] for w in words)
        right=max(w['x']+w['width'] for w in words);bottom=max(w['y']+w['height'] for w in words)
        result.append(dict(text=line['text'],x=x,y=y,width=right-x,height=bottom-y))
    return result


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--image',type=Path,required=True);parser.add_argument('--output',type=Path,required=True)
    parser.add_argument('--ocr',type=Path);parser.add_argument('--reviewed-text',type=Path)
    parser.add_argument('--ocr-language',default='en-US',help='Run local Windows OCR when reviewed text is supplied without --ocr')
    parser.add_argument('--colors',type=int,default=32);parser.add_argument('--min-area',type=float,default=2)
    parser.add_argument('--epsilon',type=float,default=0.22)
    args=parser.parse_args();output=args.output.resolve();source=args.image.resolve()
    preview=output.with_suffix('.png')
    for destination in (output,preview):
        if destination.exists():raise FileExistsError(destination)
    if not 2<=args.colors<=256 or args.min_area<0 or args.epsilon<0:raise ValueError('Invalid trace settings')
    started=time.perf_counter();labels=[];ocr_seconds=0
    if args.reviewed_text:
        if args.ocr:
            ocr=json.loads(args.ocr.read_text(encoding='utf-8-sig'))
        else:
            t=time.perf_counter()
            run=subprocess.run(['powershell','-NoProfile','-ExecutionPolicy','Bypass','-File',str(Path(__file__).with_name('ocr_windows.ps1')),
                                '-ImagePath',str(source),'-Language',args.ocr_language],capture_output=True,encoding='utf-8',timeout=30)
            if run.returncode:raise RuntimeError(run.stderr)
            ocr=json.loads(run.stdout);ocr_seconds=time.perf_counter()-t
        labels=reviewed_labels(ocr,json.loads(args.reviewed_text.read_text(encoding='utf-8-sig')))
    cv2.setNumThreads(1)
    model,stats=trace(source,args.colors,args.min_area,args.epsilon,labels)
    t=time.perf_counter();native.reconstruct(model,output,readback=False);stats['package_seconds']=time.perf_counter()-t
    # Request files and diagnostics stay out of the delivery directory.
    with tempfile.TemporaryDirectory(prefix='visio-raster-') as work:
        request=Path(work)/'request.json'
        request.write_text(json.dumps(dict(action='export',document=str(output),
                detail='summary',png_dpi=96,exports=[str(preview)])),encoding='utf-8')
        t=time.perf_counter()
        r=subprocess.run([sys.executable,str(Path(__file__).with_name('visio_bridge.py')),
               '--request',str(request)],capture_output=True,encoding='utf-8',timeout=120)
    data=json.loads(r.stdout)
    if not data['ok']:raise RuntimeError(data)
    count=stats['vector_components']+stats['editable_text_labels']+1
    if data['result']['state']['shape_count']!=count:raise RuntimeError('Visio readback shape count mismatch')
    stats['visio_verify_export_seconds']=time.perf_counter()-t;stats['ocr_seconds']=ocr_seconds;stats['total_seconds']=time.perf_counter()-started
    print(json.dumps(stats,ensure_ascii=False))


if __name__=='__main__':main()
