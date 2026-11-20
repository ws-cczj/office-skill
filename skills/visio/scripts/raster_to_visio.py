"""Raster-only color contour tracing to native editable VSDX geometry.

This is appearance tracing, not a chemical/flowchart semantic recognizer.
Optional reviewed OCR labels become text; unreviewed glyphs remain outlines.
No source VSDX, SVG, or previous native scene is read by this pipeline.
"""
import argparse
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import time
import cv2
import numpy as np
from PIL import Image,ImageFont


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
    width, height = w * scale, h * scale
    operations = [dict(op='add', key='Trace_Canvas', x=width/2, y=height/2,
                       width=width, height=height, fill='#FFFFFF', no_line=True)]
    vertices = components = 0
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
            rings = [cv2.approxPolyDP(contour, epsilon * upsample, True).reshape(-1, 2)
                     for contour in contour_set]
            if len(rings[0]) < 3:
                continue
            points = []
            anchor = rings[0][0]
            for ring in rings:
                if len(ring) < 3:
                    continue
                points.extend(ring)
                points.append(ring[0])
                # Retrace the same bridge to preserve winding in a filled Visio shape.
                if ring is not rings[0]:
                    points.append(anchor)
            coordinates = [value for px, py in points for value in
                           (float(px) * scale / upsample, (h - float(py) / upsample) * scale)]
            hex_color = '#' + ''.join(f'{int(channel):02X}' for channel in color)
            components += 1
            vertices += len(points)
            operations.append(dict(op='add', kind='polyline', key=f'trace_{components}',
                                   points=coordinates, fill=hex_color, no_line=True))
    if text_items:
        font = ImageFont.truetype('timesbd.ttf', size=1000)
        for index, label in enumerate(text_items, 1):
            x, y, bw, bh = (label[k] for k in ('x', 'y', 'width', 'height'))
            font_pixels = min(bh * 1.4, bw / max(1, font.getlength(label['text'])) * 1000)
            operations.append(dict(
                op='add', kind='text', key=f'label_{index}', text=label['text'],
                x=(x + bw/2) * scale, y=(h - y - bh/2) * scale,
                width=(bw + 4) * scale, height=bh * 1.6 * scale,
                font='Times New Roman', font_size=font_pixels * scale * 72,
                text_color='#000000', bold=True, align=1, text_margin=0,
            ))
    request = dict(action='render', width=width, height=height, page='Raster reconstruction',
                   detail='summary', operations=operations)
    return request, {'raster_width':w,'raster_height':h,'palette':colors,'min_area':min_area,'epsilon':epsilon,
                     'vector_components':components,'editable_text_labels':len(text_items),'vertices':vertices,
                     'trace_seconds':time.perf_counter()-started,'source_kind':'raster only'}


def run_bridge(request, path):
    path.write_text(json.dumps(request, ensure_ascii=False), encoding='utf-8')
    result = subprocess.run(
        [sys.executable, str(Path(__file__).with_name('visio_bridge.py')), '--request', str(path)],
        capture_output=True, encoding='utf-8', timeout=120,
    )
    if result.returncode:
        raise RuntimeError(result.stdout.strip() or result.stderr.strip())
    return json.loads(result.stdout)['result']


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
    request, stats = trace(source, args.colors, args.min_area, args.epsilon, labels)
    # Keep bridge requests separate from the delivered drawing and preview.
    with tempfile.TemporaryDirectory(prefix='visio-raster-') as work:
        request_path = Path(work) / 'request.json'
        t = time.perf_counter()
        run_bridge(dict(request, document=str(output)), request_path)
        stats['visio_render_seconds'] = time.perf_counter() - t
        t = time.perf_counter()
        reopened = run_bridge(dict(action='export', document=str(output), page=request['page'],
                                   detail='summary', png_dpi=96, exports=[str(preview)]), request_path)
    if reopened['state']['shape_count'] != len(request['operations']):
        raise RuntimeError('Visio readback shape count mismatch')
    stats['visio_verify_export_seconds'] = time.perf_counter() - t
    stats['ocr_seconds'] = ocr_seconds
    stats['total_seconds'] = time.perf_counter() - started
    print(json.dumps(stats,ensure_ascii=False))


if __name__=='__main__':main()
