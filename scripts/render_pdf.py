"""Render an existing PDF to a NEW directory for visual QA (pdf2image + Poppler)."""
import argparse
import json
from pathlib import Path
from pdf2image import convert_from_path, pdfinfo_from_path
from PIL import Image, ImageDraw


def main():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('pdf',type=Path);ap.add_argument('output',type=Path)
    ap.add_argument('--dpi',type=int,default=120)
    args=ap.parse_args()
    args.output.mkdir(parents=True,exist_ok=False)
    count=pdfinfo_from_path(str(args.pdf))['Pages'];paths=[];thumbs=[]
    for n in range(1,count+1):
        im=convert_from_path(str(args.pdf),dpi=args.dpi,first_page=n,last_page=n)[0]
        path=args.output/f'page-{n:03}.png';im.save(path);paths.append(str(path.resolve()))
        im.thumbnail((300,430));thumbs.append((n,im.copy()))
    for offset in range(0,len(thumbs),12):
        sheet=Image.new('RGB',(1200,1380),'#cccccc');draw=ImageDraw.Draw(sheet)
        for j,(n,im) in enumerate(thumbs[offset:offset+12]):
            x=(j%4)*300;y=(j//4)*460;sheet.paste(im,(x,y+25));draw.text((x+10,y+5),f'Page {n}',fill='black')
        sheet.save(args.output/f'overview-{offset//12+1}.png')
    (args.output/'pages.json').write_text(json.dumps(paths,indent=2),encoding='utf8')
    print(json.dumps({'pages':count,'directory':str(args.output.resolve())}))


if __name__=='__main__':main()
