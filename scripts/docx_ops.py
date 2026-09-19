"""Conservative Word OOXML tools. Requires lxml; never overwrite an output.

Commands: inspect, bibliography, figures, lead-sentence, normalize, audit-assets.
All mutations preserve package media/embeddings byte-for-byte and refuse revisions.
"""
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from collections import Counter
from copy import deepcopy
import argparse
import json
import re
from lxml import etree as E

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS = {'w': W[1:-1], 'o': 'urn:schemas-microsoft-com:office:office'}
SPACE = '{http://www.w3.org/XML/1998/namespace}space'
PORDER = 'pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange'.split()
CITE = re.compile(r'\[(\d+(?:\s*[-–,，]\s*\d+)*)\]')
BOUNDARY = re.compile(r'(?<=[\u3400-\u9fff，。；：！？、（）【】《》])[^\S\r\n]+(?=[A-Za-z0-9])|(?<=[A-Za-z0-9])[^\S\r\n]+(?=[\u3400-\u9fff，。；：！？、（）【】《》])')


def require(ok, message):
    if not ok:
        raise ValueError(message)


def read(path):
    with ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}


def text(p):
    return ''.join(p.xpath('.//w:t/text()', namespaces=NS))


def xml(root):
    return E.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)


def pprop(p, tag):
    pp = p.find(W+'pPr')
    if pp is None:
        pp = E.Element(W+'pPr'); p.insert(0, pp)
    node = pp.find(W+tag)
    if node is None:
        node = E.Element(W+tag)
        rank = PORDER.index(tag)
        at = next((i for i,c in enumerate(pp) if E.QName(c).localname in PORDER and PORDER.index(E.QName(c).localname)>rank), len(pp))
        pp.insert(at, node)
    return node


def run(s=None, rp=None):
    r = E.Element(W+'r')
    if rp is not None: r.append(deepcopy(rp))
    if s is not None:
        t = E.SubElement(r,W+'t'); t.set(SPACE,'preserve'); t.text=s
    return r


def field(code, result, rp=None):
    result_runs=[]
    for kind in ['begin','code','separate','result','end']:
        r=run(rp=rp)
        if kind in ['begin','separate','end']:
            E.SubElement(r,W+'fldChar').set(W+'fldCharType',kind)
        elif kind=='code':
            t=E.SubElement(r,W+'instrText'); t.set(SPACE,'preserve'); t.text=' '+code+' '
        else: E.SubElement(r,W+'t').text=result
        result_runs.append(r)
    return result_runs


def stream(p):
    """Visible editable run text; field results, math and hyperlinks are barriers."""
    out=[]; refs=[]; depth=0
    for r in p:
        if r.tag==W+'pPr': continue
        if r.tag!=W+'r':
            out.append('\ufffc');refs.append(None);continue
        fld=r.find(W+'fldChar')
        if fld is not None:
            kind=fld.get(W+'fldCharType')
            if kind=='begin': depth+=1
            elif kind=='end': depth=max(0,depth-1)
            out.append('\ufffc');refs.append(None);continue
        ts=r.findall(W+'t')
        if depth or len(ts)!=1 or any(c.tag not in {W+'rPr',W+'t',W+'lastRenderedPageBreak'} for c in r):
            out.append('\ufffc');refs.append(None);continue
        t=ts[0]
        s=t.text or '';out.extend(s);refs.extend((t,i) for i in range(len(s)))
    return ''.join(out),refs


def replace_match(p,a,b,make):
    """Run-preserving replacement. Call in reverse position order."""
    s,refs=stream(p); selected=refs[a:b]
    require(selected and all(x is not None for x in selected),'Match crosses protected content')
    first,offset=selected[0];r=first.getparent();rp=r.find(W+'rPr')
    groups={}
    for t,i in selected: groups.setdefault(t,[]).append(i)
    tail=(first.text or '')[max(groups[first])+1:]
    prefix=(first.text or '')[:offset]
    for t,indices in groups.items():
        t.text=''.join(c for i,c in enumerate(t.text or '') if i not in indices)
    first.text=prefix
    nodes=make(rp)
    if tail:nodes.append(run(tail,rp))
    at=p.index(r)+1
    for i,n in enumerate(nodes):p.insert(at+i,n)


def fonts(rp, cn='宋体', latin='Times New Roman'):
    rf=rp.find(W+'rFonts')
    if rf is None:
        rf=E.Element(W+'rFonts');rp.insert(1 if len(rp) and rp[0].tag==W+'rStyle' else 0,rf)
    for a in list(rf.attrib):
        if E.QName(a).localname.endswith('Theme') or E.QName(a).localname in ['hint','cstheme']:del rf.attrib[a]
    for a in ['ascii','hAnsi','cs']:rf.set(W+a,latin)
    rf.set(W+'eastAsia',cn)


def font_size(rp, points):
    if points is None:return
    require(points > 0 and points*2 == int(points*2),'Font size must be positive half-points')
    # Preserve baseline/position and other run properties; update size only.
    for tag in ['sz','szCs']:
        node=rp.find(W+tag)
        if node is None:
            node=E.Element(W+tag)
            later={'highlight','u','effect','bdr','shd','fitText','vertAlign','rtl','cs','em','lang','eastAsianLayout','specVanish','oMath','rPrChange'}
            if tag=='sz':later.add('szCs')
            at=next((i for i,c in enumerate(rp) if E.QName(c).localname in later),len(rp))
            rp.insert(at,node)
        node.set(W+'val',str(int(points*2)))


def field_inventory(root):
    fields=[];stack=[]
    for e in root.iter():
        if e.tag==W+'fldChar':
            kind=e.get(W+'fldCharType')
            if kind=='begin':stack.append({'code':'','result':'','show':False,'locked':e.get(W+'fldLock')})
            elif kind=='separate' and stack:stack[-1]['show']=True
            elif kind=='end' and stack:fields.append(stack.pop())
        elif e.tag==W+'instrText' and stack:stack[-1]['code']+=e.text or ''
        elif e.tag==W+'t' and stack and stack[-1]['show']:stack[-1]['result']+=e.text or ''
    for f in root.iter(W+'fldSimple'):
        fields.append({'code':f.get(W+'instr',''),'result':text(f),'simple':True,'locked':f.get(W+'fldLock')})
    return fields


def inventory(parts):
    d=E.fromstring(parts['word/document.xml'])
    return {'paragraphs':[{'index':i,'text':text(p),'numId':p.xpath('./w:pPr/w:numPr/w:numId/@w:val',namespaces=NS)} for i,p in enumerate(d.xpath('//w:p',namespaces=NS))],
            'fields':field_inventory(d),'MathType':len(d.xpath('//o:OLEObject[@ProgID="Equation.DSMT4"]',namespaces=NS)),
            'revisions':len(d.xpath('//w:ins|//w:del',namespaces=NS)),
            'bookmarks':d.xpath('//w:bookmarkStart/@w:name',namespaces=NS)}


def assets(parts,prefix):
    # Compare payloads, allowing Word to rename package members without false alarms.
    return Counter(v for k,v in parts.items() if k.startswith(prefix))


def audit_assets(before,after):
    for prefix in ['word/embeddings/','word/media/']:
        require(assets(before,prefix)==assets(after,prefix),prefix+' changed')
    return {'media_and_embeddings_identical':True}


def new_bookmark_id(root):
    return 1+max([int(n) for n in root.xpath('//w:bookmarkStart/@w:id',namespaces=NS)]+[0])


def bookmark(p,name,bid):
    start=E.Element(W+'bookmarkStart');start.set(W+'id',str(bid));start.set(W+'name',name)
    p.insert(1 if p.find(W+'pPr') is not None else 0,start)
    end=E.SubElement(p,W+'bookmarkEnd');end.set(W+'id',str(bid))


def reference_indent(parts,p):
    """Keep an existing hanging indent, including one inherited through a style."""
    candidates=[p.find('w:pPr/w:ind',NS)]
    styles=E.fromstring(parts['word/styles.xml']) if 'word/styles.xml' in parts else None
    mapping={s.get(W+'styleId'):s for s in styles} if styles is not None else {}
    ids=p.xpath('./w:pPr/w:pStyle/@w:val',namespaces=NS);seen=set()
    sid=ids[0] if ids else None
    while sid in mapping and sid not in seen:
        seen.add(sid);style=mapping[sid];candidates.append(style.find('w:pPr/w:ind',NS))
        based=style.find(W+'basedOn');sid=based.get(W+'val') if based is not None else None
    for ind in candidates:
        if ind is not None and ind.get(W+'left') and ind.get(W+'hanging'):
            return ind.get(W+'left'),ind.get(W+'hanging')
    return '560','560'


def bibliography(parts,root,heading):
    ps=root.xpath('//w:p',namespaces=NS)
    indices=[i for i,p in enumerate(ps) if text(p).strip()==heading]
    require(len(indices)==1,'Expected one exact bibliography heading')
    refs=[]
    for p in ps[indices[0]+1:]:
        s=text(p)
        if not s.strip():
            if refs:break
            continue
        m=re.match(r'^\[(\d+)\][ \t\u3000]*',s)
        if not m:break
        refs.append((int(m[1]),p,m.end()))
    require(refs and [n for n,_,_ in refs]==list(range(1,len(refs)+1)),'Reference list must have sequential literal [1]..[N] labels; refusing ambiguity')
    require(not any(p.find('w:pPr/w:numPr',NS) is not None for _,p,_ in refs),'Already numbered; inspect instead of applying twice')
    require(not any(n.startswith('AWBib') for n in root.xpath('//w:bookmarkStart/@w:name',namespaces=NS)),'AWBib bookmarks already exist')
    require(not any('ADDIN' in f['code'].upper() for f in field_inventory(root)), 'Managed ADDIN citations present; use owning citation manager')
    left,hanging=reference_indent(parts,refs[0][1])
    num=E.fromstring(parts['word/numbering.xml']) if 'word/numbering.xml' in parts else E.Element(W+'numbering',nsmap={'w':NS['w']})
    aid=max([int(x) for x in num.xpath('./w:abstractNum/@w:abstractNumId',namespaces=NS)]+[-1])+1
    nid=max([int(x) for x in num.xpath('./w:num/@w:numId',namespaces=NS)]+[0])+1
    abstract=E.Element(W+'abstractNum');abstract.set(W+'abstractNumId',str(aid))
    E.SubElement(abstract,W+'multiLevelType').set(W+'val','singleLevel')
    E.SubElement(abstract,W+'name').set(W+'val','AcademicBibliography')
    lvl=E.SubElement(abstract,W+'lvl');lvl.set(W+'ilvl','0')
    for tag,val in [('start','1'),('numFmt','decimal'),('suff','tab'),('lvlText','[%1]'),('lvlJc','left')]:E.SubElement(lvl,W+tag).set(W+'val',val)
    pp=E.SubElement(lvl,W+'pPr')
    tabs=E.SubElement(pp,W+'tabs');tab=E.SubElement(tabs,W+'tab');tab.set(W+'val','num');tab.set(W+'pos',left)
    ind=E.SubElement(pp,W+'ind');ind.set(W+'left',left);ind.set(W+'hanging',hanging)
    rp=deepcopy(refs[0][1].find('w:r/w:rPr',NS))
    if rp is None:rp=E.Element(W+'rPr')
    fonts(rp);lvl.append(rp)
    at=next((i for i,c in enumerate(num) if c.tag==W+'num'),len(num));num.insert(at,abstract)
    inst=E.SubElement(num,W+'num');inst.set(W+'numId',str(nid));E.SubElement(inst,W+'abstractNumId').set(W+'val',str(aid))
    parts['word/numbering.xml']=xml(num)
    # Add the package hooks only for documents that did not already use lists.
    relpath='word/_rels/document.xml.rels';rels=E.fromstring(parts[relpath]);R=E.QName(rels).namespace
    if not any(x.get('Type','').endswith('/numbering') for x in rels):
        used={x.get('Id') for x in rels};i=1
        while f'rId{i}' in used:i+=1
        rel=E.SubElement(rels,'{'+R+'}Relationship');rel.set('Id',f'rId{i}');rel.set('Type','http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering');rel.set('Target','numbering.xml')
        parts[relpath]=xml(rels)
    ct=E.fromstring(parts['[Content_Types].xml']);C=E.QName(ct).namespace
    if not any(x.get('PartName')=='/word/numbering.xml' for x in ct):
        el=E.SubElement(ct,'{'+C+'}Override');el.set('PartName','/word/numbering.xml');el.set('ContentType','application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml')
        parts['[Content_Types].xml']=xml(ct)
    bid=new_bookmark_id(root);refparas=set()
    for n,p,end in refs:
        before=text(p)
        replace_match(p,0,end,lambda rp:[])
        require(text(p)==before[end:],'Reference entry text changed')
        np=pprop(p,'numPr');E.SubElement(np,W+'ilvl').set(W+'val','0');E.SubElement(np,W+'numId').set(W+'val',str(nid))
        # Explicit number-tab position prevents one/two digit labels using different default tabs.
        oldind=pprop(p,'ind')
        for attr in ['firstLine','firstLineChars','leftChars','hangingChars']:oldind.attrib.pop(W+attr,None)
        oldind.set(W+'left',left);oldind.set(W+'hanging',hanging)
        tabs=pprop(p,'tabs')
        for tab in list(tabs):
            if tab.get(W+'pos')==left:tabs.remove(tab)
        tab=E.SubElement(tabs,W+'tab');tab.set(W+'val','num');tab.set(W+'pos',left)
        tabs[:]=sorted(tabs,key=lambda t:int(t.get(W+'pos','0')))
        bookmark(p,f'AWBib{n}',bid);bid+=1;refparas.add(p)
    count=0;groups=[];first=[]
    for p in ps:
        if p in refparas:continue
        original_text=text(p)
        s,_=stream(p)
        matches=list(CITE.finditer(s))
        for m in matches:
            nums=[int(x) for x in re.findall(r'\d+',m[1])]
            require(all(1<=n<=len(refs) for n in nums),'Out-of-range citation: '+m[0])
            expanded=[]
            for item in re.split('[,，]',m[1]):
                ns=[int(x) for x in re.findall(r'\d+',item)]
                require(len(ns)<3 and (len(ns)==1 or ns[0]<=ns[1]),'Invalid citation range')
                expanded.extend(range(ns[0],ns[-1]+1))
            for n in expanded:
                if n not in first:first.append(n)
        for m in reversed(matches):
            def build(rp,m=m):
                if m[1].isdigit():return field(f'REF AWBib{int(m[1])} \\n \\h',m[0],rp)
                nodes=[run('[',rp)]
                for token in re.split(r'(\d+)',m[1]):
                    if token.isdigit():nodes.extend(field(f'REF AWBib{int(token)} \\n \\h \\# "0"',str(int(token)),rp))
                    elif token:nodes.append(run(token,rp))
                nodes.append(run(']',rp));return nodes
            replace_match(p,m.start(),m.end(),build)
            count+=len(re.findall(r'\d+',m[1]));groups.append(m[0])
        require(text(p)==original_text,'Body text changed during citation conversion')
    return {'references':len(refs),'native_numId':nid,'REF_fields_added':count,'citation_groups':len(groups),'first_appearance_order':first,'first_appearance_sequential':first==list(range(1,len(refs)+1))}


def figures(root,label_separator=' ',caption_size=None):
    require(label_separator in {'',' '},'Label separator must be empty or one space')
    captions=[];ps=root.xpath('//w:p',namespaces=NS)
    for p in ps:
        fs=field_inventory(p)
        if any(re.match(r'\s*SEQ\s+图(?:\s|$)',f['code']) for f in fs):
            require(len(fs)==1 and re.fullmatch(r'\s*SEQ\s+图(?:\s+\\\*\s+ARABIC)?\s*',fs[0]['code'],re.I),
                    'Caption has extra fields or SEQ switches; requires targeted editing')
            require(fs[0].get('locked') not in {'true','1','on'},'Caption field is locked')
            captions.append(p)
    require(captions,'No native figure SEQ captions')
    require(not any(n.startswith('AWFig') for n in root.xpath('//w:bookmarkStart/@w:name',namespaces=NS)),'AWFig bookmarks already exist')
    require(not any(p.xpath('.//w:bookmarkStart|.//w:bookmarkEnd',namespaces=NS) for p in captions),'Caption bookmarks already exist; do not invalidate them')
    require(not any(p.xpath('.//w:object|.//w:drawing|.//w:pict|.//w:hyperlink',namespaces=NS) for p in captions),'Complex caption content requires targeted editing')
    bid=new_bookmark_id(root);mapping={}
    for p in captions:
        m=re.fullmatch(r'图\s*(\d+)\s*(.*)',text(p));require(m is not None,'Malformed caption')
        n=int(m[1]);require(n not in mapping,'Duplicate figure number');mapping[n]=p
        rp=deepcopy(p.find('w:r/w:rPr',NS))
        if rp is None:rp=E.Element(W+'rPr')
        fonts(rp);font_size(rp,caption_size)
        for c in list(p):
            if c.tag!=W+'pPr':p.remove(c)
        start=E.SubElement(p,W+'bookmarkStart');start.set(W+'id',str(bid));start.set(W+'name',f'AWFig{n}')
        p.append(run('图'+label_separator,rp));p.extend(field('SEQ 图 \\* ARABIC',str(n),rp))
        end=E.SubElement(p,W+'bookmarkEnd');end.set(W+'id',str(bid));bid+=1
        p.append(run(' '+m[2],rp))
    count=0
    for p in ps:
        if p in captions:continue
        s,_=stream(p)
        for m in reversed(list(re.finditer(r'图\s*(\d+)',s))):
            n=int(m[1]);require(n in mapping,'Missing figure caption')
            replace_match(p,m.start(),m.end(),lambda rp,n=n:field(f'REF AWFig{n} \\h',f'图{label_separator}{n}',rp));count+=1
    return {'captions':len(captions),'REF_fields_added':count}


def normalize(root,args):
    counts={'spaces':0,'blank_paragraphs':0,'captions':0,'word_replacements':0}
    for p in list(root.xpath('//w:p',namespaces=NS)):
        caption=bool(re.match(r'^图\s*\d+ ',text(p)))
        if args.boundary_spaces and not caption:
            s,_=stream(p)
            for m in reversed(list(BOUNDARY.finditer(s))):
                replace_match(p,m.start(),m.end(),lambda rp:[]);counts['spaces']+=1
        if args.adopt_wording:
            s,_=stream(p)
            for m in reversed(list(re.finditer('借鉴',s))):
                replace_match(p,m.start(),m.end(),lambda rp:[run('采用',rp)]);counts['word_replacements']+=1
        if args.boundary_spaces:
            for tag in ['autoSpaceDE','autoSpaceDN']:pprop(p,tag).set(W+'val','0')
        sp=p.find('w:pPr/w:spacing',NS)
        if args.real_blank_lines and sp is not None and int(sp.get(W+'afterLines','0'))>0:
            lines=int(sp.get(W+'afterLines'));require(lines%100==0,'Fractional afterLines requires layout judgment')
            sp.attrib.pop(W+'afterLines');sp.set(W+'after','0')
            following=p.getnext()
            blank=following is not None and following.tag==W+'p' and not text(following).strip() and not following.xpath('.//w:object|.//w:drawing|.//w:br|.//w:pict',namespaces=NS)
            for _ in range(max(0,lines//100-int(blank))):
                q=E.Element(W+'p');spacing=pprop(q,'spacing')
                for k,v in {'before':'0','after':'0','line':'240','lineRule':'auto'}.items():spacing.set(W+k,v)
                rp=deepcopy(p.find('w:pPr/w:rPr',NS))
                if rp is not None:q.find(W+'pPr').append(rp)
                p.addnext(q);counts['blank_paragraphs']+=1
        if (args.caption_fonts or getattr(args,'caption_size',None) is not None) and caption:
            rp=pprop(p,'rPr')
            if args.caption_fonts:fonts(rp)
            font_size(rp,getattr(args,'caption_size',None))
            for r in p.xpath('./w:r|./w:fldSimple/w:r',namespaces=NS):
                if r.find(W+'object') is not None:continue
                rp=r.find(W+'rPr')
                if rp is None:rp=E.Element(W+'rPr');r.insert(0,rp)
                if args.caption_fonts:fonts(rp)
                font_size(rp,getattr(args,'caption_size',None))
            counts['captions']+=1
    return counts


def lead_sentence(root,args):
    """Move a final sentence to paragraph start without rebuilding its REF field."""
    require(args.start_heading and args.end_heading and args.phrase and args.expected_count,
            'Specify --start-heading, --end-heading, --phrase and --expected-count')
    ps=root.xpath('//w:p',namespaces=NS)
    starts=[i for i,p in enumerate(ps) if text(p).strip()==args.start_heading]
    ends=[i for i,p in enumerate(ps) if text(p).strip()==args.end_heading]
    require(len(starts)==len(ends)==1 and starts[0]<ends[0],'Section boundaries ambiguous')
    count=0
    for p in ps[starts[0]+1:ends[0]]:
        before=text(p)
        if args.phrase not in before:continue
        require(before.count(args.phrase)==1 and not before.startswith(args.phrase),'Phrase repeated or already at start')
        a=before.index(args.phrase);sentence=before[a:]
        require(sentence.endswith('。') and sentence.count('。')==1,'Phrase is not a single final sentence')
        target=[t for t in p.xpath('./w:r/w:t',namespaces=NS) if args.phrase in (t.text or '')]
        require(len(target)==1,'Sentence start must be in a direct text run')
        t=target[0];r=t.getparent()
        require(all(c.tag in {W+'rPr',W+'t'} for c in r),'Complex sentence start run')
        at=p.index(r);suffix=list(p)[at+1:]
        require(not any(c.xpath('.//w:bookmarkStart|.//w:bookmarkEnd|self::w:bookmarkStart|self::w:bookmarkEnd',namespaces=NS) for c in suffix),'Moving bookmarks requires scope-aware editing')
        split=t.text.index(args.phrase);new=deepcopy(r);new.find(W+'t').text=t.text[split:];t.text=t.text[:split]
        nodes=[new]+suffix
        for c in suffix:p.remove(c)
        at=1 if p.find(W+'pPr') is not None else 0
        for i,c in enumerate(nodes):p.insert(at+i,c)
        require(text(p)==sentence+before[:a],'Sentence movement changed text')
        count+=1
    require(count==args.expected_count,f'Expected {args.expected_count} sentence moves; found {count}')
    return {'sentences_moved':count}


def save(parts,output):
    require(not output.exists(),'Refusing to overwrite output: '+str(output))
    output.parent.mkdir(parents=True,exist_ok=True)
    with ZipFile(output,'x',ZIP_DEFLATED) as z:
        for n,data in parts.items():z.writestr(n,data)


def main():
    ap=argparse.ArgumentParser(description=__doc__)
    ap.add_argument('command',choices=['inspect','bibliography','figures','lead-sentence','normalize','audit-assets'])
    ap.add_argument('input',type=Path);ap.add_argument('--output',type=Path);ap.add_argument('--report',type=Path)
    ap.add_argument('--compare',type=Path);ap.add_argument('--heading',default='参考文献')
    ap.add_argument('--start-heading');ap.add_argument('--end-heading');ap.add_argument('--phrase');ap.add_argument('--expected-count',type=int)
    ap.add_argument('--label-separator',choices=['space','none'],default='space',help='figures: space between label and number (default), or none')
    ap.add_argument('--caption-size',type=float,help='figures/normalize: caption size in points; preserve when omitted')
    for flag in ['boundary-spaces','real-blank-lines','caption-fonts','adopt-wording']:ap.add_argument('--'+flag,action='store_true')
    args=ap.parse_args()
    if args.caption_size is not None:
        require(args.command in {'figures','normalize'},'--caption-size requires figures or normalize')
        require(args.caption_size>0 and args.caption_size*2==int(args.caption_size*2),'Invalid caption size')
    if args.report:
        require(not args.report.exists(),'Refusing to overwrite report: '+str(args.report))
        require(args.report.resolve()!=args.input.resolve(),'Report must not overwrite input')
        require(args.output is None or args.report.resolve()!=args.output.resolve(),'Report and output paths must differ')
    before=read(args.input);parts=before.copy();root=E.fromstring(parts['word/document.xml'])
    if args.command=='inspect':report=inventory(parts)
    elif args.command=='audit-assets':
        require(args.compare is not None,'--compare required');report=audit_assets(before,read(args.compare))
    else:
        require(args.output is not None,'--output required')
        require(args.input.resolve()!=args.output.resolve(),'Input must not be overwritten')
        require(not root.xpath('//w:ins|//w:del',namespaces=NS),'Tracked revisions present; obtain disposition before mutation')
        if args.command=='normalize':require(any([args.boundary_spaces,args.real_blank_lines,args.caption_fonts,args.adopt_wording,args.caption_size is not None]),'Select at least one normalization flag')
        if args.command=='bibliography':report=bibliography(parts,root,args.heading)
        elif args.command=='figures':report=figures(root,' ' if args.label_separator=='space' else '',args.caption_size)
        elif args.command=='lead-sentence':report=lead_sentence(root,args)
        else:report=normalize(root,args)
        parts['word/document.xml']=xml(root);report.update(audit_assets(before,parts));save(parts,args.output)
        report['output']=str(args.output)
    if args.report:
        args.report.parent.mkdir(parents=True,exist_ok=True);args.report.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf8')
    print(json.dumps(report,ensure_ascii=False,indent=2))


if __name__=='__main__':main()
