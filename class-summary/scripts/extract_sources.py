"""Extract local lecture sources with the Python standard library; not a summarizer."""
import argparse
import hashlib
import json
import posixpath
import re
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
STAMP = re.compile(r'(?<![\d:])(?:(\d{1,3}):)?(\d{1,2}):(\d{2})(?:[,.](\d{1,3}))?(?!\d)')


def read_text(path):
    raw = path.read_bytes()
    for encoding in ('utf-8-sig', 'utf-16', 'gb18030'):
        try:
            return raw.decode(encoding)
        except UnicodeError:
            pass
    raise ValueError('Cannot decode subtitle text; convert it to UTF-8.')


def transcript(path):
    if path.suffix.lower() == '.docx':
        with zipfile.ZipFile(path) as z:
            root = ET.fromstring(z.read('word/document.xml'))
        text = '\n'.join(''.join(t.text or '' for t in p.iter(W + 't'))
                         for p in root.iter(W + 'p'))
    elif path.suffix.lower() in ('.txt', '.srt', '.vtt', '.md'):
        text = read_text(path)
    else:
        raise ValueError('Supported subtitle formats: DOCX, TXT, SRT, VTT, MD.')
    # Keep every source character. Offsets reference this exact extracted text.
    matches = list(STAMP.finditer(text))
    segments = []
    if matches and matches[0].start():
        segments.append({'time': None, 'start_offset': 0,
                         'text': text[:matches[0].start()]})
    for i, m in enumerate(matches):
        hours, minutes, seconds, fraction = m.groups()
        seconds_value = int(hours or 0) * 3600 + int(minutes) * 60 + int(seconds)
        if fraction:
            seconds_value += int(fraction) / 10 ** len(fraction)
        segments.append({'time': m.group(), 'seconds': seconds_value,
                         'start_offset': m.start(),
                         'text': text[m.start():matches[i + 1].start() if i + 1 < len(matches) else len(text)]})
    if not matches:
        segments = [{'time': None, 'start_offset': 0, 'text': text}]
    for i, segment in enumerate(segments, 1):
        segment['id'] = f'T{i:05d}'
    return {'text': text, 'segments': segments,
            'timestamp_note': 'Timestamp candidates, not verified topic boundaries. SRT/VTT include cue end times; numeric times in speech can also match.'}


def slides(path, out):
    if path.suffix.lower() != '.pptx':
        raise ValueError('This helper accepts PPTX only; read PDF/PPT with another available tool.')
    pages = []
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        presentation = ET.fromstring(z.read('ppt/presentation.xml'))
        rels = ET.fromstring(z.read('ppt/_rels/presentation.xml.rels'))
        targets = {r.get('Id'): r.get('Target') for r in rels}
        ids = presentation.findall('.//{http://schemas.openxmlformats.org/presentationml/2006/main}sldId')
        for number, slide_id in enumerate(ids, 1):
            target = targets[slide_id.get(R + 'id')]
            name = target.lstrip('/') if target.startswith('/') else posixpath.normpath('ppt/' + target)
            root = ET.fromstring(z.read(name))
            paragraphs = [''.join(t.text or '' for t in p.iter(A + 't')) for p in root.iter(A + 'p')]
            text = '\n'.join(x for x in paragraphs if x.strip())
            page = {'page': number, 'part': name, 'text': text, 'images': [],
                    'needs_visual_read': True, 'image_only': not bool(text.strip()),
                    'verification': 'unreviewed'}
            relname = posixpath.join(posixpath.dirname(name), '_rels', posixpath.basename(name) + '.rels')
            if relname in names:
                for rel in ET.fromstring(z.read(relname)):
                    if not rel.get('Type', '').endswith('/image'):
                        continue
                    target = rel.get('Target', '')
                    if rel.get('TargetMode') == 'External':
                        page.setdefault('external_images', []).append(target)
                        continue
                    image_name = target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.dirname(name) + '/' + target)
                    if image_name not in names:
                        page.setdefault('missing_images', []).append(image_name)
                        continue
                    data = z.read(image_name)
                    suffix = Path(image_name).suffix.lower()
                    filename = hashlib.sha256(data).hexdigest()[:24] + suffix
                    image_path = out / 'images' / filename
                    image_path.parent.mkdir(parents=True, exist_ok=True)
                    image_path.write_bytes(data)
                    page['images'].append({'relationship': rel.get('Id'), 'file': 'images/' + filename})
            pages.append(page)
    return pages


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--transcript', required=True, type=Path)
    parser.add_argument('--slides', required=True, type=Path)
    parser.add_argument('--out', required=True, type=Path)
    args = parser.parse_args()
    for path in (args.transcript, args.slides):
        if not path.is_file():
            parser.error(f'Input file not found: {path}')
    # Refuse to replace an earlier index or write into a source folder by accident.
    if args.out.exists() and any(args.out.iterdir()):
        parser.error('Output directory must be new or empty; choose another task directory.')
    args.out.mkdir(parents=True, exist_ok=True)
    try:
        evidence = {'inputs': {'transcript': args.transcript.name, 'slides': args.slides.name},
                    'transcript': transcript(args.transcript),
                    'slides': slides(args.slides, args.out)}
    except (ValueError, KeyError, zipfile.BadZipFile, ET.ParseError) as error:
        parser.exit(1, f'Extraction failed: {error}\n')
    (args.out / 'evidence.json').write_text(json.dumps(evidence, ensure_ascii=False, indent=2), encoding='utf-8')
    md = ['# Lecture evidence index', '', 'Extraction only. Visually review slides and verify timestamp candidates before summarizing.', '', '## Transcript', '']
    for segment in evidence['transcript']['segments']:
        md.extend([f"### {segment['id']} {segment['time'] or 'No timestamp'}", '', segment['text'], ''])
    md.extend(['## Slides', ''])
    for slide in evidence['slides']:
        md.extend([f"### PPT page {slide['page']}", '', slide['text'] or '[No native text: visual reading required]', ''])
        if slide.get('missing_images'):
            md.extend(['[Broken image references: page may be incomplete. Verify the original/rendered slide.]', ''])
        for img in slide['images']:
            md.extend([f"![Page {slide['page']}]({img['file']})", ''])
    (args.out / 'evidence.md').write_text('\n'.join(md), encoding='utf-8')
    print(json.dumps({'pages': len(evidence['slides']),
                      'image_only_pages': sum(p['image_only'] for p in evidence['slides']),
                      'pages_with_missing_images': sum(bool(p.get('missing_images')) for p in evidence['slides']),
                      'subtitle_segments': len(evidence['transcript']['segments']),
                      'output': str(args.out)}, ensure_ascii=False))


if __name__ == '__main__':
    main()
