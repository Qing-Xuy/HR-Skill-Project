# -*- coding: utf-8 -*-
"""
素材读取器（经验萃取专家 · 模式 B 用）

支持格式：.txt  .md  .docx  .srt  .vtt
输出 JSON：
{
  "source": "文件路径",
  "format": "srt",
  "speakers": ["张三", "李四"],
  "segments": [{"index": 1, "time": "00:00:01", "speaker": "张三", "text": "..."}],
  "plain": "清洗后的纯文本"
}

用法：
    python load_source.py <文件路径> [--out 输出.json] [--no-clean]
"""
import argparse
import json
import os
import re
import sys


def read_text_file(path):
    encodings = ['utf-8', 'utf-8-sig', 'gbk', 'gb18030']
    for enc in encodings:
        try:
            with open(path, 'r', encoding=enc) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    raise RuntimeError('无法识别文件编码：%s' % path)


def read_docx(path):
    try:
        from docx import Document
    except ImportError:
        raise RuntimeError('读取 .docx 需要 python-docx，请先安装：pip install python-docx')
    doc = Document(path)
    lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    for t in doc.tables:
        for row in t.rows:
            cells = [c.text.strip() for c in row.cells]
            line = ' | '.join([c for c in cells if c])
            if line:
                lines.append(line)
    return '\n'.join(lines)


def parse_srt(text):
    """解析 srt / vtt"""
    blocks = re.split(r'\n\s*\n', text.replace('\r\n', '\n'))
    segs = []
    idx = 0
    for b in blocks:
        b = b.strip()
        if not b:
            continue
        lines = [l for l in b.split('\n') if l.strip()]
        if not lines:
            continue
        # vtt 头部
        if lines[0].upper().startswith('WEBVTT'):
            continue
        time = ''
        body_start = 0
        if re.match(r'^\d+$', lines[0].strip()):
            if len(lines) > 1 and '-->' in lines[1]:
                time = lines[1].split('-->')[0].strip()
                body_start = 2
            else:
                body_start = 1
        elif '-->' in lines[0]:
            time = lines[0].split('-->')[0].strip()
            body_start = 1
        content = ' '.join(l.strip() for l in lines[body_start:])
        content = re.sub(r'<[^>]+>', '', content)          # 去 vtt 标签
        content = re.sub(r'\{[^}]*\}', '', content)
        content = content.strip()
        if not content:
            continue
        idx += 1
        segs.append({'index': idx, 'time': time, 'speaker': '', 'text': content})
    return segs


def clean_speech(text):
    """口语清洗：去掉明显的语气 filler，保留实词与原话细节"""
    # 去掉重复的同一词语造成的口吃，如 "这个这个"
    text = re.sub(r'([\u4e00-\u9fa5]{2,4})\1(?=[\u4e00-\u9fa5，。])', r'\1', text)
    return text.strip()


def split_speaker(line):
    """识别 '张三：xxx' / '张三: xxx' / '[00:12] 张三：xxx' 形式"""
    m = re.match(r'^\s*(?:\[[^\]]*\])?\s*([^\s：:]{1,12})[：:]\s*(.+)$', line)
    if m and len(m.group(1)) <= 12:
        return m.group(1).strip(), m.group(2).strip()
    return '', line.strip()


def build_segments(raw, fmt, do_clean):
    if fmt in ('srt', 'vtt'):
        segs = parse_srt(raw)
        if do_clean:
            for s in segs:
                s['text'] = clean_speech(s['text'])
        return segs

    if fmt == 'docx':
        lines = raw.split('\n')
    else:
        lines = raw.split('\n')

    segs = []
    idx = 0
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if line.startswith('#'):        # markdown 标题也保留，作为情境标记
            segs.append({'index': 0, 'time': '', 'speaker': '', 'text': line})
            continue
        sp, content = split_speaker(line)
        if do_clean:
            content = clean_speech(content)
        if not content:
            continue
        idx += 1
        segs.append({'index': idx, 'time': '', 'speaker': sp, 'text': content})
    return segs


def main():
    ap = argparse.ArgumentParser(description='读取萃取素材')
    ap.add_argument('path', help='素材文件路径')
    ap.add_argument('--out', help='输出 JSON 路径（不填则只打印摘要）')
    ap.add_argument('--no-clean', action='store_true', help='跳过口语清洗')
    args = ap.parse_args()

    path = args.path
    if not os.path.exists(path):
        print('文件不存在：%s' % path, file=sys.stderr)
        sys.exit(1)

    ext = os.path.splitext(path)[1].lower().lstrip('.')
    if ext == 'docx':
        raw, fmt = read_docx(path), 'docx'
    elif ext in ('srt', 'vtt'):
        raw, fmt = read_text_file(path), ext
    elif ext in ('txt', 'md', 'markdown', 'text'):
        raw, fmt = read_text_file(path), ext
    else:
        print('不支持的格式：.%s（支持 txt/md/docx/srt/vtt）' % ext, file=sys.stderr)
        sys.exit(2)

    segs = build_segments(raw, fmt, not args.no_clean)
    speakers = sorted({s['speaker'] for s in segs if s['speaker']})
    plain = '\n'.join(
        (('%s：%s' % (s['speaker'], s['text'])) if s['speaker'] else s['text'])
        for s in segs
    )

    result = {
        'source': os.path.abspath(path),
        'format': fmt,
        'segments_count': len(segs),
        'chars': len(plain),
        'speakers': speakers,
        'segments': segs,
        'plain': plain,
    }

    print('格式：%s | 片段数：%d | 字数：%d' % (fmt, len(segs), len(plain)))
    print('说话人：%s' % ('、'.join(speakers) if speakers else '（未识别到，建议在素材里加「姓名：」标签）'))
    print('--- 前 500 字预览 ---')
    print(plain[:500])

    if args.out:
        with open(args.out, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        print('\n已写入：%s' % args.out)


if __name__ == '__main__':
    main()
