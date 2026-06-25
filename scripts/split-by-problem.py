#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
split-by-problem.py - 把 docx 拆分成 solution/*.md

用法：
    python3 scripts/split-by-problem.py <input.docx> [--output-dir solution]

功能：
- 识别章节标题（一、二、xxx）和题号（数字. 标题 【难度】）
- 拆出每个题为独立 md 文件 solution/NNNN.标题.md
- 抽取 docx 内的图片到 solution/assets/
- md 里用相对路径 assets/NNNN_img_K.png 引用
- 清洗 docx 残留的 HTML 实体 (&lt; &gt; &amp;)
- 修复常见 OCR 错误（如「商单」→「简单」）
"""

import argparse
import os
import re
import sys
import shutil
import zipfile
from collections import OrderedDict
from pathlib import Path

try:
    from docx import Document
    from docx.oxml.ns import qn
except ImportError:
    sys.exit("❌ python-docx 未安装。运行：pip install python-docx")


# ============== 配置 ==============
# 已知的 OCR 错误 → 正确难度（基于 LeetCode 官方数据）
KNOWN_DIFF_FIXES = {
    1: '简单',   # 商单 → 简单
}

# 中文数字 → 阿拉伯数字（用于章节标题排序）
CHINESE_NUMS = '一二三四五六七八九十'


# ============== 工具函数 ==============
def clean_text(text: str) -> str:
    """清洗 docx 残留的 HTML 实体、多余空白。"""
    if not text:
        return ''
    # HTML 实体反转义
    text = text.replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&')
    text = text.replace('&quot;', '"').replace('&apos;', "'").replace('&nbsp;', ' ')
    # 去除行尾空白
    text = re.sub(r'[ \t]+', ' ', text)
    return text.strip()


def extract_paragraphs(docx_path: str):
    """
    解析 docx，返回带图片位置信息的段落列表。
    每项是 dict: {'text': str, 'images': [(rid, image_bytes, ext)], 'style': str}
    """
    doc = Document(docx_path)
    
    # 第一步：先把所有图片的二进制读出来，存到 rid → (bytes, ext) 映射
    rid_to_image = {}
    with zipfile.ZipFile(docx_path) as z:
        for name in z.namelist():
            if name.startswith('word/media/'):
                ext = name.rsplit('.', 1)[-1].lower()
                if ext in ('png', 'jpg', 'jpeg', 'gif', 'bmp', 'svg', 'emf', 'wmf'):
                    rid_to_image[name] = (z.read(name), ext)
    
    # 第二步：从 document.xml.rels 拿 rid → 实际文件路径
    rels_xml = ''
    with zipfile.ZipFile(docx_path) as z:
        try:
            rels_xml = z.read('word/_rels/document.xml.rels').decode('utf-8')
        except KeyError:
            pass
    
    rid_to_path = {}
    for m in re.finditer(r'Id="([^"]+)"\s+Type="[^"]*image[^"]*"\s+Target="([^"]+)"', rels_xml):
        rid_to_path[m.group(1)] = m.group(2)
    
    # 第三步：遍历每个段落，提取文本 + 关联的图片
    paragraphs = []
    for para in doc.paragraphs:
        text_parts = []
        images = []
        
        # 遍历段落里所有的 run（含图片）
        for run in para.runs:
            # 文本
            if run.text:
                text_parts.append(run.text)
            # 图片：run 里嵌了 <w:drawing> 或 <w:pict> 的 blip
            drawing = run._element.findall('.//' + qn('w:drawing'))
            pict = run._element.findall('.//' + qn('w:pict'))
            for d in drawing + pict:
                blip = d.find('.//' + qn('a:blip')) or d.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
                if blip is not None:
                    rid = blip.get(qn('r:embed')) or blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                    if rid and rid in rid_to_path:
                        target = rid_to_path[rid]  # e.g. "media/image3.png"
                        full_name = 'word/' + target
                        if full_name in rid_to_image:
                            images.append(rid_to_image[full_name])
        
        full_text = clean_text(''.join(text_parts))
        if full_text or images:
            paragraphs.append({
                'text': full_text,
                'images': images,
            })
    
    return paragraphs


def extract_images(paragraphs, asset_dir: Path, problem_num: int):
    """
    把段落里的图片抽取到 asset_dir，返回 markdown 引用列表。
    文件命名：{problem_num:04d}_img_{seq}.png
    """
    asset_dir.mkdir(parents=True, exist_ok=True)
    refs = []
    seq = 1
    for para in paragraphs:
        for img_bytes, ext in para.get('images', []):
            fname = f'{problem_num:04d}_img_{seq}.{ext}'
            fpath = asset_dir / fname
            with open(fpath, 'wb') as f:
                f.write(img_bytes)
            refs.append(f'assets/{fname}')
            seq += 1
    return refs


def parse_problems(paragraphs):
    """
    核心解析：把段落列表切成 (chapter, problem_dict) 列表。
    """
    chapter_pattern = re.compile(r'^([一二三四五六七八九十]+)、(.+)$')
    problem_pattern = re.compile(r'^(\d+)\.\s*(.+?)\s*【(.+?)】')
    
    problems = []
    current_cat = None
    current_problem = None
    skip_until_next_problem = False
    
    # 我们需要按顺序遍历：识别章节标题 → 识别题目 → 收集题目后续段落（描述、示例、笔记）→ 直到下一题
    
    i = 0
    while i < len(paragraphs):
        para = paragraphs[i]
        text = para['text']
        
        # 跳过空文本
        if not text and not para['images']:
            i += 1
            continue
        
        # 章节标题检测
        cat_m = chapter_pattern.match(text)
        if cat_m:
            cat_name = cat_m.group(2).strip()
            if len(cat_name) <= 20:
                # 确认是章节标题（后面应该有题号）
                nearby_texts = ' '.join(p['text'] for p in paragraphs[i+1:i+6])
                if re.search(r'^\d+\.', nearby_texts, re.M) or '本类共' in nearby_texts:
                    current_cat = cat_name.split('本类共')[0].strip()
                    i += 1
                    continue
        
        # 题目检测
        p_m = problem_pattern.match(text)
        if p_m:
            num = int(p_m.group(1))
            title = p_m.group(2).strip()
            diff_raw = p_m.group(3)
            
            # 校正 OCR 错误
            if num in KNOWN_DIFF_FIXES:
                diff = KNOWN_DIFF_FIXES[num]
            elif diff_raw in ('简单', '中等', '困难'):
                diff = diff_raw
            else:
                # 未知难度，保守设为「中等」并打印警告
                print(f'  ⚠️  #{num} 难度原始为「{diff_raw}」，设为「中等」')
                diff = '中等'
            
            # 开始收集这一题的内容
            current_problem = {
                'num': num,
                'title': title,
                'diff': diff,
                'category': current_cat or '未分类',
                'paragraphs': [para],  # 含题号那行
            }
            problems.append(current_problem)
            i += 1
            continue
        
        # 既不是章节也不是题目：归属于当前题目
        if current_problem is not None:
            current_problem['paragraphs'].append(para)
        i += 1
    
    return problems


def problem_to_markdown(problem, asset_dir: Path) -> str:
    """
    把单题的所有段落渲染成 markdown。
    - 题号那行变成 # 标题
    - 学霸笔记（以 ✎ 开头）变成 ## 学霸笔记
    - 示例段落格式化
    - 图片以 markdown 引用插入到合适位置
    """
    paras = problem['paragraphs']
    num = problem['num']
    title = problem['title']
    diff = problem['diff']
    category = problem['category']
    
    out = []
    out.append(f'# {num}. {title}\n')
    out.append(f'\n> 难度：{diff} · 章节：{category}\n')
    out.append('\n---\n')
    
    # 跳过第一段（题号那行已经用作了标题）
    # 剩下的段落智能归类：
    #   - 包含「示例 N：」 → 后面是示例
    #   - 包含「✎」或以「学霸笔记」开头 → 学霸笔记
    #   - 其他 → 题目描述
    
    in_notes = False
    in_example = False
    
    # 图片累计到合适位置插入（这里简化：每段如果含图片，紧跟该段文本后插入）
    img_refs = []
    img_counter = [1]  # 用 list 解决闭包
    
    for idx, para in enumerate(paras[1:], start=1):  # 跳过第一行（题号）
        text = para['text']
        images = para.get('images', [])
        
        if not text and not images:
            continue
        
        # 标题段（学霸笔记/示例）
        if re.match(r'^✎\s*学霸笔记', text) or text.startswith('学霸笔记'):
            out.append(f'\n## 学霸笔记\n\n')
            in_notes = True
            in_example = False
            # 如果标题行有内容（如 "学霸笔记："），去掉前缀
            text = re.sub(r'^[✎]?\s*学霸笔记[：:]?\s*', '', text).strip()
            if text:
                out.append(f'{text}\n')
        elif re.match(r'^示例\s*\d+', text):
            out.append(f'\n{text}\n')
            in_example = True
            in_notes = False
        elif re.match(r'^(输入|输出|解释|说明)', text):
            out.append(f'- {text}\n')
            in_example = True
        elif in_notes:
            out.append(f'{text}\n')
        elif in_example:
            # 在示例块里，直接输出
            out.append(f'{text}\n')
        else:
            # 题目描述部分
            if '题目描述' not in '\n'.join(out[-5:]):  # 第一次进入描述时加标题
                out.append(f'\n## 题目描述\n\n')
            out.append(f'{text}\n')
        
        # 图片插入到段落后面
        for img_bytes, ext in images:
            fname = f'{num:04d}_img_{img_counter[0]}.{ext}'
            fpath = asset_dir / fname
            with open(fpath, 'wb') as f:
                f.write(img_bytes)
            out.append(f'\n![配图](assets/{fname})\n')
            img_counter[0] += 1
    
    return ''.join(out)


# ============== 主流程 ==============
def main():
    parser = argparse.ArgumentParser(description='把 docx 拆分成 solution/*.md')
    parser.add_argument('input', help='输入 docx 文件路径')
    parser.add_argument('--output-dir', '-o', default='solution',
                        help='输出目录（默认 ./solution）')
    parser.add_argument('--dry-run', action='store_true', help='只解析，不写文件')
    args = parser.parse_args()
    
    input_path = Path(args.input).resolve()
    if not input_path.exists():
        sys.exit(f'❌ 输入文件不存在: {input_path}')
    
    output_dir = Path(args.output_dir).resolve()
    asset_dir = output_dir / 'assets'
    
    print(f'📖 解析 docx: {input_path.name}')
    paragraphs = extract_paragraphs(str(input_path))
    print(f'   总段落数: {len(paragraphs)}')
    
    print(f'🔍 切分题目...')
    problems = parse_problems(paragraphs)
    print(f'   识别到 {len(problems)} 道题')
    
    if args.dry_run:
        for p in problems:
            print(f'   - #{p["num"]:04d} {p["title"]} 【{p["diff"]}】 ({p["category"]}) · {len(p["paragraphs"])} 段')
        return
    
    # 写入
    output_dir.mkdir(parents=True, exist_ok=True)
    asset_dir.mkdir(parents=True, exist_ok=True)
    
    # 清空旧文件
    for old in output_dir.glob('[0-9]*.md'):
        old.unlink()
    if asset_dir.exists():
        shutil.rmtree(asset_dir)
        asset_dir.mkdir(parents=True)
    
    written_count = 0
    skipped_count = 0
    
    for problem in problems:
        # 过滤：只保留有实质内容的题（描述/示例/笔记/图片至少有其一）
        para_text_combined = ' '.join(p['text'] for p in problem['paragraphs'][1:])  # 跳过题号那行
        has_content = bool(para_text_combined.strip())
        has_images = any(p.get('images') for p in problem['paragraphs'])
        
        if not has_content and not has_images:
            print(f'   ⏭️  跳过 #{problem["num"]:04d} {problem["title"]}（无描述/笔记/图片，标记为未写）')
            skipped_count += 1
            continue
        
        md = problem_to_markdown(problem, asset_dir)
        fname = f'{problem["num"]:04d}.{problem["title"]}.md'
        # 清理文件名里的非法字符
        fname = re.sub(r'[\\/:*?"<>|]', '_', fname)
        fpath = output_dir / fname
        with open(fpath, 'w', encoding='utf-8') as f:
            f.write(md)
        img_count = sum(len(p.get('images', [])) for p in problem['paragraphs'])
        print(f'   ✓ {fname}  ({img_count} 张图, {len(md)} 字符)')
        written_count += 1
    
    print(f'\n✅ 完成！共识别 {len(problems)} 题')
    print(f'   生成 md: {written_count} 个（已写）')
    print(f'   跳过:    {skipped_count} 个（未写，等后续补）')
    print(f'   目录: {output_dir}')
    print(f'   图片: {asset_dir}')


if __name__ == '__main__':
    main()