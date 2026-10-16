# -*- coding: utf-8 -*-
"""
打包「经验萃取专家 / Experience Extraction Expert」Skill，便于发给同事。

用法：
    python package_skill.py                # 默认输出到 Skill 同级目录
    python package_skill.py --out D:\\tmp   # 指定输出目录
    python package_skill.py --extract D:\\  # 反向：把 zip 解压到指定 skills 目录

产出：<技能名>.zip，解压到自己本机的 skills 目录即可使用。
skills 目录通常是：
    Windows: %USERPROFILE%\\.workbuddy\\skills
    macOS/Linux: ~/.workbuddy/skills
"""
import argparse
import os
import zipfile

SKILL_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SKILL_NAME = os.path.basename(SKILL_ROOT)

EXCLUDE_DIRS = {'__pycache__', '.git', '.idea', '.vscode', '.pytest_cache'}
EXCLUDE_FILES = {'.DS_Store', 'Thumbs.db'}
EXCLUDE_EXT = {'.pyc', '.pyo', '.zip'}


def build_zip(out_dir):
    os.makedirs(out_dir, exist_ok=True)
    zip_path = os.path.join(out_dir, '%s.zip' % SKILL_NAME)
    count = 0
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for root, dirs, files in os.walk(SKILL_ROOT):
            dirs[:] = [d for d in dirs if d not in EXCLUDE_DIRS]
            for fn in files:
                if fn in EXCLUDE_FILES or os.path.splitext(fn)[1].lower() in EXCLUDE_EXT:
                    continue
                full = os.path.join(root, fn)
                arc = os.path.relpath(full, os.path.dirname(SKILL_ROOT))
                z.write(full, arc)
                count += 1
    size = os.path.getsize(zip_path) / 1024
    print('已打包 %d 个文件 -> %s（%.1f KB）' % (count, zip_path, size))
    print('\n对方使用方法：')
    print('  1. 解压 zip，得到文件夹「%s」' % SKILL_NAME)
    print('  2. 整个文件夹复制到本机的 skills 目录：')
    print('     Windows: %%USERPROFILE%%\\.workbuddy\\skills')
    print('     macOS/Linux: ~/.workbuddy/skills')
    print('  3. 重启或刷新对话即可生效')
    return zip_path


def extract_zip(zip_path, target_dir):
    if not os.path.exists(zip_path):
        print('zip 不存在：%s' % zip_path, file=sys.stderr)
        raise SystemExit(1)
    with zipfile.ZipFile(zip_path, 'r') as z:
        z.extractall(target_dir)
    print('已解压到：%s' % target_dir)
    print('目标路径应为：%s' % os.path.join(target_dir, SKILL_NAME))


if __name__ == '__main__':
    import sys
    ap = argparse.ArgumentParser(description='打包 / 解包 Skill')
    ap.add_argument('--out', help='zip 输出目录，默认 Skill 同级')
    ap.add_argument('--extract', help='解压 *.zip 到指定 skills 目录')
    args = ap.parse_args()

    if args.extract:
        extract_zip(args.extract, args.extract)
    else:
        build_zip(args.out or os.path.dirname(SKILL_ROOT))
