#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
日報ファイル (******_日報.txt) の文字コードを検出し、
SHIFT-JIS のものをバックアップフォルダに退避して UTF-8 に変換する。
"""
import os
import shutil
from pathlib import Path
import chardet

def detect_encoding(filepath, max_bytes=1024*100):
    """ファイルの文字コードを検出（大きなファイル対策）"""
    with open(filepath, 'rb') as f:
        raw = f.read(max_bytes)  # 最初の100KB のみ読む
    if not raw:
        return 'utf-8', 1.0
    try:
        result = chardet.detect(raw)
        return result['encoding'], result['confidence']
    except Exception as e:
        print(f"  [WARN] 検出失敗 ({filepath.name}): {e}")
        return 'utf-8', 0.5

def main():
    # 作業ディレクトリ
    work_dir = Path(r"c:\Users\10001179776\OneDrive - DENSO\2017\else\memo")
    backup_dir = work_dir / "backups_shiftjis_all"
    backup_dir.mkdir(exist_ok=True)
    
    # すべての .txt ファイルを対象に変換
    nippo_files = list(work_dir.glob("*.txt"))
    # バックアップスクリプト自身と backups フォルダを除外
    nippo_files = [f for f in nippo_files if not f.name.startswith('convert_') and 'backup' not in f.name.lower()]
    # 既にバックアップ済みのファイルはスキップ
    already_backed = {b.name for b in backup_dir.glob("*.txt")}
    nippo_files = [f for f in nippo_files if f.name not in already_backed]
    print(f"[INFO] 対象ファイル数: {len(nippo_files)} (バックアップ済みを除く)")
    
    converted = []
    already_utf8 = []
    
    for fpath in nippo_files:
        enc, conf = detect_encoding(fpath)
        # SHIFT-JIS / CP932 系を検出（confidence が低い場合は警告）
        is_sjis = enc and ('shift' in enc.lower() or 'cp932' in enc.lower())
        
        if is_sjis:
            print(f"[SHIFT-JIS] {fpath.name} (detected: {enc}, confidence: {conf:.2f})")
            # バックアップ
            backup_path = backup_dir / fpath.name
            shutil.copy2(fpath, backup_path)
            print(f"  → バックアップ: {backup_path}")
            
            # UTF-8 に変換して上書き
            with open(fpath, 'r', encoding='shift_jis', errors='replace') as f:
                content = f.read()
            with open(fpath, 'w', encoding='utf-8', errors='replace') as f:
                f.write(content)
            print(f"  → UTF-8 変換完了")
            converted.append(fpath.name)
        else:
            # UTF-8 などその他
            already_utf8.append(fpath.name)
    
    print("\n" + "="*60)
    print(f"[完了] SHIFT-JIS → UTF-8 変換: {len(converted)} 件")
    print(f"[完了] 既に UTF-8 等: {len(already_utf8)} 件")
    print(f"[バックアップ先] {backup_dir}")
    print("="*60)

if __name__ == "__main__":
    main()
