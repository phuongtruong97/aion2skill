#!/usr/bin/env python3
"""
build_all.py
============
Wrapper cho minify_data.py — dùng trong GitHub Action.
Tự quét TẤT CẢ thư mục patch trong data/ (vd: data/2026_06_24, data/2026_07_15...)
và build lại data_min/ tương ứng cho từng patch, không cần biết trước tên patch.

Cũng copy data/versions.json -> data_min/versions.json.

Cách dùng (từ root repo):
    python3 build/minify_data_ci.py
"""

import os
import shutil
import subprocess
import sys

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(REPO_ROOT, "data")
DATA_MIN_DIR = os.path.join(REPO_ROOT, "data_min")
MINIFY_SCRIPT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "minify_data.py")


def is_patch_folder(path):
    """Một thư mục patch hợp lệ phải chứa Skill.json (file bắt buộc)."""
    return os.path.isfile(os.path.join(path, "Skill.json"))


def main():
    if not os.path.isdir(DATA_DIR):
        print(f"Không tìm thấy thư mục {DATA_DIR}, dừng.")
        sys.exit(1)

    os.makedirs(DATA_MIN_DIR, exist_ok=True)

    # 1. Copy versions.json nếu có
    versions_src = os.path.join(DATA_DIR, "versions.json")
    if os.path.isfile(versions_src):
        shutil.copy2(versions_src, os.path.join(DATA_MIN_DIR, "versions.json"))
        print("Đã copy versions.json")

    # 2. Tìm tất cả thư mục patch (chứa Skill.json) trong data/
    patch_folders = []
    for name in sorted(os.listdir(DATA_DIR)):
        full_path = os.path.join(DATA_DIR, name)
        if os.path.isdir(full_path) and is_patch_folder(full_path):
            patch_folders.append(name)

    if not patch_folders:
        print("Không tìm thấy thư mục patch nào (cần có Skill.json bên trong).")
        sys.exit(1)

    print(f"Tìm thấy {len(patch_folders)} patch: {', '.join(patch_folders)}")
    print()

    # 3. Chạy minify_data.py cho từng patch
    any_failed = False
    for patch in patch_folders:
        in_path = os.path.join(DATA_DIR, patch)
        out_path = os.path.join(DATA_MIN_DIR, patch)
        print(f"=== Đang lọc patch: {patch} ===")
        result = subprocess.run(
            [sys.executable, MINIFY_SCRIPT, in_path, out_path],
            cwd=REPO_ROOT,
        )
        if result.returncode != 0:
            any_failed = True
            print(f"  [LỖI] Lọc patch {patch} thất bại.")
        print()

    if any_failed:
        sys.exit(1)

    print("Hoàn tất build data_min/ cho tất cả patch.")


if __name__ == "__main__":
    main()
