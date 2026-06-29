#!/usr/bin/env python3
"""
minify_data.py
==============
Lọc các file JSON dữ liệu gốc (export từ game) xuống chỉ còn những field
mà aion2skill (script.js / compare.js) thực sự đọc tới.

Vấn đề gốc: Skill.json (83MB), SkillEffect.json (25MB), SkillEffectLv.json (32MB)...
chứa hàng chục field debug/engine (FX, animation, AI...) không bao giờ được
dùng để hiển thị skill, nhưng vẫn bị tải về máy người dùng mỗi lần mở trang.

Cách dùng:
    python3 minify_data.py <input_folder> <output_folder>

Ví dụ:
    python3 minify_data.py data/2026_06_24 data_min/2026_06_24

Script này KHÔNG đổi cấu trúc top-level "Properties.Data: [...]" để code
script.js/compare.js hiện tại không cần sửa logic parse, chỉ cần đổi
đường dẫn fetch (hoặc dùng nguyên data_min/ thay cho data/).
"""

import json
import os
import sys

# Field cần giữ cho từng file, dựa theo những gì scanSkill/scanJson/scanJsonLv
# trong script.js và compare.js thực sự đọc.
KEEP_FIELDS = {
    "Skill.json": [
        "ID", "SkillIcon", "NeedCoolTime", "NeedCostMp",
        "SkillLvGroupId", "SkillType",
    ],
    "SkillLv.json": [
        "SkillLvGroupId", "SkillLv", "NeedCostMp", "NeedCoolTime",
    ],
    "SkillEffect.json": [
        "ID", "EffectValueList", "SkillEffectLvGroupId", "AggroAbsolute",
    ],
    "SkillAbnormalEffect.json": [
        "ID", "Values", "AbnormalEffectLvGroupId",
    ],
    "SkillEffectFilter.json": [
        "ID", "TargetCountMax", "TargetCount",
    ],
    # Các file *Lv.json có cấu trúc lồng khác (đọc qua scanJsonLv), field
    # nằm thẳng trong từng item của Properties.Data
    "SkillEffectLv.json": [
        "SkillEffectLvGroupId", "SkillLvGroupId", "SkillEffectLv",
        "EffectValueList", "CostMPoint", "CoolTime",
    ],
    "SkillAbnormalEffectLv.json": [
        "AbnormalEffectLevelGroupId", "SkillLvGroupId", "AbnormalEffectLevel",
        "Values", "CostMPoint", "CoolTime",
    ],
}

# ID có thể nằm trong dạng {"Value": x} (wrapped) — giữ nguyên dạng này
# vì code parse có kiểm tra d.ID.Value
WRAPPED_ID_KEYS = {"ID"}


def prune_item(item, keep_fields):
    """Giữ lại chỉ các field trong keep_fields, bỏ hết phần còn lại."""
    if not isinstance(item, dict):
        return item
    pruned = {}
    for k in keep_fields:
        if k in item:
            pruned[k] = item[k]
    return pruned


def minify_file(input_path, output_path, keep_fields):
    with open(input_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    if not isinstance(data, dict) or "Properties" not in data or "Data" not in data.get("Properties", {}):
        print(f"  [SKIP] {input_path}: cấu trúc không như mong đợi, copy nguyên bản")
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, separators=(",", ":"))
        return os.path.getsize(input_path), os.path.getsize(output_path)

    original_items = data["Properties"]["Data"]
    pruned_items = [prune_item(item, keep_fields) for item in original_items]

    # Giữ nguyên field "Ids" nếu code cần (script hiện tại không dùng, nhưng
    # giữ lại để an toàn — chi phí gần như 0 vì chỉ là list số ID)
    out_data = {
        "Properties": {"Data": pruned_items}
    }
    if "Ids" in data:
        out_data["Ids"] = data["Ids"]

    with open(output_path, "w", encoding="utf-8") as f:
        # separators gọn (không có khoảng trắng) để giảm thêm dung lượng
        json.dump(out_data, f, ensure_ascii=False, separators=(",", ":"))

    return os.path.getsize(input_path), os.path.getsize(output_path)


def human(n):
    for unit in ["B", "KB", "MB", "GB"]:
        if n < 1024:
            return f"{n:.1f}{unit}"
        n /= 1024
    return f"{n:.1f}TB"


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(1)

    input_folder = sys.argv[1]
    output_folder = sys.argv[2]
    os.makedirs(output_folder, exist_ok=True)

    total_before = 0
    total_after = 0

    for filename, keep_fields in KEEP_FIELDS.items():
        in_path = os.path.join(input_folder, filename)
        out_path = os.path.join(output_folder, filename)
        if not os.path.exists(in_path):
            print(f"  [WARN] Không tìm thấy {in_path}, bỏ qua")
            continue

        before, after = minify_file(in_path, out_path, keep_fields)
        total_before += before
        total_after += after
        pct = 100 * (1 - after / before) if before else 0
        print(f"  {filename:35s} {human(before):>10s} -> {human(after):>10s}  (giảm {pct:.1f}%)")

    # text.xlsx: copy nguyên, không nén ở đây (xử lý riêng nếu cần)
    xlsx_path = os.path.join(input_folder, "text.xlsx")
    if os.path.exists(xlsx_path):
        import shutil
        out_xlsx = os.path.join(output_folder, "text.xlsx")
        shutil.copy2(xlsx_path, out_xlsx)
        size = os.path.getsize(xlsx_path)
        total_before += size
        total_after += size
        print(f"  {'text.xlsx':35s} {human(size):>10s} -> {human(size):>10s}  (giữ nguyên)")

    print()
    print(f"TỔNG: {human(total_before)} -> {human(total_after)}  "
          f"(giảm {100*(1-total_after/total_before):.1f}%)")


if __name__ == "__main__":
    main()
