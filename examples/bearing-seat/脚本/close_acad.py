"""按完整路径关闭一份已保存的图纸。AutoCAD 会处理关闭时的保存提示。"""
import argparse
from pathlib import Path

import win32com.client


def close_saved_drawing(app, target):
    for doc in app.Documents:
        if doc.Path and Path(doc.FullName).resolve() == target:
            if not doc.Saved:
                raise RuntimeError(f"图纸有未保存的修改：{target}")
            doc.Close(True)
            return
    raise RuntimeError(f"未找到已打开的图纸：{target}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("drawing", type=Path, help="要关闭的 DWG 完整路径")
    args = parser.parse_args()
    if not args.drawing.is_absolute() or args.drawing.suffix.lower() != ".dwg":
        parser.error("请提供 DWG 的完整路径")
    app = win32com.client.GetActiveObject("AutoCAD.Application")
    close_saved_drawing(app, args.drawing.resolve())
    print(f"已关闭：{args.drawing}")


if __name__ == "__main__":
    main()
