from __future__ import annotations

import sys
from pathlib import Path


def status_line(name: str, ok: bool, detail: str = "") -> None:
    marker = "正常" if ok else "异常"
    suffix = f"：{detail}" if detail else ""
    print(f"{name}：{marker}{suffix}")


def main() -> int:
    failed = False
    print("PPT 图片互转工具 v3.3 运行自检\n")

    modules = [
        ("Pillow", "PIL"),
        ("python-pptx", "pptx"),
        ("PyMuPDF", "fitz"),
    ]
    for display, module_name in modules:
        try:
            __import__(module_name)
            status_line(display, True)
        except Exception as exc:
            failed = True
            status_line(display, False, str(exc))

    try:
        from native_drop import NativeFileDrop  # noqa: F401

        status_line("原生拖拽模块", True if sys.platform.startswith("win") else True, "Windows 下自动启用" if not sys.platform.startswith("win") else "可加载")
    except Exception as exc:
        failed = True
        status_line("原生拖拽模块", False, str(exc))

    try:
        from converter_core import detect_export_backends

        result = detect_export_backends()
        print(f"Microsoft PowerPoint：{'已检测到' if result['powerpoint'] else '未检测到'}")
        print(f"LibreOffice：{'已检测到' if result['libreoffice'] else '未检测到'}")
        if not result["powerpoint"] and not result["libreoffice"]:
            print("提示：图片合成 PPT 仍可使用；PPT 导出图片需要安装 PowerPoint 或 LibreOffice。")
    except Exception as exc:
        failed = True
        status_line("导出引擎检测", False, str(exc))

    icon_ok = (Path(__file__).parent / "app_icon.png").exists()
    status_line("应用图标", icon_ok)
    failed = failed or not icon_ok

    print("\n自检完成。")
    if sys.stdin.isatty():
        input("按回车键退出...")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
