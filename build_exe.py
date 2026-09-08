"""build_exe.py - 用 PyInstaller 打包为单文件 EXE

运行：
    pip install -r requirements.txt
    python build_exe.py

输出：
    dist/价格牌处理工具.exe  （双击直接运行，无需 Python 环境）

参数说明：
    --onefile     : 打包成单个 exe（启动稍慢，但分发简单）
    --windowed    : 不弹黑色控制台窗口
    --add-data    : 把 samples 目录打包进 exe（相对路径解析需要）
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path

import PyInstaller.__main__


HERE = Path(__file__).parent.resolve()
APP_NAME = "价格牌处理工具"


def build() -> None:
    # 清理旧产物（环境的安全删除 shim 可能拦截删除，忽略失败）
    for d in ["build", "dist"]:
        p = HERE / d
        try:
            if p.exists():
                shutil.rmtree(p, ignore_errors=True)
        except OSError:
            pass
    spec = HERE / f"{APP_NAME}.spec"
    try:
        if spec.exists():
            spec.unlink()
    except OSError:
        pass

    # 打包参数
    args = [
        str(HERE / "main.py"),
        f"--name={APP_NAME}",
        "--onefile",
        "--windowed",
        "--noconfirm",
        "--clean",
        "--noupx",  # 不压缩（部分杀软误报）
        f"--icon={HERE / 'assets' / 'logo.ico'}",
        f"--add-data={HERE / 'samples'}{os.pathsep}samples",
        f"--add-data={HERE / 'assets'}{os.pathsep}assets",
        # 显式包含子模块，确保 PyInstaller 能识别
        "--collect-submodules=core",
        "--hidden-import=fitz",
        "--hidden-import=PIL",
        "--hidden-import=openpyxl",
        # 排除系统 Python 中损坏/无关的第三方库，避免 PyInstaller 触发其 hook 而崩溃
        "--exclude-module=matplotlib",
        "--exclude-module=numpy",
    ]

    print("==> 打包参数：", args)
    PyInstaller.__main__.run(args)

    exe_path = HERE / "dist" / f"{APP_NAME}.exe"
    if exe_path.exists():
        size_mb = exe_path.stat().st_size / 1024 / 1024
        print(f"\n✅ 打包完成：{exe_path}  ({size_mb:.1f} MB)")
    else:
        print("\n❌ 打包失败：未找到 dist/*.exe")
        sys.exit(1)


if __name__ == "__main__":
    build()