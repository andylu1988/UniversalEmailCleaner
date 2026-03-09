"""
UniversalEmailCleaner - 发布构建脚本
===================================
仅构建当前正式发布版本。

用法:
        python build_dual.py                   # 使用源码中的版本号
        python build_dual.py --version v1.14.5 # 指定版本号并自动生成 spec
"""

import argparse
import os
import re
import subprocess
import sys
import shutil

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

SPEC_TEMPLATE_LICENSE = r"""# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['universal_email_cleaner.py'],
    pathex=[],
    binaries=[],
    datas=[('graph-mail-delete.ico', '.'), ('avatar_b64.txt', '.'), ('license_manager.py', '.')],
    hiddenimports=['license_manager'],
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='UniversalEmailCleaner_{version}',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['graph-mail-delete.ico'],
)
"""

def get_version_from_source() -> str:
    """从 universal_email_cleaner.py 中提取 APP_VERSION"""
    src = os.path.join(SCRIPT_DIR, 'universal_email_cleaner.py')
    with open(src, 'r', encoding='utf-8') as f:
        for line in f:
            m = re.match(r'^APP_VERSION\s*=\s*["\'](.+?)["\']', line)
            if m:
                return m.group(1)
    return 'v0.0.0'


def run_pyinstaller(spec_path: str, label: str) -> bool:
    """运行 PyInstaller 构建"""
    print(f"\n{'='*60}")
    print(f"  构建: {label}")
    print(f"  Spec: {os.path.basename(spec_path)}")
    print(f"{'='*60}\n")

    cmd = [sys.executable, '-m', 'PyInstaller', '--clean', spec_path]
    result = subprocess.run(cmd, cwd=SCRIPT_DIR)

    if result.returncode == 0:
        print(f"\n  [OK] {label} 构建成功!")
        return True
    else:
        print(f"\n  [FAIL] {label} 构建失败 (exit code: {result.returncode})")
        return False


def main():
    parser = argparse.ArgumentParser(description='UniversalEmailCleaner 发布构建工具')
    parser.add_argument('--version', '-v', help='版本号 (例: v1.13.0)，不指定则从源码读取')
    args = parser.parse_args()

    version = args.version or get_version_from_source()
    print(f"版本号: {version}")

    # 生成 spec 文件
    spec_license = os.path.join(SCRIPT_DIR, f'UniversalEmailCleaner_{version}.spec')
    with open(spec_license, 'w', encoding='utf-8') as f:
        f.write(SPEC_TEMPLATE_LICENSE.format(version=version))

    results = []

    ok = run_pyinstaller(spec_license, f'正式发布版 ({version})')
    results.append((f'UniversalEmailCleaner_{version}.exe', ok))

    # 汇总
    print(f"\n{'='*60}")
    print("  构建结果汇总")
    print(f"{'='*60}")

    dist_dir = os.path.join(SCRIPT_DIR, 'dist')
    for name, ok in results:
        path = os.path.join(dist_dir, name)
        if ok and os.path.exists(path):
            size_mb = os.path.getsize(path) / (1024 * 1024)
            print(f"  [OK]   {name:<50} {size_mb:.1f} MB")
        else:
            print(f"  [FAIL] {name:<50} 构建失败")

    print(f"\n  输出目录: {dist_dir}")

    all_ok = all(ok for _, ok in results)
    sys.exit(0 if all_ok else 1)


if __name__ == '__main__':
    main()
