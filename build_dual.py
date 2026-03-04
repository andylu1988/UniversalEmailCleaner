"""
UniversalEmailCleaner - 双版本构建脚本
======================================
自动构建两个版本:
  1. 带 License 的版本 (需要激活码)
  2. 不带 License 的版本 (名称后缀 _nolicense)

用法:
    python build_dual.py                          # 使用当前 spec 中的版本号
    python build_dual.py --version v1.13.0        # 指定版本号 (自动生成 spec)
    python build_dual.py --keygen                 # 同时构建密钥生成器
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

SPEC_TEMPLATE_NOLICENSE = r"""# -*- mode: python ; coding: utf-8 -*-
# No-License 版本 — 不包含 license_manager，启动无需激活


a = Analysis(
    ['universal_email_cleaner.py'],
    pathex=[],
    binaries=[],
    datas=[('graph-mail-delete.ico', '.'), ('avatar_b64.txt', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes=['license_manager'],
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
    name='UniversalEmailCleaner_{version}_nolicense',
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

SPEC_TEMPLATE_KEYGEN = r"""# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['license_keygen.py'],
    pathex=[],
    binaries=[],
    datas=[('license_manager.py', '.')],
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
    name='UniversalEmailCleaner_KeyGen',
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
    parser = argparse.ArgumentParser(description='UniversalEmailCleaner 双版本构建工具')
    parser.add_argument('--version', '-v', help='版本号 (例: v1.13.0)，不指定则从源码读取')
    parser.add_argument('--keygen', '-k', action='store_true', help='同时构建密钥生成器')
    parser.add_argument('--license-only', action='store_true', help='仅构建带 License 版本')
    parser.add_argument('--nolicense-only', action='store_true', help='仅构建不带 License 版本')
    args = parser.parse_args()

    version = args.version or get_version_from_source()
    print(f"版本号: {version}")

    # 生成 spec 文件
    spec_license = os.path.join(SCRIPT_DIR, f'UniversalEmailCleaner_{version}.spec')
    spec_nolicense = os.path.join(SCRIPT_DIR, f'UniversalEmailCleaner_{version}_nolicense.spec')
    spec_keygen = os.path.join(SCRIPT_DIR, 'UniversalEmailCleaner_KeyGen.spec')

    build_license = not args.nolicense_only
    build_nolicense = not args.license_only

    if build_license:
        with open(spec_license, 'w', encoding='utf-8') as f:
            f.write(SPEC_TEMPLATE_LICENSE.format(version=version))

    if build_nolicense:
        with open(spec_nolicense, 'w', encoding='utf-8') as f:
            f.write(SPEC_TEMPLATE_NOLICENSE.format(version=version))

    if args.keygen:
        with open(spec_keygen, 'w', encoding='utf-8') as f:
            f.write(SPEC_TEMPLATE_KEYGEN.format())

    results = []

    # 构建
    if build_license:
        ok = run_pyinstaller(spec_license, f'带 License 版 ({version})')
        results.append((f'UniversalEmailCleaner_{version}.exe', ok))

    if build_nolicense:
        ok = run_pyinstaller(spec_nolicense, f'无 License 版 ({version}_nolicense)')
        results.append((f'UniversalEmailCleaner_{version}_nolicense.exe', ok))

    if args.keygen:
        ok = run_pyinstaller(spec_keygen, 'KeyGen 密钥生成器')
        results.append(('UniversalEmailCleaner_KeyGen.exe', ok))

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
