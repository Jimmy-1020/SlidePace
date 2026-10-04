"""Offline build for the native .NET Framework PowerPoint add-in and installer."""
from pathlib import Path
import os
import shutil
import subprocess
import sys

ROOT = Path(__file__).resolve().parent.parent

def find_msbuild():
    configured = os.environ.get('SLIDEPACE_MSBUILD')
    if configured and Path(configured).is_file():
        return Path(configured)
    for program_files in ('ProgramFiles(x86)', 'ProgramFiles'):
        directory = os.environ.get(program_files)
        if not directory:
            continue
        vswhere = Path(directory) / 'Microsoft Visual Studio/Installer/vswhere.exe'
        if vswhere.is_file():
            result = subprocess.run([str(vswhere), '-latest', '-products', '*',
                                     '-requires', 'Microsoft.Component.MSBuild',
                                     '-find', r'MSBuild\**\Bin\MSBuild.exe'],
                                    capture_output=True, text=True, encoding='utf-8')
            for line in result.stdout.splitlines():
                if Path(line.strip()).is_file():
                    return Path(line.strip())
    available = shutil.which('MSBuild.exe')
    if available:
        return Path(available)
    legacy = Path(os.environ.get('WINDIR', 'C:/Windows')) / 'Microsoft.NET/Framework64/v4.0.30319/MSBuild.exe'
    return legacy if legacy.is_file() else None

def main():
    msbuild = find_msbuild()
    if msbuild is None:
        raise SystemExit('MSBuild not found. Install Visual Studio with .NET desktop development.')
    projects = ['src/SlidePace.AddIn/SlidePace.AddIn.csproj']
    for optional in ['src/SlidePace.Setup/SlidePace.Setup.csproj', 'tests/SlidePace.Tests.csproj']:
        if (ROOT / optional).exists():
            projects.append(optional)
    for project in projects:
        args = [str(msbuild), str(ROOT / project), '/t:Build', '/p:Configuration=Release', '/nologo', '/v:minimal']
        completed = subprocess.run(args, cwd=str(ROOT), stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        try:
            output = completed.stdout.decode('utf-8')
        except UnicodeDecodeError:
            output = completed.stdout.decode('gbk', errors='replace')
        print(output)
        if completed.returncode:
            return completed.returncode
    print('Build completed.')
    return 0

if __name__ == '__main__':
    sys.exit(main())
