"""Collect the release EXE, user documents, source archive, and checksums."""
from pathlib import Path
import hashlib
import shutil
import zipfile

root = Path(__file__).resolve().parent.parent
dist = root / 'dist'
dist.mkdir(exist_ok=True)
source_exe = root / 'src/SlidePace.Setup/bin/Release/SlidePace-Setup.exe'
if not source_exe.exists():
    raise SystemExit('Build Release before packaging.')
shutil.copy2(source_exe, dist / 'SlidePace-Setup.exe')
documents = [('README.md', '使用说明.md'), ('验证报告.md', '验证报告.md'),
             ('SlidePace-PowerPoint插件需求文档.md', '需求文档.md'),
             ('GitHub发布说明.md', 'GitHub发布说明.md')]
evidence = ['office/powerpoint-overlay.png',
            'host-load-visibility/fullscreen-overlay-1.png',
            'host-load-visibility/fullscreen-overlay-2.png',
            'setup-ui/installer.png',
            'v1.0.1/ui/overlay-overtime-compact.png',
            'v1.0.1/ui/overlay-overtime.png',
            'v1.0.2/ui/overlay-compact.png',
            'v1.0.2/ui/overlay-overtime-compact.png',
            'v1.0.2/ui/overlay-overtime.png',
            'v1.0.2/office/powerpoint-overlay.png',
            'v1.0.2/office/audience-overlay.png',
            'v1.0.2/office/presenter-overlay.png',
            'v1.0.3/ui/settings-compact.png',
            'v1.0.3/ui/overlay-overtime.png',
            'v1.0.3/ui/overlay-overtime-compact.png',
            'v1.0.3/office/preselected-windowed.png',
            'v1.0.3/office/powerpoint-stopped.png',
            'v1.0.3/office/audience-stopped.png',
            'v1.0.3/office/presenter-stopped.png',
            'v1.0.3/setup-ui/installer.png']
for source, target in documents:
    content = (root / source).read_text(encoding='utf-8')
    content = content.replace('./SlidePace-PowerPoint插件需求文档.md', './需求文档.md')
    content = content.replace('./README.md', './使用说明.md')
    content = content.replace('./docs/validation/v1.0.4/logs', './验证日志/v1.0.4')
    for relative in evidence:
        content = content.replace('./docs/validation/' + relative,
                                  './验证截图/' + relative)
    (dist / target).write_text(content, encoding='utf-8')
screenshots = dist / '验证截图'
screenshots.mkdir(exist_ok=True)
for relative in evidence:
    target = screenshots / relative
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(root / 'docs/validation' / relative, target)
    # Earlier packages flattened paths, causing different versions to collide.
    legacy = screenshots / Path(relative).name
    if legacy.is_file():
        legacy.unlink()

validation_logs = root / 'docs/validation/v1.0.4/logs'
for path in sorted(validation_logs.glob('*.txt')):
    target = dist / '验证日志/v1.0.4' / path.name
    target.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(path, target)

with zipfile.ZipFile(dist / 'SlidePace-Source.zip', 'w', compression=zipfile.ZIP_DEFLATED) as archive:
    folders = [root / 'src', root / 'tests', root / 'tools', root / 'assets', root / 'docs']
    for folder in folders:
        for path in sorted(folder.rglob('*')):
            if not path.is_file() or any(part in {'bin', 'obj', '__pycache__'} for part in path.relative_to(root).parts):
                continue
            archive.write(path, path.relative_to(root).as_posix())
    for name in ['SlidePace.sln', 'build.cmd', 'README.md', '验证报告.md', 'SlidePace-PowerPoint插件需求文档.md', 'GitHub发布说明.md', '.gitignore', '.gitattributes']:
        path = root / name
        if path.exists():
            archive.write(path, name)

lines = []
for path in sorted(dist.rglob('*')):
    if not path.is_file() or path.name == 'SHA256SUMS.txt':
        continue
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    lines.append(f'{digest}  {path.relative_to(dist).as_posix()}')
(dist / 'SHA256SUMS.txt').write_text('\n'.join(lines) + '\n', encoding='utf-8')
for path in sorted(dist.iterdir()):
    if path.is_file():
        print(f'{path.name}: {path.stat().st_size:,} bytes')
