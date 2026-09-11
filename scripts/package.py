"""Build a native installer on its target OS with JDK 21 and Maven available."""
import os
from pathlib import Path
import platform
import plistlib
import shutil
import subprocess
import time
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
os.chdir(ROOT)
VERSION = ET.parse('pom.xml').find('{http://maven.apache.org/POM/4.0.0}version').text
SYSTEM = platform.system()
if SYSTEM not in ('Darwin', 'Windows'):
    raise SystemExit('Run on macOS or Windows to build its native installer.')


def run(*args):
    subprocess.run([str(arg) for arg in args], check=True)


def tool(name):
    java_home = os.environ.get('JAVA_HOME')
    candidate = Path(java_home) / 'bin' / (name + ('.exe' if SYSTEM == 'Windows' else '')) if java_home else None
    return str(candidate) if candidate and candidate.exists() else name


run(os.environ.get('MAVEN', 'mvn.cmd' if SYSTEM == 'Windows' else 'mvn'),
    '-B', 'clean', 'package', 'org.apache.maven.plugins:maven-dependency-plugin:3.8.1:copy-dependencies',
    '-DincludeScope=runtime', '-DoutputDirectory=target/installer-input')
run(tool('java'), '-cp', os.pathsep.join(['target/test-classes', 'target/classes']),
    'cz.leaguesheets.XlsxPrintLayoutTest')
input_dir = ROOT / 'target/installer-input'
jar = f'league-sheets-{VERSION}.jar'
shutil.copy2(ROOT / 'target' / jar, input_dir / jar)
image_dir = ROOT / 'target/app-image'
dist = ROOT / 'target/installers'
dist.mkdir(parents=True, exist_ok=True)
args = [tool('jpackage'), '--type', 'app-image', '--dest', image_dir,
        '--input', input_dir, '--name', 'LeagueSheets', '--app-version', VERSION,
        '--main-jar', jar, '--main-class', 'cz.leaguesheets.Main',
        '--vendor', 'LeagueSheets', '--description', 'Priprava ligovych zapisu',
        '--add-modules', 'java.se,jdk.unsupported,jdk.crypto.ec']
if SYSTEM == 'Windows':
    args += ['--icon', ROOT / 'src/main/resources/cz/leaguesheets/bowling-icon.ico']
else:
    args += ['--mac-package-identifier', 'cz.leaguesheets.app',
             '--icon', ROOT / 'src/main/resources/cz/leaguesheets/bowling-icon.icns']
run(*args)
app = image_dir / ('LeagueSheets.app' if SYSTEM == 'Darwin' else 'LeagueSheets')
if SYSTEM == 'Darwin':
    plist = app / 'Contents/Info.plist'
    data = plistlib.loads(plist.read_bytes())
    data['NSAppleEventsUsageDescription'] = 'LeagueSheets ovládá Microsoft Excel pro tisk ligových zápisů a export PDF.'
    plist.write_bytes(plistlib.dumps(data))
    run('codesign', '--force', '--deep', '--sign', '-', app)

# Launch the packaged application, including its bundled Java and JavaFX.
launcher = app / ('Contents/MacOS/LeagueSheets' if SYSTEM == 'Darwin' else 'LeagueSheets.exe')
with (ROOT / 'target/launcher-smoke.log').open('w') as log:
    process = subprocess.Popen([str(launcher)], stdout=log, stderr=subprocess.STDOUT)
    try:
        time.sleep(10)
        if process.poll() is not None:
            raise RuntimeError('Packaged app exited during startup; see target/launcher-smoke.log')
    finally:
        if process.poll() is None:
            process.terminate()
            try:
                process.wait(timeout=10)
            except subprocess.TimeoutExpired:
                process.kill()
                process.wait()

kind = 'dmg' if SYSTEM == 'Darwin' else 'exe'
args = [tool('jpackage'), '--type', kind, '--app-image', app, '--dest', dist,
        '--name', 'LeagueSheets', '--app-version', VERSION, '--vendor', 'LeagueSheets']
if SYSTEM == 'Windows':
    args += ['--win-per-user-install', '--win-dir-chooser', '--win-menu', '--win-shortcut',
             '--win-upgrade-uuid', '87673e27-01c4-41fa-a58e-227aa7feaa56']
run(*args)
arch = 'arm64' if platform.machine().lower() in ('arm64', 'aarch64') else 'x64'
installer = next(dist.glob(f'*.{kind}'))
installer.rename(dist / f'LeagueSheets-{VERSION}-{ "macOS" if SYSTEM == "Darwin" else "Windows"}-{arch}.{kind}')
print(f'Installer ready in {dist}')
