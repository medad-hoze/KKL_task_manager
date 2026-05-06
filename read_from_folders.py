"""
Project folder scanner — robust Hebrew/Unicode matching.

Structure:
    <data_folder>/<username>/<project_name>/...files...

Outputs:
    projects_files.json
    projects_files.js   (defines window.PROJECT_FILES, works on file://)

Usage:
    python scan_projects.py <data_folder> [output_base]
"""

import json
import re
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

# --- file classification ----------------------------------------------------
DIR_AS_FILE_EXTENSIONS = {'.gdb'}


SCRIPT_EXTENSIONS = {
    '.py', '.js', '.ts', '.tsx', '.jsx', '.ipynb',
    '.sh', '.bat', '.ps1', '.cmd',
    '.r', '.rb', '.php', '.go', '.rs', '.java', '.c', '.cpp', '.cs',
    '.sql', '.html', '.css', '.scss', '.vue', '.svelte',
}

DOCUMENT_EXTENSIONS = {
    '.pdf', '.docx', '.doc', '.xlsx', '.xls', '.pptx', '.ppt',
    '.txt', '.md', '.rtf', '.odt', '.ods',
    '.csv', '.json', '.yaml', '.yml',
    '.gdb',
}

# Auto-generated / low-value files we never want to surface.
IGNORED_EXTENSIONS = {
    # .NET service reference clutter
    '.xsd', '.xss', '.disco', '.discomap', '.wsdl',
    '.svcmap', '.svcinfo',
    # Lucene / ArcGIS search-index internals
    '.fdt', '.fdx', '.fnm', '.frq', '.nrm', '.prx', '.tii', '.tis', '.gen',
    # GIS / config sidecars
    '.xml', '.sde', '.rsd','.txt', '.lock','.pyproj','.config','.ini','.log','gitignore','.gitattributes',
    '.rdl','.rptproj','.data','.pyt','.gpkx','.js','.cs','.db','.ts','.png','.jpg','.aspx','.asmx','.css'}

# Filename stems (no extension) that are auto-generated index/lock files.
IGNORED_STEM_RE = re.compile(r'^segments(_\d+)?$', re.IGNORECASE)

# Filenames whose stem is just a number (e.g. "1234.jpg", "-720784746.jpg") —
# these are almost always cache / thumbnail / hash-named files, not content.
NUMERIC_NAME_RE = re.compile(r'^-?\d+$')

SKIP_DIRS = {'.git', '.vscode', '.idea', '__pycache__', 'node_modules', '.venv', 'venv'}

# Characters to strip during normalization (in addition to unicode-category-based stripping):
# - whitespace, common separators, brackets
# - all forms of quote/apostrophe (regular, gershayim ״, geresh ׳, smart quotes)
DROP_CHARS = set(
    ' \t\n\r_-.,;:!?'
    '\'"`'
    '\u05F3\u05F4'  # Hebrew geresh ׳, gershayim ״
    '\u2018\u2019\u201C\u201D'  # smart quotes
    '\u2032\u2033'  # primes
    '\\/(){}[]<>'
)


# --- project-as-single-file detection --------------------------------------

def is_angular_project(path: Path) -> bool:
    """Directory is an Angular workspace (has angular.json at its root)."""
    return path.is_dir() and (path / 'angular.json').is_file()


def is_csharp_project(path: Path) -> bool:
    """Directory is a C#/.NET project or solution (has .csproj or .sln at root)."""
    if not path.is_dir():
        return False
    try:
        for child in path.iterdir():
            if child.is_file() and child.suffix.lower() in ('.csproj', '.sln'):
                return True
    except OSError:
        return False
    return False


def is_collapsed_dir(path: Path) -> bool:
    """Any directory that should be treated as a single file entry."""
    if not path.is_dir():
        return False
    if path.suffix.lower() in DIR_AS_FILE_EXTENSIONS:
        return True
    return is_angular_project(path) or is_csharp_project(path)


def is_ignored_file(path: Path) -> bool:
    """File-level filters: extension blacklist + numeric-only stems + index stems."""
    if path.suffix.lower() in IGNORED_EXTENSIONS:
        return True
    if NUMERIC_NAME_RE.match(path.stem):
        return True
    if IGNORED_STEM_RE.match(path.stem):
        return True
    return False


def classify(path: Path) -> str:
    if path.is_dir() and (is_angular_project(path) or is_csharp_project(path)):
        return 'scripts'
    ext = path.suffix.lower()
    if ext in SCRIPT_EXTENSIONS:
        return 'scripts'
    if ext in DOCUMENT_EXTENSIONS:
        return 'documents'
    return 'other'


def file_info(path: Path, root: Path) -> dict:
    if is_collapsed_dir(path):
        size = 0
        latest = path.stat().st_mtime
        for p in path.rglob('*'):
            try:
                if p.is_file():
                    s = p.stat()
                    size += s.st_size
                    if s.st_mtime > latest:
                        latest = s.st_mtime
            except OSError:
                pass
        mtime = latest
    else:
        s = path.stat()
        size = s.st_size
        mtime = s.st_mtime

    return {
        'name': path.name,
        'path': str(path.relative_to(root)).replace('\\', '/'),
        'extension': path.suffix.lower(),
        'size_bytes': size,
        'modified': datetime.fromtimestamp(mtime).isoformat(timespec='seconds'),
    }


def normalize_name(name: str) -> str:
    """Aggressive normalization for cross-source matching:
    - NFC unicode normalization (collapses NFC vs NFD)
    - Strip format chars (Cf): RTL/LTR marks, zero-width, BOM, etc.
    - Strip combining marks (Mn)
    - Lowercase
    - Drop whitespace, separators, all quote variants, brackets, punctuation
    """
    if not name:
        return ''
    name = unicodedata.normalize('NFC', name)
    # Drop invisible/format chars and combining marks
    name = ''.join(c for c in name if unicodedata.category(c) not in ('Cf', 'Mn', 'Cc'))
    name = name.lower()
    return ''.join(c for c in name if c not in DROP_CHARS)


# --- scanning ---------------------------------------------------------------

def iter_files(folder: Path):
    for entry in folder.iterdir():
        if entry.is_dir():
            if entry.name in SKIP_DIRS or entry.name.startswith('.'):
                continue
            if is_collapsed_dir(entry):
                yield entry          # whole directory becomes one entry
                continue
            yield from iter_files(entry)
        elif entry.is_file():
            if is_ignored_file(entry):
                continue
            yield entry


def scan_project(project_dir: Path, root: Path) -> dict:
    grouped = {'scripts': [], 'documents': [], 'other': []}
    if is_collapsed_dir(project_dir):
        # whole project is itself a collapsed unit (Angular / C# / .gdb)
        grouped[classify(project_dir)].append(file_info(project_dir, root))
    else:
        for p in iter_files(project_dir):
            grouped[classify(p)].append(file_info(p, root))

    for key in grouped:
        grouped[key].sort(key=lambda f: f['path'])

    stat = project_dir.stat()
    return {
        'path': str(project_dir.relative_to(root)).replace('\\', '/'),
        'modified': datetime.fromtimestamp(stat.st_mtime).isoformat(timespec='seconds'),
        'total_files': sum(len(v) for v in grouped.values()),
        **grouped,
    }


def scan_user(user_dir: Path, root: Path) -> dict:
    projects = {}
    for project_dir in sorted(user_dir.iterdir(), key=lambda p: p.name):
        if not project_dir.is_dir():
            continue
        if project_dir.name in SKIP_DIRS or project_dir.name.startswith('.'):
            continue
        projects[project_dir.name] = scan_project(project_dir, root)
    return {
        'path': str(user_dir.relative_to(root)).replace('\\', '/'),
        'project_count': len(projects),
        'projects': projects,
    }


def scan_data(data_dir) -> dict:
    root = Path(data_dir).resolve()
    if not root.is_dir():
        raise NotADirectoryError(f'{root} is not a directory')

    users = {}
    for user_dir in sorted(root.iterdir(), key=lambda p: p.name):
        if not user_dir.is_dir():
            continue
        if user_dir.name in SKIP_DIRS or user_dir.name.startswith('.'):
            continue
        users[user_dir.name] = scan_user(user_dir, root)

    by_project = {}
    duplicates = []
    for username, udata in users.items():
        for proj_name, pdata in udata['projects'].items():
            key = normalize_name(proj_name)
            if not key:
                continue
            if key in by_project:
                duplicates.append((proj_name, by_project[key]['user'], username))
            by_project[key] = {'user': username, 'original_name': proj_name, **pdata}

    return {
        'data_folder': str(root),
        'scanned_at': datetime.now().isoformat(timespec='seconds'),
        'user_count': len(users),
        'users': users,
        'by_project': by_project,
        '_duplicates': duplicates,
    }


# --- entry point ------------------------------------------------------------

def write_outputs(result: dict, output_base: str):
    json_path = f'{output_base}.json'
    js_path = f'{output_base}.js'

    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    with open(js_path, 'w', encoding='utf-8') as f:
        f.write('window.PROJECT_FILES = ')
        json.dump(result, f, ensure_ascii=False, indent=2)
        f.write(';\n')

    return json_path, js_path



data_path = r'C:\Users\Medad\OneDrive - Keren Kayemeth LeIsrael, Jewish National Fund\Desktop\KKL\ניהול פרוייקטים\data'
output_base = 'projects_files'



result = scan_data(data_path)
json_path, js_path = write_outputs(result, output_base)

print(f'Scanned: {result["data_folder"]}')
print(f'Users: {result["user_count"]}, Projects: {len(result["by_project"])}')
for user, data in result['users'].items():
    print(f'  {user}: {data["project_count"]} projects')
    # Print normalized keys so you can compare to what the dashboard logs
    for proj_name in list(data['projects'].keys())[:5]:
        print(f'    "{proj_name}"  →  "{normalize_name(proj_name)}"')
if result['_duplicates']:
    print(f'\n⚠ duplicate project names across users:')
    for name, u1, u2 in result['_duplicates']:
        print(f'  - "{name}" in both {u1} and {u2}')
print(f'\nWrote: {json_path}')
print(f'Wrote: {js_path}')