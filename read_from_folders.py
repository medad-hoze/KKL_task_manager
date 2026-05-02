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
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

# --- file classification ----------------------------------------------------

SCRIPT_EXTENSIONS = {
    '.py', '.js', '.ts', '.tsx', '.jsx', '.ipynb',
    '.sh', '.bat', '.ps1', '.cmd',
    '.r', '.rb', '.php', '.go', '.rs', '.java', '.c', '.cpp', '.cs',
    '.sql', '.html', '.css', '.scss', '.vue', '.svelte',
}

DOCUMENT_EXTENSIONS = {
    '.pdf', '.docx', '.doc', '.xlsx', '.xls', '.pptx', '.ppt',
    '.txt', '.md', '.rtf', '.odt', '.ods',
    '.csv', '.json', '.xml', '.yaml', '.yml',
}

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


def classify(path: Path) -> str:
    ext = path.suffix.lower()
    if ext in SCRIPT_EXTENSIONS:
        return 'scripts'
    if ext in DOCUMENT_EXTENSIONS:
        return 'documents'
    return 'other'


def file_info(path: Path, root: Path) -> dict:
    stat = path.stat()
    return {
        'name': path.name,
        'path': str(path.relative_to(root)).replace('\\', '/'),
        'extension': path.suffix.lower(),
        'size_bytes': stat.st_size,
        'modified': datetime.fromtimestamp(stat.st_mtime).isoformat(timespec='seconds'),
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
            yield from iter_files(entry)
        elif entry.is_file():
            yield entry


def scan_project(project_dir: Path, root: Path) -> dict:
    grouped = {'scripts': [], 'documents': [], 'other': []}
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
