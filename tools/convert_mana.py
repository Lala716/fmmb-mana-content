#!/usr/bin/env python3
"""
Convertit un Word "Mana isan'andro" (6 mois) en JSON prêt à héberger
pour l'app FMMB. Appelle pandoc automatiquement.

Usage:
    python3 convert_mana.py MonFichier.docx --bible Bible.json \\
        --taona 2027 --laharana 1 --titre "Enim-bolana Voalohany — 2027" \\
        --out mana-2027-s1.json

Voir : python3 convert_mana.py --help
"""
import re, json, sys, os, shutil, subprocess, argparse, unicodedata
from datetime import date

MOIS_MG = {
    'JANOARY': 1, 'FEBROARY': 2, 'MARTSA': 3, 'APRILY': 4, 'MEY': 5, 'JONA': 6,
    'JOLAY': 7, 'AOGOSITRA': 8, 'SEPTAMBRA': 9, 'OKTOBRA': 10, 'NOVAMBRA': 11, 'DESAMBRA': 12
}

# Alias pour les variantes/coquilles de noms de livres rencontrées dans le
# document Word, mappées vers le long_name EXACT utilisé dans Bible.json.
BOOK_ALIASES = {
    'APOKALYPSY': 'APOKALIPSY',       # coquille -> orthographe de Bible.json
    'DEOTERONOMIA': 'DEOTORONOMIA',   # orthographe Bible.json (transposition)
}

DAY_RE = re.compile(r"^#\s+([A-ZÀ-Ÿ]+)\s+(\d{1,2})\s+([A-ZÀ-Ÿ]+)(.*)$")
REF_RE = re.compile(
    r"^(?P<book>.+?)\s+(?P<c1>\d+)\.\s*(?P<v1>\d+)"
    r"(?:\s*-{1,2}\s*(?:(?P<c2>\d+)\.)?(?P<v2>\d+))?\s*$"
)
# Livres à un seul chapitre (ex: OBADIA) : "OBADIA 1-15" au lieu de "OBADIA 1.1-15"
REF_RE_1CHAP = re.compile(
    r"^(?P<book>.+?)\s+(?P<v1>\d+)(?:\s*-{1,2}\s*(?P<v2>\d+))?\s*$"
)
FQ_RE  = re.compile(r"^(.*?)\s+F\.\s*([\d,\s]+)$")
CITATION_RE = re.compile(r"^\*{2,3}([^*]+?)\s*:\*{2,3}\s*(.*)$")


def norm(s: str) -> str:
    s = unicodedata.normalize('NFKD', s).upper()
    return ''.join(c for c in s if c.isalnum())


def md_inline_to_html(text: str) -> str:
    text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    text = re.sub(r'\*\*\*(.+?)\*\*\*', r'<strong><em>\1</em></strong>', text)
    text = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', text)
    text = re.sub(r'\*(.+?)\*', r'<em>\1</em>', text)
    return text.strip()


def lines_to_html(lines):
    """Groupe les lignes en paragraphes / listes à puces et convertit en HTML."""
    paras, buf = [], []
    for l in lines + ['']:
        if l.strip() == '':
            if buf:
                paras.append(' '.join(buf).strip())
                buf = []
        else:
            buf.append(l.strip())

    html_parts, in_list = [], False
    for p in paras:
        is_item = p.startswith('-   ') or p.startswith('- ')
        content = md_inline_to_html(re.sub(r'^-+\s+', '', p)) if is_item else md_inline_to_html(p)
        if is_item:
            if not in_list:
                html_parts.append('<ul>'); in_list = True
            html_parts.append(f'<li>{content}</li>')
        else:
            if in_list:
                html_parts.append('</ul>'); in_list = False
            html_parts.append(f'<p>{content}</p>')
    if in_list:
        html_parts.append('</ul>')
    return ''.join(html_parts)


def parse_reference(ref_text, books_index):
    ref_text = ref_text.replace('*', '').replace('\u2003', ' ').strip()
    fq = FQ_RE.match(ref_text)
    if not fq:
        return None
    ref_part, nums_part = fq.groups()
    fanontaniana_ref = [int(n) for n in re.findall(r'\d+', nums_part)]

    m = REF_RE.match(ref_part.strip())
    if m:
        book_raw = m.group('book').strip()
        c1, v1 = int(m.group('c1')), int(m.group('v1'))
        c2 = int(m.group('c2')) if m.group('c2') else c1
        v2 = int(m.group('v2')) if m.group('v2') else v1
    else:
        m1 = REF_RE_1CHAP.match(ref_part.strip())
        if not m1:
            return None
        book_raw = m1.group('book').strip()
        c1 = c2 = 1
        v1 = int(m1.group('v1'))
        v2 = int(m1.group('v2')) if m1.group('v2') else v1

    key = norm(book_raw)
    key = BOOK_ALIASES.get(key, key)
    book = books_index.get(key)
    if not book:
        return None

    if c1 == c2:
        display = f"{book['long_name']} {c1}:{v1}" + (f"-{v2}" if v2 != v1 else "")
    else:
        display = f"{book['long_name']} {c1}:{v1}-{c2}:{v2}"

    return {
        'display': display,
        'book_number': book['book_number'],
        'chapter_start': c1, 'verse_start': v1,
        'chapter_end': c2, 'verse_end': v2,
        'fanontaniana_ref': fanontaniana_ref,
    }


def main(md_path, bible_path, taona, laharana, lohateny_lehibe, description, out_path):
    with open(bible_path, encoding='utf-8') as f:
        bible = json.load(f)
    books_index = {norm(b['long_name']): b for b in bible['books']}

    with open(md_path, encoding='utf-8') as f:
        lines = f.read().split('\n')

    heading_idx = [i for i, l in enumerate(lines) if l.startswith('# ')]
    heading_idx.append(len(lines))

    intro_extra_html = ''
    jours = []
    warnings = []

    for bi in range(len(heading_idx) - 1):
        start, end = heading_idx[bi], heading_idx[bi + 1]
        block = lines[start:end]
        raw_heading = block[0][2:].strip()
        clean_heading = re.sub(r'[*_]', '', raw_heading)

        dm = DAY_RE.match('# ' + clean_heading) or DAY_RE.match(clean_heading)
        # DAY_RE attend un "# " en tête -> on le simule
        dm = DAY_RE.match('# ' + clean_heading)

        if not dm:
            # Bloc d'introduction (ex: "Fampidirana ny Bokin'ny Apokalypsy")
            intro_extra_html = lines_to_html([l for l in block[1:] if not l.startswith('>')])
            continue

        weekday_mg, day_num, month_mg, heading_tail = dm.groups()
        day_num = int(day_num)
        month_num = MOIS_MG.get(month_mg)
        if not month_num:
            warnings.append(f"Mois inconnu: {month_mg} ({raw_heading})")
            continue
        daty = date(taona, month_num, day_num).isoformat()

        heading_tail = heading_tail.strip()
        body_lines = block[1:]

        if heading_tail:
            # La référence est collée dans le heading -> la ligne suivante est direct le titre
            ref_info = parse_reference(heading_tail, books_index)
            idx = 0
            while idx < len(body_lines) and body_lines[idx].strip() == '':
                idx += 1
            title = re.sub(r'[*_]', '', body_lines[idx]).strip() if idx < len(body_lines) else ''
            rest = body_lines[idx + 1:]
        else:
            idx = 0
            while idx < len(body_lines) and body_lines[idx].strip() == '':
                idx += 1
            refline = body_lines[idx] if idx < len(body_lines) else ''
            ref_info = parse_reference(refline, books_index)
            idx += 1
            while idx < len(body_lines) and body_lines[idx].strip() == '':
                idx += 1
            title = re.sub(r'[*_]', '', body_lines[idx]).strip() if idx < len(body_lines) else ''
            rest = body_lines[idx + 1:]

        if not ref_info:
            warnings.append(f"Référence non résolue pour {daty} ({raw_heading!r})")
            continue

        # Sépare corps / citation finale (dernier groupe de lignes '> ...')
        last_quote_start = None
        for i, l in enumerate(rest):
            if l.startswith('>'):
                if last_quote_start is None:
                    last_quote_start = i
            elif l.strip() != '':
                last_quote_start = None
        body_part = rest[:last_quote_start] if last_quote_start is not None else rest
        quote_part = rest[last_quote_start:] if last_quote_start is not None else []

        body_html = lines_to_html(body_part)
        if bi == 0 or (intro_extra_html and len(jours) == 0):
            pass  # (intro déjà gérée séparément, fusionnée plus bas si besoin)

        citation_label, citation_text = None, None
        if quote_part:
            quote_text = ' '.join(l.lstrip('>').strip() for l in quote_part if l.strip() != '')
            cm = CITATION_RE.match(quote_text)
            if cm:
                citation_label = cm.group(1).strip()
                citation_text = re.sub(r'\*+', '', cm.group(2)).strip()
            else:
                citation_text = re.sub(r'\*+', '', quote_text).strip()

        jour = {
            'daty': daty,
            'toko_sy_andininy': ref_info['display'],
            'ref': {
                'book_number': ref_info['book_number'],
                'chapter_start': ref_info['chapter_start'],
                'verse_start': ref_info['verse_start'],
                'chapter_end': ref_info['chapter_end'],
                'verse_end': ref_info['verse_end'],
            },
            'fanontaniana_ref': ref_info['fanontaniana_ref'],
            'lohateny_mana': title,
            'famelabelarana_mana': body_html,
            'citation_label': citation_label,
            'citation_text': citation_text,
        }
        jours.append(jour)

    # Fusionne le texte d'introduction (ex: intro du livre de l'Apokalipsy)
    # au tout début du corps du 1er jour concerné (le tout 1er jour rencontré
    # juste après le bloc d'intro dans le document original).
    if intro_extra_html:
        for j in jours:
            if 'Apokalipsy' in j['toko_sy_andininy'] or 'APOKAL' in j['toko_sy_andininy'].upper():
                j['famelabelarana_mana'] = intro_extra_html + j['famelabelarana_mana']
                break

    result = {
        'enimbolana': {
            'taona': taona,
            'laharana': laharana,
            'lohateny_lehibe': lohateny_lehibe,
            'description': description,
        },
        'andro': jours,
    }

    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"OK : {len(jours)} jours écrits dans {out_path}")
    if warnings:
        print(f"\n⚠️  {len(warnings)} avertissement(s) :")
        for w in warnings:
            print(' -', w)
    dates = [j['daty'] for j in jours]
    dupes = set(d for d in dates if dates.count(d) > 1)
    if dupes:
        print("⚠️  Dates en double:", dupes)
    print(f"Plage: {min(dates)} -> {max(dates)}")


def docx_to_markdown(docx_path: str) -> str:
    """Convertit le .docx en markdown via pandoc (doit être installé)."""
    if shutil.which('pandoc') is None:
        sys.exit(
            "❌ 'pandoc' n'est pas installé ou pas dans le PATH.\n"
            "   Windows : winget install --id JohnMacFarlane.Pandoc\n"
            "   Mac     : brew install pandoc\n"
            "   Linux   : sudo apt install pandoc"
        )
    md_path = os.path.basename(os.path.splitext(docx_path)[0]) + '.md'
    subprocess.run(['pandoc', '-t', 'markdown', docx_path, '-o', md_path], check=True)
    return md_path


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description="Convertit un Word 'Mana isan'andro' en JSON prêt à héberger.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""Exemple :
  python3 convert_mana.py Mana_Voalohany_2027.docx \\
      --bible ../fihiranaFMMB/src/assets/databases/Bible.json \\
      --taona 2027 --laharana 1 \\
      --titre "Enim-bolana Voalohany — 2027" \\
      --description "Mana isan'andro ho an'ny enim-bolana voalohany taona 2027." \\
      --out mana-2027-s1.json"""
    )
    parser.add_argument('docx', help="Chemin du fichier Word (.docx) du Mana isan'andro")
    parser.add_argument('--bible', required=True, help="Chemin vers Bible.json (dans assets/databases/)")
    parser.add_argument('--taona', type=int, required=True, help="Année, ex: 2027")
    parser.add_argument('--laharana', type=int, required=True, choices=[1, 2],
                         help="1 = enim-bolana voalohany, 2 = enim-bolana faharoa")
    parser.add_argument('--titre', required=True, dest='lohateny_lehibe',
                         help="Titre affiché de la période, ex: 'Enim-bolana Voalohany — 2027'")
    parser.add_argument('--description', default='')
    parser.add_argument('--out', required=True, help="Fichier JSON de sortie")
    args = parser.parse_args()

    md_path = docx_to_markdown(args.docx)

    main(
        md_path=md_path,
        bible_path=args.bible,
        taona=args.taona,
        laharana=args.laharana,
        lohateny_lehibe=args.lohateny_lehibe,
        description=args.description,
        out_path=args.out
    )