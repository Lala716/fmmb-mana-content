#!/usr/bin/env python3
"""
Convertit les deux Word "Fanilo" (Malagasy + Français) en un seul JSON
bilingue prêt à héberger pour l'app FMMB.

Usage:
    python3 convert_fanilo.py Fanilo_MG.docx Fanilo_FR.docx \
        --bible Bible.json --taona 2026 \
        --out fanilo-2026.json

Voir : python3 convert_fanilo.py --help
"""
import re, json, sys, os, shutil, subprocess, argparse, unicodedata
from datetime import date

# ═══════════════════════════════
# MOIS (MG + FR, avec variantes/coquilles rencontrées dans les Word)
# ═══════════════════════════════
MOIS_MG = {
    'JANOARY': 1, 'FEBROARY': 2, 'MARTSA': 3, 'MARSA': 3, 'MARS': 3, 'APRILY': 4,
    'MEY': 5, 'MAI': 5, 'JONA': 6, 'JOLAY': 7, 'AOGOSITRA': 8,
    'SEPTAMBRA': 9, 'SEPTAMBRE': 9, 'OKTOBRA': 10, 'NOVAMBRA': 11, 'DESAMBRA': 12
}
MOIS_FR = {
    'JANVIER': 1, 'FEVRIER': 2, 'FÉVRIER': 2, 'MARS': 3, 'AVRIL': 4, 'MAI': 5,
    'JUIN': 6, 'JUILLET': 7, 'AOUT': 8, 'AOÛT': 8, 'SEPTEMBRE': 9,
    'OCTOBRE': 10, 'NOVEMBRE': 11, 'DECEMBRE': 12, 'DÉCEMBRE': 12
}

# Correspondance livres : ordre canonique identique à Bible.json (book_number croissant)
FR_BOOK_NAMES_IN_ORDER = [
    "Genèse", "Exode", "Lévitique", "Nombres", "Deutéronome", "Josué", "Juges", "Ruth",
    "1 Samuel", "2 Samuel", "1 Rois", "2 Rois", "1 Chroniques", "2 Chroniques", "Esdras",
    "Néhémie", "Esther", "Job", "Psaumes", "Proverbes", "Ecclésiaste",
    "Cantique des Cantiques", "Ésaïe", "Jérémie", "Lamentations", "Ézéchiel", "Daniel",
    "Osée", "Joël", "Amos", "Abdias", "Jonas", "Michée", "Nahum", "Habacuc", "Sophonie",
    "Aggée", "Zacharie", "Malachie", "Matthieu", "Marc", "Luc", "Jean", "Actes",
    "Romains", "1 Corinthiens", "2 Corinthiens", "Galates", "Éphésiens", "Philippiens",
    "Colossiens", "1 Thessaloniciens", "2 Thessaloniciens", "1 Timothée", "2 Timothée",
    "Tite", "Philémon", "Hébreux", "Jacques", "1 Pierre", "2 Pierre", "1 Jean",
    "2 Jean", "3 Jean", "Jude", "Apocalypse"
]

# Noms anglais parfois laissés tels quels dans le Word "FR" + coquilles fréquentes du Word "MG"
BOOK_NAME_ALIASES = {
    # anglais -> français (le Word FR contient quelques restes non traduits)
    'GENESIS': 'GENÈSE', 'EXODUS': 'EXODE', 'LEVITICUS': 'LÉVITIQUE', 'NUMBERS': 'NOMBRES',
    'DEUTERONOMY': 'DEUTÉRONOME', 'JOSHUA': 'JOSUÉ', 'JUDGES': 'JUGES',
    '1SAMUEL': '1 SAMUEL', '2SAMUEL': '2 SAMUEL', '1KINGS': '1 ROIS', '2KINGS': '2 ROIS',
    'PSALM': 'PSAUMES', 'PSALMS': 'PSAUMES', 'PROVERBS': 'PROVERBES',
    'ISAIAH': 'ÉSAÏE', 'JEREMIAH': 'JÉRÉMIE', 'EZEKIEL': 'ÉZÉCHIEL', 'DANIEL': 'DANIEL',
    'MATTHEW': 'MATTHIEU', 'MARK': 'MARC', 'LUKE': 'LUC', 'JOHN': 'JEAN', 'ACTS': 'ACTES',
    'ROMANS': 'ROMAINS', 'GALATIANS': 'GALATES', 'EPHESIANS': 'ÉPHÉSIENS',
    'PHILIPPIANS': 'PHILIPPIENS', 'COLOSSIANS': 'COLOSSIENS', 'TITUS': 'TITE',
    'HEBREWS': 'HÉBREUX', 'JAMES': 'JACQUES', 'JUDE': 'JUDE', 'REVELATION': 'APOCALYPSE',
    # singulier / variantes FR
    'PSAUME': 'PSAUMES', 'ISAIE': 'ÉSAÏE', 'ISAÏE': 'ÉSAÏE',
    # coquilles MG rencontrées dans le Word
    'KSODOSY': 'EKSODOSY', 'EKOSODOSY': 'EKSODOSY',
    'ASA': "ASANNYAPOSTOLY",
}

# DAY_RE = re.compile(r"^(\d{1,2})(?:er)?\s+([A-Za-zÀ-ÿ]+)\.?\s*$")
DAY_RE = re.compile(
    r'^\s*(\d{1,2})\s*(?:er)?\s+([A-Za-zÀ-ÿ]+)\.?\s*$',
    re.IGNORECASE
)
BLOCK_RE = re.compile(r"^\\?\*?\s*\*{1,3}\s*([A-Za-zÀ-ÿ' \-]{2,40}?)\s*\*{1,3}\s*:\s*(.*)$")
# Blocs "jeu" (mots croisés, grilles à remplir, consignes de jeu...) —
# non exploitables en simple texte, exclus dès la génération.
GAME_LABEL_RE = re.compile(r'lalao|hilalao|jeu|jouons|alamino.*soraty|veuillez.*crire', re.IGNORECASE)

def clean_line(line):
    line = line.strip()

    # retire le gras Markdown
    line = re.sub(r'^\*+', '', line)
    line = re.sub(r'\*+$', '', line)

    # retire les # des titres Markdown
    line = re.sub(r'^#+\s*', '', line)

    # espaces
    line = re.sub(r'\s+', ' ', line)

    return line.strip()

def norm(s: str) -> str:
    s = unicodedata.normalize('NFKD', s).upper()
    s = ''.join(c for c in s if c.isalnum())
    return BOOK_NAME_ALIASES.get(s, s) if s in BOOK_NAME_ALIASES else s


def _norm_key(s: str) -> str:
    """Normalise puis résout un alias éventuel (l'alias lui-même doit être re-normalisé)."""
    s = unicodedata.normalize('NFKD', s).upper()
    raw = ''.join(c for c in s if c.isalnum())
    alias = BOOK_NAME_ALIASES.get(raw)
    if alias:
        return ''.join(c for c in unicodedata.normalize('NFKD', alias).upper() if c.isalnum())
    return raw


def md_inline_to_html(text: str) -> str:
    text = text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    text = text.replace('\\*', '*').replace("\\'", "'")
    # Retire les images résiduelles du Word converti (syntaxe markdown,
    # avec ou sans attributs {width=...}) — non exploitables en HTML simple.
    text = re.sub(r'!\[[^\]]*\]\([^)]*\)(\{[^}]*\})?', '', text)
    text = re.sub(r'\*\*\*(.+?)\*\*\*', r'<strong><em>\1</em></strong>', text)
    text = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', text)
    text = re.sub(r'\*(.+?)\*', r'<em>\1</em>', text)
    return text.strip()


def paragraphs_from(lines):
    paras, buf = [], []
    for l in lines + ['']:
        if l.strip() == '':
            if buf:
                paras.append(' '.join(x.strip() for x in buf))
                buf = []
        else:
            buf.append(l)
    return paras


def docx_to_markdown(docx_path: str) -> str:
    if shutil.which('pandoc') is None:
        sys.exit(
            "\u274c 'pandoc' n'est pas install\u00e9 ou pas dans le PATH.\n"
            "   Windows : winget install --id JohnMacFarlane.Pandoc\n"
            "   Mac     : brew install pandoc\n"
            "   Linux   : sudo apt install pandoc"
        )
    md_path = os.path.basename(os.path.splitext(docx_path)[0]) + '.md'
    subprocess.run(['pandoc', '-t', 'markdown', docx_path, '-o', md_path], check=True)
    return md_path


def build_books_index(bible_path, lang):
    with open(bible_path, encoding='utf-8') as f:
        bible = json.load(f)
    books_sorted = sorted(bible['books'], key=lambda b: b['book_number'])

    index = {}
    if lang == 'mg':
        for b in bible['books']:
            index[_norm_key(b['long_name'])] = b
    else:
        for b, fr_name in zip(books_sorted, FR_BOOK_NAMES_IN_ORDER):
            index[_norm_key(fr_name)] = b
    return index


# Code F/Q en fin de ligne : "F2", "F 2", "F.2", "F . 5", "Q.1", "F.1,4", "F . 2, 4"...
CODE_RE = re.compile(r'[FQ]\s*\.?\s*((?:\d\s*[,:]?\s*)+)$', re.IGNORECASE)

# livre + chapitre.verset[-verset] ; séparateur "." ou ":" ; espaces flexibles
REF_FULL_RE = re.compile(
    r"^(?P<book>.+?)\s*(?P<c1>\d+)\s*[.:]\s*(?P<v1>\d+)"
    r"(?:\s*[-.]{1,2}\s*(?:(?P<c2>\d+)\s*[.:]\s*)?(?P<v2>\d+))?\s*$"
)
# livre à un seul chapitre (ex: OBADIA 1-15, sans point du tout)
REF_1CHAP_RE = re.compile(r"^(?P<book>.+?)\s+(?P<v1>\d+)(?:\s*-\s*(?P<v2>\d+))?\s*$")


def parse_reference(ref_text, books_index):
    ref_text = ref_text.replace('*', '').strip()

    # Convertit les chiffres romains en chiffres arabes
    ref_text = re.sub(r'^I\s+', '1 ', ref_text, flags=re.IGNORECASE)
    ref_text = re.sub(r'^II\s+', '2 ', ref_text, flags=re.IGNORECASE)
    ref_text = re.sub(r'^III\s+', '3 ', ref_text, flags=re.IGNORECASE)

    # 1) Extrait et retire le(s) code(s) F/Q en fin de ligne (optionnel)
    codes = []
    cm = CODE_RE.search(ref_text)
    if cm:
        codes = [int(d) for d in re.findall(r'\d', cm.group(1))]
        ref_text = ref_text[:cm.start()].strip()

    # 2) Retire les lettres de sous-verset ("23a" / "43b" -> "23" / "43")
    ref_text = re.sub(r'(\d)\s*[a-cà]\b', r'\1', ref_text, flags=re.IGNORECASE)

    # 3) Parse "livre chapitre.verset[-verset]"
    m = REF_FULL_RE.match(ref_text)
    if m:
        book_raw = m.group('book').strip()
        c1, v1 = int(m.group('c1')), int(m.group('v1'))
        c2 = int(m.group('c2')) if m.group('c2') else c1
        v2 = int(m.group('v2')) if m.group('v2') else v1
    else:
        m1 = REF_1CHAP_RE.match(ref_text)
        if not m1:
            return None
        book_raw = m1.group('book').strip()
        c1 = c2 = 1
        v1 = int(m1.group('v1'))
        v2 = int(m1.group('v2')) if m1.group('v2') else v1

    key = _norm_key(book_raw)
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
        'codes': codes,
    }


def parse_language_file(md_path, books_index, month_aliases):
    with open(md_path, encoding='utf-8') as f:
        lines = f.read().split('\n')

    header_idx = []
    for i, l in enumerate(lines):
        # m = DAY_RE.match(l.strip())
        m = DAY_RE.match(clean_line(l))
        if m and norm(m.group(2)) in {norm(k) for k in month_aliases}:
            header_idx.append(i)
    header_idx.append(len(lines))

    entries = {}
    warnings = []

    for bi in range(len(header_idx) - 1):
        start, end = header_idx[bi], header_idx[bi + 1]
        block = lines[start:end]
         # m = DAY_RE.match(block[0].strip())
        m = DAY_RE.match(clean_line(block[0]))
        day_num = int(m.group(1))
        month_key = m.group(2).upper().rstrip('.')
        month_num = month_aliases.get(month_key) or month_aliases.get(norm(month_key))
        if not month_num:
            for k, v in month_aliases.items():
                if norm(k) == norm(month_key):
                    month_num = v
                    break
        if not month_num:
            warnings.append(f"Mois inconnu: {month_key!r} (ligne {start+1})")
            continue

        idx = 1
        while idx < len(block) and block[idx].strip() == '':
            idx += 1
        refline = block[idx].strip() if idx < len(block) else ''
        idx += 1

        ref_info = parse_reference(refline, books_index)
        if not ref_info:
            warnings.append(f"Référence non résolue: {refline!r} ({day_num}/{month_num}, ligne {start+1})")
            continue

        while idx < len(block) and block[idx].strip() == '':
            idx += 1
        title = block[idx].strip() if idx < len(block) else ''
        idx += 1

        rest_paragraphs = paragraphs_from(block[idx:])

        mode = 'body'
        body_paras = []
        blocs = []
        skip_current_bloc = False
        for para in rest_paragraphs:
            bm = BLOCK_RE.match(para)
            if bm:
                mode = 'blocs'
                label = bm.group(1).strip()
                if GAME_LABEL_RE.search(label):
                    skip_current_bloc = True
                else:
                    skip_current_bloc = False
                    blocs.append({'label': label, 'text': md_inline_to_html(bm.group(2).strip())})
            elif mode == 'blocs':
                if skip_current_bloc:
                    pass  # suite d'un bloc jeu exclu -> ignorée
                elif blocs:
                    blocs[-1]['text'] += ' ' + md_inline_to_html(para)
            else:
                body_paras.append(para)

        body_html = ''.join(f'<p>{md_inline_to_html(p)}</p>' for p in body_paras)

        entries[(month_num, day_num)] = {
            'toko_sy_andininy': ref_info['display'],
            'ref': {
                'book_number': ref_info['book_number'],
                'chapter_start': ref_info['chapter_start'],
                'verse_start': ref_info['verse_start'],
                'chapter_end': ref_info['chapter_end'],
                'verse_end': ref_info['verse_end'],
            },
            'fanontaniana_ref': ref_info['codes'],
            'lohateny': title,
            'famelabelarana': body_html,
            'blocs': blocs,
        }

    return entries, warnings


def main(mg_md, fr_md, bible_path, taona, out_path):
    books_mg = build_books_index(bible_path, 'mg')
    books_fr = build_books_index(bible_path, 'fr')

    entries_mg, warn_mg = parse_language_file(mg_md, books_mg, MOIS_MG)
    entries_fr, warn_fr = parse_language_file(fr_md, books_fr, MOIS_FR)

    all_keys = sorted(set(entries_mg.keys()) | set(entries_fr.keys()))
    andro = []
    missing_mg, missing_fr = [], []

    for (month_num, day_num) in all_keys:
        try:
            daty = date(taona, month_num, day_num).isoformat()
        except ValueError:
            continue

        mg = entries_mg.get((month_num, day_num))
        fr = entries_fr.get((month_num, day_num))
        if not mg:
            missing_mg.append(daty)
        if not fr:
            missing_fr.append(daty)

        andro.append({'daty': daty, 'mg': mg, 'fr': fr})

    andro.sort(key=lambda a: a['daty'])

    result = {
        'periode': {'taona': taona, 'description': f"Fanilo — teny fanazavana isan'andro, taona {taona}"},
        'andro': andro
    }

    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"OK : {len(andro)} jours écrits dans {out_path}")
    print(f"  MG : {len(entries_mg)} jours parsés, {len(warn_mg)} avertissement(s)")
    print(f"  FR : {len(entries_fr)} jours parsés, {len(warn_fr)} avertissement(s)")
    if missing_mg:
        print(f"\n⚠️  {len(missing_mg)} jour(s) sans version MG (ex: {missing_mg[:5]})")
    if missing_fr:
        print(f"⚠️  {len(missing_fr)} jour(s) sans version FR (ex: {missing_fr[:5]})")
    if warn_mg:
        print(f"\n--- Avertissements MG ({len(warn_mg)}) ---")
        for w in warn_mg[:30]:
            print(' -', w)
    if warn_fr:
        print(f"\n--- Avertissements FR ({len(warn_fr)}) ---")
        for w in warn_fr[:30]:
            print(' -', w)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description="Convertit les Word Fanilo (MG+FR) en un JSON bilingue prêt à héberger.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""Exemple :
  python3 convert_fanilo.py Fanilo_MG_2027.docx Fanilo_FR_2027.docx \\
      --bible ../fihiranaFMMB/src/assets/databases/Bible.json \\
      --taona 2027 --out fanilo-2027.json"""
    )
    parser.add_argument('docx_mg', help="Chemin du Word Fanilo en malgache")
    parser.add_argument('docx_fr', help="Chemin du Word Fanilo en français")
    parser.add_argument('--bible', required=True, help="Chemin vers Bible.json")
    parser.add_argument('--taona', type=int, required=True, help="Année, ex: 2026")
    parser.add_argument('--out', required=True, help="Fichier JSON de sortie")
    args = parser.parse_args()

    mg_md = docx_to_markdown(args.docx_mg)
    fr_md = docx_to_markdown(args.docx_fr)

    main(mg_md, fr_md, args.bible, args.taona, args.out)
