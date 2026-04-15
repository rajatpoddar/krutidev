# coding=utf-8
"""
Krutidev to Unicode (Devanagari) converter — Python 3 port.
Original algorithm: Language Technology Research Center, IIIT Hyderabad
Source: https://github.com/ltrc/kru2uni
"""
import re

# ---------------------------------------------------------------------------
# Mapping table: (krutidev_sequence, unicode_sequence)
# Order is critical — longer / more-specific patterns must come first.
# ---------------------------------------------------------------------------
K2U = [
    ('\xf1',    '\u0970'),   # ñ  → ॰
    ('Q+Z',     'QZ+'),
    ('sas',     'sa'),
    ('aa',      'a'),
    (')Z',      '\u0930\u094d\u0926\u094d\u0927'),  # र्द्ध
    ('ZZ',      'Z'),
    ('\u2018',  '"'),
    ('\u2019',  '"'),
    ('\u201c',  "'"),
    ('\u201d',  "'"),

    # Devanagari digits (Krutidev uses Windows-1252 private chars)
    ('\xe5',    '\u0966'),   # å → ०
    ('\u0192',  '\u0967'),   # ƒ → १
    ('\u201e',  '\u0968'),   # „ → २
    ('\u2026',  '\u0969'),   # … → ३
    ('\u2020',  '\u096a'),   # † → ४
    ('\u2021',  '\u096b'),   # ‡ → ५
    ('\u02c6',  '\u096c'),   # ˆ → ६
    ('\u2030',  '\u096d'),   # ‰ → ७
    ('\u0160',  '\u096e'),   # Š → ८
    ('\u2039',  '\u096f'),   # ‹ → ९

    # Nukta / special consonants
    ('\xb6+',   '\u095e\u094d'),  # ¶+ → फ़्
    ('d+',      '\u0958'),        # क़
    ('[+k',     '\u0959'),        # ख़
    ('[+',      '\u0959\u094d'),  # ख़्
    ('x+',      '\u095a'),        # ग़
    ('T+',      '\u091c\u093c\u094d'),  # ज़्
    ('t+',      '\u095b'),        # ज़
    ('M+',      '\u095c'),        # ड़
    ('<+',      '\u095d'),        # ढ़
    ('Q+',      '\u095e'),        # फ़
    (';+',      '\u095f'),        # य़
    ('j+',      '\u0931'),        # ऱ
    ('u+',      '\u0929'),        # ऩ

    # Conjuncts
    ('\xd9k',   '\u0924\u094d\u0924'),        # Ùk → त्त
    ('\xd9',    '\u0924\u094d\u0924\u094d'),  # Ù  → त्त्
    ('\xe4',    '\u0915\u094d\u0924'),        # ä  → क्त
    ('\u2013',  '\u0926\u0943'),              # –  → दृ
    ('\u2014',  '\u0915\u0943'),              # —  → कृ
    ('\xe9',    '\u0928\u094d\u0928'),        # é  → न्न
    ('\u2122',  '\u0928\u094d\u0928\u094d'), # ™  → न्न्
    ('=kk',     '=k'),
    ('f=k',     'f='),

    ('\xe0',    '\u0939\u094d\u0928'),        # à → ह्न
    ('\xe1',    '\u0939\u094d\u092f'),        # á → ह्य
    ('\xe2',    '\u0939\u0943'),              # â → हृ
    ('\xe3',    '\u0939\u094d\u092e'),        # ã → ह्म
    ('\xbaz',   '\u0939\u094d\u0930'),        # ºz → ह्र
    ('\xba',    '\u0939\u094d'),              # º  → ह्
    ('\xed',    '\u0926\u094d\u0926'),        # í  → द्द
    ('{k',      '\u0915\u094d\u0937'),        # क्ष
    ('{',       '\u0915\u094d\u0937\u094d'),  # क्ष्
    ('=',       '\u0924\u094d\u0930'),        # त्र
    ('\xab',    '\u0924\u094d\u0930\u094d'),  # « → त्र्
    ('N\xee',   '\u091b\u094d\u092f'),        # Nî → छ्य
    ('V\xee',   '\u091f\u094d\u092f'),        # Vî → ट्य
    ('B\xee',   '\u0920\u094d\u092f'),        # Bî → ठ्य
    ('M\xee',   '\u0921\u094d\u092f'),        # Mî → ड्य
    ('<\xee',   '\u0922\u094d\u092f'),        # <î → ढ्य
    ('|',       '\u0926\u094d\u092f'),        # द्य
    ('K',       '\u091c\u094d\u091e'),        # ज्ञ
    ('}',       '\u0926\u094d\u0935'),        # द्व
    ('J',       '\u0936\u094d\u0930'),        # श्र
    ('V\xaa',   '\u091f\u094d\u0930'),        # Vª → ट्र
    ('M\xaa',   '\u0921\u094d\u0930'),        # Mª → ड्र
    ('<\xaa\xaa', '\u0922\u094d\u0930'),      # <ªª → ढ्र
    ('N\xaa',   '\u091b\u094d\u0930'),        # Nª → छ्र
    ('\xd8',    '\u0915\u094d\u0930'),        # Ø  → क्र
    ('\xdd',    '\u092b\u094d\u0930'),        # Ý  → फ्र
    ('nzZ',     '\u0930\u094d\u0926\u094d\u0930'),  # र्द्र
    ('\xe6',    '\u0926\u094d\u0930'),        # æ → द्र
    ('\xe7',    '\u092a\u094d\u0930'),        # ç → प्र
    ('\xc1',    '\u092a\u094d\u0930'),        # Á → प्र
    ('xz',      '\u0917\u094d\u0930'),        # ग्र
    ('#',       '\u0930\u0941'),              # रु
    (':',       '\u0930\u0942'),              # रू

    # Independent vowels
    ('v\u201a', '\u0911'),   # v‚ → ऑ
    ('vks',     '\u0913'),   # ओ
    ('vkS',     '\u0914'),   # औ
    ('vk',      '\u0906'),   # आ
    ('v',       '\u0905'),   # अ
    ('b\xb1',   '\u0908\u0902'),  # b± → ईं
    ('\xc3',    '\u0908'),   # Ã → ई
    ('bZ',      '\u0908'),   # ई
    ('b',       '\u0907'),   # इ
    ('m',       '\u0909'),   # उ
    ('\xc5',    '\u090a'),   # Å → ऊ
    (',s',      '\u0910'),   # ऐ
    (',',       '\u090f'),   # ए
    ('_',       '\u090b'),   # ऋ

    # Consonants (multi-char first)
    ('\xf4',    '\u0915\u094d\u0915'),  # ô → क्क
    ('Dk',      '\u0915'),   # क
    ('d',       '\u0915'),   # क
    ('D',       '\u0915\u094d'),  # क्
    ('[k',      '\u0916'),   # ख
    ('[',       '\u0916\u094d'),  # ख्
    ('Xk',      '\u0917'),   # ग
    ('x',       '\u0917'),   # ग
    ('X',       '\u0917\u094d'),  # ग्
    ('\xc4',    '\u0918'),   # Ä → घ
    ('?k',      '\u0918'),   # घ
    ('?',       '\u0918\u094d'),  # घ्
    ('\xb3',    '\u0919'),   # ³ → ङ
    ('pkS',     '\u091a\u0948'),  # चै
    ('Pk',      '\u091a'),   # च
    ('p',       '\u091a'),   # च
    ('P',       '\u091a\u094d'),  # च्
    ('N',       '\u091b'),   # छ
    ('Tk',      '\u091c'),   # ज
    ('t',       '\u091c'),   # ज
    ('T',       '\u091c\u094d'),  # ज्
    ('>',       '\u091d'),   # झ
    ('\xf7',    '\u091d\u094d'),  # ÷ → झ्
    ('\xa5',    '\u091e'),   # ¥ → ञ
    ('\xea',    '\u091f\u094d\u091f'),  # ê → ट्ट
    ('\xeb',    '\u091f\u094d\u0920'),  # ë → ट्ठ
    ('V',       '\u091f'),   # ट
    ('B',       '\u0920'),   # ठ
    ('\xec',    '\u0921\u094d\u0921'),  # ì → ड्ड
    ('\xef',    '\u0921\u094d\u0922'),  # ï → ड्ढ
    ('M',       '\u0921'),   # ड
    ('<',       '\u0922'),   # ढ
    ('.k',      '\u0923'),   # ण
    ('.',       '\u0923\u094d'),  # ण्
    ('Rk',      '\u0924'),   # त
    ('r',       '\u0924'),   # त
    ('R',       '\u0924\u094d'),  # त्
    ('Fk',      '\u0925'),   # थ
    ('F',       '\u0925\u094d'),  # थ्
    (')',       '\u0926\u094d\u0927'),  # द्ध
    ('n',       '\u0926'),   # द
    ('/k',      '\u0927'),   # ध
    ('/',       '\u0927\u094d'),  # ध्
    ('\xcb',    '\u0927\u094d'),  # Ë → ध्
    ('\xe8',    '\u0927'),   # è → ध
    ('Uk',      '\u0928'),   # न
    ('u',       '\u0928'),   # न
    ('U',       '\u0928\u094d'),  # न्
    ('Ik',      '\u092a'),   # प
    ('i',       '\u092a'),   # प
    ('I',       '\u092a\u094d'),  # प्
    ('Q',       '\u092b'),   # फ
    ('\xb6',    '\u092b\u094d'),  # ¶ → फ्
    ('Ck',      '\u092c'),   # ब
    ('c',       '\u092c'),   # ब
    ('C',       '\u092c\u094d'),  # ब्
    ('Hk',      '\u092d'),   # भ
    ('H',       '\u092d\u094d'),  # भ्
    ('Ek',      '\u092e'),   # म
    ('e',       '\u092e'),   # म
    ('E',       '\u092e\u094d'),  # म्
    (';',       '\u092f'),   # य
    ('\xb8',    '\u092f\u094d'),  # ¸ → य्
    ('j',       '\u0930'),   # र
    ('Yk',      '\u0932'),   # ल
    ('y',       '\u0932'),   # ल
    ('Y',       '\u0932\u094d'),  # ल्
    ('G',       '\u0933'),   # ळ
    ('Ok',      '\u0935'),   # व
    ('o',       '\u0935'),   # व
    ('O',       '\u0935\u094d'),  # व्
    ("'k",      '\u0936'),   # श
    ("'",       '\u0936\u094d'),  # श्
    ('"k',      '\u0937'),   # ष
    ('"',       '\u0937\u094d'),  # ष्
    ('Lk',      '\u0938'),   # स
    ('l',       '\u0938'),   # स
    ('L',       '\u0938\u094d'),  # स्
    ('g',       '\u0939'),   # ह

    # Special combined matras
    ('\xc8',    '\u0940\u0902'),  # È → ीं
    ('saz',     '\u094d\u0930\u0947\u0902'),  # ्रें
    ('z',       '\u094d\u0930'),  # ्र

    # More conjuncts (alternate glyphs)
    ('\xcc',    '\u0926\u094d\u0926'),  # Ì → द्द
    ('\xcd',    '\u091f\u094d\u091f'),  # Í → ट्ट
    ('\xce',    '\u091f\u094d\u0920'),  # Î → ट्ठ
    ('\xcf',    '\u0921\u094d\u0921'),  # Ï → ड्ड
    ('\xd1',    '\u0915\u0943'),        # Ñ → कृ
    ('\xd2',    '\u092d'),              # Ò → भ
    ('\xd3',    '\u094d\u092f'),        # Ó → ्य
    ('\xd4',    '\u0921\u094d\u0922'),  # Ô → ड्ढ
    ('\xd6',    '\u091d\u094d'),        # Ö → झ्
    ('\xd9',    '\u0924\u094d\u0924\u094d'),  # Ù → त्त्
    ('\xdck',   '\u0936'),              # Ük → श
    ('\xdc',    '\u0936\u094d'),        # Ü → श्

    # Matras (vowel signs)
    ('\u201a',  '\u0949'),   # ‚ → ॉ
    ('kas',     '\u094b\u0902'),  # ों
    ('ks',      '\u094b'),   # ो
    ('kS',      '\u094c'),   # ौ
    ('\xa1k',   '\u093e\u0901'),  # ¡k → ाँ
    ('ak',      'k\u0902'),  # ak → k + ं
    ('k',       '\u093e'),   # ा
    ('ah',      '\u0940\u0902'),  # ीं
    ('h',       '\u0940'),   # ी
    ('aq',      '\u0941\u0902'),  # ुं
    ('q',       '\u0941'),   # ु
    ('aw',      '\u0942\u0902'),  # ूं
    ('\xa1w',   '\u0942\u0901'),  # ¡w → ूँ
    ('w',       '\u0942'),   # ू
    ('`',       '\u0943'),   # ृ
    ('\u0300',  '\u0943'),   # ̀ → ृ
    ('as',      '\u0947\u0902'),  # ें
    ('\xb1s',   's\xb1'),    # ±s → s±
    ('s',       '\u0947'),   # े
    ('aS',      '\u0948\u0902'),  # ैं
    ('S',       '\u0948'),   # ै
    ('a\xaa',   '\u094d\u0930\u0902'),  # aª → ्र + ं
    ('\xaa',    '\u094d\u0930'),  # ª → ्र
    # NOTE: 'fa' is NOT in the main table — it's handled in post-processing for glyph Ç
    ('a',       '\u0902'),   # ं
    ('\xa1',    '\u0901'),   # ¡ → ँ
    ('%',       ':'),        # % → : (visarga colon)
    ('W',       '\u0945'),   # ॅ
    ('\u2022',  '\u093d'),   # • → ऽ
    ('\xb7',    '\u093d'),   # · → ऽ
    ('\u2219',  '\u093d'),   # ∙ → ऽ
    ('~j',      '\u094d\u0930'),  # ्र
    ('~',       '\u094d'),   # ्
    ('\\',      '?'),
    ('+',       '\u093c'),   # ़
    ('^',       '\u2018'),   # '
    ('*',       '\u2019'),   # '
    ('\xde',    '\u201c'),   # Þ → "
    ('\xdf',    '\u201d'),   # ß → "
    ('(',       ';'),
    ('\xbc',    '('),        # ¼ → (
    ('\xbd',    ')'),        # ½ → )
    ('\xbf',    '{'),        # ¿ → {
    ('\xc0',    '}'),        # À → }
    ('\xbe',    '='),        # ¾ → =
    ('A',       '\u0964'),   # । (danda)
    # NOTE: '-' maps to '.' in Krutidev but we handle digit-ranges specially in post-processing
    ('-',       '.'),
    ('&',       '-'),
    ('\u0152',  '\u0970'),   # Œ → ॰
    ('\u0178k',  '\u0924\u094d\u0924'),   # Ÿk → त्त  (e.g. foŸkh; = वित्तीय)
    ('\u0178',   '\u0924\u094d\u0924'),   # Ÿ  → त्त  (fallback)
    (']',       ','),
    ('~ ',      '\u094d '),
    ('@',       '/'),
    ('\xae',    '\u0948\u0902'),  # ® → ैं
]

# Vowel signs set for reph repositioning
_VOWEL_SIGNS = set(
    'अआइईउऊएऐओऔािीुूृेैोौंःँॅ '
)


# Pre-build a lookup: first_char → list of (full_kd, unicode) sorted longest-first
_LOOKUP: dict = {}
for _kd, _uni in K2U:
    _fc = _kd[0]
    _LOOKUP.setdefault(_fc, []).append((_kd, _uni))
# Sort each bucket longest-first so greedy match works correctly
for _fc in _LOOKUP:
    _LOOKUP[_fc].sort(key=lambda x: -len(x[0]))
_MAX_KD_LEN = max(len(k) for k, _ in K2U)


def _apply_mapping(text: str) -> str:
    """
    Left-to-right greedy longest-match replacement.
    Also handles 'f' (chhoti-i matra) inline: f + next_char → next_char_unicode + ि
    This must happen here so 'f' grabs the consonant BEFORE 'a' → ं fires.
    """
    out = []
    i = 0
    n = len(text)
    while i < n:
        ch = text[i]

        # Special handling for 'f' (chhoti-i matra in Krutidev)
        # It appears BEFORE the consonant but must come AFTER in Unicode.
        # It may also skip over 'a' (anusvara) to find the consonant: falg = सिंह
        if ch == 'f' and i + 1 < n:
            # Collect any anusvara/chandrabindu chars between f and the consonant
            j = i + 1
            pre_matras = []
            # Skip 'a' (anusvara ं) that comes between f and the consonant
            while j < n and text[j] == 'a' and j + 1 < n:
                pre_matras.append('\u0902')  # anusvara
                j += 1

            next_ch = text[j] if j < n else ''
            found_consonant = False
            if next_ch and next_ch in _LOOKUP:
                for kd, uni in _LOOKUP[next_ch]:
                    end = j + len(kd)
                    if text[j:end] == kd:
                        # Only reposition if the mapped value starts with a consonant
                        if uni and ('\u0915' <= uni[0] <= '\u0939' or
                                    '\u0958' <= uni[0] <= '\u095f'):
                            out.append(uni)           # consonant first
                            out.append('\u093f')      # then ि
                            out.extend(pre_matras)    # then any anusvara etc.
                            i = end
                            found_consonant = True
                            break
            if not found_consonant:
                # Could not find a consonant — pass f through as-is
                out.append(ch)
                i += 1
            continue

        matched = False
        if ch in _LOOKUP:
            for kd, uni in _LOOKUP[ch]:
                end = i + len(kd)
                if text[i:end] == kd:
                    out.append(uni)
                    i = end
                    matched = True
                    break
        if not matched:
            out.append(ch)
            i += 1
    return ''.join(out)


def krutidev_to_unicode(text: str) -> str:
    """Convert a Krutidev-encoded string to Unicode Devanagari (Python 3)."""
    if not isinstance(text, str) or not text.strip():
        return text

    # Protect digit-hyphen-digit patterns (e.g. 2022-23) from '-' → '.' mapping
    _DASH_PLACEHOLDER = '\uE000DASH\uE000'  # Use Unicode private-use area, safe in Excel
    text = re.sub(r'(\d)-(\d)', lambda m: m.group(1) + _DASH_PLACEHOLDER + m.group(2), text)

    # Protect English text in parentheses like (ST, SC, OBC, GEN) from Krutidev mapping
    _PAREN_STORE: list = []
    def _protect_paren(m):
        _PAREN_STORE.append(m.group(0))
        return f'\uE001{len(_PAREN_STORE)-1}\uE001'
    text = re.sub(r'\([A-Z][A-Z0-9,\s/]+\)', _protect_paren, text)

    # Pre-processing: collapse spurious spaces before ्र
    text = text.replace(' \xaa', '\xaa')
    text = text.replace(' ~j', '~j')
    text = text.replace(' z', 'z')

    # Step 1: table-driven replacements using greedy left-to-right longest-match.
    # We cannot use sequential str.replace because shorter patterns (e.g. "'")
    # would fire before longer ones (e.g. "'k") and corrupt the output.
    text = _apply_mapping(text)

    # Step 2: special glyph ± → Zं
    text = text.replace('\xb1', 'Z\u0902')

    # Step 3: Æ → र् + f  (then f-reposition handles it inline in _apply_mapping)
    text = text.replace('\xc6', '\u0930\u094df')

    # NOTE: Step 4 (f-reposition) is now handled inline in _apply_mapping

    # Step 5: Ç / ¯ → fa,  É → र्fa
    text = text.replace('\xc7', 'fa')
    text = text.replace('\xaf', 'fa')
    text = text.replace('\xc9', '\u0930\u094dfa')

    misplaced = re.search(r'fa(.)', text)
    while misplaced:
        ch = misplaced.group(1)
        text = text.replace('fa' + ch, ch + '\u093f\u0902', 1)
        misplaced = re.search(r'fa(.)', text)

    # Step 6: Ê → ीZ
    text = text.replace('\xca', '\u0940Z')

    # Step 7: fix ि् + consonant → ् + consonant + ि
    misplaced = re.search('\u093f\u094d(.)', text)
    while misplaced:
        ch = misplaced.group(1)
        text = text.replace('\u093f\u094d' + ch, '\u094d' + ch + '\u093f', 1)
        misplaced = re.search('\u093f\u094d(.)', text)

    # Step 8: ्Z → Z  (halant before reph is redundant)
    text = text.replace('\u094dZ', 'Z')

    # Step 9: resolve reph 'Z' — place र् before the syllable it belongs to
    misplaced = re.search(r'(.)Z', text)
    while misplaced:
        ch = misplaced.group(1)
        idx = text.index(ch + 'Z')
        # walk left past vowel signs
        while idx >= 0 and text[idx] in _VOWEL_SIGNS:
            idx -= 1
            ch = text[idx] + ch
        text = text.replace(ch + 'Z', '\u0930\u094d' + ch, 1)
        misplaced = re.search(r'(.)Z', text)

    # Step 10: clean up illegal matra placements
    _UNATTACHED = (
        '\u093e\u093f\u0940\u0941\u0942\u0943'
        '\u0947\u0948\u094b\u094c\u0902\u0903\u0901\u0945'
    )
    for m in _UNATTACHED:
        text = text.replace(' ' + m, m)
        text = text.replace(',' + m, m + ',')
        text = text.replace('\u094d' + m, m)

    # Step 11: normalise double halant sequences
    text = text.replace('\u094d\u094d\u0930', '\u094d\u0930')
    text = text.replace('\u094d\u0930\u094d', '\u0930\u094d')
    text = text.replace('\u094d\u094d', '\u094d')
    text = text.replace('\u094d ', ' ')

    # Restore protected hyphens and English parenthetical content
    text = text.replace('\uE000DASH\uE000', '-')
    for idx, original in enumerate(_PAREN_STORE):
        text = text.replace(f'\uE001{idx}\uE001', original)

    return text


# ---------------------------------------------------------------------------
# Unicode → Krutidev (reverse mapping)
# ---------------------------------------------------------------------------

# Build reverse lookup from K2U (unicode → krutidev), longest unicode first
_U2K: list = []
_seen_uni: set = set()
for _kd, _uni in reversed(K2U):
    if _uni and _uni not in _seen_uni and len(_uni) >= 1:
        _seen_uni.add(_uni)
        _U2K.append((_uni, _kd))
# Sort longest unicode sequence first for greedy match
_U2K.sort(key=lambda x: -len(x[0]))


def unicode_to_krutidev(text: str) -> str:
    """
    Approximate reverse conversion: Unicode Devanagari → Krutidev encoding.
    Note: this is a best-effort reverse; some conjuncts may not round-trip perfectly.
    """
    if not isinstance(text, str) or not text.strip():
        return text

    out = []
    i = 0
    n = len(text)
    while i < n:
        matched = False
        for uni, kd in _U2K:
            end = i + len(uni)
            if text[i:end] == uni:
                out.append(kd)
                i = end
                matched = True
                break
        if not matched:
            out.append(text[i])
            i += 1
    return ''.join(out)
