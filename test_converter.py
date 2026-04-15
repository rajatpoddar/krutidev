# coding=utf-8
import sys
sys.path.insert(0, '.')
from converter import krutidev_to_unicode

tests = [
    ('fofgr izi=',               'विहित प्रपत्र'),
    ('ftyk dk uke%&nso?kj',      'जिला का नाम:-देवघर'),
    ('dEI;wVj lgk;d',            'कम्प्यूटर सहायक'),
    ('eksguiqj',                  'मोहनपुर'),
    ('lat; dqekj oekZ',          'संजय कुमार वर्मा'),
    ('e/kqiqj',                   'मधुपुर'),
    ('inLFkkiu dk;kZy; dk uke',  'पदस्थापन कार्यालय का नाम'),
    ('dehZ dk uke',               'कर्मी का नाम'),
    ('mez',                       'उम्र'),
    ('xksih egFkk',               'गोपी महथा'),
    ('vkyksd dqekj',              'आलोक कुमार'),
    ('cztuanu oekZ',              'ब्रजनंदन वर्मा'),
    ('fcgkjh yky',                'बिहारी लाल'),
]

print('--- Conversion Test ---')
passed = 0
for kd, expected in tests:
    result = krutidev_to_unicode(kd)
    ok = result == expected
    status = 'PASS' if ok else 'FAIL'
    if ok:
        passed += 1
    print(f'[{status}]  {kd!r}')
    print(f'       got:      {result}')
    if not ok:
        print(f'       expected: {expected}')

print(f'\n{passed}/{len(tests)} passed')
