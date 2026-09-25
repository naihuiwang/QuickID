"""QuickID marker matching core, ported from the cleaned Windows Forms source."""
import csv
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
import configparser
import zipfile
import xml.etree.ElementTree as ET
from urllib.parse import urljoin, urlparse, unquote

HEADINGS = ('Easy Label', 'Common Name', 'Species/Genus', 'P1', 'A', "A'", 'B', 'C', 'P2', 'D', 'E', 'F', "F'", 'G', "G'")
MARKERS = ('P1', 'A', "A'", 'B', 'C', 'P2', 'D', 'E', 'F', "F'", 'G', "G'")
MAIN = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
PKG = '{http://schemas.openxmlformats.org/package/2006/relationships}'
DEFAULTS = Path(__file__).with_name('app.ini.example')


def default_settings():
    parser = configparser.ConfigParser()
    parser.optionxform = str
    parser.read(DEFAULTS, encoding='utf-8')
    return dict(parser['database'])


def _col(reference):
    letters = ''.join(c for c in reference if c.isalpha()).upper()
    index = 0
    for letter in letters:
        index = index * 26 + ord(letter) - ord('A') + 1
    return index - 1


def _cells(row, strings):
    values = {}
    for cell in row.findall(MAIN + 'c'):
        col = _col(cell.attrib.get('r', ''))
        if col < 0:
            continue
        kind = cell.attrib.get('t')
        raw = cell.findtext(MAIN + 'v', default='')
        if kind == 's':
            raw = strings[int(raw)]
        elif kind == 'inlineStr':
            raw = ''.join(t.text or '' for t in cell.iter(MAIN + 't'))
        elif kind == 'e':
            raise ValueError('Excel reference contains an error cell')
        values[col] = raw
    return values


def read_reference(path, include_headers=False, required_markers=MARKERS):
    """Read the entire first worksheet in its original column order."""
    with zipfile.ZipFile(path) as archive:
        workbook = ET.fromstring(archive.read('xl/workbook.xml'))
        first = workbook.find('.//' + MAIN + 'sheet')
        if first is None:
            raise ValueError('Reference workbook has no sheets')
        rels = ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
        relation = next((r for r in rels.findall(PKG + 'Relationship')
                         if r.attrib['Id'] == first.attrib.get(REL + 'id')), None)
        if relation is None or relation.attrib.get('TargetMode') == 'External':
            raise ValueError('First worksheet is missing')
        target = unquote(urlparse(urljoin('https://quickid.invalid/xl/workbook.xml',
                                        relation.attrib['Target'])).path).lstrip('/')
        if not target.startswith('xl/'):
            raise ValueError('Invalid worksheet path')
        strings = []
        if 'xl/sharedStrings.xml' in archive.namelist():
            shared = ET.fromstring(archive.read('xl/sharedStrings.xml'))
            strings = [''.join(t.text or '' for t in si.iter(MAIN + 't'))
                       for si in shared.findall(MAIN + 'si')]
        sheet = ET.fromstring(archive.read(target))
        rows = sheet.findall('.//' + MAIN + 'sheetData/' + MAIN + 'row')
        if not rows:
            raise ValueError('Reference worksheet is empty')
        first_row = _cells(rows[0], strings)
        width = max(first_row, default=-1) + 1
        if width == 0:
            raise ValueError('Reference worksheet has no headings')
        output_headers = [first_row.get(index, '') for index in range(width)]
        names = {index: value.strip() for index, value in enumerate(output_headers) if value.strip()}
        folded = [value.casefold() for value in names.values()]
        if len(folded) != len(set(folded)):
            duplicate = next(value for value in names.values() if folded.count(value.casefold()) > 1)
            raise ValueError(f'Duplicate reference heading: {duplicate}')
        for heading in required_markers:
            matches = [index for index, value in names.items() if value.casefold() == heading.casefold()]
            if len(matches) != 1:
                raise ValueError(f"Reference must have exactly one heading '{heading}' (found {len(matches)})")
        canonical = {heading: next(index for index, value in names.items()
                                   if value.casefold() == heading.casefold()) for heading in required_markers}
        # Blank headings still have distinct internal keys; the CSV keeps them blank.
        keys = [names.get(index, f'\0quickid_column_{index}') for index in range(width)]
        result = []
        for row in rows[1:]:
            data = _cells(row, strings)
            if any(value.strip() for value in data.values()):
                values = {key: data.get(index, '') for index, key in enumerate(keys)}
                values.update({heading: data.get(index, '') for heading, index in canonical.items()})
                result.append(values)
        return (result, output_headers, keys) if include_headers else result


def read_peaks(path):
    # Round from the second decimal digit to one decimal place for matching.
    with open(path, encoding='utf-8-sig') as stream:
        masses = []
        for line in stream:
            raw = line.split('\t', 1)[0].strip()
            try:
                mass = Decimal(raw)
            except InvalidOperation:
                continue
            if mass.is_finite():
                masses.append(mass.quantize(Decimal('0.1'), rounding=ROUND_HALF_UP))
    return masses


def _number(raw):
    if '(?)' in raw or not raw.strip():
        return None
    try:
        return Decimal(raw.strip())
    except InvalidOperation:
        return None


def _hits(raw, peaks, settings, name, bounds=None):
    mass = _number(raw)
    if mass is None:
        return []
    if bounds is None:
        bounds = (Decimal(settings['tb' + name.replace("'", '1') + '1']),
                  Decimal(settings['tb' + name.replace("'", '1') + '2']))
    low, high = bounds
    return [peak for peak in peaks if mass + low <= peak <= mass + high]


def _filter_one(rows, peaks, settings, name, bounds=None):
    missing = [row for row in rows if _number(row[name]) is not None and
               not _hits(row[name], peaks, settings, name, bounds)]
    # Legacy behavior: keep all candidates if this marker matches no row.
    return [row for row in rows if row not in missing] if missing and any(
        _hits(row[name], peaks, settings, name, bounds) for row in rows) else rows


def _filter_pair(rows, peaks, settings, first, second, first_bounds=None, second_bounds=None):
    # A, F, G: a row survives when either member of the pair matches.
    keep = []
    any_match = False
    for row in rows:
        a, b = _number(row[first]), _number(row[second])
        ha = _hits(row[first], peaks, settings, first, first_bounds) if a is not None else []
        hb = _hits(row[second], peaks, settings, second, second_bounds) if b is not None else []
        any_match |= bool(ha or hb)
        if a is None or b is None or ha or hb:
            keep.append(row)
    return keep if any_match and keep else rows


def fixed_markers(settings):
    """Names can change; matching order and paired positions stay fixed."""
    slots = ('P1', 'A', 'A1', 'B', 'C', 'P2', 'D', 'E', 'F', 'F1', 'G', 'G1')
    names = [settings.get('marker_' + key + '_name', default).strip()
             for key, default in zip(slots, MARKERS)]
    if not any(names):
        return []
    if not all(names):
        raise ValueError('Fill all twelve built-in marker names or leave all twelve blank')
    if len({name.casefold() for name in names}) != len(names):
        raise ValueError('Built-in marker names must be unique')
    return [(name,
             (Decimal(settings['tb' + key + '1']), Decimal(settings['tb' + key + '2'])))
            for key, name in zip(slots, names)]


def match(rows, peaks, settings, extra_markers=(), built_in=None):
    built_in = fixed_markers(settings) if built_in is None else built_in
    candidates = list(rows)
    filter_order = (3, 6, (1, 2), 4, 7, (8, 9), (10, 11), 0, 5) if built_in else ()
    for item in filter_order:
        if isinstance(item, tuple):
            (first, first_bounds), (second, second_bounds) = (built_in[i] for i in item)
            candidates = _filter_pair(candidates, peaks, settings, first, second,
                                      first_bounds, second_bounds)
        else:
            name, bounds = built_in[item]
            candidates = _filter_one(candidates, peaks, settings, name, bounds)
    # Process the optional markers after all twelve built-in markers.
    for name, bounds in extra_markers:
        candidates = _filter_one(candidates, peaks, settings, name, bounds)
    found = []
    for marker, bounds in [*built_in, *extra_markers]:
        unique = dict.fromkeys(row[marker].strip() for row in candidates)
        matched = dict.fromkeys(str(peak) for mass in unique
                                for peak in _hits(mass, peaks, settings, marker, bounds))
        found.append('/'.join(matched))
    flags = []
    pairs = ((1, 2), (8, 9), (10, 11)) if built_in else ()
    for first_index, second_index in pairs:
        (first, first_bounds), (second, second_bounds) = (built_in[i]
                                                         for i in (first_index, second_index))
        for row in candidates:
            if _number(row[first]) is None or _number(row[second]) is None:
                continue
            first_hits = _hits(row[first], peaks, settings, first, first_bounds)
            if any(peak in _hits(other[second], peaks, settings, second, second_bounds)
                   for peak in first_hits for other in candidates):
                flags.append(second + ' don\u2019t match')
                break
    return candidates, found, flags


def selected_extra_markers(rows, settings, log=None, built_in=None):
    """Enable only fully specified names that identify a reference column."""
    available = {name.casefold(): name for name in rows[0]}
    selected = []
    used = {name.casefold() for name in HEADINGS[:3]}
    used.update(name.casefold() for name, _ in (fixed_markers(settings)
                                                if built_in is None else built_in))
    for index in range(1, 9):
        prefix = f'custom{index}_'
        name = settings.get(prefix + 'name', '').strip()
        low = settings.get(prefix + 'low', '').strip()
        high = settings.get(prefix + 'high', '').strip()
        if not all((name, low, high)):
            if name and log:
                log(f'Skipped optional marker {name}: both tolerances are required')
            continue
        if name.casefold() in used:
            if log:
                log(f'Skipped optional marker {name}: name is already in use')
            continue
        actual = available.get(name.casefold())
        if actual is None:
            if log:
                log(f'Skipped optional marker {name}: column missing from reference')
            continue
        bounds = (Decimal(low), Decimal(high))
        if bounds[0] > bounds[1]:
            raise ValueError(f'Invalid tolerance range for {name}')
        selected.append((actual, bounds))
        used.add(name.casefold())
    return selected


def process(input_dir, output_file, settings=None, log=None, reference_file=None):
    settings = settings or default_settings()
    warning_enabled = str(settings.get('peak_warning_enabled', '1')).strip()
    if warning_enabled not in ('0', '1'):
        raise ValueError('Peak-count warning must be enabled or disabled')
    warning_limit = None
    if warning_enabled == '1':
        raw_limit = str(settings.get('peak_warning_threshold', '500')).strip()
        if not raw_limit.isdecimal() or int(raw_limit) < 1:
            raise ValueError('Enter a positive whole-number peak-count warning threshold')
        warning_limit = int(raw_limit)
    built_in = fixed_markers(settings)
    fixed_names = [name for name, _ in built_in]
    directory = Path(input_dir)
    if not directory.is_dir():
        raise ValueError('Select an existing input folder')
    reference = Path(reference_file or settings.get('reference') or '')
    if reference.suffix.casefold() != '.xlsx' or not reference.is_file():
        raise ValueError('Select an existing .xlsx reference file (save older .xls as .xlsx first)')
    peaks_files = sorted(path for path in directory.iterdir()
                         if path.suffix.casefold() in ('.txt', '.tab'))
    if not peaks_files:
        raise ValueError('No .txt or .tab peak lists in input folder')
    rows, reference_headers, reference_keys = read_reference(
        reference, include_headers=True, required_markers=fixed_names)
    if not rows:
        raise ValueError('Reference worksheet has no data rows')
    extras = selected_extra_markers(rows, settings, log, built_in)
    if not built_in and not extras:
        raise ValueError('Enable at least one optional marker present in the reference workbook')
    output_headers = ['Sample name', 'QuickID marker count',
                      *reference_headers, 'QuickID mismatch flags']
    reference_positions = {name.strip().casefold(): index for index, name in enumerate(reference_headers)
                           if name.strip()}
    output = Path(output_file)
    output.parent.mkdir(parents=True, exist_ok=True)
    # Produce complete results in memory; failed reads do not truncate old output.
    records = [output_headers]
    for peak_path in peaks_files:
        peaks = read_peaks(peak_path)
        candidates, marker_hits, flags = match(rows, peaks, settings, extras, built_in)
        if warning_limit is not None and len(peaks) > warning_limit:
            flags.append(f'Possible false positive (>{warning_limit} peaks)')
        count = sum(bool(item) for item in marker_hits)
        summary = [''] * len(output_headers)
        for marker, hit in zip([*fixed_names, *(name for name, _ in extras)], marker_hits):
            summary[reference_positions[marker.casefold()] + 2] = hit
        summary[0] = peak_path.name
        summary[1] = str(count)
        summary[-1] = '; '.join(flags)
        records.append(summary)
        if count >= int(settings['tbminrows']):
            records.extend([['', '', *(row[key] for key in reference_keys), '']
                            for row in candidates])
        records.append([''] * len(output_headers))
        if log:
            log(f'{peak_path.name}: {count} marker(s), {len(candidates)} candidate row(s)')
    with output.open('w', encoding='utf-8-sig', newline='') as stream:
        csv.writer(stream).writerows(records)
    return len(peaks_files)
