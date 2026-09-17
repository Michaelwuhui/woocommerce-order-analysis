"""Shared product-name recognition. Pure parser with explicit cached rules."""
import re


def parse_product_name(name, brands_cache, series_cache):
    """
    Parse product name to extract brand, series, puff count, and flavor.
    
    Examples:
    - "IGET ONE 12000 puffs - Mixed Berries" → brand: IGET, puffs: 12000, flavor: Mixed Berries
    - "Crystal Blind 25000 Puffs" → brand: Crystal Blind, puffs: 25000
    - "FUMO king 6000 puffs Disposable Vape 20mg" → brand: FUMO, series: king, puffs: 6000
    """
    if not name:
        return {'brand': None, 'series': None, 'puffs': None, 'flavor': None, 'normalized': None}
    
    result = {'brand': None, 'series': None, 'puffs': None, 'flavor': None, 'normalized': None}
    
    # 1. Extract puff count
    # Handle optional "+" and Polish term "zaciągnięć"
    # Support spaced/dotted/comma'd thousands: "50 000", "50.000", "50,000"
    puffs_end = None  # end index of the puffs marker in `name`, used by flavor fallback below
    puff_match = re.search(r'\b(\d{1,3}(?:[\s.,]\d{3})+|\d+)\+?\s*(?:puffs?|zaciągnięć)', name, re.IGNORECASE)
    if puff_match:
        result['puffs'] = int(re.sub(r'[\s.,]', '', puff_match.group(1)))
        puffs_end = puff_match.end()
    else:
        # Fallback: Look for "Number Disposable" pattern e.g. "9000 Disposable"
        disposable_match = re.search(r'\b(\d{1,3}(?:[\s.,]\d{3})+|\d+)\s*Disposable', name, re.IGNORECASE)
        if disposable_match and int(re.sub(r'[\s.,]', '', disposable_match.group(1))) >= 100:
            result['puffs'] = int(re.sub(r'[\s.,]', '', disposable_match.group(1)))
            puffs_end = disposable_match.end()
    
    # 3. Match brand (earliest position first, then longest match)
    name_upper = name.upper()
    matched_brand = None
    matched_pos = len(name_upper)
    matched_len = 0

    for brand in brands_cache:
        for pattern in brand['patterns']:
            pos = name_upper.find(pattern)
            if pos >= 0 and (pos < matched_pos or (pos == matched_pos and len(pattern) > matched_len)):
                matched_brand = brand
                matched_pos = pos
                matched_len = len(pattern)
    
    if matched_brand:
        result['brand'] = matched_brand['name']
        result['brand_id'] = matched_brand['id']

    # 3b. Match series (if brand found and series_cache available)
    if matched_brand and series_cache:
        matched_series = None
        matched_series_len = 0
        for s in series_cache:
            if s['brand_id'] == matched_brand['id']:
                if s['name'].upper() in name_upper and len(s['name']) > matched_series_len:
                    matched_series = s
                    matched_series_len = len(s['name'])
        if matched_series:
            result['series'] = matched_series['name']
            result['series_id'] = matched_series['id']

    # 4. Extract flavor (usually after separator)
    flavor = None
    for sep in [' - ', ' – ', ' | ', ' / ']:
        if sep in name:
            parts = name.split(sep)
            if len(parts) > 1:
                flavor = parts[-1].strip()
                # Don't treat product type/specs as flavor
                if any(x in flavor.lower() for x in ['puff', 'disposable', 'vape', 'mg', 'ml']):
                    flavor = None
                break

    # 4b. Fallback: if no separator-based flavor was found but we know where the
    # puffs marker ends, treat the remaining text as a flavor candidate. Handles
    # formats like "Merrymi Blade 30000 Puffs Aperol" where the flavor is just
    # appended without a separator.
    if not flavor and puffs_end is not None:
        tail = name[puffs_end:].strip()
        # Drop leading separators/punctuation (keep periods inside flavors like "Mr. Blue")
        tail = re.sub(r'^[\s\-–|/,:;]+', '', tail).strip()
        if tail:
            # Strip standalone product-spec words (English + Polish equivalents)
            cleaned = re.sub(
                r'\b(disposable|vape|jednorazowy|e-?papieros|puffs?)\b',
                ' ',
                tail,
                flags=re.IGNORECASE,
            )
            # Strip standalone unit values like "20mg", "5ml"
            cleaned = re.sub(r'\b\d+\s*(mg|ml)\b', ' ', cleaned, flags=re.IGNORECASE)
            # Collapse separator residue (but preserve periods)
            cleaned = re.sub(r'[\-–|/,:;]+', ' ', cleaned)
            cleaned = ' '.join(cleaned.split()).strip()
            if cleaned:
                flavor = cleaned

    result['flavor'] = flavor
    
    # 5. Build normalized name
    parts = []
    if result['brand']:
        parts.append(result['brand'])
    if result['puffs']:
        parts.append(f"{result['puffs']} Puffs")
    if result['flavor']:
        parts.append(result['flavor'])
    
    result['normalized'] = ' - '.join(parts) if parts else name
    
    return result


