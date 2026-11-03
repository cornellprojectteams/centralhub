"""Turn raw.json (from scrape.py) into data/team-photos.json, merging with what is already there.

    python3 tools/team-photos/build.py                  # merge a fresh scrape
    python3 tools/team-photos/build.py --dry-run        # show what would change
    python3 tools/team-photos/build.py --prune          # also delete everyone graduated or no longer listed

Merge rules (the README in this folder has the full list):
- A person found on a team site is added, or updated if already there.
- A person whose record has "locked": true is never changed by a scrape.
- A person with "source": "manual" is never changed or flagged by a scrape.
- A scraped person who is no longer on their team's site gets "status": "left".
  They stay in the file (hidden by default on the page) until someone deletes them.
"""
import json, re, sys, os, datetime, hashlib, urllib.parse

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))
RAW = os.path.join(HERE, 'raw.json')
SITES = os.path.join(HERE, 'sites.json')
OUT = os.path.join(ROOT, 'data', 'team-photos.json')

ROLE_WORDS = re.compile(r'\b(lead|leads|co-?lead|member|members|president|vice|vp|chair|co-?chair|advisor|adviser|captain|'
                        r'director|manager|officer|head|coordinator|pm|treasurer|secretary|founder|instructor|'
                        r'engineer|developer|designer|strategist|associate|analyst|liaison|team|subteam|sub-team|'
                        r'mentor|executive|chief|ceo|cto|coo|cfo|webmaster|historian|intern|consultant|specialist)\b', re.I)
LEAD_WORDS = re.compile(r'\b(lead|co-?lead|president|vice president|vp|chair|co-?chair|captain|director|manager|officer|'
                        r'head|founder|chief|treasurer|secretary|coordinator|pm|e-?board|executive)\b', re.I)
ADVISOR_WORDS = re.compile(r'\b(advisor|adviser|prof\.?|professor|faculty)\b', re.I)
LEAD_HEADINGS = re.compile(r'lead|board|captain|leadership|chairs|officer|executive|president', re.I)
SKIP_HEADINGS = re.compile(r'alumni|former|inactive|graduated|past members|class of 20', re.I)
JOB = re.compile(r'(engineer|scientist|director|associate|analyst|consultant|professor|manager|resident|fellow|attending)\b.*,\s*[A-Z0-9]', re.I)
NOT_NAME = re.compile(r'\b(cornell|solutions|energy|university|subteam|project|members?|info|placeholder|portrait|headshot)\b', re.I)
JUNK = re.compile(r'^(linkedin|coffee chat|read more.*|bio|stats|close|boarding pass|crew member|member info|subteam|ticket no\.?|·|class of|'
                  r'learn more.*|website|github|email|resume|portfolio|x|\W+)$', re.I)
DETAIL = re.compile(r"^(major|minor|college|year|netid|hometown|clubs and activities|favorite plane|fun fact)\s*:", re.I)
PLACEHOLDER = re.compile(r'no_headshot|placeholder|default\.|blank\.|via\.placeholder|stand-in', re.I)
STUDY = re.compile(r"engineering|science|computer|economics|math|biology|physics|chemistry|studies|management|"
                   r"\bCS\b|\bECE\b|\bMAE\b|\bORIE\b|\bHADM\b|\bAEM\b|\bILR\b|architecture|statistics|design|"
                   r"policy|communication|sustainability|environment|nutrition|psychology|government|history", re.I)
NAME_RE = re.compile(r"^(prof\.\s+|dr\.\s+)?[A-Z][A-Za-z'’.\-]+(\s+\(?[A-Z][A-Za-z'’.\-]*\)?){0,4}$")


def clean(s):
    return re.sub(r'\s+', ' ', s or '').strip()


def is_name(s):
    s = clean(s)
    if not s or len(s) > 45 or any(c.isdigit() for c in s) or '@' in s:
        return False
    if ROLE_WORDS.search(s) or STUDY.search(s) or NOT_NAME.search(s):
        return False
    t = s.title() if s.isupper() else s
    return bool(NAME_RE.match(t)) and len(t.split()) >= 2


def titlecase(s):
    if s.isupper():
        return ' '.join(w if len(w) <= 2 and '.' in w else w.capitalize() for w in s.split())
    return s


def name_from_alt(alt):
    a = clean(alt)
    a = re.sub(r"(?i)^(headshot|portrait|photo|picture|image) of\s+", '', a)
    a = re.sub(r"(?i)'s profile picture$|\s+headshot$", '', a)
    a = re.split(r'\s+[-–]\s+|,\s', a)[0]
    return a if is_name(a) else ''


def class_year(lines):
    for s in lines:
        if re.search(r'(?i)former|\d{4}\s*[-–]\s*\d{2}', s):
            continue
        m = re.search(r"\b(20[2-3]\d)\b", s) or re.search(r"['’‘](2\d|3\d)\b", s)
        if m:
            y = m.group(1)
            return int(y) if len(y) == 4 else 2000 + int(y)
    return None


def parse_item(team, it, page_url):
    alt = clean(it.get('alt'))
    silly = bool(re.match(r'(?i)silly', alt))  # CUAir: a second photo whose card holds the role
    lines = [clean(s) for s in (it.get('lines') or [])]
    if not any(is_name(s) for s in lines):
        lines = [clean(s) for s in (it.get('after') or [])] or lines
    # "Name - Role" in one line (EWH)
    split = []
    for s in lines:
        if re.search(r'\s[-–]\s', s) and len(s) < 90:
            split += [p.strip() for p in re.split(r'\s[-–]\s', s, 1)]
        elif '·' in s:
            split += [p.strip() for p in s.split('·')]
        else:
            split.append(s)
    lines = [s for s in split if s and not JUNK.match(s) and '@' not in s and len(s) < 70]
    raw_lines = [clean(x) for x in (it.get('lines') or []) + (it.get('after') or [])[:3]]
    year = class_year(lines)
    details = [s for s in lines if DETAIL.match(s)]
    lines = [s for s in lines if not DETAIL.match(s)]

    name = name_from_alt(alt)
    if not name:
        name = next((s for s in lines if is_name(s)), '')
    if not name:
        return None
    name = titlecase(name)
    rest = [s for s in lines if titlecase(s) != name]
    role = next((s for s in rest if ROLE_WORDS.search(s) and not STUDY.search(s.replace('Team', ''))), '')
    if not role:
        role = next((s for s in rest if ROLE_WORDS.search(s) and len(s) < 45 and not re.search(r'\d', s)), '')
    if role and re.match(r'(?i)former\b', role):
        role = ''
    head = clean(it.get('head'))
    if head.isupper():
        head = head.title()
    if SKIP_HEADINGS.search(head) or re.search(r'(?i)\balumn', role):
        return None
    # alumni listed with a job ("Software Engineer, Google") or a finished degree
    if re.search(r',\s*[A-Z0-9]', role) and not re.search(r'(?i)\b(lead|member)\b', role) \
            or any(re.search(r'\b(B\.S\.|B\.A\.|M\.D\.|Ph\.D)', x) and 'Class of' not in x or JOB.search(x) for x in raw_lines):
        return None
    if is_name(head) or head.lower() in ('our team', 'team', 'meet the team', 'meet our team', 'our team.', 'subteams', 'meet our team!'):
        head = ''
    if not role and head:
        role = 'Lead' if LEAD_HEADINGS.search(head) else ''
    role = titlecase(role)
    if ADVISOR_WORDS.search(role) or name.lower().startswith('prof'):
        kind = 'advisor'
    elif LEAD_WORDS.search(role):
        kind = 'lead'
    else:
        kind = 'member'
    if LEAD_HEADINGS.search(head) or ROLE_WORDS.search(re.sub(r'(?i)\s*(sub)?team members$|^members$|\s*subteam$', '', head)) or len(head) > 50:
        subteam = ''
    else:
        subteam = re.sub(r'(?i)\s*(sub)?team members$|^members$|\s*subteam$', '', head)
    subteam = subteam or subteam_from_url(page_url)
    photo = '' if silly else (it.get('src') or '')
    return {
        'team': team, 'name': name, 'role': role or 'Member', 'roleType': kind,
        'subteam': subteam, 'classYear': year,
        'photo': '' if PLACEHOLDER.search(photo) else photo,
        'sourcePage': page_url,
    }


def subteam_from_url(u):
    pu = urllib.parse.urlparse(u)
    q = urllib.parse.parse_qs(pu.query).get('subteam')
    m = q[0] if q else None
    if not m:
        mm = re.search(r'/(?:subteam|team)/([a-z][\w-]+)/?$', pu.path + pu.fragment.replace('/team', '/team', 1))
        m = mm.group(1) if mm else None
    if not m and pu.netloc.endswith('cuair.org'):
        mm = re.search(r'/([a-z-]+)\.html$', pu.path)
        m = mm.group(1) if mm and mm.group(1) not in ('team', 'full-team-leads') else None
    if not m or m in ('leads', 'leadership'):
        return ''
    return ' '.join(w if w in ('and',) else w.capitalize() for w in m.replace('-', ' ').split())


def key(p):
    return hashlib.md5((p['team'] + '|' + re.sub(r'[^a-z]', '', p['name'].lower())).encode()).hexdigest()[:10]


def parse_raw(raw):
    found = {}
    for team, pages in raw.items():
        for spec, page in pages.items():
            for it in page.get('items', []):
                p = parse_item(team, it, page.get('url') or spec.split('|')[0])
                if not p:
                    continue
                k = key(p)
                old = found.get(k)
                if not old:
                    found[k] = p
                    continue
                # the same person on more than one page or card: keep the fullest details
                for f in ('photo', 'subteam', 'classYear'):
                    old[f] = old[f] or p[f]
                if old['role'] == 'Member' and p['role'] != 'Member':
                    old['role'], old['roleType'] = p['role'], p['roleType']
    return found


def main():
    dry = '--dry-run' in sys.argv
    raw = json.load(open(RAW))
    sites = json.load(open(SITES))
    today = datetime.date.today().isoformat()
    found = parse_raw(raw)
    db = json.load(open(OUT)) if os.path.exists(OUT) else {'people': []}
    people = {p['id']: p for p in db.get('people', [])}
    scraped_teams = {t for t, pages in raw.items() if any(pg.get('items') for pg in pages.values())}
    failed = sorted(t for t, pages in raw.items() if any('error' in pg for pg in pages.values()))
    # A team whose site suddenly lists far fewer people has probably moved its page.
    # Its missing people are left alone instead of being marked "left".
    before = {}
    for p in people.values():
        if p.get('source') == 'scraped' and p.get('status') == 'current':
            before[p['team']] = before.get(p['team'], 0) + 1
    now_found = {}
    for p in found.values():
        now_found[p['team']] = now_found.get(p['team'], 0) + 1
    shrunk = sorted(t for t, n in before.items() if n >= 6 and now_found.get(t, 0) < n / 2)
    added = updated = left = 0
    for k, p in found.items():
        cur = people.get(k)
        if cur and (cur.get('locked') or cur.get('source') == 'manual'):
            continue
        if cur:
            changed = any(cur.get(f) != p[f] for f in ('role', 'roleType', 'subteam', 'classYear', 'photo'))
            cur.update(p, lastSeen=today)
            if cur.get('status') != 'graduated':  # a graduate marked by hand stays graduated
                cur['status'] = 'current'
            updated += changed
        else:
            people[k] = {'id': k, **p, 'status': 'current', 'source': 'scraped', 'added': today, 'lastSeen': today}
            added += 1
    for k, cur in people.items():
        if cur.get('source') == 'scraped' and not cur.get('locked') and k not in found \
                and cur['team'] in scraped_teams and cur['team'] not in shrunk and cur.get('status') == 'current':
            cur['status'] = 'left'
            left += 1
    pruned = 0
    if '--prune' in sys.argv:
        now = datetime.date.today()
        last_grad = now.year if now.month >= 6 else now.year - 1
        gone = [k for k, p in people.items()
                if p.get('status') in ('left', 'graduated') or (p.get('classYear') or 9999) <= last_grad]
        for k in gone:
            del people[k]
        pruned = len(gone)
    order = {t: i for i, t in enumerate(sites)}
    kind = {'lead': 0, 'advisor': 2, 'member': 1}
    out = {
        'updated': today,
        'lastRun': {
            'date': datetime.datetime.now(datetime.timezone.utc).isoformat(timespec='seconds'),
            'runId': os.environ.get('RUN_ID', ''),
            'found': len(found), 'added': added, 'changed': updated, 'left': left,
            'failedTeams': failed, 'shrunkTeams': shrunk,
        },
        'teams': sites,
        'people': sorted(people.values(), key=lambda p: (order.get(p['team'], 999), kind.get(p['roleType'], 3), p['name'])),
    }
    print(f'{len(found)} people on team sites; {added} new, {updated} changed, {left} no longer listed (marked "left")'
          + (f', {pruned} deleted by --prune' if pruned else ''))
    if failed:
        print('could not load (kept as they were):', ', '.join(failed))
    if shrunk:
        print('far fewer people than before, nobody marked left (check targets.json):', ', '.join(shrunk))
    if not dry:
        os.makedirs(os.path.dirname(OUT), exist_ok=True)
        with open(OUT, 'w') as f:
            json.dump(out, f, indent=1, ensure_ascii=False)
            f.write('\n')
        print('wrote', os.path.relpath(OUT, ROOT))


if __name__ == '__main__':
    main()
