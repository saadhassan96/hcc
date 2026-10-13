#!/usr/bin/env python3
"""Build the HCC Group website from content/hcc-content.xlsx.

Usage:  python3 build.py
Needs:  openpyxl  (pip install openpyxl)

Reads every tab of the spreadsheet, fills templates/index.template.html, and writes
index.html. Nothing else is generated, so the CSS, JS and
images can be edited by hand as usual.
"""
import datetime as dt
import html
import math
import os
import re
import sys

try:
    import openpyxl
except ImportError:
    sys.exit('This script needs openpyxl. Install it with:  pip install openpyxl')

ROOT = os.path.dirname(os.path.abspath(__file__))
SHEET = os.path.join(ROOT, 'content', 'hcc-content.xlsx')
TEMPLATE = os.path.join(ROOT, 'templates', 'index.template.html')
OUT = os.path.join(ROOT, 'index.html')
TODAY = dt.date.today()
MONTHS = 'Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec'.split()


# --------------------------------------------------------------------------- reading
def read_tab(wb, name):
    if name not in wb.sheetnames:
        return []
    ws = wb[name]
    it = ws.iter_rows(values_only=True)
    head = [str(h).strip() if h else None for h in next(it)]
    out = []
    for row in it:
        if not any(c not in (None, '') for c in row):
            continue
        rec = {}
        for h, v in zip(head, row):
            if not h:
                continue
            if isinstance(v, str):
                v = v.strip()
            if isinstance(v, float) and v.is_integer():
                v = int(v)
            rec[h] = v if v not in ('',) else None
        out.append(rec)
    return out


def esc(s):
    return html.escape(str(s), quote=True) if s is not None else ''


def slug(s):
    return re.sub(r'[^a-z0-9]+', '-', str(s).lower()).strip('-')


def split(v, sep=';'):
    return [x.strip() for x in str(v).split(sep) if x.strip()] if v else []


def is_yes(v):
    return str(v or '').strip().lower() in ('yes', 'y', 'true', '1', 'x')


def money(v):
    if v is None:
        return ''
    if isinstance(v, (int, float)):
        return '${:,.0f}'.format(v)
    return str(v)


def date_key(v):
    if isinstance(v, dt.datetime):
        return (v.year, v.month, v.day)
    if isinstance(v, dt.date):
        return (v.year, v.month, v.day)
    if isinstance(v, int):
        return (v, 6, 15)
    m = re.match(r'(\d{4})(?:-(\d{1,2}))?', str(v or ''))
    return (int(m.group(1)), int(m.group(2) or 6), 15) if m else (0, 0, 0)


def date_label(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return '{} {}'.format(MONTHS[v.month - 1], v.year), v.strftime('%Y-%m-%d')
    y, m, _ = date_key(v)
    if isinstance(v, int) or re.fullmatch(r'\d{4}', str(v or '')):
        return str(y), str(y)
    return '{} {}'.format(MONTHS[m - 1], y), '{}-{:02d}'.format(y, m)


# --------------------------------------------------------------------------- people
class People:
    """Knows every group member so names can be linked and bolded anywhere."""

    def __init__(self, faculty, students, coauthors):
        self.faculty = sorted(faculty, key=lambda f: (f.get('Order') or 99, f['Name']))
        self.students = students
        self.first = {}            # 'Saad' -> faculty record
        self.links = {}            # name -> url
        self.members = set()
        for f in self.faculty:
            self.first[f['Name'].split()[0]] = f
        for p in self.faculty + self.students:
            name = p['Name']
            self.members.add(name)
            url = (p.get('Webpage URL') or p.get('Lab webpage URL') or p.get('LinkedIn URL') or p.get('Scholar URL')
                   or p.get('ResearchGate URL'))
            for variant in self.variants(name):
                self.members.add(variant)
                if url:
                    self.links.setdefault(variant, url)
        for c in coauthors:
            if c.get('Name') and c.get('URL'):
                self.links.setdefault(c['Name'], c['URL'])

    @staticmethod
    def variants(name):
        """'Hoang Dieu Linh (Evelyn) Nguyen' also matches 'Evelyn Nguyen';
        'Yu (Demi) Qin' also matches 'Yu Qin' and 'Demi Qin'."""
        out = {name}
        m = re.match(r'(.*?)\s*\((.+?)\)\s*(.*)', name)
        if m:
            before, nick, last = m.groups()
            out |= {f'{nick} {last}'.strip(), f'{before} {last}'.strip()}
        out.add(re.sub(r'[“”"]', '', name))
        return out

    def advisor_names(self, v):
        names = []
        for a in split(str(v or '').replace(' and ', ','), ','):
            f = self.first.get(a)
            names.append(f['Name'] if f else a)
        return names

    def author(self, raw):
        star = raw.endswith('*')
        name = raw.rstrip('*').strip()
        url = self.links.get(name)
        text = esc(name) + ('<sup>*</sup>' if star else '')
        if url:
            text = f'<a href="{esc(url)}">{text}</a>'
        if name in self.members:
            text = f'<b>{text}</b>'
        return text

    def link_names(self, text_html, skip=()):
        """Link the first mention of each known person in already-escaped text."""
        names = sorted((n for n in self.links if n not in skip and len(n.split()) >= 2), key=len, reverse=True)
        tokens = {}
        for i, n in enumerate(names):
            en = esc(n)
            if en in text_html:
                tok = f'\x00{i}\x00'
                tokens[tok] = f'<a href="{esc(self.links[n])}">{en}</a>'
                text_html = text_html.replace(en, tok, 1)
        for tok, a in tokens.items():
            text_html = text_html.replace(tok, a)
        return text_html


# --------------------------------------------------------------------------- icons
ICON_LABEL = {'web': 'website', 'group': 'group website', 'scholar': 'Google Scholar profile', 'linkedin': 'LinkedIn profile',
              'rg': 'ResearchGate profile'}


def icon(name):
    return f'<svg class="icon" aria-hidden="true" focusable="false"><use href="#i-{name}"></use></svg>'


def avatar(p, cls='avatar'):
    initials = ''.join(w[0] for w in re.sub(r'\(.*?\)', '', p['Name']).split() if w[0].isalpha())[:2]
    src = p.get('Photo') or ''
    img = ''
    if src:
        alt = esc(p.get('Photo alt text') or '')
        img = (f'<img src="{esc(src)}" alt="{alt}" width="320" height="320" loading="lazy" decoding="async" '
               f'onerror="this.remove()">')
    return f'<div class="{cls}"><span aria-hidden="true">{esc(initials)}</span>{img}</div>'


# --------------------------------------------------------------------------- sections
def render_faces(people):
    """Header photo grid: everyone currently in the group who has a photo, in People-section order.
    Decorative (the People section has the same photos with names), so it is hidden from screen readers."""
    cat_order = {'1 - PhD': 0, 'Postdoc': 1, "Master's": 2, 'Undergraduate': 3, 'Capstone': 4, 'High school': 5}
    cur = [s for s in people.students if (s.get('Status') or '').lower() == 'current']
    cur.sort(key=lambda s: (cat_order.get(s.get('Degree / category'), 9), s.get('Expected graduation year') or 9999,
                            re.sub(r'\(.*?\)', '', s['Name']).split()[-1].lower()))
    everyone = [p for p in people.faculty + cur if p.get('Photo')]
    return '\n'.join(f'<li><img src="{esc(p["Photo"])}" alt="" width="160" height="160" decoding="async" title="{esc(p["Name"])}"></li>'
                     for p in everyone)


def render_codirector_chips(people):
    out = []
    for f in people.faculty:
        if f.get('Group role') != 'Co-director':
            continue
        out.append(f'<li><a href="#{slug(f["Name"])}">{avatar(f, "avatar sm")}<span>{esc(f["Name"])}</span></a></li>')
    return '\n'.join(out)


def render_news(news, people):
    news = sorted(news, key=lambda n: date_key(n.get('Date')), reverse=True)
    featured = [n for n in news if is_yes(n.get('Highlight?'))][:6]
    rest = [n for n in news if n not in featured]

    def body(n):
        text = n.get('Description') or n.get('Title') or ''
        return people.link_names(esc(text))

    feat_html = []
    for i, n in enumerate(featured):
        label, iso = date_label(n.get('Date'))
        title = esc(n.get('Title') or '')
        desc = n.get('Description') or ''
        more = (f'<p class="more"><a href="{esc(n["URL"])}" aria-describedby="feat-{i}">Read more</a></p>' if n.get('URL') else '')
        extra = ''
        if desc and desc != n.get('Title') and desc.startswith(n.get('Title') or '~'):
            extra = f'<p>{people.link_names(esc(desc[len(n["Title"]):].strip()))}</p>'
        elif desc and desc != n.get('Title'):
            extra = f'<p>{people.link_names(esc(desc))}</p>'
        feat_html.append(f'''<article class="feature">
          <p class="date"><time datetime="{iso}">{label}</time></p>
          <h3 id="feat-{i}">{people.link_names(title)}</h3>
          {extra}
          {more}
        </article>''')

    items = []
    for i, n in enumerate(rest):
        label, iso = date_label(n.get('Date'))
        link = (f' <a class="inline-more" href="{esc(n["URL"])}" aria-label="Details: {esc((n.get("Title") or "")[:90])}">Details</a>'
                if n.get('URL') else '')
        hidden = ' data-extra' if i >= 8 else ''
        items.append(f'<li{hidden}><time datetime="{iso}">{label}</time><p>{body(n)}{link}</p></li>')
    toggle = ''
    if len(rest) > 8:
        toggle = (f'<button type="button" class="more-toggle" data-toggle-list="news-list" aria-expanded="false" '
                  f'data-more="Show all {len(rest)} news items" data-less="Show fewer news items">Show all {len(rest)} news items</button>')
    return '\n'.join(feat_html), '\n'.join(items), toggle


def funder_short(f):
    f = str(f or '')
    if f.lower().startswith('tulane'):
        return 'Tulane University'
    return re.split(r' \(| — | - ', f)[0].strip()


def render_projects(projects, papers, grants, people):
    out = []
    for p in sorted(projects, key=lambda p: str(p.get('Title') or '').lower()):
        key = p.get('Key')
        n = sum(1 for x in papers if x.get('Project') == key)
        leads = [people.first[l]['Name'] if l in people.first else l for l in split(p.get('Leads'))]
        lead_html = ' and '.join(f'<a href="#{slug(l)}">{esc(l)}</a>' for l in leads)
        funders = []
        for g in grants:
            if g.get('Project') == key:
                s = funder_short(g.get('Funder'))
                if s not in funders:
                    funders.append(s)
        kw = ', '.join(esc(k) for k in split(p.get('Keywords')))
        pubs_btn = (f'<a class="see-pubs" href="#all-pubs" data-filter-project="{esc(key)}">'
                    f'{n} publication{"s" if n != 1 else ""}</a>' if n else '')
        fund = f'<p class="funding"><span class="k">Funding</span> <a href="#grants">{esc(", ".join(funders))}</a></p>' if funders else ''
        out.append(f'''<article class="project" id="project-{esc(key)}">
          <h3>{esc(p.get("Title"))}</h3>
          <p class="leads">{lead_html}</p>
          <p class="summary">{esc(p.get("Summary"))}</p>
          <p class="keywords">{kw}</p>
          <div class="project-foot">{pubs_btn}{fund}</div>
        </article>''')
    return '\n'.join(out)


LINK_COLS = [('PDF', 'PDF'), ('DL', None), ('Code', 'Code'), ('Slides', 'Slides'), ('Poster', 'Poster'), ('Video', 'Video'),
             ('Dataset', 'Dataset'), ('Supplementary Material', 'Supplement'), ('Project webpage', 'Project page')]


def dl_label(url):
    u = url.lower()
    if 'dl.acm.org' in u:
        return 'ACM DL'
    if 'arxiv.org' in u:
        return 'arXiv'
    if 'doi.org' in u or 'ieee' in u:
        return 'DOI'
    if 'thecvf.com' in u:
        return 'CVF'
    return 'Publisher'


def pub_links(p, tid):
    links = []
    for col, label in LINK_COLS:
        for i, url in enumerate(split(p.get(col))):
            lab = label or dl_label(url)
            if col == 'PDF' and 'arxiv.org' in url:
                lab = 'arXiv'
            if i:
                lab = f'{lab} {i + 1}'
            links.append(f'<a href="{esc(url)}" aria-describedby="{tid}">{esc(lab)}</a>')
    for part in split(p.get('Other links')):
        m = re.match(r'(.+?):\s*(https?://\S+)$', part)
        if m:
            links.append(f'<a href="{esc(m.group(2))}" aria-describedby="{tid}">{esc(m.group(1))}</a>')
    return ''.join(links)


def venue_label(p):
    v = f'{p.get("Venue short") or ""} {p.get("Year") or ""}'.strip()
    t = p.get('Type') or ''
    if t and t not in ('Full Paper', 'Journal Article'):
        v += f' {t.lower()}' if t != 'Preprint' else ' preprint'
    return v


def leads_of(p, people):
    return [people.first[n]['Name'].split()[-1].lower() for n in split(p.get('PI/Student Lead'), ',')
            if n in people.first]


def render_highlights(papers, people):
    hl = [p for p in papers if is_yes(p.get('Highlight?'))]
    hl.sort(key=lambda p: (-(p.get('Year') or 0), p['Title'].lower()))
    out = []
    for i, p in enumerate(hl):
        tid = f'hl-{i}'
        thumb = p.get('Thumbnail URL')
        title_url = p.get('DL') or p.get('PDF') or p.get('Project webpage')
        title = esc(p['Title'])
        if title_url:
            title = f'<a href="{esc(split(title_url)[0])}">{title}</a>'
        if thumb:
            fig = f'<div class="thumb"><img src="{esc(thumb)}" alt="" decoding="async"></div>'
        else:
            fig = f'<div class="thumb placeholder" aria-hidden="true"><span>{esc(p.get("Venue short") or "")}<br>{esc(p.get("Year") or "")}</span></div>'
        award = f'<p class="award">{esc(p["Award"])}</p>' if p.get('Award') else ''
        authors = ', '.join(people.author(a) for a in split(p.get('Authors')))
        out.append(f'''<article class="hl">
          {fig}
          <div class="hl-text">
            <p class="venue">{esc(venue_label(p))}</p>
            <h3 id="{tid}">{title}</h3>
            <p class="authors">{authors}</p>
            {award}
            <p class="links">{pub_links(p, tid)}</p>
          </div>
        </article>''')
    return '\n'.join(out)


def render_publications(papers, projects, people):
    papers = sorted(papers, key=lambda p: (-(p.get('Year') or 0), p['Title'].lower()))
    by_year = {}
    for p in papers:
        by_year.setdefault(p.get('Year'), []).append(p)
    out = []
    n = 0
    for year, ps in by_year.items():
        out.append(f'<li class="year" data-year="{year}"><h3>{year}</h3><ol class="pubs">')
        for p in ps:
            n += 1
            tid = f'pub-{n}'
            title_url = p.get('DL') or p.get('PDF') or p.get('Project webpage')
            title = esc(p['Title'])
            if title_url:
                title = f'<a href="{esc(split(title_url)[0])}">{title}</a>'
            authors = ', '.join(people.author(a) for a in split(p.get('Authors')))
            award = f' <span class="award">{esc(p["Award"])}</span>' if p.get('Award') else ''
            leads = ' '.join(leads_of(p, people))
            out.append(f'''<li class="pub" data-lead="{esc(leads)}" data-project="{esc(p.get("Project") or "")}">
              <p class="title" id="{tid}">{title}</p>
              <p class="authors">{authors}</p>
              <p class="meta"><span class="venue">{esc(venue_label(p))}</span>{award}</p>
              <p class="links">{pub_links(p, tid)}</p>
            </li>''')
        out.append('</ol></li>')
    def chip(attr, value, label, count, pressed=False):
        return f'<option value="{esc(value)}">{esc(label)} ({count})</option>'
    lead_buttons = chip('lead', '', 'All co-directors', len(papers), True) + ''.join(
        chip('lead', f['Name'].split()[-1].lower(), f['Name'],
             sum(1 for p in papers if f['Name'].split()[-1].lower() in leads_of(p, people)))
        for f in people.faculty if f.get('Group role') == 'Co-director')
    area_buttons = chip('project', '', 'All areas', len(papers), True) + ''.join(
        chip('project', p['Key'], p.get('Short title') or p['Title'], sum(1 for x in papers if x.get('Project') == p['Key']))
        for p in sorted(projects, key=lambda p: str(p.get('Short title') or p.get('Title') or '').lower()))
    return '\n'.join(out), lead_buttons, area_buttons, len(papers)


def person_links(p, kinds):
    cols = {'web': 'Webpage URL', 'group': 'Lab webpage URL', 'scholar': 'Scholar URL', 'linkedin': 'LinkedIn URL',
            'rg': 'ResearchGate URL'}
    out = []
    for k in kinds:
        url = p.get(cols[k])
        if url:
            out.append(f'<a href="{esc(url)}" aria-label="{esc(p["Name"])}, {ICON_LABEL[k]}" title="{ICON_LABEL[k].capitalize()}">{icon(k)}</a>')
    return f'<p class="plinks">{"".join(out)}</p>' if out else ''


def render_people(people):
    fac = [f for f in people.faculty if f.get('Group role') == 'Co-director']
    staff = [f for f in people.faculty if f.get('Group role') != 'Co-director']
    co = []
    for f in fac:
        cols = {'web': 'Website', 'group': 'Lab', 'scholar': 'Scholar', 'linkedin': 'LinkedIn', 'rg': 'ResearchGate'}
        colmap = {'web': 'Webpage URL', 'group': 'Lab webpage URL', 'scholar': 'Scholar URL', 'linkedin': 'LinkedIn URL',
                  'rg': 'ResearchGate URL'}
        links = ''.join(f'<a href="{esc(f[colmap[k]])}" aria-label="{esc(f["Name"])}, {ICON_LABEL[k]}">{icon(k)}<span>{v}</span></a>'
                        for k, v in cols.items() if f.get(colmap[k]))
        name = esc(f['Name'])
        lab = ''
        if f.get('Lab name'):
            lab_name = esc(f['Lab name'])
            lab = f'<p class="lab">{"<a href=" + chr(34) + esc(f["Lab webpage URL"]) + chr(34) + ">" + lab_name + "</a>" if f.get("Lab webpage URL") else lab_name}</p>'
        interests = f'<p class="interests">{esc(", ".join(split(f.get("Interests"))))}</p>' if f.get('Interests') else ''
        email = f'<p class="email"><a href="mailto:{esc(f["Email"])}">{esc(f["Email"])}</a></p>' if f.get('Email') else ''
        co.append(f'''<article class="codirector" id="{slug(f["Name"])}">
          <div class="cd-head">
            {avatar(f, "avatar lg")}
            <div>
              <h4>{name}</h4>
              <p class="role">{esc(f.get("Rank / title") or "")} of {esc(f.get("Department") or "Computer Science")}</p>
              {lab}
            </div>
          </div>
          {interests}{email}
          <p class="flinks">{links}</p>
        </article>''')
    staff_html = []
    for s in staff:
        staff_html.append(f'''<li class="card">{avatar(s)}<div><p class="name">{esc(s["Name"])}</p>
          <p class="sub">{esc(s.get("Rank / title") or "")}</p>{person_links(s, ["web", "scholar", "linkedin", "rg"])}</div></li>''')

    cat_order = {'1 - PhD': 0, 'Postdoc': 1, "Master's": 2, 'Undergraduate': 3, 'Capstone': 4, 'High school': 5}
    cur = [s for s in people.students if (s.get('Status') or '').lower() == 'current']
    former = [s for s in people.students if (s.get('Status') or '').lower() != 'current']

    def stu_card(s, alum=False):
        adv = people.advisor_names(s.get('Advisor(s)'))
        cat = s.get('Degree / category') or ''
        if alum:
            label = 'Postdoc' if cat == 'Postdoc' else 'PhD'
            years = ''
            if s.get('Start year') and s.get('Last year in lab'):
                years = f', {s["Start year"]}–{s["Last year in lab"]}'
            elif s.get('Last year in lab'):
                years = f', {s["Last year in lab"]}'
            sub = f'{label}{years}'
            if s.get('Notes') and cat != 'Postdoc' and 'PhD' in str(s.get('Notes')):
                sub = esc(s['Notes']) + (years and esc(years))
            else:
                sub = esc(sub)
            pos = s.get('Next position') or ''
            # "Research Scientist, NLR" reads as a role; "FDA" reads as a place.
            nxt = f'<p class="next">Now {"" if "," in pos else "at "}{esc(pos)}</p>' if pos else ''
            adv_html = f'<p class="sub">With {esc(" and ".join(adv))}</p>' if adv else ''
            return f'''<li class="card">{avatar(s)}<div><p class="name">{esc(s["Name"])}</p><p class="sub">{sub}</p>{adv_html}{nxt}
              {person_links(s, ["web", "scholar", "linkedin", "rg"])}</div></li>'''
        kind = {'1 - PhD': 'PhD student', 'Undergraduate': 'Undergraduate researcher', "Master's": "Master's student",
                'Capstone': 'Capstone student', 'High school': 'High school researcher', 'Postdoc': 'Postdoc'}.get(cat, cat)
        adv_html = f'<p class="sub">Advised by {esc(" and ".join(adv))}</p>' if adv else ''
        exp = f'<p class="sub">Expected {esc(s["Expected graduation year"])}</p>' if s.get('Expected graduation year') else ''
        return f'''<li class="card">{avatar(s)}<div><p class="name">{esc(s["Name"])}</p><p class="kind">{esc(kind)}</p>{adv_html}{exp}
          {person_links(s, ["web", "scholar", "linkedin", "rg"])}</div></li>'''

    def last_name(s):
        return re.sub(r'\(.*?\)', '', s['Name']).split()[-1].lower()

    def sort_cur(s):
        # Most senior first (earliest expected graduation), then by last name.
        return (cat_order.get(s.get('Degree / category'), 9), s.get('Expected graduation year') or 9999,
                last_name(s), s['Name'])
    cur.sort(key=sort_cur)
    groups = []
    phd = [s for s in cur if s.get('Degree / category') == '1 - PhD']
    other_cur = [s for s in cur if s.get('Degree / category') != '1 - PhD']
    if phd:
        groups.append(('PhD students', ''.join(stu_card(s) for s in phd)))
    # One heading per kind of student, in the usual order (master's, undergraduate, capstone, high school).
    titles = {"Master's": 'Master’s students', 'Undergraduate': 'Undergraduate researchers',
              'Capstone': 'Capstone students', 'High school': 'High school researchers', 'Postdoc': 'Postdocs'}
    for cat in sorted({s.get('Degree / category') for s in other_cur}, key=lambda c: cat_order.get(c, 9)):
        members = [s for s in other_cur if s.get('Degree / category') == cat]
        groups.append((titles.get(cat, str(cat)), ''.join(stu_card(s) for s in members)))

    alum_cards = [s for s in former if s.get('Degree / category') in ('1 - PhD', 'Postdoc')]
    alum_cards.sort(key=lambda s: (cat_order.get(s.get('Degree / category'), 9), -(s.get('Last year in lab') or 0), s['Name']))
    others = {}
    for s in former:
        c = s.get('Degree / category')
        if c in ('1 - PhD', 'Postdoc'):
            continue
        others.setdefault(c, []).append(s)
    other_html = []
    labels = {"Master's": 'Master’s', 'Undergraduate': 'Undergraduate', 'Capstone': 'Capstone', 'High school': 'High school'}
    for c in sorted(others, key=lambda c: cat_order.get(c, 9)):
        ppl = sorted(others[c], key=lambda s: (-(s.get('Last year in lab') or 0), s['Name']))
        names = []
        for s in ppl:
            url = s.get('Webpage URL') or s.get('LinkedIn URL') or s.get('Scholar URL')
            nm = esc(s['Name'])
            if url:
                nm = f'<a href="{esc(url)}">{nm}</a>'
            yr = f' <span class="yr">{s["Last year in lab"]}</span>' if s.get('Last year in lab') else ''
            names.append(f'<li>{nm}{yr}</li>')
        other_html.append(f'<div class="alum-group"><h5>{labels.get(c, esc(c))} <span class="count">{len(ppl)}</span></h5><ul>{"".join(names)}</ul></div>')

    html_out = [f'<h3 class="sub-h">Co-directors</h3><div class="codirectors">{"".join(co)}</div>']
    if staff_html:
        html_out.append(f'<h3 class="sub-h">Research staff</h3><ul class="cards">{"".join(staff_html)}</ul>')
    for title, cards in groups:
        html_out.append(f'<h3 class="sub-h">{title}</h3><ul class="cards">{cards}</ul>')
    html_out.append('<h3 class="sub-h alumni-h">Alumni</h3>')
    if alum_cards:
        html_out.append(f'<h4 class="sub-h2">PhD and postdoctoral alumni</h4><ul class="cards">{"".join(stu_card(s, True) for s in alum_cards)}</ul>')
    if other_html:
        html_out.append(f'<h4 class="sub-h2">Student alumni</h4><div class="alum-groups">{"".join(other_html)}</div>')
    return '\n'.join(html_out)


def dollars(v):
    """A number from a cell that may hold 449868 or text like 'Up to $22.9M' or '$416K in FY2025'."""
    if isinstance(v, (int, float)):
        return float(v)
    m = re.search(r'\$?\s*([\d,.]+)\s*([KkMm])?', str(v or ''))
    if not m:
        return 0.0
    n = float(m.group(1).replace(',', ''))
    return n * {'k': 1e3, 'm': 1e6}.get((m.group(2) or '').lower(), 1)


def grant_value(g):
    """The amount shown on the card: the grant's own amount, or the team total when that is all we have."""
    return dollars(g.get('Amount')) or dollars(g.get('Collaborative total'))


def render_grants(grants, people, minimum=10000, top=5):
    """External awards at or above `minimum`, largest first. The first `top` show; the rest sit
    behind a Show more button. Tulane internal awards are left off the site."""
    year = TODAY.year
    shown = [g for g in grants if not str(g.get('Funder') or '').lower().startswith('tulane')
             and grant_value(g) >= minimum]
    shown.sort(key=lambda g: (-grant_value(g), str(g.get('Title') or '').lower()))

    def active(g):
        end = g.get('End year')
        return (end or 9999) >= year and (g.get('Start year') or 0) <= year

    def who(g):
        out = []
        for n in split(g.get('PI'), ','):
            f = people.first.get(n)
            out.append(f'<a href="#{slug(f["Name"])}">{esc(f["Name"])}</a>' if f else esc(n))
        return ' and '.join(out)

    def years(g):
        s, e = g.get('Start year'), g.get('End year')
        if s and e and s != e:
            return f'{s}–{e}'
        if s and not e:
            return f'Since {s}'
        return str(s or e or '')

    cards = []
    for i, g in enumerate(shown):
        amt = money(g.get('Amount'))
        tot = g.get('Collaborative total')
        amount_html = ''
        if amt and tot:
            amount_html = f'<p class="amount">{esc(amt)}</p><p class="amount-note">Part of a {esc(money(tot))} collaborative award</p>'
        elif amt:
            amount_html = f'<p class="amount">{esc(amt)}</p>'
        elif tot:
            amount_html = f'<p class="amount">{esc(money(tot))}</p><p class="amount-note">Total for the full team</p>'
        status = '<span class="status on">Active</span>' if active(g) else '<span class="status">Completed</span>'
        link = f'<a class="award-link" href="{esc(g["URL"])}" aria-describedby="gr-{i}">Award details</a>' if g.get('URL') else ''
        extra = ' data-extra' if i >= top else ''
        cards.append(f'''<li class="grant"{extra}>
          <p class="funder">{esc(g.get("Funder"))}</p>
          <h3 id="gr-{i}">{esc(g.get("Title"))}</h3>
          <p class="gwho">{who(g)}, {esc(g.get("Role") or "")}</p>
          <p class="gmeta">{status}<span>{years(g)}</span></p>
          {amount_html}
          {link}
        </li>''')
    toggle = ''
    if len(shown) > top:
        more = len(shown) - top
        toggle = (f'<button type="button" class="more-toggle" data-toggle-list="grant-list" aria-expanded="false" '
                  f'data-more="Show {more} more grants" data-less="Show fewer grants">Show {more} more grants</button>')
    funders = []
    for g in shown:
        f = funder_short(g.get('Funder'))
        f = {'CRA / NSF Computing Innovation Fellows': 'the Computing Research Association',
             'U.S. Department of Energy': 'the U.S. Department of Energy'}.get(f, f)
        if f not in funders:
            funders.append(f)
    funders.append('Tulane University')
    funder_text = ', '.join(funders[:-1]) + ', and ' + funders[-1] if len(funders) > 2 else ' and '.join(funders)
    n_active = sum(1 for g in shown if active(g))
    return '\n'.join(cards), toggle, funder_text, n_active, len(shown)


# --------------------------------------------------------------------------- main
def main():
    wb = openpyxl.load_workbook(SHEET, data_only=True)
    settings = {r['Setting']: r.get('Value') for r in read_tab(wb, 'Settings') if r.get('Setting')}
    faculty = read_tab(wb, 'FacultyStaff')
    students = read_tab(wb, 'Students')
    projects = read_tab(wb, 'Projects')
    papers = read_tab(wb, 'Papers')
    grants = read_tab(wb, 'Grants')
    news = read_tab(wb, 'News')
    coauthors = read_tab(wb, 'Coauthors')
    people = People(faculty, students, coauthors)

    feat, news_items, news_toggle = render_news(news, people)
    pubs, lead_buttons, area_buttons, n_pubs = render_publications(papers, projects, people)
    grant_cards, grant_toggle, funder_text, n_active, n_ext = render_grants(
        grants, people, minimum=settings.get('Grant minimum') or 10000, top=settings.get('Grants shown before Show more') or 5)

    labs = []
    for f in people.faculty:
        if f.get('Group role') != 'Co-director':
            continue
        # One line per co-director: their lab if they have one, otherwise their own site.
        if f.get('Lab name'):
            url = f.get('Lab webpage URL') or f.get('Webpage URL')
            name = f'<a href="{esc(url)}">{esc(f["Lab name"])}</a>' if url else esc(f['Lab name'])
            labs.append(f'<li>{name} <span>{esc(f["Name"])}</span></li>')
        else:
            url = f.get('Webpage URL')
            name = f'<a href="{esc(url)}">{esc(f["Name"])}</a>' if url else esc(f['Name'])
            labs.append(f'<li>{name} <span>{esc(", ".join(split(f.get("Interests"))[:2]))}</span></li>')
    emails = ''.join(f'<li><a href="mailto:{esc(f["Email"])}">{esc(f["Name"])}</a></li>'
                     for f in people.faculty if f.get('Group role') == 'Co-director' and f.get('Email'))

    values = {
        'GROUP_NAME': esc(settings.get('Group name', 'Human-Centered Computing Group')),
        'SHORT_NAME': esc(settings.get('Short name', 'HCC Group')),
        'INSTITUTION': esc(settings.get('Institution', 'Tulane University')),
        'TAGLINE': esc(settings.get('Tagline', '')),
        'ABOUT': people.link_names(esc(settings.get('About', ''))),
        'ADDRESS': esc(settings.get('Address', '')),
        'DEPT_URL': esc(settings.get('Department URL', '')),
        'JOIN_TEXT': esc(settings.get('Join text', '')),
        'PHD_URL': esc(settings.get('PhD program URL', '') or settings.get('Department URL', '')),
        'CODIRECTOR_CHIPS': render_codirector_chips(people),
        'FACES': render_faces(people),
        'NEWS_FEATURED': feat,
        'NEWS_ITEMS': news_items,
        'NEWS_TOGGLE': news_toggle,
        'PROJECTS': render_projects(projects, papers, grants, people),
        'N_PROJECTS': str(len(projects)),
        'HIGHLIGHTS': render_highlights(papers, people),
        'PUBLICATIONS': pubs,
        'LEAD_BUTTONS': lead_buttons,
        'AREA_BUTTONS': area_buttons,
        'N_PUBS': str(n_pubs),
        'PEOPLE': render_people(people),
        'GRANT_CARDS': grant_cards,
        'GRANT_TOGGLE': grant_toggle,
        'FUNDERS': esc(funder_text),
        'N_ACTIVE': str(n_active),
        'N_GRANTS': str(n_ext),
        'LABS': '\n'.join(labs),
        'EMAILS': emails,
        'UPDATED_ISO': TODAY.isoformat(),
        'UPDATED': f'{TODAY.strftime("%B")} {TODAY.day}, {TODAY.year}',
        'YEAR': str(TODAY.year),
        'SITE_URL': esc(settings.get('Site URL') or ''),
    }
    with open(TEMPLATE, encoding='utf8') as fh:
        page = fh.read()
    page = re.sub(r'\{\{\s*(\w+)\s*\}\}', lambda m: values.get(m.group(1), m.group(0)), page)
    left = re.findall(r'\{\{\s*\w+\s*\}\}', page)
    if left:
        print('Warning: unfilled placeholders:', ', '.join(sorted(set(left))))
    with open(OUT, 'w', encoding='utf8') as fh:
        fh.write(page)
    print(f'Built index.html: {len(projects)} projects, {n_pubs} publications, {len(grants)} grants, '
          f'{len(news)} news items, {len(faculty) + len(students)} people.')


if __name__ == '__main__':
    main()
