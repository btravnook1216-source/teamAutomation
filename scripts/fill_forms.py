#!/usr/bin/env python3
"""Fill the four Travnook Team Leader PDF forms from data/2026-10.json.

The original blank PDFs in forms/ are never changed. Text is written on top of them,
so the layout stays exactly as the company forms.

  python3 scripts/fill_forms.py form1 2026-10-02     -> reports/daily/Form1_Daily_Report_2026-10-02.pdf
  python3 scripts/fill_forms.py form2 1              -> reports/weekly/Form2_Case_Log_Week1.pdf
  python3 scripts/fill_forms.py form3 1              -> reports/weekly/Form3_Weekly_Report_Week1.pdf
  python3 scripts/fill_forms.py week 1               -> Form 2 + Form 3 for the week
  python3 scripts/fill_forms.py day-all              -> Form 1 for every day in the data file
  python3 scripts/fill_forms.py combine              -> reports/Travnook_All_Reports_Oct2026.pdf (every report, with bookmarks)
"""
import datetime as dt
import json
import os
import sys

import pymupdf

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, 'data', '2026-10.json')
INK = (0.08, 0.16, 0.30)


def load():
    with open(DATA) as f:
        return json.load(f)


def fmt_date(s, style='long'):
    d = dt.date.fromisoformat(s)
    return d.strftime('%d %b %Y') if style == 'long' else f'{d.day} {d.strftime("%b")}'


def text_width(s, size, font='helv'):
    return pymupdf.get_text_length(s, fontname=font, fontsize=size)


def put(page, rect, text, size=9, minsize=4.8, font='helv', align='left', pad=3):
    """Write text inside rect; one line if it fits, otherwise wrapped and shrunk to fit."""
    if text is None or text == '':
        return
    text = str(text)
    r = pymupdf.Rect(rect)
    w = r.width - 2 * pad
    s = size
    while s >= minsize:
        if text_width(text, s) <= w:
            x = r.x0 + pad
            tw = text_width(text, s)
            if align == 'center':
                x = r.x0 + (r.width - tw) / 2
            elif align == 'right':
                x = r.x1 - pad - tw
            y = r.y0 + (r.height + s * 0.72) / 2
            page.insert_text((x, y), text, fontsize=s, fontname=font, color=INK)
            return
        s -= 0.2
    s = size
    while s >= minsize:
        box = pymupdf.Rect(r.x0 + pad, r.y0 + 1, r.x1 - pad, r.y1 - 0.5)
        if page.insert_textbox(box, text, fontsize=s, fontname=font, color=INK, align=0 if align == 'left' else 1) >= 0:
            return
        s -= 0.2
    raise ValueError(f'text does not fit: {text!r} in {rect}')


def line_text(page, x0, x1, y, text, size=10, font='helv'):
    """Text sitting on an underline at y."""
    put(page, (x0, y - 15, x1, y - 1), text, size=size, font=font)


# ------------------------------------------------------------------ Form 1
F1_BOX = {1: (39, 232, 284, 254), 2: (312, 232, 556, 264), 3: (39, 342, 147, 363), 4: (175, 342, 284, 363),
          5: (311, 342, 420, 363), 6: (447, 342, 556, 363), 7: (39, 451, 120, 473), 8: (148, 451, 229, 473),
          9: (257, 451, 338, 473), 10: (366, 451, 447, 473), 11: (475, 451, 556, 473)}
F1_URGENT_Y = (648.8, 668.2, 687.8)


def names_text(count, names):
    if count in (None, ''):
        return ''
    if names:
        return f'{count} ({", ".join(names)})'
    return str(count)


def fill_form1(day, data):
    rec = data['days'][day]
    doc = pymupdf.open(os.path.join(ROOT, 'forms', 'Form1_Daily_Report_blank.pdf'))
    p = doc[0]
    line_text(p, 30, 277, 153, data['team_leader'], 11, 'hebo')
    line_text(p, 290, 436, 153, data['team'], 11, 'hebo')
    line_text(p, 449, 566, 153, fmt_date(day), 11, 'hebo')
    put(p, F1_BOX[1], rec.get('chats_reviewed'), size=10)
    c2 = rec.get('chats_with_issues')
    note = rec.get('chats_issue_note', '')
    t2 = (f'{c2} - {note}' if c2 not in (None, '') else note) if note else (c2 if c2 not in (None, '') else '')
    put(p, F1_BOX[2] if len(str(t2)) > 60 else (312, 232, 556, 254), t2, size=9, minsize=4.8)
    put(p, F1_BOX[3], names_text(rec.get('meetings_held'), rec.get('meetings_with')), size=9)
    put(p, F1_BOX[4], names_text(rec.get('agents_absent'), rec.get('absent_names')), size=9)
    line_text(p, 311, 420, 363, rec.get('white_phone_holder', ''), 8.5)
    put(p, F1_BOX[6], rec.get('white_phone_calls'), size=10)
    for k, key in ((7, 'lost_deals'), (8, 'refunds'), (9, 'complaints'), (10, 'discounts'), (11, 'system_issues')):
        put(p, F1_BOX[k], rec.get(key + '_display', rec.get(key)), size=10 if key + '_display' not in rec else 7.5)
    for i, txt in enumerate(rec.get('urgent', [])[:3]):
        line_text(p, 30, 566, F1_URGENT_Y[i] + 0, txt, 9)
    out = os.path.join(ROOT, 'reports', 'daily', f'Form1_Daily_Report_{day}.pdf')
    os.makedirs(os.path.dirname(out), exist_ok=True)
    doc.save(out)
    return out


# ------------------------------------------------------------------ Form 2
def short_date(s):
    return fmt_date(s, 'short') if s else ''


def table_rows(page_no, doc, header_index=0):
    sys.path.insert(0, os.path.join(ROOT, 'scripts'))
    from analyze import tables
    return tables(os.path.join(ROOT, 'forms', 'Form2_Case_Log_blank.pdf'))


def fill_form2(week, data):
    wk = data['weeks'][str(week)]
    L = wk.get('logs', {})
    doc = pymupdf.open(os.path.join(ROOT, 'forms', 'Form2_Case_Log_blank.pdf'))
    wtxt = f"{fmt_date(wk['from'], 'short')} to {fmt_date(wk['to'])}"
    # page 1 header fields
    p = doc[0]
    line_text(p, 30, 288, 133.5, data['team_leader'], 10.5, 'hebo')
    line_text(p, 301.5, 462.8, 133.5, data['team'], 10.5, 'hebo')
    line_text(p, 476.2, 637.5, 133.5, fmt_date(wk['from']), 10.5, 'hebo')
    line_text(p, 651, 812.2, 133.5, fmt_date(wk['to']), 10.5, 'hebo')
    for pi in range(1, 7):
        line_text(doc[pi], 503.2, 638.2, 31.5, data['team_leader'], 9, 'hebo')
        line_text(doc[pi], 677.2, 812.2, 31.5, wtxt, 9, 'hebo')
    tabs = table_rows(0, doc)
    # order of tables in the form: (page index, table position on page) -> log name, column keys
    spec = [('chat_issues', 0, 0, ['date', 'agent', 'client', 'code', 'feedback']),
            ('one_to_one', 1, 0, ['date', 'agent', 'topic', 'key_points', 'action']),
            ('lost_deals', 2, 0, ['date', 'client', 'agent', 'quote', 'reason', 'comment']),
            ('refunds', 2, 1, ['date', 'client', 'agent', 'amount', 'reason', 'comment']),
            ('complaints', 3, 0, ['date', 'agent', 'what', 'cause', 'fixed', 'status']),
            ('discounts', 3, 1, ['date', 'client', 'agent', 'requested', 'approved', 'reason', 'outcome']),
            ('system_issues', 4, 0, ['date', 'system', 'what', 'impact', 'reported_to', 'status']),
            ('white_phone', 4, 1, ['date', 'holder', 'calls', 'asked', 'notes']),
            ('pending', 5, 0, ['client', 'agent', 'pending', 'expected']),
            ('agents_behind', 5, 1, ['agent', 'reason', 'action', 'result']),
            ('training', 6, 0, ['date', 'topic', 'roleplay', 'key_points'])]
    by_page = {}
    for t in tabs:
        by_page.setdefault(t[0], []).append(t)
    for name, pi, pos, keys in spec:
        t = by_page[pi][pos]
        cols, rows = t[3][1:], t[4]
        for ri, row in enumerate(L.get(name, [])):
            if ri >= len(rows):
                raise ValueError(f'{name}: more rows than the form holds ({len(rows)}). Continue on a new sheet.')
            y0, y1 = rows[ri]
            for (x0, x1), k in zip(cols, keys):
                v = row.get(k, '')
                if k in ('date', 'expected') and v:
                    v = short_date(v)
                if k in ('quote', 'amount') and v != '':
                    v = f'{float(v):,.2f}'
                put(doc[pi], (x0, y0, x1, y1), v, size=8.5, minsize=4.8)
    # weekly summary (page 7)
    S = wk.get('summary', {})
    boxes = {'chats_with_issues': (39.4, 325.1, 210.4, 342.4), 'meetings': (236.6, 325.1, 407.6, 342.4), 'lost_deals': (433.9, 325.1, 604.9, 342.4),
             'refunds_number': (631.1, 325.1, 713.6, 342.4), 'refunds_total': (720.4, 325.1, 802.9, 342.4),
             'complaints_number': (39.4, 388.9, 121.9, 406.1), 'complaints_open': (128.6, 388.9, 210.4, 406.1),
             'discounts_number': (236.6, 388.9, 319.1, 406.1), 'discounts_approved': (325.9, 388.9, 407.6, 406.1),
             'system_number': (433.9, 388.9, 516.4, 406.1), 'system_open': (523.1, 388.9, 604.9, 406.1), 'white_phone_calls': (631.1, 388.9, 802.9, 406.1)}
    for k, rect in boxes.items():
        v = S.get(k)
        if v not in (None, ''):
            put(doc[6], rect, v, size=10)
    sfx = '_DRAFT' if wk.get('status', 'draft') == 'draft' else ''
    out = os.path.join(ROOT, 'reports', 'weekly', f'Form2_Case_Log_Week{week}{sfx}.pdf')
    os.makedirs(os.path.dirname(out), exist_ok=True)
    doc.save(out)
    return out


# ------------------------------------------------------------------ Form 3
def week_totals(week, data):
    wk = data['weeks'][str(week)]
    start, end = wk['from'], wk['to']
    sales = data.get('sales', {})
    tot = {}
    seen = False
    for day, row in sales.items():
        if start <= day <= end:
            seen = True
            for ag, v in row.items():
                tot[ag] = tot.get(ag, 0) + v
    return tot if seen else None


def money(v):
    if v in (None, ''):
        return ''
    v = float(v)
    return f'{int(v):,}' if v.is_integer() else f'{v:,.2f}'


def fill_form3(week, data):
    wk = data['weeks'][str(week)]
    f3 = wk.get('form3', {})
    doc = pymupdf.open(os.path.join(ROOT, 'forms', 'Form3_Weekly_Report_blank.pdf'))
    p = doc[0]
    line_text(p, 30, 201.8, 153, f3.get('team_leader', data['team_leader']), 10.5, 'hebo')
    line_text(p, 215.2, 323.2, 153, f3.get('team', data['team']), 10, 'hebo')
    line_text(p, 336.8, 444, 153, fmt_date(wk['from']), 10.5, 'hebo')
    line_text(p, 457.5, 565.5, 153, fmt_date(wk['to']), 10.5, 'hebo')
    totals = week_totals(week, data)
    agents = data['agents']
    targets = f3.get('agent_targets', {})
    team_target = f3.get('team_target')
    team_ach = sum(totals.values()) if totals else f3.get('team_achieved')
    put(p, (39.4, 214.1, 193.1, 232.9), money(team_target), size=10)
    put(p, (220.9, 214.1, 374.6, 232.9), money(team_ach), size=10)
    if team_target and team_ach is not None:
        put(p, (402.4, 214.1, 556.1, 232.9), f'{team_ach / team_target * 100:.1f}', size=10)
    ys = [306.0 + 21 * i for i in range(9)]
    for i, ag in enumerate(agents[:9]):
        y0, y1 = ys[i], ys[i] + 21
        put(p, (53.2, y0, 298.5, y1), ag, size=9)
        t = targets.get(ag)
        a = totals.get(ag) if totals else f3.get('agent_achieved', {}).get(ag)
        put(p, (298.5, y0, 396.0, y1), money(t), size=9)
        put(p, (396.0, y0, 493.5, y1), money(a), size=9)
        if t and a is not None:
            put(p, (493.5, y0, 565.5, y1), f'{a / t * 100:.1f}', size=9)
    for key, rect in (('open_now', (39.4, 568.9, 124.9, 587.6)), ('expected_next_week', (152.6, 568.9, 238.1, 587.6)),
                      ('grey_contacted', (270.4, 559.9, 346.9, 578.6)), ('grey_replied', (374.6, 559.9, 451.1, 578.6)), ('grey_converted', (478.9, 559.9, 556.1, 578.6))):
        put(p, rect, f3.get(key), size=10)
    for i, ytxt in enumerate(f3.get('top_actions', [])[:3]):
        y = (648.8, 668.2, 687.8)[i]
        if text_width(ytxt, 9) <= 526 - 6:
            line_text(p, 40, 566, y, ytxt, 9)
        else:  # long action: wrap onto two lines between the previous underline and this one
            sz = 8.2
            while sz >= 6:
                words, lines, cur = ytxt.split(), [], ''
                for wd in words:
                    t = (cur + ' ' + wd).strip()
                    if text_width(t, sz) <= 520:
                        cur = t
                    else:
                        lines.append(cur); cur = wd
                lines.append(cur)
                if len(lines) <= 2:
                    break
                sz -= 0.2
            for k, ln in enumerate(lines):
                p.insert_text((42, y - 10.8 + k * 9.6 - (0 if len(lines) == 2 else 5)), ln, fontsize=sz, fontname='helv', color=INK)
    out = os.path.join(ROOT, 'reports', 'weekly', f'Form3_Weekly_Report_Week{week}{"_DRAFT" if wk.get("form3_status", wk.get("status", "draft")) == "draft" else ""}.pdf')
    os.makedirs(os.path.dirname(out), exist_ok=True)
    doc.save(out)
    return out



def combine(data):
    """One PDF with every report, with bookmarks, so everything can be checked in one place."""
    out = pymupdf.open()
    toc = []
    days = sorted(data['days'])
    if days:
        toc.append([1, 'Form 1 - Daily Reports', out.page_count + 1])
        for d in days:
            path = fill_form1(d, data)
            toc.append([2, fmt_date(d), out.page_count + 1])
            out.insert_pdf(pymupdf.open(path))
    for w in sorted(data['weeks'], key=int):
        wk = data['weeks'][w]
        tag = ' (DRAFT)' if wk.get('status', 'draft') == 'draft' else ''
        span = f"{fmt_date(wk['from'], 'short')} to {fmt_date(wk['to'])}"
        toc.append([1, f'Form 2 - Case Log, Week {w} ({span}){tag}', out.page_count + 1])
        out.insert_pdf(pymupdf.open(fill_form2(int(w), data)))
        toc.append([1, f'Form 3 - Weekly Report, Week {w} ({span}){tag}', out.page_count + 1])
        out.insert_pdf(pymupdf.open(fill_form3(int(w), data)))
    for rv in data.get('monthly_reviews', []):
        pass
    out.set_toc(toc)
    path = os.path.join(ROOT, 'reports', 'Travnook_All_Reports_Oct2026.pdf')
    out.save(path)
    return path


if __name__ == '__main__':
    data = load()
    cmd = sys.argv[1]
    if cmd == 'form1':
        print(fill_form1(sys.argv[2], data))
    elif cmd == 'form2':
        print(fill_form2(int(sys.argv[2]), data))
    elif cmd == 'form3':
        print(fill_form3(int(sys.argv[2]), data))
    elif cmd == 'week':
        print(fill_form2(int(sys.argv[2]), data)); print(fill_form3(int(sys.argv[2]), data))
    elif cmd == 'combine':
        print(combine(data))
    elif cmd == 'day-all':
        for d in sorted(data['days']):
            print(fill_form1(d, data))
