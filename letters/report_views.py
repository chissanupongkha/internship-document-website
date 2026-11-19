# -*- coding: utf-8 -*-
"""Reports (under Tracking): one dashboard, one data set, exported as Excel or Word.

Groups: placement, pipeline, volume, companies, changes, internship period,
letters, turnaround. Filters: academic year, semester, subject.
"""
import io
import re
import statistics
from zoneinfo import ZoneInfo
from collections import Counter, defaultdict

import pandas as pd
from openpyxl.styles import Font, PatternFill
from django.http import HttpResponse
from django.shortcuts import render
from django.utils import timezone

from .decorators import get_role, staff_required
import datetime

from .models import AuditLog, ChangeLog, Enrollment, GeneratedDocument, StudentRecord
from .status_views import (ISSUE_LABELS, _check_records, STAGE_DRAFT, STAGE_INFO, STAGE_RECEIVED, STAGE_SENT, STAGE_SIGNED,
                           _list_stage, _norm_code, _period_dates, _company_key, _same_company)

DEFAULT_STUCK_DAYS = 7
DEFAULT_TARGET_DAYS = 7
PERIOD_BUCKETS = [(0, 29, 'under 30 days'), (30, 44, '30-44 days'), (45, 59, '45-59 days'),
                  (60, 89, '60-89 days'), (90, 10 ** 6, '90 days or more')]


def _code(text):
    """'ITCS499 การฝึกงาน' -> 'ITCS499' (falls back to the cleaned text)."""
    m = re.search(r'[A-Za-z]{2,5}\s?-?\d{3}', str(text or ''))
    return _norm_code(m.group(0)) if m else _norm_code(text)


def _int(value, default):
    try:
        return max(1, int(value))
    except (TypeError, ValueError):
        return default


def _pct(part, whole):
    return round(part * 100.0 / whole, 1) if whole else 0.0


def _days(a, b):
    """Days between two datetimes (b - a), None when either is missing or negative."""
    if not a or not b:
        return None
    d = (b - a).total_seconds() / 86400.0
    return round(d, 1) if d >= 0 else None


def _summary(values):
    vals = [v for v in values if v is not None]
    if not vals:
        return {'n': 0, 'avg': None, 'median': None, 'max': None}
    return {'n': len(vals), 'avg': round(sum(vals) / len(vals), 1),
            'median': round(statistics.median(vals), 1), 'max': round(max(vals), 1)}


def _filters(params):
    return {
        'year': (params.get('year') or '').strip(),
        'semester': (params.get('semester') or '').strip(),
        'subject': (params.get('subject') or '').strip().upper(),
        'stuck': _int(params.get('stuck'), DEFAULT_STUCK_DAYS),
        'target': _int(params.get('target'), DEFAULT_TARGET_DAYS),
    }


def _options():
    sems = set(StudentRecord.objects.exclude(semester='').values_list('semester', flat=True))
    sems |= set(Enrollment.objects.values_list('semester', flat=True))
    semesters = sorted(sems, reverse=True)
    years = sorted({s.split('/')[0] for s in semesters}, reverse=True)
    subjects = sorted({_code(c) for c in StudentRecord.objects.exclude(course_id='').values_list('course_id', flat=True)})
    return semesters, years, subjects


def _scope_records(f):
    qs = StudentRecord.objects.select_related('batch')
    if f['semester']:
        qs = qs.filter(semester=f['semester'])
    elif f['year']:
        qs = qs.filter(semester__startswith=f['year'] + '/')
    records = list(qs.order_by('id'))
    if f['subject']:
        records = [r for r in records if f['subject'] == _code(r.course_id)]
    return records


def _placement(f):
    enr = Enrollment.objects.select_related('student', 'course')
    if f['semester']:
        enr = enr.filter(semester=f['semester'])
    elif f['year']:
        enr = enr.filter(semester__startswith=f['year'] + '/')
    if f['subject']:
        enr = enr.filter(course__code=f['subject'])
    placed = set()
    recs = StudentRecord.objects.exclude(company='')
    if f['semester']:
        recs = recs.filter(semester=f['semester'])
    elif f['year']:
        recs = recs.filter(semester__startswith=f['year'] + '/')
    for sid, sem in recs.values_list('student_id', 'semester'):
        placed.add(((sid or '').strip(), sem))
    people, seen = [], set()
    for e in enr.order_by('semester', 'student__student_id'):
        key = (e.student.student_id.strip(), e.semester)
        if key in seen:
            continue
        seen.add(key)
        people.append({'student_id': e.student.student_id, 'name': e.student.display_name,
                       'subject': e.course.label, 'semester': e.semester, 'has_place': key in placed})
    return people, [p for p in people if not p['has_place']]


def _cluster_companies(names):
    """name -> cluster label. Different spellings of one company share a label."""
    labels, clusters = {}, []                      # clusters: list of (label, key)
    for name in sorted({n.strip() for n in names if n and n.strip()}, key=lambda n: (-len(n), n)):
        for label, key in clusters:
            if _same_company(label, name):
                labels[name] = label
                break
        else:
            clusters.append((name, _company_key(name)))
            labels[name] = name
    return labels


def build(params):
    f = _filters(params)
    records = _scope_records(f)
    now = timezone.now()
    D = {'filters': f}

    # ---- latest letter per request
    docs = {}
    for d in GeneratedDocument.objects.filter(record__in=[r.pk for r in records]).order_by('id'):
        docs[d.record_id] = d
    rows = []
    for r in records:
        d = docs.get(r.pk)
        stage = _list_stage(d)
        received = r.batch.uploaded_at if r.batch else None
        draft = d.created_at if (d and stage >= STAGE_DRAFT) else None
        signed = d.signed_at if (d and stage >= STAGE_SIGNED) else None
        sent = d.email_sent_at if (d and stage == STAGE_SENT) else None
        last = max([x for x in (received, draft, signed, sent) if x], default=None)
        rows.append({'r': r, 'd': d, 'stage': stage, 'received': received, 'draft': draft,
                     'signed': signed, 'sent': sent, 'age': _days(last, now)})
    D['total'] = len(rows)

    # ---- placement
    people, missing = _placement(f)
    D['placement'] = {'people': len(people), 'without': len(missing), 'percent': _pct(len(missing), len(people)),
                      'missing': missing}

    # ---- pipeline
    stage_counts = Counter(x['stage'] for x in rows)
    D['pipeline'] = [{'label': STAGE_INFO[s][0], 'count': stage_counts.get(s, 0),
                      'percent': _pct(stage_counts.get(s, 0), len(rows))}
                     for s in (STAGE_RECEIVED, STAGE_DRAFT, STAGE_SIGNED, STAGE_SENT)]
    stuck = [x for x in rows if x['stage'] != STAGE_SENT and x['age'] is not None and x['age'] >= f['stuck']]
    stuck.sort(key=lambda x: -x['age'])
    D['stuck'] = [{'student_id': x['r'].student_id, 'name': x['r'].display_name, 'subject': x['r'].course_id,
                   'status': STAGE_INFO[x['stage']][0], 'days': x['age'], 'record_id': x['r'].pk} for x in stuck]

    # ---- volume
    D['by_semester'] = sorted(Counter(r.semester or '(none)' for r in records).items())
    D['by_type'] = Counter((r.letter_type or '(none)').strip() for r in records).most_common()
    D['by_subject'] = Counter(_code(r.course_id) or '(none)' for r in records).most_common()
    repeat = [r for r in records if (r.times or 1) > 1]
    D['repeat'] = {'count': len(repeat), 'percent': _pct(len(repeat), len(records)),
                   'rows': [{'student_id': r.student_id, 'name': r.display_name, 'subject': r.course_id,
                             'semester': r.semester, 'times': r.times} for r in repeat]}

    # ---- companies
    labels = _cluster_companies([r.company for r in records])
    students_by_company = defaultdict(set)
    for r in records:
        if r.company.strip():
            students_by_company[labels[r.company.strip()]].add(r.student_id)
    D['companies'] = {'distinct': len(students_by_company),
                      'top': sorted(((c, len(s)) for c, s in students_by_company.items()), key=lambda x: (-x[1], x[0]))[:10]}
    # new vs seen before (only meaningful for one semester)
    if f['semester']:
        earlier = StudentRecord.objects.exclude(company='').exclude(semester__gte=f['semester']).values_list('company', flat=True)
        earlier_labels = _cluster_companies(list(earlier) + [r.company for r in records])
        before = {earlier_labels[c.strip()] for c in earlier if c.strip()}
        now_set = set(students_by_company)
        mapped = {earlier_labels.get(c, c) for c in now_set}
        D['companies']['new'] = len(mapped - before)
        D['companies']['returning'] = len(mapped & before)
    else:
        D['companies']['new'] = D['companies']['returning'] = None

    # ---- changes
    ch = list(ChangeLog.objects.filter(record__in=[r.pk for r in records]).select_related('record', 'changed_by')
              .order_by('-changed_at'))
    D['changes'] = {'count': len(ch), 'records': len({c.record_id for c in ch}),
                    'percent': _pct(len({c.record_id for c in ch}), len(records)),
                    'company': sum(1 for c in ch if c.field == 'company'),
                    'period': sum(1 for c in ch if c.field == 'period'),
                    'rows': [{'when': c.changed_at, 'student_id': c.record.student_id, 'field': c.field,
                              'old': c.old_value, 'new': c.new_value,
                              'by': c.changed_by.username if c.changed_by else '', 'reason': c.reason} for c in ch]}

    # ---- internship period
    lengths, months = [], Counter()
    for r in records:
        pd_ = _period_dates(r.period)
        if pd_:
            lengths.append((pd_[1] - pd_[0]).days + 1)
            months[f"{pd_[0]:%Y-%m}"] += 1
    buckets = [{'label': lab, 'count': sum(1 for n in lengths if lo <= n <= hi)} for lo, hi, lab in PERIOD_BUCKETS]
    D['period'] = {'readable': len(lengths), 'average': round(sum(lengths) / len(lengths), 1) if lengths else None,
                   'buckets': buckets, 'start_months': sorted(months.items())}

    # ---- letters
    edited = [x for x in rows if x['d'] and (x['d'].version or 1) > 1]
    D['letters'] = {'edited': len(edited), 'percent': _pct(len(edited), len(rows))}

    # ---- turnaround
    done = [x for x in rows if x['stage'] == STAGE_SENT and x['sent']]
    total = [_days(x['received'], x['sent']) for x in done]
    within = [t for t in total if t is not None and t <= f['target']]
    D['turnaround'] = {
        'finished': len(done), 'target': f['target'],
        'within_percent': _pct(len(within), len([t for t in total if t is not None])),
        'total': _summary(total),
        'steps': [
            {'label': 'Received to draft', **_summary([_days(x['received'], x['draft']) for x in rows])},
            {'label': 'Draft to signed', **_summary([_days(x['draft'], x['signed']) for x in rows])},
            {'label': 'Signed to emailed', **_summary([_days(x['signed'], x['sent']) for x in rows])},
        ],
        'slowest': sorted([{'student_id': x['r'].student_id, 'name': x['r'].display_name,
                            'days': _days(x['received'], x['sent'])} for x in done
                           if _days(x['received'], x['sent']) is not None], key=lambda z: -z['days'])[:10],
    }

    # ---- timeline: requests received per month, letters made per week
    months_in = Counter(f"{x['received']:%Y-%m}" for x in rows if x['received'])
    D['monthly'] = sorted(months_in.items())[-12:]
    week0 = (now - datetime.timedelta(days=now.weekday())).replace(hour=0, minute=0, second=0, microsecond=0)
    weeks = [week0 - datetime.timedelta(weeks=i) for i in range(7, -1, -1)]
    made = Counter()
    for x in rows:
        if x['draft']:
            for w in reversed(weeks):
                if x['draft'] >= w:
                    made[w] += 1
                    break
    D['weekly'] = [(f"{w:%d %b}", made.get(w, 0)) for w in weeks]

    # ---- who asks: study year and popular positions
    ycount = Counter((re.sub(r'\D', '', r.year or '') or '?') for r in records)
    D['study_year'] = sorted(ycount.items())
    pos = Counter((r.internship_position or '').strip().lower() for r in records if (r.internship_position or '').strip())
    D['positions'] = pos.most_common(8)

    # ---- data quality (same checks as the Data check page)
    checked = _check_records(records) if records else []
    qcount = Counter()
    flagged = 0
    for res in checked:
        if res['issues']:
            flagged += 1
        for i in res['issues']:
            qcount[i['code']] += 1
    D['quality'] = {'flagged': flagged, 'percent': _pct(flagged, len(records)),
                    'rows': [(ISSUE_LABELS.get(k, k), n) for k, n in qcount.most_common()]}

    # ---- staff activity, last 30 days
    since = now - datetime.timedelta(days=30)
    logs = list(AuditLog.objects.filter(timestamp__gte=since).select_related('user'))
    D['staff'] = {'total': len(logs),
                  'by_user': Counter((l.user.username if l.user else 'deleted account') for l in logs).most_common(8),
                  'by_action': Counter(l.action for l in logs).most_common(8)}

    # ---- year on year (ignores the filters)
    D['yearly'] = _yearly()
    return D


def _yearly():
    all_recs = list(StudentRecord.objects.select_related('batch'))
    by_year = defaultdict(list)
    for r in all_recs:
        y = (r.semester or '').split('/')[0]
        if y:
            by_year[y].append(r)
    docs = {}
    for d in GeneratedDocument.objects.filter(record__in=[r.pk for r in all_recs]).order_by('id'):
        docs[d.record_id] = d
    out = []
    for y in sorted(by_year, reverse=True)[:6]:
        recs = by_year[y]
        names = _cluster_companies([r.company for r in recs])
        done = [(docs[r.pk], r) for r in recs if r.pk in docs and docs[r.pk].email_sent_at]
        spans = [_days(r.batch.uploaded_at if r.batch else None, d.email_sent_at) for d, r in done]
        spans = [v for v in spans if v is not None]
        people, missing = _placement({'year': y, 'semester': '', 'subject': ''})
        out.append({'year': y, 'requests': len(recs),
                    'companies': len({names[r.company.strip()] for r in recs if r.company.strip()}),
                    'finished': _pct(len(done), len(recs)),
                    'avg_days': round(sum(spans) / len(spans), 1) if spans else None,
                    'no_place': _pct(len(missing), len(people)) if people else None})
    return out


def _conic(parts):
    """CSS conic-gradient for a donut from [(value, color), ...]."""
    total = sum(v for v, _ in parts)
    if not total:
        return 'conic-gradient(#E5E7EB 0 100%)'
    pos, stops = 0.0, []
    for v, col in parts:
        end = pos + v * 100.0 / total
        stops.append(f"{col} {pos:.2f}% {end:.2f}%")
        pos = end
    return 'conic-gradient(' + ', '.join(stops) + ')'


def _cols(pairs):
    """[(label, n)] -> [(label, n, height%)] for column charts."""
    top = max([n for _, n in pairs] + [1])
    return [(k, n, max(int(n * 100 / top), 3 if n else 0)) for k, n in pairs]


def _bars(pairs):
    top = max([n for _, n in pairs] + [1])
    return [(k, n, int(n * 100 / top)) for k, n in pairs]


def _context(request):
    D = build(request.GET)
    semesters, years, subjects = _options()
    D.update({'semesters': semesters, 'years': years, 'subjects': subjects,
              'show_names': get_role(request.user) != 'viewer'})
    maxes = {
        'pipeline': max([p['count'] for p in D['pipeline']] + [1]),
        'semester': max([n for _, n in D['by_semester']] + [1]),
        'type': max([n for _, n in D['by_type']] + [1]),
        'subject': max([n for _, n in D['by_subject']] + [1]),
        'company': max([n for _, n in D['companies']['top']] + [1]),
        'bucket': max([b['count'] for b in D['period']['buckets']] + [1]),
    }
    D['maxes'] = maxes
    colors = ['#9CA3AF', '#B4772A', '#3F7A62', '#006C6C']
    D['pipeline_gradient'] = _conic([(p['count'], colors[i]) for i, p in enumerate(D['pipeline'])])
    for i, p in enumerate(D['pipeline']):
        p['color'] = colors[i]
    pl = D['placement']
    pl['have'] = pl['people'] - pl['without']
    D['nav'] = [('overview', 'Overview'), ('time', 'Over time'), ('breakdown', 'Breakdowns'), ('companies', 'Companies'), ('speed', 'Turnaround'), ('quality', 'Quality and staff')] + ([('lists', 'Lists')] if D['show_names'] else [])
    D['placement_gradient'] = _conic([(pl['people'] - pl['without'], '#3F7A62'), (pl['without'], '#D9A441')])
    D['monthly'] = _cols(D['monthly'])
    D['weekly'] = _cols(D['weekly'])
    D['study_year'] = _cols(D['study_year'])
    D['positions'] = _bars(D['positions'])
    D['staff']['by_user'] = _bars(D['staff']['by_user'])
    D['staff']['by_action'] = _bars(D['staff']['by_action'])
    D['quality']['rows'] = _bars(D['quality']['rows'])
    for key, name in (('pipeline', 'pipeline'),):
        for p in D['pipeline']:
            p['bar'] = int(p['count'] * 100 / maxes['pipeline'])
    D['by_semester'] = [(k, n, int(n * 100 / maxes['semester'])) for k, n in D['by_semester']]
    D['by_type'] = [(k, n, int(n * 100 / maxes['type'])) for k, n in D['by_type']]
    D['by_subject'] = [(k, n, int(n * 100 / maxes['subject'])) for k, n in D['by_subject']]
    D['companies']['top'] = [(k, n, int(n * 100 / maxes['company'])) for k, n in D['companies']['top']]
    for b in D['period']['buckets']:
        b['bar'] = int(b['count'] * 100 / maxes['bucket'])
    return D


@staff_required
def report(request):
    return render(request, 'letters/report.html', _context(request))


def _frames(D, show_names=True):
    """Every table of the report as {sheet name: DataFrame} (used by Excel and Word)."""
    f = D['filters']
    scope = f['semester'] or f['year'] or 'all semesters'
    if f['subject']:
        scope += f" / {f['subject']}"
    pl, ch, tu, co = D['placement'], D['changes'], D['turnaround'], D['companies']
    kpi = [
        ['Scope', scope],
        ['Requests', D['total']],
        ['Students on the subject list', pl['people']],
        ['Students without an internship place', pl['without']],
        ['Percentage without a place (%)', pl['percent']],
        ['Distinct companies', co['distinct']],
        ['New companies this semester', co['new'] if co['new'] is not None else 'n/a'],
        ['Returning companies', co['returning'] if co['returning'] is not None else 'n/a'],
        ['Company / period changes', ch['count']],
        ['Requests changed (%)', ch['percent']],
        ['Repeat submissions', D['repeat']['count']],
        ['Letters edited after generating', D['letters']['edited']],
        ['Average internship length (days)', D['period']['average'] if D['period']['average'] is not None else 'n/a'],
        [f"Finished within {tu['target']} days (%)", tu['within_percent']],
        ['Average days received to emailed', tu['total']['avg'] if tu['total']['avg'] is not None else 'n/a'],
    ]
    F = {
        'Summary': pd.DataFrame(kpi, columns=['Metric', 'Value']),
        'Pipeline': pd.DataFrame([[p['label'], p['count'], p['percent']] for p in D['pipeline']],
                                 columns=['Status', 'Requests', '%']),
        'By semester': pd.DataFrame([[k, n] for k, n, *_ in D['by_semester']], columns=['Semester', 'Requests']),
        'By letter type': pd.DataFrame([[k, n] for k, n, *_ in D['by_type']], columns=['Letter type', 'Requests']),
        'By subject': pd.DataFrame([[k, n] for k, n, *_ in D['by_subject']], columns=['Subject', 'Requests']),
        'Top companies': pd.DataFrame([[k, n] for k, n, *_ in co['top']], columns=['Company', 'Students']),
        'Internship length': pd.DataFrame([[b['label'], b['count']] for b in D['period']['buckets']],
                                          columns=['Length', 'Requests']),
        'Turnaround': pd.DataFrame(
            [['Received to emailed (total)', tu['total']['n'], tu['total']['avg'], tu['total']['median'], tu['total']['max']]]
            + [[s['label'], s['n'], s['avg'], s['median'], s['max']] for s in tu['steps']],
            columns=['Step', 'Requests', 'Average days', 'Median days', 'Slowest days']),
    }
    F['Requests per month'] = pd.DataFrame([[k, n] for k, n, *_ in D['monthly']], columns=['Month', 'Requests'])
    F['Letters per week'] = pd.DataFrame([[k, n] for k, n, *_ in D['weekly']], columns=['Week starting', 'Letters made'])
    F['Study year'] = pd.DataFrame([[k, n] for k, n, *_ in D['study_year']], columns=['Study year', 'Requests'])
    F['Popular positions'] = pd.DataFrame([[k, n] for k, n, *_ in D['positions']], columns=['Position', 'Requests'])
    F['Data quality'] = pd.DataFrame([[k, n] for k, n, *_ in D['quality']['rows']], columns=['Issue', 'Requests'])
    F['Year on year'] = pd.DataFrame(
        [[y['year'], y['requests'], y['companies'], y['finished'], y['avg_days'], y['no_place']] for y in D['yearly']],
        columns=['Academic year', 'Requests', 'Companies', 'Emailed (%)', 'Average days', 'Without a place (%)'])
    if show_names:
        F['Staff activity'] = pd.DataFrame([[k, n] for k, n, *_ in D['staff']['by_user']], columns=['Account', 'Actions (30 days)'])
        F['Without a place'] = pd.DataFrame([[m['student_id'], m['name'], m['subject'], m['semester']] for m in pl['missing']],
                                            columns=['Student ID', 'Name', 'Subject', 'Semester'])
        F['Stuck requests'] = pd.DataFrame([[s['student_id'], s['name'], s['subject'], s['status'], s['days']] for s in D['stuck']],
                                           columns=['Student ID', 'Name', 'Subject', 'Status', f"Days (>= {f['stuck']})"])
        F['Repeat submissions'] = pd.DataFrame([[s['student_id'], s['name'], s['subject'], s['semester'], s['times']] for s in D['repeat']['rows']],
                                               columns=['Student ID', 'Name', 'Subject', 'Semester', 'Time #'])
        F['Changes'] = pd.DataFrame([[c['when'].strftime('%Y-%m-%d %H:%M'), c['student_id'], c['field'], c['old'], c['new'], c['by'], c['reason']]
                                     for c in ch['rows']], columns=['When', 'Student ID', 'Field', 'Old', 'New', 'By', 'Reason'])
    return F


@staff_required
def report_export(request):
    D = build(request.GET)
    show_names = get_role(request.user) != 'viewer'
    F = _frames(D, show_names)
    f = D['filters']
    label = re.sub(r'[^A-Za-z0-9_-]+', '-', (f['semester'] or f['year'] or 'all') + (f"-{f['subject']}" if f['subject'] else ''))
    if request.GET.get('format') == 'docx':
        return _word(D, F, label)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as xl:
        for name, df in F.items():
            df.to_excel(xl, sheet_name=name[:31], index=False)
            ws = xl.sheets[name[:31]]
            ws.freeze_panes = 'A2'
            for cell in ws[1]:
                cell.font = Font(bold=True, color='FFFFFF')
                cell.fill = PatternFill('solid', fgColor='006C6C')
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(48, max(10, max(len(str(c.value or '')) for c in col) + 2))
    resp = HttpResponse(buf.getvalue(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    resp['Content-Disposition'] = f'attachment; filename="internship_report_{label}.xlsx"'
    return resp


def _shade(cell, hex_fill):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_fill)
    tc_pr.append(shd)


def _font(run, name, size=None, bold=None, color=None):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt, RGBColor
    run.font.name = name
    r_pr = run._element.get_or_add_rPr()
    fonts = r_pr.find(qn('w:rFonts'))
    if fonts is None:
        fonts = OxmlElement('w:rFonts')
        r_pr.insert(0, fonts)
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        fonts.set(qn(attr), name)
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def _word(D, F, label):
    """Word version: title, short summary sentence, then every table with a shaded header row."""
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.shared import Cm, Pt
    from .docgen import THAI_FONT

    doc = Document()
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Cm(2)
    sec.top_margin = sec.bottom_margin = Cm(2)
    doc.styles['Normal'].font.name = THAI_FONT
    doc.styles['Normal'].font.size = Pt(14)

    def para(text, size=14, bold=False, color=None, space_after=6):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(space_after)
        _font(p.add_run(text), THAI_FONT, size, bold, color)
        return p

    pl, ch, tu, co = D['placement'], D['changes'], D['turnaround'], D['companies']
    f = D['filters']
    scope = f['semester'] or f['year'] or 'all semesters'
    if f['subject']:
        scope += f" / {f['subject']}"
    para('Internship report', 24, True, '006C6C', 2)
    para(f"Scope: {scope}  |  Created {timezone.now().astimezone(ZoneInfo('Asia/Bangkok')):%d %b %Y %H:%M}", 12, False, '666666', 10)
    para(f"{D['total']} requests from {co['distinct']} companies. {pl['without']} of {pl['people']} students "
         f"({pl['percent']}%) do not have an internship place yet. {ch['count']} company or period changes were "
         f"recorded. {tu['within_percent']}% of finished requests were emailed within {tu['target']} days.", 14, False, None, 10)

    for name, df in F.items():
        h = para(name, 16, True, '006C6C', 4)
        h.paragraph_format.space_before = Pt(10)
        h.paragraph_format.keep_with_next = True
        if df.empty:
            para('No rows.', 12, False, '666666')
            continue
        t = doc.add_table(rows=1, cols=len(df.columns))
        t.style = 'Table Grid'
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, c in enumerate(df.columns):
            cell = t.rows[0].cells[i]
            cell.text = ''
            _font(cell.paragraphs[0].add_run(str(c)), THAI_FONT, 12, True, 'FFFFFF')
            _shade(cell, '006C6C')
        for n, (_, r) in enumerate(df.iterrows()):
            cells = t.add_row().cells
            for i, v in enumerate(r):
                cells[i].text = ''
                _font(cells[i].paragraphs[0].add_run('' if pd.isna(v) else str(v)), THAI_FONT, 12)
                if n % 2:
                    _shade(cells[i], 'F3F0E8')
    buf = io.BytesIO()
    doc.save(buf)
    resp = HttpResponse(buf.getvalue(), content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    resp['Content-Disposition'] = f'attachment; filename="internship_report_{label}.docx"'
    return resp
