# -*- coding: utf-8 -*-
"""주간 스케쥴 구글시트 링크 → 기간(첫날~마지막날) 추출, index.html DAY_SHEETS 에 붙여넣을 줄 출력

사용법:  python tools/day_range.py "<구글시트 웹에 게시 링크>" [링크2 ...]
  예)    python tools/day_range.py "https://docs.google.com/spreadsheets/d/e/2PACX-.../pubhtml?gid=780158230&single=true"

시트 표의 '10월 11일 (일)' 같은 날짜 칸을 읽어 기간을 정한다. 연도는 표에 없으므로 오늘과 가장 가까운 쪽으로 정하고
(12월→1월 넘어가면 다음 해), 날짜가 하루도 빠짐없이 이어지는지 확인한다.
"""
import sys, re, html, datetime, urllib.request, urllib.parse

sys.stdout.reconfigure(encoding='utf-8')
WD = '월화수목금토일'


def fetch(url):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.read().decode('utf-8', 'ignore')


def sheet_range(link):
    m = re.match(r'(https://docs\.google\.com/spreadsheets/d/e/[^/]+)/pubhtml', link)
    if not m:
        raise SystemExit(f'구글시트 "웹에 게시" 링크가 아님: {link}')
    gid = (urllib.parse.parse_qs(urllib.parse.urlparse(link).query).get('gid') or ['0'])[0]
    page = fetch(f'{m.group(1)}/pubhtml/sheet?headers=false&gid={gid}')
    cells = [html.unescape(re.sub(r'<[^>]+>', '', c)).strip() for c in re.findall(r'<td[^>]*>(.*?)</td>', page, re.S)]
    md = []
    for c in cells:
        for mo, d in re.findall(r'(\d{1,2})월\s*(\d{1,2})일', c):
            if (int(mo), int(d)) not in md:
                md.append((int(mo), int(d)))
    if not md:
        raise SystemExit('시트에서 "N월 N일" 날짜를 찾지 못함')
    today = datetime.date.today()
    best = None
    for y0 in (today.year - 1, today.year, today.year + 1):
        y, prev, ds = y0, None, []
        for mo, d in md:
            if prev and mo < prev:            # 12월 → 1월
                y += 1
            prev = mo
            ds.append(datetime.date(y, mo, d))
        mid = ds[0] + (ds[-1] - ds[0]) / 2
        if best is None or abs((mid - today).days) < abs((best[0] - today).days):
            best = (mid, ds)
    ds = sorted(best[1])
    gaps = [ds[i] for i in range(1, len(ds)) if (ds[i] - ds[i - 1]).days != 1]
    return ds[0], ds[-1], len(ds), gaps


def main():
    if len(sys.argv) < 2:
        raise SystemExit(__doc__)
    lines = []
    for link in sys.argv[1:]:
        a, b, n, gaps = sheet_range(link)
        ok = '✓ 연속' if not gaps else f'⚠ 중간에 빠진 날 있음: {[g.isoformat() for g in gaps]}'
        print(f'[기간] {a}({WD[a.weekday()]}) ~ {b}({WD[b.weekday()]}) · {n}일 {ok}')
        lines.append(f"        {{ from: '{a}', to: '{b}', url: '{link}' }},")
    print('\nindex.html 의 DAY_SHEETS 에 붙여넣기:')
    print('\n'.join(lines))


if __name__ == '__main__':
    main()
