# -*- coding: utf-8 -*-
"""주간 전체 스케쥴(부산2·부산3): 쿠팡 어드민 양식 엑셀 → day-data.js 변환 + 검증

사용법:
  python tools/convert_day.py <어드민.xlsx> [<어드민2.xlsx> ...]
  예) python tools/convert_day.py source/day/*.xlsx

  - 입력: 유스프에서 뽑은 어드민 업로드용 엑셀, 또는 쿠팡 어드민 '엑셀 다운로드'(schedule-v2-…xlsx).
    머리행: 업무일 | 벤더명 | 사업자등록번호 | 캠프명 | 웨이브 | 이름 | 아이디 | 업무상태 | 회전 | 업무라우트
  - 웨이브 WAVE2(주간), 캠프 부산2·부산3 만 씀 (CAMPS). 아이디·사업자등록번호는 읽지도 출력하지도 않음
    (repo 가 공개라 절대 올리면 안 됨 — 원본 엑셀은 git 제외 폴더 source/day/ 에만 보관).
  - 같은 날짜가 여러 파일에 있으면 **나중 파일이 그 날짜를 통째로 대체** (수정본을 다시 뽑아 준 경우).
    '나중' = 파일명 끝 숫자(schedule-v2-…_<숫자>.xlsx 의 내보낸 시각), 없으면 파일 수정 시각.
  - 한 사람이 두 캠프에 등록돼 하루 2줄(예: 부산2 휴무 + 부산3 출근)이면 '출근' 줄이 실제 근무.
    같은 날 두 캠프 모두 출근이면 중단.
  - 사람은 그 주에 출근을 더 많이 한 캠프 아래에 한 번만 나옴. 다른 캠프 라우트를 뛴 날은 칸에 그 캠프를 작게 표시.
  - 주(일~토)마다 따로 묶음 → 다음 주 파일을 추가해도 이미 올린 주의 사람 목록·순서는 그대로.
  - 회전이 '1회전,2회전'(전체)이 아니면 칸에 '2회전' 처럼 작게 표시 (한 라우트를 둘이 나눠 뛴 날).

■ 매주 절차
  1. 유스프에서 어드민 엑셀을 뽑을 때 그 파일을 source/day/ 에 넣는다 (이름 그대로 두면 됨).
  2. python tools/convert_day.py source/day/*.xlsx   → day-data.js 갱신
  3. git add day-data.js && git commit -m "주간 전체 스케쥴 갱신" && git push

필요: pip install openpyxl
"""
import sys, os, re, io, json, glob, datetime, warnings
import openpyxl

warnings.filterwarnings('ignore', message='Workbook contains no default style')
sys.stdout.reconfigure(encoding='utf-8')
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.environ.get('DAY_OUT') or os.path.join(ROOT, 'day-data.js')
CAMPS = ['부산2', '부산3']          # 게시판에 보여줄 캠프 (이 순서로)
WAVE = 'WAVE2'                     # 주간
FULL_ROT = {'', '1회전,2회전'}
ROUTE_RE = re.compile(r'\d{3}[A-Z]')
NEED = ['업무일', '캠프명', '웨이브', '이름', '업무상태', '회전', '업무라우트']


def fail(msg, items=()):
    print(f'[중단] {msg} — day-data.js를 쓰지 않았습니다.')
    for x in list(items)[:30]:
        print('   ✗', x)
    if len(items) > 30:
        print(f'   … 외 {len(items) - 30}건')
    sys.exit(1)


def stamp(path):
    m = re.search(r'_(\d{10,})\.xlsx$', os.path.basename(path))
    return int(m.group(1)) if m else int(os.path.getmtime(path) * 1000)


def as_date(v, where):
    if isinstance(v, datetime.datetime):
        return v.date().isoformat()
    if isinstance(v, datetime.date):
        return v.isoformat()
    if isinstance(v, (int, float)):                      # 엑셀 날짜 숫자
        return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(v))).isoformat()
    s = str(v or '').strip()
    m = re.fullmatch(r'(\d{4})[-./](\d{1,2})[-./](\d{1,2})', s)
    if m:
        return datetime.date(*map(int, m.groups())).isoformat()
    fail(f"{where}: 업무일 '{s}' 를 날짜로 못 읽음")


def load(path):
    """→ [(날짜, 캠프, 이름, 상태, 회전, [라우트])]  (주간·대상 캠프만, 아이디 등은 버림)"""
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    it = ws.iter_rows(values_only=True)
    hdr = [str(c).strip() if c is not None else '' for c in next(it)]
    ix = {h: i for i, h in enumerate(hdr)}
    miss = [h for h in NEED if h not in ix]
    if miss:
        fail(f"'{os.path.basename(path)}' 어드민 양식 머리행이 아님 (없는 칸: {miss})")
    rows, errs, other = [], [], set()
    for n, r in enumerate(it, start=2):
        if not any(v not in (None, '') for v in r):
            continue
        g = lambda k: r[ix[k]] if ix[k] < len(r) else None
        where = f'{os.path.basename(path)} {n}행'
        camp, wave = str(g('캠프명') or '').strip(), str(g('웨이브') or '').strip().upper()
        if camp not in CAMPS or not wave.startswith(WAVE):
            other.add(f'{camp}/{wave}')
            continue
        d = as_date(g('업무일'), where)
        name = str(g('이름') or '').strip()
        st = str(g('업무상태') or '').strip()
        rot = re.sub(r'\s+', '', str(g('회전') or ''))
        routes = [x for x in re.split(r'[,\s]+', str(g('업무라우트') or '').strip().upper()) if x]
        if not name:
            errs.append(f'{where}: 이름 빈칸')
        elif st == '출근':
            bad = [x for x in routes if not ROUTE_RE.fullmatch(x)]
            if not routes or bad:
                errs.append(f"{where} {d} {name}: 출근인데 업무라우트 '{g('업무라우트')}' 이상")
        elif st == '휴무':
            if routes:
                errs.append(f"{where} {d} {name}: 휴무인데 업무라우트 '{g('업무라우트')}' 있음")
        else:
            errs.append(f"{where} {d} {name}: 업무상태 '{st}' (출근/휴무 만)")
        rows.append((d, camp, name, st, rot, routes))
    if errs:
        fail(f"'{os.path.basename(path)}' 검증 실패 {len(errs)}건", errs)
    if other:
        print(f"   (무시) '{os.path.basename(path)}' 의 다른 캠프/웨이브: {sorted(other)}")
    return rows


def main():
    paths = []
    for a in sys.argv[1:]:
        paths += sorted(glob.glob(a)) if any(c in a for c in '*?[') else [a]
    paths = [p for p in paths if p.lower().endswith('.xlsx') and not os.path.basename(p).startswith('~$')]
    if not paths:
        raise SystemExit(__doc__)
    paths.sort(key=stamp)                                  # 오래된 → 최신

    by_date, src_of = {}, {}                               # 날짜 → 그 날짜를 가진 최신 파일의 행들
    for p in paths:
        rows = load(p)
        ds = sorted({r[0] for r in rows})
        replaced = [d for d in ds if d in by_date]
        for d in ds:
            by_date[d] = [r for r in rows if r[0] == d]
            src_of[d] = os.path.basename(p)
        print(f"[읽음] {os.path.basename(p)} | {ds[0]} ~ {ds[-1]} ({len(ds)}일, {len(rows)}줄)"
              + (f' | 이전 파일의 {len(replaced)}일을 대체' if replaced else '') if ds else f'[읽음] {os.path.basename(p)} | 해당 행 없음')
    if not by_date:
        fail(f'{"·".join(CAMPS)} {WAVE}(주간) 행이 하나도 없음')
    dates = sorted(by_date)

    # 주(일~토)별로 따로: 사람 목록·순서·두 캠프 등록자 위치는 그 주 데이터로만 정함
    #   → 다음 주 파일을 추가해도 이미 올린 주의 모양이 바뀌지 않음
    def sunday(d):
        x = datetime.date.fromisoformat(d)
        return (x - datetime.timedelta(days=(x.weekday() + 1) % 7)).isoformat()
    week_dates = {}
    for d in dates:
        week_dates.setdefault(sunday(d), []).append(d)
    weeks_out = [build_week(wd, by_date) for _, wd in sorted(week_dates.items())]

    # 날짜가 이어지지 않으면 안내만 (빠진 주가 있을 수 있음)
    d0 = datetime.date.fromisoformat(dates[0])
    gaps = [(d0 + datetime.timedelta(days=i)).isoformat() for i in range((datetime.date.fromisoformat(dates[-1]) - d0).days + 1)]
    gaps = [d for d in gaps if d not in by_date]
    if gaps:
        print(f'   (참고) 빠진 날 {len(gaps)}일: {gaps[0]} ~ {gaps[-1]} 등 — 그 주는 게시판에 없음으로 나옴')

    for w in weeks_out:
        print(f"[주] {w['dates'][0]} ~ {w['dates'][-1]} | " + ' · '.join(f"{c['name']} {len(c['people'])}명" for c in w['camps']))
    data = {'title': '부산 주간 전체 스케쥴', 'generated': datetime.date.today().isoformat(),
            'dates': dates, 'weeks': weeks_out}
    with io.open(OUT, 'w', encoding='utf-8', newline='\n') as f:
        f.write('// 자동 생성 파일 — 직접 수정하지 말고 tools/convert_day.py 로 다시 만드세요 (이름·라우트·휴무만, 아이디 없음)\n')
        f.write('window.DAY_DATA = ')
        json.dump(data, f, ensure_ascii=False, separators=(',', ':'))
        f.write(';\n')
    print(f'[저장] {OUT} ({os.path.getsize(OUT):,} bytes)')


def build_week(dates, by_date):
    """한 주 → {'dates': [...], 'camps': [{'name', 'people': [{'name', 'days'}]}]}"""
    # 사람·날짜별로 합치기: 출근 줄이 실제 근무
    cell, errs = {}, []
    order = {c: [] for c in CAMPS}                         # 캠프별 사람 순서 (처음 나온 순서 = 어드민 순서)
    work_cnt = {}                                          # (이름, 캠프) → 출근 일수
    for d in dates:
        for (dd, camp, name, st, rot, routes) in by_date[d]:
            if name not in order[camp]:
                order[camp].append(name)
            prev = cell.get((name, d))
            if st == '출근':
                work_cnt[(name, camp)] = work_cnt.get((name, camp), 0) + 1
                if prev and prev != '휴':
                    errs.append(f'{d} {name}: {prev["c"]}·{camp} 둘 다 출근')
                    continue
                v = {'r': ','.join(routes), 'c': camp}
                if rot not in FULL_ROT:
                    v['t'] = rot
                cell[(name, d)] = v
            elif prev is None:
                cell[(name, d)] = '휴'
    if errs:
        fail('한 사람이 같은 날 두 캠프에서 출근', errs)

    # 사람은 출근을 더 많이 한 캠프 아래 한 번만 (동률이면 먼저 나온 캠프)
    home = {}
    for camp in CAMPS:
        for name in order[camp]:
            best = home.get(name)
            if best is None or work_cnt.get((name, camp), 0) > work_cnt.get((name, best), 0):
                home[name] = camp
    moved = sorted(n for n in home if sum(n in order[c] for c in CAMPS) > 1)
    camps_out = []
    for camp in CAMPS:
        people = []
        for name in order[camp]:
            if home[name] != camp:
                continue
            days = []
            for d in dates:
                v = cell.get((name, d))
                if isinstance(v, dict) and v['c'] == camp:
                    v = {k: x for k, x in v.items() if k != 'c'}    # 자기 캠프면 캠프 표시 생략
                days.append(v)
            people.append({'name': name, 'days': days})
        camps_out.append({'name': camp, 'people': people})
    if moved:
        print(f'   (참고) {dates[0]} 주: 두 캠프에 등록된 사람은 출근 많은 캠프에 한 번만: '
              + ', '.join(f'{n}→{home[n]}' for n in moved))
    return {'dates': dates, 'camps': camps_out}


if __name__ == '__main__':
    main()
