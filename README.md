# 유니온물류 스케쥴 게시판

- 사이트: https://gleaming-mermaid-cc23e0.netlify.app  (야간 바로가기: `?g=night`, 주간: `?g=day`)
- `master`에 push하면 Netlify가 1분 안에 자동 배포합니다.

## 구성
| 파일 | 내용 |
|---|---|
| `index.html` | 게시판 화면. 상단 `[주간 \| 야간]` 전환 (선택은 기기에 기억됨) |
| `night-data.js` | 야간 근무표 데이터 — **자동 생성, 직접 수정 금지** |
| `tools/convert_night.py` | 야간 월별 엑셀 → `night-data.js` 변환 + 검증 |

## 매달 업데이트
**주간** — `index.html` 위쪽 `DAY_MONTHS`에 구글시트 '웹에 게시' 링크 한 줄 추가, 가장 오래된 달 한 줄 삭제 (3개월 유지).
연도가 바뀌면 `y`도 바꿀 것 (예: `{ y: 2027, m: 1, url: '...' }`).

**야간** — 새 월별 엑셀(+어드민 양식 zip)을 받으면:
```
python tools/convert_night.py 야간근무표_월별.xlsx 야간근무표_어드민양식.zip
git add night-data.js && git commit -m "야간 근무표 갱신" && git push
```
- 출력은 통째로 새로 만들어지므로 **이번 주가 포함된 파일**을 넣을 것 (필요하면 지난 파일도 같이, 여러 개 가능).
- 엑셀 오타·빈칸·요일 불일치·어드민 불일치·보여줄 기간 축소는 모두 **중단**됩니다. 사유가 출력되니 엑셀을 고친 뒤 다시 실행.
- 규칙·오류별 대처는 `tools/convert_night.py` 맨 위 설명 참고.
