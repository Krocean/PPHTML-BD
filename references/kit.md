# 디자인 키트 (assets/kit, assets/templates)

테마·부품·차트·발표 도구 묶음이다. **두 엔진 모두에서 그대로 돈다**(최신 엔진에서는 쓴 그대로, 대비용 기존 엔진에서도 변환 없이). 그래서 키트로 쓴 장은 WebView2 가 없는 오래된 PC 에서도 살아 있다.
전부 HTML과 SVG로 그리므로 PowerPoint가 화면을 몇 배로 키워도 글자와 선이 또렷하다.
여기 적힌 것은 실제 슬라이드 쇼(숨은 화면, 4K)에서 눌러 가며 확인한 것이다(2026-11, PowerPoint LTSC 2024). 확인하지 못한 것은 그렇게 적었다.

**키트는 틀이 아니라 도구 상자다.** 페이지는 그냥 웹 페이지이고(`modern-web.md`), 사용자의 내용에 맞게 자유롭게 짠다(`new-deck.ps1 -Blank`, 출발점은 `assets/page-blank.html`). 여기서는 필요한 것만 꺼내 쓴다:
테마 변수(색·글꼴·모양을 한 곳에서 바꾸게 해 준다), 부품, 차트, 발표 도구. 맞는 부품이 없으면 그 장의 `<style>` 과 스크립트로 직접 만든다. 키트 밖의 어떤 CSS·JavaScript 와도 섞어 쓴다.

**두 가지 쓰는 법.**
- **자유 페이지** `<html class="ph-free">` (빈 페이지가 이렇다): 키트의 고정 틀(1280 × 720 에 꽉 찬 `overflow: hidden` 화면, 글자 선택 막기)이 적용되지 않는다. 페이지가 스스로 크기와 스크롤과 바탕을 정한다. 테마 변수와 발표 도구, 그리고 골라 쓰는 부품만 남는다.
- **키트 페이지** (예시 페이지 23종이 이렇다, 클래스 없음): 아래 1장의 뼈대(`ph-page`, `PH.nav`)를 쓰는 고정 화면.

아래에서 "IE11", "변환 없이", 4장의 끊김 규칙, 10장의 ES5 규칙은 **기존 엔진에서도 돌아야 하는 장**을 쓸 때의 이야기다. 최신 엔진만 생각하는 장에는 해당하지 않는다(실측 60프레임, 문법 제한 없음).
디자인을 레퍼런스나 `design.md` 에서 가져오는 법은 `design-from-reference.md`.

## 목차
1. 페이지의 뼈대
2. 테마 (17종, 회사 색에 맞추기)
3. 보기 좋게 만드는 규칙
4. 끊기지 않게 만드는 규칙 (실측)
5. 배치·글자 클래스
6. 부품
7. 스크립트 없이 되는 동작
8. 덱 도구: 슬라이드 고르기, 고치기, 펜, 확대, 보관 (`kit-deck.js`)
9. `PH.*` 함수 (차트, 표, 슬라이더, 일정선 …)
10. 내 스크립트를 쓸 때의 규칙
11. 예시 페이지(템플릿) 23종과 "바꿀 곳"
12. 슬라이드 일부 영역에 쓸 때

## 1. 페이지의 뼈대

`new-deck.ps1` 이 만든 폴더는 `kit\`(키트 사본) · `pages\`(페이지) · `job.json` 이다. 페이지는 이렇게 생겼다:

```html
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<title>핵심 지표</title>                                   <!-- 슬라이드 고르기 화면에 나오는 이름 -->
<link rel="stylesheet" href="../kit/theme-cobalt.css">   <!-- 테마는 하나만, kit.css 보다 먼저 -->
<link rel="stylesheet" href="../kit/kit.css">
<style> /* 이 페이지만의 배치 */ </style>
</head>
<body>
<div class="ph-page"> ... 본문 ... </div>
<script src="../kit/kit.js"></script>                       <!-- 본문 뒤, 내 스크립트 앞 -->
<script src="../kit/kit-deck.js"></script>                  <!-- 덱 도구(8장). 빼도 된다 -->
<script>
PH.nav({ brand: "발표 이름", links: [["표지", 1], ["지표", 2]], page: 2 });
PH.ring();
PH.deck();
/* 내 코드 */
</script>
</body>
</html>
```

빌드가 CSS·JS를 페이지 안으로 합치고, 테마의 `--ph-*` 변수를 값으로 풀어 쓰고, 쓰인 글꼴을 내장하고,
큰 그림을 줄이고, 각 페이지에 덱 정보(`window.__ihDeck`: 몇 번째 장인지, 모든 장의 제목과 작은 그림)를 넣는다.

## 2. 테마

한 파일이 색·글꼴·모서리·그림자를 모두 정한다. 같은 페이지가 테마 한 줄로 전혀 다른 화면이 된다. 모습은 `assets/kit/themes.png`.

| 테마 | 바탕 | 무엇을 따랐나 | 글꼴 (모두 내장되어 다른 PC에서도 같다) |
|---|---|---|---|
| `cobalt` | 밝음 | 키트 기본. 흰 바탕 + 파랑 | Noto Sans KR |
| `paper` | 밝음 | 종이색 + 주황, 읽는 글 | Noto Serif KR 제목 |
| `ink` | 어두움 | 큰 화면 발표 | Noto Serif KR 제목 |
| `forest` | 밝음 | 연한 초록 | Noto Serif KR 제목 |
| `snow` | 밝음 | 순백 + 검정, 주황 한 색 | Noto Sans KR |
| `material` | 밝음 | 구글 Material 3 기본 색 체계(공개 색 도구로 계산) | Roboto + Noto Sans KR |
| `fluent` | 밝음 | 마이크로소프트 Fluent 2 공개 토큰 | Segoe UI(Windows 기본) + Noto Sans KR |
| `carbon` / `carbon-dark` | 밝음 / 어두움 | IBM Carbon 공개 테마(White / Gray 100). 각진 모서리 | IBM Plex Sans KR |
| `cloudscape` | 밝음 | AWS Cloudscape 공개 토큰 | Open Sans + Noto Sans KR |
| `primer` / `primer-dark` | 밝음 / 어두움 | 깃허브 Primer 공개 토큰 | Segoe UI + Pretendard |
| `cupertino` | 밝음 | 애플 제품 소개 화면의 느낌(직접 맞춘 값) | Pretendard |
| `geist` | 밝음 | 버셀의 개발 도구 화면 느낌(직접 맞춘 값) | Geist + Pretendard |
| `indigo` | 밝음 | 스트라이프의 핀테크 화면 느낌(직접 맞춘 값) | Inter + Pretendard |
| `midnight` | 어두움 | 리니어의 어두운 제품 화면 느낌(직접 맞춘 값) | Inter + Pretendard |
| `seoul` | 밝음 | 한국 소비자 앱의 느낌(직접 맞춘 값) | Pretendard |

- "공개 토큰"이라 적은 것은 각 회사가 공개한 디자인 값 패키지에서 뽑았고(`dev/src/themes.js` 가 만든다), "느낌"이라 적은 것은 손으로 맞춘 근사치다.
  어느 것도 그 회사의 공식 자료가 아니며, 회사 로고나 이름을 넣어 공식 자료처럼 보이게 쓰지 않는다. 각 회사의 전용 글꼴은 들어 있지 않고 공개 글꼴로 대신했다.
- 테마를 바꾸려면 페이지 위쪽의 `theme-*.css` 한 줄만 바꾼다. 한 덱의 페이지는 같은 테마를 쓴다(강조하려는 한두 장만 어두운 테마로 바꾸는 것은 괜찮다).
- **내 디자인으로 테마 만들기**: 레퍼런스(HTML·웹 주소·PPT·그림)가 있으면 `get-style.ps1` 이 읽어서, 값만 알고 있으면 `new-theme.ps1 -Bg … -Ink … -Accent … -Display … -Radius …` 가
  아래 변수 전체를 채운 `theme-이름.css` 와 미리보기 그림을 만든다(옅은 색·차트 색·눌림 색은 준 값에서 계산한다). 절차와 옵션은 `design-from-reference.md`.
- 손으로 맞출 때: 가장 가까운 테마 파일을 `theme-우리회사.css` 로 복사해 값만 바꾼다. 이름은 전부 그대로 두고, 한 파일 안에서 한 번씩만 정의한다.
  꼭 바꿀 것: `--ph-accent`(강조색) · `--ph-accent-ink`(그 위의 글자) · `--ph-accent-soft`(아주 옅은 바탕) · `--ph-accent-deep`(진한 글자용) · `--ph-c1`~`--ph-c6`(차트 색) · `--ph-press`(단추).
  글꼴은 `--ph-display` / `--ph-body`. 회사 글꼴이 PC에 설치돼 있고 내장이 허용된 글꼴이면 빌드가 내장한다.
- 페이지의 `<style>` 에서 `--ph-*` 를 다시 정의하지 않는다(두 번 정의되면 자동으로 풀어 쓰지 못한다). 색을 쓸 때는 `var(--ph-accent)` 처럼 가져다 쓴다.

변수 전체: 색 `--ph-bg` `bg2` `surface` `ink` `ink2` `ink3` `line` `accent` `accent-ink` `accent-soft` `accent-deep` `good` `bad` `warn`(각각 `-soft`) `c1`~`c6` `hover` `hover2` `press` `press-ink` ·
글자 `--ph-display` `display-weight` `display-tracking` `body` `strong-weight` `label-tracking` `label-case` · 모양 `--ph-radius` `radius-s` `radius-btn` `card-border` `shadow` `shadow-pop`.

## 3. 보기 좋게 만드는 규칙

HTML을 쓰는 이유는 슬라이드보다 나은 화면을 만들기 위해서다. **슬라이드처럼 만들지 말고 잘 만든 웹 제품 화면처럼 만든다.**

- 위에는 사이트 머리(`PH.nav`)를 둔다. 발표 중 어디로든 건너뛸 수 있고, 화면이 "제품"처럼 보인다.
- 한 장에 말하려는 것 하나. 제목은 그 장의 결론을 문장으로 쓴다(`references/storytelling.md`). 강조할 낱말만 `<em>`.
- 여백을 아끼지 않는다. 카드 안쪽 여백 22px 이상, 요소 사이 16~28px. 글자는 13px 아래로 내려가지 않는다(발표 화면에서 안 읽힌다).
- 누를 수 있는 것은 눌러 보고 싶게: 마우스를 올리면 반응하고(`ph-hot`, `ph-chip`, `ph-btn` 은 스스로 한다), 눌렀을 때 화면의 다른 부분이 바뀐다.
- 숫자는 큰 글자(`ph-num`, `ph-kpi-value`)로, 단위는 작게. 증감은 `ph-good`/`ph-bad` 색으로.
- 색과 모서리, 그림자를 페이지에서 직접 정하지 않는다. 테마 변수를 쓰면 테마를 바꿔도 화면이 깨지지 않는다.
- 장식용 흐림·그림자 남발, 여러 색 섞어 쓰기, 테두리 두른 상자 나열, 맑은 고딕 본문은 피한다(슬라이드처럼 보이게 하는 것들이다).
- `<canvas>` 는 확대하면 흐려진다. 그림은 SVG와 HTML로 그린다.

## 4. 끊기지 않게 만드는 규칙 (실측)

**이 장은 기존 엔진(IE11, 대비용) 이야기다.** 최신 엔진에서는 전체 화면 캔버스·흐림·혼합 모드를 겹쳐 움직여도 60프레임이 나왔다(`modern-web.md` 4장). 기존 엔진에서도 부드러워야 하는 장만 아래를 지킨다.

기존 엔진은 **움직이는 것의 넓이만큼** 매 프레임을 다시 그린다. 4K 화면에서 페이지는 3배로 확대되어 있으므로 큰 것을 움직이면 바로 끊긴다.
아래는 숨은 4K 슬라이드 쇼에서 잰 값이다(아무것도 안 움직일 때 약 53fps, 잴 때마다 ±10 흔들린다).

| 움직인 것 | 초당 프레임 | 판단 |
|---|---|---|
| 큰 카드 여섯 장을 함께 올리거나 흐리게 나타내기 | 14~24 (한 프레임 100ms 넘음) | 하지 않는다 |
| 화면 전체의 투명도 바꾸기 | 23 | 하지 않는다 |
| 스크립트가 매 프레임 SVG 막대 36개 다시 그리기 | 20 | 하지 않는다 |
| HTML 막대를 CSS `transform: scaleY` 전환으로 키우기 | 36 | 쓴다 (`PH.bars`) |
| 카드 안의 작은 글자만 올라오기 | 35~47 | 쓴다 (`ph-rise`) |
| 숫자 세어 올리기(11번에 나눠) | 32~39 | 쓴다 (`PH.countUp`) |
| 선이 그려지기, 도넛이 차오르기 | 47 | 쓴다 |
| 슬라이더 끌기 | 57 | 쓴다 |

키트는 이 결과대로 만들어져 있다(템플릿 23종: 나타날 때 36~53fps, 누를 때 40~54fps, 가장 긴 프레임 35~75ms):

- **카드는 나타날 때 움직이지 않는다.** `ph-rise` 를 큰 것(넓이 70,000px² 초과, 또는 카드를 품은 것)에 붙이면 키트가 `ph-still` 을 붙여 그냥 보이게 한다. 움직이는 것은 그 안의 글자와 차트다.
- PowerPoint는 쇼를 시작할 때 모든 페이지를 미리 불러온다. 그때는 움직임 없이 완성된 모습으로 그려 두고(타이머도 돌지 않는다), 그 장이 나타날 때 움직임을 튼다.
- 장이 처음 나타날 때 0.7초 동안 프레임을 세어 22fps가 안 되면 그 페이지의 움직임을 모두 끈다(`html.ph-lite`). 느린 PC에서는 끊기는 대신 바로 보인다.
- 내 페이지에서 지킬 것: 움직임은 작은 요소의 `opacity` 와 `transform` 으로. 흐림이 40px 넘는 그림자, `transition: all`, 끝없이 도는 큰 애니메이션은 쓰지 않는다(빌드가 `perf-*` 경고로 알려 준다).
  슬라이더처럼 값이 빠르게 바뀌는 조작에서 차트를 매번 다시 그리지 않는다. 숫자만 바로 바꾸고 차트는 손이 멈춘 뒤(160ms)에 그린다(`whatif`, `ask` 템플릿).
- 빌드 보고의 `fps while appearing` 이 30 아래면 경고가 붙는다. 그 장에서 큰 것이 움직이고 있지 않은지 본다.

## 5. 배치·글자 클래스

| 클래스 | 뜻 |
|---|---|
| `ph-page` | 화면 전체를 채우는 본문 틀(머리 아래 96px 부터, 좌우 64px). 머리가 없는 장은 `ph-page ph-plain` |
| `ph-row` | 가로로 나란히. 자식은 같은 폭으로 나뉘고 사이는 20px. 간격: `ph-g12` `ph-g28` `ph-g40` |
| `ph-row` 의 자식에 `ph-w25` `ph-w33` `ph-w40` `ph-w50` `ph-w60` / `ph-auto` | 그 자식의 폭 고정 / 내용만큼 |
| `ph-row ph-mid` / `ph-end` | 세로 가운데 / 아래 맞춤 |
| `ph-stack` | 세로로 쌓기, 사이 16px. `ph-s8` `ph-s24` `ph-s32` |
| `ph-colflex` + `ph-fill` | 세로 flex 틀과, 남은 높이를 다 차지하는 자식 |
| `ph-mt8/16/24/32/48`, `ph-right`, `ph-center` | 위 여백, 정렬 |

- 높이는 페이지의 `<style>` 에서 px 로 준다(`.charts .ph-card { height: 318px; }`). 본문에 쓸 수 있는 높이는 720 − 96 − 48 = **576px** 이다.
- 카드 안의 차트 자리는 크기가 정해진 상자여야 한다. 흔한 방법: 카드(`ph-card` 는 `position: relative`) 안에 `position: absolute; left/right/top/bottom` 으로 잡은 `<div>`.
- 정확한 자리가 필요하면 `position: absolute` 와 px 을 쓴다. 캔버스 크기는 변하지 않으므로 안전하다.
- `gap`, `display: grid` 는 쓰지 않는다(쓰면 빌드가 자동 변환을 거치게 되어 Edge가 필요해진다).

글자: `ph-label`(작은 머리말, 앞에 `<i></i>` 를 넣으면 짧은 선) · `ph-mega` 92px · `ph-display` 58px · `ph-h1` 40px · `ph-h2` 27px · `ph-h3` 18px ·
`ph-lead` 19px 보조 문단 · `ph-small` 13.5px · `ph-muted` / `ph-faint` · `ph-num`(제목 글꼴의 숫자) · `ph-good` `ph-bad` `ph-warn` `ph-accent`.
제목 안의 `<em>` 은 강조색이 된다. 제목의 굵기와 자간은 테마가 정한다.

## 6. 부품

| 부품 | 마크업 |
|---|---|
| 카드 | `<div class="ph-card">`. `ph-flat` 그림자 대신 선 · `ph-tint` 옅은 강조 바탕 · `ph-hot` 누를 수 있음 · `is-on` 선택됨 |
| 단추 | `<span class="ph-btn">글자<i>→</i></span>`, 보조 단추 `ph-btn ph-ghost` |
| 칩 | `<span class="ph-chip is-on">` (거르기, 지표 고르기) |
| 묶음 단추 | `<div class="ph-seg"><span class="is-on">1분기</span><span>2분기</span></div>` |
| 탭 | 칩/묶음 단추에 `data-tab="p1"`, 내용은 `<div class="ph-panel is-on" id="p1">` (같은 부모 아래) |
| 배지 | `<span class="ph-badge ph-good">정상</span>` (`ph-bad` `ph-warn` `ph-plain`) |
| 숫자 카드 | `<div class="ph-card ph-kpi"><div class="ph-kpi-label">…</div><div class="ph-kpi-value"><span data-count="1284"></span><small>명</small></div><div class="ph-kpi-delta ph-good">+12%</div></div>` |
| 진행 막대 | `<div class="ph-meter"><i data-w="72"></i></div>` (장이 나타날 때 0에서 72%까지 찬다. 스크립트로는 `i.style.width = "72%"`) |
| 단계 | `<div class="ph-step"><span class="ph-step-n">1</span><div class="ph-step-t">제목</div><div class="ph-step-d">설명</div></div>` + `PH.steps` |
| 질문 | `<div class="ph-acc"><div class="ph-acc-q">질문</div><div class="ph-acc-a"><div>답</div></div></div>` (부모에 `data-single`: 하나만 열림) |
| 스위치 | `<span class="ph-switch"><i></i></span>` (누르면 `is-on`, `ph-change` 이벤트) |
| 구분선·띠 | `<hr class="ph-rule">`, `<div class="ph-band">` |

상태는 클래스로 표시한다: `is-on`(선택·열림), `is-done`(지난 단계), `is-off`(범례에서 끔).

## 7. 스크립트 없이 되는 동작

| 마크업 | 동작 |
|---|---|
| `data-goto="3"` | 누르면 3번 슬라이드로 |
| `data-next`, `data-prev` | 다음 / 이전 슬라이드. **길게 누르면 슬라이드 고르기**(`PH.deck()` 을 부른 페이지) |
| `data-tab="패널id"` | 탭 전환 (`ph-tab` 이벤트가 난다) |
| `data-count="1284"` (`data-digits="1"`) | 장이 나타날 때 0부터 세어 올라간다 |
| `ph-meter > i[data-w]` | 장이 나타날 때 채워진다 |
| `ph-rise`, `ph-d1`~`ph-d6` | 장이 나타날 때마다 차례로 올라온다(작은 것만, 4장) |
| `ph-frag` | 처음엔 숨어 있다가 페이지의 "다음" 단추를 누를 때마다 하나씩 나타난다. 다 나오면 다음 장으로 |
| `data-zoom` | 누르면 화면 가운데로 크게, 다시 누르면 제자리 |
| `data-tip="설명"` | 마우스를 올리면 설명이 뜬다 |

## 8. 덱 도구 (`kit-deck.js`)

슬라이드가 웹 페이지라서 할 수 있는 것들이다. 페이지에서 `PH.deck()` 을 한 번 부르면 오른쪽 아래에 작은 도구 막대(이전 · 쪽 번호 · 다음 · 연필 · 펜)가 준비된다.
**막대는 평소에 보이지 않는다.** 마우스를 화면 오른쪽 아래 구석으로 가져가면 나타나고, 도구(고치기·펜·슬라이드 고르기)를 쓰는 동안에는 계속 보인다. 슬라이드 그림(편집 화면)에는 찍히지 않는다.
(사용자가 "오른쪽 아래에 늘 떠 있어서 불편하다"고 해서 바꾼 것이다. 늘 보이게 되돌리지 않는다.)
`PH.deck({ dock: false })` 로 막대를 아예 없애고, `{ edit: false }` `{ pen: false }` `{ progress: false }` 로 하나씩 끈다. `{ autoNext: 12 }` 는 12초 뒤 다음 장으로 넘긴다.

| 도구 | 쓰는 법 | 숨은 쇼에서 확인한 것 |
|---|---|---|
| 슬라이드 고르기 | "다음"·"이전" 단추나 쪽 번호를 **0.45초 넘게 누르고 있으면** 모든 장의 작은 그림이 뜬다. 누른 채로 가려는 장 위로 가서 손을 떼거나, 손을 뗀 뒤 눌러 고른다. 바깥을 누르면 닫힌다 | 22장 덱에서 열림, 그림 위에서 떼어 12번 장으로 이동, 바깥을 눌러 닫힘 |
| 고치기 | 오른쪽 아래 구석에서 연필을 누른다. **글자**: 고칠 글자를 누르고 입력(Enter 나 다른 곳을 누르면 끝). **옮기기**: 글자·카드·차트·그림을 끌어서 옮긴다(두 번 누르면 제자리). **가− / 가+**: 방금 고른 글자를 작게 / 크게. **원래대로**: 이 장에서 고친 것을 모두 되돌린다 | 글자 입력, 카드 옮기기, 크기 바꾸기, 저장 후 다시 열어도 유지, **쇼를 끝낸 뒤 슬라이드 그림에도 고친 글자가 보임** |
| 펜 | 펜을 누르고 화면에 그린다. "지우기"로 지운다. 장을 떠나면 꺼진다(그린 것은 보관하지 않는다) | 선이 그려짐 |
| 확대 | `data-zoom` 을 붙인 요소를 누른다 | 켜짐, 다시 눌러 꺼짐 |
| 하나씩 보이기 | `ph-frag` | "다음"을 한 번 눌러 셋 중 하나가 나타남 |
| 설명 풍선 | `data-tip` | 화면에서 직접 확인하지 못함(마우스를 올리는 동작) |
| 타이머 | `PH.timer(el, 300)`: 5:00 을 보여 주고, 누르면 시작/멈춤, 0:00 에 빨갛게 | 눌러서 시작(5:00 → 4:57), 다시 눌러 멈춤 |
| 보관 | `PH.keep("이름", 값)` 으로 넣고 `PH.kept("이름", 기본값)` 으로 꺼낸다. 투표 수, 켠 스위치처럼 발표 중에 생긴 작은 값 | 투표 수와 스위치가 저장 후 다시 열어도 유지, 쇼를 끝낸 뒤 슬라이드 그림에도 표 수가 보임 |

**고친 것과 보관한 값이 남는 곳**

- PowerPoint: 그 슬라이드의 꼬리표(tag)에 적힌다. PowerPoint를 닫을 때 **저장하겠느냐고 묻고, 저장해야 남는다.** 저장한 파일을 다른 PC에서 열어도 그대로다.
- **쇼 밖에서도 보인다.** 슬라이드 쇼가 끝나면 호스트가 고친 장(무언가를 보관한 장)을 보관한 내용과 함께 한 번 더 불러 그리고, 그 그림으로 슬라이드의 정지 그림을 바꾼다.
  그래서 편집 화면, 축소판, 인쇄·PDF, 매크로가 꺼진 PC에서도 고친 모습으로 보인다. 고친 장 하나에 2초쯤 걸리고(빌드의 `settleMs` 만큼 기다린다), 그동안 PowerPoint가 잠깐 바쁘다.
  그림은 "처음 상태 + 고친 내용"이다: 쇼에서 눌러 바꿔 둔 탭이나 펼친 질문은 반영되지 않고, 하나씩 보이기(`ph-frag`)는 모두 보이는 상태로 찍힌다.
- 고친 `.pptm` 을 `ppt2html.ps1` 로 웹 버전으로 바꾸면 고친 내용이 함께 간다(슬라이드 그림과 살아 있는 페이지 모두).
- 웹 버전 안에서 고친 것은 그 브라우저에만 남는다(다른 사람에게는 보이지 않는다).
- HTML 원본 파일은 바뀌지 않는다. `html2ppt.ps1` 로 다시 빌드하면 새 파일이 만들어지므로 발표 중 고친 내용은 새 파일에 없다. 계속 쓸 문구는 원본 HTML에서도 고친다.

**고치기가 닿는 범위**

- 화면에 보이는 글자는 스크립트가 만든 것까지 누르면 고칠 수 있다(`compare` 의 방안 이름을 고쳐 저장하고 다시 열어 확인). 머리 메뉴, 차트, 표(`PH.table`) 안의 글자와 `ph-noedit` 를 붙인 곳은 빠진다(자료 블록에서 고친다).
- 고친 글자는 "어디에 있는 글자였는지"와 "원래 무엇이었는지"를 함께 기억한다. 단추를 눌러 스크립트가 그 자리의 글자를 다시 쓰면(탭 전환, 숫자 세기) 고친 것은 그 순간 덮인다.
  바뀌는 문구를 오래 남기려면 페이지 아래 자료 블록에서 고친다.
- 원본 HTML의 그 문장을 나중에 다르게 써서 다시 빌드하면, 옛 문장에 대고 고쳤던 내용은 덮어씌우지 않는다.
- 슬라이드 그림이 바뀌는 것은 쇼가 **끝날 때**다. 쇼 도중에 파일을 저장하거나 PowerPoint가 비정상 종료되면 그림은 옛 모습으로 남는다(보관한 내용 자체는 남아 다음 쇼에서 보인다).

함수로도 쓸 수 있다: `PH.picker.open()` · `PH.edit.on()` `off()` `mode("text" | "move")` `reset()` · `PH.pen.on()` `off()` `clear()` · `PH.zoom.on(el)` `off()`.

## 9. `PH.*` 함수

**머리와 포인터**
- `PH.nav({ brand, links: [[이름, 슬라이드번호], …], page: 지금 슬라이드번호, next: true })` — 사이트 머리. 마지막 장에서는 "처음으로" 단추가 된다. 8개가 넘으면 글자가 작아지고 13개가 넘으면 더 작아진다(22개까지 확인).
- `PH.ring()` — 포인터를 따라다니는 고리(누를 수 있는 것 위에서 커진다). 발표 때 어디를 누르는지 청중이 본다.

**차트** — 첫 인자는 **크기가 정해진 빈 요소**. 장이 나타날 때마다 다시 그려지고, 마우스를 올리면 값이 뜬다.
- `PH.bars(el, { labels: ["1월", …], values: [..], unit: "명", max: 40000, digits: 0 }, onPick)` — 막대. 여러 계열은 `series: [{ name, values }, …]`(범례를 눌러 끄고 켠다). `onPick(i, label)` 을 주면 막대를 누를 수 있다.
- `PH.lines(el, { labels, series: [{ name, values }], unit, digits, area: true })` — 선(첫 계열 아래를 옅게 칠한다).
- `PH.donut(el, { items: [{ label, value }, …], unit: "%", center: "가운데 글자", digits }, onPick)` — 왼쪽에 도넛, 오른쪽에 목록. 폭이 높이의 2배쯤 되는 자리에 둔다.
- `PH.spark(el, [3, 5, 4, 8], "c2")` — 작은 추이 선.
- 차트가 돌려주는 것의 `update({ values / series / labels / items … })` 로 내용을 바꾼다(다시 그려진다).
- 색: 계열 색은 테마의 `c1`~`c6` 순서. 특정 색은 계열에 `color: "c3"` 또는 `"#RRGGBB"`.

**표** — `PH.table(el, { cols: [{ key, label, num: true, digits, unit, bar: true, badge: { "값": "ph-good" } }, …], rows: [{…}], sort: "키", desc: true }, onPick)`
제목을 누르면 정렬, 줄을 누르면 `onPick(row)`. 돌려주는 것: `update(rows)`(거른 목록으로 바꾸기), `pick(row)`.

**조작**
- `PH.slider(el, { min, max, step, value }, onChange)` — 끌 수 있는 막대. `set(v)` / `get()`.
- `PH.steps(box, onChange(i, el), 자동전환ms)` — `.ph-step` 묶음. 누르면 그 단계가 `is-on`, 앞 단계는 `is-done`. `set(i)`.
- `PH.timeline(el, [{ when, what, … }, …], onPick(i, item), 시작번호)` — 일정선을 그린다(높이 150px).
- `PH.selectTab("패널id")` — 스크립트로 탭 고르기.

**움직임과 시간**
- `PH.onShow(fn)` / `PH.onHide(fn)` — 장이 나타날 때마다 / 떠날 때.
- `PH.wait(ms, fn)`, `PH.every(ms, fn)`, `PH.stop(id)` — 장을 떠나면 저절로 멈추는 타이머(`setTimeout` 대신 이것을 쓴다). 미리 불러올 때는 `wait` 가 바로 실행되고 `every` 는 돌지 않는다.
- `PH.animate(ms, function (t) { … }, done)` — `t` 가 0에서 1까지(감속 곡선 적용). 작은 것에만 쓴다.
- `PH.countUp(el, 값, 소수자리, ms)`, `PH.type(el, "글", done)`(타자 치듯 나타남).
- `PH.onResize(fn)` — 화면 배율이 바뀌었을 때. 직접 그린 그림을 다시 그릴 곳.
- `PH.still()` — 움직임이 꺼져 있으면 참(미리 불러오는 중이거나 느린 PC). 직접 만든 움직임은 이때 마지막 모습을 바로 그린다.

**작은 도구** — `PH.$(sel)` `PH.$$(sel)`(배열) · `PH.make(tag, class, text, parent)` · `PH.svg(tag, attrs, parent)` · `PH.on(el, type, fn)` · `PH.fire(el, type, detail)` ·
`PH.has/add/remove/toggle(el, class)` · `PH.fmt(1234.5, 1)` → `"1,234.5"` · `PH.color("c1" | "line" | "ink" | "surface" | "accent")` · `PH.goto(n)` `PH.next()` `PH.prev()`.

## 10. 내 스크립트를 쓸 때의 규칙

- **ES5 문법으로 쓴다**: `var`, `function () {}`, 문자열은 `+` 로 잇기. (최신 문법을 쓰면 빌드가 Edge로 자동 변환하지만, 키트 페이지는 변환 없이 도는 쪽이 단순하고 가볍다.)
- 화면을 채우는 코드는 `PH.onShow(function () { … })` 안에 둔다. PowerPoint는 쇼를 시작할 때 모든 페이지를 미리 불러오고, 그 장이 나타날 때 `onShow` 를 부른다.
- **불러올 때 잰 크기를 믿지 않는다.** PowerPoint는 페이지를 불러온 뒤에 배율을 맞춘다. 크기를 재서 그리는 것은 `onShow` 와 `PH.onResize` 에서.
- 요소 찾기는 `PH.$`, 만들기는 `PH.make`. `innerHTML` 로 한꺼번에 써도 된다(SVG 안쪽은 `PH.svg` 로).
- 사용자의 데이터는 스크립트 맨 위의 `var DATA = …` 한 곳에 모은다. 나중에 숫자만 바꾸기 쉽다.
- 글자 입력은 고치기 도구처럼 `contentEditable` 요소로 받는다(숨은 쇼에서 글자가 들어가는 것을 확인). `<input>` 입력창은 확인하지 못했다.
- 문서 전체의 `click` 을 가로채는 코드를 쓰면 호스트의 신호(숨은 `#ihCmd` `#ihShow` `#ihHide` 를 "누르는" 것)까지 막지 않게 한다: 그 셋이 대상이면 그냥 지나간다.

## 11. 예시 페이지(템플릿) 23종과 "바꿀 곳"

기본 길은 빈 페이지에 내용에 맞는 화면을 새로 짜는 것이다. 아래 23종은 "이런 화면이 된다"는 견본이고, 내용이 딱 맞을 때 가져와 고쳐 쓰거나 부품 쓰는 법을 베끼는 용도다.
모습은 `assets/templates/gallery.png`, 목록은 `new-deck.ps1 -List`, 가져오기는 `-Pages "cover,blank,kpi"`(`blank` 는 빈 페이지). 예시 내용은 모두 지어낸 것이다(커피 로스터리, 머그, 빵집 앱).

| 템플릿 | 무엇 | 바꿀 곳 |
|---|---|---|
| `cover` | 표지 | 제목·설명 문구, 오른쪽 카드의 숫자 셋(`data-count`)과 `PH.spark` 값 |
| `agenda` | 목차 | 카드 문구와 `data-goto`(갈 슬라이드 번호) |
| `statement` | 한 문장 | 큰 문장, 뒷받침 문장, `ph-frag` 숫자 셋 (HTML) |
| `bignumber` | 큰 숫자 | `FACTS` (칩, 숫자, 단위, 뜻, 설명) |
| `problem` | 문제와 해법 | `VIEWS` ("지금" / "쓰고 난 뒤" 카드 세 장씩) |
| `market` | 시장 | `TIERS` (바깥 원부터, 숫자와 셈법 세 줄) |
| `matrix` | 경쟁 | `AXES`, `ITEMS` (x, y 0~100, `us: true`) |
| `features` | 기능 | 왼쪽 `ph-step` 문구(HTML), `FEATURES`(상태별 숫자·설명·그림 모습). 그림은 내 제품 그림으로 |
| `pricing` | 가격 | `PLANS` (한 번에 / 달마다 값, 들어 있는 것, `best`) |
| `kpi` | 지표 | `DATA` (기간 이름 → 지표 4개, 월별 값, 비중) |
| `trend` | 추이 | `LABELS`, `METRICS` (지표마다 계열 1~3개와 읽는 법 세 줄) |
| `table` | 표 | `ROWS` (한 줄 = 한 항목), 열 정의(`cols`), 거르기 이름(`FILTERS`) |
| `timeline` | 일정 | `STEPS`, `NOW`(진행 중인 단계) |
| `process` | 절차 | 왼쪽 단계 문구(HTML), `DETAIL`, 자동 전환 시간 |
| `compare` | 비교 | `CRITERIA`, `OPTIONS`(기준별 점수) |
| `whatif` | 가정 바꿔 보기 | `NOW`, `START`, `calc()` 의 계산식, 막대 범위 |
| `team` | 팀 | `TEAM` (이름, 맡은 일, 햇수, 이유 세 줄) |
| `quote` | 후기 | `QUOTES` (한 말, 이름, 어떤 사람) |
| `ask` | 투자 요청 | `START`, `BURN`, `USE`(쓰임새 비율), `GOALS`(몇 개월째에 무엇) |
| `poll` | 현장 투표 | `OPTIONS`. 표는 `PH.keep` 으로 자료 안에 남는다 |
| `faq` | 질문 | 질문과 답(HTML), `data-tip` 설명 |
| `closing` | 정할 일 | `TODOS`. 켠 스위치는 자료 안에 남는다 |
| `end` | 끝 | 큰 문장, 다음에 할 일 카드 셋(HTML), `PH.timer` 초 |

묶음(`-Recipe`): `review` 분기·월간 보고 · `pitch` 투자 유치 · `keynote` 제품 발표 · `report` 업무 보고 · `workshop` 회의. 순서의 이유는 `references/storytelling.md`.

머리 메뉴(`PH.nav` 의 `links`)는 `new-deck.ps1` 이 장 수에 맞게 넣어 준다. 장을 더하거나 빼면 모든 페이지의 `links` 와 `page` 를 같이 고친다.
템플릿은 출발점이다. 내용에 맞지 않으면 같은 부품으로 페이지를 새로 쓴다(예시 덱 "반죽"의 제품 장은 `features` 의 뼈대에 다른 그림을 넣은 것이다).

## 12. 슬라이드 일부 영역에 쓸 때

템플릿은 1280 × 720 전체용이다. 기존 슬라이드의 일부(예: 아래쪽 65% = 1280 × 468)에 넣을 때는:
- 사이트 머리를 빼고 `ph-page ph-plain` 을 쓰되 위 여백을 줄인다(`style="padding: 20px 28px"`). 슬라이드에 이미 제목이 있으니 페이지 안 제목도 뺀다.
- 부품은 그대로 쓴다. 한 영역에 차트 하나 + 숫자 두세 개 정도가 알맞다.
- 덱 도구가 필요 없으면 `PH.deck({ dock: false })` 로 막대를 없애거나 `kit-deck.js` 를 뺀다(막대는 평소에 보이지 않으므로 그대로 둬도 화면을 가리지는 않는다).
- 바탕색을 슬라이드와 맞춘다(흰 슬라이드면 `snow` 나 `cobalt`, 테마 파일의 `--ph-bg` 를 슬라이드 배경색으로 바꾼 사본).
- `check-html.ps1 -Canvas 1280x468` 처럼 그 영역의 크기로 확인한다.
