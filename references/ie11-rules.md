# 기존 엔진(IE11)에서도 돌아가게 쓰는 법 — 대비용

**v3부터 이 문서는 대비용 엔진 이야기다.** 슬라이드 쇼는 Edge 의 엔진(WebView2)을 먼저 쓰고, 페이지는 쓴 그대로 돈다(`modern-web.md`). 그때는 아래 규칙이 하나도 적용되지 않는다.
이 문서가 필요한 때는 둘뿐이다: ① WebView2 가 없는 오래된 PC 에서도 그 장이 살아 있어야 할 때(빌드 보고의 `no fallback` 을 `fallback ready` 로 바꾸고 싶을 때) ② `-Engine classic` 으로 기존 엔진용만 만들 때.

기존 엔진은 Windows 에 들어 있는 Internet Explorer 11 엔진(Trident)이다.
이 문서는 (1) 그 엔진용으로 쓸 때의 규칙, (2) 빌드가 자동으로 바꿔 주는 것과 못 하는 것, (3) 직접 고치는 변환표, (4) 페이지에서 쓸 수 있는 연결 기능을 담는다.
디자인 키트(`references/kit.md`)의 부품과 예시 페이지는 아래 규칙을 이미 지켜서 만들어져 있어 두 엔진 모두에서 돈다.

## 목차
1. 캔버스(화면 크기)
2. 빌드가 자동으로 해 주는 것 / 직접 고쳐야 하는 것
3. JavaScript 변환표
4. CSS 변환표
5. HTML 요소
6. 배치 요령
7. PowerPoint와의 연결 (pptHtml)
8. 검사 규칙 끄기
9. 자주 겪는 문제

## 1. 캔버스

- 페이지는 **영역의 pt × 4/3** px 크기의 창으로 보인다. (HTML 1px = 슬라이드 0.75pt, 즉 24px 글자 ≈ 18pt)

| 영역 | 캔버스 |
|---|---|
| 16:9 슬라이드 전체 (960×540pt) | 1280 × 720 |
| 4:3 슬라이드 전체 (720×540pt) | 960 × 720 |
| 16:9 아래쪽 65% (960×351pt) | 1280 × 468 |
| 16:9 왼쪽 절반 (480×540pt) | 640 × 720 |

- 발표 화면이 FHD든 4K든 창이든 캔버스 크기는 같다. 브라우저 확대 기능으로 맞추는 것이라 `vw`, `%`, 미디어 쿼리, 마우스 좌표 계산이 일반 브라우저와 똑같이 동작한다.
- 확대 배율이 정수 %라서 실제 창은 캔버스보다 **작아지지는 않고 최대 1% 정도 넓을 수 있다**(예: 1282×721). 그래서:
  - `html, body { width:100%; height:100%; margin:0; overflow:hidden }` 로 시작하고 배경은 body 에 칠한다.
  - 오른쪽·아래에 딱 붙어야 하는 것은 `right:` / `bottom:` 으로 붙인다.
- 영역이 캔버스보다 좁은 페이지를 넣으면 겹치거나 잘린다. 빌드 보고의 `content is ... but the canvas is ...` 경고와 그림으로 확인한다.

## 2. 자동 / 수동

빌드는 두 단계로 페이지를 엔진에 맞춘다.

**(가) 언제나 하는 것 (Edge 불필요)**
- 상대 경로의 CSS·JS·이미지·글꼴 → 페이지 안으로 합쳐짐(이미지는 data URI).
- `:root`/`html`/`body` 에 **한 번만** 정의된 CSS 변수 → 값으로 치환.
- IE11에 없는 **함수** → 필요할 때 보완 코드가 들어간다:
  `Object.assign/entries/values/fromEntries`, `Array.from/of`, 배열 `find/findIndex/includes/fill/flat/flatMap/at`,
  문자열 `includes/startsWith/endsWith/repeat/padStart/padEnd`, `Number.isFinite/isNaN/isInteger`, `Math.sign/trunc/log2/log10/hypot/cbrt`,
  `NodeList.forEach`, 요소 `matches/closest/remove/append/prepend/before/after/replaceWith/replaceChildren/toggleAttribute`,
  `classList.add/remove` 여러 개·`toggle(이름, 참거짓)`·`replace`, `new CustomEvent()/new Event()`, `Promise`(then/catch/finally/all/race).
- 문서 모드·문자 인코딩·보안 영역 표식·PowerPoint 연결 코드.
- `dataset` 으로 `data-`·`aria-` 속성을 바꿔도 화면이 다시 칠해지게 하는 보정(IE11은 `setAttribute` 일 때만 다시 칠한다).

**(나) 자동 변환 — 페이지가 최신 문법을 쓸 때만, Edge를 도구로 써서**
검사에서 IE11에 없는 문법·CSS가 보이면(또는 엔진에서 스크립트가 실패하면) 그 페이지를 통째로 변환한다. 보고에 `[converted automatically for IE11]` 과 `note: converted: …` 로 나온다.

| 종류 | 하는 일 |
|---|---|
| JavaScript 문법 | 화살표 함수, `let`/`const`, class(비공개 필드 포함), 템플릿 문자열, 전개·나머지, 구조 분해, 기본 인자, `for…of`, `?.`, `??`, `**`, `async`/`await`, 제너레이터 → ES5. 이미 ES5인 스크립트(라이브러리 등)는 건드리지 않는다 |
| JavaScript 기능 | 변환한 스크립트가 있으면 ES2015+ 표준 기능 모음(core-js, 약 270KB)을 페이지에 넣는다: `Symbol`, `Map`/`Set` 순회, `URL` 등 |
| 모듈 스크립트 | `<script type="module">` 에 `import`/`export` 가 없으면 일반 스크립트로 바꿔 문서 끝에서 실행한다 |
| CSS 변수 | 전부 값으로 풀어 쓴다. `.dark { --bg: … }`, `[data-theme="dark"] { … }` 처럼 **클래스·속성으로 바꾸는 변수**는 그 조건이 붙은 덮어쓰기 규칙으로 만들어 주므로, 클래스·속성을 바꾸는 테마 전환이 그대로 동작한다 |
| 화면 크기에 따른 CSS | 캔버스 크기가 고정이므로 `@media`(폭·높이·방향·`prefers-color-scheme: light`·`hover`)는 지금 맞는 것만 남기고, `clamp()`/`min()`/`max()`·`vmin`·`dvh` 는 px 로 계산해 둔다 |
| grid | `display: grid` 를 엔진의 옛 grid(`-ms-grid`)로 바꾸고, 페이지에 넣는 작은 스크립트(레이아웃 어댑터)가 칸 배치를 대신한다: `grid-template-columns/rows/areas`, `repeat()`, `auto-fit`/`auto-fill` + `minmax()`, `span`, `1 / -1`, `grid-area`, `place-items`, `gap`. 스크립트가 자식을 바꾸면 다시 배치한다 |
| flex 의 `gap` | 자식의 여백으로 옮긴다(줄바꿈 flex 포함, 글자와 아이콘이 섞인 경우 포함) |
| 그 밖의 CSS | 중첩, `:is()`/`:where()`, `inset`, `place-*`, 논리 속성(`margin-inline` 등), `#RRGGBBAA`, `rgb(0 0 0 / 50%)`, `color-mix()`/`oklch()`(색을 계산해 고정), `aspect-ratio`, `<img>` 의 `object-fit`, `user-select` 접두사 |
| 글꼴 | 페이지가 쓰는 글꼴을 그 페이지의 글자만 남겨 페이지 안에 넣는다(글꼴 파일의 사용 조건이 내장을 허락할 때만). 이 PC에 없는 Noto Sans/Serif KR 은 스킬의 글꼴 묶음에서 가져온다 |

변환 뒤에는 **최신 브라우저의 그림과 엔진의 그림을 비교**해 `conversion check: 99.x% like` 와 나란히 놓은 그림(`page_NN.compare.png`)을 낸다. 낮게 나오면(`warn`) 그 그림에서 다른 곳을 찾는다.
웹 버전에는 원래의 최신 CSS가 그대로 살아 있다(엔진용 대체 값을 **덧붙이는** 방식이라 최신 브라우저는 원래 값을 쓴다).

**(다) 자동으로 안 되는 것 — 원본을 고친다**
- 효과만 빠지는 것(보고에 `look`): `backdrop-filter`, CSS `filter`, `mix-blend-mode`, `position: sticky`(슬라이드는 스크롤이 없어 대개 무해). 빠져도 읽을 수 있는지 그림으로 확인하고, 필요하면 반투명 배경색 등으로 바꾼다.
- 아예 안 되는 것(보고에 `ERROR`): `import` 가 있는 모듈, `:has()`·`:focus-within`, `@container`, WebGL, `fetch`(로컬 파일이라 받아올 곳이 없다 — 자료는 페이지 안에 적는다), `IntersectionObserver`/`ResizeObserver`/`Proxy`.
- 눈에 안 띄게 틀리는 것: **스크립트가 문자열로 만들어 넣는 인라인 스타일**(`el.innerHTML = '<div style="display:grid;gap:8px">'`)은 변환기가 보지 못한다 → 클래스로 옮긴다.
  **스크립트로 CSS 변수를 읽거나 바꾸는 코드**(`style.setProperty('--x', …)`)는 효과가 없다(`warn css-var-js`) → 클래스를 바꾸거나 값을 직접 넣는다.
  요소 **자신의 인라인 `style="--i: 3"`** 로 주는 변수도 풀지 못한다.
- 외부 주소(`https://…`)의 스크립트·글꼴·그림은 합쳐지지 않는다. 발표장 인터넷에 기대게 되니 파일로 받아 옆에 두고 상대 경로로 건다.
- Edge를 쓸 수 없는 PC(`no automatic conversion (…)`)에서는 (나)가 통째로 빠진다. 그때는 아래 변환표대로 직접 고친다.

끄기: 명령에 `-Convert never`, 한 페이지만은 `<!-- ppt-html-convert: off -->`. 손으로 맞춘 IE11용 페이지는 검사에 걸리는 것이 없으므로 원래 변환되지 않는다.

## 3. JavaScript 변환표

자동 변환이 켜져 있고 Edge를 쓸 수 있으면 이 표의 일은 빌드가 한다. 아래는 **변환 없이 도는 페이지를 직접 쓸 때**(키트 페이지처럼)와 **자동 변환을 쓸 수 없는 PC**를 위한 것이다.

| 쓰면 안 됨 | 이렇게 |
|---|---|
| `const a = 1; let b = 2;` | `var a = 1, b = 2;` |
| `x => x * 2`, `(a, b) => { ... }` | `function (x) { return x * 2; }` — 안에서 `this` 를 썼다면 바깥에 `var self = this;` |
| `` `합계 ${n}개` `` | `"합계 " + n + "개"` |
| `class A { constructor(){} m(){} }` | `function A(){ } A.prototype.m = function(){ };` |
| `[...list]`, `f(...args)` | `Array.prototype.slice.call(list)`, `f.apply(null, args)` |
| `const {a, b} = o;` `const [x, y] = p;` | `var a = o.a, b = o.b;` `var x = p[0], y = p[1];` |
| `function f(a = 1) {}` | `function f(a) { if (a === undefined) { a = 1; } }` |
| `for (const x of list) {}` | `for (var i = 0; i < list.length; i++) { var x = list[i]; }` |
| `a?.b`, `a ?? b` | `a && a.b`, `(a !== null && a !== undefined) ? a : b` |
| `2 ** 3` | `Math.pow(2, 3)` |
| `async function` / `await` | `Promise` 의 `.then()` 또는 콜백 |
| `{ a, b }` (줄임 속성), `{ [k]: v }` | `{ a: a, b: b }`, `var o = {}; o[k] = v;` |
| `import` / `export`, `<script type="module">` | 한 개의 일반 `<script>` |
| `fetch(url)` | 데이터는 페이지 안에 변수로 넣는다(로컬 파일이라 네트워크 요청은 믿을 수 없다) |
| `localStorage` | 변수에 보관(로컬 파일에서는 쓸 수 없다) |
| `e.key === "ArrowLeft"` | `e.keyCode === 37` (IE는 `"Left"`, `"Esc"` 같은 옛 이름을 쓴다) |
| `touchstart` 등 | `mousedown/mousemove/mouseup` (또는 `pointerdown` 등 — IE11 지원) |
| `new Map([[k, v]])`, `new Set([1,2])` | `var m = new Map(); m.set(k, v);` (생성자 인자·반복자 없음) |
| `el.dataset.x` | 그대로 됨 |
| `requestAnimationFrame`, `addEventListener`, `querySelector`, `JSON`, `canvas 2D`, `SVG` | 그대로 됨 |

반복문 안에서 `let` 을 `var` 로 바꾸면 클릭 함수들이 마지막 값만 보게 된다. 값을 함수로 넘겨 가둔다:

```js
// for (let i = 0; i < n; i++) { el[i].onclick = () => pick(i); }
function bind(i) { el[i].onclick = function () { pick(i); }; }
for (var i = 0; i < n; i++) { bind(i); }
```

드래그·슬라이더는 `clientX` 와 `getBoundingClientRect()` 를 **같이** 쓴다(둘 다 화면 기준이라 확대와 무관):

```js
var r = track.getBoundingClientRect();
var ratio = Math.max(0, Math.min(1, (e.clientX - r.left) / r.width));
```

## 4. CSS 변환표

| 쓰면 안 됨 | 이렇게 |
|---|---|
| `var(--x)` (자동 치환이 안 된 경우) | 값을 직접 쓴다. 아래 "테마 변수" 참고 |
| `display: grid` | `position:absolute` 로 자리 지정, 또는 `inline-block` + 폭, `float`, `display:table` |
| `gap` (flex 포함) | 자식에 `margin-right` / `margin-bottom` |
| `clamp()`, `min()`, `max()` | 캔버스가 고정이므로 값 하나로. 필요하면 `calc()` |
| `inset: 0` | `top:0; right:0; bottom:0; left:0` |
| `position: sticky` | `position: absolute` (캔버스는 스크롤하지 않는다) |
| `object-fit: cover` | `<div style="background:url(...) center / cover no-repeat">` |
| `aspect-ratio` | `width` 와 `height` 를 둘 다 |
| `:is()`, `:where()`, `:has()`, `:focus-visible` | 선택자를 풀어 쓴다(`a:hover, b:hover`), `:focus` |
| `filter: blur()` 등, `backdrop-filter`, `mix-blend-mode`, `clip-path` | 쓰지 않는다. 그림자는 `box-shadow`, 반투명은 `rgba()` |
| `#RRGGBBAA` | `rgba(r, g, b, a)` |
| `user-select: none` | `-ms-user-select: none` 을 함께 |
| `@supports`, `@container`, `@layer` | 쓰지 않는다 |
| WOFF2 글꼴, WebP 이미지 | WOFF, PNG/JPEG |

**테마 변수** (`:root{--bg:#fff}` + `body.dark{--bg:#111}` 처럼 같은 변수를 두 번 정의한 페이지 — 보고에 `CSS variables NOT written out` 으로 나온다):
- 그 페이지가 **한 테마로만** 쓰인다면(예: `<body class="dark">` 고정): 그 테마의 값을 골라 `var(--bg)` 자리를 전부 그 값으로 바꾸고, 변수 정의 두 블록은 지운다. 가장 흔한 경우다.
- 실행 중에 테마가 **바뀐다면**: 기본 테마 값으로 규칙을 쓰고, 다른 테마는 `body.dark .panel { background:#232321 }` 처럼 바뀌는 속성만 규칙으로 덧붙인다.
- 한 번만 정의된 나머지 변수는 자동으로 값이 들어가므로 그대로 둬도 된다.

그대로 되는 것: `border-radius`, `box-shadow`, `linear-gradient`, `transform`(2D/3D), `transition`, `@keyframes`/`animation`, `opacity`, `rgba`, `calc`, `rem`/`vw`/`vh`, `::before/::after`, `nth-child`, 미디어 쿼리, `@font-face`(WOFF/TTF).

**flexbox** 는 되지만 버그가 있다: 자식에 `flex: 1 1 auto` 또는 폭을 명시하고, `gap` 대신 margin, 컨테이너의 `min-height` 대신 `height`, 세로 가운데 정렬이 어긋나면 `position:absolute; top:50%; transform:translateY(-50%)` 로. 반드시 그림으로 확인한다.

## 5. HTML 요소

- `<template>`, `<dialog>`, `<details>`, `<picture>` 는 동작이 없다 → `div` 와 스크립트로.
- `<button>` 은 IE에서 안쪽 내용이 버튼 밖으로 나가면(확대 `transform:scale`, 음수 margin 등) 잘라 버린다. 빌드가 `button{overflow:visible}` 을 기본으로 넣어 주지만, 페이지가 `overflow` 를 따로 정했다면 그 값이 이긴다. 누르는 것은 `div`/`span` + `cursor:pointer` 로 만드는 편이 모양이 일정하다.
- `<input type="range">` 는 끌 때 `input` 이 아니라 `change` 이벤트가 난다 → 둘 다 듣거나, `div` 로 직접 만든 슬라이더를 쓴다(권장: 모양도 일정하다).
- `<input type="date|color|time">` 은 글상자로 보인다. 글상자에 **타이핑하는 것은 지원 범위 밖**으로 본다(확인되지 않음). 버튼·선택 칩으로 설계한다.
- `<video>/<audio>` 는 PC 코덱에 좌우되고 정지 그림에는 안 찍힌다. 꼭 필요하면 발표 PC에서 직접 확인.
- `<iframe>` 으로 인터넷 페이지를 넣을 수는 있지만 발표장 네트워크·보안에 좌우된다.
- 인라인 SVG, `<canvas>` 2D 는 된다. WebGL 은 기대하지 않는다.

## 6. 배치 요령

- 캔버스가 고정이므로 **절대 좌표(px)** 가 가장 예측 가능하다. `assets/page-template.html` 참고.
- 글꼴은 `"Malgun Gothic", "맑은 고딕", sans-serif`. PowerPoint 쪽 글꼴과 맞추면 슬라이드 제목과 자연스럽게 이어진다.
- 배경색을 슬라이드 배경과 같게 하면 경계가 보이지 않는다(영역이 슬라이드 일부일 때 특히).
- 처음 한 번 재생되는 애니메이션은 정지 그림에 중간 상태로 찍힐 수 있다 → job 파일의 `settleMs` 를 늘리거나,
  `window.pptHtmlOnShow` 에서 애니메이션을 시작해 정지 그림은 완성 상태가 되게 한다.
- 스크롤은 피한다(발표 중 휠 조작은 불편하다). 내용이 많으면 탭·단계로 나눈다.

## 7. PowerPoint와의 연결 (pptHtml)

빌드 때 페이지에 `window.pptHtml` 이 추가된다. 원본 HTML을 브라우저에서 그냥 열면 없으므로 `if (window.pptHtml)` 로 감싼다.

| 기능 | 설명 |
|---|---|
| `pptHtml.next()` / `pptHtml.prev()` | 다음/이전으로(PowerPoint의 →/← 와 같다: 애니메이션 단계도 따른다) |
| `pptHtml.gotoSlide(n)` | n번 슬라이드로 |
| `window.pptHtmlOnShow = function () {}` | 이 슬라이드가 화면에 나올 때마다 호출. 애니메이션 다시 시작, 초기화 등에 |
| `window.pptHtmlOnHide = function () {}` | 이 슬라이드가 화면에서 사라질 때 호출. 반복 애니메이션·타이머를 여기서 멈춘다 |
| `pptHtml.host` | `"powerpoint"`(슬라이드 쇼) 또는 `"web"`(웹 버전) |
| `pptHtml.keepKeys = true` | 키보드를 페이지가 계속 갖는다(기본은 클릭 뒤 슬라이드 쇼로 돌려준다). 특별한 경우만 |
| `pptHtml.saveState(글자)` / `pptHtml.loadState()` | 이 페이지의 값을 발표 자료 안에 보관 / 불러올 때 들어 있던 값. PowerPoint에서는 슬라이드 꼬리표에(파일을 저장해야 남는다), 웹 버전에서는 그 브라우저에 남는다. 글자 하나(보통 JSON)이고 40만 자까지. 키트의 `PH.keep` / `PH.kept` 가 이것을 쓴다 |
| `pptHtml.deck` | `{ id, n, total, slides: [{ n, title, thumb }] }`: 이 페이지가 몇 번째인지와 모든 HTML 장의 제목·작은 그림(슬라이드 고르기 화면용) |

- 페이지 상태(선택한 탭 등)는 다른 슬라이드에 다녀와도 유지된다. 쇼를 다시 시작하면 처음 상태가 된다(`saveState` 로 보관한 것만 남는다).
- `contentEditable` 요소에 글자 넣을 곳이 있는 동안은 키보드가 페이지로 온다(Esc 를 누르면 슬라이드 쇼로 돌아간다).
- 마우스 오른쪽 메뉴와 드래그 선택은 꺼진다(발표 중 오동작 방지).
- 키보드는 슬라이드 쇼가 받는다: 페이지를 클릭한 뒤에도 ←/→/Space/Esc 는 PowerPoint로 간다.
- 웹 버전(`.html`)에서도 같은 API가 뷰어를 움직인다.

### 웹답게, 또렷하게, 부드럽게 — 실측한 규칙 (2026-11, 4K 슬라이드 쇼)

HTML을 쓰는 이유는 웹다운 화면이다. 제목+본문+카드로 짠 "슬라이드 모양"이나 시스템 글꼴 그대로면 PPT와 다를 것이 없다(사용자가 두 번 퇴짜 놓았다).

**모양**
- 사이트처럼 짠다: 위쪽 메뉴(표시 + 장 이름 + 다음 단추), 큰 제목, 넉넉한 여백, 실제 제품 같은 부품(입력창, 작업 창, 스위치, 대화 말풍선).
  큰 좌우 화살표나 쪽 번호를 화면 한가운데 두는 슬라이드식 부품은 쓰지 않는다(키트의 도구 막대는 구석에 작고 흐리게 둔 것이다). 장 이동은 메뉴에서 `pptHtml.gotoSlide(n)`.
- **글꼴은 빌드가 내장한다.** 맑은 고딕 그대로면 낡아 보인다. `font-family` 에 글꼴 이름을 적어 두면 빌드가 그 PC에 설치된 글꼴이나
  스킬의 글꼴 묶음(`assets\fonts`: Pretendard, Inter, Roboto, Open Sans, Geist, IBM Plex Sans KR, Noto Sans/Serif KR)에서 쓰인 글자만 잘라 페이지에 넣는다.
  영문 글꼴 뒤에 한글 글꼴을 이어 적으면(`"Inter", "Pretendard", sans-serif`) 둘 다 들어간다. 내장이 허락되지 않은 글꼴은 넣지 않고 보고에 적는다.
- 못 쓰는 것: 글자 윤곽선(`stroke` — 가변 글꼴은 획이 겹친 자리가 드러난다), `background-clip:text`, `filter: blur`, `backdrop-filter`, `mix-blend-mode`.

**선명도**
- 쇼에서는 화면이 통째로 확대된다(4K 는 300%). **DOM 글자와 SVG 는 확대해도 또렷하다 → 그림은 SVG + DOM 으로 그린다.**
- 캔버스는 CSS 크기대로 만들면 3배로 늘어나 뭉개진다(실제로 "화질이 구리다"는 말을 들었다). 꼭 캔버스를 쓸 때는
  `devicePixelRatio` 만큼 키운다(`c.width = w * dpr; ctx.scale(dpr, dpr)`). **속도는 그대로였다**(36 → 37장/초).
  단, 확대는 페이지를 불러온 뒤에 걸리므로 불러올 때 읽은 `devicePixelRatio` 는 1 이다. `pptHtmlOnShow` 나 `resize` 에서 다시 읽는다.
- 캔버스에 그린 글자는 흐리다. 글자는 DOM 으로 얹는다.
- 정지 그림은 job 의 `pictureScale: 3` 으로 4K 급으로 만들 수 있다. 실측 그림을 원본 크기로 보려면 `shotWidth: 8000`.

**속도** (엔진이 GPU 없이 그린다. 아래는 1000×280 영역을 움직일 때 초당 장수)

| 조건 | 장/초 |
|---|---|
| 단색 바탕 위: 캔버스 37 · DOM 점 30개 41 · SVG 선 48개 45 | 36~45 |
| 같은 것을 body 의 `radial-gradient` 바탕 위에서 | 16~19 (절반 이하) |
| body 의 `linear-gradient` 바탕 위에서 | 25 |
| 그라데이션 덩어리가 있어도 움직이는 자리와 겹치지 않으면 | 38~47 (손해 없음) |
| 그림자 달린 600×260 카드를 통째로 옮기기 / 1000×260 면의 투명도 바꾸기 | 28 / 20 |

- **움직이는 것 밑에는 단색만.** 그라데이션은 움직이지 않는 구석에 따로 놓는다(`.bg` 안의 `.blob`). 화면 크기 캔버스를 밑에 까는 것도 같다(절반).
- 계속 도는 것은 작은 자리에서만(햇살 표시 88px, 흐르는 글자 띠 1280×124 → 41~50장). 나머지는 눌렀을 때만 움직이는 전환으로.
- **모든 페이지가 쇼 시작 때 한꺼번에 불려 와 계속 살아 있다.** 반복·타이머는 `pptHtmlOnShow` 에서 시작하고 `pptHtmlOnHide` 에서 모두 멈춘다.
  불려 온 뒤 보이지 않은 페이지도 몇 초 뒤에는 쉬게 한다(정지 그림은 불러온 직후 찍히므로 처음 몇 초는 돌려야 한다).
- 캔버스 좌표와 크기는 정수로(`fillRect(x | 0, y | 0, 2, 2)`). 입자마다 `Math.sin` 을 부르지 않는다.
- 속도는 추측하지 말고 잰다: 빌드의 쇼 시험이 장마다 `fps while appearing` / `fps during the test actions` 와 가장 긴 프레임을 적는다(`references/job-file.md`).

**부드러운 움직임 — 끊긴다는 말을 듣고 다시 잰 것 (v2.1, 페이지 전체 1280×720 을 4K 쇼에서. 가만히 있을 때 약 53)**

| 움직인 것 | 초당 프레임 |
|---|---|
| 큰 카드 여섯 장을 함께 올리거나 흐리게 나타내기 | 14~24, 한 프레임이 100ms 를 넘는다 |
| 페이지 전체의 `opacity` 전환 | 23 |
| 스크립트가 매 프레임 SVG 막대 36개를 고쳐 그리기 | 20 |
| 같은 막대를 HTML 상자 + CSS `transform: scaleY` 전환으로 | 36 |
| 카드는 그대로 두고 그 안의 글자만 올라오기 | 35~47 |
| 숫자 세어 올리기(매 프레임이 아니라 11번에 나눠) | 32~39 |
| 선 그려지기(`stroke-dashoffset`), 도넛 차오르기, 따라다니는 고리 | 43~49 |
| 슬라이더 끌기(숫자만 바로 바꾸고 차트는 손이 멈춘 뒤에) | 57 |

- 엔진은 **움직이는 것의 넓이만큼** 다시 그린다. `left`/`top` 이냐 `transform` 이냐보다 **얼마나 큰 것이 움직이느냐**가 훨씬 크게 작용한다.
- 그래서: 카드·패널·페이지 전체는 나타날 때 움직이지 않는다. 움직이는 것은 그 안의 작은 것(글자 한 줄, 숫자, 막대, 선)이다.
- 값이 빠르게 바뀌는 조작(끌기)에서는 가벼운 것(숫자)만 바로 바꾸고 무거운 것(차트 다시 그리기)은 160ms 쯤 멈춘 뒤에 한다.
- 흐림이 40px 넘는 `box-shadow`, `transition: all`, 큰 요소의 `animation ... infinite` 는 빌드가 `perf-shadow-blur` / `perf-transition-all` / `perf-large-motion` 경고로 알려 준다
  (`perf-large-motion` 은 엔진에서 실제로 그린 뒤, 전환·애니메이션이 걸린 요소 가운데 화면의 8분의 1보다 큰 것을 찾는다).
- 쇼가 시작될 때 미리 불려 오는 페이지는 움직임 없이 마지막 모습으로 그려 둔다(키트는 그렇게 한다: `PH.still()`). 스무 장이 한꺼번에 움직이며 시작하지 않게.
- 느린 PC를 위한 대비: 키트는 장이 처음 나타날 때 프레임을 세어 22 아래면 그 페이지의 움직임을 모두 끈다.

**동작**
- 눌러야만 보이는 내용은 두지 않는다: 들어오면 저절로 한 번 시연하고(자동 재생), 사용자가 누르면 그때부터 사용자를 따른다.
- 마우스를 올리면 열리는 것은 `onclick` 으로도 열리게 한다(누르면 닫히는 토글로 만들면 올리자마자 누를 때 도로 닫힌다).
- `body` 에 `position: relative` 를 주지 않으면 절대 위치 요소가 창 기준이 되어, 창이 캔버스보다 몇 px 넓어져도 바탕 장식(`width:100%`)이 끝까지 간다.

## 8. 검사 규칙 끄기

검사는 패턴으로 찾으므로 틀릴 수 있다(예: jQuery의 `.find()`). 확실히 문제없으면 페이지에 주석을 넣어 그 규칙만 끈다:

```html
<!-- ppt-html-allow: api-fetch css-flex -->
```

규칙 이름은 보고의 `[...]` 안에 나온다. 문법 규칙(`js-arrow` 등)을 끄면 페이지가 통째로 멈추니 끄지 않는다.

## 9. 자주 겪는 문제

| 증상 | 원인과 처리 |
|---|---|
| 보고에 `script error ... (line N)` | `page_NN.prepared.txt` 의 N번째 줄을 본다. 대개 문법(화살표·백틱·const) |
| 그림이 하얗거나 일부만 나옴 | CSS 변수/그리드가 남아 있다. 보고의 `ERROR [css-...]` |
| 요소가 겹치거나 줄이 틀어짐 | flexbox 버그 또는 캔버스보다 큰 설계. 절대 좌표로 바꾸거나 캔버스 크기에 맞춘다 |
| 불러올 땐 정상인데 누르면 반응 없음 | 클릭 함수 안의 문법/함수 문제. `testClicks` 로 눌러서 확인하고 그 함수를 변환표대로 고친다 |
| 한글이 깨짐 | 원본 파일이 UTF-8이 아니다. UTF-8로 저장한다(`<meta charset>` 은 없어도 된다) |
| 글꼴이 다르게 보임 | 웹 폰트(CDN)는 안 온다. 이 PC에 설치된 글꼴 이름을 쓰면 빌드가 페이지 안에 넣어 준다(보고의 `fonts embedded`). `fonts left to the PC` 로 나온 글꼴은 사용 조건상 넣지 못한 것이니 받는 PC에 없으면 바뀐다 |
| 차트·그림이 2배로 크거나 한쪽에 몰려 나옴 | 스크립트가 **불러올 때** 잰 크기로 그렸다. PowerPoint는 불러온 뒤에 배율을 맞추므로 `window.pptHtmlOnShow` 와 `resize` 에서 다시 재서 그린다 |
| 속성을 바꿨는데 색·모양이 그대로 | 스크립트로 CSS 변수를 바꾸고 있다(효과 없음). 클래스나 `data-` 속성을 바꾸고 CSS 에서 그 선택자로 값을 준다 |
| 변환 비교 그림에서 격자가 다르게 나옴 | 복잡한 grid 다. `check-html.ps1 -Probe` 로 값을 확인하거나, 그 부분만 flex + 여백(또는 절대 좌표)으로 고쳐 쓴다 |
| 자동 변환 뒤 스크립트 오류 | 변환할 수 없는 문법이거나 엔진에 없는 기능이다. 보고의 오류 줄을 `page_NN.prepared.txt` 에서 찾고, 그 기능을 다른 방법으로 바꾼다 |
