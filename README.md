# PPHTML-BD — PowerPoint 안에서 움직이는 HTML

**지금은 Windows 전용입니다**(Windows 10/11 + PowerPoint 데스크톱 앱). Mac 지원은 추가할 예정입니다.
만들어진 웹 버전 `.html` 은 Mac·휴대폰 등 어디서나 브라우저로 볼 수 있습니다.

**PowerPoint 슬라이드 위에서, 요즘 웹 페이지가 그대로 돌아가게 만드는 오픈소스입니다.**
스크롤, 애니메이션, 유리 효과, 3D, 캔버스, 입력 양식, 화면 전환 — 브라우저에서 되는 것이 슬라이드 쇼 안에서도 됩니다.
AI(Claude 등)에게 "이 내용을 눌러 볼 수 있는 발표 자료로 만들어 줘"라고 말하면, 예쁘고 살아 있는 `.pptm` 이 나옵니다.

> 누구나 PowerPoint 를 AI 로 쉽게, 예쁘게 쓸 수 있도록. 함께 키워 가는 프로젝트입니다.

## 기능

### 1. 내 내용으로 눌러 보는 PPT 만들기
- 글·숫자·표를 웹 페이지로 만들어, 슬라이드 쇼 안에서 그 자리에서 움직이는 `.pptm` 을 만듭니다(`new-deck.ps1` → `html2ppt.ps1`).
- 브라우저에서 되는 것은 거의 다 됩니다: 스크롤, 애니메이션, 유리 효과, 3D, 캔버스·WebGL, 입력창, 화면 전환, 외부 라이브러리.
- 서버·팝업·별도 창 없이 파일 하나. 글꼴도 파일 안에 넣어서 받는 PC 에 글꼴이 없어도 모양이 같습니다.
- 출발점으로 쓸 예시 페이지 23종(표지, KPI, 가격표, 타임라인, 투표 등)과 순서 묶음 5종(투자 피칭, 분기 보고, 제품 발표, 업무 보고, 회의)이 들어 있습니다. 모습은 [assets/templates/gallery.png](assets/templates/gallery.png).

### 2. 디자인 따라 하기
- 웹 주소·HTML·PPT·그림을 주면 색·글꼴·모양을 읽어 테마와 `design.md` 로 만듭니다(`get-style.ps1`).
- `design.md` 같은 디자인 규칙 문서대로 만들거나, 값 몇 개로 테마를 새로 만듭니다(`new-theme.ps1`).
- 바로 쓸 수 있는 테마 17종(cobalt, paper, ink, midnight, seoul 등). 모습은 [assets/kit/themes.png](assets/kit/themes.png).

### 3. 이미 있는 HTML 올리기
- 완성된 HTML 을 슬라이드 한 장씩 쓴 그대로 올립니다. React·Vue 빌드 결과물도 됩니다.
- 기존 PPT 의 맨 뒤나 정해진 자리("3번 슬라이드 아래쪽 절반" 등)에 넣을 수 있습니다. 원본 PPT 는 바뀌지 않습니다.

### 4. PPT 를 웹 페이지로 바꾸기
- `.pptx` 를 어디서나 열리는 `.html` 한 파일로 바꿉니다(`ppt2html.ps1`).
- ←/→·클릭·스와이프로 넘기기, 목록(G), 발표자 메모(N), 전체 화면(F), 인쇄.

### 5. 발표 중 도구
- 도구 막대는 화면 오른쪽 아래 구석에 마우스를 가져가면 나타납니다.
- 쪽 번호를 길게 누르면 모든 장의 작은 그림이 떠서 원하는 장으로 건너뜁니다.
- 연필로 글자를 고치고 옮깁니다. 쇼가 끝나면 슬라이드 그림에도 반영됩니다.
- 펜으로 그리기, 눌러서 확대하기. 투표처럼 고른 값은 파일 안에 남습니다.

### 6. 자동 확인
- 만든 뒤 화면에 보이지 않는 PowerPoint 에서 진짜 슬라이드 쇼를 돌려, 장마다 오류 없이 움직이는지 확인하고 그림과 문제 목록을 보여 줍니다.
- 페이지 하나만 몇 초 만에 확인하는 기능도 있습니다(`check-html.ps1`, PowerPoint 불필요).

### 7. 공유와 점검
- `.pptm` + 웹 버전 + "먼저 읽어주세요" 안내를 공유 폴더와 zip 으로 묶습니다.
- 이 PC 에서 쓸 수 있는지 점검합니다(`doctor.ps1`). 스킬을 zip 으로 묶어 다른 사람에게 줄 수 있습니다(`pack.ps1`).

## 어떻게 동작하나요

- 슬라이드 쇼가 시작되면 파일 안의 작은 매크로가 **Microsoft Edge 의 엔진(WebView2)** 을 슬라이드 쇼 창 안, HTML 자리 위에 정확히 띄웁니다. 페이지는 파일 안에 들어 있습니다.
- WebView2 가 없는 PC 에서는 Internet Explorer 11 엔진과 자동 변환한 사본으로 내려가고, 그것도 안 되면 같은 자리의 정지 그림이 보입니다. 파일이 깨지지 않습니다.
- 매크로를 쓸 수 없는 곳(Mac, 웹 PowerPoint, 매크로 금지 회사)에는 함께 만들어지는 **웹 버전 `.html`** 을 주면 됩니다.

자세한 원리는 [references/how-it-works.md](references/how-it-works.md).

## 필요한 것

- Windows 10 / 11 PC. Mac 은 아직 안 됩니다.
- PowerPoint 데스크톱 앱(Microsoft 365 또는 Office 2013 이상). 웹 PowerPoint 로는 만들 수 없습니다.
- Claude Code(Claude 데스크톱 앱의 Code 탭, 또는 터미널의 `claude`). claude.ai 채팅창에 zip 을 올리는 방식으로는 동작하지 않습니다. 채팅창 쪽 환경에는 Windows 와 PowerPoint 가 없기 때문입니다.
- Microsoft Edge 는 Windows 에 기본으로 들어 있습니다. 그 밖에 따로 설치할 프로그램은 없고, 만들 때 인터넷도 필요 없습니다.

## 설치

세 방법 중 하나만 하면 됩니다.

### 방법 1. Claude Code 에게 맡기기 (가장 쉬움)

Claude Code 를 열고 이렇게 말합니다.

```text
https://github.com/Krocean/PPHTML-BD 이 스킬 설치해 줘
```

### 방법 2. 명령 한 줄

PowerShell 에 붙여 넣습니다(git 이 있어야 합니다).

```powershell
git clone https://github.com/Krocean/PPHTML-BD.git "$env:USERPROFILE\.claude\skills\ppt-html"
```

### 방법 3. 직접 내려받기 (명령 없이)

1. 이 페이지 위쪽의 초록색 **Code** 단추 → **Download ZIP** 을 누릅니다.
2. 받은 zip 파일을 마우스 오른쪽 단추로 눌러 **속성**을 열고, 아래쪽 '보안'의 **차단 해제**에 체크한 뒤 확인을 누릅니다(항목이 없으면 넘어갑니다).
3. 압축을 풀면 `PPHTML-BD-main` 폴더가 나옵니다. 이 폴더 이름을 `ppt-html` 로 바꿉니다.
4. 파일 탐색기 주소창에 `%USERPROFILE%\.claude\skills` 를 넣고 Enter 를 누릅니다. 없는 폴더라고 나오면 `%USERPROFILE%\.claude` 로 가서 `skills` 폴더를 새로 만듭니다.
5. `ppt-html` 폴더를 그 안으로 옮깁니다. 아래처럼 되면 맞습니다.

```text
C:\Users\(내 이름)\.claude\skills\ppt-html\SKILL.md
```

`ppt-html\PPHTML-BD-main\SKILL.md` 처럼 폴더가 한 겹 더 들어가 있으면 Claude 가 스킬을 찾지 못합니다. 압축을 풀 때 폴더가 두 번 생기는 경우가 많으니 확인해 주세요.

### 설치 확인

Claude Code 를 다시 시작하고 이렇게 말합니다.

```text
ppt-html 스킬로 이 PC 에서 쓸 수 있는지 점검해 줘
```

OK / NOTE / PROBLEM 목록이 나옵니다. PROBLEM 이 있으면 그 내용을 [이슈](https://github.com/Krocean/PPHTML-BD/issues)로 알려 주세요.

### 써 보기

- "이 내용을 PPT 안에서 눌러 볼 수 있는 화면으로 만들어 줘" 라고 말하고, 그 아래에 내용(글·숫자·표)을 붙여 넣습니다.
- "이 사이트(또는 이 PPT) 디자인으로 내 내용을 슬라이드에 올려 줘"
- "이 PPT 를 어디서나 열리는 웹 페이지로 바꿔 줘"

AI 없이 직접 쓸 수도 있습니다.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\doctor.ps1 -Full
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\new-deck.ps1 -To C:\work\mydeck -Blank 3 -Name "내 발표"
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\html2ppt.ps1 -Job C:\work\mydeck\job.json
```

사용법 전체는 [SKILL.md](SKILL.md), 페이지 쓰는 법은 [references/modern-web.md](references/modern-web.md).

### 처음 쓸 때 Star 를 한 번 여쭤봅니다

처음으로 무언가를 만들어 달라고 하면, AI 가 시작하기 전에 이 저장소에 Star(★)를 눌러도 되는지 **한 번** 묻습니다.

- "예"라고 하면: [GitHub CLI](https://cli.github.com/)(`gh`)로 로그인돼 있으면 그 계정으로 바로 누르고, 아니면 저장소 페이지를 브라우저로 열어 드립니다(직접 누릅니다).
- "아니요"라고 하면: 그냥 넘어가고 다시 묻지 않습니다. **기능은 똑같습니다.**
- 대답은 이 PC 의 `%LOCALAPPDATA%\ppt-html\star.json` 에 날짜와 함께만 남습니다(어디로도 보내지 않습니다). 지우면 다시 묻습니다.
- 허락 없이 누르는 일은 없습니다. 하는 일은 [scripts/star.ps1](scripts/star.ps1) 에서 그대로 볼 수 있습니다.

## 만든 파일을 열 때

- 처음 열면 노란 줄의 **[콘텐츠 사용]** 을 눌러야 움직입니다(매크로). **슬라이드 쇼(F5)** 에서만 움직이고 편집 화면에서는 정지 그림입니다.
- 메일로 받은 파일은 Windows 가 매크로를 막습니다. 파일 속성에서 **[차단 해제]** 후 엽니다.

## 자주 막히는 곳

| 상황 | 이렇게 해 보세요 |
|---|---|
| Mac 이에요 | 아직 안 됩니다(추가 예정). Windows PC 에서 만든 웹 버전 `.html` 은 Mac 에서도 열립니다. |
| claude.ai 채팅창에 올렸는데 안 돼요 | Claude Code 에서만 동작합니다. 위 "필요한 것"을 봐 주세요. |
| Claude 가 스킬을 모르는 것 같아요 | `...\.claude\skills\ppt-html\SKILL.md` 위치가 맞는지 보고, Claude Code 를 다시 시작합니다. |
| HTML 이 안 움직이고 그림만 보여요 | 슬라이드 쇼(F5)인지, 노란 줄의 [콘텐츠 사용]을 눌렀는지 봅니다. 메일·메신저로 받은 파일이면 속성에서 [차단 해제]를 합니다. |
| 회사 PC 라 매크로가 막혀 있어요 | 회사 정책은 풀 수 없습니다. 함께 만들어지는 웹 버전 `.html` 을 쓰세요. |
| 그래도 안 돼요 | [이슈](https://github.com/Krocean/PPHTML-BD/issues) → "PC 실행 결과"에 점검 결과를 붙여 주세요. Claude 에게 "ppt-html 스킬로 전체 점검해 줘"라고 하면 나옵니다. 사용자 이름·경로는 지워 주세요. |

## 지금 상태 (정직하게)

- 개발한 PC 한 대(Windows 11, PowerPoint LTSC 2024 64비트, Edge 157)의 숨은 슬라이드 쇼에서 확인했습니다. 다른 PC·다른 Office 버전은 아직 확인하지 못했습니다. `doctor.ps1 -Full` 이 그 PC 에서 직접 시험합니다.
- 실제 모니터에서 사람이 눌러 보는 것, 한글 입력기 조합, 터치, 소리·동영상 재생은 아직 확인 못 했습니다.
- 확인 범위와 한계는 [SKILL.md](SKILL.md) 의 "한계" 와 [references/how-it-works.md](references/how-it-works.md) 끝에 자세히 적어 두었습니다.

다른 PC 에서 돌려 본 결과(되든 안 되든)를 이슈로 알려 주시는 것이 가장 큰 도움입니다.

## 함께 만들기

누구나 코드를 보고, 고쳐서 제안(Pull Request)할 수 있습니다. 제안은 관리자가 확인하고 수락해야 반영됩니다.
[CONTRIBUTING.md](CONTRIBUTING.md) 를 먼저 읽어 주세요.

## 라이선스

[MIT](LICENSE). 함께 들어 있는 글꼴과 빌드 도구는 각자의 공개 라이선스를 따릅니다: [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).

---

## English (short)

**Windows only for now** (Windows 10/11 + PowerPoint desktop). macOS support is planned. The web version (`.html`) it builds opens anywhere, including Mac and phones.

**PPHTML-BD** puts live, modern web pages inside PowerPoint slides. During a slide show, a small macro in the `.pptm` hosts Microsoft Edge WebView2 exactly over the slide area, so anything a current browser can do works on the slide. No server, no add-in, no install — one file. Falls back to an IE11 copy, then to a still picture, where WebView2 or macros are unavailable.

### Features

1. **Interactive decks from your content** — turns your text, numbers and tables into web pages that run in place during the slide show (`new-deck.ps1` → `html2ppt.ps1`). Scrolling, animation, glass effects, 3D, canvas/WebGL, forms, View Transitions and libraries all work. Fonts are embedded. 23 example pages (cover, KPI, pricing, timeline, poll …) and 5 ready-made sequences (pitch, quarterly review, keynote, report, workshop) to start from.
2. **Design from a reference** — reads colours, fonts and shapes from a web address, HTML, PPT or picture into a theme and a `design.md` (`get-style.ps1`), follows a design rules document, or builds a theme from a few values (`new-theme.ps1`). 17 themes included.
3. **Existing HTML onto slides** — places finished HTML (React/Vue builds included) as is, one page per slide, at the end of an existing deck or in a given area of a given slide. The original deck is never changed.
4. **PPT to a web page** — converts a `.pptx` into one `.html` file that opens anywhere (`ppt2html.ps1`): arrow keys, click or swipe, overview (G), speaker notes (N), full screen (F), print.
5. **Presenter tools** — a toolbar that appears in the bottom-right corner: long-press the page number to jump to any slide, edit text during the show (carried into the slide picture afterwards), pen, click to zoom. Values such as poll results are kept in the file.
6. **Automatic checking** — after building, runs a real slide show in a hidden PowerPoint instance, confirms each slide is live without script errors, and reports pictures and problems. `check-html.ps1` checks a single page in seconds without PowerPoint.
7. **Sharing and diagnostics** — packs the `.pptm`, the web version and a read-me into a share folder and zip; `doctor.ps1` checks whether a PC can run it; `pack.ps1` zips the skill for someone else.

On first use, the assistant asks once whether it may star this repository (`scripts/star.ps1`). It stars only after a clear yes, and everything works the same if you say no.

It is packaged as a Claude Code skill (`SKILL.md`) but the PowerShell scripts work on their own. Docs are in Korean for now; translations are welcome. MIT licensed.
