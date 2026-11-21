# 동작 원리 · 파일 구성 · 유지보수

## 1. 슬라이드 위에서 HTML이 도는 원리

**엔진이 둘이다(v3).** 쇼가 시작되면 호스트 코드가 먼저 **최신 엔진**(Edge WebView2: 페이지를 쓴 그대로 돌린다)을 띄워 보고, 그 PC 에 없거나 뜨지 않으면 **기존 엔진**(Internet Explorer 11 브라우저 컨트롤: 빌드가 변환해 둔 사본)으로 내려간다.
둘 다 "페이지를 슬라이드 쇼 창의 자식 창에 띄워 `htmlArea` 그림 자리에 정확히 얹는다"는 점은 같다. 아래 1~4번은 공통의 뼈대와 기존 엔진의 방식이고, 최신 엔진의 방식은 이 장 끝의 "최신 엔진"에 있다.

PowerPoint는 슬라이드에 웹 브라우저 컨트롤(ActiveX)을 직접 올리는 것을 막아 두었다(Office 2024 LTSC 기준: `AddOLEObject("Shell.Explorer.2")` 실패, XML에 직접 써넣어도 개체를 만들지 않음, 웹 브라우저 CLSID에 Office 킬 비트). 추가 기능(웹 add-in)은 서버·등록이 필요하다. 그래서 다음 방식을 쓴다.

1. **파일 안에 들어가는 것**
   - 슬라이드의 HTML 자리에는 그 페이지의 **정지 그림**(이름 `htmlArea`)이 놓인다. 편집 화면, 매크로가 꺼진 PC, Mac·웹에서는 이 그림이 보인다.
   - 슬라이드에는 태그 `HtmlPage=N`, 프레젠테이션에는 태그 `InlineHtml=1`.
   - 페이지(준비된 HTML 전체)는 **사용자 지정 XML 파트**(네임스페이스 `urn:ppt-html:inline:v1`)에 들어간다: `<page><bg>배경색</bg><w>캔버스 폭</w><h>높이</h><settle>ms</settle><html>기존 엔진용 문서</html><htmlm>최신 엔진용 문서</htmlm></page>`. 둘 중 하나는 비어 있을 수 있다.
   - VBA 프로젝트(모든 덱에 같은 코드) + 파일을 열 때 코드를 시작시키는 리본 파트(`customUI14.xml` 의 `onLoad`).
2. **슬라이드 쇼가 시작되면** VBA가
   - XML 파트에서 페이지를 읽어 `%TEMP%\ppthtml_<프로세스>_N.html` 로 쓰고(UTF-8, CR LF, 인터넷 영역 표식 포함),
   - 테두리·제목 줄 없는 폼에 페이지마다 브라우저 컨트롤을 하나씩 만들어 불러오고,
   - 그 폼을 **슬라이드 쇼 창의 자식 창**으로 넣어 `htmlArea` 도형 자리에 정확히 맞춘다. 별도 창이 아니라 슬라이드 표면의 일부가 된다.
   - 브라우저 확대 기능(OLECMDID_OPTICAL_ZOOM)으로 페이지의 창 폭을 캔버스 폭에 맞춘다(화면 배율·해상도와 무관하게).
3. **쇼 동안** 슬라이드가 바뀌면 해당 페이지를 보이고 나머지는 숨긴다(상태 유지). 페이지는 숨은 요소 `#ihCmd` 를 "클릭"하는 방식으로 VBA에 말을 건다(`tick`/`focus`/`key..`/`next`/`prev`/`goto..`). 0.4초마다의 `tick` 으로 위치·슬라이드·종료 상태를 스스로 맞춘다.
4. **쇼가 끝나면** 폼을 떼어 내고 임시 파일을 지운다.

시작 경로는 셋이다(앞의 것이 안 되면 뒤의 것): 리본 `onLoad` → 응용 프로그램 이벤트, PowerPoint 자동 매크로 `OnSlideShowPageChange`, 정지 그림의 실행 설정(클릭·마우스 오버 → `InlineHtml_Activate`).
쇼 창에 넣는 방법도 셋이다: 직접 자식 창 → 혼합 DPI 컨테이너를 거친 자식 창(고해상도 PC에서 보통 이것, 로그의 `mode=2`) → 쇼 창이 소유한 테두리 없는 창(`mode=3`).

진단 매크로(Application.Run 으로 호출, 문자열 반환): `InlineHtml_Status`(상태와 로그), `InlineHtml_Probe n`(n번 페이지의 데이터·임시 파일·브라우저 상태),
`InlineHtml_Eval n, "식"`(n번 페이지 안에서 계산한 값. 쇼 시험의 값 읽기와 프레임 재기가 이것을 쓴다).

**페이지가 보관하는 값** (v2.1). 페이지가 `pptHtml.saveState("글자")` 를 부르면 그 글자를 URI 인코딩해 `#ihCmd` 의 `data-arg` 에 싣고 `save` 명령을 보낸다.
VBA는 그것을 그 슬라이드의 꼬리표 `HtmlState0`, `HtmlState1`, … (4,000자씩)에 적고 프레젠테이션을 "바뀜"으로 표시한다(그래서 닫을 때 저장을 묻는다).
쇼를 시작할 때 VBA는 임시 페이지 파일을 쓰면서 머리 끝(`<!--/ih:head-->`) 뒤에 `<script>window.__ihState="…";</script>` 를 넣고, `pptHtml.loadState()` 가 그것을 푼다.
웹 버전에서는 같은 함수가 `localStorage` 의 `ppthtml:<덱 id>:<장>` 을 쓴다. 따옴표·꺾쇠·줄바꿈이 든 값이나 40만 자를 넘는 값은 VBA가 받지 않는다(인코딩된 글자만 온다).

**고친 내용이 쇼 밖에서도 보이게** (v2.2). 쇼 동안 `save` 를 보낸 페이지는 표시해 둔다. 쇼가 끝나면(`InlineShowEnded`) 호스트 폼을 떼어 낸 뒤, 표시된 페이지마다
보관한 내용을 넣은 문서를 **새 이름의 임시 파일**(`ppthtml_<프로세스>_<N>s.html`)로 써서 그 페이지의 브라우저에 다시 불러오고, 캔버스 배율을 맞추고, 빌드 때의 `settleMs`(페이지 데이터의 `<settle>`)만큼 기다린 뒤
문서를 `OleDraw` 로 메모리 그림에 그려 GDI+ 로 PNG 를 만들고, 슬라이드의 `htmlArea` 그림을 같은 자리·크기·이름·겹침 순서·실행 설정의 새 그림으로 바꾼다(`RefreshStills`, `frmInline.Snap`, `ReplaceStill`).
그림의 크기는 그 화면에서의 페이지 크기다(4K 전체 화면이면 3840×2160). 그 뒤에 폼을 내리고 임시 파일을 지운다. 이 동안 다른 진입점은 `mBusy` 로 기다린다.
여기서 걸렸던 것 세 가지: ① 컨트롤이 이미 보여 주고 있는 파일 이름으로 다시 `Navigate` 하면 캐시된 옛 문서가 나온다 → 다른 이름의 파일로 부른다.
② 쇼가 끝나는 동안 `SlideShowWindows` 에 창이 남아 있어도 그 `Presentation` 은 읽히지 않는다. `On Error Resume Next` 아래에서 실패한 `If` 조건은 블록 **안으로** 들어가므로 `FindShow` 가 죽은 창을 돌려줬다 →
값을 먼저 변수에 읽고 판단하며, 문서를 만들 때는 호스트가 기억해 둔 덱(`mShowPres`)을 먼저 쓴다. ③ MSForms 가 감싼 브라우저 컨트롤(`.Object`)은 `OleDraw` 를 받지 않는다(DV_E_NOIVIEWOBJECT) → 문서 개체(`.Document`)에 그리게 한다.
웹 쪽: `deck_extract.ps1` 이 슬라이드 꼬리표의 보관 내용을 읽고 `webdeck.ps1` 이 해당 페이지에 `window.__ihState` 로 넣는다. 웹용 연결 코드의 `loadState` 는 브라우저에 저장된 것이 없을 때 그것을 쓴다.

**최신 엔진** (v3, `vba\modWv2.bas` + `modInline.bas` 의 "modern engine" 부분).
- VBA 가 Office 에 들어 있는 `WebView2Loader.dll`(POWERPNT.EXE 옆)을 불러 `CreateCoreWebView2EnvironmentWithOptions` 를 부른다. COM 호출은 메서드 표의 칸 번호로 `DispCallFunc` 를 거쳐 하고(참조·등록 불필요), 완료 처리기는 VBA 로 만든 작은 COM 개체(메서드 표 + `AddressOf`)다. 칸 번호는 Microsoft 의 interop 어셈블리에서 읽은 값으로, 공개된 인터페이스에서는 바뀌지 않는다.
- 브라우저 프로필은 `%LOCALAPPDATA%\ppt-html\wv2`(빌드용 숨은 PowerPoint 는 `PPTHTML_WV2_DATA` 로 따로). Office 자신의 WebView2 와 같은 프로필을 쓰면 창 만들기가 `8007139F` 로 실패한다. 같은 이유로 보이지 않는 그림용 창도 쇼 창과 같은 DPI 문맥에서 만든다.
- 페이지마다 WebView 하나, 모두 쇼 창 안의 컨테이너 창(`Static`, 창 글자 `InlineHtmlHost_7C41`)의 자식이다. 페이지 파일은 `%TEMP%\ppthtml_<프로세스>_m\` 에 쓰고 `SetVirtualHostNameToFolderMapping` 으로 `https://ppt-html.example/…` 주소로 연다(보안 컨텍스트). 쇼가 끝나면 지운다.
- 창 폭은 확대 비율(`put_ZoomFactor`)로 캔버스 폭에 맞추고, 페이지가 알려 오는 실제 폭(`vp`)으로 한 번 더 고친다.
- 페이지 → 호스트: `chrome.webview.postMessage("명령\n인자")`(`ready`/`vp`/`err`/`tick`/`key..`/`next`/`prev`/`goto..`/`save`). WebView2 의 처리기 안에서는 메시지 루프를 돌릴 수 없으므로 처리기는 줄에 넣기만 하고, `SetTimer` 콜백이 호스트에 넘긴다. 호스트 → 페이지: `PostWebMessageAsString("show"|"hide")` → 페이지의 `pptshow`/`ppthide` 이벤트.
- 다음 장의 페이지는 미리 불러 두고, 처음 화면에 나올 때까지 CSS 애니메이션을 멈춰 둔다(연결 코드 `harness-head-wv.txt`). 동시에 10장까지, 넘으면 오래 안 본 장부터 내린다.
- 쇼 뒤 그림 갱신: 보이지 않는 창에 그 페이지를 보관값과 함께 다시 불러 `CapturePreview`(PNG)로 찍고 `ReplaceStill`.
- 쇼 시험의 입력은 `CallDevToolsProtocolMethod`(`Input.dispatchMouseEvent` 등)로 넣는다(`InlineHtml_Cdp`). 진단: `InlineHtml_Engine`, 호스트 기록 파일(`PPTHTML_LOG`, 빌드 때는 `%LOCALAPPDATA%\ppt-html\hd\host.log`).
- 엔진 고르기: 프레젠테이션 꼬리표 `HtmlEngine=classic` 이면 기존 엔진. 쇼 시작 전에 매크로를 부르면(PowerPoint 가 코드를 읽는 중) 그 뒤의 모든 `Application.Run` 이 실패하는 것을 봤다. 그래서 시험 도구는 매크로가 아니라 꼬리표로 엔진을 고른다.

**덱 정보** (v2.1). 빌드가 각 페이지의 머리에 `window.__ihDeck = { id, n, total, slides: [{ n, title, thumb }] }` 를 넣는다(`html2ppt.ps1`).
`title` 은 페이지의 `<title>`, `thumb` 은 정지 그림을 줄인 JPEG(data URI, 장이 15개를 넘으면 더 작게)다. 슬라이드 고르기 화면이 이것으로 그려진다.

## 2. 빌드 파이프라인 (html2ppt.ps1)

| 단계 | 하는 일 | 어디서 |
|---|---|---|
| 덱 읽기 | 원본을 읽기 전용으로 열어 작업 사본 저장, 슬라이드 크기·표식 수집 | 숨은 PowerPoint (`lib\deck_stage.ps1`) |
| 최신 엔진용 준비 | 자료 합치기, 글꼴 내장, 연결 코드(`harness-*-wv.txt`)만 넣고 그대로. Edge 로 그려 정지 그림을 얻고 스크립트 오류·스크롤 길이를 잰다 | `lib\modern.ps1`, `HtmlPrep.FinishModern` |
| 기존 엔진용 준비 (대비용) | 자료 합치기(큰 그림은 줄이고 사진은 JPEG로, 누운 사진은 세워서), (필요한 페이지만) 최신 CSS·JavaScript 자동 변환, 글꼴 내장, CSS 변수 치환, 보완 함수·연결 코드 삽입, 규칙 검사(성능 규칙 포함) | `cs\HtmlPrep.cs`(`Slim` 이 그림 담당) + Edge 헤드리스(`lib\edge.ps1`, `assets\tools\*.js`) |
| 렌더 | IE11 엔진으로 불러 그림 저장, 스크립트 오류·배경색·넘침·크게 움직이는 요소 수집. 변환한 페이지는 최신 브라우저 그림과 비교 | `lib\render.ps1` (PowerPoint 불필요), `cs\ImgKit.cs` |
| 덱 정보 | 장마다 제목과 작은 그림을 모아 모든 페이지에 넣음 | `html2ppt.ps1` |
| 배치 | 정지 그림·태그·XML 파트 넣고 `.pptm` 으로 저장 | 숨은 PowerPoint (`lib\deck_apply.ps1`) |
| 코드 삽입 | VBA 프로젝트를 텍스트 원본에서 만들어 패키지에 넣음 + 리본 파트 | `cs\VbaKit.cs`, `Add-VbaToPackage` |
| 마무리 | 정지 그림에 실행 설정 연결, PowerPoint가 직접 다시 저장 | `lib\deck_finalize.ps1` |
| 실측 | 숨은 화면에서 실제 슬라이드 쇼, 페이지별 라이브 확인·캡처, 나타날 때와 시험 동작 중의 초당 프레임(페이지의 `__ihFps` 가 `requestAnimationFrame` 간격을 센다) | `lib\show_test.ps1` |
| 전달 | 결과 복사(PowerShell이 못 쓰는 폴더면 PowerPoint가 사본 저장), 웹 버전·공유 묶음, 웹 버전을 Edge로 열어 확인 | `lib\webdeck.ps1`, `lib\webcheck.ps1` |

### 자동 변환과 Edge

Edge(없으면 Chrome)를 창 없이 띄워 DevTools 프로토콜(이 PC 안의 임시 포트)로 조종한다(`cs\EdgeKit.cs`). 사용자의 Edge 프로필은 쓰지 않는다(작업 폴더에 임시 프로필을 만들고 끝나면 지운다).
그 안에서 도는 도구: `assets\tools\convert.js`(PostCSS·Autoprefixer·acorn 과 이 스킬의 변환 규칙), `babel.min.js`(JavaScript 문법 변환), `fonttool.js`(HarfBuzz 로 글꼴 자르기 + WOFF 포장), `tool.js`(페이지를 캔버스 크기로 띄워 그림과 글꼴 사용 내역을 얻는다).
페이지에 들어가는 것: 변환한 CSS·JS, 엔진이 표현하지 못하는 배치 정보의 표(`window.__ihAdapt`)와 그것을 엔진의 옛 grid·여백으로 옮기는 `assets\adapt-ie11.txt`, 필요할 때 `assets\corejs.min.txt`.
원래의 최신 CSS는 지우지 않고 엔진용 값을 앞에 덧붙이므로(`display: -ms-grid; display: grid;`) 같은 페이지가 웹 버전의 최신 브라우저에서는 원래대로 동작한다. 어댑터는 엔진 안에서만 움직인다.
Edge를 쓸 수 없으면 이 단계들이 통째로 빠지고 보고에 이유가 적힌다(빌드는 계속된다).

VBA 개체 모델 접근은 보통 막혀 있으므로(보안 설정을 바꾸지 않는다) `vbaProject.bin` 을 파일 형식대로 직접 쓴다. 원본만 든 프로젝트이고 Office가 열 때 컴파일한다.

## 3. 숨은 PowerPoint

- 별도의 Windows 데스크톱(`PptHtmlDesk`)에서 PowerPoint를 하나 더 띄운다(`POWERPNT.EXE /B`). 화면에 아무것도 뜨지 않고, 사용자가 열어 둔 PowerPoint와 프로세스가 다르다.
- 그 인스턴스의 문서 창을 통해서만 개체에 접근한다(다른 인스턴스에 붙을 수 없다). 빈 프레젠테이션 하나가 기준점으로 열려 있다.
- 자식 스크립트가 도는 동안 그 인스턴스에 뜨는 대화 상자(VBA 오류 등)는 기록·캡처하고 닫는다(자동화가 멈추지 않게).
- 실행이 끝나면 항상 종료한다. 중단됐을 때는 `scripts\hd.ps1 stop`.
- 작업 폴더: `%LOCALAPPDATA%\ppt-html\work\<종류-시각>` (3일 지나면 자동 삭제). 실행 기록은 그 안의 `run\`.

## 4. 폴더 구성

```
ppt-html/
  SKILL.md   VERSION
  scripts/
    html2ppt.ps1  ppt2html.ps1  check-html.ps1  new-deck.ps1  get-style.ps1  new-theme.ps1  doctor.ps1  pack.ps1  hd.ps1      ← 실행하는 것
    lib/     공통 모듈과 단계별 자식 스크립트 (pages 준비, edge 도구, fonts 글꼴, render 렌더, webdeck/webcheck 웹 버전, deck_* PowerPoint 작업, style 테마 계산, deck_style PPT 디자인 읽기)
    cs/      DeskKit(숨은 데스크톱·창)  HtmlPrep(HTML 준비·검사)  VbaKit(VBA 프로젝트 쓰기)  EdgeKit(Edge 조종)  FontKit(글꼴 파일 읽기)  ImgKit(그림 비교·모아 보기)
    vba/     modInline.bas  modWv2.bas  frmInline.frm  clsAppEvents.cls  clsDomBridge.cls   ← 덱에 들어가는 코드의 원본 (modWv2 = 최신 엔진)
  assets/
    kit/                    디자인 키트: theme-*.css 17종(+ themes.json, themes.png), kit.css, kit.js, kit-deck.js(덱 도구)
    templates/              키트로 만든 예시 페이지 23종 + index.json(템플릿·테마·묶음 목록) + gallery.png(모습)
    page-blank.html         빈 페이지의 출발점(new-deck.ps1 -Blank 가 쓴다)
    specimen.html           테마 미리보기에 쓰는 견본 화면   design-template.md  get-style 이 채우는 design.md 의 틀
    tools/                  Edge 안에서 도는 빌드 도구(변환·글꼴)와 style.js(레퍼런스 페이지에서 디자인 값을 재는 것). 발표 파일에는 들어가지 않는다
    fonts/                  줄인 글꼴 묶음(글꼴이 없는 PC용): Noto Sans/Serif KR, Pretendard, Inter, Roboto, Open Sans, Geist, IBM Plex Sans KR + 사용 조건
    adapt-ie11.txt          변환한 페이지에 들어가는 레이아웃 어댑터
    corejs.min.txt          변환한 스크립트가 쓰는 표준 기능 모음(core-js)
    page-template.html      키트 없이 시작하는 빈 틀
    viewer-template.txt     웹 버전 뷰어
    harness-*.txt           페이지에 넣는 연결 코드(PowerPoint용 / 웹용)
    polyfills-ie11.txt      IE11 보완 함수
    strings.json            한글 문구(대체 텍스트, 공유 안내문, 설치 안내)
  dev/                      도구 묶음(assets/tools, assets/fonts)을 다시 만드는 원본과 스크립트. 쓰는 데는 필요 없다
  references/
```

실행에 필요한 것은 Windows에 이미 있는 것뿐이다(PowerShell 5.1, PowerPoint, Edge). C# 원본은 실행할 때 PowerShell이 컴파일한다.
`.ps1`/`.cs`/VBA 원본은 **ASCII만** 쓴다: Windows PowerShell 5.1은 BOM 없는 스크립트를 PC의 옛 코드 페이지로 읽기 때문이다. 한글 문구는 `assets\strings.json`(UTF-8)에 둔다.
`assets\tools` 와 `assets\fonts` 는 공개 라이브러리·글꼴을 묶은 것이다(MIT / SIL OFL, 각 파일 머리와 `LICENSE-fonts.txt` 참고).

## 5. 다른 PC로 스킬 옮기기

- `scripts\pack.ps1 -Out "폴더"` 가 `ppt-html.zip` 과 `ppt-html 설치 방법.txt` 를 만든다. 받는 사람은 압축을 풀어 `%USERPROFILE%\.claude\skills\ppt-html` 로 넣는다(폴더를 통째로 복사해도 같다). 설치할 것이 없다.
- 받은 PC에서 `scripts\doctor.ps1` 로 점검한다. `-Full` 을 붙이면 실제로 한 장을 만들어 숨은 슬라이드 쇼와 웹 버전까지 시험한다.
- 필요한 것: Windows 10/11, PowerPoint 데스크톱(2013 이상 대상), Windows PowerShell 5.1(기본 내장). Edge(기본 내장)가 있으면 자동 변환·글꼴 내장·웹 버전 확인까지 된다.
- 경로에 한글·공백이 있어도 된다. 사용자 이름이 한글이어도 된다.
- PowerShell 실행 정책이 그룹 정책으로 `AllSigned`/`Restricted` 로 강제된 PC, 제한 언어 모드(AppLocker/WDAC)인 PC에서는 스크립트가 돌지 않는다(`doctor.ps1` 이 알려 준다).
- Noto Sans/Serif KR 이 설치되지 않은 PC(Windows 10 등)에서는 스킬의 글꼴 묶음(자주 쓰는 한글 2,350자 + 영문·기호)에서 잘라 넣는다. 묶음에 없는 드문 글자는 그 PC의 다른 글꼴로 보인다.

## 6. 이 스킬이 조심하는 것

- 사용자의 PowerPoint 창·열린 파일을 건드리지 않는다. 원본 덱은 읽기 전용으로만 연다.
- 레지스트리·보안 설정·신뢰 위치를 바꾸지 않는다. 받은 파일의 차단 표식을 대신 풀어 주지 않는다(안내문으로 방법만 알린다).
- 기존 매크로가 있는 덱은 거부한다(덮어쓰지 않는다).
- `pack.ps1` 은 보안 프로그램이 PowerShell의 읽기를 막는 파일을 만나면 그 파일만 빼고 알려 준다(다른 방법으로 읽어 넣지 않는다). 실측한 PC에서는 방금 만든 그림 파일(`gallery.png`, 페이지 그림)이 한동안 그렇게 막혔다가 몇 분 뒤 풀렸다(보안 프로그램이 새 그림을 검사하는 동안으로 보인다). 빠졌다고 나오면 잠시 뒤 다시 실행한다.
- 같은 PC에서 2026-11-20 에는 방금 만든 테마 파일 셋(`theme-material.css`, `theme-carbon-dark.css`, `theme-primer-dark.css`)을 PowerShell만 읽지 못했다(내용이 같은 사본을 다른 폴더에 두면 읽혔다. 파일 내용 때문이 아니다).
  `new-deck.ps1` 은 그런 파일을 빼고 복사한 뒤 알려 주고, 고른 테마가 막힌 것이면 멈춘다(그 파일을 자기 파일 도구로 덱의 `kit\` 에 복사해 두면 계속된다: `ACCESS-DENIED` 때와 같은 방법).
  `dev\make-gallery.ps1` 도 같은 식으로 빼고 알려 주며, `-KitDir` 로 읽을 수 있는 키트 사본을 줄 수 있다.
- 쇼 시험은 덱을 읽기 전용으로 연다. 시험 중에 페이지가 보관한 값(고친 글자, 투표)은 파일에 남지 않는다.
- 일부 PC의 보안 프로그램은 PowerShell이 바탕 화면·문서 폴더의 문서·그림을 읽는 것을 막는다. 덱은 PowerPoint가 직접 열어 사본을 만들고, HTML이 막히면 `ACCESS-DENIED` 로 멈춰 사본 위치를 안내한다(우회하지 않는다).

## 7. 유지보수

- **VBA를 고쳤다면** 아무 페이지로 `html2ppt.ps1` 을 돌려 본다. 컴파일 오류가 있으면 실측 단계에서 PowerPoint 대화 상자가 잡혀 보고에 나오고, `work\...\run\*.vbe*.png` 에 편집기 캡처가 남는다. 원본은 ASCII, 한 줄 1000자 이하.
- **검사 규칙**은 `cs\HtmlPrep.cs` 의 `JsRules`/`CssRules`/`HtmlRules`. 보완 함수로 덮는 규칙은 `Polyfilled` 배열과 `assets\polyfills-ie11.txt` 를 함께 고친다. 자동 변환이 처리해서 더는 오류가 아닌 규칙은 변환기가 돌려주는 `allow` 목록(`dev\src\convert.js` 의 `convertPage`)에 있다.
- **준비 단계의 흐름**은 `HtmlPrep.Begin`(파일 합치기, 변환이 필요한지 판단) → `ApplyConversion`(Edge가 돌려준 것 반영) → `Finish`(뼈대·연결 코드·검사·보완 코드). `lib\pages.ps1` 이 순서를 잡고, 엔진에서 스크립트가 실패한 페이지는 변환해서 한 번 더 그린다.
- **변환 규칙·글꼴 도구를 고치려면** `dev\` 폴더에서(Node 필요): `npm install` → `node build.js`(→ `assets\tools\*.js`, `corejs.min.txt`) / `node fontpack.js`(→ `assets\fonts`). 쓰는 PC에는 Node가 필요 없다.
- **레이아웃 어댑터**(`assets\adapt-ie11.txt`)는 ES5, 줄 주석(`//`) 없이 쓴다(한 줄로 접어서 넣는다). 표의 열쇠: `cols`/`rows`(트랙), `areas`, `cg`/`rg`(간격), `cs`/`ce`/`rs`/`re`(자리), `ji`/`js`(가로 정렬), `ratio`, `fit`/`fpos`.
- **페이지 연결 코드**는 `assets\harness-*.txt`(한 줄). VBA가 찾는 이름: `#ihCmd`, `#ihShow`, `body[data-ih-ready]`, `html[data-ih-err]`.
- **웹 뷰어**는 `assets\viewer-template.txt`. `/*__DECK__*/null` 자리에 데이터가 들어간다.
- **키트**를 고치면 `dev\make-gallery.ps1` 을 돌린다(템플릿 전체와 테마 전체를 엔진으로 다시 그려 `assets\templates\gallery.png`, `assets\kit\themes.png` 를 바꾼다. PowerShell만 쓴다). 두 그림을 열어 본다. 키트 CSS에는 `gap`/grid 를 쓰지 않고 키트 JS는 ES5 로 쓴다(변환 없이 돌아야 한다).
  움직임을 고쳤다면 템플릿 전체를 한 덱으로 빌드해 장마다의 `fps while appearing` 을 본다(v2.1 기준 36~53, 가장 긴 프레임 35~75ms).
- **테마**는 `dev\src\themes.js` 가 만든다(`node build.js themes` → `assets\kit\theme-*.css`, `themes.json`). 공개 디자인 값 패키지(Material 색 도구, Fluent·Carbon·Cloudscape·Primer 토큰)에서 뽑는 것과 손으로 적은 것이 한 파일에 있다. 테마를 더하면 `assets\templates\index.json` 의 `themes` 에도 적는다.
- **템플릿**을 더하면 `assets\templates\index.json` 의 `templates`(와 어울리는 `recipes`)에 적는다. `new-deck.ps1` 은 그 목록만 본다.
- **레퍼런스 디자인 읽기**(v2.3): HTML·웹 주소는 Edge 가 열어 `assets\tools\style.js` 로 잰다(넓이로 가중한 배경·글자색, 채도 높은 색, 제목·본문 글꼴과 크기, 모서리, 그림자, 단추·카드의 생김새).
  PPT 는 숨은 PowerPoint 에서 `lib\deck_style.ps1` 이 읽는다(테마 색·글꼴, 장마다의 바탕, 도형 채움을 넓이로, 글자는 처음과 가운데 글자의 서식으로). 그림은 `ImgKit.Compare.Palette` 로 색만.
  잰 값을 테마 변수로 옮기는 규칙은 `lib\style.ps1` 의 `New-ThemeTokens` 한 곳에 있다(강조색은 바탕과 대비 2.0 이상인 채도 높은 색, 카드 색은 바탕과 대비 1.03~1.8 인 채움, 글꼴은 내장할 수 있거나 Windows 기본인 것만 남기고 한글 글꼴을 뒤에 붙인다).
  고르는 규칙을 고쳤다면 `get-style.ps1` 을 HTML 한 장, PPT 하나에 돌려 `ref_*.png` 와 `preview.png` 를 나란히 본다. `lib\style.ps1` 안의 한글 글꼴 이름은 문자 코드로 적혀 있다(원본은 ASCII).
- **쇼 시험으로 보관 기능 시험하기**: `show_test.ps1` 에 `keep = true`(쓰기로 열고 쇼 뒤에 저장)와 `only = [장 번호]` 를 주면, 한 번 고치고 저장한 뒤 다시 열어 읽는 시험을 할 수 있다(사본으로만 한다).
- **다른 환경 흉내**(시험용 환경 변수): `PPTHTML_NO_EDGE=1`(Edge 없는 PC처럼), `PPTHTML_NO_SYSTEM_FONTS=1`(글꼴 묶음만 있는 PC처럼), `PPTHTML_EDGE_DROP=N`(N번째 장을 준비하는 도중 Edge 도구의 연결이 끊긴 것처럼).
- **Edge 도구가 도중에 사라질 때**: `lib\edge.ps1` 의 `Test-EdgeTool`(아직 대답하는지) / `Restart-EdgeTool`(한 번의 실행에 두 번까지 다시 띄움). `lib\pages.ps1` 이 장마다 시작할 때와 변환·글꼴 단계가 실패했을 때 확인한다. 글꼴 내장은 Edge 가 사라진 경우 그 장에 아무것도 넣지 않고 넘겨서 다시 하게 한다.

## 8. 실측한 것 / 못 한 것 (2026-11-20, Windows 11, PowerPoint LTSC 2024 64비트, Edge 157, 4K 150%)

실측함(숨은 데스크톱의 실제 슬라이드 쇼): 새 덱·기존 덱·표식 도형·부분 영역 배치, 전체 화면/창 모드, 캔버스 확대(1280×720 정확, 창 모드 1282×721), 클릭 반응, `pptHtml.next/prev/gotoSlide`, `pptHtmlOnShow`, 페이지에 포커스가 있을 때의 ←/→ 전달, 쇼 종료 후 임시 파일 삭제.
키트 템플릿 11종: 쇼에서 전부 라이브, 템플릿마다 한두 번씩 눌러 화면이 바뀌는 것 확인(기간 단추, 지표 칩, 표 정렬·줄 선택, 일정 점, 단계, 우선순위, 슬라이더 클릭, 질문, 스위치).
자동 변환: 최신 문법으로 쓴 시험 페이지 두 장(grid 영역·`auto-fit`·`span`, flex `gap`, CSS 변수와 `[data-theme]` 테마 전환, 중첩, `:is()`, `color-mix()`, `aspect-ratio`, `object-fit`, class·비공개 필드·`async`/`await`·`Map`/`Set`·`?.`·모듈 스크립트) → 엔진 그림이 최신 브라우저 그림과 99.7~99.8% 일치, 쇼에서 클릭 뒤 다시 배치·테마 전환 확인.
글꼴 내장(설치된 글꼴에서 / 글꼴 묶음에서), 웹 버전을 Edge 헤드리스로 열어 모든 라이브 페이지 실행 확인, Edge 없는 경우·글꼴 묶음만 있는 경우의 흉내.

v2.1 (같은 날, 같은 PC): 템플릿 23종 가운데 22종을 한 덱으로 빌드해 쇼에서 전부 라이브·스크립트 오류 없음, 나타날 때 36~53fps(가장 긴 프레임 35~75ms), 시험 동작 중 40~54fps. `end` 템플릿은 예시 덱에서 확인.
장마다 실제로 누르거나 끌어 값이 바뀌는 것 확인(문제/해법 전환, 시장 원, 경쟁 점, 기능 단계, 가격 전환·구성·수량, 팀, 후기, 투자 금액 슬라이더 끌기, 가정 슬라이더 끌기, 투표, 스위치).
덱 도구: 하나씩 보이기(1/3), 길게 눌러 슬라이드 고르기(22장 그림, 바깥 눌러 닫기, 누른 채 12번 그림 위에서 떼어 이동), 확대 켜고 끄기, 고치기(글자 입력, 카드 끌어 옮기기, 글자 크게), 펜 선 그리기.
보관: 사본을 쓰기로 열어 투표·글자 고치기(HTML에 적힌 제목, 스크립트가 만든 방안 이름)·카드 옮기기·스위치를 한 뒤 저장하고, 다시 열어 쇼에서 그대로인 것을 읽음(꼬리표 126~335자). 타이머는 눌러 시작·멈춤 확인.
테마: 17종 모두 엔진으로 그려 봄(`themes.png`). 쇼까지 돌린 것은 `cobalt`, `paper`, `indigo`, `midnight`, `carbon`. 예시 덱 3종(8·9·7장)은 쇼에서 전부 라이브, 38~54fps.
빌드 최적화: 4200×2800 PNG 사진 31.7MB → 3840×2560 JPEG 1.9MB, 누운 JPEG 바로 세움, 작은 도형 PNG 는 그대로, 큰 전환·흐린 그림자·`transition: all` 경고.

v2.2 (같은 날, 같은 PC): 도구 막대가 평소 숨어 있고 고치기를 켜면 보이는 것(계산된 `visibility` 로 확인). 숨은 쇼에서 제목을 고치고 투표를 누른 뒤 쇼를 끝내 저장하고, 파일을 다시 열어 슬라이드를 그림으로 내보내
고친 제목과 표 수가 **쇼 밖의 슬라이드 그림**에 나오는 것. 그 파일을 `ppt2html.ps1` 로 바꾼 웹 버전을 Edge 로 열어 고친 제목이 나오는 것.

v2.3 (같은 날, 같은 PC): 지어낸 레퍼런스 셋으로 `get-style.ps1` 을 돌림. 잡지풍 HTML 한 장(종이색 바탕·명조 제목·초록 강조·각진 모서리), 어두운 남색 PPT 다섯 장(노란 강조·둥근 카드), 그림 한 장(색만).
각각 뽑힌 테마의 미리보기가 레퍼런스 그림과 같은 인상인 것을 눈으로 확인. HTML 에서 뽑은 테마로 빈 페이지 세 장(최신 CSS grid·`gap`, 한 장은 최신 JavaScript)을 자유롭게 써서 빌드: 쇼에서 전부 라이브, 34~40fps, 엔진 그림이 최신 브라우저 그림과 97.5~99.3% 일치.
그 과정에서 찾은 변환 빈틈(`align-content: start` 를 엔진이 모른다) 하나를 고침.
`design.md` 만 받은 길도 끝까지: 지어낸 디자인 문서 → `new-theme.ps1`(미리보기 확인) → `new-deck.ps1 -Theme <파일> -Blank 2` → 자유 페이지 두 장 → 쇼에서 라이브(30~32fps: 같은 시각 다른 덱도 낮게 나온 때였다), 누른 뒤 값이 바뀌는 것을 읽음.
빌드 도중 Edge 도구의 연결이 한 번 끊긴 일이 있었다(변환과 글꼴 내장이 건너뛰어져 "페이지를 고치라"는 엉뚱한 오류로 끝남). 이제 끊기면 Edge 를 다시 띄워 그 장을 한 번 더 한다(`PPTHTML_EDGE_DROP` 으로 흉내 내어 확인).
비교 그림의 기준(최신 브라우저 그림)은 글꼴을 내장한 뒤 다시 찍는다: 이 PC 에 없는 글꼴(Inter 등)을 쓴 장에서 글꼴 차이가 배치 차이로 잡혀 97.9% + 경고였던 것이 99.3% 가 됨.

실측 못 함: 실제 웹 주소로 `get-style.ps1` 을 돌리는 것(로컬 HTML 파일로만 확인), 남이 만든 크고 복잡한 레퍼런스(수십 장 PPT, 프레임워크로 만든 사이트), 레퍼런스의 CSS 를 그대로 쓰거나 PPT 템플릿을 바탕 덱으로 쓰는 "그대로 가져오기" 흐름, 실제 모니터에서 사람이 누르는 것, 받은 파일의 [콘텐츠 사용] 직후 첫 쇼, `<input>` 입력창과 한글 입력기 조합, Enter 로 글자 고치기 끝내기(숨은 쇼에서는 키가 페이지에 닿지 않았다), 설명 풍선(마우스 올리기), 마우스를 구석으로 가져갔을 때 도구 막대가 나타나는 것(코드로만 넣음. 숨은 쇼에서는 마우스 움직임을 흉내 내지 않았다), PowerPoint를 닫을 때 저장을 묻는 대화 상자(저장 자체는 확인), 사용자가 Esc 로 쇼를 끝냈을 때의 그림 갱신(시험은 프로그램으로 쇼를 끝냈다), 웹 버전 안에서의 고치기·보관(PowerPoint에서 고친 것이 웹 버전에 나오는 것만 확인), 테마 12종의 쇼 실행(그림으로만 확인. 그중 `material`·`carbon-dark`·`primer-dark` 는 위 6장의 사정으로 다른 폴더의 사본으로 그렸다), 발표자 도구 두 화면, 배율이 다른 두 모니터, 32비트 Office, Office 2013~2021/Microsoft 365, Windows 10, 다른 보안 정책의 PC, Edge가 정책으로 막힌 PC.
자동 변환은 위 두 장과 같은 흔한 배치에서 확인한 것이다. 남의 HTML은 저마다 다르므로 비교 그림을 보는 단계를 건너뛰지 않는다.
