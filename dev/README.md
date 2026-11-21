# dev — 도구 묶음의 원본

스킬을 **쓰는 데는 필요 없다.** `assets/tools`(Edge 안에서 도는 변환·글꼴 도구), `assets/fonts`(글꼴 묶음), `assets/kit/theme-*.css`(테마)를 다시 만들 때만 쓴다. Node 가 필요하다.

```
npm install
node build.js          # src/convert.js, src/fonttool.js -> ../assets/tools/*.js, ../assets/corejs.min.txt
node build.js themes   # src/themes.js -> ../assets/kit/theme-*.css, themes.json (공개 디자인 값 패키지에서 계산)
node fontpack.js       # 글꼴 패키지와 C:WindowsFonts 의 Noto Sans/Serif KR -> ../assets/fonts/*.subset.ttf (build.js 를 먼저 돌린다)
```

- `src/convert.js` — CSS·JavaScript 변환 규칙(PostCSS 플러그인, 격자 정보 추출, Babel 호출).
- `src/fonttool.js` — HarfBuzz 로 글꼴 자르기, WOFF 포장.
- `src/themes.js` — 테마 17종. Material(색 도구로 계산), Fluent·Carbon·Cloudscape·Primer(공개 토큰에서 뽑음), 나머지는 손으로 적은 값. 변수 이름은 모든 테마가 같아야 한다.
- `fontpack.js` — 글꼴마다 자주 쓰는 한글 2,350자 + 영문·기호만 남긴다. 글꼴의 사용 조건은 `../assets/fonts/LICENSE-fonts.txt`.
- 고친 뒤에는 최신 문법 시험 페이지와 키트 템플릿을 `check-html.ps1` 로 다시 그려 비교 그림을 확인한다.
