# Third-party notices

PPHTML-BD bundles the following openly licensed works. Their licences apply to those files; the MIT licence in `LICENSE` covers the rest of the project.

## Fonts (`assets/fonts`) — SIL Open Font License 1.1

Subsets (fewer characters, otherwise unchanged) of: Noto Sans KR, Noto Serif KR, Pretendard, Inter, Roboto, Open Sans, Geist, IBM Plex Sans KR.
Copyright statements of each font are in [`assets/fonts/LICENSE-fonts.txt`](assets/fonts/LICENSE-fonts.txt). Licence text: https://openfontlicense.org

## Build tools (`assets/tools`, `assets/corejs.min.txt`) — run at build time in Edge, never inside a presentation

| File | Contains | Licence |
|---|---|---|
| `assets/tools/babel.min.js` | @babel/standalone | MIT |
| `assets/tools/convert.js` | PostCSS, Autoprefixer, postcss-nesting, postcss-css-variables, postcss-flexbugs-fixes, postcss-selector-parser, postcss-value-parser, acorn, plus this project's conversion rules | MIT |
| `assets/tools/fonttool.js` | harfbuzzjs (HarfBuzz), pako | MIT (HarfBuzz: Old MIT) |
| `assets/corejs.min.txt` | core-js 3.51 | MIT |

Theme values in `assets/kit/theme-material.css`, `theme-fluent.css`, `theme-carbon*.css`, `theme-cloudscape.css`, `theme-primer*.css` are derived from the public design-token packages of those design systems (`@material/material-color-utilities` Apache-2.0, `@fluentui/tokens` MIT, `@carbon/themes` Apache-2.0, `@cloudscape-design/design-tokens` Apache-2.0, `@primer/primitives` MIT). They are not official material of those companies and carry no logos or trademarks.

Sources and the scripts that rebuild these files are in `dev/` (see `dev/README.md`).
