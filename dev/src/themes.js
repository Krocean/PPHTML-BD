// Dev-time: writes the theme files of the design kit (assets\kit\theme-*.css).
// The themes named after a public design system take their values from that system's own published tokens
// (npm packages, open licences): Material 3 (computed with material-color-utilities), Fluent 2, Carbon,
// Primer, Cloudscape. The other themes are written by hand; where one follows the look of a product whose
// design system is not public, the name is generic and the values are approximations.
// Run through build.js:  node build.js themes
import fs from "fs";
import path from "path";
import * as carbon from "@carbon/themes";
import { webLightTheme as fluent } from "@fluentui/tokens";
import { Hct, SchemeTonalSpot, MaterialDynamicColors, argbFromHex, hexFromArgb } from "@material/material-color-utilities";
import cloudscape from "@cloudscape-design/design-tokens/index-visual-refresh.json";
import primerLight from "@primer/primitives/dist/css/functional/themes/light.css";
import primerDark from "@primer/primitives/dist/css/functional/themes/dark.css";

const outDir = process.argv[2];
const rgba = (hex, a) => { const n = parseInt(hex.replace("#", ""), 16); return "rgba(" + (n >> 16) + "," + ((n >> 8) & 255) + "," + (n & 255) + "," + a + ")"; };
const up = h => String(h).toUpperCase();

const NOTO = '"Noto Sans KR", "Malgun Gothic", sans-serif';
const SERIF = '"Noto Serif KR", "Batang", serif';
const PRE = '"Pretendard", "Noto Sans KR", "Malgun Gothic", sans-serif';

// every theme defines every name (the build writes the values out; nothing is inherited)
function theme(o) {
  const dark = !!o.dark;
  const t = Object.assign({
    bg2: o.bg, ink3: o.ink2, "accent-ink": "#FFFFFF",
    hover: rgba(o.ink, dark ? 0.07 : 0.05), hover2: rgba(o.ink, dark ? 0.11 : 0.08),
    press: o.ink, "press-ink": o.bg,
    display: o.body, "display-weight": "700", "display-tracking": "-0.02em", "strong-weight": "700",
    "label-tracking": "0.04em", "label-case": "none",
    radius: "14px", "radius-s": "999px", "radius-btn": "999px",
    shadow: "0 0 0 1px " + rgba(o.ink, dark ? 0.1 : 0.07) + ", 0 2px 6px -2px " + rgba(dark ? "#000000" : o.ink, dark ? 0.5 : 0.12),
    "shadow-pop": "0 0 0 1px " + rgba(o.ink, dark ? 0.14 : 0.08) + ", 0 12px 28px -8px " + rgba(dark ? "#000000" : o.ink, dark ? 0.7 : 0.28),
    "card-border": "0 solid transparent"
  }, o);
  delete t.dark; delete t.note; delete t.source;
  return t;
}

const themes = {};

// ---------------------------------------------------------------- the kit's own themes
themes.cobalt = { note: "clean white and blue, sans headlines (business reports) - the default", v: theme({ bg: "#F4F6FA", bg2: "#E9EDF4", surface: "#FFFFFF", ink: "#0F1B2D", ink2: "#5B6B82", ink3: "#8E9BAE", line: "#DCE2EC", accent: "#2F6BFF", "accent-soft": "#E3ECFF", "accent-deep": "#1E4FD0", good: "#12905F", "good-soft": "#DBF1E7", bad: "#D64545", "bad-soft": "#FADEDE", warn: "#B7791F", "warn-soft": "#F7EBD0", c1: "#2F6BFF", c2: "#14B8A6", c3: "#F59E0B", c4: "#8B5CF6", c5: "#EF6A6A", c6: "#94A3B8", body: NOTO, "press-ink": "#FFFFFF" }) };
themes.paper = { note: "warm paper, serif headlines, clay accent (editorial)", v: theme({ bg: "#F3EFE7", bg2: "#EAE5DA", surface: "#FFFDF9", ink: "#1B1A16", ink2: "#6F6A5F", ink3: "#9C968A", line: "#DDD6C8", accent: "#D97757", "accent-soft": "#F6E3DA", "accent-deep": "#B5532F", good: "#2E7D5B", "good-soft": "#DDEEE4", bad: "#C2432F", "bad-soft": "#F6DDD7", warn: "#A9740B", "warn-soft": "#F5E8C8", c1: "#D97757", c2: "#3F6B8C", c3: "#7A9A62", c4: "#C9A24B", c5: "#8B6BA6", c6: "#8C877B", body: NOTO, display: SERIF, "display-weight": "600", "display-tracking": "-0.025em", radius: "18px" }) };
themes.ink = { note: "dark, serif headlines, clay accent (keynote)", v: theme({ dark: true, bg: "#131311", bg2: "#1C1B18", surface: "#211F1B", ink: "#F3EFE7", ink2: "#A8A296", ink3: "#7A756A", line: "#37342D", accent: "#E08D6F", "accent-ink": "#131311", "accent-soft": "#3A2A22", "accent-deep": "#EBA58B", good: "#6CC79B", "good-soft": "#1F3329", bad: "#EE8272", "bad-soft": "#3B2420", warn: "#E2B552", "warn-soft": "#3A3018", c1: "#E08D6F", c2: "#7FB0D6", c3: "#A6C98A", c4: "#E2C16B", c5: "#B79AD1", c6: "#8E897E", body: NOTO, display: SERIF, "display-weight": "600", "display-tracking": "-0.025em", radius: "18px" }) };
themes.forest = { note: "soft green, serif headlines (public sector, sustainability)", v: theme({ bg: "#F1F4EE", bg2: "#E5EBE0", surface: "#FFFFFF", ink: "#14201A", ink2: "#5C6B60", ink3: "#8D9A90", line: "#D5DDD0", accent: "#1F8A5B", "accent-soft": "#DCEFE4", "accent-deep": "#156B45", good: "#1F8A5B", "good-soft": "#DCEFE4", bad: "#C4472F", "bad-soft": "#F5DDD6", warn: "#A9740B", "warn-soft": "#F3E7C6", c1: "#1F8A5B", c2: "#2C6E91", c3: "#D9A441", c4: "#7C6BB0", c5: "#C9694F", c6: "#93A096", body: NOTO, display: SERIF, "display-weight": "600", "display-tracking": "-0.025em", radius: "18px" }) };
themes.snow = { note: "pure white, black type, one signal colour (minimal)", v: theme({ bg: "#FFFFFF", bg2: "#F4F4F2", surface: "#FFFFFF", ink: "#111111", ink2: "#62625E", ink3: "#9B9B96", line: "#E4E4E0", accent: "#F0522B", "accent-soft": "#FDE6DF", "accent-deep": "#C93C18", good: "#1D8A5A", "good-soft": "#DFF1E8", bad: "#D23B2B", "bad-soft": "#F9DDD9", warn: "#B07A12", "warn-soft": "#F6EACB", c1: "#111111", c2: "#F0522B", c3: "#8D8D88", c4: "#2B6FE0", c5: "#C9C9C4", c6: "#62625E", body: NOTO, radius: "12px", "card-border": "1px solid #E4E4E0", shadow: "0 0 0 0 transparent" }) };

// ---------------------------------------------------------------- Material 3 (Google): computed from the baseline seed colour
{
  const scheme = new SchemeTonalSpot(Hct.fromInt(argbFromHex("#6750A4")), false, 0);
  const c = name => up(hexFromArgb(MaterialDynamicColors[name].getArgb(scheme)));
  themes.material = { note: "Material 3 baseline (Google): tonal surfaces, Roboto + Noto Sans KR", source: "@material/material-color-utilities (Apache-2.0), SchemeTonalSpot from the baseline seed #6750A4; type and shape from the Material 3 guidelines",
    v: theme({ bg: c("surface"), bg2: c("surfaceContainer"), surface: c("surfaceContainerLow"), ink: c("onSurface"), ink2: c("onSurfaceVariant"), ink3: c("outline"), line: c("outlineVariant"),
      accent: c("primary"), "accent-ink": c("onPrimary"), "accent-soft": c("primaryContainer"), "accent-deep": c("onPrimaryContainer"),
      good: "#146C2E", "good-soft": "#C4EED0", bad: c("error"), "bad-soft": c("errorContainer"), warn: "#7C5800", "warn-soft": "#FFDEA6",
      c1: c("primary"), c2: c("tertiary"), c3: "#006A6A", c4: c("secondary"), c5: "#7C5800", c6: c("outline"),
      body: '"Roboto", "Noto Sans KR", sans-serif', "display-weight": "400", "display-tracking": "0", "strong-weight": "500", "label-tracking": "0.02em",
      radius: "12px", "radius-s": "8px", "radius-btn": "999px", press: c("primary"), "press-ink": c("onPrimary"),
      shadow: "0 1px 2px rgba(0,0,0,.3), 0 1px 3px 1px rgba(0,0,0,.15)", "shadow-pop": "0 1px 3px rgba(0,0,0,.3), 0 4px 8px 3px rgba(0,0,0,.15)" }) };
}

// ---------------------------------------------------------------- Fluent 2 (Microsoft)
themes.fluent = { note: "Fluent 2 (Microsoft): neutral greys, brand blue, soft small shadows; Segoe UI + Noto Sans KR", source: "@fluentui/tokens (MIT), webLightTheme",
  v: theme({ bg: up(fluent.colorNeutralBackground2), bg2: up(fluent.colorNeutralBackground3), surface: up(fluent.colorNeutralBackground1), ink: up(fluent.colorNeutralForeground1), ink2: up(fluent.colorNeutralForeground3), ink3: up(fluent.colorNeutralForeground4), line: up(fluent.colorNeutralStroke2),
    accent: up(fluent.colorBrandBackground), "accent-soft": up(fluent.colorBrandBackground2), "accent-deep": up(fluent.colorBrandForeground2),
    good: up(fluent.colorPaletteGreenForeground1), "good-soft": up(fluent.colorPaletteGreenBackground1), bad: up(fluent.colorPaletteRedForeground1), "bad-soft": up(fluent.colorPaletteRedBackground1), warn: up(fluent.colorPaletteDarkOrangeForeground1), "warn-soft": up(fluent.colorPaletteDarkOrangeBackground1),
    c1: up(fluent.colorBrandBackground), c2: up(fluent.colorPaletteTealForeground2), c3: up(fluent.colorPaletteMarigoldForeground2), c4: up(fluent.colorPalettePurpleForeground2), c5: up(fluent.colorPaletteCranberryForeground2), c6: up(fluent.colorNeutralForeground4),
    body: '"Segoe UI Variable Text", "Segoe UI", "Noto Sans KR", "Malgun Gothic", sans-serif', display: '"Segoe UI Variable Display", "Segoe UI", "Noto Sans KR", "Malgun Gothic", sans-serif', "display-weight": "600", "display-tracking": "-0.01em", "strong-weight": "600",
    radius: fluent.borderRadiusXLarge, "radius-s": fluent.borderRadiusMedium, "radius-btn": fluent.borderRadiusMedium, press: up(fluent.colorBrandBackground), "press-ink": "#FFFFFF",
    shadow: fluent.shadow4, "shadow-pop": fluent.shadow16 }) };

// ---------------------------------------------------------------- Carbon (IBM)
{
  const w = carbon.white, d = carbon.g100;
  themes.carbon = { note: "Carbon (IBM), white theme: square corners, flat layers, IBM Plex Sans KR", source: "@carbon/themes (Apache-2.0), theme 'white'; chart colours from the Carbon categorical palette",
    v: theme({ bg: up(w.background), bg2: up(w.layer01), surface: up(w.layer01), ink: up(w.textPrimary), ink2: up(w.textSecondary), ink3: up(w.textHelper), line: up(w.borderSubtle01),
      accent: up(w.interactive), "accent-soft": "#EDF5FF", "accent-deep": up(w.linkPrimaryHover || "#0043CE"),
      good: "#198038", "good-soft": "#DEFBE6", bad: up(w.supportError), "bad-soft": "#FFF1F1", warn: "#8E6A00", "warn-soft": "#FCF4D6",
      c1: "#6929C4", c2: "#1192E8", c3: "#005D5D", c4: "#9F1853", c5: "#FA4D56", c6: "#8D8D8D",
      body: '"IBM Plex Sans KR", "IBM Plex Sans", "Noto Sans KR", sans-serif', "display-weight": "300", "display-tracking": "0", "strong-weight": "600", "label-tracking": "0.02em",
      radius: "0px", "radius-s": "0px", "radius-btn": "0px", press: up(w.interactive), "press-ink": "#FFFFFF",
      shadow: "0 0 0 0 transparent", "shadow-pop": "0 2px 6px rgba(0,0,0,.3)" }) };
  themes["carbon-dark"] = { note: "Carbon (IBM), gray 100 theme: the dark counterpart", source: "@carbon/themes (Apache-2.0), theme 'g100'",
    v: theme({ dark: true, bg: up(d.background), bg2: up(d.layer01), surface: up(d.layer01), ink: up(d.textPrimary), ink2: up(d.textSecondary), ink3: up(d.textHelper), line: up(d.borderSubtle01),
      accent: up(d.interactive), "accent-soft": "#001D6C", "accent-deep": "#A6C8FF",
      good: "#42BE65", "good-soft": "#022D0D", bad: "#FA4D56", "bad-soft": "#2D0709", warn: "#F1C21B", "warn-soft": "#302400",
      c1: "#8A3FFC", c2: "#33B1FF", c3: "#007D79", c4: "#FF7EB6", c5: "#FA4D56", c6: "#6F6F6F",
      body: '"IBM Plex Sans KR", "IBM Plex Sans", "Noto Sans KR", sans-serif', "display-weight": "300", "display-tracking": "0", "strong-weight": "600", "label-tracking": "0.02em",
      radius: "0px", "radius-s": "0px", "radius-btn": "0px", press: up(d.interactive), "press-ink": "#FFFFFF",
      shadow: "0 0 0 0 transparent", "shadow-pop": "0 2px 6px rgba(0,0,0,.6)" }) };
}

// ---------------------------------------------------------------- Cloudscape (Amazon Web Services)
{
  const t = cloudscape.tokens, g = k => { const v = t[k] && t[k].$value; return up(v && v.light !== undefined ? v.light : v); };
  themes.cloudscape = { note: "Cloudscape (AWS): white containers, rounded, bold headings; Open Sans + Noto Sans KR", source: "@cloudscape-design/design-tokens (Apache-2.0), visual refresh, light",
    v: theme({ bg: g("color-background-cell-shaded"), bg2: "#EBEBF0", surface: g("color-background-container-content"), ink: g("color-text-body-default"), ink2: g("color-text-body-secondary"), ink3: g("color-text-status-inactive"), line: g("color-border-divider-default"),
      accent: g("color-background-button-primary-default"), "accent-soft": g("color-background-status-info"), "accent-deep": "#004A9E",
      good: g("color-text-status-success"), "good-soft": g("color-background-status-success"), bad: g("color-text-status-error"), "bad-soft": g("color-background-status-error"), warn: g("color-text-status-warning"), "warn-soft": g("color-background-status-warning"),
      c1: "#688AE8", c2: "#C33D69", c3: "#2EA597", c4: "#8456CE", c5: "#E07941", c6: "#656871",
      body: '"Open Sans", "Noto Sans KR", sans-serif', "display-weight": "700", "display-tracking": "-0.02em",
      radius: String(t["border-radius-container"].$value), "radius-s": "8px", "radius-btn": String(t["border-radius-button"].$value), press: g("color-background-button-primary-default"), "press-ink": "#FFFFFF",
      shadow: "0 0 0 1px #C6C6CD", "shadow-pop": "0 4px 20px 1px rgba(0,7,22,.1)" }) };
}

// ---------------------------------------------------------------- Primer (GitHub)
{
  const read = css => name => { const m = new RegExp("--" + name + ":\\s*([^;]+);").exec(css); return m ? up(m[1].trim()) : ""; };
  const l = read(primerLight), d = read(primerDark);
  themes.primer = { note: "Primer (GitHub), light: white canvas, hairline borders, small radius; Segoe UI + Pretendard", source: "@primer/primitives (MIT), functional theme 'light'",
    v: theme({ bg: l("bgColor-muted"), bg2: l("bgColor-neutral-muted") || "#EFF2F5", surface: l("bgColor-default"), ink: l("fgColor-default"), ink2: l("fgColor-muted"), ink3: "#818B98", line: l("borderColor-default"),
      accent: l("bgColor-accent-emphasis"), "accent-soft": l("bgColor-accent-muted"), "accent-deep": "#0550AE",
      good: l("fgColor-success"), "good-soft": "#DAFBE1", bad: l("fgColor-danger"), "bad-soft": "#FFEBE9", warn: l("fgColor-attention"), "warn-soft": "#FFF8C5",
      c1: "#0969DA", c2: "#1A7F37", c3: "#8250DF", c4: "#BF3989", c5: "#BC4C00", c6: "#818B98",
      body: '"Segoe UI", "Pretendard", "Noto Sans KR", sans-serif', "display-weight": "600", "display-tracking": "-0.015em", "strong-weight": "600",
      radius: "6px", "radius-s": "999px", "radius-btn": "6px", press: "#1F883D", "press-ink": "#FFFFFF",
      "card-border": "1px solid " + l("borderColor-default"), shadow: "0 1px 0 rgba(31,35,40,.04)", "shadow-pop": "0 0 0 1px " + l("borderColor-default") + ", 0 8px 24px rgba(140,149,159,.2)" }) };
  themes["primer-dark"] = { note: "Primer (GitHub), dark", source: "@primer/primitives (MIT), functional theme 'dark'",
    v: theme({ dark: true, bg: d("bgColor-default"), bg2: d("bgColor-muted"), surface: d("bgColor-muted"), ink: d("fgColor-default"), ink2: d("fgColor-muted"), ink3: "#656C76", line: d("borderColor-default"),
      accent: "#1F6FEB", "accent-soft": "#121D2F", "accent-deep": d("fgColor-accent"),
      good: "#3FB950", "good-soft": "#12261E", bad: "#F85149", "bad-soft": "#25171C", warn: "#D29922", "warn-soft": "#272115",
      c1: "#4493F8", c2: "#3FB950", c3: "#AB7DF8", c4: "#DB61A2", c5: "#DB6D28", c6: "#656C76",
      body: '"Segoe UI", "Pretendard", "Noto Sans KR", sans-serif', "display-weight": "600", "display-tracking": "-0.015em", "strong-weight": "600",
      radius: "6px", "radius-s": "999px", "radius-btn": "6px", press: "#238636", "press-ink": "#FFFFFF",
      "card-border": "1px solid " + d("borderColor-default"), shadow: "0 0 0 0 transparent", "shadow-pop": "0 0 0 1px " + d("borderColor-default") + ", 0 8px 24px rgba(1,4,9,.8)" }) };
}

// ---------------------------------------------------------------- looks written by hand (approximations; no public token source)
themes.cupertino = { note: "keynote look in the manner of Apple's product pages: big tight type, grey tiles, pill buttons; Pretendard", source: "hand-written approximation (the original typeface is not open; Pretendard stands in)",
  v: theme({ bg: "#FFFFFF", bg2: "#F5F5F7", surface: "#F5F5F7", ink: "#1D1D1F", ink2: "#6E6E73", ink3: "#86868B", line: "#D2D2D7", accent: "#0071E3", "accent-soft": "#E8F2FF", "accent-deep": "#0066CC", good: "#248A3D", "good-soft": "#E3F5E7", bad: "#D70015", "bad-soft": "#FFE9EA", warn: "#B25000", "warn-soft": "#FFF1E0",
    c1: "#0071E3", c2: "#34C759", c3: "#FF9500", c4: "#AF52DE", c5: "#FF2D55", c6: "#8E8E93", body: PRE, "display-weight": "700", "display-tracking": "-0.03em", "strong-weight": "600",
    radius: "22px", "radius-s": "999px", "radius-btn": "999px", press: "#0071E3", "press-ink": "#FFFFFF", shadow: "0 0 0 0 transparent" }) };
themes.geist = { note: "developer-tool look in the manner of Vercel: black on white, hairlines, very tight headlines; Geist + Pretendard", source: "hand-written approximation; Geist typeface (SIL OFL)",
  v: theme({ bg: "#FAFAFA", bg2: "#F2F2F2", surface: "#FFFFFF", ink: "#171717", ink2: "#666666", ink3: "#8F8F8F", line: "#EAEAEA", accent: "#0070F3", "accent-soft": "#EBF5FF", "accent-deep": "#0060DF", good: "#297A3A", "good-soft": "#EBFAEB", bad: "#CB2A2F", "bad-soft": "#FFEBEB", warn: "#A35200", "warn-soft": "#FFF4D5",
    c1: "#171717", c2: "#0070F3", c3: "#7928CA", c4: "#FF0080", c5: "#F5A623", c6: "#8F8F8F", body: '"Geist", "Pretendard", "Noto Sans KR", sans-serif', "display-weight": "600", "display-tracking": "-0.045em", "strong-weight": "600",
    radius: "8px", "radius-s": "999px", "radius-btn": "6px", "card-border": "1px solid #EAEAEA", shadow: "0 0 0 0 transparent" }) };
themes.indigo = { note: "fintech look in the manner of Stripe: cool grey page, navy type, indigo accent; Inter + Pretendard", source: "hand-written approximation; Inter typeface (SIL OFL)",
  v: theme({ bg: "#F6F9FC", bg2: "#E9EFF6", surface: "#FFFFFF", ink: "#0A2540", ink2: "#425466", ink3: "#8898AA", line: "#E3E8EE", accent: "#635BFF", "accent-soft": "#EFEEFF", "accent-deep": "#4B45C6", good: "#0E6245", "good-soft": "#CBF4C9", bad: "#DF1B41", "bad-soft": "#FFE7EC", warn: "#B13600", "warn-soft": "#FCEDB9",
    c1: "#635BFF", c2: "#00B4D8", c3: "#FFB020", c4: "#FF80B5", c5: "#0A2540", c6: "#8898AA", body: '"Inter", "Pretendard", "Noto Sans KR", sans-serif', "display-weight": "700", "display-tracking": "-0.03em", "strong-weight": "600",
    radius: "12px", "radius-s": "999px", "radius-btn": "999px", press: "#635BFF", "press-ink": "#FFFFFF", shadow: "0 0 0 1px rgba(10,37,64,.06), 0 2px 5px -1px rgba(50,50,93,.2)" }) };
themes.midnight = { note: "dark product look in the manner of Linear: near-black, hairline cards, soft indigo; Inter + Pretendard", source: "hand-written approximation; Inter typeface (SIL OFL)",
  v: theme({ dark: true, bg: "#08090A", bg2: "#101113", surface: "#141516", ink: "#F7F8F8", ink2: "#8A8F98", ink3: "#62666D", line: "#23252A", accent: "#6E79E6", "accent-soft": "#1B1D3A", "accent-deep": "#A4ABFF", good: "#4CB782", "good-soft": "#12261C", bad: "#EB5757", "bad-soft": "#2C1616", warn: "#F2C94C", "warn-soft": "#2B2610",
    c1: "#7C86F0", c2: "#4CB782", c3: "#F2C94C", c4: "#26B5CE", c5: "#EB5757", c6: "#62666D", body: '"Inter", "Pretendard", "Noto Sans KR", sans-serif', "display-weight": "600", "display-tracking": "-0.03em", "strong-weight": "600",
    radius: "12px", "radius-s": "999px", "radius-btn": "8px", press: "#F7F8F8", "press-ink": "#08090A", "card-border": "1px solid #23252A", shadow: "0 0 0 0 transparent" }) };
themes.seoul = { note: "Korean consumer-app look (large rounded cards on a grey page, vivid blue); Pretendard", source: "hand-written approximation",
  v: theme({ bg: "#F2F4F6", bg2: "#E5E8EB", surface: "#FFFFFF", ink: "#191F28", ink2: "#6B7684", ink3: "#8B95A1", line: "#E5E8EB", accent: "#3182F6", "accent-soft": "#E8F3FF", "accent-deep": "#1B64DA", good: "#03A65A", "good-soft": "#E5F8F1", bad: "#F04452", "bad-soft": "#FFEBEE", warn: "#D97706", "warn-soft": "#FFF3E0",
    c1: "#3182F6", c2: "#15C47E", c3: "#FFC342", c4: "#A234C7", c5: "#F04452", c6: "#8B95A1", body: PRE, "display-weight": "700", "display-tracking": "-0.03em",
    radius: "20px", "radius-s": "999px", "radius-btn": "14px", press: "#3182F6", "press-ink": "#FFFFFF", shadow: "0 0 0 0 transparent" }) };

const ORDER = ["bg", "bg2", "surface", "ink", "ink2", "ink3", "line", "accent", "accent-ink", "accent-soft", "accent-deep", "good", "good-soft", "bad", "bad-soft", "warn", "warn-soft", "c1", "c2", "c3", "c4", "c5", "c6",
  "hover", "hover2", "press", "press-ink", "display", "display-weight", "display-tracking", "body", "strong-weight", "label-tracking", "label-case", "radius", "radius-s", "radius-btn", "card-border", "shadow", "shadow-pop"];
const index = [];
for (const name of Object.keys(themes)) {
  const t = themes[name];
  for (const k of ORDER) { if (t.v[k] === undefined || t.v[k] === "") { throw new Error("theme " + name + " lacks " + k); } }
  let css = "/* ppt-html design kit - theme \"" + name + "\": " + t.note + ".\n";
  if (t.source) { css += "   Values: " + t.source + ". Unofficial; product and company names belong to their owners.\n"; }
  css += "   Link ONE theme file before kit.css. To match a company style, copy this file and change the values\n   (keep every name; each is defined exactly once, here). */\n:root {\n";
  for (const k of ORDER) { css += "  --ph-" + k + ": " + t.v[k] + ";\n"; }
  css += "}\n";
  fs.writeFileSync(path.join(outDir, "theme-" + name + ".css"), css);
  index.push({ id: name, note: t.note, dark: /^#(0|1|2)/.test(t.v.bg) });
}
fs.writeFileSync(path.join(outDir, "themes.json"), JSON.stringify(index, null, 2) + "\n");
console.log("themes written: " + index.map(x => x.id).join(", "));
