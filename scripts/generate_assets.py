from __future__ import annotations

from html import escape
from pathlib import Path
from textwrap import dedent

ROOT = Path(__file__).resolve().parents[1]
ASSETS_DIR = ROOT / "assets"

PALETTE = {
    "paper": "#F3EEE6",
    "paper_alt": "#FBF8F2",
    "ink": "#162028",
    "ink_alt": "#2B3B46",
    "line": "#D6C8B6",
    "shadow": "#BDAE9B",
    "text": "#FFF8EF",
    "muted": "#D8CEC0",
    "dark_text": "#223039",
    "accent": "#C36A35",
    "teal": "#2F6E72",
    "blue": "#4C6F94",
    "gold": "#B8913C",
    "sage": "#5E755B",
    "brick": "#9A4E45",
    "slate": "#465A6D",
}

FONT_MONO = (
    '"JetBrainsMono Nerd Font", "Hack Nerd Font", "CaskaydiaMono Nerd Font", '
    '"SFMono-Regular", Menlo, Monaco, Consolas, "Liberation Mono", monospace'
)


def write_asset(filename: str, content: str) -> None:
    (ASSETS_DIR / filename).write_text(content, encoding="utf-8")


def layered_card(
    x: int,
    y: int,
    width: int,
    height: int,
    radius: int,
    fill: str,
    *,
    stroke: str,
    shadow_dx: int = 8,
    shadow_dy: int = 8,
) -> str:
    return dedent(
        f"""\
        <rect x="{x + shadow_dx}" y="{y + shadow_dy}" width="{width}" height="{height}" rx="{radius}" fill="{PALETTE["shadow"]}" />
        <rect x="{x}" y="{y}" width="{width}" height="{height}" rx="{radius}" fill="{fill}" stroke="{stroke}" stroke-width="2" />
        """
    )


def create_banner() -> str:
    return dedent(
        f"""\
        <svg width="1200" height="360" viewBox="0 0 1200 360" fill="none" xmlns="http://www.w3.org/2000/svg">
          <defs>
            <pattern id="ledger" width="48" height="48" patternUnits="userSpaceOnUse">
              <path d="M 48 0 L 0 0 0 48" fill="none" stroke="{PALETTE["line"]}" stroke-width="1" />
            </pattern>
            <clipPath id="frame-clip">
              <rect width="1200" height="360" rx="28" />
            </clipPath>
            <style><![CDATA[
              .badge {{ font-family: {FONT_MONO}; font-size: 13px; font-weight: 700; fill: {PALETTE["dark_text"]}; }}
              .eyebrow {{ font-family: {FONT_MONO}; font-size: 15px; font-weight: 700; letter-spacing: 0.14em; fill: {PALETTE["accent"]}; }}
              .title {{ font-family: {FONT_MONO}; font-size: 40px; font-weight: 700; fill: {PALETTE["text"]}; }}
              .subtitle {{ font-family: {FONT_MONO}; font-size: 18px; font-weight: 700; fill: {PALETTE["muted"]}; }}
              .footer {{ font-family: {FONT_MONO}; font-size: 14px; font-weight: 700; fill: {PALETTE["dark_text"]}; }}
              .panel-label {{ font-family: {FONT_MONO}; font-size: 12px; font-weight: 700; letter-spacing: 0.1em; fill: {PALETTE["accent"]}; }}
              .panel-title {{ font-family: {FONT_MONO}; font-size: 16px; font-weight: 700; fill: {PALETTE["dark_text"]}; }}
              .panel-copy {{ font-family: {FONT_MONO}; font-size: 15px; fill: {PALETTE["dark_text"]}; }}
              .panel-copy-muted {{ font-family: {FONT_MONO}; font-size: 14px; fill: {PALETTE["ink_alt"]}; }}
            ]]></style>
          </defs>

          <g clip-path="url(#frame-clip)">
            <rect width="1200" height="360" rx="28" fill="{PALETTE["paper"]}" />
            <rect width="1200" height="360" rx="28" fill="url(#ledger)" opacity="0.72" />
            <rect x="0" y="298" width="1200" height="62" fill="{PALETTE["paper_alt"]}" />

            {layered_card(42, 34, 724, 250, 22, PALETTE["ink"], stroke="#2F404A").strip()}
            <rect x="70" y="58" width="286" height="32" rx="9" fill="{PALETTE["paper"]}" />
            <text x="90" y="79" class="badge">&gt; stream-transcript-extractor</text>
            <rect x="70" y="112" width="28" height="3" rx="1.5" fill="{PALETTE["accent"]}" />
            <text x="114" y="118" class="eyebrow">BROWSER-DRIVEN STREAM EXTRACTION</text>
            <text x="70" y="168" class="title">Stream Transcript Extractor</text>
            <text x="70" y="212" class="subtitle">browser-profile attach</text>
            <text x="70" y="240" class="subtitle">network-first transcript capture</text>
            <text x="70" y="268" class="subtitle">automatic panel recovery and DOM fallback</text>

            {layered_card(806, 48, 314, 132, 18, PALETTE["paper_alt"], stroke=PALETTE["line"], shadow_dx=6, shadow_dy=6).strip()}
            <text x="834" y="80" class="panel-label">[ modes ]</text>
            <text x="834" y="106" class="panel-title">capture paths</text>
            <text x="834" y="136" class="panel-copy">network  automatic  dom</text>
            <text x="834" y="160" class="panel-copy-muted">payload first, UI fallback</text>

            {layered_card(806, 188, 314, 136, 18, PALETTE["paper_alt"], stroke=PALETTE["line"], shadow_dx=6, shadow_dy=6).strip()}
            <text x="834" y="220" class="panel-label">[ outputs ]</text>
            <text x="834" y="246" class="panel-title">artifacts</text>
            <text x="834" y="274" class="panel-copy">json  markdown</text>
            <text x="834" y="296" class="panel-copy">.network.json</text>
            <text x="834" y="318" class="panel-copy-muted">debug traces and release builds</text>

            <rect x="786" y="60" width="10" height="216" rx="5" fill="{PALETTE["gold"]}" />
            <rect x="1144" y="62" width="16" height="16" rx="3" fill="{PALETTE["teal"]}" />
            <rect x="1144" y="88" width="16" height="16" rx="3" fill="{PALETTE["accent"]}" />
            <rect x="1144" y="114" width="16" height="16" rx="3" fill="{PALETTE["blue"]}" />

            <text x="70" y="334" class="footer">Chrome  •  Edge  •  Bun  •  network capture  •  automatic fallback  •  standalone builds</text>
          </g>
        </svg>
        """
    )


def create_header(text: str, accent: str) -> str:
    label = escape(text)
    card_width = max(320, min(620, 176 + len(text) * 16))
    return dedent(
        f"""\
        <svg width="920" height="82" viewBox="0 0 920 82" fill="none" xmlns="http://www.w3.org/2000/svg">
          <defs>
            <style><![CDATA[
              .label {{ font-family: {FONT_MONO}; font-size: 23px; font-weight: 700; fill: {PALETTE["dark_text"]}; }}
              .meta {{ font-family: {FONT_MONO}; font-size: 11px; font-weight: 700; letter-spacing: 0.14em; fill: {accent}; }}
            ]]></style>
          </defs>
          <rect x="8" y="14" width="{card_width}" height="48" rx="14" fill="{PALETTE["shadow"]}" />
          <rect x="0" y="6" width="{card_width}" height="48" rx="14" fill="{PALETTE["paper_alt"]}" stroke="{PALETTE["line"]}" stroke-width="2" />
          <rect x="22" y="0" width="92" height="18" rx="6" fill="{accent}" />
          <text x="33" y="13" class="meta">SECTION</text>
          <rect x="24" y="30" width="26" height="3" rx="1.5" fill="{accent}" />
          <text x="62" y="35" class="label" dominant-baseline="middle">{label}</text>
        </svg>
        """
    )


def main() -> None:
    ASSETS_DIR.mkdir(parents=True, exist_ok=True)

    write_asset("banner.svg", create_banner())
    write_asset("header-overview.svg", create_header("Overview", PALETTE["accent"]))
    write_asset("header-install.svg", create_header("Install and run", PALETTE["teal"]))
    write_asset("header-modes.svg", create_header("Modes", PALETTE["blue"]))
    write_asset(
        "header-outputs.svg",
        create_header("Outputs and diagnostics", PALETTE["sage"]),
    )
    write_asset(
        "header-architecture.svg",
        create_header("How it works", PALETTE["gold"]),
    )
    write_asset(
        "header-scripts.svg",
        create_header("Scripts and assets", PALETTE["brick"]),
    )
    write_asset("header-build.svg", create_header("Build", PALETTE["slate"]))
    write_asset(
        "header-troubleshooting.svg",
        create_header("Troubleshooting", PALETTE["teal"]),
    )
    write_asset(
        "header-repository.svg",
        create_header("Repository layout", PALETTE["accent"]),
    )


if __name__ == "__main__":
    main()
