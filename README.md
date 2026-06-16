# Notespack — Frontend Build Deck

A premium, fully animated presentation for Marco's frontend contribution to the
**Notespack** campus events platform (CSE 325, BYU-Idaho · Team "The Syntax Errors").

## ⭐ Main deliverable — the live web presentation

**`index.html`** — a self-contained animated presentation that runs in any browser.
This is the showpiece. **Just double-click `index.html`** (or open it in Chrome/Edge/Safari).

What it does that a slideshow can't:
- **Real 3D card tilt** — cards lean toward your cursor in true perspective
- **Cursor glow + halo** — a light follows your pointer; cards light up from inside
- **Animated gradient mesh** background with parallax orbs that react to the mouse
- **Count-up stat numbers** (0 → 9 / 6 / 0) and an animated timeline spine
- **Staggered entrance reveals** on every slide + 3D slide transitions

Controls: **→ ← or Space** to move, **F** for fullscreen, click/scroll/swipe also work,
and the dots at the bottom jump to any slide. Present it fullscreen (F) for max impact.

> Tip: it loads premium web fonts when online and falls back to clean system fonts
> offline — either way it looks great. To share a link, drop it on GitHub Pages
> (Settings ▸ Pages ▸ deploy from branch) and the URL serves `index.html` directly.

## Secondary deliverable — the PowerPoint version

- **`deck/notespack_frontend_v3.pptx`** — the presentation. Open in **Microsoft
  PowerPoint** for the full effect (entrance animations + transitions auto-play on
  slide enter). 12 widescreen slides (13.33 × 7.5"), ~70 KB, system fonts only (Calibri).
- **`deck/notespack_frontend_v3_preview.pdf`** — static page-per-slide preview.
- **`deck/preview_contact_sheet.png`** — all 12 slides at a glance.

> Animations and transitions are stored as raw OOXML (`p:timing` / `p:transition`).
> They render in **PowerPoint** (Mac/Windows). LibreOffice/Google Slides do not
> render PowerPoint animations reliably, so preview motion in PowerPoint.

## Design system

| Role | Hex |
|------|-----|
| Background (deep) | `#07071A` |
| Card surface | `#141432` |
| Card border | `#252548` |
| Primary text | `#EAF1FF` |
| Secondary text | `#9AA3C4` |
| Accent (cyan) | `#00D4FF` |
| Violet | `#6B3EFF` |
| Pink | `#FF4EA8` |

Signature gradient G1 (cyan→violet) and energy gradient G2 (violet→pink) are used
for full-bleed accent slides, the cover blobs, chips, and accent rules.

## Motion language

- Three verbs only: **fade**, **fade + 12px rise**, **wipe**.
- Entrances 0.4–0.8s, staggered 0.06–0.15s between siblings.
- Transitions: **Morph** where elements persist (1→2, 9→10, 11→12),
  **Push from right** between content sections, **Fade through black** for the two
  full-bleed accessibility slides (5→6, 6→7).
- Every slide auto-plays on enter — no clicking through builds.

## Rebuilding

```bash
pip install python-pptx pillow
python3 build_deck.py        # writes /tmp/output/notespack_frontend_v3.pptx
python3 render_qa.py         # renders QA PNGs + contact sheet to /tmp/qa/
```

`build_deck.py` contains the full layout, a gradient/shadow/glow engine, and the
animation/transition OOXML injector. `render_qa.py` rasterizes the *actual* saved
`.pptx` (gradients, alpha compositing, word-wrapped text) for visual QA.
