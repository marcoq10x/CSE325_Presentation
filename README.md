# Notespack — Frontend Build Deck

A premium, fully animated keynote presentation for Marco's frontend contribution
to the **Notespack** campus events platform (CSE 325, BYU-Idaho · Team "The Syntax Errors").

Built to feel like a modern product launch: Apple-keynote pacing, Stripe-level
polish, Linear-level restraint. Every slide opens with a sequenced entrance and
every slide change uses a deliberate transition.

## Deliverable

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
