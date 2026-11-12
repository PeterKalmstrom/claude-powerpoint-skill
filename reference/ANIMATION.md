# Animation

*Runs on: rules: any OS; COM/XML details as marked.* Part of the `building-powerpoint-decks` skill — start at [`SKILL.md`](../SKILL.md). Paths like `scripts/…` are relative to the skill folder.

## Animation — when and how

### The rule: animation shows cause and effect, or order

> **If the audience learns nothing from seeing it *move* that they wouldn't learn from the still, it is decoration.**

Use at most one decorative effect per slide, or none. Pick the effect by what it explains:

| What the motion explains | Use for | PowerPoint effect | Not for |
|---|---|---|---|
| **Sequence** — the next point arrives | headline, then sentence, then card | Float Up (or Fade) | something that *happens* |
| **A new part** with weight | a node, icon or part joining a system | Zoom | long text |
| **Before / after**, a reveal | a result replacing a blank, a panel revealed | Wipe | — |
| **Path of travel** | arrows, flow lines, connectors | Wipe in the direction of the line | filled shapes |
| **Scale** — how many | one big number | Fade (number already final) | a number people must read mid-count |
| **Outcome confirmed** | check mark, "approved", "passed" | Zoom from large to 100 % | anything that isn't an outcome |
| **Order of a list** | 3–6 steps, cards, rows | Float Up, one item after another | unrelated items |

Keep the whole build of a slide short — about **2 seconds** from slide start to the last element in place. People
wait for motion to finish before they read.

### Native animation traps

- **Start everything automatically from one trigger, with absolute delays.** First effect "After Previous" (it
  starts when the slide shows), every other effect "With Previous" plus its own delay from slide start. A chain of
  "After Previous" effects waits for each one in turn, so one slow effect shifts everything after it.
- **Click-triggered builds** are right for a live talk where the speaker paces the reveal; automatic builds are
  right for recorded or self-running decks. Decide per deck, not per slide.
- **A filled shape with text:** animate it as one object (in XML, `animBg="1"` on the build entry), or only the
  text moves and the box sits there from the start.
- **In XML** (no-COM builds): effects live in `<p:timing>`, inserted after `p:clrMapOvr`. Use the right preset ids —
  Fade 10, Float Up 37, Wipe 22 (subtype by direction: left 8, right 2, top 1, bottom 4), Zoom 53 (subtype 16),
  Appear 1. Only `p:sp` shapes get a `grpId`/`bldP` entry; a picture or group given the shape form is rejected.
- **PowerPoint silently drops animation XML it doesn't like.** Re-open the file and read the effects back (COM:
  `slide.TimeLine.MainSequence.Count`) — a file that opens is not proof the animation survived.
- Honour the audience: no endless loops on content people must read, and no motion lasting more than ~5 s without
  the presenter's control (`a11y_motion_overload`).

### Animation from the spec (`animate`, any OS)

`build_deck.py` writes native animation from a slide's `animate` list - no PowerPoint needed, and every effect
follows the traps above by construction (one automatic sequence, absolute delays, a filled shape built as one
object). Each step:

| Field | Meaning |
|---|---|
| `target` | Shape name(s) on the built slide; `*` and `?` match (`"Card*"`). A name that matches nothing is a `fit:` error that lists the slide's shapes. |
| `effect` | `appear`, `fade`, `float_up`, `wipe`, `zoom`, `path` |
| `at` / `dur` | Start in seconds from slide start (default: when the step before ends) / seconds per shape |
| `stagger` | Seconds between groups of matched shapes; a group = shapes sharing the first number in their name, so `Card1`, `Card1Edge`, `Value1`, `Label1` arrive together |
| `ease` | `linear`, `in`, `out`, `in_out` (PowerPoint accel/decel). Default: `out` for float_up and zoom, `in_out` for path |
| `direction` / `from` / `rise` | wipe: side it comes from · zoom: start scale (0.5) · float_up: pt it rises from (30) |
| `path` / `curve` | path: points `[x, y]` in pt from the shape's place (y grows downwards); `curve: true` draws a smooth curve through them |

```json
"animate": [
  {"target": "Title 1", "effect": "fade", "dur": 0.4},
  {"target": ["Card*", "Value*", "Label*"], "effect": "float_up", "stagger": 0.3},
  {"target": "Marker1", "effect": "path", "path": [[120, -60], [240, 0]], "curve": true, "at": 1.4}
]
```

A build that runs past 5 s is a `fit:` error (`animate_long`). Use a motion path for **travel** - a token moving
along a process, a marker to its place - not to decorate.

### Check the build without playing it: the onion sheet

`scripts/onion_sheet.py deck.pptx [--slides 4] [--out folder]` writes `onion-sNNN.png` per animated slide: a strip
of frames (before the build, then after every beat - effects that start together) and one onion frame - the
finished slide over a faint ghost of where the moving shapes start, each beat numbered where its shapes arrive,
every motion path drawn in amber with its end marked. Look at it once: wrong order, a path that overshoots or
leaves the slide, or too many beats show at a glance. It reads any deck's native animation, built or not; renders
through LibreOffice (any OS) or PowerPoint (Windows).

---
