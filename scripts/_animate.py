"""Native PowerPoint animation from a spec: entrance effects, motion paths and easing, written as <p:timing>.

A slide's `animate` field lists steps; each step names its shapes by name or pattern ("Card*", "Value?"),
an effect (appear, fade, float_up, wipe, zoom, path), when it starts (`at`, seconds from slide start; default
right after the step before) and how long it runs (`dur`). `stagger` reveals the matched shapes one group after
another, a group being every shape that shares the first number in its name (Card1, Card1Edge, Value1, Label1).
`ease` (linear, in, out, in_out) becomes PowerPoint's accel / decel. A `path` step moves a shape along points in
pt relative to its own place - straight segments, or a smooth curve through them with `curve`.

Everything starts automatically from one trigger with absolute delays (reference/ANIMATION.md). The XML recipes
mirror the ones verified against PowerPoint in the kPptx TypeScript library (PptxTimingAtoms / PptxTimingEffects).
kTimingReader reads them back for the onion sheet (onion_sheet.py).
"""
import fnmatch
import re

from lxml import etree

from kShared import ToolInputException, kS, kToolException

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
EMU_PER_PT = 12700
MAX_BUILD_S = 5.0  # longer than this without the presenter's control: a11y_motion_overload (ANIMATION.md)


class kAnimSchema:
    """The effects, easings and the JSON Schema of a slide's `animate` field."""

    # effect -> (presetID, presetClass, default duration s, default ease)
    EFFECTS = {"appear": (1, "entr", 0.0, "linear"), "fade": (10, "entr", 0.5, "linear"),
               "float_up": (37, "entr", 0.6, "out"), "wipe": (22, "entr", 0.6, "linear"),
               "zoom": (53, "entr", 0.5, "out"), "path": (0, "path", 1.0, "in_out")}
    EASE = {"linear": (0, 0), "in": (100000, 0), "out": (0, 100000), "in_out": (50000, 50000)}
    WIPE = {"left": 8, "right": 2, "top": 1, "bottom": 4}  # direction it comes from -> presetSubtype
    WIPE_FILTER = {"left": "wipe(left)", "right": "wipe(right)", "top": "wipe(up)", "bottom": "wipe(down)"}

    @staticmethod
    def Schema():
        """JSON Schema of `animate`: a list of steps."""
        if kS.ErrorMode:
            return {}
        try:
            Names = {"oneOf": [{"type": "string"}, {"type": "array", "minItems": 1, "items": {"type": "string"}}]}
            Point = {"type": "array", "minItems": 2, "maxItems": 2, "items": {"type": "number"}}
            Step = {"type": "object", "additionalProperties": False, "required": ["target", "effect"], "properties": {
                "target": dict(Names, description="Shape name(s) on the built slide; * and ? match (\"Card*\"). "
                               "build_deck.py lists the names when one matches nothing."),
                "effect": {"enum": list(kAnimSchema.EFFECTS)},
                "at": {"type": "number", "minimum": 0, "description": "Start, seconds from slide start "
                       "(default: when the step before ends)."},
                "dur": {"type": "number", "minimum": 0, "maximum": 10, "description": "Seconds per shape."},
                "ease": {"enum": list(kAnimSchema.EASE), "description": "Default: out for float_up and zoom, "
                         "in_out for path, else linear."},
                "stagger": {"type": "number", "minimum": 0, "description": "Seconds between the groups of "
                            "matched shapes (a group = the shapes sharing the first number in their name)."},
                "direction": {"enum": list(kAnimSchema.WIPE), "description": "wipe: the side it comes from."},
                "from": {"type": "number", "exclusiveMinimum": 0, "description": "zoom: start scale (default 0.5)."},
                "rise": {"type": "number", "description": "float_up: pt it rises from (default 30)."},
                "path": {"type": "array", "minItems": 1, "items": Point, "description": "path: points [x, y] in "
                         "pt from the shape's place (it starts at [0, 0]); y grows downwards."},
                "curve": {"type": "boolean", "description": "path: a smooth curve through the points."}}}
            return {"type": "array", "minItems": 1, "items": Step,
                    "description": "Native animation: steps that start automatically from slide start "
                                   "(reference/ANIMATION.md). Keep the whole build near 2 s."}
        except Exception as e:
            kS.GlobalErrorHandler(e, "kAnimSchema.Schema")
            return {}


class kMotionPath:
    """PowerPoint's path language: points in pt <-> "M 0 0 L x y ... E" in fractions of the slide."""

    @staticmethod
    def Num(Value):
        """A path number: 5 decimals, no trailing zeros."""
        if kS.ErrorMode:
            return "0"
        try:
            Text = f"{Value:.5f}".rstrip("0").rstrip(".")
            return "0" if Text in ("", "-0") else Text
        except Exception as e:
            kS.GlobalErrorHandler(e, "kMotionPath.Num")
            return "0"

    @staticmethod
    def Build(Points, WidthPt, HeightPt, Curve=False):
        """The path from [0, 0] through Points (pt from the shape's place); Curve = Catmull-Rom through them."""
        if kS.ErrorMode:
            return ""
        try:
            P = [(0.0, 0.0)] + [(float(X) / WidthPt, float(Y) / HeightPt) for X, Y in Points]
            Out = ["M 0 0"]
            for I in range(1, len(P)):
                X, Y = P[I]
                if not Curve or len(P) < 3:
                    Out.append(f"L {kMotionPath.Num(X)} {kMotionPath.Num(Y)}")
                    continue
                A, B, C = P[max(I - 2, 0)], P[I - 1], P[min(I + 1, len(P) - 1)]
                C1 = (B[0] + (X - A[0]) / 6, B[1] + (Y - A[1]) / 6)
                C2 = (X - (C[0] - B[0]) / 6, Y - (C[1] - B[1]) / 6)
                Out.append("C " + " ".join(kMotionPath.Num(V) for V in (*C1, *C2, X, Y)))
            return " ".join(Out) + " E"
        except Exception as e:
            kS.GlobalErrorHandler(e, "kMotionPath.Build")
            return ""

    @staticmethod
    def Parse(Path, Samples=8):
        """The points a path passes through, in fractions of the slide (curves sampled Samples times)."""
        if kS.ErrorMode:
            return []
        try:
            Tokens = re.findall(r"[A-Za-z]|-?\d*\.?\d+(?:[eE]-?\d+)?", Path or "")
            Out, Cur, I, Cmd = [], (0.0, 0.0), 0, "M"
            while I < len(Tokens):
                if Tokens[I].isalpha():
                    Cmd, I = Tokens[I].upper(), I + 1
                    if Cmd in ("E", "Z"):
                        continue
                    if I >= len(Tokens):
                        break
                Nums = 6 if Cmd == "C" else 2
                if Cmd not in ("M", "L", "C") or I + Nums > len(Tokens) or Tokens[I].isalpha():
                    I += 1  # a command this reader does not draw: skip its token
                    continue
                V = [float(T) for T in Tokens[I:I + Nums]]
                I += Nums
                if Cmd == "C":
                    for K in range(1, Samples + 1):
                        T = K / Samples
                        Out.append(kMotionPath.Bezier(Cur, (V[0], V[1]), (V[2], V[3]), (V[4], V[5]), T))
                    Cur = (V[4], V[5])
                else:
                    Cur = (V[0], V[1])
                    Out.append(Cur)
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kMotionPath.Parse")
            return []

    @staticmethod
    def Bezier(P0, P1, P2, P3, T):
        """A point on a cubic Bezier at T."""
        if kS.ErrorMode:
            return (0.0, 0.0)
        try:
            U = 1 - T
            return tuple(U ** 3 * A + 3 * U * U * T * B + 3 * U * T * T * C + T ** 3 * D
                         for A, B, C, D in zip(P0, P1, P2, P3))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kMotionPath.Bezier")
            return (0.0, 0.0)


class kTimingXml:
    """The XML atoms and effect bodies (ids are @ID@ until Number)."""

    @staticmethod
    def Ctn(DurMs=None, DelayMs=0, Ease=(0, 0)):
        """A behaviour's timing node."""
        if kS.ErrorMode:
            return ""
        try:
            A = '<p:cTn id="@ID@"' + (f' dur="{round(DurMs)}"' if DurMs is not None else "")
            A += (f' accel="{Ease[0]}"' if Ease[0] else "") + (f' decel="{Ease[1]}"' if Ease[1] else "")
            A += ' fill="hold"'
            if DelayMs > 0:
                return A + f'><p:stCondLst><p:cond delay="{round(DelayMs)}"/></p:stCondLst></p:cTn>'
            return A + "/>"
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTimingXml.Ctn")
            return ""

    @staticmethod
    def Target(Spid):
        """The shape a behaviour acts on."""
        if kS.ErrorMode:
            return ""
        try:
            return f'<p:tgtEl><p:spTgt spid="{Spid}"/></p:tgtEl>'
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTimingXml.Target")
            return ""

    @staticmethod
    def Visible(Spid):
        """style.visibility -> visible (a 1 ms node, as PowerPoint writes it)."""
        if kS.ErrorMode:
            return ""
        try:
            return ("<p:set><p:cBhvr>" + kTimingXml.Ctn(1) + kTimingXml.Target(Spid) +
                    '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
                    '<p:to><p:strVal val="visible"/></p:to></p:set>')
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTimingXml.Visible")
            return ""

    @staticmethod
    def Filter(Spid, Filter, DurMs, Ease):
        """A built-in transition filter (fade, wipe(left) ...) coming in."""
        if kS.ErrorMode:
            return ""
        try:
            return (f'<p:animEffect transition="in" filter="{Filter}"><p:cBhvr>' + kTimingXml.Ctn(DurMs, 0, Ease) +
                    kTimingXml.Target(Spid) + "</p:cBhvr></p:animEffect>")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTimingXml.Filter")
            return ""

    @staticmethod
    def Prop(Spid, Attr, Start, DurMs, Ease):
        """ppt_x / ppt_y / ppt_w / ppt_h from Start (a formula) to the shape's own value."""
        if kS.ErrorMode:
            return ""
        try:
            return ('<p:anim calcmode="lin" valueType="num"><p:cBhvr>' + kTimingXml.Ctn(DurMs, 0, Ease) +
                    kTimingXml.Target(Spid) + f"<p:attrNameLst><p:attrName>{Attr}</p:attrName></p:attrNameLst>"
                    f'</p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="{Start}"/></p:val></p:tav>'
                    f'<p:tav tm="100000"><p:val><p:strVal val="#{Attr}"/></p:val></p:tav></p:tavLst></p:anim>')
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTimingXml.Prop")
            return ""

    @staticmethod
    def Motion(Spid, Path, DurMs, Ease):
        """A motion path relative to the shape's place, in fractions of the slide."""
        if kS.ErrorMode:
            return ""
        try:
            return (f'<p:animMotion origin="layout" path="{Path}" pathEditMode="relative" ptsTypes=""><p:cBhvr>' +
                    kTimingXml.Ctn(DurMs, 0, Ease) + kTimingXml.Target(Spid) +
                    "<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst>"
                    "</p:cBhvr></p:animMotion>")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTimingXml.Motion")
            return ""

    @staticmethod
    def Body(Step, Spid, DurMs, Ease, Path, HeightPt):
        """The behaviours of one effect on one shape."""
        if kS.ErrorMode:
            return ""
        try:
            Fx = Step["effect"]
            if Fx == "path":
                return kTimingXml.Motion(Spid, Path, DurMs, Ease)
            Vis = kTimingXml.Visible(Spid)
            if Fx == "appear":
                return Vis
            if Fx == "fade":
                return Vis + kTimingXml.Filter(Spid, "fade", DurMs, Ease)
            if Fx == "wipe":
                return Vis + kTimingXml.Filter(Spid, kAnimSchema.WIPE_FILTER[Step.get("direction", "left")], DurMs, Ease)
            if Fx == "float_up":
                Rise = kMotionPath.Num(float(Step.get("rise", 30)) / HeightPt)
                return (Vis + kTimingXml.Filter(Spid, "fade", DurMs * 0.8, (0, 0)) +
                        kTimingXml.Prop(Spid, "ppt_y", f"#ppt_y+{Rise}", DurMs, Ease))
            if Fx == "zoom":
                From = kMotionPath.Num(float(Step.get("from", 0.5)))
                return (Vis + kTimingXml.Prop(Spid, "ppt_w", f"#ppt_w*{From}", DurMs, Ease) +
                        kTimingXml.Prop(Spid, "ppt_h", f"#ppt_h*{From}", DurMs, Ease) +
                        kTimingXml.Filter(Spid, "fade", DurMs, (0, 0)))
            raise ValueError(f"no body for effect '{Fx}'")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTimingXml.Body")
            return ""

    @staticmethod
    def Number(Xml):
        """Replace every @ID@ with 1, 2, 3 ... in document order."""
        if kS.ErrorMode:
            return ""
        try:
            Parts = Xml.split("@ID@")
            return "".join(P + (str(I + 1) if I < len(Parts) - 1 else "") for I, P in enumerate(Parts))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTimingXml.Number")
            return ""

    @staticmethod
    def Timing(Effects):
        """The whole <p:timing> for planned effects (dicts from kSlideAnimation.Plan); "" when there are none."""
        if kS.ErrorMode:
            return ""
        try:
            if not Effects:
                return ""
            Fx, Bld, Groups = "", "", {}
            for I, E in enumerate(Effects):
                Grp = ""
                if E["Sp"]:
                    Groups[E["Spid"]] = Groups.get(E["Spid"], -1) + 1
                    Grp = f' grpId="{Groups[E["Spid"]]}"'
                    Bld += f'<p:bldP spid="{E["Spid"]}"{Grp} animBg="1"/>'
                Node = "afterEffect" if I == 0 else "withEffect"
                Fx += (f'<p:par><p:cTn id="@ID@" presetID="{E["Preset"]}" presetClass="{E["Class"]}" '
                       f'presetSubtype="{E["Subtype"]}" fill="hold"{Grp} nodeType="{Node}"><p:stCondLst>'
                       f'<p:cond delay="{E["DelayMs"]}"/></p:stCondLst><p:childTnLst>{E["Body"]}'
                       "</p:childTnLst></p:cTn></p:par>")
            Step = ('<p:par><p:cTn id="@ID@" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
                    f"<p:childTnLst>{Fx}</p:childTnLst></p:cTn></p:par>")
            Group = ('<p:par><p:cTn id="@ID@" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>'
                     '<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst>'
                     f"<p:childTnLst>{Step}</p:childTnLst></p:cTn></p:par>")
            Xml = (f'<p:timing xmlns:p="{NS_P}"><p:tnLst><p:par><p:cTn id="@ID@" dur="indefinite" restart="never" '
                   'nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="@ID@" '
                   f'dur="indefinite" nodeType="mainSeq"><p:childTnLst>{Group}</p:childTnLst></p:cTn>'
                   '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>'
                   '</p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl>'
                   "</p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>" +
                   (f"<p:bldLst>{Bld}</p:bldLst>" if Bld else "") + "</p:timing>")
            return kTimingXml.Number(Xml)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTimingXml.Timing")
            return ""


class kSlideAnimation:
    """Writes a slide's `animate` steps into the slide as native animation."""

    def __init__(self, Slide, SlideWidth, SlideHeight):
        try:
            self.Slide = Slide
            self.WidthPt, self.HeightPt = SlideWidth / EMU_PER_PT, SlideHeight / EMU_PER_PT
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideAnimation.__init__")

    def Shapes(self, Target):
        """The slide's top-level shapes whose name matches Target (a name, a pattern or a list), in slide order."""
        if kS.ErrorMode:
            return []
        try:
            Patterns = Target if isinstance(Target, list) else [Target]
            return [S for S in self.Slide.shapes if any(fnmatch.fnmatchcase(S.name, P) for P in Patterns)]
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideAnimation.Shapes")
            return []

    @staticmethod
    def GroupKey(Name):
        """The first number in a shape name (Card1Edge -> 1); -1 for a name without one."""
        if kS.ErrorMode:
            return -1
        try:
            Found = re.search(r"\d+", Name)
            return int(Found.group()) if Found else -1
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideAnimation.GroupKey")
            return -1

    @staticmethod
    def Groups(Shapes):
        """Shapes grouped by their first number, groups in number order (unnumbered first)."""
        if kS.ErrorMode:
            return []
        try:
            By = {}
            for S in Shapes:
                By.setdefault(kSlideAnimation.GroupKey(S.name), []).append(S)
            return [By[K] for K in sorted(By)]
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideAnimation.Groups")
            return []

    def Plan(self, Steps, Report):
        """The effects of every step (dicts for kTimingXml.Timing), sorted by start. Report(shape, code, message)
        receives a target that matches nothing and a build longer than MAX_BUILD_S."""
        if kS.ErrorMode:
            return []
        try:
            Effects, Clock = [], 0.0
            for No, Step in enumerate(Steps, 1):
                Preset, Cls, Dur, Ease = kAnimSchema.EFFECTS[Step["effect"]]
                Dur = float(Step.get("dur", Dur))
                Ease = kAnimSchema.EASE[Step.get("ease", Ease)]
                Shapes = self.Shapes(Step["target"])
                if not Shapes:
                    Names = ", ".join(S.name for S in self.Slide.shapes)
                    Report("animate", "animate_target", f"animate step {No}: target {Step['target']!r} matches no "
                                                        f"shape; the slide has: {Names}")
                    continue
                if Step["effect"] == "path" and not Step.get("path"):
                    raise ToolInputException(f"animate step {No}: effect 'path' needs 'path' points")
                Path = kMotionPath.Build(Step.get("path", []), self.WidthPt, self.HeightPt, Step.get("curve", False))
                Start = float(Step.get("at", Clock))
                Subtype = kAnimSchema.WIPE[Step.get("direction", "left")] if Step["effect"] == "wipe" else \
                    16 if Step["effect"] == "zoom" else 0
                for I, Group in enumerate(self.Groups(Shapes)):
                    At = Start + I * float(Step.get("stagger", 0))
                    for S in Group:
                        Effects.append({"Spid": S.shape_id, "Sp": etree.QName(S._element).localname == "sp",
                                        "Preset": Preset, "Class": Cls, "Subtype": Subtype,
                                        "DelayMs": round(At * 1000), "End": At + Dur,
                                        "Body": kTimingXml.Body(Step, S.shape_id, Dur * 1000, Ease, Path,
                                                                self.HeightPt)})
                    Clock = max(Clock, At + Dur)
            End = max((E["End"] for E in Effects), default=0.0)
            if End > MAX_BUILD_S:
                Report("animate", "animate_long", f"the build runs {End:.1f} s (more than {MAX_BUILD_S:.0f} s "
                                                  "without the presenter's control); shorten it, about 2 s reads best")
            return sorted(Effects, key=kSlideAnimation.StartOf)
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideAnimation.Plan")
            return []

    @staticmethod
    def StartOf(Effect):
        """Sort key: an effect's delay."""
        if kS.ErrorMode:
            return 0
        try:
            return Effect["DelayMs"]
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideAnimation.StartOf")
            return 0

    def Apply(self, Steps, Report):
        """Plan the steps and write <p:timing> into the slide (replacing one already there); returns the effects."""
        if kS.ErrorMode:
            return 0
        try:
            Effects = self.Plan(Steps, Report)
            Xml = kTimingXml.Timing(Effects)
            if not Xml:
                return 0
            Root = self.Slide._element
            for Old in Root.findall(f"{{{NS_P}}}timing"):
                Root.remove(Old)
            After = None
            for Tag in ("transition", "clrMapOvr", "cSld"):  # schema order: cSld, clrMapOvr, transition, timing
                After = Root.find(f"{{{NS_P}}}{Tag}")
                if After is not None:
                    break
            if After is None:
                raise ValueError("slide has no p:cSld")
            After.addnext(etree.fromstring(Xml))
            return len(Effects)
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSlideAnimation.Apply")
            return 0


class kTimingReader:
    """Reads a slide's animation back: one entry per effect, in start order."""

    @staticmethod
    def Read(Slide):
        """[{'spid', 'class', 'preset', 'delay_ms', 'dur_ms', 'path'}] of the slide's main sequence; [] when none.
        path: the motion path's points in fractions of the slide (path effects only)."""
        if kS.ErrorMode:
            return []
        try:
            Out = []
            for Ctn in Slide._element.iter(f"{{{NS_P}}}cTn"):
                Cls = Ctn.get("presetClass")
                if not Cls:
                    continue
                Cond = Ctn.find(f"{{{NS_P}}}stCondLst/{{{NS_P}}}cond")
                Delay = int(Cond.get("delay", "0")) if Cond is not None and Cond.get("delay", "").isdigit() else 0
                Target = next(Ctn.iter(f"{{{NS_P}}}spTgt"), None)
                if Target is None:
                    continue
                Durs = [int(C.get("dur")) for C in Ctn.iter(f"{{{NS_P}}}cTn") if (C.get("dur") or "").isdigit()]
                Motion = next(Ctn.iter(f"{{{NS_P}}}animMotion"), None)
                Out.append({"spid": int(Target.get("spid")), "class": Cls, "preset": int(Ctn.get("presetID", "0")),
                            "delay_ms": Delay, "dur_ms": max(Durs, default=0),
                            "path": kMotionPath.Parse(Motion.get("path")) if Motion is not None else []})
            return sorted(Out, key=kTimingReader.DelayOf)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTimingReader.Read")
            return []

    @staticmethod
    def DelayOf(Entry):
        """Sort key: an entry's delay."""
        if kS.ErrorMode:
            return 0
        try:
            return Entry["delay_ms"]
        except Exception as e:
            kS.GlobalErrorHandler(e, "kTimingReader.DelayOf")
            return 0
