"""Self-test checks for native animation from the spec (_animate.py) and the onion-skin sheet (onion_sheet.py).
Run by selftest.py (kSelfTestAnimate(Test).Run()); uses its Check / RunScript and temp folder. The onion sheet is
checked with kBoxRenderer (each shape drawn as a box), so it runs without LibreOffice or PowerPoint.
"""
import json
import os

from lxml import etree
from PIL import Image, ImageDraw
from pptx import Presentation

from _animate import NS_P, kMotionPath, kTimingReader
from kShared import kS
from onion_sheet import kBuildBeats, kOnionSheet

SAY = {"say": "This slide is part of the animation self-test and says what the presenter would say aloud."}
STEPS = [{"target": "Title 1", "effect": "fade", "dur": 0.4},
         {"target": ["Card*", "Value*", "Label*"], "effect": "float_up", "stagger": 0.3},
         {"target": "Value1", "effect": "path", "path": [[120, -60], [240, 0]], "curve": True, "at": 1.4},
         {"target": "Card2", "effect": "wipe", "direction": "top", "at": 1.4, "ease": "in_out"}]


class kBoxRenderer:
    """A stand-in renderer: every slide as a white PNG with each shape's outline."""

    def Render(self, Path, OutDir):
        """The renders of Path in slide order."""
        if kS.ErrorMode:
            return []
        try:
            os.makedirs(OutDir, exist_ok=True)
            Prs, Out = Presentation(Path), []
            Scale = 640 / Prs.slide_width
            for No, Slide in enumerate(Prs.slides, 1):
                Img = Image.new("RGB", (640, round(Prs.slide_height * Scale)), "white")
                Draw = ImageDraw.Draw(Img)
                for S in Slide.shapes:
                    Draw.rectangle((S.left * Scale, S.top * Scale, (S.left + S.width) * Scale,
                                    (S.top + S.height) * Scale), outline=(90, 140, 90), width=2)
                File = os.path.join(OutDir, f"s{No:03d}.png")
                Img.save(File)
                Out.append(File)
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kBoxRenderer.Render")
            return []


class kSelfTestAnimate:
    """The checks; Test is the running kSelfTest."""

    def __init__(self, Test):
        try:
            self.Test = Test
            self.Dir = os.path.join(Test.Tmp, "animate")
            os.makedirs(self.Dir, exist_ok=True)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTestAnimate.__init__")

    @staticmethod
    def Spec(Steps):
        """A cover and a kpi slide whose build is Steps."""
        if kS.ErrorMode:
            return {}
        try:
            return {"direction": "clean-corporate", "footer": "Animation test", "slides": [
                {"pattern": "title", "title": "Animation arrives in order", "subtitle": "Self-test"},
                {"pattern": "kpi", "title": "Three numbers arrive one after another", "notes": SAY, "animate": Steps,
                 "metrics": [{"value": "12", "label": "Teams"}, {"value": "40 %", "label": "Faster"},
                             {"value": "3", "label": "Sites"}]}]}
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTestAnimate.Spec")
            return {}

    def Build(self, Name, Spec):
        """Write Spec, build it; (exit code, output, deck path)."""
        if kS.ErrorMode:
            return 1, "", ""
        try:
            Path, Deck = os.path.join(self.Dir, f"{Name}.json"), os.path.join(self.Dir, f"{Name}.pptx")
            with open(Path, "w", encoding="utf-8") as Fh:
                json.dump(Spec, Fh)
            Code, Out = self.Test.RunScript("build_deck.py", Path, "--out", Deck)
            return Code, Out, Deck
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTestAnimate.Build")
            return 1, "", ""

    def CheckTiming(self):
        """The steps become one automatic sequence: effects, stagger, easing, wipe side and a curved path."""
        if kS.ErrorMode:
            return
        try:
            T = self.Test
            Code, Out, Deck = self.Build("animate", self.Spec(STEPS))
            T.Check("animate: a spec with build steps builds", Code == 0 and os.path.exists(Deck), Out[-600:])
            if not os.path.exists(Deck):
                return
            Slide = Presentation(Deck).slides[1]
            Names = {S.shape_id: S.name for S in Slide.shapes}
            Fx = [(Names[E["spid"]], E["class"], E["preset"], E["delay_ms"]) for E in kTimingReader.Read(Slide)]
            Cards = [D for N, _, P, D in Fx if N.startswith("Card") and P == 37]
            T.Check("animate: the title fades first, the cards float up 0.3 s apart (a card with its value and label)",
                    Fx and Fx[0][:3] == ("Title 1", "entr", 10) and sorted(set(Cards)) == [400, 700, 1000]
                    and ("Value2", "entr", 37, 700) in Fx and ("Label3", "entr", 37, 1000) in Fx, repr(Fx))
            Xml = etree.tostring(Slide._element).decode()
            T.Check("animate: one automatic main sequence, the first effect after the slide starts, every shape "
                    "built as one object", Xml.count('nodeType="mainSeq"') == 1 and Xml.count("afterEffect") == 1
                    and 'animBg="1"' in Xml and "onNext" in Xml, Xml[:300])
            T.Check("animate: float_up eases out, wipe from the top is wipe(up) eased in and out",
                    'decel="100000"' in Xml and 'filter="wipe(up)"' in Xml and 'accel="50000" decel="50000"' in Xml
                    and 'presetSubtype="1"' in Xml, "")
            Motion = Slide._element.find(f".//{{{NS_P}}}animMotion")
            End = kMotionPath.Parse(Motion.get("path"))[-1] if Motion is not None else (0, 0)
            T.Check("animate: the path is a curve ending 240 pt right of the shape's place, in slide fractions",
                    Motion is not None and " C " in Motion.get("path") and abs(End[0] - 240 / 1440) < 1e-4
                    and abs(End[1]) < 1e-4, Motion.get("path") if Motion is not None else "no animMotion")
            T.Check("animate: the schema lists the animate field",
                    '"animate"' in self.Test.RunScript("build_deck.py", "--print-schema")[1], "")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTestAnimate.CheckTiming")
            return

    def CheckErrors(self):
        """A target that matches nothing and a build over 5 s are reported; a path without points is refused."""
        if kS.ErrorMode:
            return
        try:
            T = self.Test
            Code, Out, _ = self.Build("animbad", self.Spec([{"target": "Nope*", "effect": "zoom"},
                                                           {"target": "Card1", "effect": "fade", "at": 5, "dur": 1}]))
            T.Check("animate: an unknown target names the slide's shapes; a 6 s build is too long",
                    Code != 0 and "matches no shape; the slide has: Title 1" in Out and "the build runs 6.0 s" in Out,
                    Out[-600:])
            Code, Out, _ = self.Build("animnopath", self.Spec([{"target": "Card1", "effect": "path"}]))
            T.Check("animate: a path step without points is refused", Code != 0 and "needs 'path' points" in Out,
                    Out[-400:])
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTestAnimate.CheckErrors")
            return

    def CheckOnion(self):
        """The onion sheet: a frame per beat, later shapes absent from earlier frames, the moved shape at its end."""
        if kS.ErrorMode:
            return
        try:
            T = self.Test
            Deck = os.path.join(self.Dir, "animate.pptx")
            if not os.path.exists(Deck):
                T.Check("onion: the animated deck exists", False, Deck)
                return
            Prs = Presentation(Deck)
            Beats = kBuildBeats(Prs.slides[1])
            T.Check("onion: five beats (title, three cards, path with wipe)", len(Beats.Beats) == 5,
                    repr([B["at"] for B in Beats.Beats]))
            Value1 = next(S for S in Prs.slides[1].shapes if S.name == "Value1")
            Left = Value1.left
            Beats.SetState(1, Prs.slide_width, Prs.slide_height)
            Left1 = [S.name for S in Prs.slides[1].shapes]
            T.Check("onion: after beat 1 only the title of the animated shapes is there",
                    "Title 1" in Left1 and "Card1" not in Left1 and "Value3" not in Left1, repr(Left1))
            Prs = Presentation(Deck)
            Beats = kBuildBeats(Prs.slides[1])
            Beats.SetState(5, Prs.slide_width, Prs.slide_height)
            Moved = next(S for S in Prs.slides[1].shapes if S.name == "Value1").left - Left
            T.Check("onion: after the last beat Value1 sits 240 pt right (the path's end)",
                    abs(Moved - 240 * 12700) < 12700, str(Moved))
            Sheets = kOnionSheet(Deck, kBoxRenderer()).Build(os.path.join(self.Dir, "onion"))
            T.Check("onion: one sheet for the one animated slide, frames over the onion frame",
                    len(Sheets) == 1 and Sheets[0].endswith("onion-s002.png")
                    and Image.open(Sheets[0]).height > Image.open(Sheets[0]).width * 0.6, repr(Sheets))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTestAnimate.CheckOnion")
            return

    def Run(self):
        if kS.ErrorMode:
            return
        try:
            self.Test.Say("animation steps and onion sheet")
            self.CheckTiming()
            self.CheckErrors()
            self.CheckOnion()
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSelfTestAnimate.Run")
            return
