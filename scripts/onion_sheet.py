"""Onion-skin sheet: every build step of an animated slide in one picture, to check the motion without playing it.

    uvx --with python-pptx --with pillow python scripts/onion_sheet.py deck.pptx [--slides 4,6] [--out folder]
    -> <out>/onion-s004.png per animated slide (default folder: <deck>-onion beside the deck)

Each sheet has a strip of frames - the slide before its build, then after every beat (the effects that start at
the same moment) - and one large onion frame: the finished slide over a faint ghost of where the moving shapes
start, each beat's shapes numbered in the order they arrive, and every motion path drawn with its end marked.
Reads the slide's native animation (_animate.kTimingReader), so it checks any deck, not only built ones.
Renders through LibreOffice (render_lo.py, any OS) or PowerPoint (render_slides.py, Windows).
"""
# /// script
# requires-python = ">=3.10"
# dependencies = ["python-pptx", "pillow>=10.1", "pywin32; sys_platform == 'win32'"]
# ///
import argparse
import os
import shutil
import sys
import tempfile

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation

from _animate import NS_P, kTimingReader
from kShared import ToolInputException, ToolReportableException, kRun, kS, kToolException

PRESET_NAMES = {1: "Appear", 10: "Fade", 22: "Wipe", 37: "Float Up", 53: "Zoom"}
PATH_INK, BADGE_INK, GHOST = (208, 138, 0), (123, 79, 191), 0.35  # amber paths, violet badges (no blue)
FRAME_W, PAD, CAPTION = 320, 12, 22


class kOnionRenderer:
    """Renders a deck to one PNG/JPEG per slide with whatever renderer this machine has."""

    def Render(self, Path, OutDir):
        """The renders of Path in slide order."""
        if kS.ErrorMode:
            return []
        try:
            from render_lo import kLoRenderer
            if kLoRenderer.Available():
                return kLoRenderer.Render(Path, OutDir, 1280)
            if sys.platform == "win32":
                from render_slides import kSlideRenderer  # PowerPoint through COM
                return kSlideRenderer.Render(Path, OutDir)
            raise ToolReportableException("no renderer: install LibreOffice (soffice on PATH or KPS_SOFFICE) or "
                                          "run on Windows with PowerPoint")
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kOnionRenderer.Render(file={Path})")
            return []


class kBuildBeats:
    """One slide's animation as beats: the effects that start at the same moment, in start order."""

    def __init__(self, Slide):
        try:
            self.Slide = Slide
            self.Shapes = {S.shape_id: S for S in Slide.shapes}
            self.Beats = []  # [{'at': ms, 'effects': [...]}]
            for E in kTimingReader.Read(Slide):
                if E["spid"] not in self.Shapes:
                    continue  # a target inside a group or gone: nothing to draw
                if not self.Beats or self.Beats[-1]["at"] != E["delay_ms"]:
                    self.Beats.append({"at": E["delay_ms"], "effects": []})
                self.Beats[-1]["effects"].append(E)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kBuildBeats.__init__")

    def EnterBeat(self, Spid):
        """The beat in which the shape enters (0-based), or -1 when it is there from the start."""
        if kS.ErrorMode:
            return -1
        try:
            for I, Beat in enumerate(self.Beats):
                if any(E["spid"] == Spid and E["class"] == "entr" for E in Beat["effects"]):
                    return I
            return -1
        except Exception as e:
            kS.GlobalErrorHandler(e, "kBuildBeats.EnterBeat")
            return -1

    def Label(self, Index):
        """A frame caption: 'start', or '1.2 s  Float Up x4' for beat Index-1."""
        if kS.ErrorMode:
            return ""
        try:
            if Index == 0:
                return "start"
            Beat = self.Beats[Index - 1]
            Names = {}
            for E in Beat["effects"]:
                Name = "Path" if E["class"] == "path" else PRESET_NAMES.get(E["preset"], f"preset {E['preset']}")
                Names[Name] = Names.get(Name, 0) + 1
            Parts = [N if C == 1 else f"{N} x{C}" for N, C in Names.items()]
            return f"{Index}. {Beat['at'] / 1000:.1f} s  " + ", ".join(Parts)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kBuildBeats.Label")
            return ""

    def SetState(self, Index, SlideWidth, SlideHeight):
        """Make the slide look as it does after Index beats: later entrances removed, paths of done beats moved
        to their end. Removes the timing (a state is a still)."""
        if kS.ErrorMode:
            return
        try:
            for Spid, S in self.Shapes.items():
                if self.EnterBeat(Spid) >= Index:
                    S._element.getparent().remove(S._element)
            for Beat in self.Beats[:Index]:
                for E in Beat["effects"]:
                    if E["class"] == "path" and E["path"] and self.EnterBeat(E["spid"]) < Index:
                        Fx, Fy = E["path"][-1]
                        S = self.Shapes[E["spid"]]
                        S.left, S.top = S.left + round(Fx * SlideWidth), S.top + round(Fy * SlideHeight)
            for Timing in self.Slide._element.findall(f"{{{NS_P}}}timing"):
                self.Slide._element.remove(Timing)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kBuildBeats.SetState")
            return


class kOnionSheet:
    """Builds the onion sheets of a deck."""

    def __init__(self, Deck, Renderer=None):
        try:
            if not os.path.isfile(Deck):
                raise ToolInputException(f"deck not found: {Deck}")
            self.Deck = Deck
            self.Renderer = Renderer or kOnionRenderer()
            self.Prs = Presentation(Deck)
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kOnionSheet.__init__")

    def Animated(self, Only=None):
        """{slide number: kBuildBeats} of the animated slides (only those in Only, when given)."""
        if kS.ErrorMode:
            return {}
        try:
            Out = {}
            for No, Slide in enumerate(self.Prs.slides, 1):
                if Only and No not in Only:
                    continue
                Beats = kBuildBeats(Slide)
                if Beats.Beats:
                    Out[No] = Beats
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kOnionSheet.Animated")
            return {}

    def StateDeck(self, Numbers, Index, Out):
        """A copy of the deck with only the slides in Numbers, each after min(Index, its beats) beats."""
        if kS.ErrorMode:
            return
        try:
            Prs = Presentation(self.Deck)
            Ids = Prs.slides._sldIdLst
            for No, SldId in reversed(list(enumerate(list(Ids), 1))):
                if No not in Numbers:
                    Prs.part.drop_rel(SldId.rId)
                    Ids.remove(SldId)
            for Slide in Prs.slides:
                Beats = kBuildBeats(Slide)
                Beats.SetState(min(Index, len(Beats.Beats)), Prs.slide_width, Prs.slide_height)
            Prs.save(Out)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kOnionSheet.StateDeck")
            return

    def Frames(self, Animated, Work):
        """{slide number: [frame image paths, state 0..beats]} - one state deck and one render per state."""
        if kS.ErrorMode:
            return {}
        try:
            Numbers = sorted(Animated)
            Frames = {No: [] for No in Numbers}
            for Index in range(max(len(B.Beats) for B in Animated.values()) + 1):
                Live = [No for No in Numbers if len(Animated[No].Beats) >= Index]
                Path = os.path.join(Work, f"state{Index:02d}.pptx")
                self.StateDeck(Live, Index, Path)
                Files = self.Renderer.Render(Path, os.path.join(Work, f"state{Index:02d}"))
                if len(Files) != len(Live):
                    raise ToolReportableException(f"state {Index}: {len(Files)} renders for {len(Live)} slides")
                for No, File in zip(Live, Files):
                    Frames[No].append(File)
            return Frames
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kOnionSheet.Frames")
            return {}

    @staticmethod
    def Font(Px):
        """Pillow's built-in font at Px pixels."""
        if kS.ErrorMode:
            return None
        try:
            return ImageFont.load_default(size=Px)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kOnionSheet.Font")
            return None

    def Onion(self, Beats, Frames):
        """The finished slide over a ghost of the start, beats numbered, motion paths drawn."""
        if kS.ErrorMode:
            return None
        try:
            Final = Image.open(Frames[-1]).convert("RGB")
            Moves = [I for I, B in enumerate(Beats.Beats) if any(E["class"] == "path" for E in B["effects"])]
            Before = Frames[Moves[0]] if Moves else Frames[0]  # the ghost: where the moving shapes start
            Img = Image.blend(Final, Image.open(Before).convert("RGB").resize(Final.size), GHOST)
            Img = Image.blend(Img, Final, 0.5)  # the result reads clearly, the ghost of the start shows through
            Draw = ImageDraw.Draw(Img)
            Sx, Sy = Final.width / self.Prs.slide_width, Final.height / self.Prs.slide_height
            Font, R = kOnionSheet.Font(max(14, Final.width // 60)), max(10, Final.width // 90)
            for I, Beat in enumerate(Beats.Beats, 1):
                Boxes = []
                for E in Beat["effects"]:
                    S = Beats.Shapes[E["spid"]]
                    Box = (S.left * Sx, S.top * Sy, (S.left + S.width) * Sx, (S.top + S.height) * Sy)
                    if E["class"] == "path" and E["path"]:
                        Cx, Cy = (Box[0] + Box[2]) / 2, (Box[1] + Box[3]) / 2
                        Pts = [(Cx, Cy)] + [(Cx + Fx * Final.width, Cy + Fy * Final.height) for Fx, Fy in E["path"]]
                        Draw.line(Pts, fill=PATH_INK, width=max(3, R // 3), joint="curve")
                        Ex, Ey = Pts[-1]
                        Draw.ellipse((Ex - R / 2, Ey - R / 2, Ex + R / 2, Ey + R / 2), fill=PATH_INK)
                    elif E["class"] == "entr":
                        Boxes.append(Box)
                if Boxes:
                    X, Y = min(B[0] for B in Boxes), min(B[1] for B in Boxes)
                    Draw.ellipse((X - R, Y - R, X + R, Y + R), fill=BADGE_INK)
                    Draw.text((X, Y), str(I), fill="white", font=Font, anchor="mm")
            return Img
        except Exception as e:
            kS.GlobalErrorHandler(e, "kOnionSheet.Onion")
            return None

    def Sheet(self, Beats, Frames, Out):
        """The strip of frames over the onion frame, saved to Out."""
        if kS.ErrorMode:
            return ""
        try:
            Onion = self.Onion(Beats, Frames)
            if Onion is None:
                return ""
            FrameH = round(FRAME_W * Onion.height / Onion.width)
            Cols = max(1, min(len(Frames), 4))
            Rows = -(-len(Frames) // Cols)
            Width = max(Cols * (FRAME_W + PAD) + PAD, Onion.width // 2 + 2 * PAD)
            Big = Onion.resize((Width - 2 * PAD, round((Width - 2 * PAD) * Onion.height / Onion.width)))
            Top = Rows * (FrameH + CAPTION + PAD) + PAD
            Img = Image.new("RGB", (Width, Top + Big.height + CAPTION + PAD), (245, 245, 240))
            Draw, Font = ImageDraw.Draw(Img), kOnionSheet.Font(14)
            for I, File in enumerate(Frames):
                X, Y = PAD + (I % Cols) * (FRAME_W + PAD), PAD + (I // Cols) * (FrameH + CAPTION + PAD)
                Img.paste(Image.open(File).convert("RGB").resize((FRAME_W, FrameH)), (X, Y))
                Draw.text((X, Y + FrameH + 3), Beats.Label(I), fill=(60, 60, 60), font=Font)
            Draw.text((PAD, Top), "onion: finished slide over the start; numbers = arrival order, amber = motion path",
                      fill=(60, 60, 60), font=Font)
            Img.paste(Big, (PAD, Top + CAPTION))
            Img.save(Out)
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kOnionSheet.Sheet")
            return ""

    def Build(self, OutDir, Only=None):
        """Write one sheet per animated slide into OutDir; returns their paths."""
        if kS.ErrorMode:
            return []
        try:
            Animated = self.Animated(Only)
            if not Animated:
                raise ToolInputException("no animated slide" + (f" among {sorted(Only)}" if Only else "") +
                                         f" in {self.Deck}")
            os.makedirs(OutDir, exist_ok=True)
            Work = tempfile.mkdtemp(prefix="onion-")
            Frames = self.Frames(Animated, Work)
            Out = [self.Sheet(Animated[No], Frames[No], os.path.join(OutDir, f"onion-s{No:03d}.png"))
                   for No in sorted(Frames)]
            shutil.rmtree(Work, ignore_errors=True)  # a leftover temp folder is harmless
            return [P for P in Out if P]
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kOnionSheet.Build")
            return []


class kOnionSheetApp:
    """Command line."""

    def Run(self):
        if kS.ErrorMode:
            return 1
        try:
            Parser = argparse.ArgumentParser(description=__doc__.split("\n")[0])
            Parser.add_argument("deck")
            Parser.add_argument("--slides", help="slide numbers, e.g. 4,6 (default: every animated slide)")
            Parser.add_argument("--out", help="folder (default <deck>-onion beside the deck)")
            Args = Parser.parse_args()
            Only = None
            if Args.slides:
                if not all(P.strip().isdigit() for P in Args.slides.split(",")):
                    raise ToolInputException(f"--slides takes numbers like 4,6, not {Args.slides!r}")
                Only = {int(P) for P in Args.slides.split(",")}
            Out = Args.out or os.path.splitext(Args.deck)[0] + "-onion"
            for Path in kOnionSheet(Args.deck).Build(Out, Only):
                print(Path)
            return 1 if kS.ErrorMode else 0
        except kToolException:
            raise
        except Exception as e:
            kS.GlobalErrorHandler(e, "kOnionSheetApp.Run")
            return 1


if __name__ == "__main__":
    kRun.Main(kOnionSheetApp)
