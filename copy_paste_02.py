import customtkinter as ctk
import pyperclip
import threading
import time
import os
import json
import random
import platform
import hashlib
import io
import queue
import tkinter as tk
import struct
import uuid
from datetime import datetime
from PIL import Image, ImageGrab, ImageTk
from pathlib import Path
from tkinterdnd2 import TkinterDnD, DND_FILES, DND_TEXT, COPY

# Auto-updater is optional — if requests isn't installed or the module
# is missing, the app still runs, it just can't check for updates.
try:
    import updater
    _UPDATER_AVAILABLE = True
except Exception as _e:
    _UPDATER_AVAILABLE = False

# In-app audio/video playback is optional — without ffpyplayer the Play button
# falls back to the OS default player (GIFs still play in-app via Pillow).
try:
    from ffpyplayer.player import MediaPlayer
    _PLAYER_AVAILABLE = True
except Exception:
    _PLAYER_AVAILABLE = False
    
# --------------------------------------------------------------------------
# Native Windows OLE drag-and-drop.
#
# tkinterdnd2's outbound drag (drag_source_register + <<DragInitCmd>>) fires
# fine, but its bundled tkdnd build frequently fails to actually hand the
# drag off to the OS when the drop target is a native (non-Tk) app such as
# Explorer or Outlook — the app-side visual shows, nothing lands. This talks
# directly to the same COM drag-drop machinery Explorer itself uses, so it
# works against any real OLE drop target. Only available on Windows with
# pywin32 installed; other platforms/setups fall back to tkinterdnd2.
# --------------------------------------------------------------------------
try:
    import pythoncom
    import win32clipboard
    import win32con
    import winerror
    from win32com.server.util import wrap as _com_wrap
    from win32com.server.exception import COMException
    _NATIVE_DND_AVAILABLE = (platform.system() == "Windows")
except ImportError:
    _NATIVE_DND_AVAILABLE = False


if _NATIVE_DND_AVAILABLE:

    # DROPEFFECT_* aren't reliably exposed as named attributes across
    # pywin32 builds (neither pythoncom nor win32con had them here) — these
    # are fixed DWORD values from oleidl.h, safe to hardcode directly.
    _DROPEFFECT_NONE = 0
    _DROPEFFECT_COPY = 1
    _DROPEFFECT_MOVE = 2
    _DROPEFFECT_LINK = 4

    class _NativeDropSource:
        """Minimal IDropSource — decides when the drag ends and what cursor
        to show. Required by DoDragDrop even though we want default behavior.
        Note: S_OK/DRAGDROP_S_* are NOT attributes of pythoncom in current
        pywin32 (verified against the pywin32 stubs) — they live in winerror."""

        _public_methods_ = ['QueryContinueDrag', 'GiveFeedback']
        _com_interfaces_ = [pythoncom.IID_IDropSource]

        def QueryContinueDrag(self, fEscapePressed, grfKeyState):
            if fEscapePressed:
                return winerror.DRAGDROP_S_CANCEL
            if not (grfKeyState & (win32con.MK_LBUTTON | win32con.MK_RBUTTON)):
                return winerror.DRAGDROP_S_DROP
            return winerror.S_OK

        def GiveFeedback(self, dwEffect):
            return winerror.DRAGDROP_S_USEDEFAULTCURSORS

    class _FormatEtcEnumerator:
        """Minimal IEnumFORMATETC. pythoncom has no WrapEnumFormatEtc helper
        (also verified absent from the stubs) — EnumFormatEtc needs a real,
        if tiny, enumerator object instead of a one-line wrapper call."""

        _public_methods_ = ['Next', 'Skip', 'Reset', 'Clone']
        _com_interfaces_ = [pythoncom.IID_IEnumFORMATETC]

        def __init__(self, formats):
            self._formats = list(formats)
            self._index = 0

        def Next(self, count):
            chunk = self._formats[self._index:self._index + count]
            self._index += len(chunk)
            return chunk

        def Skip(self, count):
            self._index = min(self._index + count, len(self._formats))
            return winerror.S_OK if self._index < len(self._formats) else winerror.S_FALSE

        def Reset(self):
            self._index = 0
            return winerror.S_OK

        def Clone(self):
            clone = _FormatEtcEnumerator(self._formats)
            clone._index = self._index
            return _com_wrap(clone, pythoncom.IID_IEnumFORMATETC)

    class _NativeFileDataObject:
        """Minimal IDataObject exposing CF_HDROP (files) and/or
        CF_UNICODETEXT (text) via a single HGLOBAL medium — the same
        DROPFILES-struct trick already used for clipboard file copies,
        just handed to DoDragDrop instead of the clipboard."""

        _public_methods_ = [
            'GetData', 'GetDataHere', 'QueryGetData', 'GetCanonicalFormatEtc',
            'SetData', 'EnumFormatEtc', 'DAdvise', 'DUnadvise', 'EnumDAdvise',
        ]
        _com_interfaces_ = [pythoncom.IID_IDataObject]

        def __init__(self, file_paths=None, text=None):
            self._file_paths = list(file_paths) if file_paths else None
            self._text = text

        def _formats(self):
            if self._file_paths:
                yield (win32con.CF_HDROP, None, pythoncom.DVASPECT_CONTENT,
                       -1, pythoncom.TYMED_HGLOBAL)
            if self._text is not None:
                yield (win32con.CF_UNICODETEXT, None, pythoncom.DVASPECT_CONTENT,
                       -1, pythoncom.TYMED_HGLOBAL)
                # Legacy ANSI text. Some native apps (Word among them) query
                # this before/instead of CF_UNICODETEXT and refuse the whole
                # drop on the first DV_E_FORMATETC rather than trying the
                # next format — unlike Chromium, which falls back gracefully.
                yield (win32con.CF_TEXT, None, pythoncom.DVASPECT_CONTENT,
                       -1, pythoncom.TYMED_HGLOBAL)

        def _hdrop_bytes(self):
            names = "\0".join(os.path.abspath(p) for p in self._file_paths) + "\0\0"
            block = names.encode("utf-16-le")
            header = struct.pack('Iiiii', struct.calcsize('Iiiii'), 0, 0, 0, 1)
            return header + block

        def _unicode_text_bytes(self):
            return (self._text + "\0").encode("utf-16-le")

        def _ansi_text_bytes(self):
            # Best-effort downgrade to the system codepage; characters
            # outside it become '?' rather than raising, since a lossy
            # CF_TEXT is still better than no fallback at all.
            return (self._text + "\0").encode("mbcs", errors="replace")

        def GetData(self, formatetc):
            cf, _ptd, _aspect, _index, tymed = formatetc
            if cf == win32con.CF_HDROP and self._file_paths and (tymed & pythoncom.TYMED_HGLOBAL):
                payload = self._hdrop_bytes()
            elif cf == win32con.CF_UNICODETEXT and self._text is not None and (tymed & pythoncom.TYMED_HGLOBAL):
                payload = self._unicode_text_bytes()
            elif cf == win32con.CF_TEXT and self._text is not None and (tymed & pythoncom.TYMED_HGLOBAL):
                payload = self._ansi_text_bytes()
            else:
                # Diagnostic: which format/tymed combo did the drop target
                # actually insist on that we don't support? cf is a numeric
                # clipboard format id — 15 is CF_HDROP, 13 is CF_UNICODETEXT,
                # anything else (often >0xC000) is a registered shell format
                # like "Preferred DropEffect" or "Shell IDList Array".
                raise COMException(hresult=winerror.DV_E_FORMATETC)
            medium = pythoncom.STGMEDIUM()
            medium.set(pythoncom.TYMED_HGLOBAL, payload)
            return medium

        def GetDataHere(self, formatetc, medium):
            raise COMException(hresult=winerror.E_NOTIMPL)

        def QueryGetData(self, formatetc):
            cf, _ptd, _aspect, _index, tymed = formatetc
            for fmt in self._formats():
                if fmt[0] == cf and (tymed & pythoncom.TYMED_HGLOBAL):
                    return 0
            return winerror.DV_E_FORMATETC

        def GetCanonicalFormatEtc(self, formatetc):
            raise COMException(hresult=winerror.E_NOTIMPL)

        def SetData(self, formatetc, medium, release):
            raise COMException(hresult=winerror.E_NOTIMPL)

        def EnumFormatEtc(self, direction):
            if direction != pythoncom.DATADIR_GET:
                raise COMException(hresult=winerror.E_NOTIMPL)
            return _com_wrap(
                _FormatEtcEnumerator(self._formats()),
                pythoncom.IID_IEnumFORMATETC
            )

        def DAdvise(self, formatetc, flags, sink):
            raise COMException(hresult=winerror.OLE_E_ADVISENOTSUPPORTED)

        def DUnadvise(self, connection):
            raise COMException(hresult=winerror.OLE_E_ADVISENOTSUPPORTED)

        def EnumDAdvise(self):
            raise COMException(hresult=winerror.OLE_E_ADVISENOTSUPPORTED)


class T:
    """Mission-control palette: near-black canvas, hairline borders, one
    white hero colour, and status colours used sparingly."""
    BG = "#050506"
    PANEL = "#0B0B0D"
    ROW = "#111114"
    ROW_HOVER = "#1A1A1F"
    LINE = "#25252B"
    LINE_STRONG = "#3B3B44"
    TEXT = "#F4F4F6"
    MUTED = "#8C8C96"
    DIM = "#55555E"
    GO = "#2BD576"
    WARN = "#FFB020"
    ABORT = "#FF453A"
    FONT_UI = "Bahnschrift"
    FONT_MONO = "Consolas"
    TOAST = {
        "green": GO, "red": ABORT, "orange": WARN,
        "blue": TEXT, "gray": MUTED, "purple": "#B794F6",
    }


def apply_spacex_theme():
    """Restyle customtkinter at the source (its theme dict) so every widget,
    dialogs included, inherits the look without per-widget colour arguments."""
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("dark-blue")  # loads a complete base dict to override

    def pair(color):
        return [color, color]

    overrides = {
        "CTk": {"fg_color": pair(T.BG)},
        "CTkToplevel": {"fg_color": pair(T.BG)},
        "CTkFrame": {
            "corner_radius": 3, "border_width": 0,
            "fg_color": pair(T.PANEL), "top_fg_color": pair(T.ROW),
            "border_color": pair(T.LINE),
        },
        "CTkButton": {
            "corner_radius": 2, "border_width": 1, "fg_color": "transparent",
            "hover_color": pair(T.ROW_HOVER), "border_color": pair(T.LINE_STRONG),
            "text_color": pair(T.TEXT), "text_color_disabled": pair(T.DIM),
        },
        "CTkLabel": {"text_color": pair(T.TEXT)},
        "CTkEntry": {
            "corner_radius": 2, "border_width": 1, "fg_color": pair(T.PANEL),
            "border_color": pair(T.LINE_STRONG), "text_color": pair(T.TEXT),
            "placeholder_text_color": pair(T.DIM),
        },
        "CTkCheckBox": {
            "corner_radius": 2, "border_width": 1, "fg_color": pair(T.TEXT),
            "hover_color": pair("#C9C9CF"), "checkmark_color": pair("#000000"),
            "border_color": pair(T.LINE_STRONG), "text_color": pair(T.TEXT),
        },
        "CTkSwitch": {
            "fg_color": pair(T.LINE_STRONG), "progress_color": pair(T.GO),
            "button_color": pair(T.TEXT), "button_hover_color": pair("#FFFFFF"),
            "text_color": pair(T.TEXT),
        },
        "CTkProgressBar": {
            "corner_radius": 0, "border_width": 0,
            "fg_color": pair(T.LINE), "progress_color": pair(T.TEXT),
        },
        "CTkScrollbar": {
            "fg_color": "transparent", "button_color": pair(T.LINE_STRONG),
            "button_hover_color": pair(T.MUTED),
        },
    }
    theme = ctk.ThemeManager.theme
    for widget, values in overrides.items():
        theme.setdefault(widget, {}).update(values)

    font_cfg = theme.setdefault("CTkFont", {})
    for cfg in [font_cfg, *[v for v in font_cfg.values() if isinstance(v, dict)]]:
        cfg["family"] = T.FONT_UI


class MediaPlayerPanel(ctk.CTkFrame):
    """In-app player for audio, video and GIFs, shown in the preview section or
    fullscreen. Audio/video decode through ffpyplayer; GIFs animate through
    Pillow. Playback always starts muted."""

    STAGE_HEIGHT = 190
    SEEK_STEP = 5.0
    VOLUME_BOOST_COLOR = "#E8A23A"  # slider past 100%

    def __init__(self, parent, app, path, start_pos=0.0, paused=False, muted=True, fullscreen=False):
        super().__init__(
            parent,
            fg_color="#000000" if fullscreen else T.PANEL,
            border_width=0 if fullscreen else 1,
            border_color=T.LINE,
            corner_radius=0 if fullscreen else 3
        )
        self.app = app
        self.path = path
        self.fullscreen = fullscreen
        self.paused = paused
        self.muted = muted

        ext = os.path.splitext(path)[1].lower()
        if ext == '.gif':
            self.kind = 'gif'
        elif ext in app.AUDIO_EXTENSIONS:
            self.kind = 'audio'
        else:
            self.kind = 'video'

        self.player = None
        self.ended = False
        self.duration = 0.0
        self.position = 0.0
        self._meta_done = False
        self._video_size = None
        self._last_stage = (0, 0)
        self._last_progress = 0.0
        self._seeking = False
        self._photo = None
        self._job = None
        self._alive = True
        self._gif = None
        self._gif_index = 0
        self._gif_count = 1
        self._gif_dirty = True
        self._ready_at = 0.0          # earliest time set_volume/get_pts are safe
        self._volume_pending = False  # a volume change waiting for that moment
        self._unpause_pending = False # a play request waiting for that moment


        self._build()
        self._start(start_pos)
        self._job = self.after(30, self._tick)

    # ------------------------------------------------------------------
    # Construction
    # ------------------------------------------------------------------

    def _build(self):
        app = self.app
        edge = 0 if self.fullscreen else 8
        name = os.path.basename(self.path)

        if not self.fullscreen:
            header = ctk.CTkFrame(self, fg_color="transparent")
            header.pack(side="top", fill="x", padx=edge, pady=(6, 2))
            short = name if len(name) <= 30 else name[:27] + "..."
            ctk.CTkLabel(
                header, text="PLAYER  ·  " + short, anchor="w",
                font=app._mono(10), text_color=T.MUTED
            ).pack(side="left")
            ctk.CTkButton(
                header, text="✕", width=24, height=22, border_width=0,
                font=ctk.CTkFont(size=11, weight="bold"), text_color=T.MUTED,
                hover_color="#3A1210", command=app.close_media_player
            ).pack(side="right")

        # Bottom-anchored rows are packed first so the stage can take the rest
        controls = ctk.CTkFrame(self, fg_color="transparent")
        controls.pack(side="bottom", fill="x", padx=10, pady=(0, 8))

        self.play_btn = app._btn(controls, "", self.toggle_play, kind="primary", width=44, height=26)
        self.play_btn.pack(side="left")

        self.time_label = ctk.CTkLabel(
            controls, text="00:00 / 00:00", font=app._mono(10), text_color=T.MUTED
        )
        self.time_label.pack(side="left", padx=10)

        self.fs_btn = app._btn(
            controls, "EXIT" if self.fullscreen else "FULL",
            self._on_fullscreen_click, width=64 if self.fullscreen else 56, height=26
        )
        self.fs_btn.pack(side="right")

        self.mute_btn = None
        self.vol_slider = None
        self.vol_label = None
        if self.kind != 'gif':
            # Packed right-to-left: label, slider, then the speaker, so the row
            # reads  [speaker][slider][NN%][fullscreen]
            self.vol_label = ctk.CTkLabel(
                controls, text="", width=38, anchor="e",
                font=app._mono(10), text_color=T.MUTED
            )
            self.vol_label.pack(side="right", padx=(2, 6))
            self.vol_slider = ctk.CTkSlider(
                controls, from_=0, to=200, number_of_steps=200,
                width=140 if self.fullscreen else 72, height=14,
                fg_color=T.LINE_STRONG, progress_color=T.TEXT,
                button_color=T.TEXT, button_hover_color="#FFFFFF",
                command=self._on_volume
            )
            self.vol_slider.set(app.player_volume * 100)
            self.vol_slider.pack(side="right")
            self.mute_btn = ctk.CTkButton(
                controls, text="", width=36, height=26, border_width=1,
                font=ctk.CTkFont(size=14), command=self.toggle_mute
            )
            self.mute_btn.pack(side="right", padx=(0, 6))

        self.seek = ctk.CTkSlider(
            self, from_=0, to=1, height=14,
            fg_color=T.LINE_STRONG, progress_color=T.TEXT,
            button_color=T.TEXT, button_hover_color="#FFFFFF",
            command=self._on_seek_drag
        )
        self.seek.set(0)
        self.seek.pack(side="bottom", fill="x", padx=10, pady=(6, 2))
        self.seek.bind("<ButtonRelease-1>", self._on_seek_release, add="+")

        self.stage = tk.Frame(
            self, bg="#000000", bd=0, highlightthickness=0, height=self.STAGE_HEIGHT
        )
        if self.fullscreen:
            self.stage.pack(side="top", fill="both", expand=True)
        else:
            self.stage.pack(side="top", fill="x", padx=edge, pady=(2, 0))
        self.stage.pack_propagate(False)

        self.view = tk.Label(
            self.stage, bg="#000000", fg=T.MUTED, bd=0,
            font=(T.FONT_MONO, 11, "bold"), justify="center"
        )
        self.view.place(relx=0.5, rely=0.5, anchor="center")

        for widget in (self.stage, self.view):
            widget.bind("<Button-1>", lambda e: self.toggle_play(), add="+")
            widget.bind("<Double-Button-1>", lambda e: self._on_fullscreen_click(), add="+")

        if self.fullscreen:
            top = self.winfo_toplevel()
            top.bind("<space>", lambda e: self.toggle_play(), add="+")
            top.bind("<m>", lambda e: self.toggle_mute(), add="+")
            top.bind("<M>", lambda e: self.toggle_mute(), add="+")
            top.bind("<Left>", lambda e: self._seek_relative(-self.SEEK_STEP), add="+")
            top.bind("<Right>", lambda e: self._seek_relative(self.SEEK_STEP), add="+")

        self._sync_mute_button()
        self._sync_play_button()

    def _start(self, start_pos=0.0):
        """(Re)open the media. Audio/video start paused and muted-as-requested
        BEFORE playback resumes, so nothing is audible until the user unmutes."""
        self.ended = False
        self._meta_done = False
        self._last_stage = (0, 0)
        name = os.path.basename(self.path)
        short = name if len(name) <= 34 else name[:31] + "..."

        if self.kind == 'gif':
            try:
                self._gif = Image.open(self.path)
                self._gif_count = max(1, getattr(self._gif, "n_frames", 1))
                self.duration = float(self._gif_count - 1)
                self._gif_index = max(0, min(self._gif_count - 1, int(start_pos or 0)))
                self.position = float(self._gif_index)
                self._gif_dirty = True
                self.view.configure(text="LOADING")
            except Exception as e:
                self._fail(e)
            self._sync_play_button()
            self._refresh_progress(force=True)
            return

        self.view.configure(
            image="", text=("♪   AUDIO\n\n" + short) if self.kind == 'audio' else "LOADING"
        )
        try:
            # loop=1: play once. Left unset, ffpyplayer may seek back to the start
            # at EOF on its own, which is what triggers the decoder's "mmco" warning
            opts = {'paused': True, 'loop': 1}
            if start_pos and start_pos > 0.5:
                opts['ss'] = float(start_pos)
            # Fixed 2x gain in the FFmpeg chain. set_volume() only reaches 1.0, so
            # the slider's 0-200% maps onto set_volume's 0-1 (see _apply_audio_state).
            # Float keeps the x2 headroom until the final 16-bit output.
            opts['af'] = "aformat=sample_fmts=flt,volume=2.0"
            self.player = MediaPlayer(self.path, ff_opts=opts)
            # set_volume() and get_pts() crash natively if called while the
            # streams are still opening, so they wait until this time
            self._ready_at = time.monotonic() + 1.2
            self._apply_audio_state()
            # Stay paused until mute/volume are applied; _tick_media unpauses
            self._unpause_pending = not self.paused
            self.position = float(start_pos or 0.0)
        except Exception as e:
            self._fail(e)
        self._sync_play_button()
        self._refresh_progress(force=True)

    def _fail(self, error):
        print(f"Player failed: {error}")
        self._release_player()
        self._photo = None
        self.view.configure(image="", text="PLAYBACK ERROR\n" + str(error)[:80])

    def _release_player(self):
        player, self.player = self.player, None
        if player is not None:
            try:
                player.close_player()
            except Exception:
                pass
            if _NATIVE_DND_AVAILABLE:
                # SDL's audio shutdown can call CoUninitialize() once more than it
                # initialised on this OLE-initialised thread. Re-add a count so
                # COM/OLE (drag-and-drop included) survives and the next player
                # can open.
                try:
                    pythoncom.CoInitialize()
                except Exception:
                    pass

    def _restart(self, start_pos=0.0):
        player = self.player
        if player is not None and self.kind != 'gif':
            try:
                # Cheap path: reuse the open player instead of closing and
                # re-probing the file on the UI thread
                player.seek(start_pos, relative=False, accurate=False)
                player.set_pause(False)
                self.ended = False
                self.paused = False
                self.position = start_pos
                self._apply_audio_state()
                self._sync_play_button()
                self._refresh_progress(force=True)
                return
            except Exception:
                pass
        self._release_player()
        self.paused = False
        self._start(start_pos)

    # ------------------------------------------------------------------
    # Controls
    # ------------------------------------------------------------------

    def toggle_play(self):
        if self.ended:
            self._restart(0.0)
            return
        self.paused = not self.paused
        if self.player is not None:
            if self._player_ready():
                self._unpause_pending = False
                try:
                    self.player.set_pause(self.paused)
                except Exception:
                    pass
            else:
                self._unpause_pending = True  # applied from _tick_media once ready
        self._sync_play_button()

    def toggle_mute(self):
        if self.kind == 'gif':
            return
        self.muted = not self.muted
        if not self.muted and self.app.player_volume <= 0:
            # Unmuting at 0% would still be silent, so land on a sensible level
            self.app.player_volume = 0.5
            if self.vol_slider is not None:
                self.vol_slider.set(50)
        self._apply_audio_state()
        self._sync_mute_button()

    def _on_volume(self, value):
        """Fires continuously while the slider glides, so the % label and the
        actual loudness track the thumb live."""
        self.app.player_volume = max(0.0, min(2.0, value / 100.0))
        boosted = self.app.player_volume > 1.0
        if boosted and not self.app._boost_active:
            self.app.show_toast(
                "Above 100% is software boost: it can distort & may damage "
                "speakers.", "orange",
                duration=2500
            )
        self.app._boost_active = boosted
        if self.muted and self.app.player_volume > 0:
            self.muted = False  # moving the slider is an explicit unmute
        self._apply_audio_state()
        self._sync_mute_button()


    def _apply_audio_state(self):
        """Mute via both the mute flag and zero volume. The audio device opens
        asynchronously after the player is created, so this is re-applied once
        the stream is up (see _tick_media)."""
        player = self.player
        if player is None:
            return
        try:
            player.set_mute(self.muted)
        except Exception:
            pass
        if not self._player_ready():
            self._volume_pending = True  # applied from _tick_media once safe
            return
        try:
            player.set_volume(0.0 if self.muted else self.app.player_volume / 2.0)
            self._volume_pending = False
        except Exception:
            pass

    def _player_ready(self):
        return self.player is not None and time.monotonic() >= self._ready_at

    def _sync_play_button(self):
        if self.ended:
            text = "↻"
        else:
            text = "▶" if self.paused else "❚❚"
        self.play_btn.configure(text=text)

    def _sync_mute_button(self):
        if self.mute_btn is None:
            return
        silent = self.muted or self.app.player_volume <= 0
        self.mute_btn.configure(
            text="🔇" if silent else "🔊",
            text_color=T.WARN if self.muted else T.TEXT,
            border_color=T.WARN if self.muted else T.LINE_STRONG
        )
        if self.vol_slider is not None:
            boosted = self.app.player_volume > 1.0
            self.vol_slider.configure(
                progress_color=self.VOLUME_BOOST_COLOR if boosted else T.TEXT,
                button_color=self.VOLUME_BOOST_COLOR if boosted else T.TEXT
            )
        if self.vol_label is not None:
            self.vol_label.configure(
                text=f"{int(round(self.app.player_volume * 100))}%",
                text_color=T.WARN if self.muted else T.MUTED
            )

    def _on_fullscreen_click(self):
        self.app.toggle_player_fullscreen(self)

    # ------------------------------------------------------------------
    # Seeking and progress
    # ------------------------------------------------------------------

    @staticmethod
    def _fmt(seconds):
        seconds = max(0, int(seconds))
        hours, rem = divmod(seconds, 3600)
        minutes, secs = divmod(rem, 60)
        return f"{hours}:{minutes:02d}:{secs:02d}" if hours else f"{minutes:02d}:{secs:02d}"

    def _time_text(self, pos):
        if self.kind == 'gif':
            return f"FRAME {int(round(pos)) + 1}/{self._gif_count}"
        return f"{self._fmt(pos)} / {self._fmt(self.duration)}"

    def _refresh_progress(self, force=False):
        now = time.monotonic()
        if not force and now - self._last_progress < 0.25:
            return
        self._last_progress = now
        self.time_label.configure(text=self._time_text(self.position))
        if not self._seeking:
            self.seek.set(min(1.0, self.position / self.duration) if self.duration > 0 else 0)

    def _on_seek_drag(self, value):
        self._seeking = True
        if self.duration > 0:
            self.time_label.configure(text=self._time_text(value * self.duration))

    def _on_seek_release(self, event=None):
        if not self._seeking:
            return
        self._seeking = False
        if self.duration > 0:
            self._seek_to(self.seek.get() * self.duration)

    def _seek_relative(self, delta):
        if self.kind == 'gif' or self.duration <= 0:
            return
        self._seek_to(self.position + delta)

    def _seek_to(self, target):
        target = max(0.0, min(self.duration, target)) if self.duration > 0 else max(0.0, target)
        if self.kind == 'gif':
            self._gif_index = int(round(target))
            self.position = float(self._gif_index)
            self._gif_dirty = True
        elif self.ended:
            self._restart(target)
            return
        elif self.player is not None:
            try:
                self.player.seek(target, relative=False, accurate=False)
            except Exception:
                pass
            self.position = target
        self._refresh_progress(force=True)

    # ------------------------------------------------------------------
    # Playback loop
    # ------------------------------------------------------------------

    def _tick(self):
        self._job = None
        if not self._alive:
            return
        try:
            delay = self._tick_gif() if self.kind == 'gif' else self._tick_media()
        except Exception as e:
            print(f"Player tick failed: {e}")
            delay = 100
        if self._alive:
            self._job = self.after(int(delay), self._tick)

    def _tick_media(self):
        player = self.player
        if player is None or self.ended:
            return 200

        frame, val = player.get_frame()
        if val == 'eof':
            self._on_eof()
            return 200

        if self._player_ready():
            if self._volume_pending:
                self._apply_audio_state()
            if self._unpause_pending:
                self._unpause_pending = False
                player.set_pause(self.paused)

        if not self._meta_done:
            meta = player.get_metadata() or {}
            duration = meta.get('duration')
            if duration:
                self.duration = float(duration)
            size = meta.get('src_vid_size')
            if size and size[0] and size[1]:
                self._video_size = (int(size[0]), int(size[1]))
            self._meta_done = bool(self.duration) and (
                self.kind == 'audio' or self._video_size is not None
            )

        # Let FFmpeg scale frames to the stage so Python never resizes video
        if self.kind == 'video' and self._video_size:
            sw, sh = self.stage.winfo_width(), self.stage.winfo_height()
            if sw > 40 and sh > 40 and (
                abs(sw - self._last_stage[0]) > 2 or abs(sh - self._last_stage[1]) > 2
            ):
                self._last_stage = (sw, sh)
                vw, vh = self._video_size
                scale = min(sw / vw, sh / vh)
                tw = max(2, int(vw * scale))
                th = max(2, int(vh * scale))
                player.set_size(tw - tw % 2, th - th % 2)

        if frame is not None and self.kind == 'video':
            img = frame[0]
            width, height = img.get_size()
            self._show(Image.frombytes('RGB', (width, height), bytes(img.to_bytearray()[0])))

        if not self._seeking and self._player_ready():
            pts = player.get_pts()
            if pts is not None and pts >= 0:
                self.position = float(pts)
            self._refresh_progress()

        if self.kind == 'audio':
            return 100
        if isinstance(val, (int, float)):
            return max(5, min(100, int(val * 1000)))
        return 50 if val == 'paused' else 30

    def _tick_gif(self):
        if self._gif is None or self._seeking:
            return 50
        if self.paused and not self._gif_dirty:
            return 50

        index = self._gif_index
        self._gif.seek(index)
        duration = self._gif.info.get('duration', 100)
        self._show(self._fit(self._gif.convert('RGBA')))
        self._gif_dirty = False
        self.position = float(index)
        self._refresh_progress()

        if not self.paused:
            self._gif_index = (index + 1) % self._gif_count
        # Browsers treat 0-10 ms GIF delays as 100 ms; match that
        return max(20, duration) if duration > 10 else 100

    def _on_eof(self):
        self.ended = True
        self.paused = True
        if self.player is not None:
            try:
                self.player.set_pause(True)
            except Exception:
                pass
        self.position = self.duration
        self._refresh_progress(force=True)
        self._sync_play_button()

    def _fit(self, pil):
        sw, sh = self.stage.winfo_width(), self.stage.winfo_height()
        if sw < 40 or sh < 40:
            return pil
        width, height = pil.size
        scale = min(sw / width, sh / height)
        return pil.resize((max(1, int(width * scale)), max(1, int(height * scale))), Image.BILINEAR)

    def _show(self, pil):
        self._photo = ImageTk.PhotoImage(pil)
        self.view.configure(image=self._photo, text="")

    # ------------------------------------------------------------------
    # State hand-off and teardown
    # ------------------------------------------------------------------

    def snapshot(self):
        """State needed to resume this file in another panel (fullscreen <-> inline)."""
        if self.ended:
            return {"path": self.path, "position": 0.0, "paused": True, "muted": self.muted}
        return {"path": self.path, "position": self.position, "paused": self.paused, "muted": self.muted}

    def destroy(self):
        self._alive = False
        if self._job is not None:
            try:
                self.after_cancel(self._job)
            except Exception:
                pass
            self._job = None
        self._release_player()
        if self._gif is not None:
            try:
                self._gif.close()
            except Exception:
                pass
            self._gif = None
        super().destroy()


class DnDCTk(TkinterDnD.DnDWrapper, ctk.CTk):
    """customtkinter's CTk doesn't inherit from TkinterDnD's Tk, so drag-and-drop
    is bolted on via this mixin — the standard pattern for combining
    customtkinter with tkinterdnd2. Still used for the (rare) non-Windows /
    no-pywin32 fallback path."""

    def __init__(self, *args, **kwargs):
        ctk.CTk.__init__(self, *args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)


class ClipboardManager:
    ROULETTE_MESSAGES = [
        "🎲 The clipboard gods have chosen item #{n}!",
        "🎲 Rolling the dice... item #{n} it is!",
        "🎲 Fate has selected item #{n}!",
        "🎲 Behold, a blast from the past — item #{n}!",
        "🎲 Random.org would be proud: #{n}!",
    ]

    # Bitmaps larger than this (~8K) are skipped rather than encoded
    MAX_IMAGE_PIXELS = 40_000_000

    PLAYABLE_EXTENSIONS = {
        '.mp3', '.wav', '.ogg', '.flac', '.m4a', '.aac', '.opus', '.wma',                  # audio
        '.mp4', '.avi', '.mkv', '.mov', '.wmv', '.webm', '.m4v', '.mpg', '.mpeg', '.3gp',  # video
        '.gif',                                                                            # animated image
    }

    # Subset of the above that has no picture — shown with an audio placeholder
    AUDIO_EXTENSIONS = {'.mp3', '.wav', '.ogg', '.flac', '.m4a', '.aac', '.opus', '.wma'}

    def __init__(self):
        # Configure appearance (the theme must be applied before any widget exists)
        apply_spacex_theme()

        # Main window
        self.window = DnDCTk()
        self.window.title("Clipboard Manager Pro")
        self.window.geometry("800x700")
        self.window.minsize(780, 640)
        self.window.protocol("WM_DELETE_WINDOW", self.on_close)

        # DoDragDrop needs OLE (not just COM) initialized on the UI thread.
        # Harmless if tkdnd already did it.
        if _NATIVE_DND_AVAILABLE:
            try:
                pythoncom.OleInitialize()
            except Exception:
                pass

        # Maximum items to store
        self.max_items = 24
        # Oldest -> newest. Record: id, type ('text'|'file'), content, pinned, ts
        self.items = []
        self.last_clipboard_content = ""

        # All-time fun stats (survive clears/restarts)
        self.stats = {"total_items": 0, "total_characters": 0}

        # Search
        self.search_query = ""

        # Persistence
        self.history_file = Path.home() / ".clipboard_manager_history.json"
        self.image_dir = Path.home() / ".clipboard_manager_images"

        # Selection is a set of stable item IDs; the dicts below only map the
        # currently *visible* rows' IDs to their widgets
        self.selected_ids = set()
        self.row_vars = {}        # id -> BooleanVar
        self.checkboxes = {}      # id -> checkbox widget
        self.row_frames = {}      # id -> row frame

        # Selection mode (True = multiple, False = single)
        self.multiple_selection_mode = False

        # Whether clipboard history is wiped on app close (True) or
        # persisted for the next launch (False) — itself persisted, so
        # the choice survives restarts
        self.clear_on_close = False

        # Toast stacking
        self.active_toasts = []

        # Warn about missing pywin32 only once per session
        self._pywin32_warning_shown = False

        # Hand-off from the monitor thread to the Tk main thread
        self._ui_queue = queue.Queue()

        # In-app media player: inline panel in the preview section, or a
        # fullscreen window (only one of the two exists at a time)
        self._player = None
        self._fs_window = None
        self._fs_panel = None
        self.player_volume = 1.0   # 0.0–2.0 (0–200%), remembered across files and fullscreen
        self._boost_active = False # True while the slider sits above 100%
        self._open_job = None      # pending deferred player build (see _defer_player)

        # Track scheduled marquee `after()` callbacks so we can cancel them
        # when the preview panel refreshes.
        self._marquee_jobs = []

        # Drag-and-drop visual feedback
        self._drag_visual = None
        self._drag_visual_job = None

        # Distinguish a genuine click (select the row) from the start of a
        # native drag (don't also toggle selection). Tracked per press since
        # only one press-drag sequence is ever in flight at a time.
        self._press_start = None       # (x_root, y_root) at ButtonPress
        self._drag_started = False     # set True once a real drag is confirmed
        self.CLICK_MOVE_THRESHOLD = 4  # pixels of movement before it counts as a drag

        # Payload for whichever row is currently pressed, consumed by
        # _on_item_motion once the pointer crosses CLICK_MOVE_THRESHOLD.
        self._pending_drag_paths = None
        self._pending_drag_text = None

        # Timestamp until which the monitor thread ignores clipboard changes —
        # prevents our own programmatic copies from being re-detected as new items
        self._suppress_monitor_until = 0.0

        # Load persisted history before building UI so it renders immediately
        self.load_history()

        # Create UI
        self.create_widgets()
        self.update_display()

        # Snapshot whatever's currently on the OS clipboard as the known
        # baseline, using the live clipboard rather than a guess from the
        # last saved history item. Guessing from history is what caused the
        # restart-duplication bug: if the last saved item was a file
        # selection, the old code fell back to last_clipboard_content = "",
        # so the very next poll saw the still-present file selection as
        # "new" and re-added it one slot after the last item, every launch.
        self.last_clipboard_content = self._read_current_clipboard_state()

        # Kick off a silent update check on a background thread. This is
        # a no-op if updater.py or requests isn't available — see the
        # try/except import at the top of the file. Silent mode only
        # shows a dialog if an update is actually found.
        if _UPDATER_AVAILABLE:
            try:
                updater.check_for_updates_async(self.window, silent=True)
            except Exception as e:
                print(f"Update check failed to start: {e}")

        # Start clipboard monitoring
        self.monitoring = True
        self.monitor_thread = threading.Thread(
            target=self.monitor_clipboard,
            daemon=True
        )
        self.monitor_thread.start()
        self._pump_ui_queue()

    # ------------------------------------------------------------------
    # Drag-and-drop visual feedback
    # ------------------------------------------------------------------

    def show_drag_visual(self, paths, event=None, is_text=False):
        """Show a visual indicator containing the filename(s) — or a text
        preview — being dragged."""

        self.hide_drag_visual()

        if not paths:
            return

        paths = list(paths)

        if is_text:

            message = (
                "📝 DRAGGING TEXT\n"
                f"{paths[0]}"
            )

        elif len(paths) == 1:

            filename = os.path.basename(paths[0])

            message = (
                "📄 DRAGGING FILE\n"
                f"{filename}"
            )

        else:

            filenames = [
                os.path.basename(path)
                for path in paths[:3]
            ]

            message = (
                f"📁 DRAGGING {len(paths)} FILES\n"
                + "\n".join(f"• {name}" for name in filenames)
            )

            if len(paths) > 3:
                message += (
                    f"\n• ... and {len(paths) - 3} more"
                )

        self._drag_visual = ctk.CTkLabel(
            self.window,
            text=message,
            font=ctk.CTkFont(
                size=12,
                weight="bold"
            ),
            fg_color=("gray85", "gray20"),
            text_color=("black", "white"),
            corner_radius=10,
            justify="left",
            padx=12,
            pady=8
        )

        try:

            if event is not None:

                x = (
                    event.x_root
                    - self.window.winfo_rootx()
                    + 20
                )

                y = (
                    event.y_root
                    - self.window.winfo_rooty()
                    + 20
                )

            else:

                x = (
                    self.window.winfo_pointerx()
                    - self.window.winfo_rootx()
                    + 20
                )

                y = (
                    self.window.winfo_pointery()
                    - self.window.winfo_rooty()
                    + 20
                )

        except Exception:

            x = 20
            y = 20

        self._drag_visual.place(
            x=x,
            y=y
        )

        self._drag_visual.lift()
        
    def hide_drag_visual(self):
        """Remove the drag visual indicator."""

        if self._drag_visual_job is not None:
            try:
                self.window.after_cancel(self._drag_visual_job)
            except Exception:
                pass

            self._drag_visual_job = None

        if self._drag_visual is not None:
            try:
                if self._drag_visual.winfo_exists():
                    self._drag_visual.destroy()
            except Exception:
                pass

        self._drag_visual = None

    def schedule_drag_visual_hide(self, delay=700):
        """Schedule removal of the drag visual."""

        if self._drag_visual_job is not None:
            try:
                self.window.after_cancel(self._drag_visual_job)
            except Exception:
                pass

        self._drag_visual_job = self.window.after(
            delay,
            self.hide_drag_visual
        )

    def _bind_file_drag_source(self, widget, file_paths):
        """Register a widget as a drag source for the given files. On
        Windows with pywin32 available this arms the native OLE path
        (see _start_native_drag); the payload itself is stashed on
        ButtonPress by _on_item_press and consumed by _on_item_motion —
        this method only needs to exist for the tkdnd fallback branch."""

        paths = (
            list(file_paths)
            if isinstance(file_paths, (list, tuple))
            else [file_paths]
        )

        valid_paths = [
            os.path.abspath(str(path))
            for path in paths
            if isinstance(path, str) and os.path.exists(path)
        ]

        if not valid_paths:
            return

        if _NATIVE_DND_AVAILABLE:
            return  # handled by _on_item_motion via native OLE drag

        try:
            widget.drag_source_register(1, DND_FILES)

            def drag_init(event, paths=valid_paths):
                self._drag_started = True
                self.show_drag_visual(paths, event)
                data = self.window.tk.call("list", *paths)
                return (COPY, DND_FILES, data)

            widget.dnd_bind("<<DragInitCmd>>", drag_init)

        except Exception as e:
            print(f"File drag source setup failed: {e}")

    def _bind_text_drag_source(self, widget, text):
        """Bind a widget as a text drag source. On Windows with pywin32
        available this arms the native OLE path (see _start_native_drag);
        otherwise falls back to tkdnd's <<DragInitCmd>>."""

        if not text:
            return

        if _NATIVE_DND_AVAILABLE:
            return  # handled by _on_item_motion via native OLE drag

        preview = text if len(text) <= 60 else text[:60] + "..."

        try:
            widget.drag_source_register(1, DND_TEXT)

            def drag_init(event, text=text, preview=preview):
                self._drag_started = True
                self.show_drag_visual([preview], event, is_text=True)
                return ((COPY,), (DND_TEXT,), text)

            widget.dnd_bind("<<DragInitCmd>>", drag_init)

        except Exception as e:
            print(f"Could not bind text drag source: {e}")

    def _start_native_drag(self, file_paths=None, text=None):
        """Kick off a real Windows OLE drag-and-drop via pywin32. This is a
        blocking, modal call — DoDragDrop runs its own message loop and
        only returns once the drop (or cancel) completes, exactly like
        Explorer's own drag does. Returns True if the native drag actually
        ran (regardless of whether the user dropped or cancelled)."""

        if not _NATIVE_DND_AVAILABLE:
            return False

        # A native "replace/skip" conflict dialog (or any other Explorer
        # window) is an ordinary, non-topmost window. If this app is pinned
        # always-on-top it renders ON TOP of that dialog for the entire
        # blocking DoDragDrop call, making it unreachable. Drop topmost for
        # the duration of the drag and restore whatever the user had it set
        # to afterward, regardless of how the drag ends.
        was_topmost = bool(self.topmost_var.get())
        if was_topmost:
            self.window.attributes('-topmost', False)

        try:
            pythoncom.CoInitialize()
        except pythoncom.com_error:
            pass  # already initialized on this thread — fine

        try:
            data_obj = _com_wrap(
                _NativeFileDataObject(file_paths, text),
                pythoncom.IID_IDataObject
            )
            drop_source = _com_wrap(
                _NativeDropSource(),
                pythoncom.IID_IDropSource
            )
            effect = _DROPEFFECT_COPY
            pythoncom.DoDragDrop(data_obj, drop_source, effect)
            return True
        except Exception as e:
            print(f"Native drag failed: {e}")
            return False
        finally:
            if was_topmost:
                self.window.attributes('-topmost', True)

    def _on_item_motion(self, event):
        """<B1-Motion> on a row. Once the pointer crosses the click
        threshold while a drag-able item is pressed, this fires the native
        OLE drag exactly once per press. Guarded by self._drag_started so
        repeated motion events (there are many) don't re-enter DoDragDrop."""

        if self._drag_started or self._press_start is None:
            return

        if self._pending_drag_paths is None and self._pending_drag_text is None:
            return

        dx = event.x_root - self._press_start[0]
        dy = event.y_root - self._press_start[1]
        if (dx * dx + dy * dy) <= (self.CLICK_MOVE_THRESHOLD ** 2):
            return

        self._drag_started = True
        self.hide_drag_visual()

        paths, text = self._pending_drag_paths, self._pending_drag_text
        self._pending_drag_paths = None
        self._pending_drag_text = None

        self._start_native_drag(file_paths=paths, text=text)

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    @staticmethod
    def _track(text):
        """Letter-spaced uppercase, the wordmark look."""
        return "   ".join(" ".join(word) for word in text.upper().split())

    @staticmethod
    def _mono(size=11, bold=False):
        return ctk.CTkFont(
            family=T.FONT_MONO, size=size, weight="bold" if bold else "normal"
        )

    def _btn(self, parent, text, command, kind="ghost", width=90, height=30):
        """Themed button. 'ghost' is the theme default (hairline outline),
        'primary' is the single white call-to-action, 'danger' an outlined abort."""
        overrides = {
            "ghost": {},
            "primary": dict(
                fg_color=T.TEXT, hover_color="#C9C9CF", text_color="#000000", border_width=0
            ),
            "danger": dict(
                text_color=T.ABORT, border_color="#5C1F1B", hover_color="#2A0E0C"
            ),
        }[kind]
        return ctk.CTkButton(
            parent, text=text, command=command, width=width, height=height,
            font=ctk.CTkFont(size=11, weight="bold"), **overrides
        )

    def _make_dialog(self, title, width, height):
        """Modal dialog shell with the tracked title and hairline rule."""
        dialog = ctk.CTkToplevel(self.window)
        dialog.title(title)
        dialog.geometry(f"{width}x{height}")
        dialog.transient(self.window)
        try:
            dialog.grab_set()
        except Exception:
            pass
        if self.topmost_var.get():
            dialog.attributes('-topmost', True)

        ctk.CTkLabel(
            dialog, text=self._track(title), font=ctk.CTkFont(size=13, weight="bold")
        ).pack(anchor="w", padx=24, pady=(20, 6))
        ctk.CTkFrame(dialog, height=1, fg_color=T.LINE, corner_radius=0).pack(
            fill="x", padx=24, pady=(0, 12)
        )
        return dialog

    def _confirm(self, title, message, confirm_text="CONFIRM"):
        """Themed replacement for messagebox.askyesno. Blocks until answered."""
        result = {"ok": False}
        dialog = self._make_dialog(title, 400, 200)

        ctk.CTkLabel(
            dialog, text=message, font=ctk.CTkFont(size=12),
            wraplength=350, justify="left", anchor="w"
        ).pack(fill="x", padx=24, pady=(4, 0))

        row = ctk.CTkFrame(dialog, fg_color="transparent")
        row.pack(side="bottom", fill="x", padx=24, pady=18)

        def accept():
            result["ok"] = True
            dialog.destroy()

        self._btn(row, confirm_text, accept, kind="danger", width=110).pack(side="right")
        self._btn(row, "CANCEL", dialog.destroy, width=90).pack(side="right", padx=(0, 8))

        dialog.wait_window()
        return result["ok"]

    def _apply_row_style(self, item_id, hover=False):
        """Selected = white outline, pinned = amber outline, hover = lifted fill."""
        frame = self.row_frames.get(item_id)
        if frame is None:
            return
        _, rec = self._lookup(item_id)
        if rec is None:
            return
        selected = item_id in self.selected_ids
        border = T.TEXT if selected else (T.WARN if rec["pinned"] else T.LINE)
        fill = T.ROW_HOVER if (hover or selected) else T.ROW
        frame.configure(border_color=border, fg_color=fill)

    def _on_row_leave(self, event, item_id):
        """<Leave> also fires when the pointer moves onto the row's own child
        widgets, so only drop the hover state if the pointer truly left the row."""
        frame = self.row_frames.get(item_id)
        if frame is None:
            return
        try:
            widget = frame.winfo_containing(event.x_root, event.y_root)
        except Exception:
            widget = None
        while widget is not None:
            if widget is frame:
                return
            widget = getattr(widget, "master", None)
        self._apply_row_style(item_id, hover=False)

    def _refresh_telemetry(self):
        """Counts plus the capacity gauge: white, amber at 80%, red when full."""
        used = len(self.items)
        pinned = sum(1 for rec in self.items if rec["pinned"])
        ratio = min(used / max(self.max_items, 1), 1.0)
        color = T.ABORT if ratio >= 1.0 else (T.WARN if ratio >= 0.8 else T.TEXT)
        self.count_label.configure(
            text=f"ITEMS {used:02d}/{self.max_items:02d}   PINNED {pinned:02d}"
        )
        self.capacity_bar.configure(progress_color=color)
        self.capacity_bar.set(ratio)

    def create_widgets(self):
        # Main container
        self.main_container = ctk.CTkFrame(self.window, fg_color=T.BG, corner_radius=0)
        self.main_container.pack(fill="both", expand=True, padx=14, pady=12)

        # Header: tracked wordmark left, live status right
        self.header = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.header.pack(fill="x", padx=10, pady=(4, 0))

        self.title_label = ctk.CTkLabel(
            self.header,
            text=self._track("Clipboard Manager Pro"),
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.title_label.pack(side="left")

        self.status_label = ctk.CTkLabel(
            self.header, text="● CAPTURING", font=self._mono(11, True), text_color=T.GO
        )
        self.status_label.pack(side="right")

        ctk.CTkFrame(
            self.main_container, height=1, fg_color=T.LINE, corner_radius=0
        ).pack(fill="x", padx=10, pady=(10, 8))

        # Top controls frame
        self.top_controls = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.top_controls.pack(fill="x", pady=4, padx=10)

        switch_font = ctk.CTkFont(size=11, weight="bold")

        # Topmost toggle (left side)
        self.topmost_var = ctk.BooleanVar(value=False)
        self.topmost_toggle = ctk.CTkSwitch(
            self.top_controls,
            text="KEEP ON TOP",
            variable=self.topmost_var,
            command=self.toggle_topmost,
            font=switch_font,
            switch_width=34,
            switch_height=16
        )
        self.topmost_toggle.pack(side="left", padx=(0, 14))

        # Selection mode toggle (middle)
        self.mode_var = ctk.BooleanVar(value=False)
        self.mode_toggle = ctk.CTkSwitch(
            self.top_controls,
            text="SINGLE-SELECT",
            variable=self.mode_var,
            command=self.toggle_selection_mode,
            font=switch_font,
            switch_width=34,
            switch_height=16
        )
        self.mode_toggle.pack(side="left", padx=14)

        # Clear-on-close toggle (middle) — when on, history is wiped on
        # exit instead of persisted for the next launch
        self.clear_on_close_var = ctk.BooleanVar(value=self.clear_on_close)
        self.clear_on_close_toggle = ctk.CTkSwitch(
            self.top_controls,
            text="CLEAR ON CLOSE",
            variable=self.clear_on_close_var,
            command=self.toggle_clear_on_close,
            font=switch_font,
            switch_width=34,
            switch_height=16
        )
        self.clear_on_close_toggle.pack(side="left", padx=14)

        # Copy files button (right side)
        self.copy_files_button = self._btn(
            self.top_controls, "+ ADD FILES", self.add_files_manually, width=96
        )
        self.copy_files_button.pack(side="right", padx=(6, 0))

        # Quick guide button (right side)
        self.guide_button = self._btn(
            self.top_controls, "GUIDE", self.open_quick_guide_dialog, width=64
        )
        self.guide_button.pack(side="right", padx=6)

        # Check-for-updates button (right side) — manual trigger for the
        # same updater that runs silently at startup. Always shows a
        # result dialog, including "you're up to date" and network errors.
        self.update_button = self._btn(
            self.top_controls, "UPDATES", self.check_for_updates_manual, width=76
        )
        self.update_button.pack(side="right", padx=6)

        # Instructions
        self.instructions = ctk.CTkLabel(
            self.main_container,
            text="CTRL+C CAPTURE  ·  CTRL+F SEARCH  ·  ESC CLEAR  ·  DRAG ROWS INTO ANY APP  ·  ◆ PIN TO PROTECT",
            font=self._mono(10),
            text_color=T.MUTED
        )
        self.instructions.pack(pady=(8, 4))

        # Telemetry strip: selection state left, counts right, capacity gauge below
        self.telemetry_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.telemetry_frame.pack(fill="x", padx=10, pady=(6, 0))

        self.selection_label = ctk.CTkLabel(
            self.telemetry_frame,
            text="NO ITEMS SELECTED",
            font=self._mono(11, True),
            text_color=T.MUTED
        )
        self.selection_label.pack(side="left")

        self.count_label = ctk.CTkLabel(
            self.telemetry_frame,
            text="",
            font=self._mono(11, True),
            text_color=T.MUTED
        )
        self.count_label.pack(side="right")

        self.capacity_bar = ctk.CTkProgressBar(self.main_container, height=3)
        self.capacity_bar.set(0)
        self.capacity_bar.pack(fill="x", padx=10, pady=(4, 8))

        # Main content frame (left: list, right: preview)
        self.content_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.content_frame.pack(fill="both", expand=True, padx=10, pady=4)

        # Left frame for clipboard items
        self.content_frame.grid_columnconfigure(0, weight=1, uniform="content_col")
        self.content_frame.grid_columnconfigure(1, weight=1, uniform="content_col")
        self.content_frame.grid_rowconfigure(0, weight=1)

        self.left_frame = ctk.CTkFrame(
            self.content_frame, fg_color=T.PANEL, border_width=1, border_color=T.LINE
        )
        self.left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        # Search bar
        self.search_frame = ctk.CTkFrame(self.left_frame, fg_color="transparent")
        self.search_frame.pack(fill="x", padx=6, pady=(6, 2))

        self.search_entry = ctk.CTkEntry(
            self.search_frame,
            height=30,
            font=self._mono(11),
            placeholder_text="SEARCH HISTORY"
        )
        self.search_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.search_entry.bind("<KeyRelease>", self.on_search_changed)

        self.clear_search_button = self._btn(
            self.search_frame, "✕", self.clear_search, width=30
        )
        self.clear_search_button.pack(side="right")

        # Scrollable frame for clipboard items
        self.scrollable_frame = ctk.CTkScrollableFrame(
            self.left_frame,
            width=350,
            height=380,
            fg_color=T.PANEL
        )
        self.scrollable_frame.pack(fill="both", expand=True, padx=4, pady=4)

        # Right frame for preview
        self.right_frame = ctk.CTkFrame(
            self.content_frame, fg_color=T.PANEL, border_width=1, border_color=T.LINE
        )
        self.right_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))

        # Preview title
        self.preview_title = ctk.CTkLabel(
            self.right_frame,
            text=self._track("Preview"),
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=T.MUTED
        )
        self.preview_title.pack(pady=(12, 4))

        # Slot for the in-app media player — zero height until a file is played
        self.player_slot = tk.Frame(
            self.right_frame, bg=T.PANEL, bd=0, highlightthickness=0
        )
        self.player_slot.pack(fill="x")

        # Preview content area
        self.preview_frame = ctk.CTkScrollableFrame(
            self.right_frame,
            width=350,
            height=350,
            fg_color=T.PANEL
        )
        self.preview_frame.pack(fill="both", expand=True, padx=6, pady=4)

        # Selection buttons row
        self.button_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.button_frame.pack(pady=(10, 4), padx=10, fill="x")

        self.select_all_button = self._btn(
            self.button_frame, "SELECT ALL", self.select_all_items, width=96
        )
        self.select_all_button.pack(side="left", padx=(0, 6))

        self.deselect_all_button = self._btn(
            self.button_frame, "DESELECT ALL", self.deselect_all_items, width=110
        )
        self.deselect_all_button.pack(side="left", padx=6)

        self.roulette_button = self._btn(
            self.button_frame, "SURPRISE ME", self.clipboard_roulette, width=110
        )
        # self.roulette_button.pack(side="left", padx=5)

        self.copy_selected_button = self._btn(
            self.button_frame, "COPY SELECTED", self.copy_selected_items,
            kind="primary", width=140
        )
        self.copy_selected_button.pack(side="left", padx=6)

        self.clear_button = self._btn(
            self.button_frame, "CLEAR ALL", self.clear_history, kind="danger", width=84
        )
        self.clear_button.pack(side="right", padx=(6, 0))

        self.max_items_button = self._btn(
            self.button_frame, "MAX ITEMS", self.open_max_items_dialog, width=86
        )
        self.max_items_button.pack(side="right", padx=6)

        self.stats_button = self._btn(
            self.button_frame, "STATS", self.open_stats_dialog, width=70
        )
        self.stats_button.pack(side="right", padx=6)

        # Keyboard shortcuts
        self.window.bind("<Control-f>", lambda e: self.search_entry.focus_set())
        self.window.bind("<Control-F>", lambda e: self.search_entry.focus_set())
        self.window.bind("<Escape>", lambda e: self.clear_search())

    # ------------------------------------------------------------------
    # Persistence
    # ------------------------------------------------------------------

    def load_history(self):
        if not self.history_file.exists():
            return
        try:
            with open(self.history_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            self.max_items = data.get("max_items", self.max_items)
            self.stats = data.get("stats", self.stats)
            self.clear_on_close = data.get("clear_on_close", self.clear_on_close)
            self.items = self._records_from_saved(data)
            if self.items:
                # Suppress the monitor briefly so whatever is on the clipboard
                # at startup isn't re-detected as a new item
                self._suppress_monitor_until = time.time() + 2.0
        except Exception:
            pass  # Corrupt or unreadable history file — start fresh rather than crash

    # ------------------------------------------------------------------
    # Thread hand-off and history records
    # ------------------------------------------------------------------

    def _post(self, fn):
        """Thread-safe: run fn on the Tk main thread."""
        self._ui_queue.put(fn)

    def _pump_ui_queue(self):
        try:
            while True:
                self._ui_queue.get_nowait()()
        except queue.Empty:
            pass
        except Exception as e:
            print(f"UI task failed: {e}")
        if self.monitoring:
            self.window.after(50, self._pump_ui_queue)

    @staticmethod
    def _normalize(content):
        return content if isinstance(content, list) else [content]

    def _new_record(self, content, item_type):
        return {
            "id": uuid.uuid4().hex[:12],
            "type": item_type,
            "content": content,
            "pinned": False,
            "ts": datetime.now().isoformat(),
        }

    def _lookup(self, item_id):
        """Return (index, record) for an ID, or (None, None)."""
        for index, rec in enumerate(self.items):
            if rec["id"] == item_id:
                return index, rec
        return None, None

    def _selected_records(self):
        """Selected records in history order."""
        return [rec for rec in self.items if rec["id"] in self.selected_ids]

    def _set_selected(self, item_id, state):
        """Single choke point for changing selection: set, row var, widget, style."""
        if state:
            self.selected_ids.add(item_id)
        else:
            self.selected_ids.discard(item_id)
        var = self.row_vars.get(item_id)
        if var is not None and bool(var.get()) != state:
            var.set(state)
        box = self.checkboxes.get(item_id)
        if box is not None:
            if state:
                box.select()
            else:
                box.deselect()
        self._apply_row_style(item_id)

    def _enforce_capacity(self, announce=True):
        """Evict the oldest unpinned items beyond max_items. If everything
        else is pinned, the newest item is the one rejected, so history can
        never grow without bound."""
        while len(self.items) > self.max_items:
            victim = next((r for r in self.items if not r["pinned"]), None)
            if victim is None:
                break
            rejected_newest = victim is self.items[-1]
            self.items.remove(victim)
            self.selected_ids.discard(victim["id"])
            if announce:
                if rejected_newest:
                    self.show_toast("All items pinned — new item not stored", "orange")
                else:
                    self.show_toast(f"Removed oldest unpinned item to maintain {self.max_items} max", "orange")

    def _records_from_saved(self, data):
        """Build records from the current format ('items') or the legacy
        four-parallel-lists format, so existing history survives the upgrade."""
        records = []
        if "items" in data:
            for raw in data["items"]:
                if not isinstance(raw, dict) or raw.get("type") not in ("text", "file"):
                    continue
                content = raw.get("content")
                if raw["type"] == "text" and not isinstance(content, str):
                    continue
                if raw["type"] == "file" and not isinstance(content, (list, str)):
                    continue
                records.append({
                    "id": raw.get("id") or uuid.uuid4().hex[:12],
                    "type": raw["type"],
                    "content": content,
                    "pinned": bool(raw.get("pinned", False)),
                    "ts": raw.get("ts", ""),
                })
            return records

        history = data.get("history", [])
        types = data.get("types", [])
        pinned = data.get("pinned", [])
        stamps = data.get("timestamps", [])
        for i, content in enumerate(history):
            item_type = types[i] if i < len(types) else ("file" if isinstance(content, list) else "text")
            records.append({
                "id": uuid.uuid4().hex[:12],
                "type": item_type,
                "content": content,
                "pinned": bool(pinned[i]) if i < len(pinned) else False,
                "ts": stamps[i] if i < len(stamps) else "",
            })
        return records

    # ------------------------------------------------------------------
    # Image capture
    # ------------------------------------------------------------------

    def _is_cached_image(self, path):
        """True if `path` is one of our cached clipboard images."""
        try:
            folder = os.path.dirname(os.path.abspath(str(path)))
            return os.path.normcase(folder) == os.path.normcase(str(self.image_dir))
        except Exception:
            return False

    def _grab_clipboard_image(self):
        """Return a PIL image if the clipboard holds bitmap data, else None.
        Monitor thread only, and only after the sequence number changed with
        neither files nor text present."""
        for _ in range(3):
            try:
                grabbed = ImageGrab.grabclipboard()
            except Exception:
                # Clipboard transiently locked by another process — retry briefly
                time.sleep(0.05)
                continue
            return grabbed if isinstance(grabbed, Image.Image) else None
        return None

    def _store_clipboard_image(self, img):
        """Persist a clipboard bitmap as PNG and return its path. The name is
        a content hash, so an identical image always maps to the same file."""
        if img.mode not in ("RGB", "RGBA", "L", "LA"):
            img = img.convert("RGBA")
        digest = hashlib.sha1(img.tobytes()).hexdigest()[:16]
        self.image_dir.mkdir(parents=True, exist_ok=True)
        path = self.image_dir / f"image_{digest}.png"
        if path.exists():
            os.utime(path)  # refresh mtime so the prune grace window covers it
        else:
            tmp = path.with_suffix(".png.tmp")
            img.save(tmp, format="PNG")
            os.replace(tmp, path)
        return str(path)

    def _ingest_image(self, img):
        """Monitor thread: store the bitmap, then hand the record to the UI thread."""
        if img.width * img.height > self.MAX_IMAGE_PIXELS:
            self._post(lambda: self.show_toast("Clipboard image too large to store", "orange"))
            return
        try:
            path = self._store_clipboard_image(img)
        except Exception as e:
            print(f"Image capture failed: {e}")
            return
        self.last_clipboard_content = [path]
        self._post(lambda p=path: self._on_image_captured(p))

    def _on_image_captured(self, path):
        """Main thread. Skip if the newest record already is this image, e.g. a
        bitmap that was already on the clipboard when the app launched."""
        if self.items:
            newest = self.items[-1]
            if newest["type"] == 'file' and self._normalize(newest["content"]) == [path]:
                return
        self.add_new_item([path], 'file')
        self.show_toast("Image captured", "green")

    def _prune_image_cache(self, grace=120):
        """Delete cached images that no history record references. Files
        touched within `grace` seconds are kept: the monitor thread may have
        just written one whose record is still waiting in the UI queue."""
        try:
            if not self.image_dir.is_dir():
                return
            referenced = {
                os.path.normcase(os.path.abspath(p))
                for rec in self.items if rec["type"] == 'file'
                for p in self._normalize(rec["content"]) if isinstance(p, str)
            }
            cutoff = time.time() - grace
            for entry in self.image_dir.glob("image_*"):
                if os.path.normcase(str(entry)) in referenced:
                    continue
                try:
                    if entry.stat().st_mtime <= cutoff:
                        entry.unlink()
                except OSError:
                    pass
        except Exception:
            pass

    def _image_clipboard_formats(self, file_list):
        """For a single cached image, the bitmap formats to publish next to
        CF_HDROP so pasting into Paint, Word or a chat box yields the picture
        itself rather than only a file. Returns [(format_id, bytes), ...]."""
        if len(file_list) != 1 or not self._is_cached_image(file_list[0]):
            return []
        try:
            import win32clipboard
            with open(file_list[0], 'rb') as f:
                png_bytes = f.read()
            with Image.open(file_list[0]) as img:
                bmp = io.BytesIO()
                img.convert("RGB").save(bmp, format="BMP")
            dib = bmp.getvalue()[14:]  # CF_DIB is a BMP without its 14-byte file header
            return [
                (win32clipboard.RegisterClipboardFormat("PNG"), png_bytes),
                (win32clipboard.CF_DIB, dib),
            ]
        except Exception as e:
            print(f"Image clipboard formats unavailable: {e}")
            return []

    def _read_current_clipboard_state(self):
        """Return whatever's currently on the OS clipboard, in the same shape
        monitor_clipboard() compares against last_clipboard_content — a list
        of file paths if files are present, otherwise the clipboard text."""
        file_paths = self.get_clipboard_files()
        if file_paths:
            return file_paths
        try:
            return pyperclip.paste()
        except Exception:
            return ""

    def save_history(self):
        try:
            data = {
                "items": self.items,
                "max_items": self.max_items,
                "stats": self.stats,
                "clear_on_close": self.clear_on_close,
            }
            tmp_file = self.history_file.with_suffix(".json.tmp")
            with open(tmp_file, 'w', encoding='utf-8') as f:
                json.dump(data, f)
            os.replace(tmp_file, self.history_file)
            self._prune_image_cache()
        except Exception:
            pass  # Saving is best-effort — never let it interrupt the user's workflow

    def on_close(self):
        if self.clear_on_close:
            self.items.clear()
        self.save_history()
        self._prune_image_cache(grace=0)
        self.close_media_player()
        if self._fs_window is not None:
            try:
                self._fs_window.destroy()
            except Exception:
                pass
        self.stop_monitoring()
        self.window.destroy()

    # ------------------------------------------------------------------
    # Toggles
    # ------------------------------------------------------------------

    def toggle_topmost(self):
        """Toggle window always-on-top mode"""
        self.window.attributes('-topmost', self.topmost_var.get())
        if self.topmost_var.get():
            self.show_toast("Window will stay on top", "green")
        else:
            self.show_toast("Topmost disabled", "gray")

    def toggle_clear_on_close(self):
        """Toggle whether clipboard history is wiped when the app closes"""
        self.clear_on_close = self.clear_on_close_var.get()
        self.save_history()
        if self.clear_on_close:
            self.show_toast("History will be cleared on close", "orange")
        else:
            self.show_toast("History will be kept on close", "gray")

    def toggle_selection_mode(self):
        """Toggle between single and multiple selection mode"""
        self.multiple_selection_mode = self.mode_var.get()

        if not self.multiple_selection_mode:
            self.mode_toggle.configure(text="SINGLE-SELECT")
            self.enforce_single_selection()
            self.show_toast("Single-select mode enabled", "blue")
        else:
            self.mode_toggle.configure(text="MULTI-SELECT")
            self.update_selection_label()
            self.show_toast("Multi-select mode enabled", "green")

    def enforce_single_selection(self):
        """Ensure only one item is selected (for single-select mode)"""
        if not self.multiple_selection_mode:
            selected = self._selected_records()

            if len(selected) > 1:
                for rec in selected[:-1]:
                    self._set_selected(rec["id"], False)
                self.copy_item_to_clipboard(selected[-1]["id"])

            self.update_selection_label()
            self.update_preview()

    # ------------------------------------------------------------------
    # Clipboard monitoring
    # ------------------------------------------------------------------

    def monitor_clipboard(self):
        """Monitor clipboard for changes in background thread"""
        last_seq = None
        while self.monitoring:
            if time.time() < self._suppress_monitor_until:
                # Skip monitoring during suppression period
                time.sleep(0.2)
                continue
            try:
                seq = win32clipboard.GetClipboardSequenceNumber() if _NATIVE_DND_AVAILABLE else None
                if seq is not None and seq == last_seq:
                    time.sleep(0.2)
                    continue
                current_content = pyperclip.paste()
                file_paths = self.get_clipboard_files()

                # Dedupe lives in add_new_item (main thread) — the worker only
                # reports changes and never touches self.items or Tk
                if file_paths:
                    if file_paths != self.last_clipboard_content:
                        self.last_clipboard_content = file_paths
                        self._post(lambda fp=file_paths: self.add_new_item(fp, 'file'))
                elif current_content != self.last_clipboard_content and current_content.strip():
                    self.last_clipboard_content = current_content
                    self._post(lambda c=current_content: self.add_new_item(c, 'text'))
                elif seq is not None and not current_content.strip():
                    # Neither files nor text: the clipboard may hold a bitmap
                    # (screenshot, "Copy image"). Only probed after the sequence
                    # number changed, since decoding a bitmap is not cheap.
                    img = self._grab_clipboard_image()
                    if img is not None:
                        self._ingest_image(img)
                last_seq = seq
            except Exception:
                pass
            time.sleep(0.5)

    def get_clipboard_files(self):
        """Get file paths from clipboard (Windows, macOS, Linux best-effort)"""
        system = platform.system()
        if system == "Windows":
            return self._get_clipboard_files_windows()
        try:
            import subprocess
            if system == "Darwin":
                check = subprocess.run(
                    ['osascript', '-e', 'clipboard info'],
                    capture_output=True, text=True, timeout=2
                )
                if 'furl' in check.stdout:
                    result = subprocess.run(
                        ['osascript', '-e', 'POSIX path of (the clipboard as «class furl»)'],
                        capture_output=True, text=True, timeout=2
                    )
                    path = result.stdout.strip()
                    if path:
                        return [path]
            elif system == "Linux":
                targets = subprocess.run(
                    ['xclip', '-selection', 'clipboard', '-t', 'TARGETS', '-o'],
                    capture_output=True, text=True, timeout=2
                )
                if 'text/uri-list' in targets.stdout:
                    result = subprocess.run(
                        ['xclip', '-selection', 'clipboard', '-t', 'text/uri-list', '-o'],
                        capture_output=True, text=True, timeout=2
                    )
                    files = [
                        f.strip().replace('file://', '')
                        for f in result.stdout.split('\n') if f.strip()
                    ]
                    return files if files else None
        except (FileNotFoundError, Exception):
            pass
        return None

    def _get_clipboard_files_windows(self):
        """Read CF_HDROP directly via the Win32 API. Far faster and more reliable
        than shelling out to PowerShell on every 500ms poll, and avoids parsing
        PowerShell's formatted table output entirely."""
        try:
            import win32clipboard
        except ImportError:
            if not self._pywin32_warning_shown:
                self._pywin32_warning_shown = True
                self._post(lambda: self.show_toast(
                    "File-copy detection needs pywin32: pip install pywin32", "orange"
                ))
            return None

        for _ in range(3):
            try:
                win32clipboard.OpenClipboard()
                try:
                    if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_HDROP):
                        files = win32clipboard.GetClipboardData(win32clipboard.CF_HDROP)
                        files = [f for f in files if os.path.exists(f)]
                        return files if files else None
                    return None
                finally:
                    win32clipboard.CloseClipboard()
            except Exception:
                # Clipboard transiently locked by another process (common — retry briefly)
                time.sleep(0.05)
        return None

    def _copy_files_to_clipboard_windows(self, file_list, text=None):
        """Write files to the clipboard as a native CF_HDROP — the same format
        Explorer produces. This replaces the old `powershell.exe -Command
        Set-Clipboard` call, which briefly flashed a console window because
        the child process gets its own console when the parent (customtkinter,
        no attached console) spawns it."""
        try:
            import win32clipboard
        except ImportError:
            if not self._pywin32_warning_shown:
                self._pywin32_warning_shown = True
                self.window.after(0, lambda: self.show_toast(
                    "File-copy needs pywin32: pip install pywin32", "orange"
                ))
            return False

        # sizeof(DROPFILES): DWORD pFiles + POINT pt (2x LONG) + BOOL fNC + BOOL fWide
        offset = 20
        file_block = ('\0'.join(file_list) + '\0\0').encode('utf-16-le')
        dropfiles = struct.pack('Iiiii', offset, 0, 0, 0, 1) + file_block
        image_formats = self._image_clipboard_formats(file_list)

        for _ in range(3):
            try:
                win32clipboard.OpenClipboard()
                try:
                    win32clipboard.EmptyClipboard()
                    win32clipboard.SetClipboardData(win32clipboard.CF_HDROP, dropfiles)
                    if text:
                        # CF_HDROP and CF_UNICODETEXT coexist on one clipboard, so a
                        # mixed selection pastes as files or as text, per target
                        win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, text)
                    for fmt, payload in image_formats:
                        win32clipboard.SetClipboardData(fmt, payload)
                    return True
                finally:
                    win32clipboard.CloseClipboard()
            except Exception:
                # Clipboard transiently locked by another process — retry briefly
                time.sleep(0.05)
        return False

    def add_files_manually(self):
        """Open file dialog to add files"""
        from tkinter import filedialog
        files = filedialog.askopenfilenames(
            title="Select files to add",
            filetypes=[
                ("All files", "*.*"),
                ("Images", "*.png *.jpg *.jpeg *.gif *.bmp"),
                ("Documents", "*.txt *.pdf *.doc *.docx"),
                ("Music", "*.mp3 *.wav *.ogg"),
                ("Videos", "*.mp4 *.avi *.mkv")
            ]
        )
        if files:
            self.add_new_item(list(files), 'file')
            self.show_toast(f"Added {len(files)} file(s)", "green")

    # ------------------------------------------------------------------
    # Core history management
    # ------------------------------------------------------------------

    def add_new_item(self, content, item_type='text'):
        """Add an item as the newest entry. Re-copying something already in
        history bumps it to the newest slot (keeping its ID and pin) instead
        of storing a duplicate."""
        existing = None
        for rec in self.items:
            if rec["type"] != item_type:
                continue
            if item_type == 'file':
                same = self._normalize(rec["content"]) == self._normalize(content)
            else:
                same = rec["content"] == content
            if same:
                existing = rec
                break

        if existing is not None:
            self.items.remove(existing)
            existing["ts"] = datetime.now().isoformat()
            self.items.append(existing)
        else:
            self.items.append(self._new_record(content, item_type))
            self.stats["total_items"] += 1
            if item_type == 'text':
                self.stats["total_characters"] += len(content)
            self._enforce_capacity()

        self.save_history()
        self.update_display()

    def item_matches_query(self, rec, query):
        if rec["type"] == 'file':
            return any(query in os.path.basename(n).lower() for n in self._normalize(rec["content"]))
        return query in str(rec["content"]).lower()

    def update_display(self):
        """Update the scrollable frame with clipboard items"""
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        self.checkboxes.clear()
        self.row_vars.clear()
        self.row_frames.clear()

        # Selection lives in self.selected_ids (stable item IDs), so it survives
        # eviction, deletion, bumping and search filtering
        self.selected_ids &= {rec["id"] for rec in self.items}

        query = self.search_query.strip().lower()
        visible_count = 0
        for index, rec in enumerate(self.items):
            if query and not self.item_matches_query(rec, query):
                continue
            visible_count += 1
            self.create_clipboard_item(index, rec)

        if query and visible_count == 0:
            no_results = ctk.CTkLabel(
                self.scrollable_frame,
                text=f'No items match "{self.search_query}"',
                font=ctk.CTkFont(size=12),
                text_color="gray"
            )
            no_results.pack(pady=20)

        self._refresh_telemetry()
        self.update_selection_label()
        self.update_preview()

    def create_clipboard_item(self, index, rec):
        """Create one clipboard history item. Callbacks capture the record's
        stable ID, never its list position."""

        item_id = rec["id"]
        item_type = rec["type"]
        content = rec["content"]
        is_pinned = rec["pinned"]

        item_frame = ctk.CTkFrame(
            self.scrollable_frame,
            fg_color=T.ROW,
            border_width=1,
            border_color=T.LINE
        )
        item_frame.pack(fill="x", pady=3, padx=4)
        self.row_frames[item_id] = item_frame

        # --------------------------------------------------------------
        # Checkbox
        # --------------------------------------------------------------

        checkbox_var = ctk.BooleanVar(value=item_id in self.selected_ids)

        checkbox = ctk.CTkCheckBox(
            item_frame,
            text="",
            variable=checkbox_var,
            width=22,
            checkbox_width=18,
            checkbox_height=18,
            command=lambda i=item_id:
                self.on_checkbox_toggle(i)
        )
        checkbox.pack(side="left", padx=(8, 4), pady=6)

        self.row_vars[item_id] = checkbox_var
        self.checkboxes[item_id] = checkbox

        # --------------------------------------------------------------
        # Determine display text and icon FIRST
        # --------------------------------------------------------------

        if item_type == 'file':

            paths = content if isinstance(content, list) else [content]

            valid_paths = [
                os.path.abspath(str(path))
                for path in paths
                if isinstance(path, str) and os.path.exists(path)
            ]

            if len(paths) == 1:
                display_text = os.path.basename(str(paths[0]))
                icon = "FILE"
                if self._is_cached_image(paths[0]):
                    display_text = "Clipboard image"
                    icon = "IMG"
            else:
                filenames = [
                    os.path.basename(str(path))
                    for path in paths[:3]
                ]

                display_text = ", ".join(filenames)

                if len(paths) > 3:
                    display_text += f" ... (+{len(paths) - 3})"

                icon = f"×{len(paths)}"

        else:

            valid_paths = []

            display_text = str(content).replace("\n", " ")

            if len(display_text) > 50:
                display_text = display_text[:50] + "..."

            icon = "TEXT"

        # --------------------------------------------------------------
        # Icon
        # --------------------------------------------------------------

        # Files that no longer exist on disk are dimmed
        is_missing = (item_type == 'file' and not valid_paths)

        icon_label = ctk.CTkLabel(
            item_frame,
            text=icon,
            width=44,
            font=self._mono(10, True),
            text_color=T.DIM if is_missing else T.MUTED,
            cursor="hand2"
        )
        icon_label.pack(side="left", padx=2)

        # --------------------------------------------------------------
        # Item number
        # --------------------------------------------------------------

        index_label = ctk.CTkLabel(
            item_frame,
            text=f"{index + 1:02d}",
            width=30,
            font=self._mono(13, True)
        )
        index_label.pack(side="left", padx=5)

        # --------------------------------------------------------------
        # Timestamp
        # --------------------------------------------------------------

        time_label = ctk.CTkLabel(
            item_frame,
            text=self.format_relative_time(rec["ts"]),
            width=56,
            font=self._mono(10),
            text_color=T.MUTED
        )
        time_label.pack(side="left", padx=2)

        # --------------------------------------------------------------
        # Main filename / text label
        # --------------------------------------------------------------

        content_label = ctk.CTkLabel(
            item_frame,
            text=display_text,
            anchor="w",
            font=self._mono(11),
            text_color=T.DIM if is_missing else T.TEXT,
            cursor="hand2"
        )
        content_label.pack(
            side="left",
            fill="x",
            expand=True,
            padx=5
        )

        # --------------------------------------------------------------
        # Normal click = selection
        # --------------------------------------------------------------

        # --------------------------------------------------------------
        # Press/release on the row: press shows a drag preview (if this
        # item is draggable) without selecting yet; release decides
        # click-vs-drag and only then toggles the checkbox. See
        # _on_item_press / _on_item_release for why this split exists.
        # --------------------------------------------------------------

        drag_paths = valid_paths if (item_type == 'file' and valid_paths) else None
        is_text_drag = (item_type == 'text')
        drag_text = str(content) if is_text_drag else None
        press_payload = drag_paths if drag_paths is not None else (
            [str(content)[:60]] if is_text_drag else None
        )

        for w in (content_label, icon_label, index_label, time_label):
            w.bind(
                "<ButtonPress-1>",
                lambda event, i=item_id, dp=press_payload, txt=is_text_drag, full_text=drag_text:
                    self._on_item_press(event, i, drag_paths=dp, is_text=txt, full_text=full_text),
                add="+"
            )
            w.bind(
                "<B1-Motion>",
                self._on_item_motion,
                add="+"
            )
            w.bind(
                "<ButtonRelease-1>",
                lambda event, i=item_id:
                    self._on_item_release(event, i),
                add="+"
            )

        # Hover highlight. Leave is filtered in _on_row_leave because moving
        # between a row's own child widgets also fires <Leave> on the others.
        for w in (item_frame, content_label, icon_label, index_label, time_label):
            w.bind("<Enter>", lambda event, i=item_id: self._apply_row_style(i, hover=True), add="+")
            w.bind("<Leave>", lambda event, i=item_id: self._on_row_leave(event, i), add="+")

        # --------------------------------------------------------------
        # NATIVE DRAG AND DROP
        # IMPORTANT: widgets now exist before binding them
        # --------------------------------------------------------------

        if item_type == 'file' and valid_paths:

            # Drag using filename, icon, or item number.
            self._bind_file_drag_source(
                content_label,
                valid_paths
            )

            self._bind_file_drag_source(
                icon_label,
                valid_paths
            )

            self._bind_file_drag_source(
                index_label,
                valid_paths
            )

        elif item_type == 'text':

            self._bind_text_drag_source(
                content_label,
                str(content)
            )

            self._bind_text_drag_source(
                icon_label,
                str(content)
            )

        # --------------------------------------------------------------
        # Delete button
        # --------------------------------------------------------------

        delete_button = ctk.CTkButton(
            item_frame,
            text="✕",
            width=26,
            height=26,
            border_width=0,
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=T.MUTED,
            hover_color="#3A1210",
            command=lambda i=item_id:
                self.delete_item(i)
        )
        delete_button.pack(side="right", padx=(2, 8))

        # --------------------------------------------------------------
        # Copy button
        # --------------------------------------------------------------

        copy_button = ctk.CTkButton(
            item_frame,
            text="COPY",
            width=54,
            height=26,
            font=ctk.CTkFont(size=10, weight="bold"),
            command=lambda i=item_id:
                self.copy_item_to_clipboard(i)
        )
        copy_button.pack(side="right", padx=4)

        # --------------------------------------------------------------
        # Pin button
        # --------------------------------------------------------------

        pin_button = ctk.CTkButton(
            item_frame,
            text="◆" if is_pinned else "◇",
            width=26,
            height=26,
            border_width=0,
            font=ctk.CTkFont(size=13),
            text_color=T.WARN if is_pinned else T.MUTED,
            command=lambda i=item_id:
                self.toggle_pin(i)
        )
        pin_button.pack(side="right", padx=2)

        self._apply_row_style(item_id)

    def toggle_checkbox_from_label(self, item_id):
        """Let clicking an item's text toggle its selection, same as the checkbox"""
        var = self.row_vars.get(item_id)
        if var is None:
            return
        var.set(not var.get())
        self.on_checkbox_toggle(item_id)

    def _on_item_press(self, event, item_id, drag_paths=None, is_text=False, full_text=None):
        """Mouse-down on a row's icon/name. Records where the press started
        and shows a drag preview, but does NOT select the item yet —
        selection only happens on release, and only if the press turns out
        to have been a click rather than a drag (see _on_item_release).
        This is what stops every drag-out from also flipping the checkbox.
        Also stashes the real drag payload for _on_item_motion, which is
        what actually fires the native OLE drag once the pointer moves."""
        self._press_start = (event.x_root, event.y_root)
        self._drag_started = False
        self._pending_drag_paths = drag_paths if not is_text else None
        self._pending_drag_text = full_text if is_text else None

        if drag_paths is not None:
            self.show_drag_visual(drag_paths, event, is_text=is_text)
            # Safety net: if this turns into a real native OS drag, the
            # modal drag loop can swallow our own ButtonRelease-1, which
            # would otherwise leave this visual stuck on screen forever.
            self.schedule_drag_visual_hide(4000)

    def _on_item_release(self, event, item_id):
        """Mouse-up on a row. Only toggles selection if the press-release
        pair was a genuine click — i.e. no native drag actually started,
        and the pointer didn't move past the click threshold in between."""
        self.schedule_drag_visual_hide(300)

        moved_far = False
        if self._press_start is not None:
            dx = event.x_root - self._press_start[0]
            dy = event.y_root - self._press_start[1]
            moved_far = (dx * dx + dy * dy) > (self.CLICK_MOVE_THRESHOLD ** 2)

        was_drag = self._drag_started or moved_far
        self._press_start = None
        self._drag_started = False
        self._pending_drag_paths = None
        self._pending_drag_text = None

        if not was_drag:
            self.toggle_checkbox_from_label(item_id)

    def toggle_pin(self, item_id):
        """Pin/unpin an item — pinned items are protected from eviction"""
        index, rec = self._lookup(item_id)
        if rec is None:
            return
        rec["pinned"] = not rec["pinned"]
        state = "pinned" if rec["pinned"] else "unpinned"
        self.show_toast(f"Item #{index + 1} {state}", "blue")
        self.save_history()
        self.update_display()

    def format_relative_time(self, iso_timestamp):
        try:
            ts = datetime.fromisoformat(iso_timestamp)
        except (ValueError, TypeError):
            return ""
        seconds = (datetime.now() - ts).total_seconds()
        if seconds < 60:
            return "just now"
        minutes = seconds / 60
        if minutes < 60:
            return f"{int(minutes)}m ago"
        hours = minutes / 60
        if hours < 24:
            return f"{int(hours)}h ago"
        return f"{int(hours / 24)}d ago"

    def on_checkbox_toggle(self, item_id):
        """Handle checkbox toggle event.

        Checking/unchecking a box is purely an internal selection action — it
        builds up the working set shown in the preview panel. It does NOT push
        anything to the system clipboard. In single-select mode we still copy
        immediately, since there checking a box unambiguously means "use this
        one item now." In multi-select mode, pushing the combined selection to
        the clipboard is a separate, explicit action (the Copy Selected button).
        """
        var = self.row_vars.get(item_id)
        if var is None:
            return
        checked = bool(var.get())
        single = not self.multiple_selection_mode

        if checked and single:
            for other in list(self.selected_ids):
                if other != item_id:
                    self._set_selected(other, False)
        self._set_selected(item_id, checked)

        if checked and single:
            index, _ = self._lookup(item_id)
            self.copy_item_to_clipboard(item_id)
            self.show_toast(f"Item #{index + 1} selected and copied", "green")

        self.update_selection_label()
        self.update_preview()

    def copy_item_to_clipboard(self, item_id):
        """Copy a specific item to clipboard"""
        index, rec = self._lookup(item_id)
        if rec is None:
            return
        if rec["type"] == 'file':
            self.copy_files_to_clipboard(rec["content"])
        else:
            pyperclip.copy(rec["content"])
            self.last_clipboard_content = rec["content"]
            self._suppress_monitor_until = time.time() + 1.0
            self.show_toast(f"Item #{index + 1} copied", "green")

    def copy_files_to_clipboard(self, files, text=None):
        """Copy files to system clipboard (Windows/macOS/Linux, with text fallback).
        `text` rides along as CF_UNICODETEXT on Windows only."""
        system = platform.system()
        file_list = files if isinstance(files, list) else [files]
        try:
            if system == "Windows":
                if not self._copy_files_to_clipboard_windows(file_list, text):
                    raise FileNotFoundError("native Windows clipboard write unavailable")
            elif system == "Darwin":
                import subprocess
                if len(file_list) == 1:
                    ascript = f'set the clipboard to (POSIX file "{file_list[0]}")'
                else:
                    posix_list = ", ".join(f'POSIX file "{f}"' for f in file_list)
                    ascript = f'set the clipboard to {{{posix_list}}}'
                subprocess.run(['osascript', '-e', ascript], timeout=2, check=True)
            elif system == "Linux":
                import subprocess
                uri_list = "\n".join(Path(f).resolve().as_uri() for f in file_list)
                subprocess.run(
                    ['xclip', '-selection', 'clipboard', '-t', 'text/uri-list'],
                    input=uri_list, text=True, timeout=2, check=True
                )
            else:
                raise OSError(f"Unsupported platform: {system}")
            self.last_clipboard_content = file_list
            self._suppress_monitor_until = time.time() + 1.0
            self.show_toast("Files copied to clipboard", "green")
        except FileNotFoundError:
            pyperclip.copy("\n".join(file_list))
            self.last_clipboard_content = "\n".join(file_list)
            self._suppress_monitor_until = time.time() + 1.0
            self.show_toast("Native file copy unavailable — copied paths as text instead", "orange")
        except Exception as e:
            self.show_toast(f"Error copying files: {str(e)}", "red")

    def copy_selected_items(self):
        """Copy all selected items to clipboard (combined if multiple).

        This is now the single explicit trigger for pushing the checked
        selection to the system clipboard — checking boxes alone no longer
        does this in multi-select mode.
        """
        selected = self._selected_records()

        if not selected:
            self.show_toast("Check one or more items first", "gray")
            return

        if len(selected) == 1:
            self.copy_item_to_clipboard(selected[0]["id"])
            return

        text_contents = [r["content"] for r in selected if r["type"] == 'text']
        file_contents = []
        for r in selected:
            if r["type"] == 'file':
                file_contents.extend(self._normalize(r["content"]))

        if text_contents and not file_contents:
            combined = "\n".join(text_contents)
            pyperclip.copy(combined)
            self.last_clipboard_content = combined
            self._suppress_monitor_until = time.time() + 1.0
            self.show_toast(f"{len(text_contents)} text items combined", "blue")
        elif file_contents and not text_contents:
            self.copy_files_to_clipboard(file_contents)
            self.show_toast(f"{len(file_contents)} files copied", "blue")
        else:
            combined = "\n".join(text_contents)
            self.copy_files_to_clipboard(file_contents, text=combined)
            if platform.system() == "Windows":
                self.show_toast(f"{len(text_contents)} text + {len(file_contents)} files copied together", "blue")
            else:
                self.show_toast("Text can't ride along with files here — files copied", "orange")

    def update_selection_label(self):
        """Update the selection count label"""
        nums = [i + 1 for i, rec in enumerate(self.items) if rec["id"] in self.selected_ids]
        mode = "MULTI-SELECT" if self.multiple_selection_mode else "SINGLE-SELECT"

        if not nums:
            text, color = f"NO ITEMS SELECTED  ·  {mode}", T.MUTED
        elif len(nums) == 1:
            text, color = f"ITEM #{nums[0]} SELECTED  ·  READY TO PASTE", T.GO
        else:
            text, color = f"ITEMS {nums} SELECTED  ·  {mode}", T.WARN

        self.selection_label.configure(text=text, text_color=color)

    def update_preview(self):
        """Update the preview panel with selected item details"""
        self._sync_player()
        for job in self._marquee_jobs:
            try:
                self.window.after_cancel(job)
            except Exception:
                pass
        self._marquee_jobs.clear()

        for widget in self.preview_frame.winfo_children():
            widget.destroy()

        selected = [(i, rec) for i, rec in enumerate(self.items) if rec["id"] in self.selected_ids]

        if not selected:
            label = ctk.CTkLabel(
                self.preview_frame, text=self._track("No selection"),
                font=self._mono(12, True), text_color=T.DIM
            )
            label.pack(pady=50)
            return

        for pos, (idx, rec) in enumerate(selected):
            content = rec["content"]
            time_note = self.format_relative_time(rec["ts"])

            num_label = ctk.CTkLabel(
                self.preview_frame,
                text=f"ITEM {idx + 1:02d}" + ("  ◆" if rec["pinned"] else ""),
                font=self._mono(13, True),
                text_color=T.WARN if rec["pinned"] else T.TEXT
            )
            num_label.pack(pady=(8, 0))

            if time_note:
                time_label = ctk.CTkLabel(
                    self.preview_frame, text=time_note.upper(), font=self._mono(10), text_color=T.MUTED
                )
                time_label.pack(pady=(0, 5))

            if rec["type"] == 'file':
                files = self._normalize(content)
                for file_path in files[:5]:
                    self.create_file_preview(file_path)
                if len(files) > 5:
                    more_label = ctk.CTkLabel(
                        self.preview_frame, text=f"... and {len(files) - 5} more files",
                        font=self._mono(11), text_color=T.MUTED
                    )
                    more_label.pack(pady=2)
            else:
                preview_text = content[:500] + "..." if len(content) > 500 else content
                text_frame = ctk.CTkFrame(
                    self.preview_frame, fg_color=T.ROW, border_width=1, border_color=T.LINE
                )
                text_frame.pack(fill="x", padx=6, pady=4)
                text_label = ctk.CTkLabel(
                    text_frame, text=preview_text, font=self._mono(11),
                    wraplength=260, justify="left", anchor="w"
                )
                text_label.pack(fill="x", padx=8, pady=8)

            if pos < len(selected) - 1:
                separator = ctk.CTkFrame(self.preview_frame, height=1, fg_color=T.LINE_STRONG, corner_radius=0)
                separator.pack(fill="x", pady=8, padx=10)

    def create_file_preview(self, file_path):
        """Create preview for a file"""
        if not os.path.exists(file_path):
            return

        file_name = os.path.basename(file_path)
        file_size = os.path.getsize(file_path)

        file_frame = ctk.CTkFrame(self.preview_frame)
        file_frame.pack(fill="x", pady=2, padx=5)

        ext = os.path.splitext(file_name)[1].lower()
        is_playable = ext in self.PLAYABLE_EXTENSIONS
        if ext in ['.png', '.jpg', '.jpeg', '.gif', '.bmp']:
            icon = "IMG"
            try:
                img = Image.open(file_path)
                img.thumbnail((220, 220))
                photo = ctk.CTkImage(light_image=img, dark_image=img, size=img.size)
                img_label = ctk.CTkLabel(self.preview_frame, image=photo, text="")
                img_label.pack(pady=5)
            except Exception:
                pass
        elif ext in ['.mp3', '.wav', '.ogg', '.flac', '.m4a', '.aac']:
            icon = "AUD"
        elif ext in ['.mp4', '.avi', '.mkv', '.mov', '.wmv', '.webm']:
            icon = "VID"
        elif ext in ['.pdf', '.doc', '.docx', '.txt']:
            icon = "DOC"
        else:
            icon = "FILE"

        icon_label = ctk.CTkLabel(
            file_frame, text=icon, font=self._mono(10, True), text_color=T.MUTED, width=36
        )
        icon_label.pack(side="left", padx=(5, 0))

        # Play button packs right after the icon — before the name — so it's
        # never squeezed out by a long filename in the fixed-width frame
        if is_playable:
            play_button = ctk.CTkButton(
                file_frame, text="▶", width=28, height=22,
                font=ctk.CTkFont(size=11), fg_color="#1565c0", hover_color="#0d47a1",
                command=lambda p=file_path: self.open_media_player(p)
            )
            play_button.pack(side="left", padx=(2, 5))

        size_label = ctk.CTkLabel(
            file_frame, text=f"({self.format_size(file_size)})",
            font=ctk.CTkFont(size=10), text_color="gray"
        )
        size_label.pack(side="right", padx=5)

        # Name gets whatever space is left. If it's too long to fit, it
        # scrolls (marquee-style) instead of pushing other widgets off-frame.
        name_label = ctk.CTkLabel(
            file_frame, text=file_name, font=ctk.CTkFont(size=11),
            anchor="w", width=130
        )
        name_label.pack(side="left", fill="x", expand=True, padx=5)
        self.start_marquee(name_label, file_name, visible_chars=20)

    def start_marquee(self, label, full_text, visible_chars=20):
        """Slowly scroll `full_text` through `label` if it's too long to fit.
        Short text is shown as-is with no animation."""
        if len(full_text) <= visible_chars:
            label.configure(text=full_text)
            return

        padded = full_text + "   •   "
        position = {"i": 0}

        def tick():
            if not label.winfo_exists():
                return
            i = position["i"]
            window_text = (padded + padded)[i:i + visible_chars]
            label.configure(text=window_text)
            position["i"] = (i + 1) % len(padded)
            job = self.window.after(280, tick)
            self._marquee_jobs.append(job)

        tick()

    def open_media_player(self, file_path, start_pos=0.0, paused=False, muted=True):
        """Play a file in the preview section. Starts muted unless a fullscreen
        hand-off explicitly passes the user's current mute state."""
        if not os.path.exists(file_path):
            self.show_toast("File no longer exists at that path", "red")
            return

        is_gif = os.path.splitext(file_path)[1].lower() == '.gif'
        if not is_gif and not _PLAYER_AVAILABLE:
            self.show_toast("In-app player needs ffpyplayer — opening in default player", "orange")
            self.play_file(file_path)
            return

        self.close_media_player()

        def build():
            self._player = MediaPlayerPanel(
                self.player_slot, self, os.path.abspath(file_path),
                start_pos=start_pos, paused=paused, muted=muted
            )
            self._player.pack(fill="x", padx=6, pady=(4, 0))
            if muted and not is_gif:
                self.show_toast("Playing muted — click the speaker to unmute", "gray")

        self._defer_player(build)

    def _defer_player(self, build):
        """Run `build` after a short gap so the previous player's audio device
        has fully shut down before a new one opens. A newer request replaces
        a pending one, so rapid clicks on several files open only the last."""
        if self._open_job is not None:
            try:
                self.window.after_cancel(self._open_job)
            except Exception:
                pass

        def run():
            self._open_job = None
            build()

        self._open_job = self.window.after(350, run)

    def close_media_player(self):
        if self._open_job is not None:
            try:
                self.window.after_cancel(self._open_job)
            except Exception:
                pass
            self._open_job = None
        panel, self._player = self._player, None
        if panel is not None:
            try:
                panel.destroy()
            except Exception:
                pass

    def _sync_player(self):
        """Close the inline player when its file is no longer in history."""
        panel = self._player
        if panel is None:
            return
        wanted = os.path.normcase(os.path.abspath(panel.path))
        in_history = any(
            os.path.normcase(os.path.abspath(p)) == wanted
            for rec in self.items if rec["type"] == 'file'
            for p in self._normalize(rec["content"]) if isinstance(p, str)
        )
        if not in_history:
            self.close_media_player()

    def toggle_player_fullscreen(self, panel):
        if panel.fullscreen:
            self._exit_player_fullscreen()
        else:
            self._enter_player_fullscreen(panel)

    def _enter_player_fullscreen(self, panel):
        state = panel.snapshot()
        self.close_media_player()
        self._defer_player(lambda: self._build_player_fullscreen(state))

    def _build_player_fullscreen(self, state):
        win = ctk.CTkToplevel(self.window)
        win.title("Player")
        win.configure(fg_color="#000000")
        win.protocol("WM_DELETE_WINDOW", self._exit_player_fullscreen)
        win.bind("<Escape>", lambda e: self._exit_player_fullscreen())
        self._fs_window = win

        self._fs_panel = MediaPlayerPanel(
            win, self, state["path"], start_pos=state["position"],
            paused=state["paused"], muted=state["muted"], fullscreen=True
        )
        self._fs_panel.pack(fill="both", expand=True)

        # CTkToplevel finishes its own setup ~200 ms after creation and can
        # drop the fullscreen attribute if it is set earlier
        def finish():
            try:
                win.attributes('-fullscreen', True)
                win.focus_force()
            except Exception:
                pass
        win.after(250, finish)

    def _exit_player_fullscreen(self):
        panel, win = self._fs_panel, self._fs_window
        if panel is None or win is None:
            return
        state = panel.snapshot()
        self._fs_panel = None
        self._fs_window = None
        try:
            win.destroy()
        except Exception:
            pass
        self.open_media_player(
            state["path"], start_pos=state["position"],
            paused=state["paused"], muted=state["muted"]
        )

    def play_file(self, file_path):
        """Open a media file with the OS default player. This just launches
        whatever the system has associated with the extension (e.g. Windows
        Media Player, QuickTime, VLC) — it doesn't embed playback in-app."""
        if not os.path.exists(file_path):
            self.show_toast("File no longer exists at that path", "red")
            return
        system = platform.system()
        try:
            if system == "Windows":
                os.startfile(file_path)
            elif system == "Darwin":
                import subprocess
                subprocess.run(['open', file_path], timeout=2)
            else:
                import subprocess
                subprocess.run(['xdg-open', file_path], timeout=2)
            self.show_toast(f"Playing {os.path.basename(file_path)}", "blue")
        except Exception as e:
            self.show_toast(f"Couldn't open file: {str(e)}", "red")

    def format_size(self, size):
        """Format file size"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024.0:
                return f"{size:.1f} {unit}"
            size /= 1024.0
        return f"{size:.1f} TB"

    # ------------------------------------------------------------------
    # Search
    # ------------------------------------------------------------------

    def on_search_changed(self, event=None):
        self.search_query = self.search_entry.get()
        self.update_display()

    def clear_search(self):
        self.search_entry.delete(0, 'end')
        self.search_query = ""
        self.update_display()

    # ------------------------------------------------------------------
    # Bulk actions
    # ------------------------------------------------------------------

    def select_all_items(self):
        """Select all items in the list"""
        if not self.items:
            return

        if not self.multiple_selection_mode:
            newest = self.items[-1]
            for item_id in list(self.selected_ids):
                self._set_selected(item_id, False)
            self._set_selected(newest["id"], True)
            self.copy_item_to_clipboard(newest["id"])
            self.show_toast(f"Item #{len(self.items)} selected (single mode)", "green")
        else:
            for rec in self.items:
                self._set_selected(rec["id"], True)

            self.copy_selected_items()
            self.show_toast(f"All {len(self.items)} items selected", "green")

        self.update_selection_label()
        self.update_preview()

    def deselect_all_items(self):
        """Deselect all items in the list"""
        for item_id in list(self.selected_ids):
            self._set_selected(item_id, False)

        self.update_selection_label()
        self.update_preview()
        self.show_toast("All items deselected", "gray")

    def delete_item(self, item_id):
        """Delete specific item from history (other selections are preserved)"""
        index, rec = self._lookup(item_id)
        if rec is None:
            return
        del self.items[index]
        self.selected_ids.discard(item_id)

        self.save_history()
        self.update_display()
        self.show_toast(f"Item #{index + 1} deleted", "red")

    def clear_history(self):
        """Clear all clipboard history (with confirmation)"""
        if not self.items:
            self.show_toast("History is already empty", "gray")
            return

        if not self._confirm(
            "Clear All",
            "This permanently deletes all clipboard history, including pinned items.",
            "CLEAR ALL"
        ):
            return

        self.items.clear()
        self.selected_ids.clear()
        self.save_history()
        self.update_display()
        self.show_toast("History cleared!", "red")

    # ------------------------------------------------------------------
    # Fun stuff
    # ------------------------------------------------------------------

    def clipboard_roulette(self):
        """Randomly re-copy a past clipboard item, just for fun"""
        if not self.items:
            self.show_toast("Nothing to roll — clipboard's empty!", "gray")
            return
        index = random.randrange(len(self.items))
        self.copy_item_to_clipboard(self.items[index]["id"])
        message = random.choice(self.ROULETTE_MESSAGES).format(n=index + 1)
        self.show_toast(message, "purple")

    def open_quick_guide_dialog(self):
        """Show a brief how-to-use popup"""
        dialog = ctk.CTkToplevel(self.window)
        dialog.title("Quick Guide")
        dialog.geometry("380x400")
        dialog.transient(self.window)
        dialog.grab_set()
        if self.topmost_var.get():
            dialog.attributes('-topmost', True)

        title = ctk.CTkLabel(
            dialog, text=self._track("Quick Guide"), font=ctk.CTkFont(size=13, weight="bold")
        )
        title.pack(pady=(15, 10))

        guide_text = (
            "• Copy text or files (Ctrl+C) — they're captured automatically\n\n"
            "• Multi-Select mode: check several items, then hit\n"
            "  \"Copy Selected\" to combine them onto the clipboard\n\n"
            "• Single-Select mode: checking an item copies it immediately\n\n"
            "• 📌 Pin an item to protect it from being evicted once you\n"
            f"  hit the {self.max_items}-item limit\n\n"
            "• Ctrl+F focuses search, Escape clears it"
        )
        label = ctk.CTkLabel(
            dialog, text=guide_text, font=ctk.CTkFont(size=12), justify="left"
        )
        label.pack(pady=10, padx=20)

        close_button = ctk.CTkButton(dialog, text="Got it", command=dialog.destroy)
        close_button.pack(pady=10)

    # ------------------------------------------------------------------
    # Updates
    # ------------------------------------------------------------------

    def check_for_updates_manual(self):
        """Manual update check triggered by the Updates button. Always
        shows a result, unlike the silent startup check."""
        if not _UPDATER_AVAILABLE:
            self.show_toast(
                "Updater not available — install 'requests' (pip install requests)",
                "orange"
            )
            return

        self.show_toast("Checking for updates...", "blue")
        try:
            updater.check_for_updates_async(self.window, silent=False)
        except Exception as e:
            self.show_toast(f"Update check failed: {e}", "red")
            
    def open_stats_dialog(self):
        """Show lifetime clipboard stats, with a silly novel/tweet comparison"""
        dialog = ctk.CTkToplevel(self.window)
        dialog.title("Clipboard Stats")
        dialog.geometry("340x340")
        dialog.transient(self.window)
        dialog.grab_set()
        if self.topmost_var.get():
            dialog.attributes('-topmost', True)

        title = ctk.CTkLabel(
            dialog, text=self._track("Clipboard Stats"), font=ctk.CTkFont(size=13, weight="bold")
        )
        title.pack(pady=(15, 10))

        total_items = self.stats["total_items"]
        total_chars = self.stats["total_characters"]
        novel_pct = (total_chars / 500_000) * 100  # ~500k chars ≈ an average novel
        tweets = total_chars // 280
        pinned_count = sum(1 for r in self.items if r["pinned"])

        stats_text = (
            f"ITEMS COPIED   {total_items:>10,}\n"
            f"CHARACTERS     {total_chars:>10,}\n\n"
            f"≈ {novel_pct:.3f}% OF AN AVERAGE NOVEL\n"
            f"≈ {tweets:,} TWEETS OF TEXT\n\n"
            f"IN HISTORY     {len(self.items):>10}\n"
            f"PINNED         {pinned_count:>10}"
        )
        label = ctk.CTkLabel(
            dialog, text=stats_text, font=self._mono(12), justify="left"
        )
        label.pack(pady=10, padx=20)

        close_button = ctk.CTkButton(dialog, text="Nice!", command=dialog.destroy)
        close_button.pack(pady=10)

    # ------------------------------------------------------------------
    # Toasts
    # ------------------------------------------------------------------

    def show_toast(self, message, color="green", duration=1800):
        """Show a temporary toast message, stacking with any already visible"""
        accent = T.TOAST.get(color, T.TEXT)
        slot = len(self.active_toasts)
        toast = ctk.CTkFrame(
            self.window, fg_color=T.PANEL, border_width=1, border_color=accent,
            width=470, height=34
        )
        toast.pack_propagate(False)
        ctk.CTkLabel(
            toast, text="●", width=18, font=ctk.CTkFont(size=10), text_color=accent
        ).pack(side="left", padx=(10, 0))
        ctk.CTkLabel(
            toast, text=message, anchor="w", font=self._mono(11), text_color=T.TEXT
        ).pack(side="left", fill="x", expand=True, padx=(4, 10))
        toast.place(relx=0.5, rely=0.9 - (slot * 0.07), anchor="center")
        self.active_toasts.append(toast)

        def remove():
            if toast in self.active_toasts:
                self.active_toasts.remove(toast)
            toast.destroy()
            for i, t in enumerate(self.active_toasts):
                t.place(relx=0.5, rely=0.9 - (i * 0.07), anchor="center")

        self.window.after(duration, remove)

    # ------------------------------------------------------------------
    # Settings
    # ------------------------------------------------------------------

    def open_max_items_dialog(self):
        """Open dialog to set maximum items"""
        dialog = ctk.CTkToplevel(self.window)
        dialog.title("Set Max Items")
        dialog.geometry("300x180")
        dialog.transient(self.window)
        dialog.grab_set()

        if self.topmost_var.get():
            dialog.attributes('-topmost', True)

        label = ctk.CTkLabel(dialog, text="Set maximum items (1-100):", font=ctk.CTkFont(size=14))
        label.pack(pady=10)

        entry = ctk.CTkEntry(dialog, width=100, placeholder_text=str(self.max_items))
        entry.pack(pady=10)

        def set_max_items():
            try:
                new_max = int(entry.get())
                if 1 <= new_max <= 100:
                    self.max_items = new_max

                    before = len(self.items)
                    self._enforce_capacity(announce=False)
                    if len(self.items) < before:
                        self.show_toast(f"Kept {len(self.items)} items (pinned items protected)", "orange")

                    self.save_history()
                    dialog.destroy()
                    self.update_display()
                    self.show_toast(f"Max items set to {new_max}", "green")
                else:
                    entry.delete(0, 'end')
                    entry.insert(0, "Invalid!")
            except ValueError:
                entry.delete(0, 'end')
                entry.insert(0, "Enter number!")

        button = ctk.CTkButton(dialog, text="Set", command=set_max_items)
        button.pack(pady=10)

    # ------------------------------------------------------------------
    # Lifecycle
    # ------------------------------------------------------------------

    def run(self):
        """Start the application"""
        self.window.mainloop()

    def stop_monitoring(self):
        """Stop clipboard monitoring"""
        self.monitoring = False


if __name__ == "__main__":
    import faulthandler
    _crash_log = open(Path.home() / ".clipboard_manager_crash.log", "a")
    faulthandler.enable(file=_crash_log, all_threads=True)
    app = ClipboardManager()
    try:
        app.run()
    finally:
        app.stop_monitoring()