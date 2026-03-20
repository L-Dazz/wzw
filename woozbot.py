"""
wzwsprite.py — woozbot auto-farmer.
Two-stage detection: shape (template matching) → color (H+S histogram).
Cross-platform: macOS (Retina) + Windows (DPI scaling).
"""

from __future__ import annotations

import json
import os
import platform
import random
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional, Tuple

import threading

import cv2
import numpy as np
import mss
import pyautogui
from pynput import keyboard as pynput_keyboard

# ---------------------------------------------------------------------------
# pyautogui safety
# ---------------------------------------------------------------------------
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.0

SETTINGS_FILE = Path("settings.json")
IS_MAC = platform.system() == "Darwin"

# ============================================================================
# ASCII banner
# ============================================================================

SHIMSHIM_BANNER = r"""
  ╔══════════════════════════════════════╗
  ║   hi shimshim ♥  woozbot is here    ║
  ║                                      ║
  ║    /\_____/\                         ║
  ║   ( =^ω^= )    screen recording?    ║
  ║    (  ♥  )     accessibility?       ║
  ║   づ     ヽ    both? good job~~     ║
  ╚══════════════════════════════════════╝
"""

# ============================================================================
# Config
# ============================================================================

@dataclass
class Config:
    sprites_dir: str = "sprites"
    template_paths: List[str] = field(default_factory=list)
    shape_threshold: float = 0.60
    color_threshold: float = 0.35
    scale_min: float = 0.15
    scale_max: float = 0.60
    scale_step: float = 0.05
    roi: Optional[dict] = None           # {"x":0,"y":0,"w":1920,"h":1080}
    min_click_distance: int = 30
    click_cooldown: float = 1.0
    click_times: int = 2                 # clicks per sprite
    click_gap_min: float = 0.2           # min seconds between repeated clicks on same sprite
    click_gap_max: float = 1.4           # max seconds between repeated clicks on same sprite
    between_clicks_min: float = 0.3      # min delay before moving to next sprite
    between_clicks_max: float = 0.8      # max delay before moving to next sprite
    loop_delay: float = 0.5
    randomize: bool = True
    random_offset_px: int = 5
    random_delay_min: float = 0.05
    random_delay_max: float = 0.20
    mouse_move_duration: float = 0.15
    debug: bool = False

    @classmethod
    def load(cls, path: Path = SETTINGS_FILE) -> "Config":
        if path.exists():
            with open(path) as f:
                data = json.load(f)
            c = cls()
            for k, v in data.items():
                if not hasattr(c, k):
                    continue
                # Expand ~ in any string path fields loaded from disk.
                if k in ("sprites_dir",) and isinstance(v, str):
                    v = str(Path(v).expanduser())
                if k == "template_paths" and isinstance(v, list):
                    v = [str(Path(p).expanduser()) for p in v]
                setattr(c, k, v)
            return c
        return cls()

    def save(self, path: Path = SETTINGS_FILE) -> None:
        with open(path, "w") as f:
            json.dump(self.__dict__, f, indent=2)
        print(f"  saved → {path}")

# ============================================================================
# PlatformHelper
# ============================================================================

class PlatformHelper:
    """Detect OS, DPI/Retina scale, and print permission hints."""

    def __init__(self, silent: bool = False) -> None:
        self.os: str = platform.system()   # 'Darwin' | 'Windows' | 'Linux'
        self.scale: float = self._detect_scale()
        if not silent:
            print(f"  [{self.os}] scale: {self.scale}")

    def _detect_scale(self) -> float:
        if platform.system() == "Darwin":
            # Primary: Quartz — works without subprocess, no display needed
            try:
                from Quartz import CGMainDisplayID, CGDisplayPixelsWide, CGDisplayBounds
                did = CGMainDisplayID()
                logical_w = CGDisplayBounds(did).size.width
                physical_w = CGDisplayPixelsWide(did)
                if logical_w > 0:
                    return physical_w / logical_w
            except Exception:
                pass
            # Fallback: system_profiler text scan
            try:
                import subprocess
                r = subprocess.run(
                    ["system_profiler", "SPDisplaysDataType"],
                    capture_output=True, text=True, timeout=5,
                )
                if "Retina" in r.stdout:
                    return 2.0
            except Exception:
                pass
            return 1.0
        elif platform.system() == "Windows":
            try:
                import ctypes
                ctypes.windll.shcore.SetProcessDpiAwareness(2)
                dpi = ctypes.windll.user32.GetDpiForSystem()
                return dpi / 96.0
            except Exception:
                return 1.0
        return 1.0

    def physical_to_logical(self, x: int, y: int) -> Tuple[int, int]:
        return int(x / self.scale), int(y / self.scale)

    @staticmethod
    def check_macos_permissions() -> bool:
        """
        Actively probe screen recording and accessibility on macOS.
        Returns True if both look okay, False if something seems missing.
        Prints a clear actionable block either way.
        """
        if platform.system() != "Darwin":
            return True

        print(SHIMSHIM_BANNER)
        print("  checking mac permissions...\n")

        screen_ok = False
        access_ok = False

        # --- screen recording: try to grab 1 pixel ---
        try:
            with mss.mss() as sct:
                # Grab a 10×10 block — large enough to detect an all-black
                # frame that macOS returns silently when permission is denied.
                mon = {"left": 0, "top": 0, "width": 10, "height": 10}
                raw = np.array(sct.grab(mon))
            # If the entire capture is zero the permission is blocked.
            # A real screen will almost never be pure black across 100 pixels.
            if raw[:, :, :3].max() == 0:
                screen_ok = False
            else:
                screen_ok = True
        except Exception:
            screen_ok = False

        # --- accessibility: pyautogui position doesn't raise, but pynput does ---
        # We can't actually verify AX at startup without triggering the dialog;
        # we mark it as "assumed OK" unless we've seen a pynput failure.
        # The real test happens at first click — so we just warn.
        access_ok = True   # optimistic; will blow up on first click if missing

        def _tick(ok: bool) -> str:
            return "✓" if ok else "✗"

        print(f"  {_tick(screen_ok)} screen recording")
        print(f"  {_tick(access_ok)} accessibility  (can't verify until first click)")

        if not screen_ok:
            print("""
  ── action required ────────────────────────────────────────────
  screen recording is BLOCKED — woozbot can't see the game.

  fix it:
    1. open System Settings (apple menu → System Settings)
    2. Privacy & Security → Screen Recording
    3. find Terminal (or iTerm / your app) → toggle ON
    4. it will ask you to quit Terminal — do it, then rerun

  if Terminal is already listed and ON:
    toggle it OFF, wait 2 sec, toggle back ON, quit & reopen
  ────────────────────────────────────────────────────────────────
""")
            return False

        print("""
  accessibility note:
    if clicks don't register, open System Settings →
    Privacy & Security → Accessibility → add Terminal → ON
    then quit and reopen Terminal.
""")
        return True

    @staticmethod
    def print_macos_reminder() -> None:
        """Short reminder printed during setup on Darwin."""
        print("""
  ── mac permissions ────────────────────────────────────────────
  required (or things silently break):

    screen recording  →  System Settings → Privacy & Security
                         → Screen Recording → Terminal → ON

    accessibility     →  System Settings → Privacy & Security
                         → Accessibility → Terminal → ON

  restart Terminal after changing either one.
  ────────────────────────────────────────────────────────────────""")


# ============================================================================
# ScreenCapture
# ============================================================================

class ScreenCapture:
    """Grab screen (or ROI) using mss, return BGR numpy array."""

    def __init__(self, roi, scale: float) -> None:
        self._sct = mss.mss()
        if isinstance(roi, (tuple, list)) and len(roi) == 4:
            self.roi = {"x": roi[0], "y": roi[1], "w": roi[2], "h": roi[3]}
        else:
            self.roi = roi
        self.scale = scale

    def grab(self) -> np.ndarray:
        if self.roi:
            monitor = {
                "left": self.roi["x"], "top": self.roi["y"],
                "width": self.roi["w"], "height": self.roi["h"],
            }
        else:
            # monitors[0] is the combined virtual display (all screens stitched).
            # monitors[1] is the primary physical display.
            # Guard against setups where mss returns fewer entries than expected.
            monitors = self._sct.monitors
            if len(monitors) >= 2:
                monitor = monitors[1]
            elif len(monitors) == 1:
                print("  [warn] mss only found the virtual monitor — using monitors[0] (full combined)")
                monitor = monitors[0]
            else:
                raise RuntimeError("mss returned no monitors — check display connection and Screen Recording permission")
        img = self._sct.grab(monitor)
        return np.array(img)[:, :, :3]   # BGRA → BGR

    @property
    def roi_offset(self) -> Tuple[int, int]:
        return (self.roi["x"], self.roi["y"]) if self.roi else (0, 0)


# ============================================================================
# TemplateEntry
# ============================================================================

@dataclass
class TemplateEntry:
    path: str
    bgr: np.ndarray
    gray: np.ndarray
    mask: Optional[np.ndarray]
    hsv_hist: np.ndarray

    @classmethod
    def load(cls, path: str) -> "TemplateEntry":
        img_bgra = cv2.imread(path, cv2.IMREAD_UNCHANGED)
        if img_bgra is None:
            raise FileNotFoundError(f"cannot load template: {path}")
        if img_bgra.ndim == 3 and img_bgra.shape[2] == 4:
            alpha = img_bgra[:, :, 3]
            mask = (alpha > 10).astype(np.uint8) * 255
            bgr = cv2.cvtColor(img_bgra, cv2.COLOR_BGRA2BGR)
        else:
            mask = None
            bgr = img_bgra if img_bgra.ndim == 3 else cv2.cvtColor(img_bgra, cv2.COLOR_GRAY2BGR)
        gray = cv2.cvtColor(bgr, cv2.COLOR_BGR2GRAY)
        hsv_hist = cls._compute_hs_hist(bgr, mask)
        return cls(path=path, bgr=bgr, gray=gray, mask=mask, hsv_hist=hsv_hist)

    @staticmethod
    def _compute_hs_hist(bgr: np.ndarray, mask: Optional[np.ndarray]) -> np.ndarray:
        hsv = cv2.cvtColor(bgr, cv2.COLOR_BGR2HSV)
        hist = cv2.calcHist([hsv], [0, 1], mask, [30, 32], [0, 180, 0, 256])
        cv2.normalize(hist, hist, 0, 1, cv2.NORM_MINMAX)
        return hist


# ============================================================================
# NMS
# ============================================================================

def _iou(a: Tuple, b: Tuple) -> float:
    ax, ay, aw, ah = a[:4]
    bx, by, bw, bh = b[:4]
    ix = max(0, min(ax + aw, bx + bw) - max(ax, bx))
    iy = max(0, min(ay + ah, by + bh) - max(ay, by))
    inter = ix * iy
    union = aw * ah + bw * bh - inter
    return inter / union if union > 0 else 0.0

def non_max_suppression(
    boxes: List[Tuple[int, int, int, int, float]],
    iou_threshold: float = 0.3,
) -> List[Tuple[int, int, int, int, float]]:
    if not boxes:
        return []
    boxes_sorted = sorted(boxes, key=lambda b: b[4], reverse=True)
    kept: List[Tuple[int, int, int, int, float]] = []
    while boxes_sorted:
        best = boxes_sorted.pop(0)
        kept.append(best)
        boxes_sorted = [b for b in boxes_sorted if _iou(best, b) < iou_threshold]
    return kept


# ============================================================================
# Detector
# ============================================================================

@dataclass
class Detection:
    x: int; y: int; w: int; h: int
    shape_score: float
    color_dist: float
    passed_color: bool
    template_path: str

class Detector:
    def __init__(self, cfg: Config, ph: PlatformHelper) -> None:
        self.cfg = cfg
        self.scale = ph.scale
        self.templates: List[TemplateEntry] = []
        self._load_templates()

    def _load_templates(self) -> None:
        paths = list(self.cfg.template_paths)
        sprites_dir = Path(self.cfg.sprites_dir)
        if sprites_dir.is_dir():
            for p in sprites_dir.glob("*.png"):
                if str(p) not in paths:
                    paths.append(str(p))
        for p in paths:
            try:
                self.templates.append(TemplateEntry.load(p))
                print(f"  loaded template: {p}")
            except Exception as e:
                print(f"  [warn] skip {p}: {e}")

    def detect(self, frame_bgr: np.ndarray, roi_offset: Tuple[int, int] = (0, 0)) -> List[Detection]:
        results: List[Detection] = []
        frame_gray = cv2.cvtColor(frame_bgr, cv2.COLOR_BGR2GRAY)
        ox, oy = roi_offset
        for tmpl in self.templates:
            for (rx, ry, rw, rh, score) in self._stage1_shape(frame_gray, tmpl):
                crop = frame_bgr[ry:ry + rh, rx:rx + rw]
                dist = self._stage2_color(crop, tmpl)
                passed = dist <= self.cfg.color_threshold
                results.append(Detection(
                    x=int((rx + ox) / self.scale),
                    y=int((ry + oy) / self.scale),
                    w=int(rw / self.scale),
                    h=int(rh / self.scale),
                    shape_score=score,
                    color_dist=dist,
                    passed_color=passed,
                    template_path=tmpl.path,
                ))
        return results

    def _stage1_shape(self, frame_gray: np.ndarray, tmpl: TemplateEntry) -> List[Tuple]:
        th_orig, tw_orig = tmpl.gray.shape[:2]
        candidates: List[Tuple] = []
        scale = self.cfg.scale_min
        while scale <= self.cfg.scale_max + 1e-6:
            tw = max(1, int(tw_orig * scale))
            th = max(1, int(th_orig * scale))
            if tw > frame_gray.shape[1] or th > frame_gray.shape[0]:
                scale = round(scale + self.cfg.scale_step, 4)
                continue
            resized_gray = cv2.resize(tmpl.gray, (tw, th))
            resized_mask = (cv2.resize(tmpl.mask, (tw, th), interpolation=cv2.INTER_NEAREST)
                            if tmpl.mask is not None else None)
            try:
                result = cv2.matchTemplate(frame_gray, resized_gray, cv2.TM_CCOEFF_NORMED, mask=resized_mask)
            except cv2.error:
                scale = round(scale + self.cfg.scale_step, 4)
                continue
            for (py, px) in zip(*np.where(result >= self.cfg.shape_threshold)):
                candidates.append((int(px), int(py), int(tw), int(th), float(result[py, px])))
            scale = round(scale + self.cfg.scale_step, 4)
        return non_max_suppression(candidates)

    def _stage2_color(self, crop_bgr: np.ndarray, tmpl: TemplateEntry) -> float:
        if crop_bgr.size == 0:
            return 999.0
        resized = cv2.resize(crop_bgr, (tmpl.bgr.shape[1], tmpl.bgr.shape[0]))
        hist = TemplateEntry._compute_hs_hist(resized, tmpl.mask)
        return float(cv2.compareHist(tmpl.hsv_hist, hist, cv2.HISTCMP_BHATTACHARYYA))


# ============================================================================
# Clicker
# ============================================================================

class Clicker:
    def __init__(self, cfg: Config) -> None:
        self.cfg = cfg
        self._cooldown_map: dict = {}

    def click(self, det: Detection) -> bool:
        cx = det.x + det.w // 2
        cy = det.y + det.h // 2
        bucket = self._bucket(cx, cy)
        now = time.time()
        if bucket in self._cooldown_map:
            if now - self._cooldown_map[bucket] < self.cfg.click_cooldown:
                return False

        if self.cfg.randomize:
            time.sleep(random.uniform(self.cfg.random_delay_min, self.cfg.random_delay_max))

        for i in range(max(1, self.cfg.click_times)):
            jx = cx + (random.randint(-self.cfg.random_offset_px, self.cfg.random_offset_px) if self.cfg.randomize else 0)
            jy = cy + (random.randint(-self.cfg.random_offset_px, self.cfg.random_offset_px) if self.cfg.randomize else 0)
            try:
                if i == 0:
                    pyautogui.moveTo(jx, jy, duration=self.cfg.mouse_move_duration, tween=pyautogui.easeInOutQuad)
                else:
                    pyautogui.moveTo(jx, jy, duration=0.05)
                pyautogui.click()
                print(f"  click {i+1}/{self.cfg.click_times} ({jx},{jy})  shape={det.shape_score:.2f}  dist={det.color_dist:.3f}")
            except pyautogui.FailSafeException:
                print("  [failsafe] mouse hit corner — stopping")
                raise SystemExit(1)
            except Exception as e:
                # On macOS this path fires when Accessibility is blocked.
                # pyautogui normally doesn't raise here, but CGEvent failures
                # can surface as exceptions from underlying calls.
                print(f"  [ERROR] click failed: {e}")
                if IS_MAC:
                    print(
                        "  ── accessibility blocked? ──────────────────────────────────\n"
                        "  System Settings → Privacy & Security → Accessibility\n"
                        "  → add Terminal (or your app) → toggle ON → restart Terminal\n"
                        "  ─────────────────────────────────────────────────────────────"
                    )
            if i < self.cfg.click_times - 1:
                time.sleep(random.uniform(self.cfg.click_gap_min, self.cfg.click_gap_max))

        self._cooldown_map[bucket] = time.time()
        return True

    def _bucket(self, x: int, y: int) -> Tuple[int, int]:
        d = max(1, self.cfg.min_click_distance)
        return (x // d, y // d)


# ============================================================================
# Visualiser
# ============================================================================

class Visualiser:
    WINDOW = "woozbot [debug]"
    FALLBACK_FILE = "debug_out.png"

    def __init__(self, scale: float = 1.0) -> None:
        self._scale = scale
        # cv2.imshow is unreliable on macOS (requires main-thread Cocoa event
        # loop, crashes in Terminal sessions).  Force file fallback immediately.
        if IS_MAC:
            self._use_window: Optional[bool] = False
            print(f"  [debug] macOS — skipping live window, writing {self.FALLBACK_FILE} each frame")
        else:
            self._use_window = None   # probe on first show()
            try:
                cv2.namedWindow(self.WINDOW, cv2.WINDOW_NORMAL | cv2.WINDOW_KEEPRATIO)
                cv2.resizeWindow(self.WINDOW, 960, 540)
            except Exception:
                pass

    def _draw(self, frame: np.ndarray, detections: List[Detection], roi_offset: Tuple[int, int]) -> np.ndarray:
        vis = frame.copy()
        ox, oy = roi_offset
        for d in detections:
            rx = int(d.x * self._scale) - ox
            ry = int(d.y * self._scale) - oy
            rw = int(d.w * self._scale)
            rh = int(d.h * self._scale)
            color = (0, 255, 0) if d.passed_color else (0, 255, 255)
            cv2.rectangle(vis, (rx, ry), (rx + rw, ry + rh), color, 2)
            cv2.putText(vis, f"{d.shape_score:.2f}/{d.color_dist:.3f}",
                        (rx, max(12, ry - 4)), cv2.FONT_HERSHEY_SIMPLEX, 0.45, color, 1, cv2.LINE_AA)
        cv2.putText(vis, "GREEN=click  YELLOW=color fail",
                    (8, 18), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 255), 1)
        return vis

    def show(self, frame: np.ndarray, detections: List[Detection], roi_offset: Tuple[int, int]) -> None:
        vis = self._draw(frame, detections, roi_offset)
        if self._use_window is None:
            try:
                cv2.imshow(self.WINDOW, vis)
                cv2.waitKey(30)
                prop = cv2.getWindowProperty(self.WINDOW, cv2.WND_PROP_VISIBLE)
                self._use_window = prop >= 0
                if not self._use_window:
                    cv2.destroyAllWindows()
                    print(f"  [debug] live window unavailable — writing {self.FALLBACK_FILE} each frame")
            except Exception:
                self._use_window = False
                print(f"  [debug] live window unavailable — writing {self.FALLBACK_FILE} each frame")
        if self._use_window:
            cv2.imshow(self.WINDOW, vis)
            cv2.waitKey(1)
        else:
            cv2.imwrite(self.FALLBACK_FILE, vis)

    def close(self) -> None:
        if self._use_window:
            cv2.destroyAllWindows()


# ============================================================================
# CLI helpers  (all robust: retry on bad input, skip option everywhere)
# ============================================================================

def yn(prompt: str, default: bool = True) -> bool:
    """y/n prompt with retry on garbage input."""
    suffix = " [Y/n]: " if default else " [y/N]: "
    while True:
        try:
            r = input(prompt + suffix).strip().lower()
        except (EOFError, KeyboardInterrupt):
            print()
            return default
        if r == "":
            return default
        if r in ("y", "yes"):
            return True
        if r in ("n", "no"):
            return False
        print("  please enter y or n")


def ask(prompt: str, default: str = "") -> str:
    suffix = f" [{default}]: " if default else ": "
    try:
        r = input(prompt + suffix).strip()
    except (EOFError, KeyboardInterrupt):
        print()
        return default
    return r if r else default


def ask_float(prompt: str, default: float, min_val: float = 0.0, max_val: float = 1e9) -> float:
    """Ask for a float, retry until valid, show default, enforce range."""
    while True:
        raw = ask(prompt, str(default))
        try:
            v = float(raw)
        except ValueError:
            print(f"  invalid — enter a number (e.g. {default})")
            continue
        if not (min_val <= v <= max_val):
            print(f"  out of range — must be between {min_val} and {max_val}")
            continue
        return v


def ask_int(prompt: str, default: int, min_val: int = 0, max_val: int = 9999) -> int:
    while True:
        raw = ask(prompt, str(default))
        try:
            v = int(raw)
        except ValueError:
            print(f"  invalid — enter a whole number (e.g. {default})")
            continue
        if not (min_val <= v <= max_val):
            print(f"  out of range — must be between {min_val} and {max_val}")
            continue
        return v


def ask_range(prompt_base: str, default_min: float, default_max: float) -> Tuple[float, float]:
    """Ask for a min/max pair, ensuring min <= max. Retries the whole pair on bad order."""
    while True:
        lo = ask_float(f"  {prompt_base} min", default_min, 0.0)
        hi = ask_float(f"  {prompt_base} max", default_max, 0.0)
        if lo > hi:
            print(f"  min ({lo}) can't be greater than max ({hi}) — try again")
            continue
        return lo, hi


# ============================================================================
# ROI selector
# ============================================================================

def _manual_roi_entry() -> Optional[Tuple[int, int, int, int]]:
    """Prompt for ROI coordinates by hand, with retry on bad input."""
    print("  enter ROI as four numbers: x y w h  (physical screen pixels)")
    print("  tip: hover over game window corners in Screenshot app to get coords")
    print("  press ENTER with nothing to skip and use full screen")
    while True:
        try:
            raw = input("  ROI: ").strip()
        except (EOFError, KeyboardInterrupt):
            print()
            return None
        if not raw:
            print("  skipping ROI — using full screen (may be slower on 4K)")
            return None
        parts = raw.split()
        if len(parts) != 4:
            print(f"  got {len(parts)} value(s), need exactly 4 — try again or press ENTER to skip")
            continue
        try:
            x, y, w, h = (int(p) for p in parts)
        except ValueError:
            print("  couldn't parse those as integers — try again")
            continue
        if w <= 0 or h <= 0:
            print("  width and height must be > 0 — try again")
            continue
        print(f"  ROI set: x={x} y={y} w={w} h={h}")
        return (x, y, w, h)


def select_roi(ph: PlatformHelper) -> Optional[Tuple[int, int, int, int]]:
    """
    Drag-to-select ROI. Falls back to manual entry if display isn't usable
    (macOS without Screen Recording, SSH sessions, etc.).
    C to cancel in the selector now falls back to manual entry instead of
    returning None silently.
    """
    # ── display probe ────────────────────────────────────────────────────────
    if IS_MAC:
        display_ok = False
        try:
            cv2.namedWindow("__probe__", cv2.WINDOW_NORMAL)
            cv2.waitKey(1)
            cv2.destroyWindow("__probe__")
            display_ok = True
        except Exception:
            pass
        if not display_ok:
            print("  [warn] display not accessible — switching to manual entry")
            return _manual_roi_entry()

    # ── screen grab ──────────────────────────────────────────────────────────
    print("  grabbing screen for ROI selector...")
    try:
        with mss.mss() as sct:
            raw = sct.grab(sct.monitors[1])
            img = cv2.cvtColor(np.array(raw), cv2.COLOR_BGRA2BGR)
    except Exception as e:
        print(f"  [warn] screen grab failed ({e}) — switching to manual entry")
        return _manual_roi_entry()

    # On Retina (macOS) mss returns physical pixels which are 2× logical size,
    # so we need to shrink the preview extra so it fits on screen.
    # On Windows the mss grab is already at logical resolution — use flat 0.6.
    if IS_MAC:
        preview_scale = max(0.15, min(0.6 / ph.scale, 0.8))
    else:
        preview_scale = 0.6
    small = cv2.resize(img, (0, 0), fx=preview_scale, fy=preview_scale)

    print("  drag a box around the game window")
    print("  ENTER / SPACE = confirm   C = cancel")
    try:
        rect = cv2.selectROI("select ROI — ENTER to confirm, C to cancel", small, showCrosshair=True)
        cv2.destroyAllWindows()
    except Exception as e:
        cv2.destroyAllWindows()
        print(f"  [warn] selector failed ({e}) — switching to manual entry")
        return _manual_roi_entry()

    # C pressed or empty drag → offer manual entry
    if rect == (0, 0, 0, 0):
        print("  nothing selected")
        if yn("  enter ROI coordinates manually instead?", default=True):
            return _manual_roi_entry()
        print("  using full screen")
        return None

    x = int(rect[0] / preview_scale)
    y = int(rect[1] / preview_scale)
    w = int(rect[2] / preview_scale)
    h = int(rect[3] / preview_scale)
    print(f"  ROI: x={x} y={y} w={w} h={h}")
    return (x, y, w, h)


# ============================================================================
# Setup wizard — individual steps  (all retry-safe)
# ============================================================================

def setup_os(cfg: Config) -> PlatformHelper:
    print("\n── platform ──────────────────────────────────────")
    ph = PlatformHelper(silent=True)
    print(f"  detected: {ph.os}  |  scale factor: {ph.scale}")
    if ph.os == "Darwin":
        PlatformHelper.print_macos_reminder()

    if not yn("  correct?"):
        while True:
            override = ask("  OS (Darwin / Windows / Linux)", ph.os)
            if override in ("Darwin", "Windows", "Linux"):
                ph.os = override
                break
            print("  enter Darwin, Windows, or Linux exactly")
        ph.scale = ask_float("  scale factor (1.0 normal, 2.0 retina)", ph.scale, 0.25, 8.0)

    return ph


def setup_sprites_dir(cfg: Config) -> None:
    print("\n── sprites ───────────────────────────────────────")
    while True:
        print(f"  current sprites dir: {cfg.sprites_dir}")
        choice = ask("  press ENTER to keep, or enter a new path", cfg.sprites_dir)
        cfg.sprites_dir = str(Path(choice).expanduser())
        p = Path(cfg.sprites_dir)
        if not p.exists():
            print(f"  '{cfg.sprites_dir}' doesn't exist")
            action = ask("  (c)reate it, (r)e-enter path, or (s)kip", "c").lower()
            if action.startswith("c"):
                p.mkdir(parents=True, exist_ok=True)
                print(f"  created {p.resolve()} — drop your sprite pngs in there")
                break
            elif action.startswith("r"):
                continue
            else:
                print("  skipping — make sure the path exists before running")
                break
        else:
            break

    pngs = list(p.glob("*.png")) if p.exists() else []
    if pngs:
        cfg.template_paths = [str(x) for x in pngs]
        print(f"  found {len(pngs)} template(s): {[x.name for x in pngs]}")
    else:
        cfg.template_paths = []
        print("  no pngs found — add some before running")


def setup_debug(cfg: Config) -> None:
    print("\n── debug ─────────────────────────────────────────")
    cfg.debug = yn("  enable debug window? (green=click, yellow=color fail)", default=False)


def setup_clicking(cfg: Config) -> None:
    print("\n── clicking ──────────────────────────────────────")
    print(f"  clicks per sprite:     {cfg.click_times}")
    print(f"  gap between clicks:    {cfg.click_gap_min}–{cfg.click_gap_max}s")
    print(f"  delay between sprites: {cfg.between_clicks_min}–{cfg.between_clicks_max}s")
    if not yn("  change clicking settings?", default=False):
        return

    cfg.click_times = ask_int("  clicks per sprite", cfg.click_times, 1, 20)

    print("  gap between repeated clicks on the same sprite:")
    cfg.click_gap_min, cfg.click_gap_max = ask_range(
        "click gap", cfg.click_gap_min, cfg.click_gap_max
    )

    print("  delay before moving to the next sprite:")
    cfg.between_clicks_min, cfg.between_clicks_max = ask_range(
        "between sprites", cfg.between_clicks_min, cfg.between_clicks_max
    )


def setup_roi(cfg: Config, ph: PlatformHelper) -> None:
    print("\n── ROI ───────────────────────────────────────────")
    if cfg.roi:
        r = cfg.roi
        print(f"  saved ROI: x={r['x']} y={r['y']} w={r['w']} h={r['h']}")
        if yn("  keep it?"):
            return
        # offer clear
        if yn("  clear ROI and use full screen?", default=False):
            cfg.roi = None
            print("  cleared — will scan full screen")
            return

    if yn("  select game window region now? (recommended)"):
        roi = select_roi(ph)
        if roi:
            cfg.roi = {"x": roi[0], "y": roi[1], "w": roi[2], "h": roi[3]}
        else:
            cfg.roi = None
    else:
        cfg.roi = None
        print("  using full screen — may be slow on 4K")


def run_wizard(existing: Optional[Config] = None) -> Tuple[Config, PlatformHelper]:
    cfg = existing if existing else Config()

    print("\n" + "=" * 52)
    print("  woozbot — setup")
    print("=" * 52)

    ph = setup_os(cfg)
    setup_sprites_dir(cfg)
    setup_debug(cfg)
    setup_clicking(cfg)
    setup_roi(cfg, ph)

    print("\n── summary ───────────────────────────────────────")
    print(f"  OS:             {ph.os} (scale {ph.scale})")
    print(f"  sprites:        {cfg.sprites_dir}  ({len(cfg.template_paths)} template(s))")
    print(f"  shape thresh:   {cfg.shape_threshold}")
    print(f"  color thresh:   {cfg.color_threshold}")
    print(f"  scale range:    {cfg.scale_min}–{cfg.scale_max}  step {cfg.scale_step}")
    print(f"  debug:          {cfg.debug}")
    print(f"  ROI:            {cfg.roi if cfg.roi else 'full screen'}")
    print(f"  randomize:      {cfg.randomize}")
    print(f"  clicks/sprite:  {cfg.click_times}  gap {cfg.click_gap_min}–{cfg.click_gap_max}s")
    print(f"  between sprites:{cfg.between_clicks_min}–{cfg.between_clicks_max}s")
    print()

    if yn("  save settings for next time?"):
        cfg.save()

    return cfg, ph


# ============================================================================
# SpriteClicker — main orchestrator
# ============================================================================

def _start_keyboard_listener(stop_event: threading.Event) -> Optional[pynput_keyboard.Listener]:
    """
    Start a pynput keyboard listener on a thread that owns a Win32 message loop.

    On Windows, pynput uses SetWindowsHookEx which requires the hook thread to
    be pumping Win32 messages.  Running from cmd.exe / PowerShell / Windows
    Terminal works fine; IDLE also works because it already has a message loop.
    The safe approach is to start the listener inside a daemon thread so pynput
    can create its own message loop there regardless of the parent environment.

    Returns the Listener if it started successfully, or None if pynput failed
    (in which case the caller falls back to msvcrt ESC polling on Windows).
    """
    listener_box: List[Optional[pynput_keyboard.Listener]] = [None]
    ready = threading.Event()
    error_box: List[Optional[Exception]] = [None]

    def _thread_main():
        def on_press(key):
            if key == pynput_keyboard.Key.esc:
                print("\n  [ESC] stopping...")
                stop_event.set()

        try:
            with pynput_keyboard.Listener(on_press=on_press) as listener:
                listener_box[0] = listener
                ready.set()
                listener.join()          # blocks until listener stops
        except Exception as e:
            error_box[0] = e
            ready.set()

    t = threading.Thread(target=_thread_main, daemon=True)
    t.start()
    ready.wait(timeout=3.0)

    if error_box[0] is not None:
        print(f"  [warn] keyboard listener failed ({error_box[0]}) — using fallback ESC polling")
        return None

    return listener_box[0]


def _poll_esc_msvcrt() -> bool:
    """Non-blocking ESC check via msvcrt (Windows only, no hook needed)."""
    try:
        import msvcrt
        if msvcrt.kbhit():
            ch = msvcrt.getwch()
            if ord(ch) == 27:   # ESC
                return True
    except Exception:
        pass
    return False


class SpriteClicker:
    def __init__(self, cfg: Config, ph: PlatformHelper) -> None:
        self.cfg = cfg
        self.ph = ph
        self._stop = threading.Event()
        self.capture = ScreenCapture(roi=cfg.roi, scale=ph.scale)
        self.detector = Detector(cfg, ph)
        self.clicker = Clicker(cfg)
        self.vis = Visualiser(scale=ph.scale) if cfg.debug else None
        self._listener = _start_keyboard_listener(self._stop)
        self._use_msvcrt = (self._listener is None and platform.system() == "Windows")
        self._pynput_failed_mac = (self._listener is None and IS_MAC)
        if self._pynput_failed_mac:
            print(
                "  [warn] pynput keyboard listener could not start on macOS.\n"
                "  This usually means Accessibility permission is missing.\n"
                "  ESC hotkey is unavailable — use Ctrl+C or mouse → top-left corner to stop.\n"
                "  To restore ESC: System Settings → Privacy & Security → Accessibility → Terminal → ON"
            )

    def run(self) -> None:
        if not self.detector.templates:
            print("\n  no templates loaded — add sprite pngs to your sprites folder and rerun\n")
            return

        if self._use_msvcrt:
            print("\n  running — press ESC to stop | mouse → top-left corner = emergency brake")
            print("  (pynput hook unavailable — using keyboard polling)\n")
        elif self._pynput_failed_mac:
            print("\n  running — Ctrl+C to stop | mouse → top-left corner = emergency brake")
            print("  (pynput failed — ESC not available, see warning above)\n")
        else:
            print("\n  running — ESC to stop | mouse → top-left corner = emergency brake\n")

        frames = 0
        t0 = time.time()

        try:
            while not self._stop.is_set():
                # fallback ESC check (when pynput hook couldn't start)
                if self._use_msvcrt and _poll_esc_msvcrt():
                    print("\n  [ESC] stopping...")
                    break

                frame = self.capture.grab()
                detections = self.detector.detect(frame, self.capture.roi_offset)
                passed = [d for d in detections if d.passed_color]
                if self.cfg.randomize:
                    random.shuffle(passed)

                for i, d in enumerate(passed):
                    if self._stop.is_set():
                        break
                    if self.clicker.click(d):
                        if i < len(passed) - 1:
                            time.sleep(random.uniform(
                                self.cfg.between_clicks_min,
                                self.cfg.between_clicks_max,
                            ))

                if self.vis:
                    self.vis.show(frame, detections, self.capture.roi_offset)

                if detections:
                    for d in detections:
                        status = "CLICK" if d.passed_color else "skip"
                        name = Path(d.template_path).stem
                        print(f"  [{status}] {name}  shape={d.shape_score:.2f}  dist={d.color_dist:.3f}  pos=({d.x},{d.y})")

                frames += 1
                # On macOS, mss capture is slower and a zero delay causes CPU
                # spikes. Enforce a minimum without modifying the saved config.
                effective_delay = self.cfg.loop_delay
                if IS_MAC:
                    effective_delay = max(effective_delay, 0.1)
                time.sleep(effective_delay)

        except KeyboardInterrupt:
            pass
        except pyautogui.FailSafeException:
            print("  [failsafe] mouse hit corner — stopped")
        finally:
            elapsed = time.time() - t0
            fps = frames / max(elapsed, 1)
            print(f"\n  stopped — {frames} frames in {elapsed:.1f}s ({fps:.1f} fps)")
            if self.vis:
                self.vis.close()
            if self._listener is not None:
                self._listener.stop()


# ============================================================================
# Entry point
# ============================================================================

if __name__ == "__main__":
    existing = Config.load() if SETTINGS_FILE.exists() else None

    if existing:
        print("\n" + "=" * 52)
        print("  woozbot — woozworld auto-farmer")
        print("=" * 52)
        print(f"  saved settings found")
        print(f"  sprites:  {existing.sprites_dir}  ({len(existing.template_paths)} template(s))")
        print(f"  debug:    {existing.debug}")
        print(f"  ROI:      {existing.roi if existing.roi else 'full screen'}")
        print()

        if yn("  use saved settings?"):
            cfg = existing
            ph = PlatformHelper(silent=True)
            print(f"\n  detected: {ph.os}  |  scale: {ph.scale}")
            if ph.os == "Darwin":
                PlatformHelper.check_macos_permissions()
            if not yn("  platform looks correct?"):
                while True:
                    override = ask("  OS (Darwin / Windows / Linux)", ph.os)
                    if override in ("Darwin", "Windows", "Linux"):
                        ph.os = override
                        break
                    print("  enter Darwin, Windows, or Linux exactly")
                ph.scale = ask_float("  scale factor", ph.scale, 0.25, 8.0)
        else:
            cfg, ph = run_wizard(existing)
    else:
        # fresh run — check mac permissions before the wizard
        if IS_MAC:
            ph_check = PlatformHelper(silent=True)
            ok = PlatformHelper.check_macos_permissions()
            if not ok:
                print("  fix screen recording first, then rerun")
                sys.exit(1)
        cfg, ph = run_wizard()

    print("\n  starting in 3 seconds — switch to the game window\n")
    time.sleep(3)

    SpriteClicker(cfg, ph).run()
