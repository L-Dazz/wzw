# woozbot

auto-farms wood and ore in wzw. really anything clickable. two-stage detection.
shape matching first, then color histogram to filter false positives. works on mac and windows etc etc

---

## how it works

1. grabs the screen fast using mss
2. runs multi-scale template matching in grayscale (stage1 shape)
3. filters matches by color histogram similarity (stage2 color)
4. clicks everything that passes both stages
5. tracks cooldowns so it doesn't spam the same spot
6. ESC to stop. mouse to top-left corner = ebrake

two-stage is slower than single-stage but you get way fewer false positives on sprites that share a shape with background elements etc etc

---

## setup

```bash
pip install -r requirements.txt
python wzwsprite.py
```

first run is the walk thru run
will detect (should detect) OS, sprites folder (or create), debug mode toggle (kind of ahh), clicking settings, ROI (region of interest). will save to settings.json 

---

## sprites

when it asks, press enter to use the default (sprites/) folder or give it another path. it'll create the folder if it doesn't exist

move your sprite as pngs in there. just the sprite, minimal background around is like 90% of whether detection works.
(currently only supports 1, well, actually, try as many as you want. i didn't test)

---

## mac permissions (actually required, not optional or tested, figure it out.)

1. system settings → privacy & security → screen recording → terminal → ON
2. system settings → privacy & security → accessibility → terminal → ON
3. restart terminal after both

script SHOULD check screen recording on startup and tell you what to do if it's blocked.
accessibility it can't verify until first click so if clicks register on wrong spots or don't register at all that may be why

---

## windows

just works (trust). 
HOWEVER 
if clicks are offset you can right click python → properties → compatibility → high dpi settings → override → application. but it should autodetect this.

---

## ROI

during setup it asks if you want to select a region. you should. drag a box around the game window or area you can collect in. scanning full screen on 4k is noticeably slower and you might get false positives from other apps.

if the drag selector doesn't work (headless, mac permission issue etc) it falls back to manual coordinate entry

---

## config (settings.json)

| setting | default | what |
|---|---|---|
| shape_threshold | 0.60 | how strict stage 1 shape matching is. lower = more candidates |
| color_threshold | 0.35 | max histogram distance for stage 2. lower = stricter color match |
| scale_min/max/step | 0.15-0.60 | template resize range for multi-scale matching |
| click_times | 2 | clicks per sprite |
| click_cooldown | 1.0 | seconds before re-clicking same spot |
| loop_delay | 0.5 | time between scan loops |
| randomize | true | jitter on positions + random delays |
| debug | false | live window. green = click. yellow = passed shape but failed color |

---

## debug mode

set debug: true in settings.json or enable during setup

green boxes = detected and clicking. yellow boxes = shape matched but color check failed (false positive caught). (its actually crazy)

if no display is available (SSH etc) it writes debug_out.png each frame instead

---

## troubleshooting (ai written, idk)

**no templates loaded**  sprites folder is empty or wrong path. check it exists and has pngs in it

**clicks in wrong places on mac**  retina scale issue. script detects it automatically but if it's still off, say no when it asks "platform looks correct?" and manually set scale to 2.0

**yellow boxes everywhere** — color threshold too strict. try raising color_threshold to 0.45 or so in settings.json (this is real)

**nothing detected** — shape threshold too high. try lowering shape_threshold to 0.50. also check your template crops are clean

**slow** — set a tighter ROI. also lower scale_max if sprites don't appear at large sizes in your game window (Full-screen scanning on 4K is usually slow.)

**anti cheat**  uhh probably not an issue lol
