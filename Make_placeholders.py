"""
make_placeholders.py — DeltaV Architecture Image Generator
============================================================
Run ONCE to populate ./images/ with placeholder PNGs.

  python make_placeholders.py

Replace any file with a real hardware photo (same filename) to upgrade
that component's look. The generator falls back to coloured shapes if
an image is missing, so nothing breaks.

Image list
──────────
Structural  : rack_frame.png, pdc_room_bg.png, operator_room_bg.png,
              operator_desk.png, fopp_connector.png
Monitors    : MONITOR.png, PRINTER.png
Components  : CONTROLLER, CIOC, CHARM_BASE, CHARM_AI, CHARM_AO,
              CHARM_DI, CHARM_DO, POWER, UPS, WORKSTATION, SWITCH,
              SCREEN, FOPP, MEDIA_CONV, FIBER, SOFTWARE, FIREWALL,
              OPERATOR_WS
"""

import os, math
from PIL import Image, ImageDraw, ImageFilter

IMAGES_DIR = os.path.join(os.path.dirname(__file__), 'images')
os.makedirs(IMAGES_DIR, exist_ok=True)


# ─── helpers ──────────────────────────────────────────────────────────────────

def _save(img: Image.Image, filename: str):
    path = os.path.join(IMAGES_DIR, filename)
    img.save(path, 'PNG', optimize=True)
    print(f'  ✓  {filename:<35}  {img.size[0]}×{img.size[1]}')


def _font(size: int, bold: bool = False):
    from PIL import ImageFont
    candidates = [
        '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'    if bold else '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
        '/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf' if bold else '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf',
        '/usr/share/fonts/truetype/freefont/FreeSansBold.ttf'      if bold else '/usr/share/fonts/truetype/freefont/FreeSans.ttf',
    ]
    for p in candidates:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, max(7, size))
            except Exception:
                pass
    return ImageFont.load_default()


def _hex(h: str) -> tuple:
    h = h.lstrip('#')
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))


def _darker(c: tuple, f: float = 0.60) -> tuple:
    return tuple(max(0, int(x * f)) for x in c)


def _lighter(c: tuple, f: float = 1.35) -> tuple:
    return tuple(min(255, int(x * f)) for x in c)


def _text_color(bg: tuple) -> tuple:
    lum = (bg[0]*299 + bg[1]*587 + bg[2]*114) / 1000
    return (255, 255, 255) if lum < 145 else (22, 22, 22)


def _centered_text(draw, text: str, box: tuple, font, color: tuple):
    x1, y1, x2, y2 = box
    bw, bh = x2 - x1, y2 - y1
    try:
        bb = draw.textbbox((0, 0), text, font=font)
        tw, th = bb[2] - bb[0], bb[3] - bb[1]
    except Exception:
        tw, th = len(text) * 7, 12
    draw.text((x1 + max(0, (bw - tw) // 2), y1 + max(0, (bh - th) // 2)),
              text, fill=color, font=font)


def _vgrad(draw, x1, y1, x2, y2, top_col, bot_col):
    """Vertical gradient fill from top_col to bot_col."""
    for y in range(y1, y2):
        t = (y - y1) / max(1, y2 - y1)
        c = tuple(int(top_col[i] + (bot_col[i] - top_col[i]) * t) for i in range(3))
        draw.line([(x1, y), (x2, y)], fill=c)


def _hgrad(draw, x1, y1, x2, y2, l_col, r_col):
    """Horizontal gradient fill."""
    for x in range(x1, x2):
        t = (x - x1) / max(1, x2 - x1)
        c = tuple(int(l_col[i] + (r_col[i] - l_col[i]) * t) for i in range(3))
        draw.line([(x, y1), (x, y2)], fill=c)


# ══════════════════════════════════════════════════════════════════════════════
# 1.  RACK FRAME  (the money shot — pixel-perfect metal cabinet)
# ══════════════════════════════════════════════════════════════════════════════

def make_rack_frame(W: int = 440, H: int = 920) -> Image.Image:
    """
    Renders a photorealistic-style 19-inch server rack cabinet.

    Proportions (matched to generator.py constants):
      Side rails  : ~9 % of W each     (INNER_PAD = 0.06")
      Top/bottom  : ~3 % of H each     (CAP_H     = 0.14")
      Inner panel : remaining area, light grey
    """
    img  = Image.new('RGB', (W, H), (25, 25, 25))
    draw = ImageDraw.Draw(img)

    RAIL = int(W * 0.092)
    CAP  = int(H * 0.032)
    IX1, IX2 = RAIL,     W - RAIL      # inner x
    IY1, IY2 = CAP,      H - CAP       # inner y

    # ── 1a.  Side rails — gradient metallic ──────────────────────────────────
    RAIL_TOP  = (52, 52, 52)
    RAIL_MID  = (42, 42, 42)
    RAIL_BOT  = (30, 30, 30)

    for x in range(RAIL):
        t = x / RAIL
        v = int(RAIL_MID[0] + 14 * math.sin(t * math.pi))  # brighter at centre
        draw.line([(x, IY1), (x, IY2)], fill=(v, v, v))

    for x in range(W - RAIL, W):
        t = (x - (W - RAIL)) / RAIL
        v = int(RAIL_MID[0] + 14 * math.sin((1 - t) * math.pi))
        draw.line([(x, IY1), (x, IY2)], fill=(v, v, v))

    # ── 1b.  Top cap ─────────────────────────────────────────────────────────
    _vgrad(draw, 0, 0, W - 1, CAP, (72, 72, 72), (50, 50, 50))

    # ── 1c.  Bottom cap ──────────────────────────────────────────────────────
    _vgrad(draw, 0, H - CAP, W - 1, H, (50, 50, 50), (22, 22, 22))

    # ── 1d.  Inner panel ─────────────────────────────────────────────────────
    _vgrad(draw, IX1, IY1, IX2, IY2, (228, 228, 228), (215, 215, 215))

    # Inner inset shadow (top + left = darker, bottom + right = lighter)
    for i in range(4):
        draw.line([(IX1 + i, IY1 + i), (IX2 - i, IY1 + i)], fill=(120 - i*8, 120 - i*8, 120 - i*8))
        draw.line([(IX1 + i, IY1 + i), (IX1 + i, IY2 - i)], fill=(120 - i*8, 120 - i*8, 120 - i*8))
    for i in range(2):
        draw.line([(IX1, IY2 - i), (IX2, IY2 - i)], fill=(205 + i*5, 205 + i*5, 205 + i*5))
        draw.line([(IX2 - i, IY1), (IX2 - i, IY2)], fill=(205 + i*5, 205 + i*5, 205 + i*5))

    # ── 1e.  Cable-management strips (top & bottom of inner) ─────────────────
    STRIP_H = max(7, int(H * 0.015))
    for gy in (IY1, IY2 - STRIP_H):
        _vgrad(draw, IX1, gy, IX2, gy + STRIP_H, (72, 72, 72), (55, 55, 55))
        # Slot cutouts
        for sx in range(IX1 + 6, IX2 - 6, 16):
            draw.rectangle([sx, gy + 2, sx + 11, gy + STRIP_H - 2],
                           fill=(40, 40, 40))
            draw.rectangle([sx + 1, gy + 3, sx + 10, gy + STRIP_H - 3],
                           fill=(28, 28, 28))

    # ── 1f.  Rack ear hole pattern (square rack holes, 2 per U) ──────────────
    U_H    = max(16, int((IY2 - IY1 - 2 * STRIP_H) / 25))   # 25U
    HS     = max(3, int(RAIL * 0.22))                          # half hole size

    for cx in (RAIL // 2, W - RAIL // 2):
        y = IY1 + STRIP_H + U_H // 2
        while y < IY2 - STRIP_H - U_H:
            # Square rack hole (standard EIA-310)
            for hx, hy in [(cx - HS, y - HS), (cx - HS, y + U_H // 2 - HS)]:
                draw.rectangle([hx, hy, hx + 2*HS, hy + 2*HS],
                               fill=(14, 14, 14), outline=(6, 6, 6))
                # Interior glint
                draw.point((hx + HS, hy + HS), fill=(35, 35, 35))
            y += U_H

    # ── 1g.  Screw heads on mounting ears ────────────────────────────────────
    SCREW_R = max(3, int(RAIL * 0.17))
    for sx in (RAIL // 2, W - RAIL // 2):
        for sy in (CAP // 2, H - CAP // 2):
            draw.ellipse([sx - SCREW_R, sy - SCREW_R, sx + SCREW_R, sy + SCREW_R],
                         fill=(28, 28, 28), outline=(80, 80, 80))
            # Phillips slot
            cr = SCREW_R - 1
            draw.line([(sx - cr, sy), (sx + cr, sy)], fill=(12, 12, 12), width=1)
            draw.line([(sx, sy - cr), (sx, sy + cr)], fill=(12, 12, 12), width=1)

    # ── 1h.  Outer edges ──────────────────────────────────────────────────────
    draw.line([(0, 0), (W - 1, 0)], fill=(95, 95, 95), width=2)
    draw.line([(0, 0), (0, H - 1)], fill=(78, 78, 78), width=1)
    draw.line([(0, H - 1), (W - 1, H - 1)], fill=(10, 10, 10), width=2)
    draw.line([(W - 1, 0), (W - 1, H - 1)], fill=(10, 10, 10), width=1)

    # Very slight blur to soften pixel edges (looks more photographic)
    img = img.filter(ImageFilter.SMOOTH_MORE)
    return img


# ══════════════════════════════════════════════════════════════════════════════
# 2.  RACK-MOUNT COMPONENT PANEL  (1U / 2U horizontal module)
# ══════════════════════════════════════════════════════════════════════════════

def make_rack_component(label: str, color_hex: str,
                         W: int = 390, H: int = 96,
                         sub_label: str = '') -> Image.Image:
    """
    Horizontal rack-mount front panel with:
      • Gradient fill in the component's brand colour
      • Left handle strip with 3 status LEDs
      • Large centre label + optional subtitle
      • 4 RJ45-style port indicators on the right
    """
    fill  = _hex(color_hex)
    dark  = _darker(fill, 0.55)
    light = _lighter(fill, 1.40)
    tc    = _text_color(fill)

    img  = Image.new('RGB', (W, H), fill)
    draw = ImageDraw.Draw(img)

    # Main gradient (top lighter → bottom slightly darker)
    top_c = tuple(min(255, int(c * 1.28)) for c in fill)
    bot_c = tuple(int(c * 0.82) for c in fill)
    _vgrad(draw, 0, 0, W, H, top_c, bot_c)

    # Outer border
    draw.rectangle([0, 0, W - 1, H - 1], outline=dark, width=2)

    # Top bevel highlight
    draw.line([(1, 1), (W - 2, 1)], fill=light, width=1)
    draw.line([(1, 1), (1, H - 2)], fill=light, width=1)

    # Bottom shadow edge
    draw.line([(1, H - 2), (W - 2, H - 2)], fill=dark, width=1)

    # ── Handle strip (left 8%) ────────────────────────────────────────────────
    HS = int(W * 0.08)
    _hgrad(draw, 0, 1, HS, H - 1, dark, fill)
    draw.line([(HS, 1), (HS, H - 2)], fill=dark, width=1)

    # LED indicators on handle
    LED_COLOURS = [(0, 220, 75), (255, 195, 0), (0, 170, 255)]
    led_cx = HS // 2
    for i, lc in enumerate(LED_COLOURS):
        ly = H // (len(LED_COLOURS) + 1) * (i + 1)
        r  = max(3, int(HS * 0.19))
        draw.ellipse([led_cx - r, ly - r, led_cx + r, ly + r],
                     fill=lc, outline=(0, 0, 0))
        gr = max(1, r // 2)
        glow = tuple(min(255, c + 90) for c in lc)
        draw.ellipse([led_cx - gr, ly - gr, led_cx + gr, ly + gr], fill=glow)

    # ── Text labels ───────────────────────────────────────────────────────────
    TX = HS + 14
    TW = W - TX - 55       # leave room for ports on right
    fsize = max(11, min(26, H // 3))

    font_main = _font(fsize, bold=True)
    _centered_text(draw, label, (TX, H // 6, TX + TW, 4 * H // 6), font_main, tc)

    if sub_label:
        font_sub = _font(max(7, fsize - 7))
        sub_c    = tuple(max(0, c - 55) for c in tc) if tc[0] > 128 else tuple(min(255, c + 55) for c in tc)
        _centered_text(draw, sub_label, (TX, 3 * H // 5, TX + TW, H - 5), font_sub, sub_c)

    # ── Port cluster (right side) ─────────────────────────────────────────────
    PX  = W - 50
    PW  = 42
    PH  = max(7, H // 9)
    PG  = max(2, H // 12)
    py  = (H - 4 * (PH + PG)) // 2

    for i in range(4):
        _vgrad(draw, PX, py, PX + PW, py + PH, dark, _darker(fill, 0.40))
        draw.rectangle([PX, py, PX + PW - 1, py + PH - 1], outline=(0, 0, 0))
        # Socket interior
        draw.rectangle([PX + 2, py + 2, PX + PW - 3, py + PH - 3], fill=(8, 8, 8))
        # Activity LED
        led = (0, 200, 0) if i < 2 else (90, 90, 90)
        draw.ellipse([PX + 4, py + PH // 4, PX + 8, py + 3 * PH // 4], fill=led)
        py += PH + PG

    return img


# ══════════════════════════════════════════════════════════════════════════════
# 3.  CHARM BASEPLATE (narrow portrait module, 8 card slots)
# ══════════════════════════════════════════════════════════════════════════════

def make_charm_baseplate(label: str, color_hex: str,
                          W: int = 72, H: int = 420) -> Image.Image:
    """
    Tall, narrow CHARM module front face with 8 card-slot windows.
    Each window shows two I/O terminal blocks (matching real hardware).
    """
    fill  = _hex(color_hex)
    dark  = _darker(fill, 0.50)
    light = _lighter(fill, 1.30)
    tc    = _text_color(fill)

    img  = Image.new('RGB', (W, H), fill)
    draw = ImageDraw.Draw(img)

    # Gradient
    top_c = tuple(min(255, int(c * 1.22)) for c in fill)
    bot_c = tuple(int(c * 0.78) for c in fill)
    _vgrad(draw, 0, 0, W, H, top_c, bot_c)

    draw.rectangle([0, 0, W - 1, H - 1], outline=dark, width=2)
    draw.line([(1, 1), (W - 2, 1)], fill=light)
    draw.line([(1, 1), (1, H - 2)], fill=light)

    # ── Header bar ────────────────────────────────────────────────────────────
    HDR = max(22, int(H * 0.068))
    _vgrad(draw, 0, 0, W, HDR, dark, _darker(fill, 0.70))
    draw.line([(0, HDR), (W - 1, HDR)], fill=dark, width=2)

    f_hdr = _font(max(7, HDR // 3), bold=True)
    _centered_text(draw, label, (1, 1, W - 1, HDR - 1), f_hdr, tc)

    # ── 8 card slot windows ───────────────────────────────────────────────────
    PAD    = 3
    SLOTS  = 8
    avail  = H - HDR - PAD * 2
    slot_h = (avail - PAD * (SLOTS - 1)) // SLOTS
    slot_w = W - PAD * 2

    SLOT_BG    = _darker(fill, 0.42)
    SLOT_HL    = _lighter(fill, 1.12)
    TERM_COLOR = (195, 195, 195)

    for i in range(SLOTS):
        sy = HDR + PAD + i * (slot_h + PAD)
        sx = PAD

        # Slot background
        draw.rectangle([sx, sy, sx + slot_w - 1, sy + slot_h - 1],
                       fill=SLOT_BG, outline=(0, 0, 0))
        # Bevel
        draw.line([(sx + 1, sy + 1), (sx + slot_w - 2, sy + 1)], fill=SLOT_HL)
        draw.line([(sx + 1, sy + 1), (sx + 1, sy + slot_h - 2)], fill=SLOT_HL)

        # Terminal blocks (left + right)
        TW = max(4, slot_w // 5)
        TH = max(3, slot_h - 4)
        ty = sy + 2
        for tx in (sx + 3, sx + slot_w - 3 - TW):
            draw.rectangle([tx, ty, tx + TW - 1, ty + TH - 1],
                           fill=TERM_COLOR, outline=(90, 90, 90))
            # Terminal screw
            draw.rectangle([tx + 1, ty + TH // 2 - 1, tx + TW - 2, ty + TH // 2 + 1],
                           fill=(130, 130, 130))

        # Slot number
        f_num = _font(max(5, slot_h // 3 - 1))
        draw.text((sx + slot_w // 2 - 3, sy + 2), str(i + 1), fill=tc, font=f_num)

    return img


# ══════════════════════════════════════════════════════════════════════════════
# 4.  MONITOR
# ══════════════════════════════════════════════════════════════════════════════

def make_monitor(W: int = 340, H: int = 270) -> Image.Image:
    img  = Image.new('RGB', (W, H), (245, 245, 245))
    draw = ImageDraw.Draw(img)

    BEZEL   = int(W * 0.045)
    STAND_H = int(H * 0.10)
    FOOT_H  = int(H * 0.05)
    SCREEN_H = H - STAND_H - FOOT_H - 2

    # Bezel gradient (dark charcoal)
    _vgrad(draw, 0, 0, W, SCREEN_H, (42, 42, 42), (30, 30, 30))
    draw.rectangle([0, 0, W - 1, SCREEN_H - 1], outline=(18, 18, 18), width=2)

    # Screen area gradient (blue-ish display)
    SX1, SY1 = BEZEL, BEZEL
    SX2, SY2 = W - BEZEL - 1, SCREEN_H - BEZEL - 1
    _vgrad(draw, SX1, SY1, SX2, SY2, (38, 110, 185), (20, 70, 140))

    # Screen content bars (simulated DeltaV UI)
    BH = (SY2 - SY1) // 10
    for i, c in enumerate([(50, 95, 175), (55, 105, 190), (45, 88, 165)]):
        by = SY1 + 2 * BH + i * (BH + 3)
        draw.rectangle([SX1 + 5, by, SX2 - 5, by + BH - 2], fill=c)
        draw.line([(SX1 + 5, by), (SX2 - 5, by)], fill=_lighter(c, 1.3))

    # Brand power LED
    draw.ellipse([W // 2 - 3, SCREEN_H - BEZEL // 2 - 3,
                  W // 2 + 3, SCREEN_H - BEZEL // 2 + 3],
                 fill=(0, 210, 0))

    # Stand neck
    NW = int(W * 0.16)
    NX = W // 2 - NW // 2
    _vgrad(draw, NX, SCREEN_H, NX + NW, SCREEN_H + STAND_H,
           (75, 75, 75), (55, 55, 55))
    draw.line([(NX, SCREEN_H), (NX, SCREEN_H + STAND_H)], fill=(40, 40, 40))
    draw.line([(NX + NW, SCREEN_H), (NX + NW, SCREEN_H + STAND_H)], fill=(90, 90, 90))

    # Stand foot
    FW = int(W * 0.38)
    FX = W // 2 - FW // 2
    FY = SCREEN_H + STAND_H
    _vgrad(draw, FX, FY, FX + FW, FY + FOOT_H, (70, 70, 70), (50, 50, 50))
    draw.rectangle([FX, FY, FX + FW - 1, FY + FOOT_H - 1], outline=(35, 35, 35))

    return img


# ══════════════════════════════════════════════════════════════════════════════
# 5.  PRINTER
# ══════════════════════════════════════════════════════════════════════════════

def make_printer(W: int = 240, H: int = 190) -> Image.Image:
    img  = Image.new('RGB', (W, H), (195, 195, 195))
    draw = ImageDraw.Draw(img)

    # Body
    _vgrad(draw, 0, 0, W, H, (185, 185, 185), (155, 155, 155))
    draw.rectangle([0, 0, W - 1, H - 1], outline=(95, 95, 95), width=2)

    # Paper input tray
    T_Y, T_H = int(H * 0.20), int(H * 0.09)
    draw.rectangle([int(W * 0.10), T_Y, int(W * 0.90), T_Y + T_H],
                   fill=(115, 115, 115), outline=(75, 75, 75))
    draw.rectangle([int(W * 0.12), T_Y + 2, int(W * 0.88), T_Y + T_H - 2],
                   fill=(238, 238, 238))

    # Control panel
    CP_Y, CP_H = int(H * 0.48), int(H * 0.21)
    _vgrad(draw, int(W * 0.08), CP_Y, int(W * 0.92), CP_Y + CP_H,
           (72, 72, 72), (55, 55, 55))
    draw.rectangle([int(W * 0.08), CP_Y, int(W * 0.92) - 1, CP_Y + CP_H - 1],
                   outline=(45, 45, 45))

    for bx, bc in [(int(W * 0.20), (0, 200, 0)),
                   (int(W * 0.40), (200, 180, 0)),
                   (int(W * 0.60), (200, 50, 50))]:
        r = max(5, W // 20)
        draw.ellipse([bx - r, CP_Y + CP_H // 2 - r, bx + r, CP_Y + CP_H // 2 + r], fill=bc)

    # Output tray
    OT_Y = int(H * 0.73)
    draw.rectangle([int(W * 0.05), OT_Y, int(W * 0.95), OT_Y + int(H * 0.09)],
                   fill=(145, 145, 145), outline=(100, 100, 100))

    f = _font(max(9, W // 18))
    _centered_text(draw, 'PRINTER', (0, H - 24, W, H - 2), f, (55, 55, 55))
    return img


# ══════════════════════════════════════════════════════════════════════════════
# 6.  OPERATOR DESK
# ══════════════════════════════════════════════════════════════════════════════

def make_operator_desk(W: int = 520, H: int = 340) -> Image.Image:
    img  = Image.new('RGB', (W, H), (200, 200, 200))
    draw = ImageDraw.Draw(img)

    # Desk top surface (dark laminate)
    TOP_H = int(H * 0.17)
    _vgrad(draw, 0, 0, W, TOP_H, (165, 135, 100), (145, 115, 82))

    # Body
    _vgrad(draw, 0, TOP_H, W, H, (188, 188, 188), (165, 165, 165))
    draw.rectangle([0, 0, W - 1, H - 1], outline=(95, 95, 95), width=2)
    draw.line([(0, TOP_H), (W - 1, TOP_H)], fill=(80, 80, 80), width=2)

    # Drawer dividers
    draw.line([(W // 2, TOP_H), (W // 2, H)], fill=(170, 170, 170))

    # Drawer handles
    for dx in (W // 4 - 45, 3 * W // 4 - 45):
        hy = TOP_H + int((H - TOP_H) * 0.38)
        _hgrad(draw, dx, hy, dx + 90, hy + 14, (130, 130, 130), (155, 155, 155))
        draw.rectangle([dx, hy, dx + 89, hy + 13], outline=(85, 85, 85))
        draw.line([(dx + 12, hy + 4), (dx + 78, hy + 4)], fill=(160, 160, 160))
        draw.line([(dx + 12, hy + 9), (dx + 78, hy + 9)], fill=(120, 120, 120))

    # Cable grommets on desk top
    for gx in (W // 4, W // 2, 3 * W // 4):
        draw.ellipse([gx - 9, TOP_H - 9, gx + 9, TOP_H + 9],
                     fill=(58, 58, 58), outline=(38, 38, 38))
        draw.ellipse([gx - 5, TOP_H - 5, gx + 5, TOP_H + 5],
                     fill=(20, 20, 20))

    return img


# ══════════════════════════════════════════════════════════════════════════════
# 7.  FOPP CONNECTOR (front-panel fibre patch)
# ══════════════════════════════════════════════════════════════════════════════

def make_fopp(W: int = 130, H: int = 88) -> Image.Image:
    img  = Image.new('RGBA', (W, H), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    FILL  = (130, 30, 155)
    DARK  = (70,  10,  85)
    LIGHT = (195, 70, 215)

    # Body
    draw.rounded_rectangle([3, 3, W - 4, H - 4], radius=8,
                            fill=FILL, outline=DARK, width=2)
    draw.line([(8, 6), (W - 9, 6)], fill=LIGHT, width=2)
    draw.line([(6, 8), (6, H - 8)], fill=LIGHT, width=1)

    # 4 fibre port circles
    PORT_R = max(5, H // 9)
    total_ports_w = 4 * PORT_R * 2 + 3 * 8
    px_start = (W - total_ports_w) // 2
    py = H // 2

    for i in range(4):
        px = px_start + i * (PORT_R * 2 + 8)
        draw.ellipse([px, py - PORT_R, px + PORT_R * 2, py + PORT_R],
                     fill=(15, 15, 15), outline=DARK)
        draw.ellipse([px + 2, py - PORT_R + 2, px + PORT_R * 2 - 2, py + PORT_R - 2],
                     fill=(0, 0, 0))
        draw.ellipse([px + PORT_R - 2, py - 2, px + PORT_R + 2, py + 2],
                     fill=(80, 195, 255))

    # Label
    f = _font(max(9, H // 6), bold=True)
    _centered_text(draw, 'FOPP', (0, 3 * H // 5, W, H - 4), f, (255, 255, 255))
    return img


# ══════════════════════════════════════════════════════════════════════════════
# 8.  ROOM BACKGROUND  (subtle grid texture)
# ══════════════════════════════════════════════════════════════════════════════

def make_room_bg(W: int = 1200, H: int = 800,
                 tint: tuple = (248, 248, 252)) -> Image.Image:
    img  = Image.new('RGB', (W, H), tint)
    draw = ImageDraw.Draw(img)

    GRID   = 48
    GRID_C = tuple(max(0, c - 7) for c in tint)
    DOT_C  = tuple(max(0, c - 15) for c in tint)

    for x in range(0, W, GRID):
        draw.line([(x, 0), (x, H)], fill=GRID_C, width=1)
    for y in range(0, H, GRID):
        draw.line([(0, y), (W, y)], fill=GRID_C, width=1)
    for x in range(0, W, GRID):
        for y in range(0, H, GRID):
            draw.ellipse([x - 1, y - 1, x + 1, y + 1], fill=DOT_C)

    return img


# ══════════════════════════════════════════════════════════════════════════════
# COMPONENT DEFINITIONS TABLE
# ══════════════════════════════════════════════════════════════════════════════

# (output_filename, display_label, color_hex, sub_label, style)
COMPONENT_DEFS = [
    # Rack-mount (horizontal 1-2U panels)
    ('CONTROLLER.png',  'PK Controller',    '1A3A6B', 'Redundant Ready',  'rack'),
    ('CIOC.png',        'CIOC',             '2E5FA3', 'CHARM I/O Card 2', 'rack'),
    ('POWER.png',       'Power Supply',     'CC3300', '24 VDC Bulk',      'rack'),
    ('UPS.png',         'APC UPS',          'B8860B', 'Battery Backup',   'rack'),
    ('WORKSTATION.png', 'EWS / ProPlus',    '2B5EA7', 'Rack Workstation', 'rack'),
    ('SWITCH.png',      'DCS Switch',       '2D6A2D', '1 Gbps Managed',   'rack'),
    ('SCREEN.png',      '19″ KVM Screen',   '555566', 'Pullout LCD',      'rack'),
    ('MEDIA_CONV.png',  'Media Converter',  '00A0B0', 'SM Fibre',         'rack'),
    ('FIBER.png',       'Fibre Cable',      '6040A0', 'Single Mode OS2',  'rack'),
    ('FIREWALL.png',    'Firewall',         'C03030', 'NextGen UTM',      'rack'),
    ('SOFTWARE.png',    'Software Licence', '4A8C3F', 'DeltaV DST',       'rack'),
    ('OPERATOR_WS.png', 'Operator WS',      '1E4F9E', 'Full Tower OWS',   'rack'),
    # CHARM modules (narrow portrait)
    ('CHARM_BASE.png',  'CHARM BP',         '8B4513', '', 'charm'),
    ('CHARM_AI.png',    'AI CHARM',         'C96A00', '', 'charm'),
    ('CHARM_AO.png',    'AO CHARM',         'B85C00', '', 'charm'),
    ('CHARM_DI.png',    'DI CHARM',         'A85000', '', 'charm'),
    ('CHARM_DO.png',    'DO CHARM',         '963D00', '', 'charm'),
]


# ══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════

def main():
    print(f'\n📁  Output directory: {IMAGES_DIR}\n')
    print('── Structural images ─────────────────────────────────────────────')
    _save(make_rack_frame(440, 920),                              'rack_frame.png')
    _save(make_operator_desk(520, 340),                           'operator_desk.png')
    _save(make_fopp(130, 88),                                     'fopp_connector.png')
    _save(make_room_bg(1200, 800, (248, 248, 252)),               'pdc_room_bg.png')
    _save(make_room_bg(800,  800, (247, 252, 248)),               'operator_room_bg.png')

    print('\n── Special components ────────────────────────────────────────────')
    _save(make_monitor(340, 270),                                 'MONITOR.png')
    _save(make_printer(240, 190),                                 'PRINTER.png')
    _save(make_fopp(130, 88),                                     'FOPP.png')   # same design

    print('\n── Rack & CHARM components ───────────────────────────────────────')
    for fname, label, chex, sub, style in COMPONENT_DEFS:
        if style == 'charm':
            img = make_charm_baseplate(label, chex)
        else:
            img = make_rack_component(label, chex, sub_label=sub)
        _save(img, fname)

    total = 5 + 3 + len(COMPONENT_DEFS)
    print(f'\n✅  {total} images written to ./images/')
    print('   ↳  Replace any file with a real hardware photo (keep same filename)')
    print('   ↳  generator.py auto-falls back to coloured shapes if an image is missing\n')

    # Write a README inside images/
    readme = os.path.join(IMAGES_DIR, 'README.md')
    with open(readme, 'w') as f:
        f.write("# DeltaV Architecture — Component Images\n\n"
                "Each `.png` in this folder maps directly to a `diagram_class` in `rules.json`.\n\n"
                "| File | Component |\n|---|---|\n")
        for fname, label, *_ in COMPONENT_DEFS:
            f.write(f"| `{fname}` | {label} |\n")
        f.write("\n**To use real hardware photos:**\n"
                "1. Name the photo exactly as shown above (e.g. `CONTROLLER.png`)\n"
                "2. Drop it in this folder — transparent PNG works best\n"
                "3. Re-run the generator; the new image appears automatically\n")
    print(f'   ↳  README written to {readme}\n')


if __name__ == '__main__':
    main()