"""Custom vector illustrations (animated SVG) for the Climate İMECE digital-screen deck.

Every illustration is a function f(t0, static) -> (svg_markup, viewbox_w, viewbox_h).
t0 is the time (s) at which its entrance animation starts; static=True drops the
endless loops so a clean final frame can be rasterised for the .pptx.
"""
import math

BLUE, GREEN, LBLUE, RED = "#124DA2", "#44B97E", "#1B75BB", "#E41E26"
INK, GRAY, LGRAY, WHITE = "#1F2937", "#AEB7C4", "#DDE3EA", "#FFFFFF"
PAL = [BLUE, GREEN, LBLUE, RED]

CSS = """
svg text{font-family:'Inter',sans-serif;font-weight:800}
.pop{transform-box:fill-box;transform-origin:center;animation:pop .55s cubic-bezier(.2,1.5,.4,1) both}
.fade{animation:fade .6s ease both}
.draw{stroke-dasharray:1;animation:draw .9s ease both}
.grow{transform-box:fill-box;transform-origin:center bottom;animation:grow .9s cubic-bezier(.2,1.3,.4,1) both}
.pulse{transform-box:fill-box;transform-origin:center;animation:pulse 4s ease-in-out infinite}
.ripple{transform-box:fill-box;transform-origin:center;animation:ripple 3.2s ease-out infinite backwards}
.sway{transform-box:fill-box;transform-origin:center bottom;animation:sway 5s ease-in-out infinite}
.turn{transform-box:fill-box;transform-origin:center;animation:turn 60s linear infinite}
@keyframes pop{from{opacity:0;transform:scale(.2)}to{opacity:1;transform:scale(1)}}
@keyframes fade{from{opacity:0}to{opacity:1}}
@keyframes draw{from{stroke-dashoffset:1}to{stroke-dashoffset:0}}
@keyframes grow{from{opacity:0;transform:scale(.05)}to{opacity:1;transform:scale(1)}}
@keyframes pulse{0%,100%{transform:scale(1)}50%{transform:scale(1.09)}}
@keyframes ripple{0%{opacity:0;transform:scale(.25)}12%{opacity:.55}100%{opacity:0;transform:scale(1.7)}}
@keyframes sway{0%,100%{transform:rotate(-3deg)}50%{transform:rotate(3deg)}}
@keyframes turn{from{transform:rotate(0)}to{transform:rotate(360deg)}}
"""


def a(cls, delay):
    return f'class="{cls}" style="animation-delay:{delay:.2f}s"'


def pt(cx, cy, r, ang):
    """Point on a circle; ang in degrees clockwise from 12 o'clock."""
    rad = math.radians(ang - 90)
    return cx + r * math.cos(rad), cy + r * math.sin(rad)


def arc(cx, cy, r, a0, a1):
    x0, y0 = pt(cx, cy, r, a0)
    x1, y1 = pt(cx, cy, r, a1)
    large = 1 if (a1 - a0) % 360 > 180 else 0
    return f"M{x0:.1f},{y0:.1f} A{r},{r} 0 {large} 1 {x1:.1f},{y1:.1f}"


def people_ring(cx, cy, R, n, t0, s=1.0, step=0.18, colors=PAL, start=0.0):
    """Ring of figures holding hands, heads pointing outward."""
    out = []
    seg = 360 / n
    # hands: one coloured arc per person, joined into a full ring
    for i in range(n):
        ang = start + i * seg
        out.append(f'<path d="{arc(cx, cy, R, ang - seg / 2, ang + seg / 2)}" fill="none" '
                   f'stroke="{colors[i % len(colors)]}" stroke-width="{9 * s:.1f}" stroke-linecap="round" '
                   f'pathLength="1" {a("draw", t0 + n * step + 0.1 * i)}/>')
    for i in range(n):
        ang = start + i * seg
        c = colors[i % len(colors)]
        out.append(
            f'<g transform="translate({cx},{cy}) rotate({ang:.1f})"><g {a("pop", t0 + i * step)}>'
            f'<circle cx="0" cy="{-R - 31 * s:.1f}" r="{13 * s:.1f}" fill="{c}"/>'
            f'<rect x="{-12 * s:.1f}" y="{-R - 15 * s:.1f}" width="{24 * s:.1f}" height="{46 * s:.1f}" '
            f'rx="{12 * s:.1f}" fill="{c}"/></g></g>')
    return "".join(out)


def standing(x, y, color, delay, s=1.0, bubble=False):
    g = (f'<g transform="translate({x},{y})"><g {a("pop", delay)}>'
         f'<circle cx="0" cy="{-58 * s:.1f}" r="{15 * s:.1f}" fill="{color}"/>'
         f'<rect x="{-14 * s:.1f}" y="{-40 * s:.1f}" width="{28 * s:.1f}" height="{58 * s:.1f}" rx="{14 * s:.1f}" fill="{color}"/></g>')
    if bubble:
        g += (f'<g {a("pop", delay + 0.35)}>'
              f'<circle cx="{14 * s:.1f}" cy="{-80 * s:.1f}" r="{3 * s:.1f}" fill="{GRAY}"/>'
              f'<circle cx="{22 * s:.1f}" cy="{-91 * s:.1f}" r="{5 * s:.1f}" fill="{GRAY}"/>'
              f'<circle cx="{40 * s:.1f}" cy="{-112 * s:.1f}" r="{19 * s:.1f}" fill="{WHITE}" stroke="{GRAY}" stroke-width="3"/>'
              f'<text x="{40 * s:.1f}" y="{-104 * s:.1f}" font-size="{24 * s:.0f}" text-anchor="middle" fill="{GRAY}">?</text></g>')
    return g + "</g>"


def sprout(cx, cy, delay, s=1.0, static=False):
    sway = "" if static else 'class="sway" style="animation-delay:0s"'
    return (f'<g transform="translate({cx},{cy})"><g {a("grow", delay)}><g {sway}>'
            f'<path d="M0,0 C0,{-30 * s:.0f} 0,{-50 * s:.0f} 0,{-74 * s:.0f}" stroke="{GREEN}" stroke-width="{8 * s:.0f}" stroke-linecap="round" fill="none"/>'
            f'<path d="M0,{-40 * s:.0f} C{-46 * s:.0f},{-46 * s:.0f} {-58 * s:.0f},{-78 * s:.0f} {-52 * s:.0f},{-92 * s:.0f} C{-22 * s:.0f},{-90 * s:.0f} {-2 * s:.0f},{-70 * s:.0f} 0,{-40 * s:.0f}Z" fill="{GREEN}"/>'
            f'<path d="M0,{-62 * s:.0f} C{44 * s:.0f},{-64 * s:.0f} {60 * s:.0f},{-96 * s:.0f} {56 * s:.0f},{-112 * s:.0f} C{26 * s:.0f},{-112 * s:.0f} {2 * s:.0f},{-92 * s:.0f} 0,{-62 * s:.0f}Z" fill="#6FCB9A"/>'
            f'</g></g></g>')


def wrap(body, w, h):
    return f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {w} {h}" width="100%" height="100%">{body}</svg>'


# --- slide 2: awareness -> action ------------------------------------------
def awareness_action(t0=0.0, static=False):
    w, h = 980, 520
    b = []
    spots = [(95, 250), (215, 190), (335, 270), (150, 430), (290, 460)]
    for i, (x, y) in enumerate(spots):
        b.append(standing(x, y, GRAY, t0 + i * 0.22, s=1.0, bubble=True))
    # arrow
    t = t0 + 1.6
    b.append(f'<path d="M440,300 L548,300" stroke="{BLUE}" stroke-width="14" stroke-linecap="round" fill="none" pathLength="1" {a("draw", t)}/>')
    b.append(f'<g {a("pop", t + 0.7)}><path d="M540,268 L590,300 L540,332 Z" fill="{BLUE}" stroke="{BLUE}" stroke-width="8" stroke-linejoin="round"/></g>')
    # action cluster
    cx, cy = 775, 285
    if not static:
        for k in range(2):
            b.append(f'<circle cx="{cx}" cy="{cy}" r="150" fill="none" stroke="{GREEN}" stroke-width="5" class="ripple" style="animation-delay:{t0 + 5 + k * 1.6:.2f}s"/>')
    b.append(f'<circle cx="{cx}" cy="{cy}" r="78" fill="#E6F5EC" {a("pop", t + 0.9)}/>')
    b.append(people_ring(cx, cy, 118, 6, t + 1.1, s=1.15, step=0.2))
    b.append(sprout(cx, cy + 46, t + 2.9, s=0.95, static=static))
    return wrap("".join(b), w, h), w, h


# --- slide 3: İMECE circle inside the PDCA cycle ------------------------------
def imece_pdca(t0=0.0, static=False):
    w = h = 720
    cx = cy = 360
    R = 268
    b = []
    labels = [("PLAN", BLUE, 0), ("DO", GREEN, 90), ("CHECK", LBLUE, 180), ("ACT", RED, 270)]
    # connecting arcs with arrow heads (clockwise)
    for i, (_, col, ang) in enumerate(labels):
        a0, a1 = ang + 24, ang + 66
        d = t0 + 2.4 + i * 0.45
        b.append(f'<path d="{arc(cx, cy, R, a0, a1)}" fill="none" stroke="{LGRAY}" stroke-width="12" stroke-linecap="round" pathLength="1" {a("draw", d)}/>')
        hx, hy = pt(cx, cy, R, a1 + 3)
        b.append(f'<g transform="translate({hx:.1f},{hy:.1f}) rotate({a1 + 3 + 90:.1f})"><g {a("pop", d + 0.6)}>'
                 f'<path d="M-4,-17 L20,0 L-4,17 Z" fill="{GRAY}" stroke="{GRAY}" stroke-width="5" stroke-linejoin="round"/></g></g>')
    for i, (lab, col, ang) in enumerate(labels):
        x, y = pt(cx, cy, R, ang)
        pw = 150 if len(lab) > 4 else 126
        d = t0 + 2.0 + i * 0.45
        inner = (f'<rect x="{-pw / 2}" y="-29" width="{pw}" height="58" rx="29" fill="{col}"/>'
                 f'<text x="0" y="9" font-size="26" text-anchor="middle" fill="{WHITE}" letter-spacing="1">{lab}</text>')
        if static:
            b.append(f'<g transform="translate({x:.1f},{y:.1f})"><g {a("pop", d)}>{inner}</g></g>')
        else:
            b.append(f'<g transform="translate({x:.1f},{y:.1f})"><g {a("pop", d)}>'
                     f'<g class="pulse" style="animation-delay:{t0 + 5 + i:.2f}s">{inner}</g></g></g>')
    # people
    b.append(f'<circle cx="{cx}" cy="{cy}" r="88" fill="#E9F0FA" {a("pop", t0)}/>')
    ring = people_ring(cx, cy, 128, 8, t0 + 0.3, s=1.2, step=0.17)
    b.append(ring if static else f'<g class="turn" style="animation-delay:{t0 + 4:.2f}s">{ring}</g>')
    b.append(sprout(cx, cy + 44, t0 + 1.9, s=0.9, static=static))
    return wrap("".join(b), w, h), w, h


# --- slide 4 icons (200 x 200) ---------------------------------------------------
def icon_model(t0=0.0, static=False):
    b = f'<circle cx="100" cy="100" r="40" fill="#E9F0FA" {a("pop", t0)}/>' + people_ring(100, 100, 56, 6, t0 + 0.1, s=0.62, step=0.1)
    return wrap(b, 200, 200), 200, 200


def icon_bulb(t0=0.0, static=False):
    b = [f'<g {a("pop", t0)}><circle cx="100" cy="92" r="50" fill="{WHITE}" stroke="{BLUE}" stroke-width="10"/>'
         f'<rect x="80" y="140" width="40" height="15" rx="5" fill="{BLUE}"/><rect x="87" y="159" width="26" height="11" rx="5" fill="{BLUE}"/></g>',
         f'<g {a("grow", t0 + 0.45)}><path d="M100,128 L100,98" stroke="{GREEN}" stroke-width="7" stroke-linecap="round"/>'
         f'<path d="M100,104 C72,100 66,78 70,64 C92,66 102,82 100,104Z" fill="{GREEN}"/>'
         f'<path d="M100,96 C126,92 134,72 130,58 C110,60 100,76 100,96Z" fill="#6FCB9A"/></g>']
    for i, (x1, y1, x2, y2) in enumerate([(100, 22, 100, 6), (44, 44, 32, 32), (156, 44, 168, 32)]):
        b.append(f'<path d="M{x1},{y1} L{x2},{y2}" stroke="{RED}" stroke-width="8" stroke-linecap="round" pathLength="1" {a("draw", t0 + 0.9 + i * 0.12)}/>')
    return wrap("".join(b), 200, 200), 200, 200


def icon_guide(t0=0.0, static=False):
    b = [f'<g {a("pop", t0)}>'
         f'<path d="M100,62 C82,48 58,44 34,50 L34,150 C58,144 82,148 100,162 Z" fill="{WHITE}" stroke="{BLUE}" stroke-width="9" stroke-linejoin="round"/>'
         f'<path d="M100,62 C118,48 142,44 166,50 L166,150 C142,144 118,148 100,162 Z" fill="{WHITE}" stroke="{BLUE}" stroke-width="9" stroke-linejoin="round"/></g>']
    for i, (x1, y1, x2, y2) in enumerate([(50, 78, 86, 82), (50, 100, 86, 104), (50, 122, 86, 126), (114, 82, 150, 78), (114, 104, 150, 100)]):
        b.append(f'<path d="M{x1},{y1} L{x2},{y2}" stroke="{LBLUE}" stroke-width="6" stroke-linecap="round" pathLength="1" {a("draw", t0 + 0.4 + i * 0.1)}/>')
    b.append(f'<g {a("pop", t0 + 1.0)}><circle cx="152" cy="140" r="27" fill="{GREEN}" stroke="{WHITE}" stroke-width="6"/>'
             f'<path d="M139,141 L149,151 L166,131" stroke="{WHITE}" stroke-width="8" stroke-linecap="round" stroke-linejoin="round" fill="none"/></g>')
    return wrap("".join(b), 200, 200), 200, 200


def icon_network(t0=0.0, static=False):
    nodes = [(100, 46, BLUE), (44, 146, GREEN), (156, 146, RED)]
    b = []
    for i in range(3):
        x1, y1, _ = nodes[i]; x2, y2, _ = nodes[(i + 1) % 3]
        b.append(f'<path d="M{x1},{y1} L{x2},{y2}" stroke="{LBLUE}" stroke-width="8" stroke-linecap="round" pathLength="1" {a("draw", t0 + 0.5 + i * 0.2)}/>')
    for i, (x, y, c) in enumerate(nodes):
        if not static:
            b.append(f'<circle cx="{x}" cy="{y}" r="30" fill="none" stroke="{c}" stroke-width="4" class="ripple" style="animation-delay:{t0 + 3 + i:.2f}s"/>')
        b.append(f'<g {a("pop", t0 + i * 0.15)}><circle cx="{x}" cy="{y}" r="27" fill="{c}"/>'
                 f'<circle cx="{x}" cy="{y - 7}" r="8" fill="{WHITE}"/><rect x="{x - 10}" y="{y + 3}" width="20" height="14" rx="7" fill="{WHITE}"/></g>')
    return wrap("".join(b), 200, 200), 200, 200


def icon_venn(t0=0.0, static=False):
    b = (f'<circle cx="74" cy="100" r="56" fill="{BLUE}" fill-opacity=".85" {a("pop", t0)}/>'
         f'<circle cx="126" cy="100" r="56" fill="{GREEN}" fill-opacity=".85" {a("pop", t0 + 0.2)}/>')
    return wrap(b, 200, 200), 200, 200


def icon_check(t0=0.0, static=False):
    b = (f'<g {a("pop", t0)}><circle cx="100" cy="100" r="84" fill="{GREEN}"/>'
         f'<path d="M60,104 L88,132 L142,72" stroke="{WHITE}" stroke-width="22" stroke-linecap="round" stroke-linejoin="round" fill="none"/></g>')
    return wrap(b, 200, 200), 200, 200


# --- slide 5: local action, global contribution ---------------------------------
def globe(t0=0.0, static=False):
    w, h = 700, 600
    cx, cy, r = 350, 300, 238
    b = [f'<defs><clipPath id="gc"><circle cx="{cx}" cy="{cy}" r="{r}"/></clipPath></defs>',
         f'<g {a("pop", t0)}><circle cx="{cx}" cy="{cy}" r="{r}" fill="#DCEAF8"/>'
         f'<g clip-path="url(#gc)">'
         # stylised land masses
         f'<path d="M250,140 C300,120 360,130 400,150 C450,140 520,160 560,200 C600,240 590,300 540,310 C520,340 500,372 482,380 C470,350 452,318 430,300 C390,300 360,290 340,268 C300,262 270,240 262,214 C240,196 232,160 250,140Z" fill="#9AD9B7"/>'
         f'<path d="M210,270 C250,250 310,262 330,300 C350,340 338,420 300,470 C276,500 250,480 244,440 C224,410 196,370 190,330 C186,304 194,282 210,270Z" fill="#9AD9B7"/>'
         f'<ellipse cx="{cx}" cy="{cy}" rx="95" ry="{r}" fill="none" stroke="{WHITE}" stroke-width="2.5" opacity=".8"/>'
         f'<ellipse cx="{cx}" cy="{cy}" rx="180" ry="{r}" fill="none" stroke="{WHITE}" stroke-width="2.5" opacity=".8"/>'
         f'<path d="M0,{cy} H{w} M0,{cy - 95} H{w} M0,{cy + 95} H{w} M0,{cy - 175} H{w} M0,{cy + 175} H{w}" stroke="{WHITE}" stroke-width="2.5" opacity=".8"/>'
         f'</g><circle cx="{cx}" cy="{cy}" r="{r}" fill="none" stroke="{BLUE}" stroke-width="7"/></g>']
    pins = [("Türkiye", 352, 214, BLUE, 0, -62), ("Algeria", 262, 300, GREEN, -6, 62), ("India", 478, 322, RED, 8, 62)]
    # arcs between partners
    links = [(0, 1, 280, 222), (0, 2, 440, 228), (1, 2, 372, 392)]
    for i, (p, q, qx, qy) in enumerate(links):
        x1, y1 = pins[p][1], pins[p][2]; x2, y2 = pins[q][1], pins[q][2]
        b.append(f'<path d="M{x1},{y1} Q{qx},{qy} {x2},{y2}" fill="none" stroke="{BLUE}" stroke-width="5" stroke-linecap="round" '
                 f'pathLength="1" {a("draw", t0 + 2.2 + i * 0.5)}/>')
    for i, (name, x, y, c, lx, ly) in enumerate(pins):
        d = t0 + 0.9 + i * 0.4
        if not static:
            for k in range(2):
                b.append(f'<circle cx="{x}" cy="{y}" r="46" fill="none" stroke="{c}" stroke-width="5" class="ripple" style="animation-delay:{t0 + 4 + i * 0.8 + k * 1.6:.2f}s"/>')
        b.append(f'<g {a("pop", d)}><circle cx="{x}" cy="{y}" r="19" fill="{c}" stroke="{WHITE}" stroke-width="6"/></g>')
        pw = 30 + 17 * len(name)
        b.append(f'<g transform="translate({x + lx},{y + ly})"><g {a("pop", d + 0.25)}>'
                 f'<rect x="{-pw / 2}" y="-24" width="{pw}" height="48" rx="24" fill="{c}"/>'
                 f'<text x="0" y="9" font-size="25" text-anchor="middle" fill="{WHITE}">{name}</text></g></g>')
    return wrap("".join(b), w, h), w, h


# --- A5 card hero: circle only ---------------------------------------------------
def imece_circle(t0=0.0, static=False):
    w = h = 400
    b = (f'<circle cx="200" cy="200" r="92" fill="#E9F0FA" {a("pop", t0)}/>'
         + people_ring(200, 200, 132, 8, t0 + 0.2, s=1.25, step=0.1)
         + sprout(200, 246, t0 + 1.2, s=0.95, static=static))
    return wrap(b, w, h), w, h


ILLUST = {f.__name__: f for f in (awareness_action, imece_pdca, icon_model, icon_bulb, icon_guide,
                                    icon_network, icon_venn, icon_check, globe, imece_circle)}
