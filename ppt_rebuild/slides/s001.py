from lib import *


def build(S):
    S.text(200, 80, 560, 80, [[("BU VISION ", {"c": NAVY}), ("2030", {"c": SKY})]], size=44, bold=True,
           align="c", valign="m")
    S.box(230, 158, 130, 8, fill=NAVY)
    S.box(358, 158, 372, 8, fill=(SKY, "EAF2FB", 0))
    S.crop(245, 226, 470, 296)  # campus photo with corner accents
    S.crop(368, 592, 228, 70)  # university logo
