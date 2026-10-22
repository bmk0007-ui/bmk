from lib import *


def build(S):
    S.text(30, 44, 200, 30, [[("BU VISION ", {"c": NAVY}), ("2030", {"c": SKY})]], size=15, bold=True,
           valign="m", margin=(2, 0, 0, 0))
    S.box(35, 73, 50, 3, fill=NAVY)
    S.box(84, 73, 132, 3, fill=(SKY, "F0F6FC", 0))
    # section divider
    S.box(34, 296, 120, 118, fill=NAVY, shape="snip", adj=[0.17])
    S.box(34, 404, 200, 10, fill=NAVY)
    S.box(230, 404, 570, 10, fill=(SKY, "EAF2FB", 0))
    S.box(152, 386, 80, 18, fill=WHITE, shape="rtri", flipH=True)
    S.text(60, 310, 70, 90, "I", size=40, bold=True, color="E8E8E8", align="c", valign="m", font="Times New Roman")
    S.text(168, 315, 640, 80, "BU-VISION 2030 수립배경", size=36, bold=True, color=NAVY, valign="m")
