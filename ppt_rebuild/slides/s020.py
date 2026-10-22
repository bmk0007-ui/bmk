from lib import *


def build(S):
    S.header(2, "내부 환경 분석", sub="Students: 학생 현황", sub_bold="7s - ③ ", sub_w=310)
    S.panel(h=566)
    S.box(65, 167, 828, 448, fill=WHITE, line="D0D7E2", shape="round", radius=0.03)
    S.box(865, 157, 27, 26, fill=WHITE, line=NAVY2, lw=1.5, shape="round", radius=0.2,
          text="1", size=13, bold=True, color=NAVY2, margin=(0, 0, 0, 0))
    GH = ("7F7F7F", "8C8C8C", 90)
    S.box(106, 200, 221, 45, fill=GH, text="학생정원", size=13, bold=True)
    S.box(337, 200, 263, 45, fill=GH, text="재학생", size=13, bold=True)
    S.box(610, 200, 243, 21, fill=GH, text="성별", size=12, bold=True)
    S.box(610, 225, 116, 20, fill=GH, shape="round", radius=0.5, text="남", size=11.5, bold=True)
    S.box(736, 225, 117, 20, fill=GH, shape="round", radius=0.5, text="여", size=11.5, bold=True)
    for x, w, t in [(106, 221, "11,760명"), (337, 263, "11,586명 (정원 내 : 10,510명\n                정원 외 : 1,076명)"),
                    (610, 116, "5,106명(44%)"), (736, 117, "6,480명(56%)")]:
        S.box(x, 252, w, 45, fill=WHITE, line="3F6A97", lw=1.25, shape="round", radius=0.2, text=t,
              size=11, color=TXT, spacing=1.0, margin=(1, 0, 1, 0))
    # 막대그래프: 중도탈락률 막대가 보조 눈금에 그려진 원본 형태를 그대로 보존하기 위해 이미지로 유지
    S.crop(118, 345, 465, 245)
    S.notes("슬라이드 20 막대그래프 원본 데이터 라벨 - 정원내 신입생 충원율: 백석대 100 / 전국 92.6, "
            "재학생 충원율: 백석대 98.5 / 전국 104.1, 중도탈락률: 백석대 5.9 / 전국 7.0")
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
    cd = CategoryChartData()
    cd.categories = ["수도권", "충청권", "그외지역"]
    cd.add_series("지역별 비율(%)", (70.3, 15.9, 13.8))
    gf = S.s.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, E(615), E(350), E(220), E(220), cd)
    ch = gf.chart
    ch.has_legend = False
    ch.font.size = Pt(9)
    ch.font.name = FONT
    plot = ch.plots[0]
    hole = plot._element.find(qn("c:holeSize"))
    if hole is None:
        hole = etree.SubElement(plot._element, qn("c:holeSize"))
    hole.set("val", "40")
    first = plot._element.find(qn("c:firstSliceAng"))
    if first is not None:
        first.set("val", "0")
    for i, c in enumerate(["44546A", "DEEBF7", "9DC3E6"]):
        pt = plot.series[0].points[i]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = rgb(c)
        pt.format.line.color.rgb = rgb("FFFFFF")
    ch.has_title = False
    fill_plot(ch)
    for x, y, t, c in [(650, 372, "그외지역\n13.8%", TXT), (612, 428, "충청권\n15.9%", TXT),
                       (752, 488, "수도권\n70.3%", WHITE)]:
        S.text(x, y, 70, 36, t, size=9, color=c, align="c", valign="m", spacing=1.2)
    S.box(447, 615, 40, 12, fill=NAVY2, shape="dtri")
    S.box(140, 632, 680, 42, fill=("1F4E8C", "0B2A5B", 90),
          text="신입생 충원은 매우 우수하나, 재학생 충원율은 전국 평균보다 다소 낮고 중도탈락률은 전국 평균보다 낮은편임",
          size=11, margin=(2, 0, 2, 0))
    source(S, "Source : 대학알리미 (2024)", y=703)
