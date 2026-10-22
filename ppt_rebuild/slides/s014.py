from lib import *


def build(S):
    S.header(1, "외부 환경 분석", sub="R: 지역 환경 분석", sub_w=198)
    S.panel(h=566)
    S.box(65, 162, 400, 492, fill=WHITE, line="D0D7E2", shape="round", radius=0.04)
    S.box(75, 174, 100, 30, fill=WHITE, line=NAVY2, lw=1.5, shape="callout", adj=[0.62, 0.0, 0.16667],
          text="인구 영역", size=12.5, bold=True, color=NAVY)
    S.text(80, 224, 375, 22, "전체인구 대비 노령인구(64세 이상) 구성비(%)", size=10.5, bold=True, color=NAVY,
           align="c", valign="m")
    bar_chart(S, 75, 244, 385, 168, ["’20년", "’25년", "’30년", "’35년", "’40년"],
              [17.9, 22.4, 27.2, 31.7, 36.4], vmin=0, vmax=40, major=5, numfmt="0.0", axis_fmt="0", series="구성비(%)")
    S.line(80, 420, 450, 420, color="BFBFBF", lw=1.0, dash="dash")
    S.text(80, 444, 375, 22, "충청남도 외국인 인구 현황 및 변화 추이", size=10.5, bold=True, color=NAVY,
           align="c", valign="m")
    bar_chart(S, 75, 466, 385, 172, ["’18년", "’19년", "’20년", "’21년", "’22년"],
              [68234, 70675, 64546, 62578, 71159], vmin=58000, vmax=72000, major=2000, numfmt="#,##0",
              series="외국인 인구(명)")
    S.box(495, 162, 398, 412, fill=WHITE, line="D0D7E2", shape="round", radius=0.04)
    paras = []
    items = [
        "충청남도의 64세 이상 노령인구는 ’20년 39만명\n(17.9%)에서 지속적으로 증가하여, ’30년 61만명\n(27.2%), ’40년 82만명(36.4%)에 이를 전망",
        "(천안) 전체인구 대비 노령인구(64세 이상)의 비중은\n‘20년 10.4%에서 증가하여 27.1%로 증가할 전망",
        "(아산) 전체인구 대비 노령인구(64세 이상)의 비중은\n‘20년 12.1%에서 증가하여 28.4%로 증가할 전망",
        "충청남도의 외국인 인구는 2018년 68,234명에서\n2022년 71,159명으로 2,925명 증가",
        "일부 시군(예산, 서산, 계룡)을 제외하면 대부분의\n시군에서 외국인 인구는 증가 추세",
        " 충남지역의 생산가능인구가 지속적으로 감소하고\n있는 상황에서 외국인 전문인력의 적극적인 유입은\n지역 경제 성장을 촉진할 것으로 전망",
    ]
    for i, it in enumerate(items):
        ps = arrow_paras(it)
        if i:
            ps[0] = {"runs": ps[0], "before": 9}
        paras += ps
    S.text(503, 170, 388, 400, paras, size=11.5, valign="m", spacing=1.18)
    S.box(675, 575, 40, 12, fill=NAVY, shape="dtri")
    S.box(495, 587, 398, 68, fill=("1F4E8C", "0B2A5B", 90),
          text="고령화 심화와 외국인 근로자 유입은 지역산업 인력 구조를\n빠르게 변화시키고 있음. 대학은 고령층 평생교육·외국인\n직업훈련 등 다문화·고령친화 교육체계를 강화해야 함.",
          size=10, color=WHITE, align="l", margin=(30, 0, 4, 0), spacing=1.0)
