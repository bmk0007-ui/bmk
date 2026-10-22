from lib import *


def build(S):
    S.header(1, "외부 환경 분석", sub="R: 지역 환경 분석", sub_w=198)
    S.panel(h=566)
    S.box(65, 162, 400, 492, fill=WHITE, line="D0D7E2", shape="round", radius=0.04)
    S.crop(108, 198, 340, 446)  # 충청남도 인구 유입·유출 지도
    S.box(75, 174, 100, 30, fill=WHITE, line=NAVY2, lw=1.5, shape="callout", adj=[0.62, 0.0, 0.16667],
          text="인구 영역", size=12.5, bold=True, color=NAVY)
    S.box(495, 185, 398, 350, fill=WHITE, line="D0D7E2", shape="round", radius=0.05)
    S.box(495, 168, 390, 30, fill=("595959", "D9D9D9", 0), shape="pentagon", adj=[0.4],
          text="산업·교통 인프라에 따른 인구유입과 유출 구조", size=12.5, bold=True, color=WHITE,
          align="l", margin=(14, 0, 0, 0))
    paras = arrow_paras([
        "7개 시군(아산, 서산, 천안, 청양, 태안, 계룡, 당진)에서는\n인구증가율 증가, 7개 시군(부여, 논산, 서천, 보령, 금산,\n"
        "공주, 홍성)에서는 인구증가율 감소 전망 → 충남\n총인구는 2020년 2,177천명에서 2040년 2,255천명\n으로 증가할 전망",
    ])
    paras += [{"runs": [("", {})], "size": 6}]
    paras += arrow_paras(["북부권은 주요 대기업 생산공장의 입지, 수도권과의\n인접성, 발달된 교통망 등 기반으로 양질의 고용 환경을\n제공하며 지속적인 인구 순유입"])
    paras += [{"runs": [("", {})], "size": 6}]
    paras += arrow_paras(["남부권은 양질의 일자리 부족, 취약한 교통 접근성,\n교육·의료 인프라 미흡 등의 요인으로 인해 인구 유출\n현상 지속"])
    S.text(500, 205, 392, 320, paras, size=11, valign="m", spacing=1.25)
    S.box(675, 543, 40, 12, fill=NAVY, shape="dtri")
    S.box(495, 555, 398, 100, fill=("1F4E8C", "0B2A5B", 90),
          text="충남은 북부권의 산업 집중은 일자리와 인프라 확충을\n이끌지만, 남부권은 산업기반 취약으로 청년층 유출이\n"
               "심화됨. 이에 따라 대학은 지역 인구 감소와 산업·고용\n불균형에 대응한 지역산업 맞춤형 교육 및 평생학습체계\n강화가 필요함",
          size=10.5, color=WHITE, align="l", margin=(40, 0, 4, 0), spacing=1.0)
