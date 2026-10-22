from lib import *


def build(S):
    S.header(1, "외부 환경 분석", sub="R: 지역 환경 분석", sub_w=198)
    S.panel(h=566)
    S.box(65, 162, 400, 503, fill=WHITE, line="D0D7E2", shape="round", radius=0.03)
    S.crop(95, 165, 366, 497)  # 전국 산업단지 현황지도
    S.box(72, 172, 200, 32, fill=WHITE, line=NAVY2, lw=1.5, shape="callout", adj=[0.56, 0.0, 0.16667],
          text="지역산업 구조 및 고용환경", size=12.5, bold=True, color=NAVY)
    S.box(495, 185, 398, 367, fill=WHITE, line="D0D7E2", shape="round", radius=0.05)
    S.box(495, 168, 390, 30, fill=("595959", "D9D9D9", 0), shape="pentagon", adj=[0.4],
          text="충남 주력산업: 디스플레이, 자동차부품, 기계, 바이오", size=11.5, bold=True, color=WHITE,
          align="l", margin=(8, 0, 0, 0))
    items = [("디스플레이 산업", "아산·천안 중심 → 글로벌 경기 변동에 민감"),
             ("자동차부품·기계 산업", "서산·당진 중심 → 고용규모 크나 경기 의존도 높음"),
             ("바이오 헬스 산업", "천안·공주 중심 → 정부 바이오 클러스터 정책과 연계"),
             ("화학·에너지 산업", "서천·보령 중심 → ESG 규제와 탄소중립 대응 필요"),
             ("신성장 산업", "반도체·로봇·스마트팩토리 등 첨단산업 확대 기반\n마련 중")]
    paras = []
    for i, (h, b) in enumerate(items):
        paras.append({"runs": [(h, {"b": True, "c": NAVY})], "before": 0 if i == 0 else 12})
        paras += arrow_paras(b, mark="- ", indent="  ")
    S.text(508, 220, 385, 325, paras, size=11.5, valign="t", spacing=1.2)
    S.box(675, 562, 40, 12, fill=NAVY, shape="dtri")
    S.box(495, 574, 398, 92, fill=("1F4E8C", "0B2A5B", 90),
          text="산업 구조 고도화와 인력 수요 변화에 대응해 대학은 전공\n재편·산학협력 중심 교육과정을 강화하고,지역 주력산업\n"
               "(디스플레이·자동차·바이오·에너지)과 연계한 현장기반\n취·창업 지원 시스템을 구축해야 함",
          size=10.5, color=WHITE, align="l", margin=(38, 0, 4, 0), spacing=1.0)
