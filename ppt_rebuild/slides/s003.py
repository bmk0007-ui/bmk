from lib import *


def build(S):
    S.header(1, "외부 환경 분석", sub="교육부 주요정책 및 추진 방향 (1/2)", sub_bold="P: 정책 환경 분석 : ")
    S.panel()
    # gray down-chevrons between the rows
    for y in (196, 256, 317):
        S.box(330, y, 470, 24, fill="D3D3D3", shape="dtri")
    labels = [("비    전", 158), ("목    표", 218), ("추진방향", 279), ("중점과제", 345)]
    for t, y in labels:
        S.box(67, y, 162, 40, fill=("7F7F7F", "A6A6A6", 90), shape="round", radius=0.3, text=t, size=13,
              bold=True, color=WHITE)
    S.box(240, 157, 653, 40, fill=("164A86", "17A8F0", 0), text="기회의 사다리가 되는 공정한 교육 실현",
          size=13, bold=True, color=WHITE)
    S.box(240, 218, 653, 38, fill=WHITE, line="D9D9D9", text="모두를 위한 맞춤교육으로 교육격차 해소",
          size=13, bold=True, color=TXT)
    S.box(240, 279, 653, 38, fill=WHITE, line="D9D9D9", text="대학개혁 통한 역동적 지방 시대 견인",
          size=13, bold=True, color=TXT)
    S.box(240, 345, 300, 40, fill=WHITE, line="D9D9D9", text="지역-대학 동반성장 혁신 생태계 구축",
          size=12, bold=True, color=TXT, margin=(1, 0, 1, 0))
    S.box(548, 345, 345, 40, fill=WHITE, line="D9D9D9", text="청년 성장과 지역혁신을 위한 기회와 투자 확대",
          size=12, bold=True, color=TXT, margin=(1, 0, 1, 0))
    # 세부과제
    S.box(66, 424, 55, 248, fill=("7F7F7F", "A6A6A6", 90), shape="round", radius=0.12,
          text="세부\n\n과제", size=13, bold=True, color=WHITE)
    num_callout(S, 148, 432, 178, 90, 1, "대학 혁신 및\n구조개혁 모델 창출")
    detail_box(S, 338, 422, 556, 115, "대학·지역 협력 기반의 자율적 구조개혁 추진",
               ["RISE 체계와 연계한 지역혁신 거점대학 육성",
                "성과기반 재정지원 확대 및 자율·책임형 거버넌스 정립"], tx=420)
    num_callout(S, 148, 570, 178, 92, 2, "지역혁신 연계\n청년 성장 생태계 조성")
    detail_box(S, 338, 560, 556, 113, "학문 간 경계 해소를 통한 융합형 인재양성 체계 구축",
               ["전공 선택권 확대 및 자율전공 활성화",
                "제도 개선을 통한 대학 자율성 강화",
                "지역산업 연계형 청년 성장·정착 생태계 조성"], tx=420)
