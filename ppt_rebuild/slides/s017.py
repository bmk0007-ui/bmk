from lib import *

GP = ("7F7F7F", "A6A6A6", 90)


def build(S):
    S.header(2, "내부 환경 분석", sub="Strategy: 중장기 발전전략", sub_bold="7s - ① ", sub_w=350)
    S.panel(h=566)
    S.box(93, 170, 102, 33, fill=GP, shape="round", radius=0.3, text="비전", size=14, bold=True)
    S.box(207, 168, 658, 36, fill=("164A86", "2F80D0", 0),
          text="4차 산업혁명시대에 필요한 역량 중심 혁신교육 글로벌 선도대학", size=14, bold=True)
    S.box(93, 217, 102, 33, fill=GP, shape="round", radius=0.3, text="발전목표", size=14, bold=True)
    S.box(207, 214, 658, 36, fill=WHITE, line="D9D9D9", text="전국 대학 상위 10% 이내 교육 성과 창출",
          size=14, bold=True, color=NAVY)
    S.box(93, 262, 102, 70, fill=GP, shape="round", radius=0.2, text="특성화", size=14, bold=True)
    S.box(207, 258, 658, 34, fill=WHITE, line=NAVY2,
          text="목표 : CT · Healthcare · IT 육성을 통한 미래사회 대응 융복합 교육 플랫폼 구축", size=12.5, color=TXT)
    S.box(207, 300, 248, 32, fill=WHITE, line=NAVY2, text="학문분야 : CT, Healthcare, IT", size=12.5, color=TXT)
    S.box(465, 300, 400, 32, fill=WHITE, line=NAVY2, text="교육분야 : 국내 · 외 산업체 맞춤형 교육 학제간 융복합 교육",
          size=11.5, color=TXT, margin=(1, 0, 1, 0))
    S.box(93, 340, 102, 252, fill=GP, shape="round", radius=0.12, text="5대 전략\n\n23개\n과제", size=14, bold=True)
    strat = [
        (340, 46, "인프라 및 시스템 혁신",
         "1. 첨단 캠퍼스 조성  2. 대학 브랜드 차별화  3. 재정 건전성 확보  4. 행정시스템 효율화\n5. 성과기반 평가체계 운영"),
        (393, 72, "교육 운영 혁신",
         "6. 전공 역량 교육체계 선진화  7. 핵심역량 중심 교양교육 강화\n8. 교과 연계 비교과 교육 활성화  9. 상호작용 촉진 교수·학습지원체계 고도화\n"
         "10. 개인 맞춤형 학생상담지도  11. 백석 인재 선발  12. 대학원 교육역량 강화\n13. 교육의 질 관리체계 선진화"),
        (472, 46, "지역사회·산학협력 혁신",
         "14. 산학협력활성화  15. 지역사회 봉사·공헌 활동 고도화  16. 지역 사회 평생교육체제\n구축  17. 경력 기반 취·창업 지원체계 고도화"),
        (526, 30, "글로벌 혁신", "18. 글로벌 교류 확대  19. 재학생 글로벌 역량 제고  20. 해외 취업률 증대"),
        (562, 30, "대학 구성원 혁신", "21. 교원 역량 강화  22. 직원 역량 제고  23. 구성원 소통 및 복지향상"),
    ]
    for y, h, lab, txt in strat:
        S.box(207, y, 175, h, fill=("1F4E8C", "0B2A5B", 90), text=lab, size=12, bold=True, margin=(1, 0, 1, 0))
        S.box(382, y, 483, h, fill=WHITE, line=NAVY2)
        S.text(386, y, 478, h, txt, size=9.5, valign="m", spacing=1.05)
    S.box(93, 600, 102, 66, fill=GP, shape="round", radius=0.25, text="58개\n실천계획", size=14, bold=True)
    S.box(207, 600, 658, 66, fill=("1F4E8C", "4F8FD0", 90), shape="round", radius=0.3)
    plans = [("인프라 및 시스템혁신", "11개"), ("교육운영혁신", "23개"), ("지역사회·산학협력혁신", "9개"),
             ("글로벌혁신", "8개"), ("대학구성원혁신", "7개")]
    for i, (t, n) in enumerate(plans):
        x = 213 + i * 130
        S.box(x, 606, 124, 54, fill=WHITE, shape="round", radius=0.3)
        S.text(x, 606, 124, 54, [[(t, {"s": 9 if len(t) < 11 else 8.5})], [(n, {"s": 15, "b": True, "c": NAVY2})]], size=9.5,
               align="c", valign="m", spacing=1.0)
    source(S, "Source : 백석대학교 2023 하반기 자체진단평가보고서", y=703)
