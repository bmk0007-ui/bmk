from lib import *


def build(S):
    S.header(1, "외부 환경 분석", sub="P: 정책 환경 분석", sub_w=185)
    S.panel()
    # goal bar
    S.box(65, 157, 830, 43, fill=WHITE, line=NAVY2, lw=1.5, shape="round", radius=0.5)
    S.box(66, 163, 85, 31, fill=NAVY, shape="pentagon", adj=[0.35], text="목표", size=12, bold=True,
          align="l", margin=(14, 0, 0, 0))
    S.text(150, 157, 740, 43, "대학별 자율혁신 통한 체질 개선으로 창의융합형 인재 양성 및 지속가능한 고등교육 생태계 구축",
           size=11.5, bold=True, color=NAVY, align="c", valign="m")
    # left diagram
    S.box(112, 298, 258, 248, fill=None, line="7F96B8", lw=1.0, shape="oval")
    S.box(172, 344, 128, 114, fill=WHITE, line=NAVY, lw=4.5, shape="hex", adj=[0.28, None])
    S.box(182, 354, 108, 94, fill=WHITE, line=NAVY2, lw=1.0, shape="hex", adj=[0.28, None],
          text="대학혁신\n지원사업\n('25~'27)", size=12, bold=True, color=NAVY, margin=(0, 0, 0, 0))
    for (x, y, t) in [(178, 232, "자율성과\n책무성의\n조화"), (72, 412, "미래사회\n대비\n교육혁신"),
                      (300, 412, "적정 규모화\n통한 경쟁력\n강화")]:
        S.box(x, y, 113, 113, fill=WHITE, line=NAVY, lw=4.5, shape="oval", text=t, size=11, bold=True,
              color=NAVY, margin=(0, 0, 0, 0))
    S.box(412, 312, 62, 140, fill=("BFBFBF", "F2F2F2", 0), shape="chevron", adj=[0.55])
    # right side
    S.line(482, 280, 482, 490, color=NAVY2, lw=1.0, dash="dot")
    rows = [
        (240, "자율성과\n책무성의\n조화",
         ["대학별 혁신성과에 따른 성과기반 차등지원 체계\n  강화사업비 집행 자율화",
          "재정지원사업의 자율운영 및 성과관리 고도화 추진",
          "대학의 자체평가·성과관리체계 확립 지원"]),
        (340, "미래사회\n대비\n교육혁신",
         ["AI·디지털 전환 대응 교육과정 및 학사구조 개편",
          "융합형·현장맞춤형 인재양성 체계 강화",
          "대학 중장기 발전계획 기반 자율혁신 로드맵 수립"]),
        (440, "적정규모화\n통한 경쟁력\n강화",
         ["자발적 구조개선 및 적정규모화 노력 지원",
          "지역혁신중심대학(RISE) 연계, 지역 거점대학 역할 강화",
          "대학의 지속가능 경영체제 전환 및 인프라 효율화 추진"]),
    ]
    for y, lab, its in rows:
        S.box(545, y - 3, 350, 86, fill=WHITE, line="C9D3E0", shape="round", radius=0.15)
        S.line(482, y + 40, 497, y + 40, color=NAVY2, lw=1.0, dash="dot")
        S.box(478, y + 36, 8, 8, fill=NAVY2, shape="oval")
        S.box(497, y + 4, 90, 74, fill=("1F4E8C", "0B2A5B", 90), shape="pentagon", adj=[0.2], text=lab,
              size=10, bold=True, color=WHITE, align="c", margin=(0, 0, 8, 0), spacing=1.0)
        S.text(590, y - 3, 305, 86, "\n".join("· " + it for it in its), size=9.5, color=TXT, valign="m",
               spacing=1.05)
    # bottom
    S.box(65, 578, 830, 98, fill=WHITE, line=NAVY2, lw=1.5, shape="round", radius=0.12)
    S.box(346, 556, 220, 30, fill=NAVY, shape="round", radius=0.3, text="혁신사업 기본운영 방향", size=12.5,
          bold=True)
    S.text(95, 590, 380, 80, "· 대학의 자율성과 책임성을 기반으로 한 성과중심 지원체계 구축\n· 대학별 여건·발전단계에 맞는 자율혁신 지원",
           size=9.5, valign="m", spacing=1.3)
    S.text(475, 590, 415, 80, "· 자발적 적정규모화와 구조개선 유도를 통한 지속가능한 대학체제 확립\n"
           "· AI·디지털·지역혁신(RISE)과의 연계 강화로 정책 간 시너지 창출\n· 대학혁신 성과의 환류 및 확산체계 내실화 추진",
           size=9.5, valign="m", spacing=1.2)
