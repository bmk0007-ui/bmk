from lib import *


def build(S):
    S.header(1, "외부 환경 분석", sub="PEST+R 분석 결과 및 시사점", sub_w=300)
    S.panel(h=566)
    S.box(65, 148, 828, 25, fill=("6B6B6B", "9A9A9A", 90), shape="round", radius=0.5)
    for x, w, t in [(65, 113, "구분"), (178, 357, "핵심 변화 요인"), (535, 358, "시사점")]:
        S.text(x, 148, w, 25, t, size=11.5, bold=True, color=WHITE, align="c", valign="m")
    S.line(178, 151, 178, 170, color="D9D9D9", lw=0.75)
    S.line(535, 151, 535, 170, color="D9D9D9", lw=0.75)
    rows = [
        (180, 102, "정책(P)",
         "- 대학혁신지원사업·RISE 등 정부정책의 자율성과\n  지역연계 강화 기조\n- 대학 구조개혁 등 고등교육 체제 개편 본격화",
         "→ 정부 정책 변화에 대응한 자율혁신 기반 대학\n    운영체계 정립 및 지역연계형 발전모델 구축 필요"),
        (288, 79, "경제(E)",
         "- 경기 둔화·산업구조 전환·재정여력 한계 지속\n- 중소기업 경기 부진, 산학협력 수요 감소",
         "→ 산학협력·평생직업교육 중심의 혁신모델 구축\n    및 재정 다변화 전략 강화"),
        (373, 78, "사회(S)",
         "- 저출산·고령화·인구 감소, 노동·삶의 질 변화\n- 청년고용 불안·평생교육 수요 확대",
         "→ 학령인구 감소 대응형 교육혁신, 평생학습·복지\n    ·문화 융합형 대학체제 전환 필요"),
        (457, 78, "기술(T)",
         "- AI·양자·로봇·클린테크 등 첨단기술 확산\n- 기술안보·데이터 거버넌스 등 정책 이슈 부상",
         "→ AI·디지털 융합교육 강화, 첨단기술 기반 연구 및\n    지속가능 혁신체계 구축 필요"),
        (541, 78, "지역(R)",
         "- 충남 북부권 산업집중, 남부권 인구이탈로 지역\n  불균형 심화\n- 고령화 및 외국인 유입 증가",
         "→ 지역산업 맞춤형 전공개편·산학연계 강화,\n    지역평생교육 및 사회공헌 확대 필요"),
    ]
    for y, h, lab, mid, right in rows:
        S.box(70, y + h / 2 - 17, 100, 34, fill=WHITE, line=NAVY2, lw=1.5, shape="pentagon", adj=[0.25],
              text=lab, size=12.5, bold=True, color=NAVY)
        S.box(180, y, 350, h, fill=WHITE, line="9DB3D0", shape="round", radius=0.12)
        S.text(190, y, 338, h, mid, size=11, valign="m", spacing=1.2)
        S.box(540, y, 353, h, fill=WHITE, line="9DB3D0", shape="round", radius=0.12)
        S.text(545, y, 345, h, right, size=11.5, bold=True, color="1A2B4A", valign="m", spacing=1.3)
    S.box(65, 625, 828, 65, fill=("1F4E8C", "0B2A5B", 90),
          text="정책·경제·사회·기술·지역 전반의 환경 변화는 대학의 기존 체제를 근본적으로 재구조화 할 것을 요구하고 있음. 대학은 AI·디지털\n"
               "융합교육, 산학협력 기반 재정 다변화, 지역산업 연계형 평생학습 생태계 구축을 통해 지속가능한 성장과 사회적 책무를 실현하는 혁신\n"
               "허브(innovation hub) 로 전환해야 함",
          size=10, color=WHITE, align="l", margin=(22, 0, 4, 0), spacing=1.0)
