from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

doc = Document()

# Page margins
section = doc.sections[0]
section.top_margin = Cm(2.5)
section.bottom_margin = Cm(2.5)
section.left_margin = Cm(3.0)
section.right_margin = Cm(2.5)

def set_font(run, name='맑은 고딕', size=11, bold=False, color=None):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)
    r = run._r
    rPr = r.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:eastAsia'), name)
    rFonts.set(qn('w:ascii'), name)
    rFonts.set(qn('w:hAnsi'), name)
    rPr.insert(0, rFonts)

def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    if level == 0:
        set_font(run, size=18, bold=True)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(12)
    elif level == 1:
        set_font(run, size=13, bold=True)
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        # Add bottom border
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '6')
        bottom.set(qn('w:space'), '1')
        bottom.set(qn('w:color'), '2E74B5')
        pBdr.append(bottom)
        pPr.append(pBdr)
    elif level == 2:
        set_font(run, size=11, bold=True)
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
    return p

def add_normal(doc, text, indent=False):
    p = doc.add_paragraph()
    if indent:
        p.paragraph_format.left_indent = Cm(0.5)
    run = p.add_run(text)
    set_font(run, size=10.5)
    p.paragraph_format.space_after = Pt(2)
    return p

def add_bullet(doc, text, indent_cm=0.5):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(indent_cm)
    p.paragraph_format.first_line_indent = Cm(-0.3)
    run = p.add_run(f'• {text}')
    set_font(run, size=10.5)
    p.paragraph_format.space_after = Pt(1)
    return p

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def set_cell_border(cell, top=None, bottom=None, left=None, right=None):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        if val:
            border = OxmlElement(f'w:{side}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), '4')
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), val)
            tcBorders.append(border)
    tcPr.append(tcBorders)

# ===========================
# TITLE
# ===========================
title_p = doc.add_paragraph()
title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
title_p.paragraph_format.space_before = Pt(6)
title_p.paragraph_format.space_after = Pt(16)
r = title_p.add_run('캡스톤디자인 진행 보고서')
set_font(r, size=20, bold=True, color=(46, 116, 181))

sub_p = doc.add_paragraph()
sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub_p.paragraph_format.space_after = Pt(18)
r2 = sub_p.add_run('출석 인정 및 주차별 진행 상황 보고')
set_font(r2, size=11, color=(89, 89, 89))

# ===========================
# 1. 교과목 정보
# ===========================
add_heading(doc, '1. 교과목 정보')

info_table = doc.add_table(rows=5, cols=2)
info_table.style = 'Table Grid'
info_table.alignment = WD_TABLE_ALIGNMENT.LEFT

col_widths = [Cm(4), Cm(11)]
for row in info_table.rows:
    for i, cell in enumerate(row.cells):
        cell.width = col_widths[i]

info_data = [
    ('교과목명', '캡스톤디자인'),
    ('제출 목적', '출석 인정 및 주차별 진행 상황 보고'),
    ('팀명', '(추후 확정 예정)'),
    ('지도교수', '○○○ 교수님'),
    ('작성자', '김태성 (모바일 앱 개발 담당)'),
]
for i, (key, val) in enumerate(info_data):
    row = info_table.rows[i]
    row.cells[0].text = ''
    row.cells[1].text = ''
    r1 = row.cells[0].paragraphs[0].add_run(key)
    set_font(r1, size=10.5, bold=True)
    set_cell_bg(row.cells[0], 'EBF3FB')
    r2 = row.cells[1].paragraphs[0].add_run(val)
    set_font(r2, size=10.5)

doc.add_paragraph()

# ===========================
# 2. 팀 구성 및 역할
# ===========================
add_heading(doc, '2. 팀 구성 및 역할')

team_table = doc.add_table(rows=6, cols=3)
team_table.style = 'Table Grid'
team_table.alignment = WD_TABLE_ALIGNMENT.LEFT

headers = ['역할', '이름', '담당 분야']
header_row = team_table.rows[0]
for i, h in enumerate(headers):
    header_row.cells[i].text = ''
    r = header_row.cells[i].paragraphs[0].add_run(h)
    r.font.bold = True
    set_font(r, size=10.5, bold=True, color=(255, 255, 255))
    set_cell_bg(header_row.cells[i], '2E74B5')
    header_row.cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

team_data = [
    ('팀장', '김미현', '프로젝트 총괄 및 일정 관리'),
    ('하드웨어 / 센서 / 서기', '백승엽', '하드웨어 설계, 데이터 수집, 회의 기록'),
    ('모바일 앱 개발', '김태성', '스마트워치 데이터 수신 앱 및 복지사용 앱 개발'),
    ('네트워크', '왕상민', '통신 및 네트워크 구성, 중앙 서버 연동'),
    ('AI', '신수형', '인공지능 모델 설계 및 데이터 분석'),
]
for i, (role, name, field) in enumerate(team_data):
    row = team_table.rows[i + 1]
    for j, val in enumerate([role, name, field]):
        row.cells[j].text = ''
        r = row.cells[j].paragraphs[0].add_run(val)
        set_font(r, size=10.5)
    if i % 2 == 0:
        for cell in row.cells:
            set_cell_bg(cell, 'F5F9FD')

doc.add_paragraph()

# ===========================
# 3. 프로젝트 개요
# ===========================
add_heading(doc, '3. 프로젝트 개요')
overview = (
    '본 팀은 IoT 및 AI 기술을 활용한 캡스톤디자인 프로젝트를 진행하고 있다. '
    '초고령화 사회에서 증가하는 독거노인·만성질환 노인을 대상으로, '
    '스마트워치를 통해 일상생활 움직임 데이터(3축 가속도·각속도)를 수집하고, '
    'AI 모델이 이상 패턴을 감지하여 사회복지사에게 실시간 알림을 전달하는 '
    '"이동건강 상태 변화 모니터링 시스템"을 구현한다. '
    '주요 구성요소는 갤럭시 워치(데이터 수집) → 스마트폰(Bluetooth 수신) → '
    'Jetson Nano(AI 분석) → 중앙 서버(PostgreSQL) → 복지사 모바일 앱(알림)으로 이루어진다.'
)
p = add_normal(doc, overview)

doc.add_paragraph()

# ===========================
# 4. 주차별 진행 내용
# ===========================
add_heading(doc, '4. 주차별 진행 내용')

weeks = [
    {
        'date': '3월 11일 (1주차)',
        'items': [
            '캡스톤디자인 경진대회(산학협력학회 / 산학기술학회) 일정 확인 – 접수 10월 초~11월 중',
            '2023~2025년 입상작 주제·일정 조사 및 트렌드 분석 과제 부여',
            '팀 구성 및 역할 분담 완료 (팀장: 김미현, 하드웨어/서기: 백승엽, 앱: 김태성, 네트워크: 왕상민, AI: 신수형)',
            '관심 분야: 의료·사회복지 분야 우선 탐색',
            '논문 마감 기한 확인: 10월 31일 / 3페이지 제한',
        ]
    },
    {
        'date': '3월 18일 (3주차)',
        'items': [
            '개인별 키워드 지정: 김태성(공장), 백승엽(의료), 신수형(의료), 왕상민(사회문제), 김미현(사회문제)',
            '객관적 데이터 기반 발표 준비 요구 (통계 자료·실제 수치 활용)',
            '팀 운영 방식 확정: 팀장 주 2회 진행상황 체크 → 서기 취합·발표',
            '주제 범위: 공장·의료·사회문제 3개 키워드 내 변경 가능',
        ]
    },
    {
        'date': '3월 25일 (4주차)',
        'items': [
            '[김태성] 주제 변경(공장→의료): 제조업 사고 원인 분석, 기존 제품과의 차별화 방향 탐색',
            '[백승엽] 만성질환 고령인 모니터링 문제 제시 (사망 원인 80%, 고령자 의료비 43%)',
            '[신수형] 급성 심정지 대응 드론 아이디어 → AED 탑재 불가로 캡스톤 진행 불가 판정',
            '[김미현] 독거노인 증가 문제 제시, 각도 조절 카메라·센서 기반 모니터링 아이디어',
            '피드백: 실현 가능성보다 아이디어와 객관적 근거 중심',
        ]
    },
    {
        'date': '4월 1일 (5주차)',
        'items': [
            '사회복지 차원의 노인 보행 모니터링 현황 자료 조사 과제 부여',
            'ICEE 캡스톤디자인 연계사업 신청 안내 (다학제융합형 우선, 탈락 시 산학연계형 재신청)',
            '예산: H/W 130만원 / S/W 90만원 이내, 1차 신청 4월 14일 / 2차 신청 4월 27일',
            '[교수님 제안 주제] 일상생활 데이터와 강건한 모델링을 통한 이동건강 상태 변화 모니터링 시스템',
            '[왕상민] 자연재해·식량 안보 주제 발표 – 목적 불명확 등 피드백 수렴',
            '[김태성] 시각장애인 이동 및 의료 접근성 향상 IoT 시스템 주제 발표',
        ]
    },
    {
        'date': '4월 8일 (6주차)',
        'items': [
            '교수님 제안 주제로 프로젝트 방향 확정: "이동건강 상태 변화 모니터링 시스템"',
            '4월 10일 신청서 제출',
            '[백승엽] 스마트워치 탐색: 가속도·자이로 센서 필수, SDK 제공 모델, 4월 9일까지 교수님 이메일 제출',
            '[김미현] 시스템 구성도 구상: 워치→베이스스테이션→서버 흐름, 4월 9일까지 교수님 이메일 제출',
            '[신수형] 수집 데이터 AI 학습: 움직임 패턴 분석, 개인 맞춤 패턴 분석 검토',
        ]
    },
    {
        'date': '4월 15일 (7주차)',
        'items': [
            '캡스톤디자인 신청서 반려 – 지원금으로 스마트워치 구매 불가 판정',
            '예산 수정 및 재제출: Jetson Nano 50만원 + SSD 1TB 40만원 구매 예정',
            '시스템 구성도 수정: 블루투스 아이콘 제거, 이상 감지→알림 생성 화살표 추가, 복지사 단말기 연결',
            '[신수형] Nvidia Jetson Nano 활용 AI 모델 구동 제안',
            '[백승엽] 스마트워치 데이터 전달 API 탐색',
            '[김태성] 3축 가속도·각속도(총 6개 정보) 출력 테스트 앱 제작, 스마트폰→Jetson Nano 전송 방법 탐색',
            '[왕상민] Jetson Nano → 서버 데이터 전달 방법 탐색, DB 구축',
        ]
    },
    {
        'date': '4월 22일 (8주차)',
        'items': [
            '교수님 소유 Galaxy Watch 2개, 김미현님 소유 Galaxy Watch 7 1개 – 실착용 데이터 수집 예정',
            '[백승엽] 교수님 연구실에서 워치 2개 수령, Samsung SDK 권한 확보, API 데이터 수집 방법 탐색 후 김태성에게 전달',
            '[김태성] 6개 센서 정보(가속도 3축·각속도 3축) 출력 테스트 앱 제작, 스마트폰→Jetson Nano 전송 방법 탐색',
            '[신수형] 워치→Jetson Nano 직접 전송 가능성 확인, 더미 데이터 활용 서버 분석 테스트',
            '[왕상민] Jetson Nano→서버 전송 방법 탐색, DB 구축 (SQLite / PostgreSQL)',
            '[김미현] 참고문헌서 작성, 예산 계산',
        ]
    },
    {
        'date': '5월 13일 (11주차) — 5월 6일 수업 없음',
        'items': [
            '미팅 2회 남음 (5월 20일, 5월 27일)',
            '5월 27일 이후 프로젝트 지속 여부 팀장에게 의견 전달 (6월 초 마감)',
            '지속 시: 7~8월 온라인 진행, 9월 테스트 기간, 학술대회 일정 맞춤',
            '작업 막힌 부분은 포장 없이 그대로 제출, 납득 가능한 근거 첨부',
            '[김태성] 블루투스 재연결 자동화 기능 구현, 젯슨나노 외부 접속 중계 방법 탐색, 인터뷰(복지사 면담) 기획',
            '[왕상민] 젯슨나노→중앙서버 연결 방법 탐색, DB 선정',
            '[백승엽] 스마트폰↔젯슨나노 통신 방법 탐색 및 시행착오 발표 준비',
            '[신수형] 공개 데이터로 AI 모델 학습 진행',
            '[김미현] DB 조사 – Firebase 유료화로 SQLite(Jetson Nano) / PostgreSQL(중앙 서버) 결정',
        ]
    },
    {
        'date': '5월 20일 (12주차)',
        'items': [
            '5월 27일 최종 발표 준비: 지금까지 작업 내용 PPT로 정리, 활동 제출용 자료(PPT/PDF) 준비',
            '교수님 제시 가이드라인 숙지 및 발표 점검',
            '[김태성] 스토리보드 제작 우선, 사회복지사 면담(앱 필요 기능 파악), 면담 불가 시 대체 근거 준비',
            '[김미현] 5월 21일 오후 4시 교수님 미팅 예정',
        ]
    },
]

for week in weeks:
    add_heading(doc, week['date'], level=2)
    for item in week['items']:
        add_bullet(doc, item)

doc.add_paragraph()

# ===========================
# 5. 현재 진행 성과
# ===========================
add_heading(doc, '5. 현재 진행 성과 (12주차 기준)')

achievements = [
    '프로젝트 주제 확정: "일상생활 데이터와 강건한 모델링을 통한 이동건강 상태 변화 모니터링 시스템"',
    '시스템 구성도 완성: Galaxy Watch → 스마트폰(BT) → Jetson Nano(AI) → PostgreSQL 서버 → 복지사 앱',
    'Samsung Health SDK 권한 확보 및 가속도·자이로 데이터 수집 API 탐색 완료',
    '6축 센서 데이터(3축 가속도 + 3축 각속도) 출력 테스트 앱 제작 진행',
    'AI 모델: 공개 데이터 기반 이상 움직임 패턴 학습 진행 (Jetson Nano 탑재 예정)',
    'DB 아키텍처 결정: Jetson Nano(SQLite) + 중앙 서버(PostgreSQL)',
    '하드웨어 예산 확정: Jetson Nano 50만원 + SSD 1TB 40만원 (총 130만원 이내)',
    'Galaxy Watch 3개 확보 (교수님 2개, 팀원 1개) – 실착용 데이터 수집 준비 완료',
]
for a in achievements:
    add_bullet(doc, a)

doc.add_paragraph()

# ===========================
# 6. 문제점 및 개선 방향
# ===========================
add_heading(doc, '6. 문제점 및 개선 방향')

issues = [
    ('캡스톤 지원금 신청 반려', 'Jetson Nano·SSD 중심으로 예산 재구성 후 재신청'),
    ('Jetson Nano 외부 접속 불가', '중계 서버(Reverse Proxy / VPN 등) 방법 탐색 중'),
    ('스마트워치 데이터 전송 경로 불확실', 'BT 페어링 앱 → 스마트폰 → Jetson Nano 경로 확정, 직접 전송 가능성도 병행 검토'),
    ('복지사 앱 기능 명세 부재', '5월 27일 이전 스토리보드 작성 및 사회복지사 면담을 통한 요구사항 수집 예정'),
    ('AI 모델 개인화 학습 구현 시기 미결', '기본 이상 패턴 탐지 모델 먼저 완성 후 개인 맞춤 기능 단계적 추가'),
]

issue_table = doc.add_table(rows=len(issues) + 1, cols=2)
issue_table.style = 'Table Grid'
header_row = issue_table.rows[0]
for i, h in enumerate(['문제점', '개선 방향']):
    header_row.cells[i].text = ''
    r = header_row.cells[i].paragraphs[0].add_run(h)
    set_font(r, size=10.5, bold=True, color=(255, 255, 255))
    set_cell_bg(header_row.cells[i], '2E74B5')
    header_row.cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

for i, (prob, sol) in enumerate(issues):
    row = issue_table.rows[i + 1]
    row.cells[0].text = ''
    row.cells[1].text = ''
    r1 = row.cells[0].paragraphs[0].add_run(prob)
    set_font(r1, size=10.5, bold=True)
    r2 = row.cells[1].paragraphs[0].add_run(sol)
    set_font(r2, size=10.5)
    if i % 2 == 0:
        set_cell_bg(row.cells[0], 'F5F9FD')
        set_cell_bg(row.cells[1], 'F5F9FD')

doc.add_paragraph()

# ===========================
# 7. 향후 계획
# ===========================
add_heading(doc, '7. 향후 계획')

plans = [
    ('5월 27일', '지금까지 작업 내용 PPT 발표 및 활동 자료 제출'),
    ('6월 초', '프로젝트 지속 여부 최종 결정'),
    ('7월 ~ 8월', '온라인(Zoom) 집중 개발: 앱·AI 모델·서버 통합 구현'),
    ('9월', '시스템 통합 테스트 및 실착용 데이터 검증'),
    ('10월 초', '시스템 완성 및 논문 초안 작성'),
    ('10월 31일', '학술 논문 마감 (3페이지)'),
    ('10월 초 ~ 11월 중', '산학협력/산학기술학회 캡스톤디자인 경진대회 접수'),
]

plan_table = doc.add_table(rows=len(plans) + 1, cols=2)
plan_table.style = 'Table Grid'
header_row = plan_table.rows[0]
for i, h in enumerate(['일정', '내용']):
    header_row.cells[i].text = ''
    r = header_row.cells[i].paragraphs[0].add_run(h)
    set_font(r, size=10.5, bold=True, color=(255, 255, 255))
    set_cell_bg(header_row.cells[i], '2E74B5')
    header_row.cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

for i, (date, content) in enumerate(plans):
    row = plan_table.rows[i + 1]
    row.cells[0].text = ''
    row.cells[1].text = ''
    r1 = row.cells[0].paragraphs[0].add_run(date)
    set_font(r1, size=10.5, bold=True)
    r2 = row.cells[1].paragraphs[0].add_run(content)
    set_font(r2, size=10.5)
    row.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    if i % 2 == 0:
        set_cell_bg(row.cells[0], 'EBF3FB')
        set_cell_bg(row.cells[1], 'F5F9FD')

# Save
out_path = '/home/user/IOT-/캡스톤디자인_진행보고서_완성.docx'
doc.save(out_path)
print(f'저장 완료: {out_path}')
