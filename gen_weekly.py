"""
캡스톤디자인 주차별 진행 보고서 생성기
템플릿 docx의 모든 스타일/설정 파일을 그대로 재사용하고,
word/document.xml 만 주차별 내용으로 교체.
"""
import zipfile, shutil, os

TEMPLATE = '/root/.claude/uploads/3cd1354f-e18f-4a3b-9b49-754487dc8a3f/bc9ceda2-_________________.docx'
OUT_DIR  = '/home/user/IOT-/주차별_보고서'
os.makedirs(OUT_DIR, exist_ok=True)

# ─────────────────────────────────────────────
# XML 헬퍼
# ─────────────────────────────────────────────
def h1(text):
    return f'''<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>'''

def h2(text):
    return f'''<w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>'''

def para(text):
    return f'''<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'''

def bullet(text):
    """글머리 기호 단락 (numbering.xml의 첫 번째 목록 스타일 사용)"""
    return (
        f'<w:p>'
        f'<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
        f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'
        f'</w:p>'
    )

def empty():
    return '<w:p/>'

def table2(rows):
    """rows = [(col1, col2), ...]  2-column simple table"""
    W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    trs = ''
    for c1, c2 in rows:
        trs += (
            f'<w:tr>'
            f'<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>'
            f'<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">{c1}</w:t></w:r></w:p></w:tc>'
            f'<w:tc><w:tcPr><w:tcW w:w="7000" w:type="dxa"/></w:tcPr>'
            f'<w:p><w:r><w:t xml:space="preserve">{c2}</w:t></w:r></w:p></w:tc>'
            f'</w:tr>'
        )
    return (
        f'<w:tbl>'
        f'<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="9000" w:type="dxa"/></w:tblPr>'
        f'<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="7000"/></w:tblGrid>'
        f'{trs}</w:tbl>'
    )

SECTION = '''<w:sectPr w:rsidR="00FC693F" w:rsidRPr="0006063C" w:rsidSect="00034616">
<w:pgSz w:w="12240" w:h="15840"/>
<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0"/>
<w:cols w:space="720"/>
<w:docGrid w:linePitch="360"/>
</w:sectPr>'''

NS = (
    'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
    'xmlns:mo="http://schemas.microsoft.com/office/mac/office/2008/main" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:mv="urn:schemas-microsoft-com:mac:vml" '
    'xmlns:o="urn:schemas-microsoft-com:office:office" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:w10="urn:schemas-microsoft-com:office:word" '
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
    'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
    'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '
    'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'mc:Ignorable="w14 wp14"'
)

def make_document_xml(week):
    w = week
    body_parts = []

    # ── 제목 ──
    body_parts.append(h1('캡스톤디자인 주차별 진행 보고서'))
    body_parts.append(empty())

    # ── 기본 정보 헤더 박스 (표) ──
    body_parts.append(table2([
        ('주차',    f"{w['week_no']}주차"),
        ('기간',    w['period']),
        ('팀명',    '(추후 확정 예정)'),
        ('작성자',  '김태성'),
        ('지도교수', '○○○ 교수님'),
    ]))
    body_parts.append(empty())

    # ── 1. 기본 정보 ──
    body_parts.append(h2('1. 기본 정보'))
    for line in w['basic_info']:
        body_parts.append(bullet(line))
    body_parts.append(empty())

    # ── 2. 이번 주 목표 ──
    body_parts.append(h2('2. 이번 주 목표'))
    for line in w['goals']:
        body_parts.append(bullet(line))
    body_parts.append(empty())

    # ── 3. 진행 내용 ──
    body_parts.append(h2('3. 진행 내용'))
    for line in w['progress']:
        body_parts.append(bullet(line))
    body_parts.append(empty())

    # ── 4. 개인별 수행 작업 ──
    body_parts.append(h2('4. 개인별 수행 작업'))
    body_parts.append(table2(w['individual']))
    body_parts.append(empty())

    # ── 5. 결과 및 성과 ──
    body_parts.append(h2('5. 결과 및 성과'))
    for line in w['results']:
        body_parts.append(bullet(line))
    body_parts.append(empty())

    # ── 6. 문제점 및 해결 과정 ──
    body_parts.append(h2('6. 문제점 및 해결 과정'))
    for line in w['issues']:
        body_parts.append(bullet(line))
    body_parts.append(empty())

    # ── 7. 다음 주 계획 ──
    body_parts.append(h2('7. 다음 주 계획'))
    for line in w['next_plan']:
        body_parts.append(bullet(line))
    body_parts.append(empty())

    # ── 8. 첨부 자료 ──
    body_parts.append(h2('8. 첨부 자료'))
    body_parts.append(para(w.get('attachments', '해당 없음')))

    body_xml = ''.join(body_parts)
    return (
        f"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        f"<w:document {NS}>"
        f"<w:body>{body_xml}{SECTION}</w:body>"
        f"</w:document>"
    )


# ─────────────────────────────────────────────
# 주차별 데이터
# ─────────────────────────────────────────────
WEEKS = [

    # ── 1주차 ──
    dict(
        week_no=1, period='2026.03.11 ~ 2026.03.17',
        filename='01주차_2026-03-11_진행보고서.docx',
        basic_info=[
            '교과목: 캡스톤디자인 (매주 수요일 10:45~12:00)',
            '팀 구성 및 역할 분담 완료',
            '관심 분야: 의료·사회복지 분야 우선 탐색',
            '경진대회: 산학협력학회/산학기술학회 캡스톤디자인 경진대회 (접수 10월 초~11월 중)',
            '논문 마감: 10월 31일 (3페이지 제한)',
        ],
        goals=[
            '캡스톤디자인 경진대회 일정 파악',
            '2023~2025년 입상작 주제·일정·상금 조사 및 트렌드 분석',
            '팀 구성 및 역할 확정',
            '팀명 결정 (다음 주까지)',
        ],
        progress=[
            '캡스톤디자인 경진대회(산학협력학회/산학기술학회) 일정 확인',
            '입상작(2023~2025) 조사 과제 부여 – 주제·일정·상금 표 정리 후 발표',
            '팀 구성 완료: 팀장 김미현, 하드웨어/센서/서기 백승엽, 모바일 앱 김태성, 네트워크 왕상민, AI 신수형',
            'IoT + AI 융합 프로젝트 방향 설정',
            '의료·사회복지 분야를 주요 탐색 분야로 설정',
        ],
        individual=[
            ('김미현 (팀장)', '팀 총괄 및 일정 관리, 카톡방 개설'),
            ('백승엽 (HW/서기)', '회의록 작성, 하드웨어 역할 파악'),
            ('김태성 (앱)', '모바일 앱 개발 역할 확인 및 기술 스택 탐색'),
            ('왕상민 (네트워크)', '네트워크 구성 역할 파악'),
            ('신수형 (AI)', 'AI 알고리즘 역할 파악'),
        ],
        results=[
            '팀 구성 및 역할 분담 완료',
            '경진대회 일정 파악 완료',
            'IoT+AI 프로젝트 방향 초안 설정',
        ],
        issues=[
            '팀명 미결정 → 다음 주까지 확정 예정',
            '프로젝트 주제 방향성 미확정 → 입상작 조사 후 트렌드 분석으로 구체화 예정',
        ],
        next_plan=[
            '팀명 확정 (3월 25일 10시 회의)',
            '2023~2025년 입상작 트렌드 분석 결과 발표',
            '개인별 탐색 키워드 지정',
        ],
    ),

    # ── 3주차 ──
    dict(
        week_no=3, period='2026.03.18 ~ 2026.03.24',
        filename='03주차_2026-03-18_진행보고서.docx',
        basic_info=[
            '3주차 수업 (2주차 수업 없음)',
            '팀 운영 방식 확정',
            '3월 25일 10시~11시 팀명 확정 회의 예정',
        ],
        goals=[
            '개인별 탐색 키워드 지정 및 초기 아이디어 수집',
            '팀 운영 방식(주간 체크, 서기 역할) 확정',
            '객관적 자료 기반 발표 준비 방향 수립',
        ],
        progress=[
            '교수님이 개인별 탐색 키워드 제시: 김태성(공장), 백승엽(의료), 신수형(의료), 왕상민(사회문제), 김미현(사회문제)',
            '팀 운영 방식 확정: 팀장 주 2회(화·금) 진행상황 체크 → 서기 취합·발표',
            '주제 범위: 공장·의료·사회문제 3개 키워드 내 변경 가능',
            '발표 방향: 교수님 질문에 대응할 객관적 근거 데이터 기반 준비',
            '3월 25일 10시 회의에서 팀명 확정 예정',
        ],
        individual=[
            ('김미현 (팀장)', '사회문제 키워드 탐색, 주 2회 진행상황 체크'),
            ('백승엽 (HW/서기)', '의료 키워드 탐색, 회의록 통합 요약'),
            ('김태성 (앱)', '공장 키워드 탐색, 앱 구동 방식 및 기능 구상'),
            ('왕상민 (네트워크)', '사회문제 키워드 탐색'),
            ('신수형 (AI)', '의료 키워드 탐색, CCTV 활용 드론 아이디어 구상'),
        ],
        results=[
            '팀 운영 방식 확정 (주 2회 체크인, 서기 취합 발표)',
            '개인별 탐색 키워드 할당 완료',
            '다음 주 발표 방향 수립',
        ],
        issues=[
            '팀명 미결정 → 3월 25일 회의에서 확정',
            '아이디어 구체성 부족 → 객관적 통계 자료 기반 발표 준비 필요',
        ],
        next_plan=[
            '팀명 확정 (3/25 회의)',
            '개인별 키워드 기반 아이디어 + 객관적 근거 자료 발표',
            '공장·의료·사회문제 중 팀 주제 방향 압축',
        ],
    ),

    # ── 4주차 ──
    dict(
        week_no=4, period='2026.03.25 ~ 2026.03.31',
        filename='04주차_2026-03-25_진행보고서.docx',
        basic_info=[
            '4주차 수업',
            '팀 방향성 점검 및 개인별 아이디어 발표',
            '실현보다 아이디어 중심, 객관적 근거로 공감 이끌기',
        ],
        goals=[
            '개인별 아이디어 발표 및 피드백 수렴',
            '프로젝트 방향성 점검',
            '팀 주제 후보 도출',
        ],
        progress=[
            '[김태성] 제조업 사고 문제 발표 (3할이 제조업) → 기존 제품 존재 확인, 차별화 방향 탐색 과제',
            '[백승엽] 만성질환 고령인 모니터링 문제 발표 (사망 원인 80% 만성질환, 고령자 의료비 43%)',
            '[신수형] 급성 심정지 대응 드론 아이디어 발표 → AED 탑재 불가로 캡스톤 진행 불가 판정',
            '[김미현] 독거노인 증가 문제 발표 + 각도 조절 카메라·센서 기반 모니터링 아이디어',
            '[왕상민] 다음 주 발표 예정',
            '회의 활발한 참여 독려: 질문 응답, 아이디어·피드백 제시',
        ],
        individual=[
            ('김태성', '제조업 사고 통계 발표, 기존 제품 차별화 방향 추가 탐색'),
            ('백승엽', '만성질환 고령인 모니터링 문제 발표, 해결방안(왜·무엇을·어떻게) 추가 준비'),
            ('신수형', '드론 AED 아이디어 발표 → 공장 분야로 주제 변경'),
            ('김미현', '독거노인 모니터링 아이디어 발표, 객관적 자료 추가 준비'),
            ('왕상민', '다음 주 발표 준비'),
        ],
        results=[
            '팀 주제 후보 구체화: 의료·독거노인·공장 안전',
            '기존 제품 차별화 방향성 확인',
            '실현보다 아이디어 중심 접근 방향 확립',
        ],
        issues=[
            '신수형 드론 아이디어: AED 탑재 불가 → 공장 분야로 주제 변경',
            '발표 내용 구체성 부족 (왜·무엇을·어떻게 구조 보완 필요)',
            '사생활 침해 문제(카메라 설치) → 추가 검토 필요',
        ],
        next_plan=[
            '왕상민 발표 (사회문제)',
            '각자 피드백 반영하여 아이디어 보완',
            '노인 보행 모니터링 현황 자료 조사',
            '팀 최종 주제 방향 압축',
        ],
    ),

    # ── 5주차 ──
    dict(
        week_no=5, period='2026.04.01 ~ 2026.04.07',
        filename='05주차_2026-04-01_진행보고서.docx',
        basic_info=[
            '5주차 수업',
            'ICEE 캡스톤디자인 연계사업 신청 안내',
            '다학제융합형 우선 신청 (사회복지학과 인원 1명 추가)',
            '예산: H/W 130만원 이내 / S/W 90만원 이내',
            '1차 신청 4월 14일 / 2차 신청 4월 27일',
        ],
        goals=[
            '왕상민 발표 (사회문제)',
            '교수님 제안 주제 검토',
            'ICEE 신청 준비',
            '노인 보행 모니터링 자료 조사',
        ],
        progress=[
            '[왕상민] 자연재해·식량 안보 주제 발표 → 목적 불명확 등 피드백 수렴',
            '[김태성] 시각장애인 이동 및 의료 접근성 향상 IoT 시스템 발표',
            '[교수님 제안] "일상생활 데이터와 강건한 모델링을 통한 이동건강 상태 변화 모니터링 시스템"',
            '교수님: 현재 사회복지사가 직접 방문해 상태 체크 → 스마트워치로 움직임 데이터 수집 후 AI 분석 알람 시스템 구축',
            '피드백: 기존 시스템과 차별성 부족 문제 제기',
            'ICEE 연계사업 신청 일정 및 예산 확인',
        ],
        individual=[
            ('왕상민', '자연재해·식량 안보 주제 발표, 피드백 수렴 후 주제 보완'),
            ('김태성', '시각장애인 IoT 시스템 발표, 각 슬라이드 의미 보강 과제'),
            ('김미현', 'ICEE 신청서 준비 보조, 다학제 인원 섭외'),
            ('백승엽', '노인 보행 모니터링 현황 자료 조사'),
            ('신수형', '교수님 제안 주제 AI 적용 방향 검토'),
        ],
        results=[
            'ICEE 캡스톤디자인 연계사업 신청 정보 파악',
            '교수님 제안 주제 도출: 이동건강 상태 변화 모니터링 시스템',
            '예산 구조 확인 (H/W 130만원, S/W 90만원)',
        ],
        issues=[
            '왕상민 주제: 대피 시스템과 식량 안보 목적 불명확 → 주제 수정 필요',
            '기존 스마트워치 제품과 차별성 부족 → 시스템(연결 구조) 관점으로 차별화',
            '카톡·문자 사용 통계 자료 부족 → 보강 필요',
        ],
        next_plan=[
            '교수님 제안 주제로 팀 방향 확정 논의',
            '4월 10일 ICEE 신청서 제출',
            '스마트워치 탐색 (백승엽) 및 시스템 구성도 초안 (김미현)',
        ],
    ),

    # ── 6주차 ──
    dict(
        week_no=6, period='2026.04.08 ~ 2026.04.14',
        filename='06주차_2026-04-08_진행보고서.docx',
        basic_info=[
            '6주차 수업',
            '프로젝트 주제 최종 확정',
            '4월 10일 ICEE 신청서 제출',
        ],
        goals=[
            '교수님 제안 주제로 방향 확정',
            'ICEE 신청서 제출 (4/10)',
            '스마트워치 탐색 및 시스템 구성도 초안 작성',
        ],
        progress=[
            '주제 최종 확정: "일상생활 데이터와 강건한 모델링을 통한 이동건강 상태 변화 모니터링 시스템"',
            '4월 10일 신청서 제출 완료',
            '시스템 흐름: 스마트워치 데이터 수집 → 베이스스테이션(상시 인터넷 연결) 전달 → AI 분석 → 알림',
            '[백승엽] 스마트워치 탐색: 가속도·자이로 센서 필수, SDK 제공 모델, 심박수 가능 모델 우선, 4/9까지 제출',
            '[김미현] 시스템 구성도 구상: 워치→베이스스테이션→서버 흐름, 4/9까지 PPT 제출',
            '[신수형] 수집 데이터 AI 학습: 움직임 패턴 분석, 개인 맞춤 패턴 추후 검토',
        ],
        individual=[
            ('김미현', '시스템 구성도 PPT 작성 및 교수님 이메일 제출 (4/9)'),
            ('백승엽', '스마트워치 모델 탐색(가속도·자이로·심박·SDK) 한글 파일 제출 (4/9)'),
            ('김태성', '스마트워치 데이터 수신 앱 개발 방향 탐색'),
            ('왕상민', '베이스스테이션~서버 통신 방법 탐색'),
            ('신수형', 'AI 움직임 패턴 학습 방법 탐색'),
        ],
        results=[
            '프로젝트 주제 최종 확정',
            'ICEE 신청서 제출',
            '역할별 세부 과제 부여 완료',
        ],
        issues=[
            '스마트워치 SDK 접근 권한 확보 여부 불확실 → 탐색 후 판단',
            '시스템 구성도 구체화 필요 → 김미현 초안 기반 보완 예정',
        ],
        next_plan=[
            '스마트워치 탐색 결과 발표 (백승엽)',
            '시스템 구성도 발표 및 피드백 (김미현)',
            '역할별 세부 기술 탐색 진행',
        ],
    ),

    # ── 7주차 ──
    dict(
        week_no=7, period='2026.04.15 ~ 2026.04.21',
        filename='07주차_2026-04-15_진행보고서.docx',
        basic_info=[
            '7주차 수업',
            'ICEE 신청서 반려 (지원금으로 스마트워치 구매 불가)',
            '예산 재구성: Jetson Nano 50만원 + SSD 1TB 40만원 (총 130만원)',
            '시스템 구성도 수정 완료',
        ],
        goals=[
            '신청서 수정 및 재제출 준비',
            '시스템 구성도 보완',
            '역할별 기술 탐색 착수',
        ],
        progress=[
            'ICEE 신청서 반려: 지원금으로 스마트워치 구매 불가 통보',
            '예산 수정: Jetson Nano 50만원 + SSD 1TB 40만원으로 재구성, 신청서 수정 예정',
            '시스템 구성도 수정 사항: 블루투스 아이콘 제거, 이상 감지→알림 생성 화살표 추가, 사회복지사 단말기 연결 화살표 추가',
            '[신수형] Nvidia Jetson Nano를 이용한 AI 모델 구동 방안 제시',
            '[백승엽] 스마트워치 데이터 전달 API 탐색',
            '[김태성] 3축 가속도+각속도(총 6개 정보) 출력 테스트 앱 제작, 스마트폰→Jetson Nano 전송 방법 탐색',
            '[왕상민] Jetson Nano→서버 데이터 전달 방법 탐색, DB 구축 탐색',
        ],
        individual=[
            ('김미현', '시스템 구성도 수정, 도움 요청 대응, 참고문헌서 작성 준비'),
            ('백승엽', '스마트워치 데이터 전달 API 탐색'),
            ('김태성', '6축 센서(가속도3+각속도3) 테스트 앱 제작, 폰→Jetson 전송 방법 탐색'),
            ('왕상민', 'Jetson Nano→중앙서버 전송 방법 탐색, DB 구축'),
            ('신수형', 'Jetson Nano 활용 AI 모델 구동 방안 제안'),
        ],
        results=[
            '시스템 구성도 1차 수정 완료',
            '예산 재구성: Jetson Nano + SSD 1TB (130만원 이내)',
            '역할별 기술 탐색 착수',
        ],
        issues=[
            '지원금으로 스마트워치 구매 불가 → Jetson Nano·SSD 중심으로 예산 변경',
            '스마트워치 데이터 수집 방법 미확정 → API/SDK 탐색 진행 중',
        ],
        next_plan=[
            '교수님 연구실에서 Galaxy Watch 수령 (백승엽)',
            'Samsung Health SDK 권한 확보',
            '6축 센서 테스트 앱 완성 (김태성)',
            'Jetson Nano→서버 전송 방법 구체화 (왕상민)',
        ],
    ),

    # ── 8주차 ──
    dict(
        week_no=8, period='2026.04.22 ~ 2026.04.28',
        filename='08주차_2026-04-22_진행보고서.docx',
        basic_info=[
            '8주차 수업',
            'Galaxy Watch 3개 확보: 교수님 2개(Galaxy Watch), 김미현 1개(Galaxy Watch 7)',
            '5월 27일 수업 종료 예정 (이후 프로젝트 지속 여부 논의)',
            '워치+폰은 Jetson Nano 케이스에 수납하여 외부 접근 차단',
        ],
        goals=[
            'Galaxy Watch 수령 및 데이터 수집 환경 구축',
            '역할별 기술 스택 탐색 진행',
            '폰↔Jetson Nano 통신 방법 확정',
        ],
        progress=[
            '[백승엽] 교수님 연구실에서 Galaxy Watch 2개 수령, Samsung Health SDK 권한 확보 시도',
            '[백승엽] API 데이터 수집 방법 탐색 후 김태성에게 전달',
            '[김태성] 가속도 3축 + 각속도 3축 = 총 6개 센서 정보 출력 테스트 앱 제작',
            '[김태성] 스마트폰→Jetson Nano 데이터 전송 방법 탐색',
            '[신수형] 워치→Jetson Nano 직접 전송 가능성 확인, 더미 데이터 활용 서버 분석 테스트',
            '[왕상민] Jetson Nano→서버 전송 방법 탐색, DB 구축 (SQLite/PostgreSQL 검토)',
            '[김미현] 참고문헌서 작성, 예산 계산',
            '7~8월 온라인(Zoom) 진행, 5월 27일까지 의견 전달 요청',
        ],
        individual=[
            ('백승엽', 'Galaxy Watch 2개 수령, Samsung SDK 권한 확보, API 탐색 결과 공유'),
            ('김태성', '6축 센서 출력 테스트 앱 제작, 스마트폰→Jetson Nano 전송 탐색'),
            ('신수형', '워치→Jetson Nano 직접 전송 가능성 확인, 더미 데이터 서버 분석'),
            ('왕상민', 'Jetson Nano→서버 전송 탐색, DB 구조 설계'),
            ('김미현', '참고문헌서 작성, 예산 계산, 팀원 지원'),
        ],
        results=[
            'Galaxy Watch 3개 확보 (실착용 데이터 수집 준비 완료)',
            '6축 센서 테스트 앱 제작 진행',
            'DB 구조 탐색 시작',
        ],
        issues=[
            'Samsung SDK 권한 확보 불확실 → 확보 시 추가 센서 데이터 활용 가능',
            'Jetson Nano 외부 접속 불가 → 중계 서버 방법 탐색 필요',
            '5월 27일 이후 프로젝트 지속 여부 미결 → 6월 초까지 의견 전달',
        ],
        next_plan=[
            'Galaxy Watch 실착용 데이터 수집 테스트',
            '스마트폰↔Jetson Nano 통신 방법 확정',
            '5월 27일 이후 프로젝트 지속 여부 팀장에게 의견 전달',
        ],
    ),

    # ── 11주차 ──
    dict(
        week_no=11, period='2026.05.13 ~ 2026.05.19',
        filename='11주차_2026-05-13_진행보고서.docx',
        basic_info=[
            '11주차 수업 (9·10주차: 결원/공백, 5월 6일 수업 없음)',
            '남은 미팅: 5월 20일, 5월 27일 2회',
            '5월 27일 이후 프로젝트 지속 여부 팀장에게 6월 초까지 의견 전달',
            '지속 시: 7~8월 온라인(Zoom), 9월 테스트 기간, 학술대회 일정 맞춤',
        ],
        goals=[
            '각자 진행 상황 점검 및 발표',
            'DB 아키텍처 확정',
            '다음 주 최종 발표 준비 착수',
        ],
        progress=[
            '[김태성] BT 기기 재연결 자동화 기능 구현, Jetson Nano 외부 접속 중계 방법 탐색, 면담(복지사) 기획',
            '[왕상민] Jetson Nano→중앙서버 연결 방법 탐색, 통신에 적합한 DB 탐색',
            '[백승엽] 스마트폰↔Jetson Nano 통신 방법 탐색, 시행착오 과정 정리 발표 예정',
            '[신수형] 공개 데이터 기반 AI 모델 학습 진행',
            '[김미현] Firebase 유료화로 SQLite(Jetson Nano) / PostgreSQL(중앙 서버) 구조 결정',
            '작업 막힌 부분 포장 없이 그대로 제출, 납득 가능한 근거 첨부 원칙 확인',
        ],
        individual=[
            ('김태성', 'BT 자동 재연결 구현, Jetson Nano 외부접속 중계 탐색, 복지사 면담 기획'),
            ('왕상민', 'Jetson→서버 연결 탐색, DB 선정 (Firebase 유료화 이슈 처리)'),
            ('백승엽', '폰↔Jetson 통신 탐색, 시행착오 발표 준비'),
            ('신수형', '공개 데이터로 AI 모델 학습, 팀원과 정보 공유'),
            ('김미현', 'DB 결정: SQLite(Jetson) + PostgreSQL(서버), 교수님 질문 대응 지원'),
        ],
        results=[
            'DB 아키텍처 확정: Jetson Nano(SQLite) + 중앙 서버(PostgreSQL)',
            'Firebase → SQLite/PostgreSQL 전환 결정',
            '각자 역할별 기술 탐색 진행 중',
        ],
        issues=[
            'Firebase 유료화 → SQLite + PostgreSQL로 전환',
            'Jetson Nano 외부 접속 불가 → Reverse Proxy/VPN 등 중계 방법 탐색 중',
            '복지사 면담 일정 미확정 → 5월 27일 이전 스토리보드 제작 후 면담 진행',
        ],
        next_plan=[
            '5월 27일 최종 발표: 지금까지 작업 내용 PPT 정리',
            '활동 제출용 자료(PPT/PDF) 개인별 준비',
            '김태성: 스토리보드 제작 완료',
            '김미현: 5월 21일 오후 4시 교수님 미팅',
        ],
    ),

    # ── 12주차 ──
    dict(
        week_no=12, period='2026.05.20 ~ 2026.05.26',
        filename='12주차_2026-05-20_진행보고서.docx',
        basic_info=[
            '12주차 수업 (최종 주차)',
            '5월 27일 10시 30분 최종 발표',
            '발표 내용: 지금까지 작업 내용 PPT 정리 + 활동 제출용 자료(PPT/PDF)',
            '교수님 제시 가이드라인 숙지 및 발표 준비 완료',
        ],
        goals=[
            '5월 27일 최종 발표 준비 완료',
            '개인별 활동 제출용 자료 준비',
            '스토리보드 및 복지사 면담 진행',
        ],
        progress=[
            '[김태성] 스토리보드 제작 우선, 사회복지사 면담(앱 필요 기능 파악), 면담 불가 시 대체 근거 준비',
            '[김미현] 5월 21일 오후 4시 교수님 미팅 예정',
            '교수님 가이드라인: 작업한 내용 PDF/PPT로 정리 후 카톡 업로드, TV 연결 기기에서 발표',
            '프로젝트 완성보다 과정과 경험에 중점',
            '막힌 부분 포장 없이 그대로 제출, 납득 가능한 근거 첨부',
        ],
        individual=[
            ('김태성', '스토리보드 제작, 사회복지사 면담(또는 대체 근거), 발표 자료 준비'),
            ('김미현', '5/21 교수님 미팅, 팀 발표 자료 총괄'),
            ('백승엽', '폰↔Jetson 통신 탐색 결과 발표 자료 준비'),
            ('왕상민', 'Jetson→서버 연결 결과 발표 자료 준비'),
            ('신수형', 'AI 모델 학습 결과 발표 자료 준비'),
        ],
        results=[
            '12주차까지 누적 성과: 시스템 구성도 완성, Galaxy Watch 3개 확보, 6축 센서 앱 제작, DB 아키텍처 확정',
            '발표 자료 준비 진행 중',
            '스토리보드 제작 진행 중',
        ],
        issues=[
            '복지사 면담 불가 시 → 대체 근거(기존 복지 현황 자료 등) 준비',
            '통합 테스트 미완료 → 7~8월 온라인 집중 개발로 보완 예정',
        ],
        next_plan=[
            '5월 27일 최종 발표 (10:30 시작)',
            '6월 초: 프로젝트 지속 여부 최종 결정',
            '7~8월: 온라인(Zoom) 집중 개발 (앱·AI·서버 통합)',
            '9월: 시스템 통합 테스트',
            '10월 31일: 학술 논문 마감 (3페이지)',
        ],
    ),
]


# ─────────────────────────────────────────────
# 파일 생성
# ─────────────────────────────────────────────
with zipfile.ZipFile(TEMPLATE, 'r') as template_zip:
    template_files = {name: template_zip.read(name) for name in template_zip.namelist()}

for week in WEEKS:
    doc_xml = make_document_xml(week).encode('utf-8')
    out_path = os.path.join(OUT_DIR, week['filename'])

    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in template_files.items():
            if name == 'word/document.xml':
                zout.writestr(name, doc_xml)
            else:
                zout.writestr(name, data)

    print(f'생성: {week["filename"]}')

print(f'\n총 {len(WEEKS)}개 파일이 {OUT_DIR} 에 저장되었습니다.')
