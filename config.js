/* ============================================================
   SJ SCIENCE 허브 — 설정 파일
   선생님이 주로 고칠 곳은 이 파일 하나입니다.
   (index.html 은 화면을 그리는 코드라서 건드리지 않아도 됩니다)
   ============================================================ */

/* ※ 연결한 뒤에는 점검 단원의 내용(개념·OX·뉴스)을 구글 시트의 '단원'·'OX'·'뉴스'·'헷갈린개념' 탭에서 관리해요.
   아래 SUBJECTS 에는 과목과, 시트에 없는 링크용 단원만 적어 두면 됩니다. */

/* ① Apps Script 웹 앱 주소
   - 비워 두면 "체험 모드"로 동작합니다. (학번 10101 / 비밀번호 1234 로 로그인, 기록은 이 브라우저에만 저장)
   - Apps Script 를 배포한 뒤 주소(…/exec)를 따옴표 안에 붙여 넣으면 구글 시트와 연결됩니다. */
const API_URL = '';

/* ①-2 선생님 대시보드에서 바로 열 구글 시트 주소 (비워 두면 버튼이 안 보여요) */
const SHEET_URL = '';

/* ② 사이트 기본 정보 */
const SITE = {
  title: 'SJ SCIENCE',
  subtitle: '단원마다 개념 정리 → 세상과 연결 → OX 확인, 5분 점검으로 내 학습을 기록해요.',
  teacher: '이상준',
  email: 'cepheid@kakao.com',
  youtube: 'https://www.youtube.com/@science_earth',
};

/* ③ 자주 쓰는 링크 (단원 활동에서 이름으로 불러 씁니다) */
const LINKS = {
  lab: 'https://earth516.github.io/lab-reports/',
  earthBank: 'https://earth516.github.io/Earthscience_test/',
  intBank: 'https://earth516.github.io/Integrated_test/',
  notes: 'https://youtube.com/playlist?list=PLctm3ytw_zEv1cqw8Faruop34tb-dSJ7h',
  earthPast: 'https://www.youtube.com/playlist?list=PLO73Sk0gcfGY',
  materials: 'https://sites.google.com/view/sjearth/%EC%9E%90%EB%A3%8C%EC%8B%A4',
};
const LAB = path => LINKS.lab + 'reports/' + path + '/index.html';

/* ④ 과목 → 대단원 → 단원 → 활동(steps)
   활동 type: video(개념 영상) · lab(탐구·실험) · quiz(문제) · note(정리·자료)
   - 활동 하나를 추가하려면 { type, title, url } 한 줄만 넣으면 됩니다.
   - 단원 code 는 문제은행의 소단원 코드와 같게 맞춰 두었습니다.
   - 수업 페이지가 있는 단원은 steps 대신 lesson: '파일 이름' 을 적습니다. (lessons 폴더의 .js 파일)
   ※ 샘플이라 일부 단원만 채워 두었어요. */
const LESSONS = {};   // lessons 폴더의 수업 파일들이 여기에 채워져요
const SUBJECTS = [
  {
    id: 'earth', title: '지구과학', color: '#6fb3e6', colorLight: '#1d6fae', bank: LINKS.earthBank, icon: '🌏',
    desc: '대기와 해양 · 지구의 역사 · 우주',
    chapters: [
      { title: '1. 대기와 해양의 상호작용', units: [
        { code: '1-2-1', title: '대기 대순환과 표층 해류', lesson: 'earth-1-2-1' },
        { code: '1-1-6', title: '수온 염분도', steps: [
          { type: 'video', title: '개념노트 영상', url: LINKS.notes },
          { type: 'lab',   title: 'Argo 자료로 해수 비교하기', url: LAB('earth-science/atmosphere-ocean/argo-seawater-comparison') },
          { type: 'quiz',  title: '문제은행 · 수온 염분도', url: LINKS.earthBank },
        ]},
        { code: '1-2-2', title: '심층 순환', steps: [
          { type: 'video', title: '개념노트 영상', url: LINKS.notes },
          { type: 'lab',   title: '심층 순환의 발생 원리 추론하기', url: LAB('earth-science/atmosphere-ocean/deep-circulation') },
          { type: 'quiz',  title: '문제은행 · 심층 순환', url: LINKS.earthBank },
          { type: 'note',  title: '수업자료실', url: LINKS.materials },
        ]},
        { code: '1-3-2', title: '고기압과 저기압', steps: [
          { type: 'video', title: '개념노트 영상', url: LINKS.notes },
          { type: 'lab',   title: '구름 생성 실험', url: LAB('earth-science/atmosphere-ocean/cloud-formation') },
          { type: 'quiz',  title: '문제은행 · 고기압과 저기압', url: LINKS.earthBank },
        ]},
      ]},
      { title: '2. 지구의 역사와 한반도의 암석', units: [
        { code: '2-1-1', title: '지사학의 법칙', steps: [
          { type: 'video', title: '개념노트 영상', url: LINKS.notes },
          { type: 'lab',   title: '지질 단면도 시뮬레이션', url: LAB('earth-science/geology/geo-section-lab') },
          { type: 'quiz',  title: '문제은행 · 지사학의 법칙', url: LINKS.earthBank },
        ]},
        { code: '2-2-1', title: '퇴적암', steps: [
          { type: 'lab',   title: '퇴적암 관찰하기', url: LAB('earth-science/geology/sedimentary-rocks') },
          { type: 'quiz',  title: '문제은행 · 퇴적암', url: LINKS.earthBank },
        ]},
        { code: '2-2-2', title: '퇴적 구조', steps: [
          { type: 'lab',   title: '퇴적 구조 실험실', url: LAB('earth-science/geology/sedimentary-structures') },
          { type: 'quiz',  title: '문제은행 · 퇴적 구조', url: LINKS.earthBank },
        ]},
        { code: '2-3-3', title: '지질시대의 특징', steps: [
          { type: 'lab',   title: '지질시대 화석 박람회', url: LAB('earth-science/geology/geo-eras-expo') },
          { type: 'quiz',  title: '문제은행 · 지질시대', url: LINKS.earthBank },
        ]},
        { code: '2-4-2', title: '화성암', steps: [
          { type: 'lab',   title: '화성암 관찰하기', url: LAB('earth-science/geology/igneous-rocks') },
          { type: 'quiz',  title: '문제은행 · 화성암', url: LINKS.earthBank },
        ]},
      ]},
      { title: '3. 태양계 천체와 별과 우주의 진화', units: [
        { code: '3-4', title: 'H-R도', steps: [
          { type: 'video', title: '개념노트 영상', url: LINKS.notes },
          { type: 'quiz',  title: '문제은행 · H-R도', url: LINKS.earthBank },
        ]},
      ]},
    ],
  },
  {
    id: 'int', title: '통합과학', color: '#7dd3c8', colorLight: '#0e8a7c', bank: LINKS.intBank, icon: '🔬',
    desc: '통합과학1 · 통합과학2',
    chapters: [
      { title: '통합과학2 · I. 변화와 다양성', units: [
        { code: 's2-1-1-1', title: '지질 시대의 환경과 생물 변화', steps: [
          { type: 'lab',   title: '지질시대 화석 박람회', url: LAB('integrated-science/earth-science/geo-eras-expo') },
          { type: 'quiz',  title: '문제은행 · 통합과학2', url: LINKS.intBank },
        ]},
        { code: 's2-1-2-1', title: '산화와 환원', steps: [
          { type: 'lab',   title: '구리판 가열 실험', url: LAB('integrated-science/chemistry/copper-oxidation') },
          { type: 'lab',   title: '아연판과 황산 구리(Ⅱ) 수용액의 반응', url: LAB('integrated-science/chemistry/zinc-copper-reaction') },
          { type: 'quiz',  title: '문제은행 · 산화와 환원', url: LINKS.intBank },
        ]},
        { code: 's2-1-2-2', title: '산과 염기', steps: [
          { type: 'lab',   title: '산성과 염기성을 나타내는 이온의 확인', url: LAB('integrated-science/chemistry/ion-movement') },
          { type: 'lab',   title: '중화 반응 - 이온 모형 실험', url: LAB('integrated-science/chemistry/neutralization-reaction') },
          { type: 'quiz',  title: '문제은행 · 산과 염기', url: LINKS.intBank },
        ]},
      ]},
      { title: '통합과학2 · II. 환경과 에너지', units: [
        { code: 's2-2-2-2', title: '전기 에너지의 생산', steps: [
          { type: 'lab',   title: '자석과 코일의 전자기 유도', url: LAB('integrated-science/physics/electromagnetic-induction') },
          { type: 'quiz',  title: '문제은행 · 전기 에너지', url: LINKS.intBank },
        ]},
      ]},
    ],
  },
];
