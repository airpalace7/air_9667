# HL9667 정비 현황 웹페이지

## 폴더 구성
```
hl9667-site/
├── index.html      # 목차 (01~04 이동)
├── insp.html       # 01 메인 정비표 (관리자 모드 지원)
├── trp.html        # 02 기체 부품 현황
├── eng1.html       # 03 #1 엔진 부품 현황
├── eng2.html       # 04 #2 엔진 부품 현황
├── style.css       # 공용 스타일
├── component.js    # 부품 현황 페이지(02~04) 공용 렌더링 스크립트
├── convert.py       # 엑셀 → JSON 변환 스크립트 (초기 세팅/엑셀 관리 항목용)
└── data/
    ├── insp.json
    ├── trp.json
    ├── eng1.json
    └── eng2.json
```

## 01 메인정비표 - 관리자 모드로 웹에서 직접 수정하기

메인 정비표(93개 항목 중 86개)는 이제 **웹페이지에서 직접 비행시간/수행일을 입력하면 잔여시간이 자동으로 재계산**되고, 저장 버튼을 누르면 GitHub 저장소에 바로 반영됩니다. 파이썬/엑셀 없이도 가능해요.

### 최초 1회: GitHub Personal Access Token 발급
1. GitHub 로그인 → 오른쪽 위 프로필 → Settings
2. 왼쪽 아래 Developer settings → Personal access tokens → Tokens (classic)
3. Generate new token (classic) 클릭
4. Note에 아무 이름 입력, 권한(Scope)은 **repo** 전체 체크
5. Generate token 클릭 → 나오는 토큰 문자열을 복사해서 안전한 곳에 보관 (다시 못 봄)

### 사용 방법
1. `insp.html` 접속 → 우측 상단 "관리자 모드" 클릭
2. GitHub 계정명(예: airpalace7), 저장소 이름(예: air_9667), 위에서 발급한 토큰 입력 → 연결
3. 상단에 나오는 "기준값 수정" 칸에서 A/C TSN(총 비행시간) 갱신 → 모든 항목의 잔여시간이 즉시 자동으로 줄어듦
4. 특정 정비를 완료했으면 해당 행의 PERFORMED DATE / PERFORMED A/C TIME 칸에 입력 → 그 항목만 재계산됨
5. 화면 하단 "변경사항 저장" 클릭 → GitHub에 자동 커밋, 1분 내 실제 사이트에도 반영

### "엑셀관리" 배지가 붙은 7개 항목은?
다른 항목이나 다른 시트를 참조하는 특수한 계산식이 있는 항목이에요(5Y hydrostatic test, ENG 150HR INSP, CARGO HOOK 등). 이 항목들은 자동 재계산이 안전하지 않아서, 여전히 아래 "엑셀 + 변환스크립트" 방식으로 관리해주세요.

### 토큰 관련 주의
- 토큰은 이 브라우저(로컬)에만 저장돼요. 다른 사람 컴퓨터에서는 다시 연결해야 함
- 토큰이 있으면 저장소에 쓰기 권한이 생기는 것과 같으니, 본인만 아는 곳에 보관하고 공유하지 마세요
- 토큰을 잃어버리거나 노출된 것 같으면 GitHub 설정에서 바로 삭제(Delete) 가능

## 02~04 부품현황 페이지 / "엑셀관리" 항목 업데이트 (기존 방식)
1. 원본 엑셀 파일(`9667_테스트.xlsx`)에서 값을 수정하고 저장
2. `python3 convert.py 9667_테스트.xlsx` 실행 → `data/` 안 JSON 4개 새로 생성
3. GitHub 저장소의 `data` 폴더에서 이 파일들을 새 파일로 교체 업로드

## 참고
- `python3 convert.py` 실행에는 `openpyxl` 패키지가 필요합니다 (`pip install openpyxl`).
- "초과" 표시(빨간 줄, 배지)는 잔여시간이 음수이거나 다음 예정일이 오늘보다 지난 경우 자동으로 붙습니다.
- 저장소를 Public으로 두면 URL을 아는 누구나 볼 수 있고, 실제 저장(수정)은 유효한 GitHub 토큰을 가진 사람만 할 수 있습니다.
