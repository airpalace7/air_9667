# HL9667 정비 현황 웹페이지

## 폴더 구성
```
hl9667-site/
├── index.html      # 목차 (01~04 이동)
├── insp.html       # 01 메인 정비표
├── trp.html        # 02 기체 부품 현황
├── eng1.html       # 03 #1 엔진 부품 현황
├── eng2.html       # 04 #2 엔진 부품 현황
├── style.css       # 공용 스타일
├── component.js    # 부품 현황 페이지(02~04) 공용 렌더링 스크립트
├── convert.py       # 엑셀 → JSON 변환 스크립트
└── data/
    ├── insp.json
    ├── trp.json
    ├── eng1.json
    └── eng2.json
```

## GitHub Pages에 올리는 방법 (최초 1회)
1. github.com 가입 → 새 저장소(Repository) 생성 (예: `hl9667`)
2. 이 폴더 안의 모든 파일/폴더를 저장소에 업로드
   - 저장소 페이지 → "Add file" → "Upload files" → 전체 드래그 앤 드롭
3. 저장소 Settings → Pages → Branch를 `main`(또는 `master`)으로 설정 후 저장
4. 1~2분 뒤 `https://내계정.github.io/hl9667/` 주소로 접속 가능

## 비행시간 갱신할 때마다 하는 작업
1. 원본 엑셀 파일(`9667_테스트.xlsx`)에서 비행시간 등 값을 수정하고 저장
2. 터미널(또는 파이썬 실행 환경)에서:
   ```
   python3 convert.py 9667_테스트.xlsx
   ```
   → `data/insp.json`, `data/trp.json`, `data/eng1.json`, `data/eng2.json` 4개가 새로 만들어짐
3. GitHub 저장소의 `data` 폴더에서 이 4개 파일을 새 파일로 교체 업로드
   (기존 파일 클릭 → 연필 아이콘 → 내용 전체 교체, 또는 Upload files로 덮어쓰기)
4. 1분 내로 웹페이지에 반영됨

## 참고
- `python3 convert.py` 실행에는 `openpyxl` 패키지가 필요합니다 (`pip install openpyxl`).
- "초과" 표시(빨간 줄, 배지)는 잔여시간이 음수이거나 다음 예정일이 오늘보다 지난 경우 자동으로 붙습니다.
- 저장소를 Public으로 두면 URL을 아는 누구나 볼 수 있고, 수정은 저장소 협업자(Collaborator)로 등록된 사람만 할 수 있습니다.
