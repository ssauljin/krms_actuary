# krms_actuary

『리스크관리연구』 투고 원고: 보험계리사 자격시험의 데이터사이언스 역량 평가 비교 (한국·SOA·CAS·CIA·IFoA)

## 파일

- `krms_actexam_ver4.docx` — 최신 원고 (2차 필수 5과목 체계, SOA·CAS VEE 통합 서술)
- `krms_actexam_ver3.docx` — 학점 인정 비교, 개선 방안 추가본 (저자 수정 반영)
- `krms_actexam_ver2.docx` — 계산기 정책 비교 추가본
- `krms_actexam_ver1.docx` — 초판
- `template/jrm_template.docx` — 학회 Word 템플릿
- `src/manuscript.py` — 템플릿에 본문·표·참고문헌·각주를 채워 넣는 스크립트 (원고 내용은 이 파일에서 수정)
- `build.sh` — 원고 재생성

## 재생성

```bash
./build.sh                 # krms_actexam_build.docx 생성
./build.sh out.docx        # 파일명 지정
```

필요 도구: Python 3, `unzip`, `zip`
