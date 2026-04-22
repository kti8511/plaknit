# 상품DATA 운영 대시보드

정적 Vercel 배포용 폴더입니다.

## 포함 파일

- `index.html`: 운영 대시보드 본문
- `vercel.json`: Vercel 정적 배포 설정

## 배포 방법

### 방법 1. Vercel 웹에서 직접 업로드

1. <https://vercel.com/new> 접속
2. `vercel-dashboard` 폴더를 새 프로젝트로 업로드
3. Framework Preset은 `Other` 또는 감지된 기본값 사용
4. Build Command는 비워둠
5. Output Directory는 비워둠
6. Deploy 실행

### 방법 2. GitHub 연동

1. 이 폴더를 GitHub 저장소에 업로드
2. Vercel에서 해당 저장소 Import
3. 배포 설정은 정적 사이트 기본값 사용

## 업데이트 방식

데이터가 바뀌면 다음 순서로 다시 생성합니다.

1. `process_selfmall_data.py` 실행
2. `build_product_dashboard.py` 실행
3. 새로 생성된 `상품DATA_운영대시보드.html`을 `vercel-dashboard/index.html`로 복사
4. Vercel에 재배포

현재 대시보드는 정적 HTML입니다. 웹에서 파일을 업로드해 자동 반영하려면 별도 웹앱/DB 구성이 필요합니다.
