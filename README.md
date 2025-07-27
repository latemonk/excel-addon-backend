# Excel Add-in Backend - API Key Protection

이 백엔드는 Excel 애드인에서 OpenAI API 키를 안전하게 사용하기 위한 프록시 서버입니다.

## 배포 옵션

### 옵션 1: Azure Functions (권장)

1. Azure 계정 생성 (무료 크레딧 제공)
2. Azure Functions 앱 생성
3. 환경 변수에 OpenAI API 키 설정
4. 이 폴더의 코드 배포

### 옵션 2: Vercel (무료)

1. Vercel 계정 생성
2. 프로젝트 생성 및 연결
3. 환경 변수 설정
4. 자동 배포

### 옵션 3: AWS Lambda

1. AWS 계정 생성
2. Lambda 함수 생성
3. API Gateway 설정
4. 환경 변수 설정

## 환경 변수

- `OPENAI_API_KEY`: OpenAI API 키
- `ALLOWED_ORIGINS`: 허용된 도메인 목록 (콤마로 구분)

## 보안 설정

- CORS 설정으로 특정 도메인만 허용
- Rate limiting 구현 권장
- API 키는 환경 변수에만 저장