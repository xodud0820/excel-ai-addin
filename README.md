# AI 셀 어시스턴트 (Excel 웹용 Office 추가 기능)

## 1. 호스팅 (HTTPS 필수)
`taskpane.html`, `taskpane.js` 를 HTTPS 정적 호스팅(GitHub Pages / Netlify / Vercel / Cloudflare Pages)에 올립니다.
`manifest.xml` 안의 `YOUR_HOST` 를 실제 도메인으로 모두 바꿉니다. (아이콘 URL은 32/64px png 를 올리거나 해당 줄 삭제)

## 2. Excel 웹에 설치(사이드로드)
Excel 웹 → 홈/삽입 → 추가 기능 → 내 추가 기능 → "내 추가 기능 관리" → **내 추가 기능 업로드** → manifest.xml 선택.

## 3. 사용
작업창 설정에서 제공자/Endpoint/모델/API Key 입력 → 범위 선택 → 자연어 지시 → 실행(미리보기) → 적용 / 되돌리기.
- Timely: OpenAI 호환 chat/completions Endpoint 와 키 입력 (CORS 허용 필요)
- Gemini: 키와 모델(gemini-2.0-flash 등)만 입력

## Gemini API 키 발급 (유료 구독과 별개)
Gemini Advanced / Google One AI Premium 구독은 API 사용권이 아닙니다.
https://aistudio.google.com/apikey 에서 키 발급 → 결제 계정(Cloud Billing) 연결 시 유료 한도로 사용(종량제).

## 이미지 기능
작업창의 이미지 목록에서 체크 → 지시 예:
- "체크한 이미지를 모두 가로 150pt로 맞춰줘"
- "첫 이미지를 B2:D6 칸에 비율 유지로 맞춰줘"
- "이미지들을 세로로 20pt 간격으로 A열에 나열해줘"
- "이미지를 셀과 함께 이동/크기변경되게 설정해줘"
