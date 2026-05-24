# 매교역팰루시드 아실 매물 Slack 자동 리포트

## 구성

- `daily_asil_slack.py`: 아실 매물 수집, 엑셀 저장, 리포트 생성, Slack 발송
- `requirements.txt`: Python 패키지 목록
- `.github/workflows/daily.yml`: 매일 한국시간 06:00 자동 실행
- `data/`: 매일 수집한 엑셀 파일 저장
- `reports/`: 매일 생성한 텍스트 리포트 저장

## GitHub Secret 설정

1. GitHub 저장소로 이동
2. Settings
3. Secrets and variables
4. Actions
5. New repository secret
6. Name: `SLACK_WEBHOOK_URL`
7. Secret: Slack Incoming Webhook URL 입력

## 실행 시간

GitHub Actions cron은 UTC 기준입니다.

- `21:00 UTC` = 한국시간 `06:00`

## 수동 실행

GitHub 저장소 → Actions → Daily Asil Slack Report → Run workflow

## 첫 실행 참고

첫 실행일에는 어제/7일 전 데이터가 없어서 일부 항목이 `데이터 없음`으로 표시됩니다.
