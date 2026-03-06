# DenJOY DentWeb Automation Agent

DentWeb 수술기록을 자동으로 추출하여 DenJOY에 업로드하는 Windows 에이전트입니다.

## 설치

1. [최신 릴리즈](https://github.com/denjoy-korea/dentweb-agent/releases/latest)에서 `dentweb-agent.exe` 다운로드
2. 병원 PC에서 실행
3. DenJOY 앱 설정 화면에서 복사한 에이전트 토큰 입력

## 동작 방식

1. DenJOY 서버에 주기적으로 폴링 (기본 30초)
2. 실행 요청 감지 시 DentWeb 프로그램에서 Excel 자동 다운로드
3. 다운로드된 Excel을 DenJOY 서버에 업로드

## 요구사항

- Windows 10 이상
- DentWeb 프로그램이 설치되어 있어야 함
- 화면 해상도 1920x1080 권장
