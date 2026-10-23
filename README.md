# 툴 검교정 매니저 — 오프라인 복구판

## 프로그램 다운로드

[실행용 ZIP 다운로드](https://github.com/geonwook-kang/Tool-management-/raw/refs/heads/tool-manager-download/downloads/Tool-Manager.zip)

채팅의 파일 링크 대신 위 GitHub 파일을 받으세요. 압축을 풀고 `Tool-Manager.html`을 Chrome 또는 Edge에서 열면 됩니다. Python 설치나 API 키가 필요하지 않습니다. 인식 엔진과 한글·영문 데이터가 포함되어 오프라인으로 실행됩니다.

직접 다운로드가 열리지 않으면 [GitHub 파일 페이지](https://github.com/geonwook-kang/Tool-management-/blob/tool-manager-download/downloads/Tool-Manager.zip)에서 **Download raw file**을 누르세요.


회사 엑셀과 사진을 선택하면 이름 또는 시리얼 번호로 툴 행을 찾고, 교정일·다음 교정일과 라벨 사진을 자동 반영한 수정본을 만듭니다. 불확실한 사진만 수동 보정합니다. 원본 파일은 덮어쓰지 않습니다.

## 실행

배포한 Tool-Manager.html을 PC에 저장해 Chrome 또는 Edge 일반 창에서 여세요. 인식 엔진·WASM·한글/영문 모델이 HTML에 포함되어 인터넷·Python·API 키 없이 동작합니다. 채팅 미리보기는 자체 보안 정책으로 인식을 제한할 수 있습니다. 상단의 프로그램 HTML/ZIP 저장을 이용해 일반 브라우저에서 실행하세요.

채팅 다운로드와 미리보기 다운로드가 모두 막히면 다운로드가 막히면 → 전체 코드 복사 → 메모장 붙여넣기 → Tool-Manager.html, 파일 형식 모든 파일, UTF-8로 저장합니다. 배포되는 코드는 실행 전의 빈 프로그램이며 업로드한 엑셀·사진·입력값을 포함하지 않습니다.

1. 회사 엑셀(.xlsx)을 선택합니다.
2. 툴 사진(JPG/PNG/WebP)을 한 장 또는 여러 장 선택합니다.
3. 자동 식별된 항목은 별도 확인 체크 없이 해당 행의 날짜와 사진을 반영합니다.
4. 자동 생성된 엑셀을 저장합니다. 확인 필요 사진은 제외되며 별도 보정할 수 있습니다.

인식 중지 버튼으로 엔진을 종료하고 즉시 다시 시작할 수 있습니다. 엔진 초기화45초, 사진 변환15초, OCR단계60초/45초 제한 및 진행 상태를 표시합니다. 엑셀 생성 실패 시 엑셀 다시 생성·저장 버튼으로 인식을 반복하지 않고 재시도합니다.

고유 시리얼 또는 툴이 하나인 담당자 이름으로 연결합니다. 의미가 명확한 교정 날짜 라벨이나 반복 OCR에서 일치한 날짜를 이용합니다. 명확한 교정일만 있고 차기일 표기가 없으면 유효한 엑셀 주기로 계산합니다. 중복·충돌·낮은 신뢰도·폐기·과거 날짜 등은 보류합니다. 손글씨의 정확도를 보장하지 않으며 실제 사용 사진을 확인해야 합니다.

## 복구한 문제

이전 환경에 소스와 실행 파일이 남아 있지 않아 복구했습니다. 이전 버전은 OCR 실행 중 CDN에 의존해 차단·인증서 문제에 영향을 받았습니다. 모든 런타임 자산을 내장하고 Tesseract5.1.1의 언어 객체 초기화 오류를 보정했습니다. IndexedDB 캐시 의존을 제거했습니다. 업스트림 라이브러리 파일은 원본 그대로이며 체크섬을 검증합니다.

## 개발과 검증

```bash
python scripts/prepare_assets.py
python web/build.py
python web/package.py
python tests/browser_test.py --program /workspace/outputs/Tool-Manager-Fixed.html
python tests/automation_test.py --program /workspace/outputs/Tool-Manager-Fixed.html
python tests/offline_ocr_test.py --program /workspace/outputs/Tool-Manager-Fixed.html
python tests/download_test.py --program /workspace/outputs/Tool-Manager-Fixed.html
```

이미 내려받은 자산은 `python scripts/prepare_assets.py --check`로 오프라인 검증합니다. 고정 npm 버전과 SHA512 아카이브 검증을 유지합니다. 빌드/자산 다운로드 외에 브라우저 실행 중 외부 통신은 없습니다.

테스트용 Python은 Playwright/openpyxl/Pillow와 Chromium을 사용합니다. 실제 회사 엑셀의30툴,352셀/스타일,59사진 보존과 원본 SHA256을 검사합니다. 자동 처리 정책 검사는 통제된 OCR출력으로 수행하고 실제 OCR은 별도 테스트에서 HTTP(S)를 전부 차단한 상태로 생성한 한글 인쇄 라벨을 인식합니다. 처음 채팅에 보인 손글씨 원본 자체의 정확도는 검증하지 않았습니다.

실행 결과·회사 엑셀·사진은 Git과 배포물에서 제외합니다. 새로운 환경 복구용 설치 스크립트는 보존된 소스 압축본을 버전별 /workspace/.tool-manager-recovery/ 아래 복원하고 자산 검증 후 배포 파일을 다시 만듭니다. 사용자가 변경한 기존 체크아웃을 덮어쓰지 않습니다.
