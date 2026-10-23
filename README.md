# 툴 검교정 매니저 — OneDrive 연결판 4.1

## 사진 인식 엔진 로딩 수정

사용자 PC에서 `importScripts`가 `blob:null/...#core.js`를 불러오지 못해 OCR이 시작되지 않았습니다. 4.1은 코어와 워커를 하나의 스크립트에 포함해 이 추가 로딩을 제거합니다. 엔진 초기화 실패는 읽기 어려운 글씨와 구분하고, 이름·시리얼 입력란 대신 재시도를 표시합니다. 직접 입력은 선택한 경우에만 제공합니다.

`opaque_origin_ocr_test.py`는 브라우저 관리 정책을 바꾸지 않고, 테스트 문서에만 추가 스크립트 로딩을 금지하는 CSP를 넣습니다. 이 조건에서 구버전의 동일한 `blob:null/...#core.js` 오류를 재현했고 수정판의 실제 OCR→엑셀 자동 반영을 검증했습니다. 일반 opaque-origin iframe만으로는 구버전 오류가 재현되지 않았으며, Windows 로컬 실행 자체나 원본 손글씨 정확도를 검증한 결과는 아닙니다. 실제 OCR 테스트는 생성한 인쇄 라벨을 사용합니다.

## 사진만 추가해 같은 엑셀에 자동 저장

회사 SharePoint 문서를 PC의 OneDrive 폴더에 동기화하고, 프로그램의 **OneDrive 엑셀 한 번 연결**로 해당 .xlsx를 선택합니다. 그 뒤에는 사진만 추가하면 식별된 툴 행의 날짜·사진을 같은 파일에 저장합니다. 새 세션에서 브라우저 권한 재확인이 필요할 수 있습니다.

이 기능은 File System Access API로 사용자가 선택한 로컬 파일에 저장하고, PC의 OneDrive가 서버로 동기화하는 방식입니다. SharePoint 공유 URL에 직접 로그인하거나 PUT하는 기능이 아닙니다. 프로그램은 선택한 파일이 올바른 동기화 파일인지, OneDrive 서버 동기화가 끝났는지 확인할 수 없습니다. 회사 링크는 사용자가 입력하면 그 브라우저에만 보관하며 배포 파일·Git에 포함하지 않습니다.

- 선택한 파일의 핸들만 연결 정보로 기억합니다. 자동 저장 전 백업 1개는 IndexedDB에 보관하고, 완료된 트랜잭션 뒤에 원본을 수정합니다.
- 사진 작업 시작 시 최신 파일을 읽습니다. 저장 전후 SHA256을 검사하고 변경이 발견되면 원본을 덮어쓰지 않습니다. 저장 오류나 충돌 시 수정본 다운로드와 명시적 재처리를 제공합니다.
- 저장이 완료된 사진은 다음 작업에 다시 적용하지 않습니다. 다른 사람이 바꾼 미선택 셀을 보존합니다.
- 이 검사는 로컬 변경을 보호하며 SharePoint의 분산 잠금·원격 ETag를 구현하지 않습니다. Excel 잠금이나 OneDrive 충돌이 발생하면 프로그램과 OneDrive 메시지를 확인해야 합니다.
- 직접 .xlsx 파일을 업로드하면 연결 모드를 해제하고 기존의 수정본 다운로드 방식으로 사용할 수 있습니다.

이 기능 검사: `python tests/connected_file_test.py --program /workspace/outputs/Tool-Manager-Fixed.html`. 실브라우저의 OPFS 파일 핸들·IndexedDB·SHA256·쓰기/취소·백업·연결 복원을 사용하며, OCR 입력은 통제된 대역입니다. 실제 회사 SharePoint 로그인/서버 동기화는 이 환경에서 접근이 차단되어 검증하지 않았습니다.


## 프로그램 다운로드

[실행용 ZIP 다운로드 — 인식 엔진 수정 4.1](https://github.com/geonwook-kang/Tool-management-/raw/refs/heads/tool-manager-download/downloads/Tool-Manager-Auto-Fix-4.1.zip)

채팅의 파일 링크 대신 위 GitHub 파일을 받으세요. 압축을 풀고 `Tool-Manager.html`을 Chrome 또는 Edge에서 열면 됩니다. Python 설치나 API 키가 필요하지 않습니다. 인식 엔진과 한글·영문 데이터가 포함되어 오프라인으로 실행됩니다.

직접 다운로드가 열리지 않으면 [GitHub 파일 페이지](https://github.com/geonwook-kang/Tool-management-/blob/tool-manager-download/downloads/Tool-Manager-Auto-Fix-4.1.zip)에서 **Download raw file**을 누르세요.


회사 엑셀과 사진을 선택하면 이름 또는 시리얼 번호로 툴 행을 찾고, 교정일·다음 교정일과 라벨 사진을 자동 반영한 수정본을 만듭니다. 불확실한 사진만 수동 보정합니다. 연결 모드는 선택한 파일을 자동 갱신하며, 미연결 모드는 원본을 보존하고 수정본을 만듭니다.

## 실행

배포한 Tool-Manager.html을 PC에 저장해 Chrome 또는 Edge 일반 창에서 여세요. 인식 엔진·WASM·한글/영문 모델이 HTML에 포함되어 인터넷·Python·API 키 없이 동작합니다. 채팅 미리보기는 자체 보안 정책으로 인식을 제한할 수 있습니다. 상단의 프로그램 HTML/ZIP 저장을 이용해 일반 브라우저에서 실행하세요.

채팅 다운로드와 미리보기 다운로드가 모두 막히면 다운로드가 막히면 → 전체 코드 복사 → 메모장 붙여넣기 → Tool-Manager.html, 파일 형식 모든 파일, UTF-8로 저장합니다. 배포되는 코드는 실행 전의 빈 프로그램이며 업로드한 엑셀·사진·입력값·회사 문서 링크·파일 핸들을 포함하지 않습니다.

연결 모드에서는 처음 한 번 동기화된 엑셀을 선택하고 이후에는 사진만 선택합니다. 자동 식별된 항목은 같은 파일에 저장합니다. 연결 없이 직접 파일 선택을 이용하면 예전과 같이 수정본을 다운로드합니다. 확인 필요 사진은 제외되며 별도 보정할 수 있습니다.

인식 중지 버튼으로 엔진을 종료하고 즉시 다시 시작할 수 있습니다. 엔진 초기화45초, 사진 변환15초, OCR단계60초/45초 제한 및 진행 상태를 표시합니다. 엑셀 생성 실패 시 엑셀 다시 생성·저장 버튼으로 인식을 반복하지 않고 재시도합니다.

고유 시리얼 또는 툴이 하나인 담당자 이름으로 연결합니다. 의미가 명확한 교정 날짜 라벨이나 반복 OCR에서 일치한 날짜를 이용합니다. 명확한 교정일만 있고 차기일 표기가 없으면 유효한 엑셀 주기로 계산합니다. 중복·충돌·낮은 신뢰도·폐기·과거 날짜 등은 보류합니다. 손글씨의 정확도를 보장하지 않으며 실제 사용 사진을 확인해야 합니다.

## 복구한 문제

이전 환경에 소스와 실행 파일이 남아 있지 않아 복구했습니다. 이전 버전은 OCR 실행 중 CDN에 의존해 차단·인증서 문제에 영향을 받았습니다. 모든 런타임 자산을 내장하고 Tesseract5.1.1의 언어 객체 초기화 오류를 보정했습니다. OCR 언어 모델의 IndexedDB 캐시 의존을 제거했습니다. OneDrive 파일 연결과 저장 전 백업은 별도의 IndexedDB 저장소를 사용합니다. 업스트림 라이브러리 파일은 원본 그대로이며 체크섬을 검증합니다.

## 개발과 검증

```bash
python scripts/prepare_assets.py
python web/build.py
python web/package.py
python tests/browser_test.py --program /workspace/outputs/Tool-Manager-Fixed.html
python tests/automation_test.py --program /workspace/outputs/Tool-Manager-Fixed.html
python tests/offline_ocr_test.py --program /workspace/outputs/Tool-Manager-Fixed.html
python tests/download_test.py --program /workspace/outputs/Tool-Manager-Fixed.html
python tests/connected_file_test.py --program /workspace/outputs/Tool-Manager-Fixed.html
python tests/opaque_origin_ocr_test.py --program /workspace/outputs/Tool-Manager-Fixed.html
python tests/engine_failure_ui_test.py --program /workspace/outputs/Tool-Manager-Fixed.html
```

이미 내려받은 자산은 `python scripts/prepare_assets.py --check`로 오프라인 검증합니다. 고정 npm 버전과 SHA512 아카이브 검증을 유지합니다. 빌드/자산 다운로드 외에 브라우저 실행 중 외부 통신은 없습니다.

테스트용 Python은 Playwright/openpyxl/Pillow와 Chromium을 사용합니다. 실제 회사 엑셀의30툴,352셀/스타일,59사진 보존과 원본 SHA256을 검사합니다. 자동 처리 정책 검사는 통제된 OCR출력으로 수행하고 실제 OCR은 별도 테스트에서 HTTP(S)를 전부 차단한 상태로 생성한 한글 인쇄 라벨을 인식합니다. 처음 채팅에 보인 손글씨 원본 자체의 정확도는 검증하지 않았습니다.

실행 결과·회사 엑셀·사진은 Git과 배포물에서 제외합니다. 새로운 환경 복구용 설치 스크립트는 보존된 소스 압축본을 버전별 /workspace/.tool-manager-recovery/ 아래 복원하고 자산 검증 후 배포 파일을 다시 만듭니다. 사용자가 변경한 기존 체크아웃을 덮어쓰지 않습니다.
