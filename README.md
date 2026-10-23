# 툴 검교정 매니저 — 툴 라벨·SMAX 화면 자동 반영 4.6

## 4.6 금속 배경과 곡면 교정표 재인식

전체 사진의 밝기만으로 글자와 배경을 나누면 밝은 바닥 앞의 회색 라벨이 글자와 뭉칠 수 있습니다. 이름·시리얼을 찾지 못한 경우 사진 각 부분의 밝기를 따로 비교하는 Sauvola 인식을 추가합니다. 같은 사진의 서로 다른 고신뢰 번호는 계속 충돌로 처리하며, 다음 사진과 일반 인식에는 기본 밝기 설정을 다시 적용합니다. 넓은 영역에서 낮은 신뢰도로 발견한 숫자도 위치 단서로만 사용해 원본 크기의 작은 영역을 다시 읽습니다. 자동 반영 신뢰도 기준은 낮추지 않습니다.

둥근 토크렌치의 교정표에서는 `2026.0`과 `6.22`처럼 날짜가 갈라지거나 표 테두리가 숫자 행을 가를 수 있습니다. 읽힌 연도 조각으로 원본 영역을 찾고, 사진에 실제로 보이는 표 선의 굽음을 측정해 날짜 칸을 펼칩니다. 해당 칸의 밝기로 글자와 종이를 나눈 뒤 서로 다른 OCR 분할 방식으로 날짜 전체를 다시 확인합니다. 엑셀의 기존 날짜나 예시 날짜로 읽히지 않은 일자를 채우지 않습니다. 원본 업로드 사진과 엑셀에 넣는 사진은 이 인식 전처리로 변경하지 않습니다.

사용자가 제공한 진단 ZIP의 `original.jpg`를 그대로 읽는 실제 OCR 검사에서 전체 시리얼을 인식하고, 등록된 담당자와 올바른 행을 자동 선택했습니다. 사진에 표기된 교정일과 차기일은 서로 다른 두 날짜 판독에서 일치했습니다. 사람이 이름·시리얼·날짜를 입력하거나 OCR 응답을 교체하지 않고 화면 값과 자동 반영까지 확인했습니다.

브라우저 테스트 저장소(OPFS)의 엑셀 복사본에 `Cal Last`, `Cal Next`, `Label Photo`가 한 번의 자동 저장으로 반영되는 것도 검사했습니다. 352개 셀과 스타일, 59장 사진을 비교해 대상 날짜·사진 이외의 내용과 SMAX 사진, 저장 전 백업이 보존됐습니다. 실제 회사 OneDrive 서버의 동기화 완료를 검사한 결과는 아닙니다.

사진에서 고유 시리얼이 명확하면 그 번호에 등록된 한글 담당자를 연결할 수 있습니다. 담당자 이름 자체를 높은 신뢰도로 읽었다는 뜻은 아닙니다. 흐림·가림으로 실제 정보가 사라진 사진까지 자동 인식을 보장하지 않습니다.

## 툴 라벨과 SMAX 화면 함께 업로드

같은 사진 선택에서 툴 교정 라벨과 SMAX 화면을 올리면 문서 종류를 판별해 각각 `Label Photo`, `SMAX Tool Photo` 칸에 반영합니다. SMAX의 **Serial/Asset Number**는 엑셀의 `Serial`과 대조하며, `Asset #`의 지역·담당자 문자열이나 SMAX Tool Code·SGPN 번호를 시리얼로 사용하지 않습니다. 영문 `Tool Owner`는 원문으로 읽어 표시하고, 정확한 등록 시리얼에 연결된 엑셀의 한글 담당자를 사용합니다. 일반 영문 이름의 자동 번역 기능으로 주장하지 않습니다.

SMAX의 **Calibration Due Date → Cal Next**만 날짜 변경에 사용합니다. `Created By`, `Last Modified` 날짜는 `Cal Last`로 사용하지 않으며 SMAX 처리는 `Cal Last`를 변경하지 않습니다. 차기일이 불명확하면 SMAX 사진만 저장하고 기존 날짜를 유지합니다. 서로 다른 명확한 시리얼·차기일, 폐기·중복 항목·과거 차기일은 보류합니다.

같은 툴의 라벨 1장과 SMAX 1장은 한 번에 처리할 수 있습니다. 같은 사진 칸에 두 장이거나 두 화면의 확정 차기일이 다르면 자동 반영하지 않습니다. SMAX 열이 없는 엑셀을 라벨 사진 칸에 대신 저장하지 않습니다. 수동 검토에서는 저장 위치와 SMAX 차기일 반영 여부를 선택할 수 있습니다.

## 라벨 후보와 날짜 영역 재탐색

밝은 영역을 찾는 과정에서 큰 한글 내부의 작은 흰 공간을 라벨로 선택하고, 후보가 있다는 이유로 넓은 탐색을 생략하던 문제를 수정합니다. 검출한 라벨과 별도로 상단을 덮는 재탐색 영역을 유지합니다. 연도 일부나 Calibration Label 제목만 찾았을 때 좁은 영역을 반복해서 읽던 날짜 경로도 보완하며, 독립된 넓은 영역에서 다시 찾습니다.

정확한 담당자 이름과 완전한 두 교정 날짜가 읽히면, 동일 담당자에게 등록된 툴에서 실제로 읽힌 4자리 숫자 조각도 대조합니다. 서로 다른 인식에서 높은 신뢰도로 두 번 확인되고, 등록 번호 길이의 40% 이상이며, 동일 담당자의 후보가 하나이고 엑셀 주기까지 맞아야 `owner-partial-serial`로 저장합니다. 같은 번호 조각을 공유하는 툴은 주기로 억지 선택하지 않습니다. `0713`을 `0716`으로 치환하거나 누락된 날짜의 일을 추측하지 않습니다. 기존의 독립적인 이름·주기 확정은 유지합니다.

문제가 남으면 선택한 사진 아래 **인식 진단 ZIP 저장**으로 원본 사진과 패스별 글자·좌표·신뢰도를 한 파일로 받을 수 있습니다. 회사 엑셀, 회사 링크, 다른 사진, 브라우저 파일 핸들은 포함하지 않습니다. 진단 ZIP은 사용자가 다운로드하며 자동 전송하지 않습니다.

## 떨어진 이름·시리얼 라벨 재인식

날짜는 읽었지만 이름·시리얼을 찾지 못하면 상단의 라벨 후보를 나누어 재인식합니다. 기존의 한 번짜리 상단 영역·하단 영역 재독으로 끝나던 경로를 보완합니다. 같은 물리적 행에서 단어가 줄바꿈으로 분리된 경우에는 좌표와 실제 글자를 대조해 정확한 조합을 찾습니다. `Serial: 3155 0075`를 전체 번호와 짧은 접두 번호로 동시에 관측하던 문제도 수정합니다.

앞자리 0과 실제 OCR 문자를 유지하며, 끝자리나 이름을 엑셀 값에 맞춰 임의 변경하지 않습니다. 원본의 고신뢰 시리얼 충돌은 계속 보류합니다.

## 일부 시리얼만 보이는 새 사진

새 촬영 사진에서 이름이나 번호 일부가 흐려졌을 때에도 실제로 읽힌 연속된 시리얼 조각을 엑셀 전체 툴 목록과 대조합니다. 최소 6글자이면서 전체 시리얼의 60% 이상인 조각이 서로 다른 재인식에서 두 번 이상 높은 신뢰도로 읽히고, 등록된 툴 한 개만 일치하며, 사진의 명확한 두 날짜가 그 툴의 엑셀 교정 주기와 맞으면 `partial-serial`로 자동 연결합니다. 이름까지 정확하면 최소 5글자·50%까지 대조합니다. 읽히지 않은 부분은 엑셀의 등록 번호를 사용하며 사진에서 전체 번호를 읽었다고 판단하지 않습니다.

폐기 항목을 포함한 목록에서 후보가 여러 개면 자동 연결하지 않습니다. 다른 명확한 이름·전체 시리얼, 서로 다른 번호 조각, 날짜와 주기의 충돌도 보류합니다. 완전히 읽힌 다른 번호의 끝자리를 잘라 비슷한 번호에 맞추지 않습니다. 한 번만 읽힌 조각과 낮은 신뢰도의 반복은 충분한 근거로 사용하지 않습니다. 부분 번호 자체를 전체 번호의 새 별칭으로 기억하지 않습니다.

## 시리얼 재인식과 교정 주기 대조

이름이 읽혀도 동일 엔지니어에게 여러 툴이 있으면 시리얼을 계속 읽습니다. 이름 하나만 읽고 재시도를 끝내던 제어 흐름을 수정했습니다. 작은 라벨은 원본 크기로 다시 읽고, 찾기 위한 탐색 OCR과 확정 날짜 판독을 구분합니다. 표의 세로 테두리가 숫자 인식을 억제할 때에는 픽셀에서 긴 가장자리 선을 찾아 그 안쪽 숫자 열을 읽습니다. 글자·날짜 값을 이미지에 추가하거나 고치지 않습니다.

시리얼을 못 읽더라도 이름과 두 교정 날짜가 명확하고, 그 간격이 동일 담당자의 현재 엑셀 주기 한 개에만 일치하면 `name-cycle`로 연결합니다. 6개월·12개월 이외의 주기도 엑셀에서 읽습니다. 동일 주기의 툴이 여러 개이거나 다른 명시적 시리얼과 충돌하면 보류합니다.

## 확인된 인식 패턴 기억

`web/profiles.js`는 독립적으로 확인한 자동 매칭 또는 사용자가 확인한 매칭에서 근접 인식 패턴을 이 브라우저의 localStorage에 기억합니다. 모델 재훈련이나 회사 공용 학습 서버는 구현하지 않습니다. 사진·날짜·회사 링크는 저장하지 않으며, 다른 PC에 자동 공유되지 않습니다.

현재 엑셀의 고유 시리얼·담당자·자산·주기를 재검증하며, 명확히 알려진 다른 툴의 값을 덮어쓰지 않습니다. 앞자리 0을 붙이거나 삭제하는 변형은 기억하지 않습니다. 기억한 별칭을 근거로 다시 새 별칭을 늘리는 방식도 사용하지 않습니다. 충돌한 별칭은 적용하지 않고, 저장소가 막히면 현재 실행에서만 사용합니다. 화면에서 활용을 끄거나 초기화할 수 있습니다.

## 라벨의 날짜를 Cal Last / Cal Next에 저장

사진의 **교정일자 → Excel Cal Last**, **차기교정예정일자 → Excel Cal Next**로 연결합니다. 셀은 열 이름으로 찾으므로 엑셀의 열 순서와 관계없이 같은 엔지니어·시리얼 행에 날짜를 씁니다. 연결 모드에서는 사진도 함께 원본 연결 파일에 자동 저장합니다.

작은 교정 라벨에서 OCR이 숫자를 정확히 읽어도 표의 글자 순서 때문에 역할을 연결하지 못하거나, 고정 영역 자르기로 재인식이 나빠지던 문제를 개선했습니다. 단어 위치로 같은 행의 날짜를 연결하고, 검출한 날짜 영역을 확대해 재인식합니다. 날짜 전용 재인식에서 나온 글자는 이름·시리얼 판단에 섞지 않습니다. 라벨의 짧은 한글 단어를 엔지니어 이름으로 잘못 판단해 고유 시리얼 일치를 막던 문제도 구분합니다. 서로 다른 명시적 이름·시리얼이나 날짜 충돌은 보류합니다.

## 사진 인식 엔진 로딩 수정

사용자 PC에서 `importScripts`가 `blob:null/...#core.js`를 불러오지 못해 OCR이 시작되지 않았습니다. 4.1은 코어와 워커를 하나의 스크립트에 포함해 이 추가 로딩을 제거합니다. 엔진 초기화 실패는 읽기 어려운 글씨와 구분하고, 이름·시리얼 입력란 대신 재시도를 표시합니다. 직접 입력은 선택한 경우에만 제공합니다.

`opaque_origin_ocr_test.py`는 브라우저 관리 정책을 바꾸지 않고, 테스트 문서에만 추가 스크립트 로딩을 금지하는 CSP를 넣습니다. 이 조건에서 구버전의 동일한 `blob:null/...#core.js` 오류를 재현했고 수정판의 실제 OCR→엑셀 자동 반영을 검증했습니다. 일반 opaque-origin iframe만으로는 구버전 오류가 재현되지 않았으며, Windows 로컬 실행 자체나 원본 손글씨 정확도를 검증한 결과는 아닙니다. 실제 OCR 테스트는 생성한 인쇄 라벨을 사용합니다.

## 사진만 추가해 같은 엑셀에 자동 저장

Teams 파일 탭의 **SharePoint에서 열기**로 문서 라이브러리를 열거나 회사 SharePoint로 이동합니다. 문서를 PC의 OneDrive 폴더에 동기화하고, 프로그램의 **OneDrive 엑셀 한 번 연결**로 해당 .xlsx를 선택합니다. 그 뒤에는 사진만 추가하면 식별된 툴 행의 날짜·사진을 같은 파일에 저장합니다. 새 세션에서 브라우저 권한 재확인이 필요할 수 있습니다.

이 기능은 File System Access API로 사용자가 선택한 로컬 파일에 저장하고, PC의 OneDrive가 서버로 동기화하는 방식입니다. SharePoint 공유 URL에 직접 로그인하거나 PUT하는 기능이 아닙니다. 프로그램은 선택한 파일이 올바른 동기화 파일인지, OneDrive 서버 동기화가 끝났는지 확인할 수 없습니다. 회사 링크는 사용자가 입력하면 그 브라우저에만 보관하며 배포 파일·Git에 포함하지 않습니다.

- 선택한 파일의 핸들만 연결 정보로 기억합니다. 자동 저장 전 백업 1개는 IndexedDB에 보관하고, 완료된 트랜잭션 뒤에 원본을 수정합니다.
- 사진 작업 시작 시 최신 파일을 읽습니다. 저장 전후 SHA256을 검사하고 변경이 발견되면 원본을 덮어쓰지 않습니다. 저장 오류나 충돌 시 수정본 다운로드와 명시적 재처리를 제공합니다.
- 저장이 완료된 사진은 다음 작업에 다시 적용하지 않습니다. 다른 사람이 바꾼 미선택 셀을 보존합니다.
- 이 검사는 로컬 변경을 보호하며 SharePoint의 분산 잠금·원격 ETag를 구현하지 않습니다. Excel 잠금이나 OneDrive 충돌이 발생하면 프로그램과 OneDrive 메시지를 확인해야 합니다.
- 직접 .xlsx 파일을 업로드하면 연결 모드를 해제하고 기존의 수정본 다운로드 방식으로 사용할 수 있습니다.

이 기능 검사: `python tests/connected_file_test.py --program /workspace/outputs/Tool-Manager-Fixed.html`. 실브라우저의 OPFS 파일 핸들·IndexedDB·SHA256·쓰기/취소·백업·연결 복원을 사용하며, OCR 입력은 통제된 대역입니다. 실제 회사 SharePoint 로그인/서버 동기화는 이 환경에서 접근이 차단되어 검증하지 않았습니다.


## 프로그램 다운로드

[실행용 ZIP 다운로드 — 토크렌치 인식 개선 4.6](https://github.com/geonwook-kang/Tool-management-/raw/refs/heads/tool-manager-download/downloads/Tool-Manager-Torque-4.6.zip)

채팅의 파일 링크 대신 위 GitHub 파일을 받으세요. 압축을 풀고 `Tool-Manager.html`을 Chrome 또는 Edge에서 열면 됩니다. Python 설치나 API 키가 필요하지 않습니다. 인식 엔진과 한글·영문 데이터가 포함되어 오프라인으로 실행됩니다.

직접 다운로드가 열리지 않으면 [GitHub 파일 페이지](https://github.com/geonwook-kang/Tool-management-/blob/tool-manager-download/downloads/Tool-Manager-Torque-4.6.zip)에서 **Download raw file**을 누르세요.


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
python tests/calibration_dates_test.py --program /workspace/outputs/Tool-Manager-Fixed.html
python tests/calibration_label_ocr_test.py --program /workspace/outputs/Tool-Manager-Fixed.html
python tests/torque_dates_test.py --program /workspace/outputs/Tool-Manager-Fixed.html
python tests/torque_ocr_test.py --program /workspace/outputs/Tool-Manager-Fixed.html
python tests/profile_test.py --program /workspace/outputs/Tool-Manager-Fixed.html
python tests/identity_retry_test.py --program /workspace/outputs/Tool-Manager-Fixed.html
python tests/identity_ocr_test.py --program /workspace/outputs/Tool-Manager-Fixed.html
python tests/partial_identity_test.py --program /workspace/outputs/Tool-Manager-Fixed.html
python tests/partial_identity_ocr_test.py --program /workspace/outputs/Tool-Manager-Fixed.html
python tests/roi_recovery_test.py --program /workspace/outputs/Tool-Manager-Fixed.html
python tests/owner_fragment_test.py --program /workspace/outputs/Tool-Manager-Fixed.html
python tests/diagnostics_test.py --program /workspace/outputs/Tool-Manager-Fixed.html
python tests/smax_test.py --program /workspace/outputs/Tool-Manager-Fixed.html
python tests/smax_workbook_test.py --program /workspace/outputs/Tool-Manager-Fixed.html
python tests/smax_app_test.py --program /workspace/outputs/Tool-Manager-Fixed.html
python tests/smax_ocr_test.py --program /workspace/outputs/Tool-Manager-Fixed.html
```


실제 촬영 파일의 회귀 검사는 아래 예시의 경로와 기대값을 실제 사진에 맞게 바꿔 실행합니다. 기대값은 테스트의 결과 비교에만 사용하며 OCR 입력에 전달하지 않습니다.

```bash
python tests/actual_photo_ocr_test.py \
  --program /workspace/outputs/Tool-Manager-Fixed.html \
  --workbook /absolute/path/company.xlsx \
  --photo /absolute/path/original.jpg \
  --serial 0123456789 --name 홍길동 \
  --last 2027-03-17 --due 2027-09-17 \
  --diagnostics /tmp/tool-actual-photo-ocr.json
```

이 검사는 실제 Tesseract와 앱 자동 판단, OPFS 연결 파일의 날짜·사진·백업, 다른 셀·사진 보존을 확인합니다. `--diagnose-only`는 실패 결과를 보고서로 남기며 성공 주장에 사용하지 않습니다. `--variant mild-blur` 또는 `--variant reduced-resolution`은 원본과 구분한 인위적 열화 실험입니다. 실제 원본과 진단 보고서는 실행 환경의 임시 경로에만 두며 Git과 배포 ZIP에 넣지 않습니다. 실제 OneDrive 서버 동기화는 이 검사 범위에 포함하지 않습니다.

이미 내려받은 자산은 `python scripts/prepare_assets.py --check`로 오프라인 검증합니다. 고정 npm 버전과 SHA512 아카이브 검증을 유지합니다. 빌드/자산 다운로드 외에 브라우저 실행 중 외부 통신은 없습니다.

테스트용 Python은 Playwright/openpyxl/Pillow와 Chromium을 사용합니다. 실제 회사 엑셀의30툴,352셀/스타일,59사진 보존과 원본 SHA256을 검사합니다. 날짜 정책·OPFS 저장 검사는 통제된 OCR출력으로 수행합니다. 실제 OCR은 HTTP(S)를 전부 차단한 상태로 생성한 한글 인쇄 라벨과 원본 엑셀에서 추출한 실제 교정 사진 2장을 읽습니다. 실제 시리얼이 있는 사진은 임시 엑셀 복사본의 과거 날짜를 사진의 날짜로 자동 갱신하고 사진 위치까지 검증합니다. 저해상도 실제 사진은 두 차례 일치한 날짜로 판정했으며 한글 라벨 자체를 정확히 읽었다는 검증은 아닙니다. 날짜만 있는 사진은 담당자를 식별할 수 없어 자동 저장을 보류합니다. 처음 채팅에 보인 손글씨 원본 자체의 정확도는 검증하지 않았습니다.

추가 토크렌치 검사는 원본 엑셀의 실제 라벨 사진과 별도 시스템 화면 사진을 각각 읽어 날짜와 선행 0 시리얼을 검사합니다. 두 사진을 합쳐 한 장짜리 실제 촬영 사진이라고 주장하지 않습니다. 새 촬영 JPG는 별도 `actual_photo_ocr_test.py`에서 원본 파일 자체로 검사합니다. 사용자께서 지적한 이름만 읽고 시리얼 재시도를 생략하는 제어 흐름은 실제 `ToolOCR.recognize`와 대역 워커로 별도 재현·검증합니다. 기억 기능은 실제 브라우저 저장소와 대역 OCR의 앱 통합 검사로 확인하며 외부 모델 훈련 검사는 아닙니다.

`identity_retry_test.py`는 날짜만 읽힌 뒤 상단 태그 재독까지 진행하는 제어 흐름, 같은 행의 분리 문자, 충돌 시 연결 파일 쓰기 차단을 통제된 워커 응답으로 검사합니다. 사용자 원문에 없는 글자를 추가해 매칭하지 않는지도 검사합니다. `identity_ocr_test.py`는 실제 OCR로 기존 원본 이미지의 시리얼과 분리 태그의 합성 배치를 구분해 검사합니다. 합성 배치의 이름 카드는 생성한 인쇄 글자이며, 별도 시스템 화면의 시리얼·라벨 사진 픽셀을 재배치한 실험입니다. 이 실험은 최신 촬영 JPG나 한 장짜리 실제 사진의 자동 저장 성공을 입증하지 않습니다.

`partial_identity_test.py`는 부분 시리얼 정책과 연결 파일 자동 저장·충돌 차단을 통제된 OCR 응답으로 검사합니다. `partial_identity_ocr_test.py`는 담당자 이름과 전체 시리얼이 없는 합성 인쇄 라벨을 실제 Tesseract로 읽습니다. 읽힌 7자리 조각을 독립적으로 반복 확인한 뒤 등록된 전체 번호로 연결하고, 실제 OPFS 파일의 Cal Last·Cal Next·사진과 원본 백업을 검사합니다. 원문에 전체 등록 번호를 추가하거나 OCR 응답을 교체하지 않습니다. 합성 라벨 검사이므로 흐린 실촬영 사진 전반의 정확도를 보장하지는 않습니다.

`smax_test.py`는 SMAX 분류·필드 역할·재인식 제어를 검사하며, `smax_workbook_test.py`와 `smax_app_test.py`는 실제 OPFS 파일에서 두 사진 칸, 혼합 배치, 날짜 보존·충돌과 백업을 확인합니다. 이들의 OCR 응답은 통제된 대역입니다. 별도 `smax_ocr_test.py`는 원본 엑셀의 실제 SMAX 화면 두 장을 실제 OCR로 읽고, 임시 파일의 차기일을 하루 이전으로 둔 뒤 사진의 정확한 차기일과 SMAX 사진으로 자동 갱신합니다. Cal Last·Label Photo·352셀·59사진·저장 전 백업을 검사합니다. 새 채팅 캡처나 임의로 흐린 사진 전반의 정확도를 검증한 결과는 아닙니다.

실행 결과·회사 엑셀·사진은 Git과 배포물에서 제외합니다. 새로운 환경 복구용 설치 스크립트는 보존된 소스 압축본을 버전별 /workspace/.tool-manager-recovery/ 아래 복원하고 자산 검증 후 배포 파일을 다시 만듭니다. 사용자가 변경한 기존 체크아웃을 덮어쓰지 않습니다.
