# 테스트 방법

## 0. 지금 단계에서 테스트하는 것 (P0 환경 확인)

대시보드 코드를 만들기 전에, 회사 PC에서 `tools/Test-EDashEnvironment.ps1` 하나로 아래 네 가지를 확인합니다.

1. 이 PC에서 PowerShell 5.1과 WinForms(차트 포함)가 동작하는지
2. 현재 Windows 계정으로 SharePoint REST에 접근되는지. 이 결과로 온프레미스인지 SharePoint Online인지, 대시보드 인증 방식이 정해집니다.
3. 두 리스트(EasDashboard, EASDB2)의 실제 컬럼 형식과 값 형식, 그리고 내 권한
4. SRN 데이터 품질: 중복 행, 따옴표·공백, 두 리스트 간 연결 여부

**읽기 전용입니다.** SharePoint에는 아무것도 쓰지 않고, 결과는 바탕화면에 텍스트 파일로만 저장합니다.

---

## 1. 준비

1. **파일 받기.** GitHub에서 브랜치 `claude/awesome-clarke-xadn7j`의 `tools` 폴더로 갑니다. 아래 두 파일을 **같은 폴더**(예: `C:\EDash\`)에 저장합니다. 각 파일을 열고 "Download raw file" 버튼을 누르면 됩니다.
   - `Test-EDashEnvironment.ps1`
   - `Test-EDashEnvironment.cmd`

   `Code → Download ZIP`으로 받아 압축을 풀어도 됩니다.
   스크립트는 영문(ASCII)으로만 작성했기 때문에, 내용을 복사해 메모장에 붙여넣고 저장해도 인코딩 문제가 없습니다. 이때 확장자가 `.ps1.txt`가 되지 않게 주의하세요.
2. **사내망 또는 VPN에 연결합니다.** 브라우저에서 EasDashboard 리스트가 열리는지 먼저 확인합니다.

---

## 2. 실행

**방법 A (가장 쉬움):** `Test-EDashEnvironment.cmd`를 더블클릭합니다. 점검이 끝나면 "계속하려면 아무 키나 누르십시오..."에서 창이 멈춥니다.

**방법 B (PowerShell 창에서):**

```powershell
cd C:\EDash
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Test-EDashEnvironment.ps1
```

- 관리자 권한은 필요 없습니다. **"다른 사용자로 실행"은 하지 마세요.** 본인 계정으로 접근되는지를 확인하는 테스트입니다.
- PowerShell 7(`pwsh`)이 설치되어 있어도 `powershell.exe`(5.1)로 실행해 주세요. 대시보드가 5.1 기준입니다.
- 보통 10초 안에 끝납니다.

### 옵션

| 옵션 | 예 | 용도 |
|---|---|---|
| `-SiteUrl` | `https://share.philips.com/sites/EASalarmhandling` | 사이트 주소가 다를 때 (기본값이 이 주소) |
| `-ListNames` | `EasDashboard,EASDB2` | 리스트 **URL 이름**이 다를 때 |
| `-MaskFields` | `Title,field_2,field_3,field_4,SRN,SystemReferenceNumber` | 값을 가릴 컬럼 (기본값이 이 목록. 추가하려면 기존 목록까지 함께 적기) |
| `-MaxDistinctValues` | `15` | 고유값이 이 개수 이하인 컬럼만 실제 값을 표시 |
| `-OutFile` | `C:\EDash\report.txt` | 보고서 저장 위치 (기본: 바탕화면) |

`.cmd`에도 같은 옵션을 붙일 수 있습니다. 예: `Test-EDashEnvironment.cmd -ListNames EasDashboard`

---

## 3. 결과 화면 구성

| 구역 | 내용 |
|---|---|
| **1. Local PC** | PowerShell·.NET 버전, 시간대, 실행 정책, WinForms·차트 로드 |
| **2. SharePoint server and sign-in** | DNS, 현재 계정으로 REST 접근, OData 형식, SharePoint 버전 |
| **3.x List ...** | 리스트 열기, 내 권한, 컬럼 표(내부명 ↔ 표시명 ↔ 형식), 컬럼별 값 형식, SRN 품질 |
| **4. SRN matching** | 두 리스트 사이의 SRN 연결 |
| **Summary** | 결론 4줄(Hosting / Provider / OData / Charts) + OK·WARN·FAIL 전체 목록 |

상태 표시:

| 표시 | 의미 |
|---|---|
| `[OK  ]` | 정상 |
| `[WARN]` | 확인 필요 (동작은 함) |
| `[FAIL]` | 해결 필요 |
| `[INFO]` | 참고 정보 |

2번 구역에서 `REST sign-in`이 `[OK]`여야 3·4번 점검이 진행됩니다.

---

## 4. 결과 해석

### 4.1 Summary 결론 4줄

| 항목 | 결과 | 의미 |
|---|---|---|
| Hosting | `SharePoint Server on-premises ...` | 설계 그대로 진행 (Windows 통합인증 + REST) |
| | `SharePoint Online ...` | 인증을 PnP.PowerShell 방식으로 바꿈. Provider만 교체하면 되고 나머지 설계는 동일 |
| Provider | `Rest ...` / `PnP.PowerShell ...` | 대시보드의 데이터 접근 방식 |
| OData | `odata=nometadata` / `odata=verbose` | verbose는 구버전(SP2013)이라는 뜻. 대시보드가 둘 다 처리 |
| Charts | `available` / `NOT available` | P4 차트를 내장 기능으로 만들지, 다른 방법을 쓸지 결정 |

### 4.2 컬럼별 값 형식 읽는 법

```
  field_6  "Err7_SupplyWaterHigh"  [Text]  filled 3/152  types: String
      shapes : "9999-99-99 99:99" x3
      values : "2026-10-22 14:05" x2 | "2026-10-21 23:10" x1

  AlertSent  "AlertSent"  [Text]  filled 152/152  types: String
      shapes : "aaaaa" x140 | "aaaa" x12
      values : "false" x140 | "true" x12
```

- `field_6 "Err7_SupplyWaterHigh"`: 내부명과 표시명. 사용자 컬럼 매핑표가 맞는지 바로 확인할 수 있습니다.
- `[Text]`: SharePoint 컬럼 형식. `types`: 실제로 받은 값의 형식.
- `filled 3/152`: 값이 들어 있는 행 수.
- `shapes`: 실제 값 대신 **모양**만 보여줍니다. `9` = 숫자, `A`/`a` = 영문 대/소문자, `k` = 한글.
- `values`: 아래 조건을 **모두** 만족할 때만 실제 값을 보여줍니다.
  - 고유값이 15개 이하
  - 마스킹 대상 컬럼이 아님
  - 값에 `@`가 없음
- `"true"`처럼 따옴표가 있으면 **텍스트**, `true (bool)`이면 **예/아니요 컬럼**입니다. AlertSent·AllOk 형식은 이것으로 판단합니다.

이 구역에서 설계 문서 12장의 2~4번 질문이 해결됩니다. 에러 컬럼 값이 시각인지 Yes인지, AllOk·AlertSent 형식, EASDB2 컬럼 구성을 여기서 확인합니다.

### 4.3 SRN 점검 (설계 문서 10장 1·2번을 실제 데이터로 확인)

| 결과 | 의미 | 조치 |
|---|---|---|
| `duplicate SRN rows` WARN | 같은 SRN 행이 2개 이상. Flow1이 동시에 실행됐을 가능성 | 트리거 동시성을 1로 설정하고, 표시된 ID의 중복 행 정리 |
| `SRN formatting` WARN | SRN에 따옴표나 공백이 섞여 있음 | Flow의 `ComposeSRN` 단계에서 정규화 |
| `SRN spelling between lists` WARN | 따옴표·공백·대소문자를 지워야 두 리스트가 일치 | 위와 같음 (두 Get 액션의 정규화 방식 차이) |
| `SRNs only in ...` WARN | 한쪽 리스트에만 있는 SRN | 두 리스트를 만든 시점이 다르면 정상일 수 있음. 숫자만 공유해 주세요 |

---

## 5. 문제 해결

| 증상 | 원인 | 조치 |
|---|---|---|
| `running scripts is disabled on this system` | 실행 정책 | 방법 A·B처럼 `-ExecutionPolicy Bypass`로 실행. 그래도 막히면 그룹 정책입니다. `Get-ExecutionPolicy -List` 결과를 공유해 주세요 (IT 협의 또는 코드 서명 필요) |
| `is not digitally signed` 또는 보안 경고 | 인터넷에서 받은 파일이라는 표시 | 파일 우클릭 → 속성 → "차단 해제" 체크, 또는 `Unblock-File .\Test-EDashEnvironment.ps1` |
| `DNS ... FAIL` | 사내망/VPN이 아님 | VPN 연결 후 다시 실행 |
| `REST sign-in FAIL : HTTP 401` (Bearer 없음) | Windows 인증이 거절됨 | 브라우저로 리스트가 열리는지 확인하고, 사이트가 인트라넷 영역에 등록돼 있는지 확인 |
| `REST sign-in FAIL : HTTP 401; WWW-Authenticate: Bearer` | SharePoint Online | **오류가 아니라 판정 결과입니다.** 그대로 공유해 주세요 (PnP 방식으로 전환) |
| `HTTP 403` + DNS 별칭이 `sharepoint.com` | SharePoint Online. 권한 문제가 아니라 토큰 인증이 필요한 것 | 정상적인 판정입니다 ([WARN]). [FLOW-EXPORT.md](FLOW-EXPORT.md)의 파일 방식을 사용 |
| `HTTP 403` (온프레미스) | 권한 없음 | 사이트 소유자에게 읽기 권한 요청 |
| `HTTP 407` | 프록시 인증 필요 | 결과 공유 |
| `redirected to a sign-in page` | ADFS 등 로그인 페이지로 이동함 | 결과 공유 (인증 방식 검토) |
| `open list FAIL ... 404 List does not exist.` | 리스트 URL 이름이 다름 | 브라우저 주소창의 `/Lists/<이름>/` 부분을 확인한 뒤 `-ListNames`로 지정 |
| `reading all columns failed ... retrying` | 특정 컬럼을 읽을 수 없음 | 자동으로 SRN 컬럼만 다시 읽습니다. 메시지를 그대로 공유해 주세요 |
| 창이 바로 닫힘 | `-File` 실행이 끝나면 창이 닫힘 | `.cmd`로 실행하거나 PowerShell 창에서 실행 |
| 콘솔에서 한글이 깨짐 | 콘솔 글꼴 문제 | 보고서 파일(UTF-8)에서는 정상적으로 보입니다 |

---

## 6. 결과 공유

- 보고서 위치: 바탕화면의 `EDash-test-yyyyMMdd-HHmm.txt`
- 기본으로 가려지는 것:
  - Title, SiteName, SystemName, FSE, SRN 컬럼
  - `@`가 들어간 값
  - 고유값이 많은 컬럼

  이런 값은 모양(`9`/`A`/`a`/`k`)으로만 나옵니다. 실제 값이 보이는 것은 Region, AlertSent, 예/아니요, 개수처럼 몇 가지 값만 반복되는 컬럼입니다.
- 공유하기 전에 한 번 훑어봐 주세요. 더 가려야 할 컬럼이 있으면 `-MaskFields`에 추가해서 다시 실행하거나, 해당 줄을 지우면 됩니다.
- 보고서 내용 전체를 이 대화에 붙여넣어 주시면 됩니다.

---

## 7. 이후 단계의 테스트 (대시보드 개발 후)

| 대상 | 방법 |
|---|---|
| 로직 단위 테스트 | Pester로 상태 판정, SRN 정규화, 에러 그룹, 변화 감지를 테스트. UI가 필요 없어서 어느 PC에서나 `Invoke-Pester .\tests`로 실행 |
| 오프라인 UI | `settings.psd1`의 `Provider = 'Mock'` + `samples/*.json`. 샘플 데이터는 이번 점검 결과의 형식에 맞춰 작성 |
| 실데이터 조회 | Rest Provider로 운영 리스트를 **읽기만** 함 (안전) |
| 쓰기 기능 (P3) | 운영 리스트에 바로 쓰지 않습니다. 리스트 복사본(예: `EasDashboard_Test`, `EASDB2_Test`)을 만들고 `settings.psd1`의 `Lists`만 바꿔서 테스트 |
| UI 확인 항목 | 갱신 중에도 창 이동·스크롤이 끊기지 않음 / 갱신 후 선택 행 유지 / 네트워크가 끊기면 표시되고 자동 복구 / 22시·08시 야간 전환 / 18:00 리셋 직후 알림이 쏟아지지 않음 |
