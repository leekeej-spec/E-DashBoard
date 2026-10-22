# 데이터 내보내기 Flow 만들기 (SharePoint Online용)

## 왜 이 방식인가

`share.philips.com`은 **SharePoint Online**(`philips.sharepoint.com`)의 별칭입니다. SharePoint Online은 Windows 계정만으로는 REST 호출을 받아 주지 않아서, 환경 점검에서 HTTP 403이 나옵니다. PowerShell에서 직접 로그인 토큰을 받으려면 회사 Entra ID(Azure AD)에 앱을 등록해야 하는데, 보통 IT 승인이 필요합니다.

그래서 이미 쓰고 있는 도구들로 데이터를 넘깁니다. IT 승인, 앱 등록, 프리미엄 커넥터가 모두 필요 없습니다.

```
Power Automate (3분마다)                  OneDrive 앱(동기화)              E-Dashboard
  EasDashboard 항목 가져오기   ─┐
  EASDB2 항목 가져오기         ─┴─▶ OneDrive\EDashboard\edash-data.json ──▶ PC의 같은 파일 ──▶ 60초마다 읽기
```

---

## 1. 빈 파일 준비 (1회)

1. 파일 탐색기에서 회사 OneDrive 폴더(`OneDrive - Philips`)를 엽니다.
2. `EDashboard` 폴더를 만듭니다.
3. 그 안에서 우클릭 → 새로 만들기 → 텍스트 문서를 만들고, 이름을 **`edash-data.json`** 으로 바꿉니다.
   - 확장자가 보이지 않으면 탐색기 "보기 → 파일 확장명"을 켜세요. `.txt`가 남아 있으면 안 됩니다.
4. 몇 초 뒤 OneDrive 웹(`https://philips-my.sharepoint.com`)에도 같은 파일이 보이면 준비 완료입니다.

---

## 2. Flow 만들기

Power Automate → **만들기 → 예약된 클라우드 흐름**

| 항목 | 값 |
|---|---|
| 흐름 이름 | `EAS Dashboard Export` |
| 반복 간격 | **3분** |

> 1분 간격은 하루 1,440회 실행이라 Power Platform 요청 한도에 걸릴 수 있습니다. 3분이면 하루 480회 × 액션 3개라 여유가 있습니다.

### 2-1. 항목 가져오기: EasDashboard

**+ 새 단계 → SharePoint → 항목 가져오기 (Get items)**

| 항목 | 값 |
|---|---|
| 사이트 주소 | `https://share.philips.com/sites/EASalarmhandling` (목록에서 선택) |
| 목록 이름 | `EasDashboard` |
| 고급 옵션 → **상위 개수 (Top Count)** | **5000** ← 비워 두면 100건까지만 가져옵니다 |

동작 이름을 **`GetDash`** 로 바꿉니다. 동작 제목의 `…` → 이름 바꾸기에서 공백 없이 입력합니다.

### 2-2. 항목 가져오기: EASDB2

같은 방법으로 하나 더 만듭니다.

| 항목 | 값 |
|---|---|
| 사이트 주소 | 위와 같음 |
| 목록 이름 | `EASDB2` |
| 상위 개수 | **5000** |

동작 이름은 **`GetAlert`** 로 바꿉니다.

### 2-3. 파일 업데이트 (OneDrive for Business)

**+ 새 단계 → OneDrive for Business → 파일 업데이트 (Update file)**

| 항목 | 값 |
|---|---|
| 파일 | 폴더 아이콘 → `/EDashboard/edash-data.json` 선택 |
| 파일 콘텐츠 | 아래 **식**을 붙여넣기 |

파일 콘텐츠 칸을 클릭 → **식(fx)** 탭 → 아래 한 줄을 그대로 붙여넣고 → 확인(추가):

```
concat('{"generatedAt":"', utcNow(), '","dashboard":', string(outputs('GetDash')?['body/value']), ',"alerts":', string(outputs('GetAlert')?['body/value']), '}')
```

- `generatedAt`은 내보낸 시각입니다. 대시보드는 이 값으로 "데이터가 몇 분 전 것인지"를 표시하고, 15분이 넘으면 경고합니다.
- 동작 이름을 `GetDash`/`GetAlert`가 아닌 다른 이름으로 했다면, 식 안의 이름도 똑같이 바꿔야 합니다. 공백은 `_`로 바뀝니다. 예: `항목 가져오기` → `항목_가져오기`.

### 2-4. 저장 → 테스트

1. 저장한 뒤 **테스트 → 수동으로 → 실행**을 누릅니다.
2. 세 단계가 모두 초록색 체크면 성공입니다.
3. PC의 `OneDrive - Philips\EDashboard\edash-data.json`을 메모장으로 열어 `{"generatedAt":"...","dashboard":[...` 로 시작하는지 확인합니다.

---

## 3. 대시보드 실행

`EDashboard.cmd`를 실행합니다. 기본 데이터 파일 위치가 `%OneDriveCommercial%\EDashboard\edash-data.json`이라서 따로 설정할 것이 없습니다.

| 상황 | 방법 |
|---|---|
| 파일을 다른 곳에 두었을 때 | `EDashboard.cmd -DataFile "D:\경로\edash-data.json"` |
| 콘솔에서 내용만 확인하고 싶을 때 | `EDashboard.cmd -CheckOnly` |

---

## 4. 문제 해결

| 증상 | 원인 / 조치 |
|---|---|
| `데이터 파일이 없습니다` | 1단계의 파일 이름·위치를 확인합니다 (`.json.txt`가 아닌지). 또는 `-DataFile`로 경로를 지정합니다. |
| `데이터 파일이 비어 있습니다` | Flow가 아직 한 번도 실행되지 않았습니다. 2-4의 테스트를 실행하세요. |
| `데이터 파일을 읽지 못했습니다` | OneDrive가 파일을 쓰는 도중에 읽은 경우로, 다음 갱신 때 자동으로 해결됩니다. 계속 나오면 2-3의 식을 다시 확인하세요. |
| 헤더에 `● 데이터 오래됨` | Flow가 멈췄거나 OneDrive 동기화가 멈췄습니다. Power Automate의 "실행 기록"과 작업 표시줄의 OneDrive 구름 아이콘을 확인하세요. |
| 시스템 수가 정확히 100개에서 멈춤 | `항목 가져오기`의 **상위 개수**를 5000으로 설정하세요. |
| Flow 실패: `GetDash` 를 찾을 수 없음 | 동작 이름과 식 안의 이름이 다릅니다 (2-3 참고). |

---

## 5. 참고

- **지연 시간:** Flow 간격(최대 3분) + OneDrive 동기화(보통 수 초~1분) + 대시보드 갱신(60초). 위험 알림 자체는 기존 Teams 알림이 즉시 처리하고, 대시보드는 현황 파악용입니다.
- **더 빠르게 하고 싶을 때:** 메일 처리 Flow(Flow1)와 18:00 초기화 Flow(Flow2)의 맨 끝에 2-1~2-3과 같은 단계를 붙이면, 변경 직후 바로 파일이 갱신됩니다. 다만 Flow1이 동시에 여러 개 실행되면 늦게 끝난 실행이 파일을 덮어쓸 수 있습니다. 그래서 3분 예약 Flow는 그대로 두는 것을 권장합니다.
---

## 6. 다른 사람과 같이 보기

Flow는 하나만 있으면 됩니다. 다른 사람은 데이터 파일만 받아 보면 됩니다.

### 방법 A: 내 OneDrive의 `EDashboard` 폴더 공유 (가장 간단)

1. **공유하는 사람(나):** OneDrive 웹에서 `EDashboard` 폴더 → 공유를 누릅니다.
   - 대상: 함께 볼 사람이나 팀 그룹
   - 권한: **보기 가능** (편집 불가)
2. **받는 사람:** OneDrive 웹 → **공유됨** → `EDashboard` 폴더 → **"내 파일에 바로 가기 추가"**를 누릅니다.
   - 그러면 그 사람 PC의 `OneDrive - Philips\EDashboard\edash-data.json`으로 동기화됩니다.
   - 대시보드 기본 경로와 같아서 `EDashboard.cmd`만 실행하면 됩니다.
   - 같은 이름의 폴더가 이미 있으면 바로 가기 이름이 달라질 수 있습니다. 이때는 `-DataFile`로 경로를 지정합니다.

주의할 점:
- **공유 범위:** EASalarmhandling 사이트에 이미 접근 권한이 있는 사람에게만 공유하세요. 파일에는 사이트명, FSE, SRN 같은 리스트 내용이 그대로 들어 있습니다. "Philips의 모든 사용자" 링크처럼 넓게 공유하면, 사이트 권한이 없는 사람도 이 데이터를 보게 됩니다.
- **내 계정에 의존:** Flow와 파일이 모두 내 계정에 묶여 있습니다. 퇴사·이동·장기 부재로 계정이 바뀌면 모두의 대시보드가 멈춥니다. Flow에 공동 소유자를 한 명 추가해 두세요(Flow 세부 정보 → 소유자 → 편집).

### 방법 B: EASalarmhandling 사이트의 문서 라이브러리에 저장 (팀 운영용)

계정에 묶이지 않는 방식입니다. 오래 쓸 계획이면 이쪽을 권장합니다.

1. **폴더 만들기:** 사이트의 **Documents** 라이브러리 맨 위(`General` 폴더 옆)에서 **+ 새로 만들기 → 폴더 → `EDashboard`** 를 만듭니다.
   - `General`은 Teams 채널용 폴더입니다. 그 안에 두면 채널의 파일 탭에 노출되어, 누가 옮기거나 지울 수 있습니다.
2. **파일 올리기:** 웹의 "새로 만들기"에는 텍스트 파일 항목이 없습니다.
   - PC에서 메모장으로 `{}` 두 글자만 적어 `edash-data.json`으로 저장합니다.
   - 그 파일을 `EDashboard` 폴더에 끌어다 놓아 업로드합니다.
3. **Flow 마지막 단계 교체:** OneDrive 단계 대신 **SharePoint → 파일 업데이트 (Update file)** 를 넣습니다.

   | 항목 | 값 |
   |---|---|
   | 사이트 주소 | EASalarmhandling |
   | 파일 식별자 | 폴더 아이콘 → `/Shared Documents/EDashboard/edash-data.json` |
   | 파일 콘텐츠 | 2-3과 같은 식 |

4. **각자 동기화:** 라이브러리의 `EDashboard` 폴더에서 **"내 파일에 바로 가기 추가"** 를 누릅니다.
   - 그러면 `OneDrive - Philips\EDashboard\edash-data.json`으로 동기화되어, 대시보드 기본 경로와 같아집니다.
   - 본인 OneDrive에 같은 이름의 `EDashboard` 폴더(방법 A)가 이미 있으면, 그 폴더를 지우거나 이름을 바꾼 뒤에 추가하세요.

방법 B는 권한이 사이트 권한을 그대로 따르기 때문에, 공유 범위를 따로 관리할 필요가 없습니다.
