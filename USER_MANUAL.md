# Vision Issue Tracker 사용자 매뉴얼

## 1. 프로그램 목적

Vision Issue Tracker는 생산 라인의 Vision Inspection Instrument 이슈를 기록, 검색, 보고서화하고 SW/Algo 버전 변경 이력을 관리하기 위한 Windows 데스크톱 프로그램입니다.

주요 사용 목적:
- 라인별/비전별 이슈 기록
- 미해결 이슈 추적
- 조건별 검색 및 Excel 보고서 생성
- 비전별 SW Version / Algo Version 관리
- 버전 업데이트 시 Monitoring 이슈 자동 등록

## 2. 실행 및 배포 파일

### 실행 파일

배포용 실행 파일:

```text
release\VisionIssueTracker.exe
```

Python 설치 여부와 관계없이 실행되도록 빌드되어 있습니다.

### DB 파일 위치

프로그램은 실행 파일이 있는 위치 기준으로 아래 DB를 사용합니다.

```text
data\vision_issues.db
```

처음 실행 시 DB가 없으면 자동으로 생성됩니다.

### 배포 시 필요한 파일 구조

새 PC에서 빈 DB로 시작할 경우:

```text
VisionIssueTracker.exe
```

기존 데이터를 같이 배포할 경우:

```text
VisionIssueTracker.exe
data\
  vision_issues.db
```

예시:
- 작업자 노트북 4대가 같은 공유 폴더의 프로그램을 실행하려면 공유 폴더에 `VisionIssueTracker.exe`와 `data\vision_issues.db`를 같이 둡니다.
- 기존 기록을 유지하려면 반드시 `data\vision_issues.db`를 함께 복사해야 합니다.

주의:
- SQLite DB는 여러 PC가 동시에 저장할 때 충돌 가능성이 있습니다.
- 동시에 여러 명이 자주 저장하는 운영이면 추후 서버 DB 방식이 더 안전합니다.

## 3. 기본 구성

### 라인

```text
1-1
1-2
2-1
2-2
```

### 비전 종류

```text
Pinhole
Pouch Align
Lead
Sealing
Lead Align
Welding(+)
Welding(-)
```

### 작업자

이슈 작성자는 `Issue Board`의 `Create Issue` / `Edit Issue` 패널 안에 있는 `Logged By`에서 선택합니다.
버전 업데이트 작성자는 `Version Update` 영역의 `Logged By`에서 선택합니다.

```text
Hojun Kwak
Kijung Kim
Jihoon Yun
Jisub Yun
```

예시:
- Jihoon Yun이 이슈를 작성하는 경우 `Create Issue` 패널에서 `Logged By`를 `Jihoon Yun`으로 선택한 뒤 저장합니다.

### 언어

우측 상단 언어 선택에서 변경할 수 있습니다.

```text
English
한국어
```

프로그램 시작 기본 언어는 한국어입니다.

## 4. Issue Board

`Issue Board`는 프로그램 시작 시 처음 보이는 화면입니다.
현재 조치가 필요하거나 모니터링 중인 이슈를 카드 형태로 보여줍니다.

보드 컬럼:
- Action Required: 조치 필요
- Monitoring: 모니터링

Resolved 상태는 Issue Board에서 제외되고 `Search / Report` 탭에서 확인할 수 있습니다.

### 이슈 카드

각 카드는 운영에 필요한 핵심 정보만 표시합니다.

표시 항목:
- Title
- Line / Vision
- Category
- Issue Time
- Downtime Duration

예시 카드:

```text
Pinhole camera disconnect
1-1 / Pinhole
Hardware / Camera
2026-06-18 08:35
Downtime Duration: 00:15
```

### 상세 패널

카드를 클릭하면 오른쪽 상세 패널에 전체 내용이 표시됩니다.
상세 패널은 마우스를 패널 위에 올린 상태에서 휠로 스크롤할 수 있습니다.

표시 항목:
- Title
- Status
- Line / Instrument
- Category
- Issue Time
- Downtime Duration
- Logged By
- Description
- Resolution Notes

### 상세 패널 버튼

Edit:
- 선택한 이슈를 오른쪽 패널의 `Edit Issue` 폼으로 전환합니다.

Move to Monitoring:
- Action Required 이슈를 Monitoring으로 이동합니다.

Move to Action Required:
- Monitoring 이슈를 Action Required로 되돌립니다.

Resolved:
- 선택한 이슈를 해결 완료 처리합니다.
- Downtime Duration이 비어 있으면 발생 시간 기준으로 자동 계산됩니다.
- 처리 후 해당 카드는 Issue Board에서 사라집니다.

Delete:
- 선택한 이슈를 삭제합니다.
- 삭제 전 확인창이 표시됩니다.

예시:
1. `1-1 Pinhole` 카드를 클릭합니다.
2. 조치 후 관찰이 필요하면 `Move to Monitoring`을 누릅니다.
3. 해결이 끝났으면 `Resolved`를 누릅니다.
4. 해당 이슈는 Issue Board에서 사라지고 `Search / Report`에서 `Resolved` 상태로 확인됩니다.

## 5. 이슈 등록 / 수정

새 이슈 입력과 기존 이슈 수정은 모두 `Issue Board` 오른쪽 패널에서 처리합니다.

### 신규 이슈 등록

`Issue Board` 우측 상단의 `Create Issue` 버튼을 누르면 항상 빈 신규 폼이 표시됩니다.
이전에 선택했던 카드 정보는 신규 폼에 남지 않습니다.

신규 저장 버튼:
- Create Issue

보조 버튼:
- Cancel

### 기존 이슈 수정

기존 카드를 클릭한 뒤 상세 패널에서 `Edit`을 누르면 `Edit Issue` 폼이 표시됩니다.

수정 저장 버튼:
- Save Changes

보조 버튼:
- Cancel

### 라인 선택

라인 버튼 중 하나를 선택합니다.

예시:

```text
1-1
```

### 비전 선택

비전 버튼은 복수 선택이 가능합니다.

예시 1:

```text
Pinhole
```

예시 2:

```text
Welding(+) / Welding(-)
```

여러 비전에 같은 이슈가 발생하면 해당 비전을 모두 선택합니다.

### Issue Time

발생 시간을 입력합니다.

형식:

```text
YYYY-MM-DD HH:MM
```

날짜 입력칸을 클릭하면 달력이 표시됩니다. 시간은 `HH : MM` 형식으로 스핀 버튼을 사용해 조정할 수 있습니다.

예시:

```text
2026-06-18 08:35
```

### Category / Subcategory

사용 가능한 분류:

```text
Hardware
- Camera
- Lighting

Software
- Program Crash
- Program Update
- UI
- PLC
- Other

Recipe
- Overkill
- Underkill
- Add Measure
- Bypass/Unbypass

Camera Grab Fail
Production
Other
```

예시:
- 카메라 연결 불량: `Hardware > Camera`
- 프로그램 다운: `Software > Program Crash`
- PLC 통신 이상: `Software > PLC`
- 과검출: `Recipe > Overkill`
- 카메라 Grab 실패: `Camera Grab Fail`

### Status

사용 가능한 상태:

```text
Action Required
Monitoring
Resolved
```

상태 의미:
- Action Required: 조치가 아직 필요한 상태
- Monitoring: 조치는 했지만 결과를 지켜보는 상태
- Resolved: 해결 완료

### Downtime Duration

다운타임 시간을 입력합니다.

기본값:

```text
00:00
```

예시:
- 다운타임 없음: `00:00`
- 35분 정지: `00:35`
- 1시간 20분 정지: `01:20`

### Title

이슈 제목을 짧고 명확하게 입력합니다.

좋은 예:

```text
Pinhole camera disconnect
Welding program crash during auto run
Lead Align overkill after recipe change
```

피하면 좋은 예:

```text
Issue
Problem
Check needed
```

### Description

상세 증상과 상황을 입력합니다.

예시:

```text
1-1 Pinhole inspection 중 camera timeout 발생.
재시작 후 정상 복귀했으나 동일 증상 재발 가능성 있어 Monitoring 필요.
```

### Resolution Notes

조치 내용을 입력합니다.

예시:

```text
Camera cable 재체결 후 vision program restart.
10 lot monitoring 결과 재발 없음.
```

### 저장 예시

예시 상황:
- 라인: 1-1
- 비전: Pinhole
- 발생 시간: 2026-06-18 08:35
- 카테고리: Hardware > Camera
- 상태: Action Required
- 제목: Pinhole camera disconnect
- 설명: Camera connection lost during production.

입력 후 `Create Issue`를 누르면 Issue Board에 카드가 표시됩니다.
기존 이슈를 수정한 경우 `Save Changes`를 누르면 카드와 상세 패널이 갱신됩니다.

## 6. Search / Report 탭

이슈를 조건별로 검색하고 Excel 파일로 저장하는 탭입니다.

### 검색 조건

사용 가능한 필터:
- Status
- Line
- Category
- Subcategory
- Keyword
- From
- To
- Vision Filter

날짜 범위는 기본적으로 DB에 있는 첫 이슈 시간과 가장 최근 이슈 시간으로 설정됩니다.
검색 결과 테이블은 가로 스크롤을 지원하므로 제목이나 비전명이 길어도 좌우로 이동해 확인할 수 있습니다.

### Vision Filter

비전 필터도 복수 선택이 가능합니다.

예시:
- `Pinhole`만 검색
- `Welding(+)`와 `Welding(-)`를 동시에 검색
- `Lead`, `Lead Align` 관련 이슈 검색

### Quick Filter

버튼:
- Today
- This Week
- Action Required
- Monitoring
- Camera Grab Fail
- Recipe Issues
- Clear

예시:
- 오늘 발생한 이슈만 보고 싶으면 `Today`
- Recipe 관련 이슈만 보고 싶으면 `Recipe Issues`
- 필터를 초기화하려면 `Clear`

### Excel 저장

검색 결과를 Excel로 저장하려면 `Excel` 버튼을 누릅니다.

Excel 보고서 특징:
- 검색 결과만 저장됩니다.
- ID는 실제 DB ID가 아니라 검색 결과의 순번으로 표시됩니다.
- Issue Time 기준으로 정렬됩니다.
- Excel 컬럼 순서: `ID`, `Line`, `Instrument`, `Issue Time`, `Downtime`, `Category`, `Title`, `Status`, `Description`, `Resolution Notes`
- `Downtime` 컬럼은 기본적으로 숨김 처리됩니다.
- `Title`, `Description`, `Resolution Notes`를 제외한 표시 컬럼은 입력 길이에 맞춰 여백이 최소화됩니다.
- `Description`은 기존 고정 폭을 유지하고, `Resolution Notes`는 약 800px 폭으로 고정되며 셀 안에서 자동 줄바꿈됩니다.

예시:
- 전체 100건 중 Camera Grab Fail 2건만 검색 후 Excel 저장하면 ID는 `1`, `2`로 저장됩니다.

### 삭제

검색 결과 테이블에서 이슈를 선택하고 `Delete`를 누르면 삭제할 수 있습니다.

삭제 전 확인창이 표시됩니다.

### 상세 확인

검색 결과를 더블클릭하면 `Issue Board` 탭의 오른쪽 상세 패널에서 해당 이슈를 확인할 수 있습니다.
Resolved 이슈도 상세 확인과 수정이 가능합니다.

## 7. 버전 기록 탭

비전별 SW Version과 Algo Version을 관리하는 탭입니다.

구성:
- Version Dashboard
- Version Update
- Version Description

## 8. Version Dashboard

각 라인/비전의 최신 버전 상태를 표시합니다.

표시 단위:
- 4개 라인
- 7개 비전

총 28개 조합의 현재 버전을 확인할 수 있습니다.

셀 표시 내용:
- SW Version
- Algo Version

라인명은 왼쪽 행 헤더에, 비전명은 상단 열 헤더에 표시됩니다.
셀 안에는 현재 적용된 SW/Algo만 compact하게 표시됩니다.

최근 7일 내 업데이트된 항목은 셀 왼쪽의 초록색 바로 표시됩니다.
같은 비전의 다른 라인 대비 버전이 낮은 경우 오른쪽에 작은 색상 점이 표시됩니다.

표시 의미:
- 주황색 점: SW Version이 해당 비전의 최신 SW보다 낮음
- 파란색 점: Algo Version이 해당 비전의 최신 Algo보다 낮음
- 초록색 바: 최근 7일 내 업데이트됨

예시:

```text
SW 260522.1450
A  1.2.3.4
```

버전 정보가 없으면 `-`로 표시됩니다.

Sealing은 별도 Algo Version이 없으므로 SW Version만 표시됩니다.

예시:

```text
SW 260522.1450
```

### 대시보드 Excel 추출

우측 상단 `대시보드 추출` 버튼으로 현재 Version Dashboard 정보를 Excel로 저장할 수 있습니다.

Excel 항목:
- Line
- Vision
- Group
- SW Version
- Algo Version
- Last Updated

예시 사용:
- 현재 전체 라인/비전 버전 현황을 회의 자료로 저장
- 특정 날짜 기준 version snapshot 보관

## 9. Version Group

버전 그룹은 버전 템플릿을 관리하기 위한 묶음입니다.

```text
Welding
- Welding(+)
- Welding(-)

Common
- Pinhole
- Pouch Align
- Lead Align

New Lead
- Lead

Sealing
- Sealing
```

중요:
- 그룹은 버전 후보를 묶는 단위입니다.
- 실제 적용 버전은 각 라인/비전별로 다를 수 있습니다.

예시:
- Common 그룹에 SW 260522.1450이 있어도 `Pinhole`만 먼저 업데이트하고 `Pouch Align`은 이전 버전을 유지할 수 있습니다.

## 10. Version Update

새 버전 적용 기록을 입력하는 영역입니다.

입력 항목:
- SW Version
- Algo Version
- Update Time
- Logged By
- 모니터링 이슈 등록
- Line
- Vision
- SW Description
- Algo Description

Version Update 패널은 라인/비전 선택 버튼이 가장 위에 있습니다.
비전을 선택하면 해당 비전이 속한 그룹 기준으로 같은 그룹의 비전들을 함께 선택할 수 있습니다.

예시:
- `Welding(+)`를 선택하면 `Welding(-)`도 함께 선택 가능
- `Pinhole`을 선택하면 `Pouch Align`, `Lead Align`도 함께 선택 가능
- `Lead`와 `Sealing`은 각각 단독 그룹으로 관리

Sealing은 별도 Algo Version이 없으므로 Algo 입력칸과 Algo 설명은 사용하지 않습니다.

### Update Time

버전 적용 시간을 수기로 입력합니다.

형식:

```text
YYYY-MM-DD HH:MM
```

예시:

```text
2026-06-18 10:20
```

### 라인/비전 복수 선택

라인과 비전을 복수로 선택할 수 있습니다.

예시:
- `1-1`, `1-2` 라인에 Welding(+)만 업데이트
- `2-1` 라인에 Welding(+)와 Welding(-) 동시 업데이트
- Common 그룹에서 Pinhole, Pouch Align만 먼저 업데이트

### 모니터링 이슈 등록

`☑ 모니터링 이슈 등록`이 켜져 있으면 버전 저장 시 이슈가 자동 생성됩니다.

자동 생성 이슈:
- Category: Software
- Subcategory: Program Update
- Status: Monitoring
- Issue Time: Version Update의 Update Time과 동일
- Downtime Duration: 00:00

예시:

버전 업데이트:

```text
Line: 1-1
Vision: Pinhole
SW Version: 260522.1450
Algo Version: 1.2.3.4
Update Time: 2026-06-18 10:20
```

자동 생성 이슈:

```text
Software > Program Update
Monitoring
Program Update - 1-1 Pinhole SW 260522.1450 / Algo 1.2.3.4
```

## 11. Version Description

Version Description 패널은 그룹별 버전 템플릿을 확인, 수정, 삭제하는 영역입니다.

### 그룹 선택

4개 그룹 중 하나를 선택합니다.

```text
Welding
Common
New Lead
Sealing
```

선택한 그룹의 SW Version 리스트와 Algo Version 리스트가 좌우로 분리되어 표시됩니다.
Sealing은 Algo Version을 사용하지 않으므로 SW 리스트만 넓게 표시됩니다.

### 버전 내용 수정

수정 가능 항목:
- SW Version
- SW Description
- Algo Version
- Algo Description

SW와 Algo는 각각 선택, 수정, 저장할 수 있습니다.
예를 들어 SW 설명만 바꾸려면 왼쪽 SW 리스트에서 버전을 선택하고 `Save SW`를 누릅니다.
Algo 설명만 바꾸려면 오른쪽 Algo 리스트에서 버전을 선택하고 `Save Algo`를 누릅니다.

Sealing 그룹은 Algo Version을 사용하지 않으므로 SW Version과 SW Description만 수정합니다.

수정 영향:
- 해당 그룹의 선택한 SW 또는 Algo로 기록된 version history도 함께 업데이트됩니다.
- Version Dashboard에도 변경 내용이 반영됩니다.

예시:

수정 전:

```text
SW Version: 260522.1450
SW Description: ROI threshold update
```

수정 후:

```text
SW Version: 260522.1530
SW Description: ROI threshold update and PLC handshake delay fix
```

### 버전 삭제

삭제할 SW 또는 Algo 버전을 선택하고 `Delete SW` 또는 `Delete Algo`를 누릅니다.

삭제 영향:
- 해당 SW 또는 Algo 버전 템플릿이 삭제됩니다.
- 해당 그룹에서 선택한 SW 또는 Algo로 적용됐던 version history 기록도 삭제됩니다.
- Version Dashboard는 남아있는 이전 기록 기준으로 자동 갱신됩니다.

예시:
- `Welding / SW 260522.1450`을 삭제하면 해당 SW로 표시되던 라인/비전은 이전 기록 기준으로 돌아가거나, 이전 기록이 없으면 `-`로 표시됩니다.
- `Welding / Algo 1.2.3.4`를 삭제하면 해당 Algo로 표시되던 라인/비전도 동일하게 갱신됩니다.

주의:
- 버전 삭제 시 자동 생성됐던 이슈 로그는 삭제되지 않습니다.
- 이슈 로그 삭제가 필요하면 Issue Board 또는 Search / Report 탭에서 별도로 삭제합니다.

## 12. 권장 입력 규칙

### 제목 작성 규칙

권장 형식:

```text
[Vision] + symptom
```

예시:

```text
Pinhole camera timeout
Lead Align overkill after recipe update
Welding(-) program crash
```

### Description 작성 규칙

아래 정보를 포함하면 검색과 분석에 유리합니다.

```text
1. 발생 상황
2. 증상
3. 임시 조치
4. 재발 여부
5. 추가 확인 필요 사항
```

예시:

```text
Auto run 중 Welding(-) inspection program crash 발생.
Lot change 직후 2회 재발.
Program restart 후 정상 복귀했으며 동일 조건에서 Monitoring 필요.
```

### Resolution Notes 작성 규칙

예시:

```text
Recipe parameter rollback.
Camera exposure value restored from 1200 to 950.
20 trays monitoring result: no recurrence.
```

## 13. 일반 작업 예시

### 예시 A: 카메라 하드웨어 이슈 등록

1. `Issue Board` 탭에서 `Create Issue` 클릭
2. Logged By: `Jihoon Yun`
3. 라인: `1-1`
4. 비전: `Pinhole`
5. Issue Time: `2026-06-18 08:35`
6. Category: `Hardware`
7. Subcategory: `Camera`
8. Status: `Action Required`
9. Downtime Duration: `00:15`
10. Title: `Pinhole camera disconnect`
11. Description 입력
12. `Create Issue`

결과:
- Issue Board의 Action Required 컬럼에 카드로 표시됩니다.

### 예시 B: 프로그램 업데이트 기록

1. `버전 기록` 탭 선택
2. Line: `1-1`, `1-2`
3. Vision: `Pinhole`, `Pouch Align`
4. SW Version: `260522.1450`
5. Algo Version: `1.2.3.4`
6. Update Time: `2026-06-18 10:20`
7. Logged By: `Jihoon Yun`
8. `☑ 모니터링 이슈 등록` 켜기
9. SW Description:

```text
False reject 개선을 위해 ROI threshold 및 Add Measure logic update.
```

10. Algo Description:

```text
Inspection threshold version 1.2.3.4 적용.
```

11. `Save Version Update`

결과:
- 선택한 라인/비전에 버전 기록 생성
- Version Dashboard 업데이트
- Software > Program Update / Monitoring 이슈 자동 생성

### 예시 C: 이전 버전으로 롤백 기록

1. `버전 기록` 탭 선택
2. Line: `2-1`
3. Vision: `Welding(+)`, `Welding(-)`
4. 이전 SW Version / Algo Version을 직접 입력
5. Update Time 입력
6. SW Description:

```text
New version에서 intermittent grab delay 발생하여 이전 안정 버전으로 rollback.
```

7. `Save Version Update`

결과:
- Version Dashboard가 이전 버전으로 표시됩니다.
- Monitoring 이슈가 자동 생성됩니다.

### 예시 D: 검색 후 Excel 보고서 저장

1. `Search / Report` 탭 선택
2. Category: `Software`
3. Subcategory: `Program Update`
4. From / To 날짜 확인
5. `Search`
6. `Excel`

결과:
- Program Update 관련 이슈만 Excel 파일로 저장됩니다.

## 14. 데이터 백업

백업해야 할 파일:

```text
data\vision_issues.db
```

권장 백업 방식:
- 매일 작업 종료 후 날짜를 붙여 복사

예시:

```text
vision_issues_2026-06-18.db
```

복구 방법:
1. 프로그램 종료
2. 기존 `data\vision_issues.db`를 백업본으로 교체
3. 프로그램 재실행

## 15. 문제 해결

### 프로그램은 실행되지만 데이터가 안 보일 때

확인할 것:
- 실행 중인 `VisionIssueTracker.exe` 옆에 `data\vision_issues.db`가 있는지 확인
- 다른 폴더의 exe를 실행하고 있지 않은지 확인

### 다른 PC에서 데이터가 다르게 보일 때

원인:
- 각 PC가 서로 다른 위치의 DB를 사용 중일 수 있습니다.

해결:
- 모든 작업자가 같은 공유 폴더의 `VisionIssueTracker.exe`를 실행하도록 합니다.
- 또는 같은 `data\vision_issues.db`를 사용하도록 폴더 구조를 맞춥니다.

### Excel 저장이 안 될 때

확인할 것:
- 저장하려는 Excel 파일이 이미 열려 있지 않은지 확인
- 저장 위치에 쓰기 권한이 있는지 확인
- 파일명이 너무 길거나 특수문자가 포함되어 있지 않은지 확인

### 삭제한 이슈/버전 복구

현재 프로그램에는 삭제 취소 기능이 없습니다.

복구하려면:
- 백업해둔 `vision_issues.db` 파일로 복구해야 합니다.

## 16. 운영 주의사항

- 삭제 전에는 반드시 확인창 내용을 확인합니다.
- 버전 삭제는 Version Dashboard 표시에도 영향을 줍니다.
- 여러 사용자가 동시에 저장하는 환경에서는 충돌 가능성이 있습니다.
- 중요한 변경 전에는 DB 파일을 백업하는 것이 좋습니다.
