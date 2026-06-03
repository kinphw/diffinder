# Diffinder

Excel 셀 텍스트 비교(diff) 추가 기능. 두 셀의 텍스트를 줄/단어/문자 단위로 비교해,
결과 셀에 **삭제(빨강 취소선)** / **추가(파랑 굵게)** 로 표시합니다.

## 리본 ("Diffinder" 탭)

| 버튼 | 동작 |
|---|---|
| **직접 입력** | 팝업으로 [이전] → [이후] → [결과] 셀을 차례로 지정해 비교 |
| **셀 선택 비교** | [이전][이후] 2개 열 또는 [이전][이후][결과] 3개 열을 선택해 일괄 비교 |
| **정보** | 버전 정보 표시 |

## 구조

```
version.json          버전·작성자 SSOT (단일 출처)
src/ModDiff.bas       diff 엔진 + 진입점 매크로 (원본, CP949 인코딩)
src/ModRibbon.bas     리본 콜백 + 정보창 (버전 상수는 빌드 시 주입)
ribbon/customUI14.xml 리본 정의
build.ps1             소스 → build/Diffinder.xlam 빌드 스크립트
```

## 버전 관리 (SSOT)

버전·작성자는 루트 [version.json](version.json) 한 곳에서만 관리합니다.

```json
{ "name": "Diffinder", "version": "0.1.0", "author": "kinphw" }
```

`build.ps1` 가 빌드 시 이 값을 `src/ModRibbon.bas` 의
`APP_NAME` / `APP_VERSION` / `APP_AUTHOR` 상수에 주입합니다(리본 "정보" 버튼이
표시하는 값). 소스의 상수 기본값은 빌드 없이 수동 임포트했을 때의 fallback일
뿐이므로, **버전을 올릴 때는 `version.json` 만 수정**하면 됩니다.

`build/Diffinder.xlam` 은 빌드 산출물이라 git에 포함하지 않습니다(`.gitignore`).
배포본은 `build.ps1` 로 재생성하거나 GitHub Release 에 첨부하세요.

## 빌드

요구사항: Microsoft Excel 설치 + Excel 신뢰 센터에서
"VBA 프로젝트 개체 모델에 대한 액세스를 신뢰함" 활성화.

```powershell
pwsh -File build.ps1
# -> build/Diffinder.xlam 생성
```

## 설치 (사용자)

1. `Diffinder.xlam` 을 임의의 폴더에 둡니다.
2. Excel → **파일 > 옵션 > 추가 기능** → 관리: **Excel 추가 기능** → **이동** →
   **찾아보기** 로 `Diffinder.xlam` 선택 → 체크 후 확인.
3. 리본에 **Diffinder** 탭이 나타납니다.

> 메일/인터넷으로 받은 경우 매크로가 차단될 수 있습니다.
> 파일 **우클릭 > 속성 > "차단 해제"** 후 사용하세요. (서명되지 않은 VBA)
