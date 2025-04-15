# SQA 주간 이슈 리포트 자동 생성기

Redmine에서 프로젝트 이슈 데이터를 가져와서 SQA 주간 리포트를 자동으로 생성해줍니다.
Python으로 작성되었고, 엑셀 파일로 결과가 저장됩니다. 각 프로젝트별 이슈 현황을 표로 정리하고, 누적 이슈 커브 그래프도 자동으로 삽입됩니다.


## 주요 기능

- 프로젝트별 이슈 수 및 미해결 수 집계
- 이벤트 단계별로 A~D 우선순위 + 이슈 상태(Open, In Progress, Closed 등)를 표로 정리
- 이번 주 등록된 신규 이슈만 따로 표로 정리
- 이슈 누적/해결 추이를 선 그래프로 시각화
- 엑셀 파일로 자동 저장 (`SQA_주간보고서_날짜.xlsx`)
- `.bat` 파일로 반복 작업 자동화


실행은 총 3가지 방법이 있습니다.

방법 1 - 터미널에서 실행
폴더에 들어가서 아래처럼 실행:

python weekly_report.py
방법 2 - VS Code에서 실행
파일 열고 ▶ 버튼 누르거나, 터미널 열고 아래래와 같이 입력합니다.

python weekly_report.py
방법 3 - .bat 파일 더블클릭
run_report.bat 파일은 이 저장소에 포함되어 있으므로, 폴더 안에서 바로 더블클릭하면 됩니다.

## 실행 방법법

1. Python 3 설치
2. 필라이브러리 설치:

    ```
    pip install requests openpyxl matplotlib pandas
    ```

3. `weekly_report.py` 파일에서 `API_KEY`를 본인의 Redmine 키로 변경경

4. 스크립트 실행:
```
python weekly_report.py
```
또는 `run_report.bat` 배치파일을 더블클릭하면 자동 실행



## 실행 결과

파일 이름은 자동으로 SQA_주간보고서_YYYYMMDD.xlsx 형식으로 저장

 - 프로젝트 이름, 전체 이슈 수, 잔여 이슈 수
 - 전체 이슈 요약 테이블 (단계별 × 우선순위 × 상태)
 - 이번 주 신규 이슈 테이블
 - 누적 이슈 커브 그래프 (자동 이미지 삽입됨)


## 작성자자

김예지 
SQA Engineer
GitHub: [@tmsnvl0415](https://github.com/tmsnvl0415)




