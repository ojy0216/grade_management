# Grade Management for CAU students(v0.5)
* 중앙대학교 학생들을 위한 성적확인 및 계산 프로그램입니다.
##  기능
* 보관 성적 및 세부 성적 확인
* 미래 성적 시뮬레이션
* 기존 성적 수정 및 미래 성적 추가 기능
##  주의사항
* **현재 재수강은 지원되지 않습니다.(재수강이 존재하는 파일을 못구함...)**
* **표기되는 성적은 모두 열람용 성적입니다.**
* **재수강이 존재할 경우, 정상적으로 작동하지 않을 확률이 높습니다.**
##  사용법
### 1. 프로그램 다운로드
1. 우측의 [releases] 클릭
2. [assets] 클릭
3. [Grade_Management.zip] 다운로드 후 압축 풀기
### 2. 성적 데이터 다운로드
1. [중앙대학교 포탈](https://mportal.cau.ac.kr/main.do) 접속
2. 상단 메뉴 🠆 [강의마당] 🠆 [보관성적조회] 접속
3. 성적내역 우측상단의 [출력] 버튼 클릭
4. 좌측상단 [저장] 버튼 클릭
5. [.xlsx] 확장자 선택 후 [확인] 버튼 클릭
* **[다운로드] 폴더에 다운로드 받으실 경우, 기본 파일명 [noname.xlsx] 를 변경하지 말아주세요.**
* **다른 위치에 저장할 경우, 파일명 변경이 가능합니다. (.xlsx 확장자는 유지)**
### 3. 프로그램 실행
* 프로그램은 우선적으로 Downloads 폴더를 탐색합니다.
* Downloads 폴더에 noname.xlsx 파일이 존재하지 않을 경우, 수동으로 불러오는 것이 가능합니다.
1. [Grade Management.exe] 실행

![fig1](https://user-images.githubusercontent.com/47859342/98086618-601e0180-1ec2-11eb-925e-506981f83379.png)
* 모든 보관성적 확인이 가능합니다.
2. 다음 학기 Simulation

![fig2](https://user-images.githubusercontent.com/47859342/98086621-614f2e80-1ec2-11eb-9cd6-e2d4ed94b102.png)
* y(Y) 입력
* 희망 학점과 평점을 입력하여 다음 학기 simulation이 가능합니다.

![fig3](https://user-images.githubusercontent.com/47859342/98086622-614f2e80-1ec2-11eb-8c9c-3277128fe5fe.png)
* r(R) 입력
* Reset을 통해 현재 시점으로 복귀가 가능합니다.
* Reset을 하지 않을 경우, 성적이 누적됩니다.

![fig4](https://user-images.githubusercontent.com/47859342/98086623-61e7c500-1ec2-11eb-9f6f-269121d5b4ad.png)
* m(M) 입력
* 학점과 평점을 일괄 입력하지않고, 과목별 학점과 평점 개별 입력이 가능합니다.
* 입력 형식: [학점 평점] (대소문자 구별X)
* 동일 평점의 과목들은 통합 가능 (ex. 3학점 A+ 2개일 경우: [6 a+])
3. 성적 수정

![fig5](https://user-images.githubusercontent.com/47859342/98086627-61e7c500-1ec2-11eb-8044-db548a9c36fc.png)
* 엑셀 파일을 통해 성적의 수정 혹은 추가가 가능합니다.
* 성적 추가 시, 기존 성적의 행을 복사 후 끝에 붙여넣어 수정하시면 편리합니다.
* **성적 추가 시, [년도, 학기, 이수구분, 학점, 평점] 항목은 꼭 입력해주셔야 합니다.**
* 과목코드, 과목명, 등급, 비고 항목은 빈칸이여도 무방합니다.
* **수정한 파일은 꼭 다른 이름으로 저장해주세요.(원본 파일 유지)**

![fig6](https://user-images.githubusercontent.com/47859342/98086629-62805b80-1ec2-11eb-9c94-fd4cf036657e.png)
* 프로그램 재시작이 가능합니다.
* 프로그램 재시작 시 수정한 xlsx 파일 선택이 가능합니다.
