# EBookToPDF
사용방법 동영상 : 
https://youtu.be/srk8iUTtb-4?si=sWExCrPJadY47SkX

EBook 를 캡처 받아 PDF 파일로 만들기
이전 버전 : 캡처 받는 프로그램 + PDF 만드는 프로그램

New : eBook TO PDF 이전 버전 두개의 프로그램을 하나로 합침

소스파일 :
실행파일:

py소스 파일로 EXE 파일 만들기
pyinstaller --onefile --windowed --name EBookToPDF_by_muttul EBookToPDF_04.py

NVIDIA 비디오 카드를 사용하면 Alt + F1 키로 캡쳐 받아짐.
** 제일 먼저 NVIDIA 프로그램을 실행해야 합니다. **
사진 설명을 입력하세요.
사진 설명을 입력하세요.
사진 설명을 입력하세요.


* 사용 방법 - NVIDIA 비디오 카드를 사용하면 Alt + F1 키로 캡쳐 받아짐.
(다른 프로그램은 모두 막혔는데 이것은 가능함. 2024.09.15. 기준)

*** eBook 캡처 받기 ***
1. 반복 횟수 설정 : 캡쳐 받을 페이지 수 입력
2. 대기 시간 설정 : 캡쳐 받을 속도
(1초 정도로 하면 됨-eBook 넘길대 화면 로딩 시간 확인)
3. 캡처할 창 선택 : [창 목록 새로고침] 버튼 클릭하여 목록 생성 후 원하는 창 선택
(여기서 eBook 창 선택)
4. 마우스 클릭 위치 설정 : eBook 다음 페이지 넘기는 버튼 위치 설정
(마우스 포인터를 버튼 위에 3초 동안 두면 자동으로 설정 됨)
5. [매크로 시작] 버튼 클릭하고 eBook 클릭하면 자동으로 한 페이지씩 넘기며 캡쳐받음

캡처파일 저장 위치 : C:\Users\user\Videos\Desktop


*** 캡처 받기 이미지를 원하는 부분 자르기 및 PDF 파일 생성 ***
캡처파일 저장 위치 : C:\Users\user\Videos\Desktop

1. [폴더 선택] 버튼 클릭 : 캡처 받은 파일이 있는 곳 선택
( 캡처파일 저장 위치 : C:\Users\user\Videos\Desktop )
- 자동으로 폴더 구조를 다음과 같이 변경함.
사진 설명을 입력하세요.
2. 이미지 미리보기 영역에 image 폴더의 첫 파일을 화면에 표시됨
3. 창을 크게해서 자르고자 하는 부분을 마우스 드래그하여 빨간색 사각형 표시
사진 설명을 입력하세요.
4. [PDF 생성] 버튼 클릭 : Cropper 폴더에 자른 이미지와 PDF 파일 생성

끝.
