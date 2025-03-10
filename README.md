# 인스타 타이틀사진 제작 자동화시키기

사진 동아리를 하면서 인스타를 관리하다 보니 같은 디자인의 타이틀사진을 제작해야 하는 상황이 온다.  
텍스트를 제외하면 전부 같은 이미지길래 자동화가 가능할 것 같아서 만들어보았다.
<br><br>

## 프로젝트 구조
```
📂 프로젝트 폴더
 ┣ 📂 assets        # 로고 이미지 보관
 ┣ 📂 img_title     # 작업할 이미지 파일 (한 개만 존재)
 ┣ 📂 output        # 결과물 저장
 ┣ 📂 src           # 파이썬 소스 코드
 ┗ 📄 README.md      # 프로젝트 설명
```
<br>

## 설치 패키지
```bash
pip install python-pptx pillow pdf2image pymupdf comtypes
```

| 패키지명          | 설명                                                              |
|-------------------|-------------------------------------------------------------------|
| **python-pptx**   | PowerPoint(.pptx) 파일을 생성 및 수정                             |
| **pillow**        | 이미지 처리 라이브러리 (PIL)                                      |
| **pdf2image**     | PDF를 이미지로 변환                                               |
| **pymupdf (fitz)**| PDF 문서 조작 및 이미지 추출                                      |
| **comtypes**      | Windows 환경에서 COM 객체를 사용하여 PowerPoint 자동화            |

pdf2image를 사용하기 위해선 **poppler**가 있어야 한다.  
아래 링크에서 설치하자.

[Poppler 다운로드](https://github.com/oschwartz10612/poppler-windows/releases/)

bin 경로로 환경변수에 추가하는 것도 잊어선 안 된다.  
<br>

## 실행 방법
![resized_img_title1](https://github.com/user-attachments/assets/213b9775-cbba-4f5c-8517-dc72c0d1ef82)

**1.** 타이틀사진으로 사용할 사진을 준비한다.  
<br><br><br>
![image](https://github.com/user-attachments/assets/6c2f7a4f-729f-4084-8b22-df2a198c0b63)
![image](https://github.com/user-attachments/assets/da32f888-19b8-479e-a10d-d15db448b3bf)


**2.** 바꾸고 싶은 타이틀사진을 "img_title"파일에 넣는다.  
<br><br><br>
![image](https://github.com/user-attachments/assets/6a55c7ed-5ac5-4dcc-bfb2-fe8d6ab78c85)

**3.** 추가해야할 로고가 있다면 logo파일에 추가해준다.  
<br><br>
![image](https://github.com/user-attachments/assets/9258d9a6-71fd-49f5-b156-a60a9cfe1227)

**4.** 실행을 시키면 이러한 창이 뜬다.  
<br><br><br>

윗줄과 아랫줄에 쓰고 싶은 텍스트를 입력하면 된다.

그리고 조금 기다리면 ✋
<br><br>


## 실행 결과
![resized_photo](https://github.com/user-attachments/assets/7a8c03d4-bd6c-4e26-a39c-dbf6af852fc9)

이러한 타이틀사진이 **PDF로 만들어진다.**

