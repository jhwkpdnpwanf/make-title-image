# 인스타 타이틀사진 제작 자동화시키기

사진 동아리를 하면서 인스타를 관리하다 보니 같은 디자인의 타이틀사진을 제작해야 하는 상황이 온다.  
텍스트를 제외하면 전부 같은 이미지길래 자동화가 가능할 것 같아서 만들어보았다.

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

<br><br>

## 실행 방법

![KakaoTalk_20250225_030048835_small](https://github.com/user-attachments/assets/95b71245-556c-4b7b-abce-4dfbc5f56b6e)

**1.** 타이틀사진으로 사용할 사진을 준비한다.  
<br><br>

![KakaoTalk_20250225_025727760_resized](https://github.com/user-attachments/assets/7a9c31a6-d757-40bf-8042-f49879e04a6d)

**2.** 바꾸고 싶은 타이틀사진의 이름을 `"img_title1.png"`으로 변경한다.  
<br><br><br><br>

![KakaoTalk_20250225_025742726_resized](https://github.com/user-attachments/assets/460a5466-c69d-41d5-b076-b9d51eb23121)

**3.** 이름을 변경한 뒤 `image_title` 파일에 넣어준다.  
<br><br>

![KakaoTalk_20250225_025742726_01_resized](https://github.com/user-attachments/assets/fa8c11d2-317e-433e-837b-be309ed58903)

**4.** 실행을 시키면 이러한 창이 뜬다.  
<br><br><br><br>

윗줄과 아랫줄에 쓰고 싶은 텍스트를 입력하면 된다.

그리고 조금 기다리면 ✋
<br><br>

## 실행 결과

![스크린샷_작게](https://github.com/user-attachments/assets/02bcdd90-7b3c-42a8-b4c6-3b9f798ad577)

이러한 타이틀사진이 **PDF로 만들어진다.**
