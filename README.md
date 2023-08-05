# no human lookup 👀❌
귀찮아서 더 귀찮은 일 만들기
<br/><br/>
\-
<br/><br/>

**기존 엑셀 문서에 작성된 테이블/컬럼 정보와 실제 DB의 컬럼 정보가 일치하는지 비교해주는 프로그램** <br/><br/>

#### 🤷‍♀️ 5W1H
>**WHEN** : 2023년 7-8월 어느 더운 여름날들 <br/>
**WHAT** : 인터페이스 명세서 작업 중... <br/>
**WHY** : 차세대 DB 스키마에 계속 변경이 있어 확인을 해줘야 하는데 한 줄씩 눈으로 확인하다가 눈알이 빠질 것 같아서 <br/>
**WHERE** : 집과 사무실에서 <br/>
**WHO** : 내가 <br/>
**HOW** :  chatGPT와 구글의 조언을 참고하여<br/>


<br/><br/>

## 💻 환경 설정 & 실행


### ✨ 개발환경
- Python3 (3.9.9/3.11.4)
- vscode
<br/><br/>


### ⚙ 실행 환경 세팅

#### [DB Connection Properties]

  - create `db.properties` file on root directory and set db properties
    
      ```properties
      # db credentials
      user=root
      passwd=mypassword1234
      host=localhost
      db=myschema
      charset=utf8
      ```

  #### [Target Excel Files]
  - 검사를 원하는 엑셀 파일들을 `/excels` 폴더 안에 넣어주세요
  - 각 엑셀 파일은 정해진 양식을 따라 작성되어 있어야 검사가 가능합니다
    - 차세대 인터페이스 명세서 양식에 따라 작성된 파일만 가능합니다
    - 테이블명이 작성된 열의 `N`번째 컬럼에는 `TABLE` 이 적혀 있어야 합니다
    - 검사할 엑셀 시트는 `세번째`에 위치해야 합니다
  - 검사가 종료되면 파일에 생성된 `비교` 탭에서 색이 칠해진 셀을 확인해주세요
    - 기존 문서와 다른 내용의 셀은 `[빨간색]`
    - 기존 문서에는 없는 내용이 추가된 셀은 `[노란색]`

#### [Install Libraries]
- Open Project with VSCode
- `` Ctrl + ` `` (open terminal)
- `$ python -m venv venv` (create virtual environment named *venv*)
- `$ source venv/Scripts/activate` (activate virtual environment - for Windows)
- `$ pip install -r requirements.txt` (install required dependencies)
  

<br/>

### 🔨 실행 (vscode)
- `$ source venv/Scripts/activate` (이미 활성화되어 있는 경우 생략)
- `$ python main.py [sample.xlsx]` (run)
  - 특정 파일만 검사하고 싶다면 해당 파일의 이름을 파라미터로 넣어주세요
  - 파라미터를 생략하면 `excels` 폴더 내의 모든 파일을 검사합니다



<br/><br/><br/>
<br/><br/><br/>




<br/><br/>
------
<br/><br/><br/><br/>

<details>
  <summary>🔉 TMI</summary>
  
> - 안구 노가다에 비해 가성비가 안나온다... (시간이 더 오래걸림)
>   - 점점 원래 하려던 일에 비해 일이 커짐..
>   - 나름 가성비를 챙겨보기 위해 라이브러리 관련 코드 작성에 chatGPT를 이용했지만 그렇게 유명한 라이브러리들이 아니라 그런지 원하는 코드가 한번에 정확하게 안나왔다
>   - 마찬가지로 에러처리를 할 수록 가성비가 떨어지기 때문에 최소한만 작성
>   - 시간을 더 쓸 수 있다면 해 볼만한 것들
>       - 꼼꼼한 에러처리
>       - 로그 기록 남기기
>       - ~~엑셀 파일 정보를 폴더에서 직접 읽어오는 방식으로 변경 (완료)~~
> - 단순한 프로그램이지만 생각보다 고려해야 할 이슈들이 있었던 것도 가성비 하락 원인
>   - 엑셀 조작을 위해 가장 유명한 라이브러리인 `openpyxl` 을 이용해 구현하다가 DRM 암호화된 파일에 접근할 수 없는 이슈가 있어 `xlwings` 로 변경
>       - 이 프로그램의 존재를 공유했을 때, 다들 어떤 원리로 암호화된 파일 조작이 가능한가를 가장 궁금해하셨다 
>       - 근데 나도 모르겠음 ㅋㅋㅋ 그냥 된다니까.. 그리고 진짜로 되니까... 썼지...
>   - 기본 양식 외에 table/column 내용은 수기작성했다 보니 표의 값만 읽고 table명/column명을 구분하기 어려움
>       - 문서마다 사람이 봤을 때 알아볼 수 있도록 나름의 구분만 해두고 고정된 기준이 없음 (ex. 테이블명 row에 색칠 / 테이블명 오른쪽 칸에 `TABLE`)
>       - 그래서 현재는 고정된 위치에서 완성된 쿼리문을 가져와서 파싱 라이브러리(`sql_metadata`)를 사용해 테이블/컬럼을 뽑고 있는데, 복잡한 쿼리에 대해서는 정확도가 좋지 않은 것 같다.. (왤까)
>       - 테이블/컬럼 정보를 읽어다가 테이블-컬럼 매핑을 만드는 방식으로 수정할까 고민중 (테이블명/컬럼명을 구분할 기준을 정하는 것이 문제)
>       - 이 기회에 문서 양식을 통일시키는 게 좋을지도
>       - **수정 후 추가**: 확인해보니 우리 파트에서 작성한 파일들은 이미 `TABLE`을 잘 써놓아서 걱정 외로 훨씬 쉽고 깔끔하게 수정됐다 

</details>

 