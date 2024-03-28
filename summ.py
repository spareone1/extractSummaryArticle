from openai import OpenAI
import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
import urllib3

# SSL 인증서 무시 warning 메시지 미출력
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 엑셀 파일을 읽습니다.
df = pd.read_excel('327.xlsx', header=1, usecols="A,C,D,E,F,I,J,K,L,M,O,P")
# A : 후보자 ID, C : 지역, D : 지역구, E : 후보자, F : 정당

# SQL query가 저장될 리스트
res = []

# 작업 로그가 저장될 리스트
logs = []

# URL만 추출하기 위한 정규표현식
url_pattern = r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+'

# OpenAI API 키
client = OpenAI(
    api_key = 'API 키 입력',
)

# 0 ~ 615개의 행 존재
for i in range(0, df.shape[0]):
  log_txt = ''

  row = df.iloc[i]
  idx = df.index[i]
  
  try:
    log_txt += str(i) + "번째 후보 (" + row.iloc[3] + ") 작업중 (후보자 ID : " + str(int(row.iloc[0])) + ") "
    print(str(i) + "번째 후보 (" + row.iloc[3] + ") 작업중 (후보자 ID : " + str(int(row.iloc[0])) + ") ", end = '')
  except:
    log_txt += "후보자 로드 실패. 건너뜁니다..."
    print("후보자 로드 실패. 건너뜁니다...")

  # 후보자 별 기사가 저장될 리스트
  urls = []

  # 후보자 별 기사 가져오기
  for j in range(5, 13):
    try:
      if pd.isnull(row.iloc[j]):
        break
      url = re.search(url_pattern, row.iloc[j]).group()
      urls.append(url)
    except:
      log_txt += str(j - 4) + "번째 기사 영역 접근 실패. "
      print(str(j - 4) + "번째 기사 영역 접근 실패. ")
      continue

  # 해당 후보자에 대한 기사가 없다면 다음 후보자로 이동
  if len(urls) == 0:
    log_txt += "기사 없음. 건너뜁니다..."
    logs.append(log_txt)
    print("기사 없음. 건너뜁니다...")
    continue

  # 여러 기사의 내용이 담길 문자열
  articles = ''
  
  # URL에 접속하여 기사 내용 가져오기
  for j in urls:
    try:
      response = requests.get(j, verify=False)
      soup = BeautifulSoup(response.text, 'html.parser')
      article_text = soup.get_text(separator=' ', strip=True)
    except:
      continue

    # AI를 이용해 기사 내용만 추출
    try:
      response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {
                    "role": "user",
                    "content": "기사 내용 추출 : " + article_text
                }
            ]
        )
      # 기사 내용 붙이기
      articles += response.choices[0].message.content.strip()
        
    except:
      continue

  # 기사 내용이 없을 경우 다음 후보자로 이동
  if len(articles) == 0:
    log_txt += "기사 URL은 존재하지만 모종의 이유로 하나도 접근하지 못했습니다. 건너뜁니다..."
    logs.append(log_txt)
    print("기사 URL은 존재하지만 모종의 이유로 하나도 접근하지 못했습니다. 건너뜁니다...")
    continue

  # 기사 모음을 AI에게 제공하여 내용 요약
  try:
    response = client.chat.completions.create(
      model="gpt-3.5-turbo",
      messages=[
            {
                "role": "user",
                "content": "기사 내용 요약 : " + articles
            }
        ]
    )

    summary = response.choices[0].message.content.strip()
  except:
    log_txt += "!!!!!!---작업 실패---!!!!!!"
    logs.append(log_txt)
    print("!!!!!!---작업 실패---!!!!!!")
    continue

  # SQL query 추가
  res.append("UPDATE \"vote-for-christ\".election_candidates SET summary='" + summary + "' WHERE id=" + str(int(row.iloc[0])))

  log_txt += "작업 완료"
  logs.append(log_txt)
  print("작업 완료")

# 파일로 저장
with open('327.sql', 'w') as file:
  for i in res:
    file.write(i + ";\n")

# 로그 저장
with open('log.txt', 'w') as file:
  for i in logs:
    file.write(i + ";\n")
  file.write("End")

print("End")
