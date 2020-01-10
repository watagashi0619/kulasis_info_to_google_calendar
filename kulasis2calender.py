from __future__ import print_function
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import pickle
import os.path
import json
import pandas as pd
import datetime
import sys
import os
import time
import locale
import yaml

SCOPES = ['https://www.googleapis.com/auth/calendar']

with open('settings.yaml') as file:
    yml=yaml.safe_load(file)

CALENDER_ID=yml['calender-id']

class GetKULASIS():

    def __init__(self):

        ecs_account = {"ecs-id":yml['ecs-account']['ecs-id'],"password":yml['ecs-account']['password']}


        options = Options()
        options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-gpu')
        options.add_argument('--window-size=1280,1024')
        #self.driver = webdriver.Chrome("/usr/local/bin/chromedriver",options=options)
        self.driver=webdriver.Chrome(options=options)

        #kulasis login
        login_url = "https://www.k.kyoto-u.ac.jp/student/la/top"
        self.driver.get(login_url)
        #self.driver.execute_script("document.getElementById('username').setAttribute('value', %s)" % ecs_account["ecs-id"])
        self.driver.find_element_by_id("username").send_keys(ecs_account["ecs-id"])
        self.driver.find_element_by_id("password").send_keys(ecs_account["password"])
        self.driver.find_element_by_name("_eventId_proceed").click()
    
    def createDF(self):
        #la
        notice_url_la="https://www.k.kyoto-u.ac.jp/student/la/notice/top"
        df_report_la=self.createDFreport(notice_url_la)
        df_cancel_la=self.createDFcancel(notice_url_la)
        exam_url_la="https://www.k.kyoto-u.ac.jp/student/la/exam/schedule/top"
        df_exam_la=self.createDFexam(exam_url_la)
        #t
        notice_url_t="https://www.k.kyoto-u.ac.jp/student/u/t/notice/top"
        df_report_t=self.createDFreport(notice_url_t)
        df_cancel_t=self.createDFcancel(notice_url_t)
        exam_url_t="https://www.k.kyoto-u.ac.jp/student/u/t/exam/schedule/top"
        df_exam_t=self.createDFexam(exam_url_t)


        df_report=pd.concat([df_report_la,df_report_t])
        df_cancel=pd.concat([df_cancel_la,df_cancel_t])
        df_exam=pd.concat([df_exam_la,df_exam_t])

        self.driver.close()
        self.driver.quit()

        return df_report,df_cancel,df_exam

    def createDFreport(self,notice_url):
        self.driver.get(notice_url)
        WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.CLASS_NAME, 'content')))
        df_report=pd.DataFrame(columns=["科目名","提出締切","提出場所","課題等","画像"])
        report_details_link_tag = self.driver.find_elements_by_class_name("content")[6].find_elements_by_tag_name("a")
        report_details_link=[]
        for item in report_details_link_tag:
            report_details_link.append(item.get_attribute("href"))

        for i in range(len(report_details_link)-2):
            self.driver.get(report_details_link[i+1])
            df_row=[]
            subject=self.driver.find_element_by_class_name("standard_list").find_elements_by_tag_name("tr")[2].find_element_by_tag_name("td").text
            df_row.append(subject)
            for item in self.driver.find_element_by_class_name("relaxed_table").find_elements_by_tag_name("tr"):
                if len(item.find_elements_by_class_name("th_normal")):
                    if item.find_element_by_class_name("th_normal").text in ["提出締切","提出場所","課題等"]:
                        df_row.append(item.find_element_by_class_name("odd_normal").text)
                    if item.find_element_by_class_name("th_normal").text=="画像":
                        if len(item.find_elements_by_tag_name("a")):
                            df_row.append(item.find_element_by_tag_name("a").get_attribute("href"))
                        else:
                            df_row.append("")
            temp=pd.Series(df_row,index=df_report.columns)
            df_report=df_report.append(temp,ignore_index=True)
        
        return df_report
    
    def createDFcancel(self,notice_url):
        self.driver.get(notice_url)
        WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.CLASS_NAME, 'content')))
        df_cancel=pd.DataFrame(columns=["科目名","休講日時"])
        cancel_cells = self.driver.find_elements_by_class_name("content")[3].find_elements_by_tag_name("tr")
        for i in range(len(cancel_cells)-4):
            subject=cancel_cells[i+2].find_elements_by_tag_name("td")[1].text
            cancel_class=cancel_cells[i+2].find_elements_by_tag_name("td")[3].text
            if cancel_class[-2]=="1":
                cancel_class=cancel_class[:-4]+" 08:45"
            elif cancel_class[-2]=="2":
                cancel_class=cancel_class[:-4]+" 10:30"
            elif cancel_class[-2]=="3":
                cancel_class=cancel_class[:-4]+" 13:00"
            elif cancel_class[-2]=="4":
                cancel_class=cancel_class[:-4]+" 14:45"
            elif cancel_class[-2]=="5":
                cancel_class=cancel_class[:-4]+" 16:30"
            temp=pd.Series([subject,cancel_class],index=df_cancel.columns)
            df_cancel=df_cancel.append(temp,ignore_index=True)
        
        return df_cancel

    def createDFexam(self,exam_url):
        self.driver.get(exam_url)
        WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.CLASS_NAME, 'content')))
        df_exam=pd.DataFrame(columns=["科目名","試験日","時限","試験開始","試験時間","試験教室","備考欄"])
        if len(self.driver.find_elements_by_class_name("standard_list")):
            exam_tables = self.driver.find_elements_by_class_name("standard_list")
            if exam_url=="https://www.k.kyoto-u.ac.jp/student/la/exam/schedule/top":
                for exam_table in exam_tables:
                    exam_cells=exam_table.find_elements_by_tag_name("tr")[1:-2]
                    #試験日と試験開始
                    exam_date,exam_period=exam_cells[0].find_elements_by_tag_name("td")[1].text.split()
                    exam_period=exam_period[-1]
                    if exam_period=="1":
                        exam_begin="08:45"
                    elif exam_period=="2":
                        exam_begin="10:30"
                    elif exam_period=="3":
                        exam_begin="13:00"
                    elif exam_period=="4":
                        exam_begin="14:45"
                    elif exam_period=="5":
                        exam_begin="16:30"
                    else:
                        exam_begin="18:15"
                    #試験時間
                    exam_time=exam_cells[0].find_elements_by_tag_name("td")[3].text[:-1]
                    #科目名
                    exam_subject=exam_cells[1].find_elements_by_tag_name("td")[1].text
                    #試験教室
                    exam_class=exam_cells[3].find_elements_by_tag_name("td")[1].text
                    #持ち込みできるもの
                    exam_bring_in=exam_cells[4].find_elements_by_tag_name("td")[1].text
                    #学生への注意事項
                    exam_notice=exam_cells[5].find_elements_by_tag_name("td")[1].text
                    if exam_notice=="なし":
                        exam_notice=""
                    else:
                        exam_notice="\n注意事項：%s" % exam_notice
                    exam_remarks=exam_bring_in+exam_notice
                    temp=pd.Series([exam_subject,exam_date,exam_period,exam_begin,exam_time,exam_class,exam_remarks],index=df_exam.columns)
                    df_exam=df_exam.append(temp,ignore_index=True)
            else:
                for exam_table in exam_tables:
                    exam_cells=exam_table.find_elements_by_tag_name("tr")[1:-2]
                    #試験日と試験開始
                    exam_date,exam_period=exam_cells[0].find_elements_by_tag_name("td")[1].text.split()
                    exam_period=exam_period[-1]
                    if exam_period=="1":
                        exam_begin="08:45"
                    elif exam_period=="2":
                        exam_begin="10:30"
                    elif exam_period=="3":
                        exam_begin="13:00"
                    elif exam_period=="4":
                        exam_begin="14:45"
                    elif exam_period=="5":
                        exam_begin="16:30"
                    else:
                        exam_begin="18:15"
                    #試験時間
                    exam_time=exam_cells[1].find_elements_by_tag_name("td")[1].text[:-1]
                    #科目名
                    exam_subject=exam_cells[2].find_elements_by_tag_name("td")[1].text
                    #試験教室
                    exam_class=exam_cells[4].find_elements_by_tag_name("td")[1].text
                    #持ち込みできるもの
                    exam_bring_in=exam_cells[5].find_elements_by_tag_name("td")[1].text
                    #学生への注意事項
                    exam_notice=exam_cells[6].find_elements_by_tag_name("td")[1].text
                    if exam_notice=="なし":
                        exam_notice=""
                    else:
                        exam_notice="\n注意事項：%s" % exam_notice
                    exam_remarks=exam_bring_in+exam_notice
                    temp=pd.Series([exam_subject,exam_date,exam_period,exam_begin,exam_time,exam_class,exam_remarks],index=df_exam.columns)
                    df_exam=df_exam.append(temp,ignore_index=True)

        return df_exam
        
def main():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    get_kulasis=GetKULASIS()
    df_report,df_cancel,df_exam=get_kulasis.createDF()

    locale.setlocale(locale.LC_TIME, 'ja_JP.UTF-8')

    # Call the Calendar API

    nowdt = datetime.datetime.utcnow()
    now = nowdt.isoformat() + 'Z' # 'Z' indicates UTC time
    enddt = nowdt+datetime.timedelta(days=120)
    end = enddt.isoformat() + 'Z'
    events_result = service.events().list(calendarId=CALENDER_ID,
                                        timeMin=now,
                                        timeMax=end,
                                        singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])

    if events:
        recorded_start=[]
        recorded_summary=[]
        print('Be scheduled:')
        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            start = datetime.datetime.strptime(start[:-6], '%Y-%m-%dT%H:%M:%S')
            recorded_start.append(start)
            recorded_summary.append(event['summary'])
            print(start, event['summary'])
    
    #休講
    for i in range(df_cancel.shape[0]):
        t_start=datetime.datetime.strptime(df_cancel.iloc[i,1],"%Y/%m/%d %H:%M")
        t_end=t_start+datetime.timedelta(minutes=90)
        if (not df_cancel.iloc[i,0] in recorded_summary) and (not t_start in recorded_start):
            event = {
                'summary': "（休講）"+df_cancel.iloc[i,0],
                'location': '',
                'description': '',
                'start': {
                    'dateTime': t_start.isoformat(),
                    'timeZone': 'Japan',
                },
                'end': {
                    'dateTime': t_end.isoformat(),
                    'timeZone': 'Japan',
                },
            }

            event = service.events().insert(calendarId=CALENDER_ID,body=event).execute()
            print('New canceled class (%d/%d):%s %s' % (i+1,df_cancel.shape[0],df_cancel.iloc[i,1],df_cancel.iloc[i,0]))
    
    #レポート
    for i in range(df_report.shape[0]):
        t_start=datetime.datetime.strptime(df_report.iloc[i,1],"%Y/%m/%d(%a) %H:%M")
        if (not df_report.iloc[i,0] in recorded_summary) and (not t_start in recorded_start):
            event = {
                'summary': "（レポート締切）"+df_report.iloc[i,0],
                'location': df_report.iloc[i,2],
                'description': df_report.iloc[i,3]+"\n"+df_report.iloc[i,4],
                'start': {
                    'dateTime': t_start.isoformat(),
                    'timeZone': 'Japan',
                },
                'end': {
                    'dateTime': t_start.isoformat(),
                    'timeZone': 'Japan',
                },
            }

            event = service.events().insert(calendarId=CALENDER_ID,body=event).execute()
            print('New report (%d/%d):%s %s' % (i+1,df_report.shape[0],df_report.iloc[i,1],df_report.iloc[i,0]))

    #試験時間割
    for i in range(df_exam.shape[0]):
        t_start=datetime.datetime.strptime(df_exam.iloc[i,1]+df_exam.iloc[i,3],"%Y/%m/%d%H:%M")
        t_end=t_start+datetime.timedelta(minutes=int(df_exam.iloc[i,4]))
        if (not df_exam.iloc[i,0] in recorded_summary) and (not t_start in recorded_start):
            event = {
                'summary': "（"+df_exam.iloc[i,2]+"限・期末試験）"+df_exam.iloc[i,0],
                'location': df_exam.iloc[i,5],
                'description': df_exam.iloc[i,6],
                'start': {
                    'dateTime': t_start.isoformat(),
                    'timeZone': 'Japan',
                },
                'end': {
                    'dateTime': t_end.isoformat(),
                    'timeZone': 'Japan',
                },
            }

            event = service.events().insert(calendarId=CALENDER_ID,body=event).execute()
            print('New exam timetable :%s from %s until %s' % (df_exam.iloc[i,0],str(t_start),str(t_end)))

if __name__ == '__main__':
    main()
