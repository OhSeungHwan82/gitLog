import cx_Oracle
import subprocess
import shutil
import os
import requests
import xml.etree.ElementTree as ET
import openpyxl

cx_Oracle.init_oracle_client(lib_dir=r"C:\oracle_client\client\19c\client\bin")
def run_git_command_prod(command):
    process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True, cwd=r"D:\Prod-IIMS-PLUS")
    output, error = process.communicate()
    return output, error

host = '10.20.20.201'
port = 1521
#sid = 'OLB19DB'
sid = 'PDB_ONE.INCAR.CO.KR'
user_name = ''
passwd = ''

get_commit_hash = ""
get_commit_msg = ""

#운영브랜치로 체크아웃
output, error = run_git_command_prod("git checkout main")
if output:
    print(output.decode("utf-8"))
if error:
    print(error)
#운영브랜치 pull
output, error = run_git_command_prod("git pull")
if output:
    print(output.decode("utf-8"))
if error:
    print(error)

insert_commit_hash =""
insert_commit_date =""
insert_commit_msg =""
insert_commit_hash_list = ""
output, error = run_git_command_prod('git log --since="2023-07-01" --format="%h,%ad,%s" --date=format:"%Y-%m-%d %H:%M:%S"')
if output:
    print("insert_commit_hash_list : ",output.decode("utf-8"))
    insert_commit_hash_list = output.decode("utf-8")
if error:
    print(error)
    
if insert_commit_hash_list!="":    
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(['해시코드','커밋날짜','배포날짜','접수번호','현재상태','인수승인 날짜','인수승인 사번','인수승인 이름','이행승인날짜','이행승인 사번','이행승인 이름','이행완료 날짜','이행완료 사번','이행완료 이름'])
    #해쉬값은 여러개가 나올 수 있으니까 반복문
    lines = insert_commit_hash_list.split("\n")
    conn = cx_Oracle.connect(f'{user_name}/{passwd}@{host}:{port}/{sid}')
    for line in lines:
        insert_data = line.split(",")
        if len(insert_data)>1:
            print("insert_commit_hash : ",insert_data[0])
            print("insert_commit_date : ",insert_data[1])
            print("insert_commit_msg : ",insert_data[2])
            insert_commit_hash = insert_data[0]
            insert_commit_date = insert_data[1]
            insert_commit_msg = insert_data[2]

            
            #insert_commit_hash 커밋 해시값 , insert_commit_msg 커밋 메시지 = 접수번호
            qry = f"""
select *
  from (
       select a.hash_code
            , to_char(a.create_date,'yyyy-mm-dd hh24:mi:ss') deploy_date
            , b.jubsu_no
            , case when b.status_cd='1' then '요청등록'
                   when b.status_cd='2' then '접수'
                   when b.status_cd='3' then '접수확정'
                   when b.status_cd='12' then '개발접수'
                   when b.status_cd='4' then '개발승인'
                   when b.status_cd='5' then '진행'
                   when b.status_cd='9' then '테스트'
                   when b.status_cd='13' then '테스트완료'
                  when b.status_cd='6' then '인수승인'
                   when b.status_cd='10' then '이행승인'
                   when b.status_cd='11' then '이행완료'
                   when b.status_cd='7' then '완료'
                   when b.status_cd='99' then '반려'
               else '' end current_status 
            , to_char(c.create_date,'yyyy-mm-dd hh24:mi:ss') status_change_date
            , c.create_by status_change_by
            , (select x.name_kor from sawon x where x.sawon_cd = c.create_by) status_change_by_nm
            , c.status_cd
         from git_inforequest_link a
            , info_request b
            , info_confirm_list c
        where a.hash_code = :hash_code
          and a.info_request_pk = b.pk
          and b.pk = c.arc_pk
          and c.status_cd in ('6','10','11')
          and c.use_yb ='1'
     )
 pivot (max(status_cd) as status_cd, max(status_change_date) as status_change_date, max(status_change_by) as status_change_by, max(status_change_by_nm) as status_change_by_nm
   for status_cd in ('6' as s6, '10' as s10, '11' as s11))
                    """
            bind_arr={"hash_code":insert_commit_hash}
            cursor = conn.cursor()
            cursor.execute(qry, bind_arr)
            info_request_pk = ""
            results = cursor.fetchall()

            for row in results:
                hash_code = row[0]
                deploy_date = row[1]
                jubsu_no = row[2]
                current_status = row[3]
                #s6_status_cd = row[4]
                s6_status_change_date = row[5]
                s6_status_change_by = row[6]
                s6_status_cange_by_nm = row[7]
                #s10_status_cd = row[8]
                s10_status_change_date = row[9]
                s10_status_change_by = row[10]
                s10_status_cange_by_nm = row[11]
                #s11_status_cd = row[12]
                s11_status_change_date = row[13]
                s11_status_change_by = row[14]
                s11_status_cange_by_nm = row[15]
                sheet.append([hash_code, insert_commit_date, deploy_date, jubsu_no, current_status
                , s6_status_change_date, s6_status_change_by, s6_status_cange_by_nm
                , s10_status_change_date, s10_status_change_by, s10_status_cange_by_nm
                , s11_status_change_date, s11_status_change_by, s11_status_cange_by_nm])

    workbook.save("output20231204.xlsx")       
    conn.close() 

    

    
