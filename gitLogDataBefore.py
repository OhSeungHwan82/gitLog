import subprocess
import shutil
import os
import requests
import xml.etree.ElementTree as ET
import openpyxl

def run_git_command_prod(command):
    process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True, cwd=r"D:\IIMS-PULS2")
    output, error = process.communicate()
    return output, error

get_commit_hash = ""
get_commit_msg = ""

#운영브랜치로 체크아웃
output, error = run_git_command_prod("git checkout master")
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
output, error = run_git_command_prod('git log master --since="2023-01-01" --until="2023-06-30" --format="%h,%ad,%s" --date=format:"%Y-%m-%d %H:%M:%S"')
if output:
    print("insert_commit_hash_list : ",output.decode("utf-8"))
    insert_commit_hash_list = output.decode("utf-8")
if error:
    print(error)
    
if insert_commit_hash_list!="":    
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(['해시코드','커밋날짜','설명'])
    #해쉬값은 여러개가 나올 수 있으니까 반복문
    lines = insert_commit_hash_list.split("\n")
    for line in lines:
        insert_data = line.split(",")
        if len(insert_data)>1:
            print("insert_commit_hash : ",insert_data[0])
            print("insert_commit_date : ",insert_data[1])
            print("insert_commit_msg : ",insert_data[2])
            insert_commit_hash = insert_data[0]
            insert_commit_date = insert_data[1]
            insert_commit_msg = insert_data[2]

            
            
            sheet.append([insert_commit_hash, insert_commit_date, insert_commit_msg])

    workbook.save("bfoutput20231205.xlsx")       

    

    
