
from openpyxl import load_workbook
from tkinter import ttk
from tkinter import *
import tkinter as tk 
import win32com.client
import pandas as pd
import tkinter.messagebox as msgbox
import os


# 필요 : class 사용하기, 업무 추가 후 안내 메시지 보내고 새로고침
# 담당팀, 담당자 주소록 만들기, 편집, 삭제 기능
# 스크롤바 추가
# class load_db:
    # def __init__(self,wb):
    #     self.wb = wb

#     def load_wb(self,path):
#         wb = load_workbook(path)
#         # print(path)
#         # self.wb = wb    
#         return wb

# class load_excel:
#     def __init__(self):                    
#         wb = load_workbook("C:/Python/Code/ToDoList/ToDoList_Form.xlsx")
#         db_ws = wb['DB']
#         # values_only=True를 해야 위치값이 아닌 실제 데이터 값이 도출됨
#         db_header = db_ws.iter_rows(min_row=1, max_row=1, max_col=9, values_only=True)
#         db_datas = db_ws.iter_rows(min_row=2, max_col=9, values_only=True)

#         db_header = [r for r in db_header]
#         db_datas = [r for r in db_datas]
#         wb.close()
#         return db_datas
# def refresh():
#     win.destroy()
#     os.popen("C:/Python/Code/ToDoList/ToDoList_v1.py")

def add_task():
    try:
        # 업무 레벨에 따라 No링 잘 하기
        # no 정렬 방식 활용 해보기
        add_tk = Tk()
        add_tk.title("업무 추가")

        def add_db():
            wb = load_workbook("C:/Python/Code/ToDoList/ToDoList_Form.xlsx")
            db_ws = wb['DB']
            # values_only=True를 해야 위치값이 아닌 실제 데이터 값이 도출됨
            db_header = db_ws.iter_rows(min_row=1, max_row=1, max_col=9, values_only=True)
            db_datas = db_ws.iter_rows(min_row=2, max_col=9, values_only=True)
            db_header = [r for r in db_header]
            db_datas = [r for r in db_datas]
            
            # print(db_datas[db_ws.max_row-2][0])
            lv1_cnt = int(db_datas[db_ws.max_row-2][0][0:2])
            # print(lv1_cnt)
            lv1_next = lv1_cnt+1
            lv1_next = format(lv1_next,'02')
            # print(format(lv1_next,'02'))
            # print(db_ws.max_row)
            r_add = db_ws.max_row + 1
            db_ws.cell(row=r_add, column=1).value = lv1_next + "-00-00"
            # 값이 없을 때 공백 하나 넣기. 값이 없으면 편집 함수 중 entry 값을 불러올 때 에러가 남
            if len(entry_task_name.get()) == 0 :
                db_ws.cell(row=r_add, column=2).value = " "    
            else : 
                db_ws.cell(row=r_add, column=2).value = entry_task_name.get()
            if len(entry_team_name.get()) == 0 :
                db_ws.cell(row=r_add, column=3).value = " "
            else :    
                db_ws.cell(row=r_add, column=3).value = entry_team_name.get()
            if len(entry_person_name.get()) == 0:
                db_ws.cell(row=r_add, column=4).value = " "
            else :    
                db_ws.cell(row=r_add, column=4).value = entry_person_name.get()
            db_ws.cell(row=r_add, column=5).value = cmb_situation.get()
            if len(entry_date.get()) == 0:
                db_ws.cell(row=r_add, column=6).value = " "    
            else :
                db_ws.cell(row=r_add, column=6).value = entry_date.get()
            if len(entry_time.get()) == 0:
                db_ws.cell(row=r_add, column=7).value = " "
            else :    
                db_ws.cell(row=r_add, column=7).value = entry_time.get()
            if len(entry_note.get()) == 0:
                db_ws.cell(row=r_add, column=8).value = " "
            else :     
                db_ws.cell(row=r_add, column=8).value = entry_note.get()
            db_ws.cell(row=r_add, column=9).value = 1

            wb.save("C:/Python/Code/ToDoList/ToDoList_Form.xlsx")

            msgbox.showinfo("알림", "신규 업무 추가가 완료되었습니다.")

            add_tk.withdraw()
            sort()
            load_task()


        # 업무명/내용
        frame_task_name = Frame(add_tk)
        frame_task_name.pack(fill="x")

        lbl_task_name = Label(frame_task_name, text="업무명/내용 - 레벨1")
        lbl_task_name.pack(side="left")

        entry_task_name = Entry(frame_task_name)
        entry_task_name.pack(side="right")

        # 담당팀
        frame_team_name = Frame(add_tk)
        frame_team_name.pack(fill="x")

        lbl_team_name = Label(frame_team_name, text="담당팀")
        lbl_team_name.pack(side="left")

        entry_team_name = Entry(frame_team_name)
        entry_team_name.pack(side="right")

        # 담당자
        frame_person_name = Frame(add_tk)
        frame_person_name.pack(fill="x")

        lbl_person_name = Label(frame_person_name, text="담당자")
        lbl_person_name.pack(side="left")

        entry_person_name = Entry(frame_person_name)
        entry_person_name.pack(side="right")

        # 상황
        frame_situation = Frame(add_tk)
        frame_situation.pack(fill="x")

        lbl_situation = Label(frame_situation, text="상황")
        lbl_situation.pack(side="left")

        opt_situation = ["진행","완료","검토","취소","중단","지연"]
        cmb_situation = ttk.Combobox(frame_situation,state="readonly",
        values=opt_situation)
        cmb_situation.current(0)
        cmb_situation.pack(side="right")

        # 완료날짜
        frame_date = Frame(add_tk)
        frame_date.pack(fill="x")

        lbl_date = Label(frame_date, text="완료날짜")
        lbl_date.pack(side="left")

        entry_date = Entry(frame_date)
        entry_date.pack(side="right")
        # entry_date.insert(END, "YY-MM-DD")

        # 완료시간
        frame_time = Frame(add_tk)
        frame_time.pack(fill="x")

        lbl_time = Label(frame_time, text="완료시간")
        lbl_time.pack(side="left")

        entry_time = Entry(frame_time)
        entry_time.pack(side="right")
        # entry_time.insert(END, "HH:MM")

        
        # 비고
        frame_note = Frame(add_tk)
        frame_note.pack(fill="x")

        lbl_note = Label(frame_note, text="비고")
        lbl_note.pack(side="left")

        entry_note = Entry(frame_note)
        entry_note.pack(side="right")

        # fucntion2 frame
        frame_func2 = Frame(add_tk)
        frame_func2.pack(fill="x")

        # add db
        btn_add_db = Button(frame_func2, text="추가",command=add_db)
        btn_add_db.pack(side="left")

        # cancel button
        # quit 함수를 쓰면 전체 프로그램이 종료되어서 withdraw 함수 사용함
        btn_cancel = Button(frame_func2, text="취소", command=add_tk.withdraw)
        btn_cancel.pack(side="right")


        add_tk.mainloop()
    

    except Exception as err:
        msgbox.showerror("에러", err)

def get_text():
    
    selected_item = tree.item(tree.selection())
    # selected_iid = tree.focus()
    # print(selected_item)
    # print(selected_item.get("text"))
    return selected_item.get("text")

def get_iid():
    selected_iid = tree.focus()
    # print(selected_iid)
    # print(selected_item.get("text"))
    return selected_iid

# def wordfinder(SearchString):
#     for i in range(1, db_ws.max_row+1):
#         if SearchString == db_ws.cell(i,1).value:
#             print(i)
#     return i


def add_task2():

    # print(get_iid()[6:8])
    # Lv3 업무 선택 후 하위 업무를 선택하면 에러 메시지 구현
    if get_iid()[6:8] != '00':
        msgbox.showerror("에러-하위 업무 추가", "최하위단 업무를 선택하셨습니다.\n\n 상위 업무를 선택해주세요.")
        return

    try:
        add_tk = Tk()
        add_tk.title("하위 업무 추가")


        def add_db():
            wb = load_workbook("C:/Python/Code/ToDoList/ToDoList_Form.xlsx")
            db_ws = wb['DB']
            # values_only=True를 해야 위치값이 아닌 실제 데이터 값이 도출됨
            db_header = db_ws.iter_rows(min_row=1, max_row=1, max_col=9, values_only=True)
            db_datas = db_ws.iter_rows(min_row=2, max_col=9, values_only=True)
            db_header = [r for r in db_header]
            db_datas = [r for r in db_datas]
            
            # 선택 항목의 iid값 가져오기
            for i in range(1, db_ws.max_row+1):
                if db_ws.cell(i,1).value == get_iid():
                    # print(i)
                    r_selected = i
            # print(db_datas[r_selected-2][0][3:8])
            # print(db_datas[r_selected-2][0])        
            
            lv1_cnt = 0
            lv2_cnt = 0
            lv3_cnt = 0
            for x in range(db_ws.max_row-2):
                # print(db_datas[x][0][0:2])
                # 선택 값이 1레벨 -> 하위 업무는 2레벨에 업무 추가
                if db_datas[r_selected-2][0][3:8] == "00-00":
                    if db_datas[x][0][0:2] == db_datas[r_selected-2][0][0:2]:
                        lv1_cnt = int(db_datas[x][0][0:2])
                        lv2_cnt = int(db_datas[x][0][3:5])

                # 선택 값이 2레벨 -> 하위 업무는 3레벨에 업무 추가
                elif db_datas[r_selected-2][0][6:8] == "00":
                    if db_datas[x][0][0:5] == db_datas[r_selected-2][0][0:5]:
                        lv1_cnt = int(db_datas[x][0][0:2])
                        lv2_cnt = int(db_datas[x][0][3:5])
                        lv3_cnt = int(db_datas[x][0][6:8])

            # 다음 iid 값 넣기        
            lv1_cnt = str(format(lv1_cnt,'02'))
            
            lv2_next = lv2_cnt+1
            lv2_next = format(lv2_cnt+1,'02')
            lv2_cnt = str(format(lv2_cnt,'02'))
            
            lv3_next = lv3_cnt+1
            lv3_next = format(lv3_cnt+1,'02')
            lv3_cnt = str(format(lv3_cnt,'02'))
            
            
            # print(lv1_cnt + "-" + lv2_next + "-00")
            # db_datas는 0부터 시작하고, 제목행이 빠져서 -2를 해야함
            # 선택 값이 1레벨 -> 하위 업무는 2레벨에 업무 추가
            if db_datas[r_selected-2][0][3:8] == "00-00":
                r_add = db_ws.max_row + 1
                db_ws.cell(row=r_add, column=1).value = lv1_cnt + "-" + lv2_next + "-00"
                db_ws.cell(row=r_add, column=2).value = entry_task_name.get()
                db_ws.cell(row=r_add, column=3).value = entry_team_name.get()
                db_ws.cell(row=r_add, column=4).value = entry_person_name.get()
                db_ws.cell(row=r_add, column=5).value = cmb_situation.get()
                db_ws.cell(row=r_add, column=6).value = entry_date.get()
                db_ws.cell(row=r_add, column=7).value = entry_time.get()
                db_ws.cell(row=r_add, column=8).value = entry_note.get()
                db_ws.cell(row=r_add, column=9).value = 2

                wb.save("C:/Python/Code/ToDoList/ToDoList_Form.xlsx")
            
            # 선택 값이 2레벨 -> 하위 업무는 3레벨에 업무 추가
            elif db_datas[r_selected-2][0][6:8] == "00":
                r_add = db_ws.max_row + 1
                db_ws.cell(row=r_add, column=1).value = lv1_cnt + "-" + lv2_cnt + "-" + lv3_next
                db_ws.cell(row=r_add, column=2).value = entry_task_name.get()
                db_ws.cell(row=r_add, column=3).value = entry_team_name.get()
                db_ws.cell(row=r_add, column=4).value = entry_person_name.get()
                db_ws.cell(row=r_add, column=5).value = cmb_situation.get()
                db_ws.cell(row=r_add, column=6).value = entry_date.get()
                db_ws.cell(row=r_add, column=7).value = entry_time.get()
                db_ws.cell(row=r_add, column=8).value = entry_note.get()
                db_ws.cell(row=r_add, column=9).value = 3

                wb.save("C:/Python/Code/ToDoList/ToDoList_Form.xlsx")

            msgbox.showinfo("알림", "하위 업무 추가가 완료되었습니다.")

            add_tk.withdraw()
            sort()
            load_task()

        # print(get_text)

        tree.bind('<ButtonRelease-1>', get_text)
        tree.bind('<ButtonRelease-1>', get_iid)

        # 상위 업무명/내용
        frame_upper_task_name = Frame(add_tk)
        frame_upper_task_name.pack(fill="x")

        lbl_upper_task = Label(frame_upper_task_name, text="상위 업무명/내용")
        lbl_upper_task.pack(side="left")
        
        lbl_upper_task_name = Label(frame_upper_task_name, text=get_text())
        lbl_upper_task_name.pack(side="right")
        
        # 업무명/내용
        frame_task_name = Frame(add_tk)
        frame_task_name.pack(fill="x")
        
        lbl_task_name = Label(frame_task_name, text="하위 추가 업무명/내용")
        lbl_task_name.pack(side="left")

        entry_task_name = Entry(frame_task_name)
        entry_task_name.pack(side="right")

        # 담당팀
        frame_team_name = Frame(add_tk)
        frame_team_name.pack(fill="x")

        lbl_team_name = Label(frame_team_name, text="담당팀")
        lbl_team_name.pack(side="left")

        entry_team_name = Entry(frame_team_name)
        entry_team_name.pack(side="right")

        # 담당자
        frame_person_name = Frame(add_tk)
        frame_person_name.pack(fill="x")

        lbl_person_name = Label(frame_person_name, text="담당자")
        lbl_person_name.pack(side="left")

        entry_person_name = Entry(frame_person_name)
        entry_person_name.pack(side="right")

        # 상황
        frame_situation = Frame(add_tk)
        frame_situation.pack(fill="x")

        lbl_situation = Label(frame_situation, text="상황")
        lbl_situation.pack(side="left")

        opt_situation = ["진행","완료","검토","취소","중단","지연"]
        cmb_situation = ttk.Combobox(frame_situation,state="readonly",
        values=opt_situation)
        cmb_situation.current(0)
        cmb_situation.pack(side="right")

        # 완료날짜
        frame_date = Frame(add_tk)
        frame_date.pack(fill="x")

        lbl_date = Label(frame_date, text="완료날짜")
        lbl_date.pack(side="left")

        entry_date = Entry(frame_date)
        entry_date.pack(side="right")
        # entry_date.insert(END, "YY-MM-DD")

        # 완료시간
        frame_time = Frame(add_tk)
        frame_time.pack(fill="x")

        lbl_time = Label(frame_time, text="완료시간")
        lbl_time.pack(side="left")

        entry_time = Entry(frame_time)
        entry_time.pack(side="right")
        # entry_time.insert(END, "HH:MM")

        
        # 비고
        frame_note = Frame(add_tk)
        frame_note.pack(fill="x")

        lbl_note = Label(frame_note, text="비고")
        lbl_note.pack(side="left")

        entry_note = Entry(frame_note)
        entry_note.pack(side="right")

        # fucntion2 frame
        frame_func2 = Frame(add_tk)
        frame_func2.pack(fill="x")

        # add db
        btn_add_db = Button(frame_func2, text="추가",command=add_db)
        btn_add_db.pack(side="left")

        # cancel button
        # quit 함수를 쓰면 전체 프로그램이 종료되어서 withdraw 함수 사용함
        btn_cancel = Button(frame_func2, text="취소", command=add_tk.withdraw)
        btn_cancel.pack(side="right")

        add_tk.mainloop()

    except Exception as err:
        msgbox.showerror("에러", err)


def del_task():
    try:
        # print(get_iid())
        wb = load_workbook("C:/Python/Code/ToDoList/ToDoList_Form.xlsx")
        db_ws = wb['DB']
        # values_only=True를 해야 위치값이 아닌 실제 데이터 값이 도출됨
        db_header = db_ws.iter_rows(min_row=1, max_row=1, max_col=9, values_only=True)
        db_datas = db_ws.iter_rows(min_row=2, max_col=9, values_only=True)
        db_header = [r for r in db_header]
        db_datas = [r for r in db_datas]
        
        # 선택 항목의 iid값 가져오기
        for i in range(1, db_ws.max_row+1):
            if db_ws.cell(i,1).value == get_iid():
                # print(i)
                r_selected = i
        
        # db_ws.delete_rows(r_selected)

        # 선택 값이 1레벨이면 하위 2,3레벨도 삭제
        if db_datas[r_selected-2][0][3:8] == "00-00":
            for x in reversed(range(db_ws.max_row-1)):
                if db_datas[x][0][0:2] == db_datas[r_selected-2][0][0:2]:
                    db_ws.delete_rows(x+2)
        # 선택 값이 2레벨이면 3레벨도 삭제
        elif db_datas[r_selected-2][0][6:8] == "00":
            for x in reversed(range(db_ws.max_row-1)):
                if db_datas[x][0][0:5] == db_datas[r_selected-2][0][0:5]:
                    db_ws.delete_rows(x+2)
        # 선택 값이 3레벨이면 해당 행만 삭제
        else:
            db_ws.delete_rows(r_selected)

        wb.save("C:/Python/Code/ToDoList/ToDoList_Form.xlsx")
        
        msgbox.showinfo("알림", "선택한 업무 삭제 되었습니다.")
        sort()
        load_task()

    except Exception as err:
        msgbox.showerror("에러", err)

def edit_task():
    try:
        add_tk = Tk()
        add_tk.title("수정")


        def apply():
            # print(entry_task_name.get())
                
            wb = load_workbook("C:/Python/Code/ToDoList/ToDoList_Form.xlsx")
            db_ws = wb['DB']    
            # values_only=True를 해야 위치값이 아닌 실제 데이터 값이 도출됨
            # db_header = db_ws.iter_rows(min_row=1, max_row=1, max_col=9, values_only=True)
            db_datas = db_ws.iter_rows(min_row=2, max_col=9, values_only=True)
            # db_header = [r for r in db_header]
            db_datas = [r for r in db_datas]
            
            
            # 선택 항목의 iid값 가져오기
            for i in range(1, db_ws.max_row+1):
                if db_ws.cell(i,1).value == get_iid():
                    # print(i)
                    # r_selected = i
                    db_ws.cell(i,2).value = entry_task_name.get()
                    db_ws.cell(i,3).value = entry_team_name.get()
                    db_ws.cell(i,4).value = entry_person_name.get()
                    db_ws.cell(i,5).value = cmb_situation.get()
                    db_ws.cell(i,6).value = entry_date.get()
                    db_ws.cell(i,7).value = entry_time.get()
                    db_ws.cell(i,8).value = entry_note.get()

            wb.save("C:/Python/Code/ToDoList/ToDoList_Form.xlsx")

            msgbox.showinfo("알림", "수정 되었습니다.")

            add_tk.withdraw()
            sort()
            load_task()

        # print(get_iid())

        wb = load_workbook("C:/Python/Code/ToDoList/ToDoList_Form.xlsx")
        db_ws = wb['DB']    
        # values_only=True를 해야 위치값이 아닌 실제 데이터 값이 도출됨
        # db_header = db_ws.iter_rows(min_row=1, max_row=1, max_col=9, values_only=True)
        db_datas = db_ws.iter_rows(min_row=2, max_col=9, values_only=True)
        # db_header = [r for r in db_header]
        db_datas = [r for r in db_datas]
        
        
        # 선택 항목의 iid값 가져오기
        for i in range(1, db_ws.max_row+1):
            if db_ws.cell(i,1).value == get_iid():
                # print(i)
                r_selected = i
        # print(db_datas[r_selected-2][1])
        
        task_name = db_datas[r_selected-2][1]
        # print(task_name)
        team_name = db_datas[r_selected-2][2]
        person_name = db_datas[r_selected-2][3]
        situation = db_datas[r_selected-2][4]
        d_date = db_datas[r_selected-2][5]
        d_time = db_datas[r_selected-2][6]
        note = db_datas[r_selected-2][7]


        # for x in range(db_ws.max_row-1):

        #     if db_datas[x][0] == db_datas[r_selected-2][0]:
                # print(db_datas[x][0])
                # print(db_datas[r_selected-2][0])

        # 업무명/내용
        frame_task_name = Frame(add_tk)
        frame_task_name.pack(fill="x")
        
        lbl_task_name = Label(frame_task_name, text="업무명/내용")
        lbl_task_name.pack(side="left")
        
        entry_task_name = Entry(frame_task_name)
        # 기존 DB값 불러오기
        entry_task_name.insert(0,task_name)

        entry_task_name.pack(side="right")

        # 담당팀
        frame_team_name = Frame(add_tk)
        frame_team_name.pack(fill="x")

        lbl_team_name = Label(frame_team_name, text="담당팀")
        lbl_team_name.pack(side="left")

        entry_team_name = Entry(frame_team_name)
        entry_team_name.insert(0,team_name)
        entry_team_name.pack(side="right")

        # 담당자
        frame_person_name = Frame(add_tk)
        frame_person_name.pack(fill="x")

        lbl_person_name = Label(frame_person_name, text="담당자")
        lbl_person_name.pack(side="left")

        entry_person_name = Entry(frame_person_name)
        entry_person_name.insert(0,person_name)
        entry_person_name.pack(side="right")

        # 상황
        frame_situation = Frame(add_tk)
        frame_situation.pack(fill="x")

        lbl_situation = Label(frame_situation, text="상황")
        lbl_situation.pack(side="left")

        opt_situation = ["진행","완료","검토","취소","중단","지연"]
        cmb_situation = ttk.Combobox(frame_situation,state="readonly",
        values=opt_situation)
        cmb_situation.current(0)
        cmb_situation.pack(side="right")

        # 완료날짜
        frame_date = Frame(add_tk)
        frame_date.pack(fill="x")

        lbl_date = Label(frame_date, text="완료날짜")
        lbl_date.pack(side="left")

        entry_date = Entry(frame_date)
        entry_date.insert(0,d_date)
        entry_date.pack(side="right")
        # entry_date.insert(END, "YY-MM-DD")

        # 완료시간
        frame_time = Frame(add_tk)
        frame_time.pack(fill="x")

        lbl_time = Label(frame_time, text="완료시간")
        lbl_time.pack(side="left")

        entry_time = Entry(frame_time)
        entry_time.insert(0,d_time)
        entry_time.pack(side="right")
        # entry_time.insert(END, "HH:MM")

        
        # 비고
        frame_note = Frame(add_tk)
        frame_note.pack(fill="x")

        lbl_note = Label(frame_note, text="비고")
        lbl_note.pack(side="left")

        entry_note = Entry(frame_note)
        entry_note.insert(0,note)
        entry_note.pack(side="right")


        # fucntion2 frame
        frame_func2 = Frame(add_tk)
        frame_func2.pack(fill="x")

        # add db
        btn_apply = Button(frame_func2, text="수정반영",command=apply)
        btn_apply.pack(side="left")

        # cancel button
        # quit 함수를 쓰면 전체 프로그램이 종료되어서 withdraw 함수 사용함
        btn_cancel = Button(frame_func2, text="취소", command=add_tk.withdraw)
        btn_cancel.pack(side="right")

        add_tk.mainloop()



    except Exception as err:
        msgbox.showerror("에러", err)

def load_task():
    
    wb = load_workbook("C:/Python/Code/ToDoList/ToDoList_Form.xlsx")
    db_ws = wb['DB']    
    # values_only=True를 해야 위치값이 아닌 실제 데이터 값이 도출됨
    db_header = db_ws.iter_rows(min_row=1, max_row=1, max_col=9, values_only=True)
    db_datas = db_ws.iter_rows(min_row=2, max_col=9, values_only=True)
    db_header = [r for r in db_header]
    db_datas = [r for r in db_datas]
    
    # treeview reset
    # 이게 없으면 업무 추가시 기존 데이터가 2번 돌면서 iid가 또 있다고 하면서 에러가 남
    tree.delete(*tree.get_children())

    for data in db_datas:
        # print(data[0][6:8])
        # text 값이 있어야 하이라키구조 처럼 보여질 수 있음 (+-표시).
        # lv1에만 open=true를 줘서 실행 시 lv2까지만 자동 보여주기    
        if data[0][3:8] == '00-00':
            lv1 = tree.insert('','end', iid=data[0], text=data[1], values=data[2:8], open=True)
        elif '00' in data[0][6:8]:
            lv2 = tree.insert(lv1,'end',iid=data[0], text=data[1] ,values=data[2:8])
        else :
            lv3 = tree.insert(lv2,'end',iid=data[0], text=data[1] ,values=data[2:8])


def sort():
    db_df = pd.read_excel("C:/Python/Code/ToDoList/ToDoList_Form.xlsx", sheet_name='DB')
    # print(db_df)
    db_df = db_df.sort_values(by=['No'])
    # print(db_df)
    db_df.to_excel("C:/Python/Code/ToDoList/test.xlsx",index=False, sheet_name="DB")
    path1 = "C:/Python/Code/ToDoList/ToDoList_Form.xlsx"
    path2 = "C:/Python/Code/ToDoList/test.xlsx"
    
    wb2 = load_workbook(filename=path2)
    ws2 = wb2['DB']

    wb1 = load_workbook(filename=path1)
    wb1.remove(wb1['DB'])
    ws1 = wb1.create_sheet()
    ws1.title = 'DB'
    # ws1 = wb1['DB']
    for row in ws2:
        for cell in row:
            ws1[cell.coordinate].value = cell.value
    wb1.save(path1)

# class Main(tk.Frame):
#     def __init__(self, parent):
#     # def __init__(self, parent):

#         tk.Frame.__init__(self, parent)
        # tk.Frame.__init__(self, parent)

# if __name__ == "__main__":

win = Tk()
win.title("To Do List")

# browse : 1개만 선택
# extended : 2개 이상 선택 가능
# tree = ttk.Treeview(win, selectmode='extended')
tree = ttk.Treeview(win, selectmode='browse')

tree.pack(expand=True, fill="both")

# # Number of rows to display, default is 10
tree['height'] = 20

tree['columns'] = ("1", "2", "3", "4", "5", "6")

# column보다 heading이 1개 더 많아서(#0) 하이라키구조 처럼 +- 표시가 됨. 
tree.heading("#0", text="업무명/내용")
tree.heading("#1", text="담당팀")
tree.heading("#2", text="담당자")
tree.heading("#3", text="상황")
tree.heading("#4", text="완료날짜")
tree.heading("#5", text="완료시간")
tree.heading("#6", text="비고")

tree.column("1", width=100)
tree.column("2", width=100)
tree.column("3", width=50, anchor='c')
tree.column("4", width=70, anchor='c')
tree.column("5", width=70, anchor='c')
tree.column("6", width=100)

# function frame
func_frame = Frame(win)
func_frame.pack(fill="x")

# 업무 추가 버튼
btn_add_task = Button(func_frame, text="신규 추가(Lv1)", command=add_task)
btn_add_task.pack(side="left")

# # 하위 업무 추가
btn_add_task2 = Button(func_frame, text="하위 추가(Lv2~3)", command=add_task2)
btn_add_task2.pack(side="left")

# 삭제 버튼
btn_del_task = Button(func_frame, text="삭제", command=del_task)
btn_del_task.pack(side="left")

# 편집 버튼
btn_edit_task = Button(func_frame, text="수정", command=edit_task)
btn_edit_task.pack(side="left")

# 종료 버튼
btn_close = Button(func_frame, text="종료", command=win.quit)
btn_close.pack(side="right")

sort()
load_task()

# column_names = ("team_name", "person_name", "situation",
#                 "date","time","note")

# # tree_main = tree(win, columns=column_names)

# tree_main.heading("#0", text="업무명/내용")
# tree_main.heading("team_name", text="담당팀")
# tree_main.heading("person_name", text="담당자")
# tree_main.heading("situation", text="상황")
# tree_main.heading("date", text="완료날짜")
# tree_main.heading("time", text="완료시간")
# tree_main.heading("note", text="비고")

# tree_main.column("1", width=100)
# tree_main.column("2", width=100)
# tree_main.column("3", width=50, anchor='c')
# tree_main.column("4", width=70, anchor='c')
# tree_main.column("5", width=70, anchor='c')
# tree_main.column("6", width=100)



win.mainloop()

