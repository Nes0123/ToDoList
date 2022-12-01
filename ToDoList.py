
from openpyxl import load_workbook
from tkinter import ttk
from tkinter import *

# 엑셀 데이터 추출
wb = load_workbook("C:/Python/Code/ToDoList/ToDoList_Form.xlsx")
db_ws = wb['Sheet2']
# values_only=True를 해야 위치값이 아닌 실제 데이터 값이 도출됨
db_header = db_ws.iter_rows(min_row=1, max_row=1, max_col=9, values_only=True)
db_datas = db_ws.iter_rows(min_row=2, max_col=9, values_only=True)

db_header = [r for r in db_header]
db_datas = [r for r in db_datas]
wb.close()
# print(db_header[0])
db_datas

def load_db():

    # tkinter 활용
    root = Tk()
    root.title("To Do List")

    # treeview 활용
    # y,x 스크롤바 추가 필요
    tree = ttk.Treeview(root, selectmode='extended')
    tree.pack(expand=True, fill="both")
    # tree.grid(row=0,column=0,columnspan=3,padx=30,pady=20)

    # Number of rows to display, default is 10
    tree['height'] = 20
    # 컬럼 제목만 보이게함?
    tree['show'] = 'headings'

    tree['columns'] = db_header[0]

    for i in db_header[0]:
        tree.column(i, width=100, anchor='c')
        tree.heading(i, text=i, anchor='c')


    for data in db_datas:
        tree.insert("",'end', iid=data[0], values=data)


    root.mainloop()

def add_task():
    # 업무 레벨에 따라 No링 잘 하기
    # no 정렬 방식 활용 해보기
    pass


# wb = load_workbook("C:/Python/Code/ToDoList/ToDoList_Form.xlsx")
out_ws = wb['Sheet1']
# values_only=True를 해야 위치값이 아닌 실제 데이터 값이 도출됨
out_header = out_ws.iter_rows(min_row=1, max_row=1, max_col=8, values_only=True)
# out_datas = out_ws.iter_rows(min_row=2, max_col=8, values_only=True)

out_header = [r for r in out_header]
# out_datas = [r for r in out_datas]
wb.close()
# print(out_header[0])
# out_datas


# tkinter 활용
win = Tk()
win.title("To Do List")

# treeview 활용
# y,x 스크롤바 추가 필요
# 클릭시 여러개 선택 가능하도록 (extended)
tree = ttk.Treeview(win, selectmode='extended')
# tree = ttk.Treeview(win, selectmode='browse')

tree.pack(expand=True, fill="both")

# # Number of rows to display, default is 10
tree['height'] = 30

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


# lv1 = tree.insert("",0, "task1", text="인력소요", values=('인사팀','박기아','진행','22-11-20','10:30','ASAP'))
# tree.insert(lv1,"end", text='전 실에 요청 우편', values=('기획팀','박기아','완료','22-11-12','10:30','ASAP'))

# lv1 = tree.insert("",0, "task1", text="인력소요", values=('인사팀','박기아','진행','22-11-20','10:30','ASAP'))
# tree.insert(lv1,"end", text='전 실에 요청 우편', values=('기획팀','박기아','완료','22-11-12','10:30','ASAP'))


# # Number of rows to display, default is 10
# tree['height'] = 20
# # 컬럼 제목만 보이게함?
# tree['show'] = 'headings'

# tree['columns'] = out_header[0]

# for i in out_header[0]:
#     tree.column(i, width=100, anchor='c')

# # tree.column('#0', width=100, anchor='c')
# tree.heading('#0', text="task", anchor='c')


# for i in out_header[0]:
#     # tree.column(i, width=100, anchor='c')
#     tree.heading(i, text=i, anchor='c')


# # task1 = tree.insert("",'end', iid=db_datas[0], values=db_datas[0])
# # task2 = tree.insert(task1,'end', iid=db_datas[1], values=db_datas[1])


# # tree.insert('','end', iid=1, text="dkdkdk")


# ## text와 values의 차이?

# for data in db_datas:
#     # print(type(data[0][6:8]))
#     # if '00' in data[0][6:8] :
#     #     print(data)
#     # else: print("no")

#     if data[0][3:8] == '00-00':
#         tree.insert('','end', 'task1', text=data[2], values=data)

#         # task1 = tree.insert("",'end', iid=data[0], text=data, tags="tag1")


#     # elif data[0][3:8] != '00-00' & '00' in data[0][6:8]:
#     elif '00' in data[0][6:8]:
#         tree.insert('task1','end', text=data[2] ,values=data)
#     else :
#         pass
#         # tree.insert(task2,'end', values=data, tags="tag3")


# # for data in db_datas:
# #     # print(type(data[0][6:8]))

# #     # if '00' in data[0][6:8] :
# #     #     print(data)
# #     # else: print("no")

# #     if data[0][3:8] == '00-00':
# #         # print(data)
# #         # task1 = tree.insert("",'end', iid=data[0], tags="tag1", text=data)
# #         # task1 = tree.insert('','end', iid=data[0], text="111")
# #         # task1 = tree.insert('','end', iid=data[0], values=data, tags="tag1", open=True)

# #         task1 = tree.insert("",'end', iid=data[0], text=data, tags="tag1")


# #     # elif data[0][3:8] != '00-00' & '00' in data[0][6:8]:
# #     elif '00' in data[0][6:8]:
# #         task2 = tree.insert(task1,'end',values=data)
# #     else :
# #         tree.insert(task2,'end', values=data, tags="tag3")

# # tree.move(task2, task1, END)


#         # print(data)
# tree.tag_bind("tag1",sequence="<<TreeviewOpen>>")

win.mainloop()

