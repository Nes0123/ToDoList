
# 강의 : https://www.youtube.com/watch?v=n5gItcGgIkk
# Tkinter 8.5 reference: a GUI for Python : http://tkdocs.com/shipman/ttk-Treeview.html
# Class 설명 : https://hwan-hobby.tistory.com/93
            # https://dojang.io/mod/page/view.php?id=2373

# 다른 곳을 클릭하면 에러가 나는데 그것에 대한 수정은 강의에 없었음.

import tkinter as tk
from tkinter import ttk

# **kw : **kwargs와 같은건가?
# # **kwargs : dictionary 변수들을 저장시켜주는 역할
# *args : key 없는 argument를 통과시켜줌
# **kwargs : key(a, b,,)가 있는 argument를 통과시켜줌
# master는 parent? 왜 붙는지 모르겠음. 아직 파악하지 못함
# ttk.Treeview libarary?를 전체 class의 super class?로 지정
class TreeviewEdit(ttk.Treeview):
    def __init__(self, master, **kwargs):
        super().__init__(master,**kwargs)
    # def __init__(self, master, **kw):
    #     super().__init__(master,**kw)
        self.bind("<Double-1>", self.on_double_click)

    def on_double_click(self, event):
        # print("Double clicked")
        # Identify the region that was double-clicked
        region_clicked = self.identify_region(event.x, event.y)
        
        # heading과 빈칸이 눌러지면 작동되지 않겠끔 함
        # tree와 cell만 나오게함
        if region_clicked not in ("tree", "cell"):
            return
        # print(region_clicked)

        # which item was double-clicked
        # ex) "#0" is the first column, followed by "#1", "#2" ,,
        column = self.identify_column(event.x)
        # print(column)

        # "#0" will become -1, "#1" will become 0 ,,,
        column_index = int(column[1:]) -1

        # ex) I001
        selected_iid = self.focus()
        # print(selected_iid)

        # this will contain both text and values from the given iid 
        selected_values = self.item(selected_iid)
        # print(selected_values)

        # text column : #0
        if column == "#0":
            selected_text = selected_values.get("text")
        # values columns
        else :
            selected_text = selected_values.get("values")[column_index]

        # print(selected_text)

        # (X position, Y position, Width, Height)
        column_box = self.bbox(selected_iid, column)

        # print(column_box)

        # entry_edit = ttk.Entry(root)
        # 강의에서는 width를 설정해줬는데 아래 width를 지정해줘서 
        #  안 넣어도 되지 않나? 
        entry_edit = ttk.Entry(root, width=column_box[2])
        
        # record the column index and item iid
        entry_edit.editing_column_index = column_index
        entry_edit.editing_item_iid = selected_iid

        entry_edit.insert(0,selected_text)

        # 해당 값을 모두 선택?
        entry_edit.select_range(0,tk.END)
        entry_edit.focus()

        entry_edit.bind("<FocusOut>", self.on_focus_out)
        entry_edit.bind("<Return>", self.on_enter_pressed)
        # pack() 대신에 쓰는듯. 위치와 너비와 높이 지정할 때
        entry_edit.place(x=column_box[0], y=column_box[1],
                        w=column_box[2], h=column_box[3])

    def on_enter_pressed(self, event):
        new_text = event.widget.get()

        # Such as I002
        selected_iid = event.widget.editing_item_iid

        # such as -1 (tree column), 0 (first self defined column),,
        column_index = event.widget.editing_column_index

        # tree column
        if column_index == -1:
            self.item(selected_iid, text=new_text)
        # value columns
        else:
            current_values = self.item(selected_iid).get("values")
            current_values[column_index] = new_text
            self.item(selected_iid,values=current_values)
        
        event.widget.destroy()

    # 해당 위젯 선택을 해제하면 entry 모양이 사라지게함
    def on_focus_out(self, event):
        event.widget.destroy()
            



# if __name__ == "__main__": 이 프로그램의 시작점을 의미함
# 이 파일을 import하여 사용할 때에는 작동되지 않고, 
# 이 파일의 함수(def)는 사용 가능
if __name__ == "__main__":
    root = tk.Tk()

    column_names = ("vehicle_name", "year", "colour")

    treeview_vehicles = TreeviewEdit(root, columns=column_names)
    
    treeview_vehicles.heading("#0", text="Vehicle Type")
    treeview_vehicles.heading("vehicle_name", text="Vehicle_Name")
    treeview_vehicles.heading("year", text="Year")
    treeview_vehicles.heading("colour", text="Colour")

    sedan_row = treeview_vehicles.insert(parent="", index=tk.END, text="Sedan")

    treeview_vehicles.insert(parent=sedan_row,
        index=tk.END,
        values=("Nissan Versa", "2010", "Silver"))

    treeview_vehicles.insert(parent=sedan_row,
        index=tk.END,
        values=("Toyota Camry","2012","Blue"))

    treeview_vehicles.pack(fill=tk.BOTH, expand=True)



    root.mainloop()

