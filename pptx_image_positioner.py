import tkinter as tk
import win32com.client

def mm_to_pt(mm):
    return float(mm) * 2.83465

def update_image():
    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        selection = ppt.ActiveWindow.Selection

        left_pt = mm_to_pt(left_entry.get())
        top_pt = mm_to_pt(top_entry.get())
        width_pt = mm_to_pt(width_entry.get())
        height_pt = mm_to_pt(height_entry.get())
        outline_on = outline_var.get()  # 윤곽선 여부

        found = False
        for shape in selection.ShapeRange:
            if shape.Type in [13, 14]:  # 13=Picture, 14=OLE Object
                shape.Left = left_pt
                shape.Top = top_pt
                shape.Width = width_pt
                shape.Height = height_pt

                # 윤곽선 적용/제거
                if hasattr(shape, "Line"):
                    shape.Line.Visible = outline_on

                found = True

        if found:
            set_message("이미지 위치/크기(mm) 변경 완료!")
        else:
            set_message("선택된 도형은 이미지가 아닙니다.")
    except Exception as e:
        set_message(f"오류 발생: {e}")

def set_message(msg):
    message_label.config(text=msg)

root = tk.Tk()
root.title("PowerPoint 이미지 위치/크기(mm단위) 조정")

tk.Label(root, text="높이(mm)").grid(row=0, column=0)
height_entry = tk.Entry(root)
height_entry.grid(row=0, column=1)

tk.Label(root, text="너비(mm)").grid(row=1, column=0)
width_entry = tk.Entry(root)
width_entry.grid(row=1, column=1)

tk.Label(root, text="가로 위치(mm)").grid(row=2, column=0)
left_entry = tk.Entry(root)
left_entry.grid(row=2, column=1)

tk.Label(root, text="세로 위치(mm)").grid(row=3, column=0)
top_entry = tk.Entry(root)
top_entry.grid(row=3, column=1)

# 윤곽선 체크박스
outline_var = tk.BooleanVar(value=True)
outline_check = tk.Checkbutton(root, text="윤곽선", variable=outline_var)
outline_check.grid(row=4, column=0, columnspan=2)

update_btn = tk.Button(root, text="변경하기", command=update_image)
update_btn.grid(row=5, column=0, columnspan=2)

# 메시지 표시용 라벨(하단)
message_label = tk.Label(root, text="", fg="blue")
message_label.grid(row=6, column=0, columnspan=2)

root.mainloop()
