import tkinter as tk
import win32com.client

def mm_to_pt(mm):
    return float(mm) * 2.83465

def update_image():
    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        selection = ppt.ActiveWindow.Selection

        left_val = left_entry.get()
        top_val = top_entry.get()
        width_val = width_entry.get()
        height_val = height_entry.get()
        outline_on = outline_var.get()  # Outline

        found = False
        for shape in selection.ShapeRange:
            if shape.Type in [13, 14]:  # 13=Picture, 14=OLE Object
                if left_val:
                    shape.Left = mm_to_pt(left_val)
                if top_val:
                    shape.Top = mm_to_pt(top_val)
                if width_val:
                    shape.Width = mm_to_pt(width_val)
                if height_val:
                    shape.Height = mm_to_pt(height_val)

                # Outline
                if hasattr(shape, "Line"):
                    shape.Line.Visible = outline_on

                found = True

        if found:
            set_message("Finished image positioning!")
        else:
            set_message("This is not an image.")
    except Exception as e:
        set_message(f"Error: {e}")

def set_message(msg):
    message_label.config(text=msg)
    root.after(5000, lambda: message_label.config(text=""))

root = tk.Tk()
root.title("PPTX_IMAGE_POSITIONER")
#root.geometry("250x200")
root.resizable(False, False)

tk.Label(root, text="Height", anchor="e").grid(row=0, column=0, padx=10)
height_entry = tk.Entry(root)
height_entry.grid(row=0, column=1)
tk.Label(root, text="mm").grid(row=0, column=2, padx=10)

tk.Label(root, text="Width").grid(row=1, column=0, padx=10)
width_entry = tk.Entry(root)
width_entry.grid(row=1, column=1)
tk.Label(root, text="mm").grid(row=1, column=2, padx=10)

tk.Label(root, text="X position").grid(row=2, column=0, padx=10)
left_entry = tk.Entry(root)
left_entry.grid(row=2, column=1)
tk.Label(root, text="mm").grid(row=2, column=2, padx=10)

tk.Label(root, text="Y position").grid(row=3, column=0, padx=10)
top_entry = tk.Entry(root)
top_entry.grid(row=3, column=1)
tk.Label(root, text="mm").grid(row=3, column=2, padx=10)

# Checkbox for Outline
outline_var = tk.BooleanVar(value=True)
outline_check = tk.Checkbutton(root, text="Outline", variable=outline_var)
outline_check.grid(row=4, column=1, columnspan=1)

# Go button
update_btn = tk.Button(root, text="     Go!     ", command=update_image)
update_btn.grid(row=5, column=1, columnspan=1)

# Message output
message_label = tk.Label(root, text="", fg="blue")
message_label.grid(row=6, column=0, columnspan=3)

root.mainloop()
