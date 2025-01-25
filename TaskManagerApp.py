from tkinter import Tk, Label, Button, Entry,  END, filedialog, Toplevel, Frame, Scrollbar
from tkinter.ttk import Combobox, Treeview, Style
from tkcalendar import Calendar
import pandas as pd

def add_task():
    task = task_entry.get()
    priority = priority_combobox.get()
    date = date_entry.get()
    
    if len(task) > 20:
        feedback_label.config(
            text="😡 Task name exceeds 20 characters. Please shorten it.",
            fg="red",
            font=("Arial", 20, "bold"),
        )
    elif task.isdigit():
        feedback_label.config(
            text="😡 Task name cannot be purely numerical.",
            fg="red",
            font=("Arial", 20, "bold"),
        )
    elif task and priority and date:
        tasks.append((task, priority, date, "Pending"))
        task_entry.delete(0, END)
        priority_combobox.set("")
        date_entry.delete(0, END)
        feedback_label.config(
            text="✅ Task added successfully.",
            fg="green",
            font=("Arial", 20, "bold"),
        )
    else:
        feedback_label.config(
            text="😡 Please fill all fields before adding a task.",
            fg="red",
            font=("Arial", 20, "bold"),
        )

def view_tasks():
    if tasks:
        for item in task_tree.get_children():
            task_tree.delete(item)
        for task in tasks:
            task_tree.insert("", END, values=task)
        feedback_label.config(text="✅ Tasks displayed successfully.", fg="green", font=("Arial", 20, "bold"))
    else:
        feedback_label.config(text="No tasks to display. Add a task first.", fg="red", font=("Arial", 20, "bold"))

def mark_complete():
    selected_item = task_tree.selection()
    if selected_item:
        for item in selected_item:
            values = task_tree.item(item, "values")
            updated_values = (values[0], values[1], values[2], "Completed")
            task_tree.item(item, values=updated_values)
            index = tasks.index((values[0], values[1], values[2], values[3]))
            tasks[index] = updated_values
        feedback_label.config(text="✅ Task(s) marked as complete.", fg="green", font=("Arial", 20, "bold"))
    else:
        feedback_label.config(text="😡 Please select a task to mark as complete.", fg="red", font=("Arial", 20, "bold"))

def delete_task():
    selected_item = task_tree.selection()
    if selected_item:
        for item in selected_item:
            values = task_tree.item(item, "values")
            tasks.remove((values[0], values[1], values[2], values[3]))
            task_tree.delete(item)
        feedback_label.config(text="✅ Task(s) deleted successfully.", fg="green", font=("Arial", 20, "bold"))
    else:
        feedback_label.config(text="😡 Please select a task to delete.", fg="red", font=("Arial", 20, "bold"))

def export_to_excel():
    if tasks:
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            df = pd.DataFrame(tasks, columns=["Task", "Priority", "Date", "Status"])
            df.to_excel(file_path, index=False)
            feedback_label.config(text=f"✅ Tasks exported to {file_path} successfully.", fg="green", font=("Arial", 12, "bold"))
    else:
        feedback_label.config(text="😡 No tasks to export. Add tasks first.", fg="red", font=("Arial", 20, "bold"))

def select_date():
    top = Toplevel(root)
    cal = Calendar(top, selectmode="day", date_pattern="yyyy-mm-dd", font="Arial 14", cursor="hand1")
    cal.pack(fill="both", expand=True)
    ok_button = Button(top, text="OK", command=lambda: handle_date_selection(cal.get_date(), top), 
                       font=("Arial", 12, "bold"), fg="white", bg="grey")
    ok_button.pack()

def handle_date_selection(selected_date, top):
    date_entry.delete(0, END)
    date_entry.insert(0, selected_date)
    top.destroy()

root = Tk()
root.title("Task Manager")
root.geometry("900x600")  # Increased window size

# Task Data
tasks = []

# Configure Treeview Header Style
style = Style()
style.theme_use("default")

style.configure(
    "Custom.Treeview.Heading",
    font=("Arial", 12, "bold"),
    foreground="white",
    background="grey",
    borderwidth=1,
    relief="raised"
)

style.configure(
    "Custom.Treeview",
    rowheight=25,
    font=("Arial", 11),
    background="white",
    foreground="black"
)

# Frames for layout
left_frame = Frame(root)
left_frame.grid(row=0, column=0, padx=10, pady=10, sticky="n")

right_frame = Frame(root)
right_frame.grid(row=0, column=1, padx=10, pady=10)

# Labels and Entry Widgets (Right Frame)
Label(right_frame, text="Task Name:", font=("Arial", 12, "bold"), fg="black").grid(row=0, column=0, padx=10, pady=5, sticky="e")
task_entry = Entry(right_frame, width=40, font=("Arial", 12))  # Wider task entry
task_entry.grid(row=0, column=1, padx=10, pady=5)

Label(right_frame, text="Select Priority:", font=("Arial", 12, "bold"), fg="black").grid(row=1, column=0, padx=10, pady=5, sticky="e")
priority_combobox = Combobox(right_frame, values=["Low", "Medium", "High"], state="readonly", width=38, font=("Arial", 12))
priority_combobox.grid(row=1, column=1, padx=10, pady=5)

Label(right_frame, text="Select Date:", font=("Arial", 12, "bold"), fg="black").grid(row=2, column=0, padx=10, pady=5, sticky="e")
date_entry = Entry(right_frame, width=40, font=("Arial", 12))
date_entry.grid(row=2, column=1, padx=10, pady=5)
select_date_button = Button(right_frame, text="Select Date", command=select_date, font=("Arial", 12, "bold"), fg="white", bg="grey")
select_date_button.grid(row=2, column=2, padx=10, pady=5)

# Buttons (Left Frame)
add_button = Button(left_frame, text="Add Task", command=add_task, width=15, font=("Arial", 12, "bold"), fg="white", bg="grey")
add_button.grid(row=0, column=0, pady=10)

view_button = Button(left_frame, text="View Tasks", command=view_tasks, width=15, font=("Arial", 12, "bold"), fg="white", bg="grey")
view_button.grid(row=1, column=0, pady=10)

mark_complete_button = Button(left_frame, text="Mark Complete", command=mark_complete, width=15, font=("Arial", 12, "bold"), fg="white", bg="grey")
mark_complete_button.grid(row=2, column=0, pady=10)

delete_button = Button(left_frame, text="Delete Task", command=delete_task, width=15, font=("Arial", 12, "bold"), fg="white", bg="grey")
delete_button.grid(row=3, column=0, pady=10)

export_button = Button(left_frame, text="Export to Excel", command=export_to_excel, width=15, font=("Arial", 12, "bold"), fg="white", bg="grey")
export_button.grid(row=4, column=0, pady=10)

# Task Treeview (Right Frame)
columns = ("Task", "Priority", "Date", "Status")
task_tree = Treeview(right_frame, columns=columns, show="headings", style="Custom.Treeview")
task_tree.heading("Task", text="Tasks")
task_tree.heading("Priority", text="Priority")
task_tree.heading("Date", text="Date")
task_tree.heading("Status", text="Status")
task_tree.column("Task", width=220, anchor="w")  # Wider column for tasks
task_tree.column("Priority", width=100, anchor="center")
task_tree.column("Date", width=100, anchor="center")
task_tree.column("Status", width=150, anchor="center")
task_tree.grid(row=5, column=0, columnspan=3, padx=10, pady=10)

# Scrollbar for Treeview
tree_scroll = Scrollbar(right_frame, orient="vertical", command=task_tree.yview)
task_tree.configure(yscroll=tree_scroll.set)
tree_scroll.grid(row=5, column=3, sticky="ns", padx=10)

# Feedback Label
feedback_label = Label(root, text="", font=("Arial", 12, "bold"), fg="green")
feedback_label.grid(row=6, column=0, columnspan=2, pady=10)

# Run the Tkinter Event Loop
root.mainloop()
