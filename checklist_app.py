import tkinter as tk
from checklist_generator.checklist_gui import ChecklistApp

if __name__ == "__main__":
    root = tk.Tk()
    app = ChecklistApp(root)
    root.mainloop()
