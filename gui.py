import tkinter as tk
from tkinter.messagebox import showinfo, showerror
from tkinter.filedialog import askopenfilename
from main import report_creating


class ExcelOperator:
    def __init__(self):
        self.main_window = tk.Tk()
        self.main_window.title("ExcelFile-Filter")
        self.label1 = tk.Label(
            self.main_window,
            text="Hello!\n\nEnter the data of required Excel-File.",
            borderwidth=1,
            relief="groove",
            font=("Georgia", 10, "bold"),
        )
        self.label1.pack(side="top", ipady=10, ipadx=10, pady=20)

        # Labels
        self.frame1 = tk.Frame(self.main_window)

        self.label2 = tk.Label(
            self.frame1, text="Input Excel-File:", font=("Georgia", 10)
        )
        self.label3 = tk.Label(
            self.frame1, text="Column for filtration:", font=("Georgia", 10)
        )
        self.label4 = tk.Label(
            self.frame1, text="Value for filtration:", font=("Georgia", 10)
        )
        self.label5 = tk.Label(
            self.frame1, text="Output Excel-File:", font=("Georgia", 10)
        )

        self.label2.pack(side="top", ipady=5)
        self.label3.pack(side="top", ipady=5)
        self.label4.pack(side="top", ipady=5)
        self.label5.pack(side="top", ipady=5)
        self.frame1.pack(side="left", padx=(120, 10))

        # Data collectors
        self.frame2 = tk.Frame(self.main_window)

        self.input_file = tk.Entry(self.frame2)
        self.column = tk.Entry(self.frame2)
        self.value = tk.Entry(self.frame2)
        self.output_file = tk.Entry(self.frame2)

        self.input_file.pack(
            side="top",
            ipadx=100,
        )
        self.column.pack(side="top", ipadx=50, ipady=5)
        self.value.pack(side="top", ipadx=50, ipady=5)
        self.output_file.pack(side="top", ipadx=50, ipady=5)

        self.input_file_button = tk.Button(
            self.input_file, text="🔍", command=self.browse_input_file
        )
        self.input_file_button.pack(side="right")

        self.frame2.pack(side="left", padx=30)

        # Submit button
        self.data_button = tk.Button(
            self.main_window, text="Submit", command=self.excel_filter
        )
        self.data_button.pack(side="right", padx=10, pady=10)

    def browse_input_file(self):
        filename = askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        self.input_file.delete(0, tk.END)
        self.input_file.insert(0, filename)

    def excel_filter(self):
        input_file = self.input_file.get()
        column = self.column.get()
        output_file = self.output_file.get()

        try:
            value = int(self.value.get())
        except ValueError:
            value = self.value.get()

        respond = report_creating(input_file, column, value, output_file)

        if respond == "success":
            showinfo(title="Respond", message="Check your Excel-File.")
        else:
            showerror(title="Error", message=respond)


if __name__ == "__main__":
    gui = ExcelOperator()
    tk.mainloop()
