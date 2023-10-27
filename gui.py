# Import required modules
from tkinter import Tk, Label, Entry, Button, messagebox, filedialog, END
from user_check import UserCheckLogic


class UserCheckGUI():
    def __init__(self, window):
        """Initialization of UserCheckGUI class"""
        self.window = window
        self.window.title("Provjera statusa korisnika")
        self.window.config(bg="light yellow")
        self.window.resizable(True, True)  # Make it resizable both ways
        self.window.columnconfigure(0, weight=1)  # Resize the width
        self.window.rowconfigure(0, weight=1)  # Resize the height

        window_width = 600
        window_height = 150

        # Initiate the GUI on the screen center
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        # Set the window to the screen center
        self.window.geometry("{}x{}+{}+{}".format(window_width, window_height, x, y))

        self.excel_file = ""

        # Description Label
        self.description_lbl = Label(window, bg="light yellow",
                                     text="Molimo odaberite Excel tablicu - podržani formati su \".xls\", \".xlsx\" i \".xlsm\".\nProgram provjerava korisnike iz stupca \"Username\" i vraća status u stupac \"Account Status\",\nneovisno o njihovoj poziciji stupaca. Pobrinite se samo da u tablici postoje stupci navedenih naziva\ni da se nazivi stupaca nalaze u prvom redu.",
                                     anchor="w", justify="left", width=75, padx=5, pady=5)
        self.description_lbl.grid(column=0, row=0)

        # Selected table entry
        self.selected_table_name_ent = Entry(window, bg="light yellow",
                                             state="disabled", width=75)
        self.selected_table_name_ent.grid(column=0, row=1, padx=10, sticky="w")

        # Is table selected label
        self.is_table_selected_lbl = Label(window, text="", state="disabled")
        self.is_table_selected_lbl.grid(column=0, row=2, sticky="w")

        # Button to search for the table
        self.browse_btn = Button(window, bg="light blue", text="Pretraži", command=self.browse_excel)
        self.browse_btn.grid(column=1, row=1, padx=5, pady=5, sticky="ew")

        # Button to perform action
        self.start_btn = Button(window, text="Kreni", state="disabled", command=self.start_check)
        self.start_btn.grid(column=1, row=2, padx=5, pady=5, sticky="ew")

    def browse_excel(self):
        """Browse for the table"""
        self.excel_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*xls;*.xlsx;*xlsm")])
        if self.excel_file:
            self.start_btn.config(bg="light green", state="normal")
            self.selected_table_name_ent.config(state="normal")
            self.selected_table_name_ent.delete(0, END)
            self.selected_table_name_ent.insert(0, self.excel_file)
            self.is_table_selected_lbl.config(state="normal",
                                              bg="light yellow",
                                              fg="dark green",
                                                text="Tablica odabrana, počni provjeru statusa korisnika.")

    def start_check(self):
        """Start the check - get it from the UserCheckLogic class"""
        try:
            if self.excel_file:
                user_check = UserCheckLogic(self.excel_file)
                user_check.check_user_status()
                user_check.save_to_file()
                messagebox.showinfo("Info", "Provjera korisničkih statusa je završena.")
            else:
                messagebox.showerror("Pogreška", "Molimo odaberite Excel tablicu prije pokretanja provjera.")
        except ValueError as e:
            messagebox.showerror("Pogreška", str(e))