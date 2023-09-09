import customtkinter as ctk
from pandas import read_excel, DataFrame
from sys import exit

ctk.set_default_color_theme('green')
ctk.set_appearance_mode('System')


class Window(ctk.CTk):
    def __init__(self, df, y_s, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.df = df
        self.y_s = y_s


class App(ctk.CTk):

    def __init__(self, df, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.df = df
        self.title("Selection Menu")
        self.geometry('300x300')
        self.branches = ('CHEM', 'CIVIL', "ECE", 'CSE', "MECH", "EEE", "MME")
        self.branch = ctk.StringVar(self, value=None)
        self.year = ctk.StringVar(self, value=None)
        self.sem = ctk.StringVar(self, value=None)
        self.Create()

    def Topclass(self):
        Window(self.df[self.branch.get()], self.year.get() + '-' + self.sem.get()).mainloop()

    def Create(self):
        ctk.CTkLabel(self, text='Branch'.ljust(10) + ':',
                     font=("Arial", 16)).grid(row=0, column=0, pady=(50, 20), padx=20)
        ctk.CTkOptionMenu(self, values=self.branches, dynamic_resizing=False,
                          variable=self.branch).grid(row=0, column=1, pady=(50, 20), padx=20)
        ctk.CTkLabel(self, text="Year".ljust(10) + ':', font=("Arial", 16)).grid(row=1, column=0, pady=20, padx=20)
        ctk.CTkOptionMenu(self, values=('E1', 'E2', 'E3', 'E4'), dynamic_resizing=False,
                          variable=self.year).grid(row=1, column=1, pady=20, padx=20)
        ctk.CTkLabel(self, text="Sem".ljust(10) + ':', font=("Arial", 16)).grid(row=2, column=0, pady=20, padx=20)
        ctk.CTkOptionMenu(self, values=('S1', 'S2'), variable=self.sem, dynamic_resizing=False,
                          command=self.Topclass).grid(row=2, column=1, pady=20, padx=20)


def ui(Q):
    App(Q).mainloop()


def get_data(q):
    br = ("CHEM", "CIVIL", "ECE", "CSE", "MECH", "EEE", "MME")
    for i in br:
        q[i] = read_excel('./GPA.xlsx', sheet_name=i)


if __name__ == '__main__':
    Q: dict = dict()
    try:
        get_data(Q)
        print(Q['CSE'].head())
        ui(Q)
    except:
        exit(0)
