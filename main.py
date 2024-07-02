from tkinter import *
from tkinter import ttk
import pandas as pd
from pandastable import Table
import matplotlib.pyplot as plt
from heapq import nlargest

class SearchButton(Button):
    def __init__(self, part_entry, run_entry, date_entry, date_end, yield_data, mate_yield, et_yield, qc_yield, drop_menu):
        super().__init__()
        self.button = self.config(text="Search", width=7, command=lambda: self.search_data(part_entry.get(), run_entry.get(), date_entry.get(), date_end.get(), yield_data, [mate_yield, et_yield, qc_yield], drop_menu))
        self.database = None
        self.search_query = ["","",""]
        self.label = Label(text=f"Part: {self.search_query[0]} | Run: {self.search_query[1]} | Date: {self.search_query[2]}")
        self.label.grid(column=3, row=0, columnspan=2, padx=5, sticky=W)
        self.defects = []

    # Method for looking up data obtained from entry fields
    def search_data(self, part_number, run_number, date, date_end, yield_data, station_code, drop_menu):
        self.defects = []
        # Filters database by part number and run number input
        search = yield_data[(yield_data["PIPartNo"].str.contains(part_number))].where((yield_data["Run"].astype("str").str.contains(run_number)))
        
        # Filters database by date input
        try:
            if date == "" and date_end != "":
                search = search[search["Date"] <= date_end]
            elif date != "" and date_end == "":
                search = search[search["Date"] >= date]
            else:
                search = search[(search["Date"] >= date) & (search["Date"] <= date_end)]
        # try:
        #     search = search[search["Date"] >= date]
        except TypeError:
            print("Invalid date format.")
            return

        # Create empty dataframe to store searched data
        search_data = pd.DataFrame().reindex_like(search).dropna()

        # Filters database by station code check buttons and appends to search_data dataframe
        if station_code[0].get() == 1:
            search_m2d = search[search["StationName"].str.contains('M2D')]
            search_data = pd.concat((search_data, search_m2d))
        if station_code[1].get() == 1:
            search_f3d = search[search["StationName"].str.contains('F3D')]
            search_data = pd.concat((search_data, search_f3d))
        if station_code[2].get() == 1:
            search_fld = search[search["StationName"].str.contains('FLD')]
            search_data = pd.concat((search_data, search_fld))

        # Sets database and search query attributes using search data
        self.database = search_data
        self.search_query = [part_number, run_number, date, station_code]

        # Extracts defect codes for searched parts
        all_defects = [defect for defect in self.database.iloc[:,9:]]

        # Calculates total relative percentage for all defects 
        self.database = self.database.fillna(0)
        self.database.loc["mean"] = (self.database[all_defects].mean() / self.database["QTY IN"].mean(skipna=False)).apply('{:.2%}'.format)

        # Updates search label with currently loaded query parameters
        self.label.config(text=f"Part: {self.search_query[0]} | Run: {self.search_query[1]} | Date: {self.search_query[2]}")
        
        # Remove defect codes for columns with values greater equal to 0% and sets option menu to contain remaining non-zero defects
        for column in self.database.iloc[:,9:]:
            if self.database.loc["mean",column] != "0.00%":
                self.defects.append(column)
            else:
                self.database.drop(labels=column, inplace=True, axis=1)
        drop_menu.option_update(self.defects)


class DropMenu():
    def __init__(self, parent):
        self.parent = parent
        self.options = ["Defects"]
        self.option_var = StringVar(self.parent)
        self.option_var.set(self.options[0])
        self.option_var.trace("w", self.option_select)

        self.menu = OptionMenu(self.parent, self.option_var, *self.options)
        #self.menu.config(width=6)
        self.menu.grid(column=4, row=4, sticky="ew", padx=5, pady=5)

        self.selection = ""

    # Method to return current dropdown selection
    def option_select(self, *args):
        self.selection = self.option_var.get()
        return self.selection

    # Method to update dropdown menu using defect values returned from latest search dataframe
    def option_update(self, defects):
        self.menu["menu"].delete(0, "end")
        for defect in defects:
            self.menu["menu"].add_command(label=defect, command=lambda value=defect: self.option_var.set(value))


def main():
    yield_data = pd.read_excel("2014-2022 stationdefects.xlsx")
 
    # Create a window for search options
    window = Tk()
    window.title("Defect Database")
    window.config(padx=10, pady=10)
    ttk.Separator(orient=VERTICAL, master=window).grid(column=3, row=0, ipady=50, rowspan=6, sticky=W)

    # Create entry fields for all searchable inputs
    part_entry = Entry(width=21)
    part_entry.grid(column=1, row=0, padx=5)
    run_entry = Entry(width=21)
    run_entry.grid(column=1, row=1)
    date_entry = Entry(width=21)
    date_entry.grid(column=1, row=2)
    date_end = Entry(width=21)
    date_end.grid(column=1, row=3, padx=5)

    # Create text labels for all input fields
    part_label = Label(text="Part Number:")
    part_label.grid(column=0, row=0, pady=5, sticky=W)
    run_label = Label(text="Run Number:")
    run_label.grid(column=0, row=1, sticky=W)
    date_label = Label(text="Date range:")
    date_label.grid(column=0, row=2, sticky=W)
    date_label = Label(text="to")
    date_label.grid(column=0, row=3, padx=5, sticky=E)

    # Create check buttons for job workstations and variables to check button status
    mate_yield = IntVar()
    et_yield = IntVar()
    qc_yield = IntVar()
    mate = Checkbutton(text="Mate Yield", variable=mate_yield,)
    et = Checkbutton(text="ET Yield", variable=et_yield)
    qc = Checkbutton(text="QC Yield", variable=qc_yield)
    mate.grid(column=2, row=0, sticky=W)
    et.grid(column=2, row=1, sticky=W)
    qc.grid(column=2, row=2, sticky=W)

    # Create drop down menu for available defects to report
    drop_menu = DropMenu(window)

    # Create search button to run search function using all input entry fields as arguments
    search_button = SearchButton(part_entry, run_entry, date_entry, date_end, yield_data, mate_yield, et_yield, qc_yield, drop_menu)
    search_button.grid(column=0, row=4)

    # Create button for loading top 5 defects chart
    top_defects_button = Button(text="Top 5 Defects", width=10, command=lambda: calc_defects(search_button, 5))
    top_defects_button.grid(column=3, row=3, padx=5)

    # Create button for loading selected defect from dropdown menu
    select_defects_button = Button(text="Add Defect", width=10, command=lambda: calc_defects(search_button, drop_menu.option_select()))
    select_defects_button.grid(column=4, row=3, padx=5)

    # Create button for loading yield chart
    yield_chart_button = Button(text="Yield", width=10, command=lambda: yield_chart(search_button))
    yield_chart_button.grid(column=4, row=1, pady=5)

    # Create button for opening spreadsheet data
    spreadsheet_button = Button(text="Spreadsheet", width=10, command=lambda: spreadsheet(search_button))
    spreadsheet_button.grid(column=3, row=1, padx=5)

    # Create button for saving excel file
    save_button = Button(text="Save", width=7, command=lambda: save_exel_file(search_button))
    save_button.grid(column=2, row=4, sticky=W)

    # Create button for clearing search data
    clear_button = Button(text="Clear", width=7, command=lambda: clear_text(part_entry, run_entry, date_entry, date_end))
    clear_button.grid(column=1, row=4, sticky=W)

    # Keep search options window open until manually closed by user
    window.mainloop()

# Define function to calculate and display yield graph
def yield_chart(search_button):
    # Calculates percent yield from database
    try:
        x = search_button.database["Date"]
        y = (1 - search_button.database["QTY Reject"] / search_button.database["QTY IN"]) * 100
    except TypeError:
        print("No search dataframe loaded.")
        return

    # Creates scatter plot using date and yield data
    plt.scatter(x, y)
    plt.xlabel("Date")
    plt.ylabel("% Yield")
    plt.title(f"{' '.join(search_button.search_query[0:3])} Percent Yield")
    plt.show()

# Define function to create a spreadsheet using current search dataframe
def spreadsheet(search_button):
    data = Tk()
    data.geometry("600x400+500+200")
    frame = Frame(data)
    frame.pack(fill=BOTH, expand=1)
    pt = Table(frame, dataframe=search_button.database)
    pt.show()
    
# Define a function to create new excel file from current search dataframe
def save_exel_file(search_button):
    try:
        search_button.database.to_excel("test.xlsx", index=False)
    except AttributeError:
        print("No search dataframe loaded.")
        return

# Define function to clear all entry fields
def clear_text(part_entry, run_entry, date_entry, date_end):
    part_entry.delete(0, END)
    run_entry.delete(0, END)
    date_entry.delete(0, END)
    date_end.delete(0, END)


# Define function to display top defects chart
def calc_defects(search_button, number_defects):
    try:
        mean_defects = search_button.database.loc["mean", search_button.defects]
    except AttributeError:
        print("No search dataframe loaded.")
        return

    if isinstance(number_defects, int): 
        top_5 = nlargest(number_defects, mean_defects.to_dict(), key=mean_defects.to_dict().get)
    elif number_defects in search_button.defects:
        top_5 = [number_defects]
    else:
        return

    # Calculates percent defects from database
    x = search_button.database["Date"]
    y= []
    for defect in search_button.database[mean_defects[top_5].keys()]:
        y.append((search_button.database[defect] / search_button.database["QTY IN"]) * 100)

    # Creates scatter plot using date and percent defect data from each defect
    for i, defect in enumerate(mean_defects[top_5].keys()):
         plt.scatter(x, y[i], label=defect)
    plt.xlabel("Date")
    plt.ylabel("% Defective")
    plt.title(f"{' '.join(search_button.search_query[0:3])}  %Defects")
    plt.legend()
    plt.show()


if __name__ == "__main__":
    main()