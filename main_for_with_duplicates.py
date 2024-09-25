import pandas
import tkinter as tk
from tkinter import filedialog
import os


# create path to file by dialog window:
def open_file_dialog(initial_dir=None):
    if initial_dir is None:
        initial_dir = r"%USERPROFILE%\Desktop"
    root = tk.Tk()
    root.withdraw() # Hide the main window
    filetypes = [("Excel files", "*.xlsx")]
    file_path = filedialog.askopenfilename(initialdir=initial_dir, filetypes=filetypes)
    if file_path:
        directory_path = os.path.dirname(file_path) #Get the directory chosen file
        root.destroy()  # Destroy the root window to close the dialog
        excel_file_name = os.path.splitext(os.path.basename(file_path))[0]

        #print(excel_file_name)
        return file_path, directory_path, excel_file_name

#load excel file
file = open_file_dialog()
#print(file)
excel_file_path = file[0]
path_to_dir = file[1]
name_excelfile = file[2]

# data file:
df = pandas.read_excel(excel_file_path)
# initialize a list to store the processed data
restructed_date = [] # will do in list, because we can have duplicates

# restructuring coordinates of points
for index, row in df.iterrows():
    num_kui = row["架台番号"]
    # print(row["架台番号"])
    restructed_date.append(num_kui)

    coordinates = []
    for x, y in zip(df.columns[4::2], df.columns[5::2]):   # Skipping the first column and 2 and 3 coloumns with PCS番号
        x0 = row[y]
        y0 = row[x]
        if isinstance(y0, float) and isinstance(x0, float):
            x1 = round((x0 / 1), 3)
            y1 = round((y0 / 1), 3)
            coordinates.append([x1, y1])
    restructed_date.append(coordinates)

print(restructed_date)
data_list = []
z = 0
for point_name, coord in zip(restructed_date[0::2], restructed_date[1::2]):
    n = 1
    for value in coord:
        kui_name = f"{point_name}.{n}"

        # make condition when we already have same name in row and then just continue count index of it
        k = n + 1
        for point in data_list:
            while kui_name in point:
                kui_name = f"{point_name}.{k}"
                k += 1

        data_list.append([kui_name, value[0], value[1], z])
        n += 1
print(data_list)
# Create a new DataFrame
new_data = pandas.DataFrame(data_list, columns=["Kui_name", "x", "y", "z"])
# if you want delete the head line just put header=False in .to_csv parameters
path = fr"{path_to_dir}/{name_excelfile}.csv"
#print(path)
new_data.to_csv(path, index=False, header=False)


