import xlrd
import os
import numpy as np
import matplotlib.pyplot as plt
import scipy.stats as stats
import PySimpleGUI as sg

while True:
    # initialize parameters
    mean = 0
    variance = 0
    standard_deviation = 0
    file_name = ""
    data_set = []

    # clean the screen
    os.system("cls")
# -----------------------------------------------------------------
#                   start interface
# -----------------------------------------------------------------
    # create layout_start
    layout_start = [
        [sg.Button("Start")],
        [sg.Button("Exit")]
    ]

    # create window_start
    window_start = sg.Window("Start", layout_start)

    # read the window_start
    event, values = window_start.read()
    if event == "Start":
        window_start.close()

    elif event in ("Exit", None):
        window_start.close()
        exit()
# -----------------------------------------------------------------
#                   choose file interface
# -----------------------------------------------------------------
    # create layout_choose_file
    layout_file = [
        [sg.InputText("File_Name", key="fileName")],
        [sg.Button("Confirm")],
        [sg.Button("Exit")]
    ]

    window_file = sg.Window("Choose File", layout_file)

    while True:
        event_file, values_file = window_file.read()

        if event_file == "Confirm":
            # check if file exists
            file_name = values_file["fileName"]
            check = os.path.isfile(file_name)
            if check == 0:
                sg.Popup("File not found")
            else:
                sg.Popup(f"{file_name} found")
                window_file.close()
                break

        elif event_file in ("Exit", None):
            window_file.close()
            exit()

# -----------------------------------------------------------------
#                   calculation interface
# -----------------------------------------------------------------
    # initialize data set
    data_set = []
    excel = xlrd.open_workbook(file_name)
    sheet = excel.sheet_by_index(0)

    # use interface to choose which specific range of data in excel to use
    # create layout_choose_data
    layout_data = [
        [sg.Text("Row range (1 ~", sheet.nrows, ")")],
        [sg.Text("Start_row："), sg.InputText(key="start_row")],
        [sg.Text("End_row："), sg.InputText(key="end_row")],
        [sg.Text("Colon range (1 ~", sheet.ncols, ")")],
        [sg.Text("Start_col："), sg.InputText(key="start_col")],
        [sg.Text("End_col："), sg.InputText(key="end_col")],
        [sg.Button("Confirm")],
        [sg.Button("Exit")]
    ]

    window_data = sg.Window("Choose Data", layout_data)

    while True:
        event_data, values_data = window_data.read()
        if event_data == "Confirm":
            # check if the input is valid
            try:
                start_row = int(values_data["start_row"])
                end_row = int(values_data["end_row"])
                start_col = int(values_data["start_col"])
                end_col = int(values_data["end_col"])
            except ValueError:
                sg.Popup("Invalid input")
                continue
            if start_row > end_row or start_col > end_col or end_row > sheet.nrows or end_col > sheet.ncols:
                sg.Popup("Invalid input")
                continue

            # append data to data set
            for i in range(int(start_row), int(end_row) + 1):
                for j in range(int(start_col), int(end_col) + 1):
                    data_set.append(sheet.cell_value(i - 1, j - 1))
            sg.Popup("Data set is：", data_set)

            # ask if keep adding
            command_add = sg.PopupYesNo("Keep adding?")
            if command_add == "No":
                window_data.close()
                break

            # exit the program
            elif event_data in ("Exit", None):
                window_data.close()
                exit()

    # calculate mean and variance
    mean = sum(data_set) / len(data_set)
    variance = sum([(i - mean) ** 2 for i in data_set]) / len(data_set)
    standard_deviation = np.sqrt(variance)

    # whats the problem with my coming code?
    # print the following command in the interface
    layout_check = [
        [sg.Text(f"Mean: {mean}")],
        [sg.Text(f"Variance: {variance}")],
        [sg.Text(f"Standard Deviation: {standard_deviation}")],
        [sg.Button("Graph")],
        [sg.Button("Calculate Possibility")],
        [sg.Button("End")],
        [sg.Button("Exit")]
    ]

    window_check = sg.Window("Check", layout_check)
    while True:
        event_check, values_check = window_check.read()
        if event_check == "Graph":
            x = np.linspace(mean - 3 * standard_deviation, mean + 3 * standard_deviation, 100)
            y = stats.norm.pdf(x, mean, standard_deviation)
            plt.plot(x, y, 'r--')

        elif event_check == "Calculate Possibility":
            # start a new interface and ask for x1 and x2
            layout_possibility = [
                [sg.Text("Insert lower bound："), sg.InputText(key="x1")],
                [sg.Text("Insert higher bound："), sg.InputText(key="x2")],
                [sg.Button("Confirm")],
                [sg.Button("Exit")]
            ]

            window_possibility = sg.Window("Calculate Possibility", layout_possibility)

            while True:
                event_possibility, values_possibility = window_possibility.read()
                if event_possibility == "Confirm":
                    # check if x1 and x2 are valid
                    try:
                        x1 = float(values_possibility["x1"])
                        x2 = float(values_possibility["x2"])
                    except ValueError:
                        sg.Popup("Invalid input")
                        continue

                    possibility = abs((stats.norm.cdf(x2, mean, standard_deviation) - stats.norm.cdf(x1, mean,
                                                                                                     standard_deviation)
                                       ) * 100)

                    # print the possibility with precision to 4 decimal places
                    sg.Popup("Possibility: ", round(possibility, 2), "%")

                elif event_check in ("Exit", None):
                    window_check.close()
                    break

        elif event_check == "End":
            excel.release_resources()
            window_check.close()
            break

        elif event_check in ("Exit", None):
            window_check.close()
            exit()


