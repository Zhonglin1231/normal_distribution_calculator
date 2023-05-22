import xlrd
import os
import numpy as np
import matplotlib.pyplot as plt
import scipy.stats as stats
import PySimpleGUI as sg
import sys


# make a main function
def main():
    cumulative_color = 1

    while True:
        # initialize parameters
        mean = 0
        variance = 0
        standard_deviation = 0

        file_name = "r--"
        color = ""
        data_set = []

        # clean the screen
        os.system("cls")
        # -----------------------------------------------------------------
        #                   start interface
        # -----------------------------------------------------------------
        # create layout_start
        layout_start = [
            # set the specific text size and font
            [sg.Text("正态分布分析器", font=("华文行楷", 20))],
            [sg.Button("开始")],
            [sg.Button("退出")]
        ]

        # create window_start
        window_start = sg.Window("开始", layout_start)

        # read the window_start
        event, values = window_start.read()
        if event == "开始":
            window_start.close()

        elif event in ("退出", None):
            window_start.close()
            sys.exit()
        # -----------------------------------------------------------------
        #                   choose file interface
        # -----------------------------------------------------------------
        # create layout_choose_file
        layout_file = [
            [sg.InputText("文件名", key="文件名")],
            [sg.Button("确认")],
            [sg.Button("退出")]
        ]

        window_file = sg.Window("选择文件", layout_file)

        while True:
            event_file, values_file = window_file.read()

            if event_file == "确认":
                # check if file exists
                file_name = values_file["文件名"]
                check = os.path.isfile(file_name)
                if check == 0:
                    sg.Popup("文件未找到")
                else:
                    sg.Popup(f"{file_name} 找到")
                    window_file.close()
                    break

            elif event_file in ("退出", None):
                window_file.close()
                sys.exit()

        # -----------------------------------------------------------------
        #                   calculation interface
        # -----------------------------------------------------------------
        # initialize data set
        data_set = []
        excel = xlrd.open_workbook(file_name)
        sheet = excel.sheet_by_index(0)

        # use interface to choose which specific range of data in Excel to use
        # create layout_choose_data
        layout_data = [
            [sg.Text(f"行范围 (1 ~ {sheet.nrows})")],
            [sg.Text("首行："), sg.InputText(key="首行")],
            [sg.Text("尾行："), sg.InputText(key="尾行")],
            [sg.Text(f"列范围 (1 ~ {sheet.ncols})")],
            [sg.Text("首列："), sg.InputText(key="首列")],
            [sg.Text("尾列："), sg.InputText(key="尾列")],
            [sg.Button("确认")],
            [sg.Button("退出")]
        ]

        window_data = sg.Window("选择数据", layout_data)

        while True:
            event_data, values_data = window_data.read()
            if event_data == "确认":
                # check if the input is valid
                try:
                    start_row = int(values_data["首行"])
                    end_row = int(values_data["尾行"])
                    start_col = int(values_data["首列"])
                    end_col = int(values_data["尾列"])
                except ValueError:
                    sg.Popup("输入无效")
                    continue
                if start_row > end_row or start_col > end_col or end_row > sheet.nrows or end_col > sheet.ncols:
                    sg.Popup("输入无效")
                    continue

                # append data to data set
                for i in range(int(start_row), int(end_row) + 1):
                    for j in range(int(start_col), int(end_col) + 1):
                        data_set.append(sheet.cell_value(i - 1, j - 1))
                sg.Popup("数据为：", data_set)

                # ask if keep adding
                command_add = sg.PopupYesNo("继续添加数据?")
                if command_add == "No":
                    window_data.close()
                    break

            # sys.exit the program
            elif event_data in ("退出", None):
                window_data.close()
                sys.exit()

        # calculate mean and variance
        mean = sum(data_set) / len(data_set)
        variance = sum([(i - mean) ** 2 for i in data_set]) / len(data_set)
        standard_deviation = np.sqrt(variance)

        # -----------------------------------------------------------------
        #                   Analysis interface
        # -----------------------------------------------------------------

        # print the following command in the interface
        layout_check = [
            [sg.Text(f"平均值: {round(mean, 4)}")],
            [sg.Text(f"方差: {round(variance, 4)}")],
            [sg.Text(f"标准差: {round(standard_deviation, 4)}")],
            [sg.Button("图像")],
            [sg.Button("概率计算")],
            [sg.Button("返回")],
            [sg.Button("退出")]
        ]

        window_check = sg.Window("检查", layout_check)
        while True:
            event_check, values_check = window_check.read()
            if event_check == "图像":
                x = np.linspace(mean - 3 * standard_deviation, mean + 3 * standard_deviation, 100)
                y = stats.norm.pdf(x, mean, standard_deviation)

                # provide different colors
                if cumulative_color % 6 == 1:
                    color = "r--"
                elif cumulative_color % 6 == 2:
                    color = "b--"
                elif cumulative_color % 6 == 3:
                    color = "g--"
                elif cumulative_color % 6 == 4:
                    color = "y--"
                else:
                    color = "c--"
                plt.plot(x, y, color)
                cumulative_color += 1
                plt.show()

            elif event_check == "概率计算":
                # start a new interface and ask for x1 and x2
                # -----------------------------------------------------------------
                #                   Give boundary interface
                # -----------------------------------------------------------------
                layout_possibility = [
                    [sg.Text("最低值："), sg.InputText(key="x1")],
                    [sg.Text("最高值："), sg.InputText(key="x2")],
                    [sg.Button("确认")],
                    [sg.Button("返回")],
                    [sg.Button("退出")]
                ]

                window_possibility = sg.Window("概率计算", layout_possibility)

                while True:
                    event_possibility, values_possibility = window_possibility.read()
                    if event_possibility == "确认":
                        # check if x1 and x2 are valid
                        try:
                            x1 = float(values_possibility["x1"])
                            x2 = float(values_possibility["x2"])
                        except ValueError:
                            sg.Popup("输入无效")
                            continue

                        possibility = abs((stats.norm.cdf(x2, mean, standard_deviation) - stats.norm.cdf(x1, mean,
                                                                                                         standard_deviation)
                                           ) * 100)

                        # print the possibility with precision to 4 decimal places
                        sg.Popup(f"概率: {round(possibility, 2)}%")

                    elif event_possibility == "返回":
                        window_possibility.close()
                        break

                    elif event_possibility in ("退出", None):
                        window_possibility.close()
                        sys.exit()

            elif event_check == "返回":
                excel.release_resources()
                window_check.close()
                break

            elif event_check in ("退出", None):
                window_check.close()
                sys.exit()