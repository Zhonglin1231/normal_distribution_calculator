import os
import numpy as np
import matplotlib.pyplot as plt
import scipy.stats as stats
import PySimpleGUI as sg
import sys
import pandas as pd
from pylab import mpl

# set the overall theme
sg.theme("LightBrown1")

# show the themes of PySimpleGUI
# sg.theme_previewer()


# set the specific text size in PySimpleGUI
sg.SetOptions(text_justification="center", font=("宋体", 15))

# 设置显示中文字体
mpl.rcParams["font.sans-serif"] = ["SimHei"]

# fix the bug of Chinese in matplotlib
plt.rcParams["axes.unicode_minus"] = False

figure_num = [1]


def choose_file_interface(cumulative_color):
    # -----------------------------------------------------------------
    #                   choose file interface
    # -----------------------------------------------------------------
    # create layout_choose_file
    # choose file by clicking file
    while True:
        layout_file = [
            [sg.FileBrowse(button_text="选择文件"), sg.In(key="文件名")],
            [sg.Button("自定义")],
            [sg.Button("快速分析")],
            [sg.Button("退出")]
        ]

        window_file = sg.Window("选择文件", layout_file, size=(800, 500), location=(350, 200), element_justification='c')

        while True:

            event_file, values_file = window_file.read()

            if event_file == "自定义":
                # data if file exists
                file_name = values_file["文件名"]
                data = os.path.isfile(file_name)
                if data == 0:
                    sg.Popup("文件未找到")
                else:
                    window_file.close()
                    calculation_interface(file_name, cumulative_color)

            elif event_file == "快速分析":
                # data if file exists
                file_name = values_file["文件名"]
                data = os.path.isfile(file_name)
                if data == 0:
                    sg.Popup("文件未找到")
                else:
                    window_file.close()
                    rapid_calculation_interface(file_name, cumulative_color)

            elif event_file in ("退出", None):
                window_file.close()
                sys.exit()

            break
            # calculate


def calculation_interface(file_name, cumulative_color):
    # -----------------------------------------------------------------
    #                   calculation interface
    # -----------------------------------------------------------------
    while True:
        # initialize data set
        data_set = []

        # read data from Excel
        # 文件加载中
        layout_loading = [
            [sg.Text("文件加载中", font=("华文行楷", 20))],
            [sg.Text("请稍等片刻", font=("华文行楷", 20))],
        ]

        window_loading = sg.Window("文件加载中", layout_loading, size=(800, 500), location=(350,200), element_justification='c')
        window_loading.read(timeout=50)

    # read file

    # customize the settings
        # pop up a window and ask for the value of header
        layout_header = [
            [sg.Text("请输入要定义的首行")],
            [sg.InputText(key="header")],
            [sg.Button("确认")],
        ]

        window_header = sg.Window("首行", layout_header, size=(800, 500), location=(350, 200), element_justification='c')
        event_header, values_header = window_header.read()
        header = int(values_header["header"])
        window_header.close()

        csv = pd.read_csv(file_name, sep=',', header=header)

        nrows, ncols = csv.shape

        window_loading.close()

        # use interface to choose which specific range of data in Excel to use
        # create layout_choose_data
        layout_data = [
            [sg.Text(f"行范围 (1 ~ {nrows})")],
            [sg.Text("首行："), sg.InputText(key="首行")],
            [sg.Text("尾行："), sg.InputText(key="尾行")],
            [sg.Text(f"列范围 (1 ~ {ncols})")],
            [sg.Text("首列："), sg.InputText(key="首列")],
            [sg.Text("尾列："), sg.InputText(key="尾列")],
            [sg.Button("确认")],
            [sg.Button("返回")],
            [sg.Button("退出")]
        ]

        window_data = sg.Window("选择数据", layout_data, size=(800, 500), location=(350, 200), element_justification='c')

        while True:
            event_data, values_data = window_data.read()
            # if End_row has no input, set it to sheet.nrows
            if values_data["尾行"] == "":
                values_data["尾行"] = nrows
            # if End_col has no input, set it to sheet.ncols
            if values_data["尾列"] == "":
                values_data["尾列"] = ncols

            if event_data == "确认":
                # data if the input is valid

                try:
                    start_row = int(values_data["首行"])
                    end_row = int(values_data["尾行"])
                    start_col = int(values_data["首列"])
                    end_col = int(values_data["尾列"])
                except ValueError:
                    sg.Popup("输入无效")
                    continue
                if start_row > end_row or start_col > end_col or end_row > nrows or end_col > ncols:
                    sg.Popup("输入无效")
                    continue



                # use csv.iat to append data in csv to data set
                for i in range(int(start_row), int(end_row) + 1):
                    for j in range(int(start_col), int(end_col) + 1):
                        # ignore the blank data
                        if csv.iat[i - 1, j - 1] == "":
                            continue

                        elif type(csv.iat[i - 1, j - 1]) == str:
                            sg.Popup("数据异常")
                            # close the window and restart
                            window_data.close()
                            return calculation_interface(file_name, cumulative_color)

                        data_set.append(csv.iat[i - 1, j - 1])

                # sg.popup_scrolled("数据为：", data_set)

                # ask if keep adding
                command_add = sg.PopupYesNo("继续添加数据?")
                if command_add == "No":
                    window_data.close()
                    break

            if event_data == "返回":
                window_data.close()
                return

            # sys.exit the program
            elif event_data in ("退出", None):
                window_data.close()
                sys.exit()

        # get the sum of the dataset, I don't want get nan
        data_set = [float(i) for i in data_set if str(i) != 'nan']

        # print(data_set)
        # calculate mean and variance
        mean = sum(data_set) / len(data_set)
        # uses the technique of point estimate
        variance = sum([(i - mean) ** 2 for i in data_set]) / len(data_set)
        standard_deviation = np.sqrt(variance)

        # analysis_interface(mean, variance, standard_deviation, cumulative_color, csv)


def rapid_calculation_interface(file_name, cumulative_color):
    # -----------------------------------------------------------------
    #                   rapid calculation interface
    # -----------------------------------------------------------------
    while True:
        # initialize data set
        data_set = []
        header = 38
        mean = 1
        variance = 1
        standard_deviation = 1
        possibility = None
        col = 0

        csv = pd.read_csv(file_name, sep=',', header=header)
        nrows, ncols = csv.shape

        # use interface to choose which specific range of data in Excel to use
        # create layout_choose_data
        layout_data = [
            [sg.Text(f"列范围 (1 ~ {ncols})"), sg.Text("列："), sg.InputText(key="列"), sg.Button("加载列")],
            # 告诉用户这是第几列数据
            [sg.Text(f"第{col}列数据为：", key="第几列")],
            [sg.Text(f"平均值: {round(mean, 4)}", key="平均值")],
            [sg.Text(f"平方差: {round(variance, 4)}", key="平方差")],
            [sg.Text(f"标准差: {round(standard_deviation, 4)}", key="标准差")],
            [sg.Button("图像"), sg.InputText(key="名称", size=(15, 1)), sg.T("x轴名称："),
             sg.InputText(key="x轴名称", size=(15, 1))],
            [sg.Button("概率计算"), sg.Text("最低值："), sg.InputText(key="x1", size=(8, 1)),
             sg.Text("最高值："), sg.In(key="x2", size=(8, 1)), sg.T("概率："), sg.T(f"{possibility}%", key="概率")],
            [sg.Button("新文件")],
            [sg.Button("退出")]
        ]

        window_data = sg.Window("选择数据", layout_data, size=(800, 500), location=(350, 200), element_justification='c')

        while True:
            while True:
                event_data, values_data = window_data.read()
                if event_data == "加载列":
                    # data if the input is valid
                    try:
                        col = int(values_data["列"])
                    except ValueError:
                        sg.Popup("输入无效")
                        continue
                    if col > ncols:
                        sg.Popup("输入无效")
                        continue

                    # use csv.iat to append data in csv to data set
                    for i in range(1, int(nrows) + 1):
                        # ignore the blank data
                        if csv.iat[i - 1, col - 1] == "":
                            continue

                        elif type(csv.iat[i - 1, col - 1]) == str:
                            sg.Popup("数据异常")
                            # close the window and restart
                            window_data.close()
                            return rapid_calculation_interface(file_name, cumulative_color)

                        data_set.append(csv.iat[i - 1, col - 1])
                    # sg.popup_scrolled("数据为：", data_set)
                    data_set = [float(i) for i in data_set if str(i) != 'nan']

                    print(data_set)
                    # calculate mean and variance
                    mean = sum(data_set) / len(data_set)
                    # uses the technique of point estimate
                    variance = sum([(i - mean) ** 2 for i in data_set]) / len(data_set)
                    standard_deviation = np.sqrt(variance)

                    # update all elements in the window
                    window_data["第几列"].update(f"第{col}列数据为：")
                    window_data["平均值"].update(f"平均值: {round(mean, 4)}")
                    window_data["平方差"].update(f"平方差: {round(variance, 4)}")
                    window_data["标准差"].update(f"标准差: {round(standard_deviation, 4)}")
                    data_set = []

                if event_data == "新文件":
                    window_data.close()
                    return

                # sys.exit the program
                elif event_data in ("退出", None):
                    window_data.close()
                    sys.exit()

                # get the sum of the dataset, I don't want get nan

                if event_data == "图像":
                    x_label = values_data["x轴名称"]
                    figure_num[0] += 1
                    plt.figure(values_data["名称"])
                    plt.title(values_data["名称"], size=20)
                    plt.xlabel(x_label, size=15)
                    plt.ylabel("概率密度", size=15)
                    x = np.linspace(mean - 3 * standard_deviation, mean + 3 * standard_deviation, 100)
                    y = stats.norm.pdf(x, mean, standard_deviation)

                    # display the value of the mean, variance and standard deviation on the graph on appropriate position
                    proper_separation = [mean + standard_deviation, (1 / (3 * standard_deviation)) / 10]
                    plt.text(proper_separation[0], stats.norm.pdf(mean, mean, standard_deviation),
                             f"平均值 = {round(mean, 4)}", size=12)
                    plt.text(proper_separation[0],
                             stats.norm.pdf(mean, mean, standard_deviation) - proper_separation[1],
                             f"平方差 = {round(variance, 4)}", size=12)
                    plt.text(proper_separation[0],
                             stats.norm.pdf(mean, mean, standard_deviation) - 2 * proper_separation[1],
                             f"标准差 = {round(standard_deviation, 4)}", size=12)

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

                elif event_data == "概率计算":
                    try:
                        x1 = float(values_data["x1"])
                        x2 = float(values_data["x2"])
                    except ValueError:
                        sg.Popup("输入无效")
                        continue

                    possibility = abs((stats.norm.cdf(x2, mean, standard_deviation) - stats.norm.cdf(x1, mean,
                                                                                                     standard_deviation)) * 100)

                    window_data["概率"].update(
                        value=f"{round(possibility, 4)}%"
                    )


        # analysis_interface(mean, variance, standard_deviation, cumulative_color, csv)

'''
def analysis_interface(mean, variance, standard_deviation, cumulative_color, csv):
    # -----------------------------------------------------------------
    #                   Analysis interface
    # -----------------------------------------------------------------
    # print the following command in the interface
    possibility = 0.0

    layout_data = [
        [sg.Text(f"平均值: {round(mean, 4)}")],
        [sg.Text(f"平方差: {round(variance, 4)}")],
        [sg.Text(f"标准差: {round(standard_deviation, 4)}")],
        [sg.Button("图像"), sg.InputText(key="名称", size=(15, 1)), sg.T("x轴名称："),
         sg.InputText(key="x轴名称", size=(15, 1))],
        [sg.Button("概率计算"), sg.Text("最低值："), sg.InputText(key="x1", size=(8, 1)),
         sg.Text("最高值："), sg.In(key="x2", size=(8, 1)), sg.T("概率："), sg.T(f"{possibility}%", key="概率")],
        [sg.Button("返回")],
        [sg.Button("退出")]
    ]

    window_data = sg.Window("检查", layout_data, size=(800, 500), location=(350, 200), element_justification='c')

    while True:
        event_data, values_data = window_data.read()

        



        elif event_data == "返回":
            window_data.close()
            return

        elif event_data in ("退出", None):
            window_data.close()
            sys.exit() '''
