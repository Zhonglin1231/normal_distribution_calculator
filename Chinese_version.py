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


def calculation_interface(file_name, cumulative_color):
    # -----------------------------------------------------------------
    #                   calculation interface
    # -----------------------------------------------------------------
    while True:
        # initialize data set
        data_set = []

        # read file

        # customize the settings
        # pop up a window and ask for the value of header
        layout_header = [
            [sg.Text("请输入要定义的首行")],
            [sg.InputText(key="header")],
            [sg.Button("确认")],
        ]

        window_header = sg.Window("首行", layout_header, size=(800, 500), location=(350, 200),
                                  element_justification='c')
        event_header, values_header = window_header.read()
        header = int(values_header["header"])
        window_header.close()

        csv = pd.read_csv(file_name, sep=',', header=header)

        nrows, ncols = csv.shape

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

        window_data = sg.Window("选择数据", layout_data, size=(800, 500), location=(350, 200),
                                element_justification='c')

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
    global csv
    while True:
        # initialize data set
        data_set = []
        header = 38
        mean = 1
        variance = 1
        standard_deviation = 1
        possibility = None
        col = 0

        nrows, ncols = 0, 0

        # use interface to choose which specific range of data in Excel to use
        # create layout_choose_data
        layout_data = [
            [sg.FileBrowse(button_text="新文件"), sg.In(key="新文件路径")],
            [sg.Button("快速分析"), sg.Button("自定义")],
            [sg.T("请提供列值...", key="请提供列值", visible=False),
             sg.T("自定义中...", key="自定义中", visible=False)],
            [sg.Text(f"列范围 (1 ~ {ncols})", key="列范围"), sg.Text("列："), sg.InputText(key="列"),
             sg.Button("加载列")],
            # 告诉用户这是第几列数据
            [sg.Text(f"第{col}列数据为：", key="第几列")],
            [sg.Text(f"平均值: {round(mean, 4)}", key="平均值")],
            [sg.Text(f"平方差: {round(variance, 4)}", key="平方差")],
            [sg.Text(f"标准差: {round(standard_deviation, 4)}", key="标准差")],
            [sg.Button("图像"), sg.InputText(key="名称", size=(15, 1)), sg.T("x轴名称："),
             sg.InputText(key="x轴名称", size=(15, 1))],
            [sg.Button("概率计算"), sg.Text("最低值："), sg.InputText(key="x1", size=(8, 1)),
             sg.Text("最高值："), sg.In(key="x2", size=(8, 1)), sg.T("概率："), sg.T(f"{possibility}%", key="概率")],
            [sg.Button("退出")]
        ]

        window_data = sg.Window("选择数据", layout_data, size=(800, 500), location=(350, 200),
                                element_justification='c')

        while True:
            event_data, values_data = window_data.read()
            if event_data == "快速分析":
                file_name = values_data["新文件路径"]
                data = os.path.isfile(file_name)
                if data == 0:
                    sg.Popup("文件未找到")
                else:
                    csv = pd.read_csv(file_name, sep=',', header=header)
                    nrows, ncols = csv.shape
                    window_data["自定义中"].update(visible=False)
                    window_data["请提供列值"].update("请提供列值...", visible=True)
                    window_data["列范围"].update(f"列范围 (1 ~ {ncols})")

            elif event_data == "自定义":
                # data if file exists
                file_name = values_data["新文件路径"]
                data = os.path.isfile(file_name)
                if data == 0:
                    sg.Popup("文件未找到")
                else:
                    calculation_interface(file_name, cumulative_color)
                    window_data["自定义中"].update("自定义中...", visible=True)
                    window_data["请提供列值"].update("请提供列值...", visible=False)

            if event_data == "加载列":
                data_set = []
                # data if the input is valid
                try:
                    col = int(values_data["列"])
                except ValueError:
                    sg.Popup("输入无效")
                    continue
                if col > ncols:
                    sg.Popup("输入无效")
                    continue

                # update 请提供列值
                window_data["请提供列值"].update("请提供列值...", visible=False)

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
                if len(data_set) == 0:
                    sg.Popup("请添加数据")
                    # close the window and restart
                    continue

                mean = sum(data_set) / len(data_set)
                # uses the technique of point estimate
                variance = sum([(i - mean) ** 2 for i in data_set]) / len(data_set)
                standard_deviation = np.sqrt(variance)

                # update all elements in the window
                window_data["第几列"].update(f"第{col}列数据为：")
                window_data["平均值"].update(f"平均值: {round(mean, 4)}")
                window_data["平方差"].update(f"平方差: {round(variance, 4)}")
                window_data["标准差"].update(f"标准差: {round(standard_deviation, 4)}")

            # sys.exit the program
            elif event_data in ("退出", None):
                window_data.close()
                sys.exit()

            # get the sum of the dataset, I don't want get nan

            if event_data == "图像":
                # provide different colors
                if cumulative_color % 6 == 1:
                    color = "r"
                elif cumulative_color % 6 == 2:
                    color = "b"
                elif cumulative_color % 6 == 3:
                    color = "g"
                elif cumulative_color % 6 == 4:
                    color = "y"
                else:
                    color = "c"

                # set the size & title & position of the graph
                plt.figure(values_data["名称"], figsize=(5, 10))
                mngr = plt.get_current_fig_manager()  # 获取当前figure manager
                mngr.window.wm_geometry("+0+0")  # 调整窗口在屏幕上弹出的位置

                # first graph
                plt.subplot(2, 1, 1)

                # set values of x and y
                x1 = np.linspace(1, len(data_set), len(data_set))
                y1 = data_set
                # draw the line
                plt.scatter(x1, y1, color="blue", s=0.5)
                # label the graph
                plt.title("各芯片数据", size=20)
                plt.xlabel("芯片编号", size=15)
                plt.ylabel("测量值", size=15)

                # second graph
                plt.subplot(2, 1, 2)
                # set values of x and y
                x2 = np.linspace(mean - 3 * standard_deviation, mean + 3 * standard_deviation, 100)
                y2 = stats.norm.pdf(x2, mean, standard_deviation)

                # display the value of the mean, variance and standard deviation on graph on appropriate position
                proper_separation = [mean + standard_deviation, (1 / (3 * standard_deviation)) / 10]
                plt.text(proper_separation[0], stats.norm.pdf(mean, mean, standard_deviation),
                         f"平均值 = {round(mean, 4)}", size=12)
                plt.text(proper_separation[0],
                         stats.norm.pdf(mean, mean, standard_deviation) - proper_separation[1],
                         f"平方差 = {round(variance, 4)}", size=12)
                plt.text(proper_separation[0],
                         stats.norm.pdf(mean, mean, standard_deviation) - 2 * proper_separation[1],
                         f"标准差 = {round(standard_deviation, 4)}", size=12)

                # draw the line
                plt.plot(x2, y2, color)
                # label the graph
                x_label = values_data["x轴名称"]
                plt.title(values_data["名称"], size=20)
                plt.xlabel(x_label, size=15)
                plt.ylabel("概率密度", size=15)

                cumulative_color += 1

                # show the graph
                plt.show()

                figure_num[0] += 1

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
