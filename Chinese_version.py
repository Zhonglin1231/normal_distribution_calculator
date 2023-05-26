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


def calculation(data_set, window_data, col):
    # use csv.iat to append data in csv to data set

    mean = sum(data_set) / len(data_set)
    # uses the technique of point estimate
    variance = sum([(i - mean) ** 2 for i in data_set]) / len(data_set)
    standard_deviation = np.sqrt(variance)

    # update all elements in the window
    window_data["第几列"].update(f"第{col}列数据为：")
    window_data["平均值"].update(f"平均值: {round(mean, 4):<8}")
    window_data["平方差"].update(f"平方差: {round(variance, 4):<8}")
    window_data["标准差"].update(f"标准差: {round(standard_deviation, 4):<8}")

    return mean, variance, standard_deviation


def rapid_calculation_interface(file_name, cumulative_color):
    # -----------------------------------------------------------------
    #                   rapid calculation interface
    # -----------------------------------------------------------------
    global csv
    while True:
        # initialize data set
        data_set = []
        header = 100
        mean = 1
        variance = 1
        standard_deviation = 1
        possibility = None
        col = 0
        nrows, ncols = 0, 0
        bool_row_num = -1
        bool_row_pow = 2
        bool_col_num = -1
        bool_col_pow = 2

        # use interface to choose which specific range of data in Excel to use
        # create layout_choose_data
        layout_data = [
            [sg.Text("请先输入文件表头行数：", text_color='red'), sg.InputText(key="表头行数", size=(10, 1))],
            [sg.FileBrowse(button_text="新文件"), sg.In(key="新文件路径")],
            [sg.Text("文件待导入...", key="导入成功", text_color="red"), sg.B("导入文件")],
            # 自定义版面
            [sg.Text(f"行范围 (1 ~ {nrows})", key="行范围"),
             sg.Text("首行：", key="首行T"),
             sg.InputText(key="首行", size=(10, 1)), sg.B("尾行：", key="尾行T"),
             sg.InputText(key="尾行", visible=False, size=(10, 1))],
            [sg.Text(f"列范围 (1 ~ {ncols})", key="列范围"),
             sg.Text("首列：", key="首列T"),
             sg.InputText(key="首列", size=(10, 1)), sg.B("尾列：", key="尾列T"),
             sg.InputText(key="尾列", visible=False, size=(10, 1))],
            [sg.T("自定义待确认...", key="确认自定义", text_color="red"), sg.Button("完成自定义")],

            # 告诉用户这是第几列数据
            [sg.Text(f"第{col}列数据为：", key="第几列", visible=False)],
            [sg.Text(f"平均值: {round(mean, 4)}", key="平均值")],
            [sg.Text(f"平方差: {round(variance, 4)}", key="平方差")],
            [sg.Text(f"标准差: {round(standard_deviation, 4)}", key="标准差")],
            [sg.Button("生成图像"), sg.T("标题: "), sg.InputText(key="名称", size=(15, 1)), sg.T("x轴名称："),
             sg.InputText(key="x轴名称", size=(15, 1))],
            [sg.Button("概率计算"), sg.Text("最低值："), sg.InputText(key="x1", size=(8, 1)),
             sg.Text("最高值："), sg.In(key="x2", size=(8, 1)), sg.T("概率："), sg.T(f"{possibility}%", key="概率")],
            [sg.Button("退出")]
        ]

        window_data = sg.Window("选择数据", layout_data, size=(800, 500), location=(350, 200),
                                element_justification='c')

        while True:
            event_data, values_data = window_data.read()

            if event_data == "导入文件":
                file_name = values_data["新文件路径"]
                if file_name == "":
                    sg.popup("请选择文件！")
                    continue
                else:

                    # read csv
                    if file_name[-4:] == ".csv":
                        header = values_data["表头行数"]
                        if header == "":
                            sg.Popup("请输入表头行数！", text_color="red")
                            window_data.close()
                            return rapid_calculation_interface(file_name, cumulative_color)
                        header = int(values_data["表头行数"])

                        csv = pd.read_csv(file_name, sep=',', header=header, skip_blank_lines=False)
                        nrows, ncols = csv.shape

                        # update the range of rows and columns, in proper position, using fstring
                        text1 = f"行范围 (1 ~ {nrows})"
                        text2 = f"列范围 (1 ~ {ncols})"
                        window_data["导入成功"].update("导入成功！", text_color="green")
                        window_data["行范围"].update(f"{text1:<15}")
                        window_data["列范围"].update(f"{text2:<15}")

                    # read excel
                    else:
                        csv = pd.read_excel(file_name)
                        nrows, ncols = csv.shape

                        # update the range of rows and columns, in proper position, using fstring
                        text1 = f"行范围 (1 ~ {nrows})"
                        text2 = f"列范围 (1 ~ {ncols})"
                        window_data["导入成功"].update("导入成功！", text_color="green")
                        window_data["行范围"].update(f"{text1:<15}")
                        window_data["列范围"].update(f"{text2:<15}")


            # 自定义尾行尾列
            tf_row = int(abs((bool_row_num ** bool_row_pow - 1)) / 2) == False
            tf_col = int(abs((bool_col_num ** bool_col_pow - 1)) / 2) == False
            if event_data == "尾行T":
                window_data["尾行"].update(visible=tf_row)
                bool_row_pow += 1
            if event_data == "尾列T":
                window_data["尾列"].update(visible=tf_col)
                bool_col_pow += 1

            if event_data == "完成自定义":
                data_set = []
                # check if the input is valid
                try:
                    first_row = int(values_data["首行"])
                    last_row = values_data["尾行"]
                    first_col = int(values_data["首列"])
                    last_col = values_data["尾列"]
                except ValueError:
                    sg.Popup("数据类型无效")
                    continue

                if first_row > nrows or first_col > ncols:
                    sg.Popup("输入无效")
                    continue

                if last_row == '':
                    last_row = nrows

                if last_col == '':
                    last_col = first_col

                # append data to data_set
                for i in range(first_row - 41, last_row + 1):
                    for j in range(first_col, last_col + 1):

                        # ignore the blank data
                        if csv.iat[i - 1, j - 1] == "":
                            continue

                        # check if the data is string
                        elif type(csv.iat[i - 1, j - 1]) == str:
                            if '.' in csv.iat[i - 1, j - 1]:
                                csv.iat[i - 1, j - 1] = float(csv.iat[i - 1, j - 1])
                            else:
                                sg.Popup("数据异常")
                                print(csv.iat[i - 1, j - 1])
                                window_data.close()
                                return rapid_calculation_interface(file_name, cumulative_color)

                            # close the window and restart


                        data_set.append(csv.iat[i - 1, j - 1])

                data_set = [float(i) for i in data_set if str(i) != 'nan']

                print(data_set)

                if len(data_set) == 0:
                    sg.Popup("请添加数据")
                    # close the window and restart
                    continue

                # update 自定义完成
                window_data["确认自定义"].update("自定义完成！", text_color="green")

                # do calculation
                mean, variance, standard_deviation = calculation(data_set, window_data, col)

            # sys.exit the program
            if event_data in ("退出", None):
                window_data.close()
                sys.exit()

            # get the sum of the dataset, I don't want get nan

            if event_data == "生成图像":
                # update 导入成功 & 自定义完成
                window_data["导入成功"].update("新文件待导入...", text_color="red")
                window_data["确认自定义"].update("自定义待确认...", text_color="red")

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
                plt.scatter(x1, y1, color=color, s=0.5)
                # label the graph
                plt.title("各芯片数据", size=20)
                plt.xlabel("芯片编号", size=15)
                plt.ylabel("测量值", size=15)
                # give grid
                plt.grid(True)

                # second graph
                plt.subplot(2, 1, 2)
                # set values of x and y
                x2 = np.linspace(mean - 3 * standard_deviation, mean + 3 * standard_deviation, 100)
                y2 = stats.norm.pdf(x2, mean, standard_deviation)

                # display the value of the mean, variance and standard deviation on graph on appropriate position
                proper_separation = [mean + standard_deviation, (1 / (3 * standard_deviation)) / 10]
                plt.text(proper_separation[0], stats.norm.pdf(mean, mean, standard_deviation),
                         f"平均值 = {round(mean, 4)}", size=12, color=color)
                plt.text(proper_separation[0],
                         stats.norm.pdf(mean, mean, standard_deviation) - proper_separation[1],
                         f"平方差 = {round(variance, 4)}", size=12, color=color)
                plt.text(proper_separation[0],
                         stats.norm.pdf(mean, mean, standard_deviation) - 2 * proper_separation[1],
                         f"标准差 = {round(standard_deviation, 4)}", size=12, color=color)

                # draw the line
                plt.plot(x2, y2, color)
                # label the graph
                x_label = values_data["x轴名称"]
                plt.title(values_data["名称"], size=20)
                plt.xlabel(x_label, size=15)
                plt.ylabel("概率密度", size=15)

                cumulative_color += 1

                # subplot adjustment
                plt.subplots_adjust(left=None, bottom=None, right=None, top=None, wspace=None, hspace=0.5)
                # add grid to the graph
                plt.grid(True)

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
