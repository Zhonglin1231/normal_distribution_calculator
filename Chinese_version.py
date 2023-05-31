import numpy as np
import matplotlib.pyplot as plt
import scipy.stats as stats
import PySimpleGUI as sg
import sys
import pandas as pd
from pylab import mpl
import seaborn as sns

sns.set()

# set the overall theme
sg.theme("LightBrown1")

# show the themes of PySimpleGUI
# sg.theme_previewer()


# set the specific text size in PySimpleGUI
sg.SetOptions(text_justification="center", font=("宋体", 15))

# add an icon to the window,
sg.SetOptions(icon='li_icon_1.ico')

# 设置显示中文字体
mpl.rcParams["font.sans-serif"] = ["SimHei"]

# fix the bug of Chinese in matplotlib
plt.rcParams["axes.unicode_minus"] = False


def alpha_transfer(trans_list):
    sum_col = 0
    length = len(trans_list)
    for i in trans_list:
        if 97 <= ord(i) <= 122:
            i = (ord(i) - 96) + 25 * (length - 1)
        elif 65 <= ord(i) <= 90:
            i = (ord(i) - 64) + 25 * (length - 1)
        else:
            return "error"
        length -= 1
        sum_col += i
    return sum_col


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


def change_color(cumulative_color):
    # provide different colors
    if cumulative_color % 6 == 1:
        color = "b"
    elif cumulative_color % 6 == 2:
        color = "g"
    elif cumulative_color % 6 == 3:
        color = "y"
    elif cumulative_color % 6 == 4:
        color = "r"
    else:
        color = "c"
    return color


def subplot_221(data_set, values_data, spot_size, color, mean):
    plt.subplot(2, 2, 1)
    plt.cla()
    # set the vision on y axis
    plt.ylim(0, mean*1.1)
    # set values of x and y
    x1 = np.linspace(1, len(data_set), len(data_set))
    y1 = data_set
    # draw the line
    plt.scatter(x1, y1, color=color, s=spot_size, alpha=0.5)
    # label the graph
    plt.title(values_data["名称"], size=20)
    plt.xlabel("芯片编号", size=15)
    plt.ylabel(values_data["x轴名称"], size=15)
    # give grid
    plt.grid(True)


def subplot_223(data_corrected, values_data, hist_pre):
    plt.subplot(2, 2, 3)
    plt.cla()

    # histogram
    # 剔除outlier
    data = pd.Series(data_corrected)  # 将数据由数组转换成series形式
    # plt.hist(data_corrected, density=True, color="b", edgecolor='w', label='直方图', bins=int(values_data["滑块"]))
    sns.distplot(data, hist=True, bins=int(hist_pre), color="r", label='正态分布曲线', fit=stats.norm)

    # let the pricision of the kde more precise
    # sns.kdeplot(data_set, linewidth=2)

    # label the graph
    plt.xlabel(values_data["x轴名称"], size=15)

    # subplot adjustment
    plt.subplots_adjust(left=None, bottom=None, right=None, top=None, wspace=None, hspace=0.5)
    # add grid to the graph
    plt.grid(True)


def subplot_224(data_set, values_data, hist_pre):
    plt.subplot(2, 2, 4)
    plt.cla()

    # histogram
    # 剔除outlier
    data = pd.Series(data_set)  # 将数据由数组转换成series形式
    # plt.hist(data_corrected, density=True, color="b", edgecolor='w', label='直方图', bins=int(values_data["滑块"]))
    sns.distplot(data, hist=True, bins=int(hist_pre), color="y", label='原始数据正态分布曲线', fit=stats.norm)

    # let the pricision of the kde more precise
    # sns.kdeplot(data_set, linewidth=2)

    # label the graph
    plt.xlabel(values_data["x轴名称"], size=15)

    # subplot adjustment
    plt.subplots_adjust(left=None, bottom=None, right=None, top=None, wspace=None, hspace=0.5)
    # add grid to the graph
    plt.grid(True)

    # show the graph, with icon
    plt.Figure()
    thismanager = plt.get_current_fig_manager()
    thismanager.window.wm_iconbitmap("li_icon_1.ico")
    plt.show()


def rapid_calculation_interface(file_name, cumulative_color):
    # -----------------------------------------------------------------
    #                   rapid calculation interface
    # -----------------------------------------------------------------
    while True:
        # initialize data set
        data_set = []
        header = 1
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
        count_size = 1
        proper_separation = [0]
        color = 'r'
        data_corrected = []
        sep_num = 0
        hist_pre = 100
        spot_size = 1
        csv = None
        graph_count = 0
        mean_rough = 0
        variance_rough = 0
        standard_deviation_rough = 0

        # use interface to choose which specific range of data in Excel to use
        # create layout_choose_data
        layout_data = [
            [sg.B("缩小"), sg.B("放大"),
             sg.B("全屏")],
            [sg.T("---------------------------1---------------------------", text_color="grey", key="1")],
            [sg.FileBrowse(button_text="新文件"), sg.In(key="新文件路径")],
            [sg.T("工作簿位置: ", key="工作簿"), sg.In(key="工作簿名称", size=(15, 1))],
            [sg.Text("文件待导入...", key="导入成功", text_color="red"), sg.B("导入文件")],
            [sg.T("---------------------------2---------------------------", text_color="grey", key="2")],
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
            [sg.T("---------------------------3---------------------------", text_color="grey", key="3")],
            # 告诉用户这是第几列数据
            [sg.Text(f"第{col}列数据为：", key="第几列")],
            [sg.Text(f"平均值: {round(mean, 4)}", key="平均值"), sg.Text(f"平方差: {round(variance, 4)}", key="平方差"),
             sg.Text(f"标准差: {round(standard_deviation, 4)}", key="标准差")],
            [sg.Button("生成图像"), sg.T(" 标题: "), sg.InputText(key="名称", size=(15, 1)),
             sg.T("x轴名称：", key="Tx轴名称"),
             sg.InputText(key="x轴名称", size=(15, 1))],
            [sg.Button("概率计算"), sg.Text("最低值：", key='T最低值'), sg.InputText(key="x1", size=(8, 1)),
             sg.Text("最高值：", key='T最高值'), sg.In(key="x2", size=(8, 1)), sg.T("概率：", key='T概率'),
             sg.T(f"{possibility}%", key="概率")],
            [sg.B('--'), sg.B('-'), sg.T(f"直方图精度: {hist_pre:>3}", key="精度"), sg.B('+'), sg.B('++')],
            [sg.B('<<'), sg.B('<'), sg.T(f" 散点大小: {spot_size:>3} ", key="散点"), sg.B('>'), sg.B('>>')],
            [sg.Button("退出")],
        ]

        window_data = sg.Window("力生美数据分析", layout_data, size=(800, 680), location=(370, 100),
                                element_justification='c', keep_on_top=True)

        while True:
            event_data, values_data = window_data.read()
            if event_data == "导入文件":
                file_name = values_data["新文件路径"]
                if file_name == "":
                    sg.popup("请选择文件！", keep_on_top=True)
                    continue
                else:
                    # choose which sheet to read
                    sheet_name = values_data["工作簿名称"]
                    if sheet_name == "":
                        sheet_name = 1
                    try:
                        sheet_name = int(sheet_name) - 1
                    except ValueError:
                        sg.popup("请输入数字！", keep_on_top=True)
                        continue

                    # read csv
                    if file_name[-4:] == ".csv":
                        header = 1
                        # auto find header
                        while True:
                            try:
                                print(header)
                                csv = pd.read_csv(file_name, sep=',', header=header - 1, skip_blank_lines=False)
                                print("1pass")
                                csv_2 = pd.read_csv(file_name, sep=',', header=header, skip_blank_lines=False)
                                print("2pass")

                                nrows, ncols = csv.shape
                                nrow_2, ncol_2 = csv_2.shape
                                print(ncols)
                                if ncols == ncol_2:
                                    break

                                header += 1

                            except pd.errors.ParserError:
                                header += 1

                    # read excel
                    elif file_name[-5:] == ".xlsx" or file_name[-4:] == ".xls" or file_name[-4:] == ".xlsm":
                        header = 1
                        # read the .xlsx or .xls or .xlsm file
                        try:
                            csv = pd.read_excel(file_name, header=header - 1, sheet_name=sheet_name)
                            nrows, ncols = csv.shape

                        except ValueError:
                            sg.Popup("查无此工作簿", keep_on_top=True)
                            continue

                    # decide if the format of the file is acceptable
                    else:
                        sg.Popup("文件格式不兼容！", keep_on_top=True)
                        continue

                    # update the range of rows and columns, in proper position, using fstring
                    text1 = f"行范围 ({header} ~ {nrows + header})"
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
                graph_count = 1
                data_set = []

                # transform the input to int
                # print(values_data["首列"])
                if values_data["首列"].isalpha():
                    # alphabet transfer to int
                    values_data["首列"] = alpha_transfer(values_data["首列"])
                    if values_data["首列"] == "error":
                        sg.Popup("请输入正确的列数", keep_on_top=True)
                        continue
                    col = values_data["首列"]
                if values_data["尾列"].isalpha():
                    # alphabet transfer to int
                    values_data["尾列"] = alpha_transfer(values_data["尾列"])
                    if values_data["尾列"] == "error":
                        sg.Popup("请输入正确的列数", keep_on_top=True)
                        continue

                # check if the input is valid
                try:
                    first_row = int(values_data["首行"])
                    last_row = (values_data["尾行"])
                    first_col = int(values_data["首列"])
                    last_col = (values_data["尾列"])
                except ValueError:
                    sg.Popup("数据类型无效", keep_on_top=True)
                    continue

                if first_row > nrows or first_col > ncols:
                    sg.Popup("输入无效", keep_on_top=True)
                    continue

                if last_row == '':
                    last_row = nrows

                if last_col == '':
                    last_col = first_col

                # append data to data_set
                for i in range(first_row - header, int(last_row) + 1 - header):
                    for j in range(first_col, int(last_col) + 1):

                        # ignore the blank data
                        if csv.iat[i - 1, j - 1] == "":
                            continue

                        # check if the data is string
                        elif type(csv.iat[i - 1, j - 1]) == str:
                            if '.' in csv.iat[i - 1, j - 1]:
                                csv.iat[i - 1, j - 1] = float(csv.iat[i - 1, j - 1])

                            else:
                                sg.Popup("数据异常", keep_on_top=True)
                                # close the window and restart
                                # print(csv.iat[i - 1, j - 1])
                                window_data.close()
                                return rapid_calculation_interface(file_name, cumulative_color)

                        data_set.append(csv.iat[i - 1, j - 1])

                data_set = [float(i) for i in data_set if str(i) != 'nan']

                # print(data_set)

                if len(data_set) == 0:
                    sg.Popup("请添加数据", keep_on_top=True)
                    # close the window and restart
                    continue

                # update 自定义完成
                window_data["确认自定义"].update("自定义完成！", text_color="green")

                # do calculation
                mean, variance, standard_deviation = calculation(data_set, window_data, values_data["首列"])
                mean_rough = mean
                variance_rough = variance
                standard_deviation_rough = standard_deviation
                data_corrected = [i for i in data_set if
                                  mean + 3 * standard_deviation > i > mean - 3 * standard_deviation]
                mean, variance, standard_deviation = calculation(data_corrected, window_data, values_data["首列"])

            # sys.exit the program
            if event_data in ("退出", None):
                window_data.close()
                sys.exit()

            if event_data == "生成图像":
                if not data_set:
                    sg.Popup("请添加数据", keep_on_top=True)
                    continue
                sep_num = 0
                # update 导入成功 & 自定义完成
                window_data["导入成功"].update("新文件待导入...", text_color="red")
                window_data["确认自定义"].update("自定义待确认...", text_color="red")

                # provide different colors
                color = change_color(cumulative_color)

                # set the size & title & position of the graph, give a icon to the figure
                plt.figure(values_data["名称"], figsize=(5, 10))
                mngr = plt.get_current_fig_manager()  # 获取当前figure manager
                mngr.window.wm_geometry("+0+0")  # 调整窗口在屏幕上弹出的位置

                # first graph
                subplot_221(data_set, values_data, spot_size, color, mean)

                # second graph
                plt.subplot(2, 2, 2)
                # set values of x and y
                # x2 = np.linspace(mean - 3 * standard_deviation, mean + 3 * standard_deviation, 100)
                # xmin, xmax = plt.xlim()
                # x2 = np.linspace(xmin, xmax, 100)
                x2 = np.linspace(mean - 3 * standard_deviation, mean + 3 * standard_deviation, 100)
                y2 = stats.norm.pdf(x2, loc=mean, scale=standard_deviation)
                # y2 = (1/(standard_deviation * ((2 * 3.141592653)**0.5))) * (2.718281828**(-(((x2 - mean)**2) / (
                # 2*variance))))

                x_label = values_data["x轴名称"]
                plt.xlabel(x_label, size=15)
                plt.ylabel("正态分布对比区", size=15)

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
                plt.xlabel(x_label, size=15)
                plt.title(values_data["名称"], size=20)

                plt.grid(True)
                # set the length of this subplot longer

                # third graph
                subplot_223(data_corrected, values_data, hist_pre)

                # fourth graph
                subplot_224(data_set, values_data, hist_pre)

                cumulative_color += 1

            if event_data == "概率计算":
                try:
                    x1 = float(values_data["x1"])
                    x2 = float(values_data["x2"])
                except ValueError:
                    sg.Popup("输入无效", keep_on_top=True)
                    continue

                possibility = abs(
                    (stats.norm.cdf(x2, mean, standard_deviation) - stats.norm.cdf(x1, mean, standard_deviation)) * 100)

                window_data["概率"].update(value=f"{round(possibility, 2)}%")

                possibility_rough = abs(
                    (stats.norm.cdf(x2, mean_rough, standard_deviation_rough) - stats.norm.cdf(x1, mean_rough,
                                                                                               standard_deviation_rough)) * 100)

                data = pd.Series(data_corrected)

                # change to the correct subplot
                plt.subplot(2, 2, 3)

                # ax = normal distribution curve
                ax = sns.distplot(data, hist=False, kde=False,
                                  kde_kws={'linewidth': 2},
                                  label="密度图", fit=stats.norm)

                # get the value of the area under the curve in certain range
                # Get all the lines used to draw density curve
                kde_lines = ax.get_lines()[-1]
                kde_x, kde_y = kde_lines.get_data()

                if values_data["x1"] != "" and values_data["x2"] != "":
                    mask = (kde_x > float(values_data["x1"])) & (kde_x < float(values_data["x2"]))
                    filled_x, filled_y = kde_x[mask], kde_y[mask]

                    # Shade the partial region
                    ax.fill_between(filled_x, y1=filled_y, alpha=0.5, color=color)

                    # Vertical lines for reference
                    plt.axvline(x=float(values_data["x1"]), linewidth=2, linestyle='--', color=color)
                    plt.axvline(x=float(values_data["x2"]), linewidth=2, linestyle='--', color=color)

                    area = np.trapz(filled_y, filled_x) * 100
                    plt.text(proper_separation[0],
                             stats.norm.pdf(mean, mean, standard_deviation) - sep_num * proper_separation[1],
                             f"{values_data['x1']} 和 {values_data['x2']} 之间的概率为: {area.round(4)}%", size=12,
                             color=color)


                    plt.subplot(2, 2, 4)
                    plt.text(proper_separation[0],
                             stats.norm.pdf(mean_rough, mean_rough, standard_deviation_rough) - (sep_num/3) * proper_separation[1],
                             f"{values_data['x1']} 和 {values_data['x2']} 之间的概率为: {round(possibility_rough, 2)}%", size=12,
                             color=color)

                    plt.show()
                    color = change_color(cumulative_color)
                    cumulative_color += 1
                    sep_num += 2

            elif event_data == "全屏":
                if count_size == 1:
                    # 使得窗口全屏
                    window_data.Maximize()

                elif count_size == -1:
                    # 使得窗口恢复
                    window_data.Normal()

                count_size *= -1

            elif event_data == "放大":
                window_data.Size = (window_data.Size[0] + 100, window_data.Size[1] + 100)

            elif event_data == "缩小":
                window_data.Size = (window_data.Size[0] - 100, window_data.Size[1] - 100)

            # change the preciseness of the graph
            elif event_data in ["+", "-", "++", "--"] and graph_count == 1:
                if event_data == "+":
                    hist_pre += 5

                elif event_data == "-":
                    if hist_pre > 10:
                        hist_pre -= 5
                    else:
                        hist_pre = 1

                elif event_data == "++":
                    hist_pre += 50

                elif event_data == "--":
                    if hist_pre > 50:
                        hist_pre -= 50
                    else:
                        hist_pre = 1

                window_data["精度"].update(f"直方图精度: {hist_pre:<3}")
                subplot_223(data_corrected, values_data, hist_pre)
                subplot_224(data_set, values_data, hist_pre)
                plt.show()

            # change spot size
            elif event_data in [">", "<", ">>", "<<"] and graph_count == 1:
                if event_data == ">":
                    spot_size += 1

                elif event_data == "<":
                    if spot_size > 1:
                        spot_size -= 1
                    else:
                        spot_size = 1

                elif event_data == ">>":
                    spot_size += 10

                elif event_data == "<<":
                    if spot_size > 10:
                        spot_size -= 10
                    else:
                        spot_size = 1
                        continue

                subplot_221(data_set, values_data, spot_size, color, mean)
                window_data["散点"].update(f"散点大小: {spot_size:<3}")
                plt.show()
