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


def collapse(layout, key):
    return sg.pin(sg.Column(layout, key=key))


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
    plt.title("预处理数据")
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
    plt.title("原始数据")
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
    plt.subplots_adjust(left=None, bottom=None, right=None, top=None, wspace=None, hspace=0.3)
    # add grid to the graph
    plt.grid(True)

    # show the graph, with icon
    plt.Figure()
    thismanager = plt.get_current_fig_manager()
    thismanager.window.wm_iconbitmap("li_icon_1.ico")


def global_update(window_data, font, font_size):
    # update all elements in the window
    window_data["缩小"].Widget.config(font=f'{font} {font_size}')
    window_data["放大"].Widget.config(font=f'{font} {font_size}')
    window_data["全屏"].Widget.config(font=f'{font} {font_size}')
    window_data["-12"].Widget.config(font=f'{font} {font_size}')
    window_data["新文件"].Widget.config(font=f'{font} {font_size}')
    window_data["新文件路径"].Widget.config(font=f'{font} {font_size}')
    window_data["工作簿"].Widget.config(font=f'{font} {font_size}')
    window_data["工作簿名称"].Widget.config(font=f'{font} {font_size}')
    window_data["导入成功"].Widget.config(font=f'{font} {font_size}')
    window_data["导入文件"].Widget.config(font=f'{font} {font_size}')
    window_data["-22"].Widget.config(font=f'{font} {font_size}')
    window_data["行范围"].Widget.config(font=f'{font} {font_size}')
    window_data["首行T"].Widget.config(font=f'{font} {font_size}')
    window_data["首行"].Widget.config(font=f'{font} {font_size}')
    window_data["尾行T"].Widget.config(font=f'{font} {font_size}')
    window_data["尾行"].Widget.config(font=f'{font} {font_size}')
    window_data["列范围"].Widget.config(font=f'{font} {font_size}')
    window_data["首列T"].Widget.config(font=f'{font} {font_size}')
    window_data["首列"].Widget.config(font=f'{font} {font_size}')
    window_data["尾列T"].Widget.config(font=f'{font} {font_size}')
    window_data["尾列"].Widget.config(font=f'{font} {font_size}')
    window_data["确认自定义"].Widget.config(font=f'{font} {font_size}')
    window_data["完成自定义"].Widget.config(font=f'{font} {font_size}')
    window_data["-3"].Widget.config(font=f'{font} {font_size}')
    window_data["第几列"].Widget.config(font=f'{font} {font_size}')
    window_data["平均值"].Widget.config(font=f'{font} {font_size}')
    window_data["平方差"].Widget.config(font=f'{font} {font_size}')
    window_data["标准差"].Widget.config(font=f'{font} {font_size}')
    window_data["生成图像"].Widget.config(font=f'{font} {font_size}')
    window_data["名称"].Widget.config(font=f'{font} {font_size}')
    window_data["Tx轴名称"].Widget.config(font=f'{font} {font_size}')
    window_data["x轴名称"].Widget.config(font=f'{font} {font_size}')
    window_data["概率计算"].Widget.config(font=f'{font} {font_size}')
    window_data["T最低值"].Widget.config(font=f'{font} {font_size}')
    window_data["x1"].Widget.config(font=f'{font} {font_size}')
    window_data["T最高值"].Widget.config(font=f'{font} {font_size}')
    window_data["x2"].Widget.config(font=f'{font} {font_size}')
    window_data["T概率"].Widget.config(font=f'{font} {font_size}')
    window_data["概率"].Widget.config(font=f'{font} {font_size}')
    window_data["--"].Widget.config(font=f'{font} {font_size}')
    window_data["-"].Widget.config(font=f'{font} {font_size}')
    window_data["精度"].Widget.config(font=f'{font} {font_size}')
    window_data["+"].Widget.config(font=f'{font} {font_size}')
    window_data["++"].Widget.config(font=f'{font} {font_size}')
    window_data["<<"].Widget.config(font=f'{font} {font_size}')
    window_data["<"].Widget.config(font=f'{font} {font_size}')
    window_data["散点"].Widget.config(font=f'{font} {font_size}')
    window_data[">"].Widget.config(font=f'{font} {font_size}')
    window_data[">>"].Widget.config(font=f'{font} {font_size}')
    window_data["空格1"].Widget.config(font=f'{font} {font_size}')
    window_data["T标题"].Widget.config(font=f'{font} {font_size}')
    window_data["字体+"].Widget.config(font=f'{font} {font_size}')
    window_data["字体-"].Widget.config(font=f'{font} {font_size}')




def rapid_calculation_interface(file_name, cumulative_color):
    # -----------------------------------------------------------------
    #                   rapid calculation interface
    # -----------------------------------------------------------------
    while True:
        # initialize data set
        data_set = [[1]]
        header = []
        mean = [1]
        variance = [1]
        standard_deviation = [1]
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
        data_corrected = [[1]]
        sep_num = 0
        hist_pre = 100
        spot_size = 1
        csv = None
        graph_count = 0
        mean_rough = [1]
        variance_rough = [1]
        standard_deviation_rough = [1]
        last_col = 0
        first_col = 0
        opened1 = True
        opened2 = True
        font_size = 12
        font = "微软雅黑"

        # use interface to choose which specific range of data in Excel to use
        # create layout_choose_data
        SIMBLE_RIGHT = "\u25B6"
        SIMBLE_DOWN = "\u25BC"
        line_1 = f"---------------------------{SIMBLE_DOWN}---------------------------"
        line_2 = f"---------------------------{SIMBLE_RIGHT}---------------------------"


        layout_1 = [
            [sg.FileBrowse(button_text="新文件", key="新文件", font=("微软雅黑", font_size)), sg.In(key="新文件路径")],
            [sg.T("       工作簿编号: ", key="工作簿"), sg.In(key="工作簿名称", size=(5, 1))],
            [sg.Text("       文件待导入...", key="导入成功", text_color="red"), sg.B("导入文件")]
        ]

        layout_2 = [
            [sg.Text(f"行范围 (1 ~ {nrows})", key="行范围"),
             sg.Text("首行：", key="首行T"),
             sg.InputText(key="首行", size=(10, 1)), sg.B("尾行：", key="尾行T"),
             sg.InputText(key="尾行", visible=False, size=(10, 1))],
            [sg.Text(f"列范围 (1 ~ {ncols})", key="列范围"),
             sg.Text("首列：", key="首列T"),
             sg.InputText(key="首列", size=(10, 1)), sg.B("尾列：", key="尾列T"),
             sg.InputText(key="尾列", visible=False, size=(10, 1))],
            [sg.T("自定义待确认...", key="确认自定义", text_color="red"), sg.Button("完成自定义")]
        ]

        layout_data = [
            [sg.B("缩小"), sg.B("放大"), sg.B("全屏"), sg.T("字体-", text_color="grey", key="字体-", enable_events=True, relief="sunken"),
             sg.T("字体+", text_color="grey", key="字体+", enable_events=True, relief="sunken")],
            [sg.T(line_1, text_color="grey", key="-12", enable_events=True)],
            [collapse(layout_1, "fold_1")],
            # 自定义版面
            [sg.T(line_1, text_color="grey", key="-22", enable_events=True)],
            [collapse(layout_2, "fold_2")],
            [sg.T("---------------------------3---------------------------", text_color="grey", key="-3")],
            # 告诉用户这是第几列数据
            [sg.Text(f"第{col}列数据为：", key="第几列", font=("微软雅黑", 12))],
            [sg.Text(f"平均值: {round(mean[0], 4)}", key="平均值", font=("微软雅黑", 12)), sg.Text(f"平方差: {round(variance[0], 4)}", key="平方差", font=("微软雅黑", 12)),
             sg.Text(f"标准差: {round(standard_deviation[0], 4)}", key="标准差", font=("微软雅黑", 12))],
            [sg.Button("生成图像"), sg.T(" 标题:", key="T标题"), sg.InputText(key="名称", size=(10, 1)),
             sg.T("x轴名称:", key="Tx轴名称"),
             sg.InputText(key="x轴名称", size=(10, 1)), sg.T("        ", key="空格1")],
            [sg.Button("概率计算"), sg.Text("最低值:", key='T最低值'), sg.InputText(key="x1", size=(8, 1)),
             sg.Text("最高值:", key='T最高值'), sg.In(key="x2", size=(8, 1)), sg.T("概率:", key='T概率'),
             sg.T(f"{possibility}%", key="概率")],
            [sg.B('--'), sg.B('-'), sg.T(f"直方图精度: {hist_pre:>3}", key="精度"), sg.B('+'), sg.B('++')],
            [sg.B('<<'), sg.B('<'), sg.T(f" 散点大小: {spot_size:>3} ", key="散点"), sg.B('>'), sg.B('>>')]
        ]

        window_data = sg.Window("力生美数据分析", layout_data, size=(800, 600), location =(370, 100),
                                element_justification='c', keep_on_top=True)

        while True:
            event_data, values_data = window_data.read()
            if event_data == "-12":
                opened1 = not opened1
                window_data["-12"].update(line_1 if opened1 else line_2)
                window_data["fold_1"].update(visible=opened1)

            if event_data == "-22":
                opened2 = not opened2
                window_data["-22"].update(line_1 if opened2 else line_2)
                window_data["fold_2"].update(visible=opened2)

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
                    window_data["导入成功"].update("       导入成功！", text_color="green")
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

                # append number of cols to data_set
                for i in range(first_col, int(last_col) + 1):
                    data_set.append([])
                    mean.append([])
                    variance.append([])
                    standard_deviation.append([])
                    data_corrected.append([])
                    mean_rough.append([])
                    variance_rough.append([])
                    standard_deviation_rough.append([])

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
                        # data_set is a 2D list
                        data_set[j-first_col].append(csv.iat[i - 1, j - 1])

                for k in range(len(data_set)):
                    data_set[k] = [float(i) for i in data_set[k] if str(i) != 'nan']

                # print(data_set)

                if len(data_set) == 0:
                    sg.Popup("请添加数据", keep_on_top=True)
                    # close the window and restart
                    continue

                # update 自定义完成
                window_data["确认自定义"].update("自定义完成！", text_color="green")

                # do calculation
                for i in range(len(data_set)):
                    if len(data_set[i]) == 0:
                        continue
                    mean[i], variance[i], standard_deviation[i] = calculation(data_set[i], window_data, str(int(values_data["首列"])+i))
                    mean_rough[i] = mean[i]
                    variance_rough[i] = variance[i]
                    standard_deviation_rough[i] = standard_deviation[i]
                    # delete the data in data[i] if the data is larger or smaller than 3*standard deviation from mean
                    data_corrected[i] = [j for j in data_set[i] if mean[i] - 3 * standard_deviation[i] < j < mean[i] + 3 * standard_deviation[i]]
                    mean[i], variance[i], standard_deviation[i] = calculation(data_corrected[i], window_data, str(int(values_data["首列"])+i))

            # sys.exit the program
            if event_data in ("退出", None):
                window_data.close()
                sys.exit()

            if event_data == "生成图像":
                if data_set == [[1]] or data_set == []:
                    sg.Popup("请添加数据", keep_on_top=True)
                    continue

                x_coo = 0
                for i in range(len(data_set)):
                    if not data_set[i]:
                        sg.Popup("请添加数据", keep_on_top=True)
                        continue
                    sep_num = 0
                    # update 导入成功 & 自定义完成
                    window_data["导入成功"].update("新文件待导入...", text_color="red")
                    window_data["确认自定义"].update("自定义待确认...", text_color="red")

                    # provide different colors
                    color = change_color(cumulative_color)

                    # set the size & title & position of the graph, give a icon to the figure
                    if values_data["尾列"] != '':
                        plt.figure(("第" + str(first_col+i) + "列"), figsize=(8, 10))
                    else:
                        plt.figure((values_data["名称"]), figsize=(8, 10))
                    # 每个窗口都往右边平移
                    mngr = plt.get_current_fig_manager()  # 获取当前figure manager
                    mngr.window.wm_geometry(f"+{x_coo}+{0}")  # 调整窗口在屏幕上弹出的位置
                    x_coo += 100

                    # first graph
                    subplot_221(data_set[i], values_data, spot_size, color, mean[i])

                    # second graph
                    plt.subplot(2, 2, 2)
                    # set values of x and y
                    # x2 = np.linspace(mean - 3 * standard_deviation, mean + 3 * standard_deviation, 100)
                    # xmin, xmax = plt.xlim()
                    # x2 = np.linspace(xmin, xmax, 100)
                    x2 = np.linspace(mean[i] - 3 * standard_deviation[i], mean[i] + 3 * standard_deviation[i], 100)
                    y2 = stats.norm.pdf(x2, loc=mean[i], scale=standard_deviation[i])
                    # y2 = (1/(standard_deviation * ((2 * 3.141592653)**0.5))) * (2.718281828**(-(((x2 - mean)**2) / (
                    # 2*variance))))

                    x_label = values_data["x轴名称"]
                    plt.xlabel(x_label, size=15)
                    plt.ylabel("正态分布对比区", size=15)

                    # display the value of the mean, variance and standard deviation on graph on appropriate position
                    proper_separation = [mean[i] + standard_deviation[i], (1 / (3 * standard_deviation[i])) / 10]
                    plt.text(proper_separation[0], stats.norm.pdf(mean[i], mean[i], standard_deviation[i]),
                             f"平均值 = {round(mean[i], 4)}", size=12, color=color)
                    plt.text(proper_separation[0],
                             stats.norm.pdf(mean[i], mean[i], standard_deviation[i]) - proper_separation[1],
                             f"平方差 = {round(variance[i], 4)}", size=12, color=color)
                    plt.text(proper_separation[0],
                             stats.norm.pdf(mean[i], mean[i], standard_deviation[i]) - 2 * proper_separation[1],
                             f"标准差 = {round(standard_deviation[i], 4)}", size=12, color=color)

                    # draw the line

                    plt.plot(x2, y2, color)
                    # label the graph
                    plt.xlabel(x_label, size=15)
                    plt.title(values_data["名称"], size=20)

                    plt.grid(True)
                    # set the length of this subplot longer

                    # third graph
                    subplot_223(data_corrected[i], values_data, hist_pre)

                    # fourth graph
                    subplot_224(data_set[i], values_data, hist_pre)

                    cumulative_color += 1

                plt.show()

            if event_data == "概率计算":
                if data_set == [[1]] or data_set == []:
                    sg.Popup("请添加数据", keep_on_top=True)
                    continue
                try:
                    x1 = float(values_data["x1"])
                    x2 = float(values_data["x2"])
                except ValueError:
                    sg.Popup("输入无效", keep_on_top=True)
                    continue

                # give select a value
                select = int(last_col)-first_col

                possibility = abs(
                    (stats.norm.cdf(x2, mean[select], standard_deviation[select]) - stats.norm.cdf(x1, mean[select], standard_deviation[select])) * 100)

                window_data["概率"].update(value=f"{round(possibility, 2)}%")

                possibility_rough = abs(
                    (stats.norm.cdf(x2, mean_rough[select], standard_deviation_rough[select]) - stats.norm.cdf(x1, mean_rough[select],
                                                                                               standard_deviation_rough[select])) * 100)

                data = pd.Series(data_corrected[select])

                # change to the correct subplot
                plt.subplot(2, 2, 3)

                # ax = normal distribution curve
                ax = sns.distplot(data, kde=False,
                                  kde_kws={'linewidth': 2},
                                  label="密度图", fit=stats.norm)

                # get the value of the area under the curve in certain range
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
                             stats.norm.pdf(mean[select], mean[select], standard_deviation[select]) - sep_num * proper_separation[1],
                             f"{values_data['x1']} ~ {values_data['x2']}: {area.round(2)}%", size=12,
                             color=color)


                    plt.subplot(2, 2, 4)
                    plt.text(proper_separation[0],
                             stats.norm.pdf(mean_rough[select], mean_rough[select], standard_deviation_rough[select]) - (sep_num/3) * proper_separation[1],
                             f"{values_data['x1']} ~ {values_data['x2']}: {round(possibility_rough, 2)}%", size=12,
                             color=color)

                    plt.show()
                    color = change_color(cumulative_color)
                    cumulative_color += 1
                    sep_num += 2

            elif event_data == "全屏":
                if count_size == 1:
                    # 使得窗口全屏
                    window_data.Maximize()
                    global_update(window_data, font, font_size)
                    # let the font size be 1/10 of the hight of window


                elif count_size == -1:
                    # 使得窗口恢复
                    window_data.Normal()
                    global_update(window_data, font, font_size)

                count_size *= -1

            elif event_data == "放大":
                window_data.Size = (window_data.Size[0] + 100, window_data.Size[1] + 70)

            elif event_data == "缩小":
                window_data.Size = (window_data.Size[0] - 100, window_data.Size[1] - 70)

            elif event_data == "字体+":
                font_size += 1
                global_update(window_data, font, font_size)
            elif event_data == "字体-":
                font_size -= 1
                global_update(window_data, font, font_size)

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

                if values_data["尾列"] != "":
                    for i in range(len(data_set)):
                        plt.figure(f"第{first_col + i}列")
                        subplot_223(data_corrected[i], values_data, hist_pre)
                        subplot_224(data_set[i], values_data, hist_pre)
                else:
                    subplot_223(data_corrected[i], values_data, hist_pre)
                    subplot_224(data_set[i], values_data, hist_pre)

                window_data["精度"].update(f"直方图精度: {hist_pre:<3}")
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

                for i in range(len(data_set)):
                    subplot_221(data_set[i], values_data, spot_size, color, mean[i])

                window_data["散点"].update(f"散点大小: {spot_size:<3}")
                plt.show()

