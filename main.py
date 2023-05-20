import xlrd
import os
import numpy as np
import matplotlib.pyplot as plt
import scipy.stats as stats

while True:
    # initialize parameters
    mean = 0
    variance = 0
    standard_deviation = 0
    file_name = ""
    data_set = []

    # clean the screen
    os.system("cls")

    # give command
    command = input("Insert command (any key / exit)：")
    if command == "exit":
        exit()

    else:
        file_name = input("Insert File name：")
        print("File name is：", file_name)

        while True:
            check = os.path.isfile(file_name)
            if check == 0:
                print("File not found")
                file_name = input("Insert File name：")
            else:
                print("File found")
                break
        
        # initialize data set
        data_set = []
        excel = xlrd.open_workbook(file_name)
        sheet = excel.sheet_by_index(0)
        print("Voltages: ", sheet.row_values(5)[1:])

        # accept more data in the future, currently only accept data on a specific row

        # calculate mean and variance
        mean = sum(sheet.row_values(5)[1:]) / len(sheet.row_values(5)[1:])
        variance = sum([(i - mean) ** 2 for i in sheet.row_values(5)[1:]]) / len(sheet.row_values(5)[1:])
        standard_deviation = np.sqrt(variance)

        # print parameters for checking
        print("Mean: ", mean)
        print("Variance: ", variance)
        print("Standard Deviation: ", standard_deviation)

        # draw the possibility curve
        # why red line always very small!
        x = np.linspace(mean-3*standard_deviation,mean+3*standard_deviation,100)
        y = stats.norm.pdf(x, mean, standard_deviation)
        plt.plot(x, y, 'r--')

        # give command for operation
        while True:
            command_in = input("Insert command (graph / cal_p / end)：")
            # print the graph
            if command_in == "graph":
                plt.show()

            elif command_in == "cal_p":
                x1 = input("Insert lower bound：")
                x2 = input("Insert higher bound：")

                # check if x1 and x2 are valid
                while True:
                    try:
                        x1 = float(x1)
                        x2 = float(x2)
                        break
                    except ValueError:
                        print("Invalid input")
                        x1 = input("Insert lower bound：")
                        x2 = input("Insert higher bound：")

                possibility = abs((stats.norm.cdf(x2, mean, standard_deviation) - stats.norm.cdf(x1, mean,
                                                                                             standard_deviation))*100)
                # print the possibility with precision to 4 decimal places
                print("Possibility: ", round(possibility, 2), "%")

            # start next analysis
            elif command_in == "end":
                break

