#Sorter and Analysis for Impedance / Cryo / CTC
#Designed by Jordan Lee

import numpy as np
import xlsxwriter
import math
import code
import os
import os.path
from shutil import copyfile, copy2
from pathlib import Path

#alpha_bet() function helps convert column numerical number to excel's column letter format for use in averaging of multiple tests later in the script
def alpha_bet(N):

    alpha = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

    phrase = ''

    while N >= 0:
        phrase = alpha[math.floor(N % 26)] + phrase
        N = (math.floor(N) / 26) - 1

    return('$' + phrase)

#m_path needs to be the path to the directory just above the files, files should be in the /HOT directory
m_path = str(os.getcwd()).replace('\\', '/') + '/'

path = m_path + 'HOT/'                                                                                                     #HOT
u_path = m_path + 'UNSORTED/'                                                                                                 #UNSORTED
s_path = m_path + 'SORTED/'                                                                                                 #SORTED

#Creation of two directories for data aggregation
#SORTED directory holds all data files that meet the criteria to be aggregated into a specific temperature grouping
#UNSORTED directory holds all data files that did not meet criteria for aggregation
if os.path.isdir(s_path):
    pass
else:
    os.system('"mkdir SORTED"')

if os.path.isdir(u_path):
    pass
else:
    os.system('"mkdir UNSORTED"')

#Creation of temperature-descripted directories to store sorted raw data and an excel file for data aggregation and averaging
for file in os.listdir(path):
    fullname = path + "/" + file
    with open(fullname) as fp:
    	small_array = np.genfromtxt(fullname, dtype=str, delimiter="\t", max_rows=2)
    	if round(float(small_array[0][1])) == round(float(small_array[1][1])):
    		if os.path.isdir(s_path + str(round(float(small_array[0][1]))) + 'K'):
    			copy2(fullname, s_path + str(round(float(small_array[0][1]))) + 'K')
    		else:
    			os.system('"mkdir SORTED\\' + str(round(float(small_array[0][1]))) + 'K"')
    			copy2(fullname, s_path + str(round(float(small_array[0][1]))) + 'K')
    		print('Sorted ' + str(file))
    	else:
    		copy2(fullname, u_path)
    		print('Sorted ' + str(file))

#User input for sample thickness and electrode thickness, electrode thickness is fixed to two options - experimentally
path = s_path
sample = str(input("What is the sample?: "))
sample_thick = int(input("What is the sample thickness? (in microns): "))
electrode_thick = float(input("What is the working diameter? (Cryo = 15.6, Ambient = 5, in mm): "))

subs = os.listdir(path)
for folders in subs:
    fname = path + "/" + folders
    print("Looking through " + path + "/" + folders)

    decider = 0

    dir_list = []
    for file in os.listdir(fname):
        dir_list.append(fname + "/" + file)

    for file in dir_list:
        if file.endswith(".xlsx"):
            print(fname + "/" + file + " exists!")
            decider = 1
        else:
            pass

    if decider == 1:
        pass
    else:
        print("Analyzing " + fname)
        temperature_average = []
        for file in os.listdir(fname):
            fullname = fname + "/" + file
            with open(fullname) as fp:
                temperature_average.append(
                    round(
                        float(
                            np.genfromtxt(
                                fullname, dtype=str, delimiter="\t", max_rows=3
                            )[0, 1]
                        )
                    )
                )

        np.array(temperature_average)
        temperature = np.average(temperature_average)
        temperature_string = sample + " at " + str(int(temperature)) + "K"

        workbook = xlsxwriter.Workbook(fname + "/" + temperature_string + ".xlsx")
        worksheet1 = workbook.add_worksheet(
            sample + " at " + str(int(temperature)) + "K"
        )
        worksheet2 = workbook.add_worksheet("RAW")
        worksheet3 = workbook.add_worksheet("Holder")

        bold = workbook.add_format({"bold": True})

        chart = workbook.add_chart({"type": "scatter", "subtype": "straight"})

        chart.set_title({"name": temperature_string})
        chart.set_x_axis({"name": "frequency", "log_base": 10, "min": 1000, "max": 5e6})
        chart.set_y_axis({"name": "permittivity", "max": 3.3, "min": 2.5})

        worksheet1.insert_chart("K2", chart)

        chart.set_size({"width": 960, "height": 520})
        chart.set_style(5)

        row = 0
        col = 0
        count = 1

        #This section handles calculation of permittivity from capacitance data along the frequency sweep.
        #Permittivity data, along with frequency and dissipation factor data pulled from raw files, is added to a temperature-descripted excel file.
        #Values for every necessary raw data file is aggregated with respect to an integer temperature value and included into the second sheet of the excel file.
        #The main sheet of the excel file handles averaging of values such as permittivity and dissipation factor for data presentation.
        for file in os.listdir(fname):
            fullname = fname + "/" + file
            with open(fullname) as fp:
                headers = np.genfromtxt(fullname, dtype=str, delimiter="\t", max_rows=3)
                raw_data = np.loadtxt(fullname, skiprows=4)

                #Permittivity calculation derived from user input is handled here, as well as error correction.
                real_p_top = sample_thick * 1e-6
                real_p_bot = math.pi * (((electrode_thick*10**-3) / 2) ** 2) * 8.854e-12
                realp = (real_p_top * raw_data[:, 0]) / real_p_bot

                ccorr = raw_data[:, 0] * ((50e-12 / (50e-12 - raw_data[:, 0])))
                realpcorr = (real_p_top * ccorr) / real_p_bot

                calculated_data = np.transpose(np.vstack((realp, ccorr, realpcorr)))
                big_array = np.concatenate((raw_data, calculated_data), axis=1)

                full_data_set = data_holder(
                    sample,
                    round(float(headers[0, 1])),
                    float(headers[2, 1]),
                    2,
                    raw_data[:, 2],
                    raw_data[:, 1],
                    raw_data[:, 0],
                    realp,
                    ccorr,
                    realpcorr,
                )

                worksheet2.write(row, col, sample)
                worksheet2.write(
                    row + 1,
                    col,
                    "Temperature: "
                    + str(headers[0, 1])
                    + "K to "
                    + str(headers[1, 1])
                    + "K",
                )
                worksheet2.write(
                    row + 2,
                    col,
                    "Temperature Change: "
                    + str(round(float(headers[1, 1]) - float(headers[0, 1]), 3))
                    + "K",
                )

                worksheet2.write(row + 4, col, "TEST " + str(count), bold)
                worksheet2.write(row + 5, col, "Cp-Data")
                worksheet2.write(row + 5, col + 1, "D-Data")
                worksheet2.write(row + 5, col + 2, "Frequency")
                worksheet2.write(row + 5, col + 3, "Real P")
                worksheet2.write(row + 5, col + 4, "C Corr")
                worksheet2.write(row + 5, col + 5, "Real P Corr")

                for i in range(6):
                    row = 6
                    for num in big_array[:, i]:
                        worksheet2.write(row, col + i, num)
                        row += 1

            count += 1

            row = 0
            col += 7

        function_count = 0

        if count > 200:
            function_count = 200
        else:
            function_count = count

        frequencies_start = [
            "$C$7",
            "$J$7",
            "$Q$7",
            "$X$7",
            "$AE$7",
            "$AL$7",
            "$AS$7",
            "$AZ$7",
            "$BG$7",
            "$BN$7",
        ]
        frequencies_end = [
            "$C$406",
            "$J$406",
            "$Q$406",
            "$X$406",
            "$AE$406",
            "$AL$406",
            "$AS$406",
            "$AZ$406",
            "$BG$406",
            "$BN$406",
        ]
        permittivities_start = [
            "$D$7",
            "$K$7",
            "$R$7",
            "$Y$7",
            "$AF$7",
            "$AM$7",
            "$AT$7",
            "$BA$7",
            "$BH$7",
            "$BO$7",
        ]
        permittivities_end = [
            "$D$406",
            "$K$406",
            "$R$406",
            "$Y$406",
            "$AF$406",
            "$AM$406",
            "$AT$406",
            "$BA$406",
            "$BH$406",
            "$BO$406",
        ]

        setters = 0

        if count > 10:
            count = 11
        else:
            count = count

        for num in range(count - 1):
            chart.add_series(
                {
                    "name": "Test " + str(num + 1),
                    "categories": "=RAW!"
                    + frequencies_start[num]
                    + ":"
                    + frequencies_end[num],
                    "values": "=RAW!"
                    + permittivities_start[num]
                    + ":"
                    + permittivities_end[num],
                }
            )

        #Quick view of all values of relevance for permittivity experimentation, either for data presentation or debugging.
        worksheet1.write("$A$6", "Cp-Data")
        worksheet1.write("$B$6", "D-Data")
        worksheet1.write("$C$6", "Frequency")
        worksheet1.write("$D$6", "Real P")
        worksheet1.write("$E$6", "C Corr")
        worksheet1.write("$F$6", "Real P Corr")
        worksheet1.write("$G$6", "Real P Dev")
        worksheet1.write("$H$6", "Real P Min")
        worksheet1.write("$I$6", "Real P Max")

        worksheet1.write("$A$5", "AVERAGE FOR " + str(function_count) + " TESTS", bold)

        #Below account for the averaging of all the tests (if tests <= 200) within a sorted temperature directory
        for num in range(400):

            #Cp-Data Averaging
            old_runner = '=AVERAGE('
            ray = 0
            crem = 0
            while ray <= function_count:
                old_runner = old_runner + 'RAW!' + alpha_bet(crem) + '$' + str(num + 7) +', '
                crem += 7
                ray += 1
            old_runner = old_runner[:len(old_runner) - 2] + ')'

            worksheet1.write(
                "$A" + "$" + str(num + 7),
                old_runner)

            #D-Data Averaging
            old_runner = '=AVERAGE('
            ray = 0
            crem = 0
            while ray <= function_count:
                old_runner = old_runner + 'RAW!' + alpha_bet(crem + 1) + '$' + str(num + 7) +', '
                crem += 7
                ray += 1
            old_runner = old_runner[:len(old_runner) - 2] + ')'

            worksheet1.write(
                "$B" + "$" + str(num + 7),
                old_runner)

            #Frequency Averaging
            old_runner = '=AVERAGE('
            ray = 0
            crem = 0
            while ray <= function_count:
                old_runner = old_runner + 'RAW!' + alpha_bet(crem + 2) + '$' + str(num + 7) +', '
                crem += 7
                ray += 1
            old_runner = old_runner[:len(old_runner) - 2] + ')'

            worksheet1.write(
                "$C" + "$" + str(num + 7),
                old_runner)

            #Real P Averaging
            old_runner = '=AVERAGE('
            ray = 0
            crem = 0
            while ray <= function_count:
                old_runner = old_runner + 'RAW!' + alpha_bet(crem + 3) + '$' + str(num + 7) +', '
                crem += 7
                ray += 1
            old_runner = old_runner[:len(old_runner) - 2] + ')'

            worksheet1.write(
                "$D" + "$" + str(num + 7),
                old_runner)

            #C Corr Averaging
            old_runner = '=AVERAGE('
            ray = 0
            crem = 0
            while ray <= function_count:
                old_runner = old_runner + 'RAW!' + alpha_bet(crem + 4) + '$' + str(num + 7) +', '
                crem += 7
                ray += 1
            old_runner = old_runner[:len(old_runner) - 2] + ')'

            worksheet1.write(
                "$E" + "$" + str(num + 7),
                old_runner)

            #Real P Corr
            old_runner = '=AVERAGE('
            ray = 0
            crem = 0
            while ray <= function_count:
                old_runner = old_runner + 'RAW!' + alpha_bet(crem + 5) + '$' + str(num + 7) +', '
                crem += 7
                ray += 1
            old_runner = old_runner[:len(old_runner) - 2] + ')'

            worksheet1.write(
                "$F" + "$" + str(num + 7),
                old_runner)

            #Real P Dev
            old_runner = '=STDEV('
            ray = 0
            crem = 0
            while ray <= function_count:
                old_runner = old_runner + 'RAW!' + alpha_bet(crem + 3) + '$' + str(num + 7) +', '
                crem += 7
                ray += 1
            old_runner = old_runner[:len(old_runner) - 2] + ')'

            worksheet1.write(
                "$G" + "$" + str(num + 7),
                old_runner)

            #Real P Min
            old_runner = '=MIN('
            ray = 0
            crem = 0
            while ray <= function_count:
                old_runner = old_runner + 'RAW!' + alpha_bet(crem + 3) + '$' + str(num + 7) +', '
                crem += 7
                ray += 1
            old_runner = old_runner[:len(old_runner) - 2] + ')'

            worksheet1.write(
                "$H" + "$" + str(num + 7),
                old_runner)

            #Real P Max
            old_runner = '=MAX('
            ray = 0
            crem = 0
            while ray <= function_count:
                old_runner = old_runner + 'RAW!' + alpha_bet(crem + 3) + '$' + str(num + 7) +', '
                crem += 7
                ray += 1
            old_runner = old_runner[:len(old_runner) - 2] + ')'

            worksheet1.write(
                "$I" + "$" + str(num + 7),
                old_runner)

        chart.add_series(
            {"name": "AVERAGE", "categories": "$C$7:$C$406", "values": "$D$7:$C$406"}
        )

        # for p, c, pc in full_data_set.realp, full_data_set.ccorr, full_data_set.realpcorr:
        # 	worksheet.write(row, col + 3, p)
        # 	worksheet.write(row, col + 4, c)
        # 	worksheet.write(row, col + 5, pc)
        # 	row += 1

        workbook.close()

#Pops the user into python's executable shell for fine tuning of data.
code.interact(local=locals())
