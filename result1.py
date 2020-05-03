import requests
from bs4 import BeautifulSoup
import pandas
import os
import re
import xlsxwriter
def results(sem,sec,join):
    global file_name, file_name1, final_result, file_name2
    if len(join) == 10 or len(join) == 12:
        payload = {
                "__EVENTTARGET":"",
                "__EVENTARGUMENT":"",
                "__VIEWSTATE":"/wEPDwULLTE3MTAzMDk3NzUPZBYCAgMPZBYCAgcPDxYCHgRUZXh0ZWRkZKKjA/8YeuWfLRpWAZ2J1Qp0eXCJ",
                "__VIEWSTATEGENERATOR":"65B05190",
                "__EVENTVALIDATION":"/wEWFAKj/sbfBgLnsLO+DQLIk+gdAsmT6B0CypPoHQLLk+gdAsyT6B0CzZPoHQLOk+gdAt+T6B0C0JPoHQLIk6geAsiTpB4CyJOgHgLIk5weAsiTmB4CyJOUHgKL+46CBgKM54rGBgK7q7GGCLOsGLAxgUwycOU5mDizjY4EVXof",
                "cbosem":"1",
                "txtreg":"1210314401",
                "Button1":"Get Result"
                }
        try:
            if len(join) == 10:
                base = join[0:8]
                ran = 68
            else:
                base = join[0:9]
                ran = 160
            try:
                payload['cbosem'] = sem
                result = []
                list = []
                for roll in range(1, ran):
                    try:
                        if ran == 68:
                            if 1 <= roll <= 9:
                                payload['txtreg'] = base + "0" + str(roll)
                            else:
                                payload['txtreg'] = base + str(roll)
                        else:
                            if 1 <= roll <= 9:
                                payload['txtreg'] = base + "00" + str(roll)
                            elif 10 <= roll <= 99:
                                payload['txtreg'] = base + "0" + str(roll)
                            else:
                                payload['txtreg'] = base + str(roll)
                        res = requests.post("https://doeresults.gitam.edu/onlineresults/pages/Newgrdcrdinput1.aspx",
                                            data=payload)
                        soup = BeautifulSoup(res.text, "html.parser")
                        name = soup.find("span", {"id": "lblname"}).text
                        reg = soup.find("span", {"id": "lblregdno"}).text
                        heads = []
                        fails = ["Name", "Roll Number", "Subject1", "Subject2", "Subject3", "Subject4", "Subject5",
                                 "Subject6", "Subject7", "Subject8", "Subject9", "Subject10"]
                        heads.append("Name")
                        heads.append("Roll No")

                        sgpa = soup.find("span", {"id": "lblgpa"}).text
                        cgpa = soup.find("span", {"id": "lblcgpa"}).text
                        temp1 = []
                        if sgpa == "0":
                            temp1.append(name)
                            temp1.append(" ")
                            temp1.append(reg)
                            temp1.append(" ")
                        table = soup.find("table", {"class": "table-responsive"})
                        rows = table.find_all("tr")[1:]
                        temp = []
                        temp.append(name)
                        temp.append(reg)
                        for row in rows:
                            count = 0
                            for i in row.findAll("td"):
                                if count == 3:
                                    temp.append(i.text)
                                    if i.text == "F" or i.text == "Ab":
                                        temp1.append(tx)
                                        temp1.append(" ")
                                elif count == 0:
                                    z = i.text
                                elif count == 1:
                                    tx = i.text
                                    heads.append(i.text + "(" + z + ")")
                                count = count + 1
                        temp.append(sgpa)
                        temp.append(cgpa)
                        result.append(temp)
                        if len(temp1) != 0:
                            list.append(temp1)
                        cx = cx + 1
                    except:
                        pass
                heads.append("SGPA")
                heads.append("CGPA")
                df = pandas.DataFrame(result, index=None)
                df1 = pandas.DataFrame(list, index=None)

                if ran == 68:
                    num = join[5:7]
                else:
                    num = join[2:4]
                n = int(num) + 4
                locname = os.getcwd()
                final_dir = os.path.join(locname, r'Results')
                if not os.path.exists(final_dir):
                    os.makedirs(final_dir)
                pa = str(final_dir) + "\\"
                file_name = "Year(" + num + "-" + str(n) + ")Sec-" + sec + "-Sem-" + sem + "-Results.csv"
                file_name1 = "Year(" + num + "-" + str(n) + ")Sec-" + sec + "-Sem-" + sem + "-Results.xlsx"
                file_name2 = "Failures_Year(" + num + "-" + str(n) + ")Sec-" + sec + "-Sem-" + sem + "-Results.csv"
                df.to_csv(pa + file_name, header=heads, index=False)

                df1.to_csv(pa + file_name2, header=None, index=None)

                final_result = "graph_" + file_name1
                workbook = xlsxwriter.Workbook(pa + final_result)
                worksheet = workbook.add_worksheet()
                df1 = pandas.read_csv(pa + file_name)
                df2 = df1.iloc[:, 2:-2]
                df3 = df2.replace({'O': 10, 'A+': 9, 'A': 8, 'B+': 7, 'B': 6, 'C': 5, 'D': 4, 'F': 0})
                x = df2.columns.values
                q = ['A11', 'J11', 'S11', 'AB11', 'A28', 'J28', 'S28', 'AB28', 'AK28']
                h = ['B1', 'E1', 'H1', 'K1', 'N1', 'Q1', 'T1', 'W1', 'Z1', 'AC1', 'AF1']
                k = ['B2', 'E2', 'H2', 'K2', 'N2', 'Q2', 'T2', 'W2', 'Z2', 'AC2', 'AF2']
                o = ['B', 'E', 'H', 'K', 'N', 'Q', 'T', 'W', 'Z', 'AC', 'AF']
                gt = ['A', 'D', 'G', 'J', 'M', 'P', 'S', 'V', 'Y', 'AB', 'AE']
                u = ['A2', 'D2', 'G2', 'J2', 'M2', 'P2', 'S2', 'V2', 'Y2', 'AB2', 'AE2']
                ux = ['A1', 'D1', 'G1', 'J1', 'M1', 'P1', 'S1', 'V1', 'Y1', 'AB1', 'AE1']
                cv = ["Grade"]
                for t in range(len(x)):
                    df4 = df2.iloc[:, t].value_counts()

                    l = df4.to_dict()
                    y = []
                    for i in l.keys():
                        y.append(i)

                    z = []
                    vb = ["O", "A+", "A", "B+", "B", "C", "D", "F"]
                    vs = []
                    # print(len(vb))
                    for it in vb:
                        #    print(it)
                        if it in l.keys():
                            vs.append(l[it])
                        else:
                            vs.append(0)
                    for i in l.values():
                        z.append(i)
                        # print(vb)
                        # print(vs)
                        # Create a new Chart object.
                    chart = workbook.add_chart({'type': 'column'})
                    # Other chart commands

                    # Writing data to different columns for different charts
                    worksheet.write(ux[t], cv[0])
                    worksheet.write_column(u[t], vb)
                    worksheet.write(h[t], x[t])
                    worksheet.write_column(k[t], vs)

                    # Configure the chart. In simplest case we add one or more data series.
                    # chart.add_series({'values': '=Sheet1!$A$1:$A$5'})
                    chart.add_series({'name': x[t], 'categories': '=Sheet1!$' + gt[t] + '$2:$' + gt[t] + '$9',
                                      'values': '=Sheet1!$' + o[t] + '$2:$' + o[t] + '$9'})

                    worksheet.insert_chart(q[t], chart)

                workbook.close()

                return True
            except:
                return False
        except:
            return False
    else:
        return False