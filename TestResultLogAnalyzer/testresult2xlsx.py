# coding=gbk
import xlsxwriter
import time
import os

date = time.strftime('%Y-%m-%d_%H-%M', time.localtime(time.time()))
xlsx_name = "TestResults_%s.xlsx" % date
workbook = xlsxwriter.Workbook(xlsx_name)

head_format = workbook.add_format({'bold': True, 'align': 'left', 'valign': 'vcenter'})
index_format = workbook.add_format({'align': "center"})
ifs_format = workbook.add_format({'valign': "vcenter"})
pass_format = workbook.add_format({'align': 'center'})
fail_format = workbook.add_format({'align': 'center', 'color': '#FFFFFF', 'bg_color': '#FF0000', 'bold': True})
total_format = workbook.add_format({'align': 'center'})
fails_format = workbook.add_format({'align': 'center', 'color': '#FF0000'})
border_format = workbook.add_format({'border': 1, 'align': 'left', 'font_size': 10})

# key words to match
fail_of_mhs = "测试方法出错"
pass_of_mhs = "测试方法成功"
start_of_ifs = "开始测试接口"
start_of_mhs = "开始测试方法"
end_of_ifs = "测试接口出错"


def time_diff(time_str_start, time_str_end):
    tm_sec1 = date_str2secs(time_str_start.split(".")[0])
    tm_sec2 = date_str2secs(time_str_end.split(".")[0])
    tm_ms1 = time_str_start.split(".")[1]
    if len(tm_ms1) == 2:
        tm_ms1 = int("%s0" % tm_ms1)
    else:
        tm_ms1 = int(tm_ms1)
    tm_ms2 = time_str_end.split(".")[1]
    if len(tm_ms2) == 2:
        tm_ms2 = int("%s0" % tm_ms2)
    else:
        tm_ms2 = int(tm_ms2)
    dur_ms = tm_sec2*1000 + tm_ms2 - tm_sec1*1000 - tm_ms1
    return "%s.%03d" % (dur_ms / 1000, dur_ms % 1000)


def date_str2secs(date_str):
    tm_list = []
    array = date_str.split(' ')
    array1 = array[0].split('-')
    array2 = array[1].split(':')
    for v in array1:
        tm_list.append(int(v))
    for v in array2:
        tm_list.append(int(v))

    tm_list.append(0)
    tm_list.append(0)
    tm_list.append(0)
    if len(tm_list) != 9:
        return 0
    return int(time.mktime(tm_list))

ifs_names = []
case_cnts = []
fail_cnts = []
sum_sheet = workbook.add_worksheet("Summary")

dst_path = '.'
ext_name = 'log'
for log_file in os.listdir(dst_path):
    if log_file.endswith(ext_name):
        if log_file.startswith("trace_main"):
            log_file_p = open(log_file, "r")

            ifs_name = ""
            row_start = 0
            pass_cnt = 0
            fail_cnt = 0
            row = 0
            i = 0

            lines = log_file_p.readlines()
            for line in lines:
                desc = line.split(" : ")[2]

                # start of interface
                if desc.find(start_of_ifs) != -1:
                    ifs_name = desc.split(":")[1].split(".")[-1]
                    row = 0
                    worksheet = workbook.add_worksheet(ifs_name)
                    headings = ["I/Fs", "idx", "Methods", "Results", "Time_start", "Time_end", "Duration(s)"]
                    heading_widths = [30, 3, 30, 7, 22, 22, 10]
                    worksheet.write_row(row, 0, headings, head_format)
                    i = 0
                    for width in heading_widths:
                        worksheet.set_column(i, i, width)
                        i += 1

                    worksheet.freeze_panes(1, 0)
                    row += 1

                    worksheet.write(row, 0, ifs_name)
                    i = 1
                    row_start = row
                    pass_cnt = 0
                    fail_cnt = 0

                if desc.find(start_of_mhs) != -1:
                    worksheet.write(row, 1, i, index_format)
                    i += 1
                    test_mth_name = desc.split(":")[1]
                    worksheet.write(row, 2, test_mth_name.split("_")[1])
                    time_start = line.split(" : ")[1]
                    worksheet.write(row, 4, time_start)
                    row += 1

                if desc.find(pass_of_mhs) != -1:
                    worksheet.write(row-1, 3, "PASS", pass_format)
                    time_end = line.split(" : ")[1]
                    worksheet.write(row-1, 5, time_end)
                    worksheet.write(row - 1, 6, time_diff(time_start, time_end), index_format)
                    pass_cnt += 1

                if desc.find(fail_of_mhs) != -1:
                    worksheet.write(row-1, 3, "FAIL", fail_format)
                    time_end = line.split(" : ")[1]
                    worksheet.write(row-1, 5, time_end)
                    worksheet.write(row-1, 6, time_diff(time_start, time_end), index_format)
                    fail_cnt += 1

                if desc.find(end_of_ifs) != -1:
                    ifs_sum = "%s (%s/%s)" % (ifs_name, fail_cnt, i-1)
                    worksheet.merge_range(row_start, 0, row-1, 0, ifs_sum, ifs_format)
                    ifs_names.append(ifs_name)
                    case_cnts.append(i-1)
                    fail_cnts.append(fail_cnt)

            # end of for line in lines

            # headings = ["I/Fs", "Total", "Fail"]
            # sum_sheet.write_row("A1", headings, head_format)
            sum_sheet.merge_range(0, 0, 1, 0, "I/Fs", head_format)
            sum_sheet.merge_range(0, 1, 0, 2, "Methods", head_format)
            sum_sheet.write("B2", "Total", head_format)
            sum_sheet.write("C2", "Fail", head_format)
            sum_sheet.write_column("A3", ifs_names)
            sum_sheet.write_column("B3", case_cnts, total_format)
            sum_sheet.write_column("C3", fail_cnts, fails_format)
            sum_sheet.set_column(0, 0, 30)
            sum_sheet.set_column(1, 1, 5)
            sum_sheet.set_column(2, 2, 5)
            rows = len(ifs_names) + 1
            sum_sheet.conditional_format(0, 0, row, 3, {'type': 'no_blanks', 'format': border_format})

            chart1 = workbook.add_chart({'type': 'column'})

            # Configure a second series. Note use of alternative syntax to define ranges.
            chart1.add_series({
                'name':       ['Summary', 1, 1],
                'categories': ['Summary', 2, 0, 10, 0],
                'values':     ['Summary', 2, 1, 10, 1],
            })
            chart1.add_series({
                'name':       ['Summary', 1, 2],
                'categories': ['Summary', 2, 0, 10, 0],
                'values':     ['Summary', 2, 2, 10, 2],
            })

            chart1.set_title({'name': 'Test Result Analytics'})
            chart1.set_x_axis({'num_font': {'rotation': -45}})
            chart1.set_y_axis({'name': 'Case Number'})
            chart1.set_style(10)
            sum_sheet.insert_chart("D1", chart1,  {'x_offset': 5, 'y_offset': 5})

workbook.close()
