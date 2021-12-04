import os
import pandas as pd
import xlsxwriter
import re
import shutil


def gen_marksheet(p, n):
    roll_num = pd.read_csv('static/files/master_roll.csv')
    resp = pd.read_csv('static/files/responses.csv')
    resp.rename(columns={'Roll Number': 'roll'}, inplace=True)
    res = pd.merge(resp, roll_num, how='outer', on=["roll"])
    s = 1
    for i in range(7, 35):
        res.rename(columns={'Unnamed: ' + str(i): 'q' + str(s)}, inplace=True)
        s += 1
    if os.path.exists('output//marksheet1'):
        shutil.rmtree('output//marksheet1')
    os.makedirs('output//marksheet1')
    for index, row in res.iterrows():
        a = str(row['roll'])
        b = str(row['name'])
        if pd.isnull(row['Score']):
            w_book = xlsxwriter.Workbook('output//marksheet1//' + str(a) + '_ABSENT.xlsx')
            w_sheet = w_book.add_worksheet()
            w_book.close()
        else:
            c = str(row['Score'])
            str1 = '(.*?) / 140'
            match = re.match(str1, c)
            if match:
                marks = int(match.group(1))
            num_right = marks / 5
            postive1 = num_right * p
            right = str(marks / 5)
            left = res.loc[[index]].isna().sum().sum()
            wrong = 28 - num_right - left
            negative = wrong * n
            tt = postive1 + negative
            w_book = xlsxwriter.Workbook('output//marksheet1//' + str(a) + '.xlsx')
            w_sheet = w_book.add_worksheet()
            w_sheet.set_column('A:E', 15)
            w_sheet.write('C5', 'marksheet1')
            w_image = 853  # width_of_image
            h_image = 126  # Hight_of_image
            w_cell = 87  # width_of_cell
            h_cell = 15  # Hight_of_cell
            scale_x = w_cell / w_image
            scale_y = h_cell / h_image
            w_sheet.insert_image('A1', 'IIT_Patna_logo.png', {'x_scale': 0.75, 'y_scale': scale_y * 6 + 0.05})
            cell_1 = w_book.add_format({'bold': True, 'font_size': 14, 'font_name': 'Calibri'})
            cell_2 = w_book.add_format({'bold': True,'font_size': 14, 'font_name': 'Calibri', 'align': 'left'})
            cell_3 = w_book.add_format(
                {'bold': True, 'font_size': 14, 'font_name': 'Calibri', 'align': 'center', 'border': 2})
            cell_4 = w_book.add_format({'border': 2})
            cell_5 = w_book.add_format({'bold': True, 'font_size': 14, 'font_name': 'Calibri', 'align': 'left'})
            cell_6 = w_book.add_format({'font_size': 14, 'font_name': 'Calibri', 'align': 'centre', 'border': 2})
            cell_7 = w_book.add_format(
                {'font_size': 14, 'font_name': 'Calibri', 'align': 'centre', 'font_color': 'green', 'border': 2})
            cell_8 = w_book.add_format(
                {'font_size': 14, 'font_name': 'Calibri', 'align': 'centre', 'font_color': 'red', 'border': 2})
            cell_9 = w_book.add_format(
                {'font_size': 14, 'font_name': 'Calibri', 'align': 'centre', 'font_color': 'blue', 'border': 2})
            cell_10 = w_book.add_format({'font_size': 14, 'font_name': 'Calibri', 'align': 'centre'})
            cell_11 = w_book.add_format({'bold': True, 'font_size': 14, 'font_name': 'Calibri', 'align': 'right'})
            w_sheet.write_string('A6', 'Name:', cell_11)
            w_sheet.write_string('A7', 'Roll number:', cell_11)
            w_sheet.write_string('E6', 'Quiz', cell_2)
            w_sheet.write_string('A10', 'No.', cell_3)
            w_sheet.write_string('A11', 'Marking', cell_3)
            w_sheet.write_string('A12', 'Total', cell_3)
            w_sheet.write_string('D6', 'Exam:', cell_11)
            w_sheet.write_string('B9', 'Right', cell_3)
            w_sheet.write_string('C9', 'Wrong', cell_3)
            w_sheet.write_string('D9', 'Not Attempted', cell_3)
            w_sheet.write_string('E9', 'Max', cell_7)
            w_sheet.write_string('C10', str(wrong), cell_8)
            w_sheet.write_string('C11', str(n), cell_8)
            w_sheet.write_string('A9', '', cell_4)
            w_sheet.write_string('D10', str(left), cell_6)
            w_sheet.write_string('D11', '0', cell_6)
            w_sheet.write_string('E10', '28', cell_6)
            w_sheet.write_string('E11', '', cell_4)
            w_sheet.write_string('D12', '', cell_4)
            w_sheet.write_string('B10', str(right), cell_7)
            w_sheet.write_string('B11', str(p), cell_7)
            w_sheet.write_string('B6', str(b), cell_5)
            w_sheet.write_string('B7', str(a), cell_5)
            w_sheet.write_string('B12', str(postive1), cell_7)
            w_sheet.write_string('C12', str(negative), cell_8)
            w_sheet.write_string('E12', str(tt) + '/' + str(28 * p), cell_9)
            w_sheet.write_string('A15', 'Student Answer', cell_3)

            for i in range(16, 44):
                if pd.isnull(row['q' + str(i - 15)]):
                    w_sheet.write_string('A' + str(i), '')
                else:
                    w_sheet.write_string('A' + str(i), str(row['q' + str(i - 15)]), cell_10)
                w_sheet.write_string('B' + str(i), str(res.loc[0, 'q' + str(i - 15)]), cell_9)

            w_sheet.conditional_format('A16:A44', {'type': 'formula',
                                                   'criteria': '=A16=B16',
                                                   'format': cell_7})
            w_sheet.conditional_format('A16:A44', {'type': 'formula',
                                                   'criteria': '=A16<>B16',
                                                   'format': cell_8})

            w_sheet.write_string('B15', 'Correct Answer', cell_3)

            w_book.close()


gen_marksheet(4, -2)


def con_marks(p, n):
    resp = pd.read_csv('static/files/responses.csv')
    resp.rename(columns={'Roll No.': 'roll'}, inplace=True)

    s = 1
    for i in range(7, 35):
        resp.rename(columns={'Unnamed: ' + str(i): 'q' + str(s)}, inplace=True)
        s += 1
    if os.path.exists('output//marksheet1//concise_marksheet1.csv'):
        os.remove('output//marksheet1//concise_marksheet1.csv')
    gbb = []
    bbabs = []
    score = []
    for index, row in resp.iterrows():

        c = str(row['Score'])
        str1 = '(.*?) / 140'
        match = re.match(str1, c)
        if match:
            marks = int(match.group(1))
        num_right = marks / 5
        postive1 = num_right * p
        right = str(marks / 5)
        left = resp.loc[[index]].isna().sum().sum()
        wrong = 28 - num_right - left
        negative = wrong * n
        tt = postive1 + negative
        total = 28 * p

        gbb.append(str(tt) + ' / ' + str(total))
        score.append(str(postive1) + '/' + str(total))
        bbabs.append('(' + str(num_right) + ' , ' + str(wrong) + ' , ' + str(left) + ')')
    resp.rename(columns={'Score': 'Google_Score'}, inplace=True)

    resp['Google_Score'] = score
    resp['Score_after_negative_marking'] = gbb
    resp['Status_Answer'] = bbabs
    resp.to_csv('output//marksheet1//concise_marksheet1.csv')


con_marks(4, -2)