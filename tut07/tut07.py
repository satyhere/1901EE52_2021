from pandas import read_csv, DataFrame
from re import match


def is_rollno(rollno):
    return match(r"\d\d\d\d\w\w\d\d", rollno)


def get_ltp(string_ltp):
    ltp = string_ltp.split('-')
    return ltp


def get_full_list(course_taken):
    res_list = set()
    for index, row in course_taken.iterrows():
        course = row["subno"]
        if is_rollno(row["rollno"]):
            if course not in course_ltp:
                print("Course is not in LTP")
            if course_ltp[course][0] != "0":
                res_list.add((row["rollno"], row["subno"], "1"))
            if course_ltp[course][1] != "0":
                res_list.add((row["rollno"], row["subno"], "2"))
            if course_ltp[course][2] != "0":
                res_list.add((row["rollno"], row["subno"], "3"))

    return res_list


def get_done_list(fb_info):
    res_list = set()
    for index, row in fb_info.iterrows():
        if is_rollno(row["stud_roll"]):
            res_list.add((row["stud_roll"], row["course_code"], str(row["feedback_type"])))
    return res_list


fb_info = read_csv("course_feedback_submitted_by_students.csv")
course_master = read_csv("course_master_dont_open_in_excel.csv")
student_info = read_csv("studentinfo.csv")
course_taken = read_csv("course_registered_by_all_students.csv")
output_file = DataFrame(
    columns=["Roll Number", "Registered Semester", "Scheduled Semester", "Course Code", "Name", "Email", "AEmail", "Contact"])

sch_sem = {}

for index, row in course_taken.iterrows():
    if is_rollno(row["rollno"]):
        # sch_sem[row["rollno"]] = {}
        # print(row)
        if row["rollno"] not in sch_sem:
            sch_sem[row["rollno"]] = {}

        sch_sem[row["rollno"]][row["subno"]] = (row["register_sem"], row["schedule_sem"])
# sch_sem[row["rollno"][row["subno"]]] = 5

course_ltp = {}

for index, row in course_master.iterrows():
    course_ltp[row["subno"]] = get_ltp(row["ltp"])

full_list = get_full_list(course_taken)
# full_list.sort()

stud_info = {}

for index, row in student_info.iterrows():
    stud_info[row["Roll No"]] = {}
    stud_info[row["Roll No"]] = row

done_list = get_done_list(fb_info)
# done_list.sort()

full_list = full_list | done_list
rem_list = done_list ^ full_list

# print(len(rem_list))
# print(len(full_list))
# print(len(done_list))

for entry in rem_list:
    # print(entry)
    rollno = entry[0]
    course = entry[1]
    feed_type = entry[2]
    regis_sem = sch_sem[rollno][course][0]
    sched_sem = sch_sem[rollno][course][1]
    if rollno in stud_info:
        name = stud_info[rollno]["Name"]
        mail = stud_info[rollno]["email"]
        amail = stud_info[rollno]["aemail"]
        contact = stud_info[rollno]["contact"]
        new_row = {"Roll Number": rollno, "Registered Semester": regis_sem, "Scheduled Semester": sched_sem, "Course Code": course,
                   "Email": mail, "AEmail": amail, "Contact": contact, "Name": name}
        # print(new_row)
        output_file = output_file.append(new_row, ignore_index=True)

# print(output_file)
# output_file.reset_index()
output_file.to_excel("course_feedback_remaining.xlsx", index=False)
