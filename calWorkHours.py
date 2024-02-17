#cal working hours

import datetime
import pandas as pd

def calWorkHours():
    while True:
        workday = input("Enter work day (YYYY-MM-DD): ")
        workday = datetime.datetime.strptime(workday, "%Y-%m-%d")
        print("Work day: ", workday.strftime("%Y-%m-%d"))

        start = input("Enter start time (HH:MM): ")
        end = input("Enter end time (HH:MM): ")
        start = datetime.datetime.strptime(start, "%H:%M")
        end = datetime.datetime.strptime(end, "%H:%M")
        diff = end - start
        rest_time_no_sal = input("Enter rest time without salary in minutes:")
        rest_time_no_sal = datetime.timedelta(minutes=int(rest_time_no_sal))
        rest_time_no_sal = rest_time_no_sal.seconds/3600
        print("Rest time without salary: ", rest_time_no_sal, "hours")
        workHours = diff.seconds/3600 - rest_time_no_sal
        if workHours > 8:
            overtimeHours = workHours - 8
        else:
            overtimeHours = 0
        regularHours = workHours - overtimeHours
        day_salary = round(regularHours * 22.92 + overtimeHours * (22.92 * 1.5),2)
        print("Total working hours: ", workHours, "hours")
        print("Including regular hours :", regularHours, "hours and overtime: ",overtimeHours, "hours")
        print("Day salary: $", day_salary)


        df = pd.DataFrame({
            "Date": [workday.strftime("%Y-%m-%d")],
            "Start": [start.strftime("%H:%M")],
            "End": [end.strftime("%H:%M")],
            "Rest Time without Salary": [rest_time_no_sal],
            "Total Work Hours": [workHours],
            "Regular Hours": [regularHours],
            "Overtime Hours": [overtimeHours],
            "Day Salary": [day_salary]
        })
        # check if the file exists
        try:
            writer = "workHours.xlsx"
            with pd.ExcelWriter(writer, engine="openpyxl", mode="a", if_sheet_exists='overlay') as writer:
                df.to_excel(writer, sheet_name="Sheet1", index=False, startrow=writer.sheets['Sheet1'].max_row, header=False)
            print("Data has been added to the file.")
        except FileNotFoundError:
            print("Excel file not found. Creating a new file...")
            writer = "workHours.xlsx"
            df.to_excel(writer, sheet_name="Sheet1", index=False, header=True)

        # ask the user whether to continue
        cont = input("Do you want to continue? (Y/N): ")
        if cont == "N":
            break
        else:
            continue

calWorkHours()