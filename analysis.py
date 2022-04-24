from datetime import date
import pandas as pd
import logging
import time
import os

# Was written by Ellie, Chodjayev - for problems please contact.


# Create logger
logging.basicConfig(filename="problems.log",
                    level=logging.DEBUG)
logger = logging.getLogger()


def read_file():
    """Filter Test-Population sheet through Pending/Priority for easier workflow"""
    start_time = time.time()
    today = date.today()
    reversed_date = today.strftime("%m-%d-%Y")
    project = ['ADL-S-BGA', 'ADL-S', 'ADL-P']

    # Pick a project
    counter = 0
    try:
        for list in project:
            if list in project[counter]:
                counter += 1
                print(f"{list} - {counter}")
            else:
                pass
    finally:
        x = int(input("Select Project : "))
        print(f"{project[x - 1]} Selected\n")  # List prints at 0
        redirect = project[x - 1]
        assert x in range(1, 4), "Cannot be outside range"

    # Redirect to project folder
    project_path = "S:/Groups/FV/Memory/everything_else/MMEX/"
    if redirect == "ADL-S-BGA":
        new_redirect = "OutputFiles-ADL-S-BGA"
    else:
        if redirect == "ADL-S":
            new_redirect = "OutputFiles-ADL-S-PRQ"
        else:
            if redirect == "ADL-P":
                new_redirect = "OutputFiles-ADL-P-PRQ"

    # Enter full project path 
    original_dir = os.getcwd()
    os.chdir(f"{project_path}/{new_redirect}")
    latest_folder = os.listdir()
    os.chdir(latest_folder[-1])
    files = os.listdir()

    # Look for Excel
    try:
        for list in files:
            if list.endswith('xlsx'):
                file_name = list
    except UnboundLocalError as err:
        logger.error(err)
    finally:
        pass

    # Verify if file downloaded is from today (date included in file name)
    if reversed_date in file_name:
        print(f"File Date Matches - {reversed_date}")
    else:
        print(f"File is not from today \n{file_name}\n{reversed_date}\n")

    # Look for sheet names
    file = file_name
    df = pd.ExcelFile(file).sheet_names

    # Filter sheets
    counter = 0
    sheets = [""]
    for sheet in df:
        if sheet[0] == "Δ" or sheet == "Log Data":
            pass
        else:
            counter += 1
            sheets.append(sheet)
            print(f"{sheets[counter]} - {counter}")

    # Sheet selection 
    try:
        x = int(input("Select Sheet Number: "))
        assert x in range(1, counter + 1), "Select a value from list"
    except ValueError as err:
        logger.error(err)
        raise
    else:
        df = pd.read_excel(f"{file}", f"{sheets[x]}")
    finally:
        print(f"{sheets[x]} Selected")

    # Compare and keep matching columns
    a = df.columns
    b = ['ch0-idc s/n', 'ch1-idc s/n', 'ch0s0-idc s/n', 'ch0s1-idc s/n',
         'ch1s0-idc s/n', 'ch1s1-idc s/n', 'Tested',
         'mem_info_part_number=[Unique Key]',
         'Step', 'Priority Score', 'Unique Key']
    keep_columns = [x for x in a if x in b]

    # Filters by Pending and by Priority - will save output in a file
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)
    pd.set_option('display.max_colwidth', None)
    Pending = df[df['Tested'] == 'Pending'].sort_values(['Priority Score'], ascending=False)[keep_columns]
    Fail = df[df['Tested'] == 'Fail'].sort_values(['Priority Score'], ascending=False)[keep_columns]
    print(f"\n{Pending.head(10)}")
    os.chdir(f"{original_dir}/Saved")
    with open('Pending.txt', "w+") as f:
        print(Pending.head(10), file=f)

    with open('Fail.txt', "w+") as f:
        print(Fail.head(5), file=f)

    # Save logged info and time it took for this script to run
    stop_time = time.time()
    dt = stop_time - start_time
    logger.info("Time required for {file} = {time}".format(file=file,
                                                           time=dt))


if __name__ == "__main__":
    read_file()
