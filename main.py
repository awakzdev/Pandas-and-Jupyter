import concurrent.futures
from datetime import date
from pathlib import Path
import pandas as pd
import logging
import shutil
import time
import sys
import os

# Was written by Ellie, Chodjayev - for problems please contact.

# Create logger
logging.basicConfig(level=logging.DEBUG,
                    format="%(asctime)s [%(levelname)s] %(message)s",
                    handlers=[
                        logging.FileHandler(filename="debug.log"),
                        logging.StreamHandler(sys.stdout)
                    ])
logger = logging.getLogger()


def project(system):
    """Filter Test-Population sheet through Pending/Priority for easier workflow"""
    start_time = time.time()
    today = date.today()
    reversed_date = today.strftime("%m-%d-%Y")
    project_path = "S:/Groups/FV/Memory/everything_else/MMEX"

    # Enter full project path
    original_dir = "C:/Users/ehodjayx/PycharmProjects/pythonProject/Pandas-and-Jupyter"
    os.chdir(f"{project_path}/{system}")
    latest_folder = os.listdir()
    os.chdir(latest_folder[-1])
    logger.info(f'current dir {os.getcwd()}')
    files = os.listdir()
    logger.info(f'current files in folder are : {files}')

    # Create a folder for later use
    new_folder = Path(f"{original_dir}/Saved/{today}/")
    if new_folder.is_dir():
        pass
    else:
        os.mkdir(f"{original_dir}/Saved/{today}/")

    # Look for Excel
    flag = False
    for list in files:
        if '~$' in list :
            logger.info(f'detected temp excel copy : {list} , continue')
            continue
        if list.endswith('xlsx'):
            file_name = list
            flag = True
    if not flag:
        logger.info(f"File was not found - {reversed_date}")
        quit()

    # Verify if file downloaded is from today (date included in file name)
    if reversed_date in file_name:
        logger.info(f"File Date Matches - {reversed_date}")
        date_match = True
    else:
        logger.info(f"File is not from today \n{file_name}\n{reversed_date}\n")
        date_match = False

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

    # Loop over sheets
    directory = os.getcwd()
    counter = 0
    try:
        for _ in sheets:
            counter += 1
            if counter == len(sheets):
                break
            else:
                # Compare and keep matching columns
                df = pd.read_excel(f"{file}", f"{sheets[counter]}")
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
                pending = df[df['Tested'] == 'Pending'].sort_values(['Priority Score'], ascending=False)[keep_columns]
                fail = df[df['Tested'] == 'Fail'].sort_values(['Priority Score'], ascending=False)[keep_columns]
                print(f"\n{pending.head(10)}")
                my_file = Path(f"{original_dir}/Saved/{today}/{system}")
                try:
                    if my_file.is_dir():
                        pass
                    else:
                        os.mkdir(f"{original_dir}/Saved/{today}/{system}")
                finally:
                    os.chdir(f"{original_dir}/Saved/{today}/{system}")
                with open(f"{date_match}", "a+") as f:
                    f.write("Will check if date and file-date match")
                with open('Pending.txt', "a+") as f:
                    f.write(f"\n{sheets[counter]}\n")
                    print(pending.head(10), file=f)
                with open('Fail.txt', "a+") as f:
                    f.write(f"\n{sheets[counter]}\n")
                    print(fail.head(5), file=f)
    finally:
        os.chdir(project_path)

    # Save logged info and time it took to run
    stop_time = time.time()
    dt = stop_time - start_time
    logger.info("Time required for {file} = {time}".format(file=file,
                                                           time=dt))


if __name__ == "__main__":
    with concurrent.futures.ProcessPoolExecutor() as executor:
        p1 = executor.submit(project, "OutputFiles-ADL-S-PRQ")
        p2 = executor.submit(project, "OutputFiles-ADL-P-PRQ")
        p3 = executor.submit(project, "OutputFiles-ADL-S-BGA")
        p4 = executor.submit(project, "OutputFiles-RPL-P-ES2")
