import logging
import os
import shutil
import pandas as pd
import numpy as np


def main():

    logging.basicConfig(
        format="%(asctime)s %(message)s",
        datefmt="%m/%d/%Y %I:%M:%S %p",
        level=logging.INFO,
    )
    report_folder = "Reports"
    if os.path.exists(report_folder):
        shutil.rmtree(report_folder, ignore_errors=True)
        logging.info("Report folder is removed.")
    os.mkdir(report_folder)
    logging.info("Report folder is created.")

    my_list = range(0, 100_000, 11)
    my_lists = np.array_split(my_list, 1_000)
    excel_file_name = f"{report_folder}\My_Excel_Report.xlsx"

    n = 0
    with pd.ExcelWriter(excel_file_name) as writer:
        for small_list in my_lists:
            n = n + 1
            wks_name = f"Tab_{n}"
            pd.DataFrame(small_list).to_excel(
                writer, sheet_name=wks_name, header=False, index=False
            )
            logging.info(f"{n}/{len(my_lists)}")
    logging.info(f"File {excel_file_name} is created.")


if __name__ == "__main__":
    main()
