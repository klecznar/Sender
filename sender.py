import sys
import time

from functions import send_mail, refresh_tracker, analyze_rows
import openpyxl
import pandas as pd


# initialize tasks

def scheduler():

    try:
        refresh_tracker()
        analyze_rows()
        send_mail()
    except Exception as e:
        print("ERROR: " + e)

    # Every day at 08:00 scheduler() is called
    schedule.every().day.at("08:00").do(scheduler)

    # Loop so that the scheduling task
    # keeps on running all time.

    while True:
    #
    #     # Checks whether a scheduled task
    #     # is pending to run or not
        schedule.run_pending()
        time.sleep(1)