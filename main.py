#!/usr/bin/env python3

import argparse
import os
import datetime
import gspread
import pandas as pd
from logzero import logger
from moneyforward_driver import MoneyforwardDriver


def update_sheet(sheet: gspread.Worksheet, df_new: pd.DataFrame):
    records = sheet.get_all_records()
    if records:
        df_current = pd.DataFrame(records)
        # df_currentからdf_newにある日付を除外する
        df_current_dropped = df_current[
            ~df_current["日付"].isin(df_new["日付"].unique())
        ]
        # df_currentをdf_newに更新する。indexがなかったら追加する
        df_paste = pd.concat([df_current_dropped, df_new])
    else:
        df_paste = df_new

    sheet.clear()
    sheet.update("A1", [df_paste.columns.values.tolist()] + df_paste.values.tolist())


def update(ss_name: str, cookie_path: str, date_since: datetime.date = None):
    """指定した月から現在までの収入・支出をスプレッドシートに更新する

    Args:
        date (datetime.date, optional): 月。指定しない場合は今月を指定する.
    """
    if (not os.environ["MF_EMAIL"]) or (not os.environ["MF_PASSWORD"]):
        logger.error("MF_EMAIL or MF_PASSWORD is not set.")
        return False

    if date_since is None:
        # 今日から2ヶ月前を指定
        date_since = datetime.date.today() - datetime.timedelta(days=60)

    mf = MoneyforwardDriver(cookie_path)
    login_success = mf.login()

    if not login_success:
        logger.error("Login failed.")
        return

    [income_new, expense_new] = mf.fetch_monthly_income_and_expenses_since(
        date_since.year, date_since.month
    )

    gc = gspread.service_account()
    ss = gc.open(ss_name)
    expense_sheet = ss.worksheet("支出")
    update_sheet(expense_sheet, expense_new)

    income_sheet = ss.worksheet("収入")
    update_sheet(income_sheet, income_new)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()

    # Required positional argument
    parser.add_argument("ss_name", help="Name of the destination spreadsheet")

    # Optional argument flag which defaults to False
    parser.add_argument("-c", "--cookie-path", action="store", dest="cookie_path")

    args = parser.parse_args()

    update(args.ss_name, args.cookie_path, date_since=datetime.date(2023, 9, 1))
