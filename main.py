#!/usr/bin/env python3

import argparse
import os
import time
import datetime
import gspread
import pandas as pd
from logzero import logger
from moneyforward_driver import MoneyforwardDriver


def update_sheet(sheet: gspread.Worksheet, df_new: pd.DataFrame):
    """シートを更新する

    シートのデータとdf_newのデータを結合する。重複する日付のデータが有る場合はdf_newの内容で上書きする。
    その後、結合したデータをシートに貼り付ける。

    Args:
        sheet (gspread.Worksheet): 更新先のシート
        df_new (pd.DataFrame): _description_
    """
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
    sheet.update([df_paste.columns.values.tolist()] + df_paste.values.tolist(), "A1")
    return


def update(ss_name: str, cookie_path: str, date_since: datetime.date):
    """指定した月から現在までの収入・支出をスプレッドシートに更新する

    Args:
        date (datetime.date, optional): 月。指定しない場合は今月を指定する.
    """
    mf = MoneyforwardDriver(cookie_path)
    login_success = mf.login()

    if not login_success:
        logger.error("Login failed.")
        return

    mf.update()
    time.sleep(3)

    [income_new, expense_new] = mf.fetch_monthly_income_and_expenses_since(
        date_since.year, date_since.month
    )

    gc = gspread.service_account()
    ss = gc.open(ss_name)
    expense_sheet = ss.worksheet("支出")
    update_sheet(expense_sheet, expense_new)

    income_sheet = ss.worksheet("収入")
    update_sheet(income_sheet, income_new)
    return


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("ss_name", help="Name of the destination spreadsheet")
    parser.add_argument("-c", "--cookie-path", action="store", dest="cookie_path")
    parser.add_argument("--since", action="store")
    args = parser.parse_args()

    if args.since:
        date_since = datetime.date.fromisoformat(args.since)
    else:
        # 今日から2ヶ月前を指定
        date_since = datetime.date.today() - datetime.timedelta(days=60)

    update(args.ss_name, args.cookie_path, date_since)
