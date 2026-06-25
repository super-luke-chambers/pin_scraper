# -*- coding: utf-8 -*-
"""
Daily runner for the corner odds model.

Reads today's pinnacle_odds_data.xlsx (output of pinnacle_write_data_2_new.py),
applies the corner_total / corner_supremacy fit, and produces:

  pinnacle_corners_current/pinnacle_corners_ddmmyyyy.xlsx
      - just today's processed matches, in a new dated file each day,
        sorted by league then date.
  pinnacle_corners_historical.xlsx
      - a running record of every match ever scraped, sorted by date then
        league. If a match (Date/League/Home/Away) is scraped again, its row
        is updated in place with the latest values rather than duplicated.

The historical file is expected to already exist in the working directory
(committed to the repo / downloaded by the workflow before this script runs).
If it doesn't exist, it's created fresh from today's data.
"""

import os
from datetime import date
import pandas as pd
from pinn_odds_proc_web import process_corner_odds

RAW_ODDS_FILE = "outputs/pinnacle_odds_data.xlsx"
CURRENT_OUT_DIR = "outputs/pinnacle_corners_current"
HISTORICAL_FILE = "outputs/pinnacle_corners_historical.xlsx"


def main():
    raw_df = pd.read_excel(RAW_ODDS_FILE)
    processed_df = process_corner_odds(raw_df)

    if processed_df.empty:
        print("No corner odds available in today's data, skipping output.")
        return

    # Save today's matches to a new dated file, sorted by league then date
    os.makedirs(CURRENT_OUT_DIR, exist_ok=True)
    current_df = processed_df.sort_values(by=['League', 'Date', 'Home'])
    today_str = date.today().strftime('%d%m%Y')
    current_out_file = os.path.join(
        CURRENT_OUT_DIR, f"pinnacle_corners_{today_str}.xlsx"
    )
    current_df.to_excel(current_out_file, index=False)
    print(f"Wrote {len(current_df)} rows to {current_out_file}")

    # Update the running historical record
    if os.path.exists(HISTORICAL_FILE):
        historical_df = pd.read_excel(HISTORICAL_FILE)
        # processed_df is concatenated last so keep='last' lets a re-scraped
        # match overwrite its earlier row with the updated values.
        combined_df = pd.concat([historical_df, processed_df], ignore_index=True)
        combined_df['Date'] = pd.to_datetime(combined_df['Date']).dt.date
        combined_df = combined_df.drop_duplicates(subset=['Date', 'League', 'Home', 'Away'], keep='last')
    else:
        print(f"No existing {HISTORICAL_FILE} found, creating new one.")
        combined_df = processed_df

    # Historical sorted by date then league
    combined_df = combined_df.sort_values(by=['Date', 'League', 'Home'])
    combined_df.to_excel(HISTORICAL_FILE, index=False)
    print(f"Historical file now has {len(combined_df)} total rows.")


if __name__ == '__main__':
    main()
