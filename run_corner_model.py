# -*- coding: utf-8 -*-
"""
Daily runner for the corner odds model.

Reads today's pinnacle_odds_data.xlsx (output of pinnacle_write_data_2_new.py),
applies the corner_total / corner_supremacy fit, and produces two files:

  pinnacle_corners_current.xlsx     - just today's processed matches
  pinnacle_corners_historical.xlsx  - today's matches appended onto the
                                       existing historical file (if present)

The historical file is expected to already exist in the working directory
(downloaded from Google Drive by the workflow before this script runs).
If it doesn't exist, it's created fresh from today's data.
"""

import os
import pandas as pd
from pinn_odds_proc_web import process_corner_odds

RAW_ODDS_FILE = "outputs/pinnacle_odds_data.xlsx"
CURRENT_OUT_FILE = "outputs/pinnacle_corners_current.xlsx"
HISTORICAL_FILE = "outputs/pinnacle_corners_historical.xlsx"


def main():
    raw_df = pd.read_excel(RAW_ODDS_FILE)
    processed_df = process_corner_odds(raw_df)

    if processed_df.empty:
        print("No corner odds available in today's data, skipping output.")
        return

    # Save today's matches on their own
    processed_df.to_excel(CURRENT_OUT_FILE, index=False)
    print(f"Wrote {len(processed_df)} rows to {CURRENT_OUT_FILE}")

    # Append onto historical file
    if os.path.exists(HISTORICAL_FILE):
        historical_df = pd.read_excel(HISTORICAL_FILE)
        combined_df = pd.concat([historical_df, processed_df], ignore_index=True)
        combined_df['Date'] = pd.to_datetime(combined_df['Date']).dt.date
        combined_df = combined_df.drop_duplicates(subset=['Date', 'League', 'Home', 'Away'], keep='last')
    else:
        print(f"No existing {HISTORICAL_FILE} found, creating new one.")
        combined_df = processed_df

    combined_df = combined_df.sort_values(by=['Date', 'League', 'Home'])
    combined_df.to_excel(HISTORICAL_FILE, index=False)
    print(f"Historical file now has {len(combined_df)} total rows.")


if __name__ == '__main__':
    main()
