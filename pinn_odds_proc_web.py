# -*- coding: utf-8 -*-
"""
Created on Sat Jun 20 11:00:35 2026

@author: Luke

Fits expected total corners and corner supremacy from Pinnacle's Asian
Corners and Corner Handicap markets, using Poisson/Skellam distributions.
"""

import pandas as pd
from scipy.stats import poisson, skellam
from scipy.optimize import brentq


### Margin remove function

def remove_margin(odds_a, odds_b):
    p_a_raw = 1 / odds_a
    p_b_raw = 1 / odds_b
    total = p_a_raw + p_b_raw
    return p_a_raw / total, p_b_raw / total

### Fits Total Corners

def asian_over_probability(lam, line):
    remainder = line % 1.0

    if remainder == 0.0:
        return (1 - poisson.cdf(line, lam)) + 0.5 * poisson.pmf(line, lam)

    elif remainder == 0.5:
        return 1 - poisson.cdf(line - 0.5, lam)

    elif remainder == 0.25:
        lower = line - 0.25
        upper = line + 0.25
        p_lower = (1 - poisson.cdf(lower, lam)) + 0.5 * poisson.pmf(lower, lam)
        p_upper = 1 - poisson.cdf(upper - 0.5, lam)
        return 0.5 * (p_lower + p_upper)

    else:
        lower = line - 0.25
        upper = line + 0.25
        p_lower = 1 - poisson.cdf(lower - 0.5, lam)
        p_upper = (1 - poisson.cdf(upper, lam)) + 0.5 * poisson.pmf(upper, lam)
        return 0.5 * (p_lower + p_upper)


def fit_corner_lambda(line, over_odds, under_odds):
    p_over, _ = remove_margin(over_odds, under_odds)
    objective = lambda lam: asian_over_probability(lam, line) - p_over
    return brentq(objective, 1.0, 30.0)

### Fits Corner Supremacies

def skellam_home_prob(delta, line, base):
    mu_home = base + delta / 2
    mu_away = base - delta / 2

    mu_home = max(mu_home, 0.01)
    mu_away = max(mu_away, 0.01)

    remainder = line % 1.0
    if remainder < 0:
        remainder += 1.0

    if remainder == 0.0:
        p_win = 1 - skellam.cdf(line, mu_home, mu_away)
        p_push = skellam.pmf(int(line), mu_home, mu_away)
        return p_win + 0.5 * p_push

    elif remainder == 0.5:
        return 1 - skellam.cdf(line - 0.5, mu_home, mu_away)

    elif remainder == 0.25:
        lower = line - 0.25
        upper = line + 0.25
        p_lower = (1 - skellam.cdf(lower, mu_home, mu_away) +
                   0.5 * skellam.pmf(int(lower), mu_home, mu_away))
        p_upper = 1 - skellam.cdf(upper - 0.5, mu_home, mu_away)
        return 0.5 * (p_lower + p_upper)

    else:
        lower = line - 0.25
        upper = line + 0.25
        p_lower = 1 - skellam.cdf(lower - 0.5, mu_home, mu_away)
        p_upper = (1 - skellam.cdf(upper, mu_home, mu_away) +
                   0.5 * skellam.pmf(int(upper), mu_home, mu_away))
        return 0.5 * (p_lower + p_upper)


def fit_corner_delta(line, home_odds, away_odds, base):
    p_home, _ = remove_margin(home_odds, away_odds)
    objective = lambda delta: skellam_home_prob(delta, line, base) - p_home
    return brentq(objective, -15.0, 15.0)


def process_corner_odds(df):
    """
    Takes a raw odds dataframe (same columns as pinnacle_odds_data.xlsx)
    and returns a dataframe with fitted corner_total and corner_supremacy
    columns, keeping only rows with usable corner markets.
    """
    df_odds = df.copy()

    # Drops rows with no corner odds
    df_odds = df_odds.dropna(subset=['Asian Corners Line'])
    df_odds['Time of scraping'] = pd.to_datetime(df_odds['Time of scraping'])
    df_odds = df_odds.sort_values(by=['Game URL', 'Time of scraping'], ascending=[True, False])

    # Converts datetime format (Match Start Time is YYYY-MM-DD HH:MM, year-first)
    df_odds['Date'] = pd.to_datetime(
        df_odds['Match Start Time'],
        format='%Y-%m-%d %H:%M'
    ).dt.date

    # Removes unnecessary columns
    df_odds = df_odds[['Date', 'League', 'Home', 'Away', 'Asian Corners Line', 'Asian Corners O Odds', 'Asian Corners U Odds',
                                'Corner Handicap', 'Home Handicap Corners Odds', 'Away Handicap Corners Odds']]

    # Also drop rows missing the handicap side of the market - both legs are needed to fit
    df_odds = df_odds.dropna(subset=['Corner Handicap', 'Home Handicap Corners Odds', 'Away Handicap Corners Odds'])

    df_odds['corner_total'] = df_odds.apply(
        lambda row: fit_corner_lambda(row['Asian Corners Line'], row['Asian Corners O Odds'], row['Asian Corners U Odds']),
        axis=1
    )

    df_odds['hcap_line_home_perspective'] = -df_odds['Corner Handicap']

    df_odds['corner_supremacy'] = df_odds.apply(
        lambda row: fit_corner_delta(
            row['hcap_line_home_perspective'],
            row['Home Handicap Corners Odds'],
            row['Away Handicap Corners Odds'],
            base=row['corner_total'] / 2
        ),
        axis=1
    )

    df_odds = df_odds[['Date', 'League', 'Home', 'Away', 'corner_total', 'corner_supremacy']]

    df_odds['corner_total'] = df_odds['corner_total'].round(2)
    df_odds['corner_supremacy'] = df_odds['corner_supremacy'].round(2)

    return df_odds


if __name__ == '__main__':
    df = pd.read_csv('C:/Users/Luke/Desktop/Researchdocs/Scraper/pinnacle_fix/pinnacle_odds_data_2526.csv')
    df_odds = process_corner_odds(df)
