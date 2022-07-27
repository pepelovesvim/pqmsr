# %%
import geckodriver_autoinstaller 
from bokeh.models import Title
from math import pi
from turtle import title
import pandas as pd
import hvplot
import holoviews as hv
from holoviews import opts
from bokeh.layouts import column
from bokeh.models import Div
from bokeh.plotting import show, save
import xlwings

hv.extension("bokeh")

pd.set_option("plotting.backend", "holoviews")

# # %%
# excel_app = xlwings.App(visible=False)
# excel_book = excel_app.books.open('./data.xlsx')
# mults_sheet = excel_book.sheets['mults_raw']

# kpmg = pd.read_csv(
#     r"../historical_data/historical_kpmg_mults.csv",
# )

# # %%
# ah_response = mults_sheet.range('kpmg_mult[[#All]]').options(pd.DataFrame).value.reset_index().rename({'Multiples_Top 20': 'Category'}, axis=1).query('Quintile=="PPOfficial"')

# # %%
# ah_response.set_index(list(ah_response.columns[:3]))['Value'].unstack('Quintile').reset_index().rename({'PPOfficial': 'AH Response'})

# # %%
# commit_df = mults_sheet.range('kpmg_mult[[#All]]').options(pd.DataFrame).value.reset_index().rename({'Multiples_Top 20': 'Category'}, axis=1).query('Quintile!="PPOfficial" and Quintile != "High" and Quintile != "Low" and Quintile != "PHX" and Quintile != "MV"') \
# .set_index(list(ah_response.columns[:3]))['Value'].unstack() \
# .reset_index().merge(ah_response.rename({'Value': "AH Response"}, axis=1).drop("Quintile", axis=1), how='left')

# %%
kpmg = pd.read_csv(
    r"../historical_data/historical_kpmg_mults.csv",
    index_col=[0, 1],
    header=[0],
    parse_dates=[0],
    date_parser=lambda x: pd.to_datetime(x) + pd.offsets.QuarterEnd(0),
)
kpmg = kpmg.loc[kpmg.index.levels[0][-13:]]
pwc = pd.read_csv(
    r"../historical_data/historical_pwc_mults.csv",
    index_col=[0, 1],
    header=[0],
    parse_dates=[0],
    date_parser=lambda x: pd.to_datetime(x) + pd.offsets.QuarterEnd(0),
)
pwc = pwc.loc[pwc.index.levels[0][-13:]]

# %%
msrprognames = kpmg.index.levels[1]
plots = []
for prog in msrprognames:
    # p = kpmg.query(f"Category == 'Top 20' and MsrProgName == '{prog}'").plot(
    p = (
        kpmg.query("Category == 'Top 20'")
        .query(f'Product == "{prog}"')
        .reset_index("Product", drop=True)
        # .reindex(actual_index)
        .plot(
            x="Quarter",
            y=["25th Pct", "Median", "75th Pct", "AH Response"],
            stacked=True,
            legend="top_left",
            fontsize={"legend_title": 12},
            fontscale=2,
            height=900,
            width=900,
            toolbar=None,
        )
    )
    p.get_dimension("Variable").label = prog
    plots.append(p)
kpmg_top20 = (
    plots[0].opts(title="KPMG Top 20", fontsize={"title": 20})
    + plots[1]
    + plots[2]
    + plots[3]
)

msrprognames = pwc.index.levels[1]
plots = []
for prog in msrprognames:
    p = (
        # pwc.query(f"Category == 'Top 10' and MsrProgName == '{prog}'")
        pwc.query("Category == 'Top 10'")
        .query(f'Product == "{prog}"')
        .reset_index("Product", drop=True)
        # .reindex(actual_index)
        .plot(
            x="Quarter",
            y=["25th Pct", "Median", "75th Pct", "AH Response"],
            stacked=True,
            legend="top_left",
            fontsize={"legend_title": 12},
            fontscale=2,
            height=900,
            width=900,
            # rot=45,
            toolbar=None,
        )
    )
    p.get_dimension("Variable").label = prog
    plots.append(p)
pwc_top10 = (
    plots[0].opts(title="PwC Top 10", fontsize={"title": 20})
    + plots[1]
    + plots[2]
    + plots[3]
)
hvplot.save(
    (kpmg_top20 + pwc_top10).opts(toolbar=None), "top_mults.html", vspace=0.5, hspace=0.5
)

# %%
Products = kpmg.index.levels[1]
plots = []
for prog in Products:
    # p = kpmg.query(f"Product == 'Top 20' and Product == '{prog}'").plot(
    p = (
        kpmg.query("Category == 'NonBank'")
        .query(f'Product == "{prog}"')
        .reset_index('Product', drop=True)
        .plot(
            x="Quarter",
            y=["25th Pct", "Median", "75th Pct", "AH Response"],
            stacked=True,
            legend="top_left",
            fontsize={"legend_title": 12},
            fontscale=2,
            height=900,
            width=900,
            toolbar=None,
        )
    )
    p.get_dimension("Variable").label = prog
    plots.append(p)
kpmg_nonbank = (
    plots[0].opts(title="KPMG NonBank", fontsize={"title": 20})
    + plots[1]
    + plots[2]
    + plots[3]
)
Products = pwc.index.levels[1]
plots = []
for prog in Products:
    p = (
        pwc.query("Category == 'All'")
        .query(f'Product == "{prog}"')
        .reset_index('Product', drop=True)
        .plot(
            x="Quarter",
            y=["25th Pct", "Median", "75th Pct", "AH Response"],
            stacked=True,
            legend="top_left",
            fontsize={"legend_title": 12},
            fontscale=2,
            height=900,
            width=900,
            # rot=45,
            toolbar=None,
        )
    )
    p.get_dimension("Variable").label = prog
    plots.append(p)
pwc_all = (
    plots[0].opts(title="PwC All", fontsize={"title": 20})
    + plots[1]
    + plots[2]
    + plots[3]
)
hvplot.save(
    (kpmg_nonbank + pwc_all).opts(toolbar=None), "all_mults.html", vspace=0.5, hspace=0.5
)

# %%