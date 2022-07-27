# %%
import imp
from bokeh.models import Title
from math import pi
from turtle import title
from matplotlib import tight_layout
import pandas as pd
import hvplot
import holoviews as hv
from holoviews import opts
from bokeh.layouts import column
from bokeh.models import Div
from bokeh.plotting import show, save
import xlwings
import matplotlib.pyplot as plt
import numpy as np

pd.set_option("plotting.backend", "matplotlib")

# %%
df = pd.read_csv('./pwc_val_models_timeseries.csv', parse_dates=['Quarter'])
df['Quarter'] = df['Quarter'].pipe(pd.PeriodIndex, freq='Q')

# %%
# colors = plt.cm.GnBu(np.linspace(0, 1, 7))
fig, ax = plt.subplots(1,1, figsize=(16,9))
df.drop(['Amerihome', 'Survey'], axis=1).iloc[-6:].set_index('Quarter').mul(100).plot.bar(stacked=True, colormap='tab20b', ax=ax, rot=0)
handles, labels = ax.get_legend_handles_labels()
ax.legend(handles[::-1], labels[::-1], title='Line', loc='upper left', bbox_to_anchor=(1.01, 1))
# iterate through each bar container
for c in ax.containers[:-1]:
    # add the annotations
    ax.bar_label(c, fmt='%0.0f%%', label_type='center')
fig.savefig('pwc_val_models.png', transparent=True, pad_inches=0)
