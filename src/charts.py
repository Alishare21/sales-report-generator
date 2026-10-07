"""Deterministic, accessible PNG chart rendering."""
from pathlib import Path
import textwrap

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np
import pandas as pd

PALETTE = ['#0072B2', '#E69F00', '#009E73', '#CC79A7', '#56B4E9', '#D55E00']


def build_charts(findings, directory):
    directory = Path(directory)
    directory.mkdir(parents=True, exist_ok=True)
    plt.rcParams.update({'font.size':11, 'axes.titlesize':13, 'axes.labelsize':11,
                         'xtick.labelsize':10, 'ytick.labelsize':10, 'font.family':'DejaVu Sans'})
    for f in findings:
        fig, ax = plt.subplots(figsize=(8,4.2), layout='constrained')
        s = f['series']
        if f['kind'] in {'bar','barh'}:
            labels = [textwrap.fill(str(x), 24) for x in s['labels']]
            if f['kind'] == 'barh':
                ax.barh(labels, s['values'], color=PALETTE[0])
                ax.invert_yaxis()
                ax.set_xlim(left=min(0,min(s['values'])))
                ax.grid(axis='x', alpha=.18)
            else:
                ax.bar(labels, s['values'], color=PALETTE[:len(labels)])
                ax.set_ylim(bottom=min(0,min(s['values'])))
                ax.grid(axis='y', alpha=.18)
        elif f['kind'] == 'line':
            dates = pd.to_datetime(s['labels'])
            line = pd.Series(s['values'], index=dates).reindex(pd.date_range(dates.min(), dates.max(),freq='D'))
            ax.plot(line.index, line.values, color=PALETTE[0], marker='o', markersize=4)
            locator = mdates.AutoDateLocator(minticks=3, maxticks=7)
            ax.xaxis.set_major_locator(locator)
            ax.xaxis.set_major_formatter(mdates.ConciseDateFormatter(locator))
            ax.grid(axis='y', alpha=.18)
        else:
            ax.scatter(s['x'],s['y'], color=PALETTE[0], alpha=.3, s=18, edgecolors='none')
            ax.grid(alpha=.18)
        ax.set_title(textwrap.fill(f['title'], 75), loc='left', pad=14, fontweight='bold')
        ax.set_xlabel(f['x_label'])
        ax.set_ylabel(f['y_label'])
        ax.spines[['top','right']].set_visible(False)
        path = directory / f"figure_{f['figure']:02d}.png"
        fig.savefig(path, dpi=200, facecolor='white')
        plt.close(fig)
        f['chart'] = str(path.resolve())
