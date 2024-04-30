import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
import pathlib

from TimeStruct import *


def createHistogram(df, params):
    def createHistogramInner(filename, selectedDirections, time_in: TimeStruct, selectedCategories, freq, yLim, drawBackgroundGrid,
                             graphTitle):
        data = computeData(df.copy(True), selectedDirections, time_in, selectedCategories)
        makePlot(data, filename, freq, yLim, drawBackgroundGrid, graphTitle)

    for idx, item in enumerate(params['figure_histogram'].values()):
        selectedDirections = item.get('selected_directions', [])
        if len(selectedDirections) == 0:
            selectedDirections = None

        time = TimeStruct.createFromDict(item, df)
        sc = item.get('selected_categories', [])
        if len(sc) == 0:
            sc = None

        freq = item.get('scaleAxisX', '30min')
        yLim = item.get('scaleAxisY', None)
        if yLim == 'auto':
            yLim = None

        drawBackgroundGrid = item.get('drawBackgroundGrid', True)
        graphTitle = item.get('graphTitle', 'Histogram počtu průjezdů')
        filename = item.get('filename', f'histogram_{idx:02}.png')
        filename = pathlib.Path(filename).absolute()
        filename.parent.mkdir(parents=True, exist_ok=True)

        createHistogramInner(filename, selectedDirections, time, sc, freq, yLim, drawBackgroundGrid, graphTitle)


def computeData(df, selectedDirections, time: TimeStruct, selectedCategories):
    # filter time
    df = df.set_index('Capture_time').sort_index()
    df = df[time.dateTimeStart: time.dateTimeEnd]

    # filter directions
    if selectedDirections is not None:
        df = df[df['Direction'].isin(selectedDirections)]

    if selectedCategories is not None:
        df = df[df['Vehicle_category'].isin(selectedCategories)]

    df = df.reset_index()
    return df


def makePlot(df, filename, freq, yLim, drawBackgroundGrid, graphTitle):
    def makePalette():
        pal = {}
        for i in range(20):
            pal[i + 1] = sns.color_palette("tab20b")[i]
        for i in range(20):
            pal[i + 1 + 19] = sns.color_palette("tab20c")[i]

        return pal

    def generateTimeStamps(freq):
        mmin = df['Capture_time'].min().floor(freq=freq)
        mmax = df['Capture_time'].max().ceil(freq=freq)
        periods = int((mmax - mmin) / pd.Timedelta(freq))
        bins = pd.date_range(start=mmin, end=mmax, periods=periods + 1)
        return periods, bins, (mmin, mmax)

    numBins, bins, (mmin, mmax) = generateTimeStamps(freq)

    # little hack to make data spaced correctly
    df.loc[df["Capture_time"] == df['Capture_time'].min(), "Capture_time"] = mmin
    df.loc[df["Capture_time"] == df['Capture_time'].max(), "Capture_time"] = mmax

    def generateBinNames():
        # keep long names only for midnight
        longBins = bins.strftime("%Y-%m-%d %H:%M")
        shortBins = bins.strftime("%H:%M")
        outBins = []
        lastPlacedIndex = 0
        for idx, (long, short) in enumerate(zip(longBins, shortBins)):
            if idx == 0 or idx == (len(longBins) - 1) or '00:00' in long:
                outBins.append(long)
                lastPlacedIndex = idx
            elif pd.Timedelta(freq) < pd.Timedelta('30m') and (idx - lastPlacedIndex) % 2 == 1:
                outBins.append("")
            else:
                outBins.append(short)
        return outBins

    plt.figure(figsize=(30, 18))
    sns.histplot(data=df, x='Capture_time', hue='Direction', multiple="stack", palette=makePalette(), stat='count',
                 bins=numBins)
    plt.xticks(bins, generateBinNames(), rotation='vertical')
    plt.ylim(yLim)

    ax = plt.gca()
    ax.tick_params(axis='both', labelsize=25)
    plt.title(graphTitle, fontsize=45, pad=20)
    plt.xlabel('Čas průjezdu', fontsize=35)
    plt.ylabel('Počet průjezdů', fontsize=35)
    ax.get_legend().set_title("Směr")
    plt.setp(ax.get_legend().get_texts(), fontsize='25')  # for legend text
    plt.setp(ax.get_legend().get_title(), fontsize='35')  # for legend title

    if drawBackgroundGrid:
        ax.set_axisbelow(True)
        ax.yaxis.grid(color='gray', linestyle='dashed')
        ax.xaxis.grid(color='gray', linestyle='dashed')

    print('\tSaving histogram to file: ', filename)
    plt.tight_layout()
    plt.savefig(filename)
