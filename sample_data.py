#!/usr/bin/python
import re
import xlwings as xw
import pandas as pd
from re import search
import matplotlib.pyplot as plt
import matplotlib.dates as mdates


def main():
    sheet = xw.Book.caller().sheets[0]
    val = sheet.range('A7').expand().options(pd.DataFrame)          # convert excel data to Pandas dataframe. The table has to start at cell A7!!!!
    df = pd.DataFrame(val.value)
    df['diff_sec'] = df.index.to_series().diff().dt.seconds

    # uniquify(df)          # not used at the moment, might be needed when there are several temperature channels

    # init some values
    has_temp, has_press = False, False
    evac_start, evac_end, kill_start, kill_end, dry_start, dry_end = -1, -1, -1, -1, -1, -1

    # for-loop to locate timestamps for each phase in sterilisation-process
    for enum, column in enumerate(df.columns):
        if search("°C", column, re.IGNORECASE):
            has_temp = True
        elif search("bar", column, re.IGNORECASE):
            has_press = True

            pressval = df.loc[:,column]
            tl = len(pressval)
            for j in range(tl):                                             # find evac_start
                if pressval.iloc[j] < 0.9:
                    evac_start = pressval.index[j]
                    tl = tl - j
                    for k in range(tl):                                     # find evac_stop
                        if pressval.iloc[k+j] > 1.6:
                            evac_end = pressval.index[k+j]
                            tl = tl - k
                            for l in range(tl):                             # find kill_start
                                if pressval.iloc[k + j + l] > 3.01:
                                    kill_start = pressval.index[k + j + l]
                                    tl = tl - l
                                    for m in range(tl):                     # find kill_stop
                                        if pressval.iloc[k + j + l + m] < 3:
                                            kill_end = pressval.index[k + j + l + m]
                                            tl = tl - m
                                            for n in range(tl):             # find dry_start
                                                if pressval.iloc[k + j + l + m + n] < 0.89:
                                                    dry_start = pressval.index[k + j + l + m + n]
                                                    tl = tl - n
                                                    for o in range(tl):     # find dry_stop
                                                        if pressval.iloc[k + j + l + m + n + o] > 0.91:
                                                            dry_end = pressval.index[k + j + l + m + n + o]
                                                            break
                                                    break

                                            break
                                    break
                            break
                    break

    # check if temperature and pressure channels are found
    if not has_temp or not has_press:
        sheet.range('H1').value = "Es fehlt entweder ein Temperatur oder Druck Kanal. Keine Auswertung möglich!"
        sheet.range('H1').autofit()
        sheet.range('H1').color = (255,0,0)
        return

    # set beginning and end of phases in order to plot them as vertical lines
    v_lines = [evac_start, evac_end, kill_start, kill_end, dry_start, dry_end]
    v_line_labels = ['evac_start', 'evac_end', 'steri_start', 'steri_end', 'dry_start', 'dry_end']
    v_line_colors = ['b', 'navy', 'r', 'darkred', 'g', 'darkgreen']

    # Plot temperature and pressure graphs
    for enum, column in enumerate(df.columns):
        fig, ax = plt.subplots(figsize=(12, 4))
        if search("°C", column, re.IGNORECASE):
            tempval = df.loc[:,column]
            timestep = df['diff_sec']

            A0 = getA0(tempval, timestep)
            F0 = getF0(tempval, timestep)

            sheet.range('H4').value = "A0-Wert: " + str(A0//60) + " min"
            sheet.range('H5').value = "F0-Wert: " + str(F0//60) + " min"

            ax.plot(tempval.index, tempval.values)
            for en, i in enumerate(v_lines):
                plt.axvline(x=i, color=v_line_colors[en], ls='--', lw=1.5, label=v_line_labels[en])
            plt.axhline(y=130, color='k', ls=':', lw=1, label="130°C")
            plt.title("Temperatur")
            plt.xlabel("Zeit")
            plt.ylabel("Temperatur in °C")
            plt.legend(bbox_to_anchor=(1.0, 1), loc='upper left')
            plt.autoscale()
            fig.autofmt_xdate()

            xfmt = mdates.DateFormatter('%H:%M:%S')
            ax.xaxis.set_major_formatter(xfmt)
            plt.tight_layout()
            sheet.pictures.add(fig, name='Plot1', anchor=sheet.range('H25'), update=True)

        elif search("bar", column, re.IGNORECASE):
            pressval = df.loc[:,column]
            ax.plot(pressval.index, pressval.values)
            for en, i in enumerate(v_lines):
                plt.axvline(x=i, color=v_line_colors[en], ls='--', lw=1.5, label=v_line_labels[en])
            plt.legend(bbox_to_anchor=(1.0, 1), loc='upper left')
            plt.title("Druck")
            plt.xlabel("Zeit")
            plt.ylabel("Druck in bar")
            plt.autoscale()
            fig.autofmt_xdate()

            xfmt = mdates.DateFormatter('%H:%M:%S')
            ax.xaxis.set_major_formatter(xfmt)
            plt.tight_layout()
            sheet.pictures.add(fig, name='Plot2', anchor=sheet.range('H6'), update=True)

    # write some values in the excel sheet
    sheet.range('H1').value = "Start-Zeit Evakuierung: " + str(evac_start)
    sheet.range('I1').value = "End-Zeit Evakuierung: " + str(evac_end)
    sheet.range('J1').value = "Dauer Evakuierung: " + str((evac_end - evac_start).seconds // 60) + "min, " + str((evac_end - evac_start).seconds % 60) + "sec"
    sheet.range('H2').value = "Start-Zeit Sterilisation: " + str(kill_start)
    sheet.range('I2').value = "End-Zeit Sterilisation: " + str(kill_end)
    sheet.range('J2').value = "Dauer Sterilisation (Druck > 3 bar): " + str((kill_end - kill_start).seconds // 60) + "min, " + str((kill_end - kill_start).seconds % 60) + "sec"
    sheet.range('H3').value = "Start-Zeit Trocknung: " + str(dry_start)
    sheet.range('I3').value = "End-Zeit Trocknung: " + str(dry_end)
    sheet.range('J3').value = "Dauer Trocknung: " + str((dry_end - dry_start).seconds // 60) + "min, " + str((dry_end - dry_start).seconds % 60) + "sec"
    sheet.range('H1:J3').columns.autofit()


def uniquify(df):
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique():
        cols[cols[cols == dup].index.values.tolist()] = [dup + ' ' + str(i+1) if i != 0 else dup for i in
                                                         range(sum(cols == dup))]
    # rename the columns with the cols list.
    df.columns = cols
    return df


def getA0(data, timestep):
    Tmin = 65       # Minimaltemperatur in °C
    Tref = 80       # Referenztemperatur in °C
    Z = 10          # Z-Wert in °C
    # tref = 1      # Referenzzeit in s
    return _getF(Tmin, Tref, Z, data, timestep)


def getF0(data, timestep):
    Tmin = 100       # Minimaltemperatur in °C
    Tref = 121.11111111       # Referenztemperatur in °C
    Z = 10          # Z-Wert in °C
    # tref = 1      # Referenzzeit in s
    return _getF(Tmin, Tref, Z, data, timestep)


def _getF(Tmin, Tref, Z, data, timestep):
    L = data.apply(lambda x: [10 ** ((y - Tref) / Z) if y > Tmin else 0 for y in x]).iloc[:, 0]
    return (L*timestep).sum()