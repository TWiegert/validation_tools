#!/usr/bin/python
import xlwings as xw
import pandas as pd



def main():
    wb = xw.Book.caller()
    df = pd.read_csv(r'C:\temp\TestData.csv')
    df['total_length'] =  df['sepal_length_(cm)'] + df['petal_length_(cm)']
    wb.sheets[0].range('A1').value = df