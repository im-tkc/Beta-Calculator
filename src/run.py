#!/usr/bin/env python

import sys
from YahooToExcel import YahooToExcel
import argparse

def export(ticker, result, workbook=None):
    filename = "{0}.json".format(result.output)
    output_file = "{0}.xlsx".format(result.output)
    frequency = result.frequency # "1d","5d","1mo","3mo","6mo","1y","2y","5y","10y","max"
    
    y2e = YahooToExcel()
    workbook = y2e.run(filename, ticker, frequency, output_file, workbook)

    return workbook

def main(argv):
    if len(sys.argv) == 1:
        print("Beta calculator")
        print("    Usage: run.py -m <index_ticker> -c <company ticker> -f <frequency> -o <output file name>")
        sys.exit(1)
    
    parser = argparse.ArgumentParser(description="Beta calculator")
    parser.add_argument('-m', action='store', dest="index_ticker")
    parser.add_argument('-s', action='store', dest="company_ticker")
    parser.add_argument('-f', action='store', dest="frequency")
    parser.add_argument('-o', action='store', dest="output")
    
    result = parser.parse_args()
    workbook = export(result.index_ticker, result)
    workbook = export(result.company_ticker, result, workbook)
    
if __name__ == '__main__':
    main(sys.argv)