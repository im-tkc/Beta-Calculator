#!/usr/bin/env python

import sys
import json
import urllib.request
import datetime
import time
import csv
from collections import deque

class PricePoint:
    timestamp = None
    open_price = 0.0
    high_price = 0.0
    low_price = 0.0
    close_price = 0.0
    adj_close_price = 0.0
    volume = 0
    percent_change = 0.0
    
    def __init__(self, timestamp, open, high, low, close, adj_close, volume):
        self.timestamp = timestamp if timestamp is not "null" else None
        self.open_price = open if open is not "null" else 0
        self.high_price = high if high is not "null" else 0
        self.low_price = low if low is not "null" else 0
        self.close_price = close if close is not "null" else 0
        self.adj_close_price = adj_close if adj_close is not "null" else 0
        self.volume = volume if volume is not "null" else 0
        
def to_datetime(unixtime):
    return datetime.datetime.fromtimestamp(unixtime).strftime('%Y-%m-%d %H:%M:%S')

def download(filename):
    url = "https://ycharts.com/charts/fund_data.json?securities=id%3AAIG%2Cinclude%3Atrue%2C%2C&calcs=id%3Aprice_to_book_value%2Cinclude%3Atrue%2C%2C&correlations=&format=real&recessions=false&zoom=100&startDate=&endDate=&chartView=&splitType=&scaleType=&note=&title=&source=&units=&quoteLegend=&partner=&quotes=&legendOnChart=&securitylistSecurityId=&clientGroupLogoUrl=&displayTicker=&ychartsLogo=&maxPoints=100000"
    urllib.request.urlretrieve(url, filename)
    
def get_json(filename):
    raw_data = None
    with open(filename, 'rb') as file:
        raw_data=file.read()
    return json.loads(raw_data)
    
    
def export_csv(json_data, filename):
    with open("output.csv", "w",  newline='') as fout:
        writer = csv.writer(fout, quoting=csv.QUOTE_MINIMAL)
        writer.writerow(["date", "pb"])
        resultSet = json_data['chart_data'][0][0]['raw_data']
        for i in range(0, len(json_data['chart_data'][0][0]['raw_data'])):
            unixtime = resultSet[i][0] / 1000
            print(unixtime)
            pb = resultSet[i][1]
            datetime = to_datetime(unixtime)
            is_valid_input = (pb is not None)
            
            if is_valid_input:
                writer.writerow([datetime, pb])
                
            # pricepoint = PricePoint(datetime, open_price, high_price, low_price, close_price, adj_close_price, volume)
            # queue.append(pricepoint)
            
            
def main(argv):
    filename = "PB.json"
    # current_time = time.mktime(datetime.datetime.now())
    # current_unixtime = current_time.strftime("%s")
    # print(int(current_time))
    download(filename)
    json_data = get_json(filename)
    # queue = deque()
    export_csv(json_data, filename)
    
if __name__ == '__main__':
    main(sys.argv)
