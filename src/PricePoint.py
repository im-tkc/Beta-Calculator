#!/usr/bin/python

import sys

class PricePoint(object):
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
        self.percent_change = percent_change if percent_change is not None else None