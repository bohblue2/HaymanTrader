import sys
import win32com.client
import numpy as np
import pandas as pd
import xlrd

'''
매수 수량 설정
매수, 매도 담당
'''

class core:

    ## CYBOS API  호출
    def __init__(self):
        self.objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')
        self.objStockOrder = win32com.client.Dispatch("CpTrade.CpTd0311")
        self.objTrade5331A = win32com.client.Dispatch("CpTrade.CpTdNew5331A")
        self.objTrade5331B = win32com.client.Dispatch("CpTrade.CpTdNew5331B")
        self.obj6033 = win32com.client.Dispatch("CpTrade.CpTd6033")
        self.objStockCur = win32com.client.Dispatch('Dscbo1.StockCur')
        self.objStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
        self.objStockMst = win32com.client.Dispatch("DsCbo1.StockMst")
        self.objCpTrade = win32com.client.Dispatch('CpTrade.CpTdUtil')

        self.initCheck = self.objCpTrade.TradeInit(0)

        if (self.initCheck != 0):
            print("주문 초기화 실패")
            return False

        self.accountnumber = self.objCpTrade.AccountNumber[0]  # 계좌번호
        
        
    def calculate_buy_stock_amount(self, price, code):
        accFlag = self.objCpTrade.GoodsList(self.accountnumber, 1)  # 주식상품 구분

        self.objTrade5331A.SetInputValue(0, self.accountnumber) # 계좌번호
        self.objTrade5331A.SetInputValue(1, accflag)
        self.objTrade5331A.SetInputValue(2, code) # 종목코드
        self.objTrade5331A.SetInputValue(3, '01') # 보통호가
        self.objTrade5331A.SetInputValue(4, int(price)) # 가격
        self.objTrade5331A.SetInputValue(6, 2)

        self.objTrade5331A.BlockRequest()

        money_to_buy = self.objTrade5331A.GetHeaderValue(18) # 현금 주문 가능수량
        amount = self.objTrade5331A.GetHeaderValue(45) # 잔고 호출
        
        ## 매수수량, 잔고 확인 및 리턴
        print(amount)
        print(money_to_buy)

        return money_to_buy       
        

    ## 매수(입력 : 코드, 수량, 가격)
    def buy(self, code, price):

        # 주문 초기화
        initCheck = self.objCpTrade.TradeInit(0)
        if (initCheck != 0):
            print("주문 초기화 실패")
            exit()        
 
        # 주식 매수 주문
        accFlag = self.objCpTrade.GoodsList(self.accountnumber, 1)   # 주식상품 구분

        self.objStockOrder.SetInputValue(0, "2")   # 2: 매수
        self.objStockOrder.SetInputValue(1, self.accountnumber)   # 계좌번호
        self.objStockOrder.SetInputValue(2, accFlag[0])   # 상품구분
        self.objStockOrder.SetInputValue(3, code)   # 종목코드
        self.objStockOrder.SetInputValue(4, self.calculate_buy_stock_amount(price, code)) # 매수수량
        self.objStockOrder.SetInputValue(5, int(price))   # 주문단가
        self.objStockOrder.SetInputValue(7, "0")   # 주문 조건 구분 코드, 0: 기본
        self.objStockOrder.SetInputValue(8, "01")  # 주문호가 구분코드 - 01: 보통
 
        # 매수 주문 요청
        self.objStockOrder.BlockRequest()
 
        rqStatus = self.objStockOrder.GetDibStatus()
        rqRet = self.objStockOrder.GetDibMsg1()
        
        print("통신상태", rqStatus, rqRet)

        if rqStatus != 0:
            exit()
            

    ## 매도(입력 : 코드, 가격)
    def sell(self, code, amount, price):
 
        # 주문 초기화
        initCheck = self.objCpTrade.TradeInit(0)
        if (initCheck != 0):
            print("주문 초기화 실패")
            exit()

        accFlag = self.objCpTrade.GoodsList(self.accountnumber, 1) # 주식상품 구분

        self.objStockOrder.SetInputValue(0, "1")  # 1: 매도
        self.objStockOrder.SetInputValue(1, self.accountnumber) # 계좌번호
        self.objStockOrder.SetInputValue(2, accFlag[0]) # 상품구분 - 주식 상품 중 첫번째
        self.objStockOrder.SetInputValue(3, code)       # 종목코드
        self.objStockOrder.SetInputValue(4, amount)     # 매도수량
        self.objStockOrder.SetInputValue(5, int(price)) # 주문단가 
        self.objStockOrder.SetInputValue(7, "0")   # 주문 조건 구분 코드, 0: 기본
        self.objStockOrder.SetInputValue(8, "01")  # 주문호가 구분코드 - 01: 지정가
 
        # 매도 주문 요청
        self.objStockOrder.BlockRequest()
 
        rqStatus = self.objStockOrder.GetDibStatus()
        rqRet = self.objStockOrder.GetDibMsg1()

        print("통신상태", rqStatus, rqRet)

        if rqStatus != 0:
            pass
