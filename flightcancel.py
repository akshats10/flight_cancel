import openpyxl
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import statsmodels.api as sm


def readfile():
	
	wb  =  openpyxl.load_workbook ( 'CancelledFlights.xlsx' )
	sheet  =  wb.get_sheet_by_name ( 'CancelledFlights' )
	Canceled = []
	Month = []
	DepartureTime = []
	Airline = []
	ScheduledFlightTime = []
	ArrDelay = []
	DepDelay = []
	Dist = []
	
	for i in range ( 6000 ):
		Canceled.append ( int ( sheet.cell(row = i + 1,   column = 1 ).value ))
		Month.append ( float ( sheet.cell ( row = i + 1,  column = 2 ).value ))
		DepartureTime.append ( sheet.cell ( row = i + 1,  column = 3 ).value )
		Airline.append ( str ( sheet.cell ( row = i + 1,  column = 4 ).value ))
		ScheduledFlightTime.append ( float ( sheet.cell ( row = i + 1,  column = 5 ).value ))
		ArrDelay.append ( int ( sheet.cell ( row = i + 1,  column = 6 ).value ))
		DepDelay.append ( int ( sheet.cell ( row = i + 1,  column = 7 ).value ))
		Dist.append ( int ( sheet.cell ( row = i + 1,  column = 8 ).value ))
	
	return ( Canceled, Month, DepartureTime, Airline, ScheduledFlightTime, ArrDelay, DepDelay, Dist )


def reg_m(y, x):
    
    ones = np.ones(len(x[0]))
    X = sm.add_constant(np.column_stack((x[0], ones)))
    
    for ele in x[1:]:
        X = sm.add_constant(np.column_stack((ele, X)))
    
    results = sm.OLS(y, X).fit()
    
    return results


def CheckCancel( month, flighttime, distance):
	
	Canceled, Month, DepartureTime, Airline, ScheduledFlightTime, ArrDelay, DepDelay, Dist = readfile()
	m_aa=[]; t_aa=[]; d_aa=[]
	m_ua=[]; t_ua=[]; d_ua=[]
	m_dl=[]; t_dl=[]; d_dl=[] 
	Y_aa=[]; Y_ua=[]; Y_dl=[]
	
	for i in range (6000):
		if ( Airline[i].lower() == 'aa' ):
			m_aa.append ( Month[i] )
			t_aa.append ( ScheduledFlightTime[i] )
			d_aa.append ( Dist[i] )
			X_aa = [m_aa,t_aa,d_aa]
			Y_aa.append ( Canceled[i] )
		elif ( Airline[i].lower() == 'ua' ):
			m_ua.append ( Month[i] )
			t_ua.append ( ScheduledFlightTime[i] )
			d_ua.append ( Dist[i] )
			X_ua = [m_ua,t_ua,d_ua]
			Y_ua.append ( Canceled[i] )
		else:
			m_dl.append ( Month[i] )
			t_dl.append ( ScheduledFlightTime[i] )
			d_dl.append ( Dist[i] )
			X_dl = [m_dl,t_dl,d_dl]
			Y_dl.append ( Canceled[i] )

	prediction_aa = reg_m ( Y_aa, X_aa ).params[2]*month + reg_m ( Y_aa, X_aa ).params[1]*flighttime + reg_m ( Y_aa, X_aa ).params[0]*distance + reg_m ( Y_aa, X_aa ).params[3]
	
	if ( prediction_aa >= 1 ):
		prediction_aa = 1
	
	print ( '\nThe chances of flight AA being cancelled is: ', prediction_aa*100, "%" )
	prediction_ua = reg_m(Y_ua, X_ua).params[2]*month + reg_m(Y_ua, X_ua).params[1]*flighttime + reg_m(Y_ua, X_ua).params[0]*distance + reg_m(Y_ua, X_ua).params[3]
	
	if ( prediction_ua >= 1 ):
		prediction_ua = 1
	
	print ( 'The chances of flight UA being cancelled is: ', prediction_ua*100, "%" )
	prediction_dl = reg_m(Y_dl, X_dl).params[2]*month + reg_m(Y_dl, X_dl).params[1]*flighttime + reg_m(Y_dl, X_dl).params[0]*distance + reg_m(Y_dl, X_dl).params[3]
	
	if ( prediction_dl >= 1 ):
		prediction_dl = 1
	
	print ( 'The chances of flight DL being cancelled is: ', prediction_dl*100, "%\n" )
	
	CheckDelay( month )


def CheckDelay( month ):
	
	Canceled, Month, DepartureTime, Airline, ScheduledFlightTime, ArrDelay, DepDelay, Dist = readfile()
	Delay_aa  =  0
	counter_aa  =  0
	Delay_ua  =  0
	counter_ua  =  0
	Delay_dl  =  0
	counter_dl  =  0

	for i in range( 6000 ):
		if ( Month[i] == month and Airline[i].lower() == 'aa' ):
			Delay_aa = Delay_aa + abs( ArrDelay[i] + DepDelay[i] )
			counter_aa = counter_aa + 1
		elif ( Month[i] == month and Airline[i].lower() == 'ua' ):
			Delay_ua = Delay_ua + abs( ArrDelay[i] + DepDelay[i] )
			counter_ua = counter_ua + 1
		if ( Month[i] == month and Airline[i].lower() == 'dl' ):
			Delay_dl = Delay_dl + abs( ArrDelay[i] + DepDelay[i] )
			counter_dl = counter_dl + 1
	
	print ( "Flight AA is most likely to be delayed by around {0:.3f}".format( Delay_aa / ( 60 * counter_aa )), "hours" )
	print ( "Flight UA is most likely to be delayed by around {0:.3f}".format( Delay_ua / ( 60 * counter_ua )), "hours" )
	print ( "Flight DL is most likely to be delayed by around {0:.3f}".format( Delay_dl / ( 60 * counter_dl )), "hours\n" )


def showairlines():
	
	print ( "\nThe available flights are:" )
	print ( "\nAmerican Airlines (AA)" )
	print ( "Delta Airlines (DL)" )
	print ( "United Airlines (UA)\n" )


if __name__ == '__main__':
	
	month = int ( input ( "Enter the month number: " ))
	while ( int ( month ) < 1 or int ( month ) > 12 ):
		month = int ( input ( "Please enter a valid month number: " ))
	
	flighttime=int ( input("Enter the scheduled flight time in minutes: "))
	
	distance=int ( input("Enter the flight distance in kms: " ))
	
	showairlines()
	
	CheckCancel ( month,flighttime, distance)
			


# source: http://stackoverflow.com/questions/11479064/multiple-linear-regression-in-python






