'******************************************************************************
' RegeionConstants.vbs  the constant script of each region 
'
' Description  : Implement constants of each region  to the common script CallSetup.wsf  and the script for package Setup.wsf.
'
'******************************************************************************

'===================================================================
' monitoring time after calling the Setup.wsf 
'===================================================================
'Monitoring time. Default: 1second（1sec×1000msec）
Const MONITORING_TIME = 1000
'The waiting time before the window is closed in the case of the successful completion. Default: 5seconds（5sec×1000msec）
Const CLOSE_TIME = 5000
'Time-out period of the installation application screen. Default:  3hours（3h×60min×60sec×1000msec）
Const MAXLOOPCNT = 10800000

'===================================================================
' form to input the password
'===================================================================
'with or without the password input 
Const IS_PASSWORD_ENTER = False
'Password - retry. Default: 3 times. As the account will be locked after 5 continuous unsuccessful tries of the password, the value should be set less than 5.
Const PASSWORD_RETRY_COUNT = 3