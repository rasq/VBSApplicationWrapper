'******************************************************************************
' Constants.vbs  the constant scripts for message  
'
' Description  : Implement the message Constants used in common script CallSetup.wft and script for package Setup.wsf.
'            
'
'******************************************************************************
'===================================================================
' Setup.wsf Application Screen URL
'===================================================================
Const ENTRY_REQUEST_URL = "http://gw3misc/rcv/Gate/EntryRequest.aspx"

'===================================================================
' Reference: Folders and Files
'===================================================================
' Folder where scripts are placed
Const SCRIPT_FOLDER     = "C:\ConfigMgr\Scripts"
' Folder where logs are stored
Const LOG_FOLDER        = "C:\etc\log"
' Name of flag file which indicates that the application has already been installed
Const FLAG_FILE_NAME_AP = "AP_INSTALLED"

' Log files
Const SETUP_WSF_LOG     = "SetupWsf.log"
Const CALL_SETUP_LOG    = "CallSetup.log"
Const WEB_INTERFACE_LOG = "WebInterface.log"
Const ENTRY_REQUEST_LOG = "EntryRequest.log"

' Folder where temporary files are output
Const TEMP_FOLDER = "C:\ConfigMgr\Temp"
' Password form HTML file
Const PASSWORD_FORM_HTML = "password.html"
' Application request result file
Const ENTRY_REQUEST_RESULT = "result"

' Script which opens application request screen
Const ENTRY_REQUEST_SCRIPT = "EntryRequest.wsf"
' PsExec
Const PSEXEC = "PsExec.exe"

'===================================================================
' Miscellaneous
'===================================================================
' Domain name
Const DOMAIN_NAME = "ASTELLAS"
' Result file output
Const SUCCESS = "SUCCESS"

'===================================================================
' Error codes
'===================================================================
Const ERRNO_AP_ALREADY_INSTALLED = -301
Const ERRNO_AP_NOT_INSTALLED     = -302
Const ERRNO_ENTRY_REQUEST_FAILED = -904

' Name of logging level
Const LOG_LEVEL_DEBUG = "DEBUG"
Const LOG_LEVEL_INFO  = "INFO"
Const LOG_LEVEL_WARN  = "WARN"
Const LOG_LEVEL_ERROR = "ERROR"
Const LOG_LEVEL_FATAL = "FATAL"

' Logging level
Const LOG_LEVEL_DEBUG_PRIORITY = 10
Const LOG_LEVEL_INFO_PRIORITY  = 20
Const LOG_LEVEL_WARN_PRIORITY  = 30
Const LOG_LEVEL_ERROR_PRIORITY = 40
Const LOG_LEVEL_FATAL_PRIORITY = 50

' Header texts of log for each logging level
Const LOG_LEVEL_DEBUG_PREFIX = "[DEBUG] "
Const LOG_LEVEL_INFO_PREFIX  = "[INFO ] "
Const LOG_LEVEL_WARN_PREFIX  = "[WARN ] "
Const LOG_LEVEL_ERROR_PREFIX = "[ERROR] "
Const LOG_LEVEL_FATAL_PREFIX = "[FATAL] "
