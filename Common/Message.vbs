'******************************************************************************
' Constants.vbs  the constant scripts for message  
'
' Description  : Implement the message Constants used in common script CallSetup.wft and script for package Setup.wsf.
'            
'
'******************************************************************************
'===================================================================
' the list of log output messages 
'===================================================================
Const START_FUNCTION               = "start the function {0}."
Const END_FUNCTION                 = "finish the function{0}."
Const START_SUB                    = "start the procedure {0}."
Const END_SUB                      = "finish the procedure{0}."

Const SART_SCRIPT                  = "start the script{0}."
Const END_SCRIPT                   = "finish the script{0}."
Const SCRIPT_PARAMETERS            = "parameters specified while starting the script {0} are as followings."
Const SCRIPT_PARAMETER_FORMAT      = "parameter : {0} / value : {1}"

Const MANDATORY_ARGUMENT_NOT_FOUND = "The essential parameters of the script  {0}({1}) are not specified. "
Const INVALID_ARGUMENT             = "The invalid {1}of {0} is specified."

Const PKGSCRIPT_EXEC_JOBID   = "The ID of the job to start the package script: {0}"
Const APSCRIPT_EXEC_JOBID    = "The ID of the job to start the application script : {0}"
Const RESOLVED_PKG_DIST_ROOT = "The root folder of the resolved package: {0}"
Const RESOLVED_AP_DIST_ROOT  = "The root folder of the resolved application: {0}"
Const CUSTOM_PKGSCRIPT_NOT_FOUND = "The custom package script {0} is not found."
Const CUSTOM_APSCRIPT_NOT_FOUND = "The custom application script {0} is not found."
Const AP_NOT_FOUND           = "The application is not found. The folder {0} is not found. "
Const PKGSCRIPT_FILE_PATH    = "the package script file path : {0}"
Const APSCRIPT_FILE_PATH     = " the application script file path: {0}"
Const PKGSCRIPT_EXEC_CMDLINE = "the command line string to start the package script: {0}"
Const APSCRIPT_EXEC_CMDLINE  = "the command line string to start the application script: {0}"
Const START_PKGSCRIPT        = "Start the package script."
Const START_APSCRIPT         = "Start the application script."
Const START_PKG_INSTALL      = "Start the software installation."
Const START_AP_INSTALL       = "Start the application installation."
Const PKGSCRIPT_FINISHED     = "The package script has been finished normally."
Const APSCRIPT_FINISHED      = "The application script has been finished normally."
Const PKG_INSTALL_SUCCEEDED  = "The software has been installed successfully. "
Const AP_INSTALL_SUCCEEDED   = "The application has been installed successfully. "
Const AP_UNINSTALL_SUCCEEDED = "The application has been uninstalled successfully"

Const PKGSCRIPT_FAILED       = "Problems occurred while implementing the package script. Return code: {0}"
Const APSCRIPT_FAILED        = "Problems occurred while implementing the application script. Return code: {0}"
Const PKG_INSTALL_FAILED     = "Problems occurred while installing the software."
Const AP_INSTALL_FAILED      = "Problems occurred while installing the application. "
Const AP_UNINSTALL_FAILED    = "Problems occurred while uninstalling the application."
Const RETURN_CODE            = "Return code : {0}"
Const PKG_INSTALL_FINISHED   = "The software installation has been finished. "
Const AP_INSTALL_FINISHED    = "The application installation has been finished. "
Const AP_UNINSTALL_FINISHED  = "The application uninstallation has been finished. "

Const INSTALL_EXEC_CMDLINE    = "The command line string to start the installation : {0}"
Const START_INSTALL           = "Start the installation. "
Const INSTALL_SUCCEEDED       = "The installation has been finished normally. "
Const UNINSTALL_EXEC_CMDLINE  = "The command line string to start uninstallation : {0}"
Const START_UNINSTALL         = "Start the uninstallation. "
Const UNINSTALL_SUCCEEDED     = "The uninstallation has been finished normally. "
Const UNINSTALL1_EXEC_CMDLINE = "The command line string to start the uninstallation (1st process): {0}"
Const START_UNINSTALL1        = "Start the uninstallation (1st process)."
Const UNINSTALL1_SUCCEEDED    = "The uninstallation (1st process) has been finished normally. "
Const UNINSTALL2_EXEC_CMDLINE = "The command line string to start the uninstallation (2nd process): {0}"
Const START_UNINSTALL2        = "Start the uninstallation (2nd process)."
Const UNINSTALL2_SUCCEEDED    = "The uninstallation (2nd process) has been finished normally."

Const INSTALL_FAILED          = "Problems occurred during installation. Return code: {0}"
Const INSTALL_FINISHED        = "The installation has been finished. "
Const UNINSTALL_FAILED        = "Problems occurred during uninstallation. Return code: {0}"
Const UNINSTALL_FINISHED      = " The uninstallation has been finished. "
Const UNINSTALL1_FAILED       = "Problems occurred during uninstallation (1st process). Return code: {0}"
Const UNINSTALL1_FINISHED     = "The uninstallation (1st process) has been finished. "
Const UNINSTALL2_FAILED       = "Problems occurred during uninstallation (2nd process). Return code: {0}"
Const UNINSTALL2_FINISHED     = "The uninstallation (2nd process) has been finished. "

Const FLAG_FILE_GENERATED       = "Create the file {0} that could judge whether the package has been installed."
Const PACKAGE_ALREADY_INSTALLED = "The package {0} - {1} has been installed. "
Const STOP_PUSH_INSTALL         = "Stop the Push installation implementation of the package{0} - {1}."
Const FLAG_FILE_EXISTS          = "There has been the file {0} that could judge whether the package has been installed. "
Const FLAG_FILE_GENERATE_FAILED = "Problems occurred whiling creating the file {0} that could judge whether the package has been installed."

Const FLAG_FILE_GENERATED_AP    = "The file {0} that could judge whether the package has been installed has been created."
Const FLAG_FILE_DELETED_AP      = "The file {0} that could judge whether the package has been installed has been deleted."
Const FLAG_FILE_EXISTS_AP       = "There has been the file {0} that could judge whether the package has been installed."
Const FLAG_FILE_NOT_EXISTS_AP   = "There is no file {0} that could judge whether the package has been installed."
Const AP_ALREADY_INSTALLED      = "The applications {0} - {1} have been installed."
Const AP_NOT_INSTALLED          = "There are no applications {0} - {1}."
Const STOP_AP_INSTALL           = "Stop the installation of applications {0} - {1}."
Const STOP_AP_UNINSTALL         = "Stop the uninstallation of application {0} - {1}."
Const FLAG_FILE_GENERATE_FAILED_AP = "Problems occurred while creating the file {0} that could judge whether the package has been installed."
Const FLAG_FILE_DELETE_FAILED_AP = "Problems occurred while deleting the file {0} that could judge whether the package has been installed."

Const ENTRY_REQUEST_START = "Show the installation application screen. "
Const ENTRY_REQUEST_END   = "The input to the installation application screen has been finished normally. "

Const ENTRY_REQUEST_FAILED           = "The installation application has been failed. "
Const ENTRY_REQUEST_WITH_PASS_START = " Start the installation application screen which is necessary to input the password. "
Const ENTRY_REQUEST_WITH_PASS_END   = "The installation application screen which is necessary to input the password has been closed. "

Const CREATE_PASSWORD_FORM_START = "Create the HTML file for the password input form. "
Const CREATE_PASSWORD_FORM_END = "The HTML file for the password input form has been created. "
Const INPUT_PASSWORD_START = "Display the form to input the password. "
Const INPUT_PASSWORD_END = "The user has input the password. "
Const NO_EXISTS_ERROR = "{0} is not existed. "
Const FILE_DELETE_ERROR = "Deletion of {0} has been failed. "

Const DELETE_TEMP_FILE_START = "Delete {0}. "
Const DELETE_TEMP_FILE_END = "{0} has been deleted. "

Const ENTRY_REQUEST_EXEC_CMDLINE    = "Command-line character string to start the installation application screen : {0}"
Const ENTRY_REQUEST_EXEC_USER    = "User to start the installation application screen : {0}"

Const CREATE_RESULT_START = "Create the result file."
Const CREATE_RESULT_END = "The result file has been created. "

Const CHECK_RESULT_START = "Confirm the result file. "
Const CHECK_RESULT_END = "The result file has been confirmed. "
Const CHECL_RESULT_SUCCESS = "The content of the result file: succeeded. "
Const CHECK_RESULT_FAILED = "The content of the result file: failed. "

'===================================================================
'the list of message box 
'===================================================================
Const NO_APPLICATION_FORM_MESSAGE_1 = "to install the application ,"
Const NO_APPLICATION_FORM_MESSAGE_2 = "「Application Form for Installing Commercial Business Software 」 is necessary."
Const NO_APPLICATION_FORM_MESSAGE_3 = "Please click the [OK] button and reinstall after the application. "
Const NO_APPLICATION_FORM_MESSAGE_4 = "standard application installation"

Const ENTRY_REQUEST_FAILED_MESSAGE_1 = "The installation application has been failed. "
Const ENTRY_REQUEST_FAILED_MESSAGE_2 = "Please click the [OK] button and reinstall from the beginning. "
Const ENTRY_REQUEST_FAILED_MESSAGE_3 = "installation application"

Const TIMEOUT_MESSAGE_1 = "The application for installation has been time out."
Const TIMEOUT_MESSAGE_2 = "time-out"

Const ERROR_TITLE = "Error"
Const CANCEL = "Cancel"

Const ENTRY_REQUEST_INTTERRUPUTION = "The installation application has been stopped. "
Const UNFORESEEN_ERROR = "Unexpected error occurred. "
Const PASSWORD_RETRY_ERROR = "You have input password for more then {0} times. "
Const PASSWORD_CANCEL = "Connection has been cancelled."
Const PASSWORD_TIMEOUT = "Connection has timed out."