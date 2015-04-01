Module GlobalConstants

Option Explicit On

    ' Used by Document construction (old) & form document construction (new)
Global Const XP_PRECEDENT_LOCATION = "c:\Documents and Settings\All Users\Documents\"
Global Const VISTA_PRECEDENT_LOCATION = "c:\Users\Public\Documents\"
Global Const WINDOWS_7_PRECEDENT_LOCATION = "c:\Users\Public\Documents\"
Global Const PRECEDENT_FOLDER = "zPrecedents"

    ' Modes that the main goDocument can be in
Global Const START_MODE = "StartMode"
Global Const CLIENT_MODE = "ClientMode"
Global Const HELP_MODE = "HelpMode"
Global Const DEFAULT_MODE = "DefaultMode"
Global Const ERROR_MODE = "ErrorMode"

    'Values that a function can return if being called by a base function
Global Const FUNCTION_ERROR = "FunctionError"
Global Const USER_EXIT = "UserExit"
Global Const NO_ERROR = "NoError"


End Module
