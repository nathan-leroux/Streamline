Attribute VB_Name = "TemplateConst"
'TemplateConst
'constants that are used throughout the program

Option Explicit

'instfile path
Public Const ROOT_DIR As String = "C:\Users\UDIB5\Desktop\local_file\Macros
Public Const INSTFILE_PATH As String = ROOT_DIR & "\instruction_file.txt"
Public Const LETTER_DIR As String = ROOT_DIR & "\Letters"

Public Const FILE_PATH_TEST As String = "C:\Users\UDIB5\Desktop\local_files\Macros\mock_current.txt"
Public Const FILE_PATH_LIVE As String = "C:\Users\UDIB5\Desktop\local_files\current.txt"

'special characters
Public Const TEXT_TAB As Integer = 9
Public Const TEXT_NEWLINE As Integer = 13

'error codes
Public Enum ActivityError
    DumbassError = 514
    CmdExpectedError = 515
    ArgExpectedError = 516
    BadInputError = 517
    BadDocumentError = 518
End Enum

'inbound flags
Public Enum InboundFlag
    ' inbound flags here ***
End Enum

'outbound codes
Public Enum OutcomeCode
    Yes
    YesApproval
    No
    NoApproval
End Enum

'outbound flags
Public Enum OutcomeFlag
    ' outcome flags here ***
End Enum
