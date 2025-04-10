Attribute VB_Name = "TemplateConst"
'TemplateConst
'constants that are used throughout the program

Option Explicit

'instfile path



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

'outbound codes
Public Enum OutcomeCode
    yes
    YesApproval
    no
    NoApproval
End Enum

'outbound flags
Public Enum OutcomeFlag
    Details
End Enum
