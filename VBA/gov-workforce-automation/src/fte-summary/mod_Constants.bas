Option Explicit
' mod_Constants.bas
' Centralized constants for branch names, status tokens, and column headers

Public Function BranchNames() As Variant
    BranchNames = Array( _
        "OCIA", "ATPAPMB", "CAB", "CENAB", "CSAB", "EAB", "EWDSAB", _
        "EWI&IT", "FIT", "HAB", "JAB", "DACoE", "RAB" _
    )
End Function

Public Function StatusTokensExcluded() As Variant
    ' Tokens that indicate a row should be excluded from FTE count
    StatusTokensExcluded = Array("(A/O)", "(LoA)", "(M/L)", "(S/O)", "(LTIP)", "(FxT)")
End Function

' Raw Data (input) column labels used by split pipeline
Public Const COL_FIRST_NAME As String = "AW"
Public Const COL_LAST_NAME  As String = "AX"
Public Const COL_POSITION   As String = "H"
Public Const COL_DEPT_ID    As String = "E"
Public Const COL_JOB_CODE   As String = "Z"
Public Const COL_POSITIONNO As String = "G"
Public Const COL_REPORTSTO  As String = "AH"
Public Const COL_INCUMBENCY As String = "AR"
Public Const COL_EXP_RETURN As String = "BU"
Public Const COL_CLASS_DESC As String = "BE"

Public Const SHEET_RAW_DATA As String = "Raw Data"
Public Const SHEET_BRANCH_LOOKUP As String = "Branch Identifier (NEW)"