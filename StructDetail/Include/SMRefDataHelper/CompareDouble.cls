VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CompareDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Const TOLERANCE_VALUE = 0.0000001
'************************************************************************
' Method : Equal --- checkes whether two double variables have equal
'          values within a given tolerance
'
'  Inputs --------- Two double variables and tolerance(optional)
'
'  Output --------- True or False
'************************************************************************
Public Function Equal(ByVal LeftVariable As Double, ByVal RightVariable As Double, _
        Optional ByVal Tolerance As Double = TOLERANCE_VALUE) As Boolean
    If (LeftVariable >= (RightVariable - Tolerance)) And _
        (LeftVariable <= (RightVariable + Tolerance)) Then
        Equal = True
    Else
        Equal = False
    End If
End Function

'************************************************************************
' Method : GreaterThan --- checkes whether left side double variable is
'          greater than right side double variable, uses given tolerance
'
'  Inputs --------- Two double variables and tolerance(optional)
'
'  Output --------- True or False
'
'  NOTE: GreaterThanZero method need to be used if right side variable
'        is zero '0#' for comparison
'************************************************************************
Public Function GreaterThan(ByVal LeftVariable As Double, ByVal RightVariable As Double, _
        Optional ByVal Tolerance As Double = TOLERANCE_VALUE) As Boolean

    If (LeftVariable > (RightVariable - Tolerance)) Then
        GreaterThan = True
    Else
        GreaterThan = False
    End If

End Function

'************************************************************************
' Method : GreaterThanZero --- checkes whether left side double variable
'          is greater than zero '0#', uses given tolerance
'
'  Inputs --------- Left side double variable and tolerance(optional)
'
'  Output --------- True or False
'************************************************************************
Public Function GreaterThanZero(ByVal LeftVariable As Double, _
        Optional ByVal Tolerance As Double = TOLERANCE_VALUE) As Boolean

    If LeftVariable > Abs(Tolerance) Then
        GreaterThanZero = True
    Else
        GreaterThanZero = False
    End If

End Function

'************************************************************************
' Method : GreaterthanOrEqualTo --- checkes whether left side double
'          variable is greater than or equal to right side double variable
'          uses given tolerance
'
'  Inputs --------- Two double variables and tolerance(optional)
'
'  Output --------- True or False
'************************************************************************
Public Function GreaterThanOrEqualTo(ByVal LeftVariable As Double, ByVal RightVariable As Double, _
        Optional ByVal Tolerance As Double = TOLERANCE_VALUE) As Boolean

    If (LeftVariable >= (RightVariable - Tolerance)) Then
        GreaterThanOrEqualTo = True
    Else
        GreaterThanOrEqualTo = False
    End If

End Function

'************************************************************************
' Method : LessThan --- checkes whether left side double variable is
'          less than right side double variable, uses given tolerance
'
'  Inputs --------- Two double variables and tolerance(optional)
'
'  Output --------- True or False
'
'  NOTE: LessThanZero method need to be used if right side variable
'        is zero '0#' for comparison
'************************************************************************
Public Function LessThan(ByVal LeftVariable As Double, ByVal RightVariable As Double, _
        Optional ByVal Tolerance As Double = TOLERANCE_VALUE) As Boolean

    If (LeftVariable < (RightVariable + Tolerance)) Then
        LessThan = True
    Else
        LessThan = False
    End If

End Function

'************************************************************************
' Method : LessThanZero --- checkes whether left side double variable
'          is less than zero '0#', uses given tolerance
'
'  Inputs --------- Left side double variable and tolerance(optional)
'
'  Output --------- True or False
'************************************************************************

Public Function LessThanZero(ByVal LeftVariable As Double, _
        Optional ByVal Tolerance As Double = TOLERANCE_VALUE) As Boolean

    If LeftVariable < -Abs(Tolerance) Then
        LessThanZero = True
    Else
        LessThanZero = False
    End If

End Function

'************************************************************************
' Method : LessThanOrEqualTo --- checkes whether left side double
'          variable is less than or equal to right side double variable
'          uses given tolerance
'
'  Inputs --------- Two double variables and tolerance(optional)
'
'  Output --------- True or False
'************************************************************************
Public Function LessThanOrEqualTo(ByVal LeftVariable As Double, ByVal RightVariable As Double, _
        Optional ByVal Tolerance As Double = TOLERANCE_VALUE) As Boolean

    If (LeftVariable <= (RightVariable + Tolerance)) Then
        LessThanOrEqualTo = True
    Else
        LessThanOrEqualTo = False
    End If

End Function
