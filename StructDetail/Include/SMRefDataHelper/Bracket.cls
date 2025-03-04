VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bracket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const MODULE As String = "S:\StructDetail\Data\Include\SMRefDataHelper\Bracket.cls"



'************************************************************************
'  For Tripping, Bracket By plane
' Method : IsBracket ------- checkes whether object is a Tripping Bracket
'                             or Bracket by plane
'
'  Inputs --------- Any object can be sent into the method
'
'  Output --------- True or False
'                   True --- if the object is Tripping Bracket or a Bracket by plane
'                   False --- if the object is neither Tripping Bracket nor Bracket by plane
'************************************************************************

Public Function IsBracket(oObject As Object) As Boolean

    On Error GoTo ErrorHandler
    Const sMETHOD = "IsBracket"

    IsBracket = False

    ' --------------------------------
    ' If a part, see if system-derived
    ' --------------------------------
    Dim oStructDetailHelper As StructDetailHelper
    Set oStructDetailHelper = New StructDetailHelper

    Dim oRootObject As Object ' either standalone or system
    If TypeOf oObject Is IJPlate Then
        oStructDetailHelper.IsPartDerivedFromSystem oObject, oRootObject, True

        If oRootObject Is Nothing Then
            Set oRootObject = oObject
        End If
    Else
        Exit Function
    End If

    ' ----------------------------------------------------------------------
    ' If root object is a bracket-type plate part, it is a standalone object
    ' ----------------------------------------------------------------------
    Dim oPlateUtil As IJPlateAttributes
    Set oPlateUtil = New PlateUtils

    Dim plateType As StructPlateType
    plateType = CollarPlate
    If TypeOf oRootObject Is IJPlate Then
        Dim oPlate As IJPlate
        Set oPlate = oRootObject
        plateType = oPlate.plateType
    End If

    If TypeOf oRootObject Is IJPlatePart And plateType = BracketPlate Then
        IsBracket = True
    ElseIf oPlateUtil.IsTrippingBracket(oRootObject) Then
        IsBracket = True
    ElseIf oPlateUtil.IsBracketByPlane(oRootObject) Then
        IsBracket = True
    End If

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number

End Function

