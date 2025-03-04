VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSquareEdgeTrim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-------------------------------------------------------------------------------
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  FILE:  SquareEdgeTrimInit.cls
'
'  DESCRIPTION:
'  Rule creating the square edge trim for the plate bounded by stiffener or ER
'  This will return the two offset values for the profile and plate trim.
'  This rule can be updated for any additional cases if needed byt he customers as well.
'
'  AUTHOR:  Neelima Bhupatiraju
'
'  HISTORY
'  03-Nov-2011    Neelima Bhupatiraju   Creation For CR#204874
'  17-Apr-2014    Neelima Bhupatiraju   Modified For DM#253641. Square edge functionality was not
'                                       working correctly for split plates cases where the profile
'                                       and plate thickness are not same.
'-------------------------------------------------------------------------------

Option Explicit
Implements IJSquareEdgeTrimRule
'-------------------------------------------------------------------------------
' Instance data
'-------------------------------------------------------------------------------
Private METHOD As String 'String used for error message
Private Const MODULE = "M:\SharedContent\Src\ShipStructure\Rules\SquareEdgeTrimRules\CSquareEdgeTrim:"  'Used for error messages
Private Const TOLERANCE = 0.000001

'-------------------------------------------------------------------------------
' Interface IJSquareEdgeTrimRule
' Method:   ComputeOffsetValuesForProfileOnPlate
' Description: This method will takes in the two thickness values and an angle
' one is plate thickness and other is the profile web thickness.
' Then it returns the offset values calculated based on the input user rules.
'
'-------------------------------------------------------------------------------
Private Sub IJSquareEdgeTrimRule_ComputeOffsetValuesForProfileOnPlate( _
                                                                        ByVal dPlateThickness As Double, _
                                                                        ByVal dProfileThickness As Double, _
                                                                        ByVal dBetaAng As Double, _
                                                                        ByRef dPlateOffset As Double, _
                                                                        ByRef dProfileOffset As Double)

    Const METHOD = "IJSquareEdgeTrimRule_ComputeOffsetValuesForProfileOnPlate"
    On Error GoTo ErrorHandler
   
    dPlateOffset = 0#
    dProfileOffset = 0#
    
    ' Modified the below condition such that angle is round off to 7 digits after
    ' decimal and checked with the tolerance value. So that it will work for all
    ' the case as required for MHE industry. For DM#253641
    ' check if the plate thickness is equal to the web thickness of the stiffener
    ' and beta angle is bigger than 90 degree, then we calc the offsets as below.
    If (Round(dBetaAng - (PI / 2), 7) > TOLERANCE) Then
        dProfileOffset = dPlateThickness * ((1 - Cos(PI - dBetaAng)) / Sin(PI - dBetaAng))
        dPlateOffset = dProfileOffset
     
    ' Check if beta angle is almost equal to 90 degree, the plate thickness is not equal to the web thickness of the stiffener
    ' and  then we calc the offsets as below.
    ElseIf (Round(Abs(dBetaAng - (PI / 2)), 7) <= TOLERANCE) Then
        dProfileOffset = dPlateThickness * ((1 - Cos(PI - dBetaAng)) / Sin(PI - dBetaAng))
        If (dPlateThickness = dProfileThickness) Then
            dPlateOffset = dPlateThickness
        ElseIf (dPlateThickness <> dProfileThickness) Then
            dPlateOffset = dProfileThickness
        End If
    ' check if beta angle is less than or equal to 90 degree and greater than or equal to 60 degree
    ' then we calc the offsets values as below for profile and plate.
    ElseIf ((Round(dBetaAng - (PI / 3), 7) >= TOLERANCE) And (Round(dBetaAng - (PI / 2), 7) <= TOLERANCE)) Then
        dPlateOffset = (dPlateThickness * ((1 + Cos(dBetaAng)) / Sin(dBetaAng)))
        dProfileOffset = (dPlateThickness + (dProfileThickness * Cos(dBetaAng))) / Sin(dBetaAng)

    ' check if beta angle is less than 60 degree
    ' then we calc the offsets values as below for profile and plate.
    ElseIf (Round(dBetaAng - (PI / 3), 7) < TOLERANCE) Then
        dPlateOffset = (dPlateThickness * (Cos(dBetaAng) / Sin(dBetaAng)))
        dProfileOffset = (dPlateThickness + (dProfileThickness * Cos(dBetaAng))) / Sin(dBetaAng)

    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number

End Sub

