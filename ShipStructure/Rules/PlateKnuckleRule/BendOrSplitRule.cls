VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BendOrSplitRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-------------------------------------------------------------------------------
'  Copyright (C) 2004, Intergraph Corporation.  All rights reserved.
'
'  FILE:  BendOrSplitRule.cls
'
'  DESCRIPTION:
'  Rule to set the manufacturing mrthod of a Plate Knuckle and if the
'  plate Knuckle is a 'bend' one, also set the 'bend' properties.
'
'  AUTHOR:  George Dreybus
'
'  HISTORY
'  12/10/2009    George Dreybus  Creation
'  1/2010 Implemented Y. Wang
'-------------------------------------------------------------------------------

Option Explicit
 Implements IJPlateKnuckleRule

Private m_eManufacturingMethod As enumKnuckleType

' Set the knuckle type to bend if the knuckle curve is linear
' otherwise, set to split
'
' Input:
'   oPlateObj:  input plate object
'   oKnuckle: input knuckle object

' Output:
'    enumKnuckleType:  eKT_Bent or eKT_Split or eKT_Ignore
'-------------------------------------------------------------------------------
Private Function ManufacturingMethod(oPlateObj As Object, oKnuckle As Object) As enumKnuckleType

    ManufacturingMethod = eKT_Split
    m_eManufacturingMethod = eKT_Split
    
    On Error GoTo ErrorHandler
        
    Dim oPlateKnuckle As IJKnuckle
    Set oPlateKnuckle = oKnuckle
    Dim oPlate As IJPlate
    Set oPlate = oPlateObj
    
    If Not (oPlate Is Nothing And oPlateKnuckle Is Nothing) Then
        If (oPlate.plateType = Hull) Then
'*** Customization
            'Users can specify the special rules for the HULL here.
'            If (oPlateKnuckle.IsKnuckleCurveLinear = False) Then  'bend
'               m_eManufacturingMethod = eKT_Ignore
'            End If
        Else
            If (oPlateKnuckle.IsKnuckleCurveLinear = True) Then
               m_eManufacturingMethod = eKT_Bent
            End If
        End If
    End If
    ManufacturingMethod = m_eManufacturingMethod
    Set oPlateKnuckle = Nothing
    Set oPlate = Nothing
 
 Exit Function
  
ErrorHandler:
  Err.Raise Err.Number
   
End Function

'-------------------------------------------------------------------------------
'Private Function InnerRadius
' Input:
'   oPlateObj:  input plate object
'   oKnuckle: input knuckle object
' Output:
'   dInnerRadius (double):  Should be a positive value for eKT_Bent and 0 for other types.
'-------------------------------------------------------------------
Private Function InnerRadius(oPlateObj As Object) As Double

    ' Decide of inner radius of a Plate Knuckle of manufacturing method eKT_Bent.
    ' If the plate knuckle manufacturing type is not eKT_Bent, set radius to 0.
    Dim dInnerRadius As Double
    InnerRadius = 0
    dInnerRadius = 0
    
    On Error GoTo ErrorHandler
        
    Dim dPlateThickness As Double
    Dim oPlate As IJPlate
    Set oPlate = oPlateObj

'*** Customization
'   Users can modify the radius based on the plate thickness (e.g  3 * plate thickness).
       
    If Not (oPlate Is Nothing) Then
        dPlateThickness = oPlate.thickness
        
        If dPlateThickness < 0.01 Then
            dInnerRadius = 0.05
        ElseIf dPlateThickness <= 0.02 Then
            dInnerRadius = 0.1
        ElseIf dPlateThickness < 0.03 Then
            dInnerRadius = 0.15
        Else
            dInnerRadius = 0.2
        End If
    End If
    Set oPlate = Nothing
    InnerRadius = dInnerRadius
    
 Exit Function
  
ErrorHandler:
   Err.Raise Err.Number
   
End Function

'-------------------------------------------------------------------------------
'Public Sub IJPlateKnuckleRule_ComputeKnuckleType
' Input:
'   oPlateObj:  input plate object
'   oKnuckle: input knuckle object

' Output:
'   enumKnuckleType:  eKT_Bent or eKT_Split or eKT_Ignore
'-------------------------------------------------------------------

Public Sub IJPlateKnuckleRule_ComputeKnuckleType(ByVal oPlate As Object, ByVal oPlateKnuckle As Object, _
                                        eManufacturingMethod As GSCADRefCurveOnSurface.enumKnuckleType)

On Error GoTo ErrorHandler
    eManufacturingMethod = ManufacturingMethod(oPlate, oPlateKnuckle)
Exit Sub
  
ErrorHandler:
    Err.Raise Err.Number
    
End Sub


'-------------------------------------------------------------------------------
'Public Sub IJPlateKnuckleRule_ComputeKnuckleRadius
' Input:
'   oPlateObj:  input plate object
' Output:
'   dInnerRadius (double):  Should be a positive value for eKT_Bent.
'-------------------------------------------------------------------

Public Sub IJPlateKnuckleRule_ComputeKnuckleRadius(ByVal oPlate As Object, _
                                         dInnerRadius As Double)
                                        
On Error GoTo ErrorHandler
    dInnerRadius = InnerRadius(oPlate)
Exit Sub
  
ErrorHandler:
    Err.Raise Err.Number
End Sub


