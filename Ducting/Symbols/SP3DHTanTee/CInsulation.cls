VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         VRK
'   Creation Date:  Wednesday, May 30 2007
'   Description:
'       Tangential Tee
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "CInsulation" 'Used for error messages
Private Const INCH = 0.0254

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double
        
    '{<(InputParamDec)>}
    Dim parHVACShape As Integer
    Dim parWidth As Double
    Dim parBWidth As Double
    Dim dHeaderLength As Double
    Dim parDepth As Double
    Dim parBDepth As Double
    Dim InsulationThickness As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parHVACShape = arrayOfInputs(2)
    parWidth = arrayOfInputs(3)
    parBWidth = arrayOfInputs(4)
    parDepth = arrayOfInputs(5)
    parBDepth = arrayOfInputs(6)
    InsulationThickness = arrayOfInputs(7)
        
    iOutput = 0

    If parHVACShape = Rectangular Then
        dHeaderLength = parBWidth + 2 * INCH
    ElseIf parHVACShape = 4 Then
        dHeaderLength = parBWidth + 2 * INCH
    End If
    
    If parHVACShape = Rectangular Then
        Dim ObjInsBody1 As Object
        Dim ostPoint   As New AutoMath.DPosition
        Dim oEnPoint   As New AutoMath.DPosition
        ostPoint.Set -(dHeaderLength / 2 + 2 * INCH), -(0.5 * parDepth + InsulationThickness), -(0.5 * parWidth + InsulationThickness)
        oEnPoint.Set (dHeaderLength / 2 + 2 * INCH), (0.5 * parDepth + InsulationThickness), (0.5 * parWidth + InsulationThickness)
        Set ObjInsBody1 = PlaceBox(m_OutputColl, ostPoint, oEnPoint)
        iOutput = iOutput + 1
        'Set the output
        m_OutputColl.AddOutput "ObjInsBody1", ObjInsBody1
        Set ObjInsBody1 = Nothing
        
    ElseIf parHVACShape = 4 Then
        ostPoint.Set (dHeaderLength / 2 + 2 * INCH), 0, 0
        oEnPoint.Set -(dHeaderLength / 2 + 2 * INCH), 0, 0
        Set ObjInsBody1 = PlaceCylinder(m_OutputColl, ostPoint, oEnPoint, parWidth + 2 * InsulationThickness, True)
        iOutput = iOutput + 1
        'Set the output
        m_OutputColl.AddOutput "ObjInsBody1", ObjInsBody1
        Set ObjInsBody1 = Nothing
    End If
    
    If parHVACShape = Rectangular Then
        Dim ObjInsBody2 As Object
        ostPoint.Set -(parBWidth / 2 + InsulationThickness), (parDepth / 2 + InsulationThickness), _
                                parWidth / 2
        oEnPoint.Set (parBWidth / 2 + InsulationThickness), (parDepth / 2 - parBDepth - InsulationThickness), _
                                parWidth / 2 + 3 * INCH
        Set ObjInsBody2 = PlaceBox(m_OutputColl, ostPoint, oEnPoint)
        iOutput = iOutput + 1
        'Set the output
        m_OutputColl.AddOutput "ObjInsBody2", ObjInsBody2
        Set ObjInsBody2 = Nothing

    ElseIf parHVACShape = 4 Then
        Dim YLength As Double, x As Double, y As Double
        y = parWidth / 2 - parBWidth
        x = Sqr((0.5 * parWidth) ^ 2 - (y * y))
        YLength = x
        ostPoint.Set 0, parWidth / 2 - parBWidth / 2, 0
        oEnPoint.Set 0, parWidth / 2 - parBWidth / 2, YLength + 3 * INCH
        Set ObjInsBody2 = PlaceCylinder(m_OutputColl, ostPoint, oEnPoint, parBWidth + 2 * InsulationThickness, True)
        iOutput = iOutput + 1
        'Set the output
        m_OutputColl.AddOutput "ObjInsBody2", ObjInsBody2
        Set ObjInsBody2 = Nothing
    End If
        Set ostPoint = Nothing
        Set oEnPoint = Nothing
        
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
    Err.HelpFile, Err.HelpContext
End Sub

