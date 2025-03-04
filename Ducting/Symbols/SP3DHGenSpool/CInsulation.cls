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
'   Author:          RUK
'   Creation Date:  Monday, Aug 27 2007
'   Description:
'    CR-124211 Generic Spool.
'       Source: STD-M13.pdf
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     -----   ------------------
'   27-Aug-2007     RUK     Created
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = Insulation: 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Double
    
    Dim parHVACShape As Double
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parPPtoPort1 As Double
    Dim parPPtoPort2 As Double
    Dim parSpoolThickness As Double
    Dim parInsulationThickness As Double

'   Inputs
    Set oPartFclt = arrayOfInputs(1)
    parHVACShape = arrayOfInputs(2)
    parWidth = arrayOfInputs(3)
    parDepth = arrayOfInputs(4)
    parPPtoPort1 = arrayOfInputs(5)
    parPPtoPort2 = arrayOfInputs(6)
    parSpoolThickness = arrayOfInputs(7)    'Not used in the code
    parInsulationThickness = arrayOfInputs(8)

    'Origin is taken at the Center of the Damper
    'For Horizontally routed Duct X - Along the flow and Z - towards UP
    iOutput = 0
    
    If parHVACShape = 4 Then 'Round
        parDepth = parWidth
    End If
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    
    Set oGeomFactory = New GeometryFactory
    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition
    
    'Create the Damper (Output 1)
    Dim objInsSpool As Object
    If parHVACShape = Rectangular Then
        oStPoint.Set -parPPtoPort1, -(parDepth / 2 + parInsulationThickness), _
                        -(parWidth / 2 + parInsulationThickness)
        oEnPoint.Set parPPtoPort2, (parDepth / 2 + parInsulationThickness), _
                         (parWidth / 2 + parInsulationThickness)
        Set objInsSpool = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    ElseIf parHVACShape = 4 Then 'Round
        oStPoint.Set -parPPtoPort1, 0, 0
        oEnPoint.Set parPPtoPort2, 0, 0
        Set objInsSpool = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                parWidth + 2 * parInsulationThickness, True)
    End If
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsSpool
    Set objInsSpool = Nothing
    
    'Remove references
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oGeomFactory = Nothing

    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
