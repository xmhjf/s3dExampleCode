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
'   Author:         RRK
'   Creation Date:  Friday, 13, Jul 2007
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    The symbol is modelled for RM-2 Type Tank Vent Check Valve of Winel
'    (source:http://www.winel.nl/Products/tvcv.htm)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ----   ---------------------------------------
'   13.Jul.2007      RRK   CR-123055: Created the symbol.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim iOutput     As Double
    
    Dim parHeight As Double
    Dim parDiameter As Double
    Dim parLength As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parHeight = arrayOfInputs(2)
    parDiameter = arrayOfInputs(3)
    parLength = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    
    iOutput = 0
'Assumption : Origin is at Port 1.

' Insert your code for Insulation of Cylindrical portion of the body (output 5)
    'Assumption : The Length of the cylinder is 80% of the Length.
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    
    Set oStPoint = New AutoMath.DPosition
    oStPoint.Set parHeight - (parDiameter / 2), 0.5 * parLength + parInsulationThickness, 0
    Set oEnPoint = New AutoMath.DPosition
    oEnPoint.Set parHeight - (parDiameter / 2), -0.5 * parLength - parInsulationThickness, 0
    
    Dim objInsCylinderBody As Object
    Set objInsCylinderBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parDiameter + (2 * parInsulationThickness), True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCylinderBody
    Set objInsCylinderBody = Nothing
    
' Insert your code for Insulated cylinder 2 (output 6)
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    oStPoint.Set 0, 0, 0
    oEnPoint.Set parHeight - (parDiameter / 2), 0, 0
    
    Dim objInsCylinderBody2 As Object
    Set objInsCylinderBody2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam + (2 * parInsulationThickness), True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCylinderBody2
    Set objInsCylinderBody2 = Nothing
    
' Insert your code for insulated port 1 (output 7)
    
    Dim dFlangeThickness As Double
    
    If CmpDblLessThanOrEqualTo(flangeThick, 0) Then
        dFlangeThickness = 0.0001
    Else
        dFlangeThickness = flangeThick
    End If
    
    oStPoint.Set 0, 0, 0
    oEnPoint.Set dFlangeThickness + parInsulationThickness, 0, 0
    
    Dim dInsulationDia As Double
    
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        dInsulationDia = flangeDiam + (2 * parInsulationThickness)
    Else
        dInsulationDia = pipeDiam + (2 * parInsulationThickness)
    End If
    
    Dim objInsPort1 As Object
    Set objInsPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsPort1
    Set objInsPort1 = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
