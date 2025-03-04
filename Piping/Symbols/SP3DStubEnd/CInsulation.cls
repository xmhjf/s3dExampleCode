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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         RRK
'   Creation Date:  Wednusday, Mar 26 2008
'   Description:

'   Change History:
'   dd.mmm.yyyy    who     change description
'   -----------    -----        ------------------
'  26.Mar.2008     RRK      CR-83045:Created new symbol for stub end
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim flangeDiam     As Double
    Dim cptOffset       As Double
    Dim depth           As Double

    Dim iOutput     As Double
    
    Dim parStubLength As Double
    Dim parLapThickness As Double
    Dim parLapDiameter As Double
    Dim parLapRadius As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parStubLength = arrayOfInputs(2)
    parLapThickness = arrayOfInputs(3)
    parLapDiameter = arrayOfInputs(4)
    parLapRadius = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
    
    iOutput = 0

' Create Output 1 (Cylinder for Hub)
    Dim obInsjHub As Object
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition

    oStPoint.Set 0, 0, 0
    oEnPoint.Set parLapThickness, 0, 0
    
    Dim dInsulationDia As Double
    dInsulationDia = parLapDiameter + 2 * parInsulationThickness
    
    Set obInsjHub = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), obInsjHub
    Set obInsjHub = Nothing
    
' Create Output 2(Stub End Cylinder)
    
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
    
    Dim objInsStubEnd As Object
    
    oStPoint.Set parLapThickness, 0, 0
    oEnPoint.Set parStubLength, 0, 0
    
    dInsulationDia = pipeDiam + 2 * parInsulationThickness
    Set objInsStubEnd = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsStubEnd
    Set objInsStubEnd = Nothing
    
    'Creating the insulation for port 2 for the cases where flange thickness is non-zero
    If CmpDblGreaterthan(flangeThick, LINEAR_TOLERANCE) Then
        ' Insert your code for Insulated Port 2
        Dim ObjInsulatedPort2 As Object
        oStPoint.Set parStubLength, 0, 0
        oEnPoint.Set (parStubLength - flangeThick - parInsulationThickness), 0, 0
        
        'Check to see that header insulation doesn't exceed stub length
        If CmpDblGreaterthan(flangeThick + parInsulationThickness, parStubLength) Then
            oEnPoint.Set 0, 0, 0
        End If
        
        dInsulationDia = flangeDiam + 2 * parInsulationThickness
        Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDia, True)
    
        ' Set the output
        m_OutputColl.AddOutput "InsulatedPort2", ObjInsulatedPort2
        Set ObjInsulatedPort2 = Nothing
    End If
        
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
