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
'   Author: KKC
'   Creation Date:  Friday, Oct 19 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     -----   ------------------
'   19.OCT.2007      KKC    CR-127640  Provide adapter symbol for male/female NPT to Kwik-Flange and Large-Flange'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Const NEGLIGIBLE_THICKNESS = 0.0001


Private Sub Class_Initialize()
    
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Long
    Dim ObjInsAdapterCyl As Object
    Dim ObjInsPort1 As Object
    Dim ObjInsPort2 As Object
    Dim parFacetoFace As Double
    Dim parLengthFlFace As Double
    Dim parInsulationThickness As Double
    Dim dInsulationDiameter As Double
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
   
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parLengthFlFace = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    iOutput = 0
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = NEGLIGIBLE_THICKNESS
    ' Insert your code for output 1(Cylinder 1)
    oStPoint.Set 0, 0, 0
    oEnPoint.Set parFacetoFace + parInsulationThickness, 0, 0
        
    Set ObjInsAdapterCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parLengthFlFace + 2 * parInsulationThickness, True)
        
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsAdapterCyl
    Set ObjInsAdapterCyl = Nothing
            
'   Insulation for the ports
'   Insulation for Port 1
    
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    oStPoint.Set 0, 0, 0
    oEnPoint.Set parFacetoFace - flangeThick - parInsulationThickness, 0, 0
    
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        dInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    Else
        dInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If
    Dim ObjInsulatedPort1 As Object
    Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort1
    Set ObjInsulatedPort1 = Nothing
    
' Place Insulation for Port 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If CmpDblEqual(flangeThick, 0) Then flangeThick = NEGLIGIBLE_THICKNESS
    oStPoint.Set 0, 0, 0
    oEnPoint.Set flangeThick + parInsulationThickness, 0, 0
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        dInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    Else
        dInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If
    Dim ObjInsulatedPort2 As Object
    Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort2
    Set ObjInsulatedPort2 = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
