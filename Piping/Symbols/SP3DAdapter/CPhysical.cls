VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author: KKC
'   Creation Date:  Friday, Oct 19 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     -----   ------------------
'   19.OCT.2007      KKC    CR-127640  Provide adapter symbol for male/female NPT to Kwik-Flange and Large-Flange
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages


Private Sub Class_Initialize()
 
''''
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
    
    Dim iOutput     As Double
    Dim ObjAdapterCyl As Object
    Dim ObjNozzle1 As Object
    Dim ObjNozzle2 As Object
    Dim parFacetoFace As Double
    Dim parLengthFlFace As Double
    Dim parInsulationThickness As Double
    
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
    ' Insert your code for output 1(Cylinder 1)
    oStPoint.Set 0, 0, 0
    oEnPoint.Set parFacetoFace, 0, 0
    Set ObjAdapterCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parLengthFlFace, True)
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjAdapterCyl
    Set ObjAdapterCyl = Nothing
        
' Insert your code for output 3(Nozzle)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As New AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set ObjNozzle1 = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle1
    Set ObjNozzle1 = Nothing
 
 ' Insert your code for output 4(Nozzle)
     RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set parFacetoFace + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    Set ObjNozzle2 = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle2
    Set ObjNozzle2 = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
