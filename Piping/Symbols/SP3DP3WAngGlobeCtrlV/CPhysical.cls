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
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          svsmylav
'   Creation Date:  Thursday, Sep 23 2004
'   Description:
'       This class module is the place for user to implement graphical part of VBSymbol for this aspect
'       Symbol Model No. is: I6AZ Page No. 355 of PDS Piping Component Data Reference Guide pcd.pdf.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was 
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private InchToMtr       As Double

Private Sub Class_Initialize()
    InchToMtr = 25.4 / 1000
End Sub
Private Function ReturnMax3(A#, B#, C#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If MaxValue < B Then MaxValue = B
    If MaxValue < C Then MaxValue = C
    ReturnMax3 = MaxValue
End Function

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim sptOffset2 As Double
    Dim depth2 As Double
    Dim sptOffset3 As Double
    Dim depth3 As Double

    Dim iOutput     As Double
    Dim ObjValveBody As Object
    Dim ObjBodyCone1 As Object
    Dim ObjBodyCone2 As Object
    Dim ObjBodyCone3 As Object
    
    Dim parFacetoCenter As Double
    Dim Face1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInsulationThickness As Double

    Dim axis            As New AutoMath.DVector
    Dim centPoint       As New AutoMath.DPosition
    Dim stPoint         As New AutoMath.DPosition
    Dim enPoint         As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)
    Face1toCenter = arrayOfInputs(3)
    parFace2toCenter = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    
    iOutput = 0

' Insert your code for output 3(ValveBody)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
 '  Following logic is taken from PDS EDEN
    Dim factor As Double
    If (pipeDiam <= 2 * InchToMtr) Then
            factor = 0.5
    Else
            If (pipeDiam > 2 * InchToMtr And pipeDiam <= 10 * InchToMtr) Then
                    factor = 0.55
            Else
                    If (pipeDiam > 10 * InchToMtr) Then
                            factor = 0.6
                    End If
            End If
    End If

    Dim sphere_radius As Double
    sphere_radius = parFacetoCenter * factor * 0.5

    Dim StartPos As New AutoMath.DPosition
    StartPos.Set 0, 0, 0
    Set ObjValveBody = PlaceSphere(m_OutputColl, StartPos, sphere_radius)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBody
    Set ObjValveBody = Nothing
    Set centPoint = Nothing
    Set axis = Nothing
    
' Insert your code for output 4(Cone - Port1 Side)
    stPoint.Set -parFacetoCenter + flangeThick, 0, 0
'   enPoint.Set 0, 0,0
'   PlaceCone is giving 'Invalid property value' failure. A very small increment in Y or Z coordinate
'   value removes the 'Invalid property value' failure.
    enPoint.Set 0, 0, 0
    Set ObjBodyCone1 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1
    Set ObjBodyCone1 = Nothing
    
' Insert your code for output 5(Cone - Port2 Side)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset2, depth2
    stPoint.Set Face1toCenter - flangeThick, 0, 0
'   enPoint.Set 0, 0, 0
'   PlaceCone is giving 'Invalid property value' failure. A very small increment in Y or Z coordinate
'   value removes the 'Invalid property value' failure.
    enPoint.Set 0, 0, 0
    Set ObjBodyCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
    Set ObjBodyCone2 = Nothing
    
' Insert your code for output 6(Cone - Port3 Side)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset3, depth3
    stPoint.Set 0, parFace2toCenter - flangeThick, 0
    '   PlaceCone is giving 'Invalid property value' failure. A very small increment in Y or Z coordinate
'   value removes the 'Invalid property value' failure.
    enPoint.Set 0, 0, 0
    Set ObjBodyCone3 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone3
    Set ObjBodyCone3 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -parFacetoCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    oPlacePoint.Set Face1toCenter + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 3
    oPlacePoint.Set 0, parFace2toCenter + sptOffset3 - depth3, 0
    oDir.Set 0, 1, 0
    Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set StartPos = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
