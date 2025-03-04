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
'   Copyright (c) 2004-06, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         svsmylav
'   Creation Date:  Thursday, Sep 23 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    Symbol Model No. is: S13AZ Page No. 346 of PDS Piping Component Data Reference Guide pcd.pdf.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'   16.Feb.2006     svsmylav       RI -28367: Updated the symbol to use Part Data Basis property
'                                          and added new input parameter 'FacetoFace' for Model 'HMMR', r629.pdf,
'                                          of Hansen, URL http://www.hantech.com.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

'''

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
    Dim ObjLeftCone As Object
    Dim ObjRightCone As Object
    Dim ObjNozzle1 As Object
    Dim ObjNozzle2 As Object

    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim parInsulationThickness As Double
    Dim parFacetoFace As Double

    Dim BaseCenPoint As AutoMath.DPosition
    Dim TopCenPoint As AutoMath.DPosition
    Set TopCenPoint = New AutoMath.DPosition
    Set BaseCenPoint = New AutoMath.DPosition
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Set oDir = New AutoMath.DVector
    Set oPlacePoint = New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoCenter = arrayOfInputs(2)
'    parFace1toCenter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
'    parFacetoFace = arrayOfInputs(5)
    
    iOutput = 0

 'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis

    If (lPartDataBasis <> 5) Then
        'Consider Face-to-center dimension basis
        parFacetoCenter = arrayOfInputs(2)
        parFace1toCenter = arrayOfInputs(3)
    
    ElseIf (lPartDataBasis = 5) Then
        'Face-to-face dimension basis
        parFacetoFace = arrayOfInputs(5)
        'Consider parFacetoCenter and parFace1toCenter to be half of parFacetoFace
        parFacetoCenter = parFacetoFace / 2
        parFace1toCenter = parFacetoFace / 2
    End If

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
 ' Insert your code for output 1(Left Cone)
    BaseCenPoint.Set -parFacetoCenter + flangeThick, 0#, 0#
    TopCenPoint.Set 0#, 0#, 0#
    Set ObjLeftCone = PlaceCone(m_OutputColl, BaseCenPoint, TopCenPoint, pipeDiam / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLeftCone
    Set ObjLeftCone = Nothing
    
 ' Insert your code for output 2(Right Cone)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    BaseCenPoint.Set parFace1toCenter - flangeThick, 0#, 0#
    TopCenPoint.Set 0#, 0#, 0#
    Set ObjRightCone = PlaceCone(m_OutputColl, BaseCenPoint, TopCenPoint, pipeDiam / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRightCone
    Set ObjRightCone = Nothing
' Insert your code for output 3(Nozzle 1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set -parFacetoCenter - sptOffset + depth, 0#, 0#
    oDir.Set -1, 0#, 0#
    Set ObjNozzle1 = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle1
    Set ObjNozzle1 = Nothing
    
 ' Insert your code for output 4(Nozzle 2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set parFace1toCenter + sptOffset - depth, 0#, 0#
    oDir.Set 1, 0#, 0#
    Set ObjNozzle2 = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle2
    Set ObjNozzle2 = Nothing
    Set TopCenPoint = Nothing
    Set BaseCenPoint = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
