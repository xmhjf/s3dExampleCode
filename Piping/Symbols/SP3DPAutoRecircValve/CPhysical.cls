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
'   Author:     svsmylav
'   Creation Date:  Thursday, Sep 23 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    Symbol Model No. is: I30AZ Page No. 359 of PDS Piping Component Data Reference Guide pcd.pdf.
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
    Dim ObjMainCone As Object
    Dim ObjUpperCone As Object
    Dim ObjRightCylinder As Object
    Dim ObjNozzle1 As Object
    Dim ObjNozzle2 As Object
    Dim ObjNozzle3 As Object
    
    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInsulationThickness As Double
    
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    Dim CylinderLen As Double
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)
    parFace1toCenter = arrayOfInputs(3)
    parFace2toCenter = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    
    iOutput = 0

 ' Insert your code for output 1(Main Cone)
    Dim BaseCenPoint As New AutoMath.DPosition
    Dim TopCenPoint As New AutoMath.DPosition
    Dim grPipeDia As Double
    Dim MiniCylinTk As Double
    
    MiniCylinTk = 0.01
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    BaseCenPoint.Set -parFacetoCenter + flangeThick, 0#, 0#
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset2, depth2
    If pipeDiam > pipeDiam2 Then                        ' Finding the greater of twi pipe Diameters
        grPipeDia = pipeDiam
    Else
        grPipeDia = pipeDiam2
    End If
    If (0.2 * grPipeDia) > MiniCylinTk Then         ' Determining the minimum thickness of Cylinder(Output 3)
        CylinderLen = MiniCylinTk
    Else
        CylinderLen = 0.2 * grPipeDia
    End If
    If flangeThick2 > 0 Then
        TopCenPoint.Set parFace1toCenter - flangeThick2, 0#, 0
    Else
        TopCenPoint.Set parFace1toCenter - CylinderLen, 0#, 0
    End If
    Set ObjMainCone = PlaceCone(m_OutputColl, BaseCenPoint, TopCenPoint, pipeDiam / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMainCone
    Set ObjMainCone = Nothing
    
 ' Insert your code for output 2(Upper Cone)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim UpperconeBasePt As New AutoMath.DPosition
    Dim UpperconeTopPt As New AutoMath.DPosition
    UpperconeBasePt.Set 0, parFace2toCenter - flangeThick, 0#
    UpperconeTopPt.Set 0#, 0#, 0#
    Set ObjUpperCone = PlaceCone(m_OutputColl, UpperconeBasePt, UpperconeTopPt, pipeDiam / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjUpperCone
    Set ObjUpperCone = Nothing
    
 ' Insert your code for output 3(Right Cylinder)
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    stPoint.Set parFace1toCenter - CylinderLen, 0, 0
    enPoint.Set parFace1toCenter, 0, 0
    Set ObjRightCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRightCylinder
    Set ObjRightCylinder = Nothing
    
 ' Insert your code for output 6(Nozzle 1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim faceToFace  As Double

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -parFacetoCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set oPartFclt = arrayOfInputs(1)
    Set ObjNozzle1 = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle1
    Set ObjNozzle1 = Nothing
    
 ' Insert your code for output 7(Nozzle 2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set parFace1toCenter + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    Set ObjNozzle2 = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle2
    Set ObjNozzle2 = Nothing
    
 ' Insert your code for output 8(Nozzle 3)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set 0, parFace2toCenter + sptOffset - depth, 0
    oDir.Set 0, 1, 0
    Set ObjNozzle3 = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle3
    Set ObjNozzle3 = Nothing
    Set BaseCenPoint = Nothing
    Set TopCenPoint = Nothing
    Set UpperconeBasePt = Nothing
    Set UpperconeTopPt = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
