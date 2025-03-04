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
'    Symbol Model No. is: I12AZ Page No. 357 of PDS Piping Component Data Reference Guide pcd.pdf.
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
    Dim ObjUpperCone As Object
    Dim ObjLowerCone As Object
    Dim ObjRightCone As Object
    Dim ObjTopCone As Object
    Dim ObjBottomCone As Object
        
    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)     'P2
    parFace1toCenter = arrayOfInputs(3)     'P3
    parFace2toCenter = arrayOfInputs(4)     'P4
    parFace3toCenter = arrayOfInputs(5)     'P5
    parInsulationThickness = arrayOfInputs(6)
    
    iOutput = 0

 ' Insert your code for output 1(Left Cone)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
'   Compute cone height and diameter from pipeDiam
 '  Following logic is taken from PDS EDEN pipeDiam
    Dim factor As Double
    If (pipeDiam <= 2 * InchToMtr) Then
            factor = 0.4
    Else
            If (pipeDiam > 2 * InchToMtr And pipeDiam <= 10 * InchToMtr) Then
                    factor = 0.45
            Else
                    If (pipeDiam > 10 * InchToMtr) Then
                            factor = 0.5
                    End If
            End If
    End If

    Dim half_plug_length As Double
    Dim cone_diameter As Double
    
    half_plug_length = pipeDiam * 0.5
    cone_diameter = parFacetoCenter * factor
'
    Dim BaseCenPoint As New AutoMath.DPosition
    Dim TopCenPoint As New AutoMath.DPosition
    
    BaseCenPoint.Set -parFacetoCenter + flangeThick, 0#, 0#

    TopCenPoint.Set 0#, 0#, 0#
    
    Set ObjLeftCone = PlaceCone(m_OutputColl, BaseCenPoint, TopCenPoint, pipeDiam / 2, 0, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLeftCone
    Set ObjLeftCone = Nothing
    Set BaseCenPoint = Nothing
    Set TopCenPoint = Nothing
    
 ' Insert your code for output 2(Upper Cone)
    Dim UpperConeBasePt As New AutoMath.DPosition
    Dim UpperConeTopPt As New AutoMath.DPosition
    Dim pipeDiam2 As Double
    
    cone_diameter = parFacetoCenter * 0.5
    
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam, sptOffset, depth
    If pipeDiam2 > pipeDiam Then
        half_plug_length = pipeDiam2 * 0.5
    Else
        half_plug_length = pipeDiam * 0.5
    End If
    UpperConeBasePt.Set 0#, 0#, 0
    UpperConeTopPt.Set 0#, half_plug_length, 0#
    
    Set ObjUpperCone = PlaceCone(m_OutputColl, UpperConeBasePt, UpperConeTopPt, cone_diameter / 2, 0, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjUpperCone
    Set ObjUpperCone = Nothing
    Set UpperConeBasePt = Nothing
    Set UpperConeTopPt = Nothing
    
 ' Insert your code for output 3(Lower Cone)
    Dim LowerConeBasePt As New AutoMath.DPosition
    Dim LowerConeTopPt As New AutoMath.DPosition
    
    LowerConeBasePt.Set 0#, 0#, 0#
    LowerConeTopPt.Set 0#, -half_plug_length, 0#
    
    Set ObjLowerCone = PlaceCone(m_OutputColl, LowerConeBasePt, LowerConeTopPt, cone_diameter / 2, 0, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLowerCone
    Set ObjLowerCone = Nothing
    Set LowerConeBasePt = Nothing
    Set LowerConeTopPt = Nothing
    
 ' Insert your code for output 4(Right Cone)
    Set oPartFclt = arrayOfInputs(1)
    
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim RightConeBasePt As New AutoMath.DPosition
    Dim RightConeTopPt  As New AutoMath.DPosition
    
    RightConeBasePt.Set parFace1toCenter - flangeThick, 0#, 0#
    RightConeTopPt.Set 0#, 0#, 0#
    
    Set ObjRightCone = PlaceCone(m_OutputColl, RightConeBasePt, RightConeTopPt, pipeDiam / 2, 0, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRightCone
    Set ObjRightCone = Nothing
    Set RightConeBasePt = Nothing
    Set RightConeTopPt = Nothing
    
' Insert your code for output 5(Top Cone)
    Set oPartFclt = arrayOfInputs(1)
    
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim TopConeBasePt As New AutoMath.DPosition
    Dim TopConeTopPt  As New AutoMath.DPosition
    
    TopConeBasePt.Set 0, parFace2toCenter - flangeThick, 0#
TopConeTopPt.Set 0#, 0#, 0#
    
    Set ObjTopCone = PlaceCone(m_OutputColl, TopConeBasePt, TopConeTopPt, pipeDiam / 2, 0, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopCone
    Set ObjTopCone = Nothing
    Set TopConeBasePt = Nothing
    Set TopConeTopPt = Nothing
    
' Insert your code for output 6(Bottom Cone)
    Set oPartFclt = arrayOfInputs(1)
    
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim BottomConeBasePt As New AutoMath.DPosition
    Dim BottomConeTopPt  As New AutoMath.DPosition
    
    BottomConeBasePt.Set 0, -parFace3toCenter + flangeThick, 0#
    BottomConeTopPt.Set 0#, 0#, 0#
    
    Set ObjBottomCone = PlaceCone(m_OutputColl, BottomConeBasePt, BottomConeTopPt, pipeDiam / 2, 0, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBottomCone
    Set ObjBottomCone = Nothing
    Set BottomConeBasePt = Nothing
    Set BottomConeTopPt = Nothing
    
' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -parFacetoCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    
    Set oPartFclt = arrayOfInputs(1)
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
      
    oPlacePoint.Set parFace1toCenter + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 3
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
     
    oPlacePoint.Set 0, parFace2toCenter + sptOffset - depth, 0
    oDir.Set 0, 1, 0
    
    Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
            
' Place Nozzle 4
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
     
    oPlacePoint.Set 0, -parFace3toCenter - sptOffset + depth, 0
    oDir.Set 0, -1, 0
    
    Set objNozzle = CreateNozzle(4, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

