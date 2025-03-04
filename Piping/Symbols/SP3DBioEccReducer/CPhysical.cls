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
'   Copyright (c) 2003-2004, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          svsmylav
'   Creation Date:  Tuesday, Mar 26 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol is created using six outputs. Out of these three are Insulation aspect outputs.
' This class contains three outputs: Body prepared by using 'CreateByCurves' function
' and the other two outputs are prepared by using CreateNozzle/CreateNozzleWithLength
' function which is decided based on the end preparation.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   26.Jul.2004     ACM                   DI-61828 changed declaration of pipeport from IJDPipePort to IJCatalogPipePort
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const CLAMPED_ENDS = 1
Private Const WELDED_ENDS = 2
Private Const CLAMPED_WELDED = 3
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
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim flangeDiam2      As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    
    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double
    
    Dim Port1toTaperBegin   As Double
    Dim TaperLength         As Double
    Dim TaperEndToPort2     As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    
    iOutput = 0

' Insert your code for output 1(Body)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset2, depth2

'   Get end-preparation
    Dim nozzleColl As IJDCollection
    'Dim pipeport As IJDPipePort
     Dim pipeport As IJCatalogPipePort
     
    Dim EndPrep1 As Long
    Dim EndPrep2 As Long
    Dim setOfEndPrep    As Integer
    Set nozzleColl = oPartFclt.GetNozzles
    Set pipeport = nozzleColl.Item(1)
    EndPrep1 = pipeport.EndPreparation
    Set pipeport = nozzleColl.Item(2)
    EndPrep2 = pipeport.EndPreparation
    Set nozzleColl = Nothing
    Set pipeport = Nothing
    
'   Determine the set of end-preparation: Clamp-Clamp, Weld-Weld, Clamp-Weld
'   Length of the taper measured along x-axis is taken as equal to parFacetoFace / 3 for
'   Weld-Weld, Clamp-Weld cases
    If EndPrep1 = 920 Then
        If EndPrep2 = 920 Then
            setOfEndPrep = CLAMPED_ENDS
            Port1toTaperBegin = flangeThick
            TaperEndToPort2 = flangeThick2
        ElseIf EndPrep2 = 301 Or EndPrep2 = 391 Then
            setOfEndPrep = CLAMPED_WELDED
            If flangeThick > parFacetoFace / 6 Then
                Port1toTaperBegin = flangeThick
            Else
                Port1toTaperBegin = parFacetoFace / 6
            End If
            TaperEndToPort2 = parFacetoFace - Port1toTaperBegin - parFacetoFace / 3
        End If
    ElseIf EndPrep1 = 301 Or EndPrep1 = 391 And EndPrep2 = 301 Or EndPrep2 = 391 Then
        setOfEndPrep = WELDED_ENDS
        Port1toTaperBegin = parFacetoFace / 3
        TaperEndToPort2 = parFacetoFace - Port1toTaperBegin - parFacetoFace / 3
    Else
        GoTo ErrorLabel
    End If
    
    Dim cirNormalX As Double, cirNormalY As Double, cirNormalZ As Double
    cirNormalX = 1
    cirNormalY = 0
    cirNormalZ = 0
    
    Dim cir1CenterX As Double, cir1CenterY As Double, cir1CenterZ As Double
    cir1CenterX = -parFacetoFace / 2 + Port1toTaperBegin
    cir1CenterY = (pipeDiam - pipeDiam2) / 4
    cir1CenterZ = 0

    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim objCircle1 As IngrGeom3D.Circle3d
    Set objCircle1 = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                    cir1CenterX, cir1CenterY, cir1CenterZ, _
                                    cirNormalX, cirNormalY, cirNormalZ, pipeDiam * 0.5)
                                
'   Create Circle2
    Dim cir2CenterX As Double, cir2CenterY As Double, cir2CenterZ As Double
    cir2CenterX = parFacetoFace / 2 - TaperEndToPort2
    cir2CenterY = -(pipeDiam - pipeDiam2) / 4
    cir2CenterZ = 0

    Dim objCircle2 As IngrGeom3D.Circle3d
    Set objCircle2 = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                    cir2CenterX, cir2CenterY, cir2CenterZ, _
                                    cirNormalX, cirNormalY, cirNormalZ, pipeDiam2 * 0.5)
'   Create Ruled Surface for body
    Dim ObjBody As IngrGeom3D.RuledSurface3d
    Set ObjBody = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                            objCircle1, objCircle2, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
    Set geomFactory = Nothing
    
'   Remove curve
    Dim sp3dElem1 As IJDObject
    Dim sp3dElem2 As IJDObject
    Set sp3dElem1 = objCircle1
    Set sp3dElem2 = objCircle2
    Set objCircle1 = Nothing
    Set objCircle2 = Nothing
    sp3dElem1.Remove
    sp3dElem2.Remove
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
'    Dim faceToFace  As Double

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
'   faceToFace = arrayOfInputs(2)
    oPlacePoint.Set -parFacetoFace / 2 - sptOffset + depth, (pipeDiam - pipeDiam2) / 4, 0
    oDir.Set -1, 0, 0
'   Set oPartFclt = arrayOfInputs(1)
    If setOfEndPrep = CLAMPED_ENDS Or flangeThick > parFacetoFace / 6 Then
        Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    Else
        Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, Port1toTaperBegin)
    End If
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

' Place Nozzle 2
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set parFacetoFace / 2 + sptOffset2 - depth2, -(pipeDiam - pipeDiam2) / 4, 0
    oDir.Set 1, 0, 0
    If setOfEndPrep = CLAMPED_ENDS Then
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    Else
        Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, TaperEndToPort2)
    End If
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub



