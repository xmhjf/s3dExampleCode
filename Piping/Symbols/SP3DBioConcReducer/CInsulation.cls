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
'   Copyright (c) 2003-2004, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          svsmylav
'   Creation Date:  Tuesday, Mar 26 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol is created using six outputs. Out of these three are Physical aspect outputs.
' This class contains three outputs: InsulatedBody prepared by using 'CreateByCurves' function
' and the other two outputs are prepared by using PlaceRevolution function (curve to be revolved
' is decided based on the end preparation).
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   26.Jul.2004     ACM                   DI-61828 changed declaration of pipeport from IJDPipePort to IJCatalogPipePort
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private Const CLAMPED_ENDS = 1
Private Const WELDED_ENDS = 2
Private Const CLAMPED_WELDED = 3
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
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
    Dim ObjInsulatedBody As Object
    Dim ObjInsBodyPort1Side As Object
    Dim ObjInsBodyPort2Side As Object
    
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

' Insert your code for output 2(Insulated Body - Taper)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset2, depth2

    If depth <= 0 Then
        flangeDiam = pipeDiam
        flangeDiam2 = pipeDiam2
    End If

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
    End If
    
    Dim cirNormalX As Double, cirNormalY As Double, cirNormalZ As Double
    cirNormalX = 1
    cirNormalY = 0
    cirNormalZ = 0
    
    Dim cir1CenterX As Double, cir1CenterY As Double, cir1CenterZ As Double
    cir1CenterX = -parFacetoFace / 2 + Port1toTaperBegin
    cir1CenterY = 0
    cir1CenterZ = 0
    
    Dim parInsulationDiameter As Double
    parInsulationDiameter = pipeDiam + 2 * parInsulationThickness

    Dim geomFactory As New IngrGeom3D.GeometryFactory
    Dim objCircleI1 As IngrGeom3D.Circle3d
    Set objCircleI1 = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                    cir1CenterX, cir1CenterY, cir1CenterZ, _
                                    cirNormalX, cirNormalY, cirNormalZ, _
                                    parInsulationDiameter * 0.5)

'   Create Circle2
    Dim cir2CenterX As Double, cir2CenterY As Double, cir2CenterZ As Double
    cir2CenterX = parFacetoFace / 2 - TaperEndToPort2
    cir2CenterY = 0
    cir2CenterZ = 0

    Dim parInsulationDiameter2 As Double
    parInsulationDiameter2 = pipeDiam2 + 2 * parInsulationThickness
    Dim objCircleI2 As IngrGeom3D.Circle3d
    Set objCircleI2 = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                    cir2CenterX, cir2CenterY, cir2CenterZ, _
                                    cirNormalX, cirNormalY, cirNormalZ, _
                                    parInsulationDiameter2 * 0.5)
'   Create Ruled Surface for body
    Set ObjInsulatedBody = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                            objCircleI1, objCircleI2, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    
'   Remove curve
    Dim sp3dElem1 As IJDObject
    Dim sp3dElem2 As IJDObject
    Set sp3dElem1 = objCircleI1
    Set sp3dElem2 = objCircleI2
    Set objCircleI1 = Nothing
    Set objCircleI2 = Nothing
    sp3dElem1.Remove
    sp3dElem2.Remove
    
' Insert your code for output 3(Insulated Body Port1 Side)
'   Determine number of points in the linestring which is used for revolution
    Dim numPoints As Integer
    If setOfEndPrep = CLAMPED_ENDS Or setOfEndPrep = WELDED_ENDS Then
        numPoints = 4
    ElseIf flangeThick > parFacetoFace / 6 Then
        numPoints = 4   'Insulation for flange will do.
    Else
        numPoints = 6   'Insulation for flange and straight piece towards Port1.
    End If
    ReDim Port1SideIns(0 To 3 * numPoints - 1) As Double

'   Start point (On the axis, Taper starts here)
    Port1SideIns(0) = -parFacetoFace / 2 + Port1toTaperBegin
    Port1SideIns(1) = 0
    Port1SideIns(2) = 0

'   End Point (On the axis, at the Port1 center)
    Port1SideIns(3 * numPoints - 3) = -parFacetoFace / 2
    Port1SideIns(3 * numPoints - 2) = 0
    Port1SideIns(3 * numPoints - 1) = 0
    
'   Penultimate point (At Port1)
    Port1SideIns(3 * numPoints - 6) = -parFacetoFace / 2
    Port1SideIns(3 * numPoints - 5) = flangeDiam / 2 + parInsulationThickness
    Port1SideIns(3 * numPoints - 4) = 0

    If numPoints = 4 Then
'       Point before the Penultimate
        Port1SideIns(3 * numPoints - 9) = -parFacetoFace / 2 + Port1toTaperBegin
        Port1SideIns(3 * numPoints - 8) = flangeDiam / 2 + parInsulationThickness
        Port1SideIns(3 * numPoints - 7) = 0
    Else
'       Point before the Penultimate
        Port1SideIns(3 * numPoints - 9) = -parFacetoFace / 2 + flangeThick + parInsulationThickness
        Port1SideIns(3 * numPoints - 8) = flangeDiam / 2 + parInsulationThickness
        Port1SideIns(3 * numPoints - 7) = 0
        
        Port1SideIns(6) = -parFacetoFace / 2 + flangeThick + parInsulationThickness
        Port1SideIns(7) = pipeDiam / 2 + parInsulationThickness
        Port1SideIns(8) = 0

'       Points on the pipe which is towards Port1
        Port1SideIns(3) = -parFacetoFace / 2 + Port1toTaperBegin
        Port1SideIns(4) = pipeDiam / 2 + parInsulationThickness
        Port1SideIns(5) = 0
    End If

    Dim oLineString As IngrGeom3D.LineString3d
    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, numPoints, Port1SideIns)

'   Create revolution of the shape line string
    Dim axis        As New AutoMath.DVector
    axis.Set -1, 0, 0
    Dim CP As New AutoMath.DPosition 'arc center point
    CP.Set 0, 0, 0
    Set ObjInsBodyPort1Side = PlaceRevolution(m_OutputColl, oLineString, axis, CP, 2 * PI, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBodyPort1Side
    Set ObjInsBodyPort1Side = Nothing
    Set axis = Nothing
    Set CP = Nothing
    
'   Remove line string
    Dim objLineString As IJDObject
    Set objLineString = oLineString
    objLineString.Remove
    Set objLineString = Nothing
    
' Insert your code for output 4(Insulated Body Port2 Side)
'   The number of points in the linestring which is used for revolution are four for all cases
    numPoints = 4
    ReDim Port2SideIns(0 To 3 * numPoints - 1) As Double

'   Start point (On the axis, Taper ends here)
    Port2SideIns(0) = parFacetoFace / 2 - TaperEndToPort2
    Port2SideIns(1) = 0
    Port2SideIns(2) = 0
    
'   Point before the Penultimate
    Port2SideIns(3) = parFacetoFace / 2 - TaperEndToPort2
    Port2SideIns(4) = flangeDiam2 / 2 + parInsulationThickness
    Port2SideIns(5) = 0
    
'   Penultimate point (At Port2)
    Port2SideIns(6) = parFacetoFace / 2
    Port2SideIns(7) = flangeDiam2 / 2 + parInsulationThickness
    Port2SideIns(8) = 0
    
'   End Point (On the axis, at the Port2 center)
    Port2SideIns(9) = parFacetoFace / 2
    Port2SideIns(10) = 0
    Port2SideIns(11) = 0

    Dim oLineString2 As IngrGeom3D.LineString3d
    Set oLineString2 = geomFactory.LineStrings3d.CreateByPoints(Nothing, numPoints, Port2SideIns)
    
    Set geomFactory = Nothing
'   Create revolution of the shape line string
    Dim axis2        As New AutoMath.DVector
    axis2.Set 1, 0, 0
    Dim CP2 As New AutoMath.DPosition 'arc center point
    CP2.Set 0, 0, 0
    Set ObjInsBodyPort2Side = PlaceRevolution(m_OutputColl, oLineString2, axis2, CP2, 2 * PI, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBodyPort2Side
    Set ObjInsBodyPort2Side = Nothing
    
'   Remove line string
    Dim objLineString2 As IJDObject
    Set objLineString2 = oLineString2
    objLineString2.Remove
    Set objLineString2 = Nothing
    Set axis2 = Nothing
    Set CP2 = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
