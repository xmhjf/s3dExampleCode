VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2010, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author: BhupatiRaju
'   Creation Date: 08 Nov 2010
'   Description: Simple physical file for creating the head section of
'   the chute symbol to be used in the Layout desgn command of MHE.
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy          who                     change description
'   -----------         -----                   ------------------
'   08-Nov-2010         Neelima                 Creating the file first time
'   27-Sep-2012         Shireesha M             TR-CP-221650    2D HeasSectionSymbol is changed in Shape & position with changing HSLength
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CSimplePhysical:" 'Used for error messages
Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private PI As Double

Private Sub Class_Initialize()

    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
     Set m_oSymGeomHelper = New SymbolServices
     PI = 4 * Atn(1)
     
     Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext

End Sub

Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
End Sub

Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt               As PartFacelets.IJDPart
    Dim oCenterPos              As IJDPosition
    Dim iOutput                 As Double
    Dim DblHeadSecLength        As Double
    Dim DblHeadSecWidth         As Double
    Dim DblHeadSecDepth         As Double
    Dim DblPulleyCPLength       As Double
    Dim DblPulleyCPDepth        As Double
    Dim DblCutOutRadius         As Double
    Dim DblHeadSecOpenLength    As Double
    Dim DblHeadSecConvAngle1    As Double
    Dim oFrontPts(0 To 10)      As IJDPosition
    Dim iCount                  As Integer
    Dim DblTempX1               As Double
    Dim DblTempZ1               As Double
    Dim ObjHeadSecColl          As Collection
    Dim ObjPointColl            As Collection
    Dim oGeomFactory            As IngrGeom3D.GeometryFactory
    Dim oGBSFactory             As IJGeneralBusinessObjectsFactory
    Dim oControlPoint           As IJControlPoint
    
    Set oCenterPos = New DPosition
    oCenterPos.Set 0, 0, 0
    
    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    DblHeadSecLength = arrayOfInputs(2)         '3.5
    DblHeadSecWidth = arrayOfInputs(3)          '4.0
    DblHeadSecDepth = arrayOfInputs(4)          '3.0
    DblPulleyCPLength = arrayOfInputs(5)        '0.5
    DblPulleyCPDepth = arrayOfInputs(6)         '1.5
    DblCutOutRadius = arrayOfInputs(7)          '0.15
    DblHeadSecOpenLength = arrayOfInputs(8)     '0.8
    DblHeadSecConvAngle1 = arrayOfInputs(9)     '45
    
    iOutput = 0
     
    m_oSymGeomHelper.outputCollection = m_outputColl
 
    ' Constructing the Points for front face of head section of chute.
    For iCount = 0 To 10
        Set oFrontPts(iCount) = New DPosition
    Next iCount
     
    'TR-CP-221650
    DblTempX1 = DblHeadSecOpenLength + DblPulleyCPLength
    DblTempZ1 = DblTempX1 * Round((Tan(DblHeadSecConvAngle1)), 2)
    
    'TR-CP-221650
    oFrontPts(0).Set -DblHeadSecLength, -DblHeadSecWidth / 2#, DblPulleyCPDepth
    oFrontPts(1).Set DblPulleyCPLength, -DblHeadSecWidth / 2#, DblPulleyCPDepth
    oFrontPts(2).Set DblPulleyCPLength, -DblHeadSecWidth / 2#, DblCutOutRadius
    oFrontPts(3).Set oCenterPos.x, -DblHeadSecWidth / 2#, DblCutOutRadius
    oFrontPts(4).Set oCenterPos.x, -DblHeadSecWidth / 2#, oCenterPos.z
    oFrontPts(5).Set -DblCutOutRadius, -DblHeadSecWidth / 2#, oCenterPos.z
    oFrontPts(6).Set oCenterPos.x, -DblHeadSecWidth / 2#, -DblCutOutRadius
    oFrontPts(7).Set DblPulleyCPLength, -DblHeadSecWidth / 2#, -DblCutOutRadius
    oFrontPts(8).Set DblPulleyCPLength, -DblHeadSecWidth / 2#, -(DblHeadSecDepth - DblTempZ1)
    oFrontPts(9).Set -(DblHeadSecOpenLength), -DblHeadSecWidth / 2#, -DblHeadSecDepth
    oFrontPts(10).Set -DblHeadSecLength, -DblHeadSecWidth / 2#, -DblHeadSecDepth

   'Create the Head Section (Output 1)
    Set ObjHeadSecColl = CreateChuteShape(m_outputColl, oFrontPts, DblHeadSecWidth)
    'Set the Output
    iOutput = iOutput + 1
    For iCount = 1 To ObjHeadSecColl.Count
        m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjHeadSecColl(iCount)
    Next iCount
    Set ObjHeadSecColl = Nothing
    
    '============================================================================
    '    Construction of  Control Point as the origin of head section ( Output 2)
    '============================================================================

    Set oGBSFactory = New GeneralBusinessObjectsFactory
    Set oControlPoint = oGBSFactory.CreateControlPoint(m_outputColl.ResourceManager, _
                                                        0, 0, 0, 0.02, , , False)

    oControlPoint.Type = cpControlPoint
    oControlPoint.SubType = cpProcessEquipment
    
    'Set the Output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), oControlPoint
    Set oControlPoint = Nothing
    Set oGBSFactory = Nothing
    
    '============================================================================
    'Create the Reference Points on the bottoms face of head section (Output 3)
    '============================================================================
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set ObjPointColl = New Collection
    
    'Creating the points on all the four vertices of bottom face
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, oFrontPts(8).x, oFrontPts(8).y, oFrontPts(8).z)
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, oFrontPts(9).x, oFrontPts(9).y, oFrontPts(9).z)
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, oFrontPts(8).x, -oFrontPts(8).y, oFrontPts(8).z)
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, oFrontPts(9).x, -oFrontPts(9).y, oFrontPts(9).z)
    
    'Set the output
    iOutput = iOutput + 1
    For iCount = 1 To ObjPointColl.Count
        m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjPointColl(iCount)
    Next iCount
    
    For iCount = 0 To 10
        Set oFrontPts(iCount) = Nothing
    Next
    Set oGeomFactory = Nothing
    Set ObjPointColl = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub


Private Function CreateChuteShape(ByVal objOutputColl As Object, ByRef topSurfacePoints() As IJDPosition, ByRef dblWidth As Double) As Collection
    Const METHOD = "CreateChuteShape"
    On Error GoTo ErrorLabel

    Dim oGeomFactory            As IngrGeom3D.GeometryFactory
    Dim objComplexString        As IngrGeom3D.ComplexString3d
    Dim oLine                   As IngrGeom3D.Line3d
    Dim oArc                    As IngrGeom3D.Arc3d
    Dim iElements               As IJElements
    Dim pProj3d                 As Projection3d
    Dim oTmpCollection          As Collection
    
    
'          0 -------------------1
'           |                   |
'           |                   |
'           |                   |
'           |         / 3-------2
'           |        /  |
'           |      5|  4|
'           |        \  |
'           |         \ 6-------7
'           |                   |
'           |                   |
'           |                   8
'           |                  /
'           |                 /
'           |                /
'           10-------------9

    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set objComplexString = New ComplexString3d
    Set iElements = New JObjectCollection ' IMSElements.DynElements
    Set oLine = New IngrGeom3D.Line3d
    Set oArc = New IngrGeom3D.Arc3d
    Set oTmpCollection = New Collection

    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, topSurfacePoints(0).x, topSurfacePoints(0).y, topSurfacePoints(0).z, topSurfacePoints(1).x, topSurfacePoints(1).y, topSurfacePoints(1).z)
    iElements.Add oLine
    Set objComplexString = oGeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
    Set iElements = Nothing
        
    oLine.DefineBy2Points topSurfacePoints(1).x, topSurfacePoints(1).y, topSurfacePoints(1).z, topSurfacePoints(2).x, topSurfacePoints(2).y, topSurfacePoints(2).z
    objComplexString.AddCurve oLine, True
    
    oLine.DefineBy2Points topSurfacePoints(2).x, topSurfacePoints(2).y, topSurfacePoints(2).z, topSurfacePoints(3).x, topSurfacePoints(3).y, topSurfacePoints(3).z
    objComplexString.AddCurve oLine, True
    
    Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, topSurfacePoints(4).x, topSurfacePoints(4).y, topSurfacePoints(4).z, topSurfacePoints(3).x, topSurfacePoints(3).y, topSurfacePoints(3).z, topSurfacePoints(5).x, topSurfacePoints(5).y, topSurfacePoints(5).z)
    objComplexString.AddCurve oArc, True
    
    oArc.DefineByCenterStartEnd topSurfacePoints(4).x, topSurfacePoints(4).y, topSurfacePoints(4).z, topSurfacePoints(5).x, topSurfacePoints(5).y, topSurfacePoints(5).z, topSurfacePoints(6).x, topSurfacePoints(6).y, topSurfacePoints(6).z
    objComplexString.AddCurve oArc, True
       
    oLine.DefineBy2Points topSurfacePoints(6).x, topSurfacePoints(6).y, topSurfacePoints(6).z, topSurfacePoints(7).x, topSurfacePoints(7).y, topSurfacePoints(7).z
    objComplexString.AddCurve oLine, True
    
    oLine.DefineBy2Points topSurfacePoints(7).x, topSurfacePoints(7).y, topSurfacePoints(7).z, topSurfacePoints(8).x, topSurfacePoints(8).y, topSurfacePoints(8).z
    objComplexString.AddCurve oLine, True
    
    oLine.DefineBy2Points topSurfacePoints(8).x, topSurfacePoints(8).y, topSurfacePoints(8).z, topSurfacePoints(9).x, topSurfacePoints(9).y, topSurfacePoints(9).z
    objComplexString.AddCurve oLine, True
       
    oLine.DefineBy2Points topSurfacePoints(9).x, topSurfacePoints(9).y, topSurfacePoints(9).z, topSurfacePoints(10).x, topSurfacePoints(10).y, topSurfacePoints(10).z
    objComplexString.AddCurve oLine, True
    
    oLine.DefineBy2Points topSurfacePoints(10).x, topSurfacePoints(10).y, topSurfacePoints(10).z, topSurfacePoints(0).x, topSurfacePoints(0).y, topSurfacePoints(0).z
    objComplexString.AddCurve oLine, True
    
    Set pProj3d = oGeomFactory.Projections3d.CreateByCurve(objOutputColl.ResourceManager, objComplexString, 0#, 1#, 0#, dblWidth, True)
    oTmpCollection.Add pProj3d
    
    Set CreateChuteShape = oTmpCollection
    
    Set oLine = Nothing
    Set oArc = Nothing
    Set objComplexString = Nothing
    Set oGeomFactory = Nothing
    Set pProj3d = Nothing
    Set oTmpCollection = Nothing
    
    Exit Function
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
End Function

