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
'   dd.mmm.yyyy         who         change description
'   -----------         -----       ------------------
'   08-Nov-2010         Neelima     Creating the file first time
'   07-Oct-2011         Nagesh      CR-CP-199489  As part of this CR symbol origin is moved to center of top face
'
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
    
    Dim oPartFclt                As PartFacelets.IJDPart
    Dim oCenterPos               As IJDPosition
    Dim iOutput                  As Double
    Dim DblSSTotalLength         As Double
    Dim DblSSTotalWidth          As Double
    Dim DblSSTotalDepth          As Double
    Dim DblSSLength              As Double
    Dim DblSSDepth1              As Double
    Dim DblSSDepth2              As Double
    Dim DblSSAngle1              As Double
    Dim DblSSAngle2              As Double
    Dim DblSSOpeningLength       As Double
    Dim oFrontPts(0 To 5)        As IJDPosition
    Dim iCount                   As Integer
    Dim DblTempX1                As Double
    Dim DblTempX2                As Double
    Dim DblTempX                 As Double
    Dim DblTempY                 As Double
    Dim DblTemp                  As Double
    Dim DblSSAngle3              As Double
    Dim ObjSpoonSecColl          As Collection
    Dim ObjPointColl             As Collection
    Dim oGeomFactory             As IngrGeom3D.GeometryFactory
    
    Set oCenterPos = New DPosition
    oCenterPos.Set 0, 0, 0
    
    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    DblSSTotalLength = arrayOfInputs(2)         '2.175
    DblSSTotalWidth = arrayOfInputs(3)          '0.85
    DblSSTotalDepth = arrayOfInputs(4)          '1.26
    DblSSLength = arrayOfInputs(5)         '1.25
    DblSSDepth1 = arrayOfInputs(6)         '0.4
    DblSSDepth2 = arrayOfInputs(7)          '0.42
    DblSSAngle1 = arrayOfInputs(8)  '70
    DblSSAngle2 = arrayOfInputs(9)    '30
    DblSSOpeningLength = arrayOfInputs(10)   '1.125
        
    iOutput = 0
     
    m_oSymGeomHelper.outputCollection = m_outputColl
 
    ' Constructing the Points for front face of head section of chute.
    For iCount = 0 To 5
        Set oFrontPts(iCount) = New DPosition
    Next iCount
    DblSSAngle3 = ((90# * PI) / 180# - DblSSAngle2) - ((90# * PI) / 180# - DblSSAngle1)
    DblTempX1 = DblSSTotalDepth / Round((Tan(DblSSAngle1)), 2)
    DblTempX2 = DblSSTotalLength - DblSSOpeningLength
    DblTemp = ((DblTempX2 - DblTempX1) * Round((Sin(DblSSAngle2)), 2)) / Round((Sin(DblSSAngle3)), 2)
    DblTempY = DblTemp * Round((Sin(DblSSAngle1)), 2)
    DblTempX = DblTemp * Round((Cos(DblSSAngle1)), 2)

    Dim oTempPos    As IJDPosition                              'CR-CP-199489
    Set oTempPos = New DPosition                                'CR-CP-199489
    oTempPos.Set -DblSSLength / 2, -DblSSTotalWidth / 2, 0      'CR-CP-199489

    'Calculating the points for spoon shape
    oFrontPts(0).Set oTempPos.x, oTempPos.y, oTempPos.z
    oFrontPts(1).Set oTempPos.x + DblSSLength, oTempPos.y, oTempPos.z
    oFrontPts(2).Set oFrontPts(1).x - (DblTempX1 - DblTempX), oTempPos.y, -(oTempPos.z + DblSSTotalDepth - DblTempY)
    oFrontPts(3).Set oFrontPts(1).x - (DblSSTotalLength - DblSSOpeningLength), oTempPos.y, -(oTempPos.z + DblSSTotalDepth)
    oFrontPts(4).Set (oTempPos.x - (DblSSTotalLength - DblSSLength)), oTempPos.y, -(oTempPos.z + DblSSTotalDepth - DblSSDepth2)
    oFrontPts(5).Set (oTempPos.x - DblSSDepth1 / Round((Tan(DblSSAngle1)), 2)), oTempPos.y, -(oTempPos.z + DblSSDepth1)
    
    Set oTempPos = Nothing                                      'CR-CP-199489

   'Create the Head Section (Output 1)
    Set ObjSpoonSecColl = CreateChuteShape(m_outputColl, oFrontPts, DblSSTotalWidth)
    'Set the Output
    iOutput = iOutput + 1
    For iCount = 1 To ObjSpoonSecColl.Count
        m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjSpoonSecColl(iCount)
    Next iCount
    Set ObjSpoonSecColl = Nothing
    
    '============================================================================
    'Create the Reference Points on the bottoms face of head section (Output 3)
    '============================================================================
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set ObjPointColl = New Collection
    
    'Creating the points on all the four vertices of bottom face
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, oFrontPts(0).x, oFrontPts(0).y, oFrontPts(0).z)
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, oFrontPts(1).x, oFrontPts(1).y, oFrontPts(1).z)
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, oFrontPts(1).x, oFrontPts(1).y + DblSSTotalWidth, oFrontPts(1).z)
    ObjPointColl.Add oGeomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, oFrontPts(0).x, oFrontPts(0).y + DblSSTotalWidth, oFrontPts(0).z)
    
    'Set the output
    iOutput = iOutput + 1
    For iCount = 1 To ObjPointColl.Count
        m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjPointColl(iCount)
    Next iCount
    
    For iCount = 0 To 5
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
    Dim iElements               As IJElements
    Dim pProj3d                 As Projection3d
    Dim oTmpCollection          As Collection
    
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set objComplexString = New ComplexString3d
    Set iElements = New JObjectCollection ' IMSElements.DynElements
    Set oLine = New IngrGeom3D.Line3d
    Set oTmpCollection = New Collection
    
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, topSurfacePoints(0).x, topSurfacePoints(0).y, topSurfacePoints(0).z, topSurfacePoints(1).x, topSurfacePoints(1).y, topSurfacePoints(1).z)
    iElements.Add oLine
    Set objComplexString = oGeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
    Set iElements = Nothing
        
    oLine.DefineBy2Points topSurfacePoints(1).x, topSurfacePoints(1).y, topSurfacePoints(1).z, topSurfacePoints(2).x, topSurfacePoints(2).y, topSurfacePoints(2).z
    objComplexString.AddCurve oLine, True
    
    oLine.DefineBy2Points topSurfacePoints(2).x, topSurfacePoints(2).y, topSurfacePoints(2).z, topSurfacePoints(3).x, topSurfacePoints(3).y, topSurfacePoints(3).z
    objComplexString.AddCurve oLine, True
       
    oLine.DefineBy2Points topSurfacePoints(3).x, topSurfacePoints(3).y, topSurfacePoints(3).z, topSurfacePoints(4).x, topSurfacePoints(4).y, topSurfacePoints(4).z
    objComplexString.AddCurve oLine, True
    
    oLine.DefineBy2Points topSurfacePoints(4).x, topSurfacePoints(4).y, topSurfacePoints(4).z, topSurfacePoints(5).x, topSurfacePoints(5).y, topSurfacePoints(5).z
    objComplexString.AddCurve oLine, True
    
    oLine.DefineBy2Points topSurfacePoints(5).x, topSurfacePoints(5).y, topSurfacePoints(5).z, topSurfacePoints(0).x, topSurfacePoints(0).y, topSurfacePoints(0).z
    objComplexString.AddCurve oLine, True
          
    Set pProj3d = oGeomFactory.Projections3d.CreateByCurve(objOutputColl.ResourceManager, objComplexString, 0#, 1#, 0#, dblWidth, True)
    oTmpCollection.Add pProj3d
    
    Set CreateChuteShape = oTmpCollection
    
    Set oLine = Nothing
    Set objComplexString = Nothing
    Set oGeomFactory = Nothing
    Set pProj3d = Nothing
    Set oTmpCollection = Nothing
    
    Exit Function
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
End Function

