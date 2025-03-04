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
'   Copyright (c) 2011, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author: BhupatiRaju Neelima
'   Creation Date: 23 Sep 2011
'   Description: Simple physical file for creating the Bolt head and shank of
'   the bolt symbol to be used in the new custome command of MHE.
'   This will also create the centerline for the symbol and a nozzle port for
'   supporting the euipment hole aspect.
'   This class module is the place for user to implement graphical part of VBSymbol for the simple physical aspect
'
'   Change History:
'   dd.mmm.yyyy         who         change description
'   -----------         -----       ------------------
'   23-Sep-2011         Neelima     Creating the file first time
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

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt               As PartFacelets.IJDPart
    Dim oDir                    As IJDVector
    Dim ObjPort                 As IJConduitPortOcc
    Dim oGeomFactory            As IngrGeom3D.GeometryFactory
    Dim oLine                   As IngrGeom3D.Line3d
    Dim oBoltHeadPts(0 To 5)    As IJDPosition
    Dim StartPos                As IJDPosition
    Dim EndPos                  As IJDPosition
    Dim oCenterPos              As IJDPosition
    Dim portPoint               As IJDPosition
    Dim iOutput                 As Double
    Dim DblBoltDiameter         As Double
    Dim DblBoltHeadThickness    As Double
    Dim DblShankLength          As Double
    Dim DblShankDiameter        As Double
    Dim iCount                  As Integer
    Dim DblTempX                As Double
    Dim DblTempY                As Double
    Dim dblAngle                As Double
    Dim ObjBoltHeadColl         As Collection
    Dim ObjBoltShank            As Object

    
    Set oCenterPos = New DPosition
    oCenterPos.Set 0, 0, 0

    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    DblBoltDiameter = arrayOfInputs(2)         '12mm
    DblBoltHeadThickness = arrayOfInputs(3)    '6mm
    DblShankLength = arrayOfInputs(8)          '40mm
    DblShankDiameter = arrayOfInputs(9)        '6mm

    iOutput = 0

    m_oSymGeomHelper.outputCollection = m_OutputColl
    
    ' Constructing the Points for Bolt head
    For iCount = 0 To 5
        Set oBoltHeadPts(iCount) = New DPosition
    Next iCount
    
    dblAngle = 0#
    For iCount = 0 To 5
        DblTempX = oCenterPos.x + (DblBoltDiameter / 2#) * Cos(dblAngle)
        DblTempY = oCenterPos.y + (DblBoltDiameter / 2#) * Sin(dblAngle)
        oBoltHeadPts(iCount).Set DblTempX, DblTempY, oCenterPos.z
        dblAngle = dblAngle + 60# * (PI / 180#)
    Next iCount
    
   'Create the Bolt Head (Output 1)
    Set ObjBoltHeadColl = CreateBoltHead(m_OutputColl, oBoltHeadPts, DblBoltHeadThickness)
    
    'Set the Output
    iOutput = iOutput + 1
    For iCount = 1 To ObjBoltHeadColl.Count
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBoltHeadColl(iCount)
    Next iCount
    For iCount = 0 To 5
        Set oBoltHeadPts(iCount) = Nothing
    Next
    Set ObjBoltHeadColl = Nothing
    
   'Create the Bolt Shank (Output 2)
    Set StartPos = New DPosition
    Set EndPos = New DPosition
    
    StartPos.Set oCenterPos.x, oCenterPos.y, oCenterPos.z
    EndPos.Set oCenterPos.x, oCenterPos.y, oCenterPos.z - DblShankLength

    Set ObjBoltShank = PlaceCylinder(m_OutputColl, StartPos, EndPos, DblShankDiameter, True)
    Set StartPos = Nothing
    Set EndPos = Nothing
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBoltShank
    Set ObjBoltShank = Nothing
    
   'Create the Bolt Centerline (Output 3)

    Set StartPos = New DPosition
    Set EndPos = New DPosition
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oLine = New IngrGeom3D.Line3d
    
    StartPos.Set oCenterPos.x, oCenterPos.y, oCenterPos.z + DblBoltHeadThickness + 0.002
    EndPos.Set oCenterPos.x, oCenterPos.y, oCenterPos.z - (DblShankLength + 0.002)

    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, StartPos.x, StartPos.y, StartPos.z, EndPos.x, EndPos.y, EndPos.z)

    Set StartPos = Nothing
    Set EndPos = Nothing
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oLine
    Set oLine = Nothing
    Set oGeomFactory = Nothing
    
    ' code for output 4 (ConduitPort)
    Set portPoint = New DPosition
    Set oDir = New DVector
    oDir.Set 0, 0, 1
    
    portPoint.Set oCenterPos.x, oCenterPos.y, oCenterPos.z
    Set ObjPort = CreateConduitNozzlePH(portPoint, oDir, m_OutputColl, oPartFclt, 1)
    
    m_OutputColl.AddOutput "ConduitPort", ObjPort
    Set oCenterPos = Nothing
    Set ObjPort = Nothing
    Set portPoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

Private Function CreateBoltHead(ByVal objOutputColl As Object, ByRef HexagonPoints() As IJDPosition, ByRef dblBoltThickness As Double) As Collection
    Const METHOD = "CreateBoltHead"
    On Error GoTo ErrorLabel

    Dim oGeomFactory            As IngrGeom3D.GeometryFactory
    Dim objComplexString        As IngrGeom3D.ComplexString3d
    Dim oLine                   As IngrGeom3D.Line3d
    Dim iElements               As IJElements
    Dim pProj3d                 As Projection3d
    Dim oTmpCollection          As Collection
    
' Hexagon shape which is getting created
'           2------1
'           /\    /\
'          /  \  /  \
'         /    \/    \
'        3-----C-----0
'         \    /\    /
'          \  /  \  /
'           \/    \/
'           4------5
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set objComplexString = New ComplexString3d
    Set iElements = New JObjectCollection ' IMSElements.DynElements
    Set oLine = New IngrGeom3D.Line3d
    Set oTmpCollection = New Collection

    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, HexagonPoints(0).x, HexagonPoints(0).y, HexagonPoints(0).z, HexagonPoints(1).x, HexagonPoints(1).y, HexagonPoints(1).z)
    iElements.Add oLine
    Set objComplexString = oGeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
    Set iElements = Nothing
        
    oLine.DefineBy2Points HexagonPoints(1).x, HexagonPoints(1).y, HexagonPoints(1).z, HexagonPoints(2).x, HexagonPoints(2).y, HexagonPoints(2).z
    objComplexString.AddCurve oLine, True
    
    oLine.DefineBy2Points HexagonPoints(2).x, HexagonPoints(2).y, HexagonPoints(2).z, HexagonPoints(3).x, HexagonPoints(3).y, HexagonPoints(3).z
    objComplexString.AddCurve oLine, True
       
    oLine.DefineBy2Points HexagonPoints(3).x, HexagonPoints(3).y, HexagonPoints(3).z, HexagonPoints(4).x, HexagonPoints(4).y, HexagonPoints(4).z
    objComplexString.AddCurve oLine, True
    
    oLine.DefineBy2Points HexagonPoints(4).x, HexagonPoints(4).y, HexagonPoints(4).z, HexagonPoints(5).x, HexagonPoints(5).y, HexagonPoints(5).z
    objComplexString.AddCurve oLine, True
    
    oLine.DefineBy2Points HexagonPoints(5).x, HexagonPoints(5).y, HexagonPoints(5).z, HexagonPoints(0).x, HexagonPoints(0).y, HexagonPoints(0).z
    objComplexString.AddCurve oLine, True
    
    Set pProj3d = oGeomFactory.Projections3d.CreateByCurve(objOutputColl.ResourceManager, objComplexString, 0#, 0#, 1#, dblBoltThickness, True)
    oTmpCollection.Add pProj3d
    
    Set CreateBoltHead = oTmpCollection
    
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

Private Function CreateConduitNozzlePH(oBasePt As AutoMath.DPosition, oAxis As AutoMath.DVector, ByVal objOutputColl As Object, ByRef oPart As PartFacelets.IJDPart, dNozzleIndex As Long) As GSCADNozzleEntities.IJConduitPortOcc
    ' This subroutine creates a ConduitNozzle  and sets it's position and direction
    Dim oNozzlePHFactory As NozzlePHFactory
    Set oNozzlePHFactory = New NozzlePHFactory
    Dim NullObj As Object
    Dim oDistribPort As IJDistribPort
    Dim oLogicalDistPort As IJLogicalDistPort
    Dim oConduitNozzle As IJConduitPortOcc
    Const METHOD = "CreateConduitNozzlePH:"

    On Error GoTo ErrHandler
    
    Set oConduitNozzle = oNozzlePHFactory.CreateConduitNozzlePHFromPart(oPart, dNozzleIndex, _
                                                                objOutputColl.ResourceManager)
    Set oLogicalDistPort = oConduitNozzle
    Set oDistribPort = oConduitNozzle
    

    oLogicalDistPort.SetCenterLocation oBasePt
    
    oDistribPort.SetDirectionVector oAxis
    
    Set CreateConduitNozzlePH = oConduitNozzle
        
    Set oNozzlePHFactory = Nothing
    Set oDistribPort = Nothing
    Set oLogicalDistPort = Nothing
    Set oConduitNozzle = Nothing
    Exit Function
    
ErrHandler:
  ReportUnanticipatedError MODULE, METHOD
End Function

