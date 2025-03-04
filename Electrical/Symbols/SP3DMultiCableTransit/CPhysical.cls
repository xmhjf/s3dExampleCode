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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.

'   ProgID          :  SP3DMultiCableTransit.MCableTransit
'   File            :  CPhysical.cls
'   Author          :  PK
'   Creation Date   :  Monday, 1 Oct 2007
'   Description     :  Multiple Cable Transit
'   Change History  :
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   19-NOV-2007      PK        TR-129855:Created the symbol.
'   02.JUN.2008      PK        CR-143345  Update MCT symbol to be able to specify offset for each port independently
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
''''

End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput                 As Double
    Dim ActualWidth             As Double
    Dim ActualDepth             As Double
    Dim parWidth                As Double
    Dim parDepth                As Double
    Dim parLength               As Double
    Dim parPortOffset           As Double
    Dim parThickness            As Double
    Dim parPartitionsAlongWidth As Long
    Dim parPartitionsAlongDepth As Long
    Dim parVerticalPortOffset   As Double
    Dim parHorizontalPortOffset As Double
    Dim parPort2Offset          As Double
    
    Dim dTotalWidth             As Double
    Dim dPort1Offset            As Double
    Dim dPort2Offset            As Double
    
    Dim oPort1   As AutoMath.DPosition
    Dim oPort2   As AutoMath.DPosition
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    
    Set oPort1 = New AutoMath.DPosition
    Set oPort2 = New AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    Dim oGeomFactory   As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    iOutput = 0
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    parLength = arrayOfInputs(4)
    parPortOffset = arrayOfInputs(5)
    parThickness = arrayOfInputs(6)
    parPartitionsAlongWidth = arrayOfInputs(7)
    parPartitionsAlongDepth = arrayOfInputs(8)
    parVerticalPortOffset = arrayOfInputs(9)
    parHorizontalPortOffset = arrayOfInputs(10)
    
    If UBound(arrayOfInputs) >= 11 Then
        parPort2Offset = arrayOfInputs(11)
        dPort1Offset = parPortOffset
        dPort2Offset = parPort2Offset
    Else
        dPort1Offset = parPortOffset
        dPort2Offset = parPortOffset
    End If
        
    dTotalWidth = (parPartitionsAlongWidth + 1) * parWidth
    'Create Upper body
    Dim objUpperBody As Object
    oStPoint.Set -parLength / 2, dTotalWidth / 2, parDepth / 2 - parThickness
    oEnPoint.Set parLength / 2, -dTotalWidth / 2, parDepth / 2
    Set objUpperBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput "UpperBody", objUpperBody
    Set objUpperBody = Nothing
    
    'Create Left body
    Dim objLeftBody As Object
    oStPoint.Set -parLength / 2, dTotalWidth / 2, parDepth / 2 - parThickness
    oEnPoint.Set parLength / 2, (dTotalWidth / 2) - parThickness, -parDepth / 2 + parThickness
    Set objLeftBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput "LeftBody", objLeftBody
    Set objLeftBody = Nothing
                    
    'Create Bottom body
    Dim objBottomBody As Object
    oStPoint.Set -parLength / 2, dTotalWidth / 2, -parDepth / 2
    oEnPoint.Set parLength / 2, -dTotalWidth / 2, -parDepth / 2 + parThickness
    Set objBottomBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput "BottomBody", objBottomBody
    Set objBottomBody = Nothing
    
    'Create Right body
    Dim objRightBody As Object
    oStPoint.Set -parLength / 2, -dTotalWidth / 2, parDepth / 2 - parThickness
    oEnPoint.Set parLength / 2, -(dTotalWidth / 2) + parThickness, -parDepth / 2 + parThickness
    Set objRightBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput "RightBody", objRightBody
    Set objRightBody = Nothing
    
    'Create Partitions along Width
    Dim oPartition                  As Object
    Dim dPartThick                  As Double
    Dim dLineStrPoints(0 To 15)     As Double
    Dim dSpaceBwPartitionsAlgWidth  As Double

    'Assumption: dPartThick is taken as parThickness / 10
    dPartThick = parThickness / 10
    dSpaceBwPartitionsAlgWidth = (dTotalWidth - 2 * parThickness) / _
                                        (parPartitionsAlongWidth + 1)

    dLineStrPoints(0) = -parLength / 2
    dLineStrPoints(1) = (dTotalWidth / 2 - parThickness - _
                                        dSpaceBwPartitionsAlgWidth + (dPartThick / 2))
    dLineStrPoints(2) = (parDepth / 2 - parThickness)

    dLineStrPoints(3) = dLineStrPoints(0)
    dLineStrPoints(4) = dLineStrPoints(1)
    dLineStrPoints(5) = -dLineStrPoints(2)

    dLineStrPoints(6) = dLineStrPoints(3)
    dLineStrPoints(7) = dLineStrPoints(4) - dPartThick
    dLineStrPoints(8) = dLineStrPoints(5)

    dLineStrPoints(9) = dLineStrPoints(6)
    dLineStrPoints(10) = dLineStrPoints(7)
    dLineStrPoints(11) = dLineStrPoints(2)

    dLineStrPoints(12) = dLineStrPoints(0)
    dLineStrPoints(13) = dLineStrPoints(1)
    dLineStrPoints(14) = dLineStrPoints(2)

    Dim oLineString As LineString3d
    Set oLineString = New LineString3d
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)

    Dim oTransMatrix    As New DT4x4
    Dim iCount          As Integer

    iOutput = iOutput + 1

    For iCount = 1 To parPartitionsAlongWidth
        'Assumption: Partition length is taken to be 0.1 * parLength
        Set oPartition = oGeomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
                                                            oLineString, 1, 0, 0, parLength, True)
        Dim oTransVector As New AutoMath.DVector
        oTransVector.Set 0, -(iCount - 1) * dSpaceBwPartitionsAlgWidth, 0

        oTransMatrix.LoadIdentity
        oTransMatrix.Translate oTransVector
        oPartition.Transform oTransMatrix
        m_OutputColl.AddOutput "PartitionsAlongWidth_", oPartition
        Set oPartition = Nothing
    Next iCount

    If parPartitionsAlongDepth <> 0 Then
        'Create Partitions along Depth
        Dim dSpaceBwPartitionsAlgDepth As Double
        dSpaceBwPartitionsAlgDepth = (parDepth - 2 * parThickness) / _
                                            (parPartitionsAlongDepth + 1)
        
        dLineStrPoints(0) = -parLength / 2
        dLineStrPoints(1) = (dTotalWidth / 2 - parThickness)
        dLineStrPoints(2) = (parDepth / 2 - parThickness - _
                                            dSpaceBwPartitionsAlgDepth - dPartThick / 2)
        
        dLineStrPoints(3) = dLineStrPoints(0)
        dLineStrPoints(4) = dLineStrPoints(1)
        dLineStrPoints(5) = dLineStrPoints(2) + dPartThick
        
        dLineStrPoints(6) = dLineStrPoints(3)
        dLineStrPoints(7) = -dLineStrPoints(1)
        dLineStrPoints(8) = dLineStrPoints(5)
        
        dLineStrPoints(9) = dLineStrPoints(6)
        dLineStrPoints(10) = -dLineStrPoints(1)
        dLineStrPoints(11) = dLineStrPoints(2)
        
        dLineStrPoints(12) = dLineStrPoints(0)
        dLineStrPoints(13) = dLineStrPoints(1)
        dLineStrPoints(14) = dLineStrPoints(2)
        
        Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)
        iOutput = iOutput + 1
        For iCount = 1 To parPartitionsAlongDepth
            'Assumption: Partition length is taken to be 0.1 * parLength
            Set oPartition = oGeomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
                                            oLineString, 1, 0, 0, parLength, True)
                                            
            oTransVector.Set 0, 0, -(iCount - 1) * dSpaceBwPartitionsAlgDepth
            oTransMatrix.LoadIdentity
            oTransMatrix.Translate oTransVector
            oPartition.Transform oTransMatrix
            
            m_OutputColl.AddOutput "PartitionsAlongDepth_", oPartition
            Set oPartition = Nothing
        Next iCount
        Set oGeomFactory = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        Set oLineString = Nothing
        Set oTransMatrix = Nothing
        Set oTransVector = Nothing
    Else
        iOutput = iOutput + 1
    End If
    
    'Place Port 1
    Dim oDir                As AutoMath.DVector
    Dim oRadialOrient       As AutoMath.DVector
    Dim oCableTrayPort      As GSCADNozzleEntities.IJCableTrayPortOcc

    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    oDir.Set -1, 0, 0
    oRadialOrient.Set 0, 0, 1
    oPort1.Set -parLength / 2 - dPort1Offset, parHorizontalPortOffset, parVerticalPortOffset
    
    Set oCableTrayPort = CreateCableTrayPort(oPartFclt, 1, oPort1, oDir, _
                                                oRadialOrient, m_OutputColl)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oCableTrayPort
    Set oCableTrayPort = Nothing
    Set oPort1 = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    
    'Place Port 2
    Set oDir = New AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    oPort2.Set parLength / 2 + dPort2Offset, parHorizontalPortOffset, parVerticalPortOffset
    oDir.Set 1, 0, 0
    oRadialOrient.Set 0, 0, 1
    
    Set oCableTrayPort = CreateCableTrayPort(oPartFclt, 2, oPort2, oDir, _
                                                oRadialOrient, m_OutputColl)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oCableTrayPort
    Set oCableTrayPort = Nothing
    Set oDir = Nothing
    Set oRadialOrient = Nothing
    Set oPort2 = Nothing
    Set oGeomFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
