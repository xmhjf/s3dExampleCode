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
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         ACM/MS
'   Creation Date:  Monday, July 27, 2004
'   Description:
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Dim m_oGeomHelper As IJSymbolGeometryHelper
Dim PI As Double
Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private Sub Class_Terminate()
    Set m_oGeomHelper = Nothing
End Sub

Private Sub Class_Initialize()
    Set m_oGeomHelper = New SymbolServices
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
     Dim pipeDiam1        As Double
    Dim pipeDiam2        As Double
    Dim pipeDiam3        As Double
    Dim pipeDiam4        As Double
    Dim pipeDiam5        As Double
    Dim pipeDiam6        As Double
    Dim pipeDiam7        As Double
    Dim pipeDiam8        As Double
    Dim flangeThick     As Double
    Dim flangeDiam1      As Double
    Dim flangeDiam2      As Double
    Dim flangeDiam3      As Double
    Dim flangeDiam4      As Double
    Dim flangeDiam5      As Double
    Dim flangeDiam6      As Double
    Dim flangeDiam7      As Double
    Dim flangeDiam8      As Double
    Dim cptOffset1       As Double
    Dim depth1           As Double
    Dim cptOffset2       As Double
    Dim depth2           As Double
    Dim cptOffset3       As Double
    Dim depth3           As Double
    Dim cptOffset4       As Double
    Dim depth4           As Double
    Dim cptOffset5       As Double
    Dim depth5           As Double
    Dim cptOffset6       As Double
    Dim depth6           As Double
    Dim cptOffset7       As Double
    Dim depth7           As Double
    Dim cptOffset8       As Double
    Dim depth8           As Double
    
    Dim iOutput     As Double
    
    Dim parValveBodyWidth As Double
    Dim parValveBodyHeight As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parFace4toCenter As Double
    Dim parFace5toCenter As Double
    Dim parFace6toCenter As Double
    Dim parFace7toCenter As Double
    Dim parFace8toCenter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parValveBodyWidth = arrayOfInputs(2)
    parValveBodyHeight = arrayOfInputs(3)
    parFace1toCenter = arrayOfInputs(4)
    parFace2toCenter = arrayOfInputs(5)
    parFace3toCenter = arrayOfInputs(6)
    parFace4toCenter = arrayOfInputs(7)
    parFace5toCenter = arrayOfInputs(8)
    parFace6toCenter = arrayOfInputs(9)
    parFace7toCenter = arrayOfInputs(10)
    parFace8toCenter = arrayOfInputs(11)
    parInsulationThickness = arrayOfInputs(12)
    
    iOutput = 0
    
    'Origin is taken at the geometry center of the valve.
    
    m_oGeomHelper.OutputCollection = m_OutputColl
    
'   Insert your code for output 1(Top Taper Body)
    Dim ObjValveBody As Object
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    Dim oTopLineString As IngrGeom3D.LineString3d
    Dim oBottomLineString As IngrGeom3D.LineString3d
    Dim oDirVec As New AutoMath.DVector
    
    Dim dLinePoints(0 To 20) As Double
    Dim dStepHeight As Double
    Dim dValveBodyTopWidth As Double
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, _
                        flangeThick, flangeDiam1, cptOffset1, depth1
    
    dValveBodyTopWidth = pipeDiam1 * 1.5 ''Side of the Triaglur surface is consiered as 1.2 times the Pipe Dia of port 1.
    
    
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, _
                        flangeThick, flangeDiam3, cptOffset3, depth3
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, _
                        flangeThick, flangeDiam4, cptOffset4, depth4
    RetrieveParameters 5, oPartFclt, m_OutputColl, pipeDiam5, _
                        flangeThick, flangeDiam5, cptOffset5, depth5
    RetrieveParameters 6, oPartFclt, m_OutputColl, pipeDiam6, _
                        flangeThick, flangeDiam6, cptOffset6, depth6
    RetrieveParameters 7, oPartFclt, m_OutputColl, pipeDiam7, _
                        flangeThick, flangeDiam7, cptOffset7, depth7
    RetrieveParameters 8, oPartFclt, m_OutputColl, pipeDiam8, _
                        flangeThick, flangeDiam8, cptOffset8, depth8
    
         
 ''Step height is assumed to be 1.5 times the biggest Pipe diameter of Port3,Port4,Port5,Port6,Port7,Port8
    dStepHeight = 1.5 * Max6(pipeDiam3, pipeDiam4, pipeDiam5, pipeDiam6, pipeDiam7, pipeDiam8)
    
   
'Valve Body top surface Line points.
    Dim dHexagonSide As Double
    dHexagonSide = pipeDiam1 / (2 * Sin(PI / 3))
    
    dLinePoints(0) = dHexagonSide / 2
    dLinePoints(1) = parValveBodyHeight / 2
    dLinePoints(2) = dHexagonSide * Cos(PI / 6)
    
    dLinePoints(3) = dHexagonSide / 2
    dLinePoints(4) = dLinePoints(1)
    dLinePoints(5) = 0
            
    dLinePoints(6) = dLinePoints(0)
    dLinePoints(7) = dLinePoints(1)
    dLinePoints(8) = -dLinePoints(2)
    
    dLinePoints(9) = -dLinePoints(6)
    dLinePoints(10) = dLinePoints(7)
    dLinePoints(11) = dLinePoints(8)
    
    dLinePoints(12) = -dLinePoints(3)
    dLinePoints(13) = dLinePoints(4)
    dLinePoints(14) = dLinePoints(5)
    
    dLinePoints(15) = -dLinePoints(0)
    dLinePoints(16) = dLinePoints(1)
    dLinePoints(17) = dLinePoints(2)
   
    dLinePoints(18) = dLinePoints(0)
    dLinePoints(19) = dLinePoints(1)
    dLinePoints(20) = dLinePoints(2)
    
    Set oTopLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLinePoints)
    
 ''Valve Body Bottom surface Line points.
    dHexagonSide = parValveBodyWidth / 2
    
    dLinePoints(0) = dHexagonSide / 2
    dLinePoints(1) = dStepHeight / 2
    dLinePoints(2) = dHexagonSide * Cos(PI / 6)
    
    dLinePoints(3) = parValveBodyWidth / 2
    dLinePoints(4) = dLinePoints(1)
    dLinePoints(5) = 0
            
    dLinePoints(6) = dLinePoints(0)
    dLinePoints(7) = dLinePoints(1)
    dLinePoints(8) = -dLinePoints(2)
    
    dLinePoints(9) = -dLinePoints(6)
    dLinePoints(10) = dLinePoints(7)
    dLinePoints(11) = dLinePoints(8)
    
    dLinePoints(12) = -dLinePoints(3)
    dLinePoints(13) = dLinePoints(4)
    dLinePoints(14) = dLinePoints(5)
    
    dLinePoints(15) = -dLinePoints(0)
    dLinePoints(16) = dLinePoints(1)
    dLinePoints(17) = dLinePoints(2)
   
    dLinePoints(18) = dLinePoints(0)
    dLinePoints(19) = dLinePoints(1)
    dLinePoints(20) = dLinePoints(2)
    
    Set oBottomLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLinePoints)
    
    Set ObjValveBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                                    oTopLineString, oBottomLineString, True)
            
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBody
    Set ObjValveBody = Nothing
    Set oTopLineString = Nothing
    
'   Insert your code for output 2(Stepped part of valve body)
    Dim ObjValveBodyStep As Object
    Dim oStepLineString As IngrGeom3D.LineString3d
    Dim iCount As Integer
    
    For iCount = 0 To 5
        dLinePoints(3 * iCount) = dLinePoints(3 * iCount)
        dLinePoints(3 * iCount + 1) = -dLinePoints(3 * iCount + 1)
        dLinePoints(3 * iCount + 2) = dLinePoints(3 * iCount + 2)
    Next
    
    dLinePoints(18) = dLinePoints(0)
    dLinePoints(19) = dLinePoints(1)
    dLinePoints(20) = dLinePoints(2)
    
    Set oStepLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLinePoints)
    
    Set ObjValveBodyStep = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, oBottomLineString, _
                                                                   oStepLineString, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBodyStep
    Set ObjValveBodyStep = Nothing

' Insert your code for output 3(Bottom Taper part of Valve Body)
     RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, _
                        flangeThick, flangeDiam2, cptOffset2, depth2
    
    dValveBodyTopWidth = pipeDiam2 * 1.5 ''Side of the Triaglur surface is consiered as 1.5 times the Pipe Dia of port 1.
    
    Dim oLineStringBotBody As IngrGeom3D.LineString3d
    Dim ObjBottomTaperBody   As Object

''Valve Body top surface Line points.
  
    dHexagonSide = pipeDiam1 / (2 * Sin(PI / 3))
    
    dLinePoints(0) = dHexagonSide / 2
    dLinePoints(1) = -parValveBodyHeight / 2
    dLinePoints(2) = dHexagonSide * Cos(PI / 6)
    
    dLinePoints(3) = dHexagonSide / 2
    dLinePoints(4) = dLinePoints(1)
    dLinePoints(5) = 0
            
    dLinePoints(6) = dLinePoints(0)
    dLinePoints(7) = dLinePoints(1)
    dLinePoints(8) = -dLinePoints(2)
    
    dLinePoints(9) = -dLinePoints(6)
    dLinePoints(10) = dLinePoints(7)
    dLinePoints(11) = dLinePoints(8)
    
    dLinePoints(12) = -dLinePoints(3)
    dLinePoints(13) = dLinePoints(4)
    dLinePoints(14) = dLinePoints(5)
    
    dLinePoints(15) = -dLinePoints(0)
    dLinePoints(16) = dLinePoints(1)
    dLinePoints(17) = dLinePoints(2)
   
    dLinePoints(18) = dLinePoints(0)
    dLinePoints(19) = dLinePoints(1)
    dLinePoints(20) = dLinePoints(2)
    
    Set oLineStringBotBody = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLinePoints)
    
    Set ObjBottomTaperBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, oStepLineString, _
                                                                 oLineStringBotBody, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBottomTaperBody
    Set ObjBottomTaperBody = Nothing

'   Insert your code for output 4(Nozzle 1 with length)
   
    Dim objNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oPlacementPoint As New AutoMath.DPosition
    
    oPlacementPoint.Set 0, parFace1toCenter + cptOffset1 - depth1, 0
    oDirVec.Set 0, 1, 0
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, _
                            oDirVec, oPlacementPoint, parFace1toCenter - parValveBodyHeight / 2)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

'   Insert your code for output 5(Nozzle 2 with length)
    oPlacementPoint.Set 0, -parFace2toCenter - cptOffset2 + depth2, 0
    
    oDirVec.Set 0, -1, 0
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, _
                            oDirVec, oPlacementPoint, parFace2toCenter - parValveBodyHeight / 2)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

'   Insert your code for output 6(Nozzle 3 with length)
    oPlacementPoint.Set (parFace3toCenter + cptOffset3 - depth3) * Cos(PI / 6), 0, (parFace3toCenter + cptOffset3 - depth3) * Sin(PI / 6)
                        
    oDirVec.Set Cos(PI / 6), 0, Sin(PI / 6)
    Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, _
                            oDirVec, oPlacementPoint, parFace3toCenter - dHexagonSide * Cos(PI / 6))
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
'  Insert your code for output 7(Nozzle 4 with length)
     
    oPlacementPoint.Set (parFace4toCenter + cptOffset4 - depth4) * Cos(PI / 6), 0, _
                        -(parFace4toCenter + cptOffset4 - depth4) * Sin(PI / 6)
    oDirVec.Set Cos(PI / 6), 0, -Sin(PI / 6)
    Set objNozzle = CreateNozzleWithLength(4, oPartFclt, m_OutputColl, _
                            oDirVec, oPlacementPoint, parFace4toCenter - dHexagonSide * Cos(PI / 6))
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Insert your code for output 8(Nozzle 5 with length)
    oPlacementPoint.Set 0, 0, -parFace5toCenter - cptOffset5 + depth5
                        
    oDirVec.Set 0, 0, -1
    Set objNozzle = CreateNozzleWithLength(5, oPartFclt, m_OutputColl, _
                            oDirVec, oPlacementPoint, parFace5toCenter - dHexagonSide * Cos(PI / 6))
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Insert your code for output 9(Nozzle 6 with length)
    oPlacementPoint.Set -(parFace6toCenter + cptOffset6 - depth6) * Cos(PI / 6), 0, _
                        -(parFace6toCenter + cptOffset6 - depth6) * Sin(PI / 6)
    oDirVec.Set -Cos(PI / 6), 0, -Sin(PI / 6)
    Set objNozzle = CreateNozzleWithLength(6, oPartFclt, m_OutputColl, _
                            oDirVec, oPlacementPoint, parFace6toCenter - dHexagonSide * Cos(PI / 6))
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
'   Insert your code for output 10(Nozzle 7 with length)
    oPlacementPoint.Set -(parFace7toCenter + cptOffset7 - depth7) * Cos(PI / 6), 0, _
                         (parFace7toCenter + cptOffset7 - depth7) * Sin(PI / 6)
                        
    oDirVec.Set -Cos(PI / 6), 0, Sin(PI / 6)
    Set objNozzle = CreateNozzleWithLength(7, oPartFclt, m_OutputColl, _
                            oDirVec, oPlacementPoint, parFace7toCenter - dHexagonSide * Cos(PI / 6))
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Insert your code for output 11(Nozzle 8 with length)
    oPlacementPoint.Set 0, 0, parFace8toCenter + cptOffset8 - depth8
                        
    oDirVec.Set 0, 0, 1
    Set objNozzle = CreateNozzleWithLength(8, oPartFclt, m_OutputColl, _
                            oDirVec, oPlacementPoint, parFace8toCenter - dHexagonSide * Cos(PI / 6))
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Insert your code for output 15(Valve Operator Body)
    Dim oDirX As IJDVector
    Dim oDirY As IJDVector
    Dim oDirZ As IJDVector
    Set oDirX = New DVector
    Set oDirY = New DVector
    Set oDirZ = New DVector

    Dim dOperatorAngle As Double

    dOperatorAngle = Atn(((parValveBodyHeight - dStepHeight) / 2) / ((parValveBodyWidth - dValveBodyTopWidth) / 2))

    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    On Error GoTo ErrorLabel

    Dim oOperatorPart As IJValveOperator
    Dim oMultiValCol As IJMultiPortValveAccCol
    Dim oOperatorOcc   As IJPartOcc

    If Not oPipeComponent Is Nothing Then
      oPipeComponent.GetValveOperatorsForPorts oMultiValCol
    End If

    Dim OpOrigin As IJDPosition
    Set OpOrigin = New DPosition
    OpOrigin.Set 0, 0, 0

'Operator for port 3
    oDirX.Set Cos(PI / 4) * Cos(PI / 6), -Sin(PI / 4), Cos(PI / 4) * Sin(PI / 6)
    oDirY.Set Cos(PI / 4) * Cos(PI / 6), Sin(PI / 4), Cos(PI / 4) * Sin(PI / 6)
    Set oDirZ = oDirX.Cross(oDirY)
      
      Set oOperatorPart = oMultiValCol.GetValveOperatorPartPerPort(3)
            If Not oOperatorPart Is Nothing Then
                Set oOperatorOcc = m_oGeomHelper.CreateChildPartOcc("ValveOperator1", _
                                        oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)

            End If
    Set oOperatorOcc = Nothing
    Set oOperatorPart = Nothing

'Operator for port 4
    oDirX.Set Cos(PI / 4) * Cos((5 * PI) / 6), -Sin(PI / 4), Cos(PI / 4) * Sin((5 * PI) / 6)
    oDirY.Set Cos(PI / 4) * Cos((5 * PI) / 6), Sin(PI / 4), Cos(PI / 4) * Sin((5 * PI) / 6)
    Set oDirZ = oDirX.Cross(oDirY)

    Set oOperatorPart = oMultiValCol.GetValveOperatorPartPerPort(4)
            If Not oOperatorPart Is Nothing Then
                Set oOperatorOcc = m_oGeomHelper.CreateChildPartOcc("ValveOperator2", _
                                        oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)

            End If

    Set oOperatorOcc = Nothing
    Set oOperatorPart = Nothing

'Operator for port 5
    oDirX.Set Cos(PI / 4) * Cos((3 * PI) / 2), -Sin(PI / 4), Cos(PI / 4) * Sin((3 * PI) / 2)
    oDirY.Set Cos(PI / 4) * Cos((3 * PI) / 2), Sin(PI / 4), Cos(PI / 4) * Sin((3 * PI) / 2)
    Set oDirZ = oDirX.Cross(oDirY)

      Set oOperatorPart = oMultiValCol.GetValveOperatorPartPerPort(5)
            If Not oOperatorPart Is Nothing Then
                Set oOperatorOcc = m_oGeomHelper.CreateChildPartOcc("ValveOperator3", _
                                        oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)

            End If
    Set oOperatorOcc = Nothing
    Set oOperatorPart = Nothing

'Operator for port 6
    oDirY.Set Cos(PI / 4) * Cos(PI / 2), -Sin(PI / 4), Cos(PI / 4) * Sin(PI / 2)
    oDirX.Set Cos(PI / 4) * Cos(PI / 2), Sin(PI / 4), Cos(PI / 4) * Sin(PI / 2)
    Set oDirZ = oDirX.Cross(oDirY)
    
    Set oOperatorPart = oMultiValCol.GetValveOperatorPartPerPort(6)
            If Not oOperatorPart Is Nothing Then
                Set oOperatorOcc = m_oGeomHelper.CreateChildPartOcc("ValveOperator4", _
                                        oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)

            End If

'Operator for port 7
    oDirY.Set Cos(PI / 4) * Cos((7 * PI) / 6), -Sin(PI / 4), Cos(PI / 4) * Sin((7 * PI) / 6)
    oDirX.Set Cos(PI / 4) * Cos((7 * PI) / 6), Sin(PI / 4), Cos(PI / 4) * Sin((7 * PI) / 6)
    Set oDirZ = oDirX.Cross(oDirY)

      Set oOperatorPart = oMultiValCol.GetValveOperatorPartPerPort(7)
            If Not oOperatorPart Is Nothing Then
                Set oOperatorOcc = m_oGeomHelper.CreateChildPartOcc("ValveOperator5", _
                                        oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)

            End If
    Set oOperatorOcc = Nothing
    Set oOperatorPart = Nothing

'Operator for port 8
    oDirY.Set Cos(PI / 4) * Cos((11 * PI) / 6), -Sin(PI / 4), Cos(PI / 4) * Sin((11 * PI) / 6)
    oDirX.Set Cos(PI / 4) * Cos((11 * PI) / 6), Sin(PI / 4), Cos(PI / 4) * Sin((11 * PI) / 6)
    Set oDirZ = oDirX.Cross(oDirY)
    
    Set oOperatorPart = oMultiValCol.GetValveOperatorPartPerPort(8)
            If Not oOperatorPart Is Nothing Then
                Set oOperatorOcc = m_oGeomHelper.CreateChildPartOcc("ValveOperator6", _
                                        oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)

            End If
            
    Set oOperatorOcc = Nothing
    Set oOperatorPart = Nothing

    Set oPlacementPoint = Nothing
    Set oDirVec = Nothing
    Set oDirX = Nothing
    Set oDirY = Nothing
    Set oDirZ = Nothing

    Set oGeomFactory = Nothing
    Set OpOrigin = Nothing
    
     
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Sub
   

Private Function Max6(x1 As Double, x2 As Double, x3 As Double, x4 As Double, x5 As Double, x6 As Double)

    Dim dmax As Double
    dmax = x1
    If dmax < x2 Then dmax = x2
    If dmax < x3 Then dmax = x3
    If dmax < x4 Then dmax = x4
    If dmax < x5 Then dmax = x5
    If dmax < x6 Then dmax = x6
    
    Max6 = dmax

End Function
