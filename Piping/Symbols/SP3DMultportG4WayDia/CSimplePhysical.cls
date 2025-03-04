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
'   Author:         ACM
'   Creation Date:  Monday, July 27, 2004
'   Description:
'       This is a 2-way diverver valve symbol. This is prepared based on Gemu's catalog.
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
    Dim flangeThick     As Double
    Dim flangeDiam      As Double
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
    
    Dim iOutput     As Double
    
    Dim parValveBodyWidth As Double
    Dim parValveBodyDepth As Double
    Dim parValveBodyHeight As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parFace4toCenter As Double
    Dim parFace5toCenter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parValveBodyWidth = arrayOfInputs(2)
    parValveBodyDepth = arrayOfInputs(3)
    parValveBodyHeight = arrayOfInputs(4)
    parFace1toCenter = arrayOfInputs(5)
    parFace2toCenter = arrayOfInputs(6)
    parFace3toCenter = arrayOfInputs(7)
    parFace4toCenter = arrayOfInputs(8)
    parFace5toCenter = arrayOfInputs(9)
    parInsulationThickness = arrayOfInputs(10)
    
    iOutput = 0
    
    'Origin is taken at the geometry center of the valve.
    
    m_oGeomHelper.OutputCollection = m_OutputColl
    
'   Insert your code for output 1(Valve Body)
    Dim ObjValveBody As Object
    
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    Dim oTopLineString As IngrGeom3D.LineString3d
    Dim oBottomLineString As IngrGeom3D.LineString3d
    Dim dLinePoints(0 To 14) As Double
    Dim dStepHeight As Double
    Dim oDirVec As New AutoMath.DVector
    
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, _
                        flangeThick, flangeDiam, cptOffset2, depth2
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, _
                        flangeThick, flangeDiam, cptOffset3, depth3
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, _
                        flangeThick, flangeDiam, cptOffset4, depth4
    RetrieveParameters 5, oPartFclt, m_OutputColl, pipeDiam5, _
                        flangeThick, flangeDiam, cptOffset5, depth5
     
 ''Step height is assumed to be 1.5 times the biggest Pipe diameter of Port2,Port3,Port4,Port5
    dStepHeight = 1.5 * Max5(pipeDiam1, pipeDiam2, pipeDiam3, pipeDiam4, pipeDiam5)
    

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, _
                        flangeThick, flangeDiam, cptOffset1, depth1
    Dim dValveBodyTopWidth As Double
    ''Width of the Body Top is assumed to be 1.5 times the Pipe diameter of Nozzle 1.
    ''Top surface is assume to be a Squre.
    dValveBodyTopWidth = 1.5 * pipeDiam1
    
    ''Valve Body top surface Line points.
    dLinePoints(0) = dValveBodyTopWidth / 2
    dLinePoints(1) = parValveBodyHeight / 2
    dLinePoints(2) = dValveBodyTopWidth / 2
    
    dLinePoints(3) = dLinePoints(0)
    dLinePoints(4) = dLinePoints(1)
    dLinePoints(5) = -dLinePoints(2)
            
    dLinePoints(6) = -dLinePoints(3)
    dLinePoints(7) = dLinePoints(4)
    dLinePoints(8) = dLinePoints(5)
    
    dLinePoints(9) = dLinePoints(6)
    dLinePoints(10) = dLinePoints(7)
    dLinePoints(11) = dLinePoints(2)
    
    dLinePoints(12) = dValveBodyTopWidth / 2
    dLinePoints(13) = parValveBodyHeight / 2
    dLinePoints(14) = dValveBodyTopWidth / 2
    
    Set oTopLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLinePoints)
            
 ''Valve Body Bottom surface Line points.
    dLinePoints(0) = parValveBodyWidth / 2
    dLinePoints(1) = dStepHeight - parValveBodyHeight / 2
    dLinePoints(2) = parValveBodyDepth / 2
    
    dLinePoints(3) = dLinePoints(0)
    dLinePoints(4) = dLinePoints(1)
    dLinePoints(5) = -dLinePoints(2)
            
    dLinePoints(6) = -dLinePoints(3)
    dLinePoints(7) = dLinePoints(4)
    dLinePoints(8) = dLinePoints(5)
    
    dLinePoints(9) = dLinePoints(6)
    dLinePoints(10) = dLinePoints(7)
    dLinePoints(11) = dLinePoints(2)
    
    dLinePoints(12) = parValveBodyWidth / 2
    dLinePoints(13) = dStepHeight - parValveBodyHeight / 2
    dLinePoints(14) = parValveBodyDepth / 2
        
    Set oBottomLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLinePoints)
    
    Set ObjValveBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, oTopLineString, _
                                                                   oBottomLineString, True)
            
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBody
    Set ObjValveBody = Nothing

'   Insert your code for output 2(Stepped part of valve body)
    Dim ObjValveBodyStep As Object
    Dim oStartPoint As New AutoMath.DPosition
    Dim oEndPoint As New AutoMath.DPosition
    
    oStartPoint.Set dLinePoints(0), dLinePoints(1), dLinePoints(2)
    oEndPoint.Set dLinePoints(6), dLinePoints(7) - dStepHeight, dLinePoints(8)
    
    Set ObjValveBodyStep = PlaceBox(m_OutputColl, oStartPoint, oEndPoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBodyStep
    Set ObjValveBodyStep = Nothing

'   Insert your code for output 3(Nozzle 1 with length)
   
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

'   Insert your code for output 4(Nozzle 2 with length)
    oPlacementPoint.Set parFace2toCenter + cptOffset2 - depth2, -parValveBodyHeight / 2 + dStepHeight / 2, 0
    
    oDirVec.Set 1, 0, 0
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, _
                            oDirVec, oPlacementPoint, parFace2toCenter - parValveBodyWidth / 2)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

'   Insert your code for output 5(Nozzle 3 with length)
    oPlacementPoint.Set 0, -parValveBodyHeight / 2 + dStepHeight / 2, _
                        -parFace3toCenter - cptOffset3 + depth3
    oDirVec.Set 0, 0, -1
    Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, _
                            oDirVec, oPlacementPoint, parFace3toCenter - parValveBodyDepth / 2)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
'  Insert your code for output 6(Nozzle 4 with length)
    oPlacementPoint.Set -parFace4toCenter - cptOffset4 + depth4, -parValveBodyHeight / 2 + dStepHeight / 2, 0
    oDirVec.Set -1, 0, 0
    Set objNozzle = CreateNozzleWithLength(4, oPartFclt, m_OutputColl, _
                            oDirVec, oPlacementPoint, parFace4toCenter - parValveBodyWidth / 2)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Insert your code for output 7(Nozzle 5 with length)
    oPlacementPoint.Set 0, -parValveBodyHeight / 2 + dStepHeight / 2, _
                        parFace5toCenter + cptOffset5 - depth5
    oDirVec.Set 0, 0, 1
    Set objNozzle = CreateNozzleWithLength(5, oPartFclt, m_OutputColl, _
                            oDirVec, oPlacementPoint, parFace5toCenter - parValveBodyDepth / 2)
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

    dOperatorAngle = Atn((parValveBodyHeight - dStepHeight) / ((parValveBodyWidth - dValveBodyTopWidth) / 2))

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

''Operator for port 2
    oDirX.Set Cos(dOperatorAngle), -Sin(dOperatorAngle), 0
    oDirY.Set Sin(dOperatorAngle), Cos(dOperatorAngle), 0
    oDirZ.Set 0, 0, 1  ''Operator local Z-axis matches with Valve Z-axis

    Set oOperatorPart = oMultiValCol.GetValveOperatorPartPerPort(2)
          If Not oOperatorPart Is Nothing Then
              Set oOperatorOcc = m_oGeomHelper.CreateChildPartOcc("ValveOperator1", _
                                      oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)

          End If
    Set oOperatorOcc = Nothing
    Set oOperatorPart = Nothing

''Operator for port 3
    oDirX.Set 0, -Sin(dOperatorAngle), -Cos(dOperatorAngle)
    oDirY.Set 0, Cos(dOperatorAngle), -Sin(dOperatorAngle)
    oDirZ.Set 1, 0, 0  ''Operator local Z-axis matches with Valve X-axis

    Set oOperatorPart = oMultiValCol.GetValveOperatorPartPerPort(3)
        If Not oOperatorPart Is Nothing Then
            Set oOperatorOcc = m_oGeomHelper.CreateChildPartOcc("ValveOperator2", _
                                    oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)
        End If

    Set oOperatorOcc = Nothing
    Set oOperatorPart = Nothing

''Operator for port 4
    oDirX.Set -Cos(dOperatorAngle), -Sin(dOperatorAngle), 0
    oDirY.Set -Sin(dOperatorAngle), Cos(dOperatorAngle), 0
    oDirZ.Set 0, 0, 1  ''Operator local Z-axis matches with Valve Z-axis

    Set oOperatorPart = oMultiValCol.GetValveOperatorPartPerPort(2)
          If Not oOperatorPart Is Nothing Then
              Set oOperatorOcc = m_oGeomHelper.CreateChildPartOcc("ValveOperator3", _
                                      oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)
          End If
    Set oOperatorOcc = Nothing
    Set oOperatorPart = Nothing

''Operator for port 5
    oDirX.Set 0, -Sin(dOperatorAngle), Cos(dOperatorAngle)
    oDirY.Set 0, Cos(dOperatorAngle), Sin(dOperatorAngle)
    oDirZ.Set 1, 0, 0  ''Operator local Z-axis matches with Valve X-axis

    Set oOperatorPart = oMultiValCol.GetValveOperatorPartPerPort(3)
        If Not oOperatorPart Is Nothing Then
            Set oOperatorOcc = m_oGeomHelper.CreateChildPartOcc("ValveOperator4", _
                                    oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)

        End If

    Set oOperatorOcc = Nothing
    Set oOperatorPart = Nothing

    Set oPlacementPoint = Nothing
    Set oDirVec = Nothing
    Set oDirX = Nothing
    Set oDirY = Nothing
    Set oDirZ = Nothing

    Set oStartPoint = Nothing
    Set oEndPoint = Nothing
    Set oGeomFactory = Nothing
    Set OpOrigin = Nothing
    
     
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Sub
   

Public Function Max5(x1 As Double, x2 As Double, x3 As Double, x4 As Double, x5 As Double)
    Dim dmax As Double
    dmax = x1
    If dmax < x2 Then dmax = x2
    If dmax < x3 Then dmax = x3
    If dmax < x4 Then dmax = x4
    If dmax < x5 Then dmax = x5
    
    Max5 = dmax

End Function
