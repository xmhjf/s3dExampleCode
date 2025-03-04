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
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'   All Rights Reserved
'
'   CInsulation.cls
'   Author:         ACM
'   Creation Date:  Monday, July 12, 2004
'   Description:
'       This is a multi port diverver valve symbol. This is prepared based on Saunder's catalog.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_oGeomHelper As IJSymbolGeometryHelper

Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private Sub Class_Terminate()
    Set m_oGeomHelper = Nothing
End Sub

Private Sub Class_Initialize()
    Set m_oGeomHelper = New SymbolServices
    
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
    Dim flangeDiam1      As Double
    Dim flangeDiam2      As Double
    Dim flangeDiam3      As Double
    Dim flangeDiam4      As Double
    Dim flangeDiam5      As Double
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
    
'   Insert your code for output 12(Insulated Valve Body)
    Dim ObjValveBody As Object
    
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    Dim oTopLineString As IngrGeom3D.LineString3d
    Dim oBottomLineString As IngrGeom3D.LineString3d
    Dim dLinePoints(0 To 14) As Double
    Dim dStepHeight As Double
    Dim oDirVec As New AutoMath.DVector
    
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, _
                        flangeThick, flangeDiam2, cptOffset2, depth2
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, _
                        flangeThick, flangeDiam3, cptOffset3, depth3
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, _
                        flangeThick, flangeDiam4, cptOffset4, depth4
    RetrieveParameters 5, oPartFclt, m_OutputColl, pipeDiam5, _
                        flangeThick, flangeDiam5, cptOffset5, depth5
     
 ''Step height is assumed to be 1.5 times the biggest Pipe diameter of Port2,Port3,Port4,Port5
    dStepHeight = 1.5 * Max5(pipeDiam1, pipeDiam2, pipeDiam3, pipeDiam4, pipeDiam5)
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, _
                        flangeThick, flangeDiam1, cptOffset1, depth1
    Dim dValveBodyTopWidth As Double
    ''Width of the Body Top is assumed to be 1.5 times the Pipe diameter of Nozzle 1.
    ''Top surface is assume to be a Squre.
    dValveBodyTopWidth = 1.5 * pipeDiam1
    
    ''Valve Body top surface Line points.
    dLinePoints(0) = parInsulationThickness + dValveBodyTopWidth / 2
    dLinePoints(1) = parInsulationThickness + parValveBodyHeight / 2
    dLinePoints(2) = parInsulationThickness + dValveBodyTopWidth / 2
    
    dLinePoints(3) = dLinePoints(0)
    dLinePoints(4) = dLinePoints(1)
    dLinePoints(5) = -dLinePoints(2)
            
    dLinePoints(6) = -dLinePoints(3)
    dLinePoints(7) = dLinePoints(4)
    dLinePoints(8) = dLinePoints(5)
    
    dLinePoints(9) = dLinePoints(6)
    dLinePoints(10) = dLinePoints(7)
    dLinePoints(11) = dLinePoints(2)
    
    dLinePoints(12) = parInsulationThickness + dValveBodyTopWidth / 2
    dLinePoints(13) = parInsulationThickness + parValveBodyHeight / 2
    dLinePoints(14) = parInsulationThickness + dValveBodyTopWidth / 2
    
    Set oTopLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLinePoints)
            
 ''Valve Body Bottom surface Line points.
    dLinePoints(0) = parInsulationThickness + parValveBodyWidth / 2
    dLinePoints(1) = dStepHeight - parValveBodyHeight / 2
    dLinePoints(2) = parInsulationThickness + parValveBodyDepth / 2
    
    dLinePoints(3) = dLinePoints(0)
    dLinePoints(4) = dLinePoints(1)
    dLinePoints(5) = -dLinePoints(2)
            
    dLinePoints(6) = -dLinePoints(3)
    dLinePoints(7) = dLinePoints(4)
    dLinePoints(8) = dLinePoints(5)
    
    dLinePoints(9) = dLinePoints(6)
    dLinePoints(10) = dLinePoints(7)
    dLinePoints(11) = dLinePoints(2)
    dLinePoints(12) = parInsulationThickness + parValveBodyWidth / 2
    dLinePoints(13) = dStepHeight - parValveBodyHeight / 2
    dLinePoints(14) = parInsulationThickness + parValveBodyDepth / 2
    
    Set oBottomLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLinePoints)
    
    Set ObjValveBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, oTopLineString, _
                                                                   oBottomLineString, True)
            
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBody
    Set ObjValveBody = Nothing

'   Insert your code for output 13(Stepped part of valve body Insualtion)
    Dim ObjValveBodyStep As Object
    Dim oStartPoint As New AutoMath.DPosition
    Dim oEndPoint As New AutoMath.DPosition
    
    oStartPoint.Set dLinePoints(0), dLinePoints(1), dLinePoints(2)
    oEndPoint.Set dLinePoints(6), -parInsulationThickness - parValveBodyHeight / 2, dLinePoints(8)
    
    Set ObjValveBodyStep = PlaceBox(m_OutputColl, oStartPoint, oEndPoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBodyStep
    Set ObjValveBodyStep = Nothing

'   Insert your code for output 14(Nozzle 1 insulation)
    Dim dInsulationDia As Double
  
    If pipeDiam1 < flangeDiam1 Then
        dInsulationDia = flangeDiam1 + 2 * parInsulationThickness
    Else
        dInsulationDia = pipeDiam1 + 2 * parInsulationThickness
    End If
    
    oStartPoint.Set 0, parFace1toCenter - parValveBodyHeight / 2, 0
    oEndPoint.Set 0, parFace1toCenter, 0
   ' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia

'   Insert your code for output 15(Nozzle 2 Insualation)
   
    If pipeDiam2 < flangeDiam2 Then
        dInsulationDia = flangeDiam2 + 2 * parInsulationThickness
    Else
        dInsulationDia = pipeDiam2 + 2 * parInsulationThickness
    End If
    
    oStartPoint.Set parFace2toCenter - parValveBodyWidth / 2, -parValveBodyHeight / 2 + dStepHeight / 2, 0
    oEndPoint.Set parFace2toCenter, -parValveBodyHeight / 2 + dStepHeight / 2, 0
   ' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia

'   Insert your code for output 16(Nozzle 3 Insualation)
   
    If pipeDiam3 < flangeDiam3 Then
        dInsulationDia = flangeDiam3 + 2 * parInsulationThickness
    Else
        dInsulationDia = pipeDiam3 + 2 * parInsulationThickness
    End If
    
    oStartPoint.Set 0, -parValveBodyHeight / 2 + dStepHeight / 2, -parFace3toCenter + parValveBodyDepth / 2
    oEndPoint.Set 0, -parValveBodyHeight / 2 + dStepHeight / 2, -parFace3toCenter
   ' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia
 
'   Insert your code for output 17(Nozzle 4 Insualation)
  
    If pipeDiam4 < flangeDiam4 Then
        dInsulationDia = flangeDiam4 + 2 * parInsulationThickness
    Else
        dInsulationDia = pipeDiam4 + 2 * parInsulationThickness
    End If
    
    oStartPoint.Set -parFace4toCenter + parValveBodyWidth / 2, -parValveBodyHeight / 2 + dStepHeight / 2, 0
    oEndPoint.Set -parFace4toCenter, -parValveBodyHeight / 2 + dStepHeight / 2, 0
   ' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia


'   Insert your code for output 18(Nozzle 5 Insualation)
    If pipeDiam5 < flangeDiam5 Then
        dInsulationDia = flangeDiam5 + 2 * parInsulationThickness
    Else
        dInsulationDia = pipeDiam5 + 2 * parInsulationThickness
    End If
    
    oStartPoint.Set 0, -parValveBodyHeight / 2 + dStepHeight / 2, parFace5toCenter - parValveBodyDepth / 2
    oEndPoint.Set 0, -parValveBodyHeight / 2 + dStepHeight / 2, parFace5toCenter
   ' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia
 
 Set oStartPoint = Nothing
 Set oEndPoint = Nothing
 
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


