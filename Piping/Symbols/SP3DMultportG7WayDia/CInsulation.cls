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
'   Author:         ACM/MS
'   Creation Date:  Monday, July 12, 2004
'   Description:
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_oGeomHelper As IJSymbolGeometryHelper
Dim PI As Double
Private Const MODULE = "Insulation:" 'Used for error messages
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
    
    Dim oStartPoint As New AutoMath.DPosition
    Dim oEndPoint As New AutoMath.DPosition
        
    m_oGeomHelper.OutputCollection = m_OutputColl
    
'   Insert your code for output 18(Top Taper Body Insulation)
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
    dHexagonSide = pipeDiam1 / (2 * Sin(PI / 3)) + parInsulationThickness
    
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
    
    dLinePoints(0) = parInsulationThickness + dHexagonSide / 2
    dLinePoints(1) = dStepHeight / 2
    dLinePoints(2) = parInsulationThickness + dHexagonSide * Cos(PI / 6)
    
    dLinePoints(3) = parInsulationThickness + parValveBodyWidth / 2
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
    
    Set ObjValveBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, oTopLineString, _
                                                                   oBottomLineString, True)
            
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBody
    Set ObjValveBody = Nothing
    Set oTopLineString = Nothing
    
'   Insert your code for output 19(Stepped part of valve body Insulation)
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

' Insert your code for output 20(Bottom Taper part of Valve Body Insulation)
     RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, _
                        flangeThick, flangeDiam2, cptOffset2, depth2
    
    dValveBodyTopWidth = pipeDiam2 * 1.5 ''Side of the Triaglur surface is consiered as 1.5 times the Pipe Dia of port 1.
    
    Dim oLineStringBotBody As IngrGeom3D.LineString3d
    Dim ObjBottomTaperBody   As Object

'Valve Body Triangular Bottom surface Line points.
    
    dHexagonSide = pipeDiam1 / (2 * Sin(PI / 3)) + parInsulationThickness
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
    
'  Insert your code for output 21(Nozzle 1 Insulation)
    Dim dInsulationDia As Double
  
    If pipeDiam1 < flangeDiam1 Then
        dInsulationDia = flangeDiam1 + 2 * parInsulationThickness
    Else
        dInsulationDia = pipeDiam1 + 2 * parInsulationThickness
    End If
    
    oStartPoint.Set 0, parValveBodyHeight / 2, 0
    oEndPoint.Set 0, parFace1toCenter, 0
   ' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia

'   Insert your code for output 22(Nozzle 2 Insulation)
    If pipeDiam2 < flangeDiam2 Then
        dInsulationDia = flangeDiam2 + 2 * parInsulationThickness
    Else
        dInsulationDia = pipeDiam2 + 2 * parInsulationThickness
    End If
    
    oStartPoint.Set 0, -parValveBodyHeight / 2, 0
    oEndPoint.Set 0, -parFace2toCenter, 0
'   Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia

'   Insert your code for output 23(Nozzle 3 Insulation)
    If pipeDiam3 < flangeDiam3 Then
        dInsulationDia = flangeDiam3 + 2 * parInsulationThickness
    Else
        dInsulationDia = pipeDiam3 + 2 * parInsulationThickness
    End If
    
    oStartPoint.Set dHexagonSide * Cos(PI / 6), 0, dHexagonSide * Sin(PI / 6)
    oEndPoint.Set parFace3toCenter * Cos(PI / 6), 0, parFace3toCenter * Sin(PI / 6)
'   Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia
    
'  Insert your code for output 24(Nozzle 4 Insulation)
    If pipeDiam4 < flangeDiam4 Then
        dInsulationDia = flangeDiam4 + 2 * parInsulationThickness
    Else
        dInsulationDia = pipeDiam4 + 2 * parInsulationThickness
    End If
    
    oStartPoint.Set dHexagonSide * Cos(PI / 6), 0, -dHexagonSide * Sin(PI / 6)
    oEndPoint.Set parFace4toCenter * Cos(PI / 6), 0, -parFace4toCenter * Sin(PI / 6)
'   Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia
         
' Insert your code for output 25(Nozzle 5 Insulation)
    If pipeDiam5 < flangeDiam5 Then
        dInsulationDia = flangeDiam5 + 2 * parInsulationThickness
    Else
        dInsulationDia = pipeDiam5 + 2 * parInsulationThickness
    End If
    
    oStartPoint.Set 0, 0, -dHexagonSide * Sin(PI / 6)
    oEndPoint.Set 0, 0, -parFace5toCenter
'   Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia
       
' Insert your code for output 26(Nozzle 6 Insulation)
    If pipeDiam6 < flangeDiam6 Then
        dInsulationDia = flangeDiam6 + 2 * parInsulationThickness
    Else
        dInsulationDia = pipeDiam6 + 2 * parInsulationThickness
    End If
    
    oStartPoint.Set -dHexagonSide * Cos(PI / 6), 0, -dHexagonSide * Sin(PI / 6)
    oEndPoint.Set -parFace6toCenter * Cos(PI / 6), 0, -parFace6toCenter * Sin(PI / 6)
'   Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia
    
'   Insert your code for output 27(Nozzle 7 Insulation)
    If pipeDiam7 < flangeDiam7 Then
        dInsulationDia = flangeDiam7 + 2 * parInsulationThickness
    Else
        dInsulationDia = pipeDiam7 + 2 * parInsulationThickness
    End If
    
    oStartPoint.Set -dHexagonSide * Cos(PI / 6), 0, dHexagonSide * Sin(PI / 6)
    oEndPoint.Set -parFace7toCenter * Cos(PI / 6), 0, parFace7toCenter * Sin(PI / 6)
'   Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia
       
' Insert your code for output 28(Nozzle 8 Insulation)
    If pipeDiam8 < flangeDiam8 Then
        dInsulationDia = flangeDiam8 + 2 * parInsulationThickness
    Else
        dInsulationDia = pipeDiam8 + 2 * parInsulationThickness
    End If
    
    oStartPoint.Set 0, 0, dHexagonSide * Sin(PI / 6)
    oEndPoint.Set 0, 0, parFace8toCenter
'   Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia
    
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

