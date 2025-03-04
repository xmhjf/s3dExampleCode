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
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 1/13/2005-6:25:24 PM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         MS
'   Creation Date:  Wednesday 6th October, 2004
'   Description:
'    The symbol consists of a box, conduit port.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   11.JAN.2005     MS     CR69200 Added cable ports at the same location of conduit ports.
'   13.Jan.2005     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  22.Apr.2008     VRK           CR-120097  Need to set proper units to the cable nozzle diameter and area
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "SimplePhysical:" 'Used for error messages

Dim m_oGeomHelper As IJSymbolGeometryHelper

Private m_oCodeListMetadata     As IJDCodeListMetaData


Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
     Set m_oGeomHelper = New SymbolServices
    
    Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

Private Sub Class_Terminate()
    Set m_oGeomHelper = Nothing
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput     As Double

    Dim parElecEquipmentHeight As Double
    Dim parElecEquipmentWidth As Double
    Dim parElecEquipmentLength As Double
    
    Dim ObjConduitPort1 As IJConduitPortOcc
    Dim ObjConduitPort2 As IJConduitPortOcc
    Dim ObjConduitPort3 As IJConduitPortOcc
    Dim ObjConduitPort4 As IJConduitPortOcc
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parElecEquipmentHeight = arrayOfInputs(2)
    parElecEquipmentWidth = arrayOfInputs(3)
    parElecEquipmentLength = arrayOfInputs(4)

    iOutput = 0
    m_oGeomHelper.OutputCollection = m_OutputColl
        
' Insert your code for output 1(Base of Transformer)
    Dim objTransformerBase As Object
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition

    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    stPoint.Set -parElecEquipmentWidth / 2, -parElecEquipmentLength / 2, 0
    enPoint.Set parElecEquipmentWidth / 2, parElecEquipmentLength / 2, 0.05 * parElecEquipmentHeight
    
    Set objTransformerBase = PlaceBox(m_OutputColl, stPoint, enPoint)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTransformerBase
    Set objTransformerBase = Nothing
    
    
' Insert your code for output 2(Body of Transformer)
    Dim objTransformerBody As Object
    stPoint.Set -0.8 * parElecEquipmentWidth / 2, -parElecEquipmentLength / 2, 0.05 * parElecEquipmentHeight
    enPoint.Set 0.8 * parElecEquipmentWidth / 2, parElecEquipmentLength / 2, 0.9 * parElecEquipmentHeight
    
    Set objTransformerBody = PlaceBox(m_OutputColl, stPoint, enPoint)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTransformerBody
    Set objTransformerBody = Nothing
    
    
' Insert your code for output 3(Top of Transformer)
    Dim objTransformerTop As Object
    stPoint.Set -parElecEquipmentWidth / 2, -parElecEquipmentLength / 2, 0.9 * parElecEquipmentHeight
    enPoint.Set parElecEquipmentWidth / 2, parElecEquipmentLength / 2, parElecEquipmentHeight
    
    Set objTransformerTop = PlaceBox(m_OutputColl, stPoint, enPoint)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTransformerTop
    Set objTransformerTop = Nothing
    
    
' Insert your code for output 4(Transformer Coil)
    Dim objTransformerCoil As Object
    Dim arrPt(1 To 3)     As New AutoMath.DPosition
    Dim centerPoint As New AutoMath.DPosition
    Dim oLine           As IngrGeom3D.Line3d
    Dim oArc As IngrGeom3D.Arc3d
    Dim oNormalVector   As IJDVector
    Dim oLinesArcCollection          As Collection
    Dim oComplexString     As IngrGeom3D.ComplexString3d
    
    Set oNormalVector = New AutoMath.DVector
    Set oLinesArcCollection = New Collection
    
    arrPt(1).Set (parElecEquipmentWidth) / 2, _
                    (0.6 * parElecEquipmentLength) / 2 + (0.1 * parElecEquipmentLength), _
                        0.815 * parElecEquipmentHeight
    arrPt(2).Set parElecEquipmentWidth / 2 + (0.1 * parElecEquipmentLength), _
                    (0.6 * parElecEquipmentLength) / 2, _
                        0.815 * parElecEquipmentHeight
    arrPt(3).Set (parElecEquipmentWidth) / 2, _
                    (0.6 * parElecEquipmentLength) / 2, _
                        0.815 * parElecEquipmentHeight
    
    
    ' Arc 1
    centerPoint.Set arrPt(3).x, arrPt(3).y, arrPt(3).z
    stPoint.Set arrPt(1).x, arrPt(1).y, arrPt(1).z
    enPoint.Set arrPt(2).x, arrPt(2).y, arrPt(2).z
    
    oNormalVector.Set 0, 0, -1
    Set oArc = PlaceArcByCen(Nothing, centerPoint, stPoint, enPoint, oNormalVector)
    
    oLinesArcCollection.Add oArc
    
    ' Line 1
    stPoint.Set arrPt(2).x, arrPt(2).y, arrPt(2).z
    enPoint.Set arrPt(2).x, -arrPt(2).y, arrPt(2).z

    Set oLine = PlaceLine(Nothing, stPoint, enPoint)
    oLinesArcCollection.Add oLine

    ' Arc 2
    centerPoint.Set arrPt(3).x, -arrPt(3).y, arrPt(3).z
    stPoint.Set arrPt(2).x, -arrPt(2).y, arrPt(2).z
    enPoint.Set arrPt(1).x, -arrPt(1).y, arrPt(1).z

    Set oArc = PlaceArcByCen(Nothing, centerPoint, stPoint, enPoint, oNormalVector)
    oLinesArcCollection.Add oArc

    ' Line 2
    stPoint.Set arrPt(1).x, -arrPt(1).y, arrPt(1).z
    enPoint.Set -arrPt(1).x, -arrPt(1).y, arrPt(1).z

    Set oLine = PlaceLine(Nothing, stPoint, enPoint)
    oLinesArcCollection.Add oLine

    ' Arc 3
    centerPoint.Set -arrPt(3).x, -arrPt(3).y, arrPt(3).z
    stPoint.Set -arrPt(1).x, -arrPt(1).y, arrPt(1).z
    enPoint.Set -arrPt(2).x, -arrPt(2).y, arrPt(2).z

    Set oArc = PlaceArcByCen(Nothing, centerPoint, stPoint, enPoint, oNormalVector)
    oLinesArcCollection.Add oArc

    ' Line 3
    stPoint.Set -arrPt(2).x, -arrPt(2).y, arrPt(2).z
    enPoint.Set -arrPt(2).x, arrPt(2).y, arrPt(2).z

    Set oLine = PlaceLine(Nothing, stPoint, enPoint)
    oLinesArcCollection.Add oLine

    ' Arc 4
    centerPoint.Set -arrPt(3).x, arrPt(3).y, arrPt(3).z
    stPoint.Set -arrPt(2).x, arrPt(2).y, arrPt(2).z
    enPoint.Set -arrPt(1).x, arrPt(1).y, arrPt(1).z

    Set oArc = PlaceArcByCen(Nothing, centerPoint, stPoint, enPoint, oNormalVector)
    oLinesArcCollection.Add oArc

    ' Line 4
    stPoint.Set -arrPt(1).x, arrPt(1).y, arrPt(1).z
    enPoint.Set arrPt(1).x, arrPt(1).y, arrPt(1).z

    Set oLine = PlaceLine(Nothing, stPoint, enPoint)
    oLinesArcCollection.Add oLine
    
    Set oComplexString = PlaceTrCString(arrPt(1), oLinesArcCollection)
    
    Dim axis As New AutoMath.DVector
    axis.Set 0, 0, -1
    
    Set objTransformerCoil = PlaceProjection(m_OutputColl, oComplexString, axis, 0.68 * parElecEquipmentHeight, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objTransformerCoil
    Set objTransformerCoil = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set arrPt(1) = Nothing
    Set arrPt(2) = Nothing
    Set arrPt(3) = Nothing
    Set centerPoint = Nothing
    Set oNormalVector = Nothing
    Set axis = Nothing
    
' Insert your code for output 5(Conduit Port 1)
    Dim pConduitPort As New AutoMath.DPosition
    Dim oDir As New AutoMath.DVector
    oDir.Set 0, 0, 1
    
    pConduitPort.Set 0.8 * parElecEquipmentWidth / 2, -0.3 * parElecEquipmentLength, parElecEquipmentHeight
    Set ObjConduitPort1 = CreateConduitNozzlePH(pConduitPort, oDir, m_OutputColl, oPartFclt, 1)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort1
    Set ObjConduitPort1 = Nothing
    
' Insert your code for output 6(Conduit Port 2)
    pConduitPort.Set 0.8 * parElecEquipmentWidth / 2, -0.1 * parElecEquipmentLength, parElecEquipmentHeight
    Set ObjConduitPort2 = CreateConduitNozzlePH(pConduitPort, oDir, m_OutputColl, oPartFclt, 2)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort2
    Set ObjConduitPort2 = Nothing
 
' Insert your code for output 7(Conduit Port 3)
    pConduitPort.Set 0.8 * parElecEquipmentWidth / 2, 0.1 * parElecEquipmentLength, parElecEquipmentHeight
    Set ObjConduitPort3 = CreateConduitNozzlePH(pConduitPort, oDir, m_OutputColl, oPartFclt, 3)
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort3
    Set ObjConduitPort3 = Nothing
 
' Insert your code for output 8(Conduit Port 4)
    pConduitPort.Set 0.8 * parElecEquipmentWidth / 2, 0.3 * parElecEquipmentLength, parElecEquipmentHeight
    Set ObjConduitPort4 = CreateConduitNozzlePH(pConduitPort, oDir, m_OutputColl, oPartFclt, 4)
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort4
    Set ObjConduitPort4 = Nothing
    Set pConduitPort = Nothing
    
' Insert your code for output 9(Cable Port)
    Dim iLogicalDistPort As IJLogicalDistPort
    Dim iDistribPort As IJDistribPort
    Dim NozzlePHFactory As NozzlePHFactory
    Dim ObjCableNozzle As CableNozzle
    Dim cablePortPoint As AutoMath.DPosition

    Set cablePortPoint = New AutoMath.DPosition
    Set NozzlePHFactory = New NozzlePHFactory
        
    Set ObjCableNozzle = NozzlePHFactory.CreateCableNozzlePHFromPart(oPartFclt, 5, _
                                                            m_OutputColl.ResourceManager)
    cablePortPoint.Set 0.8 * parElecEquipmentWidth / 2, -0.3 * parElecEquipmentLength, parElecEquipmentHeight
                                
    Set iLogicalDistPort = ObjCableNozzle
    iLogicalDistPort.SetCenterLocation cablePortPoint
    Set iDistribPort = ObjCableNozzle
    iDistribPort.SetDirectionVector oDir

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCableNozzle
    Set ObjCableNozzle = Nothing

' Insert your code for output 10(Cable Port)
    Set ObjCableNozzle = NozzlePHFactory.CreateCableNozzlePHFromPart(oPartFclt, 6, _
                                                            m_OutputColl.ResourceManager)
    cablePortPoint.Set 0.8 * parElecEquipmentWidth / 2, -0.1 * parElecEquipmentLength, parElecEquipmentHeight
                                
    Set iLogicalDistPort = ObjCableNozzle
    iLogicalDistPort.SetCenterLocation cablePortPoint
    Set iDistribPort = ObjCableNozzle
    iDistribPort.SetDirectionVector oDir

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCableNozzle
    Set ObjCableNozzle = Nothing
    
' Insert your code for output 11(Cable Port)
    Set ObjCableNozzle = NozzlePHFactory.CreateCableNozzlePHFromPart(oPartFclt, 7, _
                                                            m_OutputColl.ResourceManager)
    cablePortPoint.Set 0.8 * parElecEquipmentWidth / 2, 0.1 * parElecEquipmentLength, parElecEquipmentHeight
                                
    Set iLogicalDistPort = ObjCableNozzle
    iLogicalDistPort.SetCenterLocation cablePortPoint
    Set iDistribPort = ObjCableNozzle
    iDistribPort.SetDirectionVector oDir

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCableNozzle
    Set ObjCableNozzle = Nothing
    
' Insert your code for output 12(Cable Port)
    Set ObjCableNozzle = NozzlePHFactory.CreateCableNozzlePHFromPart(oPartFclt, 8, _
                                                            m_OutputColl.ResourceManager)
    cablePortPoint.Set 0.8 * parElecEquipmentWidth / 2, 0.3 * parElecEquipmentLength, parElecEquipmentHeight
                                
    Set iLogicalDistPort = ObjCableNozzle
    iLogicalDistPort.SetCenterLocation cablePortPoint
    Set iDistribPort = ObjCableNozzle
    iDistribPort.SetDirectionVector oDir

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCableNozzle
    Set ObjCableNozzle = Nothing
    Set NozzlePHFactory = Nothing
    Set iLogicalDistPort = Nothing
    Set cablePortPoint = Nothing
    Set iDistribPort = Nothing
    Set oDir = Nothing

    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

''' This function creates persistent/Transient Line based on
''' start and end points of the line
'''<{(Line begin)}>
Private Function PlaceLine(pResourceMgr As IUnknown, ByRef startPoint As IJDPosition, _
                            ByRef endPoint As IJDPosition) _
                            As IngrGeom3D.Line3d

    Const METHOD = "PlaceLine"
    On Error GoTo ErrorHandler
        
    Dim oLine As IngrGeom3D.Line3d
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    ' Create Line object
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(pResourceMgr, _
                                startPoint.x, startPoint.y, startPoint.z, _
                                endPoint.x, endPoint.y, endPoint.z)
    
    
    Set PlaceLine = oLine
    Set oLine = Nothing
    Set oGeomFactory = Nothing
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Function
'''<{(Line end)}>

''' This function creates persistent/Transient Arc based on centre point
''' NormalVector, start and end points of the curve
'''<{(PlaceArcByCen begin)}>
Private Function PlaceArcByCen(pResourceMgr As IUnknown, _
                            ByRef centerPoint As IJDPosition, _
                            ByRef startPoint As IJDPosition, _
                            ByRef endPoint As IJDPosition, _
                            ByRef normVector As IJDVector) _
                            As IngrGeom3D.Arc3d
    Const METHOD = "PlaceArcByCen"
    On Error GoTo ErrorHandler
        
    Dim oArc As IngrGeom3D.Arc3d
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oArc = oGeomFactory.Arcs3d.CreateByCtrNormStartEnd(pResourceMgr, _
                              centerPoint.x, centerPoint.y, centerPoint.z, _
                             normVector.x, normVector.y, normVector.z, _
                             startPoint.x, startPoint.y, startPoint.z, _
                            endPoint.x, endPoint.y, endPoint.z)
   
    Set PlaceArcByCen = oArc
    Set oArc = Nothing
    Set oGeomFactory = Nothing

    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Function
'''<{(PlaceArcByCen end)}>




