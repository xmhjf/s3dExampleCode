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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 1/13/2005-6:23:52 PM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004-08, Intergraph Corporation. All rights reserved.
'
'   Csimplephysical.cls
'   Author:         BBA
'   Creation Date:  Monday, Oct 18 2004
'   Description:
'   Details are taken from www.lithonia.com site.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   11.JAN.2005     MS     CR69200 Added cable ports at the same location of conduit ports.
'   13.Jan.2005     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  22.Apr.2008     VRK           CR-120097  Need to set proper units to the cable nozzle diameter and area
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Dim m_oGeomHelper As IJSymbolGeometryHelper
Dim PI As Double

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
    Dim ConduitOD        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim parFacetoCenter As Double
    Dim ostPoint As New AutoMath.DPosition
    Dim oenPoint As New AutoMath.DPosition
    Dim iOutput     As Double

'height along z-axis & width along x and y axis'

'origin is at the Center of Output1


' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)
           
    iOutput = 0

' Insert the code for Output1(cylinder)
    Dim objCylinder  As Object
    ostPoint.Set 0, 0, -0.15 * parFacetoCenter
    oenPoint.Set 0, 0, 0.15 * parFacetoCenter
    Set objCylinder = PlaceCylinder(m_OutputColl, ostPoint, oenPoint, 1.2 * parFacetoCenter, True)
 
' Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
    Set objCylinder = Nothing
    Set ostPoint = Nothing
    Set oenPoint = Nothing
    
    
' Insert the code for Output2(cylinder1)
    ostPoint.Set -parFacetoCenter, 0, 0
    oenPoint.Set parFacetoCenter, 0, 0
    Set objCylinder = PlaceCylinder(m_OutputColl, ostPoint, oenPoint, 0.208333333333333 * parFacetoCenter, True)
 
' Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
    Set objCylinder = Nothing
    Set ostPoint = Nothing
    Set oenPoint = Nothing
    
' Insert the code for Output3(cylinder2)
    ostPoint.Set 0, -parFacetoCenter, 0
    oenPoint.Set 0, parFacetoCenter, 0
    Set objCylinder = PlaceCylinder(m_OutputColl, ostPoint, oenPoint, 0.208333333333333 * parFacetoCenter, True)
 
' Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
    Set objCylinder = Nothing
    Set ostPoint = Nothing
    Set oenPoint = Nothing

' Insert your code for Output4(Conduit Port)
    Dim oDir As IJDVector
    Dim ObjConduitPort As IJConduitPortOcc
    Dim portPoint As AutoMath.DPosition
    Set portPoint = New AutoMath.DPosition
    Set oDir = New DVector
    oDir.Set -1, 0, 0
   
    portPoint.Set -parFacetoCenter, 0, 0
    Set ObjConduitPort = CreateConduitNozzlePH(portPoint, oDir, m_OutputColl, oPartFclt, 1)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort
    Set ObjConduitPort = Nothing
    Set oDir = Nothing
    Set portPoint = Nothing
    
' Insert your code for Output5(Conduit Port)
    Set portPoint = New AutoMath.DPosition
    Set oDir = New DVector
    oDir.Set 1, 0, 0
   
    portPoint.Set parFacetoCenter, 0, 0
    Set ObjConduitPort = CreateConduitNozzlePH(portPoint, oDir, m_OutputColl, oPartFclt, 2)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort
    Set ObjConduitPort = Nothing
    Set oDir = Nothing
    Set portPoint = Nothing
    
' Insert your code for Output6(Conduit Port)
    Set portPoint = New AutoMath.DPosition
    Set oDir = New DVector
    oDir.Set 0, -1, 0
   
    portPoint.Set 0, -parFacetoCenter, 0
    Set ObjConduitPort = CreateConduitNozzlePH(portPoint, oDir, m_OutputColl, oPartFclt, 3)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort
    Set ObjConduitPort = Nothing
    Set oDir = Nothing
    Set portPoint = Nothing
    
' Insert your code for Output7(Conduit Port)
    Set portPoint = New AutoMath.DPosition
    Set oDir = New DVector
    oDir.Set 0, 1, 0
   
    portPoint.Set 0, parFacetoCenter, 0
    Set ObjConduitPort = CreateConduitNozzlePH(portPoint, oDir, m_OutputColl, oPartFclt, 4)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort
    Set ObjConduitPort = Nothing
    Set portPoint = Nothing
    
' Insert your code for output 8(Cable Port)
    Dim iLogicalDistPort As IJLogicalDistPort
    Dim iDistribPort As IJDistribPort
    Dim NozzlePHFactory As NozzlePHFactory
    Dim ObjCableNozzle As CableNozzle
    Dim cablePortPoint As AutoMath.DPosition

    Set cablePortPoint = New AutoMath.DPosition
    Set NozzlePHFactory = New NozzlePHFactory
        
    Set ObjCableNozzle = NozzlePHFactory.CreateCableNozzlePHFromPart(oPartFclt, 5, _
                                                            m_OutputColl.ResourceManager)
    cablePortPoint.Set -parFacetoCenter, 0, 0
                                
    Set iLogicalDistPort = ObjCableNozzle
    iLogicalDistPort.SetCenterLocation cablePortPoint
    Set iDistribPort = ObjCableNozzle
    oDir.Set -1, 0, 0
    iDistribPort.SetDirectionVector oDir
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCableNozzle
    Set ObjCableNozzle = Nothing

    
' Insert your code for output 9(Cable Port)
    Set ObjCableNozzle = NozzlePHFactory.CreateCableNozzlePHFromPart(oPartFclt, 6, _
                                                            m_OutputColl.ResourceManager)
    cablePortPoint.Set parFacetoCenter, 0, 0
    Set iLogicalDistPort = ObjCableNozzle
    iLogicalDistPort.SetCenterLocation cablePortPoint
    Set iDistribPort = ObjCableNozzle
    oDir.Set 1, 0, 0
    iDistribPort.SetDirectionVector oDir
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCableNozzle
    Set ObjCableNozzle = Nothing
    
' Insert your code for output 10(Cable Port)
    Set ObjCableNozzle = NozzlePHFactory.CreateCableNozzlePHFromPart(oPartFclt, 7, _
                                                            m_OutputColl.ResourceManager)
    cablePortPoint.Set 0, -parFacetoCenter, 0
                                
    Set iLogicalDistPort = ObjCableNozzle
    iLogicalDistPort.SetCenterLocation cablePortPoint
    Set iDistribPort = ObjCableNozzle
    oDir.Set 0, -1, 0
    iDistribPort.SetDirectionVector oDir
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCableNozzle
    Set ObjCableNozzle = Nothing
    
' Insert your code for output 11(Cable Port)
    Set ObjCableNozzle = NozzlePHFactory.CreateCableNozzlePHFromPart(oPartFclt, 8, _
                                                            m_OutputColl.ResourceManager)
    cablePortPoint.Set 0, parFacetoCenter, 0
                                
    Set iLogicalDistPort = ObjCableNozzle
    iLogicalDistPort.SetCenterLocation cablePortPoint
    Set iDistribPort = ObjCableNozzle
    oDir.Set 0, 1, 0
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
