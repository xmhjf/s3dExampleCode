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
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:04:36 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          NN
'   Creation Date:  Tuesday, Mar 12 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

'''

End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput     As Double
    
    Dim ObjEnclosureBody As Object
    Dim ObjConduitPort1 As IJConduitPortOcc
    Dim ObjConduitPort2 As IJConduitPortOcc
    Dim ObjConduitPort3 As IJConduitPortOcc
    Dim ObjConduitPort4 As IJConduitPortOcc
    Dim ObjCableTrayPort   As IJCableTrayPortOcc
    Dim ObjEnclosureOp As Object
    
    Dim parBasicAsmHeight As Double
    Dim parBasicAsmWidth As Double
    Dim parBasicAsmLength As Double
    
    Dim CenterPos       As IJDPosition
    Dim pBody1   As IJDPosition
    Dim pCableTrayPort   As IJDPosition
    Dim pConduitPort1   As IJDPosition
    Dim pConduitPort2   As IJDPosition
    Dim pConduitPort3   As IJDPosition
    Dim pConduitPort4   As IJDPosition
    Dim oDir        As IJDVector
    Dim oRadialOrient As IJDVector

    Set CenterPos = New DPosition
    Set pCableTrayPort = New DPosition
    Set pConduitPort1 = New DPosition
    Set pConduitPort2 = New DPosition
    Set pConduitPort3 = New DPosition
    Set pConduitPort4 = New DPosition
    Set pBody1 = New DPosition
    Set oDir = New DVector
    Set oRadialOrient = New DVector
    
    CenterPos.Set 0, 0, 0


' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBasicAsmHeight = arrayOfInputs(2)
    parBasicAsmWidth = arrayOfInputs(3)
    parBasicAsmLength = arrayOfInputs(4)
    
    pCableTrayPort.Set CenterPos.x + parBasicAsmLength / 2, CenterPos.y + parBasicAsmWidth / 2, CenterPos.z + parBasicAsmHeight
    
    iOutput = 0
    
    oDir.Set 0, 0, 1
    oRadialOrient.Set 1, 0, 0
 ' Insert your code for output 1(Enclosure Body)
    pBody1.Set CenterPos.x + parBasicAsmLength, CenterPos.y + parBasicAsmWidth, CenterPos.z + parBasicAsmHeight
    Set ObjEnclosureBody = PlaceBox(m_OutputColl, CenterPos, pBody1)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjEnclosureBody
    Set ObjEnclosureBody = Nothing
    
' Insert Output 2:  place Conduit Port 1
    pConduitPort1.Set pCableTrayPort.x, pCableTrayPort.y - parBasicAsmWidth / 4, pCableTrayPort.z
    Set ObjConduitPort1 = CreateConduitNozzlePH(pConduitPort1, oDir, m_OutputColl, oPartFclt, 1)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort1
    Set ObjConduitPort1 = Nothing
    
' Insert Output 3: place Conduit Port 2
    pConduitPort2.Set pCableTrayPort.x, pCableTrayPort.y - parBasicAsmWidth / 12, pCableTrayPort.z
    Set ObjConduitPort2 = CreateConduitNozzlePH(pConduitPort2, oDir, m_OutputColl, oPartFclt, 2)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort2
    Set ObjConduitPort2 = Nothing
 
' Insert your code for output 4(Conduit Port 3)
    pConduitPort3.Set pCableTrayPort.x, pCableTrayPort.y + parBasicAsmWidth / 12, pCableTrayPort.z
    Set ObjConduitPort3 = CreateConduitNozzlePH(pConduitPort3, oDir, m_OutputColl, oPartFclt, 3)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort3
    Set ObjConduitPort3 = Nothing
 
' Insert your code for output 5(Conduit Port 4)
    pConduitPort4.Set pCableTrayPort.x, pCableTrayPort.y + parBasicAsmWidth / 4, pCableTrayPort.z
    Set ObjConduitPort4 = CreateConduitNozzlePH(pConduitPort4, oDir, m_OutputColl, oPartFclt, 4)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConduitPort4
    Set ObjConduitPort4 = Nothing
     
' Insert your code for output 6(Cable Tray Port)
    Set ObjCableTrayPort = CreateCableTrayPortPH(oPartFclt, 5, pCableTrayPort, oDir, oRadialOrient, m_OutputColl)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCableTrayPort
    Set ObjCableTrayPort = Nothing

    
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

