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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:09:27 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         svsmylav
'   Creation Date:  Thursday, Mar 20 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjBox1 As Object
    Dim ObjBox2 As Object
    Dim ObjBox3 As Object
    Dim dblXLength As Double
    Dim dblYLength As Double
    Dim dblZLength As Double
    Dim strOrigin As String
    
    Dim parNominalShellDiameter As Double
    Dim parA As Double
    Dim parB As Double
    Dim parD As Double
    Dim parDG As Double
    Dim parE As Double
    Dim parG As Double
    Dim parJ As Double
    Dim parT As Double
    Dim parVesselDesignPressureH As Double
    Dim parVesselDesignPressureR As Double
    Dim parVesselHeight As Double
    Dim oriOrientation As Orientation
    

    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parNominalShellDiameter = arrayOfInputs(2)
    parA = arrayOfInputs(3)
    parB = arrayOfInputs(4)
    parD = arrayOfInputs(5)
    parDG = arrayOfInputs(6)
    parE = arrayOfInputs(7)
    parG = arrayOfInputs(8)
    parJ = arrayOfInputs(9)
    parT = arrayOfInputs(10)
    parVesselDesignPressureH = arrayOfInputs(11)
    parVesselDesignPressureR = arrayOfInputs(12)
    parVesselHeight = arrayOfInputs(13)
    
    iOutput = 0
    
' Insert your code for output 1(Box 1)

    strOrigin = "E 0 N 0 U 0"
    dblXLength = parG
    dblYLength = parB
    dblZLength = parE
    Set ObjBox1 = CreateBox(m_outputColl, strOrigin, Nothing, parG, parB, parE, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBox1
    Set ObjBox1 = Nothing
    
' Insert your code for output 2(Box 2)
    strOrigin = "E -" & parJ / 2 - parD / 2 & " N 0 U " & parVesselHeight / 2 + parE / 2
    dblXLength = parD
    dblYLength = parB - 0.012
    dblZLength = parVesselHeight
    Set ObjBox2 = CreateBox(m_outputColl, strOrigin, Nothing, dblXLength, dblYLength, dblZLength, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), ObjBox2
    Set ObjBox2 = Nothing

' Insert your code for output 3(Box3)
    If (parNominalShellDiameter >= 0.5) Then
        strOrigin = "E " & parJ / 2 - parD / 2 & " N 0 U " & parVesselHeight / 2 + parE / 2
        dblXLength = parD
        dblYLength = parB - 0.012
        dblZLength = parVesselHeight
        Set ObjBox3 = CreateBox(m_outputColl, strOrigin, Nothing, dblXLength, dblYLength, dblZLength, True)
          
    ' Set the output
        m_outputColl.AddOutput "Box_", ObjBox3
        Set ObjBox3 = Nothing
    End If
    
    Exit Sub
    
ErrorLabel:
'    MsgBox Err.Number & " " & Err.Description & " " & Err.Source
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
