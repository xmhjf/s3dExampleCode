VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:01:40 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   COperation.cls
'   Author:          JTF
'   Creation Date:  Friday, Mar 22 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   20.May.2005    MS  CR-76070: Modified the nozzle creation code(as members),added datum points.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   26.Sep.2006    RRK  TR84724   If variables of type double are compared, appropriate CmpDblXXX functions are used.
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
    Dim ObjSection1Support As Object
    Dim ObjSection1Dome As Object
    Dim ObjSection1Tank As Object
    Dim ObjSection2Support As Object
    Dim ObjSection2Dome As Object
    Dim ObjSection2Tank As Object
    Dim ObjSection3Support As Object
    Dim ObjSection3Tank As Object
    Dim ObjSection4Support As Object
    Dim ObjSection4Tank As Object
    Dim ObjSection4Dome As Object
    Dim ObjCVMaintenance As Object
    Dim ObjCVOperation As Object
    Dim ObjCVInsulation As Object
    
    Dim parSupportHeight1 As Double
    Dim parSupportDiameter1 As Double
    Dim parDomeHeight1 As Double
    Dim parVesselHeight1 As Double
    Dim parVesselDiameter1 As Double
    Dim parSupportHeight2 As Double
    Dim parDomeHeight2 As Double
    Dim parVesselHeight2 As Double
    Dim parVesselDiameter2 As Double
    Dim parSupportHeight3 As Double
    Dim parVesselHeight3 As Double
    Dim parVesselDiameter3 As Double
    Dim parSupportHeight4 As Double
    Dim parVesselHeight4 As Double
    Dim parVesselDiameter4 As Double
    Dim parDomeHeight4 As Double


' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parSupportHeight1 = arrayOfInputs(2)
    parSupportDiameter1 = arrayOfInputs(3)
    parDomeHeight1 = arrayOfInputs(4)
    parVesselHeight1 = arrayOfInputs(5)
    parVesselDiameter1 = arrayOfInputs(6)
    parSupportHeight2 = arrayOfInputs(7)
    parDomeHeight2 = arrayOfInputs(8)
    parVesselHeight2 = arrayOfInputs(9)
    parVesselDiameter2 = arrayOfInputs(10)
    parSupportHeight3 = arrayOfInputs(11)
    parVesselHeight3 = arrayOfInputs(12)
    parVesselDiameter3 = arrayOfInputs(13)
    parSupportHeight4 = arrayOfInputs(14)
    parVesselHeight4 = arrayOfInputs(15)
    parVesselDiameter4 = arrayOfInputs(16)
    parDomeHeight4 = arrayOfInputs(17)

    
    iOutput = 0

 ' Insert your code for output 13()
    Dim TotalHeightofAssm As Double
    TotalHeightofAssm = parSupportHeight1 + parVesselHeight1 + parSupportHeight2 + parVesselHeight2 + parSupportHeight3 + parVesselHeight3 + parSupportHeight4 + parVesselHeight4 + parDomeHeight4

    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition

    stPoint.Set 0, 0, 0
    enPoint.Set 0, 0, TotalHeightofAssm * 1.125
    
    'obtain largest diameter
    Dim LargestDiameter As Double
    
    LargestDiameter = parSupportHeight1
    If CmpDblLessThan(LargestDiameter, parVesselDiameter1) Then LargestDiameter = parVesselDiameter1
    If CmpDblLessThan(LargestDiameter, parVesselDiameter2) Then LargestDiameter = parVesselDiameter2
    If CmpDblLessThan(LargestDiameter, parVesselDiameter3) Then LargestDiameter = parVesselDiameter3
    If CmpDblLessThan(LargestDiameter, parVesselDiameter4) Then LargestDiameter = parVesselDiameter4
    
    Set ObjCVOperation = PlaceCylinder(m_OutputColl, stPoint, enPoint, LargestDiameter * 1.125, True)



' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCVOperation
    Set ObjCVOperation = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
