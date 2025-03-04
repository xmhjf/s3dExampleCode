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
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:19:42 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          svsmylav
'   Creation Date:  Thursday, Apr 18 2002
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    Symbol is created using one Insulation aspect output: Insulation for the vessel using 'PlaceRevolution'
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     -----                   ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   April 1 2003    svsmylav        Single complex string revolution is replaced with
'                                           cylinder/cone primitives.
'   29.Nov.2004     V6UpgradeSO             Made compatible with Smart Occurrence based Equipments
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'   06.May.2005     svsmylav        CR-76070: Removed the Nozzle related parameters and code to create insulation for nozzles.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  26.Sep.2006     RRK  TR-84724  If variables of type double are compared, appropriate CmpDblXXX functions are used.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private Const TOP_BODY = 1

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjInsEquipment As Object
    Dim ObjInsNozzle1WithLength As Object
    Dim ObjInsNozzle2WithLength As Object
    Dim ObjInsNozzle3WithLength As Object
    
    Dim parBottomBodyHeight1 As Double
    Dim parBottomBodyHeight2 As Double
    Dim parBottomBodyDiameter As Double
    Dim parShaftLength As Double
    Dim parShaftDiameter As Double
    Dim parTaperBodyHeight As Double
    Dim parTopBodyHeight As Double
    Dim parTopBodyDiameter As Double
    Dim parBundlePullingLength As Double
    Dim parInsulationThickness As Double
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    Dim parInsulationDiameter As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBottomBodyHeight1 = arrayOfInputs(2) 'P1
    parBottomBodyHeight2 = arrayOfInputs(3) 'P2
    parBottomBodyDiameter = arrayOfInputs(4) 'P3
    parShaftLength = arrayOfInputs(5)       'P4
    parShaftDiameter = arrayOfInputs(6)     'P5
    parTaperBodyHeight = arrayOfInputs(7)   'P6
    parTopBodyHeight = arrayOfInputs(8)     'P7
    parTopBodyDiameter = arrayOfInputs(9)   'P8
    parBundlePullingLength = arrayOfInputs(10) 'P9
    parInsulationThickness = arrayOfInputs(11)
   
    iOutput = 0
    
'   Input validation
    If CmpDblLessThanOrEqualTo(parBottomBodyHeight1, LINEAR_TOLERANCE) Or CmpDblLessThanOrEqualTo(parBottomBodyHeight2, LINEAR_TOLERANCE) Or CmpDblLessThanOrEqualTo(parBottomBodyDiameter, LINEAR_TOLERANCE) Then GoTo ErrorLabel
    If CmpDblGreaterthan(parShaftDiameter, LINEAR_TOLERANCE) And CmpDblGreaterthan(parShaftDiameter, parBottomBodyDiameter) Then GoTo ErrorLabel
    If CmpDblGreaterthan(parTopBodyDiameter, LINEAR_TOLERANCE) And CmpDblLessThan(parTopBodyDiameter, parBottomBodyDiameter) Then GoTo ErrorLabel

' Insert your code for output 1(InsEquipment)
    Dim flag      As Integer
    flag = 0
    
'   Place Top Body
    If CmpDblGreaterthan(parTopBodyHeight, LINEAR_TOLERANCE) And CmpDblGreaterthan(parTopBodyDiameter, LINEAR_TOLERANCE) Then
        Dim ObjInsTopBody As Object
        
        flag = TOP_BODY
        stPoint.Set 0, 0, _
                parInsulationThickness + parTopBodyHeight + parTaperBodyHeight + parBottomBodyHeight1
            
        enPoint.Set 0, 0, parTaperBodyHeight + parBottomBodyHeight1
        Set ObjInsTopBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                    parTopBodyDiameter + 2 * parInsulationThickness, True)
        
        'Set the output
        m_OutputColl.AddOutput "InsTopBody", ObjInsTopBody
        Set ObjInsTopBody = Nothing
    End If
    
'   Place Taper Body
    If CmpDblGreaterthan(parTopBodyDiameter, parBottomBodyDiameter) And CmpDblGreaterthan(parTaperBodyHeight, LINEAR_TOLERANCE) Then
        Dim ObjInsTaperBody As Object
        
        stPoint.Set 0, 0, parTaperBodyHeight + parBottomBodyHeight1
        enPoint.Set 0, 0, parBottomBodyHeight1
        Set ObjInsTaperBody = PlaceCone(m_OutputColl, stPoint, enPoint, _
                parTopBodyDiameter / 2 + parInsulationThickness, _
                parBottomBodyDiameter / 2 + parInsulationThickness, True)
        
        'Set the output
        m_OutputColl.AddOutput "InsTaperBody", ObjInsTaperBody
        Set ObjInsTaperBody = Nothing
    End If
    
'   Place Bottom Body
    stPoint.Set 0, 0, parBottomBodyHeight1
    enPoint.Set 0, 0, -parBottomBodyHeight2 - parInsulationThickness
    Set ObjInsEquipment = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
        parBottomBodyDiameter + 2 * parInsulationThickness, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsEquipment
    Set ObjInsEquipment = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
