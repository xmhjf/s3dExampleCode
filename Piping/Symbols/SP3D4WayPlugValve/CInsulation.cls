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
'   Copyright (c) 2003-04, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          MS
'   Creation Date:  Thursday, Jun 14 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol Model No. is: V19, Page No. D-18 PDS Piping Component Data Reference Guide.
' The Symbol is made of Thirteen output objects. Eleven are physical aspect outputs and Two are insulation aspect outputs.
' The physical outputs are made using 'PlaceCone' and 'CreateNozzle' functions.
'The Insulation outputs are made using 'PlaceCylinder' functions.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'   03.Dec.2004     svsmylav        Defined new input parameter HandwheelAngle.
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
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjLeftCone As Object
    Dim ObjUpperCone As Object
    Dim ObjLowerCone As Object
    Dim ObjRightCone As Object
    Dim ObjTopCone As Object
    Dim ObjBottomCone As Object
    Dim ObjInsulCylinder1 As Object
    Dim ObjInsulCylinder2 As Object
    
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parFace4toCenter As Double
    Dim parInsulationThickness As Double
    
    Dim parHandwheelAngle   As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)     'A
    parFace2toCenter = arrayOfInputs(3)     'B
    parFace3toCenter = arrayOfInputs(4)     'C
    parFace4toCenter = arrayOfInputs(5)     'C
    parInsulationThickness = arrayOfInputs(6)
    parHandwheelAngle = arrayOfInputs(7)

    
    iOutput = 0
 ' Insert your code for output 7(Insul Cylinder1)
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim InsulationDiameter As Double
    Dim InsulationDiameter1 As Double
    Dim InsulationDiameter2 As Double
    
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set parFace2toCenter, 0, 0
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeDiam > pipeDiam Then
        InsulationDiameter1 = flangeDiam + parInsulationThickness * 2
    Else
        InsulationDiameter1 = pipeDiam + parInsulationThickness * 2
    End If
       
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeDiam > pipeDiam Then
        InsulationDiameter2 = flangeDiam + parInsulationThickness * 2
    Else
        InsulationDiameter2 = pipeDiam + parInsulationThickness * 2
    End If
    
    If InsulationDiameter1 > InsulationDiameter2 Then
        InsulationDiameter = InsulationDiameter1
    Else
        InsulationDiameter = InsulationDiameter2
    End If
    
    Set ObjInsulCylinder1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulCylinder1
    Set ObjInsulCylinder1 = Nothing
    
' Insert your code for output 8(Insul Cylinder2)
    Dim InsulationDiameter3 As Double
    Dim InsulationDiameter4 As Double
    
    stPoint.Set 0, parFace3toCenter, 0
    enPoint.Set 0, -parFace4toCenter, 0
    
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeDiam > pipeDiam Then
        InsulationDiameter3 = flangeDiam + parInsulationThickness * 2
    Else
        InsulationDiameter3 = pipeDiam + parInsulationThickness * 2
    End If
       
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeDiam > pipeDiam Then
        InsulationDiameter4 = flangeDiam + parInsulationThickness * 2
    Else
        InsulationDiameter4 = pipeDiam + parInsulationThickness * 2
    End If
    
    If InsulationDiameter3 > InsulationDiameter4 Then
        InsulationDiameter = InsulationDiameter3
    Else
        InsulationDiameter = InsulationDiameter4
    End If
    
    Set ObjInsulCylinder2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulCylinder2
    Set ObjInsulCylinder2 = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
