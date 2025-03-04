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
'
'   CInsulation.cls
'   Author:          MS
'   Creation Date:  Tuesday, Sept 14 2004
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Slide Valve details were taken from PDS Piping Component Data Reference Manual
'   at Page No - D-21 and SN=V33. This Symbol consist of Physical and Insulation aspects.
'   Physical aspect consist of a Nozzle with Length , Nozzle, branch Cylinder  ,a Box and a output for Valve Operator
'   Insulation aspect consist of a Cylinder for main body ,another cylinder for branch and a Box for
'   Projection at top of BranchOption Explicit'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Const MODULE = "Insulation:" 'Used for error messages

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
    Dim ObjBodyInsulation As Object
    Dim ObjBranchInsulation As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parBranchLength As Double
    Dim parProjectionHeight As Double
    Dim parProjectionLength As Double
    Dim parProjectionWidth As Double
    Dim parInsulationThickness As Double
    Dim InsulationDiameter1 As Double
    Dim InsulationDiameter2 As Double
    Dim InsulationDiameter As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parBranchLength = arrayOfInputs(4)
    parProjectionHeight = arrayOfInputs(5)
    parProjectionLength = arrayOfInputs(6)
    parProjectionWidth = arrayOfInputs(7)
    parInsulationThickness = arrayOfInputs(8)
    
    iOutput = 0
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim startPoint As AutoMath.DPosition
    Dim endPoint As AutoMath.DPosition
    
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
  ' Insert your code for output 4(Branch Insulation)
    startPoint.Set 0, 0, 0
    endPoint.Set 0, 0, parBranchLength - parProjectionHeight
    InsulationDiameter = pipeDiam + parInsulationThickness * 2
    Set ObjBranchInsulation = PlaceCylinder(m_OutputColl, startPoint, endPoint, InsulationDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBranchInsulation
    Set ObjBranchInsulation = Nothing
    
  ' Insert your code for output Projection Box
    Dim lowerpt As AutoMath.DPosition
    Dim upperpt As AutoMath.DPosition
    Set lowerpt = New AutoMath.DPosition
    Set upperpt = New AutoMath.DPosition
    
    lowerpt.Set -parProjectionLength / 2 - parInsulationThickness, _
                        -parProjectionWidth / 2 - parInsulationThickness, _
                        parBranchLength - parProjectionHeight - parInsulationThickness
    upperpt.Set parProjectionLength / 2 + parInsulationThickness, _
                        parProjectionWidth / 2 + parInsulationThickness, _
                        parBranchLength + parInsulationThickness
    Set ObjBoxInsulation = PlaceBox(m_OutputColl, lowerpt, upperpt)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBoxInsulation
    Set ObjBoxInsulation = Nothing
    Set lowerpt = Nothing
    Set upperpt = Nothing
    
 ' Insert your code for output 3(Body Insulation)
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
    startPoint.Set -parFace1toCenter, 0, 0
    endPoint.Set parFace2toCenter, 0, 0
    Set ObjBodyInsulation = PlaceCylinder(m_OutputColl, startPoint, endPoint, InsulationDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyInsulation
    Set ObjBodyInsulation = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
