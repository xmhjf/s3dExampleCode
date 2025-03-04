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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         VRK
'   Creation Date:  Tuesday, April 1 2008
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   01.04.2008      VRK     CR-136575:Enhance rupture disc symbol to provide more realistic dimensional basis
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
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
    Dim pipeDiam2       As Double
    Dim flangeThick2    As Double
    Dim sptOffset2      As Double
    Dim flangeDiam2      As Double
    Dim depth2          As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Integer
    Dim ObjInsulatedCylinder As Object
    Dim ObjInsulatedCone As Object
    
    Dim parHeight As Double
    Dim parBodyOutsideDiameter As Double
    Dim parHubLength As Double
    Dim parNutHeight As Double
    Dim parNutWidth As Double
    Dim parInsulationThickness As Double
    
    Dim oStPoint   As New AutoMath.DPosition
    Dim oEnPoint   As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parHeight = arrayOfInputs(2)
'    parBodyOutsideDiameter= arrayOfInputs(3)
'    parHubLength = arrayOfInputs(4)
'    parNutHeight = arrayOfInputs(5)
'    parNutWidth = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)
    
    iOutput = 0

    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    Dim dhalfHeight As Double
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, _
                                    flangeDiam, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, _
                                    flangeDiam2, sptOffset2, depth2

     Select Case lPartDataBasis
        Case Is <= 1, 785
            parHeight = arrayOfInputs(2)
        Case 786
            parHeight = arrayOfInputs(2)
        Case 787
            parHeight = arrayOfInputs(2)
        Case 788
            parHeight = arrayOfInputs(2)
            parBodyOutsideDiameter = arrayOfInputs(3)
        Case 789
            parHeight = arrayOfInputs(2)
            parBodyOutsideDiameter = arrayOfInputs(3)
            parHubLength = arrayOfInputs(4)
        Case 790
            parHeight = arrayOfInputs(2)
            parHubLength = arrayOfInputs(4)
        Case 791
            parHeight = arrayOfInputs(2)
            parBodyOutsideDiameter = arrayOfInputs(3)
            parHubLength = arrayOfInputs(4)
        Case 792
            parHeight = arrayOfInputs(2)
            parBodyOutsideDiameter = arrayOfInputs(3)
            parHubLength = arrayOfInputs(4)
        Case 793
            parHeight = arrayOfInputs(2)
        Case 794
            parHeight = arrayOfInputs(2)
            parHubLength = arrayOfInputs(4)
        Case 795
            parHeight = arrayOfInputs(2)
            parNutHeight = arrayOfInputs(5)
            parNutWidth = arrayOfInputs(6)
        Case Else
            GoTo ErrorLabel: 'Invalid specification
        
    End Select
    
    If lPartDataBasis = 787 Or lPartDataBasis = 790 Then
        flangeDiam = flangeDiam2
    ElseIf lPartDataBasis = 788 Or lPartDataBasis = 789 Or lPartDataBasis = 791 _
                    Or lPartDataBasis = 792 Then
        flangeDiam = parBodyOutsideDiameter
    ElseIf lPartDataBasis = 795 Then
        flangeDiam = pipeDiam
    End If
    
    dhalfHeight = parHeight / 2
    'Create Cylinder
    oStPoint.Set -dhalfHeight, 0, 0
    oEnPoint.Set dhalfHeight, 0, 0
    Set ObjInsulatedCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, flangeDiam + 2 * parInsulationThickness, True)
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedCylinder
    Set ObjInsulatedCylinder = Nothing
    
    'Insulation for Hub
    If lPartDataBasis = 787 Or lPartDataBasis = 788 Or lPartDataBasis = 789 Then
            oStPoint.Set CenterPos.x - dhalfHeight, CenterPos.y, CenterPos.z
            oEnPoint.Set CenterPos.x - dhalfHeight - flangeThick, CenterPos.y, CenterPos.z
            Set ObjInsulatedCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam + 2 * parInsulationThickness, True)
            m_OutputColl.AddOutput "InsCylinder1", ObjInsulatedCylinder
            Set ObjInsulatedCylinder = Nothing
    End If
    If lPartDataBasis = 788 Or lPartDataBasis = 791 Or lPartDataBasis = 793 Then
            oStPoint.Set CenterPos.x + dhalfHeight, CenterPos.y, CenterPos.z
            oEnPoint.Set CenterPos.x + dhalfHeight + flangeThick2, CenterPos.y, CenterPos.z
            Set ObjInsulatedCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam2 + 2 * parInsulationThickness, True)
        m_OutputColl.AddOutput "InsCylinder2", ObjInsulatedCylinder
        Set ObjInsulatedCylinder = Nothing
    End If
    
    'Insulation for Weldneck
    If lPartDataBasis = 789 Or lPartDataBasis = 792 Or lPartDataBasis = 794 Then
        oStPoint.Set dhalfHeight + parHubLength, 0, 0
        oEnPoint.Set dhalfHeight, 0, 0
        Set ObjInsulatedCone = PlaceCone(m_OutputColl, oStPoint, oEnPoint, (0.5 * pipeDiam2) + parInsulationThickness, _
                    (0.25 * (pipeDiam2 + flangeDiam)) + parInsulationThickness, True)
    '   Set the output
        m_OutputColl.AddOutput "ObjInsCone1", ObjInsulatedCone
        Set ObjInsulatedCone = Nothing
    End If
    If lPartDataBasis = 790 Or lPartDataBasis = 791 Or lPartDataBasis = 792 Then
        oStPoint.Set -(dhalfHeight + parHubLength), 0, 0
        oEnPoint.Set -dhalfHeight, 0, 0
        Set ObjInsulatedCone = PlaceCone(m_OutputColl, oStPoint, oEnPoint, (0.5 * pipeDiam) + parInsulationThickness, _
                     (0.25 * (pipeDiam + flangeDiam)) + parInsulationThickness, True)
    '   Set the output
        m_OutputColl.AddOutput "ObjInsCone2", ObjInsulatedCone
        Set ObjInsulatedCone = Nothing
    End If
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set CenterPos = Nothing
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub


