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
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          NN
'   Creation Date:  Thursday, Jan 25 2001
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
    '
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   08.SEP.2006     KKC    DI-95670     Replace names with initials in all revision history sheets and symbols
'   16.JUL.2008     MP     CR-145604    Implemented part data basis options 1028,1029,1030,1031 and 1032
'   24.Sep.2008     MP     CR-149961    Implemented the part data basis options 1084, 1085 and 1086
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI As Double

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim StartPos       As New AutoMath.DPosition
    Dim lPartDataBasis As Long
    Dim oPipeComponent As IJDPipeComponent
    
    StartPos.Set 0, 0, 0

    Dim iOutput     As Double
    Dim ObjBody As Object
    Dim ObjBody2 As Object
    Dim ObjBody3 As Object
    Dim ObjInsulatedBody As Object
    Dim ObjInsulatedBody2 As Object
    Dim parFacetoEnd As Double
    Dim parInsulationThickness As Double
    Dim parLength1 As Double
    Dim parLength2 As Double
    Dim parWidth As Double
    Dim dLength1 As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoEnd = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis

    If lPartDataBasis <= 1 Or lPartDataBasis = PLUG_SQUARE Or lPartDataBasis = PLUG_PENTAGON Or _
        lPartDataBasis = PLUG_HEXAGON Or lPartDataBasis = PLUG_OCTAGON Then
            parLength1 = arrayOfInputs(4)
            parLength2 = arrayOfInputs(5)
    End If
    
    parWidth = arrayOfInputs(6)
    
    iOutput = 0

    'Insert your code for output 2(Insulated Body)
    Dim objPlugEndPosition   As New AutoMath.DPosition
    Dim parInsulationDiameter As Double
    
    If lPartDataBasis <= 1 Or lPartDataBasis = PLUG_DEFAULT Then
        objPlugEndPosition.Set parFacetoEnd + parInsulationThickness, 0, 0
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
        parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
        
        Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, StartPos, objPlugEndPosition, parInsulationDiameter, True)
        
        ' Set the output
        iOutput = iOutput + 1
        
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
        Set ObjInsulatedBody = Nothing
    
    ElseIf lPartDataBasis = PLUG_SQUARE Or lPartDataBasis = PLUG_PLAIN Or lPartDataBasis = PLUG_ROUND Or lPartDataBasis = PLUG_COUNTERSUNK Then
        
        If lPartDataBasis = PLUG_PLAIN Or lPartDataBasis = PLUG_ROUND Then
            dLength1 = 0.2 * parFacetoEnd
            objPlugEndPosition.Set parFacetoEnd - dLength1 + parInsulationThickness, 0, 0
        ElseIf lPartDataBasis = PLUG_COUNTERSUNK Then
            objPlugEndPosition.Set parFacetoEnd + parInsulationThickness, 0, 0
        Else
            objPlugEndPosition.Set parLength1 + parInsulationThickness, 0, 0
        End If
        
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
        parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
        
        Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, StartPos, objPlugEndPosition, parInsulationDiameter, True)
    
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
        
        Set ObjInsulatedBody = Nothing
        Set StartPos = Nothing
        
        ' set output (Insulated Body2)
        If parWidth = 0 Then
            parWidth = 0.6 * pipeDiam
        End If
        
        If lPartDataBasis = PLUG_PLAIN Then
            
            StartPos.Set parFacetoEnd - dLength1, -parWidth / 2 - parInsulationThickness, -parWidth / 2 - parInsulationThickness
            objPlugEndPosition.Set parFacetoEnd + parInsulationThickness, parWidth / 2 + parInsulationThickness, parWidth / 2 + parInsulationThickness
            Set ObjInsulatedBody2 = PlaceBox(m_OutputColl, StartPos, objPlugEndPosition)
            ' Set the output
            m_OutputColl.AddOutput "InsulatedBody2", ObjInsulatedBody2
            Set ObjInsulatedBody2 = Nothing
            
        ElseIf lPartDataBasis = PLUG_ROUND Then
            Dim dInsDia As Double
            dInsDia = parWidth + 2 * parInsulationThickness
            
            StartPos.Set parFacetoEnd - dLength1, 0, 0
            objPlugEndPosition.Set parFacetoEnd + parInsulationThickness, 0, 0
            Set ObjInsulatedBody2 = PlaceCylinder(m_OutputColl, StartPos, objPlugEndPosition, dInsDia, True)
            ' Set the output
            m_OutputColl.AddOutput "InsulatedBody2", ObjInsulatedBody2
            Set ObjInsulatedBody2 = Nothing
    
        ElseIf lPartDataBasis = PLUG_SQUARE Then
            
            StartPos.Set parLength1, -parWidth / 2 - parInsulationThickness, -parWidth / 2 - parInsulationThickness
            objPlugEndPosition.Set parLength1 + parLength2 + parInsulationThickness, parWidth / 2 + parInsulationThickness, parWidth / 2 + parInsulationThickness
            Set ObjInsulatedBody2 = PlaceBox(m_OutputColl, StartPos, objPlugEndPosition)
            ' Set the output
            m_OutputColl.AddOutput "InsulatedBody2", ObjInsulatedBody2
            Set ObjInsulatedBody2 = Nothing
    
        End If
        
'        ' Set the output
'        m_OutputColl.AddOutput "InsulatedBody2", ObjInsulatedBody2
'        Set ObjInsulatedBody2 = Nothing
    
    ElseIf lPartDataBasis = PLUG_HEXAGON Or lPartDataBasis = PLUG_OCTAGON Or lPartDataBasis = PLUG_PENTAGON Then
        
        objPlugEndPosition.Set parLength1 + parInsulationThickness, 0, 0
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
        parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
        
        Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, StartPos, objPlugEndPosition, parInsulationDiameter, True)
    
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
        Set ObjInsulatedBody = Nothing
        Set StartPos = Nothing
         
        'Insert your code for Insulated Body
        Dim lNumSides   As Long
        Dim oTransMat   As AutoMath.DT4x4
        Dim oVec        As AutoMath.DVector
        Set oVec = New AutoMath.DVector
        Set oTransMat = New AutoMath.DT4x4
        
        If lPartDataBasis = PLUG_HEXAGON Then
            lNumSides = 6
        ElseIf lPartDataBasis = PLUG_OCTAGON Then
            lNumSides = 8
        ElseIf lPartDataBasis = PLUG_PENTAGON Then
            lNumSides = 10
        End If
        
        Dim dInculdedAngle As Double
        Dim dSideLength As Double
        
        dInculdedAngle = 360 / lNumSides
        dInculdedAngle = (dInculdedAngle * PI) / 180
        
        If CmpDblEqual(parWidth, 0) Then
            'Assuming that side length of the polygon shaped hub is 0.7 times the Piping Outside Diameter of Port1
            dSideLength = 1.2 * pipeDiam
        Else
            dSideLength = parWidth * Sin(dInculdedAngle / 2)
        End If
        
        Set ObjInsulatedBody2 = PlaceNnagon(m_OutputColl, lNumSides, dSideLength + parInsulationThickness, _
                                            parLength2 + parInsulationThickness, True)
        oVec.Set parLength1, 0, 0
        oTransMat.LoadIdentity
        oTransMat.Translate oVec
        ObjInsulatedBody2.Transform oTransMat
        
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "InsulatedBody2", ObjInsulatedBody2
        
        Set ObjInsulatedBody2 = Nothing
        Set oTransMat = Nothing
        Set oVec = Nothing
     End If
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
End Sub
