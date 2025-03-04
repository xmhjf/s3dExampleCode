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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          NN
'   Creation Date:  Tuesday, Nov 16 1999
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy         who                 change description
'   -----------         -----               ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   05.Jan.2004     svsmylav        TR -53768: Code is corrected to avoid failure for Sqr function for negative values.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  18.APR.2007     KKC  CR-117893 Modified Olet symbol to be Trimmable
'   2.Jul.2007     svsmylav  TR-123363: Added PI/2 value to parAngle to support V6.1 instances.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private PI           As Double
Private RAD          As Double

Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()

    PI = 4 * Atn(1)
    RAD = 180 / PI
    
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
   
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Dim parBranchFittingPlacementType As Long
        
    Set oPartFclt = arrayOfInputs(1)
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
     
    parBranchFittingPlacementType = arrayOfInputs(7)
        
    If lPartDataBasis <= 1 Then
        If parBranchFittingPlacementType = 5 Then
            Call runInsPDB1(m_OutputColl, arrayOfInputs, arrayOfOutputs)
        Else
            Call runInsPDB1onFitting(m_OutputColl, arrayOfInputs, arrayOfOutputs)
        End If
    End If
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub


Private Sub runInsPDB1(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
        
    Const METHOD = "runInsPDB1"
    On Error GoTo ErrorLabel
   
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim Objbody As Object
    Dim ObjInsulatedBody As Object
    Dim parA As Double
    Dim parB As Double
    Dim parC As Double
    Dim parInsulationThickness As Double
    Dim parAngle As Double
   
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parA = arrayOfInputs(2)
    parB = arrayOfInputs(3)
    parC = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    If UBound(arrayOfInputs) >= 6 Then 'Not to fail the prior to V7 instances by accessing index 6
        parAngle = arrayOfInputs(6)
    Else
        parAngle = PI / 2 'To suppport V6.1 instances
    End If

    iOutput = 0

' Insert your code for output 2(Insulated Body)
    Dim HeaderOD As Double 'Header Outside pipe diameter
    
    'get the Header data
    RetrieveParameters 1, oPartFclt, m_OutputColl, HeaderOD, _
                        flangeThick, flangeDiam, cptOffset, depth
   
    Dim InsuPoint         As New AutoMath.DPosition
    Dim FacePoint         As New AutoMath.DPosition
    FacePoint.Set 0, HeaderOD / 2 + parA, 0
    
    Dim BranchOD As Double 'Branch Outside pipe diameter
    'get the Branch data

    RetrieveParameters 2, oPartFclt, m_OutputColl, BranchOD, _
                        flangeThick, flangeDiam, cptOffset, depth
    Dim CentertoInsu As Double

    'CentertoInsu = (0.5 * HeaderOD) * Cos(40# / RAD)
    CentertoInsu = -(parA / 2) * Cos(40# / RAD)
    InsuPoint.Set 0, CentertoInsu, 0
    Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, InsuPoint, FacePoint, _
                                parB + parInsulationThickness * 2, True)
                                
    Dim oDirVector As AutoMath.DVector
    Dim oTransformationMat  As New AutoMath.DT4x4
    Dim dRotAboutZaxis As Double
    Set oDirVector = New AutoMath.DVector
    
    oTransformationMat.LoadIdentity
    dRotAboutZaxis = PI / 2 - parAngle
    oDirVector.Set 0, 0, -1
    oTransformationMat.Rotate dRotAboutZaxis, oDirVector
    ObjInsulatedBody.Transform oTransformationMat
   
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    Set oDirVector = Nothing
    Set oTransformationMat = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
End Sub

Private Sub runInsPDB1onFitting(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
        
    Const METHOD = "runInsPDB1"
    On Error GoTo ErrorLabel
   
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim Objbody As Object
    Dim ObjInsulatedBody As Object
    Dim parA As Double
    Dim parB As Double
    Dim parC As Double
    Dim parInsulationThickness As Double
    Dim parAngle As Double
   
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parA = arrayOfInputs(2)
    parB = arrayOfInputs(3)
    parC = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    If UBound(arrayOfInputs) >= 6 Then 'Not to fail the prior to V7 instances by accessing index 6
        parAngle = arrayOfInputs(6)
    Else
        parAngle = PI / 2 'To suppport V6.1 instances
    End If

    iOutput = 0

' Insert your code for output 2(Insulated Body)
    Dim HeaderOD As Double 'Header Outside pipe diameter
    
    'get the Header data
    RetrieveParameters 1, oPartFclt, m_OutputColl, HeaderOD, _
                        flangeThick, flangeDiam, cptOffset, depth
   
    Dim InsuPoint         As New AutoMath.DPosition
    Dim FacePoint         As New AutoMath.DPosition
    FacePoint.Set 0, parA / 2, 0
    
    Dim BranchOD As Double 'Branch Outside pipe diameter
    'get the Branch data

    RetrieveParameters 2, oPartFclt, m_OutputColl, BranchOD, _
                        flangeThick, flangeDiam, cptOffset, depth
    Dim CentertoInsu As Double

    CentertoInsu = (0.5 * HeaderOD) * Cos(40# / RAD)
    InsuPoint.Set 0, CentertoInsu, 0
    Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, InsuPoint, FacePoint, _
                                parB + parInsulationThickness * 2, True)
                                
    Dim oDirVector As AutoMath.DVector
    Dim oTransformationMat  As New AutoMath.DT4x4
    Dim dRotAboutZaxis As Double
    Set oDirVector = New AutoMath.DVector
    
    oTransformationMat.LoadIdentity
    dRotAboutZaxis = PI / 2 - parAngle
    oDirVector.Set 0, 0, -1
    oTransformationMat.Rotate dRotAboutZaxis, oDirVector
    ObjInsulatedBody.Transform oTransformationMat
   
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    Set oDirVector = Nothing
    Set oTransformationMat = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
End Sub


