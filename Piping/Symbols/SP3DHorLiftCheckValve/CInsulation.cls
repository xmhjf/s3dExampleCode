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
'   Author:          ACM
'   Creation Date:  Tuesday, Nov 26 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Horizontal Lift Check Valve, Taken from www.resistoflex.com web catalog.
'   This symbol has total 11 outputs among them 9 are Physical and 2 are Insulation.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   02.Jan.2006     KKK       CR-91106  Updated the symbol to be based on either face-to-face
'                                           or face-to-center(Implementation of PartDataBasis Property)
'   31.Mar.2006     SymbolTeam(India)       Modified to accept default value of lPartDataBasis to be less than or equal to 1.
'   25.sep.06       KKC            TR-102869 Modified symbol code to use Face to Center or Face to Face to Dimension.
'                                            Modified symbol code to use Face to Face as a input parameter.
'   17.Jul.2007     MA             CR-123497 Implemented Part data basis to use Face 1 to Center and Face 2 to Center Dimensions.
'   07.Sept.2007    RRK            CR-126919  Enhanced horizontal lift check valve to support Y pattern by adding three part data basis values(385,386 and 387)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Dim PI As Double
Const NEGLIGIBLE_THICKNESS = 0.0001
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
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
    Dim ObjBonetIns As Object
    Dim ObjBodyIns As Object
    
    Dim parFacetoFace As Double
    Dim parValveHeight As Double
    Dim parInsulationThickness As Double
    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parBranchLength As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
'    parValveHeight = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
'    parFacetoCenter=arrayOfInputs(5)
'    parFace1toCenter = arrayOfInputs(6)
'    parFace2toCenter = arrayOfInputs(7)
'    parBranchLength = arrayOfInputs(8)

    iOutput = 0
 ' Insert your code for output 8(Bonet Insulation)
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Dim InsDia As Double
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
 'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    Dim dFace1toCenter As Double
    Dim dFace2toCenter As Double

' This symbol is based on the following Part data basis values that govern its geometry
'   (Part data Basis value -5): Face-to-face dimension basis
'   (Part data Basis value -10): Face-to-center dimension basis
'   (Part data Basis value -13): Asymmetrical Face-to-Center dimension basis
'   (Part data Basis value -385): Horizontal lift check valve, Y pattern, face-to-face dimension basis
'   (Part data Basis value -386): Horizontal lift check valve, Y pattern, face-to-center dimension basis
'   (Part data Basis value -387): Horizontal lift check valve, Y pattern, asymmetrical face-to-center dimension basis
    
    Select Case lPartDataBasis
    
    Case Is <= 1, 5
        parFacetoFace = arrayOfInputs(2)
        parValveHeight = arrayOfInputs(3)
        
        dFace1toCenter = parFacetoFace / 2
        dFace2toCenter = parFacetoFace / 2
    
    Case 10
        parValveHeight = arrayOfInputs(3)
        parFacetoCenter = arrayOfInputs(5)
        
        dFace1toCenter = parFacetoCenter
        dFace2toCenter = parFacetoCenter
    
    Case 13
        parValveHeight = arrayOfInputs(3)
        parFace1toCenter = arrayOfInputs(6)
        parFace2toCenter = arrayOfInputs(7)
        
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter
    
    Case 385
        parFacetoFace = arrayOfInputs(2)
        parBranchLength = arrayOfInputs(8)
        dFace1toCenter = parFacetoFace / 2
        dFace2toCenter = parFacetoFace / 2

    Case 386
        parFacetoCenter = arrayOfInputs(5)
        parBranchLength = arrayOfInputs(8)
        dFace1toCenter = parFacetoCenter
        dFace2toCenter = parFacetoCenter
        
    Case 387
        parFace1toCenter = arrayOfInputs(6)
        parFace2toCenter = arrayOfInputs(7)
        parBranchLength = arrayOfInputs(8)
        
        dFace1toCenter = parFace1toCenter
        dFace2toCenter = parFace2toCenter
    Case Else
        GoTo ErrorLabel:
        
    End Select

    Select Case lPartDataBasis
    Case Is <= 1, 5, 10, 13
'Default Horizontal Lift Check Valve Cases
       Dim dMinFacetoCenter As Double
       
       If CmpDblLessThan(dFace1toCenter, dFace2toCenter) Then
           dMinFacetoCenter = dFace1toCenter
       Else
           dMinFacetoCenter = dFace2toCenter
       End If
       
       InsDia = (Sqr(2) * ((dMinFacetoCenter) - (2 * flangeThick))) + 2 * parInsulationThickness
       
       oStPoint.Set 0, parValveHeight + parInsulationThickness, 0
       oEnPoint.Set 0, 0, 0
       
       Set ObjBonetIns = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, InsDia, True)
       
       ' Set the output
       iOutput = iOutput + 1
       m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBonetIns
       Set ObjBonetIns = Nothing
       
    Case 385, 386, 387
'Non-Default Horizontal Lift Check Valve Cases
      Dim objInsYBranch As Object
      
      Dim oCenterPos As AutoMath.DPosition
      Set oCenterPos = New AutoMath.DPosition
      
      Dim dAngle As Double
      Dim dAngle2 As Double
      Dim dTotAngle As Double
      Dim dBranchDia As Double
      Dim dLimBranchDia1 As Double
      Dim dLimBranchDia2 As Double
      Dim dLimLength1 As Double
      
      'Assumptions: Branch Angle is assumed to be equal to 45
      dAngle = PI / 4
      
      'Parameters useful for calculating limiting branch diameter
      dLimLength1 = Sqr((pipeDiam / 2) ^ 2 + (dFace2toCenter ^ 2))
      dAngle2 = Atn(0.5 * pipeDiam / dFace2toCenter)
      dTotAngle = dAngle + dAngle2
      
      'Calculating limiting branch diameter above which the top face of the branch crosses port 2
      dLimBranchDia1 = 2 * dLimLength1 * Cos(dTotAngle)
      'Calculating limiting branch diameter above which the bottom face of the branch crosses port 1
      dLimBranchDia2 = 2 * dFace1toCenter / Sin(dAngle)
      
      'Assumption for branch diameter equal to length of the line between ends of cylinder inclined at dAngle
      dBranchDia = pipeDiam / Cos(dAngle)
      
      'Checking with limiting branch diameter values
      If CmpDblGreaterthan(dBranchDia, dLimBranchDia1) Or CmpDblGreaterthan(dBranchDia, dLimBranchDia2) Then
        dBranchDia = 0.95 * SmallerDim(dLimBranchDia1, dLimBranchDia2)
      End If
    
      If lPartDataBasis = 385 Then 'Face to Face basis
        oCenterPos.Set -0.25 * parFacetoFace, 0, 0
      ElseIf lPartDataBasis = 386 Or lPartDataBasis = 387 Then 'Face to Center or Assymmetrical Face to Center basis
        oCenterPos.Set 0, 0, 0
      End If
      
      oStPoint.Set oCenterPos.x, oCenterPos.y, oCenterPos.z
      oEnPoint.Set oStPoint.x + parBranchLength * Cos(dAngle), parBranchLength * Sin(dAngle), 0
          
      Set objInsYBranch = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dBranchDia + (2 * parInsulationThickness), True)
    ' Set the output
      iOutput = iOutput + 1
      m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsYBranch
      Set objInsYBranch = Nothing
      Set oCenterPos = Nothing
    End Select
       
    ' Insert your code for output 9(Body Insulation)
    If CmpDblLessThanOrEqualTo(flangeDiam, -LINEAR_TOLERANCE) Then
       flangeDiam = NEGLIGIBLE_THICKNESS
    End If
    
    InsDia = flangeDiam + 2 * parInsulationThickness
    If CmpDblGreaterthan(pipeDiam, flangeDiam) Then InsDia = pipeDiam + 2 * parInsulationThickness
    
    oStPoint.Set -dFace1toCenter, 0, 0
    oEnPoint.Set dFace2toCenter, 0, 0
    
    Set ObjBodyIns = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, InsDia, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyIns
        
    Set ObjBodyIns = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oPipeComponent = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
