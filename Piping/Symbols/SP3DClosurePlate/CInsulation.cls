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
'   Copyright (c) 2004-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         ACM
'   Creation Date:  Monday, July 12, 2004
'   Description:
'       This is a Closure Palte symbol.  Sketch is taken from Design document.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   08.SEP.2006     KKC     DI-95670  Replace names with initials in all revision history sheets and symbols
'   15.Feb.2008     RUK         CR-33401  Create symbols, as required, for jacketed piping specs
'                               Sketch is taken from design document appendix E 105
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Insulation:"    'Used for error messages

Private Sub Class_Terminate()
    '''
End Sub

Private Sub Class_Initialize()
    '''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim pipeDiam As Double
    Dim flangeThick As Double
    Dim cptOffset As Double
    Dim flangeDiam As Double
    Dim depth As Double

'    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double
'    Dim parPlateInsideDia As Double
'    Dim parPlateThick As Double
'    Dim parJackPipeOffset As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
'    parInsulationThickness = arrayOfInputs(3)
'    parPlateOutsideDia = arrayOfInputs(3)
'    parPlateInsideDia = arrayOfInputs(4)
'    parPlateThick = arrayOfInputs(5)
'    parJackPipeOffset = arrayOfInputs(6)
    
    Dim iOutput As Double
    iOutput = 0
     
     'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    Dim dPlateThick As Double
    Dim dPlateOutSideDia As Double
    Dim dPlateInsideDia As Double
    Dim dJackPipeOffset As Double
    Dim dInsDia As Double
    Dim dStartX As Double
    Dim dEndX As Double
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    
    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                        cptOffset, depth

    Select Case lPartDataBasis
    Case Is <= 1
        Dim parFacetoFace As Double
        
        'Inputs
        parFacetoFace = arrayOfInputs(2)
        parInsulationThickness = arrayOfInputs(3)
        
        dPlateOutSideDia = pipeDiam
        dPlateInsideDia = 0.5 * dPlateOutSideDia
        dPlateThick = parFacetoFace
        dJackPipeOffset = parFacetoFace / 2
        dInsDia = BodyOD(pipeDiam, flangeDiam) + 2 * parInsulationThickness
        
        oStPoint.Set parFacetoFace / 2, 0, 0
        oEnPoint.Set -(parFacetoFace / 2 + parInsulationThickness), 0, 0
    Case 817, 818   'Jacket closure plate, outside diameter fits flush with inside diameter of jacket pipe (817).
                    'Jacket closure plate, outside diameter larger than outside diameter of jacket pipe (818).
        Dim parPlateOutsideDia As Double
        Dim parPlateInsideDia As Double
        Dim parPlateThick As Double
        Dim parJackPipeOffset As Double
        
        'Inputs
        parInsulationThickness = arrayOfInputs(3)
        parPlateOutsideDia = arrayOfInputs(4)
        parPlateInsideDia = arrayOfInputs(5)
        parPlateThick = arrayOfInputs(6)
        parJackPipeOffset = arrayOfInputs(7)
        
        If CmpDblEqual(parPlateOutsideDia, 0) Then
            Dim dPipeOD As Double
            Dim dPipeWallThick As Double
            RetrievePipeOD 1, oPartFclt, m_OutputColl, dPipeOD
            RetrievePipeWallThick 1, oPartFclt, dPipeWallThick
            dPlateOutSideDia = dPipeOD - 2 * dPipeWallThick
        Else
            dPlateOutSideDia = parPlateOutsideDia
        End If
        If CmpDblEqual(parPlateInsideDia, 0) Then
            dPlateInsideDia = 0.5 * parPlateOutsideDia
        Else
            dPlateInsideDia = parPlateInsideDia
        End If
        dPlateThick = parPlateThick
        If lPartDataBasis = 817 Then
            dJackPipeOffset = parJackPipeOffset
            oStPoint.Set 0, 0, 0
            oEnPoint.Set dJackPipeOffset - dPlateThick - parInsulationThickness, 0, 0
        ElseIf lPartDataBasis = 818 Then
            dJackPipeOffset = 0
            oStPoint.Set parInsulationThickness, 0, 0
            oEnPoint.Set -(dPlateThick + parInsulationThickness), 0, 0
        End If
        dInsDia = dPlateOutSideDia + 2 * parInsulationThickness
    Case Else
        GoTo ErrorLabel
    End Select
    
    'Create the Insulation for closure plate body
    Dim objInsBody As Object
    Set objInsBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsBody
    Set objInsBody = Nothing
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
    Err.HelpFile, Err.HelpContext
End Sub



