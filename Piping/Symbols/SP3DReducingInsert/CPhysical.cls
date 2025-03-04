VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          svsmylav
'   Creation Date:  Tuesday, Aug 27 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The Details were taken from PDS Piping Component Data - Reference Guide at Page no D-52
'   SN=F43_2 and MC=BUSH,INSR1,INSR2. Symbol is created with Six Outputs - three are Insulation aspect
'   outputs. Others are physical aspect outputs and corresponding function are as follows:
'   Flange - using 'PlaceCylinder',
'   Nozzle1  - using 'CreateNozzleWithLength' and
'   Nozzle2  - using 'CreateNozzle' function.
'   Symbol is created with three Insulation aspect
'   outputs which are prepared by using 'PlaceCylinder' function.
'
'   Change History:
'   dd.mmm.yyyy     who            change description
'   -----------     -----          ------------------
'   11.Dec.2003   svsmylav  TR-52642:   The input parameters are changed
'                                       to match with the prepared PDS data. One physical output
'                                       'ObjFlange' is removed, so number of outputs are Five.
'   05.Oct.2005   svsmylav  TR-85849:   Port locations are corrected.
'   08.SEP.2006   KKC       DI-95670    Replace names with initials in all revision history sheets and symbols
'   16.JUL.2008   PK        CR-145835   Implemented part data basis options 1020,1021 and 1022
'   16.Sep.2008   MP        CR-149885   Implemented partdata basis option 1083
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

    Dim iOutput                 As Double
    Dim parFacetoFace           As Double
    Dim parInsulationThickness  As Double
    Dim parLength1              As Double
    Dim parLength2              As Double
    Dim parWidth                As Double
    
    Dim stPoint                 As AutoMath.DPosition
    Dim enPoint                 As AutoMath.DPosition
    Dim dNozzleLength           As Double
    Dim dFacetoFace             As Double
    Dim dFacetoCenter           As Double
    
    'Inputss
    Set oPartFclt = arrayOfInputs(1)
    'parFacetoFace = arrayOfInputs(2)  'O
    parInsulationThickness = arrayOfInputs(3)
    'parLength1 = arrayOfInputs(4)
    'parLength2 = arrayOfInputs(5)
    'parWidth = arrayOfInputs(6)
    
    Dim lPartdatabasis As Long
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    lPartdatabasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
        
    If lPartdatabasis <= 1 Or lPartdatabasis = 1023 Or lPartdatabasis = 1083 Then
        parFacetoFace = arrayOfInputs(2)
        dFacetoFace = parFacetoFace
        If lPartdatabasis = 1083 Then
            parLength1 = arrayOfInputs(4)
            parWidth = arrayOfInputs(6)
        End If
    ElseIf lPartdatabasis = 1020 Or lPartdatabasis = 1021 Or lPartdatabasis = 1022 Then
        parLength1 = arrayOfInputs(4)
        parLength2 = arrayOfInputs(5)
        parWidth = arrayOfInputs(6)
        dFacetoFace = parLength1 + parLength2
    End If
            
    iOutput = 0
    
    'Origin location is considered at half of the port1 to port2 distance for partdatabasis 1023,1020,1021,1022
    Dim flangeThick2    As Double
    Dim sptOffset2      As Double
    Dim depth2          As Double
    Dim pipeDiam2       As Double
    Dim flangeDiam2     As Double

    'Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset2, depth2

    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    dFacetoCenter = (dFacetoFace - depth2) * 0.5
    oPlacePoint.Set -dFacetoCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    
    If lPartdatabasis <= 1 Or lPartdatabasis = 1023 Or lPartdatabasis = 1083 Then
        If lPartdatabasis = 1083 Then
            dFacetoCenter = dFacetoFace / 2
            oPlacePoint.Set -dFacetoFace / 2 - sptOffset + depth, 0, 0
            dNozzleLength = parLength1
        Else
            dNozzleLength = dFacetoFace - flangeThick2
        End If
        If dNozzleLength < flangeThick Then
            dNozzleLength = flangeThick
        End If
    ElseIf lPartdatabasis = 1020 Or lPartdatabasis = 1021 Or lPartdatabasis = 1022 Then
        dNozzleLength = parLength1
    End If
    
    
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    'Place Nozzle 2
    If lPartdatabasis <= 1 Or lPartdatabasis = 1023 Then
        oPlacePoint.Set (dFacetoFace - dFacetoCenter) + sptOffset2 - depth2, 0, 0
        oDir.Set 1, 0, 0

        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    ElseIf lPartdatabasis = 1020 Or lPartdatabasis = 1021 Or lPartdatabasis = 1022 Then
        oPlacePoint.Set (dFacetoFace - dFacetoCenter) + sptOffset2 - depth2, 0, 0
        oDir.Set 1, 0, 0

        Set objNozzle = CreateNozzleJustaCircle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    ElseIf lPartdatabasis = 1083 Then
        oPlacePoint.Set (parFacetoFace * 0.5) + sptOffset2 - depth2, 0, 0
        oDir.Set 1, 0, 0

        dNozzleLength = parFacetoFace - parLength1
        Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

    End If
            
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
        
    If lPartdatabasis = 1020 Or lPartdatabasis = 1021 Or lPartdatabasis = 1022 Or lPartdatabasis = 1083 Then
        
        'Create the female part
        Dim oHexHead    As Object
        Dim lNumSides   As Long
        Dim oTransMat   As AutoMath.DT4x4
        Dim oVec        As AutoMath.DVector
        
        Set oVec = New AutoMath.DVector
        Set oTransMat = New AutoMath.DT4x4
        
        If lPartdatabasis = 1020 Or lPartdatabasis = 1083 Then
            lNumSides = 6
        ElseIf lPartdatabasis = 1021 Then
            lNumSides = 8
        ElseIf lPartdatabasis = 1022 Then
            lNumSides = 10
        End If
        
        Dim dSideLength As Double
        If CmpDblEqual(parWidth, 0) Then
            'Assuming that side length of the polygon shaped hub is 0.7 times the Piping Outside Diameter of Port1
            dSideLength = 0.7 * pipeDiam
        Else
            dSideLength = parWidth * Sin(360 / (2 * lNumSides))
        End If
        If lPartdatabasis = 1083 Then
            Dim dThick As Double
            dThick = 0.2 * parLength1
            Set oHexHead = PlaceNnagon(m_OutputColl, lNumSides, dSideLength, dThick, True)
            oVec.Set parFacetoFace * 0.5 - (parFacetoFace - parLength1) - dThick, 0, 0

        Else
            Set oHexHead = PlaceNnagon(m_OutputColl, lNumSides, dSideLength, parLength2, True)
            oVec.Set parLength1 - dFacetoCenter, 0, 0
        End If
            oTransMat.LoadIdentity
            oTransMat.Translate oVec
            oHexHead.Transform oTransMat
            m_OutputColl.AddOutput "HexagonalHead", oHexHead
            Set oHexHead = Nothing
    End If
    Set oVec = Nothing
    Set oTransMat = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
 
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

