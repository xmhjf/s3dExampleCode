VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         ACM
'   Creation Date:  Monday, July 14, 2004
'   Description:
'       This is a Jacketed Piping Insert Flange Symbol.
'       The symbol has three ports.Port 1 is flanged end, port 2 is the other end for core pipe
'       and port 3 is for jacket piping
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----    -----------------
'  08.SEP.2006     KKC      DI-95670  Replace names with initials in all revision history sheets and symbols
'  01-Feb.2008     RRK      CR-33401  Create symbols, as required, for jacketed piping specs
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
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim cptOffset2       As Double
    Dim depth2           As Double
    Dim flangeThick2     As Double
    
    Dim flangeThick3     As Double
    Dim pipeDiam3        As Double
    Dim cptOffset3       As Double
    Dim depth3           As Double
    
    Dim parFacetoFace As Double
    Dim parInsulationThickness As Double
    Dim parGasketFacetoPipe As Double

    Dim dPort1Location As Double
    Dim dPort2Location As Double
    Dim dPort3Location As Double
    
    Dim oStPoint   As Object
    Dim oEnPoint   As Object
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
'    parGasketFacetoPipe = arrayOfInputs(4)

    Dim lPartdatabasis   As Long
    Dim dCompanyPracticeGap As Double
    Dim iOutput     As Double
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    lPartdatabasis = oPipeComponent.PartDataBasis
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth, False
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick2, flangeDiam, cptOffset2, depth2
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, flangeDiam, cptOffset3, depth3
    
    iOutput = 0
    Select Case lPartdatabasis
    Case Is <= 1, 801 'Insert flange, jacket, slip-on, non-reducing, length through hub inclusive of flange face projection
        parFacetoFace = arrayOfInputs(2)
        parGasketFacetoPipe = arrayOfInputs(4)
        
        'Retrieving company practice gap
        oPipeComponent.GetCompanyPracticeGap dCompanyPracticeGap
        
        dPort1Location = 0
        dPort2Location = dCompanyPracticeGap
        dPort3Location = parGasketFacetoPipe
        
        oStPoint.Set flangeThick, 0, 0
    Case 802 'Insert flange, jacket, slip-on, non-reducing, length through hub inclusive of flange face projection, flange thickness exclusive of flange face projection
        parFacetoFace = arrayOfInputs(2)
        parGasketFacetoPipe = arrayOfInputs(4)
        
        'Retrieving company practice gap
        oPipeComponent.GetCompanyPracticeGap dCompanyPracticeGap
        
        dPort1Location = 0
        dPort2Location = dCompanyPracticeGap
        dPort3Location = parGasketFacetoPipe
        
        oStPoint.Set flangeThick + cptOffset, 0, 0
    
    Case 803 'Insert flange, jacket, slip-on, non-reducing, length through hub exclusive of flange face projection
        parFacetoFace = arrayOfInputs(2)
        parGasketFacetoPipe = arrayOfInputs(4)
        
        'Retrieving company practice gap
        oPipeComponent.GetCompanyPracticeGap dCompanyPracticeGap
        
        dPort1Location = cptOffset
        dPort2Location = dCompanyPracticeGap - cptOffset
        dPort3Location = parGasketFacetoPipe
        
        oStPoint.Set flangeThick, 0, 0
    Case 804  'Insert flange, jacket, slip-on, reducing, length through hub inclusive of flange face projection
        parFacetoFace = arrayOfInputs(2)
        
        'Retrieving company practice gap
        oPipeComponent.GetCompanyPracticeGap dCompanyPracticeGap
        
        dPort1Location = 0
        dPort2Location = dCompanyPracticeGap
        dPort3Location = flangeThick
        
        oStPoint.Set flangeThick, 0, 0
    
    Case 805  'Insert flange, jacket, slip-on, reducing, length through hub inclusive of flange face projection, flange thickness exclusive of flange face projection

        parFacetoFace = arrayOfInputs(2)
        
        'Retrieving company practice gap
        oPipeComponent.GetCompanyPracticeGap dCompanyPracticeGap
        
        dPort1Location = 0
        dPort2Location = dCompanyPracticeGap
        dPort3Location = flangeThick + cptOffset
        
        oStPoint.Set flangeThick + cptOffset, 0, 0
    
    Case 806  'Insert flange, jacket, slip-on, reducing, length through hub exclusive of flange face projection
        parFacetoFace = arrayOfInputs(2)
    
        'Retrieving company practice gap
        oPipeComponent.GetCompanyPracticeGap dCompanyPracticeGap
        
        dPort1Location = cptOffset
        dPort2Location = dCompanyPracticeGap - cptOffset
        dPort3Location = flangeThick
        
        oStPoint.Set flangeThick, 0, 0
                   
    Case 807 'Insert flange, jacket, buttweld, non-reducing, length through hub inclusive of flange face projection

        parFacetoFace = arrayOfInputs(2)
        parGasketFacetoPipe = arrayOfInputs(4)
        
        dPort1Location = 0
        dPort2Location = parFacetoFace + cptOffset2 - depth2
        dPort3Location = parGasketFacetoPipe
        
        oStPoint.Set flangeThick, 0, 0
    Case 808 'Insert flange, jacket, buttweld, non-reducing, length through hub inclusive of flange face projection
        
        parFacetoFace = arrayOfInputs(2)
        parGasketFacetoPipe = arrayOfInputs(4)
        
        dPort1Location = 0
        dPort2Location = parFacetoFace + cptOffset2 - depth2
        dPort3Location = parGasketFacetoPipe
        
        oStPoint.Set flangeThick + cptOffset, 0, 0
    Case 809  'Insert flange, jacket, buttweld, non-reducing, length through hub exclusive of flange face projection
        parFacetoFace = arrayOfInputs(2)
        parGasketFacetoPipe = arrayOfInputs(4)
        
        dPort1Location = cptOffset
        dPort2Location = parFacetoFace + cptOffset2 - depth2
        dPort3Location = parGasketFacetoPipe
        
        oStPoint.Set flangeThick, 0, 0
        

    Case 810 'Insert flange, jacket, buttweld, reducing, length through hub inclusive of flange face projection
        parFacetoFace = arrayOfInputs(2)
        
        dPort1Location = 0
        dPort2Location = parFacetoFace + cptOffset2 - depth2
        dPort3Location = flangeThick
        
        oStPoint.Set flangeThick, 0, 0
    
    Case 811 'Insert flange, jacket, buttweld, reducing, length through hub inclusive of flange face projection, flange thickness exclusive of flange face projection
        parFacetoFace = arrayOfInputs(2)
    
        dPort1Location = 0
        dPort2Location = parFacetoFace + cptOffset2 - depth2
        dPort3Location = flangeThick + cptOffset
        
        oStPoint.Set flangeThick + cptOffset, 0, 0
        
    Case 812 'Insert flange, jacket, buttweld, reducing, length through hub exclusive of flange face projection
        parFacetoFace = arrayOfInputs(2)
        
        dPort1Location = cptOffset
        dPort2Location = parFacetoFace + cptOffset2 - depth2
        dPort3Location = flangeThick
        
        oStPoint.Set flangeThick, 0, 0
        
    Case Else
        GoTo ErrorLabel:
    End Select
    
' Place Insert Body 1(Output 1)
    
    Select Case lPartdatabasis
    Case 801, 802, 803, 807, 808, 809 'Creating the body only for non-reducing cases
        Dim ObjInsertBody1 As Object
        Dim dHubDia1 As Double
        dHubDia1 = pipeDiam3
        
        oEnPoint.Set dPort3Location, 0, 0
        
        'Check to see that for the cases where gasket face to pipe is same as flange thickness plus flange face projection symbol do not fail
        If CmpDblEqual(oStPoint.x, oEnPoint.x) Then
            oEnPoint.Set oStPoint.x + 0.0001, 0, 0
        End If

        Set ObjInsertBody1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dHubDia1, True)
        
        ' Set the output
        m_OutputColl.AddOutput "InsertBody1", ObjInsertBody1
        Set ObjInsertBody1 = Nothing
    End Select
    
' Place Insert Body 2(Output 2)
    Dim oPipePort  As IJCatalogPipePort
    Dim oCollection As IJDCollection
    Dim pPortIndex As Integer
    Set oCollection = oPartFclt.GetNozzles()
    Set oPipePort = oCollection.Item(3)
    
    Dim dWallThickness As Double
    dWallThickness = oPipePort.WallThicknessOrGrooveSetback
    
    Dim ObjInsertBody2 As Object
    Dim dHubDia2 As Double
    dHubDia2 = pipeDiam3 - (2 * dWallThickness)
    
    oEnPoint.Set parFacetoFace, 0, 0
    
    Set ObjInsertBody2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dHubDia2, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsertBody2
    Set ObjInsertBody2 = Nothing
    
' Place Raised Face body
    'Creating the raised face portion based on part data basis values.It is not required where flange thickness includes face projection
    Select Case lPartdatabasis
    Case 802, 803, 805, 806, 808, 809, 811, 812
        Dim ObjRaisedFaceCyl As Object
        Dim dRaisedFaceDiam As Double
        
        Set oPipePort = oCollection.Item(1)
        dRaisedFaceDiam = oPipePort.RaisedFaceOrSocketDiameter
        
        oStPoint.Set -dPort1Location, 0, 0
        oEnPoint.Set -dPort1Location + cptOffset, 0, 0
        
        Set ObjRaisedFaceCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dRaisedFaceDiam, True)
        
        ' Set the output
        m_OutputColl.AddOutput "RaisedFaceCyl", ObjRaisedFaceCyl
        Set ObjRaisedFaceCyl = Nothing
    End Select
    
' Place Nozzle 1
    
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim oPlacePoint As AutoMath.DPosition
    
    Set oDir = New AutoMath.DVector
    Set oPlacePoint = New AutoMath.DPosition
    
    oPlacePoint.Set -dPort1Location, 0, 0
    oDir.Set -1, 0, 0
    
    Set oPipeComponent = Nothing

    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2  'Core Pipe nozzle

    
    oPlacePoint.Set dPort2Location, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzleJustaCircle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 3 ''Jacket Pipe nozzle

    oPlacePoint.Set dPort3Location, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
    
End Sub

