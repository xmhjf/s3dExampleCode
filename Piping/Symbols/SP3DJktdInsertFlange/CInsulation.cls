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
'   Creation Date:  Monday, July 14, 2004
'   Description:
'       This is a Jacketed Piping Insert Flange Symbol.
'       On One side it has,One for Jacketed Pipe and One port of core pipe.
'       Other side it has one port.
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC      DI-95670  Replace names with initials in all revision history sheets and symbols
'  01-Feb.2008     RRK      CR-33401  Create symbols, as required, for jacketed piping specs
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Terminate()
    '
End Sub

Private Sub Class_Initialize()
    '
    
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


    Dim dPort3Location As Double
    
    Dim oStPoint   As Object 'Used in this symbol to locate back of the flange
    Dim oEnPoint   As Object
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
'    parGasketFacetoPipe = arrayOfInputs(4)

    Dim lPartdatabasis   As Long
    Dim iOutput     As Double
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    lPartdatabasis = oPipeComponent.PartDataBasis
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth, False
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick2, flangeDiam, cptOffset2, depth2
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, flangeDiam, cptOffset3, depth3
    
    iOutput = 0
    
    Select Case lPartdatabasis
    Case Is <= 1, 801, 803, 807, 809 'Non-Reducing cases where length through hub inclusive of flange face projection and
                                     'length through hub exclusive of flange face projection
        parFacetoFace = arrayOfInputs(2)
        parGasketFacetoPipe = arrayOfInputs(4)
    
        dPort3Location = parGasketFacetoPipe
        
        'Location for the back of the flange
        oStPoint.Set flangeThick, 0, 0
    Case 802  'Insert flange, jacket, slip-on, non-reducing, length through hub inclusive of flange face projection, flange thickness exclusive of flange face projection
        parFacetoFace = arrayOfInputs(2)
        parGasketFacetoPipe = arrayOfInputs(4)
        
        dPort3Location = parGasketFacetoPipe
        
        'Location for the back of the flange
        oStPoint.Set flangeThick + cptOffset, 0, 0
        
    Case 804, 806, 810 'Insert flange, jacket, slip-on, reducing, length through hub inclusive of flange face projection (804)
                       'Insert flange, jacket, slip-on, reducing, length through hub exclusive of flange face projection (806)
                       'Insert flange, jacket, buttweld, reducing, length through hub inclusive of flange face projection (810)
        parFacetoFace = arrayOfInputs(2)
        
        dPort3Location = flangeThick
        
        'Location for the back of the flange
        oStPoint.Set flangeThick, 0, 0
        
    Case 805 'Insert flange, jacket, slip-on, reducing, length through hub inclusive of flange face projection, flange thickness exclusive of flange face projection
        parFacetoFace = arrayOfInputs(2)
        
        dPort3Location = flangeThick + cptOffset
        
        'Location for the back of the flange
        oStPoint.Set flangeThick + cptOffset, 0, 0
    
    Case 808  'Insert flange, jacket, buttweld, non-reducing, length through hub inclusive of flange face projection, flange thickness exclusive of flange face projection
        parFacetoFace = arrayOfInputs(2)
        parGasketFacetoPipe = arrayOfInputs(4)
        
        dPort3Location = parGasketFacetoPipe
        
        'Location for the back of the flange
        oStPoint.Set flangeThick + cptOffset, 0, 0
        
    Case 811  'Insert flange, jacket, buttweld, reducing, length through hub inclusive of flange face projection, flange thickness exclusive of flange face projection
        parFacetoFace = arrayOfInputs(2)
        
        dPort3Location = flangeThick + cptOffset
        
        'Location for the back of the flange
        oStPoint.Set flangeThick + cptOffset, 0, 0
        
    Case 812  'Insert flange, jacket, buttweld, reducing, length through hub exclusive of flange face projection
        parFacetoFace = arrayOfInputs(2)
        
        dPort3Location = flangeThick
        
        'Location for the back of the flange
        oStPoint.Set flangeThick, 0, 0
    Case Else
        GoTo ErrorLabel:
    End Select
    
' Place Insert Body 1(Output 1)
    Select Case lPartdatabasis
    Case 801, 802, 803, 807, 808, 809 'Creating the body only for non-reducing cases
        Dim ObjInsertBody1Ins As Object
        Dim dInsHubDia1 As Double
        dInsHubDia1 = pipeDiam3 + (2 * parInsulationThickness)
        
        oEnPoint.Set dPort3Location, 0, 0
        
        'Check to see that for the cases where gasket face to pipe is same as flange thickness plus flange face projection symbol do not fail
        If CmpDblEqual(oStPoint.x, oEnPoint.x) Then
            oEnPoint.Set oStPoint.x + 0.0001, 0, 0
        End If
        
        Set ObjInsertBody1Ins = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsHubDia1, True)
        
        ' Set the output
        m_OutputColl.AddOutput "InsertBodyIns", ObjInsertBody1Ins
        Set ObjInsertBody1Ins = Nothing
    End Select
    
' Place Insert Body 2(Output 2)
    Dim oPipePort  As IJCatalogPipePort
    Dim oCollection As IJDCollection
    Dim pPortIndex As Integer
    Set oCollection = oPartFclt.GetNozzles()
    Set oPipePort = oCollection.Item(3)

    Dim dWallThickness As Double
    dWallThickness = oPipePort.WallThicknessOrGrooveSetback

    Dim ObjInsertBody2Ins As Object
    Dim dInsHubDia2 As Double
    dInsHubDia2 = pipeDiam3 - (2 * dWallThickness) + 2 * parInsulationThickness

    oEnPoint.Set parFacetoFace, 0, 0

    Set ObjInsertBody2Ins = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsHubDia2, True)

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsertBody2Ins
    Set ObjInsertBody2Ins = Nothing

 ' Insert your code for output 1(Insulation for port 3)
    
    Dim ObjInsPort1 As Object
    Dim dInsulationDiameter As Double
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth
            
    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    
    Dim dFlangeBackPos As Double
    dFlangeBackPos = oStPoint.x
    
    oStPoint.Set dFlangeBackPos - flangeThick, 0, 0
    
    Dim dPortInsEndPos As Double
    dPortInsEndPos = dFlangeBackPos + parInsulationThickness
    
    'Check for insluation length not to exceed face to face
    If CmpDblGreaterthanOrEqualTo(dPortInsEndPos, parFacetoFace) Then
        dPortInsEndPos = parFacetoFace
    End If
    
    oEnPoint.Set dPortInsEndPos, 0, 0
    
    dInsulationDiameter = flangeDiam + parInsulationThickness * 2
    
    If CmpDblGreaterthan(pipeDiam, flangeDiam) Then dInsulationDiameter = pipeDiam + parInsulationThickness * 2
    Set ObjInsPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDiameter, True)
        

    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort1
    Set ObjInsPort1 = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
     
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
    
End Sub


