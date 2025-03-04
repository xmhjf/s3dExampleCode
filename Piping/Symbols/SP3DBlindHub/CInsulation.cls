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
'   Copyright (c) 2007-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         RUK
'   Creation Date:  Wednesday 25, July 2007
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This symbol consists of three physical outputs: two ports and a cone and one insulation output.
'   Source:
'       http://www.woodcousa.com/clp_r_f.htm
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----    ------------------
'  13.MAR.2008     MA   CR-136876 Implemeted Part Data Basis for the value 872.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI As Double
Const NEGLIGIBLE_THICKNESS = 0.0001

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

    Dim parHubBackFaceDia As Double
    Dim parBlindThick As Double
    Dim parInsulationThickness As Double
    Dim parBlindLength As Double
    Dim parBlindFaceDia As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parHubBackFaceDia = arrayOfInputs(2)
'    parBlindThick = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
'    parBlindLength = arrayOfInputs(5)
'    parBlindFaceDia = arrayOfInputs(6)
    
    iOutput = 0
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                            sptOffset, depth

    'Center is taken at port 1
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    ' This symbol is based on the following three Part data basis values that govern its geometry,
    '   Blind hub, specified by hub back face diameter and blind thickness (Default case)
    '   Blind hub, specified by blind length and blind face diameter - 872

    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Dim oAxis As AutoMath.DVector
    Dim objBlindHub As Object
    
    Select Case lPartDataBasis
        
        Case Is <= 1
            parHubBackFaceDia = arrayOfInputs(2)
            parBlindThick = arrayOfInputs(3)

            'Insert your code for output 1(BlindHub Body)
            'Create the Line string and revolve it about X-Axis
                       '2
                    '-------
                    '       ' 3
                    '        '  4
                 '1 '         ----
                    '             ' 5
            '_______'____________________X-Axis
                    '             '
                    '             '
                    '         -----
                    '       '
                    '      '
                    '------
                    
            Dim oLine As Object
            Dim oCollection As Collection
            Dim oComplexStr As IngrGeom3D.ComplexString3d
        
            Set oCollection = New Collection
            Set oComplexStr = New ComplexString3d
            Set oAxis = New AutoMath.DVector
        
            'Line 1
            oStPoint.Set 0, 0, 0
            oEnPoint.Set 0, (flangeDiam / 2) + parInsulationThickness, 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine
            Set oLine = Nothing
            
            Dim dTheta As Double
            dTheta = Atn(2 * (flangeDiam - parHubBackFaceDia) / (parBlindThick))
            
            'Line 2
            oStPoint.Set 0, (flangeDiam / 2) + parInsulationThickness, 0
            oEnPoint.Set flangeThick + Tan(dTheta / 2) * parInsulationThickness, _
                                (flangeDiam / 2) + parInsulationThickness, 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine
            Set oLine = Nothing
        
            'Line 3
            Dim dTempX As Double 'X Co-ordinate of the bottom of Inclined line
            dTempX = flangeThick + parBlindThick / 4
        
            oStPoint.Set flangeThick + Tan(dTheta / 2) * parInsulationThickness, _
                                (flangeDiam / 2) + parInsulationThickness, 0
            oEnPoint.Set dTempX + Tan(dTheta / 2) * parInsulationThickness, _
                                (parHubBackFaceDia / 2) + parInsulationThickness, 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine
            Set oLine = Nothing
        
            'Line 4
            oStPoint.Set dTempX + Tan(dTheta / 2) * parInsulationThickness, (parHubBackFaceDia / 2) + parInsulationThickness, 0
            oEnPoint.Set (flangeThick + parBlindThick + parInsulationThickness), _
                                (parHubBackFaceDia / 2) + parInsulationThickness, 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine
            Set oLine = Nothing
        
            'Line 5
            oStPoint.Set (flangeThick + parBlindThick + parInsulationThickness), _
                                (parHubBackFaceDia / 2) + parInsulationThickness, 0
            oEnPoint.Set (flangeThick + parBlindThick + parInsulationThickness), 0, 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine
            Set oLine = Nothing
            
            'Create the Complex String
            oStPoint.Set 0, 0, 0
            Set oComplexStr = PlaceTrCString(oStPoint, oCollection)
        
            'Revovle the Complex string about X-Axis
            oAxis.Set 1, 0, 0
            oStPoint.Set 0, 0, 0
            Set objBlindHub = PlaceRevolution(m_OutputColl, oComplexStr, _
                                                    oAxis, oStPoint, 2 * PI, True)
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBlindHub
            Set objBlindHub = Nothing
        
            'Remove the references
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            Set oAxis = Nothing
            Set oComplexStr = Nothing
            Dim iCount As Integer
            iCount = 1
            For iCount = 1 To oCollection.Count
                oCollection.Remove 1
            Next iCount
            Set oCollection = Nothing
            
        Case 872
        
            parBlindLength = arrayOfInputs(5)
            parBlindFaceDia = arrayOfInputs(6)
            
            Dim dLineStringPoints(0 To 17) As Double
            Dim oLineString As IngrGeom3D.LineString3d
            Dim oCenterPoint As AutoMath.DPosition
            Set oCenterPoint = New AutoMath.DPosition
            Set oAxis = New AutoMath.DVector
            Dim oGeomFactory As IngrGeom3D.GeometryFactory
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            
            'Insert your code for output 1(Body of Blind Hub)
            'Point 1
            dLineStringPoints(0) = flangeThick
            dLineStringPoints(1) = 0
            dLineStringPoints(2) = 0
            
            'Point 2
            dLineStringPoints(3) = flangeThick
            dLineStringPoints(4) = flangeDiam / 2 + parInsulationThickness
            dLineStringPoints(5) = 0
            
            'Point 3
            dLineStringPoints(6) = 0.4 * parBlindLength
            dLineStringPoints(7) = parBlindFaceDia / 2 + parInsulationThickness
            dLineStringPoints(8) = 0
            
            'Point 4
            dLineStringPoints(9) = parBlindLength + parInsulationThickness
            dLineStringPoints(10) = parBlindFaceDia / 2 + parInsulationThickness
            dLineStringPoints(11) = 0
            
            'Point 5
            dLineStringPoints(12) = parBlindLength + parInsulationThickness
            dLineStringPoints(13) = 0
            dLineStringPoints(14) = 0
            
            'Point 6
            dLineStringPoints(15) = flangeThick
            dLineStringPoints(16) = 0
            dLineStringPoints(17) = 0
            
            Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 6, dLineStringPoints)
            oAxis.Set 1, 0, 0
            oCenterPoint.Set 0, -0.00001, 0
            Set objBlindHub = PlaceRevolution(m_OutputColl, oLineString, oAxis, oCenterPoint, 2 * PI, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBlindHub
            Set objBlindHub = Nothing
            Set oCenterPoint = Nothing
            Set oAxis = Nothing
            Set oLineString = Nothing
            Set oGeomFactory = Nothing
            
            'Place Insulation for Port
            Dim dFlangeBodyThickness As Double
            If CmpDblGreaterthan(flangeDiam, LINEAR_TOLERANCE) Then
                If CmpDblEqual(flangeThick, 0) Then flangeThick = NEGLIGIBLE_THICKNESS
                oStPoint.Set 0, 0, 0
                dFlangeBodyThickness = flangeThick + parInsulationThickness
                If CmpDblGreaterthan(dFlangeBodyThickness, parBlindLength) Then dFlangeBodyThickness = parBlindLength
                oEnPoint.Set dFlangeBodyThickness, 0, 0
                
                Dim objInsulatedPort As Object
                Set objInsulatedPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                flangeDiam + 2 * parInsulationThickness, True)
        
                'Set the output
                m_OutputColl.AddOutput "InsPort", objInsulatedPort
                Set objInsulatedPort = Nothing
             End If
            
        Case Else
            GoTo ErrorLabel:      ' Invalid Part data basis.

    End Select

    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

