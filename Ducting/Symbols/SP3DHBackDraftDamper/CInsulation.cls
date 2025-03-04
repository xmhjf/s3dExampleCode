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
'   Copyright (c) 2007-08, Intergraph Corporation. All rights reserved.

'   CInsulation.cls
'   Author:         RRK
'   Creation Date:  Friday, Aug 17 2007
'   Description:
'   This symbol is prepared for Back draft damper that caters to Round and Rectangular shapes as per CR-124909

'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   02.SEP.2008     PK      CR-149330 Provided insualtion for the PDB cases 145,146 and 147
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Insulation:" 'Used for error messages
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput             As Double
    
    Dim parWidth                As Double
    Dim parDepth                As Double
    Dim parDamperLength         As Double
    Dim parLeverLength          As Double
    Dim parLeverPositionX       As Double
    Dim parLeverPositionY       As Double
    Dim parLeverPositionZ       As Double
    Dim parCounterWeightAngle   As Double
    Dim parHVACShape            As Double
    Dim parInsulationThickness  As Double
    Dim parControlLength        As Double
    Dim parControlWidth         As Double
    Dim parControlHeight        As Double
        
    Dim CornerRadius As Double
    CornerRadius = 0

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parWidth = arrayOfInputs(2)
    'parDepth = arrayOfInputs(3)
    'parDamperLength = arrayOfInputs(4)
    'parLeverLength = arrayOfInputs(5)
    'parLeverPositionX = arrayOfInputs(6)
    'parLeverPositionY = arrayOfInputs(7)
    'parLeverPositionZ = arrayOfInputs(8)
    'parCounterWeightAngle = arrayOfInputs(9)
    'parHVACShape = arrayOfInputs(10)
    parInsulationThickness = arrayOfInputs(11)
    'parControlHeight = arrayOfInputs(12)
    'parControlLength = arrayOfInputs(13)
    'parControlWidth = arrayOfInputs(14)
    
    'Checking for the Part Data Basis Property
    Dim lPartDataBasis As Long
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPartFclt
    lPartDataBasis = oHvacPart.PartDataBasis
    Set oHvacPart = Nothing
        
    Select Case lPartDataBasis
    
       Case Is <= 1, 145                'Default or Back draft damper with counter weight
           parWidth = arrayOfInputs(2)
           parDepth = arrayOfInputs(3)
           parDamperLength = arrayOfInputs(4)
           parLeverLength = arrayOfInputs(5)
           parLeverPositionX = arrayOfInputs(6)
           parLeverPositionY = arrayOfInputs(7)
           parLeverPositionZ = arrayOfInputs(8)
           parCounterWeightAngle = arrayOfInputs(9)
           parHVACShape = arrayOfInputs(10)
       
       Case 146                         'Back draft damper, type A
           parWidth = arrayOfInputs(2)
           parDepth = arrayOfInputs(3)
           parDamperLength = arrayOfInputs(4)
       
       Case 147                         'Back draft damper, type B
           parWidth = arrayOfInputs(2)
           parDepth = arrayOfInputs(3)
           parDamperLength = arrayOfInputs(4)
           parControlHeight = arrayOfInputs(12)
           parControlLength = arrayOfInputs(13)
           parControlWidth = arrayOfInputs(14)
           
    End Select
    
    If CmpDblEqual(parDepth, 0) Then
        parDepth = parWidth
    End If
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    Dim objInsDamperBody As Object
    
    Select Case lPartDataBasis
        
        Case Is <= 1, 145
            'Creating Insulation for Damper Body (Output 1)
            If (parHVACShape = 4) Then 'Round
                oStPoint.Set -parDamperLength / 2, 0, 0
                oEnPoint.Set parDamperLength / 2, 0, 0
        
                Set objInsDamperBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                    parWidth + (2 * parInsulationThickness), True)
            ElseIf (parHVACShape = Rectangular) Then
                oStPoint.Set -parDamperLength / 2, _
                            (parDepth / 2) + parInsulationThickness, _
                            (parWidth / 2) + parInsulationThickness
                
                oEnPoint.Set parDamperLength / 2, -(parDepth / 2) - parInsulationThickness, _
                            -(parWidth / 2) - parInsulationThickness
                            
                Set objInsDamperBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            End If
        
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsDamperBody
                
        Case 146, 147
            'Creating Insulation for Damper Body (Output 1)
            oStPoint.Set -parDamperLength / 2, _
                                parDepth / 2 + parInsulationThickness, _
                                parWidth / 2 + parInsulationThickness

            oEnPoint.Set parDamperLength / 2, _
                                -parDepth / 2 - parInsulationThickness, _
                                -parWidth / 2 - parInsulationThickness
            
            Set objInsDamperBody = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsDamperBody
            
    End Select
    Set objInsDamperBody = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
            
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError2 MODULE, METHOD
End Sub
