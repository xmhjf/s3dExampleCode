VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ButtWeldSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const m_SelectorProgid As String = CUSTOMERID + "PhysConnRules.ButtWeldSel"
'Const m_SelectorProgid As     String = m_sProjectID & ".ButtWeldSel"
Const m_SelectorName As String = CUSTOMERID + "PhysConnRules.ButtWeldSel"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "PhysConnRules\ButtWeldSel.cls"

Implements IJDUserSymbolServices

Public Sub SelectorInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
  
    'Add port inputs
    pIH.SetInput INPUT_CONN_OBJECT1
    pIH.SetInput INPUT_CONN_OBJECT2
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "SelectorInputs").Number
End Sub
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)
  On Error GoTo ErrorHandler
  
    pQH.SetQuestion "WeldingType_FromAssyConn", "Manual", "WeldingTypeCol"
    pQH.SetQuestion "Shipyard_FromAssyConn", "DefaultShipyard", "ShipyardNameCol"
    pQH.SetQuestion "WeldingPosition", "Vertical", "WeldingPositionCol", "SetWeldPosition", CUSTOMERID + "PhysConnRules.PhysConnSelCM"
    pQH.SetQuestion "Category", "OneSided", "ButtWeldCategory", "SetCategory", CUSTOMERID + "PhysConnRules.PhysConnSelCM"
    pQH.SetQuestion "ChamferType", "None"
    pQH.SetQuestion "FlipBevel", "No", "BooleanCol", "SetFlip", CUSTOMERID + "PhysConnRules.PhysConnSelCM"
    pQH.SetQuestion "FirstWeldingSide", "Molded", "FirstWeldingSideCol", "SetFirstWeldingSide", CUSTOMERID + "PhysConnRules.PhysConnSelCM"
    pQH.SetQuestion "Workcenter", gsMachine1, "WorkcenterCol", "SetWorkcenter", CUSTOMERID + "PhysConnRules.PhysConnSelCM"

  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "SelectorQuestions").Number
End Sub

Public Sub SelectorLogic(pSLH As IJDSelectorLogic)

    Dim sError As String
    
    On Error GoTo ErrorHandler
      
    sError = ""
      
    ' ----------------------------------------------
    ' Use filter in place of selection logic, if set
    ' ----------------------------------------------
    Dim oParmCM As New PhysConnParmCM
    
    If oParmCM.UsePreSelection(pSLH) Then
        Exit Sub
    End If
    
    ' Get Class Arguments
    Dim oPhysConn As New StructDetailObjects.PhysicalConn
    Dim dThickness As Double
    Set oPhysConn.object = pSLH.SmartOccurrence
    
     ' ------------------------------------------------------------
    ' First make sure that the geometry is not removed by features
    ' ------------------------------------------------------------
    ' If an AC created the PC, it does not know about slots and corner features that
    ' may have removed all the geometry.  This is possible when the seams/boundaries
    ' two objects are offset slightly.  Example: A corner feature on part 1 can remove
    ' the geoemtry that would support a PC for AC2.
    '
    '                             | |
    '                             | |
    '                             | |
    '  +--------------------------+ |
    '  |                          | |
    '  |                          | |
    '  |          part 1          | |
    '  |                          | |
    '  |                          | |
    '  +-------AC1--------+--AC2--+-+--------------
    '  +     part 2       |           part3
    '  +-------------------------------------------
    Dim oSysChild As IJDesignChild
    Set oSysChild = oPhysConn.object
    
    ' We limit the test to AC parents to avoid ignoring bad end cut geometry, which often
    ' results in the same error.
    If TypeOf oSysChild.GetParent Is IJAssemblyConnection Then
        
        Dim oPartGeom As IJDModelBody
        Set oPartGeom = oPhysConn.ConnectedObject1
        
        Dim oConnection As IJStructConnection
        Dim oConnGeom As IJDModelBody
'        Set oConnGeom = oPhysConn.object
        
        Dim oParent As Object
        Set oParent = oSysChild.GetParent
        
        If TypeOf oParent Is IJModelBody Then
        
          Set oConnGeom = oParent
          
          Dim oPointOnPart As IJDPosition
          Dim oPointOnAC As IJDPosition
          Dim minDist As Double
          
          oConnGeom.GetMinimumDistance oPartGeom, oPointOnPart, oPointOnAC, minDist
        
          ' 4mm is the distance tolerance for PCs (really should be exposed as a global constant or property)
          ' We are assuming here that the user does not allow the feature edge to come within 4mm of a seam
          ' on the bounding object.  This is typical practice. If the distance is greater than 4mm, we pick
          ' no item.  No PC will be placed and nothing will go on the ToDO list as a result
          If minDist > 0.004 Then
              Exit Sub
          End If
        End If
    End If
      
    ' get thicknesses
    Dim dObject1Thickness As Double
    Dim dObject2Thickness As Double
  
    dObject1Thickness = 0
    dObject2Thickness = 0
    GetPhysicalConnPartsThickness _
                       pSLH.SmartOccurrence, _
                       dObject1Thickness, _
                       dObject2Thickness
    If dObject1Thickness < 0.000001 Then
       dObject1Thickness = dObject1Thickness
    End If
    If dObject2Thickness < 0.000001 Then
       dObject2Thickness = dObject2Thickness
    End If
    
    If TypeOf oPhysConn.ConnectedObject1 Is IJPlate And _
       TypeOf oPhysConn.ConnectedObject2 Is IJPlate Then
       Dim oPlate1 As New StructDetailObjects.PlatePart
       Dim oPlate2 As New StructDetailObjects.PlatePart
       Set oPlate1.object = oPhysConn.ConnectedObject1
       Set oPlate2.object = oPhysConn.ConnectedObject2
          If oPlate1.OffsetToBaseFace(oPlate2) < oPlate2.OffsetToBaseFace And _
              oPlate1.OffsetToOffsetFace(oPlate2) < oPlate2.OffsetToOffsetFace Then
              dThickness = oPlate1.OffsetToOffsetFace(oPlate2) - oPlate2.OffsetToBaseFace
          ElseIf oPlate1.OffsetToBaseFace(oPlate2) > oPlate2.OffsetToBaseFace And _
              oPlate1.OffsetToOffsetFace(oPlate2) > oPlate2.OffsetToOffsetFace Then
              dThickness = oPlate2.OffsetToOffsetFace - oPlate1.OffsetToBaseFace(oPlate2)
          Else
              If dObject1Thickness > dObject2Thickness Then
                  dThickness = dObject2Thickness
              Else
                  dThickness = dObject1Thickness
              End If
          End If
    Else
      If dObject1Thickness > dObject2Thickness Then
          dThickness = dObject2Thickness
      Else
          dThickness = dObject1Thickness
      End If
    End If

    ' Get the Question answers
    Dim sSided As String
    sSided = pSLH.Answer("Category")
    Dim sWeldingType As String
    sWeldingType = pSLH.Answer("WeldingType_FromAssyConn")
    Dim sWeldingPosition As String
    sWeldingPosition = pSLH.Answer("WeldingPosition")
    Dim sShipyard As String
    sShipyard = pSLH.Answer("Shipyard_FromAssyConn")
    
    Dim oStiffener1 As New StructDetailObjects.ProfilePart
    Dim oStiffener2 As New StructDetailObjects.ProfilePart
    ' ********************* Selection code - TODO  ****************
    Select Case sShipyard
    
    Case "DefaultShipyard"
        If sWeldingType = "Manual" Then
            If dThickness < 0.006 Then
                pSLH.Add "ButtWeldI"
            ElseIf dThickness >= 0.006 Then 'thickness is greater than 6 mm
                pSLH.Add "ButtWeldV"
                pSLH.Add "ButtWeldIV"
                
                'allow the X if the weld is not horizontal
                If Not sWeldingPosition = "Horizontal" Then
                    pSLH.Add "ButtWeldX"
                Else
                    'don't add anything other than the V
                End If
                    
                If dThickness >= 0.02 Then 'if greater than 20 mm, also add the IX
                    pSLH.Add "ButtWeldIX"
                Else
                    'don't add anything new
                End If
                
                If TypeOf oPhysConn.ConnectedObject1 Is IJStiffenerPart And _
                   TypeOf oPhysConn.ConnectedObject2 Is IJStiffenerPart Then
                      
                      Set oStiffener1.object = oPhysConn.ConnectedObject1
                      Set oStiffener2.object = oPhysConn.ConnectedObject2
                      
                            If (oStiffener1.SectionType = "HalfR" And oStiffener2.SectionType = "HalfR") Or _
                               (oStiffener1.SectionType = "R" And oStiffener2.SectionType = "R") Then
                                pSLH.Add "ButtWeldY"
                            End If
                        End If
            End If
                    
        ElseIf sWeldingType = "CO2" Then
            If sWeldingPosition = "Horizontal" Then
                pSLH.Add "ButtWeldIV"
            Else
                pSLH.Add "ButtWeldV"
            End If
            
        ElseIf sWeldingType = "UnionMeltWeld" Then
            If dThickness <= 0.02 Then
                pSLH.Add "ButtWeldI"
            Else
                pSLH.Add "ButtWeldK"
            End If
        
        ElseIf sWeldingType = "NSOneSideMAG" Then
            If dThickness >= 0.012 And dThickness <= 0.022 Then
                pSLH.Add "ButtWeldV"
            Else
                'this weld type is not used
            End If
        
        ElseIf sWeldingType = "FCB" Then
            pSLH.Add "ButtWeldY"
        
        ElseIf sWeldingType = "FAB" Then
            pSLH.Add "ButtWeldV"
            
        ElseIf sWeldingType = "ElectroGasWeld" Then
            pSLH.Add "ButtWeldV"
            
        ElseIf sWeldingType = "ConsumptionNozzleWeld" Then
            pSLH.Add "ButtWeldI"
            
        ElseIf sWeldingType = "PicoMaxWeld" Then
            pSLH.Add "ButtWeldV"
            
        ElseIf sWeldingType = "SEGARCWeld" Then
            pSLH.Add "ButtWeldV"
        
        ElseIf sWeldingType = "Chill" Then
            pSLH.Add "ButtWeldV"
        
        End If
    
    End Select
    'Update edge features which are dependent on this PC (CR-164293)
    UpdateDependentEdgeFeature oPhysConn.object
    
    Set oStiffener1 = Nothing
    Set oStiffener2 = Nothing
    Set oPhysConn = Nothing
  
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "SelectorLogic", sError).Number
End Sub


' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = m_SelectorName
End Function
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)
  
  ' Remove all existing defined Input and Output (Representations)
  ' before defining the current Inputs and Outputs
  ' *
  pSelector.IJDInputs.RemoveAllInput
  pSelector.IJDRepresentations.RemoveAllRepresentation
  
  Dim pSelectorFact As New DefinitionFactory
  pSelectorFact.InitAbstractSelector pSelector
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.definition = pSelector
  pIH.InitAs m_FamilyProgid
  SelectorInputs pIH
  Dim pQH As IJDQuestionsHelper
  Set pQH = New QuestionHelper
  pQH.Selector = pSelector
  SelectorQuestions pQH
End Sub
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pSelectorFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pSelectorFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function
Public Sub CMSelector(pRep As IJDRepresentation)
  Dim pSLH As IJDSelectorLogic
  Set pSLH = New SelectorLogic
  pSLH.Representation = pRep
  SelectorLogic pSLH
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************

