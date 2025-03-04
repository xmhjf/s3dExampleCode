VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ASCableWayHole2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------------------------------'
'    Copyright (C) 1998, 1999 Intergraph Corporation. All rights reserved.
'
'
'Abstract
'
'This class module contains the code that instanciates and recomputes a GSCAD symbol driven by a RAD2D .sym file
'it is currently used to define cableway hole.
'
'Notes
'1. There is one input as HoleTrace
'2. Get the working Plane and cable way from HoleTraceAE
'3. Get the Intersection point between workingPlane and cable way and normal vector of workingPlane.
'4. Get the collection of parameter descriptions from RAD2D
'5. Check the Angle Value of parameter, If angle is 90 degree, switch the width and length.
'6. Set the parameters as inputs for GSCAD symbol
'7. SetRepresentation
'8. Modify the transformation matrix to the symbol to move on the intersection with the normal vector.
' The graphic inputs of the GSCAD symbol are defined thanks to the attributes set on the RAD2D geometry while defining the
' .sym file in Symbol2d application.
'
'History
'
'    sypark@ship.samsung.co.kr    01/30/02                Creation.
'                                                         Originally created for StructDetailing
'                                                         brackets and features
'--------------------------------------------------------------------------------------------'

Option Explicit

Private Const MODULE = "ASCableWayHole2D"
Private m_oErrors As New IMSErrorLog.JServerErrors

'Interfaces needed
Implements IJDUserSymbolServices
Implements IJDDefinitionPlayerExOverride
Implements IJDDefinitionPlayerOverride


Private Sub IJDDefinitionPlayerExOverride_BindInputs(ByVal pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    
End Sub

Private Sub IJDDefinitionPlayerExOverride_BindOutputs(ByVal pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition, ByVal pRep As IMSSymbolEntities.IJDRepresentation)

End Sub


Private Sub IJDDefinitionPlayerExOverride_ComputeEllipsisOutputs(ByVal pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition, ByVal pRep As IMSSymbolEntities.IJDRepresentation)

End Sub

Private Sub IJDDefinitionPlayerExOverride_ComputeMatrix(ByVal pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)

End Sub

Private Sub IJDDefinitionPlayerExOverride_ConvertInputs(ByVal pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)

End Sub

Private Sub IJDDefinitionPlayerExOverride_ConvertOutputs(ByVal pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition, ByVal pRep As IMSSymbolEntities.IJDRepresentation)

End Sub

Private Sub IJDDefinitionPlayerExOverride_OpenGame(ByVal pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition, ByVal pPlayingSymbol As Object)

End Sub
Private Sub IJDDefinitionPlayerExOverride_ComputeConverter(ByVal pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    
    Const MT = "IJDDefinitionPlayerExOverride_ComputeConverter"
    On Error GoTo ErrorHandler
    
    Dim pEnumJDArgument As IEnumJDArgument
    Dim arg1 As IJDArgument
    Dim found As Long
    Dim ServerDoc As RAD2D.Document
    Dim oApplication As RAD2D.Application
    Dim DefPlayerEx As IMSSymbolEntities.IJDDefinitionPlayerEx
    Dim oInputs As IMSSymbolEntities.IJDInputs
    Dim oGeometryConv As MarineGenericSymbolLib.CSTGeomConverter
    Dim oSymbolOcc As IJDSymbol
    Dim oGraphicInput As IMSSymbolEntities.IJDInput
    Dim oPlane As IJPlane
    Dim StiffenerSolid As Object
    Dim PlateXStiffener As Object

    Set ServerDoc = pSymbolDefinition.IJDServerIdentification.object

    'Query Def PlayerEx interface
    Set DefPlayerEx = pSymbolDefinition
    
    'Query IJDInputs interface
    Set oInputs = pSymbolDefinition

    'Point on the symbol occurrence
    Set oSymbolOcc = DefPlayerEx.PlayingSymbol
    
    'Instanciate Struct geometry converter
    Set oGeometryConv = New MarineGenericSymbolLib.CSTGeomConverter

    Dim oHoleObject As Object
    Dim oHoleTrace As IJHoleTraceAE
    'Get the enum of arguments set by reference on the symbol if any
    Set pEnumJDArgument = oSymbolOcc.IJDReferencesArg.GetReferences()
    
    'If some inputs are set by reference, enumerate them
    If Not pEnumJDArgument Is Nothing Then
        'Look for U direction and V Direction curves
        'Will soon have to look for origin and curves.
        pEnumJDArgument.Reset
        Do
            pEnumJDArgument.Next 1, arg1, found
            If found = 0 Then Exit Do
            Set oGraphicInput = oInputs.Item(arg1.Index)
            If oGraphicInput.Properties <> igINPUT_IS_A_PARAMETER Then
                  Set oHoleObject = arg1.Entity
                  If TypeOf oHoleObject Is IJHoleTraceAE Then
                    Set oHoleTrace = oHoleObject
                    Set oHoleObject = Nothing
                  End If
            End If
        Loop
    End If

    ' The 3d context corresponds to the GSCAD ressource manager
    Dim oObject As IJDObject
    Set oObject = oSymbolOcc
    oGeometryConv.Context3D = oObject.ResourceManager
    Set oObject = Nothing

' there should be one Parameter Inputs not contained in the Rad2D .sym file
'   Horientation      - The angle of Hole of CW

    Dim dHorientation As Double
    GetCWHoleSymbolParameterInputs pSymbolDefinition, dHorientation
    
    If Not oHoleTrace Is Nothing Then
        Dim oConvertedPlane As IJPlane
        Set oConvertedPlane = ConvertWorkingPlane(oHoleTrace)
        
        'Compute a transformation matrix given the sketching plane only
        oGeometryConv.SketchingPlane = oConvertedPlane
        oGeometryConv.ViewTransform = ComputeMatrixForGeometryConverter(oConvertedPlane, dHorientation)

        Set oConvertedPlane = Nothing
    End If

    'The 2d context is the document of the server
    oGeometryConv.Context2D = ServerDoc

    'Set StructConverter in definition during the time of the game
    DefPlayerEx.Converter = oGeometryConv
           
'    'Optimizing:Turn off updating of the screen and Calculate when importing graphics
    If pSymbolDefinition.IJDDefinitionPlayerEx.GamePurpose = igCompute Then ServerDoc.CalculateOff

    If Not oHoleTrace Is Nothing Then
        oHoleTrace.State = HM_StateGood
        Set oHoleTrace = Nothing
    End If

'Release
    Set ServerDoc = Nothing
    Set oSymbolOcc = Nothing
    Set oPlane = Nothing
    Set pEnumJDArgument = Nothing
    Set ServerDoc = Nothing
    Set pEnumJDArgument = Nothing
    
    Exit Sub

ErrorHandler:
    'HandleError MODULE, MT
    m_oErrors.Add Err.Number, MODULE & " - " & MT, Err.Description
End Sub

Private Sub IJDDefinitionPlayerOverride_SetInputs(ByVal pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition, ByVal pSymbolOccurrence As IMSSymbolEntities.IJDSymbol, ByVal pEnumJDArgument As IMSSymbolEntities.IEnumJDArgument)

End Sub

Private Sub IJDDefinitionPlayerOverride_Calculate(ByVal pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition, ByVal pSymbolOccurrence As IMSSymbolEntities.IJDSymbol)
    Const MT = "IJDDefinitionPlayerOverride_Calculate"
    On Error GoTo ErrorHandler
    
    Dim ServerDoc As RAD2D.Document
    Dim DimsOnSheet As RAD2D.Dimensions
    Dim oDimension As RAD2D.Dimension
    
    'Get document
    Set ServerDoc = pSymbolDefinition.IJDServerIdentification.object
    
    '    OBM 25/11/99: Kludge to notify assoc that bsplines must be updated at next  RAD2D BatchFlushOn
    Set DimsOnSheet = ServerDoc.ActiveSheet.Dimensions
    For Each oDimension In DimsOnSheet
        oDimension.Value = oDimension.Value + 0.0001
        oDimension.Value = oDimension.Value - 0.0001
    Next
    'End Kludge
    
    'Call Batch on to update the document
    ' was ServerDoc.CalculateOn
    ServerDoc.CalculateOnGeomReplaced

    Set ServerDoc = Nothing
    Set DimsOnSheet = Nothing
    Exit Sub

ErrorHandler:
'    HandleError MODULE, MT
    m_oErrors.Add Err.Number, MODULE & " - " & MT, Err.Description
End Sub

Private Sub IJDDefinitionPlayerOverride_ReturnResult(ByVal pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition, ByVal pSymbolOccurrence As IMSSymbolEntities.IJDSymbol, ByVal repName As String, ByVal pOutput As IMSSymbolEntities.IJDOutputCollection)
End Sub

Private Sub IJDDefinitionPlayerExOverride_CloseGame(ByVal pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
' This sub is no more used
End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    Const MT = "IJDUserSymbolServices_GetDefinitionName"
    On Error GoTo ErrorHandler
     
     Dim insertTypeString As String
     Dim insertTypeVal As IMSInsertionType
     Dim serverName As String
    
     ' Extract the Server File Name from the Definition Name
     insertTypeString = "Link"
     insertTypeVal = igLinked
    
    '  Get the server path
'     Dim oQuery As IJSRDQuery
'     Set oQuery = New SRDQuery
     Dim strFilePath As String
'     strFilePath = oQuery.GetCatalogServerPath() + "\"
'     Set oQuery = Nothing
    
    strFilePath = "%OLE_SERVER%" + "\"
    
     GetInfoFromDefParameters definitionParameters, insertTypeString, insertTypeVal, serverName
     IJDUserSymbolServices_GetDefinitionName = SYMBOL_MACRO_PROGID + "|" + _
                            Str(insertTypeVal) + "|" + strFilePath + serverName

    Exit Function

ErrorHandler:
'    HandleError MODULE, MT
    m_oErrors.Add Err.Number, MODULE & " - " & MT, Err.Description
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
'This sub scans the document of the Symbol 's server in order to
'determine the inputs, the outputs and the representation of the GSCAD symbol.
    Const MT = "IJDUserSymbolServices_InitializeSymbolDefinition"
    On Error GoTo ErrorHandler
 
    If pSymbolDefinition Is Nothing Then
        MsgBox "IJDUserSymbolServices_InitializeSymbolDefinition : cannot initialize a symbol definition if not created"
        Exit Sub
    End If
 '                       Remove all previous GSCAD Symbol Definition information
 
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
   
  
  '                                    Get RAD2D document
    Dim ServerDoc As RAD2D.Document
    On Error Resume Next
    'Set ServerDoc = GetObject(pSymbolDefinition.IJDServerIdentification.serverName)
    'Fix for TR 48694
    Set ServerDoc = pSymbolDefinition.IJDServerIdentification.object
    On Error GoTo ErrorHandler
 
    If ServerDoc Is Nothing Then
        MsgBox "Fail in ASSymbol to open " + pSymbolDefinition.IJDServerIdentification.serverName + " document!"
        GoTo ErrorHandler
    End If

''''''''''''''''''''''''''  SCAN RAD2D DOCUMENT to describe the inputs and outputs of the GSCAD symbol''''''''''''''''''
  
    Dim oActiveSheet As RAD2D.Sheet
    Set oActiveSheet = ServerDoc.ActiveSheet
    Dim RAD2DGroups As RAD2D.Groups
    Set RAD2DGroups = oActiveSheet.Groups
   
    Dim oCommonMiscUtils As New CommonSymbolUtils.MiscUtils
    Dim oGroup As RAD2D.group
    Dim oCollOfGraphicInput As Collection
    Set oCollOfGraphicInput = New Collection
    For Each oGroup In RAD2DGroups
        oCommonMiscUtils.FillInputCollection oGroup, oCollOfGraphicInput  ' new method with now axis
    Next
 
''''''''''''''''''''''''''''''''''''' SET GRAPHIC INPUTS OF GSCAD SYMBOL: ''''''''''''''''''''''''''''''''''''''''''''
    ' Create the the Custom Method Utils
    Dim oSymbolDefUtils As New CommonSymbolUtils.SymbolDefinitionUtils
    Dim IJDinputsIf As IMSSymbolEntities.IJDInputs
    Set IJDinputsIf = pSymbolDefinition
    Dim oInputGraphic As IMSSymbolEntities.IJDInput
    Set oInputGraphic = New IMSSymbolEntities.DInput
    Dim InputIndex As Long
    InputIndex = 0

    'dummy graphic input as placeholder for HoleTrace input to SmartOccurrence
    oInputGraphic.Name = "Dummy"
    oInputGraphic.Description = "dummy"
  
    oSymbolDefUtils.SetCMBindOnInput oInputGraphic, "ResymbCableWayHole", CABLEHOLECM_LIB, pSymbolDefinition
    InputIndex = InputIndex + 1
    IJDinputsIf.SetInput oInputGraphic, InputIndex
    oInputGraphic.Reset

'''''''''''''''''''''  SET PARAMETERS INPUTS OF GSCAD SYMBOL '''''''''''''''''''''''''
    'graphic inputs have been set as inputs ,
    'now set the parameters as inputs for GSCAD symbol:
    Dim RAD2DSymbolDef As RAD2D.SymbolProperties
    Set RAD2DSymbolDef = ServerDoc.SymbolProperties
    'Test if not a .sha file, if not then it's a .sym file, retrieve its symbolproperties.
    If Not RAD2DSymbolDef Is Nothing Then
        oSymbolDefUtils.SetSymbolParameterInputs RAD2DSymbolDef, InputIndex, COMMONCM_LIB, pSymbolDefinition
    End If
    
    ' The CableWay Hole Symbol requires additional Parameter Inputs that are NOT
    ' defined in the Rectangle .sym file
    ' add the additional Parameter Inputs
    
    ' retrieve number of Inputs
    InputIndex = IJDinputsIf.Count
    
    Dim oInputParameter As IMSSymbolEntities.IJDInput
    Set oInputParameter = New IMSSymbolEntities.DInput
    
    Dim PC As IMSSymbolEntities.DParameterContent
    Set PC = New IMSSymbolEntities.DParameterContent
    PC.Type = igValue
    
    ' -----------------------------------------------------------------
    ' -----------------------------------------------------------------
    ' -----------------------------------------------------------------
    oInputParameter.Name = "Horientation"
    oInputParameter.Properties = igINPUT_IS_A_PARAMETER
    PC.UomValue = 0
    oInputParameter.DefaultParameterValue = PC
    oSymbolDefUtils.SetCMConvertOnInput oInputParameter, _
                                        "ResymbParameterAsDimOrVar", _
                                        CABLEHOLECM_LIB, _
                                        pSymbolDefinition
    InputIndex = InputIndex + 1
    IJDinputsIf.SetInput oInputParameter, InputIndex
    oInputParameter.Reset
  
''''''''''''''''''''''''' SET REPRESENTATION AND OUTPUTS OF GSCAD SYMBOL ''''''''''''
    ' Define the representation
    Dim RepsIf As IMSSymbolEntities.IJDRepresentations

    'Query the collection of representations by interfaces
    Set RepsIf = pSymbolDefinition

    'Define the type of representation and outputs that we need.
 
    DefineRepresentationForCableWayHole ServerDoc, RepsIf, pSymbolDefinition
               
    ServerDoc.Close False

    ' Release
    Set RepsIf = Nothing
    Set RAD2DSymbolDef = Nothing
    Exit Sub

ErrorHandler:
'    HandleError MODULE, MT
    m_oErrors.Add Err.Number, MODULE & " - " & MT, Err.Description
End Sub

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal definitionParameters As Variant, ByVal ActiveConnection As Object) As Object
    Const MT = "IJDUserSymbolServices_InstanciateDefinition"
    On Error GoTo ErrorHandler

    Dim pSymbolDefinition As IJDSymbolDefinition
    Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
    Set pSymbolDefinition = oSymbolFactory.CreateEntity(definition, ActiveConnection)
    Set oSymbolFactory = Nothing

    Dim ProgId As String
    Dim insertType As String
    Dim insertTypeVal As IMSInsertionType
    Dim serverName As String

    'Define the document of the server as linked!!!!!!!!!!!!!!!
    insertType = "Link"
    insertTypeVal = igLinked

    GetInfoFromDefParameters definitionParameters, insertType, insertTypeVal, serverName

    pSymbolDefinition.ProgId = SYMBOL_MACRO_PROGID
    pSymbolDefinition.CodeBase = CodeBase
    pSymbolDefinition.Name = IJDUserSymbolServices_GetDefinitionName(definitionParameters)
    pSymbolDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA

'''' support only -> go away when no occurences
    pSymbolDefinition.SupportOnlyOption = igSYMBOL_SUPPORT_ONLY
 
     ' get the server path
'    Dim oQuery As IJSRDQuery
'    Set oQuery = New SRDQuery
    Dim strFilePath As String
'    strFilePath = oQuery.GetCatalogServerPath() + "\" + serverName
'    Set oQuery = Nothing

    strFilePath = "%OLE_SERVER%" + "\" + serverName
    
    ' Set the server file name
    pSymbolDefinition.IJDServerIdentification.serverName = strFilePath

     ' Set the insertion type
    pSymbolDefinition.IJDServerIdentification.InsertionType = insertTypeVal

    'Define  custom methods and their libraries

    Dim UserMethod As IJDUserMethods
    Dim LibDesc As New DLibraryDescription
    Dim MethodDesc As New DMethodDescription

    Set UserMethod = pSymbolDefinition

    LibDesc.Name = "StdLib"
    LibDesc.Type = imsLIBRARY_IS_ACTIVEX
    LibDesc.Properties = imsLIBRARY_HAS_DEFINITIONPLAYEREX_INTERFACE + imsLIBRARY_AUTO_EXTRACT_METHOD_COOKIES + imsLIBRARY_HAS_DEFINITIONPLAYER_INTERFACE
    LibDesc.Source = SYMBOL_MACRO_PROGID
    'Set library in the symbol definition
    UserMethod.SetLibrary LibDesc

    'Declare the library of Custom Methods for Inputs, Outputs, and representation
     Dim LibDescCustomMethods As New DLibraryDescription

     LibDescCustomMethods.Name = "SketchingAndMacroCustomMethodsLib"
     LibDescCustomMethods.Type = imsLIBRARY_IS_ACTIVEX
     LibDescCustomMethods.Properties = imsLIBRARY_AUTO_EXTRACT_METHOD_COOKIES
     LibDescCustomMethods.Source = LIBRARYNAME_OF_CUSTOMMETHODS_SKETCH
    'Set library in the symbol definition
     UserMethod.SetLibrary LibDescCustomMethods

        'Declare the library of Custom Methods for Inputs, Outputs, and representation
     Dim LibDescCommonCustomMethods As New DLibraryDescription

     LibDescCommonCustomMethods.Name = COMMONCM_LIB
     LibDescCommonCustomMethods.Type = imsLIBRARY_IS_ACTIVEX
     LibDescCommonCustomMethods.Properties = imsLIBRARY_AUTO_EXTRACT_METHOD_COOKIES
     LibDescCommonCustomMethods.Source = LIBRARYNAME_OF_COMMONCUSTOMMETHODS
    'Set library in the symbol definition
     UserMethod.SetLibrary LibDescCommonCustomMethods


            'Declare the library of Custom Methods for Inputs, Outputs, and representation
     Dim LibDescCableHoleCustomMethods As New DLibraryDescription

     LibDescCableHoleCustomMethods.Name = CABLEHOLECM_LIB
     LibDescCableHoleCustomMethods.Type = imsLIBRARY_IS_ACTIVEX
     LibDescCableHoleCustomMethods.Properties = imsLIBRARY_AUTO_EXTRACT_METHOD_COOKIES
     LibDescCableHoleCustomMethods.Source = LIBRARYNAME_OF_CUSTOMMETHODS
    'Set library in the symbol definition
     UserMethod.SetLibrary LibDescCableHoleCustomMethods


    ' Override methods of interface IJDDefinitionPlayerEx

     ' Override the ComputeConverter method
      MethodDesc.Name = "ComputeConverter"
      MethodDesc.Properties = imsMETHOD_OVERRIDE
      MethodDesc.Cookie = imsCOOKIE_IS_NULL
      MethodDesc.Library = "StdLib"
      UserMethod.SetMethod MethodDesc

      ' Override the Calculate method
      MethodDesc.Name = "Calculate"
      MethodDesc.Properties = imsMETHOD_OVERRIDE
      MethodDesc.Cookie = imsCOOKIE_IS_NULL
      MethodDesc.Library = "StdLib"
      UserMethod.SetMethod MethodDesc


      'Define the Evaluation Custom Methods for the representations:

      MethodDesc.Name = "ResymbCableWayHole"
      MethodDesc.Properties = 0
      MethodDesc.Cookie = imsCOOKIE_IS_NULL
      MethodDesc.Library = CABLEHOLECM_LIB
      UserMethod.SetMethod MethodDesc

      'Define the Evaluation Custom Methods for the representations:
      MethodDesc.Name = "StdEvaluationForMacro"
      MethodDesc.Properties = 0
      MethodDesc.Cookie = imsCOOKIE_IS_NULL
      MethodDesc.Library = "SketchingAndMacroCustomMethodsLib"
      UserMethod.SetMethod MethodDesc


     Set MethodDesc = Nothing
     Set LibDesc = Nothing
     Set UserMethod = Nothing

    IJDUserSymbolServices_InitializeSymbolDefinition pSymbolDefinition

    'returned symbol defintion
    Set IJDUserSymbolServices_InstanciateDefinition = pSymbolDefinition

    Set pSymbolDefinition = Nothing
    Exit Function

ErrorHandler:
'    HandleError MODULE, MT
    m_oErrors.Add Err.Number, MODULE & " - " & MT, Err.Description
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub
 
