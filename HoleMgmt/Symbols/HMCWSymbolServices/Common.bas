Attribute VB_Name = "Common"
'--------------------------------------------------------------------------------------------'
'    Copyright (C)  1999 Intergraph Corporation. All rights reserved.
'
'
'Abstract
'    Helper math functions used for symbol recompute
'
'
'Notes
'

'
'History
'    O. Bernard-Millard      10/01/99                Creation.
'   P. Williams             Oct 2000            Add method to Get the plane from a Slot Symbol
'--------------------------------------------------------------------------------------------'
Option Explicit

Private Const MODULE = "Common"

Public Const SYMBOL_SKETCH_PROGID = "HMCWSymbolServices.ASSketcher2d"
Public Const SYMBOL_MACRO_PROGID = "HMCWSymbolServices.ASCableWayHole2D"
Public Const LIBRARYNAME_OF_CUSTOMMETHODS_SKETCH = "StructSymbolCustomMethods.CustomMethods"
Public Const SYMBOL_SKETCH_REPRESENTATION_NAME = "MFSymbolContourRep"
Public Const SYMBOL_MACRO_REPRESENTATION_NAME = "Shape3D"
Public Const cAttName = "EdgeName"
Public Const cAttSetName = "EdgeNames"
Public Const cPtAttName = "Load Point"
Public Const cPtAttSetName = "PointNames"
Public Const STRUCT_RELATION_SUBINPUT_IID = "{6034AD40-FA0B-11d1-B2FD-080036024603}"
Public Const STRUCT_RELATION_SUBINPUT_COLNAME = "StructSymbolSubInput_DEST"



Public Sub GetInfoFromDefParameters(ByVal defParameters As String, insertTypeString As String, insertTypeVal As IMSInsertionType, serverName As String)
Const MT = "GetInfoFromDefParameters"
On Error GoTo ErrorHandler

    Dim BeginPosFileName
    Dim tmpString As String
    
    serverName = ""
    
    BeginPosFileName = 0
    BeginPosFileName = InStr(1, defParameters, "|", vbTextCompare)
    
    If BeginPosFileName = 0 Then
        serverName = defParameters
    Else
        insertTypeString = Left(defParameters, BeginPosFileName - 1)
        serverName = Mid(defParameters, BeginPosFileName + 1)
    End If
    
    If StrComp(insertTypeString, "embed", vbTextCompare) = 0 Then
        insertTypeVal = igEmbedded
    ElseIf StrComp(insertTypeString, "link", vbTextCompare) = 0 Then
        insertTypeVal = igLinked
    ElseIf StrComp(insertTypeString, "shared", vbTextCompare) = 0 Then
        insertTypeVal = igSharedEmbedded
    Else
        MsgBox "InsertionType can not be found"
        insertTypeVal = igNone
    End If
Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub



Public Sub HandleError(sModule As String, sMethod As String)
    MsgBox sModule & "." & sMethod & ":" & _
            " Source= " & Err.Source & _
            " Number=" & Trim$(Str$(Err.Number)) & _
            " Description=" & Err.Description
End Sub

Public Sub FillInputDescription(pGraphicObj As Object, pInput As IMSSymbolEntities.IJDInput)
Const MT = "FillInputDescription"
On Error GoTo ErrorHandler

    Dim TheGraphicAttributes As RAD2D.AttributeSets
    Dim oAttributeSet As RAD2D.AttributeSet
    Dim oAttribute As RAD2D.Attribute
    
    Set TheGraphicAttributes = pGraphicObj.AttributeSets
    Set oAttributeSet = TheGraphicAttributes.Item("Input")
    For Each oAttribute In oAttributeSet
        With oAttribute
            If .Name = "Name" Then
                pInput.Name = .Value
            End If
        End With
    Next
    pInput.Key = pGraphicObj.Key
    
    Exit Sub
ErrorHandler:
  HandleError MODULE, MT
End Sub
''''''''''''''
''''''''''''''   PW: Attempt to make this one more generic by adding an input by value
'''''
Public Sub FillCollectionsOfInputs(RadObject As Object, CollOfAxis As Collection, CollOfInputs As Collection)
    Dim TheGraphicAttributes As RAD2D.AttributeSets
    Dim oAttributeSet As RAD2D.AttributeSet
    Dim oAttribute As RAD2D.Attribute
    Dim bAdded As Boolean
    Dim oGroup As RAD2D.group
    Set TheGraphicAttributes = RadObject.AttributeSets
    If Not TheGraphicAttributes Is Nothing Then
        For Each oAttributeSet In TheGraphicAttributes
            bAdded = False
            If oAttributeSet.SetName = "Input" Then
            'Find all of the attributes named "Axis"
            'These must be added first to support bracket orientation / translation
            'if we do not loop through the attribute set twice,
            'the collections (axis and other inputs) will be
            'populated based on the order of the
            'attribute names.  This means that, for brackets, this
            'method could return an axis collection with only one
            'item.

                For Each oAttribute In oAttributeSet
                     
                    If oAttribute.Name = "Axis" Then
                        'Be careful: Axis can be set as attribute but without value!
                        If oAttribute.Value = "U" Then
                            CollOfAxis.Add RadObject, "U"
                            bAdded = True
                        ElseIf oAttribute.Value = "V" Then
                            CollOfAxis.Add RadObject, "V"
                            bAdded = True
                        End If
                    End If
                Next
                ' the input was not named "Axis"
                For Each oAttribute In oAttributeSet
                    If oAttribute.Name = "Name" Then
                        If Not bAdded Then
                            CollOfInputs.Add RadObject, oAttribute.Value
                        End If
                    End If
                Next
            End If
            
        Next
    End If
End Sub

Private Sub Fill4x4Matrix(p4x4Matrix As AutoMath.DT4x4, M11 As Double, M21 As Double, M31 As Double, _
                          M41 As Double, M12 As Double, M22 As Double, M32 As Double, _
                          M42 As Double, M13 As Double, M23 As Double, M33 As Double, _
                          M34 As Double, TransX As Double, TransY As Double, TransZ As Double, _
                          ScalingCoeff As Double)
Const MT = "Fill4x4Matrix"
On Error GoTo ErrorHandler

    Dim ValueTransf(16) As Double
    
    p4x4Matrix.Get ValueTransf(0)
    
    ValueTransf(0) = M11
    ValueTransf(1) = M21
    ValueTransf(2) = M31
    ValueTransf(3) = M41
    ValueTransf(4) = M12
    ValueTransf(5) = M22
    ValueTransf(6) = M32
    ValueTransf(7) = M42
    ValueTransf(8) = M13
    ValueTransf(9) = M23
    ValueTransf(10) = M33
    ValueTransf(11) = M34
    ValueTransf(12) = TransX
    ValueTransf(13) = TransY
    ValueTransf(14) = TransZ
    ValueTransf(15) = ScalingCoeff

    p4x4Matrix.Set ValueTransf(0)

Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub


Public Function GetSymbol2DAppDocAttribute(ServerDoc As RAD2D.Document) As String
 Const MT = "GetSymbol2DAppDocAttribute"
On Error GoTo ErrorHandler
 Dim oAttributeSets As AttributeSets
 Dim oAttributeSet As AttributeSet
 Dim oAttribute As RAD2D.Attribute
 Set oAttributeSets = ServerDoc.AttributeSets
 For Each oAttributeSet In oAttributeSets
    If oAttributeSet.SetName = "Symbol2D AddIns" Then
        For Each oAttribute In oAttributeSet
            If oAttribute.Name = "Active Name""" Then
                GetSymbol2DAppDocAttribute = oAttribute.Value
            End If
        Next
    End If
 Next
 Exit Function
ErrorHandler:
    HandleError MODULE, MT
End Function
Public Sub SetCMOnRepresentation(pRep As IMSSymbolEntities.IJDRepresentation, LibraryName As String, MethodName As String, pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
 Const MT = "SetCMOnRepresentation"
On Error GoTo ErrorHandler

 Dim mthCookie As Long
 Dim UserMethod As IJDUserMethods
 Dim oLib As IMSSymbolEntities.DLibraryDescription
 
 Set UserMethod = pSymbolDefinition
 mthCookie = UserMethod.GetMethodCookie(MethodName, LibraryName)
 Set oLib = UserMethod.GetLibrary(LibraryName)
 pRep.IJDRepresentationStdCustomMethod.SetCMEvaluate oLib.Cookie, mthCookie
 
 Set UserMethod = Nothing
Exit Sub
ErrorHandler:
    HandleError MODULE, MT

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' order object so that for every object,
' the next one has either its end or its start point common to it
Public Sub GetOrderedSymbolDrawing(ByRef group As RAD2D.group, ByRef oColl As Collection)
    Dim oObjectColl As New Collection
    Dim tempColl As New Collection
    Dim firstObject As Object
    
    Dim currentObject As Object
    Dim bFound As Boolean
    Dim bFirstTimeNotFound As Boolean
    Dim ii As Integer
    Dim x As Double
    Dim y As Double
    
    Dim oTemp As Object
    
    On Error GoTo ErrorHandler

    For Each currentObject In group
        If currentObject.Type = igLine2d Or currentObject.Type = igArc2d Or currentObject.Type = igEllipticalArc2d _
            Or currentObject.Type = igLineString2d Then
            ' JC 11/17/99: removed Circle and Ellipse because they don't have end points
            ' Or currentObject.Type = igCircle2d Or currentObject.Type = igEllipse2d Then
            oObjectColl.Add currentObject
        End If
    Next currentObject
    
    Set currentObject = oObjectColl.Item(1)
    Set firstObject = currentObject
    
    currentObject.GetEndPoint x, y
    oObjectColl.Remove 1
    tempColl.Add currentObject
    
    bFound = True
    bFirstTimeNotFound = True
    
    While oObjectColl.Count > 0 And bFound
        bFound = False
        
        For ii = 1 To oObjectColl.Count
                    
            Set oTemp = oObjectColl.Item(ii)
            
            If PointFitObjectEnd(x, y, oTemp, x, y) Then
                oObjectColl.Remove ii
                Set currentObject = oTemp
                bFound = True
                Exit For
            End If
        Next ii
    
        If bFound Then
            ' add the object into the list, either before the first element or after it
            If bFirstTimeNotFound Then
                tempColl.Add currentObject
            Else
                tempColl.Add currentObject, , 1
            End If
        ElseIf bFirstTimeNotFound Then '<=> Not found and bFirstTimeNotFound
            ' if no end is found, go reverse (for open contour)
            firstObject.GetStartPoint x, y
            bFound = True
            bFirstTimeNotFound = False
        End If
    Wend
    
    Dim obj As Object
    ReDim nameArr(tempColl.Count)
    
    For Each currentObject In tempColl
        Set obj = currentObject
        oColl.Add obj
        nameArr(oColl.Count) = obj.Key
    Next currentObject
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
    Err.Clear
    Resume Next
    Err.Clear
End Sub
Private Function PointFitObjectEnd(x As Double, y As Double, obj As Object, newX As Double, newY As Double) As Boolean
    Dim sx As Double
    Dim sy As Double
    Dim ex As Double
    Dim ey As Double

    obj.GetStartPoint sx, sy
    obj.GetEndPoint ex, ey
    
    ' if the coordinates are equal to the object start point
    If (Abs(x - sx) < 0.000001 And Abs(y - sy) < 0.000001) Then
        PointFitObjectEnd = True
        newX = ex
        newY = ey
    ' if the coordinates are equal to the object end point
    ElseIf (Abs(x - ex) < 0.000001 And Abs(y - ey) < 0.000001) Then
        PointFitObjectEnd = True
        newX = sx
        newY = sy
    Else
        PointFitObjectEnd = False
        newX = x
        newY = y
    End If
    
End Function
Public Function GetRADAttributeValueOrKey(strAttrSetName As String, _
                                        oRADObject As Object, _
                                        strAttName As String) As String

    Dim oAttribute As RAD2D.Attribute
    Dim bAttrFound As Boolean
    Dim oAttributeSet As RAD2D.AttributeSet
    
    On Error GoTo ErrorHandler
    
    GetRADAttributeValueOrKey = ""
    
    If oRADObject Is Nothing Then
        Exit Function
    End If
    
    GetAttributeSetOnRADObj oRADObject, strAttrSetName, oAttributeSet
    
    If Not oAttributeSet Is Nothing Then
        For Each oAttribute In oAttributeSet
        
            If oAttribute.Name = strAttName Then
                GetRADAttributeValueOrKey = oAttribute.Value
                bAttrFound = True
                Exit For
            End If
        Next oAttribute
    End If
    
    If Not bAttrFound Then
        GetRADAttributeValueOrKey = oRADObject.Key
    End If
    
 Exit Function
 
ErrorHandler:
   MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
    
End Function
Public Sub GetAttributeSetOnRADObj(oGraphicObj As Object, _
                                    strAttrSetName As String, _
                                    oNamedAttributeSet As Object)
                                    
    Dim oAttributeSet As RAD2D.AttributeSet

 On Error GoTo ErrorHandler
  
    Set oNamedAttributeSet = Nothing
    
    If Not oGraphicObj.AttributeSets Is Nothing Then
'
        For Each oAttributeSet In oGraphicObj.AttributeSets
'
            If oAttributeSet.SetName = strAttrSetName Then
                Set oNamedAttributeSet = oAttributeSet
                Exit For
            End If
'
        Next oAttributeSet
        
    End If
    
    Exit Sub
    
ErrorHandler:
   MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description

End Sub

Public Sub SymbolDefinition_SetCMBindOnInput(pInput As IMSSymbolEntities.IJDInput, BindingMethodName As String, LibraryName As String, pSymbolDef As IMSSymbolEntities.IJDSymbolDefinition)
 Const MT = "SymbolDefinition_SetCMBindOnInput"
On Error GoTo ErrorHandler
    ' associate a custom method to the Input for Binding  process
    Dim oLibrary As IJDLibraryDescription
    Dim oUserMeths As IJDUserMethods
    Dim LibCookie As Long
    Dim MethodCookieBindConvert As Long
    
    Set oUserMeths = pSymbolDef
    
    Set oLibrary = oUserMeths.GetLibrary(LibraryName)
    'Get the cookie that defines the librairy
    LibCookie = oLibrary.Cookie
    'Get the cookie that defines the convert method
    MethodCookieBindConvert = oUserMeths.GetMethodCookie(BindingMethodName, LibCookie)
    
    pInput.IJDInputStdCustomMethod.SetCMBind LibCookie, MethodCookieBindConvert
Exit Sub
    
ErrorHandler:
   MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description

End Sub
Public Sub SymbolDefinition_SetCMConvertOnInput(pInput As IMSSymbolEntities.IJDInput, BindingMethodName As String, LibraryName As String, pSymbolDef As IMSSymbolEntities.IJDSymbolDefinition)
 Const MT = "SymbolDefinition_SetCMConvertOnInput"
On Error GoTo ErrorHandler
    ' associate a custom method to the Input for Binding  process
    Dim oLibrary As IJDLibraryDescription
    Dim oUserMeths As IJDUserMethods
    Dim LibCookie As Long
    Dim MethodCookieBindConvert As Long
    Set oUserMeths = pSymbolDef
    
    Set oLibrary = oUserMeths.GetLibrary(LibraryName)
    'Get the cookie that defines the librairy
    LibCookie = oLibrary.Cookie
    'Get the cookie that defines the convert method
    MethodCookieBindConvert = oUserMeths.GetMethodCookie(BindingMethodName, LibCookie)
    
    pInput.IJDInputStdCustomMethod.SetCMConvert LibCookie, MethodCookieBindConvert
Exit Sub
    
ErrorHandler:
   MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description

End Sub
Public Sub SetGraphicOutputsInRepresentation(CollectionOfGraphics As Collection, pSymbolDef As IMSSymbolEntities.IJDSymbolDefinition)
Const MT = "SetGraphicOutputsInRepresentation"
'Recreate representation according geometry present in Imaginer file
    
    Dim Rep As IMSSymbolEntities.IJDRepresentation
    Dim oOutputs As IMSSymbolEntities.IJDOutputs
    
    Dim OGraphicOutput As IJDOutput
    Dim oGraphicObject As Object
    Dim reps As IJDRepresentations
   
    Dim OutputDescription As String
       
    Set reps = pSymbolDef.IJDRepresentations
    'Get a pointer on representation to update it.
    Set Rep = reps.GetRepresentationByName(SYMBOL_SKETCH_REPRESENTATION_NAME)
    OutputDescription = Rep.Description
    
    'Query IJDOutputs interface
    Set oOutputs = Rep
    
    'Destroy previous outputs
    oOutputs.RemoveAllOutput
           
    'Declare new outputs
    Set OGraphicOutput = New IMSSymbolEntities.DOutput
   
    For Each oGraphicObject In CollectionOfGraphics
        Dim oCollOfOrderedSegments As New Collection
        Dim oSubOutput As IJDOutput
        Dim oSegment As Object
        Set oSubOutput = New DOutput
        If oGraphicObject.Type = igGroup Then
            Dim aGroup As RAD2D.group
            Set aGroup = oGraphicObject
            If aGroup.Count > 1 Then
                OrderObjectsInGroup aGroup, oCollOfOrderedSegments
                For Each oSegment In oCollOfOrderedSegments
                     oSubOutput.Key = oSegment.Key
                     oSubOutput.Name = oSegment.Name
                     OGraphicOutput.IJDOutputs.Add oSubOutput
                     oSubOutput.Reset
                Next
            Else
                oSubOutput.Key = oGraphicObject.Key
                oSubOutput.Name = oGraphicObject.Name
                OGraphicOutput.IJDOutputs.Add oSubOutput
                oSubOutput.Reset
            End If
        ElseIf oGraphicObject.Type = igSymbol2d Then
            Dim SymbolDrawings As RAD2D.DrawingObjects
            Dim aSymbol As RAD2D.Symbol2d
            Set aSymbol = oGraphicObject
            Set SymbolDrawings = aSymbol.DrawingObjects
            If SymbolDrawings.Count = 1 Then
                'Be careful that single object inside symbol can be a group!
                Dim RadObject As Object
                Set RadObject = SymbolDrawings.Item(1)
                If RadObject.Type = igGroup Then
                    Dim TheGroup As RAD2D.group
                    Set TheGroup = RadObject
                    If TheGroup.Count > 1 Then
                        OrderObjectsInGroup TheGroup, oCollOfOrderedSegments
                        For Each oSegment In oCollOfOrderedSegments
                             oSubOutput.Key = oSegment.Key
                             oSubOutput.Name = oSegment.Name
                             OGraphicOutput.IJDOutputs.Add oSubOutput
                             oSubOutput.Reset
                        Next
                    Else
                        oSubOutput.Key = TheGroup.Key
                        oSubOutput.Name = TheGroup.Name
                        OGraphicOutput.IJDOutputs.Add oSubOutput
                        oSubOutput.Reset
                    End If
                Else
                   
                End If
            Else
                OrderElemsInRAD2DSymbol SymbolDrawings, oCollOfOrderedSegments
                For Each oSegment In oCollOfOrderedSegments
                     oSubOutput.Key = oSegment.Key
                     oSubOutput.Name = oSegment.Name
                     OGraphicOutput.IJDOutputs.Add oSubOutput
                     oSubOutput.Reset
                Next
            End If
        End If
        OGraphicOutput.Name = "Output" & "_" & oGraphicObject.Key
        OGraphicOutput.Description = OutputDescription
        OGraphicOutput.Properties = 0
        OGraphicOutput.Key = oGraphicObject.Key
        oOutputs.SetOutput OGraphicOutput
        OGraphicOutput.Reset
    Next
    
    'Release
    
    Set Rep = Nothing
   
Exit Sub
ErrorHandler: HandleError MODULE, MT
End Sub
Public Sub SymbolDefinition_SetInputsForForConstrainedRADGeometry(pSymbolDef As IMSSymbolEntities.IJDSymbolDefinition, pPlayingSymbol As IJDSymbol)
Const MT = "SymbolDefinition_SetInputsForForConstrainedRADGeometry"
On Error GoTo ErrorHandler
    
    Dim oInput As IJDInput
    Dim RAD2DKey As String
    Dim LoopOnKey As Long
    Dim RadObject As Object
    Dim AssistantCollection As Object
    Dim RAD2DkeyColl As Object
    Dim UserListOfDispatch As IJElements
    Dim ServerDoc As RAD2D.Document
    Dim aSheet As RAD2D.Sheet
    
    Dim Index As Long
    Dim SubReferenceCollOthers As IMSSymbolEntities.IJDReferencesCollection
    Dim SubReferenceCollRefPlanes As IMSSymbolEntities.IJDReferencesCollection
    Dim oSubInput As IMSSymbolEntities.IJDInput
    Dim MainRefColl As IMSSymbolEntities.IJDReferencesCollection
    
    Dim IsConstrained As Boolean

    Set ServerDoc = pSymbolDef.IJDServerIdentification.object
    Set aSheet = ServerDoc.ActiveSheet
    
    Set MainRefColl = pPlayingSymbol.IJDReferencesArg.GetReferencesCollection
    
    Dim pEnumJDArgument As IEnumJDArgument
    Dim arg1 As IJDArgument
    Dim found As Long
    Dim SubCount As Long
    Dim LoopCount As Long
    Dim bOthers As Boolean
    Dim bRefPlanes As Boolean
    
    
    Set pEnumJDArgument = pPlayingSymbol.IJDReferencesArg.GetReferences()
    pEnumJDArgument.Reset
    Do
        pEnumJDArgument.Next 1, arg1, found
        If found = 0 Then Exit Do
        Select Case arg1.Index
            Case 2
                    'The input is the plate geometry before cutout
                    'Set logical edge as sub-input if its 2D geometry is constrained
                    Dim proxyEdge As Object
                    Dim ProxyName As String
                    Dim SymbolHelper As Object
                    Set SymbolHelper = SymbolDefinition_AssistantGetObject(pSymbolDef, "SymbolHelper")
                    Set oInput = pSymbolDef.IJDInputs.GetInputByIndex(arg1.Index)
                    Set UserListOfDispatch = SymbolDefinition_InputAssistantGetUserListDG(oInput)
                    Set RAD2DkeyColl = SymbolDefinition_InputAssistantGetRAD2DListDG(oInput)
                    
                    'Treat the case a constraint relation has been removed, we have to kill the input relation
                    Index = 0
                    SubCount = oInput.IJDInputs.InputCount
                    If SubCount <> 0 Then
                       For LoopCount = SubCount To 1 Step -1
                         Set oSubInput = oInput.IJDInputs.GetInputAtIndex(LoopCount)
                         If Not IsConstrainedOrDeleted(oSubInput.Key, aSheet.DrawingObjects) Then
                             'Remove subinput from the symbol's definition
                             oInput.IJDInputs.RemoveInput LoopCount
                         End If
                         Set oSubInput = Nothing
                       Next
                    End If
                    'Now look for new
                    For LoopOnKey = RAD2DkeyColl.Count To 1 Step -1
                        If IsConstrainedOrDeleted(RAD2DkeyColl.Item(LoopOnKey), aSheet.DrawingObjects) Then
                            Set oSubInput = New DInput
                            'The geometry is constrained: Create a sub-input description for
                            'recompute. The sub-input has the property PENDING because we want to manage
                            'ourself the binding and the conversion
                             oSubInput.Properties = igDESCRIPTION_PENDING
                             oSubInput.Key = RAD2DkeyColl.Item(LoopOnKey)
                             Index = Index + 1
                             Set proxyEdge = UserListOfDispatch.Item(LoopOnKey)
                             oSubInput.Name = SymbolHelper.GetProxyName(arg1.Entity, proxyEdge)
                             'Set subinput in the definition
                             oInput.IJDInputs.SetInput oSubInput
                             Set oSubInput = Nothing
                             Set proxyEdge = Nothing
                       End If
                    Next
                   ' pSymbolDef.IJDInputs.SetInput oInput, arg1.Index
                    Set RAD2DkeyColl = Nothing
                    Set UserListOfDispatch = Nothing
            Case 3
                    'If input index = 3, deals with struct objects to intersect
                    ' Dispatchs are set as sub-inputs if they are constrained in RAD2D
                    Set oInput = pSymbolDef.IJDInputs.GetInputByIndex(arg1.Index)
                    Set SubReferenceCollOthers = arg1.Entity
                    Set UserListOfDispatch = SymbolDefinition_InputAssistantGetUserListDG(oInput)
                    Set RAD2DkeyColl = SymbolDefinition_InputAssistantGetRAD2DListDG(oInput)
                    If oInput.Key = "" Then oInput.Key = 0
                    Index = oInput.Key
                     'Treat the case a constraint relation has been removed, we have to kill the input relation
                    SubCount = oInput.IJDInputs.InputCount
                    If SubCount <> 0 Then
                       
                       For LoopCount = SubCount To 1 Step -1
                         Set oSubInput = oInput.IJDInputs.GetInputAtIndex(LoopCount)
                         If Not IsConstrainedOrDeleted(oSubInput.Key, aSheet.DrawingObjects) Then
                             'The RAD2D object is no more constrained, Remove GSCAD object from Symbol's
                             'Reference Collection
                             SubReferenceCollOthers.IJDEditJDArgument.RemoveByIndex (oSubInput.Name)
                             'Remove subinput in the definition
                             oInput.IJDInputs.RemoveInput (oSubInput.Name)
                         End If
                         Set oSubInput = Nothing
                       Next
                    End If
                    If Not RAD2DkeyColl Is Nothing Then
                        For LoopOnKey = RAD2DkeyColl.Count To 1 Step -1
                            If IsConstrainedOrDeleted(RAD2DkeyColl.Item(LoopOnKey), aSheet.DrawingObjects) Then
                                Set oSubInput = New DInput
                                'The geometry is constrained: Create a sub-input description for
                                'recompute and set the associated dispatch object in the REf Coll
                                'with the appropriate interface
                                oSubInput.Properties = igDESCRIPTION_OPTIONAL
                                oSubInput.Key = RAD2DkeyColl.Item(LoopOnKey)
                                 Index = Index + 1
                                 oSubInput.Name = Index
                                 'Set object as argument of ReferenceCollection
                                 SubReferenceCollOthers.IJDEditJDArgument.SetEntity Index, UserListOfDispatch.Item(LoopOnKey), STRUCT_RELATION_SUBINPUT_IID, STRUCT_RELATION_SUBINPUT_COLNAME
                                 'Set subinput in the definition
                                 oInput.IJDInputs.SetInput oSubInput
                                 Set oSubInput = Nothing
                           End If
                        Next
                        
                        oInput.Key = Index
                    End If
                    
                        'pSymbolDef.IJDInputs.SetInput oInput, arg1.Index
                    
                    Set RAD2DkeyColl = Nothing
                    Set UserListOfDispatch = Nothing
                    Set oInput = Nothing
            Case 4
                    'if input index = 4, deals with reference planes
                    'Dispatchs are set as sub-inputs if they are constrained in RAD2D
                    Set oInput = pSymbolDef.IJDInputs.GetInputByIndex(arg1.Index)
                    Set SubReferenceCollRefPlanes = arg1.Entity
                    Set UserListOfDispatch = SymbolDefinition_InputAssistantGetUserListDG(oInput)
                    Set RAD2DkeyColl = SymbolDefinition_InputAssistantGetRAD2DListDG(oInput)
                    If oInput.Key = "" Then oInput.Key = 0
                    Index = oInput.Key
                     'Treat the case a constraint relation has been removed, we have to kill the input relation
                    SubCount = oInput.IJDInputs.InputCount
                    If SubCount <> 0 Then
                       For LoopCount = SubCount To 1 Step -1
                         Set oSubInput = oInput.IJDInputs.GetInputAtIndex(LoopCount)
                         If Not IsConstrainedOrDeleted(oSubInput.Key, aSheet.DrawingObjects) Then
                             'The RAD2D object is no more constrained, Remove GSCAD object from Symbol's
                             'Reference Collection
                             SubReferenceCollRefPlanes.IJDEditJDArgument.RemoveByIndex (oSubInput.Name)
                             'Remove subinput in the definition
                             oInput.IJDInputs.RemoveInput (oSubInput.Name)
                         End If
                         Set oSubInput = Nothing
                       Next
                    End If
                    If Not RAD2DkeyColl Is Nothing Then
                        For LoopOnKey = RAD2DkeyColl.Count To 1 Step -1
                            If IsConstrainedOrDeleted(RAD2DkeyColl.Item(LoopOnKey), aSheet.DrawingObjects) Then
                                Set oSubInput = New DInput
                                'The geometry is constrained: Create a sub-input description for
                                'recompute and set the associated dispatch object in the REf Coll
                                'with the appropriate interface
                                oSubInput.Properties = igDESCRIPTION_OPTIONAL
                                oSubInput.Key = RAD2DkeyColl.Item(LoopOnKey)
                                 Index = Index + 1
                                 oSubInput.Name = Index
                                 'Set object as argument of ReferenceCollection
                                 SubReferenceCollRefPlanes.IJDEditJDArgument.SetEntity Index, UserListOfDispatch.Item(LoopOnKey), STRUCT_RELATION_SUBINPUT_IID, STRUCT_RELATION_SUBINPUT_COLNAME
                                 'Set subinput in the definition
                                 oInput.IJDInputs.SetInput oSubInput
                                 Set oSubInput = Nothing
                           End If
                        Next
                        
                        oInput.Key = Index
                    End If
                    'pSymbolDef.IJDInputs.SetInput oInput, arg1.Index
                    Set RAD2DkeyColl = Nothing
                    Set UserListOfDispatch = Nothing
                    Set oInput = Nothing
            
            Case 5
                    'treat shared contours
                    Set oInput = pSymbolDef.IJDInputs.GetInputByIndex(arg1.Index)
                    If Not oInput Is Nothing Then
                        Set UserListOfDispatch = SymbolDefinition_InputAssistantGetUserListDG(oInput)
                        Set RAD2DkeyColl = SymbolDefinition_InputAssistantGetRAD2DListDG(oInput)
                        If Not RAD2DkeyColl Is Nothing Then
                            For LoopOnKey = RAD2DkeyColl.Count To 1 Step -1
                               Set RadObject = aSheet.DrawingObjects(RAD2DkeyColl.Item(LoopOnKey))
                               RadObject.Delete
                            Next
                        End If
                    End If
            Case Else
                
        End Select
    Loop
                
    'Clean
    Set SubReferenceCollOthers = Nothing
    Set SubReferenceCollRefPlanes = Nothing
    Set ServerDoc = Nothing
    Set MainRefColl = Nothing

Exit Sub
ErrorHandler: HandleError MODULE, MT
End Sub
Public Function CheckThatToolbarCtlExists(ToolbarName As String, ToolbarCtlProgID, pApplication As RAD2D.Application) As Boolean
    Dim oToolBar As RAD2D.ToolBar
    Dim oToolbarCtl As RAD2D.ToolbarControl
    
    CheckThatToolbarCtlExists = False
    For Each oToolBar In pApplication.ToolBars
        If oToolBar.Name = ToolbarName Then
            For Each oToolbarCtl In oToolBar.Controls
                If oToolbarCtl.DLLName = ToolbarCtlProgID Then
                    CheckThatToolbarCtlExists = True
                    Exit Function
                End If
            Next
        End If
    Next

End Function
Public Function FindRevolutionAxisDefined(pAttributeSetName As String, pSheet As RAD2D.Sheet) As Object
Const MT = "FindRevolutionAxisDefined"
On Error GoTo ErrorHandler

'this sub returns either a line or a group which owns the attribute set "RevolutionAttributeSet"
    Dim oline2d As RAD2D.Line2d
    Dim oGroup As RAD2D.group
    
    Set oline2d = FindLine2DByNamedAttributeSet(pSheet.Lines2d, pAttributeSetName)
    If oline2d Is Nothing Then
        Set oGroup = FindGroupOwningLineWithNamedAttributeSet(pSheet.Groups, pAttributeSetName)
        If Not oGroup Is Nothing Then 'The axis belonged to a group
            Set FindRevolutionAxisDefined = oGroup
        Else 'The line belongs to a complexstring2d, just returns the line
            Set oline2d = FindLineInComplexString2dWithNamedAttributeSet(pSheet.ComplexStrings2d, pAttributeSetName)
            If Not oline2d Is Nothing Then Set FindRevolutionAxisDefined = oline2d
        End If
    Else 'The axis was a line that didn't belong to a group
        Set FindRevolutionAxisDefined = oline2d
    End If
Exit Function
ErrorHandler:
    HandleError MODULE, MT
End Function

Public Function SymbolDefinition_AssistantGetObject(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition, Key As String) As Object
    Dim oAssistantCollection As Object
    Set oAssistantCollection = pSymbolDefinition.IJDDefinitionPlayerEx.Assistant
    Set SymbolDefinition_AssistantGetObject = oAssistantCollection.Item(Key)
End Function
Public Function SymbolDefinition_InputAssistantGetUserListDG(pInput As IMSSymbolEntities.IJDInput) As Object
    Const MT = "SymbolDefinition_InputAssistantGetUserListDG"
    On Error GoTo ErrorHandler
    Dim oAssistantColl As Object
    Set oAssistantColl = pInput.IJDInputDuringGame.Assistant
    If Not oAssistantColl Is Nothing Then
        On Error Resume Next
        Set SymbolDefinition_InputAssistantGetUserListDG = oAssistantColl.Item("UserList")
    End If
Exit Function
ErrorHandler: HandleError MODULE, MT & " " & "index is " & pInput.Index
End Function
Public Function SymbolDefinition_InputAssistantGetRAD2DListDG(pInput As IMSSymbolEntities.IJDInput) As Object
    Const MT = "SymbolDefinition_InputAssistantGetRAD2DListDG"
    On Error GoTo ErrorHandler

    Dim oAssistantColl As Object
    Set oAssistantColl = pInput.IJDInputDuringGame.Assistant
    If Not oAssistantColl Is Nothing Then
        On Error Resume Next
        Set SymbolDefinition_InputAssistantGetRAD2DListDG = oAssistantColl.Item("RAD2DList")
    End If
Exit Function
ErrorHandler: HandleError MODULE, MT & " " & "index is " & pInput.Index
End Function

Public Function IsConstrainedOrDeleted(RAD2DKey As String, ObjsOnSheet As RAD2D.DrawingObjects) As Boolean
Const MT = "IsConstrainedOrDeleted"
On Error GoTo ErrorHandler
    Dim BeginPosFileName As Long
    Dim RadObject As Object
    BeginPosFileName = 0
    'Look if the key is made of RAD keys concatenated with "|"
    BeginPosFileName = InStr(1, RAD2DKey, "|", vbTextCompare)
    IsConstrainedOrDeleted = False
    
     If BeginPosFileName <> 0 Then 'It may be a composite key
         Dim CompositeKey As String
         Dim pLength As Long
         Dim pos As Long
         Dim begin As Long
         Dim range As Long
         Dim Key As String
         
         CompositeKey = RAD2DKey
         pLength = Len(CompositeKey)
         begin = 2
         pos = -1
         While pos <> pLength
             pos = InStr(begin, CompositeKey, "|")
             range = pos - begin
             Key = Mid(CompositeKey, begin, range)
             Set RadObject = ObjsOnSheet.Item(Key)
             If Not RadObject Is Nothing Then
                If RADObjectIsConstrained(RadObject) Then
                     IsConstrainedOrDeleted = True
                 End If
             End If
             begin = pos + 1
         Wend

         If Not IsConstrainedOrDeleted Then
             DeleteRADObjectWithCompositeKey RAD2DKey, ObjsOnSheet
         End If

    Else
         Set RadObject = ObjsOnSheet(RAD2DKey)
         If Not RadObject Is Nothing Then
            If Not RADObjectIsConstrained(RadObject) Then
                RadObject.Delete
            Else
                IsConstrainedOrDeleted = True
            End If
         End If
    End If
    Exit Function
ErrorHandler:
    HandleError MODULE, MT
End Function
Public Function FindLine2DByNamedAttributeSet(pLines2d As RAD2D.Lines2d, pNamedAttributeSet As String) As RAD2D.Line2d
Const MT = "FindLine2DByNamedAttributeSet"
On Error GoTo ErrorHandler

    Dim oAttributeSet As RAD2D.AttributeSet
    Dim oline2d As RAD2D.Line2d
    
    For Each oline2d In pLines2d
            For Each oAttributeSet In oline2d.AttributeSets
                If oAttributeSet.SetName = pNamedAttributeSet Then
                    Set FindLine2DByNamedAttributeSet = oline2d
                    Exit Function
                End If
            Next
       
    Next
    
   Exit Function
ErrorHandler:
    HandleError MODULE, MT
End Function

Public Function FindGroupOwningLineWithNamedAttributeSet(pGroups As RAD2D.Groups, pNamedAttributeSet As String) As RAD2D.group
Const MT = "FindGroupOwningLineWithNamedAttributeSet"
On Error GoTo ErrorHandler
    Dim group As RAD2D.group
    Dim ObjInGroup As Object
    Dim Index As Long
    Dim oAttributeSet As RAD2D.AttributeSet
    
        For Each group In pGroups
            For Index = 1 To group.Count
                Set ObjInGroup = group.Item(Index)
                If ObjInGroup.Type = igLine2d Then
                    For Each oAttributeSet In ObjInGroup.AttributeSets
                        If oAttributeSet.SetName = pNamedAttributeSet Then
                            Set FindGroupOwningLineWithNamedAttributeSet = group
                            Exit Function
                        End If
                    Next
                End If
            Next
        Next
        
Exit Function
ErrorHandler:
    HandleError MODULE, MT
End Function

Public Function FindLineInComplexString2dWithNamedAttributeSet(pRADCompStrings2d As RAD2D.ComplexStrings2d, pNamedAttributeSet As String) As RAD2D.Line2d
Const MT = "FindLineInComplexString2dWithNamedAttributeSet"
On Error GoTo ErrorHandler
    Dim CS2d As RAD2D.ComplexString2d
    Dim ObjInCS2D As Object
    Dim Index As Long
    Dim oAttributeSet As RAD2D.AttributeSet
    Dim Drawings As RAD2D.DrawingObjects
    
        For Each CS2d In pRADCompStrings2d
            Set Drawings = CS2d.DrawingObjects
            For Each ObjInCS2D In Drawings
                If ObjInCS2D.Type = igLine2d Then
                    For Each oAttributeSet In ObjInCS2D.AttributeSets
                        If oAttributeSet.SetName = pNamedAttributeSet Then
                            Set FindLineInComplexString2dWithNamedAttributeSet = ObjInCS2D
                            Exit Function
                        End If
                    Next
                End If
            Next
        Next
Exit Function
ErrorHandler:
    HandleError MODULE, MT
End Function
Public Function RADObjectIsConstrained(pRADObj As Object) As Boolean
Const MT = "Line2dhasConstraints"
On Error GoTo ErrorHandler

    Dim objRelation As Object
    Dim relType As RAD2D.ObjectType
    Dim oCollection As New Collection
    Dim objRelations As RAD2D.Relationships2d
    
    RADObjectIsConstrained = False
    On Error Resume Next
    Set objRelations = pRADObj.Relationships
    If Not objRelations Is Nothing Then
        For Each objRelation In objRelations
            relType = objRelation.Type
            If relType <> igFixRelation2d Then
                RADObjectIsConstrained = True
            End If
        Next
    End If
     
    Exit Function
ErrorHandler:
    HandleError MODULE, MT
End Function
Public Sub DeleteRADObjectWithCompositeKey(RAD2DKey As String, pDrawObjColl As RAD2D.DrawingObjects)
Const MT = "DeleteRADObjectWithCompositeKey"
On Error GoTo ErrorHandler
    Dim CompositeKey As String
    Dim pLength As Long
    Dim pos As Long
    Dim begin As Long
    Dim range As Long
    Dim Key As String
    Dim RadObject As Object
    
     CompositeKey = RAD2DKey
     pLength = Len(CompositeKey)
     begin = 2
     pos = -1
     While pos <> pLength
        pos = InStr(begin, CompositeKey, "|")
        range = pos - begin
        Key = Mid(CompositeKey, begin, range)
        begin = pos + 1
        Set RadObject = pDrawObjColl.Item(Key)
        If Not RadObject Is Nothing Then RadObject.Delete
     Wend
     
    Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub
Public Function SymbolDefinition_GetSymbolArgumentAtIndex(Index As Long, pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition) As Object
Const MT = "SymbolDefinition_GetSymbolArgument"
On Error GoTo ErrorHandler
    'With this sub, we retrieve the inputs of the symbol of a cutout plate
    Dim pEnumJDArgument As IEnumJDArgument
    Dim arg1 As IJDArgument
    Dim found As Long
    Dim oRefColl As IJDReferencesCollection
    Dim Inp As IJDInput
    Dim oSymbol As IJDSymbol
            
         Set oSymbol = pSymbolDefinition.IJDDefinitionPlayerEx.PlayingSymbol
         Set pEnumJDArgument = oSymbol.IJDReferencesArg.GetReferences()
         pEnumJDArgument.Reset
         Do
             pEnumJDArgument.Next 1, arg1, found
             If found = 0 Then Exit Do
             If arg1.Index = Index Then      'OBM : gets objects that were inputs of root plate
                Set SymbolDefinition_GetSymbolArgumentAtIndex = arg1.Entity
                Exit Do
             End If
         Loop
        
        'release
        Set oRefColl = Nothing
        Set pEnumJDArgument = Nothing
Exit Function
ErrorHandler: HandleError MODULE, MT & ", index of input is " & Index
End Function

Private Sub OrderObjectsInGroup(group As RAD2D.group, oColl As Collection)
Const MT = "OrderObjectsInGroup"
On Error GoTo ErrorHandler
    Dim oObjectColl As New Collection
    Dim tempColl As New Collection
    Dim tempColl2 As New Collection
    Dim firstObject As Object
    
    Dim currentObject As Object
    Dim bFound As Boolean
    Dim ii As Integer
    Dim x As Double
    Dim y As Double
    Dim xx As Double, yy As Double
    Dim oTemp As Object
    Dim sx As Double
    Dim sy As Double
    Dim ex As Double
    Dim ey As Double
    
    Dim objAssocSegStyle As SEGSTYLELib.AsSegStyle
    Dim oBspline As RAD2D.BSplineCurve2d
    Dim oGeomBsp As RAD2D.BSplineCurve2d
    Dim ocurves As RAD2D.BSplineCurves2d
    Dim oActiveSheet As Sheet
    
    Set oActiveSheet = group.Document.ActiveSheet
    
    For Each currentObject In group
            'MsgBox currentObject.Name
            On Error Resume Next
            Set oBspline = currentObject
            On Error GoTo ErrorHandler
            If Not oBspline Is Nothing Then
                Set objAssocSegStyle = oBspline.ControllingSegmentedStyle
                If Not objAssocSegStyle Is Nothing Then
                    Set ocurves = oActiveSheet.BSplineCurves2d
                    Set oGeomBsp = ocurves.AddByGeometry(oBspline.GetGeometry)
                    objAssocSegStyle.SetTrimGeometry oGeomBsp
                    Set currentObject = oGeomBsp
                End If
            End If
            If currentObject.Type = igLine2d Or currentObject.Type = igArc2d Or currentObject.Type = igBsplineCurve2d _
            Or currentObject.Type = igCircle2d Or currentObject.Type = _
            igComplexString2d Or currentObject.Type = igEllipse2d Then
                oObjectColl.Add currentObject
            End If
            Set oBspline = Nothing
            Set currentObject = Nothing
    Next currentObject
    
    If oObjectColl.Count <> 0 Then
        Set currentObject = oObjectColl.Item(1)
        Set firstObject = currentObject
        If currentObject.Type = igBsplineCurve2d Then
            Set oGeomBsp = currentObject
            oGeomBsp.GetGeometry.GetEndPoints xx, yy, x, y
        Else
            currentObject.GetEndPoint x, y
        End If
        
        oObjectColl.Remove 1
        tempColl.Add currentObject
        
        bFound = True
        While oObjectColl.Count > 0 And bFound
            bFound = False
            
            For ii = 1 To oObjectColl.Count
                        
                Set oTemp = oObjectColl.Item(ii)
                If oTemp.Type = igBsplineCurve2d Then
                    Set oGeomBsp = Nothing
                    Set oGeomBsp = oTemp
                    oGeomBsp.GetGeometry.GetEndPoints sx, sy, ex, ey
                Else
                    oTemp.GetStartPoint sx, sy
                    oTemp.GetEndPoint ex, ey
                End If
                ' if the coordinates are equal to the object start point
                If (Abs(x - sx) < 0.000001 And Abs(y - sy) < 0.000001) Then
                    oObjectColl.Remove ii
                    'set current coordinate (X,Y) to object end point
                    x = ex
                    y = ey
                    Set currentObject = oTemp
                    bFound = True
                    Exit For
                ' if the coordinates are equal to the object end point
                ElseIf (Abs(x - ex) < 0.000001 And Abs(y - ey) < 0.000001) Then
                    oObjectColl.Remove ii
                    'set current coordinate (X,Y) to object start point
                    x = sx
                    y = sy
                    Set currentObject = oTemp
                    bFound = True
                    Exit For
                End If
            Next ii
        
            If bFound Then
                ' add the object into the list
                tempColl.Add currentObject
            End If
        Wend
        
        ' test if the first item in the collection, is in the middle of the graphic
        If Not bFound Then
            Set currentObject = firstObject
            currentObject.GetStartPoint x, y
            
            bFound = True
    
            
            While oObjectColl.Count > 0 And bFound
                bFound = False
                
                For ii = 1 To oObjectColl.Count
                    Set oTemp = oObjectColl.Item(ii)
                    If oTemp.Type = igBsplineCurve2d Then
                        Set oGeomBsp = Nothing
                        Set oGeomBsp = oTemp
                        oGeomBsp.GetGeometry.GetEndPoints sx, sy, ex, ey
                    Else
                        oTemp.GetStartPoint sx, sy
                        oTemp.GetEndPoint ex, ey
                    End If
                    
                    ' if the coordinates are equal to the object start point
                    If (Abs(x - sx) < 0.000001 And Abs(y - sy) < 0.000001) Then
                        oObjectColl.Remove ii
                        'set current coordinate (X,Y) to object end point
                        x = ex
                        y = ey
                        Set currentObject = oTemp
                        bFound = True
                        Exit For
                    ' if the coordinates are equal to the object end point
                    ElseIf (Abs(x - ex) < 0.000001 And Abs(y - ey) < 0.000001) Then
                        oObjectColl.Remove ii
                        'set current coordinate (X,Y) to object start point
                        x = sx
                        y = sy
                        Set currentObject = oTemp
                        bFound = True
                        Exit For
                    End If
                Next ii
                
                If bFound Then
                    ' add the object into the list
                    tempColl2.Add currentObject
                End If
            Wend
        End If
        
        Dim i As Long
        
        For i = tempColl2.Count To 1 Step -1
            oColl.Add tempColl2(i)
        Next i
        
        For Each currentObject In tempColl
            oColl.Add currentObject
        Next currentObject
    End If
    Exit Sub
ErrorHandler:
  'HandleError MODULE, MT
  MsgBox "Function OrderObjectsInGroup: Contour (RAD2D group) could not be evaluated after recompute"
End Sub

Private Sub OrderElemsInRAD2DSymbol(ByRef SymbolDrawings As RAD2D.DrawingObjects, ByRef oColl As Collection)
Const MT = "GetOrderSymbolDrawing"
On Error GoTo ErrorHandler
    Dim oObjectColl As New Collection
    Dim tempColl As New Collection
    Dim tempColl2 As New Collection
    Dim firstObject As Object
    
    Dim currentObject As Object
    Dim bFound As Boolean
    Dim ii As Integer
    Dim x As Double
    Dim y As Double
    Dim xx As Double, yy As Double
    Dim oTemp As Object
    Dim sx As Double
    Dim sy As Double
    Dim ex As Double
    Dim ey As Double
    
    Dim objAssocSegStyle As SEGSTYLELib.AsSegStyle
    Dim oBspline As RAD2D.BSplineCurve2d
    Dim oGeomBsp As RAD2D.BSplineCurve2d
    Dim ocurves As RAD2D.BSplineCurves2d
    Dim oActiveSheet As RAD2D.Sheet
    Dim oApp As RAD2D.Application
    
    Set oApp = SymbolDrawings.Application.radapplication
    Set oActiveSheet = oApp.ActiveDocument.ActiveSheet
    
    For Each currentObject In SymbolDrawings
            'MsgBox currentObject.Name
            On Error Resume Next
            Set oBspline = currentObject
            On Error GoTo ErrorHandler
            If Not oBspline Is Nothing Then
                Set objAssocSegStyle = oBspline.ControllingSegmentedStyle
                If Not objAssocSegStyle Is Nothing Then
                    Set ocurves = oActiveSheet.BSplineCurves2d
                    Set oGeomBsp = ocurves.AddByGeometry(oBspline.GetGeometry)
                    objAssocSegStyle.SetTrimGeometry oGeomBsp
                    Set currentObject = oGeomBsp
                End If
            End If
            If currentObject.Type = igLine2d Or currentObject.Type = igArc2d Or currentObject.Type = igBsplineCurve2d _
            Or currentObject.Type = igCircle2d Or currentObject.Type = _
            igComplexString2d Or currentObject.Type = igEllipse2d Then
                oObjectColl.Add currentObject
            End If
            Set oBspline = Nothing
            Set currentObject = Nothing
    Next currentObject
    
    If oObjectColl.Count <> 0 Then
        Set currentObject = oObjectColl.Item(1)
        Set firstObject = currentObject
        If currentObject.Type = igBsplineCurve2d Then
            Set oGeomBsp = currentObject
            oGeomBsp.GetGeometry.GetEndPoints xx, yy, x, y
        Else
            currentObject.GetEndPoint x, y
        End If
        
        oObjectColl.Remove 1
        tempColl.Add currentObject
        
        bFound = True
        While oObjectColl.Count > 0 And bFound
            bFound = False
            
            For ii = 1 To oObjectColl.Count
                        
                Set oTemp = oObjectColl.Item(ii)
                If oTemp.Type = igBsplineCurve2d Then
                    Set oGeomBsp = Nothing
                    Set oGeomBsp = oTemp
                    oGeomBsp.GetGeometry.GetEndPoints sx, sy, ex, ey
                Else
                    oTemp.GetStartPoint sx, sy
                    oTemp.GetEndPoint ex, ey
                End If
                ' if the coordinates are equal to the object start point
                If (Abs(x - sx) < 0.000001 And Abs(y - sy) < 0.000001) Then
                    oObjectColl.Remove ii
                    'set current coordinate (X,Y) to object end point
                    x = ex
                    y = ey
                    Set currentObject = oTemp
                    bFound = True
                    Exit For
                ' if the coordinates are equal to the object end point
                ElseIf (Abs(x - ex) < 0.000001 And Abs(y - ey) < 0.000001) Then
                    oObjectColl.Remove ii
                    'set current coordinate (X,Y) to object start point
                    x = sx
                    y = sy
                    Set currentObject = oTemp
                    bFound = True
                    Exit For
                End If
            Next ii
        
            If bFound Then
                ' add the object into the list
                tempColl.Add currentObject
            End If
        Wend
        
        ' test if the first item in the collection, is in the middle of the graphic
        If Not bFound Then
            Set currentObject = firstObject
            currentObject.GetStartPoint x, y
            
            bFound = True
    
            
            While oObjectColl.Count > 0 And bFound
                bFound = False
                
                For ii = 1 To oObjectColl.Count
                    Set oTemp = oObjectColl.Item(ii)
                    If oTemp.Type = igBsplineCurve2d Then
                        Set oGeomBsp = Nothing
                        Set oGeomBsp = oTemp
                        oGeomBsp.GetGeometry.GetEndPoints sx, sy, ex, ey
                    Else
                        oTemp.GetStartPoint sx, sy
                        oTemp.GetEndPoint ex, ey
                    End If
                    
                    ' if the coordinates are equal to the object start point
                    If (Abs(x - sx) < 0.000001 And Abs(y - sy) < 0.000001) Then
                        oObjectColl.Remove ii
                        'set current coordinate (X,Y) to object end point
                        x = ex
                        y = ey
                        Set currentObject = oTemp
                        bFound = True
                        Exit For
                    ' if the coordinates are equal to the object end point
                    ElseIf (Abs(x - ex) < 0.000001 And Abs(y - ey) < 0.000001) Then
                        oObjectColl.Remove ii
                        'set current coordinate (X,Y) to object start point
                        x = sx
                        y = sy
                        Set currentObject = oTemp
                        bFound = True
                        Exit For
                    End If
                Next ii
                
                If bFound Then
                    ' add the object into the list
                    tempColl2.Add currentObject
                End If
            Wend
        End If
        
        Dim i As Long
        
        For i = tempColl2.Count To 1 Step -1
            oColl.Add tempColl2(i)
        Next i
        
        For Each currentObject In tempColl
            oColl.Add currentObject
        Next currentObject
    End If
    Exit Sub

ErrorHandler:
    HandleError MODULE, MT
End Sub

Function GetPlaneFromSlot(oSlotSym As Object) As IJPlane
 Const MT = "GetPlaneFromSlot"
On Error GoTo ErrorHandler
    
    Dim oSymbol                  As IMSSymbolEntities.IJDSymbol

'''
    Dim oSymbolDefinition        As IMSSymbolEntities.IJDSymbolDefinition
    Dim oSymbDefReps As IJDRepresentations
    Dim oSymbDefRep As IJDRepresentation
    Dim RepresentationName As String
    
    Set oSymbol = oSlotSym
    Set oSymbolDefinition = oSymbol.IJDSymbolDefinition(1)
    Set oSymbDefReps = oSymbolDefinition.IJDRepresentations
    RepresentationName = "SlotPerpendicularToFace"
    Set oSymbDefRep = oSymbDefReps.GetRepresentationByName(RepresentationName)
    
    Dim nOutputCount As Long
    Dim nOutputloop As Long
    Dim oIJDOutputs As IJDOutputs
    Dim oOutput As Object
    Dim oRepOutput As IJDOutput
    Dim sOutputName As String

    Set oIJDOutputs = oSymbDefRep
    
    nOutputCount = oIJDOutputs.OutputCount
    
    On Error Resume Next
    ' loop on outputs
    For nOutputloop = 1 To nOutputCount
        ' find output name
        Set oRepOutput = oIJDOutputs.GetOutputAtIndex(nOutputloop)
        sOutputName = oRepOutput.Name
        ''' DBG MsgBox "slot output name:" & sOutputName
        Set oRepOutput = Nothing
        If InStr(sOutputName, "Plane") > 0 Then
            ' binds to output from its name
            Set oOutput = oSymbol.BindToOutput(RepresentationName, sOutputName)
            ' add to collection
            If Not oOutput Is Nothing Then
    '''        OutputElements.Add oOutput
                Dim oIJPlane As IJPlane
                 Set oIJPlane = oOutput
                 If Not oIJPlane Is Nothing Then
                    ''' DBG MsgBox "found slot's plane"
'''                    oInputDesc.IJDInputDuringGame.Assistant = oIJPlane
'''                    oStructConverter.SketchingPlane = oIJPlane
                    Set GetPlaneFromSlot = oIJPlane
                    Exit For
                 Else
                    MsgBox "no ijplane on Slot"
                End If
            ' release output
                Set oOutput = Nothing
            Else
                MsgBox "no output from Slot"
            End If
'
        End If  ' the plane
'
    Next nOutputloop

 Exit Function
'
ErrorHandler:
    HandleError MODULE, MT

End Function
 