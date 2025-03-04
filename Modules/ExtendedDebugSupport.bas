Attribute VB_Name = "ExtendedDebugSupport"
Option Explicit
' depends on DebugSupport (IDebugSupport)
' depends on IMSDObject (IJNamedItem)
' depends on scrrun (FileSystemObject)
    
Private g_pDebugSupport0 As IDebugSupport

' private variable to record what is to debug
Private DEBUG_SUPPORT As Boolean
Private g_bIsDebugOn As Boolean

Private g_sDebugSource As String
Private g_sDebugSources(1 To 100) As String
Private g_iCountOfDebugSources As Integer

Private g_pDebugSupport As IDebugSupport
Private g_pDebugSupports(1 To 100) As IDebugSupport

Private g_sListOfProjectsToDebug() As String
Private g_sListOfClassesToDebug() As String
Private g_sListOfMethodsToDebug() As String

Private g_iCountOfCalledMethods As Integer
Private g_sListOfCalledMethods(1 To 100) As String
'
' public methods
'
Public Sub DebugStart(sSource As String)
    If g_pDebugSupport0 Is Nothing Then
        Set g_pDebugSupport0 = New DebugSupport
        If g_pDebugSupport0.DEBUG_REGISTRY_VALUE = 1 Then Let DEBUG_SUPPORT = True Else Let DEBUG_SUPPORT = False
        If DEBUG_SUPPORT Then Call DebugInitListsToDebug
    End If
    
    If DEBUG_SUPPORT Then
        If DebugExtractSource(sSource) <> "" Then
            If sSource <> g_sDebugSource Then
                Set g_pDebugSupport = DebugGetDebugSupportFromDebugSource(sSource)
                Let g_sDebugSource = sSource
            End If
            ' dimension to the maximum length pf a progid (39 characters + 1)
            Dim sLongNameOfDebugSource As String: Let sLongNameOfDebugSource = "                                        "
            Mid(sLongNameOfDebugSource, 1, Len(sSource)) = sSource
            Let g_pDebugSupport.DEBUG_SOURCE = sLongNameOfDebugSource
            Call DebugIn(sSource + "::" + "Start")
        Else
            Call DebugValue("Invalid syntax for source (expected 'project.class')", sSource)
        End If
    End If
End Sub
Public Sub DebugStop(sSource As String)
    If DEBUG_SUPPORT Then
        Call DebugOut(sSource + "::" + "Stop")
'''        Call DebugOut(DebugExtractSource(DebugGetLastCalledMethod) + "::" + "Stop")
'''        Let g_sDebugSource = ""
'''
'''        Dim sCallingMethod As String: Let sCallingMethod = DebugGetLastCalledMethod()
'''        If sCallingMethod <> "" Then
'''            Dim sSource As String: Let sSource = DebugExtractSource(sCallingMethod)
'''            If sSource <> g_sDebugSource Then
'''                Set g_pDebugSupport = DebugGetDebugSupportFromDebugSource(sSource)
'''                Let g_sDebugSource = sSource
'''            End If
'''        End If
    End If
End Sub
Public Sub DebugIn(sCalledMethod As String)
    If DEBUG_SUPPORT Then
        Dim sSource As String: Let sSource = DebugExtractSource(sCalledMethod)
        Dim sFullNameOfMethod As String:
        If sSource <> "" Then
            Let sFullNameOfMethod = sCalledMethod
        Else
            Let sSource = g_sDebugSource
            Let sFullNameOfMethod = g_sDebugSource + "::" + sCalledMethod
        End If
        
        If sSource <> g_sDebugSource Then
            Set g_pDebugSupport = DebugGetDebugSupportFromDebugSource(sSource)
            Let g_sDebugSource = sSource
        End If
        
        Dim sShortNameOfMethod As String: Let sShortNameOfMethod = DebugExtractMethod(sFullNameOfMethod)
        If sShortNameOfMethod <> "Start" Then Call DebugAddCalledMethod(sFullNameOfMethod)
        
        Let g_bIsDebugOn = IsAllToDebug() Or IsProjectToDebug() Or IsClassToDebug() Or IsMethodToDebug(sFullNameOfMethod)
        
        If g_bIsDebugOn Then
            If sShortNameOfMethod <> "" Then
                If sShortNameOfMethod = "Start" Then
                    Call g_pDebugSupport.DEBUG_MSG("++ " + sShortNameOfMethod)
                Else
                    Call g_pDebugSupport.DEBUG_MSG(">> " + sShortNameOfMethod)
                    Call g_pDebugSupport.DEBUG_DEEP_BEGIN
                End If
            Else
                Call DebugValue("Invalid syntax for called method (expected 'project.class::method' or 'method')", sCalledMethod)
            End If
        End If
    End If
End Sub
Public Sub DebugOut(Optional sCalledMethod As String = "")
    If DEBUG_SUPPORT Then
        Dim sShortNameOfMethod As String
        Dim sFullNameOfMethod As String: Let sFullNameOfMethod = sCalledMethod
        If sFullNameOfMethod = "" Then Let sFullNameOfMethod = DebugGetLastCalledMethod
        Let sShortNameOfMethod = DebugExtractMethod(sFullNameOfMethod)
            
        Dim sDebugSource As String: Let sDebugSource = DebugExtractSource(sFullNameOfMethod)
        If sDebugSource <> g_sDebugSource Then
            Set g_pDebugSupport = DebugGetDebugSupportFromDebugSource(sDebugSource)
            Let g_sDebugSource = sDebugSource
        End If
        
        Let g_bIsDebugOn = IsAllToDebug() Or IsProjectToDebug() Or IsClassToDebug() Or IsMethodToDebug(sFullNameOfMethod)
        If g_bIsDebugOn Then
            On Error Resume Next
            If sShortNameOfMethod = "Stop" Then
                Call g_pDebugSupport.DEBUG_MSG("-- " + sShortNameOfMethod)
            Else
                Call g_pDebugSupport.DEBUG_DEEP_END
                Call g_pDebugSupport.DEBUG_MSG("<< " + sShortNameOfMethod)
            End If
            On Error GoTo 0
        End If
        
        If sShortNameOfMethod <> "Stop" Then Call DebugRemoveCalledMethod
        
        Dim sCallingMethod As String: Let sCallingMethod = DebugGetLastCalledMethod()
        Dim sSource As String: Let sSource = DebugExtractSource(sCallingMethod)
        If sSource <> g_sDebugSource Then
            Set g_pDebugSupport = DebugGetDebugSupportFromDebugSource(sSource)
            Let g_sDebugSource = sSource
        End If
        If sCallingMethod <> "" Then
            Let g_bIsDebugOn = IsAllToDebug() Or IsProjectToDebug() Or IsClassToDebug() Or IsMethodToDebug(sCallingMethod)
        Else
            Let g_bIsDebugOn = False
        End If
    End If
End Sub
Public Sub DebugInput(sName As String, vValue As Variant)
    If DEBUG_SUPPORT Then
        If g_bIsDebugOn Then Call DebugValue("Input [" + sName + "]", vValue)
    End If
End Sub
Public Sub DebugOutput(sName As String, vValue As Variant)
    If DEBUG_SUPPORT Then
        If g_bIsDebugOn Then Call DebugValue("Output [" + sName + "]", vValue)
    End If
End Sub
Public Sub DebugValue(sMessage As String, vValue As Variant)
    If DEBUG_SUPPORT Then
        If g_bIsDebugOn Then
            Dim sValue As String: Let sValue = ""
            Select Case VarType(vValue)
                Case vbString: Let sValue = vValue
                Case vbInteger, vbLong: Dim lValue As Long: Let lValue = vValue: Let sValue = CStr(lValue)
                Case vbBoolean: Dim bValue As Boolean: Let bValue = vValue: If bValue Then Let sValue = "True" Else Let sValue = "False"
                Case vbDouble: Dim dValue As Double: Let dValue = vValue: Let sValue = CStr(Round(dValue, 6))
                Case vbObject, vbDataObject
                    If vValue Is Nothing Then
                        If VarType(vValue) = vbObject Then Let sValue = "Object is set to Nothing" Else Let sValue = "DataObject is not an Object"
                    Else
                        Dim oValue As Object
                        If VarType(vValue) = vbObject Then
                             Set oValue = vValue
                        Else
                            If TypeOf vValue Is IJDObject Then
                                Dim pObject As IJDObject:  Set pObject = vValue
                                Set oValue = pObject
                            ElseIf TypeOf vValue Is IMoniker Then
                                Dim pMoniker As IMoniker: Set pMoniker = vValue
                                Let sValue = Moniker_GetDisplayName(pMoniker)
                            End If
                        End If
    
                        If Not oValue Is Nothing Then
'''                            If TypeOf oValue Is IMoniker Then
'''                                Let sValue = Moniker_GetDisplayName(oValue)
'''                            Else
                            If TypeOf oValue Is IJNamedItem Then
                                Let sValue = NamedItem_GetName(oValue)
                            End If
                            If TypeOf oValue Is IJDPosition Then
                                Let sValue = sValue + Position_ToString(oValue)
                            ElseIf TypeOf oValue Is IJDVector Then
                                Let sValue = sValue + Vector_ToString(oValue)
                            ElseIf TypeOf oValue Is Collection Then
                                Dim pCollection As Collection: Set pCollection = oValue
                                Call DebugMsg(sMessage + ".Count= " + CStr(pCollection.Count))
                                Dim oItem As Object
'''                                For Each oItem In pCollection
                                Dim i As Integer
                                For i = 1 To pCollection.Count
                                    Set oItem = pCollection(i)
                                    Call DebugValue(sMessage + "[" + "]= ", oItem)
                                Next
                                sValue = "NoValue"
'''                            ElseIf TypeOf oValue Is IJMonikerElements Then
'''                                Dim pMonikerElements As IJMonikerElements: Set pMonikerElements = oValue
'''                                Call DebugMsg(sMessage + ".Count= " + CStr(pMonikerElements.Count))
''''''                                For Each oItem In pElements
'''                                For i = 1 To pMonikerElements.Count
''''''                                    Dim pMoniker As IMoniker
'''                                    Set pMoniker = pMonikerElements(i)
'''                                    Call DebugValue(sMessage + "[" + "]= ", pMoniker)
'''                                Next
'''                                sValue = "NoValue"
                            ElseIf TypeOf oValue Is IJElements Then
                                Dim pElements As IJElements: Set pElements = oValue
                                Call DebugMsg(sMessage + ".Count= " + CStr(pElements.Count))
'''                                For Each oItem In pElements
'''                                For i = 1 To pElements.Count
'''                                    Set oItem = pElements(i)
'''                                    Call DebugValue(sMessage + "[" + "]= ", oItem)
'''                                Next
                                sValue = "NoValue"
                            Else
                                Let sValue = "Unsupported type of object"
                            End If
                       End If
                    End If
                Case Else: Let sValue = "DebugInput: Unsupported variant type= " + CStr(VarType(vValue))
            End Select
            If sValue <> "NoValue" Then Call DebugMsg(sMessage + "= " + sValue)
        End If
    End If
End Sub
Public Sub DebugMsg(sMessage As String)
    If DEBUG_SUPPORT Then
        If g_bIsDebugOn Then Call g_pDebugSupport.DEBUG_MSG(sMessage)
    End If
End Sub
Public Sub DebugNamedItem(oObject As Object, sMessage As String)
    If DEBUG_SUPPORT Then
        If g_bIsDebugOn Then Call g_pDebugSupport.DEBUG_MSG(sMessage + "= " + NamedItem_GetName(oObject))
    End If
End Sub

'
' private methods
'
Private Function DebugAddCalledMethod(sMethod As String)
    If g_iCountOfCalledMethods < 100 Then
        Let g_iCountOfCalledMethods = g_iCountOfCalledMethods + 1
        g_sListOfCalledMethods(g_iCountOfCalledMethods) = sMethod ' g_sDebugSource + "::" + sMethod
    Else
        DebugMsg ("Stack of called methods > 100")
    End If
End Function
Private Sub DebugRemoveCalledMethod()
    If g_iCountOfCalledMethods > 0 Then
        g_sListOfCalledMethods(g_iCountOfCalledMethods) = ""
        Let g_iCountOfCalledMethods = g_iCountOfCalledMethods - 1
    Else
        DebugMsg ("Stack of called methods < 1")
    End If
End Sub
Private Function DebugGetLastCalledMethod() As String
    ' prepare result
    Dim sMethod As String: Let sMethod = ""
    
    If g_iCountOfCalledMethods > 0 Then Let sMethod = g_sListOfCalledMethods(g_iCountOfCalledMethods)

    ' return result
    Let DebugGetLastCalledMethod = sMethod
End Function
Private Function DebugExtractSource(sCalledMethod As String) As String
    ' prepare result
    Dim sSource As String: Let sSource = ""

    Dim iDelimiter As Integer: Let iDelimiter = InStr(1, sCalledMethod, "::", vbTextCompare)
    If iDelimiter > 0 Then Let sSource = Mid(sCalledMethod, 1, iDelimiter - 1) Else Let sSource = sCalledMethod
    
    If DebugExtractProject(sSource) = "" Or DebugExtractClass(sSource) = "" Then sSource = ""

    ' return result
    Let DebugExtractSource = sSource
End Function
Private Function DebugExtractProject(sSource As String) As String
    ' prepare result
    Dim sProject As String: Let sProject = ""
    
    Dim iDelimiter As Integer: Let iDelimiter = InStr(1, sSource, ".", vbTextCompare)
    If iDelimiter > 0 Then Let sProject = Mid(sSource, 1, iDelimiter - 1)

    ' return result
    Let DebugExtractProject = sProject
End Function
Private Function DebugExtractClass(sSource As String) As String
    ' prepare result
    Dim sClass As String: Let sClass = ""
    
    Dim iDelimiter As Integer: Let iDelimiter = InStr(1, sSource, ".", vbTextCompare)
    If iDelimiter > 0 Then Let sClass = Mid(sSource, iDelimiter + 1)

    ' return result
    Let DebugExtractClass = sClass
End Function
Private Function DebugExtractMethod(sCalledMethod As String) As String
    ' prepare result
    Dim sMethod As String: Let sMethod = ""
    
    Dim iDelimiter As Integer: Let iDelimiter = InStr(1, sCalledMethod, "::", vbTextCompare)
    If iDelimiter > 0 Then Let sMethod = Mid(sCalledMethod, iDelimiter + 2) Else Let sMethod = sCalledMethod
    
    ' return result
    Let DebugExtractMethod = sMethod
End Function
Private Function NamedItem_GetName(oObject As Object) As String
    ' prepare result
    Dim sName As String: Let sName = ""
    
    If TypeOf oObject Is IJNamedItem Then
        Dim pNamedItem As IJNamedItem: Set pNamedItem = oObject
        On Error Resume Next: Let sName = "[" + pNamedItem.Name + "]": On Error GoTo 0
    Else
        Let sName = "Object does not support IJNamedItem"
    End If
    
    ' return result
    Let NamedItem_GetName = sName
End Function
Function Moniker_GetDisplayName(pMoniker As IMoniker) As String
    ' prepare result
    Dim sDisplayName As String: Let sDisplayName = "Unable to get DisplayName"
    
'''    Dim pBindCtx As IBindCtx: Set pBindCtx = Nothing
'''    Call pMoniker.GetDisplayName(pBindCtx, Nothing, sDisplayName)
    
    ' return result
    Let Moniker_GetDisplayName = sDisplayName
End Function
Private Sub DebugAddItemToListsToDebug(sItem As String)
    Dim iSize As Integer:
    If InStr(1, sItem, ".", vbTextCompare) = 0 Then
        Let iSize = SizeOfArrayOfStrings(g_sListOfProjectsToDebug, 1)
        ReDim Preserve g_sListOfProjectsToDebug(1 To iSize + 1)
        Let g_sListOfProjectsToDebug(iSize + 1) = sItem
    ElseIf InStr(1, sItem, "::", vbTextCompare) = 0 Then
        Let iSize = SizeOfArrayOfStrings(g_sListOfClassesToDebug, 1)
        ReDim Preserve g_sListOfClassesToDebug(1 To iSize + 1)
        Let g_sListOfClassesToDebug(iSize + 1) = sItem
    Else
        Let iSize = SizeOfArrayOfStrings(g_sListOfMethodsToDebug, 1)
        ReDim Preserve g_sListOfMethodsToDebug(1 To iSize + 1)
        Let g_sListOfMethodsToDebug(iSize + 1) = sItem
    End If
End Sub
Private Sub DebugInitListsToDebug()
    ' format of the file
    ' Project
    ' Project.ClassName
    ' Project.ClassName::MethodName
    On Error GoTo wrapup:
    Dim oFileSystemObject As New FileSystemObject
    If oFileSystemObject.FileExists("C:\Temp\DebugSupport.txt") Then
        Dim oTextStream As TextStream
        Set oTextStream = oFileSystemObject.OpenTextFile("C:\Temp\DebugSupport.txt")
        While Not oTextStream.AtEndOfStream
            DebugAddItemToListsToDebug (oTextStream.ReadLine)
        Wend
    End If
wrapup:
    Let g_bIsDebugOn = False
End Sub
Private Function IsProjectToDebug() As Boolean
    Let IsProjectToDebug = False
    
    Dim iSize As Integer: Let iSize = SizeOfArrayOfStrings(g_sListOfProjectsToDebug, 1)
    If iSize > 0 Then
        Dim sProjectToDebug As String: Let sProjectToDebug = DebugExtractProject(g_sDebugSource)
        Dim i As Integer
        For i = 1 To iSize
            If g_sListOfProjectsToDebug(i) = sProjectToDebug Then
                Let IsProjectToDebug = True
                Exit For
            End If
        Next
    End If
End Function
Private Function IsClassToDebug() As Boolean
    Let IsClassToDebug = False
    
    Dim iSize As Integer: Let iSize = SizeOfArrayOfStrings(g_sListOfClassesToDebug, 1)
    If iSize > 0 Then
        Dim sClassToDebug As String: Let sClassToDebug = g_sDebugSource
        Dim i As Integer
        For i = 1 To iSize
            If g_sListOfClassesToDebug(i) = sClassToDebug Then
                Let IsClassToDebug = True
                Exit For
            End If
        Next
    End If
End Function
Private Function IsMethodToDebug(sMethod As String) As Boolean
    Let IsMethodToDebug = False
    
    Dim iSize As Integer: Let iSize = 0
    On Error Resume Next: Let iSize = SizeOfArrayOfStrings(g_sListOfMethodsToDebug, 1): On Error GoTo 0
    
    If iSize > 0 Then
        Dim i As Integer
        For i = 1 To iSize
            If g_sListOfMethodsToDebug(i) = sMethod Then
                Let IsMethodToDebug = True
                Exit For
            End If
        Next
    End If
End Function
Private Function IsAllToDebug() As Boolean
    Let IsAllToDebug = True
    
    If SizeOfArrayOfStrings(g_sListOfProjectsToDebug, 1) > 0 _
    Or SizeOfArrayOfStrings(g_sListOfClassesToDebug, 1) > 0 _
    Or SizeOfArrayOfStrings(g_sListOfMethodsToDebug, 1) > 0 Then Let IsAllToDebug = False
End Function
Private Function DebugGetDebugSupportFromDebugSource(sDebugSource As String) As IDebugSupport
    Dim pDebugSupport As IDebugSupport
    Dim bFound As Boolean: Let bFound = False
    If g_iCountOfDebugSources > 0 Then
        Dim i As Integer
        For i = 1 To g_iCountOfDebugSources
            If sDebugSource = g_sDebugSources(i) Then
                Set pDebugSupport = g_pDebugSupports(i)
                Let bFound = True
                Exit For
            End If
        Next
    End If
    If Not bFound Then
        Let g_iCountOfDebugSources = g_iCountOfDebugSources + 1
        Set g_pDebugSupports(g_iCountOfDebugSources) = New DebugSupport
        Let g_sDebugSources(g_iCountOfDebugSources) = sDebugSource
        
        Set pDebugSupport = g_pDebugSupports(g_iCountOfDebugSources)
    End If
    
    Set DebugGetDebugSupportFromDebugSource = pDebugSupport
End Function
Private Function SizeOfArrayOfStrings(sArray() As String, iDimension As Integer) As Integer
    Dim iSize As Integer: Let iSize = 0
    On Error Resume Next: Let iSize = UBound(sArray, iDimension): On Error GoTo 0
    Let SizeOfArrayOfStrings = iSize
End Function
