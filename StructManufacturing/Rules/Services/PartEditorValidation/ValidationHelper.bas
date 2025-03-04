Attribute VB_Name = "ValidationHelper"
'Key for the Units Of Measure Service
Public Const TKUnitsOfMeasure = "UnitsOfMeasure"
Private Const ERROR_LOG_FILE_NAME = "PartMonitorError.log"

'SMS_SCHEMA Settings XML DOM DOCUMENT files
Public m_oSMS_ViewerDOM As DOMDocument
Public m_oSMS_AnnotationDOM As DOMDocument

Private Declare Function GetTempPath Lib "kernel32" _
            Alias "GetTempPathA" (ByVal nBufferLength As Long, _
            ByVal lpBuffer As String) As Long

'Logging DebugData
Public Sub LogDebugData(sSource As String, Optional sData As String = "", _
Optional sMethodName As String = "", Optional lineNo As Long = 0)
    Dim FileNumber
    Dim sLogFileName    As String
    Dim sDebugInfo      As String
    
    sDebugInfo = " Source File  :   " & sSource & vbCrLf & _
                " Method Name   :   " & sMethodName & vbCrLf & _
                " Line No       :   " & lineNo & vbCrLf & _
                " Data          :   " & sData & vbCrLf & vbCrLf
    
    sLogFileName = GetTempFolder() & ERROR_LOG_FILE_NAME
    
    FileNumber = FreeFile
    
    'Open the file in Append mode
    Open sLogFileName For Append As #FileNumber
    
    'Write the debug data into log file
    Write #FileNumber, vbCrLf & _
        "###################################################################################" & _
        vbCrLf & vbCrLf & _
        "Debug data: " & Format(Date, "Long Date") & "|" & Format(Time, "Long Time") & "." & _
        Right(Format(Timer, "#0.000"), 3) & _
        vbCrLf & vbCrLf & _
        sDebugInfo & vbCrLf & _
        "###################################################################################"
    'Close the file
    Close #FileNumber
End Sub

Public Function GetTempFolder() As String
    'Returns the path to the user's Temp folder. To boot, Windows
    'requires that a temporary folder exist, so this should always
    'safely return a path to one. Just in case, though, check the
    'return value of GetTempPath.
    Dim lngTempPath As Long
    Dim strTempPath As String
    
    'Fill string with null characters.
    strTempPath = String(144, vbNullChar)
    'Get length of string.
    lngTempPath = Len(strTempPath)
    'Call GetTempPath, passing in string length and string.
    If (GetTempPath(lngTempPath, strTempPath) > 0) Then
        'Get TempPath returns path into string.
        'Truncate string at first null character.
        GetTempFolder = Left(strTempPath, _
            InStr(1, strTempPath, vbNullChar) - 1)
    Else
        GetTempFolder = ""
    End If
End Function

''**************************************************************************************
'' Routine      : GetTotalLengthOfObject
'' Description  : Returns the total length of a Rad.Object
''                For a Line we return the .Length property
''                For an Arc we return SweepAngle * Radius
''                For a group we return the sum of the lengths of the Sub elements
''**************************************************************************************
Public Function GetTotalLengthOfObject(oRadObject As Object) As Double
    On Error GoTo ErrorHandler
    Const METHOD = "GetTotalLengthOfObject"
    'StartTimer m_oTimer, strPerfTrack_SupportingFunc, MODULE & "::" & METHOD
    Dim oLine2D             As Line2d
    Dim oArc2DGeom          As ArcGeometry2d
    Dim oArc2d              As Arc2d
    Dim oEllipticalArcGeom  As EllipticalArcGeometry2d
    Dim oEllipticalArc      As EllipticalArc2d
    Dim oBSplineCurve2DGeom As BSplineCurveGeometry2d
    Dim oBSplineCurve2D     As BSplineCurve2d
    Dim oCircle             As Circle2d
    Dim oEllipse            As Ellipse2d
    Dim oGroup              As Group
    Dim oSubObject          As Object
    Dim dTotalLength        As Double
    
    GetTotalLengthOfObject = 0
    If Not oRadObject Is Nothing Then
        If oRadObject.Type = igLine2d Then
            Set oLine2D = oRadObject
            If Not oLine2D Is Nothing Then
                GetTotalLengthOfObject = oLine2D.length
            End If
        ElseIf oRadObject.Type = igArc2d Then
            Set oArc2d = oRadObject
            If Not oArc2d Is Nothing Then
                Set oArc2DGeom = oArc2d.GetGeometry
                If Not oArc2DGeom Is Nothing Then
                    GetTotalLengthOfObject = oArc2DGeom.length
                    GoTo CleanUp
                End If
                'Survival technique, in case of failure above
                GetTotalLengthOfObject = Abs(oArc2d.SweepAngle * oArc2d.Radius)
            End If
        ElseIf oRadObject.Type = igEllipticalArc2d Then
            Set oEllipticalArc = oRadObject
            If Not oEllipticalArc Is Nothing Then
                Set oEllipticalArcGeom = oEllipticalArc.GetGeometry
                If Not oEllipticalArcGeom Is Nothing Then
                    GetTotalLengthOfObject = oEllipticalArcGeom.length
                    GoTo CleanUp
                End If
                'Survival technique, in case of failure above
                'We don't have a survival technique for elliptical arcs!
            End If
        ElseIf oRadObject.Type = igBsplineCurve2d Then
            Set oBSplineCurve2D = oRadObject
            If Not oBSplineCurve2D Is Nothing Then
                Set oBSplineCurve2DGeom = oBSplineCurve2D.GetGeometry
                If Not oBSplineCurve2DGeom Is Nothing Then
                    GetTotalLengthOfObject = oBSplineCurve2DGeom.length
                End If
                'Survival technique, in case of failure above
                'We don't have a survival technique for BSplines!
            End If
        ElseIf oRadObject.Type = igCircle2d Then
            Set oCircle = oRadObject
            If Not oCircle Is Nothing Then
                GetTotalLengthOfObject = oCircle.Circumference
            End If
        ElseIf oRadObject.Type = igEllipse2d Then
            Set oEllipse = oRadObject
            If Not oEllipse Is Nothing Then
                GetTotalLengthOfObject = oEllipse.Circumference
            End If
        ElseIf oRadObject.Type = igGroup Then
            Set oGroup = oRadObject
            If Not oGroup Is Nothing Then
                For Each oSubObject In oGroup
                    If Not oSubObject Is Nothing Then
                        GetTotalLengthOfObject = GetTotalLengthOfObject + GetTotalLengthOfObject(oSubObject)
                    End If
                Next oSubObject
            End If
        Else
            'This is an error condition, not sure what kind of object it is!
        End If
    End If
CleanUp:
    Set oLine2D = Nothing
    Set oArc2d = Nothing
    Set oBSplineCurve2D = Nothing
    Set oGroup = Nothing
    Set oSubObject = Nothing
    'StopTimer m_oTimer, strPerfTrack_SupportingFunc, MODULE & "::" & METHOD
    Exit Function
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Function

''**************************************************************************************
'' Routine      : ReportRADError
'' Abstract     : Reports an error against a RAD object.
''**************************************************************************************
Public Sub ReportRADError(oObj As Object, oErrorCollection As Collection, oErrColObjs As Collection, _
sError As String, sSolution As String, sValidType As String, ErrorLevel As ErrorLevel, sErrorType As String)
    Const METHOD = "ReportRADError"
    On Error GoTo ErrorHandler
    'Need to check if the given object already has an error reported against it.
    'If there is an error reported against an object there isn't any point in reporting the same object
    'again in the event that there is another problem with it.
    'If there is another problem with the contour and it's not fixed it will get caught and reported
    'the next time the validation is run, so no worries.
    If IsObjInCol(oObj, oErrColObjs) = False Then
        Dim tempErrorObj As STRMFGPartEditorError.CPartEditErrorObj
        Set tempErrorObj = New STRMFGPartEditorError.CPartEditErrorObj
        If tempErrorObj.SetData(oObj, sValidType, sError, sSolution, _
        ErrorLevel, sErrorType, "") Then
            oErrorCollection.Add tempErrorObj
            oErrColObjs.Add oObj
        End If
        Set tempErrorObj = Nothing
    End If
CleanUp:
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Sub

''**************************************************************************************
'' Routine      : IsObjInCol
'' Abstract     : Checks to see if a given object is in a given collection.
''**************************************************************************************
Public Function IsObjInCol(oObj As Object, oErrColObjs As Collection) As Boolean
    Const METHOD = "IsObjInCol"
    On Error GoTo ErrorHandler
    
    Dim oCompObj As Object
    IsObjInCol = False
    If oObj Is Nothing Or oErrColObjs Is Nothing Then GoTo CleanUp
    For Each oCompObj In oErrColObjs
        If Not oCompObj Is Nothing Then
            If oCompObj = oObj Then
                IsObjInCol = True 'Found the object, get out!
                GoTo CleanUp
            End If
        End If
    Next oCompObj
CleanUp:
    Set oCompObj = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Function

''**************************************************************************************
'' Routine      : FlattenGroups
'' Abstract     : Flattens all the objects in groups into a collection.
''**************************************************************************************
Public Sub FlattenGroups(oCollection As Collection, oGroup As Group, RetCollection As Collection, oRadDoc As RAD2D.Document)
    Const METHOD = "FlattenGroups"
    On Error GoTo ErrorHandler
    'Make the outer contour collection a group by removing all geometry primitives
    'from their current groups and placing them in one master group
    Set oGroup = MakeGroup(oCollection, oRadDoc.ActiveSheet, RetCollection)
    Set oCollection = Nothing
CleanUp:
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Sub

''**************************************************************************************
'' Routine      : UnFlattenGroups
'' Abstract     : UnFlattens all the objects in groups and returns them to their original groups.
''**************************************************************************************
Public Sub UnFlattenGroups(oCollection As Collection, oGroup As Group, oRadDoc As RAD2D.Document)
    Const METHOD = "UnFlattenGroups"
    On Error GoTo ErrorHandler
    If Not oGroup Is Nothing Then
        'Place the primitives back in their original groups
        ReAssignParentGroupToRadObjects oRadDoc.ActiveSheet, oCollection
        Set oCollection = Nothing
        oGroup.Delete
    End If
CleanUp:
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Sub

''**************************************************************************************
'' Routine      : MakeGroup
'' Abstract     : Create a group from the Rad object collection
''**************************************************************************************
Public Function MakeGroup(oObjectCollection As Collection, oSheet As Sheet, oGroupedObjCollection As Collection) _
As Group
    Const METHOD = "MakeGroup"
    On Error GoTo ErrorHandler
    Dim oSelectSet As RAD2D.SelectSet
    Dim oObject As Object
    Dim oTempArry() As Object
    Dim i As Integer
    
    If oGroupedObjCollection Is Nothing Then
        Set oGroupedObjCollection = New Collection
    End If
    
    Set oSelectSet = oSheet.Parent.SelectSet
    
    For Each oObject In oObjectCollection
        If oObject.Type = igGroup Then
            RetrieveObjectsFromGroup oObject, oGroupedObjCollection
        Else
            oGroupedObjCollection.Add oObject, oObject.Name
        End If
    Next oObject
    If oGroupedObjCollection.Count >= 1 Then
        ReDim oTempArry(1 To oGroupedObjCollection.Count)
    Else
        ReDim oTempArry(1 To 1)
    End If
    oSelectSet.removeAll
    
    If oGroupedObjCollection.Count <> 0 Then
        For i = 1 To oGroupedObjCollection.Count
            Set oTempArry(i) = oGroupedObjCollection.Item(i)
        Next i
        If oSelectSet.Count <= 0 Then GoTo CleanUp
        Set MakeGroup = oSheet.Groups.AddByObjects(UBound(oTempArry), oTempArry)
    End If
CleanUp:
    oSelectSet.removeAll
    Set oSelectSet = Nothing
    Set oObject = Nothing
    Erase oTempArry
    Exit Function
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Function

''**************************************************************************************
'' Routine      : RetrieveObjectsFromGroup
'' Abstract     : Call recursively to put all the grouped objects into a collection
''**************************************************************************************
Public Sub RetrieveObjectsFromGroup(oGroup As RAD2D.Group, oObjectCollection As Collection)
    Const METHOD = "RetrieveObjectsFromGroup"
    On Error GoTo ErrorHandler
    Dim oAttribute As RAD2D.Attribute
    Dim oAttributeSet As RAD2D.AttributeSet
    Dim oObject As Object
    Dim oParent As Object
    Dim iIndex As Integer
    
    If oObjectCollection Is Nothing Then
        Set oObjectCollection = New Collection
    End If
    For iIndex = 1 To oGroup.Count
        Set oObject = oGroup.Item(iIndex)
        If Not oObject Is Nothing Then
            If oObject.Type = igGroup Then 'Call recursively
                RetrieveObjectsFromGroup oObject, oObjectCollection
            Else
                Set oParent = oObject.Parent
                
                Set oAttributeSet = oObject.AttributeSets.Add("Parent")
                Set oAttribute = oAttributeSet.Add("Name", igAttrTypeString)
                oAttribute.Value = oParent.Name
                
                oObjectCollection.Add oObject
                
                Set oAttributeSet = Nothing
                Set oAttribute = Nothing
            End If
        End If
    Next iIndex
CleanUp:
    Set oAttribute = Nothing
    Set oAttributeSet = Nothing
    Set oObject = Nothing
    Set oParent = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Sub

''**************************************************************************************
'' Routine      : ReAssignParentGroupToRadObjects
'' Abstract     : put objects back into their collections
''**************************************************************************************
Public Sub ReAssignParentGroupToRadObjects(oSheet As Sheet, oObjectCollection As Collection)
    Const METHOD = "ReAssignParentGroupToRadObjects"
    On Error GoTo ErrorHandler
    Dim oAttributeSet       As AttributeSet
    Dim oObject             As Object
    Dim oGroup              As RAD2D.Group
    Dim oGroupCollection    As Collection
    Dim strParentName       As String
    Dim iIndex              As Integer
    Dim oObjectArray()      As Object
    
    If oObjectCollection Is Nothing Then
        GoTo CleanUp
    End If
    
    Set oGroupCollection = New Collection
    
    For Each oGroup In oSheet.Groups
        oGroupCollection.Add oGroup
        If oGroup.HasNested = True Then
            AddSubGroupsInToCollection oGroup, oGroupCollection
        End If
    Next oGroup
    
    For Each oObject In oObjectCollection
        For Each oAttributeSet In oObject.AttributeSets
            If oAttributeSet.SetName = "Parent" Then
                strParentName = GetAttributeValueFromAttrSet(oAttributeSet, "NAME")
                Set oAttributeSet = Nothing
                RemoveAttributeSet oObject, "Parent"
                Exit For
            End If
        Next oAttributeSet
        For Each oGroup In oGroupCollection
            If strParentName = oGroup.Name Then
                ReDim oObjectArray(1 To 1)
                Set oObjectArray(1) = oObject
                oGroup.AddObjectsToGroup 1, oObjectArray
                strParentName = ""
                Exit For
            End If
        Next oGroup
        Erase oObjectArray
    Next oObject
CleanUp:
    Set oAttributeSet = Nothing
    Set oObject = Nothing
    Set oGroup = Nothing
    Erase oObjectArray
    Set oGroupCollection = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Sub

''**************************************************************************************
'' Routine      : AddSubGroupsInToCollection
'' Abstract     :
''**************************************************************************************
Public Sub AddSubGroupsInToCollection(oGroup As Group, oGroupCollection As Collection)
    Const METHOD = "AddSubGroupsInToCollection"
    On Error GoTo ErrorHandler
    Dim oSubGroup As Group
    Dim i As Integer
    
    If oGroup Is Nothing Then GoTo CleanUp
    If oGroupCollection Is Nothing Then
        Set oGroupCollection = New Collection
    End If
    For i = 1 To oGroup.Count
        Set oSubGroup = oGroup.Item(i)
        If Not oSubGroup Is Nothing Then
            oGroupCollection.Add oSubGroup
            If oSubGroup.HasNested = True Then
                AddSubGroupsInToCollection oSubGroup, oGroupCollection
            End If
        End If
    Next i
CleanUp:
    Set oSubGroup = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Sub

'********************************************************************
' Routine: GetOuterContoursCollection
' Abstract: Will gather together all the outer contour objects.
'********************************************************************
Public Function GetOuterContoursCollection(oSheet As RAD2D.Sheet) As Collection
    On Error GoTo ErrorHandler
    Const METHOD As String = "GetOuterContoursCollection"
    
    Dim oAttributeSet       As RAD2D.AttributeSet
    Dim oAttributeSets      As RAD2D.AttributeSets
    Dim oRadObject          As Object
    Dim oCollection         As New Collection
    Dim strAttributeValue   As String
    Dim sTemp               As String
    
    If oSheet Is Nothing Then GoTo CleanUp
    
    'Get the Marking line objects from the sheet Drawing objects
    For Each oRadObject In oSheet.DrawingObjects
        'Add if they are of type either line, arc, bspline, circle, ellipse, elliptical arc, or group,
        'Need to report circles & ellipses!
        If oRadObject.Type = igLine2d Or oRadObject.Type = igArc2d _
        Or oRadObject.Type = igBsplineCurve2d Or oRadObject.Type = igCircle2d _
        Or oRadObject.Type = igEllipse2d Or oRadObject.Type = igGroup _
        Or oRadObject.Type = igEllipticalArc2d Then
            'Get the Attributesets
            Set oAttributeSets = oRadObject.AttributeSets
            'If they have attributesets then only add the object into collection
            If oAttributeSets.Count <> 0 Then
                For Each oAttributeSet In oAttributeSets
                    'if the AttributeSet name is "Contour" then it may be outer or inner contour
                    If InStr(1, oAttributeSet.SetName, "Contour", vbTextCompare) > 0 Then
                        strAttributeValue = GetAttributeValueFromAttrSet(oAttributeSet, "SMS_EDGE" & "||" & "TYPE")
                        If strAttributeValue <> "" And InStr(1, strAttributeValue, "hole", vbTextCompare) <= 0 Then
                            oCollection.Add oRadObject, oRadObject.Name
                        End If
                    End If
                Next oAttributeSet
            End If
        End If
    Next oRadObject
CleanUp:
    Set GetOuterContoursCollection = oCollection
    Set oAttributeSet = Nothing
    Set oAttributeSets = Nothing
    Set oRadObject = Nothing
    Set oCollection = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Function

'********************************************************************
' Routine: GetInnerContoursCollection
' Abstract: Will gather together all the inner contour objects.
'********************************************************************
Public Function GetInnerContoursCollection(oSheet As RAD2D.Sheet) As Collection
    On Error GoTo ErrorHandler
    Const METHOD As String = "GetInnerContoursCollection"
    
    Dim oAttributeSet As RAD2D.AttributeSet
    Dim oAttributeSets As RAD2D.AttributeSets
    Dim oRadObject As Object
    Dim oCollection As New Collection
    Dim strAttributeValue As String
    Dim sTemp As String
    
    If oSheet Is Nothing Then GoTo CleanUp
    
    'Get the Marking line objects from the sheet Drawing objects
    For Each oRadObject In oSheet.DrawingObjects
        'Add if they are of type either line, arc, bspline, circle, ellipse or group!
        If Not oRadObject Is Nothing Then
            If oRadObject.Type = igLine2d Or oRadObject.Type = igArc2d _
            Or oRadObject.Type = igBsplineCurve2d Or oRadObject.Type = igCircle2d _
            Or oRadObject.Type = igEllipse2d Or oRadObject.Type = igGroup _
            Or oRadObject.Type = igEllipticalArc2d Then
                'Get the AttributeSets
                Set oAttributeSets = oRadObject.AttributeSets
                If Not oAttributeSets Is Nothing Then
                    'If they have attribute sets then only add the object into the collection
                    If oAttributeSets.Count <> 0 Then
                        For Each oAttributeSet In oAttributeSets
                            If Not oAttributeSet Is Nothing Then
                                'If the AttributeSet name is "Contour" then it may be outer or inner contour
                                If InStr(1, oAttributeSet.SetName, "Contour", vbTextCompare) > 0 Then
                                    strAttributeValue = GetAttributeValueFromAttrSet(oAttributeSet, "SMS_EDGE" & "||" & "TYPE")
                                    If InStr(1, strAttributeValue, "hole", vbTextCompare) > 0 Then
                                        oCollection.Add oRadObject, oRadObject.Name
                                    End If
                                End If
                            End If
                        Next oAttributeSet
                    End If
                End If
            End If
        End If
    Next oRadObject
CleanUp:
    Set GetInnerContoursCollection = oCollection
    Set oAttributeSet = Nothing
    Set oAttributeSets = Nothing
    Set oRadObject = Nothing
    Set oCollection = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Function

'********************************************************************
' Routine: GetValidationErrorLevel
' Abstract: Gets the error level of a particular type of validation from the SMS_VIEWER.xml config file.
'********************************************************************
Public Function GetValidationErrorLevel(strKey As String) As ErrorLevel
    Const METHOD = "GetValidationErrorLevel"
    On Error GoTo ErrorHandler
    Dim oViewerDOM As DOMDocument
    Dim oValidationNode As IXMLDOMNode
    Dim oValidationAttribs As IXMLDOMNamedNodeMap
    Dim oValidationAttrib As IXMLDOMAttribute
    Dim sValidationErrLvl As String
    Set oViewerDOM = m_oSMS_ViewerDOM
    If Not oViewerDOM Is Nothing Then
        '<VALIDATION_TYPE NAME="Bevel Grind Custom" KEY="ValidateBevelGrindAttribCustom" VALUE="2" LEVEL="1" PROGID=""/>
        '//VALIDATION_TYPES/VALIDATION_TYPE[@KEY='ValidateBevelGrindAttribCustom']
        Set oValidationNode = oViewerDOM.selectSingleNode("//" & "VALIDATION_TYPES" & "/" & "VALIDATION_TYPE" & _
                                                        "[@" & "KEY" & "='" & strKey & "']")
        If Not oValidationNode Is Nothing Then
            Set oValidationAttribs = oValidationNode.attributes
            If Not oValidationAttribs Is Nothing Then
                For Each oValidationAttrib In oValidationAttribs
                    If Not oValidationAttrib Is Nothing Then
                        If oValidationAttrib.nodeName = "LEVEL" Then
                            sValidationErrLvl = oValidationAttrib.nodeValue
                            If IsNumeric(sValidationErrLvl) = True Then
                                GetValidationErrorLevel = CInt(sValidationErrLvl)
                                GoTo CleanUp
                            End If
                        End If
                    End If
                Next oValidationAttrib
            End If
        End If
    End If
CleanUp:
    Set oValidationAttrib = Nothing
    Set oValidationAttribs = Nothing
    Set oValidationNode = Nothing
    Set oViewerDOM = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Function

'********************************************************************
'Routine      :   GetAttribute
'Abstract     :   Get a Attribute value from attribute set of RAD object given Attribute name
'             :     In order to differentiate between the String and Double attributes, the attributes are created
'             :     with igAttrTypeString for strings and igAttrTypeArray types for double values
'             :     (Just to make Units to be same as UOM units in Attribute viewer and
'             :     Drawing Area. This is just a way to identify the attribute value units.
'NOTE         : The below information was in support of the old way of handling attributes,
'             : all attributes are now of type string, this is considered obsolete, we just keep it shown here
'             : because the below unit code is still in this function. We've kept it there because this is
'             : a critically important function, and it is risky to make changes.
' Inputs      :   The UBount of Array attribute is always two.
'                   Array(1) - Unit Type(igUnitDistance,igUnitArea,igUnitAngle,igUnitMass)'
'                   Array(2) - Double Value
'********************************************************************
Public Function GetAttribute(oObject As Object, AttrName As String) As RAD2D.Attribute
    On Error GoTo ErrorHandler
    Const METHOD = "GetAttribute"
    Dim oAttribute      As RAD2D.Attribute
    Dim oAttributeSet   As RAD2D.AttributeSet
    Dim oAttributeSets  As RAD2D.AttributeSets
    
    If oObject Is Nothing Then
        GoTo CleanUp
    End If
    Set oAttributeSets = oObject.AttributeSets
    If Not oAttributeSets Is Nothing Then
        If oAttributeSets.Count > 0 Then
            For Each oAttributeSet In oAttributeSets
                If Not oAttributeSet Is Nothing Then
                    For Each oAttribute In oAttributeSet
                        If Not oAttribute Is Nothing Then
                            If InStr(1, oAttribute.Name, "||") > 0 Then
                                'If GetSystemName(oAttribute.Name) = AttrName Then
                                If oAttribute.Name = AttrName Then
                                    Set GetAttribute = oAttribute
                                    GoTo CleanUp
                                End If
                            ElseIf oAttribute.Name = AttrName Then
                                Set GetAttribute = oAttribute
                                GoTo CleanUp
                            End If
                        End If
                    Next oAttribute
                End If
            Next oAttributeSet
        End If
    End If
CleanUp:
    Set oAttribute = Nothing
    Set oAttributeSet = Nothing
    Set oAttributeSets = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Function

'********************************************************************
'Routine      :   RemoveAttributeSet
'Abstract     :   Removes an entire attribute set given the name of the attribute set to be removed.
'NOTE         : The names must match exactly, this is a case sensitive operation!
'********************************************************************
Public Function RemoveAttributeSet(oObject As Object, strAttrSetName As String)
    On Error GoTo ErrorHandler
    Const METHOD = "RemoveAttributeSet"
    
    Dim oAttributeSets  As RAD2D.AttributeSets
    Dim oAttributeSet   As RAD2D.AttributeSet
    Dim i               As Integer
    If oObject Is Nothing Then
        GoTo CleanUp
    End If
    If oObject.IsAttributeSetPresent(strAttrSetName) Then
        Set oAttributeSets = oObject.AttributeSets
        If Not oAttributeSets Is Nothing Then
            For i = 1 To oAttributeSets.Count 'Each oAttributeSet In oAttributeSets
                Set oAttributeSet = oAttributeSets.Item(i)
                If Not oAttributeSet Is Nothing Then
                    If oAttributeSet.SetName = strAttrSetName Then
                        oAttributeSets.Remove oAttributeSet.SetName
                    End If
                End If
                Set oAttributeSet = Nothing
            Next i
        End If
    End If
CleanUp:
    Set oAttributeSet = Nothing
    Set oAttributeSets = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Function

'********************************************************************
'Routine      :   GetAttributeValueFromAttrSet
'Abstract     :   Given the attribute set and the attribute name, this function returns the attribute value
'                 In order to differentiate between the String and Double attributes, the attributes are created
'                 with igAttrTypeString for strings and igAttrTypeArray types for double values.
'                 (Just to make Units to be same as UOM units in Attribute viewer and
'                 Drawing Area. This is just a way to identify the attribute value units.)
'
'NOTE         : The below information was in support of the old way of handling attributes,
'             : all attributes are now of type string, this is considered obsolete, we just keep it shown here
'             : because the below unit code is still in this function. We've kept it there because this is
'             : a critically important function, and it is risky to make changes.
'                 The UBound of Array attribute is always two.
'                   Array(1) - Unit Type(igUnitDistance,igUnitArea,igUnitAngle,igUnitMass)'
'                   Array(2) - Double Value'
'
'Description:   Used in  AddSymbol() method..basically for Multiple margin case
'********************************************************************
Public Function GetAttributeValueFromAttrSet(oAttrSet As AttributeSet, sAttrName As String) As String
    Const METHOD = "GetAttributeValueFromAttrSet"
    On Error GoTo ErrorHandler
    'StartTimer m_oTimer, strPerfTrack_Attributes, MODULE & "::" & METHOD
    Dim dblDBU      As Double
    Dim i           As Long
    Dim strName2Use As String
    Dim sParentName As String
    Dim sSysName    As String
    
    If Not oAttrSet Is Nothing Then
        For i = 1 To oAttrSet.Count
            strName2Use = oAttrSet.Item(i).Name
            If strName2Use = sAttrName Then
                GetAttributeValueFromAttrSet = oAttrSet.Item(i).Value
                Exit For
            End If
        Next i
    End If
    'StopTimer m_oTimer, strPerfTrack_Attributes, MODULE & "::" & METHOD
CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Function

'********************************************************************
'Routine      :   RadGetAttributeValue
'Abstract     :   Get a Attribute value from attribute, attributeSet or attributeSets or RAD object given Attribute name
'
'NOTE         : The below information was in support of the old way of handling attributes,
'             : all attributes are now of type string, this is considered obsolete, we just keep it shown here
'             : because the below unit code is still in this function. We've kept it there because this is
'             : a critically important function, and it is risky to make changes.
'                 The UBound of Array attribute is always two.
'                   Array(1) - Unit Type(igUnitDistance,igUnitArea,igUnitAngle,igUnitMass)'
'                   Array(2) - Double Value'
'********************************************************************
Public Function RadGetAttributeValue(oObject As Object, AttrName As String, useUnits As Boolean, oDOMDoc As DOMDocument, _
Optional eOutputUnit As IMSICDPInterfacesLib.Units = UNIT_NOT_SET, Optional bUseRnd As Boolean = True, _
Optional bLookUpAnnot As Boolean = False, Optional sParentNodeName As String = "") As String
    On Error GoTo ErrorHandler
    Const METHOD = "RadGetAttributeValue"
    Dim oAttribute      As RAD2D.Attribute
    Dim oAttributeSet   As RAD2D.AttributeSet
    Dim oAttributeSets  As RAD2D.AttributeSets
    Dim oUnitsOfMeasure As IJUomVBInterface
    Dim dblDBU          As Double
    Dim dValue          As Double
    Dim eUnit           As Long
    Dim i               As Integer
    Dim j               As Integer
    Dim iSetCount       As Integer
    Dim iAttrCount      As Integer
    Dim strName         As String
    Dim strCurName      As String
    Dim strParentName   As String
    Dim strSysName      As String
    Dim bFound          As Boolean
    
    bFound = False
    If oObject Is Nothing Then
        GoTo CleanUp
    End If
    
    Set oUnitsOfMeasure = New UnitsOfMeasureServicesLib.UomVBInterface
    
    iSetCount = 0
    iAttrCount = 0
    
    On Error Resume Next
    If TypeOf oObject Is RAD2D.Attribute Then
        iSetCount = 1
        iAttrCount = 1
        Set oAttribute = oObject
        AttrName = oAttribute.Name 'We don't need to find a name, expect the user passed in an emptry string.
    ElseIf TypeOf oObject Is RAD2D.AttributeSet Then
        iSetCount = 1
        Set oAttributeSet = oObject
        iAttrCount = oAttributeSet.Count
    ElseIf TypeOf oObject Is RAD2D.AttributeSets Then
        Set oAttributeSets = oObject
        On Error GoTo CleanUp
        iSetCount = oAttributeSets.Count
        On Error GoTo ErrorHandler
    Else
        On Error Resume Next
        Set oAttributeSets = oObject.AttributeSets
        If oAttributeSets Is Nothing Then GoTo CleanUp
        iSetCount = oAttributeSets.Count
    End If
    On Error GoTo ErrorHandler
    
    If iSetCount > 0 Then
        If AttrName <> "" And AttrName <> "NOTHING" Then
            'There is no way to find an attribute for an emptry string, just convert it to the AttrName.
            'Unless it really is supposed to be NOTHING.
            strName = AttrName
        End If
        For i = 1 To iSetCount
            If Not oAttributeSets Is Nothing Then
                Set oAttributeSet = oAttributeSets.Item(i)
                iAttrCount = oAttributeSet.Count
            End If
            
            If iAttrCount < 1 Then GoTo NextSet
            For j = 1 To iAttrCount
                If Not oAttributeSet Is Nothing Then
                    Set oAttribute = oAttributeSet.Item(j)
                End If
                
                If Not oAttribute Is Nothing Then
                    strCurName = oAttribute.Name 'PARENT_NODE_NAME||SYSTEM_NAME
                    If InStr(1, strCurName, "||") <= 0 Then
                        strSysName = strCurName
                    Else
                        If strCurName <> strName Then
                            'Only do this if it's absolutely necessary! Should be quite rare!
                            strParentName = GetParentName(strCurName)
                            strCurName = GetSystemName(strCurName)
                            If strSysName = "" Then
                                strSysName = strCurName
                            End If
                        ElseIf strCurName = strName Then
                            strSysName = strCurName
                        End If
                    End If
                    If (strSysName = strName And sParentNodeName = "") _
                    Or (strSysName = strName And sParentNodeName = strParentName) Then
                        If eUnit <> 0 And eUnit <> UNIT_NOT_SET Then
                            dValue = oAttribute.Value
                            Select Case eUnit
                                Case DISTANCE_CENTIMETER, _
                                DISTANCE_CHAIN, _
                                DISTANCE_FOOT, _
                                DISTANCE_FURLONG, _
                                DISTANCE_HUNDREDTH, _
                                DISTANCE_INCH, _
                                DISTANCE_KILOMETER, _
                                DISTANCE_LINK, _
                                DISTANCE_METER, _
                                DISTANCE_MILE, _
                                DISTANCE_MILLIMETER, _
                                DISTANCE_NANOMETER, _
                                DISTANCE_POINT, _
                                DISTANCE_POLE, _
                                DISTANCE_ROD, _
                                DISTANCE_TENTH, _
                                DISTANCE_THOUSANDTH, _
                                DISTANCE_YARD, _
                                UNIT_DISTANCE
                                    RadGetAttributeValue = ConvertAttributeValueUnit(dValue, eUnit, UNIT_DISTANCE, _
                                                                                useUnits, eOutputUnit, bUseRnd)
                                
                                Case ANGLE_DEGREE, _
                                ANGLE_MINUTE, _
                                ANGLE_SECOND, _
                                ANGLE_GRADIAN, _
                                ANGLE_RADIAN, _
                                ANGLE_REVOLUTION, _
                                UNIT_ANGLE
                                    RadGetAttributeValue = ConvertAttributeValueUnit(dValue, eUnit, UNIT_ANGLE, _
                                                                                useUnits, eOutputUnit, bUseRnd)
                                
                                Case AREA_ACRE, _
                                AREA_HECTARE, _
                                AREA_SQUARE_ACRE, _
                                AREA_SQUARE_CENTIMETER, _
                                AREA_SQUARE_FOOT, _
                                AREA_SQUARE_INCH, _
                                AREA_SQUARE_KILOMETER, _
                                AREA_SQUARE_METER, _
                                AREA_SQUARE_MILE, _
                                AREA_SQUARE_MILLIMETER, _
                                AREA_SQUARE_YARD, _
                                UNIT_AREA
                                    RadGetAttributeValue = ConvertAttributeValueUnit(dValue, eUnit, UNIT_AREA, _
                                                                                useUnits, eOutputUnit, bUseRnd)
                                    
                                Case MASS_GRAIN, _
                                MASS_GRAM, _
                                MASS_KILOGRAM, _
                                MASS_LONG_TON, _
                                MASS_MEGAGRAM, _
                                MASS_METRIC_TON, _
                                MASS_MILLIGRAM, _
                                MASS_OUNCE, _
                                MASS_POUND_MASS, _
                                MASS_SHORT_TON, _
                                MASS_SLINCH, _
                                MASS_SLUG, _
                                UNIT_MASS
                                    RadGetAttributeValue = ConvertAttributeValueUnit(dValue, eUnit, UNIT_MASS, _
                                                                                useUnits, eOutputUnit, bUseRnd)
                                
                                Case Else
                                    'We aren't sure what the unit type is, can't convert anything.
                                    bFound = True 'Just return the value as best we can!
                                    If Not VarType(oAttribute.Value) = vbNull Then
                                        RadGetAttributeValue = CStr(oAttribute.Value)
                                    Else
                                        'Not sure of type of value either!
                                        RadGetAttributeValue = ""
                                    End If
                            End Select
                        ElseIf (oAttribute.AttributeType = igAttrTypeDouble) Or _
                               (oAttribute.AttributeType = igAttrTypeInteger) Or _
                               (oAttribute.AttributeType = igAttrTypeLong) Or _
                               (oAttribute.AttributeType = igAttrTypeString) Then
                            dblDBU = oAttribute.Value
                            RadGetAttributeValue = CStr(dblDBU)
                            bFound = True
                        ElseIf oAttribute.AttributeType = igAttrTypeString Then
                            RadGetAttributeValue = oAttribute.Value
                            bFound = True
                        Else
                            If Not VarType(oAttribute.Value) = vbNull Then
                                RadGetAttributeValue = CStr(oAttribute.Value)
                            Else
                                'Not sure of type of value
                                RadGetAttributeValue = ""
                            End If
                            bFound = True
                        End If
                        GoTo CleanUp
                    End If
                End If
                strSysName = ""
                strParentName = ""
            Next j
NextSet:
        Next i
    End If
CleanUp:
    ''Below code for debugging purposes
    'If bFound = False Then
    '    If Not strName = "BevelType" And Not strName = "GrindType" And Not strName = "MarginType" _
    '    And Not strName = "BEVEL_TYPE" And Not strName = "GRIND_TYPE" And Not strName = "MARGIN_TYPE" Then
    '        'MsgBox "Didn't find any matching names, was looking for Attribute Name: " & strName
    '    End If
    'End If
    Set oAttribute = Nothing
    Set oAttributeSet = Nothing
    Set oAttributeSets = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Function

Public Function ConvertAttributeValueUnit(dValue As Double, eUnit As Long, eUnitType As UnitTypes, _
bUseUnits As Boolean, Optional eOutputUnit As IMSICDPInterfacesLib.Units = UNIT_NOT_SET, _
Optional bUseRnd As Boolean = True) As String
    Const METHOD = "ConvertAttributeValueUnit"
    On Error GoTo ErrorHandler
    Dim oUnitsOfMeasure As IJUomVBInterface
    Dim dblDBU          As Double
    
    Set oUnitsOfMeasure = New UnitsOfMeasureServicesLib.UomVBInterface
    'Convert the value into Database units
    dblDBU = oUnitsOfMeasure.ConvertUnitToDbu(eUnitType, dValue, eUnit)
    'Format the string with Units specified in UnitsOfMeasure
    'ie if distance=ft
    'the string will be 5 ft
    If bUseUnits = False Then
        'Just output the value, no units!
        If Not eOutputUnit = UNIT_NOT_SET And eOutputUnit > 0 Then
            dblDBU = oUnitsOfMeasure.ConvertUnitToUnit(eUnitType, dValue, eUnit, eOutputUnit)
        End If
        If dUseRnd = True Then
            ConvertAttributeValueUnit = Round2ThirdDecPlace(dblDBU)
        Else
            ConvertAttributeValueUnit = CStr(dblDBU)
        End If
    Else
        oUnitsOfMeasure.FormatUnit UNIT_DISTANCE, dblDBU, ConvertAttributeValueUnit
    End If
CleanUp:
    Set oUnitsOfMeasure = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Function

'********************************************************************
'Routine      :   GetParentName
'Abstract     :   Given the full system name, get the parent node name
'********************************************************************
Public Function GetParentName(sNameIn As String) As String
    Const METHOD = "GetParentName"
    On Error GoTo ErrorHandler
    Dim iPos As Integer
    GetParentName = sNameIn 'Return what was passed in by default.
    iPos = InStr(1, sNameIn, "||")
    If iPos > 0 Then
        GetParentName = Mid(sNameIn, 1, iPos - 1)
    End If
CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Function

'********************************************************************
'Routine      :   GetSystemName
'Abstract     :   Given the full property name, get the system name, also called the system name of the property
'********************************************************************
Public Function GetSystemName(sNameIn As String) As String
    Const METHOD = "GetSystemName"
    On Error GoTo ErrorHandler
    Dim iPos As Integer
    GetSystemName = sNameIn 'Return what was passed in by default.
    iPos = InStr(1, sNameIn, "||")
    If iPos > 0 Then
        GetSystemName = Mid(sNameIn, iPos + 2, Len(sNameIn) - iPos)
    End If
CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Function

'********************************************************************
'Routine      :   Round2ThirdDecPlace
'Abstract     :   Rounds to 3 significant figures,
' ensures that we always return 3 digits past the decimal point.
'********************************************************************
Public Function Round2ThirdDecPlace(dblInput As Double) As String
    Const METHOD = "Round2ThirdDecPlace"
    On Error GoTo ErrorHandler
    'We always need to round to the 3rd significant digit.
    Round2ThirdDecPlace = Round2NDecPlace(dblInput, 3)
CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Function

'********************************************************************
'Routine      :   Round2SecondDecPlace
'Abstract     :   Rounds to 2 significant figures,
' ensures that we always return 2 digits past the decimal point.
'********************************************************************
Public Function Round2SecondDecPlace(dblInput As Double) As String
    Const METHOD = "Round2SecondDecPlace"
    On Error GoTo ErrorHandler
    'We always need to round to the 2nd significant digit.
    Round2SecondDecPlace = Round2NDecPlace(dblInput, 2)
CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Function

'********************************************************************
'Routine      :   Round2NDecPlaceAsString
'Abstract     :   Takes a string that must be a number and rounds it to the specified number of significant figures
'********************************************************************
Public Function Round2NDecPlaceAsString(sInput As String, iSigFigDesired As Integer) As String
    Const METHOD = "Round2NDecPlaceAsString"
    On Error GoTo ErrorHandler
    
    Round2NDecPlaceAsString = sInput 'Return what was passed in, just in case there is an error!
    If sInput <> "" And iSigFigDesired > 0 Then
        If IsNumeric(sInput) = True Then
            Round2NDecPlaceAsString = Round2NDecPlace(CDbl(sInput), iSigFigDesired)
        End If
    End If
CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Function

'********************************************************************
'Routine      :   Truncate2NDecPlace
'Abstract     :   Truncates to n-th significant figures,
' ensures that we always return n-number of digits past the decimal point. No more, no less
'********************************************************************
Public Function Truncate2NDecPlace(dblInput As Double, iSigFigDesired As Integer) As String
    Const METHOD = "Truncate2NDecPlace"
    On Error GoTo ErrorHandler
    Dim iSigFig     As Integer
    Dim decPos      As Integer
    Dim iSigFigDiff As Integer
    Dim i           As Integer
    Dim sSigFig     As String
    'We always need to truncate to the n-th significant digit.
    If dblInput = 0 Then
        Truncate2NDecPlace = AddNZeros2String("0.0", "", iSigFigDesired)
    ElseIf InStr(1, CStr(dblInput), ".") <> 0 Then
        decPos = InStr(1, CStr(dblInput), ".")
        dSigFig = Len(CStr(dblInput)) - decPos
        sSigFig = Right(CStr(dblInput), dSigFig)
        iSigFigDiff = iSigFigDesired - dSigFig
        If iSigFigDiff > 0 Then
            'We need to add some zeros
            Truncate2NDecPlace = AddNZeros2String(CStr(dblInput), "0", iSigFigDiff)
        ElseIf iSigFigDiff = 0 Then
            'Perfect!
            Truncate2NDecPlace = CStr(dblInput)
        ElseIf iSigFigDiff < 0 Then
            'Truncate some numbers off
            Truncate2NDecPlace = Left(CStr(dblInput), decPos + iSigFigDesired) 'Example: x.xxx
        Else
            'Not really sure what to do in this case! Definately an error condition!
            Err.Raise 1, METHOD, "Error Truncating value, invalid significant digits desired input"
        End If
    ElseIf InStr(1, CStr(dblInput), ".") = 0 Then
        Truncate2NDecPlace = AddNZeros2String(CStr(dblInput), ".0", iSigFigDesired)
    End If
CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Function

'********************************************************************
'Routine      :   Round2NDecPlace
'Abstract     :   Rounds to n-th significant figures,
' ensures that we always return n-number of digits past the decimal point.
'********************************************************************
Private Function Round2NDecPlace(dblInput As Double, iSigFigDesired As Integer) As String
    Const METHOD = "Round2NthDecPlace"
    On Error GoTo ErrorHandler
    Dim iSigFig     As Integer
    Dim decPos      As Integer
    Dim iSigFigDiff As Integer
    Dim i           As Integer
    Dim sSigFig     As String
    'We always need to round to the n-th significant digit.
    If iSigFigDesired >= 1 Then
        dblInput = Round(dblInput, iSigFigDesired)
    End If
    If dblInput = 0 Then
        Round2NDecPlace = AddNZeros2String("0.0", "", iSigFigDesired)
    ElseIf InStr(1, CStr(dblInput), ".") <> 0 Then
        decPos = InStr(1, CStr(dblInput), ".")
        dSigFig = Len(CStr(dblInput)) - decPos
        sSigFig = Right(CStr(dblInput), dSigFig)
        iSigFigDiff = iSigFigDesired - dSigFig
        If iSigFigDiff > 0 Then
            'We need to add some zeros
            Round2NDecPlace = AddNZeros2String(CStr(dblInput), "0", iSigFigDiff)
        ElseIf iSigFigDiff = 0 Then
            'Perfect
            Round2NDecPlace = CStr(dblInput)
        ElseIf iSigFigDiff < 0 Then
            'Truncate some numbers off
            Round2NDecPlace = Left(CStr(dblInput), decPos + iSigFigDesired) 'Example: x.xxx
        Else
            'Not really sure what to do in this case! Definately an error condition!
            Err.Raise 1, METHOD, "Error rounding: Invalid desired significant figures."
        End If
    ElseIf InStr(1, CStr(dblInput), ".") = 0 Then
        Round2NDecPlace = AddNZeros2String(CStr(dblInput), ".0", iSigFigDesired)
    End If
CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Function

'********************************************************************
'Routine      :   AddNZeros2String
'Abstract     :   Adds "N" number of Zeros to the end of a given string, First time through the loop another string
' can be appended to the string, before adding other zeros.
' If there is an invalid input for iNZeros2Add, then we just return the input string, with no changes.
'********************************************************************
Public Function AddNZeros2String(sIn As String, sFirstStr As String, iNZeros2Add As Integer) As String
    Const METHOD = "AddNZeros2String"
    On Error GoTo ErrorHandler
    Dim i As Integer
    If iNZeros2Add >= 1 Then
        For i = 1 To iNZeros2Add
            If i = 1 Then
                AddNZeros2String = sIn & sFirstStr
            Else
                AddNZeros2String = AddNZeros2String & "0"
            End If
        Next i
    Else
        'There was an invalid input, just return the string the same as it was passed in.
        AddNZeros2String = sIn
    End If
CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Function


