Attribute VB_Name = "CommonMethods"

Public Sub WriteToLog(strErr As String, m_oStream As TextStream)
    If Not m_oStream Is Nothing Then
        m_oStream.WriteLine strErr
    End If
End Sub

Public Sub CloseLogFile(m_oStream As TextStream)
    Dim strMsg As String
    Dim eStyle As VbMsgBoxStyle
    If Not m_oStream Is Nothing Then
        WriteToLog vbNewLine & "Ending report: " & Now, m_oStream

        strMsg = "Finished reporting Candidates Rule Information" & vbNewLine
        eStyle = vbInformation

        strMsg = strMsg & _
                "Please check the logfile for more information:" & vbNewLine & vbTab & m_strLogFile

        m_oStream.Close
        Set m_oStream = Nothing
    End If
End Sub

Public Sub OpenLogFile(m_strLogFile As String, m_oStream As TextStream)
    Dim strTempPath As String
    Dim oFSO As FileSystemObject

    strTempPath = Environ("TMP")
    If strTempPath = "" Then
        strTempPath = Environ("TEMP")
        If strTempPath = "" Then
            strTempPath = Environ("SystemRoot")
        End If
    End If

    Set oFSO = New FileSystemObject
    m_strLogFile = strTempPath & "\CompareRule.log"
    Set m_oStream = oFSO.OpenTextFile(m_strLogFile, ForWriting, True)
    Set oFSO = Nothing

    If Not m_oStream Is Nothing Then
        WriteToLog "Starting report at: " & Now & vbNewLine, m_oStream
    End If
End Sub

Public Function GetRelatedObjects _
                 (oGivenObject, strGivenObjIfaceGUID, strInterestedRoleName) As IJElements
Const METHOD = "GetProductionRouting"
On Error GoTo ErrorHandler

    Dim oAssocRel               As IJDAssocRelation
    Dim oTargetObjCol           As IJDTargetObjectCol
    Dim oResultColl             As IJElements
    Dim oProductionRouting      As IJDProductionRouting
    Dim oRoutingActionColl      As IJElements
    
    Set oResultColl = New JObjectCollection
    
    Set oAssocRel = oGivenObject
    Set oTargetObjCol = oAssocRel.CollectionRelations(strGivenObjIfaceGUID, strInterestedRoleName)
    
    If oTargetObjCol.Count > 0 Then
        Dim i As Long
        
        For i = 1 To oTargetObjCol.Count
            oResultColl.Add oTargetObjCol.Item(i)
        Next
    End If
    
    Set GetRelatedObjects = oResultColl

Exit Function
ErrorHandler:
    Set m_oError = m_oErrors.AddFromErr(Err, sSOURCEFILE & " - " & METHOD)
    m_oError.Raise
End Function


Public Function IsBracket(oPlateObj As Object) As Boolean
Const METHOD = "IsBracket"
On Error GoTo ErrorHandler

    If TypeOf oPlateObj Is IJSmartPlate Then
        IsBracket = True
        GoTo CleanUp
    End If
    
    Dim oBracketAttr As IJPlateAttributes
    Set oBracketAttr = New PlateUtils
    
    On Error Resume Next
    IsBracket = oBracketAttr.IsBracketByPlane(oPlateObj)
    
    If Not IsBracket Then
        IsBracket = oBracketAttr.IsTrippingBracket(oPlateObj)
    End If
    
    
CleanUp:
    Set oBracketAttr = Nothing
    
Exit Function
                 
ErrorHandler:
    GoTo CleanUp
End Function

