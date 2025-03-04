Attribute VB_Name = "BktUtilities"
Option Explicit

Public Const BrBPToDoMsgCodelist = "IntellishipMsgs"
Public Const TDL_BRACKET_BY_PLANE_INVALID_SUPPORTS = 4
                            '"Corner gusset connection has invalid inputs. Delete and replace."
Public Const E_FAIL = -2147467259

Public Const dDebug As Boolean = False

Public Const INPUT_PLANE As String = "Plane"
Public Const INPUT_BRACKETPLANE As String = "Bracket Plane"
Public Const INPUT_BRACKETPLATE As String = "Bracket Plate System"
Public Const INPUT_SUPPORTS As String = "Supports"
Public Const INPUT_SUPPORT1 As String = "Support1"
Public Const INPUT_SUPPORT2 As String = "Support2"
Public Const INPUT_SUPPORT3 As String = "Support3"
Public Const INPUT_SUPPORT4 As String = "Support4"
Public Const INPUT_SUPPORT5 As String = "Support5"
Public Const INPUT_DIRECTION As String = "Direction"

Public Const QUESTION_BracketSupport1Type As String = "Support1Type"
Public Const CODELIST_BracketSupport1or2Type As String = "BracketSupport1or2Type"
Public Const DEFAULT_BracketSupport1or2Type_TRIMMED As String = "Trimmed"
Public Const QUESTION_BracketSupport2Type As String = "Support2Type"
Public Const QUESTION_BracketSupport3Type As String = "Support3Type"
Public Const QUESTION_BracketSupport4Type As String = "Support4Type"
Public Const QUESTION_BracketSupport5Type As String = "Support5Type"
Public Const CODELIST_BracketSupport3or4or5Type As String = "BracketSupport3or4or5Type"
Public Const DEFAULT_BracketSupport3or4or5Type_TRIMMED As String = "Trimmed"
Public Const CODELIST_BracketSupport4or5Type As String = "BracketSupport4or5Type"
Public Const DEFAULT_BracketSupport4or5Type_TRIMMED As String = "Trimmed"
Public Const QUESTION_BracketContourType As String = "BracketContourType"
Public Const CODELIST_BracketContourType As String = "BracketContourType"
Public Const DEFAULT_BracketContourType_LINEAR As String = "Linear"
Public Const CMLIBRARY_BRACKETSEL As String = "BktUnifiedRules.BracketSelCM"



Private Const MODULE = "M:\SharedContent\Src\StructDetail\SmartOccurrence\BktUnifiedRules\BktUtilities.bas"

'*************************************************************************
'Function
'ToDoErrorNotify
'
'Abstract
' Called to notify the SmartOccurrence of a ToDo error that occurred during a
' smart occurrence custom evaluate
'
'***************************************************************************
Public Sub ToDoErrorNotify(strCodelistTable As String, nToDoListErrorNum As Long, oObjectInError As Object, oObjectToUpdate As Object)
    Const METHOD = "ToDoErrorNotify"
    On Error GoTo ErrHandler:
    Dim oToDoListHelper As IJToDoListHelper

    Set oToDoListHelper = oObjectInError ' Set ToDoListHelper = pointer to the CAO Object
    If Not oToDoListHelper Is Nothing Then
        If oObjectToUpdate Is Nothing Then
          oToDoListHelper.SetErrorInfo strCodelistTable, nToDoListErrorNum
        Else
          oToDoListHelper.SetErrorInfo strCodelistTable, nToDoListErrorNum, oObjectToUpdate
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    ' Report the error but do not return the error since this call maybe in their
    '   error handler
    Err.Raise LogError(Err, MODULE, METHOD).Number
    Err.Clear
End Sub
'*************************************************************************
'Function
'GetIntersection
'
'Abstract
' Gets the intersection object between the two input objects
'
'***************************************************************************
Public Function GetIntersection(pObject1 As Object, pObject2 As Object) As Object
    Const METHOD = "GetIntersection"
    On Error GoTo ErrorHandler
 
    Dim oStructIntersector As IMSModelGeomOps.DGeomOpsIntersect
    Dim oIntersectedUnknown As IUnknown
    Dim pAgtorUnk As IUnknown
    Dim NullObject As Object
    
    On Error Resume Next
    Set oStructIntersector = New IMSModelGeomOps.DGeomOpsIntersect
    oStructIntersector.PlaceIntersectionObject NullObject, pObject1, pObject2, pAgtorUnk, oIntersectedUnknown
    If Err.Number <> 0 Then
        Err.Clear
    End If
    On Error GoTo ErrorHandler
   
    Set GetIntersection = oIntersectedUnknown
    
    Set oStructIntersector = Nothing
    Set oIntersectedUnknown = Nothing
    Set pAgtorUnk = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
    Err.Clear
End Function


'*************************************************************************
'Function
'IsFlangeIn
'
'Purpose:
'   Determines is the flange for a giving support is in the direction of
'   the bracket.
'Returns:
'   True   - if Flange is IN (orientation is towards the direction of the bracket.)
'   False - if Flange is OUT (orientation is away from the direction of the bracket.)
'*********************************************************************************************
Public Function IsFlangeTowardsBracket(oProfileSupport As Object, oBracketByPlaneSO As IJSmartOccurrence, uBracketPoint As IJPoint, vBracketPoint As IJPoint) As Boolean

'Assume true...
IsFlangeTowardsBracket = False
                 
'*********************************************************************************************
 '1. Get interstion point between supports and bracket plate
 Dim oBracketUtils As GSCADCreateModifyUtilities.IJBracketAttributes
Set oBracketUtils = New GSCADCreateModifyUtilities.PlateUtils
Dim oBracketPlate As IJPlate
Set oBracketPlate = oBracketUtils.GetBracketByPlaneFromBracketContour(oBracketByPlaneSO)
Dim oIntersectUtil As IIntersect
Set oIntersectUtil = New Intersect
Dim oCommonBody As IUnknown
Dim oStructGeometry As Object
'
'Debug.Print "TypeOf oBracketPlate Is IJDModelBody = " & TypeOf oBracketPlate Is IJDModelBody
'Debug.Print "TypeOf oBracketPlate Is IJPlate = " & TypeOf oBracketPlate Is IJPlate
'Debug.Print "TypeOf oBracketPlate Is IJDGeometry = " & TypeOf oBracketPlate Is IJDGeometry

'During recursion due, need to check if the object is a struct geometry type...

On Error Resume Next
If TypeOf oBracketPlate Is IJDModelBody Then
    ' Just query
    Set oStructGeometry = oBracketPlate
    oIntersectUtil.GetCommonGeometry oProfileSupport, oStructGeometry, oCommonBody, 0
End If
If oCommonBody Is Nothing Then
    Set oStructGeometry = oBracketUtils.GetLimitedBracketPlane(oBracketPlate)
    oIntersectUtil.GetCommonGeometry oProfileSupport, oStructGeometry, oCommonBody, 0
End If
On Error GoTo ErrorHandler:
If dDebug = True Then
'Debug.Print ""
    Dim oModelBody As IJDModelBody
    Set oModelBody = oProfileSupport
    If Not oModelBody Is Nothing Then
        oModelBody.DebugToSATFile "C:\\temp\\Profile.sat"
    End If
    Set oModelBody = Nothing
    Set oModelBody = oStructGeometry
    If Not oModelBody Is Nothing Then
        oModelBody.DebugToSATFile "C:\\temp\\oStructGeometry.sat"
    End If
    Set oModelBody = Nothing
    Set oModelBody = oCommonBody
    If Not oModelBody Is Nothing Then
        oModelBody.DebugToSATFile "C:\\temp\\oCommonBody.sat"
    End If
End If
'oIntersectUtil.GetCommonGeometry oProfileSupport, oStructGeometry, oCommonBody, 0

Dim oFlangeOrientVec As IJDVector
Dim oWebOrientVec As IJDVector
Dim oPointsGraphBody As IJPointsGraphBody
Dim oPntGraphUtils As New SGOPointsGraphUtilities
If TypeOf oCommonBody Is IJPointsGraphBody Then
    Set oPointsGraphBody = oCommonBody
    Dim oPointColl As Collection
    Set oPointColl = oPntGraphUtils.GetPositionsFromPointsGraph(oPointsGraphBody)
    Dim oOrientationPosition As IJDPosition
    Set oOrientationPosition = oPointColl.Item(1)
                     
    Set oBracketUtils = Nothing
    
    '*********************************************************************************************
    '2. Get the direction vector for first and seondary orientation of profile sent...
    Dim oProfilehelper As IJProfileAttributes
    Set oProfilehelper = New ProfileUtils
    oProfilehelper.GetProfileOrientation oProfileSupport, oOrientationPosition, oFlangeOrientVec, oWebOrientVec
     
    '*********************************************************************************************
    '3. Found out what support was sent and choice either the U or V vector from the bracket...
    Dim oBracketDir As IJDVector
    Dim oS1 As Object
    Dim oS2 As Object
    Dim oS3 As Object
    Dim oS4 As Object
    Dim oS5 As Object
    Dim oSupports As IJElements
    Dim nNumSupports As Long
    Set oBracketUtils = New GSCADCreateModifyUtilities.PlateUtils
    
    oBracketUtils.GetSupportsFromBracketContourSO oBracketByPlaneSO, oSupports, nNumSupports
    
    Dim i As Long
    For i = 1 To oSupports.Count
        Select Case i
            Case 1
                Set oS1 = oSupports.Item(i)
            Case 2
                Set oS2 = oSupports.Item(i)
            Case 3
                Set oS3 = oSupports.Item(i)
            Case 4
                Set oS4 = oSupports.Item(i)
            Case 5
                Set oS5 = oSupports.Item(i)
        End Select
    Next i
    
    Dim xDirPoint As Double
    Dim yDirPoint As Double
    Dim zDirPoint As Double
    
    If oProfileSupport Is oS1 Then
        ' Bracket is attached to flange, therefore no need to determine flange direction...
        Exit Function
    ElseIf (oProfileSupport Is oS2) Or (oProfileSupport Is oS3) Then
        
        uBracketPoint.GetPoint xDirPoint, yDirPoint, zDirPoint
        Set oBracketDir = New DVector
        
        oBracketDir.x = xDirPoint
        oBracketDir.y = yDirPoint
        oBracketDir.z = zDirPoint
    
    ElseIf (oProfileSupport Is oS4) Or (oProfileSupport Is oS5) Then
    
        vBracketPoint.GetPoint xDirPoint, yDirPoint, zDirPoint
        Set oBracketDir = New DVector
        
        oBracketDir.x = xDirPoint
        oBracketDir.y = yDirPoint
        oBracketDir.z = zDirPoint
        
    End If
    
    '*********************************************************************************************
    '4. Determine if the direction between the flange direction is in the same direction as the
    '    the bracket...
    Dim fDot As Double
    
    fDot = oBracketDir.x * oFlangeOrientVec.x + _
               oBracketDir.y * oFlangeOrientVec.y + _
               oBracketDir.z * oFlangeOrientVec.z
    
    If fDot < 0 Then
        IsFlangeTowardsBracket = True
    Else
        IsFlangeTowardsBracket = False
    End If
End If
Exit Function
'Clean up...
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "IsFlangeTowardsBracket").Number
    Err.Clear
End Function


'********************************************************************
' ' Routine: LogError
'
' Description:  default Error logger
'********************************************************************
Public Function LogError(oErrObject As ErrObject, _
                            Optional strSourceFile As String = "", _
                            Optional strMethod As String = "", _
                            Optional strExtraInfo As String = "") As IJError
     
    Dim strErrSource As String
    Dim strErrDesc As String
    Dim lErrNumber As Long
    Dim oEditErrors As IJEditErrors
     
    lErrNumber = oErrObject.Number
    strErrSource = oErrObject.Source
    strErrDesc = oErrObject.Description
     
     ' retrieve the error service
    Set oEditErrors = GetJContext().GetService("Errors")
       
    ' add the error to the service : the error is also logged to the file specified by
    ' "HKEY_LOCAL_MACHINE/SOFTWARE/Intergraph/Sp3D/Core/OperationParameter/ReportErrors_Log"
    Set LogError = oEditErrors.Add(lErrNumber, _
                                      strErrSource, _
                                      strErrDesc, _
                                      , _
                                      , _
                                      , _
                                      strMethod & ": " & strExtraInfo, _
                                      , _
                                      strSourceFile)
    Set oEditErrors = Nothing
End Function

Public Sub Add2SBrackets(ByRef pSL As IJDSelectorLogic, ByRef sBracketName As String)
'*************************************************************************
'Sub:
'Add2SL_TrimmedBrackets
'
'Purpose:
'   This Sub will append the defaults sizes as noted in the catalog and add all items
'   to the selection logic input.
'
'Inputs:
'   IJSelectionLogic - Selector to add items to.
'   sBracketName     - Brackets name... Needs to match the
'Returns:
'   Will add all the predifined brackets to the selector
'*********************************************************************************************
On Error GoTo ErrHandler:

    pSL.Add sBracketName & " HxWxN1xN2"
    pSL.Add sBracketName & " 400x400x50x50"
    pSL.Add sBracketName & " 600x600x75x75"
    pSL.Add sBracketName & " 800x800x75x75"
    pSL.Add sBracketName & " 1000x1000x100x100"
    pSL.Add sBracketName & " 2000x2000x150x150"

ErrHandler:
    Err.Raise LogError(Err, MODULE, "Add2SBrackets").Number
    Err.Clear
End Sub
