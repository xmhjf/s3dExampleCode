VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTextMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IJMfgTextMarkService

Const SMALL_PART_TEXT_SIZE = 25 ' mm
Const LARGE_PART_TEXT_SIZE = 40 ' mm
Const SMALL_PART_MAX_SIZE = 2#  ' m^2

Const MARKING_CHAR As String = "%"
Const ANTI_MARKING_CHAR As String = "$"

Const INGRPATH = "SOFTWARE\Intergraph\"
Const SYMBOLPATH = "SP3D\Common\TaskInfo\StructManufacturing"

Private Declare Function GetTempPath Lib "kernel32" _
         Alias "GetTempPathA" (ByVal nBufferLength As Long, _
         ByVal lpBuffer As String) As Long

Private m_oSettingsXML As DOMDocument
Private m_bDoesXMLExist As Boolean

Private Const PI As Double = 3.14159265358979
Private Const DECLIVITY_MARK_LENGTH As Double = 0.125
Private Const LABEL_OFFSET_DIST As Double = 0.3
Private Const LABEL_LENGTH_COND As Double = 3
Private Const MAX_TEXT_SIZE_PLATE As Double = 30            'Dimension in mm
Private Const MIN_TEXT_SIZE_PLATE As Double = 20            'Dimension in mm
Private Const SECONDARY_TEXT_SIZE_PLATE As Double = 70      'Dimension in mm
Private Const TEXT_SIZE_PROFILE As Double = 21              'Dimension in mm
Private Const TEXT_SIZE_COLLAR As Double = 21               'Dimension in mm
Private Const TEXT_SIZE_BRACKET As Double = 21              'Dimension in mm
Private Const CHECK_PLATE_AREA As Double = 0.09             'Dimension in square meters
Private Const MAX_PLATE_LENGTH As Double = 12               'Dimension in meters
Private Const MAX_PLATE_WIDTH As Double = 3                 'Dimension in meters
Private Const ROLLER_RADIUS_CONDITION As Double = 75        'Dimension in mm
Private Const ROLLER_DIST_CONDITION As Double = 100         'Dimension in mm
Private Const ROLLER_DIST_TEXT_CONDITION As Double = 300    'Dimension in meters
Private Const MAX_ROLL_RADIUS = 50000
Private Const ERROR_LOG_FILE_NAME = "AnnotationControl.log"
Private Const TWIST_ANGLE = 0.0471238898
Private m_strXMLData As String

Private Const MODULE As String = "StrMfgCustomXML.CTextMark"

Private Function PlaceTextAtPositionAtAngle(sText As String, dLocX As Double, dLocY As Double, dRotAngle As Double, lJustification As Long, Optional dSize As Double = LARGE_PART_TEXT_SIZE, Optional strFont As String = "Arial", Optional strType As String = "") As IJMfgTextMark
    Const METHOD = "PlaceTextAtPosition"
    On Error GoTo ErrorHandler
    Set PlaceTextAtPositionAtAngle = New MfgTextMark
    PlaceTextAtPositionAtAngle.Text = sText
    PlaceTextAtPositionAtAngle.LocX = dLocX
    PlaceTextAtPositionAtAngle.LocY = dLocY
    PlaceTextAtPositionAtAngle.FontSize = dSize
    PlaceTextAtPositionAtAngle.Font = strFont
    PlaceTextAtPositionAtAngle.Justification = lJustification
    PlaceTextAtPositionAtAngle.Purpose = 0
    PlaceTextAtPositionAtAngle.FontStyle = 0
    PlaceTextAtPositionAtAngle.RotAngle = dRotAngle
    PlaceTextAtPositionAtAngle.Type = strType
CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Function

Private Function PlaceTextAtRelativeDistanceAlongCurve(oCurve As IJCurve, dFraction As Double, sText As String, lJustification As Long, Optional dSize As Double = LARGE_PART_TEXT_SIZE, Optional strFont As String = "Arial", Optional strType As String = "") As IJMfgTextMark
    Const METHOD = "PlaceTextAtRelativeDistanceAlongCurve"
    On Error GoTo ErrorHandler
    
    Dim dLocX As Double
    Dim dLocY As Double
    Dim dRotAngle As Double
    FindPositionAndAngleOnCurve oCurve, dFraction, dLocX, dLocY, dRotAngle
    Set PlaceTextAtRelativeDistanceAlongCurve = PlaceTextAtPositionAtAngle(sText, dLocX, dLocY, dRotAngle, lJustification, dSize, strFont, strType)

CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Function

Private Function PlaceTextAtAbsoluteDistanceAlongCurve(oCurve As IJCurve, dLength As Double, sText As String, lJustification As Long, Optional dSize As Double = LARGE_PART_TEXT_SIZE, Optional strFont As String = "Arial", Optional strType As String = "") As IJMfgTextMark
    Const METHOD = "PlaceTextAtAbsoluteDistanceAlongCurve"
    On Error GoTo ErrorHandler
    Dim dLocX As Double
    Dim dLocY As Double
    Dim dRotAngle As Double
    FindPositionAndAngleOnCurve oCurve, dLength / oCurve.length, dLocX, dLocY, dRotAngle
    Set PlaceTextAtAbsoluteDistanceAlongCurve = PlaceTextAtPositionAtAngle(sText, dLocX, dLocY, dRotAngle, lJustification, dSize, strFont, strType)

CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Function


Private Function GenerateRollBoundaryAnnotation(oCurve As IJCurve, dLengthPercent As Double, _
                                                dAdditionalRotation As Double, _
                                                sAnnotType As String, StrForAnnot As String, _
                                                BendRadius As Double, SweepAngle As Double, _
                                                ChordHeight As Double, BendDir As String, _
                                                dSize As Double, Justification As String) As String
    Const METHOD = "GenerateRollBoundaryAnnotation"
    On Error GoTo ErrorHandler

    '*** Get the Position ***'
    Dim dLocX As Double
    Dim dLocY As Double
    Dim dRotAngle As Double
    FindPositionAndAngleOnCurve oCurve, dLengthPercent, dLocX, dLocY, dRotAngle

    Dim oPosition As IJDPosition
    Set oPosition = New DPosition
    oPosition.Set 1000 * dLocX, 1000 * dLocY, 0

    '*** Set the Vector ***'
    Dim oDirVector As IJDVector
    Set oDirVector = New DVector
    oDirVector.Set Cos(dRotAngle + dAdditionalRotation), Sin(dRotAngle + dAdditionalRotation), 0
    oDirVector.length = -1 * oDirVector.length
    
    
    '*** Call Evaluate Method ***'

    Dim strProgID As String
    Dim oAnnotSettings As IXMLDOMElement
    Set oAnnotSettings = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='" & sAnnotType & "']")

    strProgID = CStr(oAnnotSettings.GetAttribute("PROGID"))

    Dim oOutputAnnotation   As IJDMfgOutputAnnotation
    Set oOutputAnnotation = SP3DCreateObject(strProgID)

    '*** Create <SMS_MARKINGS> xml node for mapping into Settings.xml node ***'
    Dim oXMLDOM As New DOMDocument
    
    Dim oXMLElem As IXMLDOMElement
    Set oXMLElem = oXMLDOM.createElement("SMS_MARKINGS")
    
    Dim oXMLChildElem As IXMLDOMElement
    Set oXMLChildElem = oXMLDOM.createElement("SMS_MARKING")
    
    oXMLElem.appendChild oXMLChildElem

    oXMLChildElem.setAttribute "BowCurvatureCode", StrForAnnot
    oXMLChildElem.setAttribute "BendRadius", Str(BendRadius)
    oXMLChildElem.setAttribute "SweepAngle", Str(SweepAngle)
    oXMLChildElem.setAttribute "ChordHeight", Str(ChordHeight)
    oXMLChildElem.setAttribute "BendDirection", BendDir
    
    oXMLChildElem.setAttribute "TextSize", Str(dSize)
    oXMLChildElem.setAttribute "ControlPoint", Justification
    
    Dim sOutputXML As String
    sOutputXML = GetAnnotationXML(oXMLElem.xml, oAnnotSettings.xml)
    
    oOutputAnnotation.SetArguments sOutputXML
    
    GenerateRollBoundaryAnnotation = oOutputAnnotation.Evaluate(oPosition, oDirVector, sOutputXML)

CleanUp:
    Set oPosition = Nothing
    Set oDirVector = Nothing
    Set oAnnotSettings = Nothing
    Set oOutputAnnotation = Nothing
    Set oXMLDOM = Nothing
    Set oXMLElem = Nothing
    Set oXMLChildElem = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp
End Function

'Logging DebugData
Private Sub LogDebugData(sSource As String, Optional sData As String = "", _
                        Optional sMethodName As String = "", _
                        Optional lineNo As Long = 0)
    Dim FileNumber
    Dim sLogFileName As String
    Dim sDebugInfo As String
    
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
        "Debug data: " & Format(Date, "Long Date") & "|" & Format(Time, "Long Time") & _
        vbCrLf & vbCrLf & _
        sDebugInfo & vbCrLf & _
        "###################################################################################"
    
    'Close the file
    Close #FileNumber
End Sub

Private Function GetTempFolder() As String
   ' Returns the path to the user's Temp folder. To boot, Windows
   ' requires that a temporary folder exist, so this should always
   ' safely return a path to one. Just in case, though, check the
   ' return value of GetTempPath.
   Dim strTempPath As String
   Dim lngTempPath As Long
   
   ' Fill string with null characters.
   strTempPath = String(144, vbNullChar)
   ' Get length of string.
   lngTempPath = Len(strTempPath)
   ' Call GetTempPath, passing in string length and string.
   If (GetTempPath(lngTempPath, strTempPath) > 0) Then
      ' GetTempPath returns path into string.
      ' Truncate string at first null character.
      GetTempFolder = Left(strTempPath, _
         InStr(1, strTempPath, vbNullChar) - 1)
   Else
      GetTempFolder = ""
   End If
End Function

Private Function PutMarkingNameAnnotation(ByVal oCurve As Object, ByVal pMfgGeom2d As GSCADStrMfgUtilities.IJMfgGeom2d, _
                                          ByVal pMfgSystemMark As GSCADStrMfgUtilities.IJMfgSystemMark, ByVal strInputXML As String, _
                                          ByVal pTransformationMatrix As AutoMath.IJDT4x4, bRemoveGraphics As Boolean, bstrOutputXML As String, _
                                          Optional ByVal sJustification As String, Optional ByVal dHorOffset As Double, Optional ByVal dVertOffset As Double, _
                                          Optional ByVal bLabelParallel As Boolean, Optional ByVal bVectorFlip As Boolean, Optional ByVal bLengthCheck As Boolean, _
                                          Optional ByVal dLengthCond As Double, Optional ByVal dLengthOffset As Double, _
                                          Optional ByVal dTextSize As Double, Optional oVectorDir As IJDVector, _
                                          Optional sPartType As String) As String

    'Added 8 New Optional Arguments to the function for Positioning of the text accurately
    '***************************** PLACE ENTITY NAME / MARKINGINFO.NAME ***************************************'
    Dim oMarkingDoc         As DOMDocument
    Dim oEntityNameElem     As IXMLDOMElement
    Dim oOutputAnnotation   As IJDMfgOutputAnnotation
    Dim oGeomHelper         As IJMfgGeomHelper
    Dim oVector             As New DVector
    Dim dLength             As Double
    Dim oNameElem           As IXMLDOMElement
    Dim oMarkInfo          As MarkingInfo
    Dim oMiscElem           As IXMLDOMElement
    Dim sRefer              As String
    Dim oSettingsXML        As New DOMDocument
    Dim vRefVal             As Variant
    'Temporarily declared here as a place holder, needed for calling GetAttributesList
    'Set oSettingsXML = Nothing
    oSettingsXML.resolveExternals = True
    oSettingsXML.validateOnParse = False
    
    'Loading the Annotation xml to a Variable of DOMDocument type
    oSettingsXML.Load m_strXMLData
    
    Dim oStartPos   As New DPosition, oEndPos As New DPosition, oPosition As New DPosition
    
    Dim dAlong As Double
    Dim dTempX  As Double, dTempY As Double, dTempZ As Double
    Dim dCurveEndX As Double, dCurveEndY As Double, dTanX As Double, dTanY As Double, dTanZ As Double
    Dim strProgID As String, sNAME As String
    
    On Error GoTo CleanUp
    
    Set oMarkingDoc = New DOMDocument
    Set oGeomHelper = New MfgGeomHelper
    
    If Not oMarkingDoc.loadXML(strInputXML) Then GoTo CleanUp
   
    '*** Get Entity Name Annotation XML ***'
    Set oEntityNameElem = oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='ENTITYNAME']")
    Set oEntityNameElem = UpdateAnnotationPropertyNodes(m_oSettingsXML, oEntityNameElem, sPartType)
    strProgID = oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='ENTITYNAME']/@PROGID").Text
    Set oOutputAnnotation = SP3DCreateObject(strProgID)
   
    Set oNameElem = oMarkingDoc.selectSingleNode("//SMS_MARKING")
    If oNameElem Is Nothing Then GoTo CleanUp
   
    '*** Get Marking Name for Annotation Text ***'
    If Not oNameElem.selectSingleNode("//SMS_MARKING[@MARKING_NAME]") Is Nothing Then
        sNAME = oNameElem.GetAttribute("MARKING_NAME")
    Else
        If Not pMfgGeom2d.SystemMark Is Nothing Then
            Set oMarkInfo = pMfgGeom2d.SystemMark
            sNAME = oMarkInfo.Name
        End If
    End If
    
    If sNAME <> "" Then
        
        '*** Clean the Annotation XML ***'
        'Need to include all the Properties as TEXT_ARG.
        'Going to try and implement GetDefaultTypeAnnotationAttributesList here is the variable name mapping
        'oDefaultXMLDom =       oSettingsXML
        'oDefaultTypeElement =  oEntityNameElem
        'sPartType =            sName??
        
        'Assume we will be able to get access to the Part_Type here as a true representation if it's a
        'Plate or Profile.
        'Set the PartType for testing ONLY!
        Set oEntityNameElem = UpdateAnnotationPropertyNodes(oSettingsXML, oEntityNameElem, sPartType)
        Dim sOutputXML As String
        
        '*** Set Marking Name for Annotation Text ***'
        If sJustification <> vbNullString Then
            Set oMiscElem = oEntityNameElem.selectSingleNode(".//SMS_GEOM_ARG[@NAME='ControlPoint']")
            oMiscElem.setAttribute "VALUE", sJustification
        End If
        
        If dHorOffset <> 0 Then
            Set oMiscElem = oEntityNameElem.selectSingleNode(".//SMS_GEOM_ARG[@NAME='HorizontalOffset']")
            oMiscElem.setAttribute "VALUE", Str(dHorOffset)
        End If
        
        If dVertOffset <> 0 Then
            Set oMiscElem = oEntityNameElem.selectSingleNode(".//SMS_GEOM_ARG[@NAME='VerticalOffset']")
            oMiscElem.setAttribute "VALUE", Str(dVertOffset)
        End If
        
        If dTextSize <> 0 Then
            Set oMiscElem = oEntityNameElem.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
            oMiscElem.setAttribute "VALUE", Str(dTextSize)
        End If
        
        oOutputAnnotation.SetArguments oEntityNameElem.xml
        
                '*** In case of overlap marking, SubgeometryType STRMFG_ANNOTATION_MARK means we must Prefix ANTI_MARKING_CHAR($) to the part name ***'
        'If pMfgGeom2d.GetSubGeometryType = STRMFG_ANNOTATION_MARK Then
'        Set oMarkInfo = pMfgGeom2d.SystemMark
'        oMarkInfo.GetAttributeValue "REFERENCE", vRefVal

        Set oMiscElem = oMarkingDoc.selectSingleNode("//SMS_MARKING")
        vRefVal = oMiscElem.GetAttribute("REFERENCE")
        
        If IsEmpty(vRefVal) Then
            sRefer = vbNullString
        Else
            sRefer = IIf(VarType(vRefVal) = vbString, vRefVal, vbNullString)
        End If
        
        If sRefer = vbNullString Then
            On Error Resume Next
            'As Some Marks Might not have the User Defined "REFERENCE" Attribute
            Set oMarkInfo = pMfgSystemMark
            If oMarkInfo Is Nothing Then
                'MsgBox "oMarkInfo is nothing for geomtype = " & pMfgGeom2d.GetGeometryType
            Else
                oMarkInfo.GetAttributeValue "REFERENCE", vRefVal
                sRefer = IIf(VarType(vRefVal) = vbString, vRefVal, "")
            End If
        End If
        
        If UCase(sRefer) = "OVERLAPNAME" Then
            sNAME = ANTI_MARKING_CHAR & sNAME
        End If
        
        PutGeomGUIDvalue strInputXML, oEntityNameElem

        Set oNameElem = oEntityNameElem.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||MARKING_NAME']")
                
                If oNameElem Is Nothing Then GoTo CleanUp
                
        oNameElem.setAttribute "VALUE", sNAME
        'sOutputXML = GetAnnotationXML(oMarkingDoc.xml, oEntityNameElem.xml)
        
        '*** Get Position and Vector ***'
        oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
        Set oStartPos = New DPosition
        Set oEndPos = New DPosition
        
        oStartPos.Set dTempX, dTempY, dTempZ
        oEndPos.Set dCurveEndX, dCurveEndY, dTanZ
        dLength = oCurve.length
        
        If bLengthCheck = True Then
            If dLength > dLengthCond Then
                Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, dLengthOffset)
                Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
                
                ' If the label to be parallel to mark then swap the vector components
                If bLabelParallel = True Then
                    dTanX = oVector.x
                    oVector.x = oVector.y
                    oVector.y = -dTanX
                End If
                
                'Fix for getting the label on the opposite to the thickness direction annotation
                If bVectorFlip = True Then oVector.length = -1
                
                oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
                bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, oEntityNameElem.xml)
                
                Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oEndPos, dLengthOffset)
                Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
                
                If bLabelParallel = True Then
                    dTanX = oVector.x
                    oVector.x = oVector.y
                    oVector.y = -dTanX
                End If
                
                'Fix for getting the label on the opposite to the thickness direction annotation
                If bVectorFlip = True Then oVector.length = -1
                
                oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
                bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, oEntityNameElem.xml)
            Else
               dAlong = dLength / 2
                Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, dAlong)
                Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
        
                ' If the label to be parallel to mark then swap the vector components
                If bLabelParallel = True Then
                    dTanX = oVector.x
                    oVector.x = oVector.y
                    oVector.y = -dTanX
                End If
                
                'Fix for getting the label on the opposite to the thickness direction annotation
                If bVectorFlip = True Then oVector.length = -1
                
                oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
                bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, oEntityNameElem.xml)
            End If
        Else
            dAlong = dLength / 2
            Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, dAlong)
            
            If oVectorDir Is Nothing Then
                Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
            Else
                Set oVector = oVectorDir
            End If
            ' If the label to be parallel to mark then swap the vector components
            If bLabelParallel = True Then
                dTanX = oVector.x
                oVector.x = oVector.y
                oVector.y = -dTanX
            End If
                
            'Fix for getting the label on the opposite to the thickness direction annotation
            If bVectorFlip = True Then oVector.length = -1
                
            oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
            bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, oEntityNameElem.xml)
        End If
        PutMarkingNameAnnotation = bstrOutputXML
    End If
    '************************************************************************************'
CleanUp:
    Set oMiscElem = Nothing
    Set oNameElem = Nothing
    Set oEntityNameElem = Nothing
    Set oMarkingDoc = Nothing
    Set oGeomHelper = Nothing
    Set oVector = Nothing
    Set oMarkInfo = Nothing
    Set oPosition = Nothing
    Set oStartPos = Nothing
    Set oEndPos = Nothing
End Function

Private Sub IJMfgTextMarkService_CreateAnnotation(ByVal pdispPart As Object, ByVal pMfgGeom2d As GSCADStrMfgUtilities.IJMfgGeom2d, ByVal pMfgSystemMark As GSCADStrMfgUtilities.IJMfgSystemMark, ByVal strInputXML As String, ByVal pTransformationMatrix As AutoMath.IJDT4x4, bRemoveGraphics As Boolean, bstrOutputXML As String)
 On Error GoTo ErrorHandler
    
    Dim oCS                 As IJComplexString
    Dim oOriginalCS         As IJComplexString
    Dim oCurve              As IJCurve
    Dim oStartPos           As IJDPosition
    Dim oEndPos             As IJDPosition
    Dim oElements           As IJElements
    Dim eGeometryType       As StrMfgGeometryType
    Dim eSubGeometryType    As StrMfgGeometryType
    Dim oVector             As IJDVector
    Dim oPosition           As IJDPosition
    Dim oOutputAnnotation   As IJDMfgOutputAnnotation
    Dim oElement            As IXMLDOMElement
    Dim oBevelSetting       As IXMLDOMElement
    Dim oGrindSetting       As IXMLDOMElement
    Dim oMarginSetting      As IXMLDOMElement
    Dim oRollLineSetting    As IXMLDOMElement
    Dim oRollboundarySetting    As IXMLDOMElement
    Dim oCorrSetting        As IXMLDOMElement
    Dim oCorrElem           As IXMLDOMElement
    Dim oFuriwakeSetting    As IXMLDOMElement
    Dim oLabelSetting       As IXMLDOMElement
    Dim oKnuckleSetting     As IXMLDOMElement
    Dim oIBLSetting         As IXMLDOMElement
    Dim oMarginElem         As IXMLDOMElement
    Dim oKnuckleElem        As IXMLDOMElement
    Dim oIBLElem            As IXMLDOMElement
    Dim oRollerElem         As IXMLDOMElement
    Dim oFuriwakeElem       As IXMLDOMElement
    Dim oShipElem           As IXMLDOMElement
    Dim oMarkingNameElem    As IXMLDOMElement
    Dim oInputNode          As IXMLDOMElement
    Dim oLabelNode          As IXMLDOMElement
    Dim oShipDirSetting     As IXMLDOMElement
    Dim oPlateIdInfoSetting As IXMLDOMElement
    Dim oProfileIdInfoSetting As IXMLDOMElement
    Dim oToshiElement       As IXMLDOMElement
    Dim oDeclElement        As IXMLDOMElement
    Dim oCollElement        As IXMLDOMElement
    Dim oCorrugateElement   As IXMLDOMElement
    Dim oEntityNameElem     As IXMLDOMElement
    Dim sMarkingName        As String
    Dim oBevelChangeElem    As IXMLDOMElement
    Dim oNodeList           As IXMLDOMNodeList
    Dim oNode               As IXMLDOMElement
    Dim oGeomHelper         As IJMfgGeomHelper
    Dim oPointList          As IXMLDOMNodeList
    Dim oMarkingDoc         As DOMDocument
    Dim oMfgMGHelper        As New MfgMGHelper
    Dim vTemp               As Variant
    Dim sDir                As String
    Dim dLength             As Double
    Dim dStartParam         As Double
    Dim dEndParam           As Double
    Dim dMidParam           As Double
    Dim dTempLocZ           As Double
    Dim dLocX               As Double
    Dim dLocY               As Double
    Dim dTanX               As Double
    Dim dTanY               As Double
    Dim dTanZ               As Double
    Dim dTempX              As Double
    Dim dTempY              As Double
    Dim dTempZ              As Double
    Dim StartX              As Double
    Dim StartY              As Double
    Dim endX                As Double
    Dim endY                As Double
    Dim dCurveEndX          As Double
    Dim dCurveEndY          As Double
    Dim dAlong              As Double
    Dim dAlong2             As Double
    Dim dAngle              As Double
    Dim bBevelExists        As Boolean
    Dim strProgID           As String
    Dim strProgID2          As String
    Dim strGrindProgID      As String
    Dim strThickDir         As String
    Dim ObjectType          As StrMfgPartType
    Dim oEdgeDoc            As New DOMDocument
    Dim sMarkedSide         As String
    Dim sOutputXML          As String
    Dim oMarkInfo           As MarkingInfo
    Dim dSize               As Double
    Dim oMfgPlate           As IJMfgPlatePart
    Dim dTextSize           As Double
    Dim dSecondaryTextSize  As Double
    Dim oTextSizeElem       As IXMLDOMElement
    Dim oSysMarkingInfo     As MarkingInfo
    Dim varRefAttrVal       As Variant
    Dim oMarkingInfo        As IJMarkingInfo
    Dim bFlipVec            As Boolean
    Dim sTempValue          As String
    Dim oPlatePart          As IJPlatePart
    Dim strGrinder          As String
    Dim strGUID             As String
    Dim sPartType           As String
    Dim sUsage              As String
    Dim vRefVal             As Variant
    Dim strReference        As String
    Dim oReferenceElem      As IXMLDOMElement
    Dim strFaceType         As String
    Dim eTemplateSetType    As MfgTemplateSetType
    Dim oNamedItem          As IJNamedItem
    Dim oMemberMarkElem     As IXMLDOMElement
    
    If Not m_bDoesXMLExist Then Exit Sub
    
    If TypeOf pdispPart Is IJMfgPlatePart Then
        ObjectType = PLATE_TYPE
        Set oMfgPlate = pdispPart
        oMfgPlate.GetDetailedPart oPlatePart
        sPartType = "PLATE"
    ElseIf TypeOf pdispPart Is IJMfgProfilePart Then
        ObjectType = PROFILE_TYPE
        sPartType = "PROFILE"
    ElseIf TypeOf pdispPart Is IJDMfgTemplateSet Then
        ObjectType = TEMPLATE_TYPE
        sPartType = "TEMPLATE"
        
        Dim oMfgTemplateSet  As IJDMfgTemplateSet
        Set oMfgTemplateSet = pdispPart
        
        eTemplateSetType = oMfgTemplateSet.GetTemplateSetType
    End If
    
    Set oGeomHelper = New MfgGeomHelper
    Set oPosition = New DPosition
    
    dTextSize = GetTextSizeForMfgPart(pdispPart)
    
    Set oOriginalCS = pMfgGeom2d.GetGeometry
    
    If ObjectType = TEMPLATE_TYPE Then
        ' In case of Template, Use the original CS as it is not needed/persisted after annotations are created
        Set oCS = oOriginalCS
    Else
        oMfgMGHelper.CloneComplexString oOriginalCS, oCS
    End If
    
    If Not pTransformationMatrix Is Nothing Then
        oCS.Transform pTransformationMatrix
    End If
    Set oCurve = oCS
    eGeometryType = pMfgGeom2d.GetGeometryType
    eSubGeometryType = pMfgGeom2d.GetSubGeometryType
 
    Set oMarkingDoc = New DOMDocument
    If Not oMarkingDoc.loadXML(strInputXML) Then
        oMarkingDoc.resolveExternals = True
        oMarkingDoc.validateOnParse = False
        oMarkingDoc.Load strInputXML
    End If
    
    '*** Read The Reference Attribute for Certain Markings ***'
    If Not oMarkingDoc.selectSingleNode("//SMS_MARKING") Is Nothing Then
        Set oReferenceElem = oMarkingDoc.selectSingleNode("//SMS_MARKING")
        vRefVal = oReferenceElem.GetAttribute("REFERENCE")
        
        If IsEmpty(vRefVal) Then
            strReference = vbNullString
        Else
            strReference = IIf(VarType(vRefVal) = vbString, vRefVal, vbNullString)
        End If
    
        If strReference = vbNullString Then
            On Error GoTo ErrorHandler
            'As Some Marks Might not have the User Defined REFERNCE Attribute
            Set oSysMarkingInfo = pMfgSystemMark
            
            If oSysMarkingInfo Is Nothing Then
                'MsgBox "oSysMarkingInfo is nothing for eGeometryType = " & eGeometryType
            Else
                oSysMarkingInfo.GetAttributeValue "REFERENCE", vRefVal
                If IsEmpty(vRefVal) Then
                    strReference = vbNullString
                Else
                    strReference = IIf(VarType(vRefVal) = vbString, vRefVal, vbNullString)
                End If
            End If
        End If
    End If
    '*********************************************************'
      
    Select Case eGeometryType
        
        Case STRMFG_LAP_REF_MARK
            ' in this case the annotation label is expected to be at an offset of 20 mm from Point of intersection of two lines
            bstrOutputXML = PutMarkingNameAnnotation(oCurve, pMfgGeom2d, pMfgSystemMark, strInputXML, pTransformationMatrix, bRemoveGraphics, bstrOutputXML, "ll", 20 / dTextSize, 20 / dTextSize, , , , , , dTextSize, , sPartType)
        
        Case STRMFG_HOLE_TRACE_MARK
            ' in this case the annotation label is expected to the top of the minor axis
            
            oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
            Set oStartPos = New DPosition

            oStartPos.Set dTempX, dTempY, dTempZ
            Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oStartPos)
            
            ' If the hole trace mark tangent is negative then flip the vector as we always want it's label to be above the mark
            If oVector.x < 0 Then
                bFlipVec = True
            End If
            
            bstrOutputXML = PutMarkingNameAnnotation(oCurve, pMfgGeom2d, pMfgSystemMark, strInputXML, pTransformationMatrix, bRemoveGraphics, bstrOutputXML, "um", 0.05, 2.6, , bFlipVec, , , , dTextSize, , sPartType)
        
        Case STRMFG_FEATURETAB_MARK, STRMFG_KNUCKLE_TAB_MARK, STRMFG_WELDTAB_MARK
            ' in this case the annotation Label is expected to be at a distance of 2/3 * Text Size below the mark
''''            If eSubGeometryType = STRMFG_FEATURETAB_MARK Or eSubGeometryType = STRMFG_WELDTAB_MARK Or eSubGeometryType = STRMFG_KNUCKLE_TAB_MARK Then
''''                Dim CurvesCol As IJElements
''''                oCS.GetCurves CurvesCol
''''                'get first two curves and get their mid points to construct a vector
''''
''''                Dim oCurve1 As IJCurve
''''                Dim oCurve2 As IJCurve
''''                Set oCurve1 = CurvesCol.Item(1)
''''                Set oCurve2 = CurvesCol.Item(2)
''''
''''                Dim dx1 As Double, dy1 As Double, dz1 As Double, dx2 As Double, dy2 As Double, dz2 As Double
''''                oCurve1.Position 0.5, dx1, dy1, dz1
''''                oCurve2.Position 0.5, dx2, dy2, dz2
''''                Dim oPos1 As IJDPosition
''''                Dim oPos2 As IJDPosition
''''
''''                Set oPos1 = New DPosition
''''                Set oPos2 = New DPosition
''''
''''                oPos1.Set dx1, dy1, dz1
''''                oPos2.Set dx2, dy2, dz2
''''
''''                Dim oNewVec As IJDVector
''''
''''                Set oNewVec = oPos1.Subtract(oPos2)
''''                bstrOutputXML = PutMarkingNameAnnotation(oCurve, pMfgGeom2d, pMfgSystemMark, strInputXML, pTransformationMatrix, bRemoveGraphics, bstrOutputXML, "um", 0, -2 / 3, True, True, , , , dTextSize, oNewVec, sPartType)
''''            End If

        Case STRMFG_TAB_REFERENCE_MARK
            
            If pMfgGeom2d.IsSupportOnly = True Then
                Dim oTabMoniker        As IMoniker
                Set oTabMoniker = pMfgGeom2d.GetMoniker
                
                Dim oPOM                As IJDPOM
                Dim oObject             As IJDObject
                
                Set oObject = pdispPart
                Set oPOM = oObject.ResourceManager
                
                Dim oTabBO As IJStrMfgTab
                Set oTabBO = oPOM.GetObject(oTabMoniker)
            
                Dim oHeightAttr As IJDAttribute
                Set oHeightAttr = GetAttribute(oTabBO, "IJUASMPlateTabHeight", "TabHeight")
                                
                Dim oLengthAttr As IJDAttribute
                Set oLengthAttr = GetAttribute(oTabBO, "IJUASMPlateTabLength", "TabLength")

                Dim dTempTextSize As Double
                dTempTextSize = oHeightAttr.Value + (Cos(45) * oLengthAttr.Value)
                
                'convert to mm and set as text size
                dTempTextSize = dTempTextSize * 1000
                dTextSize = dTempTextSize / 2

                bstrOutputXML = PutMarkingNameAnnotation(oCurve, pMfgGeom2d, pMfgSystemMark, strInputXML, pTransformationMatrix, bRemoveGraphics, bstrOutputXML, "lm", 0, -0.225, , True, , , , dTextSize, , sPartType)
            End If
        Case STRMFG_MARGIN_MARK, STRMFG_FEATURECUT_MARK
            ' in this case the annotation Label is expected to be at a distance of 2/3 * Text Size below the mark
            bstrOutputXML = PutMarkingNameAnnotation(oCurve, pMfgGeom2d, pMfgSystemMark, strInputXML, pTransformationMatrix, bRemoveGraphics, bstrOutputXML, "um", 0, -2 / 3, , , , , , dTextSize, , sPartType)

        Case STRMFG_FRAMELINE_MARK
            If ObjectType = PLATE_TYPE Then
                Dim oSDPlatePart As New StructDetailObjects.PlatePart
                Set oSDPlatePart.object = oPlatePart
                
                If oSDPlatePart.plateType = DeckPlate Then
                    bstrOutputXML = PutMarkingNameAnnotation(oCurve, pMfgGeom2d, pMfgSystemMark, strInputXML, pTransformationMatrix, bRemoveGraphics, bstrOutputXML, , , , , , True, LABEL_LENGTH_COND, LABEL_OFFSET_DIST, dTextSize, , sPartType)
                Else
                    ' in this case the annotation Label is expected to be oriented parallel to the mark
                    bstrOutputXML = PutMarkingNameAnnotation(oCurve, pMfgGeom2d, pMfgSystemMark, strInputXML, pTransformationMatrix, bRemoveGraphics, bstrOutputXML, , 2, , True, , True, LABEL_LENGTH_COND, LABEL_OFFSET_DIST, dTextSize, , sPartType)
                End If
            Else
                bstrOutputXML = PutMarkingNameAnnotation(oCurve, pMfgGeom2d, pMfgSystemMark, strInputXML, pTransformationMatrix, bRemoveGraphics, bstrOutputXML, , , , , , True, LABEL_LENGTH_COND, LABEL_OFFSET_DIST, dTextSize, , sPartType)
            End If
            
            If ObjectType = TEMPLATE_TYPE Then
                Exit Sub
            End If
            
        Case STRMFG_BUTTOCKLINE_MARK
            Set oMarkingDoc = New DOMDocument
            If Not oMarkingDoc.loadXML(strInputXML) Then GoTo CleanUp
            
            'CenterLine mark(Y = 0) doesn't need label
            If UCase(strReference) <> "CENTERLINE" Then
                If ObjectType = PLATE_TYPE Then
                    ' in this case the annotation Label is expected to be oriented parallel to the mark
                    bstrOutputXML = PutMarkingNameAnnotation(oCurve, pMfgGeom2d, pMfgSystemMark, strInputXML, pTransformationMatrix, bRemoveGraphics, bstrOutputXML, , 2, , True, , True, LABEL_LENGTH_COND, LABEL_OFFSET_DIST, dTextSize, , sPartType)
                Else
                    bstrOutputXML = PutMarkingNameAnnotation(oCurve, pMfgGeom2d, pMfgSystemMark, strInputXML, pTransformationMatrix, bRemoveGraphics, bstrOutputXML, , , , , , True, LABEL_LENGTH_COND, LABEL_OFFSET_DIST, dTextSize, , sPartType)
                End If
            End If
            
            If ObjectType = TEMPLATE_TYPE Then
                Exit Sub
            End If
        
        Case STRMFG_WATERLINE_MARK
        
            oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
            Set oStartPos = New DPosition

            oStartPos.Set dTempX, dTempY, dTempZ
            Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oStartPos)
            
            ' If the water line mark tangent is negative then flip the vector as we always want it's label to be above the mark
            If oVector.x < 0 Then
                bFlipVec = True
            End If
            
            bstrOutputXML = PutMarkingNameAnnotation(oCurve, pMfgGeom2d, pMfgSystemMark, strInputXML, pTransformationMatrix, bRemoveGraphics, bstrOutputXML, , , , , bFlipVec, True, LABEL_LENGTH_COND, LABEL_OFFSET_DIST, dTextSize, , sPartType)
            
            If ObjectType = TEMPLATE_TYPE Then
                Exit Sub
            End If
            
        Case STRMFG_FEATURE_LABEL_MARK
            ' in this case the annotation Label is expected to be oriented parallel to the mark
            bstrOutputXML = PutMarkingNameAnnotation(oCurve, pMfgGeom2d, pMfgSystemMark, strInputXML, pTransformationMatrix, bRemoveGraphics, bstrOutputXML, , , 1, True, , , , , dTextSize, , sPartType)
        
        Case STRMFG_FITTING_MARK, STRMFG_CONN_PART_MARK, _
             STRMFG_FLAT_OF_BOTTOM_MARK, STRMFG_FLAT_OF_SIDE_MARK, STRMFG_GENERAL_MARK, STRMFG_END_MARK, STRMFG_LAP_MARK, _
             STRMFG_TEMPLATE_MARK, STRMFG_DIRECTION, STRMFG_ELONGATION_MARK

            ' Other Marking types where Annotation Labels are required
            If ObjectType = TEMPLATE_TYPE Then
                If eGeometryType = STRMFG_FITTING_MARK Then
                    bstrOutputXML = PutMarkingNameAnnotation(oCurve, pMfgGeom2d, pMfgSystemMark, strInputXML, pTransformationMatrix, bRemoveGraphics, bstrOutputXML, "lm", 4, 0.5, False, True, False, , , dTextSize, , sPartType)
                    
                ElseIf eGeometryType = STRMFG_DIRECTION Then
                    
                    oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
                    Set oStartPos = New DPosition
                    Set oEndPos = New DPosition
                    oStartPos.Set dTempX, dTempY, dTempZ
                    oEndPos.Set dCurveEndX, dCurveEndY, dTanZ
                    
                    Set oVector = oEndPos.Subtract(oStartPos)
                    oVector.length = 1
                    
                    If Abs(oVector.x) < 0.01 Then
                        If eTemplateSetType = ProfileTemplate Then
                            bstrOutputXML = PutMarkingNameAnnotation(oCurve, pMfgGeom2d, pMfgSystemMark, strInputXML, pTransformationMatrix, bRemoveGraphics, bstrOutputXML, "lm", 0.6, 0.5, True, True, False, , , 15, oVector, sPartType)
                        Else
                            bstrOutputXML = PutMarkingNameAnnotation(oCurve, pMfgGeom2d, pMfgSystemMark, strInputXML, pTransformationMatrix, bRemoveGraphics, bstrOutputXML, "lm", 0.6, 1.5, True, True, False, , , dTextSize, oVector, sPartType)
                        End If
                    Else
                        If eTemplateSetType = ProfileTemplate Then
                            bstrOutputXML = PutMarkingNameAnnotation(oCurve, pMfgGeom2d, pMfgSystemMark, strInputXML, pTransformationMatrix, bRemoveGraphics, bstrOutputXML, "lm", 2, -0.2, False, False, False, , , 15, oVector, sPartType)
                        Else
                            bstrOutputXML = PutMarkingNameAnnotation(oCurve, pMfgGeom2d, pMfgSystemMark, strInputXML, pTransformationMatrix, bRemoveGraphics, bstrOutputXML, "lm", 3, -0.2, False, False, False, , , dTextSize, oVector, sPartType)
                        End If
                    End If
                    
                    Set oVector = oEndPos.Subtract(oStartPos)
                    oVector.length = 1
                    
                Else
                    bstrOutputXML = PutMarkingNameAnnotation(oCurve, pMfgGeom2d, pMfgSystemMark, strInputXML, pTransformationMatrix, bRemoveGraphics, bstrOutputXML, "lm", 1, 0, True, True, False, , , dTextSize, , sPartType)
                End If
            Else
                'We dont want to place marking name if strReference is "Identification"
                ' ... for strReference is "Identification", PlateId/ProfileId will be placed
                If Not (eGeometryType = STRMFG_DIRECTION And UCase(strReference) = "IDENTIFICATION") Then
                bstrOutputXML = PutMarkingNameAnnotation(oCurve, pMfgGeom2d, pMfgSystemMark, strInputXML, pTransformationMatrix, bRemoveGraphics, bstrOutputXML, , , , , , , , , dTextSize, , sPartType)
            End If
            End If
            
        Case STRMFG_FRAMES_CHECKLINES_MARK, STRMFG_USERDEFINED_MARK, STRMFG_SEAM_MARK, STRMFG_BASELINE_MARK, STRMFG_BLOCK_MARK
            If ObjectType = TEMPLATE_TYPE Then
                If Not eGeometryType = STRMFG_SEAM_MARK Then
                    Exit Sub
                End If
            Else
                ' Other Marking types where Annotation Labels are required
                bstrOutputXML = PutMarkingNameAnnotation(oCurve, pMfgGeom2d, pMfgSystemMark, strInputXML, pTransformationMatrix, bRemoveGraphics, bstrOutputXML, , , , , , True, LABEL_LENGTH_COND, LABEL_OFFSET_DIST, dTextSize, , sPartType)
            End If
            
        Case STRMFG_PLATELOCATION_MARK, STRMFG_PROFILELOCATION_MARK, STRMFG_BRACKETLOCATION_MARK
            'Start of Logic for Fixing the Label Location
            Set oMarkingDoc = New DOMDocument
            If Not oMarkingDoc.loadXML(strInputXML) Then GoTo CleanUp
            
            
            Set oElement = oMarkingDoc.selectSingleNode("//SMS_EDGE")
            If oElement Is Nothing Then GoTo CleanUp
            Dim bFlipSinceOverlap As Boolean
            strThickDir = oElement.GetAttribute("PART_DIR")
            
            If UCase(strThickDir) <> "L" Then
                bFlipSinceOverlap = False
                If UCase(strReference) = "OVERLAPFLIP" Then
                    bFlipSinceOverlap = True
                End If
                bstrOutputXML = PutMarkingNameAnnotation(oCurve, pMfgGeom2d, pMfgSystemMark, strInputXML, pTransformationMatrix, bRemoveGraphics, bstrOutputXML, "lm", 0, 0, False, bFlipSinceOverlap, True, LABEL_LENGTH_COND, LABEL_OFFSET_DIST, dTextSize, , sPartType)
            Else
                bFlipSinceOverlap = True
                If UCase(strReference) = "OVERLAPFLIP" Then
                    bFlipSinceOverlap = False
                End If
                bstrOutputXML = PutMarkingNameAnnotation(oCurve, pMfgGeom2d, pMfgSystemMark, strInputXML, pTransformationMatrix, bRemoveGraphics, bstrOutputXML, "lm", 0, 0, False, bFlipSinceOverlap, True, LABEL_LENGTH_COND, LABEL_OFFSET_DIST, dTextSize, , sPartType)
            End If
            
        Case STRMFG_ROLL_LINES_MARK
        
            Set oMarkingDoc = New DOMDocument
            If Not oMarkingDoc.loadXML(strInputXML) Then GoTo CleanUp
            
            If UCase(strReference) = "TRIMMED" Then
                bstrOutputXML = PutMarkingNameAnnotation(oCurve, pMfgGeom2d, pMfgSystemMark, strInputXML, pTransformationMatrix, bRemoveGraphics, bstrOutputXML, "lm", 0, 0, False, False, False, , , dTextSize, , sPartType)
            End If
            
        Case STRMFG_REF_MARK
            If ObjectType = TEMPLATE_TYPE Then
                bstrOutputXML = PutMarkingNameAnnotation(oCurve, pMfgGeom2d, pMfgSystemMark, strInputXML, pTransformationMatrix, bRemoveGraphics, bstrOutputXML, "lm", 1, 0, True, True, False, , , dTextSize, , sPartType)
                Exit Sub
            End If
        
        Case STRMFG_SIGHTLINE_MARK
            Exit Sub
            'End of the Logic for fixing the Label Location
    End Select
    
    If m_oSettingsXML Is Nothing Then
        Exit Sub
    End If
    
    Select Case eGeometryType
        Case STRMFG_FEATURE_MARK
            'Removed Obsolete code
        Case STRMFG_KNUCKLE_MARK
            
            If ObjectType = TEMPLATE_TYPE Then
                
                '***************************** TEMPLATE SEAM_MARK ***************************************'
                Set oMarkingDoc = New DOMDocument
                
                If Not oMarkingDoc.loadXML(strInputXML) Then GoTo CleanUp
                Set oElement = oMarkingDoc.selectSingleNode("//SMS_EDGE")
                If oElement Is Nothing Then GoTo CleanUp
                
                Dim oKnuckleMarkElem    As IXMLDOMElement
                
                Set oKnuckleMarkElem = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='TEMPLATE_KNUCKLE']")
                Set oKnuckleMarkElem = UpdateAnnotationPropertyNodes(m_oSettingsXML, oKnuckleMarkElem, sPartType)
                strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='TEMPLATE_KNUCKLE']/@PROGID").Text
                Set oOutputAnnotation = SP3DCreateObject(strProgID)
            
                Set oTextSizeElem = oKnuckleMarkElem.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
                oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
            
                oOutputAnnotation.SetArguments oKnuckleMarkElem.xml
                
                oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
                Set oStartPos = New DPosition
                Set oEndPos = New DPosition
                
                oStartPos.Set dTempX, dTempY, dTempZ
                oEndPos.Set dCurveEndX, dCurveEndY, dTanZ
                dLength = oCurve.length
                
                dAlong = dLength / 2
                Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, dAlong)
                
                Set oVector = oStartPos.Subtract(oEndPos)
                oVector.length = 1
                
                oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
                
                sOutputXML = GetAnnotationXML(strInputXML, oKnuckleMarkElem.xml)
                
                bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, sOutputXML)
                '************************************************************************************'
            Else
                
                '******************* PUT KNUCKLE MARK ANNOTATION ***************************'
                Dim oKnuckleDoc As DOMDocument
                Set oKnuckleDoc = New DOMDocument
    
                If Not oKnuckleDoc.loadXML(strInputXML) Then GoTo CleanUp

                Set oKnuckleSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='KnuckleLineMark']")
                Set oKnuckleSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oKnuckleSetting, sPartType)
                If Not oKnuckleSetting Is Nothing Then
                    strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='KnuckleLineMark']/@PROGID").Text
                    
                    Set oTextSizeElem = oKnuckleSetting.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
                    oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
    
                    'Create Annotation Object
                    Set oOutputAnnotation = SP3DCreateObject(strProgID)
                    
                    If Not oOutputAnnotation Is Nothing Then
                    
                        'Call SetArguments method
                        oOutputAnnotation.SetArguments oKnuckleSetting.xml
                        
                        'Get Point & Vector
                        oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
                        Set oStartPos = New DPosition
                        Set oEndPos = New DPosition
                        
                        oStartPos.Set dTempX, dTempY, dTempZ
                        oEndPos.Set dCurveEndX, dCurveEndY, dTanZ
                        dLength = oCurve.length
                        
                        '*** Formulate oKnuckleSetting.xml with Angle & Direction ***'
                        'Get Knuckle Angle & convert into degrees
                        Dim dKnuckleAngle As Double
                        Dim oKElement As IXMLDOMElement
                        Dim dMarkedSide As String
                        Dim sKnuckleGUID As String
                        
                        If Not oKnuckleDoc.selectSingleNode(".//SMS_MARKING_KNUCKLE_INFO") Is Nothing Then
                        
                            '*** Set Knuckle Direction ***'
                            Set oKElement = oKnuckleDoc.selectSingleNode(".//SMS_MARKING")
                            dMarkedSide = oKElement.GetAttribute("MARKED_SIDE")
                            
                            If UCase(dMarkedSide) = "MARKING" Then
                                dMarkedSide = MARKING_CHAR
                            ElseIf UCase(dMarkedSide) = "ANTI_MARKING" Then
                                dMarkedSide = ANTI_MARKING_CHAR
                            End If
                            
                            ' If sub geometry is STRMFG_KNUCKLE_FIT_MARK, that means its Knuckle along ..
                            '    .. web, hence for Flat Bar we dont want to show the bendside info.
                            ' In case of Knuckle along Molded Side, we need bendside info for flat bar
                            
                            If ObjectType = PROFILE_TYPE And UCase(strReference) = "KNUCKLEFIT" Then
                                Dim oDetailedProfPart As IJProfilePart
                                Dim oMfgObj2 As IJMfgProfilePart
                                Set oMfgObj2 = pdispPart
                                  
                                Set oDetailedProfPart = oMfgObj2.GetDetailedPart
                                    
                                'Create the SD profile Wrapper and initialize it
                                Dim oSDProfilePartSupport As IJProfilePartSupport
                                Dim oPartSupp As IJPartSupport
                                        
                                'Create the SD profile Wrapper and initialize it
                                If TypeOf oDetailedProfPart Is IJProfilePart Then
                                    Set oSDProfilePartSupport = New ProfilePartSupport
                                    Set oPartSupp = oSDProfilePartSupport
                                    Set oPartSupp.Part = oDetailedProfPart
                                End If
                                Dim eSectionType As ProfileSectionType
                                eSectionType = oSDProfilePartSupport.SectionType
                                
                                If eSectionType = Flat_Bar Then dMarkedSide = ""
                                
                            End If
                            
                            Set oKnuckleElem = oKnuckleSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||KNUCKLE_DIR']")
                            oKnuckleElem.setAttribute "VALUE", Str(dMarkedSide)
                            '*****************************'
                            
                            '*** Set Knuckle Angle ***'
                            Set oKElement = oKnuckleDoc.selectSingleNode(".//SMS_MARKING_KNUCKLE_INFO")
                            vTemp = oKElement.GetAttribute("KNUCKLE_ANGLE")
                            
                            'Set pointer to Settings Knuckle node & set value of KNUCKLE_DIR
    
                            If (isNumericUS(vTemp)) Then
                                'Precision to 1 decimal Point for Knuckle Angle
                                dKnuckleAngle = Trim(Str(Round(Abs(Val(vTemp) * 180 / PI), 1)))
                            End If
                            
                            
                            
                            Set oKnuckleElem = oKnuckleSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||KNUCKLE_ANGLE']")
                            oKnuckleElem.setAttribute "VALUE", Str(dKnuckleAngle)
                            '*****************************'
                            
                            '*** Set GEOM2D_GUID *********'
                            Set oKElement = oKnuckleDoc.selectSingleNode(".//SMS_EDGE")
                            sKnuckleGUID = oKElement.GetAttribute("GEOM2D_GUID")
                            
                            Set oKnuckleElem = oKnuckleSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||GEOM2D_GUID']")
                            oKnuckleElem.setAttribute "VALUE", sKnuckleGUID
                            '*****************************'
                            
                            Set oKElement = Nothing
                           
                            'Apply Length Condition
                            If oCurve.length <= LABEL_LENGTH_COND Then
                                '*** OFFSET 1 ***'
                                dAlong = dLength / 2
                                Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, dAlong)
                                Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
                                
'                            dTanX = oVector.x
'                            oVector.x = -oVector.y
'                            oVector.y = dTanX
                                
                                oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
                                
                                bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, oKnuckleSetting.xml)
                            Else
                                '*** OFFSET 2 ***'
                                Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, LABEL_OFFSET_DIST)
                                Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
                                
                                oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
                               
                                'bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, strInputXML)
                                bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, oKnuckleSetting.xml)
                                
                                '*** OFFSET 2 ***'
                                Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oEndPos, LABEL_OFFSET_DIST)
                                Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
                                
'                            dTanX = oVector.x
'                            oVector.x = -oVector.y
'                            oVector.y = dTanX
                                
                                oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
                                
                                'sOutputXML = GetAnnotationXML(strInputXML, oKnuckleSetting.xml)
                                bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, oKnuckleSetting.xml)
                            End If
                        End If
                    End If
                End If
            End If
            '***************************************************************************'
            
        Case STRMFG_OUTER_CONTOUR, STRMFG_INNER_CONTOUR, STRMFG_COMMON_SEAM_MARK
            If InStr(strInputXML, "SMS_ANNOTATION") > 0 Then
                bBevelExists = True
            Else
                bBevelExists = False
            End If
            
            '**************************** BEVEL AND GRIND ***************************************'
            Dim strBevelXML As String
            Dim bBevel As Boolean
            Dim oGrind2 As IXMLDOMElement
            Dim sProgID2 As String
            Dim oOutputAnnotation2 As IJDMfgOutputAnnotation
            Dim bStartGrind As Boolean
            Dim bEndGrind As Boolean
            Dim bCenterGrind As Boolean
            Dim oPos As New DPosition
            Dim dHeight As Double
            Dim oPointOnCurve As New DPosition
            Dim oDir As New DVector
            Dim sGUID As String
            Dim oContourDoc As DOMDocument
            
            Set oContourDoc = New DOMDocument
            
            If Not oContourDoc.loadXML("<SMS_EDGES>" & strInputXML & "</SMS_EDGES>") Then GoTo CleanUp
            
            Dim oEdgeElem As IXMLDOMElement
            Set oEdgeElem = oContourDoc.selectSingleNode("//*[@GEOM2D_GUID]")
            
            Dim sThisBevelGrindType As String
            sThisBevelGrindType = oEdgeElem.GetAttribute("TYPE")
            
            sGUID = oEdgeElem.GetAttribute("GEOM2D_GUID")
            
            'Set oNodeList = m_oSettingsXML.selectNodes("//SMS_OUTPUT_ANNOTATION[@TYPE='BEVEL_NORMAL_TO_CONTOUR_1']")
            'strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='BEVEL_NORMAL_TO_CONTOUR_1']/@PROGID").Text
                       
            'SupportOnly means color should be green
            If eGeometryType = STRMFG_COMMON_SEAM_MARK And pMfgGeom2d.IsSupportOnly = True Then
                Set oBevelSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='BevelCommonCutSO']")
            'NOT SupportOnly means color should be blue
            ElseIf eGeometryType = STRMFG_COMMON_SEAM_MARK And pMfgGeom2d.IsSupportOnly <> True Then
                Set oBevelSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='BevelCommonCut']")
            'Regular bevel
            Else
            Set oBevelSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='BEVEL NEW 1']")
            End If
            
            Set oBevelSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oBevelSetting, sPartType)
            vTemp = oBevelSetting.GetAttribute("PROGID")
            strProgID = IIf(VarType(vTemp) = vbString, vTemp, "")
            
            Set oTextSizeElem = oBevelSetting.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
            oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
            
            Set oGrind2 = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='GRIND_CENTER']")
            Set oGrind2 = UpdateAnnotationPropertyNodes(m_oSettingsXML, oGrind2, sPartType)
            sProgID2 = oGrind2.GetAttribute("PROGID")
            '***'
            
            Set oGrindSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='GRIND_NORMAL_TO_CONTOUR_1']")
            Set oGrindSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oGrindSetting, sPartType)
            vTemp = oGrindSetting.GetAttribute("PROGID")
            strGrindProgID = IIf(VarType(vTemp) = vbString, vTemp, "")
            
            Set oTextSizeElem = oGrindSetting.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
            oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
            
            If Not (oGrindSetting Is Nothing And oBevelSetting Is Nothing) Then
                If Len(strBevelXML) > -1 Then
                    Dim oBevelDoc As DOMDocument
                    Dim oBevelList As IXMLDOMNodeList
                    Dim oBevelNode As IXMLDOMNode
                    Set oBevelDoc = New DOMDocument
                    Set oBevelList = oContourDoc.selectNodes("//SMS_BEVEL | //SMS_GRIND")
                    
                    For Each oBevelNode In oBevelList
                        Set oElement = oBevelNode
                        'oElement.getElementsByTagName("GUID")
                        'sGUID = oElement.getAttribute("GEOM2D_GUID")
                        oElement.setAttribute "GEOM2D_GUID", sGUID
                        
                        Set oPointList = oElement.getElementsByTagName("CVG_POINT")
                        Set oStartPos = New DPosition
                        Set oEndPos = New DPosition
     
                        oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
                        StartX = Val(oPointList.Item(0).Attributes.getNamedItem("X").Text) / 1000
                        StartY = Val(oPointList.Item(0).Attributes.getNamedItem("Y").Text) / 1000
                        endX = Val(oPointList.Item(1).Attributes.getNamedItem("X").Text) / 1000
                        endY = Val(oPointList.Item(1).Attributes.getNamedItem("Y").Text) / 1000
                        
                        If oCurve.IsPointOn(StartX, StartY, 0) Then
                            oStartPos.Set StartX, StartY, 0
                        ElseIf oCurve.IsPointOn(endX, endY, 0) Then
                            oStartPos.Set endX, endY, 0
                        End If
                        
                        If oBevelNode.nodeName = "SMS_GRIND" Then 'GRIND
                            
                            '*** Get the Grind type ***'
                            Dim oGElem As IXMLDOMElement
                            Set oGElem = oBevelNode
                            
                            strGrinder = oGElem.GetAttribute("NAME")
                            If strGrinder = vbNullString Then
                                strGrinder = "G" 'Default
                            End If
                            
                            '*** Get the Grind type ***'
                            If CheckIfAllGrindsAreSameForOneEdge(oContourDoc, oCurve) Then
                                Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, oCurve.length / 2)
                                Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
                                
                                Set oTextSizeElem = oGrind2.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
                                oTextSizeElem.setAttribute "VALUE", Str(dTextSize)

                                Set oTextSizeElem = oGrind2.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||GRINDER']")
                                oTextSizeElem.setAttribute "VALUE", strGrinder
                                
                                bstrOutputXML = bstrOutputXML & Evaluate(oCurve, oPosition, sProgID2, oGrind2, sGUID)
                                
                                Exit For
                            End If
                            
                            'CASE 1 - Grind_StartPoint = Curve_StartPoint and Grind_EndPoint = Curve_EndPoint
                            If ((Abs(StartX - dTempX) < 0.001 And Abs(StartY - dTempY) < 0.001) And (Abs(endX - dCurveEndX) < 0.001 And Abs(endY - dCurveEndY) < 0.001)) Or (Abs(dTempX - dCurveEndX) < 0.001 And Abs(dTempY - dCurveEndY) < 0.001) Then
  
                                'Get the ProgId
                                Set oGrind2 = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='GRIND_CENTER']")
                                Set oGrind2 = UpdateAnnotationPropertyNodes(m_oSettingsXML, oGrind2, sPartType)
                                sProgID2 = oGrind2.GetAttribute("PROGID")
                                
                                Set oTextSizeElem = oGrind2.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
                                oTextSizeElem.setAttribute "VALUE", Str(dTextSize)

                                Set oTextSizeElem = oGrind2.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||GRINDER']")
                                oTextSizeElem.setAttribute "VALUE", strGrinder
                                
                                bCenterGrind = True
                                
                                'CASE 1.1  If Curve_StartPoint = Curve_EndPoint, then its a closed contour,
                                '          Place Annotation on Right Side
                                If Abs(StartX - endX) < 0.001 And Abs(StartX - endY) < 0.001 Then

                                    oDir.Set 1, 0, 0
                                    oGeomHelper.GetMaxHeightAndExtremaPointOnCurve oCurve, oDir, dHeight, oPointOnCurve

                                    Set oPosition = oPointOnCurve
                                    
                                    bstrOutputXML = bstrOutputXML & Evaluate(oCurve, oPosition, sProgID2, oGrind2, sGUID)
                                    Exit For 'All Grinds on this edge are same, just place annotation once.
                                    
                                'CASE 1.2  If Curve_StartPoint <> Curve_EndPoint, then its NOT a closed contour,
                                '          Place Annotation on Mid-Point
                                Else
                                
                                    '*** Get the Mid-point ***'
                                    If Abs(dCurveEndX - endX) < 0.001 And Abs(dCurveEndY - endY) < 0.001 Then
                                        oCurve.ParamRange dTempZ, dEndParam
                                        oCurve.Parameter StartX, StartY, 0, dStartParam
                                        oCurve.Evaluate ((dEndParam - dStartParam) * 0.95) + dStartParam, endX, endY, dTempZ, dTempLocZ, dTempY, dTanZ, dTempX, dTempY, dTempZ
                                    End If
                                    oEndPos.Set endX, endY, 0
                                    Set oElements = oGeomHelper.DivideCurveEqualSpace(oCurve, oStartPos, oEndPos, 3)
    
                                    If Not oElements Is Nothing Then
                                        oCurve.Parameter oElements.Item(2).x, oElements.Item(2).y, 0, dMidParam
                                        oCurve.Evaluate dMidParam, dLocX, dLocY, dTempLocZ, dTanX, dTanY, dTanZ, dTempX, dTempY, dTempZ
                                        oPosition.Set oElements.Item(2).x, oElements.Item(2).y, 0
                                    Else: GoTo CleanUp
                                    End If
                                    dTanZ = 0
                                    '*************************'
                                    bstrOutputXML = bstrOutputXML & Evaluate(oCurve, oPosition, sProgID2, oGrind2, sGUID)
                                End If
                                
                            Else 'Means either Grind-Start or Grind-End
                            
                                If (Abs(StartX - dTempX) > 0.001 Or Abs(StartY - dTempY) > 0.001) Then 'Means Start Points are Different
                                    Set oGrind2 = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='GRIND_START']")
                                    Set oGrind2 = UpdateAnnotationPropertyNodes(m_oSettingsXML, oGrind2, sPartType)
                                    sProgID2 = oGrind2.GetAttribute("PROGID")
                                        
                                    Set oTextSizeElem = oGrind2.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
                                    oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
                                    
                                    Set oTextSizeElem = oGrind2.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||GRINDER']")
                                    oTextSizeElem.setAttribute "VALUE", strGrinder
                                    
                                    bStartGrind = True
                                    oPos.Set StartX, StartY, 0
                                    
                                    bstrOutputXML = bstrOutputXML & Evaluate(oCurve, oPos, sProgID2, oGrind2, sGUID)
                                    
                                End If
                                'Means Grind-End
                                If (Abs(endX - dCurveEndX) > 0.001 Or Abs(endY - dCurveEndY) > 0.001) Then 'Means End Points are different
                                    Set oGrind2 = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='GRIND_END']")
                                    Set oGrind2 = UpdateAnnotationPropertyNodes(m_oSettingsXML, oGrind2, sPartType)
                                    sProgID2 = oGrind2.GetAttribute("PROGID")
                                    
                                    Set oTextSizeElem = oGrind2.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
                                    oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
                                    
                                    Set oTextSizeElem = oGrind2.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||GRINDER']")
                                    oTextSizeElem.setAttribute "VALUE", strGrinder
                                    
                                    bEndGrind = True
                                    
                                    oPos.Set endX, endY, 0
                                    
                                    bstrOutputXML = bstrOutputXML & Evaluate(oCurve, oPos, sProgID2, oGrind2, sGUID)
                                    
                                End If
                                
                            End If
                            '****************'
                        Else 'BEVELS
                            If (Not bBevelExists) Then
                            
                                Dim dAngleA As Double, dAngleB As Double, dAngleN As Double, dAngleD As Double, dAngleE As Double
                                Dim dChamferAngleM As Double, dChamferDepthM As Double
                                Dim dChamferAngleUM As Double, dChamferDepthUM As Double
                                Dim dTotal As Double, dTotalChamfer As Double
                                Dim bIsDashigiri As Boolean
                                
                                bIsDashigiri = False 'Initialization
                                
                                FillBevelValuesFromXML oBevelNode.xml, sGUID, dAngleA, _
                                                        dAngleB, dAngleN, dAngleD, dAngleE, dChamferAngleM, dChamferDepthM, _
                                                        dChamferAngleUM, dChamferDepthUM
                                
                                dTotal = Round(dAngleA + dAngleB + dAngleN + dAngleD + dAngleE, 1)
                                
                                dTotalChamfer = Round(dChamferAngleM + dChamferDepthM + dChamferAngleUM + dChamferDepthUM, 1)
                                
                                
                                If dTotal = 0 And dTotalChamfer = 0 And eGeometryType <> STRMFG_COMMON_SEAM_MARK Then
                                    GoTo NextNode
                                ElseIf dTotal <> 0 Or dTotalChamfer <> 0 Or eGeometryType = STRMFG_COMMON_SEAM_MARK Then
                                    Dim strPurpose As String
                                    Dim oMyElement As IXMLDOMElement
                                    Set oMyElement = oBevelNode
                                    strPurpose = oMyElement.GetAttribute("PURPOSE")
                                    
                                    If UCase(strPurpose) = "NO_OFFSET" Then
                                        Set oBevelSetting = m_oSettingsXML.selectSingleNode(".//SMS_OUTPUT_ANNOTATION[@TYPE='DashigiriType']")
                                        Set oBevelSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oBevelSetting, sPartType)
                                        vTemp = oBevelSetting.GetAttribute("PROGID")
                                        strProgID = IIf(VarType(vTemp) = vbString, vTemp, "")
                                        Set oTextSizeElem = oBevelSetting.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
                                        oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
                                        Set oOutputAnnotation = SP3DCreateObject(strProgID)

                                        Set oMyElement = oBevelSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||DESCRIPTION']")
                                        oMyElement.setAttribute "VALUE", Str(dTotal)
                                        bIsDashigiri = True
                                    Else
                                        Set oOutputAnnotation = SP3DCreateObject(strProgID)
                                    End If
                                Else
                                    GoTo NextNode
                                End If
                                
                                If Not oOutputAnnotation Is Nothing Then
                                    oOutputAnnotation.SetArguments oBevelSetting.xml
                                    bBevel = True
                                    
                                    Dim dtest As Double
                                    '*** Get the Mid-Point ***'
                                    If Abs(dTempX - StartX) < 0.002 And Abs(dTempY - StartY) < 0.002 Then
                                        oCurve.Parameter StartX, StartY, 0, dStartParam
                                        oCurve.Evaluate (dStartParam) + 0.05, StartX, StartY, dtest, dTempLocZ, dTempY, dTanZ, dTempX, dTempY, dTempZ
                                    End If
                                    If Abs(dCurveEndX - endX) < 0.002 And Abs(dCurveEndY - endY) < 0.002 Then
                                                oCurve.ParamRange dTempZ, dEndParam
                                                oCurve.Parameter StartX, StartY, 0, dStartParam
                                        oCurve.Evaluate (dEndParam) - 0.05, endX, endY, dtest, dTempLocZ, dTempY, dTanZ, dTempX, dTempY, dTempZ
                                    End If
                                    oStartPos.Set StartX, StartY, 0
                                    oEndPos.Set endX, endY, 0
                                    Set oElements = oGeomHelper.DivideCurveEqualSpace(oCurve, oStartPos, oEndPos, 3)
    
                                    If Not oElements Is Nothing Then
                                        oCurve.Parameter oElements.Item(2).x, oElements.Item(2).y, 0, dMidParam
                                        oCurve.Evaluate dMidParam, dLocX, dLocY, dTempLocZ, dTanX, dTanY, dTanZ, dTempX, dTempY, dTempZ
                                        oPosition.Set oElements.Item(2).x, oElements.Item(2).y, 0
                                    Else
                                        GoTo CleanUp
                                    End If
                                    dTanZ = 0
                                    
                                    Dim sStrXML As String
                      
                                    sStrXML = GetAnnotationXML(oBevelNode.xml, oBevelSetting.xml)
                                    Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
                                    
                                    'If bIsDashigiri = True And ObjectType = PROFILE_TYPE Then
                                    If ObjectType = PROFILE_TYPE And UCase(sThisBevelGrindType) <> "EDGE" Then
                                        ' We want to orient bevel annotations in standard x/y/z axis orientation only if its not on Edge Feature
                                        Dim oXVector As IJDVector, oYVector As IJDVector, oZVector As IJDVector
                                        Set oXVector = New DVector
                                        Set oYVector = New DVector
                                        Set oZVector = New DVector
                                        
                                        oXVector.Set 1, 0, 0
                                        oYVector.Set 0, 1, 0
                                        oZVector.Set 0, 0, 1
                                        
                                        Dim dXDotP As Double, dYDotP As Double, dZDotP As Double
                                        
                                        dXDotP = oVector.Dot(oXVector)
                                        dYDotP = oVector.Dot(oYVector)
                                        dZDotP = oVector.Dot(oZVector)
                                        
                                        If ((Abs(dXDotP) > Abs(dYDotP)) And (Abs(dXDotP) > Abs(dZDotP))) Then
                                            Set oVector = oXVector
                                            If dXDotP < 0 Then
                                                oVector.length = -1
                                            End If
                                        ElseIf ((Abs(dYDotP) > Abs(dXDotP)) And (Abs(dYDotP) > Abs(dZDotP))) Then
                                            Set oVector = oYVector
                                            If dYDotP < 0 Then
                                                oVector.length = -1
                                            End If
                                        Else
                                            Set oVector = oZVector
                                            If dZDotP < 0 Then
                                                oVector.length = -1
                                            End If
                                        End If
                                    End If
                                    
                                    oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
                                    
                                    bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, sStrXML)
                                    'bstrOutputXML = bstrOutputXML & Evaluate(oCurve, oPosition, strProgID, oBevelNode, sGUID)
                                    
                                Else
                                    GoTo NextNode
                                End If
                            End If
                        End If
NextNode:
                    Next oBevelNode
                End If
            End If
            '************************************************************************************'
            
            '*** IF COMMON_SEAM_MARK, then place arrow annotation ***'
            If eGeometryType = STRMFG_COMMON_SEAM_MARK And pMfgGeom2d.IsSupportOnly = True Then
            
                
                Dim pResMgr As IUnknown
                Set pResMgr = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
                Dim oMoniker As IMoniker
                Set oMoniker = pMfgGeom2d.GetMoniker
                If oMoniker Is Nothing Then GoTo CleanUp
                Dim oTempObj As IUnknown
                Dim m_oMfgRuleHelper As New MfgRuleHelpers.Helper
                Set oTempObj = m_oMfgRuleHelper.BindToObject(pResMgr, oMoniker)
                Dim oFacePort As IJPort
                Set oFacePort = oTempObj ' lateral faceport

                Dim oConnectedPart As Object
                Set oConnectedPart = oFacePort.Connectable
                Dim sThisPartName As String, sParentName As String, sGeomName As String, sChildName As String
'''                Set oNamedItem = pdispPart
'''                sThisPartName = oNamedItem.Name

                Dim sConnPartName As String
'''                Set oNamedItem = oConnectedPart
'''                sConnPartName = oNamedItem.Name
                
                'Get the parent object of connected part (to make sure we are placing annotation
                '  ... only for common seam between parent and child
                Dim oEntityHelper As GSCADStrMfgUtilities.MfgEntityHelper
                Set oEntityHelper = New MfgEntityHelper
                
                Dim oParentObj As Object, oParentAssy As Object
                oEntityHelper.GetParentPartAndAssembly pdispPart, oParentObj, oParentAssy
                                
                'Proceeed only if the common seam is between child and parent
                Set oNamedItem = oParentObj
                sParentName = oNamedItem.Name
                
                Set oNamedItem = pMfgGeom2d
                sGeomName = oNamedItem.Name
                                
                Dim nFoundPos   As Long
                nFoundPos = 0
                nFoundPos = VBA.InStr(1, sGeomName, "|")
                If nFoundPos <> 0 Then
                    'Left Part is Parent Name, RightPart is Child Name
                    sChildName = VBA.Right(sGeomName, (Len(sGeomName) - nFoundPos))
                    sGeomName = VBA.Left(sGeomName, (nFoundPos - 1))
                    
                End If
                
                'Getting the last section name of the name (SKDY requirement)
                sChildName = "&" & Mid(sChildName, InStrRev(sChildName, "-") + 1)
                
                If sGeomName <> sParentName Then GoTo CleanUp
                
                Dim bIfLongestCommonSeam As Boolean
                bIfLongestCommonSeam = CheckIfLongestCommonSeam(pdispPart, pMfgGeom2d)
                
                '2 = Block, 21 = Ogumi, 22 = Chugumi, 23 = Kogumi, 24 = T-Block, 25 = PS-Block
                If bIfLongestCommonSeam And GetAssemblyType(oConnectedPart) = 1 And UCase(GetProductionRoutingStageCode(oConnectedPart) = "StageA") Then
                
                    Set oMarkingDoc = New DOMDocument
        
                    If Not oMarkingDoc.loadXML(strInputXML) Then GoTo CleanUp
                    Set oElement = oMarkingDoc.selectSingleNode("//SMS_EDGE")
                    If oElement Is Nothing Then GoTo CleanUp
                    
                    Set oMemberMarkElem = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='ArrowMark1']")
                    Set oMemberMarkElem = UpdateAnnotationPropertyNodes(m_oSettingsXML, oMemberMarkElem, sPartType)
                    strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='ArrowMark1']/@PROGID").Text
                    Set oOutputAnnotation = SP3DCreateObject(strProgID)
                
                    Set oTextSizeElem = oMemberMarkElem.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
                    oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
                    
                    Set oTextSizeElem = oMemberMarkElem.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||DESCRIPTION']")
                    oTextSizeElem.setAttribute "VALUE", sChildName
                    
                    oOutputAnnotation.SetArguments oMemberMarkElem.xml
                    
                    oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
                    Set oStartPos = New DPosition
                    Set oEndPos = New DPosition
                    
                    oStartPos.Set dTempX, dTempY, dTempZ
                    oEndPos.Set dCurveEndX, dCurveEndY, dTanZ
                    dLength = oCurve.length
                    
                    dAlong = dLength / 2
                    Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, dAlong)
                    Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
                    
                    dTanX = oVector.x
                    oVector.x = oVector.y
                    oVector.y = -dTanX
                    oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
                    
                    sOutputXML = GetAnnotationXML(strInputXML, oMemberMarkElem.xml)
                        
                    bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, sOutputXML)
                    
                End If
            End If
            '********************************************************'
            
            
            
        Case STRMFG_PLATELOCATION_MARK, STRMFG_PROFILELOCATION_MARK, STRMFG_BRACKETLOCATION_MARK
            '***************************** PLATE LOCATION ***************************************'
            Set oMarkingDoc = New DOMDocument

            If Not oMarkingDoc.loadXML(strInputXML) Then GoTo CleanUp
            Set oElement = oMarkingDoc.selectSingleNode("//SMS_EDGE")
            If oElement Is Nothing Then GoTo CleanUp
            
            If ObjectType = PLATE_TYPE Or (ObjectType = PROFILE_TYPE And pMfgGeom2d.IsSupportOnly = False) Then
                Set oNode = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='PLATE_THICKNESS_1']")
                Set oNode = UpdateAnnotationPropertyNodes(m_oSettingsXML, oNode, sPartType)
                strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='PLATE_THICKNESS_1']/@PROGID").Text
                Set oOutputAnnotation = SP3DCreateObject(strProgID)
                
                Set oTextSizeElem = oNode.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
                oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
                
                sOutputXML = GetAnnotationXML(strInputXML, oNode.xml)
                oOutputAnnotation.SetArguments sOutputXML
                
                oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
                Set oStartPos = New DPosition
                Set oEndPos = New DPosition
                
                oStartPos.Set dTempX, dTempY, dTempZ
                oEndPos.Set dCurveEndX, dCurveEndY, dTanZ
                dLength = oCurve.length
                
                If dLength > LABEL_LENGTH_COND Then
                    Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, LABEL_OFFSET_DIST)
                    Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
    
                    dTanX = oVector.x
                    oVector.x = oVector.y
                    oVector.y = -dTanX
                    oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
                    bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, sOutputXML)
    
                    Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oEndPos, LABEL_OFFSET_DIST)
                    Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
    
                    dTanX = oVector.x
                    oVector.x = oVector.y
                    oVector.y = -dTanX
                    oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
                    bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, sOutputXML)
                Else
                    dAlong = dLength / 2
                    Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, dAlong)
                    Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
                    
                    dTanX = oVector.x
                    oVector.x = oVector.y
                    oVector.y = -dTanX
                    oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0

                    bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, sOutputXML)
                End If
            
            '************************************************************************************'
                If eGeometryType = STRMFG_PROFILELOCATION_MARK Then
                   '**************************** PROFILE LOCATION ********************************'
                   Set oMarkingDoc = New DOMDocument
                   Set oGeomHelper = New MfgGeomHelper
                   If Not oMarkingDoc.loadXML(strInputXML) Then GoTo CleanUp
                   Set oElement = oMarkingDoc.selectSingleNode("//SMS_EDGE")
                   If oElement Is Nothing Then GoTo CleanUp
                   strThickDir = oElement.GetAttribute("PART_DIR")
                   
                   If strThickDir <> "R" And strThickDir <> "L" Then GoTo CleanUp
                   
                   Set oNode = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='PROFILE_FLANGE_THICKNESS_1']")
                   Set oNode = UpdateAnnotationPropertyNodes(m_oSettingsXML, oNode, sPartType)
                   strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='PROFILE_FLANGE_THICKNESS_1']/@PROGID").Text
                   Set oOutputAnnotation = SP3DCreateObject(strProgID)
            
                   Set oTextSizeElem = oNode.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
                   oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
                   
                   sOutputXML = GetAnnotationXML(strInputXML, oNode.xml)
                   
                   oOutputAnnotation.SetArguments sOutputXML
                   oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
                   Set oStartPos = New DPosition
                   Set oEndPos = New DPosition
                   oStartPos.Set dTempX, dTempY, dTempZ
                   oEndPos.Set dCurveEndX, dCurveEndY, dTanZ
                   dLength = oCurve.length
        'Commented code is for TSN Bubbles:
                   'If dLength < 5 Then
                       If dLength < 0.075 Then
                           GoTo CleanUp
                       Else
                           dAlong = dLength / 2
                       End If
            
                       Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, dAlong)
                       Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
                       
                       If strThickDir = "R" Then
                           dTanX = oVector.x
                           oVector.x = oVector.y
                           oVector.y = -dTanX
                       ElseIf strThickDir = "L" Then
                           dTanX = oVector.x
                           oVector.x = -oVector.y
                           oVector.y = dTanX
                       End If
            
                       oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
                       bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, sOutputXML)
                End If
                '************************************************************************************'
            End If
        Case STRMFG_BEVEL_MARK
            '***************************** BEVEL CHANGE MARK/TOMARI ***************************************'
            Set oMarkingDoc = New DOMDocument

            If Not oMarkingDoc.loadXML(strInputXML) Then GoTo CleanUp
            Set oElement = oMarkingDoc.selectSingleNode("//SMS_EDGE")
            If oElement Is Nothing Then GoTo CleanUp
            
            Set oBevelChangeElem = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='TOMARI_MARK_1']")
            Set oBevelChangeElem = UpdateAnnotationPropertyNodes(m_oSettingsXML, oBevelChangeElem, sPartType)
            strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='TOMARI_MARK_1']/@PROGID").Text
            Set oOutputAnnotation = SP3DCreateObject(strProgID)
        
            Set oTextSizeElem = oBevelChangeElem.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
            oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
        
            oOutputAnnotation.SetArguments oBevelChangeElem.xml
            
            oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
            Set oStartPos = New DPosition
            Set oEndPos = New DPosition
            
            oStartPos.Set dTempX, dTempY, dTempZ
            oEndPos.Set dCurveEndX, dCurveEndY, dTanZ
            dLength = oCurve.length
            
            dAlong = dLength / 2
            Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, dAlong)
            Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
            
            dTanX = oVector.x
            oVector.x = oVector.y
            oVector.y = -dTanX
            oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
            
            sOutputXML = GetAnnotationXML(strInputXML, oBevelChangeElem.xml)
                
            bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, sOutputXML)
            '************************************************************************************'
            
        Case STRMFG_DIRECTION
        
            Set oMarkingDoc = New DOMDocument
            If Not oMarkingDoc.loadXML(strInputXML) Then GoTo CleanUp
            
            Set oReferenceElem = oMarkingDoc.selectSingleNode("//SMS_MARKING")
            vRefVal = oReferenceElem.GetAttribute("REFERENCE")

            If IsEmpty(vRefVal) Then
                strReference = vbNullString
            Else
                strReference = IIf(VarType(vRefVal) = vbString, vRefVal, vbNullString)
            End If
            
            '*** This is customer specific ShipDirection & PlateId, so check for subgeometry type ***'
            If UCase(strReference) = "IDENTIFICATION" Then
                If ObjectType = PLATE_TYPE Then
                    Set oMarkInfo = pMfgSystemMark
                    Dim oMarkVector As IJDVector
                    Dim oVVec As IJDVector, oZVec As IJDVector
                    Dim sDirection As Variant
                    Dim oOutputAnnot1   As IJDMfgOutputAnnotation
                    Dim oOutputAnnot2   As IJDMfgOutputAnnotation
    
    
                    sDirection = oMarkInfo.FlangeDirection
                    'Dim HDir As Variant
                    Dim sType As Variant
                    Dim PlusDir As Variant
    
                    oMarkInfo.GetAttributeValue "PART_TYPE", sType
        '            oMarkInfo.GetAttributeValue "H", HDir
                    oMarkInfo.GetAttributeValue "ANNOTATION_SIDE", PlusDir
    
                    If oEdgeDoc.loadXML(strInputXML) Then
    
                        oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
                        Set oStartPos = New DPosition
                        Set oEndPos = New DPosition
    
                        oStartPos.Set dTempX, dTempY, dTempZ
                        oEndPos.Set dCurveEndX, dCurveEndY, dTanZ
                        Set oZVec = New DVector
                        oZVec.Set 0, 0, 1
    
                        Set oMarkVector = oEndPos.Subtract(oStartPos)
    
                        If sDirection = "Opposite" Then
                            oMarkVector.length = -1#
                        End If
    
                        Set oVVec = oZVec.Cross(oMarkVector)
    
                        If sType = 1 Then
                            If PlusDir = "R" Then
                                Set oShipDirSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='ShipDirection']")
                                Set oShipDirSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oShipDirSetting, sPartType)
                                Set oPlateIdInfoSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='PlateIdInfo']")
                                Set oPlateIdInfoSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oPlateIdInfoSetting, sPartType)
                            Else
                                Set oShipDirSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='ShipDirectionInvert']")
                                Set oShipDirSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oShipDirSetting, sPartType)
                                Set oPlateIdInfoSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='PlateIdInfoInvert']")
                                Set oPlateIdInfoSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oPlateIdInfoSetting, sPartType)
                            End If
                        ElseIf sType = 2 Then
                            If PlusDir = "R" Then
                                Set oShipDirSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='ShipDirection']")
                                Set oShipDirSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oShipDirSetting, sPartType)
                                Set oPlateIdInfoSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='PlateIdInfo']")
                                Set oPlateIdInfoSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oPlateIdInfoSetting, sPartType)
                            Else
                                Set oShipDirSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='ShipDirectionInvert']")
                                Set oShipDirSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oShipDirSetting, sPartType)
                                Set oPlateIdInfoSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='PlateIdInfoInvert']")
                                Set oPlateIdInfoSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oPlateIdInfoSetting, sPartType)
                            End If
                        ElseIf sType = 3 Then
                            If PlusDir = "L" Then
                                Set oShipDirSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='ShipDirection']")
                                Set oShipDirSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oShipDirSetting, sPartType)
                                Set oPlateIdInfoSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='PlateIdInfo']")
                                Set oPlateIdInfoSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oPlateIdInfoSetting, sPartType)
                            Else
                                Set oShipDirSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='ShipDirectionInvert']")
                                Set oShipDirSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oShipDirSetting, sPartType)
                                Set oPlateIdInfoSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='PlateIdInfoInvert']")
                                Set oPlateIdInfoSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oPlateIdInfoSetting, sPartType)
                            End If
                        End If
    
                        Set oTextSizeElem = oShipDirSetting.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
                        oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
                       
                        
                        dSecondaryTextSize = GetSecondaryTextSizeForMfgPart(pdispPart)
                        
                        If dSecondaryTextSize <> 0 Then
                            Set oTextSizeElem = oPlateIdInfoSetting.selectSingleNode(".//SMS_GEOM_ARG[@NAME='ShipNumber']")
                            oTextSizeElem.setAttribute "TEXTSIZE2", dSecondaryTextSize
    
                            Set oTextSizeElem = oPlateIdInfoSetting.selectSingleNode(".//SMS_GEOM_ARG[@NAME='PartName']")
                            oTextSizeElem.setAttribute "TEXTSIZE2", dSecondaryTextSize
    
                            Set oTextSizeElem = oPlateIdInfoSetting.selectSingleNode(".//SMS_GEOM_ARG[@NAME='ErectionStage']")
                            oTextSizeElem.setAttribute "TEXTSIZE2", dSecondaryTextSize
                        End If
                        
                        If Not oShipDirSetting Is Nothing Then
    
                            '*** Get ShipDirection ProgId ***'
                            vTemp = oShipDirSetting.GetAttribute("PROGID")
                            strProgID = IIf(VarType(vTemp) = vbString, vTemp, "")
    
                            '*** Get Part Id Info ProgId ***'
                            vTemp = oPlateIdInfoSetting.GetAttribute("PROGID")
                            strProgID2 = IIf(VarType(vTemp) = vbString, vTemp, "")
    
                            Set oShipElem = oEdgeDoc.selectSingleNode("//*[@GUID]")
            '                    sGUID = oShipElem.getAttribute("GUID")
    
                            dLength = oCurve.length
                            dAlong = 1 * dLength / 2
    
                            Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, dAlong)
                            Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
    
                            dTanX = oVector.x
                            oVector.x = oVector.y
                            oVector.y = -dTanX
    
                            oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
    
                            oVector.length = 1
    
                            '*** Create Ship Direction Annotation Object ***'
                            sOutputXML = GetAnnotationXML(strInputXML, oShipDirSetting.xml)
    
                            '*** For LBHD Exchange the U & V values ***'
                            If sType = 3 Then
                                Dim oSetting As New DOMDocument
                                Dim oUElement As IXMLDOMElement, oVElement As IXMLDOMElement
                                Dim sUSide As String, sVSide As String
    
                                oSetting.loadXML sOutputXML
    
                                Set oUElement = oSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||USIDE']")
                                sUSide = oUElement.GetAttribute("VALUE")
                                Set oVElement = oSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||VSIDE']")
                                sVSide = oVElement.GetAttribute("VALUE")
    
                                oUElement.setAttribute "VALUE", sVSide
                                oVElement.setAttribute "VALUE", sUSide
                                sOutputXML = oSetting.xml
                            End If
                            '******************************************'
    
                            Set oOutputAnnot1 = SP3DCreateObject(strProgID)
                            oOutputAnnot1.SetArguments sOutputXML
    
                            'bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, strInputXML)
                            bstrOutputXML = bstrOutputXML & oOutputAnnot1.Evaluate(oPosition, oVector, sOutputXML)
    
                            '*** Create Part Id Info Annotation Object ***'
                            sOutputXML = GetAnnotationXML(strInputXML, oPlateIdInfoSetting.xml)
    
                            '*** Customized Output ***'
                            'Append "T" to Thickness
                            Dim oTempElement As IXMLDOMElement
                            Dim dThick As Double
                            Dim lNum As Long
                            Dim lPortQty As Long, lCentQty As Long, lSBQty As Long
                            
                            oSetting.loadXML sOutputXML
                            Set oTempElement = oSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||THICKNESS']")
                            dThick = Val(oTempElement.GetAttribute("VALUE"))
                            'To have it in format e.g. 45.0 -> 45 (NOT 45.)
                            dThick = Format(dThick, "0.#")
                            oTempElement.setAttribute "VALUE", "T" & dThick
                            
                            ' Append "P", "C", & "S" to QuantityPort, QuantityCenter & QuantityStarboard respectively
                            Set oTempElement = oSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||QUANTITY_PORT']")
                            If Not oTempElement Is Nothing And oTempElement.GetAttribute("VALUE") <> "" Then
                                lPortQty = oTempElement.GetAttribute("VALUE")
                                If lPortQty <> 0 Then
                                    oTempElement.setAttribute "VALUE", "P" & lPortQty
                                Else
                                    oTempElement.setAttribute "VALUE", ""
                                End If
                            End If
                            
                            Set oTempElement = oSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||QUANTITY_CENTER']")
                            If Not oTempElement Is Nothing And oTempElement.GetAttribute("VALUE") <> "" Then
                                lCentQty = oTempElement.GetAttribute("VALUE")
                                If lCentQty <> 0 Then
                                    oTempElement.setAttribute "VALUE", "C" & lCentQty
                                Else
                                    oTempElement.setAttribute "VALUE", ""
                                End If
                            End If
                            
                            Set oTempElement = oSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||QUANTITY_STARBOARD']")
                            If Not oTempElement Is Nothing And oTempElement.GetAttribute("VALUE") <> "" Then
                                lSBQty = oTempElement.GetAttribute("VALUE")
                                If lSBQty <> 0 Then
                                    oTempElement.setAttribute "VALUE", "S" & lSBQty
                                Else
                                    oTempElement.setAttribute "VALUE", ""
                                End If
                            End If
                            
                            '*** If no quantity information available in SMS_COMMON_PART_INFO,  ***'
                            '  then check PART_BOARDSIDE attribute and update the respective quantity
                            If lPortQty = 0 And lCentQty = 0 And lSBQty = 0 Then
                                Set oTempElement = oSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||PART_BOARDSIDE']")
                                If Not oTempElement Is Nothing And oTempElement.GetAttribute("VALUE") <> "" Then
                                    sTempValue = oTempElement.GetAttribute("VALUE")
                                End If
                                
                                Select Case UCase(sTempValue)
                                    Case "P"
                                        Set oTempElement = oSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||QUANTITY_PORT']")
                                        oTempElement.setAttribute "VALUE", "P1"
                                    Case "C"
                                        Set oTempElement = oSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||QUANTITY_CENTER']")
                                        oTempElement.setAttribute "VALUE", "C1"
                                    Case "S"
                                        Set oTempElement = oSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||QUANTITY_STARBOARD']")
                                        oTempElement.setAttribute "VALUE", "S1"
                                End Select
                            End If
                            '**********************************************************************'
                            
                            
                            '*** Add Quotes / Special Characters to Grade/Primer/Routing etc  *****'
                            Dim sPrefixText As String, sSuffixText As String
                            
                            Set oTempElement = oSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||MATERIAL_GRADE']")
                            If Not oTempElement Is Nothing And oTempElement.GetAttribute("VALUE") <> "" Then
                            sTempValue = oTempElement.GetAttribute("VALUE")
                                
                                If oTempElement.GetAttribute("PREFIX") <> "" Then
                                    sPrefixText = GetXMLEquivalentCharacter(oTempElement.GetAttribute("PREFIX"))
                                End If
                                If oTempElement.GetAttribute("SUFFIX") <> "" Then
                                    sSuffixText = GetXMLEquivalentCharacter(oTempElement.GetAttribute("SUFFIX"))
                                End If
                                
                            If sTempValue = "A" Then
                                oTempElement.setAttribute "VALUE", ""
                                
                                ElseIf sTempValue <> "" Then
                                    oTempElement.setAttribute "VALUE", sPrefixText & sTempValue & sSuffixText
                                End If
                            End If
                            
                            Set oTempElement = oSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||PRIMER']")
                            If Not oTempElement Is Nothing And oTempElement.GetAttribute("VALUE") <> "" Then
                                sTempValue = oTempElement.GetAttribute("VALUE")
                                
                                If oTempElement.GetAttribute("PREFIX") <> "" Then
                                    sPrefixText = GetXMLEquivalentCharacter(oTempElement.GetAttribute("PREFIX"))
                                End If
                                If oTempElement.GetAttribute("SUFFIX") <> "" Then
                                    sSuffixText = GetXMLEquivalentCharacter(oTempElement.GetAttribute("SUFFIX"))
                                End If
                                
                                If sTempValue <> "" Then
                                    oTempElement.setAttribute "VALUE", sPrefixText & sTempValue & sSuffixText
                                End If
                            End If
                            
                            Set oTempElement = oSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||ROUTING']")
                            If Not oTempElement Is Nothing And oTempElement.GetAttribute("VALUE") <> "" Then
                                sTempValue = oTempElement.GetAttribute("VALUE")
                                
                                If oTempElement.GetAttribute("PREFIX") <> "" Then
                                    sPrefixText = GetXMLEquivalentCharacter(oTempElement.GetAttribute("PREFIX"))
                                End If
                                If oTempElement.GetAttribute("SUFFIX") <> "" Then
                                    sSuffixText = GetXMLEquivalentCharacter(oTempElement.GetAttribute("SUFFIX"))
                                End If
                                
                                If sTempValue <> "" Then
                                    oTempElement.setAttribute "VALUE", sPrefixText & sTempValue & sSuffixText
                                End If
                            End If
                            '**********************************************************************'
                            
                            sOutputXML = oSetting.xml
                            '*************************'
                            Set oOutputAnnot2 = SP3DCreateObject(strProgID2)
                            oOutputAnnot2.SetArguments sOutputXML
    
                            bstrOutputXML = bstrOutputXML & oOutputAnnot2.Evaluate(oPosition, oVector, sOutputXML)
                            
                        End If
                    End If
                ElseIf ObjectType = PROFILE_TYPE Then
                    Set oProfileIdInfoSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='PROFILE_ID1']")
                    Set oProfileIdInfoSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oProfileIdInfoSetting, sPartType)
                    strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='PROFILE_ID1']/@PROGID").Text

                    Set oOutputAnnotation = SP3DCreateObject(strProgID)
                    
                    Set oTextSizeElem = oProfileIdInfoSetting.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
                    oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
                    
                    '*** Check Twisted Info ***'
                    Set oShipElem = oMarkingDoc.selectSingleNode("//SMS_PART_INFO")
                   
                    'Get Plate Part
                    Dim oDetailProfPart As IJProfilePart
                    Dim oMfgProfPart As IJMfgProfilePart
                    Set oMfgProfPart = pdispPart
                    Set oDetailProfPart = oMfgProfPart.GetDetailedPart
                    'oMfgProfPart.GetDetailedPart oDetailProfPart
                    
                    Dim oRuleHelpr As New MfgProfilePartHlpr
                    Set oRuleHelpr.object = pdispPart
                    Dim oProfPartHelper As New ProfilePartHlpr
                    Set oProfPartHelper.object = oDetailProfPart
                   
                    'If Not oMarkingDoc.selectSingleNode("//SMS_PART_INFO[@CURVATURE]") Is Nothing Then
                    If oRuleHelpr.GetTwistAngle > TWIST_ANGLE Or CheckIfProfileIsPlaneBend(oDetailProfPart) = True Then
                        oShipElem.setAttribute "CURVATURE", "NEJIRE ARI"
                    Else
                        oShipElem.setAttribute "CURVATURE", ""
                    End If
                    'Since oMarkingDoc is changed, we need to copy it to strInputXML
                    strInputXML = oMarkingDoc.xml
                    '**************************'
                    
                    If Not oMarkingDoc.selectSingleNode(".//SMS_MARKING[@MARKING_NAME]") Is Nothing Then
                        Dim oShipDirMarkInfo As MarkingInfo
                        'set oShipDirMarkInfo =
                        Set oShipElem = oMarkingDoc.selectSingleNode(".//SMS_MARKING[@MARKING_NAME]")
                        
                        Dim sUString As String
                        sUString = oShipElem.GetAttribute("MARKING_NAME")
                        
                        If Not oMarkingDoc.selectSingleNode(".//SMS_PART_INFO[@DEVELOPMENT_DIR_U]") Is Nothing Then
                            Set oShipElem = oMarkingDoc.selectSingleNode(".//SMS_PART_INFO[@DEVELOPMENT_DIR_U]")
                            oShipElem.setAttribute "DEVELOPMENT_DIR_U", sUString
                        End If

                        strInputXML = oMarkingDoc.xml
                    End If
                    
                    '*** Check for conditions on Global Direction Attribute ***'
                    'Do NOT put Global Direction as part of annotation in following conditions
                    '   1. No Bevel on Profile
                    '   2. No marking lines on profile
                    '   3. No Scallop/Slots on Profile
                    '   4. If end cuts on both ends have same shape
                    Dim oRuleHlpr As New MfgProfilePartHlpr
                    Set oRuleHlpr.object = pdispPart
                   
                    If IfBevelExistsForProfile(pdispPart) = False And _
                        IfMarkingExistsForProfile(pdispPart) = False And _
                        oRuleHlpr.CheckIfFeaturesExist = False And _
                        oRuleHlpr.CheckIfEndsAreIdentical(WEB_END, 0.01) = True Then
                        'This code is to NOT put the ShipDirection (U-V direction) of Profile Id
                            If Not oMarkingDoc.selectSingleNode(".//SMS_PART_INFO[@DEVELOPMENT_DIR_U]") Is Nothing Then
                                Set oShipElem = oMarkingDoc.selectSingleNode(".//SMS_PART_INFO[@DEVELOPMENT_DIR_U]")
                                oShipElem.setAttribute "DEVELOPMENT_DIR_U", ""

                                strInputXML = oMarkingDoc.xml
                            End If
                    End If
                    '**********************************************************'
                    
                    Dim oProfReport As IJMfgProfileReport
                    Set oProfReport = pdispPart
                    Dim dProfLength As Double
                    
                    dProfLength = oProfReport.GetProcessingLength(STRMFG_UnfoldedLength, JXSEC_WEB_LEFT)
                   
                    If dProfLength <= 1 Then
                        Set oShipElem = oProfileIdInfoSetting.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
                        oShipElem.setAttribute "VALUE", "9"
                    End If
                    
                    If dProfLength < 0.4 Then
                        If Not oProfileIdInfoSetting.selectSingleNode(".//SMS_GEOM_ARG[@NAME='PartName']") Is Nothing Then
                            Set oShipElem = oProfileIdInfoSetting.selectSingleNode(".//SMS_GEOM_ARG[@NAME='PartName']")
                            oShipElem.setAttribute "VALUE", "NEWLINE"
                        End If
                        If Not oProfileIdInfoSetting.selectSingleNode(".//SMS_GEOM_ARG[@NAME='PartNameOffset']") Is Nothing Then
                            Set oShipElem = oProfileIdInfoSetting.selectSingleNode(".//SMS_GEOM_ARG[@NAME='PartNameOffset']")
                            oShipElem.setAttribute "VALUE", "3"
                        End If
                    End If

                    
                    oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
                    Set oStartPos = New DPosition
                    Set oEndPos = New DPosition

                    oStartPos.Set dTempX, dTempY, dTempZ
                    oEndPos.Set dCurveEndX, dCurveEndY, dTanZ
                    dLength = oCurve.length

                    dAlong = dLength / 2
                    Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, dAlong)
                    Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)

                    oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0

                    sOutputXML = GetAnnotationXML(strInputXML, oProfileIdInfoSetting.xml)
                    
                    oOutputAnnotation.SetArguments sOutputXML
                    
                    bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, sOutputXML)
                End If
            Else
                If ObjectType = TEMPLATE_TYPE Then
                    
                    ' Exit for horizontal ship direction mark
                    If Abs(oVector.x) > 0.01 Then
                        Exit Sub
                    End If
                    
                    Dim oTemplateIdInfoSetting  As IXMLDOMElement
                    Set oTemplateIdInfoSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='TEMPLATE_SHIP_ANNOTATION']")
                    Set oTemplateIdInfoSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oTemplateIdInfoSetting, sPartType)
                    strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='TEMPLATE_SHIP_ANNOTATION']/@PROGID").Text
                    
                    Set oOutputAnnotation = SP3DCreateObject(strProgID)
                    
                    Set oTextSizeElem = oTemplateIdInfoSetting.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
                    If eTemplateSetType = ProfileTemplate Then
                        oTextSizeElem.setAttribute "VALUE", 15
                    Else
                        oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
                    End If
                    
                    Dim oTemplateSetRep    As IJMfgTemplateSetReport
                    Set oTemplateSetRep = pdispPart
                    

                    Set oNamedItem = oTemplateSetRep.GetParentPart
                    
                    ' Get SHIP_NAME
            
                    Dim oMfgEntityHelper As IJMfgEntityHelper
                    Set oMfgEntityHelper = New MfgEntityHelper
                    
                    Dim oProjectRoot As IJProjectRoot
                    Set oProjectRoot = oMfgEntityHelper.GetConfigProjectRoot
        
                    Dim strShipName As String
                    strShipName = oProjectRoot.Name
                    
                    Dim oShipNameElem  As IXMLDOMElement
                    Dim oPartNameElem  As IXMLDOMElement
                    
                    Set oShipNameElem = oTemplateIdInfoSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||SHIP_NAME']")
                    oShipNameElem.setAttribute "VALUE", strShipName
                    
                    Set oPartNameElem = oTemplateIdInfoSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||PART_NAME']")
                    oPartNameElem.setAttribute "VALUE", oNamedItem.Name
                    
                    '*** Check Twisted Info ***'
                    Set oShipElem = oMarkingDoc.selectSingleNode("//SMS_PART_INFO")
                   
                    strInputXML = oMarkingDoc.xml
                    '**************************'
                    
                    oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
                    Set oStartPos = New DPosition
                    Set oEndPos = New DPosition

                    oStartPos.Set dTempX, dTempY, dTempZ
                    oEndPos.Set dCurveEndX, dCurveEndY, dTanZ
                    
                    Set oVector = New DVector 'oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
                    
                    oVector.Set 1, 0, 0

                    oPosition.Set oEndPos.x * 1000, oEndPos.y * 1000, 0

                    sOutputXML = GetAnnotationXML(strInputXML, oTemplateIdInfoSetting.xml)
                    
                    oOutputAnnotation.SetArguments sOutputXML
                    
                    bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, sOutputXML)
                    
                End If
                
            End If

        Case STRMFG_SEAM_MARK, STRMFG_FRAMES_CHECKLINES_MARK
        
            If ObjectType = TEMPLATE_TYPE Then
                
                '***************************** TEMPLATE SEAM_MARK ***************************************'
                Set oMarkingDoc = New DOMDocument
                
                If Not oMarkingDoc.loadXML(strInputXML) Then GoTo CleanUp
                Set oElement = oMarkingDoc.selectSingleNode("//SMS_EDGE")
                If oElement Is Nothing Then GoTo CleanUp
                
                Dim oSeamMarkElem    As IXMLDOMElement
                
                Set oSeamMarkElem = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='TEMPLATE_SEAM']")
                Set oSeamMarkElem = UpdateAnnotationPropertyNodes(m_oSettingsXML, oSeamMarkElem, sPartType)
                strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='TEMPLATE_SEAM']/@PROGID").Text
                Set oOutputAnnotation = SP3DCreateObject(strProgID)
            
                Set oTextSizeElem = oSeamMarkElem.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
                oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
            
                oOutputAnnotation.SetArguments oSeamMarkElem.xml
                
                oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
                Set oStartPos = New DPosition
                Set oEndPos = New DPosition
                
                oStartPos.Set dTempX, dTempY, dTempZ
                oEndPos.Set dCurveEndX, dCurveEndY, dTanZ
                dLength = oCurve.length
                
                dAlong = dLength / 2
                Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, dAlong)
                
                Set oVector = oStartPos.Subtract(oEndPos)
                oVector.length = 1
                oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
                
                sOutputXML = GetAnnotationXML(strInputXML, oSeamMarkElem.xml)
                
                bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, sOutputXML)
                '************************************************************************************'
            
            Else
                '**************************** SEAM CONTROL ******************************************'
                
                Dim dSeamOffset1 As Double
                Dim dLengthCondition As Double
                
                Set oMarkingDoc = New DOMDocument
                If Not oMarkingDoc.loadXML(strInputXML) Then GoTo CleanUp
                
                Set oReferenceElem = oMarkingDoc.selectSingleNode("//SMS_MARKING")
                vRefVal = oReferenceElem.GetAttribute("REFERENCE")
                
                If IsEmpty(vRefVal) Then
                    strReference = vbNullString
                Else
                    strReference = IIf(VarType(vRefVal) = vbString, vRefVal, vbNullString)
                End If
                
                Set oElement = oMarkingDoc.selectSingleNode("//SMS_EDGE")
                If oElement Is Nothing Then GoTo CleanUp
                
                Set oNode = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='SeamControl2']")
                Set oNode = UpdateAnnotationPropertyNodes(m_oSettingsXML, oNode, sPartType)
                strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='SeamControl2']/@PROGID").Text
            
                If UCase(strReference) = "SEAMCONTROL" Then
                    Set oNode = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='SeamControl3']")
                    Set oNode = UpdateAnnotationPropertyNodes(m_oSettingsXML, oNode, sPartType)
                    strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='SeamControl3']/@PROGID").Text
                End If
                
                Set oOutputAnnotation = SP3DCreateObject(strProgID)
                
                Set oTextSizeElem = oNode.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
                oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
                
                oOutputAnnotation.SetArguments oNode.xml
                
                oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
                Set oStartPos = New DPosition
                Set oEndPos = New DPosition
                
                oStartPos.Set dTempX, dTempY, dTempZ
                oEndPos.Set dCurveEndX, dCurveEndY, dTanZ
                dLength = oCurve.length
                
                'For WebFrame Check Lines Mark
                If eGeometryType = STRMFG_FRAMES_CHECKLINES_MARK Then
                    
                    dAlong = dLength / 2
                    Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, dAlong)
                    Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
                    
                    oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
                    
                    sOutputXML = GetAnnotationXML(strInputXML, oNode.xml)
                    
                    'bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, strInputXML)
                    bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, sOutputXML)
                
                
                Else
                    'For SeamControl Mark
                    '*** If Length < 3000 mm, put one X on center ***'
                    If dLength <= LABEL_LENGTH_COND Then
                        dAlong = dLength / 2
                        Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, dAlong)
                        Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
                        
                        oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
                        
                        sOutputXML = GetAnnotationXML(strInputXML, oNode.xml)
                        
                        'bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, strInputXML)
                        bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, sOutputXML)
                    Else
                        '*** OFFSET 1 ***'
                        Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, LABEL_OFFSET_DIST)
                        Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
                        
                        oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
                        
                        sOutputXML = GetAnnotationXML(strInputXML, oNode.xml)
                        
                        'bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, strInputXML)
                        bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, sOutputXML)
                        
                        '*** OFFSET 2 ***'
                        Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oEndPos, LABEL_OFFSET_DIST)
                        Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
                        
                        oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
                        
                        sOutputXML = GetAnnotationXML(strInputXML, oNode.xml)
                        bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, sOutputXML)
                    End If
                End If
            End If
            '************************************************************************************'
        Case STRMFG_CONN_PART_MARK
            '**************************** SEAM CONTROL ******************************************'
            Dim dTOLERANCE As Double
            dTOLERANCE = 0.0001
            'If eSubGeometryType <> STRMFG_CONN_PART_MARK Then
            
            Set oMarkingDoc = New DOMDocument

            If Not oMarkingDoc.loadXML(strInputXML) Then GoTo CleanUp
            
            If UCase(strReference) <> "OFFSETEDGE" Then
                
                Set oElement = oMarkingDoc.selectSingleNode("//SMS_EDGE")
                If oElement Is Nothing Then GoTo CleanUp
    
                Dim dMoldedSideDifference As Variant, dBaseBaseDist As Variant, dBaseOffsetDist As Variant
                Dim dOffsetBaseDist As Variant, dOffsetOffsetDist As Variant, dDistDifference As Variant
                Dim sUpside As Variant
                Dim sThickness As String
                
                Set oMarkInfo = pMfgSystemMark
                oMarkInfo.GetAttributeValue "MoldedSideDifference", dMoldedSideDifference
                oMarkInfo.GetAttributeValue "BaseBaseDifference", dBaseBaseDist
                oMarkInfo.GetAttributeValue "BaseOffsetDifference", dBaseOffsetDist
                oMarkInfo.GetAttributeValue "OffsetBaseDifference", dOffsetBaseDist
                oMarkInfo.GetAttributeValue "OffsetOffsetDifference", dOffsetOffsetDist
                oMarkInfo.GetAttributeValue "UpSideDifference", dDistDifference
                oMarkInfo.GetAttributeValue "UpSide", sUpside
                
                If dMoldedSideDifference > dTOLERANCE And dDistDifference > dTOLERANCE Then
                    If dMoldedSideDifference > dTOLERANCE And dBaseBaseDist > dTOLERANCE And dOffsetOffsetDist > dTOLERANCE And _
                        dBaseOffsetDist > dTOLERANCE And dOffsetBaseDist > dTOLERANCE Then
                        Set oFuriwakeSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='FURIWAKEMARK']")
                        Set oFuriwakeSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oFuriwakeSetting, sPartType)
                        strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='FURIWAKEMARK']/@PROGID").Text
                        Set oFuriwakeElem = oFuriwakeSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||THICKNESS']")
                        sThickness = Str(Round(dDistDifference * 1000, 2))
                        oFuriwakeElem.setAttribute "VALUE", sThickness 'Convert to mmm
                    Else
                        Set oFuriwakeSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='TSURAICHIMARK']")
                        Set oFuriwakeSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oFuriwakeSetting, sPartType)
                        strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='TSURAICHIMARK']/@PROGID").Text
                    End If
                Else
                    GoTo CleanUp
                End If
                
                Set oFuriwakeElem = oFuriwakeSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||UPSIDE']")
    
                If sUpside = 1 Then sDir = MARKING_CHAR Else sDir = ANTI_MARKING_CHAR
                oFuriwakeElem.setAttribute "VALUE", sDir
                   
                Set oOutputAnnotation = SP3DCreateObject(strProgID)
                Set oTextSizeElem = oFuriwakeSetting.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
                oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
                oOutputAnnotation.SetArguments oFuriwakeSetting.xml
                
                oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
                Set oStartPos = New DPosition
                Set oEndPos = New DPosition
                
                oStartPos.Set dTempX, dTempY, dTempZ
                oEndPos.Set dCurveEndX, dCurveEndY, dTanZ
                dLength = oCurve.length
            
            '*** Put on center ***'
                dAlong = dLength / 2
                Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, dAlong)
                Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
                
                strThickDir = oElement.GetAttribute("PART_SIDE")
                On Error GoTo ErrorHandler
            
                If strThickDir = "R" Then
                    oVector.length = -1
                End If
                
                
                oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
                
                sOutputXML = GetAnnotationXML(strInputXML, oFuriwakeSetting.xml)
                
                'bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, strInputXML)
                bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, sOutputXML)
            '************************************************************************************'
            End If
        Case STRMFG_LAP_MARK
            '***************************** LAP MARK ANNOTATION ***************************************'
            
            If TypeOf oPlatePart Is IJCollarPart Then
                Exit Sub
            End If
            
            Dim oElem As IXMLDOMElement
            Set oMarkingDoc = New DOMDocument

            If Not oMarkingDoc.loadXML(strInputXML) Then GoTo CleanUp
            Set oElement = oMarkingDoc.selectSingleNode("//SMS_EDGE")
            If oElement Is Nothing Then GoTo CleanUp
            Set oElem = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='LappedMark']")
            Set oElem = UpdateAnnotationPropertyNodes(m_oSettingsXML, oElem, sPartType)
            strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='LappedMark']/@PROGID").Text
            Set oOutputAnnotation = SP3DCreateObject(strProgID)
            
            Set oTextSizeElem = oElem.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
            oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
            
            oOutputAnnotation.SetArguments oElem.xml
                        
            oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
            Set oStartPos = New DPosition
            Set oEndPos = New DPosition
            
            oStartPos.Set dTempX, dTempY, dTempZ
            oEndPos.Set dCurveEndX, dCurveEndY, dTanZ
            
            Dim oGeomCS As IJComplexString
            Set oGeomCS = oCurve
            
            Dim oCurveColl As IJElements
            oGeomCS.GetCurves oCurveColl
            
            Dim dMaxLength As Double
            Dim dMaxLengthCurve As IJCurve
            
            If oCurveColl.Count > 1 Then
                Dim lCount As Long
                For lCount = 1 To oCurveColl.Count
                    Dim oTempCurve As IJCurve
                    Set oTempCurve = oCurveColl.Item(lCount)
                    
                    If dMaxLength < oTempCurve.length Then
                        dMaxLength = oTempCurve.length
                        Set dMaxLengthCurve = oTempCurve
                    End If
                Next
            Else
                Set dMaxLengthCurve = oCurve
            End If
            
            dLength = dMaxLengthCurve.length
            
            If ObjectType = PROFILE_TYPE Then
                dAlong = dLength / 2
            Else
                dAlong = dLength * (1 / 3)
            End If
            
            Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(dMaxLengthCurve, oStartPos, dAlong)
            Set oVector = oGeomHelper.GetTangentByPointOnCurve(dMaxLengthCurve, oPosition)
            
            On Error Resume Next ' as the direction is not set always on the marking info
            strThickDir = oElement.GetAttribute("PART_SIDE")
            On Error GoTo ErrorHandler
            
            If strThickDir = "R" Then
                oVector.length = -1
            End If
            
            oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
            sOutputXML = GetAnnotationXML(strInputXML, oElem.xml)

            bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, sOutputXML)

            '************************************************************************************'
            
        Case STRMFG_LAP_TRACE_MARK
            '**************************** TOSHI MARK Annotation******************************************'
            Set oMarkingDoc = New DOMDocument

            If Not oMarkingDoc.loadXML(strInputXML) Then GoTo CleanUp

            Set oElem = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='LappedMark2']")
            Set oElem = UpdateAnnotationPropertyNodes(m_oSettingsXML, oElem, sPartType)
            strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='LappedMark2']/@PROGID").Text
            Set oOutputAnnotation = SP3DCreateObject(strProgID)

            Set oCS = pMfgGeom2d.GetGeometry
            GetLappedPlateAnnotationPosition oCS, oPosition
            
            Set oTextSizeElem = oElem.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
            oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
            
            oOutputAnnotation.SetArguments oElem.xml

            Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)

            If oVector.x < 0 Or oVector.y < 0 Then
                oVector.x = Abs(oVector.x)
                oVector.y = Abs(oVector.y)
            End If
            
            oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
            sOutputXML = GetAnnotationXML(strInputXML, oElem.xml)
            
            bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, sOutputXML)
            
            '************************************************************************************'
            
        Case STRMFG_BUTTOCKLINE_MARK
            
            '**************************** CENTER LINE MARK Annotation******************************************'
            'CenterLine annotation is needed only when the type of grid plane is centerline

            Set oMarkingDoc = New DOMDocument
            If Not oMarkingDoc.loadXML(strInputXML) Then GoTo CleanUp
            
            If UCase(strReference) <> "CENTERLINE" Then
                Exit Sub
            End If
            
            Set oElement = oMarkingDoc.selectSingleNode("//SMS_EDGE")
            If oElement Is Nothing Then GoTo CleanUp
            
            Set oNode = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='CenterlineMark']")
            Set oNode = UpdateAnnotationPropertyNodes(m_oSettingsXML, oNode, sPartType)
            strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='CenterlineMark']/@PROGID").Text
            
            Set oOutputAnnotation = SP3DCreateObject(strProgID)
            
            Set oTextSizeElem = oNode.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
            oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
            
            oOutputAnnotation.SetArguments oNode.xml

            oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
            Set oStartPos = New DPosition
            
            oStartPos.Set dTempX, dTempY, dTempZ
            
            dLength = oCurve.length
            
            dAlong = dLength / 2
            Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, dAlong)
            Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
            
            oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
            
            sOutputXML = GetAnnotationXML(strInputXML, oNode.xml)
            
            bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, sOutputXML)
            
            '************************************************************************************'
            
            
                
'Commenting since Margin Mark is placed as a Mark (not Annotation)
'        Case STRMFG_MARGIN_MARK
'            '**************************** MARGIN ************************************************'
'
'            If eSubGeometryType <> STRMFG_ANNOTATION_MARK Then
'                If oEdgeDoc.loadXML(strInputXML) Then
'                    Set oMarginSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='MARGIN_MARK_2']")
'                    Set oMarginSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oMarginSetting, sPartType)
'
'                    vTemp = oMarginSetting.getAttribute("PROGID")
'                    strProgID = IIf(VarType(vTemp) = vbString, vTemp, "")
'
'                    Set oMarginElem = oEdgeDoc.selectSingleNode("//*[@GEOM2D_GUID]")
'                    sGUID = oMarginElem.getAttribute("GEOM2D_GUID")
'
'                    'Set Annotation's LENGTH attribute
'                    '  According to the LENGTH value, the annotation service will place
'                    '     the Big/Small Margin Mark
'                    Set oMarginElem = oMarginSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||LENGTH']")
'                    oMarginElem.setAttribute "VALUE", oCurve.Length * 1000  'Convert to mmm
'
'                    oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
'                    Set oStartPos = New DPosition
'                    Set oEndPos = New DPosition
'
'                    oStartPos.Set dTempX, dTempY, dTempZ
'                    dAlong = oCurve.Length / 2
'                    Set oGeomHelper = New MfgGeomHelper
'                    Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, dAlong)
'                    Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
'
'                    oVector.Length = 1
'                    bstrOutputXML = bstrOutputXML & vbNewLine _
'                                    & Evaluate(oCurve, oPosition, strProgID, oMarginSetting, sGUID)
'                End If
'SkipMargin:
'                Set oEdgeDoc = Nothing
'                Set oMarginElem = Nothing
'            End If
'            '************************************************************************************'
        
        Case STRMFG_BENDING_CONTROLLINES_MARK
            '******************* PUT BENDING LINE ANNOTATION ***************************'
            Dim oIBLDoc As DOMDocument
            Set oIBLDoc = New DOMDocument
            
            If Not oIBLDoc.loadXML(strInputXML) Then GoTo CleanUp
          
            Set oIBLSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='IBL_CONTROL']")
            Set oIBLSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oIBLSetting, sPartType)
            If Not oIBLSetting Is Nothing Then
                strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='IBL_CONTROL']/@PROGID").Text
                
                'Create Annotation Object
                Set oOutputAnnotation = SP3DCreateObject(strProgID)
                
                If Not oOutputAnnotation Is Nothing Then
                    Set oTextSizeElem = oIBLSetting.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
                    oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
                    
                    'Call SetArguments method
                    oOutputAnnotation.SetArguments oIBLSetting.xml
                    
                    'Get Point & Vector
                    oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
                    Set oStartPos = New DPosition
                    Set oEndPos = New DPosition
                    
                    oStartPos.Set dTempX, dTempY, dTempZ
                    oEndPos.Set dCurveEndX, dCurveEndY, dTanZ
                    dLength = oCurve.length
                    
                    '*** Formulate oIBLSetting.xml with Angle & Direction ***'
                    'Get Knuckle Angle & convert into degrees
                    Dim oIElement As IXMLDOMElement
                    Dim sIBLLength As String, sIBLHeight As String
                   
                    
                    '*** Set IBL LANDING_STRAIGHT_LENGTH ***'
                    Set oIElement = oIBLDoc.selectSingleNode(".//SMS_MARKING")
                    vTemp = oIElement.GetAttribute("LANDING_STRAIGHT_LENGTH")
                    If (isNumericUS(vTemp)) Then
                        sIBLLength = Trim(Str(Round(Val(vTemp) * 1000))) ' Convert into mm
                    End If
                    
                    Set oIBLElem = oIBLSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||LANDING_STRAIGHT_LENGTH']")
                    oIBLElem.setAttribute "VALUE", sIBLLength
                    '*****************************'
                    
                    '*** Set IBL MARK_HEIGHT ***'
                    Set oIElement = oIBLDoc.selectSingleNode(".//SMS_MARKING")
                    vTemp = oIElement.GetAttribute("MARK_HEIGHT")
                    If (isNumericUS(vTemp)) Then
                        sIBLHeight = Trim(Str(Round(Val(vTemp) * 1000))) ' Convert into mm
                    End If
                    
                    Set oIBLElem = oIBLSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||MARK_HEIGHT']")
                    oIBLElem.setAttribute "VALUE", sIBLHeight
                    '*****************************'
                    Set oIElement = Nothing
                    
                    dAlong = dLength / 2
                    Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, dAlong)
                    Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
                    
                    dTanX = oVector.x
                    oVector.x = -oVector.y
                    oVector.y = dTanX
                    
                    oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
                    
                    bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, oIBLSetting.xml)

                End If
            End If
            '***************************************************************************'

        Case STRMFG_FITTING_MARK, STRMFG_PROFILE_TO_PLATE_MARK
            
            If ObjectType = TEMPLATE_TYPE Then
                
                Set oMarkingDoc = New DOMDocument
                If Not oMarkingDoc.loadXML(strInputXML) Then GoTo CleanUp
                Set oElement = oMarkingDoc.selectSingleNode("//SMS_EDGE")
                If oElement Is Nothing Then GoTo CleanUp
                
                Set oNode = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='SeamControl2']")
                Set oNode = UpdateAnnotationPropertyNodes(m_oSettingsXML, oNode, sPartType)
                strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='SeamControl2']/@PROGID").Text
                                
                Set oOutputAnnotation = SP3DCreateObject(strProgID)
                
                Set oTextSizeElem = oNode.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
                oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
                
                oOutputAnnotation.SetArguments oNode.xml
                
                oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
                Set oStartPos = New DPosition
                Set oEndPos = New DPosition
                
                oStartPos.Set dTempX, dTempY, dTempZ
                oEndPos.Set dCurveEndX, dCurveEndY, dTanZ
                dLength = oCurve.length
                                    
                dAlong = dLength / 2.2
                Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, dAlong)
                Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
                
                oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
                
                sOutputXML = GetAnnotationXML(strInputXML, oNode.xml)
                
                bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, sOutputXML)
            Else
            
                '**************************** MARGIN ************************************************'
                If oEdgeDoc.loadXML(strInputXML) Then GoTo CleanUp
                
                If UCase(strReference) = "BRACKETLOCATION" Or UCase(strReference) = "LAP" Then
                    
                    Set oMarginSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='BracketStiffener1']")
                    Set oMarginSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oMarginSetting, sPartType)
                    
                    vTemp = oMarginSetting.GetAttribute("PROGID")
                    strProgID = IIf(VarType(vTemp) = vbString, vTemp, "")
                    
                    Set oOutputAnnotation = SP3DCreateObject(strProgID)
                    
                    Set oTextSizeElem = oMarginSetting.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
                    oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
                    
                    Set oMarginElem = oEdgeDoc.selectSingleNode("//*[@GEOM2D_GUID]")
                    sGUID = oMarginElem.GetAttribute("GEOM2D_GUID")
    
                    oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
                    Set oStartPos = New DPosition
                    Set oEndPos = New DPosition
                    
                    oStartPos.Set dTempX, dTempY, dTempZ
                    oEndPos.Set dCurveEndX, dCurveEndY, dTanZ
                    dLength = oCurve.length
                    
                    If UCase(strReference) = "BRACKETLOCATION" Then
                        dAlong = 3 * dLength / 4
                    Else
                        dAlong = 2 * dLength / 3
                    End If
    
                    
                    Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, dAlong)
                    Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
                    
                    oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
                    
                    sOutputXML = GetAnnotationXML(strInputXML, oMarginSetting.xml)
    
                    oOutputAnnotation.SetArguments sOutputXML
                    
                    oVector.length = 1
                    
                    'bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, strInputXML)
                    bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, sOutputXML)
    
SkipMark:
                    Set oEdgeDoc = Nothing
                    Set oMarginElem = Nothing
                
                Else
                   
                    Set oElement = oEdgeDoc.selectSingleNode("//SMS_EDGE")
                    If oElement Is Nothing Then GoTo CleanUp
                    strThickDir = oElement.GetAttribute("PART_SIDE")
                    
                    On Error GoTo ErrorHandler
                    
                    ' thickness direction "C" indicates that toshi mark is needed on the ER fitting mark
                    If strThickDir = "C" Then
                    
                        Set oMarginSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='TOSHI_MARK_1']")
                        Set oMarginSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oMarginSetting, sPartType)
                        
                        Set oElem = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='TOSHI_MARK_1']")
                        Set oElem = UpdateAnnotationPropertyNodes(m_oSettingsXML, oElem, sPartType)
                        
                        strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='TOSHI_MARK_1']/@PROGID").Text
                        
                    ElseIf UCase(strReference) = "OVERLAP" Then
                        Set oMarginSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='END_FITTING_MARK_1']")
                        Set oMarginSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oMarginSetting, sPartType)
                        
                        vTemp = oMarginSetting.GetAttribute("PROGID")
                        strProgID = IIf(VarType(vTemp) = vbString, vTemp, "")
                                            
                        Set oMarginElem = oEdgeDoc.selectSingleNode("//*[@GEOM2D_GUID]")
                        sGUID = oMarginElem.GetAttribute("GEOM2D_GUID")
                        
                        Set oMarginElem = oEdgeDoc.selectSingleNode("//*[@MARKED_SIDE]")
                        sMarkedSide = oMarginElem.GetAttribute("MARKED_SIDE")
    
    
    '                                               ANTI_MARKING_CHAR will be added to Marking Name/EntityName Annotation
    '                        If UCase(sMarkedSide) <> "MARKING" Then
    '                            Set oMarginSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='END_FITTING_MARK_2']")
    '                            Set oMarginSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oMarginSetting, sPartType)
    '                            vTemp = oMarginSetting.getAttribute("PROGID")
    '                            strProgID = IIf(VarType(vTemp) = vbString, vTemp, "")
    '                        End If
                    Else
                        Exit Sub
                    End If
                    
                    Set oOutputAnnotation = SP3DCreateObject(strProgID)
                    
                    Set oTextSizeElem = oMarginSetting.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
                    oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
                    
                    oOutputAnnotation.SetArguments sOutputXML
                    
                    'Set Annotation's LENGTH attribute
                    '  According to the LENGTH value, the annotation service will place
                    '     the Big/Small Margin Mark
                    
                    oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
                    Set oStartPos = New DPosition
                    Set oEndPos = New DPosition
                    
                    oStartPos.Set dTempX, dTempY, dTempZ
                    dAlong = oCurve.length / 2
    
                    Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, dAlong)
                    Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
                    
                    sOutputXML = GetAnnotationXML(oEdgeDoc.xml, oMarginSetting.xml)
    
                    oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
                    
                    oVector.length = 1
                    'bstrOutputXML = bstrOutputXML & vbNewLine _
                                    & Evaluate(oCurve, oPosition, strProgID, oMarginSetting, sGUID)
                    bstrOutputXML = bstrOutputXML & vbNewLine & oOutputAnnotation.Evaluate(oPosition, oVector, sOutputXML)
                End If
            End If
SkipFittingMark:
            Set oEdgeDoc = Nothing
            Set oMarginElem = Nothing
            '************************************************************************************'
            
        Case STRMFG_BASELINE_MARK
        '************************************ BASE LINE MARK (OR  SIGHTLINE MARK)*********************************
            
            Set oMarkingDoc = New DOMDocument

            If Not oMarkingDoc.loadXML(strInputXML) Then GoTo CleanUp
            Set oElement = oMarkingDoc.selectSingleNode("//SMS_EDGE")
            If oElement Is Nothing Then GoTo CleanUp
            
            Set oNode = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='SightLineMark']")
            Set oNode = UpdateAnnotationPropertyNodes(m_oSettingsXML, oNode, sPartType)
            strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='SightLineMark']/@PROGID").Text
                       
            Set oOutputAnnotation = SP3DCreateObject(strProgID)
            
            Set oTextSizeElem = oNode.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
            oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
            
            oOutputAnnotation.SetArguments oNode.xml
            
            oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
            Set oStartPos = New DPosition
            Set oEndPos = New DPosition
                        
            oStartPos.Set dTempX, dTempY, dTempZ
            oEndPos.Set dCurveEndX, dCurveEndY, dTanZ
            
            dLength = oCurve.length
            
            dAlong = dLength / 2
            Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, dAlong)
            Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
            
            oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
            
            sOutputXML = GetAnnotationXML(strInputXML, oNode.xml)
            
            bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, sOutputXML)
            
            
        Case STRMFG_ROLL_BOUNDARIES_MARK
            
             '**************************** ROLL BOUNDARY MARK Annotation******************************************'
            Set oMarkingDoc = New DOMDocument

            If Not oMarkingDoc.loadXML(strInputXML) Then GoTo CleanUp
            
            Set oRollboundarySetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='ROLL_BOUNDARY']")
            Set oRollboundarySetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oRollboundarySetting, sPartType)
            strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='ROLL_BOUNDARY']/@PROGID").Text
            Set oOutputAnnotation = SP3DCreateObject(strProgID)
            
            Set oTextSizeElem = oRollboundarySetting.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
            oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
            
            'Added below code for adding the Guid to the Part XML
            Set oMarkingNameElem = oMarkingDoc.selectSingleNode(".//SMS_EDGE")
            strGUID = oMarkingNameElem.GetAttribute("GEOM2D_GUID")
            
            Set oTextSizeElem = oRollboundarySetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||GEOM2D_GUID']")
            oTextSizeElem.setAttribute "VALUE", strGUID
             'End of Code
                        
            oOutputAnnotation.SetArguments oRollboundarySetting.xml
        
            '*************** Get Position & Vector ****************'
            oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
            Set oStartPos = New DPosition
            Set oEndPos = New DPosition
            
             oStartPos.Set dTempX, dTempY, dTempZ
             oEndPos.Set dCurveEndX, dCurveEndY, dTanZ
            
             dLength = oCurve.length
         
            dAlong = dLength / 2
            Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, dAlong)
            Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
             
             oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
             
             sOutputXML = GetAnnotationXML(strInputXML, oRollboundarySetting.xml)
            
             '*** Adjust direction for "RE" text ***'
             Dim sREND_DIR As String
             sREND_DIR = GetAttributeValueFromXML(sOutputXML, "REND_BEND")
             
             Select Case sREND_DIR
                 Case "R"
                         oVector.length = oVector.length * -1#
                 Case "L", "B", "I"
                 
                 Case "C"
                     GoTo CleanUp
                 Case Else
             End Select
             
             '*** Add Upside and Radius for Ref Mark ***'
             Dim sRollText As String
             Dim sRollUpside As String
             Dim dRollRadius As Double
             Dim dREND_ANGLE As Double
             
             sRollUpside = GetAttributeValueFromXML(sOutputXML, "REND_DIR")
             If UCase(sRollUpside) = "UP" Then
                 sRollUpside = MARKING_CHAR
             Else
                 sRollUpside = ANTI_MARKING_CHAR
             End If
             dREND_ANGLE = GetAttributeValueFromXML(sOutputXML, "REND_ANGLE")
             dRollRadius = GetAttributeValueFromXML(sOutputXML, "REND_R")
             dRollRadius = Format(dRollRadius, 0)
             
             If UCase(strReference) = "ROLLBOUNDARY" Then
                If dRollRadius < MAX_ROLL_RADIUS Then
                    sRollText = sRollUpside & dRollRadius & "R"
                Else
                    sRollText = sRollUpside & "*****" & "R"
                End If
            Else
                sRollText = sRollUpside & dRollRadius & "R"
            End If
            
            Dim oTempElement2 As IXMLDOMElement
'            Set oMarkingDoc = New DOMDocument

'            If Not oMarkingDoc.loadXML(sOutputXML) Then GoTo CleanUp
'            Set oRollboundarySetting = oMarkingDoc.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='ROLL_BOUNDARY']")
'            Set oRollboundarySetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oRollboundarySetting, sPartType)
            Set oTempElement2 = oRollboundarySetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||ROLL_LINE_TEXT']")
                 
             If UCase(strReference) = "ROLLBOUNDARY" Then
                oTempElement2.setAttribute "VALUE", sRollText
             Else
                oTempElement2.setAttribute "VALUE", "RE"
             End If
 
            sOutputXML = GetAnnotationXML(strInputXML, oRollboundarySetting.xml)
             
             bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, sOutputXML)
             '************************************************************************************'

        Case STRMFG_ROLL_LINES_MARK
            ''**************************** ROLL LINE ******************************************'
            
            Dim oRollerMark As MarkingInfo
            Set oRollerMark = pMfgGeom2d.SystemMark
        
            Dim oDetaiedPart As IJPlatePart
            Dim oMfgObj As IJMfgPlatePart
            Set oMfgObj = pdispPart
                
            oMfgObj.GetDetailedPart oDetaiedPart
       
            Set oMarkingDoc = New DOMDocument

            If Not oMarkingDoc.loadXML(strInputXML) Then GoTo CleanUp
            
            'No annotation for Trimmed Roll Lines Mark
            If UCase(strReference) = "TRIMMED" Then
                Exit Sub
            End If
            
            Set oElement = oMarkingDoc.selectSingleNode("//SMS_EDGE")
            If oElement Is Nothing Then GoTo CleanUp
            
            Set oRollLineSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='RollDirMark']")
            Set oRollLineSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oRollLineSetting, sPartType)
            strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='RollDirMark']/@PROGID").Text
            
            '*** Set Roller Annotation Type ***'
            Set oRollerElem = oRollLineSetting.selectSingleNode(".//SMS_GEOM_ARG[@NAME='RollerCondition']")
            
            Dim dRollDistance As Double
            'oRollerMark.thickness here means Sweep Angle
            dRollDistance = oRollerMark.Radius * 1000 * (oRollerMark.thickness) 'in mm
                
            '*** Conditions to place Roller Annotations ***'
            '"1": Places Roller Arrow Mark
            '"2": Places Roll Text parallel to Roll Line Mark
            '"3": Places Roll Text Normal to Roll Line Mark
            '"4": Places Roll Text + Knuckle Text parallel to Roll Line Mark
            
            'CHECK 1 - Roll line coming from Knuckle

            If UCase(strReference) = "KNUCKLE" Then
                If IsLBH(oDetaiedPart) Or _
                    IsBilgeHopper(oDetaiedPart) Or _
                    IsShoulderTank(oDetaiedPart) Or _
                    IsCorrugated(oDetaiedPart) Or _
                    IsCladPlate(oDetaiedPart) Then
                    oRollerElem.setAttribute "VALUE", "4"
                Else
                    oRollerElem.setAttribute "VALUE", "2"
                End If
            Else
                If pMfgGeom2d.IsSupportOnly And IsCylidrical(oDetaiedPart) Then
                    oRollerElem.setAttribute "VALUE", "1"
                ElseIf oRollerMark.Radius * 1000 < ROLLER_RADIUS_CONDITION Then
                    oRollerElem.setAttribute "VALUE", "4"
                ElseIf oRollerMark.Radius * 1000 > ROLLER_RADIUS_CONDITION Then
                    If dRollDistance <= ROLLER_DIST_CONDITION Then
                    oRollerElem.setAttribute "VALUE", "2"
                ElseIf dRollDistance > ROLLER_DIST_CONDITION Then
                    oRollerElem.setAttribute "VALUE", "3"
                End If
                
                Else 'Default
                    oRollerElem.setAttribute "VALUE", "2"
                End If
            End If
            '**********************************************'
                
            Set oOutputAnnotation = SP3DCreateObject(strProgID)
            
            Set oRollerElem = oRollLineSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||UPSIDE']")
            If UCase(oRollerMark.Direction) = "UP" Then sDir = MARKING_CHAR Else sDir = ANTI_MARKING_CHAR
            oRollerElem.setAttribute "VALUE", sDir
            
            'ROLL_RADIUS should be INNER RADIUS.
            Dim oRElem As IXMLDOMElement
            Dim dInnerRadius As Double
            If Not oMarkingDoc.selectSingleNode("//SMS_MARKING[@INNER_RADIUS]") Is Nothing Then
                Set oRElem = oMarkingDoc.selectSingleNode("//SMS_MARKING[@INNER_RADIUS]")
                dInnerRadius = oRElem.GetAttribute("INNER_RADIUS")
            End If
            
             Dim oRollEndInfo As IXMLDOMElement
            Set oRollEndInfo = oMarkingDoc.selectSingleNode("//SMS_MARKING_REND_INFO")
            
            Set oRollerElem = oRollLineSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||ROLL_RADIUS']")
            oRollerElem.setAttribute "VALUE", Str(dInnerRadius * 1000) 'Convert to mmm
            
            Set oRollerElem = oRollLineSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||SWEEP_ANGLE']")
            Dim dThickVal As Double
            If oRollEndInfo Is Nothing Then
                oRollerElem.setAttribute "VALUE", 0
            Else
                dThickVal = oRollEndInfo.GetAttribute("REND_SANGLE")
                oRollerElem.setAttribute "VALUE", dThickVal
            End If
            
            Set oRollerElem = oRollLineSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||KNUCKLE_ANGLE']")
            Dim dFitAngle As Double
            If oRollEndInfo Is Nothing Then
                oRollerElem.setAttribute "VALUE", 0
            Else
                dFitAngle = oRollEndInfo.GetAttribute("REND_ANGLE")
                oRollerElem.setAttribute "VALUE", dFitAngle
            End If
            
            oOutputAnnotation.SetArguments oRollLineSetting.xml
                
            oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
            Set oStartPos = New DPosition
            Set oEndPos = New DPosition
            
            oStartPos.Set dTempX, dTempY, dTempZ
            oEndPos.Set dCurveEndX, dCurveEndY, dTanZ
            dLength = oCurve.length
            
            '*** Put annotation at center ***'
            dAlong = dLength / 2
            Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, dAlong)
            Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
            
            oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0

            sOutputXML = GetAnnotationXML(strInputXML, oRollLineSetting.xml)
            
            'bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, strInputXML)
            bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, sOutputXML)

            '************************************************************************************'

        Case STRMFG_MOUNT_ANGLE_MARK

            Set oMarkingDoc = New DOMDocument
            If Not oMarkingDoc.loadXML(strInputXML) Then GoTo CleanUp
            
            Set oElem = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='DeclivityMark']")
            Set oElem = UpdateAnnotationPropertyNodes(m_oSettingsXML, oElem, sPartType)
            strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='DeclivityMark']/@PROGID").Text
            Set oOutputAnnotation = SP3DCreateObject(strProgID)
            
            Set oTextSizeElem = oElem.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
            oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
            
            oOutputAnnotation.SetArguments oElem.xml
            
            Set oMarkingNameElem = oMarkingDoc.selectSingleNode("//SMS_MARKING[@TYPE='strmfg_mount_angle_mark']")
           
            sMarkingName = oMarkingNameElem.GetAttribute("MARKSIDE")
            Set oDeclElement = oElem.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||UPSIDE']")

            Dim sDeclDirection                      As String
            If sMarkingName = "anti_marking" Then
                sDeclDirection = ANTI_MARKING_CHAR
            Else
                sDeclDirection = ""
            End If
            oDeclElement.setAttribute "VALUE", sDeclDirection
            
            Dim sDeclAngle                          As String
            Dim dDeclAngle                          As Double
            Dim sGlobalDir                          As String
            
            sDeclAngle = oMarkingNameElem.GetAttribute("DECLIVITY")
            dDeclAngle = Val(sDeclAngle) * 180 / PI
            dDeclAngle = Format(dDeclAngle, "0.#")
            
            sGlobalDir = oMarkingNameElem.GetAttribute("MOUNTINGDIRECTION")
            
            Set oDeclElement = oElem.selectSingleNode(".//PROPERTY[@NAME = 'SMS_ANNOTATION||ANGLE']")
            oDeclElement.setAttribute "VALUE", Trim(Str(dDeclAngle))
            
            Dim oDeclCS                             As IJComplexString
            Set oDeclCS = pMfgGeom2d.GetGeometry
            
            'Checking Whether there are 2 Declivity Marks for Annotation
            If oDeclCS.CurveCount = 2 Then
            
                'A\   |D  /C
                '  \  |  /
                '   \ | /
                '    \|/
                '     B
                '
                ' AB and BC are the Declivity Mark Lines
                ' AB and BC also represent the Vectors alng the Marking Lines
                ' BD is the Internal Angular Bisector Vector of angle ABC
            
                Dim oDeclCurve1                     As IJCurve
                Dim oDeclCurve2                     As IJCurve
                
                oDeclCS.GetCurve 1, oDeclCurve1
                oDeclCS.GetCurve 2, oDeclCurve2
                
                Dim dXstart As Double, dYstart As Double, dZstart As Double, dXend As Double, dYend As Double, dZend As Double
                Dim dXstart2 As Double, dYstart2 As Double, dZstart2 As Double, dXend2 As Double, dYend2 As Double, dZend2 As Double
    
                oDeclCurve1.EndPoints dXstart, dYstart, dZstart, dXend, dYend, dZend
                oDeclCurve2.EndPoints dXstart2, dYstart2, dZstart2, dXend2, dYend2, dZend2
    
                Dim oStartPosition1                 As IJDPosition
                Dim oEndPosition1                   As IJDPosition
                Dim oStartPosition2                 As IJDPosition
                Dim oEndPosition2                   As IJDPosition
                Dim oOtherPoint1                    As IJDPosition
                Dim oOtherPoint2                    As IJDPosition
                Dim oIntPoint                       As IJDPosition
                
                Set oStartPosition1 = New DPosition
                Set oEndPosition1 = New DPosition
                Set oStartPosition2 = New DPosition
                Set oEndPosition2 = New DPosition
    
                oStartPosition1.Set dXstart, dYstart, dZstart
                oEndPosition1.Set dXend, dYend, dZend
                
                oStartPosition2.Set dXstart2, dYstart2, dZstart2
                oEndPosition2.Set dXend2, dYend2, dZend2
                
                If oStartPosition1.DistPt(oStartPosition2) < 0.001 Then
                    Set oIntPoint = oStartPosition1
                    Set oOtherPoint1 = oEndPosition1
                    Set oOtherPoint2 = oEndPosition2
                ElseIf oStartPosition1.DistPt(oEndPosition2) < 0.001 Then
                    Set oIntPoint = oStartPosition1
                    Set oOtherPoint1 = oEndPosition1
                    Set oOtherPoint2 = oStartPosition2
                ElseIf oEndPosition1.DistPt(oStartPosition2) < 0.001 Then
                    Set oIntPoint = oStartPosition2
                    Set oOtherPoint1 = oStartPosition1
                    Set oOtherPoint2 = oEndPosition2
                Else
                    Set oIntPoint = oEndPosition1
                    Set oOtherPoint1 = oStartPosition1
                    Set oOtherPoint2 = oStartPosition2
                End If
                
                Dim oMidPoint                       As IJDPosition
                Dim oVectorAlongCurve1              As IJDVector
                Dim oVectorAlongCurve2              As IJDVector
                Dim oDeclVector                     As IJDVector
                
                Set oMidPoint = New DPosition
                Set oVectorAlongCurve1 = oOtherPoint1.Subtract(oIntPoint)
                Set oVectorAlongCurve2 = oOtherPoint2.Subtract(oIntPoint)
                
                'Normalizing the Vector
                oVectorAlongCurve1.length = 1
                oVectorAlongCurve2.length = 1
                
                'Finding the Internal Angular Bisector of the Declivity Marks
                Set oDeclVector = oVectorAlongCurve1.Add(oVectorAlongCurve2)
                
                'Normalize the Vector
                oDeclVector.length = 1
                
                oMidPoint.x = oIntPoint.x + (DECLIVITY_MARK_LENGTH * Cos(Val(sDeclAngle) / 2) * oDeclVector.x)
                oMidPoint.y = oIntPoint.y + (DECLIVITY_MARK_LENGTH * Cos(Val(sDeclAngle) / 2) * oDeclVector.y)
                oMidPoint.z = oIntPoint.z + (DECLIVITY_MARK_LENGTH * Cos(Val(sDeclAngle) / 2) * oDeclVector.z)
                
                oMidPoint.x = oMidPoint.x * 1000
                oMidPoint.y = oMidPoint.y * 1000
                oMidPoint.z = oMidPoint.z * 1000
                                
                
                'oDeclVector.x
                dTanX = oDeclVector.x
                oDeclVector.x = oDeclVector.y
                oDeclVector.y = -dTanX
                                
                sOutputXML = GetAnnotationXML(strInputXML, oElem.xml)
                bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oMidPoint, oDeclVector, sOutputXML)
            End If
            
        Case STRMFG_COLLARPLATELOCATION_MARK
            'Fix for adjusting the Annotation of Collar Location Mark
            Set oMarkingDoc = New DOMDocument
            'Collecting the Information in strInputXML into oMarkingDoc
            If Not oMarkingDoc.loadXML(strInputXML) Then GoTo CleanUp
            
            Set oElement = oMarkingDoc.selectSingleNode("//SMS_EDGE")
            If oElement Is Nothing Then GoTo CleanUp
            
            strThickDir = oElement.GetAttribute("PART_SIDE")

            
            Dim oCollarPlateSetting As IXMLDOMElement
            Dim sTextElem()             As String
            Dim sWebLength              As String       'Stores the Web Length of Profile
            Dim sSlotType               As String       'Stores the Slot Type of the Collar
            
            'Collecting the information in SMS_ANNOTATION.xml in oCollarPlateSetting element
            Set oCollarPlateSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='CollarPlateInfo']")
            Set oCollarPlateSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oCollarPlateSetting, sPartType)
            strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='CollarPlateInfo']/@PROGID").Text
            Set oOutputAnnotation = SP3DCreateObject(strProgID)
            
            Set oTextSizeElem = oCollarPlateSetting.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
            oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
            
            Set oCollElement = oCollarPlateSetting.selectSingleNode(".//SMS_GEOM_ARG[@NAME = 'ControlPoint']")

            If UCase(strThickDir) = "L" Then
                oCollElement.setAttribute "VALUE", "ll"
            Else
                oCollElement.setAttribute "VALUE", "lr"
            End If
            
            oOutputAnnotation.SetArguments oCollarPlateSetting.xml
            
            Set oMarkingNameElem = oMarkingDoc.selectSingleNode("//SMS_MARKING[@TYPE='strmfg_collarplatelocation_mark']")
            sSlotType = oMarkingNameElem.GetAttribute("MARKING_NAME")
            
            If sSlotType = vbNullString Then Exit Sub
            sTextElem = VBA.Split(Trim(sSlotType), " ")
            
            'As the Last Element of sTextElem Array is th Web Length
            sWebLength = sTextElem(UBound(sTextElem))
            
            'Removing the Web Length part of string from sSlotType
            sSlotType = Trim(Replace(sSlotType, sWebLength, ""))
            'Replacing the Values in SMS_ANNOTATION.xml
            Set oCollElement = oCollarPlateSetting.selectSingleNode(".//PROPERTY[@NAME = 'SMS_ANNOTATION||SLOT_TYPE']")
            oCollElement.setAttribute "VALUE", sSlotType
            
            Set oCollElement = oCollarPlateSetting.selectSingleNode(".//PROPERTY[@NAME = 'SMS_ANNOTATION||WEB_LENGTH']")
            oCollElement.setAttribute "VALUE", sWebLength
            
            'Getting the Position and the vector for placing the Annotation
            oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
            Set oStartPos = New DPosition
                                    
            oStartPos.Set dTempX, dTempY, dTempZ
            Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oStartPos)
            
            'Rotating the Vector by 90 degrees clockwise as the Annotation need to be Perpendicular to the Mark
            dTanX = oVector.x
            oVector.x = oVector.y
            oVector.y = -dTanX
            
            'For Adjusting the Annotation position
            If UCase(strThickDir) = "L" Then
                oVector.length = -1
            End If

            oStartPos.Set oStartPos.x * 1000, oStartPos.y * 1000, 0
            
            sOutputXML = GetAnnotationXML(strInputXML, oCollarPlateSetting.xml)
            bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oStartPos, oVector, sOutputXML)
            
        Case STRMFG_CORRUGATE_MARK
                
            Set oMarkingDoc = New DOMDocument
            
            If Not oMarkingDoc.loadXML(strInputXML) Then GoTo CleanUp
            Set oElement = oMarkingDoc.selectSingleNode("//SMS_EDGE")
            If oElement Is Nothing Then GoTo CleanUp
                
            Dim sBump As String
            
            Set oCorrugateElement = oMarkingDoc.selectSingleNode("//SMS_MARKING")
            If oCorrugateElement Is Nothing Then GoTo CleanUp
            
            If Not oMarkingDoc.selectSingleNode("//SMS_MARKING[@USAGE]") Is Nothing Then
            sUsage = oCorrugateElement.GetAttribute("USAGE")
            sUsage = UCase(sUsage)
            End If
            
            If sUsage = "MAIN" Then
                '..................................... _
                'STRMFG_REFERENCE_CENTER_MARK - Put "_/ \_" Annotation
            
                Set oCorrSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='CorrugateMark']")
                Set oCorrSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oCorrSetting, sPartType)
                strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='CorrugateMark']/@PROGID").Text
            
                Set oOutputAnnotation = SP3DCreateObject(strProgID)
                
                Set oTextSizeElem = oCorrSetting.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
                oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
                
                sBump = oCorrugateElement.GetAttribute("BUMP")
                sBump = UCase(sBump)
                
                If sBump = "UP" Then
                    dMarkedSide = MARKING_CHAR
                ElseIf sBump = "DOWN" Then
                    dMarkedSide = ANTI_MARKING_CHAR
                End If
                
                Set oCorrElem = oCorrSetting.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||PREFIX']")
                oCorrElem.setAttribute "VALUE", Str(dMarkedSide)
                
                oOutputAnnotation.SetArguments oCorrSetting.xml
                
                oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
                Set oStartPos = New DPosition
                Set oEndPos = New DPosition
                
                oStartPos.Set dTempX, dTempY, dTempZ
                oEndPos.Set dCurveEndX, dCurveEndY, dTanZ
                dLength = oCurve.length
                                   
                dAlong = dLength / 2
                Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, dAlong)
                Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
                
                oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
                
                sOutputXML = GetAnnotationXML(strInputXML, oCorrSetting.xml)
                
                bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, sOutputXML)
                            
            ElseIf sUsage = "REF" Then
                
                'REF_MARK - Put "X" Annotation
                Set oCorrSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='CorrugateMark2']")
                Set oCorrSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oCorrSetting, sPartType)
                strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='CorrugateMark2']/@PROGID").Text
                
                Set oOutputAnnotation = SP3DCreateObject(strProgID)
            
                Set oTextSizeElem = oCorrSetting.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
                oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
                
                oOutputAnnotation.SetArguments oCorrSetting.xml
                
                oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
                Set oStartPos = New DPosition
                Set oEndPos = New DPosition
                
                oStartPos.Set dTempX, dTempY, dTempZ
                oEndPos.Set dCurveEndX, dCurveEndY, dTanZ
                dLength = oCurve.length
                   
                dAlong = dLength / 2
                Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, dAlong)
                Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
                
                'Rotating the Vector by 90 degrees clockwise as the Annotation need to be Perpendicular to the Mark
                dTanX = oVector.x
                oVector.x = oVector.y
                oVector.y = -dTanX
                
                oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
                
                sOutputXML = GetAnnotationXML(strInputXML, oCorrSetting.xml)
                
                bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, sOutputXML)
            
            End If
        
        Case STRMFG_REF_MARK
            'For Corrugate Dimensions Marks
            '*** Get Entity Name Annotation XML ***'
            
            Dim sCorrLabel As String
            Dim dLabelValue As Double
            
            Set oMarkingDoc = New DOMDocument
            If Not oMarkingDoc.loadXML(strInputXML) Then GoTo CleanUp
                        
            Set oCorrugateElement = oMarkingDoc.selectSingleNode("//SMS_MARKING")
            If oCorrugateElement Is Nothing Then GoTo CleanUp
            
            sUsage = oCorrugateElement.GetAttribute("USAGE")
            
            If UCase(sUsage) = "DIM" Then
                                    
                Set oEntityNameElem = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='ENTITYNAME']")
                Set oEntityNameElem = UpdateAnnotationPropertyNodes(m_oSettingsXML, oEntityNameElem, sPartType)
                strProgID = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='ENTITYNAME']/@PROGID").Text
                Set oOutputAnnotation = SP3DCreateObject(strProgID)
                
                sCorrLabel = oCorrugateElement.GetAttribute("LABEL")
                dLabelValue = Val(sCorrLabel)
                            
                Set oCorrElem = oEntityNameElem.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||MARKING_NAME']")
                oCorrElem.setAttribute "VALUE", Str(dLabelValue)
                
                Set oTextSizeElem = oEntityNameElem.selectSingleNode(".//SMS_GEOM_ARG[@NAME='TextSize']")
                oTextSizeElem.setAttribute "VALUE", Str(dTextSize)
                oOutputAnnotation.SetArguments oEntityNameElem.xml
                
                oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
                Set oStartPos = New DPosition
                Set oEndPos = New DPosition
                
                oStartPos.Set dTempX, dTempY, dTempZ
                oEndPos.Set dCurveEndX, dCurveEndY, dTanZ
                dLength = oCurve.length
                   
                dAlong = dLength / 2
                Set oPosition = oGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, dAlong)
                Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
                  
                oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0
                
                sOutputXML = GetAnnotationXML(strInputXML, oEntityNameElem.xml)
                
                bstrOutputXML = bstrOutputXML & oOutputAnnotation.Evaluate(oPosition, oVector, sOutputXML)

            End If
        
        Case Else
        
    End Select
    
CleanUp:
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number ', MODULE + METHOD
    GoTo CleanUp

End Sub
Public Function GetAttributeValueFromXML(sXML As String, sAttribute As String) As String

   
    Dim oXMLDOM As New DOMDocument
    Dim oXMLElement As IXMLDOMElement
    Dim sAttributeValue As String
    
    If Not oXMLDOM.loadXML(sXML) Then GoTo CleanUp
    
    Set oXMLElement = oXMLDOM.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||" & sAttribute & "']")
    
    If oXMLElement Is Nothing Then
        Set oXMLElement = oXMLDOM.selectSingleNode(".//SMS_GEOM_ARG[@NAME='" & sAttribute & "']")
    End If
        
    sAttributeValue = oXMLElement.GetAttribute("VALUE")
    
    GetAttributeValueFromXML = sAttributeValue
    
CleanUp:
    Set oXMLDOM = Nothing
    Set oXMLElement = Nothing

End Function

Public Function GetAnnotationXML(sInputXML As String, sSettingXML As String) As String
    
    Dim oSMSMarkingDOM As New DOMDocument
    Dim oMarkingElem As IXMLDOMElement
    Dim sGUID As String
    Dim oSettingsDOM As New DOMDocument
    Dim oTextElem As IXMLDOMElement
    Dim oSettingsElem As IXMLDOMElement
    Dim oTextArgList As IXMLDOMNodeList
    Dim oContext As IJContext
    Dim sAttribute As String
    Dim sValue As String
    Dim oAttr As IXMLDOMAttribute
    
    'For each attribute in <SMS_TEXT_ARGS>, search its value in sInputXML
    oSMSMarkingDOM.loadXML (sInputXML)
    oSettingsDOM.loadXML (sSettingXML)
    
    Set oTextArgList = oSettingsDOM.selectNodes(".//PROPERTY")
    
    For Each oTextElem In oTextArgList
        sAttribute = getAttributeName(oTextElem.GetAttribute("NAME"))
        
        Set oAttr = oSMSMarkingDOM.selectSingleNode("//@" & sAttribute)
        If Not oAttr Is Nothing Then
            Set oMarkingElem = oSMSMarkingDOM.selectSingleNode("//*[@" & sAttribute & "]")
            sValue = oMarkingElem.GetAttribute(sAttribute)
                
            Set oSettingsElem = oSettingsDOM.selectSingleNode(".//PROPERTY[@NAME='SMS_ANNOTATION||" & sAttribute & "']")
            oSettingsElem.setAttribute "VALUE", sValue
                
        End If
    Next
    
    Dim oGeomArgList
    Set oGeomArgList = oSettingsDOM.selectNodes(".//SMS_GEOM_ARG")
    
    For Each oTextElem In oGeomArgList
        sAttribute = oTextElem.GetAttribute("NAME")
        
        Set oAttr = oSMSMarkingDOM.selectSingleNode("//@" & sAttribute)
        If Not oAttr Is Nothing Then
            Set oMarkingElem = oSMSMarkingDOM.selectSingleNode("//*[@" & sAttribute & "]")
            sValue = oMarkingElem.GetAttribute(sAttribute)
                
            Set oSettingsElem = oSettingsDOM.selectSingleNode(".//SMS_GEOM_ARG[@NAME='" & sAttribute & "']")
            oSettingsElem.setAttribute "VALUE", sValue
        End If
    Next
    
    GetAnnotationXML = oSettingsDOM.xml
CleanIp:
    
End Function

Private Function getAttributeName(strIn As String) As String
    Const METHOD = "UpdateXMLWithAnnotationInfo"
    On Error GoTo ErrorHandler
    
    If InStr(strIn, "||") > 0 Then
        getAttributeName = Right(strIn, Len(strIn) - InStr(strIn, "||") - 1)
    Else
        getAttributeName = strIn
    End If
Exit Function
ErrorHandler:
    Resume Next
End Function

Private Function Evaluate(oCurve As IJCurve, oPosition As IJDPosition, _
                            sProgId As String, oAnnotElem As IXMLDOMElement, sGUID As String) As String
     
    Dim dTanX               As Double, dTanY As Double, dTanZ As Double
    Dim dLocX               As Double, dLocY As Double, dLocZ As Double
    Dim dTempX              As Double, dTempY As Double, dTempZ As Double
    Dim dTempLocX           As Double, dTempLocY As Double, dTempLocZ As Double
    Dim StartX              As Double, StartY As Double
    Dim endX                As Double
    Dim endY                As Double
    Dim dCurveEndX          As Double
    Dim dCurveEndY          As Double
    Dim dStartParam         As Double
    Dim dMidParam           As Double
    Dim dEndParam           As Double
    Dim oElements           As IJElements
    Dim oXMLDOM             As New DOMDocument
    Dim oXMLElem            As IXMLDOMElement
    
    Dim oStartPos           As New DPosition
    Dim oEndPos             As New DPosition
    
    Dim oGeomHelper         As IJMfgGeomHelper
    Set oGeomHelper = New MfgGeomHelper
    Dim oVector As New DVector
    
    Dim bstrOutputXML As String

    oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
    
    oStartPos.Set StartX, StartY, 0
                        
    '*** Create Annotation Object ***'
    Dim oOutputAnnotation As IJDMfgOutputAnnotation
    Set oOutputAnnotation = SP3DCreateObject(sProgId)
    If Not oOutputAnnotation Is Nothing Then
        'Set GUID
        If oXMLDOM.loadXML(oAnnotElem.xml) Then
            Set oXMLElem = oXMLDOM.selectSingleNode("//*[@NAME='SMS_ANNOTATION||GEOM2D_GUID']")
            If Not oXMLElem Is Nothing Then
                oXMLElem.setAttribute "VALUE", sGUID
            End If
        End If
        
        oOutputAnnotation.SetArguments oXMLDOM.xml
    Else
        GoTo CleanUp
    End If
     
    '*** Set Vector ***'
    Set oVector = oGeomHelper.GetTangentByPointOnCurve(oCurve, oPosition)
    oVector.length = 1
    
    '*** Set Position ***'
    oPosition.Set oPosition.x * 1000, oPosition.y * 1000, 0

    '*** Call Evaluate Method ***'
    bstrOutputXML = oOutputAnnotation.Evaluate(oPosition, oVector, oXMLDOM.xml)
     
    Evaluate = bstrOutputXML

CleanUp:
    Set oOutputAnnotation = Nothing
    Set oXMLElem = Nothing
    Set oXMLDOM = Nothing
    
 
Exit Function


End Function

Private Sub IJMfgTextMarkService_CreateBevelAnnotation(ByVal pdispPart As Object, ByVal pMfgGeom2d As GSCADStrMfgUtilities.IJMfgGeom2d, ByVal strInputXML As String, ByVal pTransformationMatrix As AutoMath.IJDT4x4, bstrOutputXML As String)
    Const METHOD = "IJMfgTextMarkService_CreateBevelAnnotation"
    On Error GoTo ErrorHandler

'*****************************************************************************************
        'BEVEL
        Dim oTempAnnotationDoc  As IXMLDOMDocument
        Dim oContourDoc         As IXMLDOMDocument
        Set oContourDoc = New DOMDocument
        Dim oBevelSetting       As IXMLDOMElement
        Dim oGrindSetting       As IXMLDOMElement
        Dim oElement            As IXMLDOMElement
        Dim oPointList          As IXMLDOMNodeList
        Dim oBevelList          As IXMLDOMNodeList
        Dim oBevelNode          As IXMLDOMNode
        Dim vTemp               As Variant
        Dim strProgID           As String
        Dim strGrindProgID      As String
        Dim sGUID               As String
        Dim bBevel              As Boolean
            
        Dim oGeomHelper         As IJMfgGeomHelper
        Dim oOriginalCS         As IJComplexString
        Dim oCS                 As IJComplexString
        Dim oCurve              As IJCurve
        Dim oStartPos           As IJDPosition
        Dim oEndPos             As IJDPosition
        Dim oPosition           As IJDPosition
        Dim oElements           As IJElements
        Dim oOrientation        As IJDVector
        Dim oMfgMGHelper        As New MfgMGHelper
        Set oOrientation = New DVector
        Dim oOutputAnnotation As IJDMfgOutputAnnotation
        
        If Not m_bDoesXMLExist Then Exit Sub

        Set oOriginalCS = pMfgGeom2d.GetGeometry
        oMfgMGHelper.CloneComplexString oOriginalCS, oCS
        If Not pTransformationMatrix Is Nothing Then
            oCS.Transform pTransformationMatrix
        End If
        Set oCurve = oCS
        Set oGeomHelper = New MfgGeomHelper
        Set oPosition = New DPosition
        
        Dim dLength As Double
        Dim dStartParam As Double, dEndParam As Double, dMidParam As Double
        Dim dTempLocZ As Double, dLocX As Double, dLocY As Double
        Dim dTanX As Double, dTanY As Double, dTanZ As Double
        Dim dTempX As Double, dTempY As Double, dTempZ As Double
        Dim StartX As Double, StartY As Double, endX As Double, endY As Double
        Dim dCurveEndX As Double, dCurveEndY As Double
        Dim dAlong As Double, dAlong2 As Double, dAngle As Double
            
        Set oBevelSetting = m_oSettingsXML.selectSingleNode("//SMS_OUTPUT_ANNOTATION[@TYPE='BEVEL NEW 1']")
        'Set oBevelSetting = UpdateAnnotationPropertyNodes(m_oSettingsXML, oBevelSetting, sPartType)
        vTemp = oBevelSetting.GetAttribute("PROGID")
        strProgID = IIf(VarType(vTemp) = vbString, vTemp, "")
            
        oContourDoc.loadXML "<TEMP>" & strInputXML & "</TEMP>"
        
        If Not oBevelSetting Is Nothing Then
            Set oBevelList = oContourDoc.selectNodes("//SMS_BEVEL")
            If oBevelList.length > 0 Then
                For Each oBevelNode In oBevelList
                    Set oElement = oBevelNode
                    
                    Set oPointList = oElement.getElementsByTagName("CVG_POINT")
                    Set oStartPos = New DPosition
                    Set oEndPos = New DPosition
        
                    oCurve.EndPoints dTempX, dTempY, dTempZ, dCurveEndX, dCurveEndY, dTanZ
                    StartX = Val(oPointList.Item(0).Attributes.getNamedItem("X").Text) / 1000
                    StartY = Val(oPointList.Item(0).Attributes.getNamedItem("Y").Text) / 1000
                    endX = Val(oPointList.Item(1).Attributes.getNamedItem("X").Text) / 1000
                    endY = Val(oPointList.Item(1).Attributes.getNamedItem("Y").Text) / 1000
                    
                    Set oOutputAnnotation = SP3DCreateObject(strProgID)
                    If Not oOutputAnnotation Is Nothing Then
                        oOutputAnnotation.SetArguments oBevelSetting.xml
                        bBevel = True
                        Dim dtest As Double
                        '*** Get the Mid-Point ***'
                        If Abs(dTempX - StartX) < 0.002 And Abs(dTempY - StartY) < 0.002 Then
                            oCurve.Parameter StartX, StartY, 0, dStartParam
                            oCurve.Evaluate (dStartParam) + 0.05, StartX, StartY, dtest, dTempLocZ, dTempY, dTanZ, dTempX, dTempY, dTempZ
                        End If
                        If Abs(dCurveEndX - endX) < 0.002 And Abs(dCurveEndY - endY) < 0.002 Then
                            oCurve.ParamRange dTempZ, dEndParam
                            oCurve.Parameter StartX, StartY, 0, dStartParam
                            oCurve.Evaluate (dEndParam) * 0.95, endX, endY, dtest, dTempLocZ, dTempY, dTanZ, dTempX, dTempY, dTempZ
                        End If
                        oStartPos.Set StartX, StartY, 0
                        oEndPos.Set endX, endY, 0
                        Set oElements = oGeomHelper.DivideCurveEqualSpace(oCurve, oStartPos, oEndPos, 3)
        
                        If Not oElements Is Nothing Then
                            oCurve.Parameter oElements.Item(2).x, oElements.Item(2).y, 0, dMidParam
                            oCurve.Evaluate dMidParam, dLocX, dLocY, dTempLocZ, dTanX, dTanY, dTanZ, dTempX, dTempY, dTempZ
                            oPosition.Set dLocX * 1000, dLocY * 1000, 0
                        Else: GoTo CleanUp
                        End If
                        oOrientation.Set dTanX, dTanY, dTanZ
                        oOrientation.length = 1
                        dTanZ = 0
                        
                        Set oTempAnnotationDoc = New DOMDocument
                        oTempAnnotationDoc.loadXML oOutputAnnotation.Evaluate(oPosition, oOrientation, _
                                                GetAnnotationXML(oBevelNode.xml, oBevelSetting.xml))
                        oBevelNode.parentNode.appendChild oTempAnnotationDoc.documentElement
                    End If
                Next oBevelNode
            End If
        End If
'*****************************************************************************************

    bstrOutputXML = oContourDoc.xml
    bstrOutputXML = Right(bstrOutputXML, Len(bstrOutputXML) - 6)
    bstrOutputXML = Left(bstrOutputXML, InStr(bstrOutputXML, "</TEMP>") - 1)
CleanUp:
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, MODULE & METHOD
    GoTo CleanUp
End Sub

Private Sub IJMfgTextMarkService_CreateCGVTEXTXML(ByVal pdispPart As Object, ByVal pMfgGeom2d As GSCADStrMfgUtilities.IJMfgGeom2d, ByVal pMfgSystemMark As GSCADStrMfgUtilities.IJMfgSystemMark, bRemoveGraphics As Boolean, ByVal pTransformationMatrix As AutoMath.IJDT4x4, pXMLDOMDoc As GSCADStrMfgUtilities.JCmnShp_CollectionAlias)
    Const METHOD = "IJMfgTextMarkService_CreateCGVTEXTXML"
    On Error GoTo ErrorHandler
    Set pXMLDOMDoc = New Collection
    'DI-CP-180851    Make Check drawings available with default delivery in 9.1
    
CleanUp:
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, MODULE & METHOD
    GoTo CleanUp
End Sub

Private Sub FindPositionAndAngleOnCurve(oCurve As IJCurve, dDistance As Double, dLocX As Double, dLocY As Double, dRotAngle As Double)
    Const METHOD = "FindPositionAndAngleOnCurve"
    On Error GoTo ErrorHandler
    
    Dim dStartParam As Double
    Dim dEndParam As Double
    Dim dTempLocZ As Double
    Dim dTanX As Double
    Dim dTanY As Double
    Dim dTanZ As Double
    Dim dTempX As Double
    Dim dTempY As Double
    Dim dTempZ As Double
    oCurve.ParamRange dStartParam, dEndParam
    oCurve.Evaluate dStartParam + dDistance * (dEndParam - dStartParam), dLocX, dLocY, dTempLocZ, dTanX, dTanY, dTanZ, dTempX, dTempY, dTempZ
    If Abs(dTanX) > 0.000001 Then
        dRotAngle = Atn(dTanY / dTanX)
        If dTanY >= 0 And dTanX < 0 Then
            dRotAngle = PI + dRotAngle
        ElseIf dTanY < 0 And dTanX < 0 Then
            dRotAngle = dRotAngle - PI
        End If
    Else
        If dTanY > 0 Then
            dRotAngle = PI / 2
        Else
            dRotAngle = -PI / 2
        End If
    End If
        
CleanUp:
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, MODULE & METHOD
    GoTo CleanUp
End Sub

Private Function GetActiveShipClass() As IJProjectRoot
    Const METHOD = "GetActiveShipClass"
    On Error GoTo ErrHandler

    Dim oMfgEntHelper As IJMfgEntityHelper
    Set oMfgEntHelper = New MfgEntityHelper

    Dim oActiveProject As IJProjectRoot
    Set oActiveProject = oMfgEntHelper.GetConfigProjectRoot

    Set GetActiveShipClass = oActiveProject

    Exit Function
ErrHandler:
End Function
'
' Method:FindPositionAndAngleForShipDirectionMarkText
'
' Description:
'   This method determines position and angle for ship direction mark text.
'   Ship direction mark consists of two mark lines along U and V direction respectively.
'   Get two marks
'   Find their start and end points
'   Use common point or cloeset point as their start points
'   Get each mark line's direction, from start to end
'   Adjust text location and rotation angle based on mark lines' directions
'
Private Sub FindPositionAndAngleForShipDirectionMarkText( _
        ByVal oCurrentMarkGeom2d As IJMfgGeom2d, _
        ByVal oMfgPart As Object, _
        ByVal dTextBoxWidth As Double, _
        ByVal dTextBoxHeight As Double, _
        ByRef dLocX As Double, _
        ByRef dLocY As Double, _
        ByRef dRotAngle As Double)
   Const METHOD = "FindPositionAndAngleForShipDirectionMarkText"
   On Error GoTo ErrorHandler
        
   ' Ship direction marks are paired, get other mark
   Dim oMfgParent As IJMfgGeomParent
   Dim oChildColl As IJDTargetObjectCol
   Dim oGeomCol2d As IJMfgGeomCol2d
   
   Set oMfgParent = oMfgPart
   Set oChildColl = oMfgParent.GetChildren
   Set oMfgParent = Nothing
   Set oGeomCol2d = oChildColl.Item(1)
   Set oChildColl = Nothing
   
   Dim oGeom2d As IJMfgGeom2d
   Dim oOtherMarkGeom2d As IJMfgGeom2d
   Dim oMarkingInfo As MarkingInfo
   Dim oThisMarkingInfo As MarkingInfo
   Dim oOtherMarkingInfo As MarkingInfo
   Dim nIndex As Long
   Dim strThisName As String
   Dim strOtherName As String
   
   Dim oRelationHelper As IMSRelation.DRelationHelper
   Dim oCollectionHelper As IMSRelation.DCollectionHelper
   
   For nIndex = oGeomCol2d.Getcount To 1 Step -1
      Set oGeom2d = oGeomCol2d.GetGeometry(nIndex)
      If oGeom2d.GetGeometryType = STRMFG_DIRECTION Then
         Set oRelationHelper = oGeom2d
         Set oCollectionHelper = oRelationHelper.CollectionRelations( _
                             "{E6B9C8CA-4AC2-11D5-8151-0090276F4297}", _
                             "SystemMark2dParent")
         Set oRelationHelper = Nothing
         Set oMarkingInfo = oCollectionHelper.Item(1)
         Set oCollectionHelper = Nothing
         If oGeom2d Is oCurrentMarkGeom2d Then
            strThisName = oMarkingInfo.Name
         Else
            strOtherName = oMarkingInfo.Name
            Select Case strOtherName
               Case "P", "S", "U", "L", "I", "O", "A", "F"
                  Set oOtherMarkGeom2d = oGeom2d
               Case Else
                  strOtherName = ""
            End Select
         End If
         Set oMarkingInfo = Nothing
      End If
      Set oGeom2d = Nothing
      If strThisName <> "" And Not oOtherMarkGeom2d Is Nothing Then
         Exit For
      End If
   Next
   Set oGeomCol2d = Nothing
   Dim dThisStartX As Double
   Dim dThisStartY As Double
   Dim dThisStartZ As Double
   Dim dThisEndX As Double
   Dim dThisEndY As Double
   Dim dThisEndZ As Double
   
   Dim oThisCurve As IJCurve
   If oOtherMarkGeom2d Is Nothing Then
        Set oThisCurve = oCurrentMarkGeom2d.GetGeometry()
        oThisCurve.EndPoints dThisStartX, dThisStartY, dThisStartZ, _
                        dThisEndX, dThisEndY, dThisEndZ
        Set oThisCurve = Nothing
        
        dLocX = dThisEndX * 1000
        dLocY = dThisEndY * 1000
      Exit Sub
   End If
   
   ' Get two mark lines' start and end points
   Dim oOtherCurve As IJCurve
   Dim dDist As Double
   
   Set oThisCurve = oCurrentMarkGeom2d.GetGeometry
   Set oOtherCurve = oOtherMarkGeom2d.GetGeometry
   Set oOtherMarkGeom2d = Nothing
   
   Dim dOtherStartX As Double
   Dim dOtherStartY As Double
   Dim dOtherStartZ As Double
   Dim dOtherEndX As Double
   Dim dOtherEndY As Double
   Dim dOtherEndZ As Double
   
   oThisCurve.EndPoints dThisStartX, dThisStartY, dThisStartZ, _
                        dThisEndX, dThisEndY, dThisEndZ
   Set oThisCurve = Nothing
   oOtherCurve.EndPoints dOtherStartX, dOtherStartY, dOtherStartZ, _
                        dOtherEndX, dOtherEndY, dOtherEndZ
   Set oOtherCurve = Nothing
   
   Dim dTempX As Double
   Dim dTempY As Double
   Dim dTempZ As Double
   
   '   X
   '   |
   '   |
   '   |
   '   X
   '     X-------X
   '
   If (dThisStartX - dOtherStartX) * (dThisStartX - dOtherStartX) + _
      (dThisStartY - dOtherStartY) * (dThisStartY - dOtherStartY) + _
      (dThisStartZ - dOtherStartZ) * (dThisStartZ - dOtherStartZ) + _
      (dThisStartX - dOtherEndX) * (dThisStartX - dOtherEndX) + _
      (dThisStartY - dOtherEndY) * (dThisStartY - dOtherEndY) + _
      (dThisStartZ - dOtherEndZ) * (dThisStartZ - dOtherEndZ) > _
      (dThisEndX - dOtherStartX) * (dThisEndX - dOtherStartX) + _
      (dThisEndY - dOtherStartY) * (dThisEndY - dOtherStartY) + _
      (dThisEndZ - dOtherStartZ) * (dThisEndZ - dOtherStartZ) + _
      (dThisEndX - dOtherEndX) * (dThisEndX - dOtherEndX) + _
      (dThisEndY - dOtherEndY) * (dThisEndY - dOtherEndY) + _
      (dThisEndZ - dOtherEndZ) * (dThisEndZ - dOtherEndZ) Then
      ' This mark's start point is far away from other mark, swap its start and end value
      dTempX = dThisStartX
      dTempY = dThisStartY
      dTempZ = dThisStartZ
      
      dThisStartX = dThisEndX
      dThisStartY = dThisEndY
      dThisStartZ = dThisEndZ
      dThisEndX = dTempX
      dThisEndY = dTempY
      dThisEndZ = dTempZ
   End If
      
   If (dOtherStartX - dThisStartX) * (dOtherStartX - dThisStartX) + _
      (dOtherStartY - dThisStartY) * (dOtherStartY - dThisStartY) + _
      (dOtherStartZ - dThisStartZ) * (dOtherStartZ - dThisStartZ) + _
      (dOtherStartX - dThisEndX) * (dOtherStartX - dThisEndX) + _
      (dOtherStartY - dThisEndY) * (dOtherStartY - dThisEndY) + _
      (dOtherStartZ - dThisEndZ) * (dOtherStartZ - dThisEndZ) > _
      (dOtherEndX - dThisStartX) * (dOtherEndX - dThisStartX) + _
      (dOtherEndY - dThisStartY) * (dOtherEndY - dThisStartY) + _
      (dOtherEndZ - dThisStartZ) * (dOtherEndZ - dThisStartZ) + _
      (dOtherEndX - dThisEndX) * (dOtherEndX - dThisEndX) + _
      (dOtherEndY - dThisEndY) * (dOtherEndY - dThisEndY) + _
      (dOtherEndZ - dThisEndZ) * (dOtherEndZ - dThisEndZ) Then
      ' Other mark's start point is far away from this mark, swap its start and end value
      dTempX = dOtherStartX
      dTempY = dOtherStartY
      dTempZ = dOtherStartZ
      
      dOtherStartX = dOtherEndX
      dOtherStartY = dOtherEndY
      dOtherStartZ = dOtherEndZ
      
      dOtherEndX = dTempX
      dOtherEndY = dTempY
      dOtherEndZ = dTempZ
   End If
   
   ' Get each mark line's direction from start to end
   Dim oThisDir As IJDVector
   Dim oOtherDir As IJDVector
   
   Set oThisDir = New DVector
   Set oOtherDir = New DVector
   
   oThisDir.Set dThisEndX - dThisStartX, _
                dThisEndY - dThisStartY, _
                dThisEndZ - dThisStartZ
                
   oOtherDir.Set dOtherEndX - dOtherStartX, _
                 dOtherEndY - dOtherStartY, _
                 dOtherEndZ - dOtherStartZ
   oThisDir.length = 1
   oOtherDir.length = 1

   ' Set default location at the end of mark line and convert to mm
   dLocX = dThisEndX
   dLocY = dThisEndY
   dLocX = dLocX * 1000
   dLocY = dLocY * 1000
   dRotAngle = 0
   
   Dim oUDir As IJDVector
   Dim oVDir As IJDVector
   
   Set oUDir = New DVector
   Set oVDir = New DVector
   
   oUDir.Set 1, 0, 0
   oVDir.Set 0, 1, 0
   
   Dim dThisDotU As Double
   Dim dThisDotV As Double
   Dim dOtherDotU As Double
   Dim dOtherDotV As Double
   
   dThisDotU = oThisDir.Dot(oUDir)
   dThisDotV = oThisDir.Dot(oVDir)
   
   ' Adjust text location and rotation angle based on mark lines' directions
   If Abs(dThisDotU) > Abs(dThisDotV) Then
      ' This is along U
      dOtherDotV = oOtherDir.Dot(oVDir)
      
      If dThisDotU > 0.001 Then
         If dOtherDotV > 0 Then
            '
            '  |
            '  |    *
            '   -----  This
            '
            dLocY = dLocY + dTextBoxHeight
         Else
            '
            '   -----  This
            '  |    *
            '  |
            '
            dRotAngle = -PI / 2
            dLocX = dLocX + dTextBoxHeight / 2
            dLocY = dLocY - dTextBoxWidth / 2
         End If
      ElseIf dThisDotU < -0.001 Then
         If dOtherDotV > 0 Then
            '
            '        |
            '  *     |
            '   -----  This
            '
            dRotAngle = PI / 2
            dLocX = dLocX - dTextBoxHeight / 2
            dLocY = dLocY + dTextBoxWidth / 2
         Else
            '
            '   -----  This
            '  *     |
            '        |
            '
            dRotAngle = PI
            dLocY = dLocY - dTextBoxHeight
         End If
      Else
         ' Should not happen
         Exit Sub
      End If
   Else
      ' This is along V
      dOtherDotU = oOtherDir.Dot(oUDir)
      
      If dThisDotV > 0.001 Then
         If dOtherDotU > 0 Then
            'This
            '    *
            '  |
            '  |
            '   -----
            '
            dLocX = dLocX + dTextBoxWidth / 2
            dLocY = dLocY + dTextBoxHeight / 2
         Else
            '    This
            '    *
            '      |
            '      |
            ' -----
            '
            dRotAngle = PI / 2
            dLocX = dLocX - dTextBoxHeight
         End If
         
      ElseIf dThisDotV < -0.001 Then
         If dOtherDotU > 0 Then
            '
            '   -----
            '  |
            '  |
            '    *
            ' This
            dRotAngle = -PI / 2
            dLocX = dLocX + dTextBoxHeight
         Else
            '
            ' -----
            '      |
            '      |
            '    *
            '     This
            dRotAngle = PI
            dLocX = dLocX - dTextBoxWidth / 2
            dLocY = dLocY - dTextBoxHeight / 2
         End If
      Else
         ' Should not happen
         Exit Sub
      End If
      
   End If

   Set oThisDir = Nothing
   Set oOtherDir = Nothing
   Set oUDir = Nothing
   Set oVDir = Nothing
   
   Exit Sub
ErrorHandler:
   
End Sub

Private Function GetRollBoundaryMarkDir(oThisGeom2D As MfgGeom2d, oMfgPart As Object) As Boolean
    Const METHOD = "FindPositionAndAngleForShipDirectionMarkText"
    On Error GoTo ErrorHandler
   
    Dim oMfgParent As IJMfgGeomParent
    Dim oChildColl As IJDTargetObjectCol
    Dim oGeomCol2d As IJMfgGeomCol2d
   
    Set oMfgParent = oMfgPart
    Set oChildColl = oMfgParent.GetChildren
    Set oMfgParent = Nothing
    Set oGeomCol2d = oChildColl.Item(1)
    Set oChildColl = Nothing

    Dim nIndex As Long
    Dim nCount As Long
    nCount = 0
    Dim bGetNext As Boolean
    bGetNext = False
    Dim oPreviousGeom2d As IJMfgGeom2d
    Dim oGeom2d As IJMfgGeom2d
    Dim oOtherMarkGeom2d As IJMfgGeom2d

    For nIndex = oGeomCol2d.Getcount To 1 Step -1
        Set oGeom2d = oGeomCol2d.GetGeometry(nIndex)
        If oGeom2d.GetGeometryType = STRMFG_ROLL_BOUNDARIES_MARK Then
            If oGeom2d.GetSubGeometryType = STRMFG_REF_MARK Or _
                oGeom2d.GetSubGeometryType = STRMFG_DIRECTION Then
                GoTo NextGeometry
            End If
            nCount = nCount + 1
            If bGetNext Then
                Set oOtherMarkGeom2d = oGeom2d
            End If
            If oGeom2d Is oThisGeom2D Then
                If (nCount Mod 2) = 0 Then
                    Set oOtherMarkGeom2d = oPreviousGeom2d
                Else
                    bGetNext = True
                End If
            Else
                Set oPreviousGeom2d = oGeom2d
            End If
        End If
NextGeometry:
        Set oGeom2d = Nothing
    Next
    Set oGeomCol2d = Nothing

    Dim oCS As IJComplexString
    Dim oCurve As IJCurve
    Dim dLength As Double
    Set oCS = oOtherMarkGeom2d.GetGeometry
    Set oCurve = oCS
    
    Dim dStartX As Double
    Dim dStartY As Double
    Dim dStartZ As Double
    Dim dEndX As Double
    Dim dEndY As Double
    Dim dEndZ As Double

    Dim dXOther As Double

    oCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
    dXOther = dStartX
    
    Set oCS = oThisGeom2D.GetGeometry
    Set oCurve = oCS
    oCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
    
    If dXOther < dStartX Then
        GetRollBoundaryMarkDir = True
    End If
    
CleanUp:
    
Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE & METHOD
    GoTo CleanUp
End Function

Private Sub IJMfgTextMarkService_Initialize(ByVal bstrXMLPath As String)
    Const METHOD = "IJMfgTextMarkService_Initialize"
    On Error GoTo ErrorHandler
    Set m_oSettingsXML = New DOMDocument
    bstrXMLPath = GetSymbolSharePath & "\" & bstrXMLPath
    m_oSettingsXML.resolveExternals = True
    m_oSettingsXML.validateOnParse = False
    
    If Dir(bstrXMLPath) <> "" Then ' Fileexists
        m_oSettingsXML.Load bstrXMLPath
        m_bDoesXMLExist = True
    Else
        m_bDoesXMLExist = False
    End If

    'Assignnig the Path of the Annotation XML to a Global Variable
    m_strXMLData = bstrXMLPath
    m_oSettingsXML.setProperty "SelectionLanguage", "XPath"
CleanUp:
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, MODULE & METHOD
    GoTo CleanUp

End Sub

''**************************************************************************************
'' Routine      : GetSymbolSharePath
'' Abstract     : This method will return the SymbolShare Path ( like M:\CatalogData\Symbols).
''                Assumption: There will be Structmanufacturing folder under symbols directory. If doesn't exist
''                Get the LogicalSymbolPath registry key (This will be applicable for developer build)
''
''**************************************************************************************
Private Function GetSymbolSharePath() As String
    Const METHOD = "GetSymbolLocationPath"
    On Error GoTo ErrorHandler
    
    Dim oContext As IJContext
    Dim strContextString As String
    Dim oFileSysObj As New FileSystemObject
    Dim strSymbolShare As String
    
    strContextString = "OLE_SERVER"
    
    'Get IJContext
    Set oContext = GetJContext()
    
    If Not oContext Is Nothing Then
        'Get the Symbol Share
        strSymbolShare = oContext.GetVariable(strContextString)
    
        'Append StructManufacturing
        strSymbolShare = strSymbolShare
    
        'Check whether the folder exists if not get the Registry key Value
        If oFileSysObj.FolderExists(strSymbolShare) Then
            GetSymbolSharePath = strSymbolShare
        Else
            GetSymbolSharePath = GetRegValue("LogicalSymbolPath")
        End If
    Else
        GetSymbolSharePath = GetRegValue("LogicalSymbolPath")
    End If
    
    Set oContext = Nothing
    Set oFileSysObj = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE & METHOD
    Resume Next
End Function

''**************************************************************************************
'' Routine      : GetRegValue
'' Abstract     : Get the value from Registry
''
''**************************************************************************************
Private Function GetRegValue(Key As String) As String
    Const METHOD = "GetRegValue"
    On Error GoTo ErrHandler
    
    Dim oRegistry As IJRegistry
    Dim varValue As Variant
    Dim sHKeyPath As String
    Dim strPath As String

    Set oRegistry = New Registry

    sHKeyPath = "\HKEY_LOCAL_MACHINE\"
    strPath = sHKeyPath & INGRPATH & SYMBOLPATH

    varValue = 0
    varValue = oRegistry.getValue(strPath, Key)

    GetRegValue = varValue
    
    Exit Function

ErrHandler:
    Err.Raise Err.Number, MODULE & METHOD
    Resume Next
End Function

''**************************************************************************************
'' Routine      : GetLappedPlateAnnotationPosition
'' Abstract     : This Function gets the point where the Toshi Mark Annotation
''                need to be placed on the input curve
''
''**************************************************************************************
Private Function GetLappedPlateAnnotationPosition(ByVal oCS As IJComplexString, ByRef oPosition As IJDPosition)
    Const METHOD = "GetLappedPlateAnnotationPosition"
    On Error GoTo ErrorHandler

    Dim dCount As Double
    Dim oCSColl As IJElements
    Set oCSColl = New JObjectCollection
    oCSColl.Add oCS
        
    Dim oVectorX As IJDVector
    Dim oVectorY As IJDVector
    Dim oVectorZ As IJDVector
        
    Set oVectorX = New DVector
    Set oVectorY = New DVector
    Set oVectorZ = New DVector
        
    oVectorX.Set 1, 0, 0
    oVectorY.Set 0, 1, 0
    oVectorZ.Set 0, 0, 1
                            
    Dim oVectorElements As IJElements
    Set oVectorElements = New JObjectCollection
        
    oVectorElements.Add oVectorX
    oVectorElements.Add oVectorY
    oVectorElements.Add oVectorZ
        
    '   Logic For Getting Minimum Bounding Box
    '                       ________________
    '                      /.              /'
    '                     /_._____________/3'
    '                     ' .             ' '
    '                     ' ..............'.'
    '                     '/______________'/
    '                     1               2
    '                    Minimum Bounding Box

    Dim oMfgGeomHelper As New MfgGeomHelper
    Dim oBoxPoints As IJElements
    Set oBoxPoints = oMfgGeomHelper.GetGeometryMinBoxByVectors(oCSColl, oVectorElements)
    
    Dim oPoint(1 To 3) As IJDPosition
    Set oPoint(1) = oBoxPoints.Item(1)
    Set oPoint(2) = oBoxPoints.Item(2)
    Set oPoint(3) = oBoxPoints.Item(3)
    
    
    Dim oPoint4 As IJDPosition
    Set oPoint4 = New DPosition
    oPoint4.x = oPoint(1).x + oPoint(3).x - oPoint(2).x
    oPoint4.y = oPoint(1).y + oPoint(3).y - oPoint(2).y
    oPoint4.z = oPoint(1).z + oPoint(3).z - oPoint(2).z
    '                    _______.3______
    '                   '               '
    '                   '               '
    '                   .4              .2          1,2,3,4 are cross points
    '                   '               '
    '                   ' ______.1______'
                            
    
    Dim oCrossPoint(1 To 4) As IJDPosition
    Set oCrossPoint(1) = New DPosition
    Set oCrossPoint(2) = New DPosition
    Set oCrossPoint(3) = New DPosition
    Set oCrossPoint(4) = New DPosition
    
    oCrossPoint(1).x = (oPoint(1).x + oPoint(2).x) / 2
    oCrossPoint(1).y = (oPoint(1).y + oPoint(2).y) / 2
    oCrossPoint(1).z = (oPoint(1).z + oPoint(2).z) / 2
    
    oCrossPoint(2).x = (oPoint(2).x + oPoint(3).x) / 2
    oCrossPoint(2).y = (oPoint(2).y + oPoint(3).y) / 2
    oCrossPoint(2).z = (oPoint(2).z + oPoint(3).z) / 2
    
    oCrossPoint(3).x = (oPoint(3).x + oPoint4.x) / 2
    oCrossPoint(3).y = (oPoint(3).y + oPoint4.y) / 2
    oCrossPoint(3).z = (oPoint(3).z + oPoint4.z) / 2
    
    oCrossPoint(4).x = (oPoint4.x + oPoint(1).x) / 2
    oCrossPoint(4).y = (oPoint4.y + oPoint(1).y) / 2
    oCrossPoint(4).z = (oPoint4.z + oPoint(1).z) / 2
        
    Dim oIntLine1 As IJLine
    Dim oIntLine2 As IJLine
                        
    'Setting the vector with the available cross points and then normalizing it
    Dim oIntVector1 As IJDVector
    Dim oIntVector2 As IJDVector
        
    Set oIntVector1 = oCrossPoint(3).Subtract(oCrossPoint(1))
    Set oIntVector2 = oCrossPoint(4).Subtract(oCrossPoint(2))
    
    oIntVector1.length = 1
    oIntVector2.length = 1
    'End of the logic for getting the vector
    
    Set oIntLine1 = New Line3d
    Set oIntLine2 = New Line3d
    oIntLine1.DefineBy2Points oCrossPoint(1).x, oCrossPoint(1).y, oCrossPoint(1).z, oCrossPoint(3).x, oCrossPoint(3).y, oCrossPoint(3).z
    oIntLine2.DefineBy2Points oCrossPoint(2).x, oCrossPoint(2).y, oCrossPoint(2).z, oCrossPoint(4).x, oCrossPoint(4).y, oCrossPoint(4).z
    
    Dim oCS1 As IJComplexString
    Dim oCS2 As IJComplexString
    Set oCS1 = New ComplexString3d
    Set oCS2 = New ComplexString3d
    
    oCS1.AddCurve oIntLine1, True
    oCS2.AddCurve oIntLine2, True
        
    Dim oIntElems1 As IJElements
    Dim oIntElems2 As IJElements
    
    'Storing thepoint of intersection of curve and the lines as a collection
    Set oIntElems1 = oMfgGeomHelper.IntersectCurveWithCurve(oCS, oCS1)
    Set oIntElems2 = oMfgGeomHelper.IntersectCurveWithCurve(oCS, oCS2)
    dCount = oIntElems1.Count + oIntElems2.Count
    
    Dim oTempPos1 As IJDPosition
    Dim oTempPos2 As IJDPosition

    Set oTempPos1 = New DPosition
    Set oTempPos2 = New DPosition
            
    'Logic for getting the point where the annotation (Toshi mark) need to be placed
    If ((oIntElems1.Count) + (oIntElems2.Count) = 3) Then
        If (oIntElems1.Count = 1) Then
            Set oPosition = oIntElems1.Item(1)
        ElseIf (oIntElems2.Count = 1) Then
            Set oPosition = oIntElems2.Item(1)
        End If
    ElseIf ((oIntElems1.Count) + (oIntElems2.Count) = 2) Then
        If (Abs(oIntVector1.y) > 0.999) Then
            Set oPosition = oIntElems1.Item(1)
        Else
            Set oPosition = oIntElems2.Item(1)
        End If
    Else
        
        If (Abs(oIntVector1.y) > 0.999) Then
            Set oTempPos1 = oIntElems1.Item(1)
            Set oTempPos2 = oIntElems1.Item(2)
            If (oTempPos1.y > oTempPos2.y) Then
                Set oPosition = oIntElems1.Item(1)
            Else
                Set oPosition = oIntElems1.Item(2)
            End If
        Set oTempPos1 = Nothing
        Set oTempPos2 = Nothing
       Else
            Set oTempPos1 = oIntElems2.Item(1)
            Set oTempPos2 = oIntElems2.Item(2)
            If (oTempPos1.y > oTempPos2.y) Then
                Set oPosition = oIntElems2.Item(1)
            Else
                Set oPosition = oIntElems2.Item(2)
            End If
        Set oTempPos1 = Nothing
        Set oTempPos2 = Nothing
       End If
    End If
    
CleanUp:
    Erase oCrossPoint
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Sub FillBevelValuesFromXML(ByVal sBevelXml As String, ByRef strGUID, ByRef dAngleA As Double, _
                              ByRef dAngleB As Double, ByRef dAngleN As Double, _
                              ByRef dAngleD As Double, ByRef dAngleE As Double, _
                              ByRef dChamferAngleM As Double, ByRef dChamferDepthM As Double, _
                              ByRef dChamferAngleUM As Double, ByRef dChamferDepthUM As Double)
    
    Const METHOD = "FillBevelValuesFromXML"
    On Error GoTo ErrorHandler
    
    Dim oBevelDom As New DOMDocument
    'Dim oNodeList As IXMLDOMNodeList
    Dim oBevelElem As IXMLDOMElement
    Dim sValue As String
    Dim vTemp As Variant
        
    strGUID = ""
    dAngleA = 0
    dAngleB = 0
    dAngleD = 0
    dAngleN = 0
   
    If Not oBevelDom.loadXML(sBevelXml) Then GoTo CleanUp
    
    'Assumes the first node found is the desired node
    Set oBevelElem = oBevelDom.selectSingleNode("//SMS_BEVEL")
    
    If oBevelElem Is Nothing Then GoTo CleanUp
    
    vTemp = oBevelElem.GetAttribute("GEOM2D_GUID")
    strGUID = IIf(VarType(vTemp) = vbString, vTemp, "")
    
    sValue = 0
    vTemp = oBevelElem.GetAttribute("ANGLE_A")
    sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
    If isNumericUS(sValue) Then
        dAngleA = Val(sValue)
    End If
    
    sValue = 0
    vTemp = oBevelElem.GetAttribute("ANGLE_B")
    sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
    If isNumericUS(sValue) Then
        dAngleB = Val(sValue)
    End If
    
    sValue = 0
    vTemp = oBevelElem.GetAttribute("ANGLE_N")
    sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
    If isNumericUS(sValue) Then
        dAngleN = Val(sValue)
    End If
    
    sValue = 0
    vTemp = oBevelElem.GetAttribute("ANGLE_D")
    sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
    If isNumericUS(sValue) Then
        dAngleD = Val(sValue)
    End If
    
    sValue = 0
    vTemp = oBevelElem.GetAttribute("ANGLE_E")
    sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
    If isNumericUS(sValue) Then
        dAngleE = Val(sValue)
    End If
    
    sValue = 0
    vTemp = oBevelElem.GetAttribute("CHAMFER_ANGLE_M")
    sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
    If isNumericUS(sValue) Then
        dChamferAngleM = Val(sValue)
    End If
    
    sValue = 0
    vTemp = oBevelElem.GetAttribute("CHAMFER_DEPTH_M")
    sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
    If isNumericUS(sValue) Then
        dChamferDepthM = Val(sValue)
    End If
    
    sValue = 0
    vTemp = oBevelElem.GetAttribute("CHAMFER_ANGLE_UM")
    sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
    If isNumericUS(sValue) Then
        dChamferAngleUM = Val(sValue)
    End If
    
    sValue = 0
    vTemp = oBevelElem.GetAttribute("CHAMFER_DEPTH_UM")
    sValue = IIf(VarType(vTemp) = vbString, vTemp, "")
    If isNumericUS(sValue) Then
        dChamferDepthUM = Val(sValue)
    End If
    
CleanUp:
    Set oBevelDom = Nothing
    Set oBevelElem = Nothing
    'Set oNodeList = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number
    Resume Next
End Sub

''**************************************************************************************
'' Routine      : GetTextSizeForMfgPart
'' Abstract     : This Function gives the Text Size of the Annotation Based
''                on the size of the Plate Dimensions & also checks whether the
''                Part is Plate or Profile
''**************************************************************************************

Private Function GetTextSizeForMfgPart(ByVal oMfgPart As Object) As Double

    Const METHOD = "GetTextSizeForMfgPart"
    On Error GoTo ErrorHandler

    Dim dTextSize           As Double

    If TypeOf oMfgPart Is IJMfgPlatePart Then
        Dim oMfgPlate           As IJMfgPlatePart
        Dim oPlatePart          As IJPlatePart
        
        Set oMfgPlate = oMfgPart
        oMfgPlate.GetDetailedPart oPlatePart
        
        If (TypeOf oPlatePart Is IJCollarPart) Then
            'this Means the Plate Part is a Collar
            dTextSize = TEXT_SIZE_COLLAR
            GetTextSizeForMfgPart = dTextSize
            Exit Function
        End If
        
        Dim oSystem             As IJSystem
        Dim oStructDetailHelper As GSCADStructDetailUtil.StructDetailHelper

        Set oStructDetailHelper = New GSCADStructDetailUtil.StructDetailHelper
        oStructDetailHelper.IsPartDerivedFromSystem oPlatePart, oSystem, True
        
        Dim oPlateUtils As IJPlateAttributes
        Set oPlateUtils = New PlateUtils

        If (IsFlangedBracket(oPlatePart) = True) Or (oPlateUtils.IsBracketByPlane(oSystem) = True) Or (oPlateUtils.IsTrippingBracket(oSystem) = True) Then
            'this Means the Plate Part is a Bracket
            dTextSize = TEXT_SIZE_BRACKET
            GetTextSizeForMfgPart = dTextSize
            Exit Function
        End If

        'Check for the Plate Dimensions
        If Round(oMfgPlate.length, 3) * Round(oMfgPlate.Width, 3) <= CHECK_PLATE_AREA Then
            dTextSize = MIN_TEXT_SIZE_PLATE
        Else
            dTextSize = MAX_TEXT_SIZE_PLATE
        End If
        Set oMfgPlate = Nothing
    Else
        'Here the Object Type is Profile Part
        dTextSize = TEXT_SIZE_PROFILE
    End If
    GetTextSizeForMfgPart = dTextSize

     Exit Function
         
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

''**************************************************************************************
'' Routine      : GetSecondaryTextSizeForMfgPart
'' Abstract     : This Function gives the Secondary Text Size of the Ship Annotation Based
''                on the size of the Plate Dimensions & also checks whether the
''                Part is Plate or Profile
''**************************************************************************************

Private Function GetSecondaryTextSizeForMfgPart(ByVal oMfgPart As Object) As Double

    Const METHOD = "GetSecondaryTextSizeForMfgPart"
    On Error GoTo ErrorHandler
    
    
    Dim oMfgPlate           As IJMfgPlatePart
    Dim oPlatePart          As IJPlatePart
    Dim oSDPlatePart        As New StructDetailObjects.PlatePart
    Dim dTextSize           As Double
    Dim dSecondaryTextSize  As Double

    dSecondaryTextSize = 0
    
    If TypeOf oMfgPart Is IJMfgPlatePart Then
        Set oMfgPlate = oMfgPart
        oMfgPlate.GetDetailedPart oPlatePart
        Set oSDPlatePart.object = oPlatePart
        
        dTextSize = GetTextSizeForMfgPart(oMfgPart)
        
        'When the Text Size is of minimum size then the Ship Annotation needs to be of same text size
        If dTextSize = MIN_TEXT_SIZE_PLATE Or dTextSize = TEXT_SIZE_COLLAR Or dTextSize = TEXT_SIZE_BRACKET Then
            Exit Function
        Else
            Select Case oSDPlatePart.plateType
                Case DeckPlate, Hull, LBulkheadPlate
                    If Round(oMfgPlate.length, 3) > MAX_PLATE_LENGTH Or Round(oMfgPlate.Width, 3) > MAX_PLATE_WIDTH Then
                    'if the Plate Part is big then the ship Annotation text size varies according to dimensions
                        dSecondaryTextSize = SECONDARY_TEXT_SIZE_PLATE
                    Else
                        'No Action is needed hence Exit the Function
                        Exit Function
                    End If
                Case Else
                    'No Action is needed hence Exit the Function
                    Exit Function
                End Select
        End If
    Else
        'A Case where the Object Type is Profile Type hence Exit the Function
        Exit Function
    End If
    
    GetSecondaryTextSizeForMfgPart = dSecondaryTextSize
    
    Set oMfgPlate = Nothing
    Set oPlatePart = Nothing
    Set oSDPlatePart = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function GetXMLEquivalentCharacter(sString As String) As String

    Const METHOD = "GetXMLEquivalentCharacter"
    On Error GoTo ErrorHandler
    
    Select Case sString
        Case "<"
            GetXMLEquivalentCharacter = "&lt;"
        Case ">"
            GetXMLEquivalentCharacter = "&gt;"
        Case "'"
            GetXMLEquivalentCharacter = "&apos;"
        Case """"
            GetXMLEquivalentCharacter = "&quot;"
        Case "&"
            GetXMLEquivalentCharacter = "&amp;"
    End Select

Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

''**************************************************************************************
'' Routine      : GetGrinderType
'' Abstract     : This Function gives the Grind type of the free edge treatment
 ''               Applied on the Part
''**************************************************************************************
Private Function GetGrinderType(ByVal oMfgPart As Object, ByVal oMfgGeom2d As IJMfgGeom2d, ByRef strGrinderType As String)
    Const METHOD = "GetGrinderType"
    On Error GoTo ErrorHandler
    
    If oMfgPart Is Nothing Or oMfgGeom2d Is Nothing Then GoTo CleanUp
  
    Dim oPortMoniker        As IMoniker
    Dim oPartSupport        As GSCADSDPartSupport.IJPartSupport
    Dim oConnColl           As Collection
    Dim oGrindFETColl       As Collection
    Dim oPOM                As IJDPOM
    Dim oObject             As IJDObject
    Dim oSO                 As IJSmartOccurrence
    Dim oSmartItem          As IJSmartItem
    Dim oMfgPlate           As IJMfgPlatePart
    Dim oMfgProfile         As IJMfgProfilePart
    Dim oPlatePart          As IJPlatePart
    Dim oProfilePart        As IJProfilePart
    Dim oStructPort         As IJStructPort
    
    strGrinderType = vbNullString
    
    Set oPartSupport = New GSCADSDPartSupport.PartSupport
    
    'From the Manufactured Knowing whether the Part Type is Plate or Profile
    If TypeOf oMfgPart Is IJMfgPlatePart Then
        Set oMfgPlate = oMfgPart
        oMfgPlate.GetDetailedPart oPlatePart
        Set oPartSupport.Part = oPlatePart
        Set oObject = oPlatePart
    Else
        Set oMfgProfile = oMfgPart
        oMfgPlate.GetDetailedPart oProfilePart
        Set oPartSupport.Part = oProfilePart
        Set oObject = oProfilePart
    End If
    
    'Getting the Port Moniker of the Edge
    Set oPortMoniker = oMfgGeom2d.GetMoniker
    Set oPOM = oObject.ResourceManager
   
    Set oStructPort = oPOM.GetObject(oPortMoniker)
    Set oPortMoniker = Nothing
    Set oPortMoniker = oStructPort.PortMoniker
    oPartSupport.GetBevel oPortMoniker, oConnColl, oGrindFETColl
    
    If oGrindFETColl Is Nothing Then
        GoTo CleanUp
    ElseIf oGrindFETColl.Count = 0 Then
        GoTo CleanUp
    End If
    
    Set oSO = oGrindFETColl.Item(1)
    
    Set oSmartItem = oSO.ItemObject
    strGrinderType = oSmartItem.Name

CleanUp:
    Set oSmartItem = Nothing
    Set oSO = Nothing
    Set oGrindFETColl = Nothing
    Set oConnColl = Nothing
    Set oPartSupport = Nothing
    Set oPortMoniker = Nothing
    Set oPOM = Nothing
    Set oObject = Nothing
    Set oMfgPlate = Nothing
    Set oMfgProfile = Nothing
    Set oPlatePart = Nothing
    Set oProfilePart = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

''**************************************************************************************
'' Routine      : IfBevelExistsForProfile
'' Abstract     : This Function checks if bevel exists for given profile
''**************************************************************************************
Private Function IfBevelExistsForProfile(oProfPart As Object) As Boolean

    Const METHOD = "IfBevelExistsForProfile"
    On Error GoTo ErrorHandler
    
    IfBevelExistsForProfile = False
    
    Dim oMfgProfilePart As IJMfgProfilePart
    Dim oOutrCntGeom3dColl As IJMfgGeomCol3d
    Dim oOutrCntGeom2dColl As IJMfgGeomCol2d
    Dim oMfgGeom3d As IJMfgGeom3d
    Dim oMfgGeom2d As IJMfgGeom2d
    Dim oMfgBevel As IJMfgBevel
    Dim j As Long
    
    If TypeOf oProfPart Is IJMfgProfilePart Then
        Set oMfgProfilePart = oProfPart
        Set oOutrCntGeom2dColl = oMfgProfilePart.FinalGeometriesAfterProcess2D
        
        'For each outer contour, check if bevel exists
        For j = 1 To oOutrCntGeom2dColl.Getcount
            Set oMfgGeom2d = oOutrCntGeom2dColl.GetGeometry(j)
            
            If oMfgGeom2d.GetGeometryType = STRMFG_OUTER_CONTOUR Or oMfgGeom2d.GetGeometryType = STRMFG_INNER_CONTOUR Then
                Set oMfgBevel = oMfgGeom2d.GetBevel
                If Not oMfgBevel Is Nothing Then 'Bevel exists, return true and exit
                    IfBevelExistsForProfile = True
                    Exit Function
                End If
            End If
        
        Next j
    End If

    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

''**************************************************************************************
'' Routine      : IfMarkingExistsForProfile
'' Abstract     : This Function checks if markings exists for given profile
''**************************************************************************************
Private Function IfMarkingExistsForProfile(oProfPart As Object) As Boolean

    Const METHOD = "IfMarkingExistsForProfile"
    On Error GoTo ErrorHandler
    
    IfMarkingExistsForProfile = False
    
    Dim oMfgProfilePart As IJMfgProfilePart
    Dim oOutrCntGeom3dColl As IJMfgGeomCol3d
    Dim oOutrCntGeom2dColl As IJMfgGeomCol2d
    Dim oMfgGeom3d As IJMfgGeom3d
    Dim oMfgGeom2d As IJMfgGeom2d
    Dim oMfgBevel As IJMfgBevel
    Dim j As Long
    Dim oMfgEntityHelper As New MfgEntityHelper
    Dim sMarkingName As String
    Dim lMarkingCategory As Long
    
    If TypeOf oProfPart Is IJMfgProfilePart Then
        Set oMfgProfilePart = oProfPart
        Set oOutrCntGeom2dColl = oMfgProfilePart.FinalGeometriesAfterProcess2D
        
        'For each outer contour, check if bevel exists
        For j = 1 To oOutrCntGeom2dColl.Getcount
            Set oMfgGeom2d = oOutrCntGeom2dColl.GetGeometry(j)
            
            'Check the marking category for each geom3d
            oMfgEntityHelper.GeomTypeToString oMfgGeom2d, sMarkingName, lMarkingCategory
            If lMarkingCategory > -1 And lMarkingCategory < 5 And oMfgGeom2d.IsSupportOnly <> True Then ' >0 means its a Marking
                    IfMarkingExistsForProfile = True
                    Exit Function
            End If
        
        Next j
    End If

    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function CheckIfAllGrindsAreSameForOneEdge(oContourDoc As DOMDocument, oCurve As IJCurve) As Boolean

    Dim oXMLCurveList As IXMLDOMNodeList
    Dim oXMLGrindList As IXMLDOMNodeList
    Dim oXMLElem As IXMLDOMElement
    Dim oXMLGrindElem1 As IXMLDOMElement
    Dim oXMLGrindElem2 As IXMLDOMElement
    Dim oXMLCurveElem1 As IXMLDOMElement
    Dim oXMLCurveElem2 As IXMLDOMElement
    Dim oXMLNode As IXMLDOMNode
    Dim oPointList1 As IXMLDOMNodeList, oPointList2 As IXMLDOMNodeList
    Dim sNAME As String, sRADIUS_M As String, sRADIUS_UM As String
    Dim dGrindStartX As Double, dGrindStartY As Double, dGrindEndX As Double, dGrindEndY As Double
    Dim dCurveStartX As Double, dCurveStartY As Double, dCurveStartZ As Double
    Dim dCurveEndX As Double, dCurveEndY As Double, dCurveEndZ As Double
    Dim bAllGrindsHaveSameName As Boolean
    
    Set oXMLCurveList = oContourDoc.selectNodes("//SMS_EDGE/CVG_CURVE")
    Set oXMLGrindList = oContourDoc.selectNodes("//SMS_EDGE/SMS_GRIND")
    
    CheckIfAllGrindsAreSameForOneEdge = True
    
    Set oXMLGrindElem1 = oXMLGrindList.Item(0)
    Set oXMLGrindElem2 = oXMLGrindList.Item(oXMLGrindList.length - 1)
    
    Set oPointList1 = oXMLGrindElem1.getElementsByTagName("CVG_POINT")
    Set oPointList2 = oXMLGrindElem2.getElementsByTagName("CVG_POINT")

    dGrindStartX = Val(oPointList1.Item(0).Attributes.getNamedItem("X").Text) / 1000
    dGrindStartY = Val(oPointList1.Item(0).Attributes.getNamedItem("Y").Text) / 1000
    dGrindEndX = Val(oPointList2.Item(1).Attributes.getNamedItem("X").Text) / 1000
    dGrindEndY = Val(oPointList2.Item(1).Attributes.getNamedItem("Y").Text) / 1000
    
    oCurve.EndPoints dCurveStartX, dCurveStartY, dCurveStartZ, dCurveEndX, dCurveEndY, dCurveEndZ
    
    'Check if all grinds are connected to each other and
    ' .. Start Pt & End Pts are matching and all have same name. This means that all grinds are same for given
    ' .. SMS_EDGE
        
    Dim ii As Integer
    'This case handles <CVG_CURVE> <SMS_GRIND> ... <CVG_CURVE> <SMS_GRIND> case
    If oXMLCurveList.length = oXMLGrindList.length And oXMLGrindList.length = 1 Then
        If (Abs(dGrindStartX - dCurveStartX) < 0.001 And Abs(dGrindStartY - dCurveStartY) < 0.001) And _
            (Abs(dGrindEndX - dCurveEndX) < 0.001 And Abs(dGrindEndY - dCurveEndY) < 0.001) Then
                CheckIfAllGrindsAreSameForOneEdge = True
            Else
                CheckIfAllGrindsAreSameForOneEdge = False
            End If

    ElseIf oXMLCurveList.length = oXMLGrindList.length And oXMLGrindList.length > 1 Then
        Set oXMLElem = oXMLGrindList.Item(0)
        sNAME = oXMLElem.GetAttribute("NAME")
        For Each oXMLElem In oXMLGrindList
            Set oXMLNode = oXMLElem
            If oXMLNode.nodeName = "SMS_GRIND" Then
                If oXMLElem.GetAttribute("NAME") <> sNAME Then
                    CheckIfAllGrindsAreSameForOneEdge = False
                    Exit Function
                End If
            End If
        Next oXMLElem
    Else
    'This case handles <CVG_CURVE> <SMS_GRIND> ...  <SMS_GRIND> case
        
        '----- Check if All Grinds Have same name --------------'
        bAllGrindsHaveSameName = True
        Set oXMLElem = oXMLGrindList.Item(0)
        sNAME = oXMLElem.GetAttribute("NAME")
        For Each oXMLElem In oXMLGrindList
            Set oXMLNode = oXMLElem
            If oXMLNode.nodeName = "SMS_GRIND" Then
                If oXMLElem.GetAttribute("NAME") <> sNAME Then
                    CheckIfAllGrindsAreSameForOneEdge = False
                    Exit Function
                End If
            End If
        Next oXMLElem
        '-------------------------------------------------------'
        
        oCurve.EndPoints dCurveStartX, dCurveStartY, dCurveStartZ, dCurveEndX, dCurveEndY, dCurveEndZ
        
        Set oXMLGrindElem1 = oXMLGrindList.Item(0)
        Set oXMLGrindElem2 = oXMLGrindList.Item(oXMLGrindList.length - 1)
         
        Set oPointList1 = oXMLGrindElem1.getElementsByTagName("CVG_POINT")
        Set oPointList2 = oXMLGrindElem2.getElementsByTagName("CVG_POINT")

        dGrindStartX = Val(oPointList1.Item(0).Attributes.getNamedItem("X").Text) / 1000
        dGrindStartY = Val(oPointList1.Item(0).Attributes.getNamedItem("Y").Text) / 1000
        dGrindEndX = Val(oPointList2.Item(1).Attributes.getNamedItem("X").Text) / 1000
        dGrindEndY = Val(oPointList2.Item(1).Attributes.getNamedItem("Y").Text) / 1000
        
        'Check if GrindStartPt = CurveStartPt and GrindEndPt = CurveEndPt
        If (Abs(dGrindStartX - dCurveStartX) < 0.001 And Abs(dGrindStartY - dCurveStartY) < 0.001) And _
            (Abs(dGrindEndX - dCurveEndX) < 0.001 And Abs(dGrindEndY - dCurveEndY) < 0.001) Then
            
            For ii = 0 To oXMLGrindList.length - 2
                Set oXMLGrindElem1 = oXMLGrindList.Item(ii)
                Set oXMLGrindElem2 = oXMLGrindList.Item(ii + 1)
                
                Set oPointList1 = oXMLGrindElem1.getElementsByTagName("CVG_POINT")
                dGrindEndX = Val(oPointList1.Item(1).Attributes.getNamedItem("X").Text) / 1000
                dGrindEndY = Val(oPointList1.Item(1).Attributes.getNamedItem("Y").Text) / 1000
                
                Set oPointList2 = oXMLGrindElem2.getElementsByTagName("CVG_POINT")
                dGrindStartX = Val(oPointList2.Item(0).Attributes.getNamedItem("X").Text) / 1000
                dGrindStartY = Val(oPointList2.Item(0).Attributes.getNamedItem("Y").Text) / 1000
                
                If (Abs(dGrindEndX - dGrindStartX) < 0.001 And Abs(dGrindEndY - dGrindStartY) < 0.001) _
                    And bAllGrindsHaveSameName Then
                    'Grinds are connected to each other
                Else
                    'Grinds are not connected to each other
                    CheckIfAllGrindsAreSameForOneEdge = False
                    Exit Function
                End If
            Next ii
        Else
            CheckIfAllGrindsAreSameForOneEdge = False
        End If
    End If
    
CleanUp:
    Exit Function
    
ErrorHandler:
End Function

Private Sub PutGeomGUIDvalue(strInputXML As String, oEntityNameElem As IXMLDOMElement)
    Const METHOD = "PutGeomGUIDvalue"
    On Error GoTo ErrorHandler
    
    Dim oMarkingDoc         As DOMDocument
    Dim oMiscElem           As IXMLDOMElement
    Dim oTextArgList        As IXMLDOMNodeList
    Dim oTextElem           As IXMLDOMElement
    Dim oSettingsElem       As IXMLDOMElement
    Dim sAttribute          As String
    Dim sValue              As String
    Dim oAttr               As IXMLDOMAttribute
    
    Set oMarkingDoc = New DOMDocument
    
    If Not oMarkingDoc.loadXML(strInputXML) Then GoTo CleanUp
   
    Set oTextArgList = oEntityNameElem.selectNodes(".//PROPERTY")
    
    For Each oTextElem In oTextArgList
        sAttribute = oTextElem.GetAttribute("NAME")
        If sAttribute = "SMS_ANNOTATION||GEOM2D_GUID" Then
            Set oAttr = oMarkingDoc.selectSingleNode("//@GEOM2D_GUID")
            If Not oAttr Is Nothing Then
                Set oMiscElem = oMarkingDoc.selectSingleNode("//*[@GEOM2D_GUID]")
                sValue = oMiscElem.GetAttribute("GEOM2D_GUID")
                    
                Set oSettingsElem = oEntityNameElem.selectSingleNode(".//PROPERTY[@NAME='" & sAttribute & "']")
                oSettingsElem.setAttribute "VALUE", sValue
                
                Exit For
            End If
        End If
    Next
   
CleanUp:
    Set oMiscElem = Nothing
    Set oMarkingDoc = Nothing
    Set oTextArgList = Nothing
    Set oTextElem = Nothing
    Set oSettingsElem = Nothing
    Set oAttr = Nothing
        
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub

' ***********************************************************************************
' Public Function GetAssemblyType
'
' Description:  Method will give the Type of Assembly
'
' ***********************************************************************************
Private Function GetAssemblyType(oPart As Object) As Long
    Const METHOD = "GetAssemblyType"
    On Error GoTo ErrorHandler

    Dim oAssemblyChild As IJAssemblyChild
    Dim oUAFolder As IJAssemblyChild
    Dim oAssemblybase As IJAssemblyBase
    Dim lAssemblyType As Long
        
    Set oAssemblyChild = oPart
    
    If oAssemblyChild.Parent Is Nothing Then 'This will Exist when the Part is under Config Root
        Exit Function
    ElseIf TypeOf oAssemblyChild.Parent Is IJAssemblyBase Then
        Set oAssemblybase = oAssemblyChild.Parent
    'Below Condition will occurs when the Part is under UnProcessedParts or FailedParts
    ElseIf TypeOf oAssemblyChild.Parent Is IJPlnUnprocessedParts Or TypeOf oAssemblyChild.Parent Is IJPlnFailedParts Then
        Set oUAFolder = oAssemblyChild.Parent
        Set oAssemblybase = oUAFolder.Parent
    End If
    
    If Not oAssemblybase Is Nothing Then
        lAssemblyType = oAssemblybase.Type
    Else
        Exit Function
    End If
    
    GetAssemblyType = lAssemblyType
    
    Exit Function
    
ErrorHandler:
   Err.Raise Err.Number, Err.Source, Err.Description

End Function





