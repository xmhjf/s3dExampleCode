Attribute VB_Name = "PF_HgrAssyUtil"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2005, Intergraph Corporation. All rights reserved.
'
'   PF_HgrAssyUtil.bas
'   ProgID:
'   Author:         Srinivas
'   Creation Date:  03.Mar.2005
'   Description:
'
'
'   Change History:
'        Date           Name                    Description
'   24.May.2006         Chinta,Mahesh           DI CP:95927(Address the code review comments for  PF Assembly Common Modules)
'   02.Jun.2006         Ravi ippili             DI-CP·96213 Modified the GetProfileDim to get the backtoback distance from hanger beam
'   06.Jul.2006         SS                      TR 98162 - Need to get Pipe OD without insulation
'   28 Jan 2013         VDP                     DI-CP-224963    Due to default note text of Supports, Not able to get correct notes in ISO
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "PF_HgrAssyUtil" 'Used for error messages

'*************************************************************************************
' Method:   AddPartsByClassName
'
' Desc:     This Function(normally called from IJHgrAssmInfo_GetAssemblyCatalogParts)
'           It Simplifies the specification of the parts included
'
' Param:
'           pDispInputConfigHlpr    -       Object
'           PartClasses             -       String
'           Collection              -       Returns a Collection of Catalogparts
'*************************************************************************************

Public Function AddPartsByClassName(ByVal pDispInputConfigHlpr As Object, PartClasses() As String) As Collection
Const METHOD = "AddPartsByClassName"
    On Error GoTo ErrorHandler

    'Get IJHgrInputConfig Hlpr Interface off of passed Helper
    Dim my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr
    Set my_IJHgrInputConfigHlpr = pDispInputConfigHlpr
    
    'Create a new collection to hold the caltalog parts
    Dim CatalogPartCollection As New Collection
        
    'Dimension Object to temporarily hold a Catalog Part
    Dim PartProxy As Object
    
    'Use the default selection rule to get a catalog part for each part class
    Dim Index As Integer
    For Index = 1 To UBound(PartClasses)
        Set PartProxy = my_IJHgrInputConfigHlpr.GetPartProxyFromName(PartClasses(Index))
        CatalogPartCollection.Add PartProxy
    Next
    
    'Return the collection of Catalog Parts
    Set AddPartsByClassName = CatalogPartCollection
    
    Set CatalogPartCollection = Nothing
    Set PartProxy = Nothing
    Set my_IJHgrInputConfigHlpr = Nothing
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "Perhaps problem with " + PartClasses(Index)).Number
End Function

'*************************************************************************************
' Method:   GetPipeRadiusDBU
'
' Desc:     This Function gets the PipeRadiusDatabaseunits
'
' Param:
'           pDispInputConfigHlpr    -       Object
'           GetPipeRadiusDBU        -       Returns a PipeRadius
'*************************************************************************************

Public Function GetPipeRadiusDBU(ByVal pDispInputConfigHlpr As Object)
    
    Const METHOD = "GetPipeRadiusDBU"
    On Error GoTo ErrorHandler
    
   'Get IJHgrInputConfig Hlpr Interface off of passed Helper
    Dim my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr
    Set my_IJHgrInputConfigHlpr = pDispInputConfigHlpr
    
    Dim PipeRadius As Double
    Dim strUnit As String
    PipeRadius = (my_IJHgrInputConfigHlpr.PrimaryPipeDiameter(strUnit)) / 2
    
    Dim oUOMService         As IJUomVBInterface
    Dim unitID              As Units
    Set oUOMService = New UnitsOfMeasureServicesLib.UomVBInterface
    unitID = oUOMService.GetUnitId(UNIT_DISTANCE, strUnit)
    GetPipeRadiusDBU = oUOMService.ConvertUnitToDbu(UNIT_DISTANCE, PipeRadius, unitID)

    Set oUOMService = Nothing
    Set my_IJHgrInputConfigHlpr = Nothing
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:   ConnectPartToRouteOrStruct
' Desc:     This function populates the collection based on the passed indicies.  Note
'           that this is only useful for single part-pipe/struct relations.
'
' Param:
'           PartIndex                    -       Integer
'           refIndex                     -       Integer
'           ConnectPartToRouteOrStruct   -  Returns a Collection of Route/Struct Connection information
'*************************************************************************************

Public Function ConnectPartToRouteOrStruct(ByVal PartIndex As Integer, ByVal refIndex As Integer) As Collection

    Const METHOD = "ConnectPartToRouteOrStruct"
    

    'Create a collection to hold the ALL Struct connection information
    Dim ConnInfoCollection As New Collection
    
    'Create a SAFEARRAY to hold Structure/route Connection information
    Dim ConnectionInfo(2) As Integer
    
    On Error GoTo ErrorHandler
    ConnectionInfo(1) = PartIndex
    ConnectionInfo(2) = refIndex
    
    'Add the connections to the Collection of Connections.
    ConnInfoCollection.Add ConnectionInfo
    
    'Return the collection of Route/struct connection information.
    Set ConnectPartToRouteOrStruct = ConnInfoCollection
    Set ConnInfoCollection = Nothing
Exit Function

ErrorHandler:
  Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

Public Function Rad(DegNumber As Double) As Double
    Rad = DegNumber / 180 * 3.14159265358979
End Function

'*************************************************************************************
'Method:   GetUserDouble
'
'Desc:
'
'Param:
'
'Inputs:
'           PartIndex (Integer)
'           ColumnName (String)
'           IJElements_PartOccCollection (IJElements)
'Output:
'           Double
'*************************************************************************************
Public Function GetUserDouble(PartIndex As Integer, ColumnName As String, IJElements_PartOccCollection As IJElements, my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr) As Double
    
    Const METHOD = "GetUserDouble"
    'Gets table information from component xls (not auxiliary)
    Dim oPartName As New PartOcc
    Dim oPart As IJDPart
    
    On Error GoTo ErrorHandler
    Set oPartName = IJElements_PartOccCollection.Item(PartIndex)
    oPartName.GetPart oPart
    
    ' Get attributes from part to calculate band material length
    Dim VarAttribValue As Variant

    my_IJHgrInputConfigHlpr.GetAttributeValue ColumnName, oPart, VarAttribValue
    GetUserDouble = VarAttribValue
    Set oPart = Nothing
    Set oPartName = Nothing
 Exit Function

ErrorHandler:
   Err.Raise LogError(Err, MODULE, METHOD, "").Number
 
End Function

'*************************************************************************************
'Method:   GetUserString
'
'Desc:
'
'Param:
'
'Inputs:
'           PartIndex (Integer)
'           ColumnName (String)
'           IJElements_PartOccCollection (IJElements)
'Output:
'           String
'*************************************************************************************

Public Function GetUserString(PartIndex As Integer, ColumnName As String, IJElements_PartOccCollection As IJElements, my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr) As String
   
    Const METHOD = "GetUserString"
    On Error GoTo ErrorHandler
    'Gets table information from component xls (not auxiliary)
    Dim oPartName As New PartOcc
    Dim oPart As IJDPart
    Set oPartName = IJElements_PartOccCollection.Item(PartIndex)
    oPartName.GetPart oPart
    
    ' Get attributes from part to calculate band material length
    Dim VarAttribValue As Variant

    my_IJHgrInputConfigHlpr.GetAttributeValue ColumnName, oPart, VarAttribValue
    GetUserString = VarAttribValue
    Set oPartName = Nothing
    Set oPart = Nothing
 Exit Function
ErrorHandler:
   Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
'Method:   GetProfileDim
'
'Desc:
'
'Param:
'
'Inputs:
'           idxProfilePart (Integer)
'           profileDim (String)
'           IJElements_PartOccCollection (IJElements)
'Output:
'           Double
'*************************************************************************************

Public Function GetProfileDim(idxProfilePart As Integer, profileDim As String, IJElements_PartOccCollection As IJElements) As Double
    'Use this function to get the profile dimensions of hanger beam
    'Ippili Ravi
    Const METHOD = "GetProfileDim"
    On Error GoTo ErrorHandler
    Dim PCSDimen() As Double
    PCSDimen = GetAngleCSDimension(IJElements_PartOccCollection.Item(idxProfilePart))
    
    If UCase(profileDim) = "WEB" Then
        GetProfileDim = PCSDimen(4)
    ElseIf UCase(profileDim) = "DEPTH" Then
        GetProfileDim = PCSDimen(2)
    ElseIf UCase(profileDim) = "FLANGE" Then
        GetProfileDim = PCSDimen(3)
    ElseIf UCase(profileDim) = "BACKTOBACK" Then
        GetProfileDim = PCSDimen(5)
    Else
        GetProfileDim = PCSDimen(1)
    End If
Exit Function

ErrorHandler:
   Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
'Method:   GetHSSDim
'
'Desc:
'
'Param:
'
'Inputs:
'           idxProfilePart (Integer)
'           profileDim (String)
'           IJElements_PartOccCollection (IJElements)
'Output:
'           Double
'*************************************************************************************

Public Function GetHSSDim(idxProfilePart As Integer, profileDim As String, IJElements_PartOccCollection As IJElements) As Double
    Const METHOD = "GetHSSDim"
    On Error GoTo ErrorHandler
    Dim PCSDimen() As Double
    PCSDimen = CrossSectionAttributeValue(IJElements_PartOccCollection.Item(idxProfilePart))
    
    If UCase(profileDim) = "WEB" Then
        GetHSSDim = PCSDimen(4)
    ElseIf UCase(profileDim) = "DEPTH" Then
        GetHSSDim = PCSDimen(2)
    ElseIf UCase(profileDim) = "FLANGE" Then
        GetHSSDim = PCSDimen(3)
    ElseIf UCase(profileDim) = "TNOM" Then
        GetHSSDim = PCSDimen(3)
    Else
        GetHSSDim = PCSDimen(1)
    End If
 Exit Function
ErrorHandler:
  Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:   SetValuesOfSymbolInputsAndOccAttributes
'
'
' Desc:
'
' Param:
'
'Inputs:
'           my_IJHgrInputConfigHlpr      -    Object
'           IJElements_PartOccCollection -    IJElements
'           iArrayOfKeys                 -    String
'           vInputValue                  -    Variant
'
'*************************************************************************************

Public Sub SetValuesOfSymbolInputsAndOccAttributes(my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr, IJElements_PartOccCollection As IJElements, iArrayOfKeys() As Integer, strAttributeName As String, vInputValue As Variant)
    Const METHOD = "SetValuesOfSymbolInputsAndOccAttributes"
    On Error GoTo ErrorHandler
    
    Dim iLowerBound As Integer
    Dim iUpperBound As Integer
    Dim i As Integer
            
    On Error Resume Next ' since in the collection, some of the Occs may not be exposing the property which we are trying to set
    
    iLowerBound = LBound(iArrayOfKeys)
    iUpperBound = UBound(iArrayOfKeys)
    If Not (iUpperBound < iLowerBound) And Not (my_IJHgrInputConfigHlpr Is Nothing) And Not (IJElements_PartOccCollection Is Nothing) Then
        For i = iLowerBound To iUpperBound
            my_IJHgrInputConfigHlpr.SetSymbolInputByName IJElements_PartOccCollection.Item(iArrayOfKeys(i)), strAttributeName, vInputValue
        Next i
    End If

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub

'*************************************************************************************
' Method:   GetLAttributeFromHelper
' Desc:      Added this function to remove redundant code
'
'
' Param:
'       ' Takes the InputHelper and Attribute name and returns the DOUBLE value
'*************************************************************************************

Public Function GetNAttributeFromHelper(ByRef my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr, ByVal AttributeName As String) As Double

Const METHOD = "GetNAttributeFromHelper"
On Error GoTo ErrorHandler

    Dim varTemp As Variant

    my_IJHgrInputConfigHlpr.GetAttributeValue AttributeName, Nothing, varTemp
    
    GetNAttributeFromHelper = varTemp
    Exit Function
    
ErrorHandler:
        Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:   GetLAttributeFromHelper
' Desc:      Added this function to remove redundant code
'
'
' Param:
'        ' Takes the InputHelper and Attribute name and returns the LONG value
'
'*************************************************************************************

Public Function GetLAttributeFromHelper(ByRef my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr, ByVal AttributeName As String) As Long

Const METHOD = "GetLAttributeFromHelper"
On Error GoTo ErrorHandler

    Dim varTemp As Variant

    my_IJHgrInputConfigHlpr.GetAttributeValue AttributeName, Nothing, varTemp
    
    GetLAttributeFromHelper = varTemp
    Exit Function
    
ErrorHandler:
        Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:   GetSAttributeFromHelper
' Desc:      Added this function to remove redundant code
'
'
' Param:
'         Takes the InputHelper and Attribute name and returns the STRING value
'*************************************************************************************

Public Function GetSAttributeFromHelper(ByRef my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr, ByVal AttributeName As String) As String


Const METHOD = "GetSAttributeFromHelper"
On Error GoTo ErrorHandler

    Dim varTemp As Variant

    my_IJHgrInputConfigHlpr.GetAttributeValue AttributeName, Nothing, varTemp

    GetSAttributeFromHelper = varTemp
    Exit Function
    
ErrorHandler:
        Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:   ProcCreateNote
' Desc:     Add a note to a part in an assembly
' see Assy_FR_IT_LS for an example of the use of this function
'
' Param:
'          sblOcc   : the part that needs the note(Object)
'          pICH     : pointer to input config helper object(IJHgrInputConfigHlpr)
'          sNoteText: the text of the note(String)
'*************************************************************************************


Public Sub ProcCreateNote(ByVal sblOcc As Object, pICH As IJHgrInputConfigHlpr, sNoteText As String)
Const METHOD = "ProcCreateNote"
On Error GoTo ErrHandler

    ' Create a new GeneralNote
    Dim oNote As IJGeneralNote
    Set oNote = CreateGenNote(sblOcc)
    If oNote Is Nothing Then
        Exit Sub
    End If

    oNote.Text = sNoteText
    oNote.Dimensioned = False
    ' The delivered HgrSup_Note label expects notes named Note 1, with a purpose of Fabrication
    oNote.Name = "Note 1"   ' set fall thru value
    oNote.Purpose = 3       ' set fall thru value to 3 = Fabrication
    If Not pICH Is Nothing Then
        Dim vData() As Variant
        On Error GoTo FallThru 'in case the rule is not defined
        vData = pICH.GetDataByRule("HgrSupNotePurpose")  ' get the note purpose from the rule
        Dim lPurpose As Long
        lPurpose = vData(1)
        oNote.Purpose = lPurpose
        vData = pICH.GetDataByRule("HgrSupNoteName")  ' get the note name from the rule
        Dim sName As String
        sName = vData(1)
        oNote.Name = sName
        On Error GoTo ErrHandler
    End If
    
FallThru:
    Exit Sub
    
ErrHandler:
        Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub

'*************************************************************************************
' Method:   CreateGenNote
' Desc:     'CreateGenNote
'            Helper function of ProcCreateNote
'            sblOcc: the part that the note should be created on
'
' Param:
'           sblOcc                       -    Input as  Object
'           CreateGenNote                -    Output as IJGeneralNote
'*************************************************************************************

Private Function CreateGenNote(ByVal sblOcc As Object) As IJGeneralNote
Const METHOD = "CreateGenNote"
On Error GoTo errHndlr
    
    Dim oNoteFac As New GeneralNoteFactory
    'Dim oWorkingSet As IJDWorkingSet
    'Dim oActiveConnection As IJDConnection
    'Dim oTrader As New Trader
    Dim oResMgr As Object
    
    If Not sblOcc Is Nothing Then
        'Set oWorkingSet = oTrader.Service("WorkingSet", "")
        'Set oActiveConnection = oWorkingSet.ActiveConnection
        'Set CreateGenNote = oNoteFac.CreateGeneralNote(oActiveConnection.ResourceManager, sblOcc)
        Set oResMgr = GetModelResourceManager
        Set CreateGenNote = oNoteFac.CreateGeneralNote(oResMgr, sblOcc)
    End If
    
    Exit Function

errHndlr:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

'*************************************************************************************
' Method:   GetPipeODwithoutInsulation
' Desc:     'GetPipeODwithoutInsulation
'            Function to get PipeOD without insulation - Takes dTotalOD as double
'            oRoute as Object as Input and returns the PipeOD without insulation
'
' Param:
'           dTotalOD                       -    dTotalOD as  double - input
'           oRoute                         -    oRoute as Object - input
'           Ouput                          -    returns the PipeOD without insulation
'*************************************************************************************
' DI#109789 ; commented during merging of files
'Public Function GetPipeODwithoutInsulation(dTotalOD As Double, oRoute As Object) As Double
'    Dim oInsulationObject As IJRteInsulation
'    Set oInsulationObject = oRoute
'
'    Dim dInsulat As Double
'    dInsulat = oInsulationObject.Thickness
'
'    GetPipeODwithoutInsulation = dTotalOD - 2 * dInsulat
'    Set oInsulationObject = Nothing
'
'End Function


'*************************************************************************************
' Method:   IsAttributeExists
' Desc:     Get the attribute value from an Object.
' Param:
'           pObject As Object
'               The Object  in which the attribute value to be get
'           AttributeName As String
'               The name of the attribute for which the value to be get
' Return:
'           True - If the attribute exists on the Object
'           False- If the attribute doesn't exist on the Object
' Common Use:
'            - used in 'GetAttr' function to find the Attribute is exists on the object or not
'               hsHlpr.IsAttributeExists(oPartOcc,"Length")
'               hsHlpr.IsAttributeExists(oPart,"Length")
'*************************************************************************************
Public Function IsAttributeExists(pObject As Object, AttributeName As String) As Boolean
Const METHOD = "GetAttributeFromObject"
Dim iErrCodeToUse As Integer
On Error GoTo ErrHandler

    Dim oAttributes             As IMSAttributes.IJDAttributes
    Dim oAttributesCol          As IJDAttributesCol
    Dim oAttribute              As IJDAttribute
    Dim oAttributeInfo          As IJDAttributeInfo
    Dim iID                     As Variant
    Dim iValue                  As Integer
    iValue = 0

    Set oAttributes = pObject
    For Each iID In oAttributes
        Set oAttributesCol = oAttributes.CollectionOfAttributes(iID)
        If oAttributesCol.Count > 0 Then
            For Each oAttribute In oAttributesCol
                Set oAttributeInfo = oAttribute.AttributeInfo
                If UCase(oAttributeInfo.Name) = UCase(AttributeName) Then
                    IsAttributeExists = True
                    'cleanup
                    Set oAttributeInfo = Nothing
                    Set oAttribute = Nothing
                    Set oAttributesCol = Nothing
                    Exit Function
                End If
                Set oAttributeInfo = Nothing
                Set oAttribute = Nothing
            Next
        End If
        Set oAttributesCol = Nothing
    Next iID
    
    IsAttributeExists = False
Exit Function
ErrHandler:
    Dim sErr As String
    Select Case iErrCodeToUse
        Case 0
            sErr = ""
    End Select
    Err.Raise LogError(Err, MODULE, METHOD, sErr).Number
End Function
