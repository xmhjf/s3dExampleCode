Imports System
Imports System.Collections.Generic
Imports Ingr.SP3D.Route.Middle
Imports System.Collections
Imports Ingr.SP3D.Common.Middle
Imports System.Data.SqlClient
Imports System.Data.OracleClient
Imports Excel = Microsoft.Office.Interop.Excel


Public Class CableVolumeRule
    Inherits CARVolumeRule

    Dim m_aVolumeZoneOID(,,) As String
    Dim m_iAvailableFloor As Integer = -1
    Dim m_oSignalTypeDeck As New Dictionary(Of String, Integer)

    '/// <summary>
    '/// Constructor of CableVolumeRule.
    '/// </summary>
    Public Sub CableVolumeRule()

    End Sub

    ''' <summary>
    ''' This Rule Method that is called .. When Rule is Enabled.
    ''' </summary>
    ''' <param name="cableRun"></param>
    ''' <param name="redundantCableColl"></param>
    ''' <param name="avoidancePlnColl"></param>
    ''' <param name="avoidanceVolColl"></param>
    ''' <remarks></remarks>
    Public Overrides Sub CARVolumeRuleMethod(ByVal cableRun As Ingr.SP3D.Route.Middle.CableRun, ByVal redundantCableColl As System.Collections.ObjectModel.Collection(Of Ingr.SP3D.Common.Middle.BusinessObject), ByVal avoidancePlnColl As System.Collections.ObjectModel.Collection(Of Ingr.SP3D.Common.Middle.BusinessObject), ByVal avoidanceVolColl As System.Collections.ObjectModel.Collection(Of Ingr.SP3D.Common.Middle.BusinessObject))
        Try

            'This to read Excel and get the avialable Floor for Corresponding Signal Type of the Cable Run
            'Remove the below Code if reading from Excel is not required
            'ReadExcelIntoDictonary()
            m_iAvailableFloor = FindAllowedFloorPerRule(cableRun)


            Dim oOrigVolColl As System.Collections.ObjectModel.Collection(Of Ingr.SP3D.Common.Middle.BusinessObject) = CARVolZoneHelper.GetVolumesGivenAnEquipment(CType(cableRun.OriginatingDevice, Ingr.SP3D.Common.Middle.BusinessObject), RteSpaceVolumeType.RTE_VOLUMETYPE_ZONE)
            Dim oTermVolColl As System.Collections.ObjectModel.Collection(Of Ingr.SP3D.Common.Middle.BusinessObject) = CARVolZoneHelper.GetVolumesGivenAnEquipment(CType(cableRun.TerminatingDevice, Ingr.SP3D.Common.Middle.BusinessObject), RteSpaceVolumeType.RTE_VOLUMETYPE_ZONE)
            Dim oCableRunLV As CableRun = cableRun

            Dim oOrigVolLot As Ingr.SP3D.Common.Middle.BusinessObject = Nothing
            Dim oDestVolLot As Ingr.SP3D.Common.Middle.BusinessObject = Nothing


            Dim oCableRunBussinessObj As Ingr.SP3D.Common.Middle.BusinessObject = cableRun
            Dim oSp3dConn As Ingr.SP3D.Common.Middle.Services.SP3DConnection = cableRun.DBConnection

            Dim oTempProp As PropertyValue
            Dim oTempCLVal As String

            Dim iXmax As Integer = 0
            Dim iXmin As Integer = 0
            Dim iYmax As Integer = 0
            Dim iYmin As Integer = 0
            Dim iZmax As Integer = 0
            Dim iZmin As Integer = 0

            Dim iOrigZoneIdx As Integer = 0
            Dim iOrigSideIdx As Integer = 0
            Dim iOrigDeckIdx As Integer = 0
            Dim iDestZoneIdx As Integer = 0
            Dim iDestSideIdx As Integer = 0
            Dim iDestDeckIdx As Integer = 0

            ' This code commented below is for DuctBankCables
            'Dim bDuctBankCables As Boolean = False

            'Below 2 For Loops is to get Volumes in which the Orig and Term Equipment is Present.
            For Each oBo As Ingr.SP3D.Common.Middle.BusinessObject In oOrigVolColl


                oTempProp = oBo.GetPropertyValue("IJUAPlantVolumeAttribute", "VolumeType")

                oTempCLVal = oTempProp.ToString()

                If (oTempCLVal = "LotZone") Then
                    oTempProp = oBo.GetPropertyValue("IJUAPlantVolumeAttribute", "SideIndex")
                    'iXmin = oTempProp.PropertyInfo.CodeListInfo.GetCodelistItem(oTempProp.ToString()).Value
                    iXmin = Convert.ToInt32(oTempProp.ToString())

                    ' In this Case it's a CodelistValue so get the long value
                    oTempProp = oBo.GetPropertyValue("IJUAPlantVolumeAttribute", "CompartmentZoneIndex")
                    'Ymin = Convert.ToInt32(oTempProp.ToString())
                    iYmin = Convert.ToInt32(oTempProp.ToString())

                    ' In this Case it's a CodelistValue so get the long value
                    oTempProp = oBo.GetPropertyValue("IJUAPlantVolumeAttribute", "FloorLevelIndex")
                    iZmin = Convert.ToInt32(oTempProp.ToString())

                    oOrigVolLot = oBo
                End If

            Next

            For Each oBo As Ingr.SP3D.Common.Middle.BusinessObject In oTermVolColl
                oTempProp = oBo.GetPropertyValue("IJUAPlantVolumeAttribute", "VolumeType")

                oTempCLVal = oTempProp.ToString()

                If (oTempCLVal = "LotZone") Then
                    oTempProp = oBo.GetPropertyValue("IJUAPlantVolumeAttribute", "SideIndex")
                    iXmax = Convert.ToInt32(oTempProp.ToString())

                    oTempProp = oBo.GetPropertyValue("IJUAPlantVolumeAttribute", "CompartmentZoneIndex")
                    'Ymax = Convert.ToInt32(oTempProp.ToString())
                    iYmax = Convert.ToInt32(oTempProp.ToString())

                    oTempProp = oBo.GetPropertyValue("IJUAPlantVolumeAttribute", "FloorLevelIndex")
                    iZmax = Convert.ToInt32(oTempProp.ToString())

                    oDestVolLot = oBo
                End If

            Next

            If (iXmax < iXmin) Then
                iXmax += iXmin
                iXmin = iXmax - iXmin
                iXmax -= iXmin
            End If

            If (iYmax < iYmin) Then
                iYmax += iYmin
                iYmin = iYmax - iYmin
                iYmax -= iYmin
            End If

            If (iZmax < iZmin) Then
                iZmax += iZmin
                iZmin = iZmax - iZmin
                iZmax -= iZmin
            End If

            Dim StrQuery As String

            ' This code commented below is for DuctBankCables.To make the below code work. Have to add the  IJUACableVolumeAttributes with CableRegulation attribute on it.
            'If (iZmax = 1 And iZmin = 1) Then
            '    oTempProp = oCableRunBussinessObj.GetPropertyValue("IJUACableVolumeAttributes", "CableRegulation")
            '    oTempCLVal = oTempProp.ToString()
            '    If (oTempCLVal = "DUCTBANKCABLES") Then bDuctBankCables = True
            '    'Ymin = oTempProp.PropertyInfo.CodeListInfo.GetCodelistItem(oTempProp.ToString()).Value
            '    'End If
            'End If

            'The Query Gets OID's of Zone That will be further processed to Get Avoidance Volumes And Avoidance Planes....
            'This Query Also takes care of Getting the OIDs In Order So that it is Added in 3D array Perfectly(it's a very Crude approach and if this don't work try a differennt approach) 
            'Along X will be SideIndex and Along Y will be CompartmentZoneIndex and Along Z will be FloorLevelIndex 
            ' This code commented below is for DuctBankCables
            'If (bDuctBankCables) Then
            'StrQuery = "select oid from JUAPlantVolumeAttribute where ( FloorLevelIndex = " & 1 & " ) and VolumeType = " & 15 & " order by SideIndex ,CompartmentZoneIndex ,FloorLevelIndex "
            'Else
            StrQuery = "select oid from JUAPlantVolumeAttribute where ( SideIndex >=" & iXmin & " and SideIndex <= " & iXmax & " and CompartmentZoneIndex >= " & iYmin & " and CompartmentZoneIndex <= " & iYmax & " and FloorLevelIndex >= " & iZmin & " and FloorLevelIndex <= " & iZmax & " ) and VolumeType = " & 15 & " order by SideIndex ,CompartmentZoneIndex ,FloorLevelIndex "
            'End If
            Dim eDBProvider As Ingr.SP3D.Common.Middle.Services.SiteManager.eDBProviderTypes
            If (cableRun.DBConnection.DBProvider = "MSSQL") Then
                eDBProvider = Ingr.SP3D.Common.Middle.Services.SiteManager.eDBProviderTypes.MSSQL
            Else
                eDBProvider = Ingr.SP3D.Common.Middle.Services.SiteManager.eDBProviderTypes.Oracle
            End If

            Dim oDT As DataTable = RunSelectQuery(cableRun.DBConnection.Server, cableRun.DBConnection.Name, StrQuery, eDBProvider) 'Ingr.SP3D.Common.Middle.Services.SiteManager.eDBProviderTypes.MSSQL)

            Dim currentRows As DataRow() = oDT.Select()

            Dim OIDindex As Integer = 0

            If currentRows.Length() = 0 Then Return

            'This code commented below is for DuctBankCables
            'If (bDuctBankCables) Then
            '    For Each CurrentRow As DataRow In currentRows
            '        Dim strOId As String

            '        If eDBProvider = Ingr.SP3D.Common.Middle.Services.SiteManager.eDBProviderTypes.MSSQL Then
            '            strOId = (CType(CurrentRow("oid"), Guid)).ToString("B").ToUpper()
            '        Else
            '            Dim oGuidTemp As Guid
            '            Dim TempByteArr As Byte()
            '            TempByteArr = CType(CurrentRow("oid"), Byte())
            '            oGuidTemp = New Guid(BitConverter.ToInt32(New Byte() {TempByteArr(3), TempByteArr(2), TempByteArr(1), TempByteArr(0)}, 0), BitConverter.ToInt16(TempByteArr, 4), BitConverter.ToInt16(TempByteArr, 6), TempByteArr(8), _
            '                                 TempByteArr(9), TempByteArr(10), TempByteArr(11), TempByteArr(12), TempByteArr(13), TempByteArr(14), TempByteArr(15))
            '            strOId = oGuidTemp.ToString("N").ToUpper()
            '        End If

            '        If Not (oDestVolLot.ObjectIDForQuery = strOId Or oOrigVolLot.ObjectIDForQuery = strOId) Then
            '            Dim oBOMon As Ingr.SP3D.Common.Middle.Services.BOMoniker
            '            Dim oZone As BusinessObject

            '            oBOMon = oSp3dConn.GetBOMonikerFromDbIdentifier(CType(strOId, String))
            '            oZone = oSp3dConn.WrapSP3DBO(oBOMon)
            '            AvoidanceVolColl.Add(oZone)
            '        End If
            '    Next
            '    Return
            'End If

            ' redefining the Volume array in which the OID's of Zone is Mapped... in 3D ...
            'Volume Array is only gets all the OID's which Exists Between the Equipments... Change the Aproach if all other Vol OIDs are also required.
            ReDim m_aVolumeZoneOID(iXmax - (iXmin), iYmax - (iYmin), iZmax - (iZmin))

            For x As Integer = 0 To m_aVolumeZoneOID.GetUpperBound(0)

                For y As Integer = 0 To m_aVolumeZoneOID.GetUpperBound(1)

                    For z As Integer = 0 To m_aVolumeZoneOID.GetUpperBound(2)
                        Dim CurrentRow As DataRow

                        CurrentRow = currentRows(OIDindex)
                        If eDBProvider = Ingr.SP3D.Common.Middle.Services.SiteManager.eDBProviderTypes.MSSQL Then
                            m_aVolumeZoneOID(x, y, z) = (CType(CurrentRow("oid"), Guid)).ToString("B").ToUpper()
                        Else
                            Dim oGuidTemp As Guid
                            Dim TempByteArr As Byte()
                            TempByteArr = CType(CurrentRow("oid"), Byte())
                            oGuidTemp = New Guid(BitConverter.ToInt32(New Byte() {TempByteArr(3), TempByteArr(2), TempByteArr(1), TempByteArr(0)}, 0), BitConverter.ToInt16(TempByteArr, 4), BitConverter.ToInt16(TempByteArr, 6), TempByteArr(8), _
                                                 TempByteArr(9), TempByteArr(10), TempByteArr(11), TempByteArr(12), TempByteArr(13), TempByteArr(14), TempByteArr(15))
                            m_aVolumeZoneOID(x, y, z) = oGuidTemp.ToString("N").ToUpper()
                        End If
                        OIDindex = OIDindex + 1
                    Next
                Next
            Next

            'The below Code is to Get position of Orig and term Eq Contained Volume in Volume Array ... Usually it's the Ends in the Volume Array...
            For x As Integer = 0 To m_aVolumeZoneOID.GetUpperBound(0)

                For y As Integer = 0 To m_aVolumeZoneOID.GetUpperBound(1)

                    For z As Integer = 0 To m_aVolumeZoneOID.GetUpperBound(2)

                        If (m_aVolumeZoneOID(x, y, z) = oOrigVolLot.ObjectIDForQuery) Then

                            iOrigSideIdx = x
                            iOrigZoneIdx = y
                            iOrigDeckIdx = z
                        End If

                        If (m_aVolumeZoneOID(x, y, z) = oDestVolLot.ObjectIDForQuery) Then

                            iDestSideIdx = x
                            iDestZoneIdx = y
                            iDestDeckIdx = z
                        End If

                    Next
                Next
            Next

            'This if Condition is satisfied when both the Equipments are in Same Zone ... and then the Cable will be retricted to This Zone Only...
            If iOrigZoneIdx = iDestZoneIdx Then

                'added the code for the lots being in the FireZone
                For x As Integer = 0 To m_aVolumeZoneOID.GetUpperBound(0)

                    For z As Integer = 0 To m_aVolumeZoneOID.GetUpperBound(2)

                        Dim eAvoidancePlanes As RteVolumePlanes = RteVolumePlanes.None
                        Dim oZone As BusinessObject = Nothing

                        eAvoidancePlanes = eAvoidancePlanes Or RteVolumePlanes.YNegative_Pln Or RteVolumePlanes.YPositive_Pln

                        If x = 0 Then eAvoidancePlanes = eAvoidancePlanes Or RteVolumePlanes.XNegative_Pln
                        If x = (m_aVolumeZoneOID.GetLength(0) - 1) Then eAvoidancePlanes = eAvoidancePlanes Or RteVolumePlanes.XPositive_Pln

                        If z = 0 Then eAvoidancePlanes = eAvoidancePlanes Or RteVolumePlanes.ZNegative_Pln

                        If z = (m_aVolumeZoneOID.GetLength(2) - 1) Then eAvoidancePlanes = eAvoidancePlanes Or RteVolumePlanes.ZPositive_Pln

                        Dim oBOMon As Ingr.SP3D.Common.Middle.Services.BOMoniker
                        oBOMon = oSp3dConn.GetBOMonikerFromDbIdentifier(CType(m_aVolumeZoneOID(x, iOrigZoneIdx, z), String))
                        oZone = oSp3dConn.WrapSP3DBO(oBOMon)

                        Dim oTempColl As System.Collections.ObjectModel.Collection(Of Ingr.SP3D.Common.Middle.BusinessObject) = CARVolZoneHelper.GetPlaneCollFromVolume(oZone, eAvoidancePlanes)
                        'AvoidancePlnColl = (Collection<BusinessObject>)(AvoidancePlnColl.Concat(oTempColl))
                        For Each oBo As Ingr.SP3D.Common.Middle.BusinessObject In oTempColl
                            avoidancePlnColl.Add(oBo)
                        Next
                    Next
                Next
                Return
            End If

            Dim bAvailbleDeckExist As Boolean = False

            'The Below Code is written ... if there is an avalaible Floor present and only through which the cable should pass ... 
            If (iZmax >= m_iAvailableFloor And iZmin <= m_iAvailableFloor) Then bAvailbleDeckExist = True

            If (bAvailbleDeckExist) Then

                Dim ZoneMin As Integer = 0
                Dim ZoneMax As Integer = 0

                If (iOrigZoneIdx > iDestZoneIdx) Then
                    ZoneMax = iOrigZoneIdx
                    ZoneMin = iDestZoneIdx

                Else
                    ZoneMax = iDestZoneIdx
                    ZoneMin = iOrigZoneIdx
                End If

                'This loop sets the Z planes Available Deck as avoided so that it doesn't pass to another deck when not in the Eq Zone....
                For z As Integer = 0 To m_aVolumeZoneOID.GetUpperBound(2)
                    If ((z + iZmin) = m_iAvailableFloor) Then

                        For x As Integer = 0 To (m_aVolumeZoneOID.GetUpperBound(0))

                            For y As Integer = 0 To m_aVolumeZoneOID.GetUpperBound(1)

                                If (y = iOrigZoneIdx Or y = iDestZoneIdx) Then Continue For

                                Dim eAvoidancePlanes As RteVolumePlanes = RteVolumePlanes.None
                                Dim oZone As BusinessObject = Nothing

                                eAvoidancePlanes = eAvoidancePlanes Or RteVolumePlanes.ZNegative_Pln Or RteVolumePlanes.ZPositive_Pln

                                Dim oBOMon As Ingr.SP3D.Common.Middle.Services.BOMoniker
                                oBOMon = oSp3dConn.GetBOMonikerFromDbIdentifier(CType(m_aVolumeZoneOID(x, y, z), String))
                                oZone = oSp3dConn.WrapSP3DBO(oBOMon)

                                Dim oTempColl As System.Collections.ObjectModel.Collection(Of Ingr.SP3D.Common.Middle.BusinessObject) = CARVolZoneHelper.GetPlaneCollFromVolume(oZone, eAvoidancePlanes)
                                For Each oBo1 As Ingr.SP3D.Common.Middle.BusinessObject In oTempColl

                                    avoidancePlnColl.Add(oBo1)
                                Next
                            Next
                        Next
                    End If
                Next

                'This code Is written for Eq Zones to Work
                For x As Integer = 0 To m_aVolumeZoneOID.GetUpperBound(0)

                    For z As Integer = 0 To m_aVolumeZoneOID.GetUpperBound(2)

                        Dim eAvoidancePlanes As RteVolumePlanes = RteVolumePlanes.None
                        Dim oZone As BusinessObject = Nothing

                        If (Not ((z + iZmin) = m_iAvailableFloor)) Then

                            eAvoidancePlanes = eAvoidancePlanes Or RteVolumePlanes.YNegative_Pln Or RteVolumePlanes.YPositive_Pln
                        End If

                        If x = 0 Then eAvoidancePlanes = eAvoidancePlanes Or RteVolumePlanes.XNegative_Pln

                        If (x = (m_aVolumeZoneOID.GetLength(0) - 1)) Then eAvoidancePlanes = eAvoidancePlanes Or RteVolumePlanes.XPositive_Pln

                        If z = 0 Then eAvoidancePlanes = eAvoidancePlanes Or RteVolumePlanes.ZNegative_Pln

                        If (z = (m_aVolumeZoneOID.GetLength(2) - 1)) Then eAvoidancePlanes = eAvoidancePlanes Or RteVolumePlanes.ZPositive_Pln

                        Dim oBOMon As Ingr.SP3D.Common.Middle.Services.BOMoniker
                        Dim oTempColl As System.Collections.ObjectModel.Collection(Of Ingr.SP3D.Common.Middle.BusinessObject)

                        'For the first Zone.
                        oBOMon = oSp3dConn.GetBOMonikerFromDbIdentifier(CType(m_aVolumeZoneOID(x, ZoneMin, z), String))
                        oZone = oSp3dConn.WrapSP3DBO(oBOMon)

                        oTempColl = CARVolZoneHelper.GetPlaneCollFromVolume(oZone, eAvoidancePlanes Or RteVolumePlanes.YNegative_Pln)
                        'AvoidancePlnColl = (Collection<BusinessObject>)(AvoidancePlnColl.Concat(oTempColl))
                        For Each oBo As Ingr.SP3D.Common.Middle.BusinessObject In oTempColl
                            avoidancePlnColl.Add(oBo)
                        Next

                        'For the Second Zone.
                        oBOMon = oSp3dConn.GetBOMonikerFromDbIdentifier(CType(m_aVolumeZoneOID(x, ZoneMax, z), String))
                        oZone = oSp3dConn.WrapSP3DBO(oBOMon)

                        oTempColl = CARVolZoneHelper.GetPlaneCollFromVolume(oZone, eAvoidancePlanes Or RteVolumePlanes.YPositive_Pln)
                        'AvoidancePlnColl = (Collection<BusinessObject>)(AvoidancePlnColl.Concat(oTempColl))
                        For Each oBo As Ingr.SP3D.Common.Middle.BusinessObject In oTempColl

                            avoidancePlnColl.Add(oBo)
                        Next
                    Next
                Next
            End If

            'This Case is for the Cables Whose Eq's Are niether in the Same Zone nor an available floor is present 
            If Not (bAvailbleDeckExist) Then
                For x = 0 To m_aVolumeZoneOID.GetUpperBound(0)

                    Dim eAvoidancePlanes As RteVolumePlanes
                    eAvoidancePlanes = RteVolumePlanes.None

                    'The below If conditions are to avoid the boundary planes... and the else condition is to clear off the flag if it is set before 
                    If (x = 0) Then
                        eAvoidancePlanes = eAvoidancePlanes Or RteVolumePlanes.XNegative_Pln
                    Else
                        eAvoidancePlanes = eAvoidancePlanes And Not (RteVolumePlanes.XNegative_Pln)
                    End If

                    If (x = (m_aVolumeZoneOID.GetLength(0) - 1)) Then
                        eAvoidancePlanes = eAvoidancePlanes Or RteVolumePlanes.XPositive_Pln
                    Else
                        eAvoidancePlanes = eAvoidancePlanes And Not (RteVolumePlanes.XPositive_Pln)
                    End If

                    For y = 0 To m_aVolumeZoneOID.GetUpperBound(1)

                        'The below If conditions are to avoid the boundary planes... and the else condition is to clear off the flag if it is set before 
                        If (y = 0) Then
                            eAvoidancePlanes = eAvoidancePlanes Or RteVolumePlanes.YNegative_Pln
                        Else
                            eAvoidancePlanes = eAvoidancePlanes And Not (RteVolumePlanes.YNegative_Pln)
                        End If

                        If (y = (m_aVolumeZoneOID.GetLength(1) - 1)) Then
                            eAvoidancePlanes = eAvoidancePlanes Or RteVolumePlanes.YPositive_Pln
                        Else
                            eAvoidancePlanes = eAvoidancePlanes And Not (RteVolumePlanes.YPositive_Pln)
                        End If

                        For z = 0 To m_aVolumeZoneOID.GetUpperBound(2)

                            'The below If conditions are to avoid the boundary planes... and the else condition is to clear off the flag if it is set before 
                            If (z = 0) Then
                                eAvoidancePlanes = eAvoidancePlanes Or RteVolumePlanes.ZNegative_Pln
                            Else
                                eAvoidancePlanes = eAvoidancePlanes And Not (RteVolumePlanes.ZNegative_Pln)
                            End If

                            If (z = (m_aVolumeZoneOID.GetLength(2) - 1)) Then
                                eAvoidancePlanes = eAvoidancePlanes Or RteVolumePlanes.ZPositive_Pln
                            Else
                                eAvoidancePlanes = eAvoidancePlanes And Not (RteVolumePlanes.ZPositive_Pln)
                            End If

                            If (Not ((y = iOrigZoneIdx) Or (y = iDestZoneIdx))) Then
                                eAvoidancePlanes = eAvoidancePlanes Or RteVolumePlanes.ZNegative_Pln Or RteVolumePlanes.ZPositive_Pln
                            End If

                            Dim oBOMon As Ingr.SP3D.Common.Middle.Services.BOMoniker
                            Dim oTempColl As System.Collections.ObjectModel.Collection(Of Ingr.SP3D.Common.Middle.BusinessObject)
                            Dim oZone As BusinessObject

                            oBOMon = oSp3dConn.GetBOMonikerFromDbIdentifier(CType(m_aVolumeZoneOID(x, y, z), String))
                            oZone = oSp3dConn.WrapSP3DBO(oBOMon)
                            oTempColl = CARVolZoneHelper.GetPlaneCollFromVolume(oZone, eAvoidancePlanes)
                            For Each oBo As Ingr.SP3D.Common.Middle.BusinessObject In oTempColl
                                avoidancePlnColl.Add(oBo)
                            Next
                            eAvoidancePlanes = eAvoidancePlanes And Not (RteVolumePlanes.ZNegative_Pln)
                            eAvoidancePlanes = eAvoidancePlanes And Not (RteVolumePlanes.ZPositive_Pln)
                        Next
                        eAvoidancePlanes = eAvoidancePlanes And Not (RteVolumePlanes.YNegative_Pln)
                        eAvoidancePlanes = eAvoidancePlanes And Not (RteVolumePlanes.YPositive_Pln)
                    Next
                    eAvoidancePlanes = eAvoidancePlanes And Not (RteVolumePlanes.XNegative_Pln)
                    eAvoidancePlanes = eAvoidancePlanes And Not (RteVolumePlanes.XPositive_Pln)
                Next
            Else
                Return
            End If

        Catch oEx As Exception
            Dim sException As String = oEx.Message
            'If (Not oLogError Is Nothing) Then
            '    oLogError.Log(sException)
            'End If
            Throw oEx
        End Try

    End Sub

    ''' <summary>
    ''' This is a private function for Query Support i.e getting the OID's of the Volume after wich we will get Corrresponding the Object.
    ''' </summary>
    ''' <param name="strServerName"></param>
    ''' <param name="strDatabaseName"></param>
    ''' <param name="strSelectQuery"></param>
    ''' <param name="eProvider"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function RunSelectQuery(ByVal strServerName As String, ByVal strDatabaseName As String, ByVal strSelectQuery As String, ByVal eProvider As Ingr.SP3D.Common.Middle.Services.SiteManager.eDBProviderTypes) As DataTable
        '  Note the code in this method will not work if S3D is used with a database where the user only has db user access does not have windows authentication access.
        Try
            Dim item As DataTable
            If (eProvider = Ingr.SP3D.Common.Middle.Services.SiteManager.eDBProviderTypes.MSSQL) Then
                Dim oConn As SqlConnection
                Dim oConnString As String = "Data Source=" + strServerName + "; Database='" + strDatabaseName + "';Integrated Security=true"
                oConn = New SqlConnection(oConnString)
                Using (oConn)
                    Dim SqlCommand As SqlCommand
                    oConn.Open()
                    SqlCommand = oConn.CreateCommand()
                    Using (SqlCommand)

                        SqlCommand.CommandText = strSelectQuery
                        SqlCommand.CommandTimeout = 900
                        Dim SqlDataAdapter1 As SqlDataAdapter = New SqlDataAdapter()
                        SqlDataAdapter1.SelectCommand = SqlCommand
                        Dim DataSet As DataSet = New DataSet()
                        SqlDataAdapter1.Fill(DataSet, "DataTable")
                        item = DataSet.Tables("DataTable")
                    End Using
                End Using
            Else

                Dim oOraConn As OracleConnection
                'Connection string for Oracle OLEDB 
                Dim sSqlConnectionString As String = "Data Source=" + strServerName + "; " + "Integrated Security=yes;"

                oOraConn = New OracleConnection(sSqlConnectionString)

                Using (oOraConn)
                    oOraConn.Open()
                    If (Not String.IsNullOrEmpty(strDatabaseName)) Then
                        Dim alterSessionCmd As OracleCommand = oOraConn.CreateCommand()
                        alterSessionCmd.CommandText = "alter session set current_schema=" + strDatabaseName
                        alterSessionCmd.ExecuteNonQuery()

                        Dim oOracleSqlCommand As OracleCommand = oOraConn.CreateCommand()

                        Using (oOracleSqlCommand)
                            Const tableResultName As String = "DataTable"
                            oOracleSqlCommand.CommandText = strSelectQuery
                            oOracleSqlCommand.CommandTimeout = 900

                            Dim oSqlDataAdapter As OracleDataAdapter = New OracleDataAdapter()
                            oSqlDataAdapter.SelectCommand = oOracleSqlCommand
                            Dim oDataSet As DataSet = New DataSet()

                            oSqlDataAdapter.Fill(oDataSet, tableResultName)
                            item = oDataSet.Tables.Item(tableResultName)
                        End Using
                    Else
                        item = Nothing
                    End If
                End Using
            End If
            Return item
        Catch oEx As Exception
            Dim sException As String = oEx.Message
            'If (Not oLogError Is Nothing) Then
            '    oLogError.Log(sException)
            'End If
            Throw oEx
        End Try
    End Function
    ''' <summary>
    ''' This is to Read Excel cells of the Rule sheet to find Values on the cell.This is just a sample. 
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ReadExcelIntoDictonary()
        Try
            Dim xlApp As New Excel.Application
            Dim xlWorkBook As Excel.Workbook
            Dim xlWorkSheet As Excel.Worksheet



            xlWorkBook = xlApp.Workbooks.Open("c:\CableVolumes.xlsx")
            xlWorkSheet = xlWorkBook.Worksheets("Rule")

            For iCounter As Integer = 2 To 10
                Dim strSignalType As String = vbNullString
                Dim iDeckNum As Integer = -1
                Try
                    strSignalType = xlWorkSheet.Cells(iCounter, 1).value.ToString()
                    iDeckNum = Convert.ToInt32(xlWorkSheet.Cells(iCounter, 2).value.ToString())
                Catch
                    Exit For
                End Try
                If strSignalType = vbNullString Then Exit For

                m_oSignalTypeDeck.Add(strSignalType, iDeckNum)
            Next


            xlWorkBook.Close()
            xlApp.Quit()

            ReleaseExcelObject(xlWorkSheet)
            ReleaseExcelObject(xlWorkBook)
            ReleaseExcelObject(xlApp)

        Catch oEx As Exception
            Dim sException As String = oEx.Message
            Throw oEx
        End Try

    End Sub

    Private Sub ReleaseExcelObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
    ''' <summary>
    ''' Find the Deck to be Allowed for cable to get Pass through.
    ''' </summary>
    ''' <param name="oCableRun"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function FindAllowedFloorPerRule(ByVal oCableRun As Ingr.SP3D.Route.Middle.CableRun) As Integer
        Try

            'Hard Coded for Sample Rule.
            m_oSignalTypeDeck.Add("Control", 2)
            m_oSignalTypeDeck.Add("Power", 4)
            m_oSignalTypeDeck.Add("Communication", 5)

            Dim SignalType As Integer = oCableRun.SignalType
            Dim strSignalType As String = Nothing

            Select Case SignalType

                Case 1
                    strSignalType = "Communication"
                Case 2
                    strSignalType = "Control"
                Case 3
                    strSignalType = "Data"
                Case 4
                    strSignalType = "Fire Alarm"
                Case 5
                    strSignalType = "Lightning"
                Case 6
                    strSignalType = "MultiConductor Power"
                Case 7
                    strSignalType = "Power"
                Case 8
                    strSignalType = "Signal"
            End Select

            m_iAvailableFloor = m_oSignalTypeDeck.Item(strSignalType)
            Return m_iAvailableFloor
        Catch oEx As Exception
            Dim sException As String = oEx.Message
            m_iAvailableFloor = -1
            Return m_iAvailableFloor
        End Try
    End Function
End Class



