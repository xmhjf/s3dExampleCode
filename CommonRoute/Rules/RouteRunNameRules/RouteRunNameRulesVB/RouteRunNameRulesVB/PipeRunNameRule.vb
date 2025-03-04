'***************************************************************************************
'  Copyright (C) 2008-2009, Intergraph Corporation.  All rights reserved.
'
'  Project  : M:\SharedContent\Src\CommonRoute\Rules\RouteRunNameRules\RouteRunNameRulesVB
'
'  Class    : PipeRunNameRule
'
'  Abstract : The file contains implementation of the naming rules for PipeRuns.
'
'***************************************************************************************
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Systems.Middle
Imports Ingr.SP3D.Route.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.ReferenceData.Middle

Public Class PipeRunNameRule
    Inherits NameRuleBase
    Private Const strCountFormat = "0000"

    ''' <summary>
    '''  Creates a name for the object passed in. The name is based on the parents
    '''  name and object name.The Naming Parents are added in AddNamingParents().
    '''  Both these methods are called from naming rule semantic.
    ''' </summary>
    ''' <param name="oEntity">Child object that needs to have the naming rule naming.</param>
    ''' <param name="oParents">Naming parents collection.</param>
    ''' <param name="oActiveEntity">Naming rules active entity on which the NamingParentsString is stored.</param>

    Public Overrides Sub ComputeName(ByVal oEntity As Ingr.SP3D.Common.Middle.BusinessObject, ByVal oParents As ReadOnlyCollection(Of Ingr.SP3D.Common.Middle.BusinessObject), ByVal oActiveEntity As Ingr.SP3D.Common.Middle.BusinessObject)

        Dim strPipelineName As String = vbNullString, strRunName As String = vbNullString
        Dim oPipeline As Pipeline, strFluidCode As String, oCLPropValue As PropertyValueCodelist
        Dim oPipeRun As PipeRun, oPipeSpec As PipeSpec, strPipeSpec As String, strNPD As String
        Dim oStrPropValue As PropertyValueString

        Try
            'Get the Parent of the PipeRun
            oPipeline = oParents.Item(0)
            strPipelineName = oPipeline.Name

            'Get the name of the PipeRun
            oPipeRun = oEntity
            strRunName = oPipeRun.Name

            'Get NPD from PipeRun
            strNPD = CStr(oPipeRun.NPD.Size)

            'Get the Fluid Code from the PipeLine
            oCLPropValue = oPipeline.GetPropertyValue("IJPipelineSystem", "FluidCode")
            strFluidCode = oCLPropValue.PropertyInfo.CodeListInfo.GetCodelistItem(oCLPropValue.PropValue).Name

            'Get Pipe Spec from PipeRun
            oPipeSpec = oPipeRun.Specification
            oStrPropValue = oPipeSpec.GetPropertyValue("IJDPipeSpec", "SpecName")
            strPipeSpec = oStrPropValue.PropValue

            Dim oParentSystem As ISystem, strUnitSystem As String = vbNullString, oSysChild As ISystemChild

            oParentSystem = oParents.Item(0)
            'To get the UnitSystem
            Do While TypeOf oParentSystem Is ISystemChild
                If TypeOf oParentSystem Is UnitSystem Then
                    'Get the Unit system string.
                    Dim oSys As Ingr.SP3D.Systems.Middle.System

                    oSys = oParentSystem
                    oStrPropValue = oSys.GetPropertyValue("IJNamedItem", "Name")
                    strUnitSystem = oStrPropValue.PropValue
                    ' Presently we are Displaying the Name of the Unitsystem but if we want to change
                    ' it to display UnitCode of the Unit system then use the below code
                    'oStrPropValue = oSys.GetPropertyValue("IJUnitSystem", "UnitCode")
                    'strUnitSystem = oStrPropValue.PropValue
                    Exit Do
                Else
                    oSysChild = oParentSystem
                    oParentSystem = oSysChild.SystemParent
                End If
            Loop

            Dim strValidateName As String, strOldName As String = vbNullString, arr() As String
            Dim intUpperBound As Integer, intLowerBound As Integer
            Dim strOldPipeSpec As String = vbNullString, strOldSeqNo As String
            Dim RunNameLength As Integer, Checklength As Integer

            If Len(strUnitSystem) > 0 Then
                strValidateName = strUnitSystem + "-" + strNPD + "-" + strFluidCode + "-"
            Else
                strValidateName = strNPD + "-" + strFluidCode + "-"
            End If

            arr = Split(strRunName, "-", , vbTextCompare)
            intUpperBound = UBound(arr)
            If intUpperBound > 0 Then
                strOldPipeSpec = arr(intUpperBound)
                intLowerBound = intUpperBound - 1
                strOldSeqNo = arr(intLowerBound)
                strOldName = strRunName
                RunNameLength = Len(Trim(strOldName))
                Checklength = Len(Trim(strOldPipeSpec)) + Len(Trim(strOldSeqNo))
                'Deleting "1" more from RunNameLength to Include "-" also .
                Checklength = RunNameLength - Checklength - 1
                strOldName = Left$(strOldName, Checklength)
            End If
            'We Compare the NewString (strUnitSystem + "-" + strNPD + "-" + strFluidCode + "-" )and Oldstring from the RunName and
            'if they are Different we generate a New Name.we also Compare whethere the PipeRunSpec has Changed or Not.
            If Not (StrComp(strOldName, strValidateName, vbTextCompare) = 0) Or Not StrComp(strOldPipeSpec, strPipeSpec, vbTextCompare) = 0 Then
                Dim lCount As Long, strSeqNo As String, strLocation As String = vbNullString, strName As String

                GetCountAndLocationID(strPipelineName, lCount, strLocation)
                strSeqNo = Format(lCount, strCountFormat)
                If Len(strUnitSystem) > 0 Then
                    strName = strUnitSystem + "-" + strNPD + "-" + strFluidCode + "-" + strSeqNo + "-" + strPipeSpec
                Else
                    strName = strNPD + "-" + strFluidCode + "-" + strSeqNo + "-" + strPipeSpec
                End If
                SetNamingParentsString(oActiveEntity, strName)
                oEntity.SetPropertyValue(strName, "IJNamedItem", "Name")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    ''' <summary>
    ''' All the Naming Parents that need to participate in an objects naming are added here to the
    ''' Collection(Of BusinessObject). The parents added here are used in computing the name of the object in
    ''' ComputeName(). Both these methods are called from naming rule semantic.
    ''' </summary>
    ''' <param name="oEntity">Child object that needs to have the naming rule naming.</param>
    ''' <returns> Collection of parents that participate </returns>

    Public Overrides Function GetNamingParents(ByVal oEntity As Ingr.SP3D.Common.Middle.BusinessObject) As Collection(Of Ingr.SP3D.Common.Middle.BusinessObject)

        Dim oParentsColl As New Collection(Of Ingr.SP3D.Common.Middle.BusinessObject)
        Dim oParent As BusinessObject

        GetNamingParents = Nothing
        Try
            oParent = GetParent(HierarchyTypes.System, oEntity) ' Get the System Parent
            If (Not oParent Is Nothing) Then
                oParentsColl.Add(oParent) 'Add the Parent to the ParentColl
            End If
            GetNamingParents = oParentsColl
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function

End Class
