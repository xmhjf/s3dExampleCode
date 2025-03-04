Imports Ingr.SP3D.Common.Middle.Services
Friend Module SectionLibraryCalculatorLocalizer
    ''' <summary>
    ''' This method is used to set the localized strings for control's text and error messages.
    ''' If the string matches with the string in resource file, appropriate message from resource file is fetched else the default message.
    ''' </summary>
    Function GetString(ByVal IID As Integer, ByVal defaultMessage As String) As String
        Return CmnLocalizer.GetString(IID, defaultMessage, SectionLibraryCalculatorResourceIDs.DEFAULT_RESOURCE, SectionLibraryCalculatorResourceIDs.DEFAULT_ASSEMBLY)
    End Function

    ''' <summary>
    '''Retrieves the string from the resource with a colon.
    ''' </summary>
    Function GetStringWithColon(ByVal IID As Integer, ByVal defaultMessage As String) As String

        Return CmnLocalizer.GetStringWithColon(IID, defaultMessage, SectionLibraryCalculatorResourceIDs.DEFAULT_RESOURCE, SectionLibraryCalculatorResourceIDs.DEFAULT_ASSEMBLY)
    End Function
End Module
