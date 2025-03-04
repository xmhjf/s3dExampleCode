Attribute VB_Name = "DefineIAViewName"
Option Explicit

Private Const VIEW_PREFIX_CMD As String = "ViewPrefix"
Private Const VIEW_NAME_CMD As String = "ViewName"
Public Sub DefineViewName(ByVal pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition, strName As String)
    Dim pCommandDescription As IJDCommandDescription
    Set pCommandDescription = New DCommandDescription

    pCommandDescription.Name = VIEW_NAME_CMD
    pCommandDescription.Properties = 0
    pCommandDescription.Type = 0
    pCommandDescription.Source = strName

    pSymbolDefinition.IJDUserCommands.SetCommand pCommandDescription
End Sub
Public Sub DefineViewPrefix(ByVal pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition, strPrefix As String)
    Dim pCommandDescription As IJDCommandDescription
    Set pCommandDescription = New DCommandDescription

    pCommandDescription.Name = VIEW_PREFIX_CMD
    pCommandDescription.Properties = 0
    pCommandDescription.Type = 0
    pCommandDescription.Source = strPrefix

    pSymbolDefinition.IJDUserCommands.SetCommand pCommandDescription
End Sub
