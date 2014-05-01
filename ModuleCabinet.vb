Module ModuleCabinet

    Public AboutShow As Boolean
    Public LicenseShow As Boolean
    Public SettingsShow As Boolean
    Public FontModule As Font
    Public ActionModuleFormat As Boolean
    Public ActionTopMost As Boolean
    Public ActionUserPass As Boolean
    Public UserName As String
    Public UserPass As String
    Public UserLoad As Boolean

    Public TUser As TestUser
    Public Sub LoadTUser()
        TUser = New TestUser
    End Sub

    Public SettingsProg As Settings
    Public Sub LoadSettingsProg()
        SettingsProg = New Settings
    End Sub

    Public Programm As MyCabinet
    Public Sub LoadProgramm()
        Programm = New MyCabinet
    End Sub


End Module
