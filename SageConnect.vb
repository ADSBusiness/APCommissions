
Module SageConnect

    Public mDBLinkCmpRW As AccpacCOMAPI.AccpacDBLink
    Public DateTime As String

    Public sSageOrgID As String
    Public sSageCompName As String
    Public sSageUserID As String
    Public sSageSessDate As String


    Function SageSession() As Boolean

        SageSession = False

        DateTime = CStr(System.DateTime.UtcNow.ToLocalTime())
        Dim dNow As Date = Format(Today, "dd/MM/yyyy")
        Dim iID As Integer

        Dim sSession As AccpacCOMAPI.AccpacSession
        Dim sSignon As AccpacSignonManager.AccpacSignonMgr

        sSession = New AccpacCOMAPI.AccpacSession
        sSignon = New AccpacSignonManager.AccpacSignonMgr

        sSession.Init("", "CS", "CS0001", "67A")
        'sSession.Open("ADMIN", "ADMIN", "ATSDAT", dNow, 0, "")
        iID = sSignon.Signon(sSession)

        If sSession.IsOpened Then
            mDBLinkCmpRW = sSession.OpenDBLink(AccpacCOMAPI.tagDBLinkTypeEnum.DBLINK_COMPANY, AccpacCOMAPI.tagDBLinkFlagsEnum.DBLINK_FLG_READWRITE)

            sSageOrgID = sSession.CompanyID
            sSageCompName = sSession.CompanyName
            sSageUserID = sSession.UserID
            sSageSessDate = sSession.SessionDate

            SageSession = True

        End If


    End Function


End Module
