
Module SageConnect

    Public mDBLinkCmpRW As AccpacCOMAPI.AccpacDBLink
    Public DateTime As String

    Public sSageOrgID As String
    Public sSageCompName As String
    Public sSageUserID As String
    Public sSageSessDate As String

    Public sSession As AccpacCOMAPI.AccpacSession
    Public sSignon As AccpacSignonManager.AccpacSignonMgr
    Public FirstRecord As Boolean
    Public Temp As Boolean

    Public APBatch As String
    Public APBtchCreated As Boolean


    Public APINVOICE1batch As AccpacCOMAPI.AccpacView
    Public APINVOICE1header As AccpacCOMAPI.AccpacView
    Public APINVOICE1detail1 As AccpacCOMAPI.AccpacView
    Public APINVOICE1detail2 As AccpacCOMAPI.AccpacView
    Public APINVOICE1detail3 As AccpacCOMAPI.AccpacView
    Public APINVOICE1detail4 As AccpacCOMAPI.AccpacView
    Public APINVOICE1batchFields As AccpacCOMAPI.AccpacViewFields
    Public APINVOICE1headerFields As AccpacCOMAPI.AccpacViewFields
    Public APINVOICE1detail1Fields As AccpacCOMAPI.AccpacViewFields
    Public APINVOICE1detail2Fields As AccpacCOMAPI.AccpacViewFields
    Public APINVOICE1detail3Fields As AccpacCOMAPI.AccpacViewFields
    Public APINVOICE1detail4Fields As AccpacCOMAPI.AccpacViewFields

    Public APInvBatch(1) As AccpacCOMAPI.AccpacView
    Public APInvBatchHeader(4) As AccpacCOMAPI.AccpacView
    Public APInvBatchDetail1(3) As AccpacCOMAPI.AccpacView
    Public APInvBatchDetail2(1) As AccpacCOMAPI.AccpacView
    Public APInvBatchDetail3(1) As AccpacCOMAPI.AccpacView
    Public APInvBatchDetail4(1) As AccpacCOMAPI.AccpacView




    Public Function SageSession() As Boolean

        SageSession = False

        DateTime = CStr(System.DateTime.UtcNow.ToLocalTime())
        Dim dNow As Date = Format(Today, "dd/MM/yyyy")
        Dim iID As Integer

        '
        ' This is designed to be run from within Sage, and will use the current session login details
        ' This will not use an additional Lanpak as it uses the current session
        ' This can be run external from Sage, however will prompt for a login, and hence will consume a Lanpak
        '
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

    Public Function DeclareViews() As Boolean

        Try
            '
            ' Accounts Payable
            '

            mDBLinkCmpRW.OpenView("AP0020", APINVOICE1batch)
            mDBLinkCmpRW.OpenView("AP0021", APINVOICE1header)
            mDBLinkCmpRW.OpenView("AP0022", APINVOICE1detail1)
            mDBLinkCmpRW.OpenView("AP0023", APINVOICE1detail2)
            mDBLinkCmpRW.OpenView("AP0402", APINVOICE1detail3)
            mDBLinkCmpRW.OpenView("AP0401", APINVOICE1detail4)
            APINVOICE1batchFields = APINVOICE1batch.Fields
            APINVOICE1headerFields = APINVOICE1header.Fields
            APINVOICE1detail1Fields = APINVOICE1detail1.Fields
            APINVOICE1detail2Fields = APINVOICE1detail2.Fields
            APINVOICE1detail3Fields = APINVOICE1detail3.Fields
            APINVOICE1detail4Fields = APINVOICE1detail4.Fields

            'APINVOICE1batch.Compose(Array(APINVOICE1header))
            APInvBatch(0) = APINVOICE1header
            APINVOICE1batch.Compose(APInvBatch)

            'APINVOICE1header.Compose Array(APINVOICE1batch, APINVOICE1detail1, APINVOICE1detail2, APINVOICE1detail3)
            APInvBatchHeader(0) = APINVOICE1batch
            APInvBatchHeader(1) = APINVOICE1detail1
            APInvBatchHeader(2) = APINVOICE1detail2
            APInvBatchHeader(3) = APINVOICE1detail3
            APINVOICE1header.Compose(APInvBatchHeader)

            'APINVOICE1detail1.Compose Array(APINVOICE1header, APINVOICE1batch, APINVOICE1detail4)
            APInvBatchDetail1(0) = APINVOICE1header
            APInvBatchDetail1(1) = APINVOICE1batch
            APInvBatchDetail1(2) = APINVOICE1detail4
            APINVOICE1detail1.Compose(APInvBatchDetail1)

            'APINVOICE1detail2.Compose Array(APINVOICE1header)
            APInvBatchDetail2(0) = APINVOICE1header
            APINVOICE1detail2.Compose(APInvBatchDetail2)

            'APINVOICE1detail3.Compose Array(APINVOICE1header)
            APInvBatchDetail3(0) = APINVOICE1header
            APINVOICE1detail3.Compose(APInvBatchDetail3)

            'APINVOICE1detail4.Compose Array(APINVOICE1detail1)
            APInvBatchDetail4(0) = APINVOICE1header
            APINVOICE1detail4.Compose(APInvBatchDetail4)




        Catch ex As Exception

            AccpacErrorHandler()

            DeclareViews = True

        End Try



    End Function

    Public Function Sage_AP_CreateBatch(sBtchDesc As String, sBtchDate As String) As String

        ''MsgBox(Sage_AP_CreateBatch("test desc", "2022/03/01"))

        Sage_AP_CreateBatch = ""
        Try

            APINVOICE1batch.Browse("((BTCHSTTS = 1) OR (BTCHSTTS = 7))", 1)
            Temp = APINVOICE1batch.Exists
            APINVOICE1batch.RecordCreate(1)
            APINVOICE1batch.Read()
            Temp = APINVOICE1header.Exists
            APINVOICE1header.RecordCreate(2)
            APINVOICE1detail1.Cancel()
            APINVOICE1batchFields.FieldByName("BTCHDESC").PutWithoutVerification(sBtchDesc)
            APINVOICE1batch.Update()
            APINVOICE1batchFields.FieldByName("DATEBTCH").Value = (sBtchDate)
            APINVOICE1batch.Update()
            APINVOICE1batch.Read()
            Sage_AP_CreateBatch = APINVOICE1batchFields.FieldByName("CNTBTCH").Value
            APBatch = APINVOICE1batchFields.FieldByName("CNTBTCH").Value

            Exit Function

        Catch ex As Exception

            AccpacErrorHandler()

        End Try




    End Function


    Public Sub AccpacErrorHandler()
        Dim idxError As Integer
        Dim OutString As String
        If sSession.Errors.Count > 0 Then
            OutString = ""

            For idxError = 0 To sSession.Errors.Count - 1
                OutString = OutString + sSession.Errors.Item(idxError) & vbLf
            Next idxError
            Select Case MsgBox("Errors reported as follows:" & vbLf & OutString & vbLf & "Do you want to continue processing?", MsgBoxStyle.YesNo)
                Case vbYes : Resume Next
                Case vbNo : Exit Sub
            End Select
            sSession.Errors.Clear()
        End If


    End Sub

End Module
