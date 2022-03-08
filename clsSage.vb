''Imports ACCPAC.Advantage
Imports AccpacCOMAPI
''Imports RICG_CloudERP_Updater.Utility

Public Class clsSage
    Dim sSageSession As AccpacCOMAPI.AccpacSession

    Public Function fConnect(ByVal SageUser As String, ByVal SagePassword As String, ByVal SageDatabase As String) As Boolean
        Dim sError As String = vbNullString
        Try
            sSageSession = New AccpacCOMAPI.AccpacSession
            sSageSession.Init("", "XY", "XY1000", "61A")
            'fLogIt("Sage Credentials: UserName [" & SageUser & "] Password [" & SagePassword & "] Database [" & SageDatabase & "]")
            sSageSession.Open(SageUser, SagePassword, SageDatabase, DateTime.Now, 0, String.Empty)
            Return True

        Catch ex As Exception
            sError = Err.Description
            fHandleSageErrors(sError)
            fLogIt("ERR~~fConnect> " & sError)
            Return False
        End Try
    End Function


    Public Function fSAGEInsertAssembly(ByVal sMirrorConnection As String, ByVal TRANSDATE As Date, ByVal ITEMNO As String, ByVal BOMNO As String, ByVal REFERENCE As String, ByVal LOCATION As String, ByVal QUANTITY As Double, ByVal SERIAL As String, ByRef sError As String) As Boolean
        Try
            Dim mDBLinkCmpRW As AccpacCOMAPI.AccpacDBLink
            mDBLinkCmpRW = sSageSession.OpenDBLink(AccpacCOMAPI.tagDBLinkTypeEnum.DBLINK_COMPANY, AccpacCOMAPI.tagDBLinkFlagsEnum.DBLINK_FLG_READWRITE)
            Dim mDBLinkSysRW As AccpacCOMAPI.AccpacDBLink
            mDBLinkSysRW = sSageSession.OpenDBLink(AccpacCOMAPI.tagDBLinkTypeEnum.DBLINK_SYSTEM, AccpacCOMAPI.tagDBLinkFlagsEnum.DBLINK_FLG_READWRITE)

            Dim temp As Boolean
            Dim ICASEN1header As AccpacCOMAPI.AccpacView
            Dim ICASEN1headerFields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("IC0160", ICASEN1header)
            ICASEN1headerFields = ICASEN1header.Fields

            Dim ICASEN1detail1 As AccpacCOMAPI.AccpacView
            Dim ICASEN1detail1Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("IC0165", ICASEN1detail1)
            ICASEN1detail1Fields = ICASEN1detail1.Fields

            Dim ICASEN1detail2 As AccpacCOMAPI.AccpacView
            Dim ICASEN1detail2Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("IC0162", ICASEN1detail2)
            ICASEN1detail2Fields = ICASEN1detail2.Fields

            Dim ICASEN1detail3 As AccpacCOMAPI.AccpacView
            Dim ICASEN1detail3Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("IC0167", ICASEN1detail3)
            ICASEN1detail3Fields = ICASEN1detail3.Fields

            ICASEN1header.Compose(New AccpacCOMAPI.AccpacView() {Nothing, Nothing, Nothing, ICASEN1detail1, ICASEN1detail2, ICASEN1detail3})

            ICASEN1detail1.Compose(New AccpacCOMAPI.AccpacView() {ICASEN1header})

            ICASEN1detail2.Compose(New AccpacCOMAPI.AccpacView() {ICASEN1header})

            ICASEN1detail3.Compose(New AccpacCOMAPI.AccpacView() {ICASEN1header})

            ICASEN1header.Order = 5
            ICASEN1header.FilterSelect("(DELETED = 0)", True, 5, 0)
            ICASEN1header.Order = 5
            ICASEN1header.Order = 0

            ICASEN1headerFields.FieldByName("ASSMENSEQ").PutWithoutVerification("0")         ' Sequence Number

            ICASEN1header.Init()
            ICASEN1header.Order = 5

            ICASEN1headerFields.FieldByName("TRANSDATE").Value = TRANSDATE      ' Transaction Date

            ICASEN1headerFields.FieldByName("ITEMNO").Value = ITEMNO                      ' Item Number
            ICASEN1headerFields.FieldByName("BOMNO").Value = BOMNO                            ' BOM Number
            ICASEN1headerFields.FieldByName("PROCESSCMD").PutWithoutVerification("1")        ' Process Command
            ICASEN1headerFields.FieldByName("REFERENCE").PutWithoutVerification(REFERENCE)    ' Reference

            ICASEN1header.Process()

            ICASEN1headerFields.FieldByName("COMPASSMTD").Value = "1"                         ' Component Assembly Method

            ICASEN1headerFields.FieldByName("LOCATION").Value = LOCATION                     ' Location
            ICASEN1headerFields.FieldByName("QUANTITY").Value = QUANTITY                    ' Quantity
            ICASEN1headerFields.FieldByName("STATUS").PutWithoutVerification("2")            ' Record Status

            ICASEN1header.Insert()
            ICASEN1header.Order = 0

            ICASEN1headerFields.FieldByName("ASSMENSEQ").PutWithoutVerification("0")         ' Sequence Number

            ICASEN1header.Init()
            ICASEN1header.Order = 5

            Return True

        Catch ex As Exception
            sError = Err.Description
            fHandleSageErrors(sError)
            sError = "ERR~~" & sError
            fLogIt("ERR~~fSAGEInsertAssembly. Error: " & sError)
            Return False
        End Try
    End Function

    Public Function fSAGEInsertARInvoice(ByVal sMirrorConnection As String, ByVal dsIn As DataSet, ByRef sError As String) As Boolean
        Dim sErrReference As String = vbNullString
        Try
            Dim sCurrentCustomer As String = vbNullString
            Dim sReferenceGuid As String = Guid.NewGuid.ToString

            Dim mDBLinkCmpRW As AccpacCOMAPI.AccpacDBLink
            mDBLinkCmpRW = sSageSession.OpenDBLink(AccpacCOMAPI.tagDBLinkTypeEnum.DBLINK_COMPANY, AccpacCOMAPI.tagDBLinkFlagsEnum.DBLINK_FLG_READWRITE)
            Dim mDBLinkSysRW As AccpacCOMAPI.AccpacDBLink
            mDBLinkSysRW = sSageSession.OpenDBLink(AccpacCOMAPI.tagDBLinkTypeEnum.DBLINK_SYSTEM, AccpacCOMAPI.tagDBLinkFlagsEnum.DBLINK_FLG_READWRITE)

            Dim temp As Boolean
            Dim ARINVOICE1batch As AccpacCOMAPI.AccpacView
            Dim ARINVOICE1batchFields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("AR0031", ARINVOICE1batch)
            ARINVOICE1batchFields = ARINVOICE1batch.Fields

            Dim ARINVOICE1header As AccpacCOMAPI.AccpacView
            Dim ARINVOICE1headerFields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("AR0032", ARINVOICE1header)
            ARINVOICE1headerFields = ARINVOICE1header.Fields

            Dim ARINVOICE1detail1 As AccpacCOMAPI.AccpacView
            Dim ARINVOICE1detail1Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("AR0033", ARINVOICE1detail1)
            ARINVOICE1detail1Fields = ARINVOICE1detail1.Fields

            Dim ARINVOICE1detail2 As AccpacCOMAPI.AccpacView
            Dim ARINVOICE1detail2Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("AR0034", ARINVOICE1detail2)
            ARINVOICE1detail2Fields = ARINVOICE1detail2.Fields

            Dim ARINVOICE1detail3 As AccpacCOMAPI.AccpacView
            Dim ARINVOICE1detail3Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("AR0402", ARINVOICE1detail3)
            ARINVOICE1detail3Fields = ARINVOICE1detail3.Fields

            Dim ARINVOICE1detail4 As AccpacCOMAPI.AccpacView
            Dim ARINVOICE1detail4Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("AR0401", ARINVOICE1detail4)
            ARINVOICE1detail4Fields = ARINVOICE1detail4.Fields

            ARINVOICE1batch.Compose(New AccpacCOMAPI.AccpacView() {ARINVOICE1header})

            ARINVOICE1header.Compose(New AccpacCOMAPI.AccpacView() {ARINVOICE1batch, ARINVOICE1detail1, ARINVOICE1detail2, ARINVOICE1detail3, Nothing})

            ARINVOICE1detail1.Compose(New AccpacCOMAPI.AccpacView() {ARINVOICE1header, ARINVOICE1batch, ARINVOICE1detail4})

            ARINVOICE1detail2.Compose(New AccpacCOMAPI.AccpacView() {ARINVOICE1header})

            ARINVOICE1detail3.Compose(New AccpacCOMAPI.AccpacView() {ARINVOICE1header})

            ARINVOICE1detail4.Compose(New AccpacCOMAPI.AccpacView() {ARINVOICE1detail1})

            ARINVOICE1batch.RecordCreate(1)

            ARINVOICE1batchFields.FieldByName("PROCESSCMD").PutWithoutVerification("1")      ' Process Command

            ARINVOICE1batch.Process()
            ARINVOICE1batch.Read()
            ARINVOICE1header.RecordCreate(2)
            ARINVOICE1detail1.Cancel()
            ARINVOICE1batchFields.FieldByName("DATEBTCH").Value = dsIn.Tables(0).Rows(0).Item("Date")    ' Batch Date
            ARINVOICE1batch.Update()
            ARINVOICE1batch.Read()
            ARINVOICE1header.RecordCreate(2)
            ARINVOICE1detail1.Cancel()

            Dim dCountHeader As Long = 0
            Dim dCountLines As Long = 0
            For Each dR As DataRow In dsIn.Tables(0).Rows

                If sCurrentCustomer <> dR("Customer") Then
                    If dCountHeader <> 0 Then

                        ARINVOICE1headerFields.FieldByName("INVCDESC").PutWithoutVerification((sReferenceGuid & " - " & dCountHeader).ToString)
                        ARINVOICE1header.Insert()
                        ARINVOICE1detail1.Read()
                        ARINVOICE1detail1.Read()
                        ARINVOICE1batch.Read()
                        ARINVOICE1header.RecordCreate(2)
                        ARINVOICE1detail1.Cancel()
                    End If

                    dCountHeader = dCountHeader + 1
                    dCountLines = 0

                    ARINVOICE1headerFields.FieldByName("CNTITEM").Value = dCountHeader                        ' Entry Number
                    ARINVOICE1header.Fetch()
                    temp = ARINVOICE1header.Exists
                    temp = ARINVOICE1header.Exists
                    ARINVOICE1batch.Read()
                    ARINVOICE1header.RecordCreate(2)
                    ARINVOICE1detail1.Cancel()

                    ARINVOICE1headerFields.FieldByName("IDCUST").Value = dR("Customer")               ' Customer Number
                    ARINVOICE1headerFields.FieldByName("PROCESSCMD").PutWithoutVerification("4")     ' Process Command
                    ARINVOICE1header.Process()

                    sCurrentCustomer = dR("Customer")
                End If

                dCountLines = dCountLines + 1

                temp = ARINVOICE1detail1.Exists
                ARINVOICE1detail1.RecordClear()
                temp = ARINVOICE1detail1.Exists
                ARINVOICE1detail1.RecordCreate(0)

                ARINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")    ' Process Command Code

                ARINVOICE1detail1.Process()

                ARINVOICE1detail1Fields.FieldByName("IDDIST").Value = dR("LineType")                    ' Distribution Code
                ARINVOICE1detail1Fields.FieldByName("TEXTDESC").Value = dR("Description")         ' Description
                ARINVOICE1detail1Fields.FieldByName("AMTEXTN").Value = dR("SubTotal")                   ' Extended Amount w/ TIP

                If dR("LedgerCode") <> "-1" Then
                    ARINVOICE1detail1Fields.FieldByName("IDACCTREV").Value = dR("LedgerCode")                  ' Revenue Account
                End If

                ARINVOICE1detail1Fields.FieldByName("COMMENT").PutWithoutVerification(dR("Reference"))

                ARINVOICE1detail1.Insert()

                ARINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification(dCountLines * -1)      ' Line Number

                ARINVOICE1detail1.Read()

            Next

            ARINVOICE1headerFields.FieldByName("INVCDESC").PutWithoutVerification(sReferenceGuid & "-" & dCountHeader)
            ARINVOICE1header.Insert()
            ARINVOICE1detail1.Read()
            ARINVOICE1detail1.Read()
            ARINVOICE1batch.Read()
            ARINVOICE1header.RecordCreate(2)
            ARINVOICE1detail1.Cancel()

            If (dsIn.Tables(0).Rows(0).Item("GLCategory") <> "") Then
                For k As Integer = 0 To dCountHeader
                    fUpdateGLCategoryOnARInvoice(sReferenceGuid & "-" & k, fGetValueInBrackets(Trim(dsIn.Tables(0).Rows(0).Item("GLCategory"))))
                Next
            End If

            Return True

        Catch ex As Exception
            sError = Err.Description
            fHandleSageErrors(sError)
            sError = "ERR~~" & sError
            fLogIt("ERR~~fSAGEInsertARInvoice. Error: " & sError)
            Return False
        End Try
    End Function

    Public Function fSAGEInsertAPInvoice(ByVal sMirrorConnection As String, ByVal dsIn As DataSet, ByRef sError As String) As Boolean
        Try
            Dim sCurrentSupplier As String = vbNullString
            Dim sReferenceGuid As String = Guid.NewGuid.ToString

            Dim mDBLinkCmpRW As AccpacCOMAPI.AccpacDBLink
            mDBLinkCmpRW = sSageSession.OpenDBLink(AccpacCOMAPI.tagDBLinkTypeEnum.DBLINK_COMPANY, AccpacCOMAPI.tagDBLinkFlagsEnum.DBLINK_FLG_READWRITE)
            Dim mDBLinkSysRW As AccpacCOMAPI.AccpacDBLink
            mDBLinkSysRW = sSageSession.OpenDBLink(AccpacCOMAPI.tagDBLinkTypeEnum.DBLINK_SYSTEM, AccpacCOMAPI.tagDBLinkFlagsEnum.DBLINK_FLG_READWRITE)

            Dim temp As Boolean
            Dim APINVOICE1batch As AccpacCOMAPI.AccpacView
            Dim APINVOICE1batchFields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("AP0020", APINVOICE1batch)
            APINVOICE1batchFields = APINVOICE1batch.Fields

            Dim APINVOICE1header As AccpacCOMAPI.AccpacView
            Dim APINVOICE1headerFields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("AP0021", APINVOICE1header)
            APINVOICE1headerFields = APINVOICE1header.Fields

            Dim APINVOICE1detail1 As AccpacCOMAPI.AccpacView
            Dim APINVOICE1detail1Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("AP0022", APINVOICE1detail1)
            APINVOICE1detail1Fields = APINVOICE1detail1.Fields

            Dim APINVOICE1detail2 As AccpacCOMAPI.AccpacView
            Dim APINVOICE1detail2Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("AP0023", APINVOICE1detail2)
            APINVOICE1detail2Fields = APINVOICE1detail2.Fields

            Dim APINVOICE1detail3 As AccpacCOMAPI.AccpacView
            Dim APINVOICE1detail3Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("AP0402", APINVOICE1detail3)
            APINVOICE1detail3Fields = APINVOICE1detail3.Fields

            Dim APINVOICE1detail4 As AccpacCOMAPI.AccpacView
            Dim APINVOICE1detail4Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("AP0401", APINVOICE1detail4)
            APINVOICE1detail4Fields = APINVOICE1detail4.Fields

            APINVOICE1batch.Compose(New AccpacCOMAPI.AccpacView() {APINVOICE1header})

            APINVOICE1header.Compose(New AccpacCOMAPI.AccpacView() {APINVOICE1batch, APINVOICE1detail1, APINVOICE1detail2, APINVOICE1detail3})

            APINVOICE1detail1.Compose(New AccpacCOMAPI.AccpacView() {APINVOICE1header, APINVOICE1batch, APINVOICE1detail4})

            APINVOICE1detail2.Compose(New AccpacCOMAPI.AccpacView() {APINVOICE1header})

            APINVOICE1detail3.Compose(New AccpacCOMAPI.AccpacView() {APINVOICE1header})

            APINVOICE1detail4.Compose(New AccpacCOMAPI.AccpacView() {APINVOICE1detail1})

            APINVOICE1batch.Browse("((BTCHSTTS = 1) OR (BTCHSTTS = 7))", 1)
            Dim APINVCPOST2 As AccpacCOMAPI.AccpacView
            Dim APINVCPOST2Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("AP0039", APINVCPOST2)
            APINVCPOST2Fields = APINVCPOST2.Fields

            APINVOICE1batch.RecordCreate(1)

            APINVOICE1batchFields.FieldByName("PROCESSCMD").PutWithoutVerification("1")      ' Process Command Code

            APINVOICE1batch.Process()
            APINVOICE1batch.Read()
            APINVOICE1header.RecordCreate(2)
            APINVOICE1detail1.Cancel()
            APINVOICE1batchFields.FieldByName("DATEBTCH").Value = dsIn.Tables(0).Rows(0).Item("Date")      ' Batch Date
            APINVOICE1batch.Update()
            APINVOICE1batch.Read()
            APINVOICE1header.RecordCreate(2)
            APINVOICE1detail1.Cancel()

            Dim dCountHeader As Long = 0
            Dim dCountLines As Long = 0
            Dim dTotal As Double = 0
            For Each dR As DataRow In dsIn.Tables(0).Rows

                If sCurrentSupplier <> Trim(dR("Customer")) Then
                    If dCountHeader <> 0 Then

                        APINVOICE1headerFields.FieldByName("INVCDESC").PutWithoutVerification(sReferenceGuid & "-" & dCountHeader)
                        APINVOICE1headerFields.FieldByName("AMTGROSTOT").Value = dTotal
                        APINVOICE1header.Insert()
                        APINVOICE1batch.Read()
                        APINVOICE1header.RecordCreate(2)
                        APINVOICE1detail1.Cancel()

                    End If

                    dCountHeader = dCountHeader + 1
                    dCountLines = 0

                    'APINVOICE1headerFields.FieldByName("CNTITEM").Value = dCountHeader
                    'APINVOICE1header.Fetch()
                    'temp = APINVOICE1header.Exists

                    'APINVOICE1header.Browse("", 1)
                    'temp = APINVOICE1header.Exists
                    'APINVOICE1header.Read()
                    'temp = APINVOICE1header.Exists

                    Dim sSupplier As String = Trim(dR("Customer"))
                    APINVOICE1headerFields.FieldByName("IDVEND").Value = sSupplier           ' Vendor Number
                    APINVOICE1headerFields.FieldByName("PROCESSCMD").PutWithoutVerification("7")     ' Process Command Code
                    APINVOICE1header.Process()
                    APINVOICE1headerFields.FieldByName("PROCESSCMD").PutWithoutVerification("4")     ' Process Command Code
                    APINVOICE1header.Process()

                    Dim sInvoiceNbr As String = "INV" & Right("0000000" & fGetLatestNumber(), 7)
                    APINVOICE1headerFields.FieldByName("IDINVC").Value = sInvoiceNbr

                    sCurrentSupplier = Trim(dR("Customer"))
                    dTotal = dR("Total")
                End If

                dCountLines = dCountLines + 1

                temp = APINVOICE1detail1.Exists
                APINVOICE1detail1.RecordClear()
                temp = APINVOICE1detail1.Exists
                APINVOICE1detail1.RecordCreate(0)

                APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")    ' Process Command Code

                APINVOICE1detail1.Process()

                APINVOICE1detail1Fields.FieldByName("IDDIST").Value = dR("LineType")            ' Distribution Code
                APINVOICE1detail1Fields.FieldByName("TEXTDESC").Value = dR("Description")         ' Description
                APINVOICE1detail1Fields.FieldByName("AMTDIST").Value = dR("SubTotal")       ' Extended Amount w/ TIP
                APINVOICE1detail1Fields.FieldByName("COMMENT").PutWithoutVerification(dR("Reference"))

                If dR("LedgerCode") <> "-1" Then
                    APINVOICE1detail1Fields.FieldByName("IDGLACCT").Value = dR("LedgerCode")                  ' Revenue Account
                End If

                APINVOICE1detail1.Insert()
                APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification(dCountLines * -1)      ' Line Number
                APINVOICE1detail1.Read()

            Next

            APINVOICE1headerFields.FieldByName("INVCDESC").PutWithoutVerification(sReferenceGuid & "-" & dCountHeader)
            APINVOICE1headerFields.FieldByName("AMTGROSTOT").Value = dTotal
            APINVOICE1header.Insert()
            APINVOICE1batch.Read()
            APINVOICE1header.RecordCreate(2)
            APINVOICE1detail1.Cancel()

            If (dsIn.Tables(0).Rows(0).Item("GLCategory") <> "") Then
                For k As Integer = 0 To dCountHeader
                    fUpdateGLCategoryOnAPInvoice(sReferenceGuid & "-" & k, fGetValueInBrackets(Trim(dsIn.Tables(0).Rows(0).Item("GLCategory"))))
                Next
            End If

            Return True

        Catch ex As Exception
            sError = Err.Description
            fHandleSageErrors(sError)
            sError = "ERR~~" & sError
            fLogIt("ERR~~fSAGEInsertAPInvoice. Error: " & sError)
            Return False
        End Try
    End Function

    Private Function fGetLatestNumber() As Long
        Dim dNextNo As Long = 0
        Try
            fLogIt("# fGetLatestNumber > Before Getting Latest Number")
            dNextNo = Convert.ToInt64(fGetSystemParam("MUS", "Sage_LastInvoiceNumber", "*"))

            fLogIt("# fGetLatestNumber > " & dNextNo)
            fUpdateSystemParam("MUS", "Sage_LastInvoiceNumber", "*", dNextNo + 1)
            fLogIt("# fGetLatestNumber > After Updating Latest Number")

            Return dNextNo
        Catch ex As Exception
            fLogIt("ERR~~fGetLatestNumber> " & Err.Description)
            Return Nothing
        End Try
    End Function

    Public Function fSAGEInsertInventoryReceipt(ByVal sMirrorConnection As String, ByVal dsIn As DataSet, ByRef sError As String) As Boolean
        Dim BOOKREF As String = IIf(IsDBNull(dsIn.Tables(0).Rows(0).Item("Reference")), "", dsIn.Tables(0).Rows(0).Item("Reference"))
        Try
            Dim mDBLinkCmpRW As AccpacCOMAPI.AccpacDBLink
            mDBLinkCmpRW = sSageSession.OpenDBLink(AccpacCOMAPI.tagDBLinkTypeEnum.DBLINK_COMPANY, AccpacCOMAPI.tagDBLinkFlagsEnum.DBLINK_FLG_READWRITE)
            Dim mDBLinkSysRW As AccpacCOMAPI.AccpacDBLink
            mDBLinkSysRW = sSageSession.OpenDBLink(AccpacCOMAPI.tagDBLinkTypeEnum.DBLINK_SYSTEM, AccpacCOMAPI.tagDBLinkFlagsEnum.DBLINK_FLG_READWRITE)

            Dim temp As Boolean
            Dim ICREE1header As AccpacCOMAPI.AccpacView
            Dim ICREE1headerFields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("IC0590", ICREE1header)
            ICREE1headerFields = ICREE1header.Fields

            Dim ICREE1detail1 As AccpacCOMAPI.AccpacView
            Dim ICREE1detail1Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("IC0580", ICREE1detail1)
            ICREE1detail1Fields = ICREE1detail1.Fields

            Dim ICREE1detail2 As AccpacCOMAPI.AccpacView
            Dim ICREE1detail2Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("IC0595", ICREE1detail2)
            ICREE1detail2Fields = ICREE1detail2.Fields

            Dim ICREE1detail3 As AccpacCOMAPI.AccpacView
            Dim ICREE1detail3Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("IC0585", ICREE1detail3)
            ICREE1detail3Fields = ICREE1detail3.Fields

            Dim ICREE1detail4 As AccpacCOMAPI.AccpacView
            Dim ICREE1detail4Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("IC0582", ICREE1detail4)
            ICREE1detail4Fields = ICREE1detail4.Fields

            Dim ICREE1detail5 As AccpacCOMAPI.AccpacView
            Dim ICREE1detail5Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("IC0587", ICREE1detail5)
            ICREE1detail5Fields = ICREE1detail5.Fields

            ICREE1header.Compose(New AccpacCOMAPI.AccpacView() {ICREE1detail1, ICREE1detail2})

            ICREE1detail1.Compose(New AccpacCOMAPI.AccpacView() {ICREE1header, Nothing, Nothing, Nothing, Nothing, Nothing, ICREE1detail3, ICREE1detail5, ICREE1detail4})

            ICREE1detail2.Compose(New AccpacCOMAPI.AccpacView() {ICREE1header})

            ICREE1detail3.Compose(New AccpacCOMAPI.AccpacView() {ICREE1detail1})

            ICREE1detail4.Compose(New AccpacCOMAPI.AccpacView() {ICREE1detail1})

            ICREE1detail5.Compose(New AccpacCOMAPI.AccpacView() {ICREE1detail1})


            ICREE1header.Order = 2
            ICREE1header.FilterSelect("(DELETED = 0)", True, 2, 0)
            ICREE1header.Order = 2

            ICREE1headerFields.FieldByName("RECPTYPE").Value = "2"                            ' Receipt Type
            ICREE1header.Init()
            ICREE1header.Order = 0

            ICREE1headerFields.FieldByName("SEQUENCENO").PutWithoutVerification("0")         ' Sequence Number

            ICREE1header.Init()
            temp = ICREE1detail1.Exists
            ICREE1detail1.RecordClear()
            ICREE1header.Order = 2

            ICREE1headerFields.FieldByName("RECPDATE").Value = dsIn.Tables(0).Rows(0).Item("Date")       ' Receipt Date
            ICREE1headerFields.FieldByName("PROCESSCMD").PutWithoutVerification("1")         ' Process Command
            ICREE1header.Process()

            ICREE1headerFields.FieldByName("REFERENCE").PutWithoutVerification(dsIn.Tables(0).Rows(0).Item("Reference"))    ' Reference
            ICREE1header.Process()

            ICREE1headerFields.FieldByName("RECPDESC").PutWithoutVerification(dsIn.Tables(0).Rows(0).Item("PONumber"))  ' Vendor Number
            ICREE1headerFields.FieldByName("PROCESSCMD").PutWithoutVerification("1")         ' Process Command

            ICREE1header.Process()

            temp = ICREE1detail1.Exists
            ICREE1detail1.RecordClear()
            temp = ICREE1detail1.Exists
            ICREE1detail1.RecordCreate(0)

            For i = 1 To dsIn.Tables(0).Rows.Count
                ICREE1detail1Fields.FieldByName("ITEMNO").Value = dsIn.Tables(0).Rows(0).Item("ItemCode")                           ' Item Number
                ICREE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("1")        ' Process Command

                ICREE1detail1.Process()

                ICREE1detail1Fields.FieldByName("LOCATION").Value = dsIn.Tables(0).Rows(0).Item("Location")                         ' Location
                ICREE1detail1Fields.FieldByName("RECPQTY").Value = dsIn.Tables(0).Rows(0).Item("Qty")                     ' Quantity Received
                ICREE1detail1Fields.FieldByName("UNITCOST").Value = dsIn.Tables(0).Rows(0).Item("ExtCost") / dsIn.Tables(0).Rows(0).Item("Qty")                   ' Unit Cost
                ICREE1detail1Fields.FieldByName("CHKBELZERO").PutWithoutVerification("1")        ' Check Below Zero

                ICREE1detail1.Process()
                ICREE1detail1.Insert()

                ICREE1detail1Fields.FieldByName("LINENO").PutWithoutVerification(i * -1)           ' Line Number

                ICREE1detail1.Read()
                temp = ICREE1detail1.Exists
                ICREE1detail1.RecordCreate(0)
            Next

            ICREE1detail1.Read()
            ICREE1headerFields.FieldByName("STATUS").PutWithoutVerification("2")             ' Record Status
            ICREE1header.Insert()
            ICREE1header.Init()
            ICREE1header.Order = 0

            ICREE1headerFields.FieldByName("SEQUENCENO").PutWithoutVerification("0")         ' Sequence Number

            ICREE1header.Init()
            temp = ICREE1detail1.Exists
            ICREE1detail1.RecordClear()
            ICREE1header.Order = 2

            If (dsIn.Tables(0).Rows(0).Item("GLCategory") <> "") Then
                fUpdateGLCategoryOnReceipt(dsIn.Tables(0).Rows(0).Item("Reference"), fGetValueInBrackets(Trim(dsIn.Tables(0).Rows(0).Item("GLCategory"))))
            End If

            Return True

        Catch ex As Exception
            sError = Err.Description
            fHandleSageErrors(sError)
            sError = "ERR~~fSAGEInsertInventoryShipment <Booking - " & BOOKREF & ">" & sError
            fLogIt("ERR~~fSAGEInsertInventoryReceipt. Error: " & sError)
            Return False
        End Try
    End Function


    Public Function fSAGEInsertInventoryShipment(ByVal sMirrorConnection As String, ByVal dsIn As DataSet, ByRef sError As String) As Boolean
        Dim BOOKREF As String = IIf(IsDBNull(dsIn.Tables(0).Rows(0).Item("Reference")), "", dsIn.Tables(0).Rows(0).Item("Reference"))
        Try
            Dim mDBLinkCmpRW As AccpacCOMAPI.AccpacDBLink
            mDBLinkCmpRW = sSageSession.OpenDBLink(AccpacCOMAPI.tagDBLinkTypeEnum.DBLINK_COMPANY, AccpacCOMAPI.tagDBLinkFlagsEnum.DBLINK_FLG_READWRITE)
            Dim mDBLinkSysRW As AccpacCOMAPI.AccpacDBLink
            mDBLinkSysRW = sSageSession.OpenDBLink(AccpacCOMAPI.tagDBLinkTypeEnum.DBLINK_SYSTEM, AccpacCOMAPI.tagDBLinkFlagsEnum.DBLINK_FLG_READWRITE)

            Dim temp As Boolean
            Dim ICSHE1header As AccpacCOMAPI.AccpacView
            Dim ICSHE1headerFields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("IC0640", ICSHE1header)
            ICSHE1headerFields = ICSHE1header.Fields

            Dim ICSHE1detail1 As AccpacCOMAPI.AccpacView
            Dim ICSHE1detail1Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("IC0630", ICSHE1detail1)
            ICSHE1detail1Fields = ICSHE1detail1.Fields

            Dim ICSHE1detail2 As AccpacCOMAPI.AccpacView
            Dim ICSHE1detail2Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("IC0645", ICSHE1detail2)
            ICSHE1detail2Fields = ICSHE1detail2.Fields

            Dim ICSHE1detail3 As AccpacCOMAPI.AccpacView
            Dim ICSHE1detail3Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("IC0635", ICSHE1detail3)
            ICSHE1detail3Fields = ICSHE1detail3.Fields

            Dim ICSHE1detail4 As AccpacCOMAPI.AccpacView
            Dim ICSHE1detail4Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("IC0632", ICSHE1detail4)
            ICSHE1detail4Fields = ICSHE1detail4.Fields

            Dim ICSHE1detail5 As AccpacCOMAPI.AccpacView
            Dim ICSHE1detail5Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("IC0636", ICSHE1detail5)
            ICSHE1detail5Fields = ICSHE1detail5.Fields

            ICSHE1header.Compose(New AccpacCOMAPI.AccpacView() {ICSHE1detail1, Nothing, ICSHE1detail2})

            ICSHE1detail1.Compose(New AccpacCOMAPI.AccpacView() {ICSHE1header, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, ICSHE1detail3, ICSHE1detail4, ICSHE1detail5})

            ICSHE1detail2.Compose(New AccpacCOMAPI.AccpacView() {ICSHE1header})

            ICSHE1detail3.Compose(New AccpacCOMAPI.AccpacView() {ICSHE1detail1})

            ICSHE1detail4.Compose(New AccpacCOMAPI.AccpacView() {ICSHE1detail1})

            ICSHE1detail5.Compose(New AccpacCOMAPI.AccpacView() {ICSHE1detail1})


            ICSHE1header.Order = 3
            ICSHE1header.FilterSelect("(DELETED = 0)", True, 3, 0)
            ICSHE1header.Order = 3
            ICSHE1header.Order = 0

            ICSHE1headerFields.FieldByName("SEQUENCENO").PutWithoutVerification("0")         ' Sequence Number

            ICSHE1header.Init()
            temp = ICSHE1detail1.Exists
            ICSHE1detail1.RecordClear()
            ICSHE1header.Order = 3

            ICSHE1headerFields.FieldByName("TRANSDATE").Value = dsIn.Tables(0).Rows(0).Item("Date")       ' Ship Date
            ICSHE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("1")         ' Process Command
            ICSHE1header.Process()

            'ICSHE1headerFields.FieldByName("CUSTNO").Value = dsIn.Tables(0).Rows(0).Item("PONumber")                    ' Customer Number
            'ICSHE1headerFields.FieldByName("PROCESSCMD").PutWithoutVerification("1")         ' Process Command

            ICSHE1header.Process()

            temp = ICSHE1detail1.Exists
            ICSHE1detail1.RecordClear()
            temp = ICSHE1detail1.Exists
            ICSHE1detail1.RecordCreate(0)

            For i = 1 To dsIn.Tables(0).Rows.Count
                ICSHE1detail1Fields.FieldByName("ITEMNO").Value = dsIn.Tables(0).Rows(0).Item("ItemCode")                           ' Item Number
                ICSHE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("1")       ' Process Command

                ICSHE1detail1.Process()

                ICSHE1detail1Fields.FieldByName("LOCATION").Value = dsIn.Tables(0).Rows(0).Item("Location")                         ' Location
                ICSHE1detail1Fields.FieldByName("QUANTITY").Value = dsIn.Tables(0).Rows(0).Item("Qty")                     ' Quantity
                ICSHE1detail1Fields.FieldByName("UNITPRICE").Value = dsIn.Tables(0).Rows(0).Item("ExtCost") / dsIn.Tables(0).Rows(0).Item("Qty")                   ' Unit Price
                'ICSHE1detail1Fields.FieldByName("FUNCTION").Value = "100"                         ' Function

                ICSHE1detail1.Process()
                ICSHE1detail1.Insert()

                ICSHE1detail1Fields.FieldByName("LINENO").PutWithoutVerification(i * -1)           ' Line Number

                ICSHE1detail1.Read()
            Next

            ICSHE1headerFields.FieldByName("HDRDESC").PutWithoutVerification(dsIn.Tables(0).Rows(0).Item("PONumber"))                    ' Customer Number
            ICSHE1headerFields.FieldByName("REFERENCE").PutWithoutVerification(dsIn.Tables(0).Rows(0).Item("Reference"))    ' Reference
            ICSHE1header.Insert()
            ICSHE1header.Order = 0

            ICSHE1headerFields.FieldByName("SEQUENCENO").PutWithoutVerification("0")         ' Sequence Number

            ICSHE1header.Init()
            temp = ICSHE1detail1.Exists
            ICSHE1detail1.RecordClear()
            ICSHE1header.Order = 3

            If (dsIn.Tables(0).Rows(0).Item("GLCategory") <> "") Then
                fUpdateGLCategoryOnShipment(dsIn.Tables(0).Rows(0).Item("Reference"), fGetValueInBrackets(Trim(dsIn.Tables(0).Rows(0).Item("GLCategory"))))
            End If

            Return True

        Catch ex As Exception
            sError = Err.Description
            fHandleSageErrors(sError)
            sError = "ERR~~fSAGEInsertInventoryShipment <Booking - " & BOOKREF & ">" & sError
            fLogIt("ERR~~fSAGEInsertInventoryShipment. <Booking - " & BOOKREF & "> Error: " & sError)
            Return False
        End Try

    End Function


    Public Function fHandleSageErrors(ByRef sErr As String) As Boolean
        Try
            If sSageSession.IsOpened Then
                If Not sSageSession.Errors Is Nothing Then
                    Dim sSageError As String = vbNullString
                    For i As Integer = 0 To sSageSession.Errors.Count - 1
                        sSageError = sSageError & " - " & sSageSession.Errors.Item(i)
                    Next
                    If Not Trim(sSageError) = vbNullString Then
                        sErr = sSageError
                    End If
                    sSageSession.Errors.Clear()
                End If
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function fGetValueInBrackets(ByVal sText As String) As String
        Try
            sText = sText.Substring(sText.LastIndexOf("[") + 1, sText.LastIndexOf("]") - (sText.LastIndexOf("[") + 1))
            Return sText
        Catch ex As Exception
            fLogIt("ERR~~fGetValueInBrackets. Error: " & Err.Description)
            Return sText
        End Try
    End Function

    Public Function fSAGEInsertSingleAPInvoice(ByVal sMirrorConnection As String, ByRef BATCHNBR As String, ByRef ENTRYNBR As String, ByRef LINENO As String, ByRef TOTAL As Double, ByVal NEWCUST As Boolean, ByVal REFERENCE As String, ByVal dRow As DataRow, ByRef sError As String) As Boolean
        Dim BOOKREF As String = IIf(IsDBNull(dRow("Reference")), "", dRow("Reference"))
        Try
            ' ================================
            ' DECLARE ENVIRONMENT
            ' ================================
            Dim sReferenceGuid As String = Guid.NewGuid.ToString
            Dim sLatestNo As String

            If NEWCUST Then
                sLatestNo = fGetLatestNumber()
            Else
                sLatestNo = "0"
            End If

            Dim mDBLinkCmpRW As AccpacCOMAPI.AccpacDBLink
            mDBLinkCmpRW = sSageSession.OpenDBLink(AccpacCOMAPI.tagDBLinkTypeEnum.DBLINK_COMPANY, AccpacCOMAPI.tagDBLinkFlagsEnum.DBLINK_FLG_READWRITE)
            Dim mDBLinkSysRW As AccpacCOMAPI.AccpacDBLink
            mDBLinkSysRW = sSageSession.OpenDBLink(AccpacCOMAPI.tagDBLinkTypeEnum.DBLINK_SYSTEM, AccpacCOMAPI.tagDBLinkFlagsEnum.DBLINK_FLG_READWRITE)

            Dim temp As Boolean
            Dim APINVOICE1batch As AccpacCOMAPI.AccpacView
            Dim APINVOICE1batchFields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("AP0020", APINVOICE1batch)
            APINVOICE1batchFields = APINVOICE1batch.Fields

            Dim APINVOICE1header As AccpacCOMAPI.AccpacView
            Dim APINVOICE1headerFields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("AP0021", APINVOICE1header)
            APINVOICE1headerFields = APINVOICE1header.Fields

            Dim APINVOICE1detail1 As AccpacCOMAPI.AccpacView
            Dim APINVOICE1detail1Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("AP0022", APINVOICE1detail1)
            APINVOICE1detail1Fields = APINVOICE1detail1.Fields

            Dim APINVOICE1detail2 As AccpacCOMAPI.AccpacView
            Dim APINVOICE1detail2Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("AP0023", APINVOICE1detail2)
            APINVOICE1detail2Fields = APINVOICE1detail2.Fields

            Dim APINVOICE1detail3 As AccpacCOMAPI.AccpacView
            Dim APINVOICE1detail3Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("AP0402", APINVOICE1detail3)
            APINVOICE1detail3Fields = APINVOICE1detail3.Fields

            Dim APINVOICE1detail4 As AccpacCOMAPI.AccpacView
            Dim APINVOICE1detail4Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("AP0401", APINVOICE1detail4)
            APINVOICE1detail4Fields = APINVOICE1detail4.Fields

            APINVOICE1batch.Compose(New AccpacCOMAPI.AccpacView() {APINVOICE1header})

            APINVOICE1header.Compose(New AccpacCOMAPI.AccpacView() {APINVOICE1batch, APINVOICE1detail1, APINVOICE1detail2, APINVOICE1detail3})

            APINVOICE1detail1.Compose(New AccpacCOMAPI.AccpacView() {APINVOICE1header, APINVOICE1batch, APINVOICE1detail4})

            APINVOICE1detail2.Compose(New AccpacCOMAPI.AccpacView() {APINVOICE1header})

            APINVOICE1detail3.Compose(New AccpacCOMAPI.AccpacView() {APINVOICE1header})

            APINVOICE1detail4.Compose(New AccpacCOMAPI.AccpacView() {APINVOICE1detail1})


            APINVOICE1batch.Browse("((BTCHSTTS = 1) OR (BTCHSTTS = 7))", 1)
            Dim APINVCPOST2 As AccpacCOMAPI.AccpacView
            Dim APINVCPOST2Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("AP0039", APINVCPOST2)
            APINVCPOST2Fields = APINVCPOST2.Fields

            ' ================================
            ' FIND BATCH OR CREATE NEW BATCH FOR AP INVOICE
            ' ================================
            If BATCHNBR = 0 Then
                APINVOICE1batch.RecordCreate(1)

                APINVOICE1batchFields.FieldByName("PROCESSCMD").PutWithoutVerification("1")      ' Process Command Code

                APINVOICE1batch.Process()
                APINVOICE1batch.Read()
                APINVOICE1header.RecordCreate(2)
                APINVOICE1detail1.Cancel()
                APINVOICE1batchFields.FieldByName("DATEBTCH").Value = dRow("Date")      ' Batch Date
                APINVOICE1batch.Update()
                APINVOICE1batch.Read()

            Else
                APINVOICE1batchFields.FieldByName("CNTBTCH").Value = BATCHNBR                        ' Batch Number
                APINVOICE1batch.Read()
                APINVOICE1batchFields.FieldByName("PROCESSCMD").PutWithoutVerification("1")      ' Process Command Code

                APINVOICE1batch.Process()

            End If


            ' ================================
            ' CREATE NEW ENTRY WITHIN BATCH FOR NEW CUSTOMER
            ' ================================
            If NEWCUST Then
                APINVOICE1header.RecordCreate(2)
                APINVOICE1detail1.Cancel()

                Dim sSupplier As String = Trim(dRow("Customer"))
                APINVOICE1headerFields.FieldByName("IDVEND").Value = sSupplier           ' Vendor Number
                APINVOICE1headerFields.FieldByName("PROCESSCMD").PutWithoutVerification("7")     ' Process Command Code
                APINVOICE1header.Process()
                APINVOICE1headerFields.FieldByName("PROCESSCMD").PutWithoutVerification("4")     ' Process Command Code
                APINVOICE1header.Process()

                If dRow("ShipTo") <> vbNullString Then
                    If Not fCheckInvoiceTogether(dRow("Customer")) Then
                        APINVOICE1headerFields.FieldByName("IDRMITTO").Value = Left(dRow("ShipTo") + "000", 3)                    ' Remit-To Location
                        APINVOICE1headerFields.FieldByName("PROCESSCMD").PutWithoutVerification("4")     ' Process Command Code
                        APINVOICE1header.Process()
                    End If
                End If

                Dim sInvoiceNbr As String = "INV" & Right("0000000" & sLatestNo, 7)
                fLogIt("fSAGEInsertSingleAPInvoice > Code #AFTER LATEST NO")
                'APINVOICE1headerFields.FieldByName("IDINVC").Value = Trim(sInvoiceNbr)
                APINVOICE1headerFields.FieldByName("IDINVC").PutWithoutVerification(sInvoiceNbr)
                fLogIt("fSAGEInsertSingleAPInvoice > Code #AFTER INV NO")
                'APINVOICE1header.Insert()
                'APINVOICE1batch.Read()
                'APINVOICE1header.RecordCreate(2)
                'APINVOICE1detail1.Cancel()

                LINENO = 0
                TOTAL = 0
            Else
                APINVOICE1headerFields.FieldByName("CNTITEM").PutWithoutVerification(ENTRYNBR)   ' Entry Number
                APINVOICE1header.Browse("", 1)
                APINVOICE1header.Fetch()
            End If

            fLogIt("fSAGEInsertSingleAPInvoice > Code #AFTER NEW ENTRY")

            ' ================================
            ' PROCESS INVOICE LINES
            ' ================================

            TOTAL = TOTAL + dRow("Total")

            temp = APINVOICE1detail1.Exists
            APINVOICE1detail1.RecordClear()
            temp = APINVOICE1detail1.Exists
            APINVOICE1detail1.RecordCreate(0)

            APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")    ' Process Command Code

            APINVOICE1detail1.Process()

            APINVOICE1detail1Fields.FieldByName("IDDIST").Value = dRow("LineType")            ' Distribution Code
            APINVOICE1detail1Fields.FieldByName("TEXTDESC").Value = dRow("Description")         ' Description
            APINVOICE1detail1Fields.FieldByName("AMTDIST").Value = dRow("SubTotal")       ' Extended Amount w/ TIP
            APINVOICE1detail1Fields.FieldByName("COMMENT").PutWithoutVerification(dRow("Reference"))

            If dRow("LedgerCode") <> "" And dRow("LedgerCode") <> "-1" Then
                APINVOICE1detail1Fields.FieldByName("IDGLACCT").Value = dRow("LedgerCode")                  ' Revenue Account
            End If

            APINVOICE1detail1.Insert()
            APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification((LINENO + 1) * -1)      ' Line Number
            APINVOICE1detail1.Read()

            LINENO = LINENO + 1

            fLogIt("fSAGEInsertSingleAPInvoice > Code #AFTER PROCESS")

            ' ================================
            ' UPDATE DOCUMENT TOTAL
            ' ================================
            If APINVOICE1headerFields.FieldByName("INVCDESC").Value = vbNullString Then
                APINVOICE1headerFields.FieldByName("INVCDESC").PutWithoutVerification(REFERENCE)
            End If
            fLogIt("fSAGEInsertSingleAPInvoice > Code #B4 PO")
            If Not IsDBNull(dRow("PONumber")) Then
                If Trim(dRow("PONumber")) <> vbNullString Then
                    APINVOICE1headerFields.FieldByName("PONBR").PutWithoutVerification(dRow("PONumber"))
                    fLogIt("fSAGEInsertSingleAPInvoice > Code #AFTER PO")
                End If
            End If


            fLogIt("fSAGEInsertSingleAPInvoice > B4 TOTAL")
            APINVOICE1headerFields.FieldByName("AMTGROSTOT").PutWithoutVerification(TOTAL.ToString)
            fLogIt("fSAGEInsertSingleAPInvoice > Code #UPDATE TOTAL")
            BATCHNBR = APINVOICE1batchFields.FieldByName("CNTBTCH").Value
            'ENTRYNBR = APINVOICE1headerFields.FieldByName("CNTITEM").Value
            fLogIt("fSAGEInsertSingleAPInvoice > STORE BATCH")


            ' ================================
            ' ADD ENTRY
            ' ================================
            If NEWCUST Then
                fLogIt("fSAGEInsertSingleAPInvoice > B4 INSERT")
                APINVOICE1header.Insert()
                APINVOICE1batch.Read()
                APINVOICE1header.RecordCreate(2)
                APINVOICE1detail1.Cancel()
                fLogIt("fSAGEInsertSingleAPInvoice > INSERTED")
            Else
                fLogIt("fSAGEInsertSingleAPInvoice > B4 UPDATE")
                APINVOICE1header.Update()
                fLogIt("fSAGEInsertSingleAPInvoice > UPDATED")
            End If

            fLogIt("fSAGEInsertSingleAPInvoice > MOVE TO FIRST")
            APINVOICE1batchFields.FieldByName("PROCESSCMD").PutWithoutVerification("1")      ' Process Command Code

            APINVOICE1batch.Process()
            APINVOICE1headerFields.FieldByName("CNTITEM").PutWithoutVerification("-9999999")   ' Entry Number
            APINVOICE1header.Browse("", 1)
            APINVOICE1header.Fetch()
            fLogIt("fSAGEInsertSingleAPInvoice > MOVED")

            ' ================================
            ' UPDATE GL CATEGORY
            ' ================================
            If (dRow("GLCategory") <> "") Then
                fUpdateGLCategoryOnAPInvoice(REFERENCE, fGetValueInBrackets(Trim(dRow("GLCategory"))))
            End If

            fUpdateBookingDetailsWithInvoice(BOOKREF, BATCHNBR & "-" & ENTRYNBR)
            fUpdateInvoiceAsProcessed(BOOKREF, dRow("LineNum"), REFERENCE, BATCHNBR & "-" & ENTRYNBR)
            fLogIt("fSAGEInsertSingleAPInvoice > UPDATE MIRROR")
            Return True


        Catch ex As Exception
            sError = Err.Description
            fHandleSageErrors(sError)
            sError = "ERR~~<Booking - " & BOOKREF & "> " & sError
            fLogIt("ERR~~fSAGEInsertSingleAPInvoice. <Booking - " & BOOKREF & "> Error: " & sError)
            fUpdateInvoiceAsProcessed(BOOKREF, dRow("LineNum"), REFERENCE, "ERR")
            Return False
        End Try
    End Function
    Public Function fSAGEInsertSingleARInvoice(ByVal sMirrorConnection As String, ByRef BATCHNBR As String, ByRef ENTRYNBR As String, ByRef LINENO As String, ByRef TOTAL As Double, ByVal NEWCUST As Boolean, ByVal REFERENCE As String, ByVal dRow As DataRow, ByRef sError As String) As Boolean
        Dim BOOKREF As String = IIf(IsDBNull(dRow("Reference")), "", dRow("Reference"))

        Try
            ' ================================
            ' DECLARE ENVIRONMENT
            ' ================================
            Dim sCurrentCustomer As String = vbNullString
            Dim sReferenceGuid As String = Guid.NewGuid.ToString

            Dim mDBLinkCmpRW As AccpacCOMAPI.AccpacDBLink
            mDBLinkCmpRW = sSageSession.OpenDBLink(AccpacCOMAPI.tagDBLinkTypeEnum.DBLINK_COMPANY, AccpacCOMAPI.tagDBLinkFlagsEnum.DBLINK_FLG_READWRITE)
            Dim mDBLinkSysRW As AccpacCOMAPI.AccpacDBLink
            mDBLinkSysRW = sSageSession.OpenDBLink(AccpacCOMAPI.tagDBLinkTypeEnum.DBLINK_SYSTEM, AccpacCOMAPI.tagDBLinkFlagsEnum.DBLINK_FLG_READWRITE)

            Dim temp As Boolean
            Dim ARINVOICE1batch As AccpacCOMAPI.AccpacView
            Dim ARINVOICE1batchFields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("AR0031", ARINVOICE1batch)
            ARINVOICE1batchFields = ARINVOICE1batch.Fields

            Dim ARINVOICE1header As AccpacCOMAPI.AccpacView
            Dim ARINVOICE1headerFields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("AR0032", ARINVOICE1header)
            ARINVOICE1headerFields = ARINVOICE1header.Fields

            Dim ARINVOICE1detail1 As AccpacCOMAPI.AccpacView
            Dim ARINVOICE1detail1Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("AR0033", ARINVOICE1detail1)
            ARINVOICE1detail1Fields = ARINVOICE1detail1.Fields

            Dim ARINVOICE1detail2 As AccpacCOMAPI.AccpacView
            Dim ARINVOICE1detail2Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("AR0034", ARINVOICE1detail2)
            ARINVOICE1detail2Fields = ARINVOICE1detail2.Fields

            Dim ARINVOICE1detail3 As AccpacCOMAPI.AccpacView
            Dim ARINVOICE1detail3Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("AR0402", ARINVOICE1detail3)
            ARINVOICE1detail3Fields = ARINVOICE1detail3.Fields

            Dim ARINVOICE1detail4 As AccpacCOMAPI.AccpacView
            Dim ARINVOICE1detail4Fields As AccpacCOMAPI.AccpacViewFields
            mDBLinkCmpRW.OpenView("AR0401", ARINVOICE1detail4)
            ARINVOICE1detail4Fields = ARINVOICE1detail4.Fields

            ARINVOICE1batch.Compose(New AccpacCOMAPI.AccpacView() {ARINVOICE1header})

            ARINVOICE1header.Compose(New AccpacCOMAPI.AccpacView() {ARINVOICE1batch, ARINVOICE1detail1, ARINVOICE1detail2, ARINVOICE1detail3, Nothing})

            ARINVOICE1detail1.Compose(New AccpacCOMAPI.AccpacView() {ARINVOICE1header, ARINVOICE1batch, ARINVOICE1detail4})

            ARINVOICE1detail2.Compose(New AccpacCOMAPI.AccpacView() {ARINVOICE1header})

            ARINVOICE1detail3.Compose(New AccpacCOMAPI.AccpacView() {ARINVOICE1header})

            ARINVOICE1detail4.Compose(New AccpacCOMAPI.AccpacView() {ARINVOICE1detail1})


            ' ================================
            ' FIND BATCH OR CREATE NEW BATCH FOR AP INVOICE
            ' ================================
            If BATCHNBR = 0 Then
                ARINVOICE1batch.RecordCreate(1)

                ARINVOICE1batchFields.FieldByName("PROCESSCMD").PutWithoutVerification("1")      ' Process Command

                ARINVOICE1batch.Process()
                ARINVOICE1batch.Read()
                ARINVOICE1header.RecordCreate(2)
                ARINVOICE1detail1.Cancel()
                ARINVOICE1batchFields.FieldByName("DATEBTCH").Value = dRow("Date")    ' Batch Date
                ARINVOICE1batch.Update()
                ARINVOICE1batch.Read()

                ARINVOICE1headerFields.FieldByName("CNTITEM").Value = ENTRYNBR                        ' Entry Number
                ARINVOICE1header.Fetch()
                temp = ARINVOICE1header.Exists
                temp = ARINVOICE1header.Exists
                ARINVOICE1batch.Read()
                ARINVOICE1header.RecordCreate(2)
                ARINVOICE1detail1.Cancel()
            Else
                ARINVOICE1batchFields.FieldByName("CNTBTCH").Value = BATCHNBR                         ' Batch Number
                ARINVOICE1batch.Read()

                ARINVOICE1batchFields.FieldByName("PROCESSCMD").PutWithoutVerification("1")      ' Process Command

                ARINVOICE1batch.Process()

            End If


            ' ================================
            ' CREATE NEW ENTRY WITHIN BATCH FOR NEW CUSTOMER
            ' ================================
            If NEWCUST Then
                ARINVOICE1header.RecordCreate(2)
                ARINVOICE1detail1.Cancel()

                ARINVOICE1headerFields.FieldByName("IDCUST").Value = dRow("Customer")               ' Customer Number
                ARINVOICE1headerFields.FieldByName("PROCESSCMD").PutWithoutVerification("4")     ' Process Command
                ARINVOICE1header.Process()

                If dRow("ShipTo") <> vbNullString Then
                    If Not fCheckInvoiceTogether(dRow("Customer")) Then
                        ARINVOICE1headerFields.FieldByName("IDSHPT").Value = Left(dRow("ShipTo") + "000", 3)                        ' Ship-To Location Code
                        ARINVOICE1headerFields.FieldByName("PROCESSCMD").PutWithoutVerification("4")     ' Process Command
                        ARINVOICE1header.Process()
                    End If
                End If

                ARINVOICE1headerFields.FieldByName("INVCDESC").PutWithoutVerification((REFERENCE).ToString)
                'ARINVOICE1header.Insert()
                'ARINVOICE1detail1.Read()
                'ARINVOICE1detail1.Read()
                'ARINVOICE1batch.Read()
                'ARINVOICE1header.RecordCreate(2)
                'ARINVOICE1detail1.Cancel()

                LINENO = 0
            Else
                ARINVOICE1headerFields.FieldByName("CNTITEM").PutWithoutVerification(ENTRYNBR)   ' Entry Number
                ARINVOICE1header.Browse("", 1)
                ARINVOICE1header.Fetch()
                ARINVOICE1detail1.Read()
            End If

            ' ================================
            ' PROCESS INVOICE LINES
            ' ================================

            temp = ARINVOICE1detail1.Exists
            ARINVOICE1detail1.RecordClear()
            temp = ARINVOICE1detail1.Exists
            ARINVOICE1detail1.RecordCreate(0)

            ARINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")    ' Process Command Code

            ARINVOICE1detail1.Process()

            ARINVOICE1detail1Fields.FieldByName("IDDIST").Value = dRow("LineType")                    ' Distribution Code
            ARINVOICE1detail1Fields.FieldByName("TEXTDESC").Value = dRow("Description")         ' Description
            ARINVOICE1detail1Fields.FieldByName("AMTEXTN").Value = dRow("SubTotal")                   ' Extended Amount w/ TIP

            If dRow("LedgerCode") <> "" And dRow("LedgerCode") <> "-1" Then
                ARINVOICE1detail1Fields.FieldByName("IDACCTREV").Value = dRow("LedgerCode")                  ' Revenue Account
            End If

            ARINVOICE1detail1Fields.FieldByName("COMMENT").PutWithoutVerification(dRow("Reference"))

            ARINVOICE1detail1.Insert()

            ARINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification((LINENO + 1) * -1)      ' Line Number

            ARINVOICE1detail1.Read()

            LINENO = LINENO + 1

            ' ================================
            ' UPDATE DOCUMENT TOTAL
            ' ================================
            If ARINVOICE1headerFields.FieldByName("INVCDESC").Value = vbNullString Then
                ARINVOICE1headerFields.FieldByName("INVCDESC").PutWithoutVerification(REFERENCE)
            End If
            If dRow("PONumber") <> vbNullString Then
                ARINVOICE1headerFields.FieldByName("CUSTPO").PutWithoutVerification(dRow("PONumber"))
            End If

            BATCHNBR = ARINVOICE1batchFields.FieldByName("CNTBTCH").Value

            ' ================================
            ' ADD ENTRY
            ' ================================
            If NEWCUST Then
                ARINVOICE1header.Insert()
                ARINVOICE1detail1.Read()
                ARINVOICE1detail1.Read()
                ARINVOICE1batch.Read()
                ARINVOICE1header.RecordCreate(2)
                ARINVOICE1detail1.Cancel()
            Else
                ARINVOICE1detail1.Read()
                ARINVOICE1header.Update()
            End If

            ' ================================
            ' UPDATE GL CATEGORY
            ' ================================
            If (dRow("GLCategory") <> "") Then
                fUpdateGLCategoryOnARInvoice(REFERENCE, fGetValueInBrackets(Trim(dRow("GLCategory"))))
            End If

            fUpdateBookingDetailsWithInvoice(BOOKREF, BATCHNBR & "-" & ENTRYNBR)
            fUpdateInvoiceAsProcessed(BOOKREF, dRow("LineNum"), REFERENCE, BATCHNBR & "-" & ENTRYNBR)
            Return True


        Catch ex As Exception
            sError = Err.Description
            fHandleSageErrors(sError)
            sError = "ERR~~<Booking - " & BOOKREF & "> " & sError
            fLogIt("ERR~~fSAGEInsertSingleARInvoice. <Booking - " & BOOKREF & "> Error: " & sError)
            fUpdateInvoiceAsProcessed(BOOKREF, dRow("LineNum"), REFERENCE, "ERR")
            Return False
        End Try
    End Function


    Public Function fCreateInventoryReceiptOrShipment(ByVal sMirrorConnection As String, ByVal sParam As String) As String
        Dim sSQL As String = vbNullString
        Dim dsI, dsO As New DataSet
        Dim sError As String = vbNullString
        Try
            sSQL = <SQL>
                    SELECT HH.*, ISNULL(WW.GLCategory, '') AS GLCategory FROM tAPI_CreateNewInventoryReceipt HH
                    INNER JOIN tBooking_WarehouseMapping WW
                    ON HH.DisposalSite = WW.DisposalSite
                    WHERE IC_Guid = '**GUID**'
                </SQL>

            sSQL = sSQL.Replace("**GUID**", sParam)

            fGetMirrorDataset(sSQL, dsO)

            If Not dsO Is Nothing Then
                If dsO.Tables.Count > 0 Then
                    If dsO.Tables(0).Rows.Count > 0 Then
                        If dsO.Tables(0).Rows(0).Item("Qty") > 0 Then
                            If dsO.Tables(0).Rows(0).Item("Type") = "Receipt" Then
                                fSAGEInsertInventoryReceipt(sMirrorConnection, dsO, sError)
                            Else
                                fSAGEInsertInventoryShipment(sMirrorConnection, dsO, sError)
                            End If
                        End If
                    Else
                        sError = "ERR~~fCreateInventoryReceiptOrShipment> No records found: " & vbCrLf & sSQL
                    End If
                End If
            End If

            If Mid(sError, 1, 5) = "ERR~~" Then
                fLogIt("ERR~~fCreateInventoryReceiptOrShipment> Error: " & sError)
                Return "ERR~~fCreateInventoryReceiptOrShipment> Error: " & sError
            End If

        Catch ex As Exception
            fLogIt("ERR~~fCreateInventoryReceiptOrShipment> Error: " & Err.Description)
            Return "ERR~~fCreateInventoryReceiptOrShipment> Error: " & Err.Description
        Finally
            If Not dsI Is Nothing Then
                dsI.Dispose()
            End If
            If Not dsO Is Nothing Then
                dsO.Dispose()
            End If
        End Try
    End Function


    Public Function fCreateARorAPInvoice(ByVal sMirrorConnection As String, ByVal sParam As String) As String
        Dim sSQL As String = vbNullString
        Dim dsI, dsO As New DataSet
        Dim sError As String = vbNullString
        Try
            sSQL = <SQL>
                    SELECT HH.INV_Guid, HH.Type, HH.Date, HH.Reference, HH.PONumber,
                    HH.Location, HH.Customer, HH.CustomerName, LL.LineType, LL.Description,
                    LL.SubTotal, LL.TaxTotal, LL.Total, ISNULL(WW.GLCategory, '') AS GLCategory, LL.LedgerCode
                    FROM tAPI_CreateNewInvoice_HDR HH
                    INNER JOIN tAPI_CreateNewInvoice_Lines LL ON HH.INV_Guid = LL.INV_Guid AND HH.Reference = LL.Reference
                    INNER JOIN tBooking_WarehouseMapping WW ON HH.Location = WW.Warehouse
                    WHERE HH.RequestGUID = '**GUID**'
                </SQL>

            sSQL = sSQL.Replace("**GUID**", sParam)

            fGetMirrorDataset(sSQL, dsO)

            If Not dsO Is Nothing Then
                If dsO.Tables.Count > 0 Then
                    If dsO.Tables(0).Rows.Count > 0 Then
                        If dsO.Tables(0).Rows(0).Item("Type") = "Invoice" Then
                            fSAGEInsertARInvoice(sMirrorConnection, dsO, sError)
                        Else
                            fSAGEInsertAPInvoice(sMirrorConnection, dsO, sError)
                        End If
                    Else
                        sError = "ERR~~fCreateARorAPInvoice> No records found: " & vbCrLf & sSQL
                    End If
                End If
            End If

            If Mid(sError, 1, 5) = "ERR~~" Then
                fLogIt("ERR~~fCreateARorAPInvoice> Error: " & sError)
                Return "ERR~~fCreateARorAPInvoice> Error: " & sError
            End If

        Catch ex As Exception
            fLogIt("ERR~~fCreateARorAPInvoice> Error: " & Err.Description)
            Return "ERR~~fCreateARorAPInvoice> Error: " & Err.Description
        Finally
            If Not dsI Is Nothing Then
                dsI.Dispose()
            End If
            If Not dsO Is Nothing Then
                dsO.Dispose()
            End If
        End Try
    End Function

    Public Function fCreateARorAPInvoice_NEW(ByVal sMirrorConnection As String, ByVal sParam As String) As String
        Dim sSQL As String = vbNullString
        Dim dsI, dsO As New DataSet
        Dim sError As String = vbNullString
        Dim sColError As String = vbNullString
        'Dim bInvoiceTogether As Boolean = False
        Try
            sSQL = <SQL>
                    SELECT HH.INV_Guid, HH.Type, HH.Date, HH.Reference, HH.PONumber, HH.DisposalSite, LL.LineNum,
                    HH.Location, HH.Customer, ISNULL(HH.ShipTo, '') AS ShipTo, HH.CustomerName, LL.LineType, LL.Description,
                    LL.SubTotal, LL.TaxTotal, LL.Total, ISNULL(WW.GLCategory, '') AS GLCategory, LL.LedgerCode
                    FROM tAPI_CreateNewInvoice_HDR HH
                    INNER JOIN tAPI_CreateNewInvoice_Lines LL ON HH.INV_Guid = LL.INV_Guid AND HH.Reference = LL.Reference
                    INNER JOIN tBooking_WarehouseMapping WW ON HH.DisposalSite = WW.DisposalSite
                    WHERE HH.RequestGUID = '**GUID**' AND LL.Processed IS NULL
                    ORDER BY HH.Type, HH.Customer, ISNULL(HH.ShipTo, ''), LL.Total DESC
                </SQL>

            sSQL = sSQL.Replace("**GUID**", sParam)

            fGetMirrorDataset(sSQL, dsO)

            If Not dsO Is Nothing Then
                If dsO.Tables.Count > 0 Then
                    If dsO.Tables(0).Rows.Count > 0 Then
                        Dim dTOTAL As Double = 0
                        Dim BATCHNBR As Long = 0
                        Dim ENTRYNBR As Long = 0
                        Dim LINENO As Long = 0
                        Dim bNEW As Boolean = False
                        Dim REFERENCE As String = vbNullString

                        Dim CurrentGuid As String = vbNullString
                        Dim CurrentCustomer As String = vbNullString
                        Dim sType As String = "Invoice"

                        If dsO.Tables(0).Rows(0).Item("Type") = "Invoice" Then
                            sType = "Invoice"
                        Else
                            sType = "Bill"
                        End If

                        Dim dRecordCount As Long = 0

                        For Each dR As DataRow In dsO.Tables(0).Rows
                            If sType <> dR("Type") Then
                                CurrentCustomer = vbNullString
                                BATCHNBR = 0
                                ENTRYNBR = 0
                                LINENO = 0
                                'bInvoiceTogether = fCheckInvoiceTogether(dR("Customer"))
                            End If

                            sType = dR("Type")

                            If Trim(dR("INV_Guid").ToString) <> CurrentGuid Then
                                bNEW = True
                                If dRecordCount > 0 Then
                                    Exit For
                                End If

                                ENTRYNBR = ENTRYNBR + 1

                                CurrentGuid = Trim(dR("INV_Guid").ToString)
                                REFERENCE = dR("INV_Guid").ToString
                                dRecordCount = dRecordCount + 1
                            Else
                                bNEW = False
                            End If

                            sError = vbNullString

                            fLogIt("# fCreateARorAPInvoice_NEW > [" & REFERENCE & "] BATCH - " & BATCHNBR & ", ENTRY - " & ENTRYNBR & ", LINE - " & LINENO)

                            'oSage = New clsSage
                            'If Not oSage.fConnect(gsSageUser, gsSagePassword, gsSageDatabase) Then
                            '    fErr("ERR~~fInitSage> Error connecting to Sage 300", 0)
                            '    Return False
                            'End If

                            If sType = "Invoice" Then
                                fSAGEInsertSingleARInvoice(sMirrorConnection, BATCHNBR, ENTRYNBR, LINENO, dTOTAL, bNEW, REFERENCE, dR, sError)
                            Else
                                fSAGEInsertSingleAPInvoice(sMirrorConnection, BATCHNBR, ENTRYNBR, LINENO, dTOTAL, bNEW, REFERENCE, dR, sError)
                            End If


                            fLogIt("# fCreateARorAPInvoice_NEW > [" & REFERENCE & "] BATCH - " & BATCHNBR & ", ENTRY - " & ENTRYNBR & ", LINE - " & LINENO)

                            If Not sError Is Nothing Then
                                If sError <> vbNullString Then
                                    sColError = sError & vbCrLf & sColError
                                End If
                            End If

                        Next
                    Else
                        'sColError = "ERR~~fCreateARorAPInvoice_NEW> No records found: " & vbCrLf & sSQL
                        gbCompleted = True
                    End If
                End If
            End If

            If Mid(sColError, 1, 5) = "ERR~~" Then
                fLogIt("ERR~~fCreateARorAPInvoice_NEW> Error: " & sColError)
                Return "ERR~~fCreateARorAPInvoice_NEW> Error: " & sColError
            End If

        Catch ex As Exception
            fLogIt("ERR~~fCreateARorAPInvoice_NEW> Error: " & Err.Description)
            Return "ERR~~fCreateARorAPInvoice_NEW> Error: " & Err.Description
        Finally
            If Not dsI Is Nothing Then
                dsI.Dispose()
            End If
            If Not dsO Is Nothing Then
                dsO.Dispose()
            End If
        End Try
    End Function
End Class

