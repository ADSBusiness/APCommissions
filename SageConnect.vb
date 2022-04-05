Imports System.Data.SqlClient
Imports Microsoft.Data.SqlClient


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

        sSession.Init("", "XY", "XY0001", "67A")
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
            APInvBatchDetail4(0) = APINVOICE1detail1
            '            APINVOICE1detail4.Compose(APInvBatchDetail1)
            APINVOICE1detail4.Compose(APINVOICE1detail1)

            '            APINVOICE1batch.Compose Array(APINVOICE1header)
            '           APINVOICE1header.Compose Array(APINVOICE1batch, APINVOICE1detail1, APINVOICE1detail2, APINVOICE1detail3)
            '          APINVOICE1detail1.Compose Array(APINVOICE1header, APINVOICE1batch, APINVOICE1detail4)
            '         APINVOICE1detail2.Compose Array(APINVOICE1header)
            '        APINVOICE1detail3.Compose Array(APINVOICE1header)
            '       APINVOICE1detail4.Compose Array(APINVOICE1detail1)




            DeclareViews = True

        Catch ex As Exception

            AccpacErrorHandler()

            DeclareViews = False

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
                Case vbYes
                    sSession.Errors.Clear()
                    Resume Next
                Case vbNo
                    sSession.Errors.Clear()
                    Exit Sub
            End Select
            sSession.Errors.Clear()
        End If


    End Sub


    Public Sub PrintRCTI()

        Try
            sSession.Errors.Clear()

            Dim rpt As AccpacCOMAPI.AccpacReport = sSession.ReportSelect("APRCTI[APRCTI01.RPT]", " ", " ")
            Dim rptPrintSetup As AccpacCOMAPI.AccpacPrintSetup = sSession.GetPrintSetupGetPrintSetup("      ", "      ")
            rpt.NumOfCopies = 1
            '            rpt.PrintDir = ""

            rpt.SetParam("BTCHNUM", "1380")
            rpt.Destination = AccpacCOMAPI.tagPrintDestinationEnum.PD_PREVIEW

            rpt.PrintReport()



        Catch ex As Exception
            AccpacErrorHandler()
        End Try







    End Sub


    Sub APInvoice()

        '        Do While Cells(RowID, 2) <> ""
        '        Cells(RowID, 1).Select
        '
        ' Add 3 lines of costs
        ' Lopp through comms and sum
        ' loop through datagrid and group by vendor
        '
        Dim sSPCode As String = ""
        Dim sOrdNumb As String = ""
        Dim sSPState As String
        Dim bFirstLine As Boolean = True
        Dim bFirstDetail As Boolean = True
        Dim dInvAmt As Double = 0
        Dim dTotComm As Double
        Dim rRow As Integer
        Dim tempp As String
        rRow = 1
        Try
            Dim AP_InvBatch As String = Sage_AP_CreateBatch("Commissions - " & sSageUserID, Format(Today, "dd/MM/yyyy"))

            ' loop through records
            For Each sitm As ListViewItem In frmMain.ListView1.Items
                Try

                    If sitm.Checked = True Then
                        '   sSPCode = ""
                        If sSPCode <> sitm.SubItems.Item(1).Text Or sSPCode = "" Then

                            sSPCode = sitm.SubItems.Item(1).Text
                            sOrdNumb = ""
                            Dim txtVal As String

                            txtVal = frmMain.ListView1.Items(rRow).SubItems(7).Text

                            sSPCode = sitm.SubItems.Item(1).Text
                            bFirstLine = True
                            bFirstDetail = True
                            APINVOICE1batch.Process()
                            APINVOICE1batchFields.FieldByName("CNTBTCH").Value = AP_InvBatch
                            '            Temp = APINVOICE1header.Exists
                            APINVOICE1batch.Read()
                            APINVOICE1header.RecordCreate(2)
                            APINVOICE1detail1.Cancel()
                            APINVOICE1headerFields.FieldByName("IDVEND").Value = sitm.SubItems.Item(1).Text
                            APINVOICE1headerFields.FieldByName("Processcmd").PutWithoutVerification("7")
                            APINVOICE1header.Process()
                            APINVOICE1headerFields.FieldByName("Processcmd").PutWithoutVerification("7")
                            APINVOICE1header.Process()
                            APINVOICE1headerFields.FieldByName("Processcmd").PutWithoutVerification("4")
                            APINVOICE1header.Process()




                            APINVOICE1headerFields.FieldByName("TAXCLASS1").Value = "1"
                            APINVOICE1headerFields.FieldByName("DATEINVC").Value = "01/03/2022"


                            tempp = Format(System.DateTime.Now, "HH:mm:ss")
                            ' TODO: change the IDINVC value to suitable for real data
                            APINVOICE1headerFields.FieldByName("IDINVC").Value = "12345c" & "-" & tempp
                            APINVOICE1headerFields.FieldByName("PONBR").PutWithoutVerification("PO1")
                            APINVOICE1headerFields.FieldByName("TEXTTRX").Value = "1"  '
                        End If

                        If bFirstLine = True Then
                            APINVOICE1detail1.Read()
                            Temp = APINVOICE1detail1.Exists
                            If Temp = True Then
                                APINVOICE1detail1.Delete()
                                APINVOICE1detail1.Process()
                                APINVOICE1detail1.Read()
                            End If
                            bFirstLine = False
                        End If
                        '
                        '
                        '
                        ' add a line for each ORDNUMBER Perf Bonus
                        If (sOrdNumb <> sitm.SubItems.Item(6).Text) And (sSPCode = sitm.SubItems.Item(1).Text) Then
                            sOrdNumb = sitm.SubItems.Item(6).Text
                            Temp = APINVOICE1detail1.Exists
                            APINVOICE1detail1.RecordCreate(0)
                            APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")
                            APINVOICE1detail1.Process()
                            APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification("-1")
                            APINVOICE1detail1.Read()
                            APINVOICE1detail1.RecordCreate(0)
                            APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")
                            APINVOICE1detail1.Process()

                            sSPState = FindSPState(sSPCode)
                            APINVOICE1detail1Fields.FieldByName("IDGLACCT").Value = "62060-" & Trim(sSPState) & "-3"
                            APINVOICE1detail1Fields.FieldByName("TEXTDESC").Value = Trim(sOrdNumb) & " - Bonus"

                            'Call AddOptField("APORD", sitm.SubItems.Item(6).Text)
                            'Call AddOptField("APCUST", sitm.SubItems.Item(3).Text)
                            'Call AddOptField("APTYPE", sitm.SubItems.Item(14).Text)
                            'Call AddOptField("APSRCE", sitm.SubItems.Item(15).Text)
                            'Call AddOptField("APFT", sitm.SubItems.Item(19).Text)
                            'Call AddOptField("APSPLIT", sitm.SubItems.Item(17).Text)
                            'Call AddOptField("APAMT", sitm.SubItems.Item(18).Text)

                            APINVOICE1detail1.Insert()
                            APINVOICE1detail1.Read()

                            APINVOICE1detail1Fields.FieldByName("AMTDIST").Value = 0 'frmMain.ListView1.Items(rRow).SubItems(18).Text
                            ' dInvAmt += 0 'frmMain.ListView1.Items(rRow).SubItems(18).Text
                            APINVOICE1detail1.Update()
                            APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification("-1")
                            APINVOICE1detail1.Read()
                            Temp = APINVOICE1detail1.Exists
                            APINVOICE1detail1.RecordCreate(0)
                            APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")
                            APINVOICE1detail1.Process()
                            APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification("-1")
                            APINVOICE1detail1.Read()

                            ' add a line for each ORDNUMBER Finance Charges
                            Temp = APINVOICE1detail1.Exists
                            APINVOICE1detail1.RecordCreate(0)
                            APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")
                            APINVOICE1detail1.Process()
                            APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification("-1")
                            APINVOICE1detail1.Read()
                            APINVOICE1detail1.RecordCreate(0)
                            APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")
                            APINVOICE1detail1.Process()

                            'sSPState = FindSPState(sSPCode)
                            APINVOICE1detail1Fields.FieldByName("IDGLACCT").Value = "50950-" & Trim(sSPState) & "-3"
                            APINVOICE1detail1Fields.FieldByName("TEXTDESC").Value = Trim(sOrdNumb) & " - Finance"

                            'Call AddOptField("APORD", sitm.SubItems.Item(6).Text)
                            'Call AddOptField("APCUST", sitm.SubItems.Item(3).Text)
                            'Call AddOptField("APTYPE", sitm.SubItems.Item(14).Text)
                            'Call AddOptField("APSRCE", sitm.SubItems.Item(15).Text)
                            'Call AddOptField("APFT", sitm.SubItems.Item(19).Text)
                            'Call AddOptField("APSPLIT", sitm.SubItems.Item(17).Text)
                            'Call AddOptField("APAMT", sitm.SubItems.Item(18).Text)

                            APINVOICE1detail1.Insert()
                            APINVOICE1detail1.Read()

                            APINVOICE1detail1Fields.FieldByName("AMTDIST").Value = 0 'frmMain.ListView1.Items(rRow).SubItems(18).Text
                            dInvAmt += 0 'frmMain.ListView1.Items(rRow).SubItems(18).Text
                            APINVOICE1detail1.Update()
                            APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification("-1")
                            APINVOICE1detail1.Read()
                            Temp = APINVOICE1detail1.Exists
                            APINVOICE1detail1.RecordCreate(0)
                            APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")
                            APINVOICE1detail1.Process()
                            APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification("-1")
                            APINVOICE1detail1.Read()


                        End If


                        If (sOrdNumb <> sitm.SubItems.Item(6).Text) And (sSPCode = sitm.SubItems.Item(1).Text) Then
                            sOrdNumb = sitm.SubItems.Item(6).Text
                            Temp = APINVOICE1detail1.Exists
                            APINVOICE1detail1.RecordCreate(0)
                            APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")
                            APINVOICE1detail1.Process()
                            APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification("-1")
                            APINVOICE1detail1.Read()
                            APINVOICE1detail1.RecordCreate(0)
                            APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")
                            APINVOICE1detail1.Process()
                            ' TODO: Determine SP State and adjust GL Code for perf bonus
                            sSPState = FindSPState(sSPCode)
                            APINVOICE1detail1Fields.FieldByName("IDGLACCT").Value = "50950-" & Trim(sSPState) & "-3"
                            APINVOICE1detail1Fields.FieldByName("TEXTDESC").Value = sOrdNumb & " - " & sitm.SubItems.Item(5).Text

                            Call AddOptField("APORD", sitm.SubItems.Item(6).Text)
                            Call AddOptField("APCUST", sitm.SubItems.Item(3).Text)
                            Call AddOptField("APTYPE", sitm.SubItems.Item(14).Text)
                            Call AddOptField("APSRCE", sitm.SubItems.Item(15).Text)
                            Call AddOptField("APFT", sitm.SubItems.Item(19).Text)
                            Call AddOptField("APSPLIT", sitm.SubItems.Item(17).Text)
                            Call AddOptField("APAMT", sitm.SubItems.Item(18).Text)
                            Call AddOptField("APITEM", sitm.SubItems.Item(7).Text)

                            'APINVOICE1detail1.Process()
                            APINVOICE1detail1.Insert()
                            APINVOICE1detail1.Read()

                            APINVOICE1detail1Fields.FieldByName("AMTDIST").Value = 0 'frmMain.ListView1.Items(rRow).SubItems(18).Text
                            ' dInvAmt += 0 'frmMain.ListView1.Items(rRow).SubItems(18).Text
                            APINVOICE1detail1.Update()
                            APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification("-1")
                            APINVOICE1detail1.Read()
                            Temp = APINVOICE1detail1.Exists
                            APINVOICE1detail1.RecordCreate(0)
                            APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")
                            APINVOICE1detail1.Process()
                            APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification("-1")
                            APINVOICE1detail1.Read()
                        End If

                        If bFirstDetail = True Then
                            ' find total for comms $ value and add a single line
                            Temp = APINVOICE1detail1.Exists
                            APINVOICE1detail1.RecordCreate(0)
                            APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")
                            APINVOICE1detail1.Process()
                            APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification("-1")
                            APINVOICE1detail1.Read()
                            APINVOICE1detail1.RecordCreate(0)
                            APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")
                            APINVOICE1detail1.Process()
                            APINVOICE1detail1Fields.FieldByName("IDGLACCT").Value = "30200-3"
                            APINVOICE1detail1Fields.FieldByName("TEXTDESC").Value = sOrdNumb & " - Total"

                            Call AddOptField("APORD", sitm.SubItems.Item(6).Text)
                            Call AddOptField("APCUST", sitm.SubItems.Item(3).Text)
                            Call AddOptField("APTYPE", sitm.SubItems.Item(14).Text)
                            Call AddOptField("APSRCE", sitm.SubItems.Item(15).Text)
                            Call AddOptField("APFT", sitm.SubItems.Item(19).Text)
                            Call AddOptField("APSPLIT", sitm.SubItems.Item(17).Text)
                            Call AddOptField("APAMT", sitm.SubItems.Item(18).Text)
                            Call AddOptField("APITEM", sitm.SubItems.Item(7).Text)

                            'APINVOICE1detail1.Process()
                            APINVOICE1detail1.Insert()
                            APINVOICE1detail1.Read()
                            dTotComm = Math.Round(FindTotalComms(sSPCode), 2)
                            APINVOICE1detail1Fields.FieldByName("AMTDIST").Value = dTotComm
                            dInvAmt += dTotComm
                            APINVOICE1detail1.Update()
                            dInvAmt += APINVOICE1detail1Fields.FieldByName("AMTTAX1").Value

                            APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification("-1")
                            APINVOICE1detail1.Read()
                            Temp = APINVOICE1detail1.Exists
                            APINVOICE1detail1.RecordCreate(0)
                            APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")
                            APINVOICE1detail1.Process()
                            APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification("-1")
                            APINVOICE1detail1.Read()

                            bFirstDetail = False
                        End If

                        Temp = APINVOICE1detail1.Exists
                        APINVOICE1detail1.RecordCreate(0)
                        APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")
                        APINVOICE1detail1.Process()
                        APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification("-1")
                        APINVOICE1detail1.Read()
                        APINVOICE1detail1.RecordCreate(0)
                        APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")
                        APINVOICE1detail1.Process()
                        APINVOICE1detail1Fields.FieldByName("IDGLACCT").Value = "30200-3"
                        APINVOICE1detail1Fields.FieldByName("TEXTDESC").Value = sOrdNumb

                        Call AddOptField("APORD", sitm.SubItems.Item(6).Text)
                        Call AddOptField("APCUST", sitm.SubItems.Item(3).Text)
                        Call AddOptField("APTYPE", sitm.SubItems.Item(14).Text)
                        Call AddOptField("APSRCE", sitm.SubItems.Item(15).Text)
                        Call AddOptField("APFT", sitm.SubItems.Item(19).Text)
                        Call AddOptField("APSPLIT", sitm.SubItems.Item(17).Text)
                        Call AddOptField("APAMT", sitm.SubItems.Item(18).Text)
                        Call AddOptField("APITEM", sitm.SubItems.Item(7).Text)

                        'APINVOICE1detail1.Process()
                        APINVOICE1detail1.Insert()
                        APINVOICE1detail1.Read()
                        '  dTotComm = sitm.SubItems.Item(18).Text 'Math.Round(FindTotalComms(sSPCode), 2)
                        APINVOICE1detail1Fields.FieldByName("AMTDIST").Value = (0) 'dTotComm
                        ' dInvAmt += dTotComm
                        APINVOICE1detail1.Update()
                        APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification("-1")
                        APINVOICE1detail1.Read()
                        Temp = APINVOICE1detail1.Exists
                        APINVOICE1detail1.RecordCreate(0)
                        APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")
                        APINVOICE1detail1.Process()
                        APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification("-1")
                        APINVOICE1detail1.Read()

                        '
                        '  AP Detail Opt Fields
                        'APINVOICE3header.Process

                        '
                        '
                        '

                        tempp = frmMain.ListView1.Items(rRow).SubItems(1).Text

                            If sSPCode <> frmMain.ListView1.Items(rRow).SubItems(1).Text Then
                                ''sitm.SubItems.selecteditem.index
                                APINVOICE1headerFields.FieldByName("TAXCLASS1").Value = "1"
                                APINVOICE1headerFields.FieldByName("AMTGROSTOT").Value = dInvAmt
                                dInvAmt = 0
                                APINVOICE1header.Insert()
                                APINVOICE1batch.Read()
                                Temp = APINVOICE1header.Exists
                                APINVOICE1header.RecordCreate(2)
                                APINVOICE1detail1.Cancel()


                            End If
                            rRow = rRow + 1

                        End If

                Catch ex As Exception
                    AccpacErrorHandler()
                End Try
                'APINVOICE3headerFields("ORDRNBR").PutWithoutVerification ("ORDNUM")
            Next

        Catch ex As Exception
            AccpacErrorHandler()
        End Try



    End Sub



    Function FindTotalComms(spCode As String) As Double

        FindTotalComms = False
        Dim A4W As New SqlConnection()
        Dim SQLConStr As String = "Server=" & frmMain.sqlServer & ";Database=" & frmMain.sqlDB & ";User ID=" & frmMain.sqlUser & ";Password=" & frmMain.sqlPswd
        Dim Read_TotalComs As SqlDataReader
        Dim vTotalComms As SqlCommand

        Dim sSQl_TotalComms As String

        ' caters for removing discounts...
        sSQl_TotalComms = "SELECT SPCode, SUM(ScommVal * TotRecComm) AS SPComm FROM v_APComm GROUP BY SPCode HAVING  (SPCode = '" & spCode & "') "
        A4W = New SqlConnection(SQLConStr)

        Try
            A4W.Open()
            vTotalComms = New SqlCommand(sSQl_TotalComms, A4W)
            Read_TotalComs = vTotalComms.ExecuteReader()

            While Read_TotalComs.Read()

                FindTotalComms = Read_TotalComs.Item(1)

            End While

            Read_TotalComs.Close()
            vTotalComms.Dispose()
            A4W.Close()


        Catch ex As Exception
            MsgBox("SQL data - Cannot total Commison $ for -  " & spCode)
            ' ================================
            frmMain.strLogLine = System.DateTime.Now & " - " & sSageOrgID & " - " & sSageUserID & "  -  " & "SQL data - Cannot update OrdUniq FOB" & spCode
            My.Computer.FileSystem.WriteAllText(frmMain.LogFile, frmMain.strLogLine & vbCrLf, True)
        End Try



    End Function


    Function FindSPState(spCode As String) As String

        FindSPState = False
        Dim A4W As New SqlConnection()
        Dim SQLConStr As String = "Server=" & frmMain.sqlServer & ";Database=" & frmMain.sqlDB & ";User ID=" & frmMain.sqlUser & ";Password=" & frmMain.sqlPswd
        Dim Read_ARSAP As SqlDataReader
        Dim vARSAP As SqlCommand

        Dim sSQl_ARSAP As String

        sSQl_ARSAP = "select CODESLSP,CODEEMPL from ARSAP  where CODESLSP= '" & spCode & "' "
        A4W = New SqlConnection(SQLConStr)

        Try
            A4W.Open()
            vARSAP = New SqlCommand(sSQl_ARSAP, A4W)
            Read_ARSAP = vARSAP.ExecuteReader()

            While Read_ARSAP.Read()

                FindSPState = Read_ARSAP.Item(1)

            End While

            Read_ARSAP.Close()
            vARSAP.Dispose()
            A4W.Close()


        Catch ex As Exception
            MsgBox("SQL data - Cannot SP State for -  " & spCode)
            ' ================================
            frmMain.strLogLine = System.DateTime.Now & " - " & sSageOrgID & " - " & sSageUserID & "  -  " & "SQL data - Cannot SP State for -  " & spCode
            My.Computer.FileSystem.WriteAllText(frmMain.LogFile, frmMain.strLogLine & vbCrLf, True)
        End Try



    End Function


    Sub APInvoice_OLD()
        ' set OLD 31-03-2022
        ' 

        '        Do While Cells(RowID, 2) <> ""
        '        Cells(RowID, 1).Select
        '
        ' Add 3 lines of costs
        ' Lopp through comms and sum
        ' loop through datagrid and group by vendor
        '
        Dim sSPCode As String = ""
        Dim sOrdNumb As String = ""
        Dim sSPState As String
        Dim bFirstLine As Boolean = True
        Dim bFirstDetail As Boolean = True
        Dim dInvAmt As Double = 0
        Dim dTotComm As Double
        Dim rRow As Integer
        Dim tempp As String
        rRow = 1
        Try
            Dim AP_InvBatch As String = Sage_AP_CreateBatch("Commissions - " & sSageUserID, Format(Today, "dd/MM/yyyy"))

            ' loop through records
            For Each sitm As ListViewItem In frmMain.ListView1.Items
                Try

                    If sitm.Checked = True Then



                        '   sSPCode = ""
                        If sSPCode <> sitm.SubItems.Item(1).Text Or sSPCode = "" Then

                            sSPCode = sitm.SubItems.Item(1).Text
                            sOrdNumb = ""
                            Dim txtVal As String

                            txtVal = frmMain.ListView1.Items(rRow).SubItems(7).Text

                            sSPCode = sitm.SubItems.Item(1).Text
                            bFirstLine = True
                            bFirstDetail = True
                            APINVOICE1batch.Process()
                            APINVOICE1batchFields.FieldByName("CNTBTCH").Value = AP_InvBatch
                            '            Temp = APINVOICE1header.Exists
                            APINVOICE1batch.Read()
                            APINVOICE1header.RecordCreate(2)
                            APINVOICE1detail1.Cancel()
                            APINVOICE1headerFields.FieldByName("IDVEND").Value = sitm.SubItems.Item(1).Text
                            APINVOICE1headerFields.FieldByName("Processcmd").PutWithoutVerification("7")
                            APINVOICE1header.Process()

                            APINVOICE1headerFields.FieldByName("Processcmd").PutWithoutVerification("4")
                            APINVOICE1header.Process()
                            APINVOICE1headerFields.FieldByName("TAXCLASS1").Value = "1"
                            APINVOICE1headerFields.FieldByName("DATEINVC").Value = "01/03/2022"

                            tempp = Format(System.DateTime.Now, "HH:mm:ss")
                            ' TODO: change the IDINVC value to suitable for real data
                            APINVOICE1headerFields.FieldByName("IDINVC").Value = "12345c" & "-" & tempp
                            APINVOICE1headerFields.FieldByName("PONBR").PutWithoutVerification("PO1")
                            APINVOICE1headerFields.FieldByName("TEXTTRX").Value = "1"  '
                        End If

                        If bFirstLine = True Then
                            APINVOICE1detail1.Read()
                            Temp = APINVOICE1detail1.Exists
                            If Temp = True Then
                                APINVOICE1detail1.Delete()
                                APINVOICE1detail1.Process()
                                APINVOICE1detail1.Read()
                            End If
                            bFirstLine = False
                        End If
                        '
                        '
                        '
                        'If bFirstDetail = True Then
                        ' find total for comms $ value and add a single line
                        Temp = APINVOICE1detail1.Exists
                        APINVOICE1detail1.RecordCreate(0)
                        APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")
                        APINVOICE1detail1.Process()
                        APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification("-1")
                        APINVOICE1detail1.Read()
                        APINVOICE1detail1.RecordCreate(0)
                        APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")
                        APINVOICE1detail1.Process()
                        APINVOICE1detail1Fields.FieldByName("IDGLACCT").Value = "30200-3"
                        APINVOICE1detail1.Insert()
                        APINVOICE1detail1.Read()
                        dTotComm = sitm.SubItems.Item(18).Text 'Math.Round(FindTotalComms(sSPCode), 2)
                        APINVOICE1detail1Fields.FieldByName("AMTDIST").Value = dTotComm
                        dInvAmt += dTotComm
                        APINVOICE1detail1.Update()
                        APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification("-1")
                        APINVOICE1detail1.Read()
                        Temp = APINVOICE1detail1.Exists
                        APINVOICE1detail1.RecordCreate(0)
                        APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")
                        APINVOICE1detail1.Process()
                        APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification("-1")
                        APINVOICE1detail1.Read()

                        bFirstDetail = False
                        'End If

                        ' add a line for each ORDNUMBER Perf Bonus

                        If (sOrdNumb <> sitm.SubItems.Item(6).Text) And (sSPCode = sitm.SubItems.Item(1).Text) Then
                            sOrdNumb = sitm.SubItems.Item(6).Text
                            Temp = APINVOICE1detail1.Exists
                            APINVOICE1detail1.RecordCreate(0)
                            APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")
                            APINVOICE1detail1.Process()
                            APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification("-1")
                            APINVOICE1detail1.Read()
                            APINVOICE1detail1.RecordCreate(0)
                            APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")
                            APINVOICE1detail1.Process()

                            sSPState = FindSPState(sSPCode)
                            APINVOICE1detail1Fields.FieldByName("IDGLACCT").Value = "62060-" & Trim(sSPState) & "-3"
                            APINVOICE1detail1Fields.FieldByName("TEXTDESC").Value = Trim(sOrdNumb) & " - Bonus"

                            APINVOICE1detail1.Insert()
                            APINVOICE1detail1.Read()

                            APINVOICE1detail1Fields.FieldByName("AMTDIST").Value = 0 'frmMain.ListView1.Items(rRow).SubItems(18).Text
                            dInvAmt += 0 'frmMain.ListView1.Items(rRow).SubItems(18).Text
                            APINVOICE1detail1.Update()
                            APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification("-1")
                            APINVOICE1detail1.Read()
                            Temp = APINVOICE1detail1.Exists
                            APINVOICE1detail1.RecordCreate(0)
                            APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")
                            APINVOICE1detail1.Process()
                            APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification("-1")
                            APINVOICE1detail1.Read()

                            ' add a line for each ORDNUMBER Finance Charges
                            Temp = APINVOICE1detail1.Exists
                            APINVOICE1detail1.RecordCreate(0)
                            APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")
                            APINVOICE1detail1.Process()
                            APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification("-1")
                            APINVOICE1detail1.Read()
                            APINVOICE1detail1.RecordCreate(0)
                            APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")
                            APINVOICE1detail1.Process()

                            'sSPState = FindSPState(sSPCode)
                            APINVOICE1detail1Fields.FieldByName("IDGLACCT").Value = "50950-" & Trim(sSPState) & "-3"
                            APINVOICE1detail1Fields.FieldByName("TEXTDESC").Value = Trim(sOrdNumb) & " - Finance"

                            APINVOICE1detail1.Insert()
                            APINVOICE1detail1.Read()

                            APINVOICE1detail1Fields.FieldByName("AMTDIST").Value = 0 'frmMain.ListView1.Items(rRow).SubItems(18).Text
                            dInvAmt += 0 'frmMain.ListView1.Items(rRow).SubItems(18).Text
                            APINVOICE1detail1.Update()
                            APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification("-1")
                            APINVOICE1detail1.Read()
                            Temp = APINVOICE1detail1.Exists
                            APINVOICE1detail1.RecordCreate(0)
                            APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")
                            APINVOICE1detail1.Process()
                            APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification("-1")
                            APINVOICE1detail1.Read()


                        End If


                        If (sOrdNumb <> sitm.SubItems.Item(6).Text) And (sSPCode = sitm.SubItems.Item(1).Text) Then
                            sOrdNumb = sitm.SubItems.Item(6).Text
                            Temp = APINVOICE1detail1.Exists
                            APINVOICE1detail1.RecordCreate(0)
                            APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")
                            APINVOICE1detail1.Process()
                            APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification("-1")
                            APINVOICE1detail1.Read()
                            APINVOICE1detail1.RecordCreate(0)
                            APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")
                            APINVOICE1detail1.Process()
                            ' TODO: Determine SP State and adjust GL Code for perf bonus
                            sSPState = FindSPState(sSPCode)
                            APINVOICE1detail1Fields.FieldByName("IDGLACCT").Value = "50950-" & Trim(sSPState) & "-3"
                            APINVOICE1detail1Fields.FieldByName("TEXTDESC").Value = sOrdNumb

                            APINVOICE1detail1.Insert()
                            APINVOICE1detail1.Read()

                            APINVOICE1detail1Fields.FieldByName("AMTDIST").Value = 0 'frmMain.ListView1.Items(rRow).SubItems(18).Text
                            dInvAmt += 0 'frmMain.ListView1.Items(rRow).SubItems(18).Text
                            APINVOICE1detail1.Update()
                            APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification("-1")
                            APINVOICE1detail1.Read()
                            Temp = APINVOICE1detail1.Exists
                            APINVOICE1detail1.RecordCreate(0)
                            APINVOICE1detail1Fields.FieldByName("PROCESSCMD").PutWithoutVerification("0")
                            APINVOICE1detail1.Process()
                            APINVOICE1detail1Fields.FieldByName("CNTLINE").PutWithoutVerification("-1")
                            APINVOICE1detail1.Read()
                        End If


                        tempp = frmMain.ListView1.Items(rRow).SubItems(1).Text

                        If sSPCode <> frmMain.ListView1.Items(rRow).SubItems(1).Text Then
                            ''sitm.SubItems.selecteditem.index
                            APINVOICE1headerFields.FieldByName("TAXCLASS1").Value = "1"
                            APINVOICE1headerFields.FieldByName("AMTGROSTOT").Value = dInvAmt
                            dInvAmt = 0
                            APINVOICE1header.Insert()
                            APINVOICE1batch.Read()
                            Temp = APINVOICE1header.Exists
                            APINVOICE1header.RecordCreate(2)
                            APINVOICE1detail1.Cancel()


                        End If
                        rRow = rRow + 1

                    End If

                Catch ex As Exception
                    AccpacErrorHandler()
                End Try
                'APINVOICE3headerFields("ORDRNBR").PutWithoutVerification ("ORDNUM")
            Next

        Catch ex As Exception
            AccpacErrorHandler()
        End Try



    End Sub

    Function AddOptField(sOptFld As String, sOptValue As String)
        Try
            'TODO: need to fix the detail line to record that opt fields exist - currently says NO
            APINVOICE1detail4.RecordClear()
            APINVOICE1detail4.RecordCreate(0)
            APINVOICE1detail4Fields.FieldByName("OPTFIELD").Value = (sOptFld)                   ' Optional Field
            APINVOICE1detail4Fields.FieldByName("SWSET").Value = "1"                          ' Value Set
            APINVOICE1detail4Fields.FieldByName("VALIFTEXT").Value = (sOptValue)                    ' Text Value
            APINVOICE1detail4.Insert()
            APINVOICE1detail4Fields.FieldByName("OPTFIELD").PutWithoutVerification(sOptFld)   ' Optional Field
            APINVOICE1detail4.Read()



            '            APINVOICE1detail4.RecordCreate(0)
            '           APINVOICE1detail4Fields.FieldByName("OPTFIELD").Value = (sOptFld)                   ' Optional Field
            '          APINVOICE1detail4Fields.FieldByName("SWSET").Value = "1"                          ' Value Set
            '         APINVOICE1detail4Fields.FieldByName("VALIFTEXT").Value = (sOptValue)                   ' Text Value
            '        APINVOICE1detail4.Insert()
            '       APINVOICE1detail1.Update()
        Catch ex As Exception
            MsgBox("optfld error")
        End Try



    End Function

End Module
