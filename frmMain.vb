Imports AccpacCOMAPI
Imports AccpacSessionManager
Imports AccpacSignonManager
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms

Imports Microsoft.Data.SqlClient

Imports SpreadsheetLight


Imports System.Xml
Imports System.IO

Imports System.Configuration



Public Class frmMain

    Inherits System.Windows.Forms.Form

    Public sqlServer As String
    Public sqlDB As String
    Public sqlUser As String
    Public sqlPswd As String
    Public smtpHost As String
    Public smtpPort As String
    Public smtpSSL As String
    Public smtpUser As String
    Public smtpPswd As String

    Public LogFile As String = System.Windows.Forms.Application.StartupPath & "APComm - " & Format(Now, "yyyyMMdd") & ".txt"
    Public strLogLine As String = ""

    Public Shared iItems As Integer = 0
    Public Shared EffDate As String
    Public bDebug As Boolean = False

    '20220309.1930    ADS      09/03/2022       Initial Release Of APCommissions
    '20220311.1430    ADS      11/03/2022       LoadGrid verify and formatting
    '                                           Tidy code and objects
    '                                           Build Sage AP Inv Batch Creation OBject
    '20220328.1545    ADS       28/03/2022      Adjust for correct AP INV entry
    '20220404.1045    ADS       04/04/2022      create OptFields for AP Details  
    '20220406.0815    ADS       06/04/2022      Adjust OptFlds compose to correctly add details
    '20220509.1045    ADS       09/05/2022      Tidy and streamline batch creation process
    '20220622.1545    ADS       22/06/2022      elease for LIVE
    '20240314.1045    ADS       14/03/2024      Adjust BatchCreate for SRCAPPL=CM - accomodate APWorkflow

    Public BuildVersion As String = "20240314.1045"

    'TODO:  Add Record counters to show # entries, and total $ (hence balance to imports)


    Sub LoadGrid()

        '

        If ListView1.Items.Count > 0 Then
            ListView1.Items.Clear()
        End If

        Me.Label5.Text = ""



        Dim vAPComms As SqlCommand
        Dim sqlAPComms As String
        Dim intCount As Decimal = 1
        Dim lRow As Integer = 0
        Dim iErr As Integer = 0
        Dim lView(21) As String
        Dim sItm As ListViewItem

        Dim Sdte As String = Me.dteExpDate.Value.ToString("yyyyMMdd")
        EffDate = Sdte

        Dim A4W As New SqlConnection()
        Dim SQLConStr As String = "Server=" & sqlServer & ";Database=" & sqlDB & ";User ID=" & sqlUser & ";Password=" & sqlPswd
        A4W = New SqlConnection(SQLConStr)
        A4W.ConnectionString = SQLConStr


        '
        'TODO: add logic to check if view exists else fail   IF EXISTS(select * FROM sys.views where name = '')
        '

        A4W = New SqlConnection(SQLConStr)

        Try
            A4W.Open()
            ' ================================
            strLogLine = System.DateTime.Now & " - " & sSageOrgID & " - " & sSageUserID & "  -  " & "Connected to SQL"
            My.Computer.FileSystem.WriteAllText(LogFile, strLogLine & vbCrLf, True)
            ' ================================

            sqlAPComms = "Select * from v_APComm order by SPCode, OrderNo, InvDate, DUMMY3 "
            vAPComms = New SqlCommand(sqlAPComms, A4W)

            Dim Sql_APComms As SqlDataReader = vAPComms.ExecuteReader()
            While Sql_APComms.Read()

                lView(0) = intCount
                lView(1) = Trim(Sql_APComms.Item(0))    ' SPCode
                lView(2) = Trim(Sql_APComms.Item(1))    ' SPName
                lView(3) = Trim(Sql_APComms.Item(2))    ' IDCust
                lView(4) = Trim(Sql_APComms.Item(3))    ' NameCust
                lView(5) = Trim(Sql_APComms.Item(5))    ' OrdDate
                lView(6) = Trim(Sql_APComms.Item(6))    ' OrdNumber
                lView(7) = Trim(Sql_APComms.Item(9))    ' FmtItemNo
                lView(8) = Trim(Sql_APComms.Item(10))   ' InvNumber
                lView(9) = Trim(Sql_APComms.Item(12))   ' InvDate
                lView(10) = Format(Sql_APComms.Item(13), "##,##0.00")  ' Sales  Format(5459.4, "##,##0.00")
                lView(11) = Format(Sql_APComms.Item(14), "##,##0.00")  ' Cost
                lView(12) = Format(Sql_APComms.Item(15), "##,##0.00")  ' Margin
                lView(13) = Format(Sql_APComms.Item(16), "##,##0.00")  ' InvcBase
                lView(14) = Trim(Sql_APComms.Item(17))  ' LeadType
                lView(15) = Trim(Sql_APComms.Item(18))  ' LeadSource
                lView(16) = Format(Sql_APComms.Item(21), "##,##0.00")  ' TotRecComm
                lView(17) = Trim(Sql_APComms.Item(22))  ' SP%
                lView(18) = Format(Sql_APComms.Item(23) * Sql_APComms.Item(21), "##,##0.00")  ' SCommValue
                lView(19) = Trim(Sql_APComms.Item(24))  ' SPGrpID
                lView(20) = Format(Sql_APComms.Item(25), "##,##0.00")  ' Comm
                lView(21) = Format(Sql_APComms.Item(25), "##,##0.00")  ' MComm

                sItm = New ListViewItem(lView)
                ListView1.Items.Add(sItm)
                ListView1.Items(lRow).UseItemStyleForSubItems = False
                ListView1.Items(lRow).Checked = True

                intCount += 1
                lRow += 1
                ListView1.Refresh()

            End While
            Me.Label5.Text = lRow
            iItems = lRow

            Sql_APComms.Close()
            vAPComms.Dispose()
            A4W.Close()

            ' ================================
            strLogLine = System.DateTime.Now & " - " & sSageOrgID & " - " & sSageUserID & "  -  " & "Data Grid Refreshed"
            My.Computer.FileSystem.WriteAllText(LogFile, strLogLine & vbCrLf, True)
            strLogLine = System.DateTime.Now & " - " & sSageOrgID & " - " & sSageUserID & "  -  " & "Rows:- " & lRow
            My.Computer.FileSystem.WriteAllText(LogFile, strLogLine & vbCrLf, True)
            ' ================================

        Catch ex As Exception
            ' ================================
            strLogLine = System.DateTime.Now & " - " & sSageOrgID & " - " & sSageUserID & "  -  " & "ERROR - Failed to open SQL connection"
            My.Computer.FileSystem.WriteAllText(LogFile, strLogLine & vbCrLf, True)
            ' ================================
            MsgBox("SQL data - Cannot open connection")
        End Try

    End Sub


    Private Sub tnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        ' ================================
        strLogLine = System.DateTime.Now & " - " & sSageOrgID & " - " & sSageUserID & "  -  " & "Application Closed normally"
        My.Computer.FileSystem.WriteAllText(LogFile, strLogLine & vbCrLf, True)
        strLogLine = System.DateTime.Now & " - " & sSageOrgID & " - " & sSageUserID & "  -------------------------------------------------------------------------  "
        My.Computer.FileSystem.WriteAllText(LogFile, strLogLine & vbCrLf, True)
        ' ================================

        System.Windows.Forms.Application.Exit()

    End Sub

    Private Sub ReadAppConfig()

        Dim xmldoc As New XmlDocument()
        Dim SQLnode As XmlNodeList
        Dim SMTPnode As XmlNodeList
        Dim i As Integer
        Dim fs As New FileStream(System.Windows.Forms.Application.StartupPath & "\APComm.xml", FileMode.Open, FileAccess.Read)
        xmldoc.Load(fs)
        SQLnode = xmlDoc.GetElementsByTagName("SQLConfig")
        For i = 0 To SQLnode.Count - 1
            SQLnode(i).ChildNodes.Item(0).InnerText.Trim()
            sqlServer = SQLnode(i).ChildNodes.Item(0).InnerText.Trim()
            sqlDB = SQLnode(i).ChildNodes.Item(1).InnerText.Trim()
            sqlUser = SQLnode(i).ChildNodes.Item(2).InnerText.Trim()
            sqlPswd = SQLnode(i).ChildNodes.Item(3).InnerText.Trim()
        Next

        SMTPnode = xmlDoc.GetElementsByTagName("SMTPConfig")
        For i = 0 To SMTPnode.Count - 1
            SMTPnode(i).ChildNodes.Item(0).InnerText.Trim()
            smtpPort = SMTPnode(i).ChildNodes.Item(0).InnerText.Trim()
            smtpSSL = SMTPnode(i).ChildNodes.Item(1).InnerText.Trim()
            smtpHost = SMTPnode(i).ChildNodes.Item(2).InnerText.Trim()
            smtpUser = SMTPnode(i).ChildNodes.Item(3).InnerText.Trim()
            smtpPswd = SMTPnode(i).ChildNodes.Item(4).InnerText.Trim()
        Next

        fs.Dispose()
    End Sub





    Private Sub LoadListview()

        ListView1.View = View.Details
        ListView1.CheckBoxes = True
        ListView1.GridLines = True
        ListView1.FullRowSelect = True

        With ListView1
            .Columns.Add("#", 45, HorizontalAlignment.Center)
            .Columns.Add("SPCode", 60, HorizontalAlignment.Left)
            .Columns.Add("SPName", 110, HorizontalAlignment.Left)
            .Columns.Add("IDCUST", 75, HorizontalAlignment.Left)
            .Columns.Add("NameCust", 260, HorizontalAlignment.Left)
            .Columns.Add("OrdDate", 75, HorizontalAlignment.Left)
            .Columns.Add("OrdNumber", 75, HorizontalAlignment.Left)
            .Columns.Add("FmtItemNo", 120, HorizontalAlignment.Left)
            .Columns.Add("InvNumber", 80, HorizontalAlignment.Left)
            .Columns.Add("InvDate", 80, HorizontalAlignment.Left)
            .Columns.Add("Sales", 80, HorizontalAlignment.Right)
            .Columns.Add("Cost", 80, HorizontalAlignment.Right)
            .Columns.Add("Margin", 80, HorizontalAlignment.Right)
            .Columns.Add("InvcBase", 80, HorizontalAlignment.Right)
            .Columns.Add("LeadType", 65, HorizontalAlignment.Center)
            .Columns.Add("LeadSource", 65, HorizontalAlignment.Center)
            .Columns.Add("TotRecComm", 80, HorizontalAlignment.Right)
            .Columns.Add("SP%", 40, HorizontalAlignment.Center)
            .Columns.Add("SComValue", 80, HorizontalAlignment.Right)
            .Columns.Add("FieldTrainer", 65, HorizontalAlignment.Right)
            .Columns.Add("Comm", 80, HorizontalAlignment.Right)
            .Columns.Add("MComm", 80, HorizontalAlignment.Right)
            .Columns.Add("FT", 80, HorizontalAlignment.Right)


        End With

        If ListView1.Items.Count > 0 Then
            ListView1.Items.Clear()
        End If


        Me.Label5.Text = ""
        'Me.Label6.Text = ""


        '
        ' TODO: Move sqlARSAP to Form Load
        ' TODO: Add a condition
        '
        Dim SQLConStr As String = "Server=" & sqlServer & ";Database=" & sqlDB & ";User ID=" & sqlUser & ";Password=" & sqlPswd
        Dim A4W As New SqlConnection()
        A4W = New SqlConnection(SQLConStr)
        Try
            A4W.Open()

            Dim vARSAP As SqlCommand
            Dim sqlARSAP As String

            sqlARSAP = " select rtrim(CODESLSP) + '  -  ' + rtrim(NAMEEMPL) from arsap where SWACTV=1 "
            vARSAP = New SqlCommand(sqlARSAP, A4W)
            Dim Sql_ARSAP As SqlDataReader = vARSAP.ExecuteReader()
            While Sql_ARSAP.Read()
                cboSalesPerson.Items.Add(Sql_ARSAP.Item(0))
            End While

            Sql_ARSAP.Close()
            vARSAP.Dispose()
            A4W.Close()

        Catch ex As Exception

        End Try




    End Sub

    Private Sub ConnectToSage()

        If SageConnect.SageSession() = True Then
            Me.ToolStripStatusLabel1.Text = sSageOrgID
            Me.ToolStripStatusLabel2.Text = sSageCompName
            Me.ToolStripStatusLabel3.Text = sSageUserID
            Me.ToolStripStatusLabel4.Text = sSageSessDate
            Me.Text = sSageOrgID & " - AP Commissions Processing            [ " & BuildVersion & " ]"

            If SageConnect.DeclareViews() = True Then

            End If
        Else
            MsgBox("Sage Connection failed")
            Me.Text = "[ ## Sage Connection Failed ## ]"
        End If




    End Sub
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Me.txtTestMobile.Visible = False
        'Me.Label4.Visible = False

        ConnectToSage()


        strLogLine = System.DateTime.Now & " - " & sSageOrgID & " - " & sSageUserID & "  -  " & "Application Opened - Success"
        My.Computer.FileSystem.WriteAllText(LogFile, strLogLine & vbCrLf, True)

        strLogLine = System.DateTime.Now & " - " & sSageOrgID & " - " & sSageUserID & "  -  " & "Sage Connected - Success"
        My.Computer.FileSystem.WriteAllText(LogFile, strLogLine & vbCrLf, True)

        ReadAppConfig()
        strLogLine = System.DateTime.Now & " - " & sSageOrgID & " - " & sSageUserID & "  -  " & "Config Read - Success"
        My.Computer.FileSystem.WriteAllText(LogFile, strLogLine & vbCrLf, True)

        LoadListview()

    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click

        LoadGrid()

    End Sub



    Private Sub btnProcess_Click(sender As Object, e As EventArgs) Handles btnProcess.Click
        APInvoice()

        '  ProcessNotifications()

    End Sub



    Sub CreateSPInvoices()

        '
        '  Create New AP Inv Batch
        '   Group daat by SP and sum Commission Due
        '   Add entry to btch for SP, and add lines for comms/charges
        '








    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub
End Class
