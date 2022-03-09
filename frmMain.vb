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


    Public bDebug As Boolean = False

    '20220309.1930    ADS      28/07/2021      Initial Release Of APCommissions
    '20220309.1930    ADS      28/07/2021      Initial Release Of APCommissions
    '                                          and here
    '                                          and here again
    Public BuildVersion As String = "20220309.1930"



    Sub LoadGrid()

        ' Connect to SQL
        ' Prompt for paramters
        ' Load grid with sales data





        If ListView1.Items.Count > 0 Then
            ListView1.Items.Clear()
        End If

        Me.Label5.Text = ""
        Me.Label6.Text = ""


        Dim SQLConStr As String = "Server=" & sqlServer & ";Database=" & sqlDB & ";User ID=" & sqlUser & ";Password=" & sqlPswd


        Dim A4W As New SqlConnection()
        A4W.ConnectionString = SQLConStr
        '  A4W.Open()

        Dim sSQLCommand As SqlCommand
        Dim sSQl As String
        Dim intCount As Decimal = 1
        Dim lRow As Integer = 0
        Dim iErr As Integer = 0
        Dim lView(16) As String
        Dim sItm As ListViewItem

        Dim Sdte As String = Me.dteExpDate.Value.ToString("yyyyMMdd")


        If Me.chkShowAllOrders.Checked = True And Me.chkClosedOrders.Checked = True Then
            sSQl = "select * from v_SMSAlert where EXPDATE>='" & Sdte & "'  order by EXPDATE"
        Else
            If Me.chkShowAllOrders.Checked = True And Me.chkClosedOrders.Checked = False Then
                sSQl = "select * from v_SMSAlert where EXPDATE>='" & Sdte & "' and COMPLETE <4 order by EXPDATE"
            Else
                sSQl = "select * from v_SMSAlert where EXPDATE='" & Sdte & "' and COMPLETE <4 order by EXPDATE"

            End If
        End If


        ' ================================
        strLogLine = System.DateTime.Now & " - " & sSageOrgID & " - " & sSageUserID & "  -  " & sSQl
        My.Computer.FileSystem.WriteAllText(LogFile, strLogLine & vbCrLf, True)
        ' ================================
        A4W = New SqlConnection(SQLConStr)
        Try
            A4W.Open()
            ' ================================
            strLogLine = System.DateTime.Now & " - " & sSageOrgID & " - " & sSageUserID & "  -  " & "Connected to SQL"
            My.Computer.FileSystem.WriteAllText(LogFile, strLogLine & vbCrLf, True)
            ' ================================

            sSQLCommand = New SqlCommand(sSQl, A4W)
            Dim SqlReader As SqlDataReader = sSQLCommand.ExecuteReader()
            While SqlReader.Read()

                lView(0) = intCount
                lView(1) = Trim(SqlReader.Item(1))  ' Order #
                lView(2) = Trim(SqlReader.Item(3))  ' Cust #
                lView(3) = Trim(SqlReader.Item(4))  ' Name
                lView(4) = Trim(SqlReader.Item(5))  ' City
                lView(5) = Trim(SqlReader.Item(20))  ' Location
                lView(6) = Trim(SqlReader.Item(7))  ' Rep
                lView(7) = Trim(SqlReader.Item(9))  ' Name
                lView(8) = Trim(SqlReader.Item(15))  ' RepMobile
                lView(9) = Trim(SqlReader.Item(16))  ' RepEmail
                lView(10) = Trim(SqlReader.Item(8))  ' Rep 2
                lView(11) = Trim(SqlReader.Item(10))  ' Name 2
                lView(12) = Trim(SqlReader.Item(17))  ' RepMobile 2
                lView(13) = Trim(SqlReader.Item(18))  ' RepEmail 2
                lView(14) = Trim(SqlReader.Item(0))  ' OrdUniq 2
                lView(15) = Trim(SqlReader.Item(19))  ' Status 2

                sItm = New ListViewItem(lView)
                ListView1.Items.Add(sItm)
                ListView1.Items(lRow).UseItemStyleForSubItems = False


                'If MOB <> "" And MOB


                If VerifyMob(Trim(SqlReader.Item(15))) = False Then
                    sItm.Checked = False
                    ListView1.Items(lRow).SubItems(8).BackColor = Color.LightPink
                Else
                    sItm.Checked = True
                End If
                If VerifyMob(Trim(SqlReader.Item(17))) = False Then
                    If VerifyMob(Trim(SqlReader.Item(15))) = True Then
                        sItm.Checked = False
                        ListView1.Items(lRow).SubItems(12).BackColor = Color.LightPink
                    Else
                        sItm.Checked = True
                    End If
                Else
                    sItm.Checked = True
                End If

                ' ------------------------------------------------------------------------------
                If Trim(SqlReader.Item(19)) = "" Then
                    sItm.Checked = True
                Else
                    sItm.Checked = False
                    ListView1.Items(lRow).SubItems(15).BackColor = Color.LightYellow
                End If

                If Trim(SqlReader.Item(2)) < 4 Then
                    sItm.Checked = True
                Else
                    sItm.Checked = False
                    ListView1.Items(lRow).SubItems(1).BackColor = Color.LightPink
                End If




                intCount += 1
                lRow += 1
                ListView1.Refresh()

            End While
            Me.Label5.Text = lRow
            Me.Label6.Text = iErr
            SqlReader.Close()
            sSQLCommand.Dispose()
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
            .Columns.Add("SPCode", 62, HorizontalAlignment.Left)
            .Columns.Add("SPName", 75, HorizontalAlignment.Left)
            .Columns.Add("IDCUST", 240, HorizontalAlignment.Left)
            .Columns.Add("NameCust", 110, HorizontalAlignment.Left)
            .Columns.Add("OrdDate", 75, HorizontalAlignment.Left)
            .Columns.Add("OrdNumber", 65, HorizontalAlignment.Left)
            .Columns.Add("FmtItemNo", 110, HorizontalAlignment.Left)
            .Columns.Add("InvNumber", 80, HorizontalAlignment.Left)
            .Columns.Add("InvDate", 0, HorizontalAlignment.Left)
            .Columns.Add("Sales", 65, HorizontalAlignment.Left)
            .Columns.Add("Cost", 110, HorizontalAlignment.Left)
            .Columns.Add("Margin", 80, HorizontalAlignment.Left)
            .Columns.Add("InvcBase", 0, HorizontalAlignment.Left)
            .Columns.Add("LeadType", 0, HorizontalAlignment.Left)
            .Columns.Add("LeadSource", 190, HorizontalAlignment.Left)
            .Columns.Add("TotRecComm", 190, HorizontalAlignment.Left)
            .Columns.Add("SP%", 190, HorizontalAlignment.Left)
            .Columns.Add("SComValue", 190, HorizontalAlignment.Left)
            .Columns.Add("SPGrpID", 190, HorizontalAlignment.Left)
            .Columns.Add("Comm", 190, HorizontalAlignment.Left)
            .Columns.Add("MComm", 190, HorizontalAlignment.Left)






        End With

        If ListView1.Items.Count > 0 Then
            ListView1.Items.Clear()
        End If

        Me.Label5.Text = ""
        Me.Label6.Text = ""

    End Sub

    Private Sub ConnectToSage()

        If SageSession() = True Then
            Me.ToolStripStatusLabel1.Text = sSageOrgID
            Me.ToolStripStatusLabel2.Text = sSageCompName
            Me.ToolStripStatusLabel3.Text = sSageUserID
            Me.ToolStripStatusLabel4.Text = sSageSessDate
            Me.Text = sSageOrgID & " - AP Commissions Processing            [ " & BuildVersion & " ]"
        Else
            MsgBox("Sage Connection failed")
            Me.Text = "[ ## Sage Connection Failed ## ]"
        End If


    End Sub
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.txtTestMobile.Visible = False
        Me.Label4.Visible = False

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

    Sub ProcessNotifications()

        Dim eBody1 As String
        Dim eBody2 As String

        For Each sitm As ListViewItem In Me.ListView1.Items
            Try

                If sitm.Checked = True Then

                    eBody1 = "<br>Hi " & sitm.SubItems.Item(7).Text & "<br>"
                    eBody1 += "<br>Your order " & sitm.SubItems.Item(1).Text & "<br>"
                    eBody1 += "for " & sitm.SubItems.Item(3).Text & " of " & sitm.SubItems.Item(4).Text & "<br>"
                    eBody1 += "will be shipped on " & Me.dteExpDate.Text & "<br>"

                    eBody2 = "<br>Hi " & sitm.SubItems.Item(11).Text & "<br>"
                    eBody2 += "<br>Your order " & sitm.SubItems.Item(1).Text & "<br>"
                    eBody2 += "for " & sitm.SubItems.Item(3).Text & " of " & sitm.SubItems.Item(4).Text & "<br>"
                    eBody2 += "will be shipped on " & Me.dteExpDate.Text & "<br>"


                    If SendEmail(sitm.SubItems.Item(8).Text, sitm.SubItems.Item(1).Text, eBody1, eBody2, "mob", sitm.SubItems.Item(8).Text, sitm.SubItems.Item(12).Text) = True Then

                        Call UpdateOENotified(System.DateTime.Now, sSageUserID, sitm.SubItems.Item(14).Text)
                        sitm.SubItems.Item(15).Text = "SMS sent " & System.DateTime.Now & " - " & sSageUserID

                        sitm.Checked = False
                        sitm.SubItems.Item(0).BackColor = Color.LightGreen

                        If Me.chkRunTest.Checked Then
                            strLogLine = System.DateTime.Now & " - " & sSageOrgID & " - " & sSageUserID & "  -  " & "TEST - Notification sent - " & Me.txtTestMobile.Text & " / " & sitm.SubItems.Item(1).Text
                            My.Computer.FileSystem.WriteAllText(LogFile, strLogLine & vbCrLf, True)
                        Else
                            strLogLine = System.DateTime.Now & " - " & sSageOrgID & " - " & sSageUserID & "  -  " & "Notification sent - " & sitm.SubItems.Item(6).Text & " / " & sitm.SubItems.Item(8).Text & " / " & sitm.SubItems.Item(1).Text
                            My.Computer.FileSystem.WriteAllText(LogFile, strLogLine & vbCrLf, True)
                            If sitm.SubItems.Item(10).Text <> "" Then
                                strLogLine = System.DateTime.Now & " - " & sSageOrgID & " - " & sSageUserID & "  -  " & "Notification sent - " & sitm.SubItems.Item(10).Text & " / " & sitm.SubItems.Item(12).Text & " / " & sitm.SubItems.Item(1).Text
                                My.Computer.FileSystem.WriteAllText(LogFile, strLogLine & vbCrLf, True)
                            End If
                        End If

                    End If

                End If

            Catch ex As Exception
                If Me.chkRunTest.Checked Then
                    strLogLine = System.DateTime.Now & " - " & sSageOrgID & " - " & sSageUserID & "  -  " & "TEST - Notification FAILED - " & Me.txtTestMobile.Text & " / " & sitm.SubItems.Item(1).Text
                    My.Computer.FileSystem.WriteAllText(LogFile, strLogLine & vbCrLf, True)
                Else
                    strLogLine = System.DateTime.Now & " - " & sSageOrgID & " - " & sSageUserID & "  -  " & "Notification FAILED - " & sitm.SubItems.Item(8).Text & " / " & sitm.SubItems.Item(1).Text
                    My.Computer.FileSystem.WriteAllText(LogFile, strLogLine & vbCrLf, True)
                End If
            End Try

        Next

    End Sub

    Private Sub btnProcess_Click(sender As Object, e As EventArgs) Handles btnProcess.Click

        ProcessNotifications()

    End Sub

    Private Sub chkRunTest_CheckedChanged(sender As Object, e As EventArgs) Handles chkRunTest.CheckedChanged

        If Me.chkRunTest.Checked = True Then
            Me.txtTestMobile.Visible = True
            Me.Label4.Visible = True
        Else
            Me.txtTestMobile.Visible = False
            Me.Label4.Visible = False
        End If

    End Sub

    Function UpdateOENotified(sDateTime As String, sUser As String, sOrdUniq As String) As Boolean

        UpdateOENotified = False
        Dim SQLConStr As String = "Server=" & sqlServer & ";Database=" & sqlDB & ";User ID=" & sqlUser & ";Password=" & sqlPswd
        Dim A4W As New SqlConnection()
        A4W.ConnectionString = SQLConStr

        Dim sSQLCommand As SqlCommand
        Dim sSQl As String
        Dim intCount As Decimal = 1
        Dim lRow As Integer = 0
        Dim iErr As Integer = 0
        Dim lView(10) As String

        sSQl = "update oeordh set fob= 'SMS sent " & sDateTime & " - " & sUser & "' where ORDUNIQ='" & sOrdUniq & "'"
        A4W = New SqlConnection(SQLConStr)
        Try
            A4W.Open()
            sSQLCommand = New SqlCommand(sSQl, A4W)
            sSQLCommand.ExecuteNonQuery()
            UpdateOENotified = True
            ' ================================
            '            strLogLine = System.DateTime.Now & " - " & sSageOrgID & " - " & sSageUserID & "  -  " & "Order " & sOrdUniq & " FOB Updated "
            '           My.Computer.FileSystem.WriteAllText(LogFile, strLogLine & vbCrLf, True)

            '            sSQLCommand.Dispose()
            A4W.Close()
        Catch ex As Exception
            MsgBox("SQL data - Cannot update OrdUniq " & sOrdUniq)
            ' ================================
            strLogLine = System.DateTime.Now & " - " & sSageOrgID & " - " & sSageUserID & "  -  " & "SQL data - Cannot update OrdUniq FOB" & sOrdUniq
            My.Computer.FileSystem.WriteAllText(LogFile, strLogLine & vbCrLf, True)
        End Try



    End Function

    Function VerifyMob(sMob1 As String) As Integer

        Dim vMob As Boolean = False

        If Trim(sMob1) <> "" Then
            If (Microsoft.VisualBasic.Left(CStr(sMob1), 2) = "04") And (Len(sMob1.Replace(" ", String.Empty)) = 10) Then
                vMob = True
            Else
                vMob = False
            End If
        Else
            vMob = False
        End If

        VerifyMob = vMob
        Return VerifyMob

    End Function


    Public Sub ExportXLS()

        Try



            Dim SL As New SLDocument()








            For i = 0 To Me.ListView1.Columns.Count - 1
                SL.SetCellValue(1, i + 1, Me.ListView1.Columns(i).Text)

            Next
            For i = 0 To Me.ListView1.Items.Count - 1
                For j = 0 To Me.ListView1.Items(i).SubItems.Count - 1
                    SL.SetCellValue(i + 2, j + 1, Me.ListView1.Items(i).SubItems(j).Text)
                Next
            Next




            SL.SaveAs("e:\test.xlsx")
            'SL.

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try




    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        ExportXLS()
    End Sub

    Private Sub test1()
        Dim ssSQL As String

        ssSQL = <SQL>

                </SQL>
    End Sub



End Class
