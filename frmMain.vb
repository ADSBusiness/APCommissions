﻿Imports AccpacCOMAPI
Imports AccpacSessionManager
Imports AccpacSignonManager
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms

Imports Microsoft.Data.SqlClient

Imports System.Xml
Imports System.IO

Imports System.Configuration



Public Class frmMain

    Inherits System.Windows.Forms.Form

    Public sqlServer As String
    Public sqlDB As String
    Public sqlUser As String
    Public sqlPswd As String
    Public BuildVersion As String = "20210713.1"
    Public smtpHost As String
    Public smtpPort As String
    Public smtpSSL As String
    Public smtpUser As String
    Public smtpPswd As String

    Public strLogLine As String = ""

    Public bDebug As Boolean = False


    Sub LoadGrid()



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
        Dim lView(11) As String
        Dim sItm As ListViewItem

        Dim Sdte As String = Me.dteExpDate.Value.ToString("yyyyMMdd")

        sSQl = "select * from v_SMSAlert where EXPDATE='" & Sdte & "' order by expdate"
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
                lView(1) = Trim(SqlReader.Item(1))
                lView(2) = Trim(SqlReader.Item(3))
                lView(3) = Trim(SqlReader.Item(4))
                lView(4) = Trim(SqlReader.Item(5))
                lView(5) = Trim(SqlReader.Item(7))
                lView(6) = Trim(SqlReader.Item(8))
                lView(7) = Trim(SqlReader.Item(11))
                lView(8) = Trim(SqlReader.Item(13))
                lView(9) = Trim(SqlReader.Item(14))
                lView(10) = Trim(SqlReader.Item(0))
                lView(11) = Trim(SqlReader.Item(15))

                sItm = New ListViewItem(lView)
                ListView1.Items.Add(sItm)
                ListView1.Items(lRow).UseItemStyleForSubItems = False
                If Trim(SqlReader.Item(13)) <> "" And Microsoft.VisualBasic.Left(CStr(Trim(SqlReader.Item(13))), 2) = "04" Then
                    sItm.Checked = True
                Else
                    sItm.Checked = False
                    ListView1.Items(lRow).SubItems(8).BackColor = Color.LightPink
                    iErr += 1
                End If
                If Trim(SqlReader.Item(15)) = "" And Trim(SqlReader.Item(13)) <> "" And Microsoft.VisualBasic.Left(CStr(Trim(SqlReader.Item(13))), 2) = "04" Then
                    sItm.Checked = True
                Else
                    sItm.Checked = False
                    ListView1.Items(lRow).SubItems(11).BackColor = Color.LightYellow
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


    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        ' ================================
        strLogLine = System.DateTime.Now & " - " & sSageOrgID & " - " & sSageUserID & "  -  " & "Application Closed normally"
        My.Computer.FileSystem.WriteAllText(LogFile, strLogLine & vbCrLf, True)
        ' ================================

        Application.Exit()

    End Sub

    Private Sub ReadAppConfig()

        Dim xmldoc As New XmlDocument()
        Dim SQLnode As XmlNodeList
        Dim SMTPnode As XmlNodeList
        Dim i As Integer
        Dim fs As New FileStream(Application.StartupPath & "\EmailAlert.xml", FileMode.Open, FileAccess.Read)
        xmlDoc.Load(fs)
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
            .Columns.Add("Order #", 75, HorizontalAlignment.Left)
            .Columns.Add("Cust #", 75, HorizontalAlignment.Left)
            .Columns.Add("Name", 250, HorizontalAlignment.Left)
            .Columns.Add("City", 110, HorizontalAlignment.Left)
            .Columns.Add("Rep", 60, HorizontalAlignment.Left)
            .Columns.Add("Name", 125, HorizontalAlignment.Left)
            .Columns.Add("Location", 60, HorizontalAlignment.Left)
            .Columns.Add("RepMobile", 80, HorizontalAlignment.Left)
            .Columns.Add("RepEmail", 80, HorizontalAlignment.Left)
            .Columns.Add("ORDUNIQ", 0, HorizontalAlignment.Left)
            .Columns.Add("Status", 190, HorizontalAlignment.Left)
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


            Me.Text = sSageOrgID & " - Order Notifications"
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

        Dim eBody As String

        For Each sitm As ListViewItem In Me.ListView1.Items
            Try

                If sitm.Checked = True Then

                    eBody = "<br>Hi " & sitm.SubItems.Item(6).Text & "<br>"
                    eBody += "<br>Your order " & sitm.SubItems.Item(1).Text & "<br>"
                    eBody += "for " & sitm.SubItems.Item(3).Text & " of " & sitm.SubItems.Item(4).Text & "<br>"
                    eBody += "will be shipped on " & Me.dteExpDate.Text & "<br>"

                    If SendEmail(sitm.SubItems.Item(8).Text, sitm.SubItems.Item(1).Text, eBody, "from", "mob") = True Then


                        Call UpdateOENotified(System.DateTime.Now, sSageUserID, sitm.SubItems.Item(10).Text)
                        sitm.SubItems.Item(11).Text = "SMS sent " & System.DateTime.Now & " - " & sSageUserID

                        sitm.Checked = False
                        sitm.SubItems.Item(0).BackColor = Color.LightGreen

                        ' ================================
                        If Me.chkRunTest.Checked Then
                            strLogLine = System.DateTime.Now & " - " & sSageOrgID & " - " & sSageUserID & "  -  " & "TEST - Notification sent - " & Me.txtTestMobile.Text & " / " & sitm.SubItems.Item(1).Text
                            My.Computer.FileSystem.WriteAllText(LogFile, strLogLine & vbCrLf, True)
                        Else
                            strLogLine = System.DateTime.Now & " - " & sSageOrgID & " - " & sSageUserID & "  -  " & "Notification sent - " & sitm.SubItems.Item(8).Text & " / " & sitm.SubItems.Item(1).Text
                            My.Computer.FileSystem.WriteAllText(LogFile, strLogLine & vbCrLf, True)
                            ' ================================
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

            '            sSQLCommand.Dispose()
            A4W.Close()
        Catch ex As Exception
            MsgBox("SQL data - Cannot open connection")
        End Try



    End Function




End Class