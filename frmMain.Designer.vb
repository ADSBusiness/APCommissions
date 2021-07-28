<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnProcess = New System.Windows.Forms.Button()
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.btnRefresh = New System.Windows.Forms.Button()
        Me.dteExpDate = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel3 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripStatusLabel4 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.chkRunTest = New System.Windows.Forms.CheckBox()
        Me.txtTestMobile = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.chkShowAllOrders = New System.Windows.Forms.CheckBox()
        Me.chkClosedOrders = New System.Windows.Forms.CheckBox()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(857, 469)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(203, 30)
        Me.btnExit.TabIndex = 0
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnProcess
        '
        Me.btnProcess.Location = New System.Drawing.Point(140, 469)
        Me.btnProcess.Name = "btnProcess"
        Me.btnProcess.Size = New System.Drawing.Size(203, 30)
        Me.btnProcess.TabIndex = 1
        Me.btnProcess.Text = "Process Notifications"
        Me.btnProcess.UseVisualStyleBackColor = True
        '
        'ListView1
        '
        Me.ListView1.HideSelection = False
        Me.ListView1.Location = New System.Drawing.Point(12, 126)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(1261, 324)
        Me.ListView1.TabIndex = 7
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.List
        '
        'btnRefresh
        '
        Me.btnRefresh.Location = New System.Drawing.Point(474, 47)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(198, 43)
        Me.btnRefresh.TabIndex = 8
        Me.btnRefresh.Text = "Refresh Data"
        Me.btnRefresh.UseVisualStyleBackColor = True
        '
        'dteExpDate
        '
        Me.dteExpDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dteExpDate.Location = New System.Drawing.Point(191, 47)
        Me.dteExpDate.Name = "dteExpDate"
        Me.dteExpDate.Size = New System.Drawing.Size(107, 23)
        Me.dteExpDate.TabIndex = 9
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(58, 53)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(114, 15)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Expected Ship Date :"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(949, 69)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(41, 15)
        Me.Label6.TabIndex = 23
        Me.Label6.Text = "Label6"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(949, 47)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(41, 15)
        Me.Label5.TabIndex = 22
        Me.Label5.Text = "Label5"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.Label3.Location = New System.Drawing.Point(834, 69)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(105, 15)
        Me.Label3.TabIndex = 21
        Me.Label3.Text = "Missing Mobile # :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point)
        Me.Label2.Location = New System.Drawing.Point(834, 47)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 15)
        Me.Label2.TabIndex = 20
        Me.Label2.Text = "Total Records :"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.AutoSize = False
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1, Me.ToolStripStatusLabel2, Me.ToolStripStatusLabel3, Me.ToolStripStatusLabel4})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 532)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(1285, 22)
        Me.StatusStrip1.TabIndex = 24
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(119, 17)
        Me.ToolStripStatusLabel1.Text = "ToolStripStatusLabel1"
        '
        'ToolStripStatusLabel2
        '
        Me.ToolStripStatusLabel2.BorderSides = System.Windows.Forms.ToolStripStatusLabelBorderSides.Left
        Me.ToolStripStatusLabel2.Name = "ToolStripStatusLabel2"
        Me.ToolStripStatusLabel2.Size = New System.Drawing.Size(123, 17)
        Me.ToolStripStatusLabel2.Text = "ToolStripStatusLabel2"
        '
        'ToolStripStatusLabel3
        '
        Me.ToolStripStatusLabel3.BorderSides = System.Windows.Forms.ToolStripStatusLabelBorderSides.Left
        Me.ToolStripStatusLabel3.Name = "ToolStripStatusLabel3"
        Me.ToolStripStatusLabel3.Size = New System.Drawing.Size(123, 17)
        Me.ToolStripStatusLabel3.Text = "ToolStripStatusLabel3"
        '
        'ToolStripStatusLabel4
        '
        Me.ToolStripStatusLabel4.BorderSides = System.Windows.Forms.ToolStripStatusLabelBorderSides.Left
        Me.ToolStripStatusLabel4.Name = "ToolStripStatusLabel4"
        Me.ToolStripStatusLabel4.Size = New System.Drawing.Size(123, 17)
        Me.ToolStripStatusLabel4.Text = "ToolStripStatusLabel4"
        '
        'chkRunTest
        '
        Me.chkRunTest.AutoSize = True
        Me.chkRunTest.Location = New System.Drawing.Point(511, 473)
        Me.chkRunTest.Name = "chkRunTest"
        Me.chkRunTest.Size = New System.Drawing.Size(95, 19)
        Me.chkRunTest.TabIndex = 25
        Me.chkRunTest.Text = "System Test ?"
        Me.chkRunTest.UseVisualStyleBackColor = True
        '
        'txtTestMobile
        '
        Me.txtTestMobile.Location = New System.Drawing.Point(705, 469)
        Me.txtTestMobile.Name = "txtTestMobile"
        Me.txtTestMobile.Size = New System.Drawing.Size(121, 23)
        Me.txtTestMobile.TabIndex = 26
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(632, 474)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(67, 15)
        Me.Label4.TabIndex = 27
        Me.Label4.Text = "Test Mobile"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(857, 502)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(200, 15)
        Me.Label7.TabIndex = 28
        Me.Label7.Text = "©2021 ADS Business Services Pty Ltd"
        '
        'chkShowAllOrders
        '
        Me.chkShowAllOrders.AutoSize = True
        Me.chkShowAllOrders.Location = New System.Drawing.Point(191, 76)
        Me.chkShowAllOrders.Name = "chkShowAllOrders"
        Me.chkShowAllOrders.Size = New System.Drawing.Size(181, 19)
        Me.chkShowAllOrders.TabIndex = 29
        Me.chkShowAllOrders.Text = "Show all orders after this date"
        Me.chkShowAllOrders.UseVisualStyleBackColor = True
        '
        'chkClosedOrders
        '
        Me.chkClosedOrders.AutoSize = True
        Me.chkClosedOrders.Location = New System.Drawing.Point(191, 101)
        Me.chkClosedOrders.Name = "chkClosedOrders"
        Me.chkClosedOrders.Size = New System.Drawing.Size(138, 19)
        Me.chkClosedOrders.TabIndex = 30
        Me.chkClosedOrders.Text = "Include closed orders"
        Me.chkClosedOrders.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(1285, 554)
        Me.Controls.Add(Me.chkClosedOrders)
        Me.Controls.Add(Me.chkShowAllOrders)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtTestMobile)
        Me.Controls.Add(Me.chkRunTest)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dteExpDate)
        Me.Controls.Add(Me.btnRefresh)
        Me.Controls.Add(Me.ListView1)
        Me.Controls.Add(Me.btnProcess)
        Me.Controls.Add(Me.btnExit)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form1"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnExit As Button
    Friend WithEvents btnProcess As Button
    Friend WithEvents ListView1 As ListView
    Friend WithEvents btnRefresh As Button
    Friend WithEvents dteExpDate As DateTimePicker
    Friend WithEvents Label1 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents ToolStripStatusLabel1 As ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel2 As ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel3 As ToolStripStatusLabel
    Friend WithEvents ToolStripStatusLabel4 As ToolStripStatusLabel
    Friend WithEvents chkRunTest As CheckBox
    Friend WithEvents txtTestMobile As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents chkShowAllOrders As CheckBox
    Friend WithEvents chkClosedOrders As CheckBox
End Class
