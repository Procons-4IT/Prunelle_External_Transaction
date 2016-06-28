Imports System.IO
Imports System.Diagnostics.Process
Imports System.Xml
Imports System.Data
Imports System.Data.SqlClient
Imports System.Threading

Public Class frmLoginScreen
    Inherits System.Windows.Forms.Form

#Region "Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If

        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnReferesh As System.Windows.Forms.Button
    Friend WithEvents txtServerPwd As System.Windows.Forms.TextBox
    Friend WithEvents txtDBUserName As System.Windows.Forms.TextBox
    Friend WithEvents txtServerName As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnConnect As System.Windows.Forms.Button
    Friend WithEvents cmbCompanyDB As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtCompanyPWD As System.Windows.Forms.TextBox
    Friend WithEvents txtSBOUserName As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    '  Friend WithEvents Show_Click As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents fldFolderBrowse As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Public WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents cmbMainSAPCompany As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents MainSAPPWD As System.Windows.Forms.TextBox
    Friend WithEvents MainSAPUID As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Public WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents MainSQLPWD As System.Windows.Forms.TextBox
    Friend WithEvents MainSQLUID As System.Windows.Forms.TextBox
    Friend WithEvents MainSQLServer As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Public WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents LocalLicenseServer As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtDirectory As System.Windows.Forms.TextBox
    Friend WithEvents mainLicenseServer As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents cmbServertype As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents cmbMainServertype As System.Windows.Forms.ComboBox
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Frame1 = New System.Windows.Forms.GroupBox
        Me.btnReferesh = New System.Windows.Forms.Button
        Me.txtServerPwd = New System.Windows.Forms.TextBox
        Me.txtDBUserName = New System.Windows.Forms.TextBox
        Me.txtServerName = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.LocalLicenseServer = New System.Windows.Forms.TextBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtDirectory = New System.Windows.Forms.TextBox
        Me.Button5 = New System.Windows.Forms.Button
        Me.btnConnect = New System.Windows.Forms.Button
        Me.cmbCompanyDB = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtCompanyPWD = New System.Windows.Forms.TextBox
        Me.txtSBOUserName = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.fldFolderBrowse = New System.Windows.Forms.FolderBrowserDialog
        Me.Button1 = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.mainLicenseServer = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.cmbMainSAPCompany = New System.Windows.Forms.ComboBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.MainSAPPWD = New System.Windows.Forms.TextBox
        Me.MainSAPUID = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.Button3 = New System.Windows.Forms.Button
        Me.MainSQLPWD = New System.Windows.Forms.TextBox
        Me.MainSQLUID = New System.Windows.Forms.TextBox
        Me.MainSQLServer = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.cmbServertype = New System.Windows.Forms.ComboBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.cmbMainServertype = New System.Windows.Forms.ComboBox
        Me.Frame1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.Label16)
        Me.Frame1.Controls.Add(Me.cmbServertype)
        Me.Frame1.Controls.Add(Me.btnReferesh)
        Me.Frame1.Controls.Add(Me.txtServerPwd)
        Me.Frame1.Controls.Add(Me.txtDBUserName)
        Me.Frame1.Controls.Add(Me.txtServerName)
        Me.Frame1.Controls.Add(Me.Label3)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(12, 12)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(327, 158)
        Me.Frame1.TabIndex = 9
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Local Server Details"
        '
        'btnReferesh
        '
        Me.btnReferesh.Location = New System.Drawing.Point(121, 127)
        Me.btnReferesh.Name = "btnReferesh"
        Me.btnReferesh.Size = New System.Drawing.Size(163, 23)
        Me.btnReferesh.TabIndex = 4
        Me.btnReferesh.Text = "Refresh/Update Company List"
        '
        'txtServerPwd
        '
        Me.txtServerPwd.Location = New System.Drawing.Point(132, 72)
        Me.txtServerPwd.Name = "txtServerPwd"
        Me.txtServerPwd.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtServerPwd.Size = New System.Drawing.Size(152, 20)
        Me.txtServerPwd.TabIndex = 3
        '
        'txtDBUserName
        '
        Me.txtDBUserName.Location = New System.Drawing.Point(132, 48)
        Me.txtDBUserName.Name = "txtDBUserName"
        Me.txtDBUserName.Size = New System.Drawing.Size(152, 20)
        Me.txtDBUserName.TabIndex = 2
        '
        'txtServerName
        '
        Me.txtServerName.Location = New System.Drawing.Point(132, 24)
        Me.txtServerName.Name = "txtServerName"
        Me.txtServerName.Size = New System.Drawing.Size(152, 20)
        Me.txtServerName.TabIndex = 1
        Me.txtServerName.Text = "(local)"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 20)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = " User Name"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 20)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Server"
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox1.Controls.Add(Me.LocalLicenseServer)
        Me.GroupBox1.Controls.Add(Me.Label15)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.txtDirectory)
        Me.GroupBox1.Controls.Add(Me.Button5)
        Me.GroupBox1.Controls.Add(Me.btnConnect)
        Me.GroupBox1.Controls.Add(Me.cmbCompanyDB)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.txtCompanyPWD)
        Me.GroupBox1.Controls.Add(Me.txtSBOUserName)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.GroupBox1.Location = New System.Drawing.Point(8, 176)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.GroupBox1.Size = New System.Drawing.Size(331, 177)
        Me.GroupBox1.TabIndex = 10
        Me.GroupBox1.TabStop = False
        '
        'LocalLicenseServer
        '
        Me.LocalLicenseServer.Location = New System.Drawing.Point(136, 16)
        Me.LocalLicenseServer.Name = "LocalLicenseServer"
        Me.LocalLicenseServer.Size = New System.Drawing.Size(152, 20)
        Me.LocalLicenseServer.TabIndex = 19
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(16, 16)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(109, 20)
        Me.Label15.TabIndex = 20
        Me.Label15.Text = "SBO License Server"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(20, 119)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(109, 20)
        Me.Label7.TabIndex = 18
        Me.Label7.Text = "Log Directory"
        '
        'txtDirectory
        '
        Me.txtDirectory.Location = New System.Drawing.Point(132, 119)
        Me.txtDirectory.Name = "txtDirectory"
        Me.txtDirectory.Size = New System.Drawing.Size(152, 20)
        Me.txtDirectory.TabIndex = 17
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(285, 118)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(40, 23)
        Me.Button5.TabIndex = 16
        Me.Button5.Text = "..."
        '
        'btnConnect
        '
        Me.btnConnect.Location = New System.Drawing.Point(93, 148)
        Me.btnConnect.Name = "btnConnect"
        Me.btnConnect.Size = New System.Drawing.Size(152, 23)
        Me.btnConnect.TabIndex = 10
        Me.btnConnect.Text = "Test Local Connection"
        '
        'cmbCompanyDB
        '
        Me.cmbCompanyDB.Location = New System.Drawing.Point(132, 42)
        Me.cmbCompanyDB.Name = "cmbCompanyDB"
        Me.cmbCompanyDB.Size = New System.Drawing.Size(152, 22)
        Me.cmbCompanyDB.TabIndex = 5
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(16, 42)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(109, 20)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Company DB"
        '
        'txtCompanyPWD
        '
        Me.txtCompanyPWD.Location = New System.Drawing.Point(132, 92)
        Me.txtCompanyPWD.Name = "txtCompanyPWD"
        Me.txtCompanyPWD.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtCompanyPWD.Size = New System.Drawing.Size(152, 20)
        Me.txtCompanyPWD.TabIndex = 7
        '
        'txtSBOUserName
        '
        Me.txtSBOUserName.Location = New System.Drawing.Point(132, 68)
        Me.txtSBOUserName.Name = "txtSBOUserName"
        Me.txtSBOUserName.Size = New System.Drawing.Size(152, 20)
        Me.txtSBOUserName.TabIndex = 6
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(16, 92)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(109, 20)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "Password"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(16, 68)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(109, 20)
        Me.Label5.TabIndex = 14
        Me.Label5.Text = "SBO User Name"
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem2, Me.MenuItem1, Me.MenuItem3})
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 0
        Me.MenuItem2.Text = "Service Log File"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 1
        Me.MenuItem1.Text = "BP Log File"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 2
        Me.MenuItem3.Text = "Exit"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(110, 148)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(152, 23)
        Me.Button1.TabIndex = 12
        Me.Button1.Text = "Test Remote Connection"
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox2.Controls.Add(Me.mainLicenseServer)
        Me.GroupBox2.Controls.Add(Me.Label8)
        Me.GroupBox2.Controls.Add(Me.Button1)
        Me.GroupBox2.Controls.Add(Me.cmbMainSAPCompany)
        Me.GroupBox2.Controls.Add(Me.Label9)
        Me.GroupBox2.Controls.Add(Me.MainSAPPWD)
        Me.GroupBox2.Controls.Add(Me.MainSAPUID)
        Me.GroupBox2.Controls.Add(Me.Label10)
        Me.GroupBox2.Controls.Add(Me.Label11)
        Me.GroupBox2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.GroupBox2.Location = New System.Drawing.Point(358, 176)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.GroupBox2.Size = New System.Drawing.Size(331, 177)
        Me.GroupBox2.TabIndex = 13
        Me.GroupBox2.TabStop = False
        '
        'mainLicenseServer
        '
        Me.mainLicenseServer.Location = New System.Drawing.Point(132, 16)
        Me.mainLicenseServer.Name = "mainLicenseServer"
        Me.mainLicenseServer.Size = New System.Drawing.Size(152, 20)
        Me.mainLicenseServer.TabIndex = 15
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(6, 19)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(109, 20)
        Me.Label8.TabIndex = 16
        Me.Label8.Text = "SBO License Server"
        '
        'cmbMainSAPCompany
        '
        Me.cmbMainSAPCompany.Location = New System.Drawing.Point(132, 46)
        Me.cmbMainSAPCompany.Name = "cmbMainSAPCompany"
        Me.cmbMainSAPCompany.Size = New System.Drawing.Size(152, 22)
        Me.cmbMainSAPCompany.TabIndex = 5
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(16, 46)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(109, 20)
        Me.Label9.TabIndex = 11
        Me.Label9.Text = "Company DB"
        '
        'MainSAPPWD
        '
        Me.MainSAPPWD.Location = New System.Drawing.Point(132, 96)
        Me.MainSAPPWD.Name = "MainSAPPWD"
        Me.MainSAPPWD.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.MainSAPPWD.Size = New System.Drawing.Size(152, 20)
        Me.MainSAPPWD.TabIndex = 7
        '
        'MainSAPUID
        '
        Me.MainSAPUID.Location = New System.Drawing.Point(132, 72)
        Me.MainSAPUID.Name = "MainSAPUID"
        Me.MainSAPUID.Size = New System.Drawing.Size(152, 20)
        Me.MainSAPUID.TabIndex = 6
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(16, 96)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(109, 20)
        Me.Label10.TabIndex = 13
        Me.Label10.Text = "Password"
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(16, 72)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(109, 20)
        Me.Label11.TabIndex = 14
        Me.Label11.Text = "SBO User Name"
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.SystemColors.Control
        Me.GroupBox3.Controls.Add(Me.Label17)
        Me.GroupBox3.Controls.Add(Me.Button3)
        Me.GroupBox3.Controls.Add(Me.cmbMainServertype)
        Me.GroupBox3.Controls.Add(Me.MainSQLPWD)
        Me.GroupBox3.Controls.Add(Me.MainSQLUID)
        Me.GroupBox3.Controls.Add(Me.MainSQLServer)
        Me.GroupBox3.Controls.Add(Me.Label12)
        Me.GroupBox3.Controls.Add(Me.Label13)
        Me.GroupBox3.Controls.Add(Me.Label14)
        Me.GroupBox3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.GroupBox3.Location = New System.Drawing.Point(362, 12)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.GroupBox3.Size = New System.Drawing.Size(336, 158)
        Me.GroupBox3.TabIndex = 11
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = " Remote Server Details"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(128, 118)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(163, 23)
        Me.Button3.TabIndex = 4
        Me.Button3.Text = "Refresh/Update Company List"
        '
        'MainSQLPWD
        '
        Me.MainSQLPWD.Location = New System.Drawing.Point(132, 72)
        Me.MainSQLPWD.Name = "MainSQLPWD"
        Me.MainSQLPWD.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.MainSQLPWD.Size = New System.Drawing.Size(152, 20)
        Me.MainSQLPWD.TabIndex = 3
        '
        'MainSQLUID
        '
        Me.MainSQLUID.Location = New System.Drawing.Point(132, 48)
        Me.MainSQLUID.Name = "MainSQLUID"
        Me.MainSQLUID.Size = New System.Drawing.Size(152, 20)
        Me.MainSQLUID.TabIndex = 2
        '
        'MainSQLServer
        '
        Me.MainSQLServer.Location = New System.Drawing.Point(132, 24)
        Me.MainSQLServer.Name = "MainSQLServer"
        Me.MainSQLServer.Size = New System.Drawing.Size(152, 20)
        Me.MainSQLServer.TabIndex = 1
        Me.MainSQLServer.Text = "(local)"
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(16, 72)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(100, 20)
        Me.Label12.TabIndex = 4
        Me.Label12.Text = "Password"
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(16, 48)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(100, 20)
        Me.Label13.TabIndex = 5
        Me.Label13.Text = " User Name"
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(16, 24)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(100, 20)
        Me.Label14.TabIndex = 6
        Me.Label14.Text = "Server"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(12, 399)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(125, 23)
        Me.Button2.TabIndex = 14
        Me.Button2.Text = "Export Data"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(144, 399)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(125, 23)
        Me.Button4.TabIndex = 15
        Me.Button4.Text = "Close"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 72)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 20)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Password"
        '
        'cmbServertype
        '
        Me.cmbServertype.Items.AddRange(New Object() {"2005", "2008"})
        Me.cmbServertype.Location = New System.Drawing.Point(132, 98)
        Me.cmbServertype.Name = "cmbServertype"
        Me.cmbServertype.Size = New System.Drawing.Size(152, 22)
        Me.cmbServertype.TabIndex = 21
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(15, 98)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(100, 20)
        Me.Label16.TabIndex = 22
        Me.Label16.Text = "Server Type"
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(15, 95)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(100, 20)
        Me.Label17.TabIndex = 24
        Me.Label17.Text = "Server Type"
        '
        'cmbMainServertype
        '
        Me.cmbMainServertype.Items.AddRange(New Object() {"2005", "2008"})
        Me.cmbMainServertype.Location = New System.Drawing.Point(132, 95)
        Me.cmbMainServertype.Name = "cmbMainServertype"
        Me.cmbMainServertype.Size = New System.Drawing.Size(152, 22)
        Me.cmbMainServertype.TabIndex = 23
        '
        'frmLoginScreen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(757, 481)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Frame1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Name = "frmLoginScreen"
        Me.Text = "Choose Company Database"
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
#End Region

#Region "Declaration"
    Public oCompany As SAPbobsCOM.Company
    Dim objPC = New clsMain
    Dim sValue As String
    Dim sPath As String
#End Region

#Region "API Calls"

    '****************************************************************************
    'Type              :   Function     
    'Name              :   WritePrivateProfileString
    'Parameter         :   
    'Return Value      : 
    'Author            :   DEV-3
    'Created Date      :   
    'Last Modified By  : 
    'Modified Date     : 
    'Purpose           :   Standard API declarations for INI access
    '****************************************************************************

    Private Declare Unicode Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringW" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, ByVal lpString As String, _
    ByVal lpFileName As String) As Int32

    Private Declare Unicode Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Int32, _
    ByVal lpFileName As String) As Int32
#End Region

#Region "Function for Read INI File"

    '****************************************************************************
    'Type              :   Function    
    'Name              :   INIRead
    'Parameter         :   
    'Return Value      : 
    'Author            :   DEV-3
    'Created Date      :   
    'Last Modified By  : 
    'Modified Date     : 
    'Purpose           :   To Read INI File
    '****************************************************************************

    Public Overloads Function INIRead(ByVal INIPath As String, _
        ByVal SectionName As String, ByVal KeyName As String, _
        ByVal DefaultValue As String) As String
        Dim n As Int32
        Dim sData As String
        sData = Space$(1024)
        n = GetPrivateProfileString(SectionName, KeyName, DefaultValue, _
        sData, sData.Length, INIPath)
        If n > 0 Then
            INIRead = sData.Substring(0, n)
        Else
            INIRead = ""
        End If
    End Function
#End Region

#Region "To Read INI File"

    '****************************************************************************
    'Type              :   Function     
    'Name              :   INIRead
    'Parameter         :   
    'Return Value      : 
    'Author            :   DEV-3
    'Created Date      :   
    'Last Modified By  : 
    'Modified Date     : 
    'Purpose           :   To Read INI File
    '****************************************************************************

    Public Overloads Function INIRead(ByVal INIPath As String, _
    ByVal SectionName As String, ByVal KeyName As String) As String
        Return INIRead(INIPath, SectionName, KeyName, "")
    End Function

    Public Overloads Function INIRead(ByVal INIPath As String, _
    ByVal SectionName As String) As String
        ' overload 2 returns all keys in a given section of the given file 
        Return INIRead(INIPath, SectionName, Nothing, "")
    End Function

    Public Overloads Function INIRead(ByVal INIPath As String) As String
        ' overload 3 returns all section names given just path 
        Return INIRead(INIPath, Nothing, Nothing, "")
    End Function
#End Region

#Region "To write INI File"

    '****************************************************************************
    'Type              :   Procedure     
    'Name              :   INIWrite
    'Parameter         :   
    'Return Value      : 
    'Author            :   DEV-3
    'Created Date      :   
    'Last Modified By  : 
    'Modified Date     : 
    'Purpose           :   To write INI File 
    '****************************************************************************

    Public Sub INIWrite(ByVal INIPath As String, ByVal SectionName As String, _
       ByVal KeyName As String, ByVal TheValue As String)
        Call WritePrivateProfileString(SectionName, KeyName, TheValue, INIPath)
    End Sub

    Public Overloads Sub INIDelete(ByVal INIPath As String, ByVal SectionName As String, _
    ByVal KeyName As String) ' delete single line from section 
        Call WritePrivateProfileString(SectionName, KeyName, Nothing, INIPath)
    End Sub

    Public Overloads Sub INIDelete(ByVal INIPath As String, ByVal SectionName As String)
        ' delete section from INI file 
        Call WritePrivateProfileString(SectionName, Nothing, Nothing, INIPath)
    End Sub


#End Region

#Region "To Save Details in ini File"
    '****************************************************************************
    'Type              :   Procedure     
    'Name              :   WriteiniFile
    'Parameter         :   
    'Return Value      : 
    'Author            :   DEV-3
    'Created Date      :   
    'Last Modified By  : 
    'Modified Date     : 
    'Purpose           :   To Save Details in ini File

    '****************************************************************************
    Private Sub WriteiniFile()
        Try

      
            sPath = System.Windows.Forms.Application.StartupPath & "\ConnectionInfo.ini"
            INIWrite(sPath, "SqlServer", "Key1-1", txtServerName.Text) ' build INI file 
            INIWrite(sPath, "SQLUID", "Key1-2", txtDBUserName.Text)
            INIWrite(sPath, "SQLPWD", "Key1-3", txtServerPwd.Text)
            INIWrite(sPath, "ServerType", "Key1-S1", strLocalServertype)
            INIWrite(sPath, "SAPCompany", "Key1-4", strSAPCompany) ' cmbCompanyDB.SelectedText)
            INIWrite(sPath, "SAPUID", "Key1-5", txtSBOUserName.Text)
            INIWrite(sPath, "SAPPWD", "Key1-6", txtCompanyPWD.Text)
            INIWrite(sPath, "LogFile", "Key1-9", txtDirectory.Text)

            INIWrite(sPath, "MainSqlServer", "Key1-10", MainSQLServer.Text) ' build INI file 
            INIWrite(sPath, "MainSQLUID", "Key1-11", MainSQLUID.Text)
            INIWrite(sPath, "MainSQLPWD", "Key1-12", MainSQLPWD.Text)
            INIWrite(sPath, "MainServerType", "Key1-M1", strMainServertype)
            INIWrite(sPath, "MainSAPCompany", "Key1-13", strMainSAPCompany) ' cmbCompanyDB.SelectedText)
            INIWrite(sPath, "MainSAPUID", "Key1-14", MainSAPUID.Text)
            INIWrite(sPath, "MainSAPPWD", "Key1-15", MainSAPPWD.Text)
            INIWrite(sPath, "MainLicense", "Key1-16", mainLicenseServer.Text)
            INIWrite(sPath, "LocalLicense", "Key1-17", LocalLicenseServer.Text)

        Catch ex As Exception

        End Try
    End Sub

    Private Sub WriteiniFile_Folder(ByVal aFileName As String)
        Try
            sPath = aFileName 'System.Windows.Forms.Application.StartupPath & "\ConnectionInfo.ini"
            INIWrite(sPath, "SqlServer", "Key1-1", txtServerName.Text) ' build INI file 
            INIWrite(sPath, "SQLUID", "Key1-2", txtDBUserName.Text)
            INIWrite(sPath, "SQLPWD", "Key1-3", txtServerPwd.Text)
            INIWrite(sPath, "ServerType", "Key1-S1", strLocalServertype)
            INIWrite(sPath, "SAPCompany", "Key1-4", strSAPCompany) ' cmbCompanyDB.SelectedText)
            INIWrite(sPath, "SAPUID", "Key1-5", txtSBOUserName.Text)
            INIWrite(sPath, "SAPPWD", "Key1-6", txtCompanyPWD.Text)
            INIWrite(sPath, "LogFile", "Key1-9", txtDirectory.Text)

            INIWrite(sPath, "MainSqlServer", "Key1-10", MainSQLServer.Text) ' build INI file 
            INIWrite(sPath, "MainSQLUID", "Key1-11", MainSQLUID.Text)
            INIWrite(sPath, "MainSQLPWD", "Key1-12", MainSQLPWD.Text)
            INIWrite(sPath, "MainServerType", "Key1-M1", strMainServertype)
            INIWrite(sPath, "MainSAPCompany", "Key1-13", strMainSAPCompany) ' cmbCompanyDB.SelectedText)
            INIWrite(sPath, "MainSAPUID", "Key1-14", MainSAPUID.Text)
            INIWrite(sPath, "MainSAPPWD", "Key1-15", MainSAPPWD.Text)
            INIWrite(sPath, "MainLicense", "Key1-16", mainLicenseServer.Text)
            INIWrite(sPath, "LocalLicense", "Key1-17", LocalLicenseServer.Text)
        Catch ex As Exception

        End Try

    End Sub

#End Region

#Region "To Read ini File"

    '****************************************************************************
    'Type              :   Procedure     
    'Name              :   ReadiniFile
    'Parameter         :   
    'Return Value      : 
    'Author            :   DEV-3
    'Created Date      :   
    'Last Modified By  : 
    'Modified Date     : 
    'Purpose           :   To Read ini File
    '****************************************************************************

    Private Sub ReadiniFile()
        sPath = System.Windows.Forms.Application.StartupPath & "\ConnectionInfo.ini"
        sValue = INIRead(sPath, "SqlServer", "key1-1", "")
        txtServerName.Text = sValue
        strSQLServer = sValue
        sValue = INIRead(sPath, "SQLUID", "Key1-2", "")
        txtDBUserName.Text = sValue
        strSQLUID = sValue
        sValue = INIRead(sPath, "SQLPWD", "Key1-3", "")
        txtServerPwd.Text = sValue
        strSQLPWD = sValue

        sValue = INIRead(sPath, "ServerType", "Key1-S1", "")
        cmbServertype.Text = sValue
        strLocalServertype = sValue

        sValue = INIRead(sPath, "SAPCompany", "Key1-4", "")
        cmbCompanyDB.SelectedText = sValue
        strSAPCompany = sValue
        sValue = INIRead(sPath, "SAPUID", "Key1-5", "")
        txtSBOUserName.Text = sValue
        strSAPUID = sValue
        sValue = INIRead(sPath, "SAPPWD", "Key1-6", "")
        txtCompanyPWD.Text = sValue
        strSAPPWD = sValue
        strSAPPWD = sValue
        sValue = INIRead(sPath, "ExportDir", "Key1-7", "")
        '  txtDirectory.Text = sValue
        strExportFilePaty = sValue
        'sValue = INIRead(sPath, "ExportFile", "Key1-8", "")
        'strFileStart = sValue
        sValue = INIRead(sPath, "LogFile", "Key1-9", "")
        strErrorLogPath = sValue
        txtDirectory.Text = strErrorLogPath


        sValue = INIRead(sPath, "MainSqlServer", "key1-10", "")
        MainSQLServer.Text = sValue
        strMainSQLServer = sValue
        sValue = INIRead(sPath, "MainSQLUID", "Key1-11", "")
        MainSQLUID.Text = sValue
        strMainSQLUID = sValue
        sValue = INIRead(sPath, "MainSQLPWD", "Key1-12", "")
        MainSQLPWD.Text = sValue
        strMainSQLPWD = sValue
        sValue = INIRead(sPath, "MainServerType", "Key1-M1", "")
        cmbMainServertype.Text = sValue
        strMainServertype = sValue

        sValue = INIRead(sPath, "MainSAPCompany", "Key1-13", "")
        cmbMainSAPCompany.SelectedText = sValue
        strMainSAPCompany = sValue
        sValue = INIRead(sPath, "MainSAPUID", "Key1-14", "")
        MainSAPUID.Text = sValue
        strMainSAPUID = sValue
        sValue = INIRead(sPath, "MainSAPPWD", "Key1-15", "")
        MainSAPPWD.Text = sValue
        strMainSAPPWD = sValue
        strMainSAPPWD = sValue

        sValue = INIRead(sPath, "MainLicense", "Key1-16", "")
        mainLicenseServer.Text = sValue
        strMainLicenseServer = sValue

        sValue = INIRead(sPath, "LocalLicense", "Key1-17", "")
        LocalLicenseServer.Text = sValue
        strLocalLicensserver = sValue

        sValue = INIRead(sPath, "StartDate", "Key1-18", "")
        strStartDate = sValue

        sValue = INIRead(sPath, "EndDate", "Key1-19", "")
        strEndDate = sValue


    End Sub

    Private Sub ReadiniFile_File(ByVal aFileName As String)
        sPath = aFileName 'System.Windows.Forms.Application.StartupPath & "\ConnectionInfo.ini"
        sValue = INIRead(sPath, "SqlServer", "key1-1", "")
        txtServerName.Text = ""
        txtServerName.Text = sValue
        strSQLServer = sValue
        sValue = INIRead(sPath, "SQLUID", "Key1-2", "")
        txtDBUserName.Text = ""
        txtDBUserName.Text = sValue
        strSQLUID = sValue
        sValue = INIRead(sPath, "SQLPWD", "Key1-3", "")
        txtServerPwd.Text = ""
        txtServerPwd.Text = sValue
        strSQLPWD = sValue

        sValue = INIRead(sPath, "ServerType", "Key1-S1", "")
        cmbServertype.Text = ""
        cmbServertype.Text = sValue
        strLocalServertype = sValue

        sValue = INIRead(sPath, "SAPCompany", "Key1-4", "")
        cmbCompanyDB.SelectedText = ""
        cmbCompanyDB.SelectedText = sValue
        strSAPCompany = sValue
        sValue = INIRead(sPath, "SAPUID", "Key1-5", "")
        txtSBOUserName.Text = ""
        txtSBOUserName.Text = sValue
        strSAPUID = sValue
        sValue = INIRead(sPath, "SAPPWD", "Key1-6", "")
        txtCompanyPWD.Text = ""
        txtCompanyPWD.Text = sValue
        strSAPPWD = sValue
        strSAPPWD = sValue
        sValue = INIRead(sPath, "ExportDir", "Key1-7", "")
        '  txtDirectory.Text = sValue
        strExportFilePaty = sValue
        'sValue = INIRead(sPath, "ExportFile", "Key1-8", "")
        'strFileStart = sValue
        sValue = INIRead(sPath, "LogFile", "Key1-9", "")
        strErrorLogPath = sValue
        txtDirectory.Text = strErrorLogPath
        sValue = INIRead(sPath, "MainSqlServer", "key1-10", "")
        MainSQLServer.Text = sValue
        strMainSQLServer = sValue
        sValue = INIRead(sPath, "MainSQLUID", "Key1-11", "")
        MainSQLUID.Text = sValue
        strMainSQLUID = sValue
        sValue = INIRead(sPath, "MainSQLPWD", "Key1-12", "")
        MainSQLPWD.Text = sValue
        strMainSQLPWD = sValue
        sValue = INIRead(sPath, "MainServerType", "Key1-M1", "")
        cmbMainServertype.Text = ""
        cmbMainServertype.Text = sValue
        strMainServertype = sValue

        sValue = INIRead(sPath, "MainSAPCompany", "Key1-13", "")
        cmbMainSAPCompany.SelectedText = ""
        cmbMainSAPCompany.SelectedText = sValue
        strMainSAPCompany = sValue
        sValue = INIRead(sPath, "MainSAPUID", "Key1-14", "")
        MainSAPUID.Text = sValue
        strMainSAPUID = sValue
        sValue = INIRead(sPath, "MainSAPPWD", "Key1-15", "")
        MainSAPPWD.Text = sValue
        strMainSAPPWD = sValue
        strMainSAPPWD = sValue

        sValue = INIRead(sPath, "MainLicense", "Key1-16", "")
        mainLicenseServer.Text = sValue
        strMainLicenseServer = sValue

        sValue = INIRead(sPath, "LocalLicense", "Key1-17", "")
        LocalLicenseServer.Text = sValue
        strLocalLicensserver = sValue

        sValue = INIRead(sPath, "StartDate", "Key1-18", "")
        strStartDate = sValue

        sValue = INIRead(sPath, "EndDate", "Key1-19", "")
        strEndDate = sValue


    End Sub

    Private Sub ReadiniFile_FileCheck(ByVal aFileName As String)
        Dim sqlServer As String

        sPath = aFileName 'System.Windows.Forms.Application.StartupPath & "\ConnectionInfo.ini"
        sValue = INIRead(sPath, "SqlServer", "key1-1", "")
        strSQLServer = sValue

        sValue = INIRead(sPath, "SQLUID", "Key1-2", "")
        strSQLUID = sValue

        sValue = INIRead(sPath, "SQLPWD", "Key1-3", "")
        strSQLPWD = sValue

        sValue = INIRead(sPath, "ServerType", "Key1-S1", "")
        strLocalServertype = sValue

        sValue = INIRead(sPath, "SAPCompany", "Key1-4", "")
        strSAPCompany = sValue

        sValue = INIRead(sPath, "SAPUID", "Key1-5", "")
        strSAPUID = sValue

        sValue = INIRead(sPath, "SAPPWD", "Key1-6", "")
        strSAPPWD = sValue

        sValue = INIRead(sPath, "ExportDir", "Key1-7", "")
        strExportFilePaty = sValue
        'sValue = INIRead(sPath, "ExportFile", "Key1-8", "")
        'strFileStart = sValue
        sValue = INIRead(sPath, "LogFile", "Key1-9", "")
        strErrorLogPath = sValue

        sValue = INIRead(sPath, "MainSqlServer", "key1-10", "")
        strMainSQLServer = sValue

        sValue = INIRead(sPath, "MainSQLUID", "Key1-11", "")
        strMainSQLUID = sValue

        sValue = INIRead(sPath, "MainSQLPWD", "Key1-12", "")
        strMainSQLPWD = sValue

        sValue = INIRead(sPath, "MainServerType", "Key1-M1", "")

        strMainServertype = sValue

        sValue = INIRead(sPath, "MainSAPCompany", "Key1-13", "")

        strMainSAPCompany = sValue
        sValue = INIRead(sPath, "MainSAPUID", "Key1-14", "")

        strMainSAPUID = sValue
        sValue = INIRead(sPath, "MainSAPPWD", "Key1-15", "")

        strMainSAPPWD = sValue
        strMainSAPPWD = sValue

        sValue = INIRead(sPath, "MainLicense", "Key1-16", "")
        strMainLicenseServer = sValue

        sValue = INIRead(sPath, "LocalLicense", "Key1-17", "")
        strLocalLicensserver = sValue

        sValue = INIRead(sPath, "StartDate", "Key1-18", "")
        strStartDate = sValue

        sValue = INIRead(sPath, "EndDate", "Key1-19", "")
        strEndDate = sValue


    End Sub

    Private Sub ReadiniFile_Folder()
       
        Dim SOURCE, DESGINGNATIONSUCCESS, DESGFAIL, strTempFolder1, strExportDirectory As String
        SOURCE = System.Windows.Forms.Application.StartupPath & "\ConnectionInfo\"
        Dim di As New IO.DirectoryInfo(sPath)
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.ini")
        Dim fi As IO.FileInfo
        Dim strsourcetrg, strdesgtrg, strImpFilename, strFile As String
        For Each fi In aryFi
            strImpFilename = fi.FullName
            strFile = fi.Name
            sPath = SOURCE & "\" & strFile & ".ini"
            sValue = INIRead(sPath, "SqlServer", "key1-1", "")
            sValue = INIRead(sPath, "SqlServer", "key1-1", "")
            txtServerName.Text = sValue
            strSQLServer = sValue
            sValue = INIRead(sPath, "SQLUID", "Key1-2", "")
            txtDBUserName.Text = sValue
            strSQLUID = sValue
            sValue = INIRead(sPath, "SQLPWD", "Key1-3", "")
            txtServerPwd.Text = sValue
            strSQLPWD = sValue

            sValue = INIRead(sPath, "ServerType", "Key1-S1", "")
            cmbServertype.Text = sValue
            strLocalServertype = sValue

            sValue = INIRead(sPath, "SAPCompany", "Key1-4", "")
            cmbCompanyDB.SelectedText = sValue
            strSAPCompany = sValue
            sValue = INIRead(sPath, "SAPUID", "Key1-5", "")
            txtSBOUserName.Text = sValue
            strSAPUID = sValue
            sValue = INIRead(sPath, "SAPPWD", "Key1-6", "")
            txtCompanyPWD.Text = sValue
            strSAPPWD = sValue
            strSAPPWD = sValue
            sValue = INIRead(sPath, "ExportDir", "Key1-7", "")
            '  txtDirectory.Text = sValue
            strExportFilePaty = sValue
            'sValue = INIRead(sPath, "ExportFile", "Key1-8", "")
            'strFileStart = sValue
            sValue = INIRead(sPath, "LogFile", "Key1-9", "")
            strErrorLogPath = sValue
            txtDirectory.Text = strErrorLogPath


            sValue = INIRead(sPath, "MainSqlServer", "key1-10", "")
            MainSQLServer.Text = sValue
            strMainSQLServer = sValue
            sValue = INIRead(sPath, "MainSQLUID", "Key1-11", "")
            MainSQLUID.Text = sValue
            strMainSQLUID = sValue
            sValue = INIRead(sPath, "MainSQLPWD", "Key1-12", "")
            MainSQLPWD.Text = sValue
            strMainSQLPWD = sValue
            sValue = INIRead(sPath, "MainServerType", "Key1-M1", "")
            cmbMainServertype.Text = sValue
            strMainServertype = sValue

            sValue = INIRead(sPath, "MainSAPCompany", "Key1-13", "")
            cmbMainSAPCompany.SelectedText = sValue
            strMainSAPCompany = sValue
            sValue = INIRead(sPath, "MainSAPUID", "Key1-14", "")
            MainSAPUID.Text = sValue
            strMainSAPUID = sValue
            sValue = INIRead(sPath, "MainSAPPWD", "Key1-15", "")
            MainSAPPWD.Text = sValue
            strMainSAPPWD = sValue
            strMainSAPPWD = sValue

            sValue = INIRead(sPath, "MainLicense", "Key1-16", "")
            mainLicenseServer.Text = sValue
            strMainLicenseServer = sValue

            sValue = INIRead(sPath, "LocalLicense", "Key1-17", "")
            LocalLicenseServer.Text = sValue
            strLocalLicensserver = sValue

        Next
    End Sub

#End Region

#Region "Get DBInformation from XML"

    '****************************************************************************
    'Type              :   Function     
    'Name              :   GetConInfo
    'Parameter         :   
    'Return Value      : 
    'Author            :   DEV-3
    'Created Date      :   
    'Last Modified By  : 
    'Modified Date     : 
    'Purpose           :   To Get DBInformation from XML
    '****************************************************************************

    Public Function GetConInfo() As String
        Dim objXMLDoc As New System.Xml.XmlDocument
        Dim objNode As Xml.XmlNode
        Dim intLoop As Integer
        Dim strConn, strserver, strDb, strUID, strPwd As String
        objXMLDoc.Load(System.Windows.Forms.Application.StartupPath & "\ConnectionInfo.xml")
        With objXMLDoc.GetElementsByTagName("Parameter")
            For intLoop = 0 To .Count - 1
                objNode = .Item(intLoop)
                If objNode.Attributes("Type").Value = "SBOServer" Then
                    strserver = objNode.Attributes("Value").Value
                ElseIf objNode.Attributes("Type").Value = "Database" Then
                    strDb = objNode.Attributes("Value").Value
                ElseIf objNode.Attributes("Type").Value = "DBUser" Then
                    strUID = objNode.Attributes("Value").Value
                ElseIf objNode.Attributes("Type").Value = "DBPass" Then
                    strPwd = objNode.Attributes("Value").Value
                End If
            Next
        End With
        strConn = "Database=" & strDb & ";Server=" & strserver & ";uid=" & strUID & ";Pwd=" & strPwd & ";"
        Return strConn
    End Function


#End Region

#Region "Events"
    Private Sub LoadForm()
        'System.Diagnostics.Debugger.Launch()
        CheckForIllegalCrossThreadCalls = False
        Me.Visible = False

        Dim strPath, strFilename, strMessage As String
        Dim SOURCE, DESGINGNATIONSUCCESS, DESGFAIL, strTempFolder1, strExportDirectory As String
        SOURCE = System.Windows.Forms.Application.StartupPath & "\ConnectionInfo"
        Dim di As New IO.DirectoryInfo(SOURCE)
        Dim aryFi As IO.FileInfo() = di.GetFiles("*.ini")
        Dim fi As IO.FileInfo
        Dim strsourcetrg, strdesgtrg, strImpFilename, strFile As String
        If CheckDuplicateProcess() = False Then
            End
        End If

        For Each fi In aryFi
            Try
                strImpFilename = fi.FullName
                strFile = fi.Name
                sPath = SOURCE & "\" & strFile ' & ".ini"
                ReadiniFile_FileCheck(sPath)
                ' ReadiniFile()
                If ValidateFilePaths() = False Then
                    End
                    'Exit Sub
                End If
                strPath = strErrorLogPath ' System.Windows.Forms.Application.StartupPath
                strFilename = Now.ToLongDateString
                strPath = strPath & "\Log_Documents" & strFilename & ".txt"
                strErrorFileName = strPath
                WriteErrorHeader(strPath, "Start connecting..")
                Dim styear, stMonth, stDate As String

                If strStartDate.Length = 8 Then
                    styear = strStartDate.Substring(0, 4)
                    stMonth = strStartDate.Substring(4, 2)
                    stDate = strStartDate.Substring(6, 2)
                    dtStartDate = New DateTime(CInt(styear), CInt(stMonth), CInt(stDate))
                End If

                If strEndDate.Length = 8 Then
                    styear = strEndDate.Substring(0, 4)
                    stMonth = strEndDate.Substring(4, 2)
                    stDate = strEndDate.Substring(6, 2)
                    dtEndDate = New DateTime(CInt(styear), CInt(stMonth), CInt(stDate))
                    ' dtEndDate = New DateTime(strEndDate.Substring(0, 4), strEndDate.Substring(5, 2), strEndDate.Substring(6, 2))
                End If
                If connectLocalCompanyCheck() = False Or connectMainCompanyCheck() = False Then
                    strMessage = ("Connection to SAP B1 failed. Check the ConnectionInfo.ini ")
                    WriteErrorlog(strMessage, strPath)
                    ' End
                Else
                    sPath = strErrorLogPath
                    Dim dtStart, dtEnd As Date
                    dtStart = dtStartDate
                    dtEnd = dtEndDate
                    WriteErrorlog("Documents Process Started  : " & objRemoteCompany.CompanyDB & " : Period  : " & dtStartDate.ToString("yyyy-MM-dd") & " To : " & dtEnd.ToString("yyyy-MM-dd"), sPath)
                    Do While dtStartDate <= dtEndDate
                        dtPostingDate = dtStartDate
                        WriteErrorlog("Documents Process Started  : " & objRemoteCompany.CompanyDB & " : Date : " & dtStartDate.ToString("yyyy-MM-dd"), sPath)

                        Try
                            ExportDocuments("PO")
                        Catch ex As Exception
                            strMessage = (ex.Message)
                            WriteErrorlog(strMessage, strPath)
                        End Try
                        Try
                            UpdateExchangerate_Main("USD", Now.Date, objMainCompany)
                        Catch ex As Exception
                        End Try
                        Dim OLoopRs, objRemoteRec1 As SAPbobsCOM.Recordset
                        Dim oDb As SAPbobsCOM.Recordset
                        oDb = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim s1 As String = "SELECT Code from [@Z_DBSYN] where U_Active='Y' and  '" & dtPostingDate.ToString("yyyy-MM-dd") & "' between U_FrmDate and U_ToDate"
                        OLoopRs = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        Dim s As String = "Select * from OJDT where isnull(U_Export,'N')='N'  and  RefDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and TransType in (" & s1 & ") order by TransID"
                        OLoopRs.DoQuery(s)
                        Dim strTransID As String
                        For intLoop1 As Integer = 0 To OLoopRs.RecordCount - 1
                            Dim strBaseType1 As String = OLoopRs.Fields.Item("TransType").Value
                            Dim strBaseRef1 As String = OLoopRs.Fields.Item("createdBy").Value
                            strTransID = OLoopRs.Fields.Item("TransID").Value
                            Select Case strBaseType1
                                Case "20"
                                    Try
                                        ExportDocuments("GRPO", strBaseRef1, strTransID)
                                    Catch ex As Exception
                                        strMessage = (ex.Message & " --> " & ex.StackTrace)
                                        WriteErrorlog(strMessage, strPath)
                                    End Try
                                Case "21"
                                    Try
                                        ExportDocuments("GR", strBaseRef1, strTransID)
                                    Catch ex As Exception
                                        strMessage = (ex.Message & " --> " & ex.StackTrace)
                                        WriteErrorlog(strMessage, strPath)
                                    End Try
                                Case "18"
                                    Try
                                        ExportDocuments("APInvoice", strBaseRef1, strTransID)
                                    Catch ex As Exception
                                        strMessage = (ex.Message & " --> " & ex.StackTrace)
                                        WriteErrorlog(strMessage, strPath)
                                    End Try

                                Case "19"
                                    Try
                                        ExportDocuments("APCR", strBaseRef1, strTransID)
                                    Catch ex As Exception

                                    End Try
                                Case "13"
                                    'sales document
                                    Try
                                        ExportDocuments("ARReser", strBaseRef1, strTransID)
                                    Catch ex As Exception
                                        strMessage = (ex.Message & " --> " & ex.StackTrace)
                                        WriteErrorlog(strMessage, strPath)
                                    End Try
                                    Try
                                        ExportDocuments_Canceled("ARReser", strBaseRef1, strTransID)
                                    Catch ex As Exception
                                        strMessage = (ex.Message & " --> " & ex.StackTrace)
                                        WriteErrorlog(strMessage, strPath)
                                    End Try
                                    Try
                                        ExportDocuments("Invoice", strBaseRef1, strTransID)
                                    Catch ex As Exception
                                        strMessage = (ex.Message & " --> " & ex.StackTrace)
                                        WriteErrorlog(strMessage, strPath)
                                    End Try
                                    Try
                                        ExportDocuments_Canceled("Invoice", strBaseRef1, strTransID)
                                    Catch ex As Exception
                                        strMessage = (ex.Message & " --> " & ex.StackTrace)
                                        WriteErrorlog(strMessage, strPath)
                                    End Try
                                Case "15"

                                    Try
                                        ExportDocuments("Delivery", strBaseRef1, strTransID)
                                    Catch ex As Exception
                                        strMessage = (ex.Message & " --> " & ex.StackTrace)
                                        WriteErrorlog(strMessage, strPath)
                                    End Try
                                Case "14"
                                    Try
                                        ExportDocuments("CreditNote", strBaseRef1, strTransID)
                                    Catch ex As Exception
                                        strMessage = (ex.Message & " --> " & ex.StackTrace)
                                        WriteErrorlog(strMessage, strPath)
                                    End Try
                                Case "30"
                                    objRemoteRec1 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    objRemoteRec1.DoQuery("SELECT * from [@Z_DBSYN] where isnull(U_Active,'N')='Y' and Code='30'")
                                    If objRemoteRec1.RecordCount > 0 Then
                                        Try
                                            ImportJournalEntries(OLoopRs.Fields.Item("TransID").Value)
                                        Catch ex As Exception
                                            strMessage = (ex.Message & " --> " & ex.StackTrace)
                                            WriteErrorlog(strMessage, strPath)
                                        End Try
                                    End If
                                Case "25"
                                    objRemoteRec1 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    objRemoteRec1.DoQuery("SELECT * from [@Z_DBSYN] where isnull(U_Active,'N')='Y' and Code='25'")
                                    If objRemoteRec1.RecordCount > 0 Then
                                        Try
                                            ImportBankDepositJournal(strBaseRef1)
                                        Catch ex As Exception
                                            strMessage = (ex.Message & " --> " & ex.StackTrace)
                                            WriteErrorlog(strMessage, strPath)
                                        End Try
                                    End If
                                Case "46"
                                    objRemoteRec1 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    objRemoteRec1.DoQuery("SELECT * from [@Z_DBSYN] where isnull(U_Active,'N')='Y' and Code='46'")
                                    If objRemoteRec1.RecordCount > 0 Then
                                        Try
                                            ExportOutgoingPaymentDocument(strBaseRef1)
                                        Catch ex As Exception
                                            strMessage = (ex.Message & " --> " & ex.StackTrace)
                                            WriteErrorlog(strMessage, strPath)
                                        End Try
                                    End If
                                Case "24"
                                    objRemoteRec1 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    objRemoteRec1.DoQuery("SELECT * from [@Z_DBSYN] where isnull(U_Active,'N')='Y' and Code='24'")
                                    If objRemoteRec1.RecordCount > 0 Then
                                        Try
                                            ExportIncomingPaymentDocument(strBaseRef1)
                                        Catch ex As Exception
                                            strMessage = (ex.Message & " --> " & ex.StackTrace)
                                            WriteErrorlog(strMessage, strPath)
                                        End Try
                                    End If

                            End Select
                            OLoopRs.MoveNext()
                            'Thread.Sleep(2000)
                        Next
                        dtStartDate = dtStartDate.AddDays(1)
                    Loop
                    WriteErrorlog("Export/Import Process Completed : " & objRemoteCompany.CompanyDB, sPath)
                End If
            Catch ex As Exception
                strMessage = (ex.Message)
                WriteErrorlog(strMessage, strPath)
                End
            End Try
        Next
        End

    End Sub
    Private Sub frmLoginScreen_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

    End Sub

    Private Sub frmLoginScreen_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        System.GC.Collect()
        End
    End Sub

#Region "Check Duplicate process"
    Private Function CheckDuplicateProcess() As Boolean
        Dim MatchingNames As System.Diagnostics.Process()
        Dim TargetName As String
        TargetName = System.Diagnostics.Process.GetCurrentProcess.ProcessName

        MatchingNames = System.Diagnostics.Process.GetProcessesByName(TargetName)
        If (MatchingNames.Length = 1) Then
            ' MsgBox("Started...")
            Return True

        Else
            '   MsgBox("Process already running")
            Return False

        End If



    End Function

    Public Function PrevInstance() As Boolean
        If UBound(Diagnostics.Process.GetProcessesByName(Diagnostics.Process.GetCurrentProcess.ProcessName)) > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

#End Region
    '****************************************************************************
    'Type              :   Procedure     
    'Name              :   
    'Parameter         :   
    'Return Value      : 
    'Author            :   DEV-3
    'Created Date      :   
    'Last Modified By  : 
    'Modified Date     : 
    'Purpose           :   To Handle Events
    '****************************************************************************
    Private Sub frmLoginScreen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim newThread As New Thread(Sub() Me.LoadForm())
        newThread.Start()

        'System.Diagnostics.Debugger.Launch()
        'Dim strPath, strFilename, strMessage As String
        'Dim SOURCE, DESGINGNATIONSUCCESS, DESGFAIL, strTempFolder1, strExportDirectory As String
        'SOURCE = System.Windows.Forms.Application.StartupPath & "\ConnectionInfo"
        'Dim di As New IO.DirectoryInfo(SOURCE)
        'Dim aryFi As IO.FileInfo() = di.GetFiles("*.ini")
        'Dim fi As IO.FileInfo
        'Dim strsourcetrg, strdesgtrg, strImpFilename, strFile As String
        'If CheckDuplicateProcess() = False Then
        '    End
        'End If

        'For Each fi In aryFi
        '    Try
        '        strImpFilename = fi.FullName
        '        strFile = fi.Name
        '        sPath = SOURCE & "\" & strFile ' & ".ini"
        '        ReadiniFile_File(sPath)
        '        ' ReadiniFile()
        '        If ValidateFilePaths() = False Then
        '            End
        '            'Exit Sub
        '        End If
        '        strPath = strErrorLogPath ' System.Windows.Forms.Application.StartupPath
        '        strFilename = Now.ToLongDateString
        '        strPath = strPath & "\Log_Documents" & strFilename & ".txt"
        '        strErrorFileName = strPath
        '        WriteErrorHeader(strPath, "Start connecting..")
        '        Dim styear, stMonth, stDate As String

        '        If strStartDate.Length = 8 Then
        '            styear = strStartDate.Substring(0, 4)
        '            stMonth = strStartDate.Substring(4, 2)
        '            stDate = strStartDate.Substring(6, 2)
        '            dtStartDate = New DateTime(CInt(styear), CInt(stMonth), CInt(stDate))
        '        End If

        '        If strEndDate.Length = 8 Then
        '            styear = strEndDate.Substring(0, 4)
        '            stMonth = strEndDate.Substring(4, 2)
        '            stDate = strEndDate.Substring(6, 2)
        '            dtEndDate = New DateTime(CInt(styear), CInt(stMonth), CInt(stDate))
        '            ' dtEndDate = New DateTime(strEndDate.Substring(0, 4), strEndDate.Substring(5, 2), strEndDate.Substring(6, 2))
        '        End If
        '        If connectLocalCompany() = False Or connectMainCompany() = False Then
        '            strMessage = ("Connection to SAP B1 failed. Check the ConnectionInfo.ini ")
        '            WriteErrorlog(strMessage, strPath)
        '            ' End
        '        Else
        '            sPath = strErrorLogPath
        '            Dim dtStart, dtEnd As Date
        '            dtStart = dtStartDate
        '            dtEnd = dtEndDate
        '            WriteErrorlog("Documents Process Started  : " & objRemoteCompany.CompanyDB & " : Period  : " & dtStartDate.ToString("yyyy-MM-dd") & " To : " & dtEnd.ToString("yyyy-MM-dd"), sPath)
        '            Do While dtStartDate <= dtEndDate
        '                dtPostingDate = dtStartDate
        '                WriteErrorlog("Documents Process Started  : " & objRemoteCompany.CompanyDB & " : Date : " & dtStartDate.ToString("yyyy-MM-dd"), sPath)
        '                Me.Hide()
        '                Try
        '                    ExportDocuments("PO")
        '                Catch ex As Exception
        '                    strMessage = (ex.Message)
        '                    WriteErrorlog(strMessage, strPath)
        '                End Try
        '                Try
        '                    UpdateExchangerate_Main("USD", Now.Date, objMainCompany)
        '                Catch ex As Exception
        '                End Try
        '                Dim OLoopRs, objRemoteRec1 As SAPbobsCOM.Recordset
        '                Dim oDb As SAPbobsCOM.Recordset
        '                oDb = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '                Dim s1 As String = "SELECT Code from [@Z_DBSYN] where U_Active='Y' and  '" & dtPostingDate.ToString("yyyy-MM-dd") & "' between U_FrmDate and U_ToDate"
        '                OLoopRs = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '                Dim s As String = "Select * from OJDT where isnull(U_Export,'N')='N'  and  RefDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and TransType in (" & s1 & ") order by TransID"
        '                OLoopRs.DoQuery(s)
        '                Dim strTransID As String
        '                For intLoop1 As Integer = 0 To OLoopRs.RecordCount - 1
        '                    Dim strBaseType1 As String = OLoopRs.Fields.Item("TransType").Value
        '                    Dim strBaseRef1 As String = OLoopRs.Fields.Item("createdBy").Value
        '                    strTransID = OLoopRs.Fields.Item("TransID").Value
        '                    Select Case strBaseType1
        '                        Case "20"
        '                            Try
        '                                ExportDocuments("GRPO", strBaseRef1, strTransID)
        '                            Catch ex As Exception
        '                                strMessage = (ex.Message & " --> " & ex.StackTrace)
        '                                WriteErrorlog(strMessage, strPath)
        '                            End Try
        '                        Case "21"
        '                            Try
        '                                ExportDocuments("GR", strBaseRef1, strTransID)
        '                            Catch ex As Exception
        '                                strMessage = (ex.Message & " --> " & ex.StackTrace)
        '                                WriteErrorlog(strMessage, strPath)
        '                            End Try
        '                        Case "18"
        '                            Try
        '                                ExportDocuments("APInvoice", strBaseRef1, strTransID)
        '                            Catch ex As Exception
        '                                strMessage = (ex.Message & " --> " & ex.StackTrace)
        '                                WriteErrorlog(strMessage, strPath)
        '                            End Try

        '                        Case "19"
        '                            Try
        '                                ExportDocuments("APCR", strBaseRef1, strTransID)
        '                            Catch ex As Exception

        '                            End Try
        '                        Case "13"
        '                            'sales document
        '                            Try
        '                                ExportDocuments("ARReser", strBaseRef1, strTransID)
        '                            Catch ex As Exception
        '                                strMessage = (ex.Message & " --> " & ex.StackTrace)
        '                                WriteErrorlog(strMessage, strPath)
        '                            End Try
        '                            Try
        '                                ExportDocuments_Canceled("ARReser", strBaseRef1, strTransID)
        '                            Catch ex As Exception
        '                                strMessage = (ex.Message & " --> " & ex.StackTrace)
        '                                WriteErrorlog(strMessage, strPath)
        '                            End Try
        '                            Try
        '                                ExportDocuments("Invoice", strBaseRef1, strTransID)
        '                            Catch ex As Exception
        '                                strMessage = (ex.Message & " --> " & ex.StackTrace)
        '                                WriteErrorlog(strMessage, strPath)
        '                            End Try
        '                            Try
        '                                ExportDocuments_Canceled("Invoice", strBaseRef1, strTransID)
        '                            Catch ex As Exception
        '                                strMessage = (ex.Message & " --> " & ex.StackTrace)
        '                                WriteErrorlog(strMessage, strPath)
        '                            End Try
        '                        Case "15"

        '                            Try
        '                                ExportDocuments("Delivery", strBaseRef1, strTransID)
        '                            Catch ex As Exception
        '                                strMessage = (ex.Message & " --> " & ex.StackTrace)
        '                                WriteErrorlog(strMessage, strPath)
        '                            End Try
        '                        Case "14"
        '                            Try
        '                                ExportDocuments("CreditNote", strBaseRef1, strTransID)
        '                            Catch ex As Exception
        '                                strMessage = (ex.Message & " --> " & ex.StackTrace)
        '                                WriteErrorlog(strMessage, strPath)
        '                            End Try
        '                        Case "30"
        '                            objRemoteRec1 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '                            objRemoteRec1.DoQuery("SELECT * from [@Z_DBSYN] where isnull(U_Active,'N')='Y' and Code='30'")
        '                            If objRemoteRec1.RecordCount > 0 Then
        '                                Try
        '                                    ImportJournalEntries(OLoopRs.Fields.Item("TransID").Value)
        '                                Catch ex As Exception
        '                                    strMessage = (ex.Message & " --> " & ex.StackTrace)
        '                                    WriteErrorlog(strMessage, strPath)
        '                                End Try
        '                            End If
        '                        Case "25"
        '                            objRemoteRec1 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '                            objRemoteRec1.DoQuery("SELECT * from [@Z_DBSYN] where isnull(U_Active,'N')='Y' and Code='25'")
        '                            If objRemoteRec1.RecordCount > 0 Then
        '                                Try
        '                                    ImportBankDepositJournal(strBaseRef1)
        '                                Catch ex As Exception
        '                                    strMessage = (ex.Message & " --> " & ex.StackTrace)
        '                                    WriteErrorlog(strMessage, strPath)
        '                                End Try
        '                            End If
        '                        Case "46"
        '                            objRemoteRec1 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '                            objRemoteRec1.DoQuery("SELECT * from [@Z_DBSYN] where isnull(U_Active,'N')='Y' and Code='46'")
        '                            If objRemoteRec1.RecordCount > 0 Then
        '                                Try
        '                                    ExportOutgoingPaymentDocument(strBaseRef1)
        '                                Catch ex As Exception
        '                                    strMessage = (ex.Message & " --> " & ex.StackTrace)
        '                                    WriteErrorlog(strMessage, strPath)
        '                                End Try
        '                            End If
        '                        Case "24"
        '                            objRemoteRec1 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '                            objRemoteRec1.DoQuery("SELECT * from [@Z_DBSYN] where isnull(U_Active,'N')='Y' and Code='24'")
        '                            If objRemoteRec1.RecordCount > 0 Then
        '                                Try
        '                                    ExportIncomingPaymentDocument(strBaseRef1)
        '                                Catch ex As Exception
        '                                    strMessage = (ex.Message & " --> " & ex.StackTrace)
        '                                    WriteErrorlog(strMessage, strPath)
        '                                End Try
        '                            End If

        '                    End Select
        '                    OLoopRs.MoveNext()
        '                    'Thread.Sleep(2000)
        '                Next
        '                dtStartDate = dtStartDate.AddDays(1)
        '            Loop
        '            WriteErrorlog("Export/Import Process Completed : " & objRemoteCompany.CompanyDB, sPath)
        '        End If
        '    Catch ex As Exception
        '        strMessage = (ex.Message)
        '        WriteErrorlog(strMessage, strPath)
        '        End
        '    End Try
        'Next
        'End
    End Sub

#Region "Export Incoming Payment Doocuments"

    Private Sub ExportOutgoingPaymentDocument(Optional ByVal aDocEntry As String = "")
        Dim oTempRec11, oTempRec2 As SAPbobsCOM.Recordset
        oTempRec11 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        oTempRec2 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRec2 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRec2.DoQuery("SELECT * from [@Z_DBSYN] where isnull(U_Active,'N')='Y' and Code='46'")
        Dim dtFrom, dtTo As Date
        If oTempRec2.RecordCount <= 0 Then
            Exit Sub
        Else
            dtFrom = oTempRec2.Fields.Item("U_FrmDate").Value
            dtTo = oTempRec2.Fields.Item("U_ToDate").Value
        End If
        Dim dtString As String
        If Year(dtFrom) = 1899 Then
            dtString = " 1=1"
        Else
            dtString = " T0.DocDate >='" & dtFrom.ToString("yyyy-MM-dd") & "'"

        End If

        If Year(dtTo) = 1899 Then
            dtString = dtString & " and  1=1"
        Else
            dtString = dtString & " and T0.DocDate <='" & dtTo.ToString("yyyy-MM-dd") & "'"

        End If

        'Post New Documents
        If aDocEntry <> "" Then
            oTempRec11.DoQuery("SELECT T0.DocEntry,T0.Series ,T0.DocType  FROM OVPM T0  inner Join  OCRD T1 on T1.CardCode=T0.CardCode WHERE T0.DocEntry=" & aDocEntry & " and  T0.DocDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and   T0.DocType<>'A' and T1.[U_sequencetype] ='P1' and isnull(T0.U_Export,'N')='N' and (" & dtString & ") order by T0.DocDate,T0.DocEntry")
        Else
            oTempRec11.DoQuery("SELECT T0.DocEntry,T0.Series ,T0.DocType  FROM OVPM T0  inner Join  OCRD T1 on T1.CardCode=T0.CardCode WHERE T0.DocDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and   T0.DocType<>'A' and T1.[U_sequencetype] ='P1' and isnull(T0.U_Export,'N')='N' and (" & dtString & ") order by T0.DocDate,T0.DocEntry")
        End If
        For intRow As Integer = 0 To oTempRec11.RecordCount - 1
            If oTempRec11.Fields.Item(2).Value = "A" Then
                If ValidateDocSeries(oTempRec11.Fields.Item(0).Value, "Incoming") = True Then
                    ExportOutgoingPayment(objMainCompany, oTempRec11.Fields.Item(0).Value)
                End If
            Else
                ExportOutgoingPayment(objMainCompany, oTempRec11.Fields.Item(0).Value)
            End If
            oTempRec11.MoveNext()
        Next
        If aDocEntry <> "" Then
            oTempRec11.DoQuery("SELECT T0.DocEntry,T0.Series ,T0.DocType FROM OVPM T0  WHERE T0.DocEntry=" & aDocEntry & " and  T0.DocDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and   T0.DocType='A'  and  isnull(T0.U_Export,'N')='N' and (" & dtString & ") order by T0.DocDate,T0.DocEntry")
        Else
            oTempRec11.DoQuery("SELECT T0.DocEntry,T0.Series ,T0.DocType FROM OVPM T0  WHERE T0.DocDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and   T0.DocType='A'  and  isnull(T0.U_Export,'N')='N' and (" & dtString & ") order by T0.DocDate,T0.DocEntry")
        End If

        For intRow As Integer = 0 To oTempRec11.RecordCount - 1
            If oTempRec11.Fields.Item(2).Value = "A" Then
                If ValidateDocSeries(oTempRec11.Fields.Item(0).Value, "OutGoing") = True Then
                    ExportOutgoingPayment(oCompany, oTempRec11.Fields.Item(0).Value)
                End If
            Else
                ExportOutgoingPayment(oCompany, oTempRec11.Fields.Item(0).Value)
            End If
            oTempRec11.MoveNext()
        Next
    End Sub
#End Region

#Region "Export InComing Payment Details"
    Private Sub ExportOutgoingPayment(ByVal aCompany As SAPbobsCOM.Company, ByVal aDocEntry As Integer)
        Dim objMainDoc, ObjRemoteDoc As SAPbobsCOM.Payments
        Dim objremoteRec, objRemoteRec1 As SAPbobsCOM.Recordset
        Dim strRemoteDocNumber As Integer
        objRemoteRec1 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objremoteRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objremoteRec.DoQuery("Select DocEntry,DocNum from OVPM where DocEntry=" & aDocEntry)
        For intRemoteloop As Integer = 0 To objremoteRec.RecordCount - 1
            Try
                ObjRemoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)
                objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)
                If ObjRemoteDoc.GetByKey(Convert.ToInt32(objremoteRec.Fields.Item(0).Value)) Then
                    If ValidateDocDate(ObjRemoteDoc.DocEntry, "46", ObjRemoteDoc.DocDate) = True Then
                        If ObjRemoteDoc.Cancelled = SAPbobsCOM.BoYesNoEnum.tNO Then
                            If checkDuplicateInvoice(ObjRemoteDoc.DocEntry, ObjRemoteDoc.DocNum, objMainCompany.CompanyName, "OVPM") = False Then
                                UpdateExchangerate_Main("USD", ObjRemoteDoc.DocDate, objMainCompany)
                                strRemoteDocNumber = ObjRemoteDoc.DocNum
                                objMainDoc.Address = ObjRemoteDoc.Address
                                objMainDoc.ApplyVAT = ObjRemoteDoc.ApplyVAT
                                objMainDoc.BankAccount = ObjRemoteDoc.BankAccount
                                If ObjRemoteDoc.BankChargeAmountInFC > 0 Then
                                    objMainDoc.BankChargeAmount = Math.Round(ObjRemoteDoc.BankChargeAmount / ObjRemoteDoc.DocRate, 2)
                                Else
                                    objMainDoc.BankChargeAmount = ObjRemoteDoc.BankChargeAmount
                                End If

                                objMainDoc.BankCode = ObjRemoteDoc.BankCode
                                objMainDoc.BillOfExchangeAgent = ObjRemoteDoc.BillOfExchangeAgent
                                objMainDoc.BillOfExchangeAgent = ObjRemoteDoc.BillOfExchangeAmount
                                objMainDoc.BillofExchangeStatus = ObjRemoteDoc.BillofExchangeStatus
                                objMainDoc.BoeAccount = ObjRemoteDoc.BoeAccount
                                objMainDoc.CardCode = ObjRemoteDoc.CardCode
                                objMainDoc.CardName = ObjRemoteDoc.CardName
                                objMainDoc.CashAccount = ObjRemoteDoc.CashAccount
                                objMainDoc.LocalCurrency = ObjRemoteDoc.LocalCurrency
                                If ObjRemoteDoc.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tYES Then
                                    objMainDoc.CashSum = ObjRemoteDoc.CashSum
                                Else
                                    If ObjRemoteDoc.DocRate > 0 Then
                                        objMainDoc.CashSum = Math.Round(ObjRemoteDoc.CashSum / ObjRemoteDoc.DocRate, 2)
                                    Else
                                        objMainDoc.CashSum = ObjRemoteDoc.CashSum
                                    End If

                                End If


                                'If ObjRemoteDoc.CashSumFC > 0 Then
                                '    objMainDoc.CashSum = ObjRemoteDoc.CashSum '/ ObjRemoteDoc.DocRate
                                'Else
                                '    objMainDoc.CashSum = ObjRemoteDoc.CashSum
                                'End If


                                objMainDoc.UserFields.Fields.Item("U_Import").Value = "Y"
                                objMainDoc.UserFields.Fields.Item("U_BaseEntry").Value = ObjRemoteDoc.DocEntry
                                objMainDoc.UserFields.Fields.Item("U_BaseNum").Value = ObjRemoteDoc.DocNum
                                objMainDoc.UserFields.Fields.Item("U_Branch").Value = objMainCompany.CompanyName

                                If ObjRemoteDoc.Checks.CheckSum <> 0 Then
                                    objMainDoc.CheckAccount = ObjRemoteDoc.CheckAccount
                                    For intCheck As Integer = 0 To ObjRemoteDoc.Checks.Count - 1
                                        If intCheck > 0 Then
                                            objMainDoc.Checks.Add()
                                            objMainDoc.Checks.SetCurrentLine(intCheck)
                                        End If
                                        ObjRemoteDoc.Checks.SetCurrentLine(intCheck)
                                        objMainDoc.Checks.AccounttNum = ObjRemoteDoc.Checks.AccounttNum
                                        objMainDoc.Checks.BankCode = ObjRemoteDoc.Checks.BankCode
                                        objMainDoc.Checks.Branch = ObjRemoteDoc.Checks.Branch
                                        objMainDoc.Checks.CheckAccount = ObjRemoteDoc.Checks.CheckAccount
                                        objMainDoc.Checks.CheckNumber = ObjRemoteDoc.Checks.CheckNumber
                                        '  MsgBox(ObjRemoteDoc.Checks.CheckSum)
                                        objMainDoc.Checks.CheckSum = ObjRemoteDoc.Checks.CheckSum

                                        If ObjRemoteDoc.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tYES Then
                                            objMainDoc.Checks.CheckSum = ObjRemoteDoc.Checks.CheckSum
                                        Else
                                            If ObjRemoteDoc.DocRate > 0 Then
                                                ' objMainDoc.Checks.CheckSum = Math.Round(ObjRemoteDoc.Checks.CheckSum * ObjRemoteDoc.DocRate, 2)
                                                objMainDoc.Checks.CheckSum = ObjRemoteDoc.Checks.CheckSum
                                            Else
                                                objMainDoc.Checks.CheckSum = ObjRemoteDoc.Checks.CheckSum
                                            End If

                                        End If


                                        objMainDoc.Checks.CountryCode = ObjRemoteDoc.Checks.CountryCode
                                        objMainDoc.Checks.Details = ObjRemoteDoc.Checks.Details
                                        objMainDoc.Checks.DueDate = ObjRemoteDoc.Checks.DueDate
                                        objMainDoc.Checks.ManualCheck = ObjRemoteDoc.Checks.ManualCheck
                                        objMainDoc.Checks.Trnsfrable = ObjRemoteDoc.Checks.Trnsfrable
                                        ' MsgBox(ObjRemoteDoc.Checks.CheckSum)
                                    Next
                                End If
                                objMainDoc.ContactPersonCode = ObjRemoteDoc.ContactPersonCode
                                objMainDoc.CounterReference = ObjRemoteDoc.CounterReference
                                If ObjRemoteDoc.CreditCards.CreditSum <> 0 Then
                                    objMainDoc.CreditCards.AdditionalPaymentSum = ObjRemoteDoc.CreditCards.AdditionalPaymentSum
                                    objMainDoc.CreditCards.CardValidUntil = ObjRemoteDoc.CreditCards.CardValidUntil
                                    objMainDoc.CreditCards.ConfirmationNum = ObjRemoteDoc.CreditCards.ConfirmationNum
                                    objMainDoc.CreditCards.CreditAcct = ObjRemoteDoc.CreditCards.CreditAcct
                                    objMainDoc.CreditCards.CreditCard = ObjRemoteDoc.CreditCards.CreditCard
                                    objMainDoc.CreditCards.CreditCardNumber = ObjRemoteDoc.CreditCards.CreditCardNumber
                                    objMainDoc.CreditCards.CreditSum = ObjRemoteDoc.CreditCards.CreditSum
                                    objMainDoc.CreditCards.CreditType = ObjRemoteDoc.CreditCards.CreditType
                                    objMainDoc.CreditCards.FirstPaymentDue = ObjRemoteDoc.CreditCards.FirstPaymentDue
                                    objMainDoc.CreditCards.FirstPaymentSum = ObjRemoteDoc.CreditCards.FirstPaymentSum
                                    objMainDoc.CreditCards.NumOfCreditPayments = ObjRemoteDoc.CreditCards.NumOfCreditPayments
                                    objMainDoc.CreditCards.NumOfPayments = ObjRemoteDoc.CreditCards.NumOfPayments
                                    objMainDoc.CreditCards.OwnerIdNum = ObjRemoteDoc.CreditCards.OwnerIdNum
                                    objMainDoc.CreditCards.OwnerPhone = ObjRemoteDoc.CreditCards.OwnerPhone
                                    objMainDoc.CreditCards.SplitPayments = ObjRemoteDoc.CreditCards.SplitPayments
                                    objMainDoc.CreditCards.PaymentMethodCode = ObjRemoteDoc.CreditCards.PaymentMethodCode
                                End If
                                objMainDoc.DeductionPercent = ObjRemoteDoc.DeductionPercent
                                objMainDoc.DeductionSum = ObjRemoteDoc.DeductionSum
                                objMainDoc.DocCurrency = ObjRemoteDoc.DocCurrency
                                objMainDoc.DocRate = ObjRemoteDoc.DocRate
                                objMainDoc.DocDate = ObjRemoteDoc.DocDate
                                '   MsgBox(ObjRemoteDoc.DocRate)
                                objMainDoc.DocTypte = ObjRemoteDoc.DocTypte
                                objMainDoc.DueDate = ObjRemoteDoc.DueDate
                                objMainDoc.DocObjectCode = ObjRemoteDoc.DocObjectCode
                                objMainDoc.HandWritten = ObjRemoteDoc.HandWritten
                                Dim intRefdocentry, intRemoteDocEntry As Integer
                                Dim intMultiple As Integer = 1
                                For intInvoice As Integer = 0 To ObjRemoteDoc.Invoices.Count - 1
                                    If ObjRemoteDoc.Invoices.DocEntry > 0 Then
                                        If intInvoice > 0 Then
                                            objMainDoc.Invoices.Add()
                                            objMainDoc.Invoices.SetCurrentLine(intInvoice)
                                        End If
                                        ObjRemoteDoc.Invoices.SetCurrentLine(intInvoice)
                                        intMultiple = 1
                                        objMainDoc.Invoices.DiscountPercent = ObjRemoteDoc.Invoices.DiscountPercent
                                        objMainDoc.Invoices.DistributionRule = ObjRemoteDoc.Invoices.DistributionRule
                                        intRemoteDocEntry = ObjRemoteDoc.Invoices.DocEntry

                                        If ObjRemoteDoc.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseCreditNote Then
                                            intRefdocentry = getbaseDocument(intRemoteDocEntry, strRemoteDocNumber, SAPbobsCOM.BoRcptInvTypes.it_PurchaseCreditNote, objMainCompany.CompanyName)
                                            intMultiple = -1
                                        ElseIf ObjRemoteDoc.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseDeliveryNote Then
                                            intRefdocentry = getbaseDocument(intRemoteDocEntry, strRemoteDocNumber, SAPbobsCOM.BoRcptInvTypes.it_PurchaseDeliveryNote, objMainCompany.CompanyName)
                                        ElseIf ObjRemoteDoc.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseReturn Then
                                            intRefdocentry = getbaseDocument(intRemoteDocEntry, strRemoteDocNumber, SAPbobsCOM.BoRcptInvTypes.it_PurchaseReturn, objMainCompany.CompanyName)
                                        ElseIf ObjRemoteDoc.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseInvoice Then
                                            intRefdocentry = getbaseDocument(intRemoteDocEntry, strRemoteDocNumber, SAPbobsCOM.BoRcptInvTypes.it_PurchaseInvoice, objMainCompany.CompanyName)
                                        ElseIf ObjRemoteDoc.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_JournalEntry Then
                                            intRefdocentry = getbaseDocument(intRemoteDocEntry, strRemoteDocNumber, SAPbobsCOM.BoRcptInvTypes.it_JournalEntry, objMainCompany.CompanyName)
                                        ElseIf ObjRemoteDoc.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PaymentAdvice Then
                                            intRefdocentry = getbaseDocument(intRemoteDocEntry, strRemoteDocNumber, SAPbobsCOM.BoRcptInvTypes.it_PaymentAdvice, objMainCompany.CompanyName)
                                            'ElseIf ObjRemoteDoc.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_PurchaseCreditNote Then
                                            '    intRefdocentry = getbaseDocument(intRemoteDocEntry, strRemoteDocNumber, SAPbobsCOM.BoRcptInvTypes.it_PurchaseCreditNote, objMainCompany.CompanyName)
                                        Else
                                            intRefdocentry = ObjRemoteDoc.Invoices.DocEntry
                                        End If
                                        'objMainDoc.Invoices.DocEntry  = ObjRemoteDoc.Invoices.DocEntry
                                        objMainDoc.Invoices.DocEntry = intRefdocentry
                                        objMainDoc.Invoices.DocLine = ObjRemoteDoc.Invoices.DocLine
                                        objMainDoc.Invoices.InstallmentId = ObjRemoteDoc.Invoices.InstallmentId

                                        objMainDoc.Invoices.InvoiceType = ObjRemoteDoc.Invoices.InvoiceType
                                        objMainDoc.Invoices.AppliedFC = ObjRemoteDoc.Invoices.AppliedFC * intMultiple
                                        If ObjRemoteDoc.Invoices.AppliedFC <> 0 Then
                                            objMainDoc.Invoices.SumApplied = Math.Round(ObjRemoteDoc.Invoices.SumApplied / ObjRemoteDoc.DocRate, 2) * intMultiple
                                        Else
                                            objMainDoc.Invoices.SumApplied = ObjRemoteDoc.Invoices.SumApplied * intMultiple
                                        End If

                                        objMainDoc.Invoices.SumApplied = ObjRemoteDoc.Invoices.SumApplied * intMultiple
                                        '   MsgBox(ObjRemoteDoc.Invoices.SumApplied)
                                    End If
                                Next
                                objMainDoc.IsPayToBank = ObjRemoteDoc.IsPayToBank
                                objMainDoc.JournalRemarks = ObjRemoteDoc.JournalRemarks
                                objMainDoc.LocalCurrency = ObjRemoteDoc.LocalCurrency
                                objMainDoc.PaymentPriority = ObjRemoteDoc.PaymentPriority
                                objMainDoc.PaymentType = ObjRemoteDoc.PaymentType
                                objMainDoc.PayToBankAccountNo = ObjRemoteDoc.PayToBankAccountNo
                                objMainDoc.PayToBankBranch = ObjRemoteDoc.PayToBankBranch
                                objMainDoc.PayToBankCode = ObjRemoteDoc.PayToBankCode
                                objMainDoc.PayToBankCountry = ObjRemoteDoc.PayToBankCountry
                                objMainDoc.ProjectCode = ObjRemoteDoc.ProjectCode
                                objMainDoc.Reference1 = ObjRemoteDoc.Reference1
                                objMainDoc.Reference2 = ObjRemoteDoc.Reference2
                                objMainDoc.Remarks = ObjRemoteDoc.Remarks
                                objMainDoc.Series = -1 ' ObjRemoteDoc.Series
                                objMainDoc.HandWritten = SAPbobsCOM.BoYesNoEnum.tYES
                                objMainDoc.DocNum = ObjRemoteDoc.DocNum
                                objMainDoc.TaxDate = ObjRemoteDoc.TaxDate
                                objMainDoc.TaxGroup = ObjRemoteDoc.TaxGroup
                                If ObjRemoteDoc.TransferSum > 0 Then
                                    objMainDoc.TransactionCode = ObjRemoteDoc.TransactionCode
                                    objMainDoc.TransferAccount = ObjRemoteDoc.TransferAccount
                                    objMainDoc.TransferRealAmount = ObjRemoteDoc.TransferRealAmount
                                    objMainDoc.TransferReference = ObjRemoteDoc.TransferReference
                                    objMainDoc.TransferDate = ObjRemoteDoc.TransferDate
                                    objRemoteRec1.DoQuery("SELECT T0.[TrsfrSum], T0.[TrsfrSumFC], T0.[TrsfrRef], T0.[TrsfrDate] FROM OVPM T0 WHERE T0.[DocNum] =" & ObjRemoteDoc.DocNum)
                                    'If objRemoteRec1.Fields.Item(1).Value > 0 Then
                                    '    objMainDoc.TransferSum = ObjRemoteDoc.TransferSum / ObjRemoteDoc.DocRate
                                    'Else
                                    '    objMainDoc.TransferSum = ObjRemoteDoc.TransferSum

                                    'End If
                                    If ObjRemoteDoc.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tYES Then
                                        objMainDoc.TransferSum = ObjRemoteDoc.TransferSum
                                    Else
                                        If ObjRemoteDoc.DocRate > 0 Then
                                            objMainDoc.TransferSum = Math.Round(ObjRemoteDoc.TransferSum / ObjRemoteDoc.DocRate, 2)
                                        Else
                                            objMainDoc.TransferSum = ObjRemoteDoc.TransferSum
                                        End If

                                    End If
                                    'objMainDoc.TransferSum = ObjRemoteDoc.TransferSum
                                End If
                                objMainDoc.VatDate = ObjRemoteDoc.VatDate
                                objMainDoc.WTAmount = ObjRemoteDoc.WTAmount
                                objMainDoc.WtBaseSum = ObjRemoteDoc.WtBaseSum
                                objMainDoc.WTCode = ObjRemoteDoc.WTCode
                                If ObjRemoteDoc.DocType = SAPbobsCOM.BoRcptTypes.rAccount Then
                                    For intAcctLoop As Integer = 0 To ObjRemoteDoc.AccountPayments.Count - 1
                                        If ObjRemoteDoc.AccountPayments.SumPaid <> 0 Then
                                            If intAcctLoop > 0 Then
                                                objMainDoc.AccountPayments.Add()
                                                objMainDoc.AccountPayments.SetCurrentLine(intAcctLoop)
                                            End If
                                            ObjRemoteDoc.AccountPayments.SetCurrentLine(intAcctLoop)
                                            objMainDoc.AccountPayments.AccountCode = ObjRemoteDoc.AccountPayments.AccountCode
                                            objMainDoc.AccountPayments.AccountName = ObjRemoteDoc.AccountPayments.AccountName
                                            objMainDoc.AccountPayments.Decription = ObjRemoteDoc.AccountPayments.Decription
                                            '   objMainDoc.AccountPayments.GrossAmount = ObjRemoteDoc.AccountPayments.GrossAmount
                                            '  objMainDoc.AccountPayments.SumPaid = ObjRemoteDoc.AccountPayments.SumPaid
                                            objMainDoc.AccountPayments.VatGroup = ObjRemoteDoc.AccountPayments.VatGroup
                                            objRemoteRec1.DoQuery("Select * from VPM4 where DocNum=" & ObjRemoteDoc.DocEntry & " and LineID=" & intAcctLoop)
                                            If objRemoteRec1.Fields.Item("AppliedFC").Value <> 0 Then
                                                objMainDoc.AccountPayments.SumPaid = Math.Round(ObjRemoteDoc.AccountPayments.SumPaid / ObjRemoteDoc.DocRate, 2)
                                                objMainDoc.AccountPayments.VatAmount = Math.Round(ObjRemoteDoc.AccountPayments.VatAmount / ObjRemoteDoc.DocRate, 2)
                                            Else
                                                objMainDoc.AccountPayments.SumPaid = ObjRemoteDoc.AccountPayments.SumPaid
                                                objMainDoc.AccountPayments.VatAmount = ObjRemoteDoc.AccountPayments.VatAmount
                                            End If
                                            objMainDoc.AccountPayments.ProfitCenter = ObjRemoteDoc.AccountPayments.ProfitCenter
                                            objMainDoc.AccountPayments.ProjectCode = ObjRemoteDoc.AccountPayments.ProjectCode
                                        End If
                                    Next
                                End If
                                If objMainDoc.Add() <> 0 Then
                                    sPath = strErrorLogPath
                                    WriteErrorlog("Failed to Export Outgoing payment  Document:" & ObjRemoteDoc.DocNum & ":" & objRemoteCompany.GetLastErrorDescription & "   : Branch : " & objMainCompany.CompanyName, sPath)
                                Else
                                    Dim strDocNumber As String
                                    objRemoteCompany.GetNewObjectCode(strDocNumber)
                                    'MsgBox(strDocNumber)
                                    sPath = strErrorLogPath
                                    Dim oInvoice As SAPbobsCOM.Payments
                                    oInvoice = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)
                                    If oInvoice.GetByKey(Convert.ToInt64(strDocNumber)) Then
                                        strDocNumber = oInvoice.DocNum
                                    End If
                                    WriteErrorlog("Outgoing Payment document  Created Successfully. DocNum : " & strDocNumber, sPath)
                                    ObjRemoteDoc.UserFields.Fields.Item("U_Export").Value = "Y"
                                    ObjRemoteDoc.Update()
                                End If
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception
                WriteErrorlog("Failed to create Outgoing payment Error  :" & ex.Message & "--->" & ex.StackTrace, sPath)
            Finally
                If Not IsNothing(ObjRemoteDoc) Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ObjRemoteDoc)
                End If
                If Not IsNothing(objMainDoc) Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objMainDoc)
                End If
            End Try
            objremoteRec.MoveNext()
        Next
    End Sub



    Private Sub ExportOutgoingPayment_Cancelled(ByVal aCompany As SAPbobsCOM.Company, ByVal aDocEntry As Integer)
        Dim objMainDoc, ObjRemoteDoc As SAPbobsCOM.Payments
        Dim objremoteRec, objRemoteRec1 As SAPbobsCOM.Recordset
        Dim strRemoteDocNumber As Integer
        objRemoteRec1 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objremoteRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objremoteRec.DoQuery("Select DocEntry,DocNum from OVPM where DocEntry=" & aDocEntry)
        For intRemoteloop As Integer = 0 To objremoteRec.RecordCount - 1
            ObjRemoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)
            objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVendorPayments)
            If ObjRemoteDoc.GetByKey(Convert.ToInt32(objremoteRec.Fields.Item(0).Value)) Then
                If ValidateDocDate(ObjRemoteDoc.DocEntry, "46", ObjRemoteDoc.DocDate) = True Then
                    If ObjRemoteDoc.Cancelled = SAPbobsCOM.BoYesNoEnum.tYES Then
                        If CheckCanceledInvoice(ObjRemoteDoc.DocEntry, ObjRemoteDoc.DocNum, objMainCompany.CompanyName, "OVPM") <> 0 Then
                            If objMainDoc.GetByKey(CheckCanceledInvoice(ObjRemoteDoc.DocEntry, ObjRemoteDoc.DocNum, objMainCompany.CompanyName, "OVPM")) Then
                                If objMainDoc.Cancelled = SAPbobsCOM.BoYesNoEnum.tNO Then
                                    If objMainDoc.Cancel <> 0 Then
                                        '  MsgBox(objRemoteCompany.GetLastErrorDescription)
                                    Else
                                        WriteErrorlog("Outgoing Payment document  Cancelled Successfully. DocNum : " & ObjRemoteDoc.DocNum, sPath)
                                        '  MsgBox("Canceld")
                                    End If
                                End If
                            End If
                        End If
                    Else
                    End If
                End If
            End If
            objremoteRec.MoveNext()
        Next

    End Sub

#End Region
#Region "Export InComing Payment Details"
    Private Sub ExportIncomingPayment1(ByVal aCompany As SAPbobsCOM.Company, ByVal aDocEntry As Integer)
        Dim objMainDoc, ObjRemoteDoc As SAPbobsCOM.Payments
        Dim objremoteRec As SAPbobsCOM.Recordset
        Dim strRemoteDocNumber As Integer
        objremoteRec = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objremoteRec.DoQuery("Select DocEntry,DocNum from ORCT where DocEntry=" & aDocEntry)
        For intRemoteloop As Integer = 0 To objremoteRec.RecordCount - 1
            ObjRemoteDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
            objMainDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
            If ObjRemoteDoc.GetByKey(Convert.ToInt32(objremoteRec.Fields.Item(0).Value)) Then
                strRemoteDocNumber = ObjRemoteDoc.DocNum
                objMainDoc.Address = ObjRemoteDoc.Address
                objMainDoc.ApplyVAT = ObjRemoteDoc.ApplyVAT
                objMainDoc.BankAccount = ObjRemoteDoc.BankAccount
                objMainDoc.BankChargeAmount = ObjRemoteDoc.BankChargeAmount
                objMainDoc.BankCode = ObjRemoteDoc.BankCode
                objMainDoc.BillOfExchangeAgent = ObjRemoteDoc.BillOfExchangeAgent
                objMainDoc.BillOfExchangeAgent = ObjRemoteDoc.BillOfExchangeAmount
                objMainDoc.BillofExchangeStatus = ObjRemoteDoc.BillofExchangeStatus
                objMainDoc.BoeAccount = ObjRemoteDoc.BoeAccount
                objMainDoc.CardCode = ObjRemoteDoc.CardCode
                objMainDoc.CardName = ObjRemoteDoc.CardName
                objMainDoc.CashAccount = ObjRemoteDoc.CashAccount
                objMainDoc.CashSum = ObjRemoteDoc.CashSum
                If ObjRemoteDoc.Checks.CheckSum > 0 Then
                    objMainDoc.CheckAccount = ObjRemoteDoc.CheckAccount
                    For intCheck As Integer = 0 To ObjRemoteDoc.Checks.Count - 1
                        If intCheck > 0 Then
                            objMainDoc.Checks.Add()
                            objMainDoc.Checks.SetCurrentLine(intCheck)
                        End If
                        ObjRemoteDoc.Checks.SetCurrentLine(intCheck)
                        objMainDoc.Checks.AccounttNum = ObjRemoteDoc.Checks.AccounttNum
                        objMainDoc.Checks.BankCode = ObjRemoteDoc.Checks.BankCode
                        objMainDoc.Checks.Branch = ObjRemoteDoc.Checks.Branch
                        objMainDoc.Checks.CheckAccount = ObjRemoteDoc.Checks.CheckAccount
                        objMainDoc.Checks.CheckAccount = ObjRemoteDoc.Checks.CheckNumber
                        objMainDoc.Checks.CheckSum = ObjRemoteDoc.Checks.CheckSum
                        objMainDoc.Checks.CountryCode = ObjRemoteDoc.Checks.CountryCode
                        objMainDoc.Checks.Details = ObjRemoteDoc.Checks.Details
                        objMainDoc.Checks.DueDate = ObjRemoteDoc.Checks.DueDate
                        objMainDoc.Checks.ManualCheck = ObjRemoteDoc.Checks.ManualCheck
                        objMainDoc.Checks.Trnsfrable = ObjRemoteDoc.Checks.Trnsfrable
                        '    MsgBox(ObjRemoteDoc.Checks.CheckSum)
                    Next
                End If
                objMainDoc.ContactPersonCode = ObjRemoteDoc.ContactPersonCode
                objMainDoc.CounterReference = ObjRemoteDoc.CounterReference
                If ObjRemoteDoc.CreditCards.CreditSum > 0 Then
                    objMainDoc.CreditCards.AdditionalPaymentSum = ObjRemoteDoc.CreditCards.AdditionalPaymentSum
                    objMainDoc.CreditCards.CardValidUntil = ObjRemoteDoc.CreditCards.CardValidUntil
                    objMainDoc.CreditCards.ConfirmationNum = ObjRemoteDoc.CreditCards.ConfirmationNum
                    objMainDoc.CreditCards.CreditAcct = ObjRemoteDoc.CreditCards.CreditAcct
                    objMainDoc.CreditCards.CreditCard = ObjRemoteDoc.CreditCards.CreditCard
                    objMainDoc.CreditCards.CreditCardNumber = ObjRemoteDoc.CreditCards.CreditCardNumber
                    objMainDoc.CreditCards.CreditSum = ObjRemoteDoc.CreditCards.CreditSum
                    objMainDoc.CreditCards.CreditType = ObjRemoteDoc.CreditCards.CreditType
                    objMainDoc.CreditCards.FirstPaymentDue = ObjRemoteDoc.CreditCards.FirstPaymentDue
                    objMainDoc.CreditCards.FirstPaymentSum = ObjRemoteDoc.CreditCards.FirstPaymentSum
                    objMainDoc.CreditCards.NumOfCreditPayments = ObjRemoteDoc.CreditCards.NumOfCreditPayments
                    objMainDoc.CreditCards.NumOfPayments = ObjRemoteDoc.CreditCards.NumOfPayments
                    objMainDoc.CreditCards.OwnerIdNum = ObjRemoteDoc.CreditCards.OwnerIdNum
                    objMainDoc.CreditCards.OwnerPhone = ObjRemoteDoc.CreditCards.OwnerPhone
                    objMainDoc.CreditCards.SplitPayments = ObjRemoteDoc.CreditCards.SplitPayments
                    objMainDoc.CreditCards.PaymentMethodCode = ObjRemoteDoc.CreditCards.PaymentMethodCode
                End If
                objMainDoc.DeductionPercent = ObjRemoteDoc.DeductionPercent
                objMainDoc.DeductionSum = ObjRemoteDoc.DeductionSum
                objMainDoc.DocCurrency = ObjRemoteDoc.DocCurrency
                objMainDoc.DocRate = ObjRemoteDoc.DocRate
                objMainDoc.DocTypte = ObjRemoteDoc.DocTypte
                objMainDoc.DueDate = ObjRemoteDoc.DueDate
                objMainDoc.DocObjectCode = ObjRemoteDoc.DocObjectCode
                objMainDoc.HandWritten = ObjRemoteDoc.HandWritten
                Dim intRefdocentry, intRemoteDocEntry As Integer
                For intInvoice As Integer = 0 To ObjRemoteDoc.Invoices.Count - 1
                    If ObjRemoteDoc.Invoices.DocEntry > 0 Then
                        If intInvoice > 0 Then
                            objMainDoc.Invoices.Add()
                            objMainDoc.Invoices.SetCurrentLine(intInvoice)
                        End If
                        ObjRemoteDoc.Invoices.SetCurrentLine(intInvoice)
                        objMainDoc.Invoices.AppliedFC = ObjRemoteDoc.Invoices.AppliedFC
                        objMainDoc.Invoices.DiscountPercent = ObjRemoteDoc.Invoices.DiscountPercent
                        objMainDoc.Invoices.DistributionRule = ObjRemoteDoc.Invoices.DistributionRule
                        intRemoteDocEntry = ObjRemoteDoc.Invoices.DocEntry
                        If ObjRemoteDoc.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice Then
                            intRefdocentry = GetBaseDocEntry(intRemoteDocEntry, strRemoteDocNumber, "Invoice", objRemoteCompany.CompanyName)
                        ElseIf ObjRemoteDoc.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_CredItnote Then
                            intRefdocentry = GetBaseDocEntry(intRemoteDocEntry, strRemoteDocNumber, "CreditNote", objRemoteCompany.CompanyName)
                        Else
                            intRefdocentry = ObjRemoteDoc.Invoices.DocEntry
                        End If
                        'objMainDoc.Invoices.DocEntry  = ObjRemoteDoc.Invoices.DocEntry
                        objMainDoc.Invoices.DocEntry = intRefdocentry
                        objMainDoc.Invoices.DocLine = ObjRemoteDoc.Invoices.DocLine
                        objMainDoc.Invoices.InstallmentId = ObjRemoteDoc.Invoices.InstallmentId
                        objMainDoc.Invoices.InvoiceType = ObjRemoteDoc.Invoices.InvoiceType
                        objMainDoc.Invoices.SumApplied = ObjRemoteDoc.Invoices.SumApplied
                        '   MsgBox(ObjRemoteDoc.Invoices.SumApplied)
                    End If
                Next
                objMainDoc.IsPayToBank = ObjRemoteDoc.IsPayToBank
                objMainDoc.JournalRemarks = ObjRemoteDoc.JournalRemarks
                objMainDoc.LocalCurrency = ObjRemoteDoc.LocalCurrency
                objMainDoc.PaymentPriority = ObjRemoteDoc.PaymentPriority
                objMainDoc.PaymentType = ObjRemoteDoc.PaymentType
                objMainDoc.PayToBankAccountNo = ObjRemoteDoc.PayToBankAccountNo
                objMainDoc.PayToBankBranch = ObjRemoteDoc.PayToBankBranch
                objMainDoc.PayToBankCode = ObjRemoteDoc.PayToBankCode
                objMainDoc.PayToBankCountry = ObjRemoteDoc.PayToBankCountry
                objMainDoc.ProjectCode = ObjRemoteDoc.ProjectCode
                objMainDoc.Reference1 = ObjRemoteDoc.Reference1
                objMainDoc.Reference2 = ObjRemoteDoc.Reference2
                objMainDoc.Remarks = ObjRemoteDoc.Remarks
                objMainDoc.Series = ObjRemoteDoc.Series
                objMainDoc.TaxDate = ObjRemoteDoc.TaxDate
                objMainDoc.TaxGroup = ObjRemoteDoc.TaxGroup
                objMainDoc.TransactionCode = ObjRemoteDoc.TransactionCode
                objMainDoc.TransferAccount = ObjRemoteDoc.TransferAccount
                objMainDoc.TransferRealAmount = ObjRemoteDoc.TransferRealAmount
                objMainDoc.TransferReference = ObjRemoteDoc.TransferReference
                objMainDoc.TransferSum = ObjRemoteDoc.TransferSum
                objMainDoc.VatDate = ObjRemoteDoc.VatDate
                objMainDoc.WTAmount = ObjRemoteDoc.WTAmount
                objMainDoc.WtBaseSum = ObjRemoteDoc.WtBaseSum
                objMainDoc.WTCode = ObjRemoteDoc.WTCode
                If ObjRemoteDoc.DocType = SAPbobsCOM.BoRcptTypes.rAccount Then
                    For intAcctLoop As Integer = 0 To ObjRemoteDoc.AccountPayments.Count - 1
                        If ObjRemoteDoc.AccountPayments.SumPaid > 0 Then
                            If intAcctLoop > 0 Then
                                objMainDoc.AccountPayments.Add()
                                objMainDoc.AccountPayments.SetCurrentLine(intAcctLoop)
                            End If
                            ObjRemoteDoc.AccountPayments.SetCurrentLine(intAcctLoop)
                            objMainDoc.AccountPayments.AccountCode = ObjRemoteDoc.AccountPayments.AccountCode
                            objMainDoc.AccountPayments.AccountName = ObjRemoteDoc.AccountPayments.AccountName
                            objMainDoc.AccountPayments.Decription = ObjRemoteDoc.AccountPayments.Decription
                            objMainDoc.AccountPayments.GrossAmount = ObjRemoteDoc.AccountPayments.GrossAmount
                            objMainDoc.AccountPayments.SumPaid = ObjRemoteDoc.AccountPayments.SumPaid
                            objMainDoc.AccountPayments.ProfitCenter = ObjRemoteDoc.AccountPayments.ProfitCenter
                            objMainDoc.AccountPayments.ProjectCode = ObjRemoteDoc.AccountPayments.ProjectCode
                            objMainDoc.AccountPayments.VatAmount = ObjRemoteDoc.AccountPayments.VatAmount
                            objMainDoc.AccountPayments.VatGroup = ObjRemoteDoc.AccountPayments.VatGroup
                        End If
                    Next
                End If
            End If

            If objMainDoc.Add() <> 0 Then
                sPath = strErrorLogPath
                WriteErrorlog("Failed to Export Incoming payment  Document:" & ObjRemoteDoc.DocNum & ":" & objMainCompany.GetLastErrorDescription & "   : Branch : " & objRemoteCompany.CompanyName, sPath)
            Else
                Dim strDocNumber As String
                objMainCompany.GetNewObjectCode(strDocNumber)
                'MsgBox(strDocNumber)
                sPath = strErrorLogPath
                Dim oInvoice As SAPbobsCOM.Payments
                oInvoice = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
                If oInvoice.GetByKey(Convert.ToInt64(strDocNumber)) Then
                    strDocNumber = oInvoice.DocNum
                End If
                WriteErrorlog("Incoming Payment document  Created Successfully. DocNum : " & strDocNumber, sPath)
                ObjRemoteDoc.UserFields.Fields.Item("U_Export").Value = "Y"
                ObjRemoteDoc.Update()

            End If
            objremoteRec.MoveNext()
        Next

    End Sub

#End Region
#Region "Update Exchange rate"
    Private Sub UpdateExchangerate(ByVal strCurrency As String, ByVal dtdate As Date, ByVal aCompany As SAPbobsCOM.Company)
        Dim oexchrate, oExchnageMain As SAPbobsCOM.Recordset
        Dim dblRate, dblRate1 As Double
        Dim strSQL As String
        oExchnageMain = aCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oExchnageMain.DoQuery("Select Currency from ORTT where convert(nvarchar(10),rateDate,101)='" & dtdate.ToString("MM/dd/yyyy") & "' and isnull(rate,0)<>0")
        For intLoop As Integer = 0 To oExchnageMain.RecordCount - 1
            strCurrency = oExchnageMain.Fields.Item(0).Value
            If strCurrency <> "" Then
                oexchrate = aCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                strSQL = "Select * from ORTT where Currency='" & strCurrency & "' and convert(nvarchar(10),rateDate,101)='" & dtdate.ToString("MM/dd/yyyy") & "'"
                oexchrate.DoQuery(strSQL)
                If oexchrate.RecordCount > 0 Then
                    dblRate1 = oexchrate.Fields.Item("Rate").Value
                End If
                oexchrate = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oexchrate.DoQuery("Select * from ORTT where Currency='" & strCurrency & "' and convert(nvarchar(10),rateDate,101)='" & dtdate.ToString("MM/dd/yyyy") & "'")
                If oexchrate.RecordCount > 0 Then
                    dblRate = oexchrate.Fields.Item("Rate").Value
                    If dblRate <= 0 Or dblRate <> dblRate1 Then
                        oexchrate = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oexchrate.DoQuery("Update ORTT set Rate=" & dblRate1 & " where Currency='" & strCurrency & "' and  convert(nvarchar(10),rateDate,101)='" & dtdate.ToString("MM/dd/yyyy") & "'")
                    End If
                Else
                    oexchrate = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    strSQL = "insert into ORTT values ('" & dtdate.ToString("yyyy-MM-dd") & "','" & strCurrency & "'," & dblRate1 & ",'I',1)"
                    oexchrate.DoQuery(strSQL)
                End If
            End If
            oExchnageMain.MoveNext()
        Next
    End Sub


    Private Sub UpdateExchangerate_Main(ByVal strCurrency As String, ByVal dtdate As Date, ByVal aCompany As SAPbobsCOM.Company)
        Dim oexchrate, oexchrate1 As SAPbobsCOM.Recordset
        Dim dblRate, dblRate1 As Double
        Dim strSQL As String
        oexchrate = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strSQL = "Select * from ORTT where convert(nvarchar(10),rateDate,101)='" & dtdate.ToString("MM/dd/yyyy") & "' and Rate<>0"
        oexchrate.DoQuery(strSQL)
        For intRow As Integer = 0 To oexchrate.RecordCount - 1
            strCurrency = oexchrate.Fields.Item("Currency").Value
            dblRate1 = oexchrate.Fields.Item("Rate").Value
            oexchrate1 = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oexchrate1.DoQuery("Select * from ORTT where Currency='" & strCurrency & "' and convert(nvarchar(10),rateDate,101)='" & dtdate.ToString("MM/dd/yyyy") & "'")
            If oexchrate1.RecordCount > 0 Then
                dblRate = oexchrate1.Fields.Item("Rate").Value
                If dblRate <= 0 Then
                    oexchrate1 = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oexchrate1.DoQuery("Update ORTT set Rate=" & dblRate1 & " where Currency='" & strCurrency & "' and rateDate=" & dtdate)
                End If
            Else
                oexchrate1 = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                strSQL = "insert into ORTT values ('" & dtdate.ToString("yyyy-MM-dd") & "','" & strCurrency & "'," & dblRate1 & ",'I',1)"
                oexchrate1.DoQuery(strSQL)
            End If

            oexchrate.MoveNext()
        Next
    End Sub
#Region "Check Duplicate Entry"
    Private Function checkDuplicateJEs(ByVal aDocEntry As Integer, ByVal aDocNum As Integer, ByVal aBranchcode As String, ByVal aTableName As String) As Boolean
        Dim oDuplicateCheck As SAPbobsCOM.Recordset
        Dim strQuery1 As String
        oDuplicateCheck = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strQuery1 = "Select * from " & aTableName & " where U_BaseEntry in (" & aDocEntry & ") and U_BaseNum in (" & aDocNum & ") and U_Branch='" & aBranchcode & "'"
        oDuplicateCheck.DoQuery(strQuery1)
        If oDuplicateCheck.RecordCount > 0 Then
            Return True
        Else
            Return False
        End If

    End Function
#End Region
#Region "Get Account Code"
    Private Function getFormatCode(ByVal aCode As String, ByVal aCompany As SAPbobsCOM.Company) As String
        Dim otest As SAPbobsCOM.Recordset
        otest = aCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest.DoQuery("Select Formatcode from OACT where AcctCode='" & aCode & "'")
        Return otest.Fields.Item(0).Value
    End Function

    Private Function getAccountCode(ByVal aCode As String, ByVal aCompany As SAPbobsCOM.Company) As String
        Dim otest As SAPbobsCOM.Recordset
        otest = aCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest.DoQuery("Select AcctCode from OACT where FormatCode='" & aCode & "'")
        Return otest.Fields.Item(0).Value
    End Function
#End Region
#End Region

    Private Sub ImportBankDepositJournal(Optional ByVal aDeposite As String = "")
        Dim objMainJE, objRemoteJE As SAPbobsCOM.JournalEntries
        Dim objremoteRec1, otemp1 As SAPbobsCOM.Recordset
        Dim strAcctCode, strFormatCode As String
        Dim intDocEntry As Integer
        objremoteRec1 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp1 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        Dim oTempRec11, oTempRec2 As SAPbobsCOM.Recordset
        oTempRec11 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRec2 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRec2 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRec2.DoQuery("SELECT * from [@Z_DBSYN] where isnull(U_Active,'N')='Y' and Code='25'")
        Dim dtFrom, dtTo As Date
        If oTempRec2.RecordCount <= 0 Then
            Exit Sub
        Else
            dtFrom = oTempRec2.Fields.Item("U_FrmDate").Value
            dtTo = oTempRec2.Fields.Item("U_ToDate").Value
        End If
        Dim dtString As String
        If Year(dtFrom) = 1899 Then
            dtString = " 1=1"
        Else
            dtString = " DeposDate >='" & dtFrom.ToString("yyyy-MM-dd") & "'"

        End If

        If Year(dtTo) = 1899 Then
            dtString = dtString & " and  1=1"
        Else
            dtString = dtString & " and DeposDate <='" & dtTo.ToString("yyyy-MM-dd") & "'"

        End If
        If aDeposite <> "" Then
            objremoteRec1.DoQuery("Select * from ODPS where DeposId=" & aDeposite & " and  DeposDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and  isnull(U_Export,'N')='N' and (" & dtString & ") order by DeposDate,DeposId")
        Else
            objremoteRec1.DoQuery("Select * from ODPS where DeposDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and  isnull(U_Export,'N')='N' and (" & dtString & ") order by DeposDate,DeposId")
        End If

        For intLoop As Integer = 0 To objremoteRec1.RecordCount - 1
            intDocEntry = objremoteRec1.Fields.Item("Deposid").Value
            If 1 = 1 Then
                If ValidateDocSeries_Deposit(intDocEntry, "25") = True Then
                    If ValidateDocDate(intDocEntry, "25", objremoteRec1.Fields.Item("DeposDate").Value) = True Then
                        If ImportBankDeposit(objremoteRec1.Fields.Item("TransAbs").Value) = True Then
                            otemp1.DoQuery("Update ODPS set U_Export='Y' where Deposid=" & intDocEntry)
                        End If
                    End If
                End If
            End If
            objremoteRec1.MoveNext()
        Next
    End Sub
    Private Function ImportBankDeposit(ByVal aTransId As Integer) As Boolean
        Dim objMainJE, objRemoteJE As SAPbobsCOM.JournalEntries
        Dim objremoteRec, otemp As SAPbobsCOM.Recordset
        Dim strAcctCode, strFormatCode As String
        objremoteRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objremoteRec.DoQuery("Select * from OJDT where [TransId]=" & aTransId & " and  transtype=25 and  isnull(U_Export,'N')='N'")
        For intLoop As Integer = 0 To objremoteRec.RecordCount - 1
            objRemoteJE = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
            If objRemoteJE.GetByKey(objremoteRec.Fields.Item("TransID").Value) Then
                If checkDuplicateJEs(objRemoteJE.JdtNum, objRemoteJE.Number, objMainCompany.CompanyName, "OJDT") = True Then
                    WriteErrorlog("Journal Entry : " & objRemoteJE.Number & " already exported --> DataBase : " & objRemoteCompany.CompanyName, sPath)
                Else
                    objMainJE = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                    objMainJE.DueDate = objRemoteJE.DueDate
                    objMainJE.Memo = objRemoteJE.Memo
                    objMainJE.ProjectCode = objRemoteJE.ProjectCode
                    objMainJE.Reference = objRemoteJE.Reference
                    objMainJE.Reference2 = objRemoteJE.Reference2
                    objMainJE.ReferenceDate = objRemoteJE.ReferenceDate
                    objMainJE.TaxDate = objRemoteJE.TaxDate
                    '     objMainJE.TransactionCode = objRemoteJE.TransactionCode
                    objMainJE.VatDate = objRemoteJE.VatDate
                    objMainJE.UserFields.Fields.Item("U_BaseEntry").Value = objRemoteJE.JdtNum
                    objMainJE.UserFields.Fields.Item("U_Branch").Value = objMainCompany.CompanyName
                    objMainJE.UserFields.Fields.Item("U_BaseNum").Value = objRemoteJE.Number
                    'UpdateExchangerate_Main("", objRemoteJE.TaxDate, objRemoteCompany)
                    UpdateExchangerate("", objRemoteJE.TaxDate, objRemoteCompany)
                    For intLines As Integer = 0 To objRemoteJE.Lines.Count - 1
                        If intLines > 0 Then
                            objMainJE.Lines.Add()
                        End If
                        objMainJE.Lines.SetCurrentLine(intLines)
                        objRemoteJE.Lines.SetCurrentLine(intLines)

                        strFormatCode = getFormatCode(objRemoteJE.Lines.AccountCode, objMainCompany)
                        strAcctCode = getAccountCode(strFormatCode, objRemoteCompany)
                        objMainJE.Lines.AccountCode = strAcctCode
                        'objMainJE.Lines.AccountCode = objRemoteJE.Lines.AccountCode
                        objMainJE.Lines.AdditionalReference = objRemoteJE.Lines.AdditionalReference
                        objMainJE.Lines.BaseSum = objRemoteJE.Lines.BaseSum
                        objMainJE.Lines.ContraAccount = objRemoteJE.Lines.ContraAccount
                        objMainJE.Lines.CostingCode = objRemoteJE.Lines.CostingCode
                        If objRemoteJE.Lines.FCCurrency <> "" Then
                            objMainJE.Lines.FCCurrency = objRemoteJE.Lines.FCCurrency
                            ' UpdateExchangerate(objRemoteJE.Lines.FCCurrency, objRemoteJE.TaxDate, objRemoteCompany)
                            objMainJE.Lines.FCDebit = objRemoteJE.Lines.FCDebit
                            objMainJE.Lines.FCCredit = objRemoteJE.Lines.FCCredit
                        Else
                            objMainJE.Lines.Credit = objRemoteJE.Lines.Credit
                            objMainJE.Lines.Debit = objRemoteJE.Lines.Debit
                        End If
                        objMainJE.Lines.DueDate = objRemoteJE.Lines.DueDate
                        objMainJE.Lines.LineMemo = objRemoteJE.Lines.LineMemo
                        '  objMainJE.Lines.LocationCode = objRemoteJE.Lines.LocationCode
                        objMainJE.Lines.ProjectCode = objRemoteJE.Lines.ProjectCode
                        objMainJE.Lines.Reference1 = objRemoteJE.Lines.Reference1
                        objMainJE.Lines.Reference2 = objRemoteJE.Lines.Reference2
                        objMainJE.Lines.ReferenceDate1 = objRemoteJE.Lines.ReferenceDate1
                        objMainJE.Lines.ReferenceDate2 = objRemoteJE.Lines.ReferenceDate2
                        objMainJE.Lines.ShortName = objRemoteJE.Lines.ShortName
                        objMainJE.Lines.TaxCode = objRemoteJE.Lines.TaxCode
                        objMainJE.Lines.TaxDate = objRemoteJE.Lines.TaxDate
                        objMainJE.Lines.TaxGroup = objRemoteJE.Lines.TaxGroup
                    Next
                    If objMainJE.Add <> 0 Then
                        WriteErrorlog("Failed to create Journal Entry  " & objRemoteJE.Number & " : Database : " & objMainCompany.CompanyName & " :Error : " & objRemoteCompany.GetLastErrorDescription, sPath)
                        Return False
                    Else
                        Dim strDocNum As String
                        objRemoteCompany.GetNewObjectCode(strDocNum)
                        WriteErrorlog("Journal entry created successully " & strDocNum & " : Database : " & objRemoteCompany.CompanyName, sPath)
                        objRemoteJE.UserFields.Fields.Item("U_Export").Value = "Y"
                        objRemoteJE.Update()

                    End If
                End If
            End If
            objremoteRec.MoveNext()
        Next
        Return True
    End Function

    Private Sub ImportJournalEntries(Optional ByVal aTransId As String = "")
        Dim objMainJE, objRemoteJE As SAPbobsCOM.JournalEntries
        Dim objremoteRec, otemp As SAPbobsCOM.Recordset
        Dim strAcctCode, strFormatCode As String
        Dim oTempRec11, oTempRec2 As SAPbobsCOM.Recordset
        objremoteRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRec11 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRec2 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRec2 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRec2.DoQuery("SELECT * from [@Z_DBSYN] where isnull(U_Active,'N')='Y' and Code='30'")
        Dim dtFrom, dtTo As Date
        If oTempRec2.RecordCount <= 0 Then
            Exit Sub
        Else
            dtFrom = oTempRec2.Fields.Item("U_FrmDate").Value
            dtTo = oTempRec2.Fields.Item("U_ToDate").Value
        End If
        Dim dtString As String
        If Year(dtFrom) = 1899 Then
            dtString = " 1=1"
        Else
            dtString = " RefDate >='" & dtFrom.ToString("yyyy-MM-dd") & "'"

        End If

        If Year(dtTo) = 1899 Then
            dtString = dtString & " and  1=1"
        Else
            dtString = dtString & " and RefDate <='" & dtTo.ToString("yyyy-MM-dd") & "'"

        End If

        If aTransId <> "" Then
            objremoteRec.DoQuery("Select * from OJDT  where transID=" & aTransId & " and  RefDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and  transtype=30 and TransCode='P1' and isnull(U_Export,'N')='N' and (" & dtString & ") order by TransID")
        Else
            objremoteRec.DoQuery("Select * from OJDT  where RefDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and  transtype=30 and TransCode='P1' and isnull(U_Export,'N')='N' and (" & dtString & ") order by TransID")

        End If
        For intLoop As Integer = 0 To objremoteRec.RecordCount - 1
            objRemoteJE = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
            If objRemoteJE.GetByKey(objremoteRec.Fields.Item("TransID").Value) Then
                If checkDuplicateJEs(objRemoteJE.JdtNum, objRemoteJE.Number, objMainCompany.CompanyName, "OJDT") = True Then
                    WriteErrorlog("Journal Entry : " & objRemoteJE.Number & " already exported --> DataBase : " & objRemoteCompany.CompanyName, sPath)
                Else
                    If ValidateDocDate(objRemoteJE.JdtNum, "30", objRemoteJE.TaxDate) = True Then
                        objMainJE = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries)
                        If File.Exists(System.Windows.Forms.Application.StartupPath & "\JE.XML") Then
                            File.Delete(System.Windows.Forms.Application.StartupPath & "\JE.XML")
                        End If
                        'objRemoteJE.SaveXML(System.Windows.Forms.Application.StartupPath & "\JE.XML")
                        'objMainJE = objRemoteCompany.GetBusinessObjectFromXML(System.Windows.Forms.Application.StartupPath & "\JE.XML", 0)
                        'If objMainJE.Add <> 0 Then
                        '    MsgBox(objRemoteCompany.GetLastErrorDescription)
                        'End If

                        objMainJE.DueDate = objRemoteJE.DueDate
                        objMainJE.Memo = objRemoteJE.Memo
                        objMainJE.ProjectCode = objRemoteJE.ProjectCode
                        objMainJE.Reference = objRemoteJE.Reference
                        objMainJE.Reference2 = objRemoteJE.Reference2
                        objMainJE.ReferenceDate = objRemoteJE.ReferenceDate
                        objMainJE.TaxDate = objRemoteJE.TaxDate
                        '     objMainJE.TransactionCode = objRemoteJE.TransactionCode
                        objMainJE.VatDate = objRemoteJE.VatDate
                        objMainJE.UserFields.Fields.Item("U_BaseEntry").Value = objRemoteJE.JdtNum
                        objMainJE.UserFields.Fields.Item("U_Branch").Value = objMainCompany.CompanyName
                        objMainJE.UserFields.Fields.Item("U_BaseNum").Value = objRemoteJE.Number
                        objMainJE.AutoVAT = objRemoteJE.AutoVAT
                        objMainJE.AutomaticWT = objRemoteJE.AutomaticWT
                        objMainJE.TransactionCode = objRemoteJE.TransactionCode
                        ' objMainJE.UseAutoStorno = objRemoteJE.UseAutoStorno
                        'UpdateExchangerate_Main("", objRemoteJE.TaxDate, objRemoteCompany)
                        UpdateExchangerate("", objRemoteJE.TaxDate, objRemoteCompany)
                        Dim intCount As Integer = 0
                        For intLines As Integer = 0 To objRemoteJE.Lines.Count - 1
                            objRemoteJE.Lines.SetCurrentLine(intLines)
                            If 1 = 1 Then ' objRemoteJE.Lines.VatLine = SAPbobsCOM.BoYesNoEnum.tNO Then
                                If intCount > 0 Then
                                    objMainJE.Lines.Add()
                                End If
                                objMainJE.Lines.SetCurrentLine(intCount)
                                strFormatCode = getFormatCode(objRemoteJE.Lines.AccountCode, objMainCompany)
                                strAcctCode = getAccountCode(strFormatCode, objRemoteCompany)
                                objMainJE.Lines.AccountCode = strAcctCode
                                'objMainJE.Lines.AccountCode = objRemoteJE.Lines.AccountCode
                                objMainJE.Lines.AdditionalReference = objRemoteJE.Lines.AdditionalReference
                                objMainJE.Lines.BaseSum = objRemoteJE.Lines.BaseSum
                                objMainJE.Lines.ContraAccount = objRemoteJE.Lines.ContraAccount
                                objMainJE.Lines.CostingCode = objRemoteJE.Lines.CostingCode
                                If objRemoteJE.Lines.FCCurrency <> "" Then
                                    objMainJE.Lines.FCCurrency = objRemoteJE.Lines.FCCurrency
                                    ' UpdateExchangerate(objRemoteJE.Lines.FCCurrency, objRemoteJE.TaxDate, objRemoteCompany)
                                    objMainJE.Lines.FCDebit = objRemoteJE.Lines.FCDebit
                                    objMainJE.Lines.FCCredit = objRemoteJE.Lines.FCCredit
                                    objMainJE.Lines.DebitSys = objRemoteJE.Lines.DebitSys
                                    objMainJE.Lines.CreditSys = objRemoteJE.Lines.CreditSys
                                    objMainJE.Lines.Credit = objRemoteJE.Lines.Credit
                                    objMainJE.Lines.Debit = objRemoteJE.Lines.Debit
                                Else
                                    '  objMainJE.Lines.FCCurrency = ""
                                    objMainJE.Lines.Credit = objRemoteJE.Lines.Credit
                                    objMainJE.Lines.Debit = objRemoteJE.Lines.Debit
                                    objMainJE.Lines.DebitSys = objRemoteJE.Lines.DebitSys
                                    objMainJE.Lines.CreditSys = objRemoteJE.Lines.CreditSys
                                    objMainJE.Lines.FCDebit = objRemoteJE.Lines.FCDebit
                                    objMainJE.Lines.FCCredit = objRemoteJE.Lines.FCCredit
                                End If

                                objMainJE.Lines.DueDate = objRemoteJE.Lines.DueDate
                                objMainJE.Lines.LineMemo = objRemoteJE.Lines.LineMemo
                                '  objMainJE.Lines.LocationCode = objRemoteJE.Lines.LocationCode
                                objMainJE.Lines.ProjectCode = objRemoteJE.Lines.ProjectCode
                                objMainJE.Lines.Reference1 = objRemoteJE.Lines.Reference1
                                objMainJE.Lines.Reference2 = objRemoteJE.Lines.Reference2
                                objMainJE.Lines.ReferenceDate1 = objRemoteJE.Lines.ReferenceDate1
                                objMainJE.Lines.ReferenceDate2 = objRemoteJE.Lines.ReferenceDate2
                                objMainJE.Lines.ShortName = objRemoteJE.Lines.ShortName
                                objMainJE.Lines.TaxCode = objRemoteJE.Lines.TaxCode
                                objMainJE.Lines.TaxDate = objRemoteJE.Lines.TaxDate
                                objMainJE.Lines.TaxGroup = objRemoteJE.Lines.TaxGroup
                                objMainJE.Lines.VatLine = objRemoteJE.Lines.VatLine
                                'If objMainJE.AutomaticWT = SAPbobsCOM.BoYesNoEnum.tYES Then
                                '    objMainJE.Lines.VatAmount = objRemoteJE.Lines.VatAmount
                                'End If


                                'objMainJE.Lines.VatDate = objRemoteJE.Lines.VatDate
                                intCount = intCount + 1
                            End If
                        Next
                        If objMainJE.Add <> 0 Then
                            WriteErrorlog("Failed to create Journal Entry  " & objRemoteJE.Number & " : Database : " & objMainCompany.CompanyName & " :Error : " & objRemoteCompany.GetLastErrorDescription, sPath)
                        Else
                            Dim strDocNum As String
                            objRemoteCompany.GetNewObjectCode(strDocNum)
                            WriteErrorlog("Journal entry created successully " & strDocNum & " : Database : " & objRemoteCompany.CompanyName, sPath)
                            objRemoteJE.UserFields.Fields.Item("U_Export").Value = "Y"
                            objRemoteJE.Update()
                        End If
                    End If
                End If
            End If
            objremoteRec.MoveNext()
        Next
    End Sub
    Private Sub ExportDocuments(ByVal aChoice As String, Optional ByVal aDocEntry As String = "", Optional ByVal aJournalNo As String = "")
        Dim conn As New SqlConnection()
        Dim objMainDoc, objremoteDoc, objMainDoc1 As SAPbobsCOM.Documents
        Dim intRemoteDocEntry As Integer
        Dim strPath, strFilename, strMessage, strTableName, stMessagetext As String
        Dim strFileName1 As String
        Dim objremoteRec As SAPbobsCOM.Recordset
        Dim strCondition As String = "1=1"
        strTableName = ""
        Dim strChildTable As String
        Dim strCode As String = ""
        Dim blnCanceled As Boolean = False
        Try

            'objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
            'objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
            Select Case aChoice
                Case "PO"
                    strCode = "22"
                    strTableName = "OPOR"
                    stMessagetext = "Purchase Document"
                    objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                    objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                Case "PQ"
                    strTableName = ""
                    stMessagetext = "Purchase Quotation"
                    objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseQuotations)
                    objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseQuotations)
                Case "GRPO"
                    strCode = "20"
                    strTableName = "OPDN"
                    stMessagetext = "Goods Receipt Purchase Order"
                    objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
                    objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
                Case "GR"
                    strCode = "21"
                    strTableName = "ORPD"
                    stMessagetext = "Goods Return"
                    objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseReturns)
                    objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseReturns)
                Case "APInvoice"
                    strCode = "18"
                    strTableName = "OPCH"
                    stMessagetext = "A/P Invoice"
                    objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                    objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                    '   blnCanceled = True
                Case "APCR"
                    strCode = "19"
                    strTableName = "ORPC"
                    stMessagetext = "A/P Credit Memo"
                    objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
                    objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)

                Case "ARReser"
                    'isIns]
                    strCode = "171"
                    strCondition = " IsIns='Y'"
                    strTableName = "OINV"
                    stMessagetext = "A/R Invoice"
                    objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                    objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                    blnCanceled = True
                Case "Delivery"
                    strCode = "15"
                    strTableName = "ODLN"
                    stMessagetext = "Delivery"
                    objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                    objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                Case "Invoice"
                    strCode = "13"
                    strCondition = " IsIns='N'"
                    strTableName = "OINV"
                    stMessagetext = "A/R Invoice"
                    blnCanceled = True
                    objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                    objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                Case "CreditNote"
                    strCode = "14"
                    strTableName = "ORIN"
                    stMessagetext = "A/R Credit Note"
                    objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                    objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
            End Select

            If strTableName = "" Then
                If Not IsNothing(objremoteDoc) Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objremoteDoc)
                End If
                If Not IsNothing(objMainDoc) Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objMainDoc)
                End If
                Exit Sub
            End If
            objremoteRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objremoteRec.DoQuery("SELECT * from [@Z_DBSYN] where isnull(U_Active,'N')='Y' and Code='" & strCode & "'")
            Dim dtFrom, dtTo As Date
            If objremoteRec.RecordCount <= 0 Then
                If Not IsNothing(objremoteDoc) Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objremoteDoc)
                End If
                If Not IsNothing(objMainDoc) Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objMainDoc)
                End If
                If Not IsNothing(objremoteRec) Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objremoteRec)
                End If
                Exit Sub
            Else
                dtFrom = objremoteRec.Fields.Item("U_FrmDate").Value
                dtTo = objremoteRec.Fields.Item("U_ToDate").Value
            End If

            Dim dtString As String
            If Year(dtFrom) = 1899 Then
                dtString = " 1=1"
            Else
                dtString = " T0.DocDate >='" & dtFrom.ToString("yyyy-MM-dd") & "'"
            End If

            If Year(dtTo) = 1899 Then
                dtString = dtString & " and  1=1"
            Else
                dtString = dtString & " and T0.DocDate <='" & dtTo.ToString("yyyy-MM-dd") & "'"
            End If
            '  blnCanceled = True
            Dim s As String
            If aDocEntry <> "" Then
                If blnCanceled = False Then
                    s = "Select T0.DocEntry from " & strTableName & " T0 inner Join OCRD T1 on T1.CardCode=T0.CardCode  where T0.DocEntry=" & aDocEntry & " and  T0.DocDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and " & strCondition & " and T0.Canceled='N' and isnull(T1.U_sequencetype,'N')='P1' and isnull(U_Export,'N')='N' and ( " & dtString & ") order by T0.DocDate,T0.DocEntry"
                Else
                    s = "Select T0.DocEntry from " & strTableName & " T0 inner Join OCRD T1 on T1.CardCode=T0.CardCode  where T0.DocEntry=" & aDocEntry & " and  T0.DocDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and " & strCondition & " and T0.Canceled<>'C' and isnull(T1.U_sequencetype,'N')='P1' and isnull(U_Export,'N')='N' and ( " & dtString & ") order by T0.DocDate,T0.DocEntry"
                End If
            Else
                If blnCanceled = False Then
                    s = "Select T0.DocEntry from " & strTableName & " T0 inner Join OCRD T1 on T1.CardCode=T0.CardCode  where T0.DocDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and " & strCondition & " and T0.Canceled='N' and isnull(T1.U_sequencetype,'N')='P1' and isnull(U_Export,'N')='N' and ( " & dtString & ") order by T0.DocDate,T0.DocEntry"
                Else
                    s = "Select T0.DocEntry from " & strTableName & " T0 inner Join OCRD T1 on T1.CardCode=T0.CardCode  where T0.DocDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and " & strCondition & " and T0.Canceled<>'C' and isnull(T1.U_sequencetype,'N')='P1' and isnull(U_Export,'N')='N' and ( " & dtString & ") order by T0.DocDate,T0.DocEntry"
                End If
            End If

            'Dim s As String = "Select T0.DocEntry from " & strTableName & " T0 inner Join OCRD T1 on T1.CardCode=T0.CardCode  where " & strCondition & "  and isnull(T1.U_sequencetype,'N')='P1' and isnull(U_Export,'N')='N' and ( " & dtString & ") order by T0.DocDate,T0.DocEntry"
            objremoteRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objremoteRec.DoQuery(s)
            '  Dim objMainDoc1 As SAPbobsCOM.Documents
            For intRemoteLoop As Integer = 0 To objremoteRec.RecordCount - 1
                Try
                    If objremoteDoc.GetByKey(Convert.ToInt32(objremoteRec.Fields.Item(0).Value)) Then

                        If ValidateDocDate(objremoteDoc.DocEntry, strCode, objremoteDoc.DocDate) = False Then

                        Else
                            If checkDuplicateInvoice(objremoteDoc.DocEntry, objremoteDoc.DocNum, objMainCompany.CompanyName, strTableName) = True Then
                                objremoteDoc.UserFields.Fields.Item("U_Export").Value = "Y"
                                objremoteDoc.Update()
                                Dim oTE As SAPbobsCOM.Recordset
                                oTE = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                If aJournalNo <> "" Then
                                    oTE.DoQuery("Update OJDT set U_Export='Y' where TransId=" & aJournalNo)
                                    If Not IsNothing(oTE) Then
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oTE)
                                    End If
                                End If
                            Else
                                intRemoteDocEntry = objremoteDoc.DocEntry
                                '    UpdateExchangerate_Main(objremoteDoc.DocCurrency, objremoteDoc.DocDate, objMainCompany)
                                Select Case aChoice
                                    Case "PO"
                                        strTableName = "OPOR"
                                        'objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                                        objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                                        objMainDoc1 = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                                    Case "PQ"
                                        strTableName = "OPQT"
                                        'objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                                        objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseQuotations)
                                        objMainDoc1 = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseQuotations)
                                    Case "GRPO"
                                        strTableName = "OPDN"
                                        ' objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
                                        objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
                                        objMainDoc1 = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
                                    Case "GR"
                                        strTableName = "ORPD"
                                        ' objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
                                        objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseReturns)
                                        objMainDoc1 = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseReturns)
                                    Case "APInvoice"
                                        strTableName = "OPCH"
                                        ' objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                                        objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                                        objMainDoc1 = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                                    Case "APCR"
                                        strTableName = "ORPC"
                                        'objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
                                        objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
                                        objMainDoc1 = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)

                                    Case "ARReser"
                                        'isIns]
                                        strCondition = " IsIns='Y'"
                                        strTableName = "OINV"
                                        stMessagetext = "A/R Invoice"
                                        ' objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                                        objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                                        objMainDoc1 = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                                    Case "Delivery"
                                        strTableName = "ODLN"
                                        stMessagetext = "Delivery"
                                        ' objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                                        objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                                        objMainDoc1 = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                                    Case "Invoice"
                                        strCondition = " IsIns='N'"
                                        strTableName = "OINV"
                                        stMessagetext = "A/R Invoice"
                                        ' objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                                        objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                                        objMainDoc1 = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                                    Case "CreditNote"
                                        strTableName = "ORIN"
                                        stMessagetext = "A/R Credit Note"
                                        ' objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                                        objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                                        objMainDoc1 = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                                End Select

                                If 1 = 2 Then 'objremoteDoc.Cancelled = SAPbobsCOM.BoYesNoEnum.tYES Then '  CheckCanceledInvoice(objremoteDoc.DocEntry, objremoteDoc.DocNum, objMainCompany.CompanyName, strTableName) <> 0 Then
                                    If objMainDoc.GetByKey(CheckCanceledInvoice(objremoteDoc.DocEntry, objremoteDoc.DocNum, objMainCompany.CompanyName, strTableName)) Then
                                        Dim otest As SAPbobsCOM.Documents
                                        If blnCanceled = True Then
                                            objMainDoc1 = objMainDoc.CreateCancellationDocument()
                                            If objMainDoc1.Add = 0 Then
                                                Dim strDocNum As String
                                                objRemoteCompany.GetNewObjectCode(strDocNum)
                                                Dim oInvoice As SAPbobsCOM.Documents
                                                Dim intInvoiceDocentry As String
                                                oInvoice = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                                                If objMainDoc.GetByKey(Convert.ToInt64(strDocNum)) Then
                                                    strDocNum = objMainDoc.DocNum
                                                    intInvoiceDocentry = objMainDoc.DocEntry
                                                End If
                                                sPath = strErrorLogPath
                                                WriteErrorlog(stMessagetext & " Created Successfully. DocNum : " & strDocNum, sPath)
                                                Dim otempRec As SAPbobsCOM.Recordset
                                                otempRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                objremoteDoc.GetByKey(intRemoteDocEntry)
                                                objremoteDoc.UserFields.Fields.Item("U_Export").Value = "Y"
                                                objremoteDoc.Update()
                                            End If
                                        End If
                                    End If
                                Else
                                    UpdateExchangerate_Main(objremoteDoc.DocCurrency, objremoteDoc.DocDate, objMainCompany)
                                    objMainDoc.CardCode = objremoteDoc.CardCode
                                    objMainDoc.DocDate = objremoteDoc.DocDate
                                    objMainDoc.DocDueDate = objremoteDoc.DocDueDate
                                    objMainDoc.HandWritten = SAPbobsCOM.BoYesNoEnum.tYES
                                    objMainDoc.Series = -1
                                    objMainDoc.DocNum = objremoteDoc.DocNum
                                    objMainDoc.NumAtCard = objremoteDoc.NumAtCard
                                    objMainDoc.Comments = objremoteDoc.Comments
                                    objMainDoc.DiscountPercent = objremoteDoc.DiscountPercent
                                    objMainDoc.DocCurrency = objremoteDoc.DocCurrency
                                    objMainDoc.DocRate = objremoteDoc.DocRate
                                    objMainDoc.ShipToCode = objremoteDoc.ShipToCode
                                    objMainDoc.SalesPersonCode = objremoteDoc.SalesPersonCode
                                    objMainDoc.TaxDate = objremoteDoc.TaxDate
                                    objMainDoc.PaymentGroupCode = objremoteDoc.PaymentGroupCode
                                    objMainDoc.PaymentMethod = objremoteDoc.PaymentMethod
                                    objMainDoc.Address = objremoteDoc.Address
                                    objMainDoc.Address2 = objremoteDoc.Address2
                                    objMainDoc.AgentCode = objremoteDoc.AgentCode
                                    objMainDoc.BPChannelCode = objremoteDoc.BPChannelCode
                                    objMainDoc.BPChannelContact = objremoteDoc.BPChannelContact
                                    '  objMainDoc.ContactPersonCode = objremoteDoc.ContactPersonCode
                                    objMainDoc.ReserveInvoice = objremoteDoc.ReserveInvoice
                                    If objremoteDoc.FederalTaxID <> "" Then
                                        ' MsgBox(objremoteDoc.FederalTaxID)
                                        Dim oBP As SAPbobsCOM.BusinessPartners
                                        'oBP = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                                        'If oBP.GetByKey(objremoteDoc.CardCode) Then
                                        '    ' MsgBox(oBP.FederalTaxID)
                                        'End If
                                        objMainDoc.FederalTaxID = objremoteDoc.FederalTaxID
                                    End If
                                    ' objMainDoc.DocRate = objremoteDoc.DocRate
                                    objMainDoc.DocType = objremoteDoc.DocType
                                    '  objMainDoc.TransportationCode = objremoteDoc.TransportationCode
                                    If objremoteDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES Then
                                        objMainDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES
                                        objMainDoc.RoundingDiffAmount = objremoteDoc.RoundingDiffAmount
                                    Else
                                        objMainDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tNO
                                    End If
                                    For intLoop As Integer = 0 To objremoteDoc.UserFields.Fields.Count - 1
                                        Try
                                            objMainDoc.UserFields.Fields.Item(intLoop).Value = objremoteDoc.UserFields.Fields.Item(intLoop).Value
                                        Catch ex As Exception

                                        End Try
                                    Next
                                    objMainDoc.UserFields.Fields.Item("U_Import").Value = "Y"
                                    objMainDoc.UserFields.Fields.Item("U_BaseEntry").Value = objremoteDoc.DocEntry
                                    objMainDoc.UserFields.Fields.Item("U_BaseNum").Value = objremoteDoc.DocNum
                                    objMainDoc.UserFields.Fields.Item("U_Branch").Value = objMainCompany.CompanyName
                                    For IntExp As Integer = 0 To objremoteDoc.Expenses.Count - 1
                                        If objremoteDoc.Expenses.LineTotal > 0 Then
                                            If IntExp > 0 Then
                                                objMainDoc.Expenses.Add()
                                                objMainDoc.Expenses.SetCurrentLine(IntExp)
                                            End If
                                            objremoteDoc.Expenses.SetCurrentLine(IntExp)
                                            objMainDoc.Expenses.BaseDocEntry = objremoteDoc.Expenses.BaseDocEntry
                                            objMainDoc.Expenses.BaseDocLine = objremoteDoc.Expenses.BaseDocLine
                                            objMainDoc.Expenses.BaseDocType = objremoteDoc.Expenses.BaseDocType
                                            objMainDoc.Expenses.DistributionMethod = objremoteDoc.Expenses.DistributionMethod
                                            objMainDoc.Expenses.DistributionRule = objremoteDoc.Expenses.DistributionRule
                                            objMainDoc.Expenses.ExpenseCode = objremoteDoc.Expenses.ExpenseCode
                                            objMainDoc.Expenses.LastPurchasePrice = objremoteDoc.Expenses.LastPurchasePrice
                                            objMainDoc.Expenses.LineTotal = objremoteDoc.Expenses.LineTotal
                                            objMainDoc.Expenses.Remarks = objremoteDoc.Expenses.Remarks
                                            objMainDoc.Expenses.TaxCode = objremoteDoc.Expenses.TaxCode
                                            objMainDoc.Expenses.VatGroup = objremoteDoc.Expenses.VatGroup
                                        End If
                                    Next
                                    If aChoice <> "PQ" Then
                                        For intLoop As Integer = 0 To objremoteDoc.Lines.Count - 1
                                            If intLoop > 0 Then
                                                objMainDoc.Lines.Add()
                                            End If
                                            objMainDoc.Lines.SetCurrentLine(intLoop)
                                            objremoteDoc.Lines.SetCurrentLine(intLoop)

                                            '  If aChoice <> "GRPO" And objremoteDoc.Lines.BaseType <> -1 And getbaseDocument(objremoteDoc.Lines.BaseEntry, objremoteDoc.Lines.BaseEntry, objremoteDoc.Lines.BaseType, objMainCompany.CompanyName) > 0 Then 'objremoteDoc.Lines.BaseEntry aChoice <> "ARReser" And aChoice <> "GRPO" And aChoice <> "PO" Then
                                            If aChoice <> "GRPO" And objremoteDoc.Lines.BaseType <> -1 And getbaseDocument_BaseLine(objremoteDoc.Lines.BaseEntry, objremoteDoc.Lines.BaseEntry, objremoteDoc.Lines.BaseType, objMainCompany.CompanyName, objremoteDoc.Lines.BaseLine) <> 9999 And getbaseDocument(objremoteDoc.Lines.BaseEntry, objremoteDoc.Lines.BaseEntry, objremoteDoc.Lines.BaseType, objMainCompany.CompanyName) > 0 Then 'objremoteDoc.Lines.BaseEntry aChoice <> "ARReser" And aChoice <> "GRPO" And aChoice <> "PO" Then
                                                objMainDoc.Lines.BaseType = objremoteDoc.Lines.BaseType
                                                objMainDoc.Lines.BaseEntry = getbaseDocument(objremoteDoc.Lines.BaseEntry, objremoteDoc.Lines.BaseEntry, objremoteDoc.Lines.BaseType, objMainCompany.CompanyName) 'objremoteDoc.Lines.BaseEntry
                                                '  objMainDoc.Lines.BaseLine = objremoteDoc.Lines.BaseLine
                                                objMainDoc.Lines.BaseLine = getbaseDocument_BaseLine(objremoteDoc.Lines.BaseEntry, objremoteDoc.Lines.BaseEntry, objremoteDoc.Lines.BaseType, objMainCompany.CompanyName, objremoteDoc.Lines.BaseLine) ' objremoteDoc.Lines.BaseLine
                                                objMainDoc.Lines.ProjectCode = objremoteDoc.Lines.ProjectCode
                                                If objremoteDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service Then
                                                    objMainDoc.Lines.AccountCode = objremoteDoc.Lines.AccountCode
                                                    objMainDoc.Lines.ItemDescription = objremoteDoc.Lines.ItemDescription
                                                    objMainDoc.Lines.Currency = objremoteDoc.Lines.Currency
                                                    objMainDoc.Lines.UnitPrice = objremoteDoc.Lines.UnitPrice
                                                    objMainDoc.Lines.DiscountPercent = objremoteDoc.Lines.DiscountPercent
                                                    objMainDoc.Lines.Quantity = objremoteDoc.Lines.Quantity
                                                    objMainDoc.Lines.WarehouseCode = objremoteDoc.Lines.WarehouseCode
                                                    objMainDoc.Lines.RowTotalFC = objremoteDoc.Lines.RowTotalFC
                                                    objMainDoc.Lines.LineTotal = objremoteDoc.Lines.LineTotal
                                                    objMainDoc.Lines.ProjectCode = objremoteDoc.Lines.ProjectCode
                                                    objMainDoc.Lines.TaxCode = objremoteDoc.Lines.TaxCode
                                                    objMainDoc.Lines.VatGroup = objremoteDoc.Lines.VatGroup
                                                Else
                                                    objMainDoc.Lines.ItemCode = objremoteDoc.Lines.ItemCode
                                                    objMainDoc.Lines.Quantity = objremoteDoc.Lines.Quantity
                                                    objMainDoc.Lines.WarehouseCode = objremoteDoc.Lines.WarehouseCode
                                                    objMainDoc.Lines.ShipDate = objremoteDoc.Lines.ShipDate
                                                    objMainDoc.Lines.ShippingMethod = objremoteDoc.Lines.ShippingMethod
                                                    objMainDoc.Lines.Currency = objremoteDoc.Lines.Currency
                                                    objMainDoc.Lines.UnitPrice = objremoteDoc.Lines.UnitPrice
                                                    objMainDoc.Lines.DiscountPercent = objremoteDoc.Lines.DiscountPercent
                                                    objMainDoc.Lines.TaxCode = objremoteDoc.Lines.TaxCode
                                                    objMainDoc.Lines.VatGroup = objremoteDoc.Lines.VatGroup
                                                    objMainDoc.Lines.UoMEntry = objremoteDoc.Lines.UoMEntry
                                                    objMainDoc.Lines.RowTotalFC = objremoteDoc.Lines.RowTotalFC
                                                    objMainDoc.Lines.LineTotal = objremoteDoc.Lines.LineTotal
                                                End If
                                                objMainDoc.Lines.CostingCode = objremoteDoc.Lines.CostingCode
                                                objMainDoc.Lines.CostingCode2 = objremoteDoc.Lines.CostingCode2
                                                objMainDoc.Lines.CostingCode3 = objremoteDoc.Lines.CostingCode3
                                                objMainDoc.Lines.CostingCode4 = objremoteDoc.Lines.CostingCode4
                                                objMainDoc.Lines.CostingCode5 = objremoteDoc.Lines.CostingCode5

                                            Else
                                                If objremoteDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service Then
                                                    objMainDoc.Lines.AccountCode = objremoteDoc.Lines.AccountCode
                                                    objMainDoc.Lines.ItemDescription = objremoteDoc.Lines.ItemDescription

                                                    ' objMainDoc.Lines.UoMEntry = objremoteDoc.Lines.UoMEntry
                                                    objMainDoc.Lines.Currency = objremoteDoc.Lines.Currency
                                                    'objMainDoc.Lines.Price = objremoteDoc.Lines.Price
                                                    objMainDoc.Lines.UnitPrice = objremoteDoc.Lines.UnitPrice
                                                    objMainDoc.Lines.DiscountPercent = objremoteDoc.Lines.DiscountPercent
                                                    objMainDoc.Lines.Quantity = objremoteDoc.Lines.Quantity
                                                    objMainDoc.Lines.WarehouseCode = objremoteDoc.Lines.WarehouseCode
                                                    ' objMainDoc.Lines.ShipDate = objremoteDoc.Lines.ShipDate
                                                    ' objMainDoc.Lines.ShippingMethod = objremoteDoc.Lines.ShippingMethod
                                                    ' objMainDoc.Lines.PriceAfterVAT = objremoteDoc.Lines.PriceAfterVAT
                                                    'objMainDoc.Lines.LineTotal = objremoteDoc.Lines.RowTotalSC 
                                                    objMainDoc.Lines.RowTotalFC = objremoteDoc.Lines.RowTotalFC
                                                    objMainDoc.Lines.LineTotal = objremoteDoc.Lines.LineTotal

                                                    objMainDoc.Lines.ProjectCode = objremoteDoc.Lines.ProjectCode
                                                    objMainDoc.Lines.TaxCode = objremoteDoc.Lines.TaxCode
                                                    objMainDoc.Lines.VatGroup = objremoteDoc.Lines.VatGroup
                                                Else
                                                    objMainDoc.Lines.ItemCode = objremoteDoc.Lines.ItemCode
                                                    objMainDoc.Lines.ItemDescription = objremoteDoc.Lines.ItemDescription
                                                    objMainDoc.Lines.BarCode = objremoteDoc.Lines.BarCode
                                                    objMainDoc.Lines.VatGroup = objremoteDoc.Lines.VatGroup
                                                    objMainDoc.Lines.UoMEntry = objremoteDoc.Lines.UoMEntry
                                                    objMainDoc.Lines.Currency = objremoteDoc.Lines.Currency
                                                    objMainDoc.Lines.UnitPrice = objremoteDoc.Lines.UnitPrice
                                                    objMainDoc.Lines.DiscountPercent = objremoteDoc.Lines.DiscountPercent
                                                    objMainDoc.Lines.Quantity = objremoteDoc.Lines.Quantity
                                                    objMainDoc.Lines.WarehouseCode = objremoteDoc.Lines.WarehouseCode
                                                    objMainDoc.Lines.ShipDate = objremoteDoc.Lines.ShipDate
                                                    objMainDoc.Lines.ShippingMethod = objremoteDoc.Lines.ShippingMethod
                                                    objMainDoc.Lines.PriceAfterVAT = objremoteDoc.Lines.PriceAfterVAT
                                                    objMainDoc.Lines.RowTotalFC = objremoteDoc.Lines.RowTotalFC
                                                    objMainDoc.Lines.LineTotal = objremoteDoc.Lines.LineTotal
                                                    objMainDoc.Lines.ProjectCode = objremoteDoc.Lines.ProjectCode
                                                    objMainDoc.Lines.TaxCode = objremoteDoc.Lines.TaxCode
                                                End If
                                                objMainDoc.Lines.CostingCode = objremoteDoc.Lines.CostingCode
                                                objMainDoc.Lines.CostingCode2 = objremoteDoc.Lines.CostingCode2
                                                objMainDoc.Lines.CostingCode3 = objremoteDoc.Lines.CostingCode3
                                                objMainDoc.Lines.CostingCode4 = objremoteDoc.Lines.CostingCode4
                                                objMainDoc.Lines.CostingCode5 = objremoteDoc.Lines.CostingCode5
                                            End If

                                            For intLoop1 As Integer = 0 To objremoteDoc.Lines.UserFields.Fields.Count - 1
                                                Try
                                                    objMainDoc.Lines.UserFields.Fields.Item(intLoop1).Value = objremoteDoc.Lines.UserFields.Fields.Item(intLoop1).Value
                                                Catch ex As Exception

                                                End Try
                                            Next
                                            Dim objItem As SAPbobsCOM.Items
                                            objItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                                            If objItem.GetByKey(objremoteDoc.Lines.ItemCode) Then
                                                If objItem.ManageBatchNumbers = SAPbobsCOM.BoYesNoEnum.tYES Then
                                                    If aChoice <> "PQ" Then
                                                        For intSer As Integer = 0 To objremoteDoc.Lines.BatchNumbers.Count - 1
                                                            If intSer > 0 Then
                                                                objMainDoc.Lines.BatchNumbers.Add()
                                                                objMainDoc.Lines.BatchNumbers.SetCurrentLine(intSer)
                                                            End If
                                                            objremoteDoc.Lines.BatchNumbers.SetCurrentLine(intSer)
                                                            objMainDoc.Lines.BatchNumbers.AddmisionDate = objremoteDoc.Lines.BatchNumbers.AddmisionDate
                                                            objMainDoc.Lines.BatchNumbers.BaseLineNumber = objremoteDoc.Lines.BatchNumbers.BaseLineNumber
                                                            objMainDoc.Lines.BatchNumbers.BatchNumber = objremoteDoc.Lines.BatchNumbers.BatchNumber
                                                            objMainDoc.Lines.BatchNumbers.ExpiryDate = objremoteDoc.Lines.BatchNumbers.ExpiryDate
                                                            objMainDoc.Lines.BatchNumbers.InternalSerialNumber = objremoteDoc.Lines.BatchNumbers.InternalSerialNumber
                                                            objMainDoc.Lines.BatchNumbers.Location = objremoteDoc.Lines.BatchNumbers.Location
                                                            objMainDoc.Lines.BatchNumbers.ManufacturingDate = objremoteDoc.Lines.BatchNumbers.ManufacturingDate
                                                            objMainDoc.Lines.BatchNumbers.Notes = objremoteDoc.Lines.BatchNumbers.Notes
                                                            objMainDoc.Lines.BatchNumbers.Quantity = objremoteDoc.Lines.BatchNumbers.Quantity
                                                        Next
                                                    End If
                                                End If
                                                If objItem.ManageSerialNumbers = SAPbobsCOM.BoYesNoEnum.tYES Then
                                                    If aChoice <> "PQ" Then
                                                        For intSer As Integer = 0 To objremoteDoc.Lines.SerialNumbers.Count - 1
                                                            If intSer > 0 Then
                                                                objMainDoc.Lines.SerialNumbers.Add()
                                                                objMainDoc.Lines.SerialNumbers.SetCurrentLine(intSer)
                                                            End If
                                                            objremoteDoc.Lines.SerialNumbers.SetCurrentLine(intSer)
                                                            objMainDoc.Lines.SerialNumbers.BaseLineNumber = objremoteDoc.Lines.SerialNumbers.BaseLineNumber
                                                            '  objMainDoc.Lines.SerialNumbers.BatchID = objremoteDoc.Lines.SerialNumbers.BatchID
                                                            objMainDoc.Lines.SerialNumbers.ExpiryDate = objremoteDoc.Lines.SerialNumbers.ExpiryDate
                                                            objMainDoc.Lines.SerialNumbers.InternalSerialNumber = objremoteDoc.Lines.SerialNumbers.InternalSerialNumber
                                                            ' objMainDoc.Lines.SerialNumbers.Location = objremoteDoc.Lines.SerialNumbers.Location
                                                            objMainDoc.Lines.SerialNumbers.ManufactureDate = objremoteDoc.Lines.SerialNumbers.ManufactureDate
                                                            objMainDoc.Lines.SerialNumbers.ManufacturerSerialNumber = objremoteDoc.Lines.SerialNumbers.ManufacturerSerialNumber
                                                            objMainDoc.Lines.SerialNumbers.Notes = objremoteDoc.Lines.SerialNumbers.Notes
                                                            objMainDoc.Lines.SerialNumbers.ReceptionDate = objremoteDoc.Lines.SerialNumbers.ReceptionDate
                                                            ' objMainDoc.Lines.SerialNumbers.SystemSerialNumber = objremoteDoc.Lines.SerialNumbers.SystemSerialNumber
                                                        Next
                                                    End If
                                                End If
                                            End If
                                        Next
                                    Else
                                    End If
                                    If objMainDoc.Add <> 0 Then
                                        sPath = strErrorLogPath
                                        WriteErrorlog("Failed to Export " & stMessagetext & " Document" & objremoteDoc.DocNum & ": Error :" & objRemoteCompany.GetLastErrorDescription & "  : Branch : " & objRemoteCompany.CompanyName, sPath)
                                    Else
                                        Dim strDocNum As String
                                        objRemoteCompany.GetNewObjectCode(strDocNum)
                                        Dim oInvoice As SAPbobsCOM.Documents
                                        Dim intInvoiceDocentry As String
                                        oInvoice = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                                        If objMainDoc.GetByKey(Convert.ToInt64(strDocNum)) Then
                                            strDocNum = objMainDoc.DocNum
                                            intInvoiceDocentry = objMainDoc.DocEntry
                                        End If
                                        sPath = strErrorLogPath
                                        WriteErrorlog(stMessagetext & " Created Successfully. DocNum : " & strDocNum, sPath)
                                        Dim otempRec As SAPbobsCOM.Recordset
                                        otempRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        objremoteDoc.GetByKey(intRemoteDocEntry)
                                        objremoteDoc.UserFields.Fields.Item("U_Export").Value = "Y"
                                        objremoteDoc.Update()
                                        Dim oTe As SAPbobsCOM.Recordset
                                        oTe = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        If aJournalNo <> "" Then
                                            oTe.DoQuery("Update OJDT set U_Export='Y' where TransId=" & aJournalNo)
                                            If Not IsNothing(oTe) Then
                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oTe)
                                            End If
                                        End If

                                    End If
                                End If
                            End If
                        End If
                    End If
                Catch ex As Exception
                    WriteErrorlog("Failed to create Documents Error  :" & ex.Message & " --> " & ex.StackTrace, sPath)
                End Try
                objremoteRec.MoveNext()
            Next

        Catch ex As Exception
            WriteErrorlog("Current Document : " & aChoice & " " & aJournalNo, sPath)
            WriteErrorlog(ex.Message & " --> " & ex.StackTrace, sPath)
        Finally
            If Not IsNothing(objremoteDoc) Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objremoteDoc)
                GC.Collect()
            End If
            If Not IsNothing(objMainDoc) Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objMainDoc)
                GC.Collect()
            End If
            If Not IsNothing(objMainDoc1) Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objMainDoc1)
                GC.Collect()
            End If
            'If Not IsNothing(objremoteRec) Then
            '    System.Runtime.InteropServices.Marshal.ReleaseComObject(objremoteRec)
            'End If
            ' System.Runtime.InteropServices.Marshal.ReleaseComObject(objMainDoc)
        End Try
    End Sub

    Private Function getbaseDocument_BaseLine(ByVal aDocEntry As Integer, ByVal aDocNum As Integer, ByVal aBasetype As Double, ByVal aBranch As String, ByVal aLineID As Integer) As Integer
        Dim str As String
        Dim str2 As String
        Dim str5 As String
        Dim businessObject As SAPbobsCOM.Recordset = DirectCast(SubMain.objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim recordset2 As SAPbobsCOM.Recordset = DirectCast(SubMain.objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim recordset3 As SAPbobsCOM.Recordset = DirectCast(SubMain.objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Select Case aBasetype
            Case 22
                str5 = "OPOR"
                str = "POR1"
                Exit Select
            Case 21
                str5 = "ORPD"
                str = "RPD1"
                Exit Select
            Case 18
                str5 = "OPCH"
                str = "PCH1"
                Exit Select
            Case 20
                str5 = "OPDN"
                str = "PDN1"
                Exit Select
            Case 19
                str5 = "ORPC"
                str = "RPC1"
                Exit Select
            Case 100
                str5 = "OWOR"
                str = "WOR1"
                Exit Select
            Case 202
                str5 = "OWOR"
                str = "WOR1"
                Exit Select
            Case 30
                str5 = "OJDT"
                str = "JDT1"
                Exit Select
            Case 13
                str5 = "OINV"
                str = "INV1"
                Exit Select
            Case 46
                str5 = "OVPM"
                str = "VPM1"
                Exit Select
        End Select
        If ((aBasetype = "-1") Or (str5 = "")) Then
            Return 9999
        End If
        If (aBasetype = 46) Then
            str2 = ("Select BaseRef from OJDT where TransID=" & aDocEntry)
            businessObject.DoQuery(str2)
            If (businessObject.RecordCount > 0) Then
                str5 = "OVPM"
                aDocEntry = (businessObject.Fields.Item(0).Value)
                str2 = ("Select TransId from OJDT where TransType=46 and   BaseRef=" & aDocEntry)
                recordset2.DoQuery(str2)
                If (recordset2.RecordCount > 0) Then
                    Return recordset2.Fields.Item(0).Value
                End If
                Return 9999
            End If
        End If
        If (str5 <> "OJDT") Then
            str2 = "Select VisOrder from " & str & " where DocEntry=" & aDocEntry & " and LineNum=" & aLineID
        Else
            str2 = "Select TransId from " & str & " where U_Branch='" & aBranch & "' and   U_BaseEntry=" & aDocEntry
        End If
        businessObject.DoQuery(str2)
        If (businessObject.RecordCount > 0) Then
            Return businessObject.Fields.Item(0).Value
        End If
        Return 9999
    End Function
    Private Sub ExportDocuments_Canceled(ByVal aChoice As String, Optional ByVal adocentry As String = "", Optional ByVal aJournalNo As String = "")
        Dim conn As New SqlConnection()
        Dim objMainDoc, objremoteDoc, objMainDoc1 As SAPbobsCOM.Documents
        Dim intRemoteDocEntry As Integer
        Dim strPath, strFilename, strMessage, strTableName, stMessagetext As String
        Dim strFileName1 As String
        Dim objremoteRec As SAPbobsCOM.Recordset
        Dim strCondition As String = "1=1"
        strTableName = ""
        Dim strCode As String = ""
        Dim blnCanceled As Boolean = False
        Try
            'objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
            'objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
            Select Case aChoice
                Case "PO"
                    strCode = "22"
                    strTableName = "OPOR"
                    stMessagetext = "Purchase Document"
                    objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                    objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                Case "PQ"
                    strTableName = ""
                    stMessagetext = "Purchase Quotation"
                    objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseQuotations)
                    objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseQuotations)
                Case "GRPO"
                    strCode = "20"
                    strTableName = "OPDN"
                    stMessagetext = "Goods Receipt Purchase Order"
                    objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
                    objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
                Case "GR"
                    strCode = "21"
                    strTableName = "ORPD"
                    stMessagetext = "Goods Return"
                    objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseReturns)
                    objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseReturns)
                Case "APInvoice"
                    strCode = "18"
                    strTableName = "OPCH"
                    stMessagetext = "A/P Invoice"
                    objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                    objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                    '   blnCanceled = True
                Case "APCR"
                    strCode = "19"
                    strTableName = "ORPC"
                    stMessagetext = "A/P Credit Memo"
                    objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
                    objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
                Case "ARReser"
                    'isIns]
                    strCode = "171"
                    strCondition = " IsIns='Y'"
                    strTableName = "OINV"
                    stMessagetext = "A/R Invoice"
                    objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                    objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                    blnCanceled = True
                Case "Delivery"
                    strCode = "15"
                    strTableName = "ODLN"
                    stMessagetext = "Delivery"
                    objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                    objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                Case "Invoice"
                    strCode = "13"
                    strCondition = " IsIns='N'"
                    strTableName = "OINV"
                    stMessagetext = "A/R Invoice"
                    blnCanceled = True
                    objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                    objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                Case "CreditNote"
                    strCode = "14"
                    strTableName = "ORIN"
                    stMessagetext = "A/R Credit Note"
                    objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                    objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
            End Select
            If strTableName = "" Then
                If Not IsNothing(objremoteDoc) Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objremoteDoc)
                End If
                If Not IsNothing(objMainDoc) Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objMainDoc)
                End If
                Exit Sub
            End If
            '  blnCanceled = True
            objremoteRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objremoteRec.DoQuery("SELECT * from [@Z_DBSYN] where isnull(U_Active,'N')='Y' and Code='" & strCode & "'")
            Dim dtFrom, dtTo As Date
            If objremoteRec.RecordCount <= 0 Then
                If Not IsNothing(objremoteDoc) Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objremoteDoc)
                End If
                If Not IsNothing(objMainDoc) Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objMainDoc)
                End If
                If Not IsNothing(objremoteRec) Then
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objremoteRec)
                End If
                Exit Sub
            Else
                dtFrom = objremoteRec.Fields.Item("U_FrmDate").Value
                dtTo = objremoteRec.Fields.Item("U_ToDate").Value
            End If
            Dim dtString As String
            If Year(dtFrom) = 1899 Then
                dtString = " 1=1"
            Else
                dtString = " T0.DocDate >='" & dtFrom.ToString("yyyy-MM-dd") & "'"
            End If
            If Year(dtTo) = 1899 Then
                dtString = dtString & " and  1=1"
            Else
                dtString = dtString & " and T0.DocDate <='" & dtTo.ToString("yyyy-MM-dd") & "'"
            End If
            'Dim s As String = "Select T0.DocEntry from " & strTableName & " T0 inner Join OCRD T1 on T1.CardCode=T0.CardCode  where " & strCondition & " and isnull(T1.U_sequencetype,'N')='P1' and isnull(U_Export,'N')='N' order by T0.DocDate,T0.DocEntry"
            Dim s As String
            If adocentry <> "" Then
                s = "Select T0.DocEntry from " & strTableName & " T0 inner Join OCRD T1 on T1.CardCode=T0.CardCode  where T0.DocEntry=" & adocentry & " and " & strCondition & " and T0.Canceled='C' and isnull(T1.U_sequencetype,'N')='P1' and ( " & dtString & ") order by T0.DocDate,T0.DocEntry"
            Else
                s = "Select T0.DocEntry from " & strTableName & " T0 inner Join OCRD T1 on T1.CardCode=T0.CardCode  where  " & strCondition & " and T0.Canceled='C' and isnull(T1.U_sequencetype,'N')='P1' and ( " & dtString & ") order by T0.DocDate,T0.DocEntry"
            End If
            'Dim s As String = "Select T0.DocEntry from " & strTableName & " T0 inner Join OCRD T1 on T1.CardCode=T0.CardCode  where " & strCondition & "  and isnull(T1.U_sequencetype,'N')='P1' and isnull(U_Export,'N')='N' and ( " & dtString & ") order by T0.DocDate,T0.DocEntry"
            objremoteRec.DoQuery(s)
            ' Dim objMainDoc1 As SAPbobsCOM.Documents
            For intRemoteLoop As Integer = 0 To objremoteRec.RecordCount - 1
                Try
                    If objremoteDoc.GetByKey(Convert.ToInt32(objremoteRec.Fields.Item(0).Value)) Then
                        If ValidateDocDate(objremoteDoc.DocEntry, strCode, objremoteDoc.DocDate) = False Then

                        Else
                            If checkDuplicateInvoice(objremoteDoc.DocEntry, objremoteDoc.DocNum, objMainCompany.CompanyName, strTableName) = True Then
                                objremoteDoc.UserFields.Fields.Item("U_Export").Value = "Y"
                                objremoteDoc.Update()
                            Else
                                intRemoteDocEntry = objremoteDoc.DocEntry
                                '    UpdateExchangerate_Main(objremoteDoc.DocCurrency, objremoteDoc.DocDate, objMainCompany)
                                Select Case aChoice
                                    Case "PO"
                                        strTableName = "OPOR"
                                        'objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                                        objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                                        objMainDoc1 = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                                    Case "PQ"
                                        strTableName = "OPQT"
                                        'objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                                        objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseQuotations)
                                        objMainDoc1 = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseQuotations)
                                    Case "GRPO"
                                        strTableName = "OPDN"
                                        ' objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
                                        objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
                                        objMainDoc1 = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
                                    Case "GR"
                                        strTableName = "ORPD"
                                        ' objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
                                        objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseReturns)
                                        objMainDoc1 = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseReturns)
                                    Case "APInvoice"
                                        strTableName = "OPCH"
                                        ' objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                                        objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                                        objMainDoc1 = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseInvoices)
                                    Case "APCR"
                                        strTableName = "ORPC"
                                        'objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
                                        objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)
                                        objMainDoc1 = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseCreditNotes)

                                    Case "ARReser"
                                        'isIns]
                                        strCondition = " IsIns='Y'"
                                        strTableName = "OINV"
                                        stMessagetext = "A/R Invoice"
                                        ' objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                                        objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                                        objMainDoc1 = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                                    Case "Delivery"
                                        strTableName = "ODLN"
                                        stMessagetext = "Delivery"
                                        ' objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                                        objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                                        objMainDoc1 = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)
                                    Case "Invoice"
                                        strCondition = " IsIns='N'"
                                        strTableName = "OINV"
                                        stMessagetext = "A/R Invoice"
                                        ' objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                                        objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                                        objMainDoc1 = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                                    Case "CreditNote"
                                        strTableName = "ORIN"
                                        stMessagetext = "A/R Credit Note"
                                        ' objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                                        objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                                        objMainDoc1 = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                                End Select

                                If objremoteDoc.Cancelled = SAPbobsCOM.BoYesNoEnum.tYES Then '  CheckCanceledInvoice(objremoteDoc.DocEntry, objremoteDoc.DocNum, objMainCompany.CompanyName, strTableName) <> 0 Then
                                    Dim inTc As Integer = CheckCanceledInvoice(objremoteDoc.DocEntry, objremoteDoc.DocNum, objMainCompany.CompanyName, strTableName)
                                    If objMainDoc1.GetByKey(inTc) Then
                                        Dim otest As SAPbobsCOM.Documents
                                        If 1 = 1 Then
                                            objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                                            If objMainDoc1.Cancelled = SAPbobsCOM.BoYesNoEnum.tNO Then
                                                objMainDoc1.SaveXML("C:\Test.xml")
                                                objMainDoc = objMainDoc1.CreateCancellationDocument()
                                                objMainDoc.SaveXML("C:\Test.xml")
                                                objMainDoc.DocNum = objremoteDoc.DocNum
                                                objMainDoc.HandWritten = SAPbobsCOM.BoYesNoEnum.tYES
                                                objMainDoc.Series = -1
                                                inTc = objMainDoc.Add()
                                                If inTc = 0 Then
                                                    Dim strDocNum As String
                                                    objRemoteCompany.GetNewObjectCode(strDocNum)
                                                    Dim oInvoice As SAPbobsCOM.Documents
                                                    Dim intInvoiceDocentry As String
                                                    oInvoice = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                                                    If objMainDoc.GetByKey(Convert.ToInt64(strDocNum)) Then
                                                        strDocNum = objMainDoc.DocNum
                                                        intInvoiceDocentry = objMainDoc.DocEntry
                                                    End If
                                                    sPath = strErrorLogPath
                                                    WriteErrorlog(stMessagetext & " Created Successfully. DocNum : " & strDocNum, sPath)
                                                    Dim otempRec As SAPbobsCOM.Recordset
                                                    otempRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                    objremoteDoc.GetByKey(intRemoteDocEntry)
                                                    objremoteDoc.UserFields.Fields.Item("U_Export").Value = "Y"
                                                    objremoteDoc.Update()
                                                    Dim oTe As SAPbobsCOM.Recordset
                                                    oTe = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                                    If aJournalNo <> "" Then
                                                        oTe.DoQuery("Update OJDT set U_Export='Y' where TransId=" & aJournalNo)
                                                    End If
                                                    If Not IsNothing(oInvoice) Then
                                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oInvoice)
                                                    End If
                                                    If Not IsNothing(oTe) Then
                                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oTe)
                                                    End If
                                                Else
                                                    WriteErrorlog(stMessagetext & "  DocNum : " & objremoteDoc.DocNum & " Not exported : Error  : " & objRemoteCompany.GetLastErrorDescription, sPath)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else
                                    UpdateExchangerate_Main(objremoteDoc.DocCurrency, objremoteDoc.DocDate, objMainCompany)
                                    objMainDoc.CardCode = objremoteDoc.CardCode
                                    objMainDoc.DocDate = objremoteDoc.DocDate
                                    objMainDoc.DocDueDate = objremoteDoc.DocDueDate
                                    objMainDoc.HandWritten = SAPbobsCOM.BoYesNoEnum.tYES
                                    objMainDoc.Series = -1
                                    objMainDoc.DocNum = objremoteDoc.DocNum
                                    objMainDoc.NumAtCard = objremoteDoc.NumAtCard
                                    objMainDoc.Comments = objremoteDoc.Comments
                                    objMainDoc.DiscountPercent = objremoteDoc.DiscountPercent
                                    objMainDoc.DocCurrency = objremoteDoc.DocCurrency
                                    objMainDoc.ShipToCode = objremoteDoc.ShipToCode
                                    objMainDoc.SalesPersonCode = objremoteDoc.SalesPersonCode
                                    objMainDoc.TaxDate = objremoteDoc.TaxDate
                                    objMainDoc.PaymentGroupCode = objremoteDoc.PaymentGroupCode
                                    objMainDoc.PaymentMethod = objremoteDoc.PaymentMethod
                                    objMainDoc.Address = objremoteDoc.Address
                                    objMainDoc.Address2 = objremoteDoc.Address2
                                    objMainDoc.AgentCode = objremoteDoc.AgentCode
                                    objMainDoc.BPChannelCode = objremoteDoc.BPChannelCode
                                    objMainDoc.BPChannelContact = objremoteDoc.BPChannelContact
                                    '  objMainDoc.ContactPersonCode = objremoteDoc.ContactPersonCode
                                    objMainDoc.ReserveInvoice = objremoteDoc.ReserveInvoice
                                    If objremoteDoc.FederalTaxID <> "" Then
                                        objMainDoc.FederalTaxID = objremoteDoc.FederalTaxID
                                    End If
                                    ' objMainDoc.DocRate = objremoteDoc.DocRate
                                    objMainDoc.DocType = objremoteDoc.DocType
                                    '  objMainDoc.TransportationCode = objremoteDoc.TransportationCode
                                    If objremoteDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES Then
                                        objMainDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES
                                        objMainDoc.RoundingDiffAmount = objremoteDoc.RoundingDiffAmount
                                    Else
                                        objMainDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tNO
                                    End If
                                    For intLoop As Integer = 0 To objremoteDoc.UserFields.Fields.Count - 1
                                        Try
                                            objMainDoc.UserFields.Fields.Item(intLoop).Value = objremoteDoc.UserFields.Fields.Item(intLoop).Value
                                        Catch ex As Exception

                                        End Try
                                    Next
                                    objMainDoc.UserFields.Fields.Item("U_Import").Value = "Y"
                                    objMainDoc.UserFields.Fields.Item("U_BaseEntry").Value = objremoteDoc.DocEntry
                                    objMainDoc.UserFields.Fields.Item("U_BaseNum").Value = objremoteDoc.DocNum
                                    objMainDoc.UserFields.Fields.Item("U_Branch").Value = objMainCompany.CompanyName
                                    For IntExp As Integer = 0 To objremoteDoc.Expenses.Count - 1
                                        If objremoteDoc.Expenses.LineTotal > 0 Then
                                            If IntExp > 0 Then
                                                objMainDoc.Expenses.Add()
                                                objMainDoc.Expenses.SetCurrentLine(IntExp)
                                            End If
                                            objremoteDoc.Expenses.SetCurrentLine(IntExp)
                                            objMainDoc.Expenses.BaseDocEntry = objremoteDoc.Expenses.BaseDocEntry
                                            objMainDoc.Expenses.BaseDocLine = objremoteDoc.Expenses.BaseDocLine
                                            objMainDoc.Expenses.BaseDocType = objremoteDoc.Expenses.BaseDocType
                                            objMainDoc.Expenses.DistributionMethod = objremoteDoc.Expenses.DistributionMethod
                                            objMainDoc.Expenses.DistributionRule = objremoteDoc.Expenses.DistributionRule
                                            objMainDoc.Expenses.ExpenseCode = objremoteDoc.Expenses.ExpenseCode
                                            objMainDoc.Expenses.LastPurchasePrice = objremoteDoc.Expenses.LastPurchasePrice
                                            objMainDoc.Expenses.LineTotal = objremoteDoc.Expenses.LineTotal
                                            objMainDoc.Expenses.Remarks = objremoteDoc.Expenses.Remarks
                                            objMainDoc.Expenses.TaxCode = objremoteDoc.Expenses.TaxCode
                                            objMainDoc.Expenses.VatGroup = objremoteDoc.Expenses.VatGroup
                                        End If
                                    Next
                                    If aChoice <> "PQ" Then
                                        For intLoop As Integer = 0 To objremoteDoc.Lines.Count - 1
                                            If intLoop > 0 Then
                                                objMainDoc.Lines.Add()
                                            End If
                                            objMainDoc.Lines.SetCurrentLine(intLoop)
                                            objremoteDoc.Lines.SetCurrentLine(intLoop)

                                            If objremoteDoc.Lines.BaseType <> -1 And getbaseDocument(objremoteDoc.Lines.BaseEntry, objremoteDoc.Lines.BaseEntry, objremoteDoc.Lines.BaseType, objMainCompany.CompanyName) > 0 Then 'objremoteDoc.Lines.BaseEntry aChoice <> "ARReser" And aChoice <> "GRPO" And aChoice <> "PO" Then
                                                objMainDoc.Lines.BaseType = objremoteDoc.Lines.BaseType
                                                objMainDoc.Lines.BaseEntry = getbaseDocument(objremoteDoc.Lines.BaseEntry, objremoteDoc.Lines.BaseEntry, objremoteDoc.Lines.BaseType, objMainCompany.CompanyName) 'objremoteDoc.Lines.BaseEntry
                                                objMainDoc.Lines.BaseLine = objremoteDoc.Lines.BaseLine
                                                objMainDoc.Lines.ProjectCode = objremoteDoc.Lines.ProjectCode
                                                objMainDoc.Lines.Quantity = objremoteDoc.Lines.Quantity
                                                objMainDoc.Lines.WarehouseCode = objremoteDoc.Lines.WarehouseCode
                                            Else
                                                If objremoteDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service Then
                                                    objMainDoc.Lines.AccountCode = objremoteDoc.Lines.AccountCode
                                                    objMainDoc.Lines.ItemDescription = objremoteDoc.Lines.ItemDescription
                                                Else
                                                    objMainDoc.Lines.ItemCode = objremoteDoc.Lines.ItemCode
                                                    objMainDoc.Lines.ItemDescription = objremoteDoc.Lines.ItemDescription
                                                    objMainDoc.Lines.BarCode = objremoteDoc.Lines.BarCode
                                                End If
                                                ' objMainDoc.Lines.UnitPrice = objremoteDoc.Lines.UnitPrice

                                                objMainDoc.Lines.VatGroup = objremoteDoc.Lines.VatGroup
                                                objMainDoc.Lines.UoMEntry = objremoteDoc.Lines.UoMEntry
                                                objMainDoc.Lines.Currency = objremoteDoc.Lines.Currency
                                                objMainDoc.Lines.Price = objremoteDoc.Lines.Price
                                                objMainDoc.Lines.DiscountPercent = objremoteDoc.Lines.DiscountPercent
                                                objMainDoc.Lines.Quantity = objremoteDoc.Lines.Quantity
                                                objMainDoc.Lines.WarehouseCode = objremoteDoc.Lines.WarehouseCode
                                                ' objMainDoc.Lines.PriceAfterVAT = objremoteDoc.Lines.PriceAfterVAT
                                                'objMainDoc.Lines.LineTotal = objremoteDoc.Lines.LineTotal
                                                objMainDoc.Lines.ProjectCode = objremoteDoc.Lines.ProjectCode

                                                objMainDoc.Lines.TaxCode = objremoteDoc.Lines.TaxCode
                                            End If

                                            For intLoop1 As Integer = 0 To objremoteDoc.Lines.UserFields.Fields.Count - 1
                                                Try
                                                    objMainDoc.Lines.UserFields.Fields.Item(intLoop1).Value = objremoteDoc.Lines.UserFields.Fields.Item(intLoop1).Value
                                                Catch ex As Exception

                                                End Try
                                            Next

                                            Dim objItem As SAPbobsCOM.Items
                                            objItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                                            If objItem.GetByKey(objremoteDoc.Lines.ItemCode) Then
                                                If objItem.ManageBatchNumbers = SAPbobsCOM.BoYesNoEnum.tYES Then
                                                    If aChoice <> "PQ" Then
                                                        For intSer As Integer = 0 To objremoteDoc.Lines.BatchNumbers.Count - 1
                                                            If intSer > 0 Then
                                                                objMainDoc.Lines.BatchNumbers.Add()
                                                                objMainDoc.Lines.BatchNumbers.SetCurrentLine(intSer)
                                                            End If
                                                            objremoteDoc.Lines.BatchNumbers.SetCurrentLine(intSer)
                                                            objMainDoc.Lines.BatchNumbers.AddmisionDate = objremoteDoc.Lines.BatchNumbers.AddmisionDate
                                                            objMainDoc.Lines.BatchNumbers.BaseLineNumber = objremoteDoc.Lines.BatchNumbers.BaseLineNumber
                                                            objMainDoc.Lines.BatchNumbers.BatchNumber = objremoteDoc.Lines.BatchNumbers.BatchNumber
                                                            objMainDoc.Lines.BatchNumbers.ExpiryDate = objremoteDoc.Lines.BatchNumbers.ExpiryDate
                                                            objMainDoc.Lines.BatchNumbers.InternalSerialNumber = objremoteDoc.Lines.BatchNumbers.InternalSerialNumber
                                                            objMainDoc.Lines.BatchNumbers.Location = objremoteDoc.Lines.BatchNumbers.Location
                                                            objMainDoc.Lines.BatchNumbers.ManufacturingDate = objremoteDoc.Lines.BatchNumbers.ManufacturingDate
                                                            objMainDoc.Lines.BatchNumbers.Notes = objremoteDoc.Lines.BatchNumbers.Notes
                                                            objMainDoc.Lines.BatchNumbers.Quantity = objremoteDoc.Lines.BatchNumbers.Quantity
                                                        Next
                                                    End If
                                                End If
                                                If objItem.ManageSerialNumbers = SAPbobsCOM.BoYesNoEnum.tYES Then
                                                    If aChoice <> "PQ" Then
                                                        For intSer As Integer = 0 To objremoteDoc.Lines.SerialNumbers.Count - 1
                                                            If intSer > 0 Then
                                                                objMainDoc.Lines.SerialNumbers.Add()
                                                                objMainDoc.Lines.SerialNumbers.SetCurrentLine(intSer)
                                                            End If
                                                            objremoteDoc.Lines.SerialNumbers.SetCurrentLine(intSer)
                                                            objMainDoc.Lines.SerialNumbers.BaseLineNumber = objremoteDoc.Lines.SerialNumbers.BaseLineNumber
                                                            '  objMainDoc.Lines.SerialNumbers.BatchID = objremoteDoc.Lines.SerialNumbers.BatchID
                                                            objMainDoc.Lines.SerialNumbers.ExpiryDate = objremoteDoc.Lines.SerialNumbers.ExpiryDate
                                                            objMainDoc.Lines.SerialNumbers.InternalSerialNumber = objremoteDoc.Lines.SerialNumbers.InternalSerialNumber
                                                            ' objMainDoc.Lines.SerialNumbers.Location = objremoteDoc.Lines.SerialNumbers.Location
                                                            objMainDoc.Lines.SerialNumbers.ManufactureDate = objremoteDoc.Lines.SerialNumbers.ManufactureDate
                                                            objMainDoc.Lines.SerialNumbers.ManufacturerSerialNumber = objremoteDoc.Lines.SerialNumbers.ManufacturerSerialNumber
                                                            objMainDoc.Lines.SerialNumbers.Notes = objremoteDoc.Lines.SerialNumbers.Notes
                                                            objMainDoc.Lines.SerialNumbers.ReceptionDate = objremoteDoc.Lines.SerialNumbers.ReceptionDate
                                                            ' objMainDoc.Lines.SerialNumbers.SystemSerialNumber = objremoteDoc.Lines.SerialNumbers.SystemSerialNumber
                                                        Next
                                                    End If
                                                End If
                                            End If

                                        Next
                                    Else

                                    End If
                                    If objMainDoc.Add <> 0 Then
                                        sPath = strErrorLogPath
                                        WriteErrorlog("Failed to Export " & stMessagetext & " docuemnt :" & objremoteDoc.DocNum & ": Error :" & objRemoteCompany.GetLastErrorDescription & "  : Branch : " & objRemoteCompany.CompanyName, sPath)
                                    Else
                                        Dim strDocNum As String
                                        objRemoteCompany.GetNewObjectCode(strDocNum)
                                        Dim oInvoice As SAPbobsCOM.Documents
                                        Dim intInvoiceDocentry As String
                                        oInvoice = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                                        If objMainDoc.GetByKey(Convert.ToInt64(strDocNum)) Then
                                            strDocNum = objMainDoc.DocNum
                                            intInvoiceDocentry = objMainDoc.DocEntry
                                        End If
                                        sPath = strErrorLogPath
                                        WriteErrorlog(stMessagetext & " Created Successfully. DocNum : " & strDocNum, sPath)
                                        Dim otempRec As SAPbobsCOM.Recordset
                                        otempRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        objremoteDoc.GetByKey(intRemoteDocEntry)
                                        objremoteDoc.UserFields.Fields.Item("U_Export").Value = "Y"
                                        objremoteDoc.Update()
                                    End If
                                End If
                            End If
                        End If
                    End If
                Catch ex As Exception
                    WriteErrorlog("Failed to create Incomint payment Error  :" & ex.Message, sPath)
                End Try
                objremoteRec.MoveNext()
            Next
        Catch ex As Exception
            WriteErrorlog(ex.Message, sPath)
        Finally
            If Not IsNothing(objremoteDoc) Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objremoteDoc)
                GC.Collect()
            End If
            If Not IsNothing(objMainDoc) Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objMainDoc)
                GC.Collect()
            End If
            If Not IsNothing(objMainDoc1) Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objMainDoc1)
                GC.Collect()
            End If
            'If Not IsNothing(objremoteRec) Then
            '    System.Runtime.InteropServices.Marshal.ReleaseComObject(objremoteRec)
            'End If
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(objMainDoc)
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(objremoteRec)
        End Try

    End Sub


    Private Sub add_Landing_Cost(ByVal strBDocEntry As String) 'Source Landing No
        Try
            'Source Company
            Dim svrSLandedCost As SAPbobsCOM.LandedCostsService = objMainCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.LandedCostsService)
            Dim oSLandedCostParams As SAPbobsCOM.LandedCostParams = svrSLandedCost.GetDataInterface(SAPbobsCOM.LandedCostsServiceDataInterfaces.lcsLandedCostParams)
            Dim oSLandedCost As SAPbobsCOM.LandedCost = svrSLandedCost.GetDataInterface(SAPbobsCOM.LandedCostsServiceDataInterfaces.lcsLandedCost)

            'Destination Company
            Dim svrLandedCost As SAPbobsCOM.LandedCostsService = objRemoteCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.LandedCostsService)
            Dim oLandedCostParams As SAPbobsCOM.LandedCostParams = svrLandedCost.GetDataInterface(SAPbobsCOM.LandedCostsServiceDataInterfaces.lcsLandedCostParams)
            Dim oLandedCost As SAPbobsCOM.LandedCost = svrLandedCost.GetDataInterface(SAPbobsCOM.LandedCostsServiceDataInterfaces.lcsLandedCost)
            Dim oLandedCost_ItemLine As SAPbobsCOM.LandedCost_ItemLine
            Dim oLandedCost_CostLine As SAPbobsCOM.LandedCost_CostLine

            'Operate on the landed cost 
            oSLandedCostParams.LandedCostNumber = strBDocEntry
            'Get the landed cost 
            Try
                oSLandedCost = svrSLandedCost.GetLandedCost(oSLandedCostParams)
                oLandedCost.VendorCode = oSLandedCost.VendorCode
                oLandedCost.VendorName = oSLandedCost.VendorName
                oLandedCost.Broker = oSLandedCost.Broker
                oLandedCost.BrokerName = oSLandedCost.BrokerName
                oLandedCost.DocumentCurrency = oSLandedCost.DocumentCurrency
                oLandedCost.DocumentRate = oSLandedCost.DocumentRate
                oLandedCost.ActualCustoms = oSLandedCost.ActualCustoms
                oLandedCost.PostingDate = oSLandedCost.PostingDate
                oLandedCost.DueDate = oSLandedCost.DueDate
                oLandedCost.Reference = oSLandedCost.Reference
                oLandedCost.FileNumber = oSLandedCost.FileNumber
                oLandedCost.ActualCustoms = oSLandedCost.ActualCustoms
                oLandedCost.Remarks = oSLandedCost.Remarks
                oLandedCost.JournalRemarks = oSLandedCost.JournalRemarks
                For Each row As SAPbobsCOM.LandedCost_ItemLine In oSLandedCost.LandedCost_ItemLines
                    oLandedCost_ItemLine = oLandedCost.LandedCost_ItemLines.Add

                    oLandedCost_ItemLine.AllocatedUnitCostsLineTotal = row.AllocatedUnitCostsLineTotal
                    oLandedCost_ItemLine.AllocatedUnitCostsLineTotalFC = row.AllocatedUnitCostsLineTotalFC
                    oLandedCost_ItemLine.AutomaticExpenditure = row.AutomaticExpenditure
                    oLandedCost_ItemLine.BaseDocumentPrice = row.BaseDocumentPrice
                    oLandedCost_ItemLine.BaseDocumentType = row.BaseDocumentType
                    oLandedCost_ItemLine.BaseEntry = GetBaseDocEntry_LandedCost(row.BaseEntry, "GRPO", objMainCompany.CompanyName) ' row.BaseEntry 
                    '   MsgBox(row.BaseLine)
                    oLandedCost_ItemLine.BaseLine = row.BaseLine
                    oLandedCost_ItemLine.Quantity = row.Quantity
                    oLandedCost_ItemLine.BlockNumber = row.BlockNumber
                    oLandedCost_ItemLine.CCDNumber = row.CCDNumber
                    oLandedCost_ItemLine.Currency = row.Currency
                    oLandedCost_ItemLine.CustomsAffectStock = row.CustomsAffectStock
                    oLandedCost_ItemLine.CustomsCost = row.CustomsCost
                    oLandedCost_ItemLine.CustomsCostFC = row.CustomsCostFC
                    oLandedCost_ItemLine.CustomsValue = row.CustomsValue
                    oLandedCost_ItemLine.CustomsValueFC = row.CustomsValueFC
                    oLandedCost_ItemLine.CustomsVat = row.CustomsVat
                    oLandedCost_ItemLine.CustomsVatAffectStock = row.CustomsVatAffectStock
                    oLandedCost_ItemLine.CustomsVatFC = row.CustomsVatFC
                    oLandedCost_ItemLine.ExciseAffectStock = row.ExciseAffectStock
                    oLandedCost_ItemLine.ExciseSum = row.ExciseSum
                    oLandedCost_ItemLine.ExciseSumFC = row.ExciseSumFC
                    oLandedCost_ItemLine.Expenditure = row.Expenditure
                    oLandedCost_ItemLine.ExpenditureFC = row.ExpenditureFC
                    oLandedCost_ItemLine.InventoryValuation = row.InventoryValuation
                    oLandedCost_ItemLine.OriginLine = row.OriginLine
                    oLandedCost_ItemLine.PriceList = row.PriceList
                    oLandedCost_ItemLine.Project = row.Project
                    oLandedCost_ItemLine.ProjectedCustoms = row.ProjectedCustoms
                    oLandedCost_ItemLine.ProjectedCustomsFC = row.ProjectedCustomsFC
                    oLandedCost_ItemLine.Quantity = row.Quantity
                    oLandedCost_ItemLine.TotalVolume = row.TotalVolume
                    oLandedCost_ItemLine.VatGroup = row.VatGroup
                    oLandedCost_ItemLine.Volume = row.Volume
                    oLandedCost_ItemLine.VolumeUoM = row.VolumeUoM
                    oLandedCost_ItemLine.Warehouse = row.Warehouse

                    oLandedCost_ItemLine.DistributionRule = row.DistributionRule
                    oLandedCost_ItemLine.DistributionRule2 = row.DistributionRule2
                    oLandedCost_ItemLine.DistributionRule3 = row.DistributionRule3

                Next
                For Each row As SAPbobsCOM.LandedCost_CostLine In oSLandedCost.LandedCost_CostLines
                    oLandedCost_CostLine = oLandedCost.LandedCost_CostLines.Add
                    oLandedCost_CostLine.LandedCostCode = row.LandedCostCode
                    oLandedCost_CostLine.amount = row.amount
                    oLandedCost_CostLine.AllocationBy = row.AllocationBy
                    oLandedCost_CostLine.CostType = row.CostType
                    oLandedCost_CostLine.IncludeForCustoms = row.IncludeForCustoms
                Next

                'Add a landed cost 
                Try
                    oLandedCostParams = svrLandedCost.AddLandedCost(oLandedCost)
                    Dim oLandedCostEntry As String = oLandedCostParams.LandedCostNumber
                Catch ex As Exception
                    'exception process 
                    MsgBox(ex.Message)
                End Try
            Catch ex As Exception
                'exception process 
                MsgBox(ex.Message)
            End Try
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
#Region "Check Duplicate Entry"
    Private Function checkDuplicateInvoice(ByVal aDocEntry As Integer, ByVal aDocNum As Integer, ByVal aBranchcode As String, ByVal aTableName As String) As Boolean
        Dim oDuplicateCheck As SAPbobsCOM.Recordset
        Dim strQuery1 As String
        oDuplicateCheck = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strQuery1 = "Select * from " & aTableName & " where U_BaseEntry=" & aDocEntry & " and U_BaseNum=" & aDocNum & " and U_Branch='" & aBranchcode & "'"
        oDuplicateCheck.DoQuery(strQuery1)
        If oDuplicateCheck.RecordCount > 0 Then
            If Not IsNothing(oDuplicateCheck) Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDuplicateCheck)
            End If
            Return True
        ElseIf Not IsNothing(oDuplicateCheck) Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oDuplicateCheck)
            GC.Collect()
        End If
        If Not IsNothing(oDuplicateCheck) Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oDuplicateCheck)
            GC.Collect()
        End If
        Return False


    End Function
#End Region

    Private Function CheckCanceledInvoice(ByVal aDocEntry As Integer, ByVal aDocNum As Integer, ByVal aBranchcode As String, ByVal aTableName As String) As Integer
        Dim oDuplicateCheck As SAPbobsCOM.Recordset
        Dim strQuery1 As String
        oDuplicateCheck = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim objMainRec As SAPbobsCOM.Recordset
        objMainRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If aTableName = "ORCT" Then
            strQuery1 = "Select * from ORCT where DocEntry=" & aDocEntry
            objMainRec.DoQuery(strQuery1)
            strQuery1 = "Select * from " & aTableName & " where U_BaseEntry=" & objMainRec.Fields.Item("DocNum").Value ' & " and U_BaseNum=" & aDocNum & " and U_Branch='" & aBranchcode & "'"
        ElseIf aTableName = "OVPM" Then
            strQuery1 = "Select * from OVPM where DocEntry=" & aDocEntry
            objMainRec.DoQuery(strQuery1)
            strQuery1 = "Select * from " & aTableName & " where U_BaseEntry=" & objMainRec.Fields.Item("DocNum").Value ' & " and U_BaseNum=" & aDocNum & " and U_Branch='" & aBranchcode & "'"

        Else
            '  strQuery1 = "Select * from INV1 where DocEntry=" & aDocEntry
            Select Case aTableName
                Case "OPDN"
                    strQuery1 = "Select * from PDN1 where DocEntry=" & aDocEntry
                Case "OPOR"
                    strQuery1 = "Select * from POR1 where DocEntry=" & aDocEntry
                Case "ORPD"
                    strQuery1 = "Select * from RPD1 where DocEntry=" & aDocEntry
                Case "OPCH"
                    strQuery1 = "Select * from PCH1 where DocEntry=" & aDocEntry
                Case "ORPC"
                    strQuery1 = "Select * from RPC1 where DocEntry=" & aDocEntry
                Case "OINV"
                    strQuery1 = "Select * from INV1 where DocEntry=" & aDocEntry
                Case "ODLN"
                    strQuery1 = "Select * from DLN1 where DocEntry=" & aDocEntry
                Case "ORIN"
                    strQuery1 = "Select * from RIN1 where DocEntry=" & aDocEntry
                Case "OINV"
                    strQuery1 = "Select * from INV1 where DocEntry=" & aDocEntry
            End Select
            objMainRec.DoQuery(strQuery1)
            strQuery1 = "Select * from " & aTableName & " where U_BaseEntry=" & objMainRec.Fields.Item("BaseEntry").Value ' & " and U_BaseNum=" & aDocNum & " and U_Branch='" & aBranchcode & "'"
            '    strQuery1 = "Select * from " & aTableName & " where U_BaseEntry=" & objMainRec.Fields.Item("DocEntry").Value ' & " and U_BaseNum=" & aDocNum & " and U_Branch='" & aBranchcode & "'"
        End If

        oDuplicateCheck.DoQuery(strQuery1)
        If oDuplicateCheck.RecordCount > 0 Then
            ' MsgBox(oDuplicateCheck.Fields.Item("DocEntry").Value)
            Return oDuplicateCheck.Fields.Item("DocEntry").Value
        Else
            Return 0
        End If

    End Function


    Private Function CheckCanceledInvoice_Documents(ByVal aDocEntry As Integer, ByVal aDocNum As Integer, ByVal aBranchcode As String, ByVal aTableName As String) As Integer
        Dim oDuplicateCheck As SAPbobsCOM.Recordset
        Dim strQuery1 As String
        oDuplicateCheck = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim objMainRec As SAPbobsCOM.Recordset
        objMainRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If aTableName = "ORCT" Then
            strQuery1 = "Select * from ORCT where DocEntry=" & aDocEntry
            objMainRec.DoQuery(strQuery1)
            strQuery1 = "Select * from " & aTableName & " where U_BaseEntry=" & objMainRec.Fields.Item("DocNum").Value ' & " and U_BaseNum=" & aDocNum & " and U_Branch='" & aBranchcode & "'"
        Else
            strQuery1 = "Select * from INV1 where DocEntry=" & aDocEntry
            objMainRec.DoQuery(strQuery1)
            strQuery1 = "Select * from " & aTableName & " where U_BaseEntry=" & objMainRec.Fields.Item("BaseEntry").Value ' & " and U_BaseNum=" & aDocNum & " and U_Branch='" & aBranchcode & "'"
        End If

        oDuplicateCheck.DoQuery(strQuery1)
        If oDuplicateCheck.RecordCount > 0 Then
            ' MsgBox(oDuplicateCheck.Fields.Item("DocEntry").Value)
            Return oDuplicateCheck.Fields.Item("DocEntry").Value
        Else
            Return 0
        End If

    End Function

#Region "Master Data Update"

    Private Sub ImportBOM()
        Dim FileName, strItem As String
        Dim Ecount As Long
        Dim ii As Long
        Dim objMainItem, objRemoteItem As SAPbobsCOM.ProductTrees
        Dim objMainRecset As SAPbobsCOM.Recordset
        objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRecset.DoQuery("Select * from OITT where   U_Action<>'N'")
        For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
            FileName = System.Windows.Forms.Application.StartupPath & "\Item.xml"
            strItem = objMainRecset.Fields.Item("Code").Value
            objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductTrees)
            If objMainItem.GetByKey(strItem) Then
                objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductTrees)
                If objRemoteItem.GetByKey(strItem) Then
                    objRemoteItem.TreeCode = objMainItem.TreeCode
                    objRemoteItem.TreeType = SAPbobsCOM.BoItemTreeTypes.iProductionTree
                    objRemoteItem.Quantity = objMainItem.Quantity
                    ' objRemoteItem.UserFields.Fields.Item("U_SaleBOM").Value = objMainItem.UserFields.Fields.Item("U_SaleBOM").Value
                    For intRow As Integer = 0 To objMainItem.Items.Count - 1
                        If intRow > 0 Then
                            objRemoteItem.Items.Add()
                            objRemoteItem.Items.SetCurrentLine(intRow)
                        End If
                        objMainItem.Items.SetCurrentLine(intRow)
                        objRemoteItem.Items.ItemCode = objMainItem.Items.ItemCode
                        objRemoteItem.Items.IssueMethod = objMainItem.Items.IssueMethod
                        objRemoteItem.Items.ParentItem = objMainItem.Items.ParentItem
                        objRemoteItem.Items.Price = objMainItem.Items.Price
                        objRemoteItem.Items.PriceList = objMainItem.Items.PriceList
                        objRemoteItem.Items.Currency = objMainItem.Items.Currency
                        objRemoteItem.Items.Quantity = objMainItem.Items.Quantity
                        objRemoteItem.Items.Warehouse = objMainItem.Items.Warehouse
                    Next
                    If objRemoteItem.Update() <> 0 Then
                        WriteErrorlog("Failed to Update BOM : " & objMainItem.TreeCode & " : " & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        objMainItem.UserFields.Fields.Item("U_Action").Value = "N"
                        objMainItem.Update()
                        WriteErrorlog("BOM Updated : " & objRemoteItem.TreeCode, sPath)
                    End If
                Else
                    objRemoteItem.TreeCode = objMainItem.TreeCode
                    objRemoteItem.TreeType = SAPbobsCOM.BoItemTreeTypes.iProductionTree
                    objRemoteItem.Quantity = objMainItem.Quantity
                    'objRemoteItem.UserFields.Fields.Item("U_SaleBOM").Value = objMainItem.UserFields.Fields.Item("U_SaleBOM").Value
                    For intRow As Integer = 0 To objMainItem.Items.Count - 1
                        If intRow > 0 Then
                            objRemoteItem.Items.Add()
                            objRemoteItem.Items.SetCurrentLine(intRow)
                        End If
                        objMainItem.Items.SetCurrentLine(intRow)
                        objRemoteItem.Items.ItemCode = objMainItem.Items.ItemCode
                        objRemoteItem.Items.IssueMethod = objMainItem.Items.IssueMethod
                        objRemoteItem.Items.ParentItem = objMainItem.Items.ParentItem
                        objRemoteItem.Items.Price = objMainItem.Items.Price
                        objRemoteItem.Items.PriceList = objMainItem.Items.PriceList
                        objRemoteItem.Items.Currency = objMainItem.Items.Currency
                        objRemoteItem.Items.Quantity = objMainItem.Items.Quantity
                        objRemoteItem.Items.Warehouse = objMainItem.Items.Warehouse
                    Next
                    If objRemoteItem.Add() <> 0 Then
                        WriteErrorlog("Failed to create BOM : " & objMainItem.TreeCode & " : " & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        objMainItem.UserFields.Fields.Item("U_Action").Value = "N"
                        objMainItem.Update()
                        WriteErrorlog("New BOM Created : " & objRemoteItem.TreeCode, sPath)
                    End If
                End If
            End If
            objMainRecset.MoveNext()
        Next
    End Sub

#Region "Export Chart of Accounts"
    Private Sub ExportChartofAccounts()
        Dim FileName, strItem, strDflWhs, strFormatCode, strAccountCode As String
        Dim objMainItem, objRemoteItem As SAPbobsCOM.ChartOfAccounts
        Dim objMainRecset, oTemp As SAPbobsCOM.Recordset
        objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRecset.DoQuery("Select AcctCode,formatcode from OACT where U_Action<>'N'")
        For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
            strAccountCode = objMainRecset.Fields.Item(0).Value
            strItem = objMainRecset.Fields.Item(0).Value
            objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oChartOfAccounts)
            If objMainItem.GetByKey(strItem) Then
                strFormatCode = objMainItem.FormatCode
                oTemp = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp.DoQuery("Select AcctCode from OACT where FormatCode='" & strFormatCode & "'")
                strItem = oTemp.Fields.Item(0).Value
                objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oChartOfAccounts)
                If objRemoteItem.GetByKey(strItem) Then
                    objRemoteItem.AccountType = objMainItem.AccountType
                    objRemoteItem.AcctCurrency = objMainItem.AcctCurrency
                    objRemoteItem.ActiveAccount = objMainItem.ActiveAccount
                    objRemoteItem.AllowChangeVatGroup = objMainItem.AllowChangeVatGroup
                    objRemoteItem.BudgetAccount = objMainItem.BudgetAccount
                    objRemoteItem.Code = objMainItem.Code
                    objRemoteItem.CashAccount = objMainItem.CashAccount
                    objRemoteItem.Category = objMainItem.Category
                    objRemoteItem.DataExportCode = objMainItem.DataExportCode
                    objRemoteItem.DefaultVatGroup = objMainItem.DefaultVatGroup
                    objRemoteItem.Details = objMainItem.Details
                    objRemoteItem.ExternalCode = objMainItem.ExternalCode
                    objRemoteItem.FatherAccountKey = objMainItem.FatherAccountKey
                    objRemoteItem.ForeignName = objMainItem.ForeignName
                    objRemoteItem.FormatCode = objMainItem.FormatCode
                    objRemoteItem.LiableForAdvances = objMainItem.LiableForAdvances
                    objRemoteItem.LockManualTransaction = objMainItem.LockManualTransaction
                    objRemoteItem.Name = objMainItem.Name
                    objRemoteItem.ProjectCode = objMainItem.ProjectCode
                    objRemoteItem.Protected = objMainItem.Protected
                    objRemoteItem.ReconciledAccount = objMainItem.ReconciledAccount
                    objRemoteItem.TaxExemptAccount = objMainItem.TaxExemptAccount
                    objRemoteItem.TaxLiableAccount = objMainItem.TaxLiableAccount
                    objRemoteItem.TransactionCode = objMainItem.TransactionCode
                    If objRemoteItem.Update() <> 0 Then
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Failed to Update Chart of Accounts : " & objRemoteItem.FormatCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Chart of Accounts : " & objRemoteItem.FormatCode & " Updated successfully", sPath)
                        objMainItem.UserFields.Fields.Item("U_Action").Value = "N"
                        objMainItem.Update()
                    End If

                Else
                    objRemoteItem.AccountType = objMainItem.AccountType
                    objRemoteItem.AcctCurrency = objMainItem.AcctCurrency
                    objRemoteItem.ActiveAccount = objMainItem.ActiveAccount
                    objRemoteItem.AllowChangeVatGroup = objMainItem.AllowChangeVatGroup
                    objRemoteItem.BudgetAccount = objMainItem.BudgetAccount
                    objRemoteItem.Code = objMainItem.Code
                    objRemoteItem.CashAccount = objMainItem.CashAccount
                    objRemoteItem.Category = objMainItem.Category
                    objRemoteItem.DataExportCode = objMainItem.DataExportCode
                    objRemoteItem.DefaultVatGroup = objMainItem.DefaultVatGroup
                    objRemoteItem.Details = objMainItem.Details
                    objRemoteItem.ExternalCode = objMainItem.ExternalCode
                    objRemoteItem.FatherAccountKey = objMainItem.FatherAccountKey
                    objRemoteItem.ForeignName = objMainItem.ForeignName
                    objRemoteItem.FormatCode = objMainItem.FormatCode
                    objRemoteItem.LiableForAdvances = objMainItem.LiableForAdvances
                    objRemoteItem.LockManualTransaction = objMainItem.LockManualTransaction
                    objRemoteItem.Name = objMainItem.Name
                    objRemoteItem.ProjectCode = objMainItem.ProjectCode
                    objRemoteItem.Protected = objMainItem.Protected
                    objRemoteItem.ReconciledAccount = objMainItem.ReconciledAccount
                    objRemoteItem.TaxExemptAccount = objMainItem.TaxExemptAccount
                    objRemoteItem.TaxLiableAccount = objMainItem.TaxLiableAccount
                    objRemoteItem.TransactionCode = objMainItem.TransactionCode
                    If objRemoteItem.Add <> 0 Then
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " :  Failed to Add Chart of Accounts : " & objRemoteItem.FormatCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Chart of Accounts : " & objRemoteItem.FormatCode & " Added successfully", sPath)
                        objMainItem.UserFields.Fields.Item("U_Action").Value = "N"
                        objMainItem.Update()
                    End If
                End If
            End If


            objMainRecset.MoveNext()
        Next
    End Sub
#End Region

    Private Sub ExportSegmentation()
        Dim FileName, strItem, strDflWhs, strName As String
        Dim objMainItem, objRemoteItem As SAPbobsCOM.AccountSegmentationCategories
        Dim objMainRecset As SAPbobsCOM.Recordset
        objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRecset.DoQuery("Select * from OASC")
        For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
            strItem = objMainRecset.Fields.Item("SegmentId").Value
            strName = objMainRecset.Fields.Item("Code").Value
            objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAccountSegmentationCategories)
            If objMainItem.GetByKey(CInt(strItem), strName) Then
                objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAccountSegmentationCategories)
                If objRemoteItem.GetByKey(CInt(strItem), strName) Then
                    objRemoteItem.Code = objMainItem.Code
                    objRemoteItem.Name = objMainItem.Name
                    objRemoteItem.SegmentID = objMainItem.SegmentID
                    objRemoteItem.ShortName = objMainItem.ShortName
                    If objRemoteItem.Update() <> 0 Then
                        ' WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Failed to Accoutns Segementation : " & objRemoteItem.FormatCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        ' WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Account Segemntation : " & objRemoteItem.Name & " Updated successfully", sPath)
                    End If

                Else
                    objRemoteItem.Code = objMainItem.Code
                    objRemoteItem.Name = objMainItem.Name
                    objRemoteItem.SegmentID = objMainItem.SegmentID
                    objRemoteItem.ShortName = objMainItem.ShortName
                    If objRemoteItem.Add <> 0 Then
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " :  Failed to Accoutns Segementation : " & objMainItem.Name & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Account Segemntation : " & objRemoteItem.Name & " Added successfully", sPath)
                    End If
                End If
            End If
            objMainRecset.MoveNext()
        Next
    End Sub

    Private Sub ExportCurrencies()
        Dim FileName, strItem, strDflWhs As String
        Dim objMainItem, objRemoteItem As SAPbobsCOM.Currencies
        Dim objMainRecset As SAPbobsCOM.Recordset
        objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRecset.DoQuery("Select CurrCode from OCRN ")
        For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
            strItem = objMainRecset.Fields.Item(0).Value
            objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCurrencyCodes)
            If objMainItem.GetByKey(strItem) Then
                objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCurrencyCodes)
                If objRemoteItem.GetByKey(strItem) Then
                    objRemoteItem.Code = objMainItem.Code
                    objRemoteItem.Decimals = objMainItem.Decimals
                    objRemoteItem.DocumentsCode = objMainItem.DocumentsCode
                    objRemoteItem.EnglishHundredthName = objMainItem.EnglishHundredthName
                    objRemoteItem.EnglishName = objMainItem.EnglishName
                    objRemoteItem.HundredthName = objMainItem.HundredthName
                    objRemoteItem.InternationalDescription = objMainItem.InternationalDescription
                    objRemoteItem.Name = objMainItem.Name
                    objRemoteItem.Rounding = objMainItem.Rounding
                    objRemoteItem.RoundingInPayment = objMainItem.RoundingInPayment
                    If objRemoteItem.Update() <> 0 Then
                        '   WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Failed to Accoutns Segementation : " & objRemoteItem.FormatCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        '   WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Account Segemntation : " & objRemoteItem.Name & " Updated successfully", sPath)
                    End If

                Else
                    objRemoteItem.Code = objMainItem.Code
                    objRemoteItem.Decimals = objMainItem.Decimals
                    objRemoteItem.DocumentsCode = objMainItem.DocumentsCode
                    objRemoteItem.EnglishHundredthName = objMainItem.EnglishHundredthName
                    objRemoteItem.EnglishName = objMainItem.EnglishName
                    objRemoteItem.HundredthName = objMainItem.HundredthName
                    objRemoteItem.InternationalDescription = objMainItem.InternationalDescription
                    objRemoteItem.Name = objMainItem.Name
                    objRemoteItem.Rounding = objMainItem.Rounding
                    objRemoteItem.RoundingInPayment = objMainItem.RoundingInPayment

                    If objRemoteItem.Add <> 0 Then
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " :  Failed to Add Currency : " & objMainItem.Code & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Currency : " & objRemoteItem.Code & " Added successfully", sPath)
                    End If
                End If
            End If
            objMainRecset.MoveNext()
        Next
    End Sub

    Private Sub ExportWarehouseLocation()
        Dim FileName, strItem, strDflWhs As String
        Dim objMainItem, objRemoteItem As SAPbobsCOM.WarehouseLocations
        Dim objMainRecset As SAPbobsCOM.Recordset
        objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRecset.DoQuery("Select Code from OLCT")
        For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
            strItem = objMainRecset.Fields.Item(0).Value
            objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWarehouseLocations)
            If objMainItem.GetByKey(strItem) Then
                objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWarehouseLocations)
                If objRemoteItem.GetByKey(strItem) Then
                    objRemoteItem.Name = objMainItem.Name
                    If objRemoteItem.Update() <> 0 Then
                        ' WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Failed to Accoutns Segementation : " & objRemoteItem.FormatCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        ' WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Account Segemntation : " & objRemoteItem.Name & " Updated successfully", sPath)
                    End If
                Else
                    objRemoteItem.Name = objMainItem.Name
                    If objRemoteItem.Add <> 0 Then
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " :  Failed to add Location : " & objMainItem.Name & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Location : " & objRemoteItem.Name & " Added successfully", sPath)
                    End If
                End If
            End If

            objMainRecset.MoveNext()
        Next
    End Sub

    Private Sub ExportProjectCodes()
        Dim FileName, strItem, strDflWhs, strProject, strProjectName As String
        Dim objMainItem, objRemoteItem As SAPbobsCOM.ProjectsService
        Dim oCmpSrv As SAPbobsCOM.CompanyService
        Dim objMainRecset, objRemoteRecc As SAPbobsCOM.Recordset
        objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRecset.DoQuery("Select prjcode,prjName from OPRJ")
        For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
            strProject = objMainRecset.Fields.Item(0).Value
            strProjectName = objMainRecset.Fields.Item(1).Value
            If 1 = 1 Then
                objRemoteRecc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRemoteRecc.DoQuery("Select * from OPRJ where PrjCode='" & strProject & "'")
                If objRemoteRecc.RecordCount > 0 Then
                    Dim projectService As SAPbobsCOM.IProjectsService
                    Dim project As SAPbobsCOM.IProject
                    Dim projectParams As SAPbobsCOM.IProjectParams
                    oCmpSrv = objRemoteCompany.GetCompanyService
                    projectService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ProjectsService)
                    'Get a project
                    projectParams = projectService.GetDataInterface(SAPbobsCOM.ProjectsServiceDataInterfaces.psProjectParams)
                    projectParams.Code = strProject
                    project = projectService.GetProject(projectParams)
                    'Update the project
                    project.Name = strProjectName
                    projectService.UpdateProject(project)
                Else
                    Dim projectService As SAPbobsCOM.IProjectsService
                    Dim project As SAPbobsCOM.IProject
                    oCmpSrv = objRemoteCompany.GetCompanyService
                    projectService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ProjectsService)
                    project = projectService.GetDataInterface(SAPbobsCOM.ProjectsServiceDataInterfaces.psProject)
                    project.Code = strProject
                    project.Name = strProject
                    projectService.AddProject(project)
                    'objRemoteItem.Name = objMainItem.Name
                    'objRemoteItem.Size = objMainItem.Size
                    'objRemoteItem.Type = objMainItem.Type
                    'If objRemoteItem.Add <> 0 Then
                    '    WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " :  Failed to Accoutns Segementation : " & objRemoteItem.FormatCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    'Else
                    '    WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Account Segemntation : " & objRemoteItem.Name & " Added successfully", sPath)
                    'End If
                End If
            End If



            objMainRecset.MoveNext()
        Next
    End Sub

    Private Sub ExportSalesTaxCodes()
        Dim FileName, strItem, strDflWhs As String
        Dim objMainItem, objRemoteItem As SAPbobsCOM.VatGroups
        Dim objMainRecset As SAPbobsCOM.Recordset
        objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRecset.DoQuery("Select Code from OVTG")
        For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
            strItem = objMainRecset.Fields.Item(0).Value
            objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVatGroups)
            If objMainItem.GetByKey(strItem) Then
                objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oVatGroups)
                If objRemoteItem.GetByKey(strItem) Then
                    objRemoteItem.Category = objMainItem.Category
                    objRemoteItem.Code = objMainItem.Code
                    objRemoteItem.Correction = objMainItem.Correction
                    objRemoteItem.DeferredTaxAcc = objMainItem.DeferredTaxAcc
                    objRemoteItem.EU = objMainItem.EU
                    objRemoteItem.GoodsShipment = objMainItem.GoodsShipment
                    objRemoteItem.Name = objMainItem.Name
                    objRemoteItem.NonDeduct = objMainItem.NonDeduct
                    objRemoteItem.NonDeductAcc = objMainItem.NonDeductAcc
                    objRemoteItem.TaxAccount = objMainItem.TaxAccount
                    objRemoteItem.TriangularDeal = objMainItem.TriangularDeal
                    objRemoteItem.VatCorrection = objMainItem.VatCorrection
                    For intRow As Integer = 0 To objMainItem.VatGroups_Lines.Count - 1
                        If intRow > 0 Then
                            objRemoteItem.VatGroups_Lines.Add()
                            objRemoteItem.VatGroups_Lines.SetCurrentLine(intRow)
                        End If
                        objMainItem.VatGroups_Lines.SetCurrentLine(intRow)
                        objRemoteItem.VatGroups_Lines.Effectivefrom = objMainItem.VatGroups_Lines.Effectivefrom
                        objRemoteItem.VatGroups_Lines.EqualizationTax = objMainItem.VatGroups_Lines.EqualizationTax
                        objRemoteItem.VatGroups_Lines.Rate = objMainItem.VatGroups_Lines.Rate
                    Next
                    If objRemoteItem.Update() <> 0 Then
                        ' WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Failed to Accoutns Segementation : " & objRemoteItem.FormatCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        ' WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Account Segemntation : " & objRemoteItem.Name & " Updated successfully", sPath)
                    End If

                Else
                    objRemoteItem.Category = objMainItem.Category
                    objRemoteItem.Code = objMainItem.Code
                    objRemoteItem.Correction = objMainItem.Correction
                    objRemoteItem.DeferredTaxAcc = objMainItem.DeferredTaxAcc
                    objRemoteItem.EU = objMainItem.EU
                    objRemoteItem.GoodsShipment = objMainItem.GoodsShipment
                    objRemoteItem.Name = objMainItem.Name
                    objRemoteItem.NonDeduct = objMainItem.NonDeduct
                    objRemoteItem.NonDeductAcc = objMainItem.NonDeductAcc
                    objRemoteItem.TaxAccount = objMainItem.TaxAccount
                    objRemoteItem.TriangularDeal = objMainItem.TriangularDeal
                    objRemoteItem.VatCorrection = objMainItem.VatCorrection
                    For intRow As Integer = 0 To objMainItem.VatGroups_Lines.Count - 1
                        If intRow > 0 Then
                            objRemoteItem.VatGroups_Lines.Add()
                            objRemoteItem.VatGroups_Lines.SetCurrentLine(intRow)
                        End If
                        objMainItem.VatGroups_Lines.SetCurrentLine(intRow)
                        objRemoteItem.VatGroups_Lines.Effectivefrom = objMainItem.VatGroups_Lines.Effectivefrom
                        objRemoteItem.VatGroups_Lines.EqualizationTax = objMainItem.VatGroups_Lines.EqualizationTax
                        objRemoteItem.VatGroups_Lines.Rate = objMainItem.VatGroups_Lines.Rate
                    Next
                    If objRemoteItem.Add <> 0 Then
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " :  Failed to add  tax Code : " & objMainItem.Code & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Tax code : " & objRemoteItem.Code & " Added successfully", sPath)
                    End If
                End If
            End If

            objMainRecset.MoveNext()
        Next
    End Sub

    Private Sub ExportBPGroups()
        Dim FileName, strItem, strDflWhs, strName, strMainBPGroup As String
        Dim objMainItem, objRemoteItem As SAPbobsCOM.BusinessPartnerGroups
        Dim objMainRecset, objtemp As SAPbobsCOM.Recordset
        objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRecset.DoQuery("Select GroupCode,GroupName from OCRG")
        For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
            strMainBPGroup = objMainRecset.Fields.Item(0).Value
            strItem = objMainRecset.Fields.Item(0).Value
            objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartnerGroups)
            If objMainItem.GetByKey(strItem) Then
                strName = objMainItem.Name
                objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartnerGroups)
                objtemp = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objtemp.DoQuery("Select GroupCode from OCRG where GroupName='" & strName & "'")
                strItem = objtemp.Fields.Item(0).Value

                If objRemoteItem.GetByKey(strItem) Then
                    objRemoteItem.Name = objMainItem.Name
                    objRemoteItem.Type = objMainItem.Type
                    If objRemoteItem.Update() <> 0 Then
                        ' WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Failed to Accoutns Segementation : " & objRemoteItem.FormatCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        'WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Account Segemntation : " & objRemoteItem.Name & " Updated successfully", sPath)
                    End If
                Else
                    objRemoteItem.Name = objMainItem.Name
                    objRemoteItem.Type = objMainItem.Type
                    If objRemoteItem.Add <> 0 Then
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " :  Failed to add BP Group : " & objMainItem.Name & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : BP Group : " & objRemoteItem.Name & " Added successfully", sPath)
                    End If
                End If
            End If
            objMainRecset.MoveNext()
        Next
    End Sub

    Private Sub ExportPaymentTerms()
        Dim FileName, strItem, strDflWhs As String
        Dim objMainItem, objRemoteItem As SAPbobsCOM.PaymentTermsTypes
        Dim objMainRecset As SAPbobsCOM.Recordset
        objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRecset.DoQuery("Select GroupNum from OCTG")
        For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
            strItem = objMainRecset.Fields.Item(0).Value
            objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentTermsTypes)
            If objMainItem.GetByKey(strItem) Then
                objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPaymentTermsTypes)
                If objRemoteItem.GetByKey(strItem) Then
                    objRemoteItem.BaselineDate = objMainItem.BaselineDate
                    objRemoteItem.CreditLimit = objMainItem.CreditLimit
                    objRemoteItem.DiscountCode = objMainItem.DiscountCode
                    objRemoteItem.DunningCode = objMainItem.DunningCode
                    objRemoteItem.GeneralDiscount = objMainItem.GeneralDiscount
                    objRemoteItem.InterestOnArrears = objMainItem.InterestOnArrears
                    objRemoteItem.LoadLimit = objMainItem.LoadLimit
                    objRemoteItem.NumberOfAdditionalDays = objMainItem.NumberOfAdditionalDays
                    objRemoteItem.NumberOfAdditionalMonths = objMainItem.NumberOfAdditionalMonths
                    objRemoteItem.NumberOfToleranceDays = objMainItem.NumberOfToleranceDays
                    objRemoteItem.OpenReceipt = objMainItem.OpenReceipt
                    objRemoteItem.PaymentTermsGroupName = objMainItem.PaymentTermsGroupName
                    objRemoteItem.PriceListNo = objMainItem.PriceListNo
                    objRemoteItem.StartFrom = objMainItem.StartFrom
                    If objRemoteItem.Update() <> 0 Then
                        ' WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Failed to Accoutns Segementation : " & objRemoteItem.FormatCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        ' WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Account Segemntation : " & objRemoteItem.Name & " Updated successfully", sPath)
                    End If

                Else
                    objRemoteItem.BaselineDate = objMainItem.BaselineDate
                    objRemoteItem.CreditLimit = objMainItem.CreditLimit
                    objRemoteItem.DiscountCode = objMainItem.DiscountCode
                    objRemoteItem.DunningCode = objMainItem.DunningCode
                    objRemoteItem.GeneralDiscount = objMainItem.GeneralDiscount
                    objRemoteItem.InterestOnArrears = objMainItem.InterestOnArrears
                    objRemoteItem.LoadLimit = objMainItem.LoadLimit
                    objRemoteItem.NumberOfAdditionalDays = objMainItem.NumberOfAdditionalDays
                    objRemoteItem.NumberOfAdditionalMonths = objMainItem.NumberOfAdditionalMonths
                    objRemoteItem.NumberOfToleranceDays = objMainItem.NumberOfToleranceDays
                    objRemoteItem.OpenReceipt = objMainItem.OpenReceipt
                    objRemoteItem.PaymentTermsGroupName = objMainItem.PaymentTermsGroupName
                    objRemoteItem.PriceListNo = objMainItem.PriceListNo
                    objRemoteItem.StartFrom = objMainItem.StartFrom
                    If objRemoteItem.Add <> 0 Then
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " :  Failed to add Payment Terms : " & objMainItem.GroupNumber & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Payment terms : " & objRemoteItem.GroupNumber & " Added successfully", sPath)
                    End If
                End If
            End If


            objMainRecset.MoveNext()
        Next
    End Sub

    Private Sub ExportBPProperties()



        Dim FileName, strItem, strDflWhs As String
        Dim objMainItem, objRemoteItem As SAPbobsCOM.BusinessPartnerPropertiesService
        Dim getLine, getLineDesg As SAPbobsCOM.BusinessPartnerPropertyParams
        Dim updateline, updatelinedesg As SAPbobsCOM.BusinessPartnerProperty

        objMainItem = objMainCompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.BusinessPartnerPropertiesService)
        objRemoteItem = objRemoteCompany.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.BusinessPartnerPropertiesService)

        For intRow As Integer = 1 To 64
            getLine = objMainItem.GetDataInterface(SAPbobsCOM.BusinessPartnerPropertiesServiceDataInterfaces.bppsBusinessPartnerPropertyParams)
            getLine.PropertyCode = intRow
            updateline = objMainItem.GetBusinessPartnerProperty(getLine)
            getLineDesg = objRemoteItem.GetDataInterface(SAPbobsCOM.BusinessPartnerPropertiesServiceDataInterfaces.bppsBusinessPartnerPropertyParams)
            getLineDesg.PropertyCode = intRow
            updatelinedesg = objRemoteItem.GetBusinessPartnerProperty(getLineDesg)
            updatelinedesg.PropertyName = updateline.PropertyName
            objRemoteItem.UpdateBusinessPartnerProperty(updatelinedesg)
        Next



    End Sub

    Private Sub ExportItemGroup()
        Dim FileName, strItem, strDflWhs, strName, strMainItemGroup As String
        Dim objMainItem, objRemoteItem As SAPbobsCOM.ItemGroups
        Dim objMainRecset, objtemp As SAPbobsCOM.Recordset
        objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRecset.DoQuery("Select ItmsGrpCod from OITB ")
        For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
            strMainItemGroup = objMainRecset.Fields.Item(0).Value
            strItem = objMainRecset.Fields.Item(0).Value
            objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItemGroups)
            If objMainItem.GetByKey(CInt(strItem)) Then
                objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItemGroups)
                strName = objMainItem.GroupName
                objtemp = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objtemp.DoQuery("Select ItmsGrpCod from OITB where ItmsGrpNam='" & strName & "'")
                strItem = objtemp.Fields.Item(0).Value
                If objRemoteItem.GetByKey(strItem) Then
                    objRemoteItem.CostAccount = objMainItem.CostAccount
                    objRemoteItem.CostInflationAccount = objMainItem.CostInflationAccount
                    objRemoteItem.CostInflationOffsetAccount = objMainItem.CostInflationOffsetAccount
                    ' objRemoteItem.CycleCode = objMainItem.CycleCode
                    objRemoteItem.DecreaseGLAccount = objMainItem.DecreaseGLAccount
                    objRemoteItem.DecreasingAccount = objMainItem.DecreasingAccount
                    objRemoteItem.EUExpensesAccount = objMainItem.EUExpensesAccount
                    objRemoteItem.EUPurchaseCreditAcc = objMainItem.EUPurchaseCreditAcc
                    objRemoteItem.EURevenuesAccount = objMainItem.EURevenuesAccount
                    objRemoteItem.ExemptedCredits = objMainItem.ExemptedCredits
                    objRemoteItem.ExemptRevenuesAccount = objMainItem.ExemptRevenuesAccount
                    objRemoteItem.ExpenseClearingAct = objMainItem.ExpenseClearingAct
                    objRemoteItem.ExpenseOffsetAccount = objMainItem.ExpenseOffsetAccount
                    objRemoteItem.ExpensesAccount = objMainItem.ExpensesAccount
                    objRemoteItem.ForeignExpensesAccount = objMainItem.ForeignExpensesAccount
                    objRemoteItem.ForeignPurchaseCreditAcc = objMainItem.ForeignPurchaseCreditAcc
                    objRemoteItem.ForeignRevenuesAccount = objMainItem.ForeignRevenuesAccount
                    objRemoteItem.GoodsClearingAccount = objMainItem.GoodsClearingAccount
                    objRemoteItem.GroupName = objMainItem.GroupName
                    objRemoteItem.IncreaseGLAccount = objMainItem.IncreaseGLAccount
                    objRemoteItem.IncreasingAccount = objMainItem.IncreasingAccount
                    objRemoteItem.InventoryAccount = objMainItem.InventoryAccount
                    objRemoteItem.InventorySystem = objMainItem.InventorySystem
                    objRemoteItem.LeadTime = objMainItem.LeadTime
                    objRemoteItem.MinimumOrderQuantity = objMainItem.MinimumOrderQuantity
                    objRemoteItem.NegativeInventoryAdjustmentAccount = objMainItem.NegativeInventoryAdjustmentAccount
                    objRemoteItem.OrderInterval = objMainItem.OrderInterval
                    objRemoteItem.OrderMultiple = objMainItem.OrderMultiple
                    objRemoteItem.PAReturnAccount = objMainItem.PAReturnAccount
                    objRemoteItem.PlanningSystem = objMainItem.PlanningSystem
                    objRemoteItem.PriceDifferencesAccount = objMainItem.PriceDifferencesAccount
                    objRemoteItem.ProcurementMethod = objMainItem.ProcurementMethod
                    objRemoteItem.PurchaseAccount = objMainItem.PurchaseAccount
                    objRemoteItem.PurchaseCreditAcc = objMainItem.PurchaseCreditAcc
                    objRemoteItem.PurchaseOffsetAccount = objMainItem.PurchaseOffsetAccount
                    objRemoteItem.ReturningAccount = objMainItem.ReturningAccount
                    objRemoteItem.RevenuesAccount = objMainItem.RevenuesAccount
                    objRemoteItem.SalesCreditAcc = objMainItem.SalesCreditAcc
                    objRemoteItem.SalesCreditEUAcc = objMainItem.SalesCreditEUAcc
                    objRemoteItem.SalesCreditForeignAcc = objMainItem.SalesCreditForeignAcc
                    objRemoteItem.ShippedGoodsAccount = objMainItem.ShippedGoodsAccount
                    objRemoteItem.StockInflationAdjustAccount = objMainItem.StockInflationAdjustAccount
                    objRemoteItem.StockInflationOffsetAccount = objMainItem.StockInflationOffsetAccount
                    objRemoteItem.TransfersAccount = objMainItem.TransfersAccount
                    objRemoteItem.VarianceAccount = objMainItem.VarianceAccount
                    objRemoteItem.VATInRevenueAccount = objMainItem.VATInRevenueAccount
                    '    objRemoteItem.WHIncomingCenvatAccount = objMainItem.WHIncomingCenvatAccount
                    '     objRemoteItem.WHOutgoingCenvatAccount = objMainItem.WHOutgoingCenvatAccount
                    objRemoteItem.WIPMaterialAccount = objMainItem.WIPMaterialAccount
                    objRemoteItem.WIPMaterialVarianceAccount = objMainItem.WIPMaterialVarianceAccount

                    If objRemoteItem.Update() <> 0 Then
                        ' WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Failed to Accoutns Segementation : " & objRemoteItem.FormatCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        ' WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Account Segemntation : " & objRemoteItem.Name & " Updated successfully", sPath)
                    End If
                Else
                    objRemoteItem.GroupName = objMainItem.GroupName
                    If objRemoteItem.Add <> 0 Then
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " :  Failed to add Item Group : " & objMainItem.GroupName & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Item Group : " & objRemoteItem.GroupName & " Added successfully", sPath)
                    End If
                End If
            End If

            objMainRecset.MoveNext()
        Next
    End Sub

    Private Sub exportItemProperties()
        Dim FileName, strItem, strDflWhs, strName, strMainItemGroup As String
        Dim objMainItem, objRemoteItem As SAPbobsCOM.ItemProperties
        Dim objMainRecset, objtemp As SAPbobsCOM.Recordset
        objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRecset.DoQuery("Select ItmsTypCod from OITG ")
        For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
            strMainItemGroup = objMainRecset.Fields.Item(0).Value
            strItem = objMainRecset.Fields.Item(0).Value
            objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItemProperties)
            If objMainItem.GetByKey(CInt(strItem)) Then
                objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItemProperties)
                strName = objMainItem.Number
                objtemp = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objtemp.DoQuery("Select ItmsTypCod from OITG where ItmsTypCod=" & objMainItem.Number)
                strItem = objtemp.Fields.Item(0).Value
                If objRemoteItem.GetByKey(CInt(strItem)) Then
                    objRemoteItem.PropertyName = objMainItem.PropertyName
                    If objRemoteItem.Update() <> 0 Then
                        ' WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Failed to Accoutns Segementation : " & objRemoteItem.FormatCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        ' WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Account Segemntation : " & objRemoteItem.Name & " Updated successfully", sPath)
                    End If
                Else
                    objRemoteItem.PropertyName = objMainItem.PropertyName
                    If objRemoteItem.Add <> 0 Then
                        '  WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " :  Failed to add Item Group : " & objMainItem.GroupName & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        '  WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Item Group : " & objRemoteItem.GroupName & " Added successfully", sPath)
                    End If
                End If
            End If


            objMainRecset.MoveNext()
        Next
    End Sub

    Private Sub ExportWarehosue()
        Dim FileName, strItem, strDflWhs As String
        Dim objMainItem, objRemoteItem As SAPbobsCOM.Warehouses
        Dim objMainRecset As SAPbobsCOM.Recordset
        objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRecset.DoQuery("Select WhsCode from OWHS")
        For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
            strItem = objMainRecset.Fields.Item(0).Value
            objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWarehouses)
            If objMainItem.GetByKey(strItem) Then
                objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWarehouses)
                If objRemoteItem.GetByKey(strItem) Then
                    objRemoteItem.AllowUseTax = objMainItem.AllowUseTax
                    objRemoteItem.Block = objMainItem.Block
                    objRemoteItem.BuildingFloorRoom = objMainItem.BuildingFloorRoom
                    objRemoteItem.City = objMainItem.City
                    objRemoteItem.CostInflationAccount = objMainItem.CostInflationAccount
                    objRemoteItem.CostInflationOffsetAccount = objMainItem.CostInflationOffsetAccount
                    objRemoteItem.CostOfGoodsSold = objMainItem.CostOfGoodsSold
                    objRemoteItem.Country = objMainItem.Country
                    objRemoteItem.County = objMainItem.County
                    objRemoteItem.DecreaseGLAccount = objMainItem.DecreaseGLAccount
                    objRemoteItem.DecreasingAccount = objMainItem.DecreasingAccount
                    objRemoteItem.DropShip = objMainItem.DropShip
                    objRemoteItem.EUExpensesAccount = objMainItem.EUExpensesAccount
                    objRemoteItem.EUPurchaseCreditAcc = objMainItem.EUPurchaseCreditAcc
                    objRemoteItem.EURevenuesAccount = objMainItem.EURevenuesAccount
                    objRemoteItem.ExchangeRateDifferencesAccount = objMainItem.ExchangeRateDifferencesAccount
                    '  objRemoteItem.Excisable = objMainItem.Excisable
                    objRemoteItem.ExemptedCredits = objMainItem.ExemptedCredits
                    objRemoteItem.ExemptRevenuesAccount = objMainItem.ExemptRevenuesAccount
                    objRemoteItem.ExpenseAccount = objMainItem.ExpenseAccount
                    objRemoteItem.ExpenseOffsetingAct = objMainItem.ExpenseOffsetingAct
                    objRemoteItem.ExpensesClearingAccount = objMainItem.ExpensesClearingAccount
                    objRemoteItem.FederalTaxID = objMainItem.FederalTaxID
                    objRemoteItem.ForeignExpensesAccount = objMainItem.ForeignExpensesAccount
                    objRemoteItem.ForeignPurchaseCreditAcc = objMainItem.ForeignPurchaseCreditAcc
                    objRemoteItem.ForeignRevenuesAcc = objMainItem.ForeignRevenuesAcc
                    objRemoteItem.GoodsClearingAcc = objMainItem.GoodsClearingAcc
                    objRemoteItem.IncreaseGLAccount = objMainItem.IncreaseGLAccount
                    objRemoteItem.IncreasingAcc = objMainItem.IncreasingAcc
                    objRemoteItem.Location = objMainItem.Location
                    objRemoteItem.NegativeInventoryAdjustmentAccount = objMainItem.NegativeInventoryAdjustmentAccount
                    objRemoteItem.Nettable = objMainItem.Nettable
                    objRemoteItem.PriceDifferencesAccount = objMainItem.PriceDifferencesAccount
                    objRemoteItem.PurchaseAccount = objMainItem.PurchaseAccount
                    objRemoteItem.PurchaseCreditAcc = objMainItem.PurchaseCreditAcc
                    objRemoteItem.PurchaseOffsetAccount = objMainItem.PurchaseOffsetAccount
                    objRemoteItem.PurchaseReturningAccount = objMainItem.PurchaseReturningAccount
                    objRemoteItem.ReturningAccount = objMainItem.ReturningAccount
                    objRemoteItem.RevenuesAccount = objMainItem.RevenuesAccount
                    objRemoteItem.SalesCreditAcc = objMainItem.SalesCreditAcc
                    objRemoteItem.SalesCreditEUAcc = objMainItem.SalesCreditEUAcc
                    objRemoteItem.SalesCreditForeignAcc = objMainItem.SalesCreditForeignAcc
                    objRemoteItem.ShippedGoodsAccount = objMainItem.ShippedGoodsAccount
                    objRemoteItem.State = objMainItem.State
                    objRemoteItem.StockAccount = objMainItem.StockAccount
                    objRemoteItem.StockInflationAdjustAccount = objMainItem.StockInflationAdjustAccount
                    objRemoteItem.StockInflationOffsetAccount = objMainItem.StockInflationOffsetAccount
                    objRemoteItem.Street = objMainItem.Street
                    objRemoteItem.TaxGroup = objMainItem.TaxGroup
                    objRemoteItem.TransfersAcc = objMainItem.TransfersAcc
                    objRemoteItem.VarianceAccount = objMainItem.VarianceAccount
                    objRemoteItem.VATInRevenueAccount = objMainItem.VATInRevenueAccount
                    objRemoteItem.WarehouseName = objMainItem.WarehouseName
                    objRemoteItem.WarehouseCode = objMainItem.WarehouseCode
                    ' objRemoteItem.WHShipToName = objMainItem.ShippedGoodsAccount
                    objRemoteItem.WIPMaterialAccount = objMainItem.WIPMaterialAccount
                    objRemoteItem.WIPMaterialVarianceAccount = objMainItem.WIPMaterialVarianceAccount
                    objRemoteItem.ZipCode = objMainItem.ZipCode
                    If objRemoteItem.Update() <> 0 Then
                        'WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Failed to Accoutns Segementation : " & objRemoteItem.FormatCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        'WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Account Segemntation : " & objRemoteItem.Name & " Updated successfully", sPath)
                    End If

                Else
                    objRemoteItem.WarehouseCode = objMainItem.WarehouseCode
                    objRemoteItem.WarehouseName = objMainItem.WarehouseName
                    If objRemoteItem.Add <> 0 Then
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " :  Failed to add Warehosue : " & objMainItem.WarehouseCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Warehouse: " & objRemoteItem.WarehouseCode & " Added successfully", sPath)
                    End If
                End If
            End If
            'objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWarehouses)


            objMainRecset.MoveNext()
        Next
    End Sub

    Private Sub ExportShippingTypes()
        Dim FileName, strItem, strDflWhs As String
        Dim objMainItem, objRemoteItem As SAPbobsCOM.ShippingTypes
        Dim objMainRecset As SAPbobsCOM.Recordset
        objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRecset.DoQuery("Select TrnspCode from OSHP")
        For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
            strItem = objMainRecset.Fields.Item(0).Value
            objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oShippingTypes)
            If objMainItem.GetByKey(strItem) Then
                objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oShippingTypes)
                If objRemoteItem.GetByKey(strItem) Then
                    objRemoteItem.Name = objMainItem.Name
                    objRemoteItem.Website = objMainItem.Website
                    If objRemoteItem.Update() <> 0 Then
                        ' WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Failed to Accoutns Segementation : " & objRemoteItem.FormatCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        ' WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Account Segemntation : " & objRemoteItem.Name & " Updated successfully", sPath)
                    End If

                Else
                    objRemoteItem.Name = objMainItem.Name
                    objRemoteItem.Website = objMainItem.Website
                    If objRemoteItem.Add <> 0 Then
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " :  Failed to add Shipping type : " & objMainItem.Name & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " :Shipping Type : " & objRemoteItem.Name & " Added successfully", sPath)
                    End If
                End If
            End If

            objMainRecset.MoveNext()
        Next
    End Sub

    Private Sub ExportManufactures()
        Dim FileName, strItem, strDflWhs As String
        Dim objMainItem, objRemoteItem As SAPbobsCOM.Manufacturers
        Dim objMainRecset, objRemoteRec As SAPbobsCOM.Recordset
        objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRecset.DoQuery("Select firmCode from OMRC")
        For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
            strItem = objMainRecset.Fields.Item(0).Value
            objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oManufacturers)
            If objMainItem.GetByKey(strItem) Then
                objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oManufacturers)
                objRemoteRec = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRemoteRec.DoQuery("Select FirmCode from OMRC where FirmName='" & objMainItem.ManufacturerName & "'")
                If objRemoteRec.RecordCount > 0 Then
                    strItem = objRemoteRec.Fields.Item(0).Value
                    If objRemoteItem.GetByKey(strItem) Then
                        objRemoteItem.ManufacturerName = objMainItem.ManufacturerName
                        If objRemoteItem.Update() <> 0 Then
                            ' WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Failed to Accoutns Segementation : " & objRemoteItem.FormatCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                        Else
                            ' WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Account Segemntation : " & objRemoteItem.Name & " Updated successfully", sPath)
                        End If
                    End If
                Else
                    objRemoteItem.ManufacturerName = objMainItem.ManufacturerName
                    If objRemoteItem.Add <> 0 Then
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " :  Failed to add Manufacture : " & objMainItem.ManufacturerName & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Manufacture : " & objRemoteItem.ManufacturerName & " Added successfully", sPath)
                    End If
                End If

            End If

            objMainRecset.MoveNext()
        Next
    End Sub

    Private Sub ExportPackagingType()
        Dim FileName, strItem, strDflWhs As String
        Dim objMainItem, objRemoteItem As SAPbobsCOM.PackagesTypes
        Dim objMainRecset, objRemoteRec As SAPbobsCOM.Recordset
        objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRecset.DoQuery("Select PkgCode from OPKG")
        For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
            strItem = objMainRecset.Fields.Item(0).Value
            objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPackagesTypes)
            If objMainItem.GetByKey(strItem) Then
                objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPackagesTypes)
                objRemoteRec = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRemoteRec.DoQuery("Select PkgCode from OPKG where PkgCode='" & objMainItem.Code & "'")
                If objRemoteRec.RecordCount > 0 Then
                    strItem = objRemoteRec.Fields.Item(0).Value
                    If objRemoteItem.GetByKey(strItem) Then
                        objRemoteItem.Type = objMainItem.Type
                        If objRemoteItem.Update() <> 0 Then
                            ' WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Failed to Accoutns Segementation : " & objRemoteItem.FormatCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                        Else
                            ' WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Account Segemntation : " & objRemoteItem.Name & " Updated successfully", sPath)
                        End If
                    End If
                Else
                    objRemoteItem.Type = objMainItem.Type
                    If objRemoteItem.Add <> 0 Then
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " :  Failed to add Packing Type : " & objMainItem.Type & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Packing Type : " & objMainItem.Type & " Added successfully", sPath)
                    End If
                End If

            End If

            objMainRecset.MoveNext()
        Next
    End Sub

    Private Sub ExportLengthandUOM()
        Dim FileName, strItem, strDflWhs As String
        Dim objMainItem, objRemoteItem As SAPbobsCOM.LengthMeasures
        Dim objMainRecset, objRemoteRec As SAPbobsCOM.Recordset
        objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRecset.DoQuery("Select UnitCode from OLGT")
        For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
            strItem = objMainRecset.Fields.Item(0).Value
            objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oLengthMeasures)
            If objMainItem.GetByKey(CInt(strItem)) Then
                objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oLengthMeasures)
                objRemoteRec = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRemoteRec.DoQuery("Select UnitCode from OLGT where UnitCode=" & objMainItem.UnitCode)
                If objRemoteRec.RecordCount > 0 Then
                    strItem = objRemoteRec.Fields.Item(0).Value
                    If objRemoteItem.GetByKey(CInt(strItem)) Then
                        objRemoteItem.UnitCodeforQuantityDisplay = objMainItem.UnitCodeforQuantityDisplay
                        objRemoteItem.UnitDisplay = objMainItem.UnitDisplay
                        objRemoteItem.UnitLengthinmm = objMainItem.UnitLengthinmm
                        objRemoteItem.UnitName = objMainItem.UnitName


                        If objRemoteItem.Update() <> 0 Then
                            ' WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Failed to Accoutns Segementation : " & objRemoteItem.FormatCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                        Else
                            ' WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Account Segemntation : " & objRemoteItem.Name & " Updated successfully", sPath)
                        End If
                    End If
                Else
                    objRemoteItem.UnitCodeforQuantityDisplay = objMainItem.UnitCodeforQuantityDisplay
                    objRemoteItem.UnitDisplay = objMainItem.UnitDisplay
                    objRemoteItem.UnitLengthinmm = objMainItem.UnitLengthinmm
                    objRemoteItem.UnitName = objMainItem.UnitName

                    If objRemoteItem.Add <> 0 Then
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " :  Failed to add Length and UoM : " & objMainItem.UnitName & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Lenght and UoM : " & objMainItem.UnitName & " Added successfully", sPath)
                    End If
                End If

            End If

            objMainRecset.MoveNext()
        Next

        ExportWeightUOM()
    End Sub

    Private Sub ExportWeightUOM()
        Dim FileName, strItem, strDflWhs As String
        Dim objMainItem, objRemoteItem As SAPbobsCOM.WeightMeasures
        Dim objMainRecset, objRemoteRec As SAPbobsCOM.Recordset
        objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRecset.DoQuery("Select UnitCode from owgt")
        For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
            strItem = objMainRecset.Fields.Item(0).Value
            objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWeightMeasures)
            If objMainItem.GetByKey(CInt(strItem)) Then
                objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oWeightMeasures)
                objRemoteRec = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRemoteRec.DoQuery("Select UnitCode from owgt where UnitCode=" & objMainItem.UnitCode)
                If objRemoteRec.RecordCount > 0 Then
                    strItem = objRemoteRec.Fields.Item(0).Value
                    If objRemoteItem.GetByKey(CInt(strItem)) Then
                        objRemoteItem.UnitWeightinmg = objMainItem.UnitWeightinmg
                        objRemoteItem.UnitDisplay = objMainItem.UnitDisplay

                        objRemoteItem.UnitName = objMainItem.UnitName


                        If objRemoteItem.Update() <> 0 Then
                            ' WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Failed to Accoutns Segementation : " & objRemoteItem.FormatCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                        Else
                            ' WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Account Segemntation : " & objRemoteItem.Name & " Updated successfully", sPath)
                        End If
                    End If
                Else
                    objRemoteItem.UnitWeightinmg = objMainItem.UnitWeightinmg
                    objRemoteItem.UnitDisplay = objMainItem.UnitDisplay
                    objRemoteItem.UnitName = objMainItem.UnitName

                    If objRemoteItem.Add <> 0 Then
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " :  Failed to add Weight UoM : " & objMainItem.UnitName & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Weight UoM : " & objMainItem.UnitName & " Added successfully", sPath)
                    End If
                End If
            End If

            objMainRecset.MoveNext()
        Next
    End Sub
#End Region
#Region "Check the Filepaths"
    Private Function ValidateFilePaths() As Boolean
        Dim strMessage, strpath, strFilename As String
        strpath = strErrorLogPath ' System.Windows.Forms.Application.StartupPath

        If Directory.Exists(strpath) = False Then
            '  strMessage = "Error Log file direcory does not exists. check the connectionInfo.ini"
            '  MsgBox(strMessage)
            ' Return False
        End If


        If Directory.Exists(strpath) = False Then
            Directory.CreateDirectory(strpath)
            'strMessage = ("Export Directory does not exist. Check the connectionInfo.ini")
            'WriteErrorlog(strMessage, strpath)
            'MsgBox(strMessage)
            'Return False
        End If



        strFilename = Now.ToLongDateString



        If Directory.Exists(strExportFilePaty) = False Then
            strMessage = ("Export Directory does not exist. Check the connectionInfo.ini")
            WriteErrorlog(strMessage, strpath)
            ' MsgBox(strMessage)
            Return False
        End If
        'If strFileStart = "" Then
        '    strMessage = "Export File Name is missing"
        '    MsgBox(strMessage)
        '    Return False
        'End If

        Return True
    End Function
#End Region

    Private Sub btnReferesh_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReferesh.Click
        Try
            If txtServerName.Text = "" Then
                MsgBox("Enter Server")
                Exit Sub
            End If
            If txtDBUserName.Text = "" Then
                MsgBox("Enter UserName")
                Exit Sub
            End If
            If txtServerPwd.Text = "" Then
                MsgBox("Enter Password")
                Exit Sub
            End If
            strLocalServertype = cmbServertype.Text
            oCompany = New SAPbobsCOM.Company
            oCompany.Server = txtServerName.Text
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
            If strLocalServertype = "2005" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005
            Else
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
            End If
            oCompany.DbUserName = txtDBUserName.Text
            oCompany.DbPassword = txtServerPwd.Text
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim lErrCode As Long
            Dim sErrMsg As String
            oRecordSet = Me.oCompany.GetCompanyList
            Me.oCompany.GetLastError(lErrCode, sErrMsg)
            If lErrCode <> 0 Then
                MsgBox(sErrMsg)
                Exit Sub
            Else
                cmbCompanyDB.Items.Clear()
                Do Until oRecordSet.EoF = True
                    cmbCompanyDB.Items.Add(oRecordSet.Fields.Item(0).Value)
                    oRecordSet.MoveNext()
                Loop
                btnConnect.Enabled = True
            End If
        Catch ex As Exception
            MsgBox("Connection to Server =" & txtServerName.Text & ".  Failed")
            Exit Sub
        End Try
    End Sub

#Region "Write into ErrorLog File"
    Private Sub WriteErrorHeader(ByVal apath As String, ByVal strMessage As String)
        Dim aSw As System.IO.StreamWriter
        Dim aMessage As String
        aMessage = Now.ToLocalTime & "---" & strMessage
        If File.Exists(apath) Then
        End If
        aSw = New StreamWriter(apath, True)
        aSw.WriteLine(aMessage)
        aSw.Flush()
        aSw.Close()
    End Sub
    Private Sub WriteErrorlog(ByVal aMessage As String, ByVal aPath As String)
        aPath = strErrorFileName
        Dim aSw As System.IO.StreamWriter
        If File.Exists(aPath) Then
        End If
        aSw = New StreamWriter(aPath, True)
        aMessage = Now.ToLocalTime & "---" & aMessage
        aSw.WriteLine(aMessage)
        aSw.Flush()
        aSw.Close()
    End Sub
#End Region

#Region "exportSO"
    Private Function GetSalesOrders() As String
        Dim oSalesRec As SAPbobsCOM.Recordset
        oSalesRec = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim strsql As String
        strsql = "SELECT convert(nvarchar(12),dbo.ORDR.DocNum) AS DocumentNoV1, dbo.ORDR.CardCode AS CustomerCode, "
        strsql = strsql & "       convert(nvarchar(10),dbo.ORDR.DocDate,112) AS DocumentDate, '' "
        strsql = strsql & " AS DeliveryDate, coalesce(dbo.ORDR.NumAtCard,'') AS CustomerRef, "
        strsql = strsql & " convert(nvarchar(12),dbo.RDR1.LineNum) as LineNum, dbo.RDR1.ItemCode, "
        strsql = strsql & " dbo.RDR1.Dscription AS ItemDesc, "
        strsql = strsql & " coalesce(dbo.[@MORPHI].Name,'') AS Presentation, convert(nvarchar(12),convert(int,round(dbo.RDR1.Quantity,0))), "
        strsql = strsql & " '' AS PickedQuantity, '' "
        strsql = strsql & " AS ExpDate, '' AS Batch FROM   dbo.ORDR INNER JOIN  dbo.RDR1 ON dbo.ORDR.DocEntry = dbo.RDR1.DocEntry left outer JOIN "
        strsql = strsql & " dbo.[@MORPHI] ON dbo.RDR1.U_TYPE = dbo.[@MORPHI].Code where u_mantisst='N'ORDER BY dbo.DocumentNoV1, dbo.RDR1.LineNum"
        oSalesRec.DoQuery(strsql)
        If oSalesRec.RecordCount > 0 Then
            Return strsql
        Else
            Return ""
        End If
        Return ""
    End Function
#End Region

    Private Function ValidateDocSeries(ByVal aDocEntry As Integer, ByVal aChoice As String) As Boolean
        Dim otempRec1 As SAPbobsCOM.Recordset
        Dim strSeries As Integer
        Dim strSql As String
        otempRec1 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If aChoice = "Incoming" Then
            strSql = "SELECT T0.[Series], T1.[Remark] FROM ORCT T0  INNER JOIN NNM1 T1 ON T0.[Series] = T1.[Series] WHERE T1.[ObjectCode] =24 and T0.DocEntry=" & aDocEntry
        Else
            strSql = "SELECT T0.[Series], T1.[Remark] FROM OVPM T0  INNER JOIN NNM1 T1 ON T0.[Series] = T1.[Series] WHERE T1.[ObjectCode] =46 and T0.DocEntry=" & aDocEntry
        End If
        otempRec1.DoQuery(strSql)
        If otempRec1.RecordCount > 0 Then
            If otempRec1.Fields.Item(1).Value = "P1" Then
                Return True
            Else
                Return False
            End If
        End If
        Return False
    End Function

    Private Function ValidateDocSeries_Deposit(ByVal aDocEntry As Integer, ByVal aChoice As String) As Boolean
        Dim otempRec1 As SAPbobsCOM.Recordset
        Dim strSeries As Integer
        Dim strSql As String
        otempRec1 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strSql = "SELECT T0.[Series], T1.[Remark] FROM ODPS T0 INNER JOIN NNM1 T1 ON T0.[Series] = T1.[Series]  WHERE T1.[ObjectCode] ='" & aChoice & "' and  T0.DeposID=" & aDocEntry
        otempRec1.DoQuery(strSql)
        If otempRec1.RecordCount > 0 Then
            If otempRec1.Fields.Item(1).Value = "P1" Then
                Return True
            Else
                Return False
            End If
        End If
        Return False
    End Function
    Private Function ValidateDocDate(ByVal aDocEntry As Integer, ByVal aChoice As String, ByVal aDate As Date) As Boolean
        Dim otempRec1 As SAPbobsCOM.Recordset
        Try
            Dim strSeries As Integer
            Dim strSql As String
            otempRec1 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If aChoice = "Incoming" Then
                '  strSql = "SELECT * from [@Z_DBSYN] where U_Active='Y' and  Code='" & aChoice & "' and '" & aDate.ToString("yyyy-MM-dd") & "' between (isnull(U_FrmDate,getdate()) and isnull(U_ToDate,getdate())"
                strSql = "SELECT * from [@Z_DBSYN] where isnull(U_Active,'N')='Y' and  Code='" & aChoice & "' and ('" & aDate.ToString("yyyy-MM-dd") & "' between (isnull(U_FrmDate,'" & aDate.ToString("yyyy-MM-dd") & "')) and isnull(U_ToDate,'" & aDate.ToString("yyyy-MM-dd") & "'))"
            Else
                strSql = "SELECT * from [@Z_DBSYN] where isnull(U_Active,'N')='Y' and  Code='" & aChoice & "' and ('" & aDate.ToString("yyyy-MM-dd") & "' between (isnull(U_FrmDate,'" & aDate.ToString("yyyy-MM-dd") & "')) and isnull(U_ToDate,'" & aDate.ToString("yyyy-MM-dd") & "'))"
            End If
            otempRec1.DoQuery(strSql)
            If otempRec1.RecordCount > 0 Then
                Return True
            Else
                Return False
            End If
            Return False
        Catch ex As Exception

        Finally
            If Not IsNothing(otempRec1) Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(otempRec1)
                GC.Collect()
            End If
        End Try
    End Function

#Region "Export Incoming Payment Doocuments"
    Private Sub ExportIncomingPaymentDocument(Optional ByVal aDocentry As String = "")
        Dim oTempRec, objremoteRec As SAPbobsCOM.Recordset
        oTempRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objremoteRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objremoteRec.DoQuery("SELECT * from [@Z_DBSYN] where isnull(U_Active,'N')='Y' and Code='24'")
        Dim dtFrom, dtTo As Date
        If objremoteRec.RecordCount <= 0 Then
            Exit Sub
        Else
            dtFrom = objremoteRec.Fields.Item("U_FrmDate").Value
            dtTo = objremoteRec.Fields.Item("U_ToDate").Value
        End If
        Dim dtString As String
        If Year(dtFrom) = 1899 Then
            dtString = " 1=1"
        Else
            dtString = " T0.DocDate >='" & dtFrom.ToString("yyyy-MM-dd") & "'"

        End If

        If Year(dtTo) = 1899 Then
            dtString = dtString & " and  1=1"
        Else
            dtString = dtString & " and T0.DocDate <='" & dtTo.ToString("yyyy-MM-dd") & "'"

        End If

        'Post Open Documents
        If aDocentry <> "" Then
            '  oTempRec.DoQuery("SELECT T0.DocEntry,T0.Series ,T0.DocType FROM ORCT T0  inner Join  OCRD T1 on T1.CardCode=T0.CardCode WHERE T0.DocEntry=" & aDocentry & " and T0.DocDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and  T0.DocType<>'A' and T1.[U_sequencetype] ='P1' and  isnull(T0.U_Export,'N')='N' and (" & dtString & ") order by T0.DocDate,T0.DocEntry")
            oTempRec.DoQuery("SELECT T0.DocEntry,T0.Series ,T0.DocType FROM ORCT T0  inner Join  OCRD T1 on T1.CardCode=T0.CardCode WHERE T0.DocEntry=" & aDocentry & " and T0.DocDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and  T0.DocType<>'A' and T1.[U_sequencetype] ='P1' and  isnull(T0.U_Export,'N')='N' and (" & dtString & ") order by T0.DocDate,T0.DocEntry")
        Else
            oTempRec.DoQuery("SELECT T0.DocEntry,T0.Series ,T0.DocType FROM ORCT T0  inner Join  OCRD T1 on T1.CardCode=T0.CardCode WHERE T0.DocDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and  T0.DocType<>'A' and T1.[U_sequencetype] ='P1' and  isnull(T0.U_Export,'N')='N' and (" & dtString & ") order by T0.DocDate,T0.DocEntry")
        End If

        '   oTempRec.DoQuery("SELECT T0.DocEntry,T0.Series ,T0.DocType FROM ORCT T0  inner Join  OCRD T1 on T1.CardCode=T0.CardCode WHERE T0.DocDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and  T0.DocType<>'A' and T1.[U_sequencetype] ='P1' and  isnull(T0.U_Export,'N')='N' and (" & dtString & ") order by T0.DocDate,T0.DocEntry")
        For intRow As Integer = 0 To oTempRec.RecordCount - 1
            If oTempRec.Fields.Item(2).Value = "A" Then
                If ValidateDocSeries(oTempRec.Fields.Item(0).Value, "Incoming") = True Then
                    ExportIncomingPayment(oCompany, oTempRec.Fields.Item(0).Value)
                    '  ExportIncomingPayment_Cancelled(oCompany, oTempRec.Fields.Item(0).Value)
                End If
            Else
                ExportIncomingPayment(oCompany, oTempRec.Fields.Item(0).Value)
                ' ExportIncomingPayment_Cancelled(oCompany, oTempRec.Fields.Item(0).Value)
            End If

            oTempRec.MoveNext()
        Next
        If aDocentry <> "" Then
            oTempRec.DoQuery("SELECT T0.DocEntry,T0.Series ,T0.DocType FROM ORCT T0   WHERE T0.DocEntry=" & aDocentry & " and T0.DocDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and  T0.DocType='A' and  isnull(T0.U_Export,'N')='N'  and (" & dtString & ")  order by T0.DocDate,T0.DocEntry")
        Else
            oTempRec.DoQuery("SELECT T0.DocEntry,T0.Series ,T0.DocType FROM ORCT T0   WHERE T0.DocDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and  T0.DocType='A' and  isnull(T0.U_Export,'N')='N'  and (" & dtString & ")  order by T0.DocDate,T0.DocEntry")
        End If
        'oTempRec.DoQuery("SELECT T0.DocEntry,T0.Series ,T0.DocType FROM ORCT T0   WHERE T0.DocDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and  T0.DocType='A' and  isnull(T0.U_Export,'N')='N'  and (" & dtString & ")  order by T0.DocDate,T0.DocEntry")
        For intRow As Integer = 0 To oTempRec.RecordCount - 1
            If oTempRec.Fields.Item(2).Value = "A" Then
                If ValidateDocSeries(oTempRec.Fields.Item(0).Value, "Incoming") = True Then
                    ExportIncomingPayment(oCompany, oTempRec.Fields.Item(0).Value)
                    '   ExportIncomingPayment_Cancelled(oCompany, oTempRec.Fields.Item(0).Value)
                End If
            Else
                ExportIncomingPayment(oCompany, oTempRec.Fields.Item(0).Value)
                '  ExportIncomingPayment_Cancelled(oCompany, oTempRec.Fields.Item(0).Value)
            End If

            oTempRec.MoveNext()
        Next


        'Post Cancelled Documents
        If aDocentry <> "" Then
            '  oTempRec.DoQuery("SELECT T0.DocEntry,T0.Series ,T0.DocType FROM ORCT T0  inner Join  OCRD T1 on T1.CardCode=T0.CardCode WHERE T0.DocEntry=" & aDocentry & " and T0.DocDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and  T0.DocType<>'A' and T1.[U_sequencetype] ='P1' and  isnull(T0.U_Export,'N')='N' and (" & dtString & ") order by T0.DocDate,T0.DocEntry")
            oTempRec.DoQuery("SELECT T0.DocEntry,T0.Series ,T0.DocType FROM ORCT T0  inner Join  OCRD T1 on T1.CardCode=T0.CardCode WHERE T0.DocEntry=" & aDocentry & " and T0.DocDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and  T0.DocType<>'A' and T1.[U_sequencetype] ='P1' and (" & dtString & ") order by T0.DocDate,T0.DocEntry")
        Else
            oTempRec.DoQuery("SELECT T0.DocEntry,T0.Series ,T0.DocType FROM ORCT T0  inner Join  OCRD T1 on T1.CardCode=T0.CardCode WHERE T0.DocDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and  T0.DocType<>'A' and T1.[U_sequencetype] ='P1'  and (" & dtString & ") order by T0.DocDate,T0.DocEntry")
        End If

        '   oTempRec.DoQuery("SELECT T0.DocEntry,T0.Series ,T0.DocType FROM ORCT T0  inner Join  OCRD T1 on T1.CardCode=T0.CardCode WHERE T0.DocDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and  T0.DocType<>'A' and T1.[U_sequencetype] ='P1' and  isnull(T0.U_Export,'N')='N' and (" & dtString & ") order by T0.DocDate,T0.DocEntry")
        For intRow As Integer = 0 To oTempRec.RecordCount - 1
            If oTempRec.Fields.Item(2).Value = "A" Then
                If ValidateDocSeries(oTempRec.Fields.Item(0).Value, "Incoming") = True Then
                    '  ExportIncomingPayment(oCompany, oTempRec.Fields.Item(0).Value)
                    ExportIncomingPayment_Cancelled(oCompany, oTempRec.Fields.Item(0).Value)
                End If
            Else
                '  ExportIncomingPayment(oCompany, oTempRec.Fields.Item(0).Value)
                ExportIncomingPayment_Cancelled(oCompany, oTempRec.Fields.Item(0).Value)
            End If

            oTempRec.MoveNext()
        Next
        If aDocentry <> "" Then
            oTempRec.DoQuery("SELECT T0.DocEntry,T0.Series ,T0.DocType FROM ORCT T0   WHERE T0.DocEntry=" & aDocentry & " and T0.DocDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and  T0.DocType='A'  and (" & dtString & ")  order by T0.DocDate,T0.DocEntry")
        Else
            oTempRec.DoQuery("SELECT T0.DocEntry,T0.Series ,T0.DocType FROM ORCT T0   WHERE T0.DocDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and  T0.DocType='A'  and (" & dtString & ")  order by T0.DocDate,T0.DocEntry")
        End If
        'oTempRec.DoQuery("SELECT T0.DocEntry,T0.Series ,T0.DocType FROM ORCT T0   WHERE T0.DocDate='" & dtPostingDate.ToString("yyyy-MM-dd") & "' and  T0.DocType='A' and  isnull(T0.U_Export,'N')='N'  and (" & dtString & ")  order by T0.DocDate,T0.DocEntry")
        For intRow As Integer = 0 To oTempRec.RecordCount - 1
            If oTempRec.Fields.Item(2).Value = "A" Then
                If ValidateDocSeries(oTempRec.Fields.Item(0).Value, "Incoming") = True Then
                    '  ExportIncomingPayment(oCompany, oTempRec.Fields.Item(0).Value)
                    ExportIncomingPayment_Cancelled(oCompany, oTempRec.Fields.Item(0).Value)
                End If
            Else
                ' ExportIncomingPayment(oCompany, oTempRec.Fields.Item(0).Value)
                ExportIncomingPayment_Cancelled(oCompany, oTempRec.Fields.Item(0).Value)
            End If

            oTempRec.MoveNext()
        Next
    End Sub
#End Region

#Region "Export InComing Payment Details"
    Private Sub ExportIncomingPayment(ByVal aCompany As SAPbobsCOM.Company, ByVal aDocEntry As Integer)
        Dim objMainDoc, ObjRemoteDoc As SAPbobsCOM.Payments
        Dim objremoteRec, objRemoteREC1 As SAPbobsCOM.Recordset
        Dim objRec1 As SAPbobsCOM.Recordset
        Dim strRemoteDocNumber As Integer
        Try
            objremoteRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRemoteREC1 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRec1 = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objremoteRec.DoQuery("Select DocEntry from ORCT where DocEntry=" & aDocEntry)
            For intRemoteloop As Integer = 0 To objremoteRec.RecordCount - 1
                Try
                    ObjRemoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
                    objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
                    If ObjRemoteDoc.GetByKey(Convert.ToInt32(objremoteRec.Fields.Item(0).Value)) Then
                        If ValidateDocDate(ObjRemoteDoc.DocEntry, "24", ObjRemoteDoc.DocDate) = True Then
                            If 1 = 2 Then ' ObjRemoteDoc.Cancelled = SAPbobsCOM.BoYesNoEnum.tYES Then
                                If CheckCanceledInvoice(ObjRemoteDoc.DocEntry, ObjRemoteDoc.DocNum, objMainCompany.CompanyName, "ORCT") <> 0 Then
                                    If objMainDoc.GetByKey(CheckCanceledInvoice(ObjRemoteDoc.DocEntry, ObjRemoteDoc.DocNum, objMainCompany.CompanyName, "ORCT")) Then
                                        If objMainDoc.Cancel <> 0 Then
                                            '  MsgBox(objRemoteCompany.GetLastErrorDescription)
                                        Else
                                            '  MsgBox("Canceld")
                                        End If
                                    End If
                                End If
                            Else
                                If checkDuplicateInvoice(ObjRemoteDoc.DocEntry, ObjRemoteDoc.DocNum, objMainCompany.CompanyName, "ORCT") = False Then
                                    UpdateExchangerate_Main("USD", ObjRemoteDoc.DocDate, objMainCompany)
                                    objMainDoc.Address = ObjRemoteDoc.Address
                                    objMainDoc.ApplyVAT = ObjRemoteDoc.ApplyVAT
                                    objMainDoc.BankAccount = ObjRemoteDoc.BankAccount
                                    If ObjRemoteDoc.BankChargeAmountInFC > 0 Then
                                        If ObjRemoteDoc.DocRate > 0 Then
                                            objMainDoc.BankChargeAmount = Math.Round(ObjRemoteDoc.BankChargeAmount / ObjRemoteDoc.DocRate, 2)
                                        Else
                                            objMainDoc.BankChargeAmount = ObjRemoteDoc.BankChargeAmount
                                        End If
                                    End If
                                    ' objMainDoc.BankChargeAmount = ObjRemoteDoc.BankChargeAmount
                                    objMainDoc.BankCode = ObjRemoteDoc.BankCode
                                    objMainDoc.BillOfExchangeAgent = ObjRemoteDoc.BillOfExchangeAgent
                                    objMainDoc.BillOfExchangeAgent = ObjRemoteDoc.BillOfExchangeAmount
                                    objMainDoc.BillofExchangeStatus = ObjRemoteDoc.BillofExchangeStatus
                                    objMainDoc.BoeAccount = ObjRemoteDoc.BoeAccount
                                    objMainDoc.CardCode = ObjRemoteDoc.CardCode
                                    objMainDoc.CardName = ObjRemoteDoc.CardName
                                    objMainDoc.CashAccount = ObjRemoteDoc.CashAccount
                                    objMainDoc.DocCurrency = ObjRemoteDoc.DocCurrency
                                    objMainDoc.LocalCurrency = ObjRemoteDoc.LocalCurrency
                                    If ObjRemoteDoc.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tYES Then
                                        objMainDoc.CashSum = ObjRemoteDoc.CashSum
                                    Else
                                        If ObjRemoteDoc.DocRate > 0 Then
                                            objMainDoc.CashSum = Math.Round(ObjRemoteDoc.CashSum / ObjRemoteDoc.DocRate, 2)
                                        Else
                                            objMainDoc.CashSum = ObjRemoteDoc.CashSum
                                        End If
                                    End If
                                    objMainDoc.CheckAccount = ObjRemoteDoc.CheckAccount
                                    objMainDoc.HandWritten = SAPbobsCOM.BoYesNoEnum.tYES
                                    objMainDoc.DocRate = ObjRemoteDoc.DocRate
                                    objRec1.DoQuery("Select * from NNM1 where ObjectCode='24' and SeriesName='Manual'")
                                    objMainDoc.Series = -1 'objRec1.Fields.Item("Series").Value
                                    objMainDoc.DocNum = ObjRemoteDoc.DocNum
                                    objMainDoc.UserFields.Fields.Item("U_Import").Value = "Y"
                                    objMainDoc.UserFields.Fields.Item("U_BaseEntry").Value = ObjRemoteDoc.DocEntry
                                    objMainDoc.UserFields.Fields.Item("U_BaseNum").Value = ObjRemoteDoc.DocNum
                                    objMainDoc.UserFields.Fields.Item("U_Branch").Value = objMainCompany.CompanyName
                                    If ObjRemoteDoc.Checks.CheckSum > 0 Then
                                        For intCheck As Integer = 0 To ObjRemoteDoc.Checks.Count - 1
                                            If intCheck > 0 Then
                                                objMainDoc.Checks.Add()
                                                objMainDoc.Checks.SetCurrentLine(intCheck)
                                            End If
                                            ObjRemoteDoc.Checks.SetCurrentLine(intCheck)
                                            objMainDoc.Checks.AccounttNum = ObjRemoteDoc.Checks.AccounttNum
                                            objMainDoc.Checks.BankCode = ObjRemoteDoc.Checks.BankCode
                                            objMainDoc.Checks.Branch = ObjRemoteDoc.Checks.Branch
                                            objMainDoc.Checks.CheckAccount = ObjRemoteDoc.Checks.CheckAccount
                                            objMainDoc.Checks.CheckAccount = ObjRemoteDoc.Checks.CheckNumber
                                            objMainDoc.Checks.CheckSum = ObjRemoteDoc.Checks.CheckSum
                                            objMainDoc.Checks.CountryCode = ObjRemoteDoc.Checks.CountryCode
                                            objMainDoc.Checks.Details = ObjRemoteDoc.Checks.Details
                                            objMainDoc.Checks.DueDate = ObjRemoteDoc.Checks.DueDate
                                            objMainDoc.Checks.ManualCheck = ObjRemoteDoc.Checks.ManualCheck
                                            objMainDoc.Checks.Trnsfrable = ObjRemoteDoc.Checks.Trnsfrable
                                            '    MsgBox(ObjRemoteDoc.Checks.CheckSum)
                                        Next
                                    End If
                                    objMainDoc.ContactPersonCode = ObjRemoteDoc.ContactPersonCode
                                    objMainDoc.CounterReference = ObjRemoteDoc.CounterReference
                                    If ObjRemoteDoc.CreditCards.CreditSum > 0 Then
                                        objMainDoc.CreditCards.AdditionalPaymentSum = ObjRemoteDoc.CreditCards.AdditionalPaymentSum
                                        objMainDoc.CreditCards.CardValidUntil = ObjRemoteDoc.CreditCards.CardValidUntil
                                        objMainDoc.CreditCards.ConfirmationNum = ObjRemoteDoc.CreditCards.ConfirmationNum
                                        objMainDoc.CreditCards.CreditAcct = ObjRemoteDoc.CreditCards.CreditAcct
                                        objMainDoc.CreditCards.CreditCard = ObjRemoteDoc.CreditCards.CreditCard
                                        objMainDoc.CreditCards.CreditCardNumber = ObjRemoteDoc.CreditCards.CreditCardNumber
                                        objMainDoc.CreditCards.CreditSum = ObjRemoteDoc.CreditCards.CreditSum
                                        objMainDoc.CreditCards.CreditType = ObjRemoteDoc.CreditCards.CreditType
                                        objMainDoc.CreditCards.FirstPaymentDue = ObjRemoteDoc.CreditCards.FirstPaymentDue
                                        objMainDoc.CreditCards.FirstPaymentSum = ObjRemoteDoc.CreditCards.FirstPaymentSum
                                        objMainDoc.CreditCards.NumOfCreditPayments = ObjRemoteDoc.CreditCards.NumOfCreditPayments
                                        objMainDoc.CreditCards.NumOfPayments = ObjRemoteDoc.CreditCards.NumOfPayments
                                        objMainDoc.CreditCards.OwnerIdNum = ObjRemoteDoc.CreditCards.OwnerIdNum
                                        objMainDoc.CreditCards.OwnerPhone = ObjRemoteDoc.CreditCards.OwnerPhone
                                        objMainDoc.CreditCards.SplitPayments = ObjRemoteDoc.CreditCards.SplitPayments
                                        objMainDoc.CreditCards.PaymentMethodCode = ObjRemoteDoc.CreditCards.PaymentMethodCode
                                    End If
                                    objMainDoc.DeductionPercent = ObjRemoteDoc.DeductionPercent
                                    objMainDoc.DeductionSum = ObjRemoteDoc.DeductionSum
                                    objMainDoc.DocCurrency = ObjRemoteDoc.DocCurrency
                                    objMainDoc.DocRate = ObjRemoteDoc.DocRate
                                    objMainDoc.DocTypte = ObjRemoteDoc.DocTypte
                                    objMainDoc.DocDate = ObjRemoteDoc.DocDate
                                    objMainDoc.DueDate = ObjRemoteDoc.DueDate
                                    objMainDoc.DocObjectCode = ObjRemoteDoc.DocObjectCode
                                    objMainDoc.IsPayToBank = ObjRemoteDoc.IsPayToBank
                                    objMainDoc.LocalCurrency = ObjRemoteDoc.LocalCurrency

                                    'objMainDoc.HandWritten = ObjRemoteDoc.HandWritten
                                    Dim intRefdocentry, intRemoteDocEntry As Integer

                                    For intInvoice As Integer = 0 To ObjRemoteDoc.Invoices.Count - 1
                                        If ObjRemoteDoc.Invoices.DocEntry > 0 Then
                                            If intInvoice > 0 Then
                                                objMainDoc.Invoices.Add()
                                                objMainDoc.Invoices.SetCurrentLine(intInvoice)
                                            End If
                                            ObjRemoteDoc.Invoices.SetCurrentLine(intInvoice)

                                            objMainDoc.Invoices.AppliedFC = ObjRemoteDoc.Invoices.AppliedFC
                                            objMainDoc.Invoices.SumApplied = ObjRemoteDoc.Invoices.SumApplied
                                            objMainDoc.Invoices.DiscountPercent = ObjRemoteDoc.Invoices.DiscountPercent
                                            objMainDoc.Invoices.DistributionRule = ObjRemoteDoc.Invoices.DistributionRule
                                            intRemoteDocEntry = ObjRemoteDoc.Invoices.DocEntry
                                            If ObjRemoteDoc.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice Then 'Invoice
                                                intRefdocentry = getbaseDocument_IncomingPayment(intRemoteDocEntry, strRemoteDocNumber, SAPbobsCOM.BoRcptInvTypes.it_Invoice, objMainCompany.CompanyName)
                                            ElseIf ObjRemoteDoc.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_CredItnote Then 'Credit Note
                                                objMainDoc.Invoices.AppliedFC = ObjRemoteDoc.Invoices.AppliedFC * -1
                                                objMainDoc.Invoices.SumApplied = ObjRemoteDoc.Invoices.SumApplied * -1
                                                intRefdocentry = getbaseDocument_IncomingPayment(intRemoteDocEntry, strRemoteDocNumber, SAPbobsCOM.BoRcptInvTypes.it_CredItnote, objMainCompany.CompanyName)
                                            ElseIf ObjRemoteDoc.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Return Then 'Sales Return
                                                intRefdocentry = getbaseDocument_IncomingPayment(intRemoteDocEntry, strRemoteDocNumber, SAPbobsCOM.BoRcptInvTypes.it_Return, objMainCompany.CompanyName)
                                            ElseIf ObjRemoteDoc.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_TaxInvoice Then 'Delivery
                                                intRefdocentry = getbaseDocument_IncomingPayment(intRemoteDocEntry, strRemoteDocNumber, SAPbobsCOM.BoRcptInvTypes.it_TaxInvoice, objMainCompany.CompanyName)
                                            ElseIf ObjRemoteDoc.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_JournalEntry Then 'Journal Entry
                                                intRefdocentry = getbaseDocument_IncomingPayment(intRemoteDocEntry, strRemoteDocNumber, SAPbobsCOM.BoRcptInvTypes.it_JournalEntry, objMainCompany.CompanyName)
                                            ElseIf ObjRemoteDoc.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Receipt Then 'Journal Entry
                                                intRefdocentry = getbaseDocument_IncomingPayment(intRemoteDocEntry, strRemoteDocNumber, SAPbobsCOM.BoRcptInvTypes.it_Receipt, objMainCompany.CompanyName)
                                            Else
                                                intRefdocentry = ObjRemoteDoc.Invoices.DocEntry
                                            End If

                                            'objMainDoc.Invoices.DocEntry  = ObjRemoteDoc.Invoices.DocEntry
                                            objMainDoc.Invoices.DocEntry = intRefdocentry
                                            objMainDoc.Invoices.DocLine = ObjRemoteDoc.Invoices.DocLine
                                            objMainDoc.Invoices.InstallmentId = ObjRemoteDoc.Invoices.InstallmentId
                                            objMainDoc.Invoices.InvoiceType = ObjRemoteDoc.Invoices.InvoiceType
                                            'objMainDoc.Invoices.AppliedFC = ObjRemoteDoc.Invoices.AppliedFC
                                        End If
                                    Next

                                    objMainDoc.IsPayToBank = ObjRemoteDoc.IsPayToBank
                                    objMainDoc.JournalRemarks = ObjRemoteDoc.JournalRemarks
                                    objMainDoc.LocalCurrency = ObjRemoteDoc.LocalCurrency
                                    objMainDoc.PaymentPriority = ObjRemoteDoc.PaymentPriority
                                    objMainDoc.PaymentType = ObjRemoteDoc.PaymentType
                                    objMainDoc.PayToBankAccountNo = ObjRemoteDoc.PayToBankAccountNo
                                    objMainDoc.PayToBankBranch = ObjRemoteDoc.PayToBankBranch
                                    objMainDoc.PayToBankCode = ObjRemoteDoc.PayToBankCode
                                    objMainDoc.PayToBankCountry = ObjRemoteDoc.PayToBankCountry
                                    objMainDoc.ProjectCode = ObjRemoteDoc.ProjectCode
                                    objMainDoc.Reference1 = ObjRemoteDoc.Reference1
                                    objMainDoc.Reference2 = ObjRemoteDoc.Reference2
                                    objMainDoc.Remarks = ObjRemoteDoc.Remarks
                                    '  objMainDoc.Series = ObjRemoteDoc.Series
                                    objMainDoc.TaxDate = ObjRemoteDoc.TaxDate
                                    objMainDoc.TaxGroup = ObjRemoteDoc.TaxGroup
                                    If ObjRemoteDoc.TransferSum > 0 Then
                                        objMainDoc.TransferAccount = ObjRemoteDoc.TransferAccount
                                        objMainDoc.TransferRealAmount = ObjRemoteDoc.TransferRealAmount
                                        objMainDoc.TransferReference = ObjRemoteDoc.TransferReference
                                        objMainDoc.TransferDate = ObjRemoteDoc.TransferDate
                                        objRemoteREC1.DoQuery("SELECT T0.[TrsfrSum], T0.[TrsfrSumFC], T0.[TrsfrRef], T0.[TrsfrDate] FROM ORCT T0 WHERE T0.[DocNum] =" & ObjRemoteDoc.DocNum)
                                        If ObjRemoteDoc.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tYES Then
                                            objMainDoc.TransferSum = ObjRemoteDoc.TransferSum
                                        Else
                                            If ObjRemoteDoc.DocRate > 0 Then
                                                objMainDoc.TransferSum = Math.Round(ObjRemoteDoc.TransferSum / ObjRemoteDoc.DocRate, 2)
                                            Else
                                                objMainDoc.TransferSum = ObjRemoteDoc.TransferSum
                                            End If
                                        End If
                                        'objMainDoc.TransferSum = ObjRemoteDoc.TransferSum
                                    End If
                                    objMainDoc.TransactionCode = ObjRemoteDoc.TransactionCode
                                    objMainDoc.WTAmount = ObjRemoteDoc.WTAmount
                                    objMainDoc.WtBaseSum = ObjRemoteDoc.WtBaseSum
                                    objMainDoc.WTCode = ObjRemoteDoc.WTCode
                                    If ObjRemoteDoc.DocType = SAPbobsCOM.BoRcptTypes.rAccount Then
                                        For intAcctLoop As Integer = 0 To ObjRemoteDoc.AccountPayments.Count - 1
                                            If ObjRemoteDoc.AccountPayments.SumPaid > 0 Then
                                                If intAcctLoop > 0 Then
                                                    objMainDoc.AccountPayments.Add()
                                                    objMainDoc.AccountPayments.SetCurrentLine(intAcctLoop)
                                                End If

                                                ObjRemoteDoc.AccountPayments.SetCurrentLine(intAcctLoop)
                                                ' objMainDoc.AccountPayments.GrossAmount = ObjRemoteDoc.AccountPayments.GrossAmount
                                                objMainDoc.AccountPayments.AccountCode = ObjRemoteDoc.AccountPayments.AccountCode
                                                objMainDoc.AccountPayments.AccountName = ObjRemoteDoc.AccountPayments.AccountName
                                                objMainDoc.AccountPayments.Decription = ObjRemoteDoc.AccountPayments.Decription
                                                objMainDoc.AccountPayments.GrossAmount = ObjRemoteDoc.AccountPayments.GrossAmount
                                                objMainDoc.AccountPayments.VatGroup = ObjRemoteDoc.AccountPayments.VatGroup
                                                objRemoteREC1.DoQuery("Select * from RCT4 where DocNum=" & ObjRemoteDoc.DocEntry & " and LineID=" & intAcctLoop)
                                                If objRemoteREC1.Fields.Item("AppliedFC").Value > 0 Then
                                                    objMainDoc.AccountPayments.SumPaid = Math.Round(ObjRemoteDoc.AccountPayments.SumPaid / ObjRemoteDoc.DocRate, 2)
                                                    'objMainDoc.AccountPayments.VatAmount = Math.Round(ObjRemoteDoc.AccountPayments.VatAmount / ObjRemoteDoc.DocRate, 2)
                                                Else
                                                    objMainDoc.AccountPayments.SumPaid = ObjRemoteDoc.AccountPayments.SumPaid
                                                    ' objMainDoc.AccountPayments.VatAmount = ObjRemoteDoc.AccountPayments.VatAmount
                                                End If
                                                objMainDoc.AccountPayments.ProfitCenter = ObjRemoteDoc.AccountPayments.ProfitCenter
                                                objMainDoc.AccountPayments.ProjectCode = ObjRemoteDoc.AccountPayments.ProjectCode

                                            End If
                                        Next
                                    End If


                                    If objMainDoc.Add() <> 0 Then
                                        sPath = strErrorLogPath
                                        WriteErrorlog("Failed to create Incomint payment  Document :" & ObjRemoteDoc.DocNum & " Error  :" & objRemoteCompany.GetLastErrorDescription, sPath)
                                    Else
                                        Dim strDocNumber As String
                                        objRemoteCompany.GetNewObjectCode(strDocNumber)
                                        'MsgBox(strDocNumber)
                                        sPath = strErrorLogPath
                                        If objMainDoc.GetByKey(Convert.ToInt64(strDocNumber)) Then
                                            strDocNumber = objMainDoc.DocNum
                                        End If
                                        WriteErrorlog("Incoming Payment document  Created Successfully. DocNum : " & strDocNumber, sPath)
                                        ObjRemoteDoc.UserFields.Fields.Item("U_Export").Value = "Y"
                                        ObjRemoteDoc.Update()
                                    End If
                                End If
                            End If
                        End If
                    End If
                Catch ex As Exception
                    WriteErrorlog("Failed to create Incoming  payment Error  :" & ex.Message, sPath)
                    WriteErrorlog("Failed to create Incoming payment Error  :" & ex.StackTrace, sPath)
                Finally
                    If Not IsNothing(ObjRemoteDoc) Then
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(ObjRemoteDoc)
                        GC.Collect()
                    End If
                    If Not IsNothing(objMainDoc) Then
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objMainDoc)
                        GC.Collect()
                    End If
                End Try
                objremoteRec.MoveNext()
            Next
        Catch ex As Exception
            WriteErrorlog(ex.Message, sPath)
            WriteErrorlog("Failed to create Incomint payment Error  :" & ex.StackTrace, sPath)
        Finally
            If Not IsNothing(ObjRemoteDoc) Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ObjRemoteDoc)
                GC.Collect()
            End If
            If Not IsNothing(objMainDoc) Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objMainDoc)
                GC.Collect()
            End If
        End Try

    End Sub


    Private Sub ExportIncomingPayment_Cancelled(ByVal aCompany As SAPbobsCOM.Company, ByVal aDocEntry As Integer)
        Dim objMainDoc, ObjRemoteDoc As SAPbobsCOM.Payments
        Dim objremoteRec, objRemoteREC1 As SAPbobsCOM.Recordset
        Dim objRec1 As SAPbobsCOM.Recordset
        Dim strRemoteDocNumber As Integer
        Try
            objremoteRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objRemoteREC1 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objremoteRec.DoQuery("Select DocEntry from ORCT where DocEntry=" & aDocEntry)
            objRec1 = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            For intRemoteloop As Integer = 0 To objremoteRec.RecordCount - 1
                Try
                    ObjRemoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
                    objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
                    If ObjRemoteDoc.GetByKey(Convert.ToInt32(objremoteRec.Fields.Item(0).Value)) Then
                        If ValidateDocDate(ObjRemoteDoc.DocEntry, "24", ObjRemoteDoc.DocDate) = True Then
                            If ObjRemoteDoc.Cancelled = SAPbobsCOM.BoYesNoEnum.tYES Then
                                If CheckCanceledInvoice(ObjRemoteDoc.DocEntry, ObjRemoteDoc.DocNum, objMainCompany.CompanyName, "ORCT") <> 0 Then
                                    If objMainDoc.GetByKey(CheckCanceledInvoice(ObjRemoteDoc.DocEntry, ObjRemoteDoc.DocNum, objMainCompany.CompanyName, "ORCT")) Then
                                        If objMainDoc.Cancelled = SAPbobsCOM.BoYesNoEnum.tNO Then
                                            Dim strDocNum As String = objMainDoc.DocNum
                                            If objMainDoc.Cancel <> 0 Then
                                                '  MsgBox(objRemoteCompany.GetLastErrorDescription)
                                            Else
                                                WriteErrorlog("Incoming Payment document  Cancelled Successfully. DocNum : " & strDocNum, sPath)
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                               
                            End If
                        End If
                    End If
                Catch ex As Exception
                    WriteErrorlog("Failed to create Incoming payment Error  :" & ex.Message, sPath)
                    WriteErrorlog("Failed to create Incoming payment Error  :" & ex.StackTrace, sPath)
                Finally
                    If Not IsNothing(ObjRemoteDoc) Then
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(ObjRemoteDoc)
                        GC.Collect()
                    End If
                    If Not IsNothing(objMainDoc) Then
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objMainDoc)
                        GC.Collect()
                    End If

                End Try
                objremoteRec.MoveNext()
            Next
        Catch ex As Exception
            WriteErrorlog(ex.Message, sPath)
            WriteErrorlog("Failed to create Incoming payment Error  :" & ex.StackTrace, sPath)
        Finally
            'If Not IsNothing(objremoteRec) Then
            '    System.Runtime.InteropServices.Marshal.ReleaseComObject(objremoteRec)
            'End If
            'If Not IsNothing(objRemoteREC1) Then
            '    System.Runtime.InteropServices.Marshal.ReleaseComObject(objRemoteREC1)
            'End If

        End Try

    End Sub

#End Region



#Region "Impot BP Master"


    Private Sub ImportBPMaster()
        Dim FileName, strItem, strGroupName As String
        Dim Ecount As Long
        Dim ii As Long
        Dim objMainItem, objRemoteItem As SAPbobsCOM.BusinessPartners
        Dim objMainRecset, objTemp As SAPbobsCOM.Recordset
        objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRecset.DoQuery("Select CardCode from OCRD where isnull(U_Action,'N')<>'N'")
        For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
            FileName = System.Windows.Forms.Application.StartupPath & "\BP.xml"
            strItem = objMainRecset.Fields.Item(0).Value
            objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            If objMainItem.GetByKey(strItem) Then
                objTemp = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objTemp.DoQuery("Select isnull(GroupName,'') from OCRG where GroupCode=" & objMainItem.GroupCode)
                strGroupName = objTemp.Fields.Item(0).Value
                objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                If objRemoteItem.GetByKey(strItem) Then
                    objMainCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                    objMainItem.SaveXML(FileName)

                    objTemp = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    objTemp.DoQuery("Select isnull(GroupCode," & objMainItem.GroupCode & ") from OCRG where GroupName='" & strGroupName & "'")
                    strGroupName = objTemp.Fields.Item(0).Value
                    If ReadBPXML(FileName, strGroupName) = False Then
                    End If

                    objRemoteCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                    objRemoteItem = objRemoteCompany.GetBusinessObjectFromXML(FileName, 0)
                    If objRemoteItem.Update() <> 0 Then
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & ": Failed to Update BP : " & objRemoteItem.CardCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & ":BP Code Updated : " & objRemoteItem.CardCode, sPath)
                        objMainItem.UserFields.Fields.Item("U_Action").Value = "N"
                        objMainItem.Update()
                    End If
                Else
                    objMainCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                    objMainItem.SaveXML(FileName)

                    objTemp = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    objTemp.DoQuery("Select isnull(GroupCode," & objMainItem.GroupCode & ") from OCRG where GroupName='" & strGroupName & "'")
                    strGroupName = objTemp.Fields.Item(0).Value
                    If ReadBPXML(FileName, strGroupName) = False Then
                    End If

                    objRemoteCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                    objRemoteItem = objRemoteCompany.GetBusinessObjectFromXML(FileName, 0)
                    If objRemoteItem.Add() <> 0 Then
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & ": Failed to create BP : " & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & ": New BP Created : " & objRemoteItem.CardCode, sPath)
                        objMainItem.UserFields.Fields.Item("U_Action").Value = "N"
                        objMainItem.Update()
                    End If
                End If
                ' objRemoteCompany1 = objRemoteCompany
                '    objRemoteItem = objRemoteCompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                '    If objRemoteItem.GetByKey(strItem) Then

                '        objTemp = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                '        objTemp.DoQuery("Select isnull(GroupName,'') from OCRG where GroupCode=" & objMainItem.GroupCode)
                '        strGroupName = objTemp.Fields.Item(0).Value

                '        objMainCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                '        objMainItem.SaveXML(FileName)

                '        objTemp = objRemoteCompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                '        objTemp.DoQuery("Select isnull(GroupCode," & objMainItem.GroupCode & ") from OCRG where GroupName='" & strGroupName & "'")
                '        strGroupName = objTemp.Fields.Item(0).Value
                '        If ReadBPXML(FileName, strGroupName) = False Then
                '        End If

                '        objRemoteCompany1.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                '        objRemoteItem = objRemoteCompany1.GetBusinessObjectFromXML(FileName, 0)
                '        If objRemoteItem.Update() <> 0 Then
                '            WriteErrorlog("Database Name : " & objRemoteCompany1.CompanyName & ": Failed to Update BP : " & objRemoteItem.CardCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                '        Else
                '            objMainItem.UserFields.Fields.Item("U_Action").Value = "N"
                '            objMainItem.Update()
                '            WriteErrorlog("Database Name : " & objRemoteCompany1.CompanyName & " : BP Code Updated : " & objRemoteItem.CardCode, sPath)
                '        End If
                '    Else
                '        objMainCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                '        objMainItem.SaveXML(FileName)

                '        objTemp = objRemoteCompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                '        objTemp.DoQuery("Select isnull(GroupCode," & objMainItem.GroupCode & ") from OCRG where GroupName='" & strGroupName & "'")
                '        strGroupName = objTemp.Fields.Item(0).Value
                '        If ReadBPXML(FileName, strGroupName) = False Then
                '        End If

                '        objRemoteCompany1.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                '        objRemoteItem = objRemoteCompany1.GetBusinessObjectFromXML(FileName, 0)
                '        If objRemoteItem.Add() <> 0 Then
                '            WriteErrorlog("Database Name : " & objRemoteCompany1.CompanyName & " : Failed to create BP : " & objRemoteCompany.GetLastErrorDescription, sPath)
                '        Else
                '            objMainItem.UserFields.Fields.Item("U_Action").Value = "N"
                '            objMainItem.Update()
                '            WriteErrorlog("Database Name : " & objRemoteCompany1.CompanyName & " : New BP Created : " & objRemoteItem.CardCode, sPath)
                '        End If
                '    End If
            End If
            objMainRecset.MoveNext()
        Next
    End Sub
    Private Sub UpdateBPMaster()
        Dim FileName, strItem As String
        Dim Ecount As Long
        Dim ii As Long
        Dim objMainItem, objRemoteItem As SAPbobsCOM.BusinessPartners
        Dim objMainRecset As SAPbobsCOM.Recordset
        objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRecset.DoQuery("Select CardCode from OCRD where U_Action='U'")
        For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
            FileName = System.Windows.Forms.Application.StartupPath & "\BP.xml"
            strItem = objMainRecset.Fields.Item(0).Value
            objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            If objMainItem.GetByKey(strItem) Then
                objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                If objRemoteItem.GetByKey(strItem) Then
                    objMainCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                    objMainItem.SaveXML(FileName)
                    objRemoteCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                    objRemoteItem = objRemoteCompany.GetBusinessObjectFromXML(FileName, 0)
                    If objRemoteItem.Update <> 0 Then
                        WriteErrorlog("Failed to update BP : " & objRemoteItem.CardCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        objMainItem.UserFields.Fields.Item("U_Action").Value = "N"
                        objMainItem.Update()
                        WriteErrorlog("BP Updated : " & objRemoteItem.CardCode, sPath)
                    End If
                End If

                objRemoteItem = objRemoteCompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
                If objRemoteItem.GetByKey(strItem) Then
                    objMainCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                    objMainItem.SaveXML(FileName)
                    objRemoteCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                    objRemoteItem = objRemoteCompany.GetBusinessObjectFromXML(FileName, 0)
                    If objRemoteItem.Update <> 0 Then
                        WriteErrorlog("Failed to update BP : " & objRemoteItem.CardCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        objMainItem.UserFields.Fields.Item("U_Action").Value = "N"
                        objMainItem.Update()
                        WriteErrorlog("BP Updated : " & objRemoteItem.CardCode, sPath)
                    End If
                End If
            End If
            objMainRecset.MoveNext()
        Next
    End Sub
    '    Private Sub ImportBPMaster()
    '        Dim FileName, strItem As String
    '        Dim Ecount As Long
    '        Dim ii As Long
    '        Dim objMainItem, objRemoteItem As SAPbobsCOM.BusinessPartners
    '        Dim objMainRecset As SAPbobsCOM.Recordset
    '        Dim objTemp As SAPbobsCOM.Recordset
    '        dim strGroupName as String 
    '        objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        objMainRecset.DoQuery("Select CardCode from OCRD where U_Action='A'")
    '        For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
    '            FileName = System.Windows.Forms.Application.StartupPath & "\BP.xml"
    '            strItem = objMainRecset.Fields.Item(0).Value
    '            objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
    '            If objMainItem.GetByKey(strItem) Then
    '                objTemp = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                objTemp.DoQuery("Select isnull(GroupName,'') from OCRG where GroupCode=" & objMainItem.GroupCode)
    '                strGroupName = objTemp.Fields.Item(0).Value
    '                objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

    '                objMainCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
    '                objMainItem.SaveXML(FileName)
    '                objRemoteCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
    '                objRemoteItem = objRemoteCompany.GetBusinessObjectFromXML(FileName, 0)
    '                If objRemoteItem.Add() <> 0 Then
    '                    WriteErrorlog("Failed to create BP : " & objRemoteCompany.GetLastErrorDescription, sPath)
    '                Else
    '                    objMainItem.UserFields.Fields.Item("U_Action").Value = "N"
    '                    objMainItem.Update()
    '                    WriteErrorlog("New BP Created : " & objRemoteItem.CardCode, sPath)
    '                End If
    '            End If
    '            objMainRecset.MoveNext()
    '        Next
    '    End Sub
    '    Private Sub UpdateBPMaster()
    '        Dim FileName, strItem As String
    '        Dim Ecount As Long
    '        Dim ii As Long
    '        Dim objMainItem, objRemoteItem As SAPbobsCOM.BusinessPartners
    '        Dim objMainRecset As SAPbobsCOM.Recordset
    '        objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        objMainRecset.DoQuery("Select CardCode from OCRD where U_Action='U'")
    '        For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
    '            FileName = System.Windows.Forms.Application.StartupPath & "\BP.xml"
    '            strItem = objMainRecset.Fields.Item(0).Value
    '            objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
    '            If objMainItem.GetByKey(strItem) Then
    '                objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
    '                If objRemoteItem.GetByKey(strItem) Then
    '                    objMainCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
    '                    objMainItem.SaveXML(FileName)
    '                    objRemoteCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
    '                    objRemoteItem = objRemoteCompany.GetBusinessObjectFromXML(FileName, 0)
    '                    If objRemoteItem.Update <> 0 Then
    '                        WriteErrorlog("Failed to update BP : " & objRemoteItem.CardCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
    '                    Else
    '                        objMainItem.UserFields.Fields.Item("U_Action").Value = "N"
    '                        objMainItem.Update()
    '                        WriteErrorlog("BP Updated : " & objRemoteItem.CardCode, sPath)
    '                    End If
    '                End If
    '            End If
    '            objMainRecset.MoveNext()
    '        Next
    '    End Sub


    Private Function ReadBPXML(ByVal XmlFilePath As String, ByVal aGroupCode As String) As Boolean
        Dim strOrdNum, strCardCode, strNumatcard, strDocDate, strDueDate, strShipid, strBillid, strCarcod, strSrvlvl, strWMSRefNo As String
        Dim strItemCode, strBaseRef, strInvoiceNumLines, strbaseNum, strBaseLine, strquantity, strheaderqty, strWhs, strShippingcharge As String
        'Dim objStaticitem As SAPbouiCOM.StaticText
        Dim dblQty, dblUnitprice As Double
        Dim oRS As SAPbobsCOM.Recordset
        Dim objXmlDocument As Xml.XmlDocument
        objXmlDocument = New Xml.XmlDocument
        Dim blnErrorFlag As Boolean = False
        Dim oRec As SAPbobsCOM.Recordset
        Dim objOrders, objOrder, objMove, objStop, objShipMent, objShipOrder, objSHipHeader, objSHipLines, objShipload, objSerialNums, objProduct, objitem, objBox, objSN As Xml.XmlNode
        'sPath = System.Windows.Forms.Application.StartupPath & "\ImportErrorLog.txt"
        'If File.Exists(sPath) Then
        '    File.Delete(sPath)
        'End If
        Dim strSerialRef As String = ""
        Try
            objXmlDocument.Load(XmlFilePath)
            objOrders = objXmlDocument.ChildNodes.Item(1)
            If objOrders.Name = "BOM" Then
                For Each objOrder In objOrders.ChildNodes
                    If objOrder.Name = "BO" Then
                        For Each objMove In objOrder.ChildNodes
                            If objMove.Name = "BusinessPartners" Then
                                For Each objStop In objMove.ChildNodes
                                    If objStop.Name = "row" Then
                                        For Each objSHipLines In objStop
                                            If objSHipLines.Name = "GroupCode" Then
                                                objSHipLines.FirstChild.Value = aGroupCode
                                            End If
                                        Next
                                    End If
                                Next

                            End If
                        Next
                    End If
                Next
            End If
            objXmlDocument.Save(XmlFilePath)
        Catch ex As Exception
            WriteErrorlog(ex.Message, sPath)
            Return False
        End Try
        Return True
    End Function

    Private Function ReadItemXML(ByVal XmlFilePath As String, ByVal aGroupCode As String) As Boolean
        Dim strOrdNum, strCardCode, strNumatcard, strDocDate, strDueDate, strShipid, strBillid, strCarcod, strSrvlvl, strWMSRefNo As String
        Dim strItemCode, strBaseRef, strInvoiceNumLines, strbaseNum, strBaseLine, strquantity, strheaderqty, strWhs, strShippingcharge As String
        'Dim objStaticitem As SAPbouiCOM.StaticText
        Dim dblQty, dblUnitprice As Double
        Dim oRS As SAPbobsCOM.Recordset
        Dim objXmlDocument As Xml.XmlDocument
        objXmlDocument = New Xml.XmlDocument
        Dim blnErrorFlag As Boolean = False
        Dim oRec As SAPbobsCOM.Recordset
        Dim objOrders, objOrder, objMove, objStop, objShipMent, objShipOrder, objSHipHeader, objSHipLines, objShipload, objSerialNums, objProduct, objitem, objBox, objSN As Xml.XmlNode
        'sPath = System.Windows.Forms.Application.StartupPath & "\ImportErrorLog.txt"
        'If File.Exists(sPath) Then
        '    File.Delete(sPath)
        'End If
        Dim strSerialRef As String = ""
        Try
            objXmlDocument.Load(XmlFilePath)
            objOrders = objXmlDocument.ChildNodes.Item(1)
            If objOrders.Name = "BOM" Then
                For Each objOrder In objOrders.ChildNodes
                    If objOrder.Name = "BO" Then
                        For Each objMove In objOrder.ChildNodes
                            If objMove.Name = "Items" Then
                                For Each objStop In objMove.ChildNodes
                                    If objStop.Name = "row" Then
                                        For Each objSHipLines In objStop
                                            If objSHipLines.Name = "ItemsGroupCode" Then
                                                objSHipLines.FirstChild.Value = aGroupCode
                                            End If
                                        Next
                                    End If
                                Next

                            End If
                        Next
                    End If
                Next
            End If
            objXmlDocument.Save(XmlFilePath)
        Catch ex As Exception
            WriteErrorlog(ex.Message, sPath)
            Return False
        End Try
        Return True
    End Function
#End Region

#Region "Get MaxCode"
    Public Function getMaxCode(ByVal sTable As String, ByVal sColumn As String, ByVal oRecSet As SAPbobsCOM.Recordset) As String
        Dim oRS As SAPbobsCOM.Recordset
        Dim MaxCode As Integer
        Dim sCode As String
        Dim strSQL As String
        Try
            strSQL = "SELECT MAX(CAST(" & sColumn & " AS Numeric)) FROM [" & sTable & "]"
            oRecSet.DoQuery(strSQL)

            If Convert.ToString(oRecSet.Fields.Item(0).Value).Length > 0 Then
                MaxCode = oRecSet.Fields.Item(0).Value + 1
            Else
                MaxCode = 1
            End If

            sCode = Format(MaxCode, "00000000")
            Return sCode
        Catch ex As Exception
            Throw ex
        Finally
            oRS = Nothing
        End Try
    End Function
#End Region

#Region "Export Stock transfer Request Details"
    Private Sub StocktransferRequest()
        Dim FileName, strItem, strSQLquery, strCode As String
        Dim Ecount As Long
        Dim ii, intDocEntry As Long
        strSQLquery = "SELECT T0.[DocEntry], T0.[U_DocDate], T0.[U_WhsCode], T1.[U_ItemCode], T1.[U_ItemName], T1.[U_Qty],T0.U_Status FROM [dbo].[@DABT_STRHEADER]  T0 "
        strSQLquery = strSQLquery & " inner join [dbo].[@DABT_STRLINES]  T1  on  (T0.[DocEntry] = T1.[DocEntry]  )  where t0.U_Status='O' and  isnull(T1.U_ItemCode,'')<>''"
        Dim objMainItem, objRemoteItem As SAPbobsCOM.UserTable
        Dim objMainRecset, objremoteRecSet As SAPbobsCOM.Recordset
        objremoteRecSet = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objremoteRecSet.DoQuery(strSQLquery)
        objMainRecset = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        For intItemLoop As Integer = 0 To objremoteRecSet.RecordCount - 1
            intDocEntry = objremoteRecSet.Fields.Item(0).Value
            objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objMainItem = objMainCompany.UserTables.Item("DABT_StImport")
            strCode = getMaxCode("@DABT_StImport", "Code", objMainRecset)
            objMainItem.Code = strCode
            objMainItem.Name = strCode
            objMainItem.UserFields.Fields.Item("U_TransDate").Value = objremoteRecSet.Fields.Item(1).Value
            objMainItem.UserFields.Fields.Item("U_TransWhs").Value = objremoteRecSet.Fields.Item(2).Value
            objMainItem.UserFields.Fields.Item("U_ItemCode").Value = objremoteRecSet.Fields.Item(3).Value
            objMainItem.UserFields.Fields.Item("U_ItemName").Value = objremoteRecSet.Fields.Item(4).Value
            objMainItem.UserFields.Fields.Item("U_ReqQty").Value = objremoteRecSet.Fields.Item(5).Value
            If objMainItem.Add <> 0 Then
                WriteErrorlog("Failed to import data " & objRemoteCompany.GetLastErrorDescription, sPath)
            Else
                objMainRecset = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objMainRecset.DoQuery("Update [@DABT_STRHEADER] set U_Status='I' where docentry=" & intDocEntry)
            End If
            objremoteRecSet.MoveNext()
        Next
    End Sub
#End Region

#Region "Impot Item Master"
    Private Sub ImportItemMaster()
        Dim FileName, strItem, strDflWhs, strGroupName As String
        Dim Ecount As Long
        Dim ii As Long
        Dim objMainItem, objRemoteItem As SAPbobsCOM.Items
        Dim objMainRecset, objTemp As SAPbobsCOM.Recordset
        objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRecset.DoQuery("Select ItemCode from OITM") ' where isnull(U_Action,'N')<>'N'")
        For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
            FileName = System.Windows.Forms.Application.StartupPath & "\Item.xml"
            strItem = objMainRecset.Fields.Item(0).Value
            Try
                objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                If objMainItem.GetByKey(strItem) Then
                    objTemp = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    objTemp.DoQuery("Select isnull(ItmsGrpNam,'') from OITB where ItmsGrpCod=" & objMainItem.ItemsGroupCode)
                    strGroupName = objTemp.Fields.Item(0).Value
                    objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                    If objRemoteItem.GetByKey(strItem) Then
                        strDflWhs = objRemoteItem.DefaultWarehouse
                        objMainCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                        objMainItem.SaveXML(FileName)
                        objTemp = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        objTemp.DoQuery("Select isnull(ItmsGrpCod," & objMainItem.ItemsGroupCode & ") from OITB where ItmsGrpNam='" & strGroupName & "'")
                        strGroupName = objTemp.Fields.Item(0).Value
                        If ReadItemXML(FileName, strGroupName) = False Then
                        End If
                        'objRemoteCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                        'objRemoteItem = objRemoteCompany.GetBusinessObjectFromXML(FileName, 0)
                        Dim x As SAPbobsCOM.Items
                        Dim y As SAPbobsCOM.Items
                        x = objRemoteItem
                        y = objMainItem
                        y.GetByKey(strItem)
                        If x.GetByKey(strItem) Then
                            x.ApTaxCode = y.ApTaxCode
                            x.ArTaxCode = y.ArTaxCode
                            x.AssetItem = y.AssetItem
                            x.AutoCreateSerialNumbersOnRelease = y.AutoCreateSerialNumbersOnRelease
                            x.AvgStdPrice = y.AvgStdPrice
                            x.BarCode = y.BarCode
                            x.BaseUnitName = y.BaseUnitName
                            x.BeverageCommercialBrandCode = y.BeverageCommercialBrandCode
                            x.BeverageGroupCode = y.BeverageGroupCode
                            x.BeverageTableCode = y.BeverageTableCode
                            x.CommissionGroup = y.CommissionGroup
                            x.CommissionPercent = y.CommissionGroup
                            x.CommissionSum = y.CommissionSum
                            x.CostAccountingMethod = y.CostAccountingMethod
                            x.CustomsGroupCode = y.CustomsGroupCode
                            x.DataExportCode = y.DataExportCode
                            x.DefaultWarehouse = strDflWhs
                            x.DesiredInventory = y.DesiredInventory
                            x.DNFEntry = y.DNFEntry
                            x.ECExpensesAccount = y.ECExpensesAccount
                            x.ECRevenuesAccount = y.ECRevenuesAccount
                            x.ExemptIncomeAccount = y.ExemptIncomeAccount
                            x.ExpanseAccount = y.ExpanseAccount
                            x.ForceSelectionOfSerialNumber = y.ForceSelectionOfSerialNumber
                            x.ForeignExpensesAccount = y.ForeignExpensesAccount
                            x.ForeignName = y.ForeignName
                            x.ForeignRevenuesAccount = y.ForeignRevenuesAccount
                            x.Frozen = y.Frozen
                            x.FrozenFrom = y.FrozenFrom
                            x.FrozenRemarks = y.FrozenRemarks
                            x.FrozenTo = y.FrozenTo
                            x.FuelID = y.FuelID
                            x.GLMethod = y.GLMethod
                            x.GTSItemSpec = y.GLMethod
                            x.GTSItemTaxCategory = y.GTSItemSpec
                            x.IncomeAccount = y.IncomeAccount
                            x.IncomingServiceCode = y.IncomingServiceCode
                            x.IndirectTax = y.IndirectTax
                            x.InventoryUOM = y.InventoryUOM
                            x.IsPhantom = y.IsPhantom
                            x.IssueMethod = y.IssueMethod
                            x.ItemClass = y.ItemClass
                            x.ItemCountryOrg = y.ItemCountryOrg
                            x.ItemsGroupCode = strGroupName
                            x.ItemType = y.ItemType
                            x.LeadTime = y.LeadTime
                            x.ItemName = y.ItemName
                            x.Mainsupplier = y.Mainsupplier
                            x.ManageBatchNumbers = y.ManageBatchNumbers
                            x.ManageSerialNumbers = y.ManageSerialNumbers
                            x.ManageSerialNumbersOnReleaseOnly = y.ManageSerialNumbersOnReleaseOnly
                            x.ManageStockByWarehouse = y.ManageStockByWarehouse
                            x.Manufacturer = y.Manufacturer
                            x.MaterialGroup = y.MaterialGroup
                            x.MaterialType = y.MaterialType
                            x.MaxInventory = y.MaxInventory
                            x.MinOrderQuantity = y.MinInventory
                            x.MinInventory = y.MinInventory
                            'x.MovingAveragePrice = y.MovingAveragePrice
                            x.NCMCode = y.NCMCode
                            Dim xPricelist, yPriceList As SAPbobsCOM.Items_Prices
                            xPricelist = x.PriceList
                            yPriceList = y.PriceList
                            For i As Integer = 0 To yPriceList.Count - 1
                                xPricelist.SetCurrentLine(i)
                                yPriceList.SetCurrentLine(i)
                                'print price list name
                                xPricelist.Price = yPriceList.Price
                            Next


                            If y.OrderIntervals <> "0" Then
                                x.OrderIntervals = y.OrderIntervals
                            End If

                            x.OrderMultiple = y.OrderMultiple
                            x.OutgoingServiceCode = y.OutgoingServiceCode
                            x.Picture = y.Picture
                            x.PlanningSystem = y.PlanningSystem
                            'x.PreferredVendors 
                            x.ProcurementMethod = y.ProcurementMethod
                            x.ProductSource = y.ProductSource
                            'For intPro As Integer = 0 To 64
                            '    x.Properties(intPro) = y.Properties(intPro)
                            'Next
                            x.PurchaseFactor1 = y.PurchaseFactor1
                            x.PurchaseFactor2 = y.PurchaseFactor2
                            x.PurchaseFactor3 = y.PurchaseFactor3
                            x.PurchaseFactor4 = y.PurchaseFactor4
                            '   x.PurchaseHeightUnit = y.PurchaseHeightUnit
                            ' x.PurchaseHeightUnit1 = y.PurchaseHeightUnit1
                            x.PurchaseItem = y.PurchaseItem
                            If y.PurchaseItemsPerUnit > 0 Then
                                x.PurchaseItemsPerUnit = y.PurchaseItemsPerUnit
                            End If

                            If y.PurchaseLengthUnit > 0 Then
                                x.PurchaseLengthUnit = y.PurchaseLengthUnit
                            End If

                            ' x.PurchaseLengthUnit1 = y.PurchaseLengthUnit1
                            'x.PurchasePackagingUnit = y.PurchasePackagingUnit
                            'x.PurchaseQtyPerPackUnit = y.PurchaseQtyPerPackUnit
                            'x.PurchaseUnit = y.PurchaseUnit
                            'x.PurchaseUnitHeight = y.PurchaseUnitHeight
                            ''   x.PurchaseUnitHeight1 = y.PurchaseUnitHeight1
                            'x.PurchaseUnitLength = y.PurchaseUnitLength
                            ''  x.PurchaseUnitLength1 = y.PurchaseUnitLength1
                            'x.PurchaseUnitVolume = y.PurchaseUnitVolume
                            'x.PurchaseUnitWeight = y.PurchaseUnitWeight
                            ''  x.PurchaseUnitWeight1 = y.PurchaseUnitWeight1
                            'x.PurchaseUnitWidth = y.PurchaseUnitWidth
                            ''   x.PurchaseUnitWidth1 = y.PurchaseUnitWidth1
                            'x.PurchaseVATGroup = y.PurchaseVATGroup
                            'x.PurchaseVolumeUnit = y.PurchaseVolumeUnit
                            'x.PurchaseWeightUnit = x.PurchaseWeightUnit
                            '   x.PurchaseWeightUnit1 = y.PurchaseWeightUnit1
                            x.SalesFactor1 = y.SalesFactor1
                            x.SalesFactor2 = y.SalesFactor2
                            x.SalesFactor3 = y.SalesFactor3
                            x.SalesFactor4 = y.SalesFactor4
                            '   x.SalesHeightUnit = y.SalesHeightUnit
                            '  x.SalesHeightUnit1 = y.SalesHeightUnit1

                            x.SalesItem = y.SalesItem
                            If y.SalesItemsPerUnit > 0 Then
                                x.SalesItemsPerUnit = y.SalesItemsPerUnit
                            End If

                            '  x.SalesLengthUnit = y.SalesLengthUnit
                            ' x.SalesLengthUnit1 = y.SalesLengthUnit1
                            'x.SalesPackagingUnit = y.SalesPackagingUnit
                            'x.SalesQtyPerPackUnit = y.SalesQtyPerPackUnit
                            x.SalesUnit = y.SalesUnit
                            'x.SalesUnitHeight = y.SalesUnitHeight
                            'x.SalesUnitHeight1 = y.SalesUnitHeight1
                            'x.SalesUnitLength = y.SalesUnitLength
                            'x.SalesUnitVolume = y.SalesUnitVolume
                            'x.SalesUnitWeight = y.SalesUnitWeight
                            'x.SalesUnitWeight1 = y.SalesUnitWeight1
                            'x.SalesUnitWidth = y.SalesUnitWidth
                            'x.SalesUnitWidth1 = y.SalesUnitWidth1
                            'x.SalesVATGroup = y.SalesVATGroup
                            'x.SalesVolumeUnit = y.SalesVolumeUnit
                            'x.SalesWeightUnit = y.SalesWeightUnit
                            'x.SalesWeightUnit1 = y.SalesWeightUnit1
                            'x.SalesWidthUnit = y.SalesWidthUnit
                            'x.SalesWidthUnit1 = y.SalesWidthUnit1

                            x.SerialNum = y.SerialNum
                            x.ServiceGroup = y.ServiceGroup
                            If y.ShipType > 0 Then
                                x.ShipType = y.ShipType
                            End If
                            x.SRIAndBatchManageMethod = y.SRIAndBatchManageMethod
                            x.SupplierCatalogNo = y.SupplierCatalogNo
                            x.SWW = y.SWW
                            x.TaxType = y.TaxType
                            x.ToleranceDays = y.ToleranceDays
                            x.User_Text = y.User_Text
                            x.Valid = y.Valid
                            x.ValidFrom = y.ValidFrom
                            x.ValidRemarks = y.ValidRemarks
                            x.ValidTo = y.ValidTo
                            x.VatLiable = y.VatLiable
                            x.WarrantyTemplate = y.WarrantyTemplate
                            For intUDF As Integer = 0 To y.UserFields.Fields.Count - 1
                                If y.UserFields.Fields.Item(intUDF).Value = "" Then
                                Else
                                    x.UserFields.Fields.Item(intUDF).Value = y.UserFields.Fields.Item(intUDF).Value
                                End If
                            Next
                            If x.Update <> 0 Then
                                WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & ": Failed to Update item : " & x.ItemCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                            Else
                                '  WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Item Code updated : " & x.ItemCode, sPath)
                                '   y.UserFields.Fields.Item("U_Action").Value = "N"
                                '  y.Update()
                                If x.GetByKey(strItem) Then
                                    x.DefaultWarehouse = strDflWhs
                                    x.Update()
                                End If
                            End If
                        End If
                        'If objRemoteItem.Update <> 0 Then
                        '    WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & ": Failed to Update item : " & objRemoteItem.ItemCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                        'Else
                        '    WriteErrorlog("Database Name : " & objRemoteCompany.CompanyName & " : Item Code updated : " & objRemoteItem.ItemCode, sPath)
                        '    objMainItem.UserFields.Fields.Item("U_Action").Value = "N"
                        '    objMainItem.Update()
                        '    If objRemoteItem.GetByKey(strItem) Then
                        '        objRemoteItem.DefaultWarehouse = strDflWhs
                        '        objRemoteItem.Update()
                        '    End If
                        'End If
                    Else
                        objMainCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                        objMainItem.SaveXML(FileName)
                        objTemp = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        objTemp.DoQuery("Select isnull(ItmsGrpCod," & objMainItem.ItemsGroupCode & ") from OITB where ItmsGrpNam='" & strGroupName & "'")
                        strGroupName = objTemp.Fields.Item(0).Value
                        If ReadItemXML(FileName, strGroupName) = False Then
                        End If
                        objRemoteCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                        objRemoteItem = objRemoteCompany.GetBusinessObjectFromXML(FileName, 0)
                        If objRemoteItem.Add() <> 0 Then
                            WriteErrorlog("DataBase : " & objRemoteCompany.CompanyName & " : Failed to create item : " & objRemoteCompany.GetLastErrorDescription, sPath)
                        Else
                            WriteErrorlog("DataBase : " & objRemoteCompany.CompanyName & " :  New Item Created : " & objRemoteItem.ItemCode, sPath)
                            objMainItem.UserFields.Fields.Item("U_Action").Value = "N"
                            objMainItem.Update()
                        End If
                    End If
                End If
            Catch ex As Exception

            End Try
            objMainRecset.MoveNext()
        Next
    End Sub

    Private Sub UpdateItemMaster()
        Dim FileName, strItem As String
        Dim Ecount As Long
        Dim ii As Long
        Dim objMainItem, objRemoteItem As SAPbobsCOM.Items
        Dim objMainRecset As SAPbobsCOM.Recordset
        objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRecset.DoQuery("Select ItemCode from OITM where U_Action='U'")
        For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
            FileName = System.Windows.Forms.Application.StartupPath & "\Item.xml"
            strItem = objMainRecset.Fields.Item(0).Value
            objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
            If objMainItem.GetByKey(strItem) Then
                objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                If objRemoteItem.GetByKey(strItem) Then
                    objMainCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                    objMainItem.SaveXML(FileName)
                    objRemoteCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
                    objRemoteItem = objRemoteCompany.GetBusinessObjectFromXML(FileName, 0)
                    If objRemoteItem.Update <> 0 Then
                        WriteErrorlog("Failed to Update item : " & objRemoteItem.ItemCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
                    Else
                        WriteErrorlog("Update Completed : " & objRemoteItem.Treecode, sPath)
                    End If
                End If
            End If
            objMainRecset.MoveNext()
        Next
    End Sub
    'Private Sub ImportItemMaster()
    '    Dim FileName, strItem As String
    '    Dim Ecount As Long
    '    Dim ii As Long
    '    Dim objMainItem, objRemoteItem As SAPbobsCOM.Items
    '    Dim objMainRecset As SAPbobsCOM.Recordset
    '    '        Dim sboItem As SAPbobsCOM.Items = oCompany.GetBusinessObject(BoObjectTypes.oItems)
    '    '        If sboItem.GetByKey(itemcode) = False Then
    '    '            Exit Sub
    '    '        End If
    '    '        oCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
    '    '        sboItem.ItemCode = itemcode + "_d"
    '    '        sboItem.SaveToFile(String.Format("c:\temp\{0}.xml", itemcode))

    '    '        Dim sboNewItem As SAPbobsCOM.Items = oCompany.GetBusinessObjectFromXML(String.Format("c:\temp\{0}.xml", itemcode), 0)

    '    '        If sboNewItem.Add()  0 Then
    '    '            sbo_application.MessageBox(oCompany.GetLastErrorCode.ToString + Space(1) + oCompany.GetLastErrorDescription)
    '    '        Else
    '    '            sbo_application.MessageBox("Item added")
    '    '        End If
    '    'Your code should be
    '    '        sboItem = oSBO.SboCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
    '    '        _oSBO.SboCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
    '    '        sboItem.ItemCode = sItemCode
    '    '        sboItem.SaveXML(sFileName)
    '    '        sboNewItem = _oSBO.SboCompany.GetBusinessObjectFromXML(sFileName, 0)
    '    '        iRet = sboNewItem.Add()
    '    'If iRet 0 Then
    '    '            '_oSBO.SboCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
    '    '            _oSBO.SboCompany.GetLastError(iRet, sErrMsg)
    '    '            _oSBO.SboCompany.MessageBox(sErrMsg & vbCrLf & "Failed to duplicate item record")
    '    '            Return ""
    '    '        Else
    '    '            Return sboNewItem.ItemCode
    '    '        End If
    '    objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    objMainRecset.DoQuery("Select ItemCode from OITM where U_Action='A'")
    '    For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
    '        FileName = System.Windows.Forms.Application.StartupPath & "\Item.xml"
    '        strItem = objMainRecset.Fields.Item(0).Value
    '        objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
    '        If objMainItem.GetByKey(strItem) Then
    '            objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
    '            objMainCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
    '            objMainItem.SaveXML(FileName)
    '            objRemoteCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
    '            objRemoteItem = objRemoteCompany.GetBusinessObjectFromXML(FileName, 0)
    '            If objRemoteItem.Add() <> 0 Then
    '                WriteErrorlog("Failed to create item : " & objRemoteCompany.GetLastErrorDescription, sPath)
    '            Else
    '                objMainItem.UserFields.Fields.Item("U_Action").Value = "N"
    '                objMainItem.Update()
    '                WriteErrorlog("New Item Created : " & objRemoteItem.ItemCode, sPath)
    '            End If
    '        End If
    '        objMainRecset.MoveNext()
    '    Next
    'End Sub
    'Private Sub UpdateItemMaster()
    '    Dim FileName, strItem As String
    '    Dim Ecount As Long
    '    Dim ii As Long
    '    Dim objMainItem, objRemoteItem As SAPbobsCOM.Items
    '    Dim objMainRecset As SAPbobsCOM.Recordset
    '    objMainRecset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '    objMainRecset.DoQuery("Select ItemCode from OITM where U_Action='U'")
    '    For intItemLoop As Integer = 0 To objMainRecset.RecordCount - 1
    '        FileName = System.Windows.Forms.Application.StartupPath & "\Item.xml"
    '        strItem = objMainRecset.Fields.Item(0).Value
    '        objMainItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
    '        If objMainItem.GetByKey(strItem) Then
    '            objRemoteItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
    '            If objRemoteItem.GetByKey(strItem) Then
    '                objMainCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
    '                objMainItem.SaveXML(FileName)
    '                objRemoteCompany.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode
    '                objRemoteItem = objRemoteCompany.GetBusinessObjectFromXML(FileName, 0)
    '                If objRemoteItem.Update <> 0 Then
    '                    WriteErrorlog("Failed to Update item : " & objRemoteItem.ItemCode & ":" & objRemoteCompany.GetLastErrorDescription, sPath)
    '                Else
    '                    objMainItem.UserFields.Fields.Item("U_Action").Value = "N"
    '                    objMainItem.Update()
    '                    WriteErrorlog("Update Completed : " & objRemoteItem.ItemCode, sPath)
    '                End If
    '            End If
    '        End If
    '        objMainRecset.MoveNext()
    '    Next
    'End Sub
#End Region

#Region "Export Item Details"


    Private Sub ExportStockTransfer(ByVal aCompany As SAPbobsCOM.Company)
        Dim conn As New SqlConnection()
        Dim objMainDoc As SAPbobsCOM.Documents
        Dim objremoteDoc As SAPbobsCOM.StockTransfer
        Dim strPath, strFilename, strMessage As String
        Dim strFileName1 As String
        Dim objremoteRec, objMainRec, objMainTemp As SAPbobsCOM.Recordset
        objMainRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainTemp = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objremoteRec = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRec.DoQuery("Select DocEntry,* from OWTR where  isnull(U_Export,'N')='N'")
        Dim strFromWhs As String
        Dim blnMainWhs As Boolean
        objMainRec.DoQuery("Select DocEntry,* from OWTR where DocStatus='O' and isnull(U_Export,'N')='N'")
        'Header fields – customer code, customer name, posting date, due date, document date, customer reference number, remarks, Total before discount, discount %, Tax, Total, Applied Amount, Balance Due
        'Line fields – item code, item name, barcode, quantity, unit price, discount %, price after discount, vat code, Gross price, Total (LC)
        For intRemoteLoop As Integer = 0 To objMainRec.RecordCount - 1
            Try
                strFromWhs = objMainRec.Fields.Item("Filler").Value
                'WriteErrorlog("Processing Main Branch Stock Transfer Number:" & objMainRec.Fields.Item(0).Value, sPath)
                blnMainWhs = False
                Dim strBranchCompanyDB As String
                Dim oBranchrS As SAPbobsCOM.Recordset
                Dim blnCurrentCompany As Boolean = False
                If strFromWhs <> "" Then
                    objMainTemp.DoQuery("Select * from [OWHS] where whscode='" & strFromWhs & "'")
                    If objMainTemp.Fields.Item("U_Warcode").Value <> "" Then
                        strBranchCompanyDB = objMainTemp.Fields.Item("U_Branch").Value
                        blnMainWhs = False
                        If strBranchCompanyDB = objRemoteCompany.CompanyDB Then
                            blnCurrentCompany = True
                        Else
                            blnCurrentCompany = False
                        End If
                    Else
                        strBranchCompanyDB = objMainTemp.Fields.Item("U_Branch").Value
                        blnMainWhs = True
                    End If
                End If
                oBranchrS = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                If blnMainWhs = True Then
                    Dim dblQty As Double
                    oBranchrS.DoQuery("Select * from WTR1 where DocEntry=" & objMainRec.Fields.Item(0).Value)
                    If oBranchrS.RecordCount > 0 Then
                        dblQty = oBranchrS.Fields.Item("Quantity").Value
                        Dim strBranchWarehouse As String = oBranchrS.Fields.Item("WhsCode").Value
                        oBranchrS.DoQuery("Select isnull(U_Branch,'') from OWHS where Whscode='" & strBranchWarehouse & "'")
                        If oBranchrS.Fields.Item(0).Value = objRemoteCompany.CompanyDB Then
                            blnCurrentCompany = True
                        Else
                            blnCurrentCompany = False
                        End If
                        If dblQty <= 0 Then
                            blnCurrentCompany = False
                        End If
                    Else
                        blnCurrentCompany = False
                    End If
                End If
                If blnCurrentCompany = False Then
                    ' WriteErrorlog("Stock Transfer Number:" & objMainRec.Fields.Item(0).Value & " Not relevant to this Branch DB : " & objRemoteCompany.CompanyDB, sPath)
                End If
                If blnCurrentCompany = True Then
                    '  WriteErrorlog("Processing Stock Transfer Number:" & objMainRec.Fields.Item(0).Value, sPath)
                    If blnMainWhs = True Then
                        objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
                        If objremoteDoc.GetByKey(Convert.ToInt32(objMainRec.Fields.Item(0).Value)) Then
                            WriteErrorlog("Processing Main Branch Stock Transfer Number:" & objremoteDoc.DocNum, sPath)
                            objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
                            objMainDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInventoryGenEntry
                            objMainDoc.GroupNumber = objremoteDoc.PriceList
                            objMainDoc.Comments = objremoteDoc.Comments
                            objMainDoc.DocDate = objremoteDoc.DocDate
                            objMainDoc.TaxDate = objremoteDoc.TaxDate
                            objMainDoc.UserFields.Fields.Item("U_Import").Value = "Y"
                            objMainDoc.UserFields.Fields.Item("U_BaseEntry").Value = objremoteDoc.DocEntry
                            objMainDoc.UserFields.Fields.Item("U_BaseNum").Value = objremoteDoc.DocNum
                            objMainDoc.UserFields.Fields.Item("U_Branch").Value = objMainCompany.CompanyName
                            For intLoop As Integer = 0 To objremoteDoc.Lines.Count - 1
                                objremoteDoc.Lines.SetCurrentLine(intLoop)
                                If 1 = 1 Then
                                    If intLoop > 0 Then
                                        objMainDoc.Lines.Add()
                                    End If
                                    objMainDoc.Lines.SetCurrentLine(intLoop)
                                    objMainDoc.Lines.ItemCode = objremoteDoc.Lines.ItemCode
                                    objMainDoc.Lines.Quantity = objremoteDoc.Lines.Quantity
                                    Dim strBranchwhs, strMainwhs As String
                                    strBranchwhs = objremoteDoc.Lines.WarehouseCode
                                    objMainTemp.DoQuery("Select isnull(U_Warcode,'" & strBranchwhs & "') from OWHS where WhsCode='" & strBranchwhs & "'")
                                    strMainwhs = objMainTemp.Fields.Item(0).Value
                                    objMainDoc.Lines.WarehouseCode = strMainwhs ' objremoteDoc.Lines.WarehouseCode 
                                    objMainDoc.Lines.UnitPrice = objremoteDoc.Lines.UnitPrice
                                    objMainDoc.GroupNumber = objremoteDoc.PriceList
                                    Dim dblLineQty, dblSerialCount As Double
                                    dblLineQty = objremoteDoc.Lines.Quantity
                                    dblSerialCount = objremoteDoc.Lines.SerialNumbers.Count
                                    Dim objItem As SAPbobsCOM.Items
                                    Dim stritem As String = objremoteDoc.Lines.ItemCode
                                    objItem = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                                    If objItem.GetByKey(objremoteDoc.Lines.ItemCode) Then
                                        If objItem.ManageSerialNumbers = SAPbobsCOM.BoYesNoEnum.tYES Then
                                            For intSer As Integer = 0 To objremoteDoc.Lines.SerialNumbers.Count - 1
                                                objremoteDoc.Lines.SerialNumbers.SetCurrentLine(intSer)
                                                'If intSer > 0 Then
                                                '    objMainDoc.Lines.SerialNumbers.Add()
                                                '    Try
                                                '        objMainDoc.Lines.SerialNumbers.SetCurrentLine(intSer)
                                                '    Catch ex As Exception
                                                '        '  MsgBox(ex.Message)
                                                '    End Try
                                                'End If

                                                'oInvGenEntry.Lines.SerialNumbers.SystemSerialNumber = CStr(i + 1)
                                                'oInvGenEntry.Lines.SerialNumbers.ManufacturerSerialNumber = manSerialNumber + "_" + CStr(i)
                                                'oInvGenEntry.Lines.SerialNumbers.InternalSerialNumber = i + 3
                                                'oInvGenEntry.Lines.SerialNumbers.SetCurrentLine(i)
                                                'oInvGenEntry.Lines.SerialNumbers.Add()

                                                '  objMainDoc.Lines.SerialNumbers.BaseLineNumber = objremoteDoc.Lines.SerialNumbers.BaseLineNumber
                                                '  objMainDoc.Lines.SerialNumbers.Quantity = 1
                                                objMainDoc.Lines.SerialNumbers.ExpiryDate = objremoteDoc.Lines.SerialNumbers.ExpiryDate
                                                objMainDoc.Lines.WarehouseCode = strMainwhs
                                                objMainDoc.Lines.SerialNumbers.InternalSerialNumber = objremoteDoc.Lines.SerialNumbers.InternalSerialNumber
                                                objMainDoc.Lines.SerialNumbers.ManufactureDate = objremoteDoc.Lines.SerialNumbers.ManufactureDate
                                                objMainDoc.Lines.SerialNumbers.ManufacturerSerialNumber = objremoteDoc.Lines.SerialNumbers.ManufacturerSerialNumber
                                                objMainDoc.Lines.SerialNumbers.Notes = objremoteDoc.Lines.SerialNumbers.Notes
                                                objMainDoc.Lines.SerialNumbers.ReceptionDate = objremoteDoc.Lines.SerialNumbers.ReceptionDate
                                                '  objMainDoc.Lines.SerialNumbers.SystemSerialNumber = objremoteDoc.Lines.SerialNumbers.SystemSerialNumber
                                                Try
                                                    objMainDoc.Lines.SerialNumbers.SetCurrentLine(intSer)
                                                Catch ex As Exception

                                                End Try

                                                objMainDoc.Lines.SerialNumbers.Add()
                                            Next

                                        End If

                                        If objItem.ManageBatchNumbers = SAPbobsCOM.BoYesNoEnum.tYES Then
                                            For intSer As Integer = 0 To objremoteDoc.Lines.BatchNumbers.Count - 1
                                                If intSer > 0 Then
                                                    objMainDoc.Lines.BatchNumbers.Add()
                                                    '   objMainDoc.Lines.BatchNumbers.SetCurrentLine(intSer)
                                                End If
                                                objMainDoc.Lines.BatchNumbers.SetCurrentLine(intSer)
                                                objremoteDoc.Lines.BatchNumbers.SetCurrentLine(intSer)
                                                objMainDoc.Lines.BatchNumbers.AddmisionDate = objremoteDoc.Lines.BatchNumbers.AddmisionDate
                                                objMainDoc.Lines.BatchNumbers.BaseLineNumber = objremoteDoc.Lines.BatchNumbers.BaseLineNumber
                                                objMainDoc.Lines.BatchNumbers.BatchNumber = objremoteDoc.Lines.BatchNumbers.BatchNumber
                                                objMainDoc.Lines.BatchNumbers.ExpiryDate = objremoteDoc.Lines.BatchNumbers.ExpiryDate
                                                objMainDoc.Lines.BatchNumbers.InternalSerialNumber = objremoteDoc.Lines.BatchNumbers.InternalSerialNumber
                                                objMainDoc.Lines.BatchNumbers.Location = objremoteDoc.Lines.BatchNumbers.Location
                                                objMainDoc.Lines.BatchNumbers.ManufacturingDate = objremoteDoc.Lines.BatchNumbers.ManufacturingDate
                                                objMainDoc.Lines.BatchNumbers.Notes = objremoteDoc.Lines.BatchNumbers.Notes
                                                objMainDoc.Lines.BatchNumbers.Quantity = objremoteDoc.Lines.BatchNumbers.Quantity
                                            Next
                                        End If
                                    End If
                                End If
                            Next
                            If objMainDoc.Add <> 0 Then
                                sPath = strErrorLogPath
                                WriteErrorlog("Failed to Create Goods Receipt docuemnt :" & objremoteDoc.DocNum & " Error : " & objRemoteCompany.GetLastErrorDescription, sPath)
                            Else
                                Dim strDocNum As String
                                objRemoteCompany.GetNewObjectCode(strDocNum)
                                sPath = strErrorLogPath
                                WriteErrorlog("Goods Receipt Created Successfully." & objRemoteCompany.CompanyDB & ": DocNum : " & strDocNum, sPath)
                                objremoteDoc.UserFields.Fields.Item("U_Export").Value = "Y"
                                objremoteDoc.Update()
                            End If
                        End If
                    Else 'Branch Whs to Headoffice
                        objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
                        If objremoteDoc.GetByKey(Convert.ToInt32(objMainRec.Fields.Item(0).Value)) Then
                            WriteErrorlog("Processing Main Branch Stock Transfer Number:" & objremoteDoc.DocNum, sPath)
                            objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
                            '   MsgBox(objremoteDoc.DocEntry)
                            'objMainDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInventoryGenExit
                            objMainDoc.Address = objremoteDoc.Address
                            'objMainDoc.Address2 = objremoteDoc.Address2
                            objMainDoc.CardCode = objremoteDoc.CardCode
                            objMainDoc.CardName = objremoteDoc.CardName
                            objMainDoc.Comments = objremoteDoc.Comments
                            objMainDoc.DocDate = objremoteDoc.DocDate
                            objMainDoc.DocDueDate = objremoteDoc.DueDate
                            objMainDoc.UserFields.Fields.Item("U_Import").Value = "Y"
                            objMainDoc.UserFields.Fields.Item("U_BaseEntry").Value = objremoteDoc.DocEntry
                            objMainDoc.UserFields.Fields.Item("U_BaseNum").Value = objremoteDoc.DocNum
                            objMainDoc.UserFields.Fields.Item("U_Branch").Value = objMainCompany.CompanyName
                            Dim strBranchwhs, strMainwhs As String
                            'strBranchwhs = objremoteDoc.FromWarehouse
                            'objremoteRec.DoQuery("Select isnull(U_Warcode,'" & strBranchwhs & "') from OWHS where WhsCode='" & strBranchwhs & "'")
                            'strMainwhs = objMainRec.Fields.Item(0).Value
                            For intLoop As Integer = 0 To objremoteDoc.Lines.Count - 1
                                If intLoop > 0 Then
                                    objMainDoc.Lines.Add()
                                End If
                                objMainDoc.Lines.SetCurrentLine(intLoop)
                                objremoteDoc.Lines.SetCurrentLine(intLoop)
                                objMainDoc.Lines.ItemCode = objremoteDoc.Lines.ItemCode
                                objMainDoc.Lines.Quantity = objremoteDoc.Lines.Quantity
                                strBranchwhs = objremoteDoc.FromWarehouse
                                objMainTemp.DoQuery("Select isnull(U_Warcode,'" & strBranchwhs & "') from OWHS where WhsCode='" & strBranchwhs & "'")
                                strMainwhs = objMainTemp.Fields.Item(0).Value
                                objMainDoc.Lines.WarehouseCode = strMainwhs ' objremoteDoc.Lines.WarehouseCode
                                '   MsgBox(objremoteDoc.Lines.UnitPrice)
                                objMainDoc.Lines.UnitPrice = objremoteDoc.Lines.UnitPrice
                                '   objMainDoc.Lines.TaxCode = objremoteDoc.Lines.TaxCode
                                For intSer As Integer = 0 To objremoteDoc.Lines.SerialNumbers.Count - 1
                                    objremoteDoc.Lines.SerialNumbers.SetCurrentLine(intSer)
                                    If intSer > 0 Then
                                        objMainDoc.Lines.SerialNumbers.Add()
                                    End If
                                    objremoteDoc.Lines.SerialNumbers.SetCurrentLine(intSer)
                                    'objMainDoc.Lines.SerialNumbers.ExpiryDate = objremoteDoc.Lines.SerialNumbers.ExpiryDate
                                    'objMainDoc.Lines.WarehouseCode = strMainwhs
                                    'objMainDoc.Lines.SerialNumbers.InternalSerialNumber = objremoteDoc.Lines.SerialNumbers.InternalSerialNumber
                                    'objMainDoc.Lines.SerialNumbers.ManufactureDate = objremoteDoc.Lines.SerialNumbers.ManufactureDate
                                    'objMainDoc.Lines.SerialNumbers.ManufacturerSerialNumber = objremoteDoc.Lines.SerialNumbers.ManufacturerSerialNumber
                                    'objMainDoc.Lines.SerialNumbers.Notes = objremoteDoc.Lines.SerialNumbers.Notes
                                    'objMainDoc.Lines.SerialNumbers.ReceptionDate = objremoteDoc.Lines.SerialNumbers.ReceptionDate

                                    objMainDoc.Lines.SerialNumbers.BaseLineNumber = objremoteDoc.Lines.SerialNumbers.BaseLineNumber
                                    objMainDoc.Lines.WarehouseCode = strMainwhs
                                    ' objMainDoc.Lines.SerialNumbers.ExpiryDate = objremoteDoc.Lines.SerialNumbers.ExpiryDate
                                    objMainDoc.Lines.SerialNumbers.InternalSerialNumber = objremoteDoc.Lines.SerialNumbers.InternalSerialNumber
                                    objMainDoc.Lines.SerialNumbers.ManufactureDate = objremoteDoc.Lines.SerialNumbers.ManufactureDate
                                    objMainDoc.Lines.SerialNumbers.ManufacturerSerialNumber = objremoteDoc.Lines.SerialNumbers.ManufacturerSerialNumber
                                    objMainDoc.Lines.SerialNumbers.Notes = objremoteDoc.Lines.SerialNumbers.Notes
                                    objMainDoc.Lines.SerialNumbers.ReceptionDate = objremoteDoc.Lines.SerialNumbers.ReceptionDate
                                    'If intSer > 0 Then
                                    '    objMainDoc.Lines.SerialNumbers.Add()
                                    'End If
                                    'objremoteDoc.Lines.SerialNumbers.SetCurrentLine(intSer)
                                Next

                                For intSer As Integer = 0 To objremoteDoc.Lines.BatchNumbers.Count - 1
                                    If intSer > 0 Then
                                        objMainDoc.Lines.BatchNumbers.Add()
                                    End If
                                    objremoteDoc.Lines.BatchNumbers.SetCurrentLine(intSer)
                                    objMainDoc.Lines.BatchNumbers.SetCurrentLine(intSer)
                                    objMainDoc.Lines.BatchNumbers.AddmisionDate = objremoteDoc.Lines.BatchNumbers.AddmisionDate
                                    objMainDoc.Lines.BatchNumbers.BaseLineNumber = objremoteDoc.Lines.BatchNumbers.BaseLineNumber
                                    objMainDoc.Lines.BatchNumbers.BatchNumber = objremoteDoc.Lines.BatchNumbers.BatchNumber
                                    objMainDoc.Lines.BatchNumbers.ExpiryDate = objremoteDoc.Lines.BatchNumbers.ExpiryDate
                                    objMainDoc.Lines.BatchNumbers.InternalSerialNumber = objremoteDoc.Lines.BatchNumbers.InternalSerialNumber
                                    objMainDoc.Lines.BatchNumbers.Location = objremoteDoc.Lines.BatchNumbers.Location
                                    objMainDoc.Lines.BatchNumbers.ManufacturingDate = objremoteDoc.Lines.BatchNumbers.ManufacturingDate
                                    objMainDoc.Lines.BatchNumbers.Notes = objremoteDoc.Lines.BatchNumbers.Notes
                                    objMainDoc.Lines.BatchNumbers.Quantity = objremoteDoc.Lines.BatchNumbers.Quantity
                                Next
                            Next
                            If objMainDoc.Add <> 0 Then
                                sPath = strErrorLogPath
                                WriteErrorlog("Failed to create Goods Issue docuemnt :" & objremoteDoc.DocNum & " Error : " & objRemoteCompany.GetLastErrorDescription, sPath)
                            Else
                                Dim strDocNum As String
                                objRemoteCompany.GetNewObjectCode(strDocNum)
                                sPath = strErrorLogPath
                                WriteErrorlog("Goods Issue Created Successfully : " & objRemoteCompany.CompanyDB & " : DocNum : " & strDocNum, sPath)
                                objremoteDoc.UserFields.Fields.Item("U_Export").Value = "Y"
                                objremoteDoc.Update()
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception
                WriteErrorlog("Failed to Process Stock transfer Number  : " & objMainRec.Fields.Item(0).Value, sPath)

            End Try
            objMainRec.MoveNext()
        Next
    End Sub
    Private Sub ExportSalesOrer(ByVal aCompany As SAPbobsCOM.Company)
        Dim conn As New SqlConnection()
        Dim objMainDoc, objremoteDoc As SAPbobsCOM.Documents
        Dim strPath, strFilename, strMessage As String
        Dim strFileName1, strBranchWhs, strMainWhs As String
        Dim objremoteRec, objMainRec As SAPbobsCOM.Recordset
        objremoteRec = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objremoteRec.DoQuery("Select DocEntry from OINV where IsICT='N' and  isnull(U_Export,'N')='N'")
        'Header fields – customer code, customer name, posting date, due date, document date, customer reference number, remarks, Total before discount, discount %, Tax, Total, Applied Amount, Balance Due
        'Line fields – item code, item name, barcode, quantity, unit price, discount %, price after discount, vat code, Gross price, Total (LC)
        For intRemoteLoop As Integer = 0 To objremoteRec.RecordCount - 1
            objremoteDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
            If objremoteDoc.GetByKey(Convert.ToInt32(objremoteRec.Fields.Item(0).Value)) Then
                objMainDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                objMainDoc.Address = objremoteDoc.Address
                objMainDoc.Address2 = objremoteDoc.Address2
                objMainDoc.CardCode = objremoteDoc.CardCode
                objMainDoc.CardName = objremoteDoc.CardName
                objMainDoc.CentralBankIndicator = objremoteDoc.CentralBankIndicator
                objMainDoc.ClosingRemarks = objremoteDoc.ClosingRemarks
                objMainDoc.Comments = objremoteDoc.Comments
                objMainDoc.ContactPersonCode = objremoteDoc.ContactPersonCode
                objMainDoc.DeferredTax = objremoteDoc.DeferredTax
                objMainDoc.DiscountPercent = objremoteDoc.DiscountPercent
                objMainDoc.DocCurrency = objremoteDoc.DocCurrency
                objMainDoc.DocDate = objremoteDoc.DocDate
                objMainDoc.DocDueDate = objremoteDoc.DocDueDate
                objMainDoc.DocRate = objremoteDoc.DocRate
                objMainDoc.DocTotal = objremoteDoc.DocTotal
                objMainDoc.DocType = objremoteDoc.DocType
                objMainDoc.DocumentSubType = objremoteDoc.DocumentSubType
                objMainDoc.NumAtCard = objremoteDoc.NumAtCard
                objMainDoc.Comments = objremoteDoc.Comments
                objMainDoc.DiscountPercent = objremoteDoc.DiscountPercent
                objMainDoc.DocCurrency = objremoteDoc.DocCurrency
                objMainDoc.ShipToCode = objremoteDoc.ShipToCode
                objMainDoc.SalesPersonCode = objremoteDoc.SalesPersonCode
                objMainDoc.TaxDate = objremoteDoc.TaxDate
                objMainDoc.PaymentGroupCode = objremoteDoc.PaymentGroupCode
                objMainDoc.PaymentMethod = objremoteDoc.PaymentMethod
                'objMainDoc.TransportationCode = objremoteDoc.TransportationCode
                objMainDoc.SalesPersonCode = objremoteDoc.SalesPersonCode
                If objremoteDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES Then
                    objMainDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES
                    objMainDoc.RoundingDiffAmount = objremoteDoc.RoundingDiffAmount
                Else
                    objMainDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tNO
                End If
                objMainDoc.UserFields.Fields.Item("U_Import").Value = "Y"
                objMainDoc.UserFields.Fields.Item("U_BaseEntry").Value = objremoteDoc.DocEntry
                objMainDoc.UserFields.Fields.Item("U_BaseNum").Value = (objremoteDoc.DocNum)
                objMainDoc.UserFields.Fields.Item("U_Branch").Value = objRemoteCompany.CompanyName
                For IntExp As Integer = 0 To objremoteDoc.Expenses.Count - 1
                    If objremoteDoc.Expenses.LineTotal > 0 Then
                        If IntExp > 0 Then
                            objMainDoc.Expenses.Add()
                            objMainDoc.Expenses.SetCurrentLine(IntExp)
                        End If
                        objremoteDoc.Expenses.SetCurrentLine(IntExp)
                        objMainDoc.Expenses.BaseDocEntry = objremoteDoc.Expenses.BaseDocEntry
                        objMainDoc.Expenses.BaseDocLine = objremoteDoc.Expenses.BaseDocLine
                        objMainDoc.Expenses.BaseDocType = objremoteDoc.Expenses.BaseDocType
                        objMainDoc.Expenses.DistributionMethod = objremoteDoc.Expenses.DistributionMethod
                        objMainDoc.Expenses.DistributionRule = objremoteDoc.Expenses.DistributionRule
                        objMainDoc.Expenses.ExpenseCode = objremoteDoc.Expenses.ExpenseCode
                        objMainDoc.Expenses.LastPurchasePrice = objremoteDoc.Expenses.LastPurchasePrice
                        objMainDoc.Expenses.LineTotal = objremoteDoc.Expenses.LineTotal
                        objMainDoc.Expenses.Remarks = objremoteDoc.Expenses.Remarks
                        objMainDoc.Expenses.TaxCode = objremoteDoc.Expenses.TaxCode
                        objMainDoc.Expenses.VatGroup = objremoteDoc.Expenses.VatGroup
                    End If
                Next

                For intLoop As Integer = 0 To objremoteDoc.Lines.Count - 1
                    If intLoop > 0 Then
                        objMainDoc.Lines.Add()
                        objMainDoc.Lines.SetCurrentLine(intLoop)
                    End If
                    objremoteDoc.Lines.SetCurrentLine(intLoop)
                    objMainDoc.Lines.AccountCode = objremoteDoc.Lines.AccountCode
                    objMainDoc.Lines.ItemDescription = objremoteDoc.Lines.ItemDescription
                    objMainDoc.Lines.ItemCode = objremoteDoc.Lines.ItemCode
                    objMainDoc.Lines.BarCode = objremoteDoc.Lines.BarCode
                    objMainDoc.Lines.UnitPrice = objremoteDoc.Lines.UnitPrice
                    objMainDoc.Lines.DiscountPercent = objremoteDoc.Lines.DiscountPercent
                    objMainDoc.Lines.VatGroup = objremoteDoc.Lines.VatGroup
                    objMainDoc.Lines.PriceAfterVAT = objremoteDoc.Lines.PriceAfterVAT
                    objMainDoc.Lines.LineTotal = objremoteDoc.Lines.LineTotal
                    objMainDoc.Lines.ProjectCode = objremoteDoc.Lines.ProjectCode
                    objMainDoc.Lines.Quantity = objremoteDoc.Lines.Quantity
                    strBranchWhs = objremoteDoc.Lines.WarehouseCode
                    objMainDoc.Lines.CostingCode = objremoteDoc.Lines.CostingCode
                    objMainDoc.Lines.CostingCode2 = objremoteDoc.Lines.CostingCode2
                    objMainDoc.Lines.CostingCode3 = objremoteDoc.Lines.CostingCode3
                    objMainDoc.Lines.CostingCode4 = objremoteDoc.Lines.CostingCode4
                    objMainDoc.Lines.CostingCode5 = objremoteDoc.Lines.CostingCode5


                    objMainRec.DoQuery("Select isnull(WhsCode,'" & strBranchWhs & "') from OWHS where U_Warcode='" & strBranchWhs & "'")
                    strMainWhs = objMainRec.Fields.Item(0).Value


                    objMainDoc.Lines.WarehouseCode = strMainWhs ' objremoteDoc.Lines.WarehouseCode
                    objMainDoc.Lines.TaxCode = objremoteDoc.Lines.TaxCode
                    ' MsgBox(objremoteDoc.Lines.SerialNumbers.Count)

                    For intSer As Integer = 0 To objremoteDoc.Lines.SerialNumbers.Count - 1
                        If intSer > 0 Then
                            objMainDoc.Lines.SerialNumbers.Add()
                            objMainDoc.Lines.SerialNumbers.SetCurrentLine(intSer)
                        End If
                        objremoteDoc.Lines.SerialNumbers.SetCurrentLine(intSer)
                        objMainDoc.Lines.SerialNumbers.BaseLineNumber = objremoteDoc.Lines.SerialNumbers.BaseLineNumber
                        '  objMainDoc.Lines.SerialNumbers.BatchID = objremoteDoc.Lines.SerialNumbers.BatchID
                        objMainDoc.Lines.SerialNumbers.ExpiryDate = objremoteDoc.Lines.SerialNumbers.ExpiryDate
                        objMainDoc.Lines.SerialNumbers.InternalSerialNumber = objremoteDoc.Lines.SerialNumbers.InternalSerialNumber
                        ' objMainDoc.Lines.SerialNumbers.Location = objremoteDoc.Lines.SerialNumbers.Location
                        objMainDoc.Lines.SerialNumbers.ManufactureDate = objremoteDoc.Lines.SerialNumbers.ManufactureDate
                        objMainDoc.Lines.SerialNumbers.ManufacturerSerialNumber = objremoteDoc.Lines.SerialNumbers.ManufacturerSerialNumber
                        objMainDoc.Lines.SerialNumbers.Notes = objremoteDoc.Lines.SerialNumbers.Notes
                        objMainDoc.Lines.SerialNumbers.ReceptionDate = objremoteDoc.Lines.SerialNumbers.ReceptionDate
                        ' objMainDoc.Lines.SerialNumbers.SystemSerialNumber = objremoteDoc.Lines.SerialNumbers.SystemSerialNumber
                    Next

                    For intSer As Integer = 0 To objremoteDoc.Lines.BatchNumbers.Count - 1
                        If intSer > 0 Then
                            objMainDoc.Lines.BatchNumbers.Add()
                            objMainDoc.Lines.BatchNumbers.SetCurrentLine(intSer)
                        End If
                        objremoteDoc.Lines.BatchNumbers.SetCurrentLine(intSer)
                        objMainDoc.Lines.BatchNumbers.AddmisionDate = objremoteDoc.Lines.BatchNumbers.AddmisionDate
                        objMainDoc.Lines.BatchNumbers.BaseLineNumber = objremoteDoc.Lines.BatchNumbers.BaseLineNumber
                        objMainDoc.Lines.BatchNumbers.BatchNumber = objremoteDoc.Lines.BatchNumbers.BatchNumber
                        objMainDoc.Lines.BatchNumbers.ExpiryDate = objremoteDoc.Lines.BatchNumbers.ExpiryDate
                        objMainDoc.Lines.BatchNumbers.InternalSerialNumber = objremoteDoc.Lines.BatchNumbers.InternalSerialNumber
                        objMainDoc.Lines.BatchNumbers.Location = objremoteDoc.Lines.BatchNumbers.Location
                        objMainDoc.Lines.BatchNumbers.ManufacturingDate = objremoteDoc.Lines.BatchNumbers.ManufacturingDate
                        objMainDoc.Lines.BatchNumbers.Notes = objremoteDoc.Lines.BatchNumbers.Notes
                        objMainDoc.Lines.BatchNumbers.Quantity = objremoteDoc.Lines.BatchNumbers.Quantity
                    Next
                Next
                If objMainDoc.Add <> 0 Then
                    sPath = strErrorLogPath
                    WriteErrorlog("Failed to create invoice docuemnt :" & objremoteDoc.DocNum & " Error : " & objMainCompany.GetLastErrorDescription, sPath)
                    '    MsgBox(objMainCompany.GetLastErrorDescription)
                Else
                    Dim strDocNum As String
                    objMainCompany.GetNewObjectCode(strDocNum)
                    sPath = strErrorLogPath
                    WriteErrorlog("Invoice Created Successfully.DocNum : " & strDocNum, sPath)
                    objremoteDoc.UserFields.Fields.Item("U_Export").Value = "Y"
                    objremoteDoc.Update()
                End If
            End If
            objremoteRec.MoveNext()
        Next
    End Sub

    Private Sub ExportInvoicePayment(ByVal aCompany As SAPbobsCOM.Company)
        Dim conn As New SqlConnection()
        Dim objMainDoc, objremoteDoc As SAPbobsCOM.Documents
        Dim intRemoteDocEntry As Integer
        Dim strPath, strFilename, strMessage As String
        Dim strFileName1, strBranchWhs, strMainWhs As String
        Dim objremoteRec, objMainRec As SAPbobsCOM.Recordset
        objremoteRec = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objMainRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'objremoteRec.DoQuery("Select DocEntry from OINV where IsICT='Y' and  DocType='I' and  isnull(U_Export,'N')='N'")
        objremoteRec.DoQuery("Select DocEntry from OINV where IsICT='Y' and  isnull(U_Export,'N')='N'")
        For intRemoteLoop As Integer = 0 To objremoteRec.RecordCount - 1
            objremoteDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
            If objremoteDoc.GetByKey(Convert.ToInt32(objremoteRec.Fields.Item(0).Value)) Then
                UpdateExchangerate_Main("USD", objremoteDoc.DocDate, objMainCompany)
                intRemoteDocEntry = objremoteDoc.DocEntry
                objMainDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                objMainDoc.CardCode = objremoteDoc.CardCode
                objMainDoc.DocDate = objremoteDoc.DocDate
                objMainDoc.DocDueDate = objremoteDoc.DocDueDate
                objMainDoc.NumAtCard = objremoteDoc.NumAtCard
                objMainDoc.Comments = objremoteDoc.Comments
                objMainDoc.DiscountPercent = objremoteDoc.DiscountPercent
                objMainDoc.DocCurrency = objremoteDoc.DocCurrency
                objMainDoc.ShipToCode = objremoteDoc.ShipToCode
                objMainDoc.SalesPersonCode = objremoteDoc.SalesPersonCode
                objMainDoc.TaxDate = objremoteDoc.TaxDate
                objMainDoc.PaymentGroupCode = objremoteDoc.PaymentGroupCode
                objMainDoc.PaymentMethod = objremoteDoc.PaymentMethod
                objMainDoc.Address = objremoteDoc.Address
                objMainDoc.Address2 = objremoteDoc.Address2
                objMainDoc.AgentCode = objremoteDoc.AgentCode
                objMainDoc.BPChannelCode = objremoteDoc.BPChannelCode
                objMainDoc.BPChannelContact = objremoteDoc.BPChannelContact
                objMainDoc.ContactPersonCode = objremoteDoc.ContactPersonCode
                objMainDoc.DocRate = objremoteDoc.DocRate
                objMainDoc.DocType = objremoteDoc.DocType
                '  objMainDoc.TransportationCode = objremoteDoc.TransportationCode
                If objremoteDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES Then
                    objMainDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES
                    objMainDoc.RoundingDiffAmount = objremoteDoc.RoundingDiffAmount
                Else
                    objMainDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tNO
                End If
                objMainDoc.UserFields.Fields.Item("U_Import").Value = "Y"
                objMainDoc.UserFields.Fields.Item("U_BaseEntry").Value = objremoteDoc.DocEntry
                objMainDoc.UserFields.Fields.Item("U_BaseNum").Value = (objremoteDoc.DocNum)
                objMainDoc.UserFields.Fields.Item("U_Branch").Value = objRemoteCompany.CompanyName
                For IntExp As Integer = 0 To objremoteDoc.Expenses.Count - 1
                    If objremoteDoc.Expenses.LineTotal > 0 Then
                        If IntExp > 0 Then
                            objMainDoc.Expenses.Add()
                            objMainDoc.Expenses.SetCurrentLine(IntExp)
                        End If
                        objremoteDoc.Expenses.SetCurrentLine(IntExp)
                        objMainDoc.Expenses.BaseDocEntry = objremoteDoc.Expenses.BaseDocEntry
                        objMainDoc.Expenses.BaseDocLine = objremoteDoc.Expenses.BaseDocLine
                        objMainDoc.Expenses.BaseDocType = objremoteDoc.Expenses.BaseDocType
                        objMainDoc.Expenses.DistributionMethod = objremoteDoc.Expenses.DistributionMethod
                        objMainDoc.Expenses.DistributionRule = objremoteDoc.Expenses.DistributionRule
                        objMainDoc.Expenses.ExpenseCode = objremoteDoc.Expenses.ExpenseCode
                        objMainDoc.Expenses.LastPurchasePrice = objremoteDoc.Expenses.LastPurchasePrice
                        objMainDoc.Expenses.LineTotal = objremoteDoc.Expenses.LineTotal
                        objMainDoc.Expenses.Remarks = objremoteDoc.Expenses.Remarks
                        objMainDoc.Expenses.TaxCode = objremoteDoc.Expenses.TaxCode
                        objMainDoc.Expenses.VatGroup = objremoteDoc.Expenses.VatGroup
                    End If
                Next

                For intLoop As Integer = 0 To objremoteDoc.Lines.Count - 1
                    If intLoop > 0 Then
                        objMainDoc.Lines.Add()
                        objMainDoc.Lines.SetCurrentLine(intLoop)
                    End If
                    objremoteDoc.Lines.SetCurrentLine(intLoop)
                    objMainDoc.Lines.AccountCode = objremoteDoc.Lines.AccountCode
                    objMainDoc.Lines.ItemDescription = objremoteDoc.Lines.ItemDescription
                    objMainDoc.Lines.ItemCode = objremoteDoc.Lines.ItemCode
                    objMainDoc.Lines.BarCode = objremoteDoc.Lines.BarCode
                    objMainDoc.Lines.UnitPrice = objremoteDoc.Lines.UnitPrice
                    objMainDoc.Lines.DiscountPercent = objremoteDoc.Lines.DiscountPercent
                    objMainDoc.Lines.VatGroup = objremoteDoc.Lines.VatGroup
                    objMainDoc.Lines.PriceAfterVAT = objremoteDoc.Lines.PriceAfterVAT
                    objMainDoc.Lines.LineTotal = objremoteDoc.Lines.LineTotal
                    objMainDoc.Lines.ProjectCode = objremoteDoc.Lines.ProjectCode
                    objMainDoc.Lines.Quantity = objremoteDoc.Lines.Quantity

                    objMainDoc.Lines.CostingCode = objremoteDoc.Lines.CostingCode
                    objMainDoc.Lines.CostingCode2 = objremoteDoc.Lines.CostingCode2
                    objMainDoc.Lines.CostingCode3 = objremoteDoc.Lines.CostingCode3
                    objMainDoc.Lines.CostingCode4 = objremoteDoc.Lines.CostingCode4
                    objMainDoc.Lines.CostingCode5 = objremoteDoc.Lines.CostingCode5

                    strBranchWhs = objremoteDoc.Lines.WarehouseCode
                    objMainRec.DoQuery("Select isnull(WhsCode,'" & strBranchWhs & "') from OWHS where U_Warcode='" & strBranchWhs & "'")
                    strMainWhs = objMainRec.Fields.Item(0).Value


                    objMainDoc.Lines.WarehouseCode = strMainWhs ' objremoteDoc.Lines.WarehouseCode
                    For intSer As Integer = 0 To objremoteDoc.Lines.SerialNumbers.Count - 1
                        If intSer > 0 Then
                            objMainDoc.Lines.SerialNumbers.Add()
                            objMainDoc.Lines.SerialNumbers.SetCurrentLine(intSer)
                        End If
                        objremoteDoc.Lines.SerialNumbers.SetCurrentLine(intSer)
                        objMainDoc.Lines.SerialNumbers.BaseLineNumber = objremoteDoc.Lines.SerialNumbers.BaseLineNumber
                        '  objMainDoc.Lines.SerialNumbers.BatchID = objremoteDoc.Lines.SerialNumbers.BatchID
                        objMainDoc.Lines.SerialNumbers.ExpiryDate = objremoteDoc.Lines.SerialNumbers.ExpiryDate
                        objMainDoc.Lines.SerialNumbers.InternalSerialNumber = objremoteDoc.Lines.SerialNumbers.InternalSerialNumber
                        ' objMainDoc.Lines.SerialNumbers.Location = objremoteDoc.Lines.SerialNumbers.Location
                        objMainDoc.Lines.SerialNumbers.ManufactureDate = objremoteDoc.Lines.SerialNumbers.ManufactureDate
                        objMainDoc.Lines.SerialNumbers.ManufacturerSerialNumber = objremoteDoc.Lines.SerialNumbers.ManufacturerSerialNumber
                        objMainDoc.Lines.SerialNumbers.Notes = objremoteDoc.Lines.SerialNumbers.Notes
                        objMainDoc.Lines.SerialNumbers.ReceptionDate = objremoteDoc.Lines.SerialNumbers.ReceptionDate
                        ' objMainDoc.Lines.SerialNumbers.SystemSerialNumber = objremoteDoc.Lines.SerialNumbers.SystemSerialNumber
                    Next

                    For intSer As Integer = 0 To objremoteDoc.Lines.BatchNumbers.Count - 1
                        If intSer > 0 Then
                            objMainDoc.Lines.BatchNumbers.Add()
                            objMainDoc.Lines.BatchNumbers.SetCurrentLine(intSer)
                        End If
                        objremoteDoc.Lines.BatchNumbers.SetCurrentLine(intSer)
                        objMainDoc.Lines.BatchNumbers.AddmisionDate = objremoteDoc.Lines.BatchNumbers.AddmisionDate
                        objMainDoc.Lines.BatchNumbers.BaseLineNumber = objremoteDoc.Lines.BatchNumbers.BaseLineNumber
                        objMainDoc.Lines.BatchNumbers.BatchNumber = objremoteDoc.Lines.BatchNumbers.BatchNumber
                        objMainDoc.Lines.BatchNumbers.ExpiryDate = objremoteDoc.Lines.BatchNumbers.ExpiryDate
                        objMainDoc.Lines.BatchNumbers.InternalSerialNumber = objremoteDoc.Lines.BatchNumbers.InternalSerialNumber
                        objMainDoc.Lines.BatchNumbers.Location = objremoteDoc.Lines.BatchNumbers.Location
                        objMainDoc.Lines.BatchNumbers.ManufacturingDate = objremoteDoc.Lines.BatchNumbers.ManufacturingDate
                        objMainDoc.Lines.BatchNumbers.Notes = objremoteDoc.Lines.BatchNumbers.Notes
                        objMainDoc.Lines.BatchNumbers.Quantity = objremoteDoc.Lines.BatchNumbers.Quantity
                    Next
                Next
                If objMainDoc.Add <> 0 Then
                    sPath = strErrorLogPath
                    WriteErrorlog("Failed to create invoice+Payment docuemnt :" & objMainCompany.GetLastErrorDescription, sPath)
                Else
                    Dim strDocNum As String
                    objMainCompany.GetNewObjectCode(strDocNum)
                    sPath = strErrorLogPath
                    WriteErrorlog("Invoice + Payment Created Successfully. DocNum : " & strDocNum, sPath)
                    Dim otempRec As SAPbobsCOM.Recordset
                    otempRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    otempRec.DoQuery("Update OINV set IsICT='Y' where Docentry=" & strDocNum)
                    intRemoteDocEntry = objremoteDoc.DocEntry
                    ExportIncomingPayment(oCompany, GetIncomingDocument(intRemoteDocEntry))
                    objremoteDoc.UserFields.Fields.Item("U_Export").Value = "Y"
                    objremoteDoc.Update()
                End If
            End If
            objremoteRec.MoveNext()
        Next

    End Sub

    Private Sub ImportCreditNotes(ByVal aCompany As SAPbobsCOM.Company)
        Dim conn As New SqlConnection()
        Dim objMainDoc, objremoteDoc As SAPbobsCOM.Documents
        Dim strPath, strFilename, strMessage As String
        Dim strFileName1, strBranchWhs, strMainWhs As String
        Dim objremoteRec, objMainrec As SAPbobsCOM.Recordset
        Try
            objremoteRec = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objMainrec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objremoteRec.DoQuery("Select DocEntry from ORIN where   isnull(U_Export,'N')='N'")

            For intRemoteLoop As Integer = 0 To objremoteRec.RecordCount - 1
                objremoteDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                If objremoteDoc.GetByKey(Convert.ToInt32(objremoteRec.Fields.Item(0).Value)) Then
                    objMainDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                    objMainDoc.CardCode = objremoteDoc.CardCode
                    objMainDoc.DocDate = objremoteDoc.DocDate
                    objMainDoc.DocDueDate = objremoteDoc.DocDueDate
                    objMainDoc.NumAtCard = objremoteDoc.NumAtCard
                    objMainDoc.Comments = objremoteDoc.Comments
                    objMainDoc.DiscountPercent = objremoteDoc.DiscountPercent
                    objMainDoc.DocCurrency = objremoteDoc.DocCurrency
                    objMainDoc.ShipToCode = objremoteDoc.ShipToCode
                    objMainDoc.SalesPersonCode = objremoteDoc.SalesPersonCode
                    objMainDoc.TaxDate = objremoteDoc.TaxDate
                    objMainDoc.PaymentGroupCode = objremoteDoc.PaymentGroupCode
                    objMainDoc.Address = objremoteDoc.Address
                    objMainDoc.Address2 = objremoteDoc.Address2
                    objMainDoc.AgentCode = objremoteDoc.AgentCode
                    objMainDoc.BPChannelCode = objremoteDoc.BPChannelCode
                    objMainDoc.BPChannelContact = objremoteDoc.BPChannelContact
                    objMainDoc.ContactPersonCode = objremoteDoc.ContactPersonCode
                    objMainDoc.DocRate = objremoteDoc.DocRate
                    objMainDoc.DocType = objremoteDoc.DocType
                    '  objMainDoc.TransportationCode = objremoteDoc.TransportationCode
                    If objremoteDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES Then
                        objMainDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES
                        objMainDoc.RoundingDiffAmount = objremoteDoc.RoundingDiffAmount
                    Else
                        objMainDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tNO
                    End If

                    'Document_LinesClass' 
                    objMainDoc.UserFields.Fields.Item("U_Import").Value = "Y"
                    objMainDoc.UserFields.Fields.Item("U_BaseEntry").Value = objremoteDoc.DocEntry
                    objMainDoc.UserFields.Fields.Item("U_BaseNum").Value = (objremoteDoc.DocNum)
                    objMainDoc.UserFields.Fields.Item("U_Branch").Value = objRemoteCompany.CompanyName
                    For IntExp As Integer = 0 To objremoteDoc.Expenses.Count - 1
                        If objremoteDoc.Expenses.LineTotal > 0 Then
                            If IntExp > 0 Then
                                objMainDoc.Expenses.Add()
                                objMainDoc.Expenses.SetCurrentLine(IntExp)
                            End If
                            objremoteDoc.Expenses.SetCurrentLine(IntExp)
                            objMainDoc.Expenses.BaseDocEntry = objremoteDoc.Expenses.BaseDocEntry
                            objMainDoc.Expenses.BaseDocLine = objremoteDoc.Expenses.BaseDocLine
                            objMainDoc.Expenses.BaseDocType = objremoteDoc.Expenses.BaseDocType
                            objMainDoc.Expenses.DistributionMethod = objremoteDoc.Expenses.DistributionMethod
                            objMainDoc.Expenses.DistributionRule = objremoteDoc.Expenses.DistributionRule
                            objMainDoc.Expenses.ExpenseCode = objremoteDoc.Expenses.ExpenseCode
                            objMainDoc.Expenses.LastPurchasePrice = objremoteDoc.Expenses.LastPurchasePrice
                            objMainDoc.Expenses.LineTotal = objremoteDoc.Expenses.LineTotal
                            objMainDoc.Expenses.Remarks = objremoteDoc.Expenses.Remarks
                            objMainDoc.Expenses.TaxCode = objremoteDoc.Expenses.TaxCode
                            objMainDoc.Expenses.VatGroup = objremoteDoc.Expenses.VatGroup
                        End If
                    Next

                    For intLoop As Integer = 0 To objremoteDoc.Lines.Count - 1
                        If intLoop > 0 Then
                            objMainDoc.Lines.Add()
                            objMainDoc.Lines.SetCurrentLine(intLoop)
                        End If
                        objremoteDoc.Lines.SetCurrentLine(intLoop)
                        '    objMainDoc.Lines.SerialNumbers.BaseLineNumber = objremoteDoc.Lines.SerialNumbers.BaseLineNumber
                        objMainDoc.Lines.AccountCode = objremoteDoc.Lines.AccountCode
                        objMainDoc.Lines.ItemDescription = objremoteDoc.Lines.ItemDescription
                        objMainDoc.Lines.ItemCode = objremoteDoc.Lines.ItemCode
                        objMainDoc.Lines.BarCode = objremoteDoc.Lines.BarCode
                        objMainDoc.Lines.UnitPrice = objremoteDoc.Lines.UnitPrice
                        objMainDoc.Lines.AccountCode = objremoteDoc.Lines.AccountCode
                        objMainDoc.Lines.DiscountPercent = objremoteDoc.Lines.DiscountPercent
                        objMainDoc.Lines.VatGroup = objremoteDoc.Lines.VatGroup
                        objMainDoc.Lines.PriceAfterVAT = objremoteDoc.Lines.PriceAfterVAT
                        objMainDoc.Lines.LineTotal = objremoteDoc.Lines.LineTotal
                        objMainDoc.Lines.ProjectCode = objremoteDoc.Lines.ProjectCode
                        objMainDoc.Lines.Quantity = objremoteDoc.Lines.Quantity

                        objMainDoc.Lines.CostingCode = objremoteDoc.Lines.CostingCode
                        objMainDoc.Lines.CostingCode2 = objremoteDoc.Lines.CostingCode2
                        objMainDoc.Lines.CostingCode3 = objremoteDoc.Lines.CostingCode3
                        objMainDoc.Lines.CostingCode4 = objremoteDoc.Lines.CostingCode4
                        objMainDoc.Lines.CostingCode5 = objremoteDoc.Lines.CostingCode5
                        'objMainDoc.Lines.COGSCostingCode = objremoteDoc.Lines.COGSCostingCode
                        'objMainDoc.Lines.COGSCostingCode2 = objremoteDoc.Lines.COGSCostingCode2
                        'objMainDoc.Lines.COGSCostingCode3 = objremoteDoc.Lines.COGSCostingCode3
                        'objMainDoc.Lines.COGSCostingCode4 = objremoteDoc.Lines.COGSCostingCode4
                        'objMainDoc.Lines.COGSCostingCode5 = objremoteDoc.Lines.COGSCostingCode5
                        'objMainDoc.Lines.CommisionPercent = objremoteDoc.Lines.CommisionPercent
                        'objMainDoc.Lines.ConsumerSalesForecast = objremoteDoc.Lines.ConsumerSalesForecast
                        'objMainDoc.Lines.CorrectionInvoiceItem = objremoteDoc.Lines.CorrectionInvoiceItem
                        'objMainDoc.Lines.CorrInvAmountToDiffAcct = objremoteDoc.Lines.CorrInvAmountToDiffAcct


                        strBranchWhs = objremoteDoc.Lines.WarehouseCode
                        objMainrec.DoQuery("Select isnull(WhsCode,'" & strBranchWhs & "') from OWHS where U_Warcode='" & strBranchWhs & "'")
                        strMainWhs = objMainrec.Fields.Item(0).Value


                        objMainDoc.Lines.WarehouseCode = strMainWhs ' objremoteDoc.Lines.WarehouseCode

                        For intSer As Integer = 0 To objremoteDoc.Lines.SerialNumbers.Count - 1
                            If intSer > 0 Then
                                objMainDoc.Lines.SerialNumbers.Add()
                                objMainDoc.Lines.SerialNumbers.SetCurrentLine(intSer)
                            End If
                            objremoteDoc.Lines.SerialNumbers.SetCurrentLine(intSer)
                            objMainDoc.Lines.SerialNumbers.BaseLineNumber = objremoteDoc.Lines.SerialNumbers.BaseLineNumber
                            '  objMainDoc.Lines.SerialNumbers.BatchID = objremoteDoc.Lines.SerialNumbers.BatchID
                            objMainDoc.Lines.SerialNumbers.ExpiryDate = objremoteDoc.Lines.SerialNumbers.ExpiryDate
                            objMainDoc.Lines.SerialNumbers.InternalSerialNumber = objremoteDoc.Lines.SerialNumbers.InternalSerialNumber
                            ' objMainDoc.Lines.SerialNumbers.Location = objremoteDoc.Lines.SerialNumbers.Location
                            objMainDoc.Lines.SerialNumbers.ManufactureDate = objremoteDoc.Lines.SerialNumbers.ManufactureDate
                            objMainDoc.Lines.SerialNumbers.ManufacturerSerialNumber = objremoteDoc.Lines.SerialNumbers.ManufacturerSerialNumber
                            objMainDoc.Lines.SerialNumbers.Notes = objremoteDoc.Lines.SerialNumbers.Notes
                            objMainDoc.Lines.SerialNumbers.ReceptionDate = objremoteDoc.Lines.SerialNumbers.ReceptionDate
                            ' objMainDoc.Lines.SerialNumbers.SystemSerialNumber = objremoteDoc.Lines.SerialNumbers.SystemSerialNumber
                        Next

                        For intSer As Integer = 0 To objremoteDoc.Lines.BatchNumbers.Count - 1
                            If intSer > 0 Then
                                objMainDoc.Lines.BatchNumbers.Add()
                                objMainDoc.Lines.BatchNumbers.SetCurrentLine(intSer)
                            End If
                            objremoteDoc.Lines.BatchNumbers.SetCurrentLine(intSer)
                            objMainDoc.Lines.BatchNumbers.AddmisionDate = objremoteDoc.Lines.BatchNumbers.AddmisionDate
                            objMainDoc.Lines.BatchNumbers.BaseLineNumber = objremoteDoc.Lines.BatchNumbers.BaseLineNumber
                            objMainDoc.Lines.BatchNumbers.BatchNumber = objremoteDoc.Lines.BatchNumbers.BatchNumber
                            objMainDoc.Lines.BatchNumbers.ExpiryDate = objremoteDoc.Lines.BatchNumbers.ExpiryDate
                            objMainDoc.Lines.BatchNumbers.InternalSerialNumber = objremoteDoc.Lines.BatchNumbers.InternalSerialNumber
                            objMainDoc.Lines.BatchNumbers.Location = objremoteDoc.Lines.BatchNumbers.Location
                            objMainDoc.Lines.BatchNumbers.ManufacturingDate = objremoteDoc.Lines.BatchNumbers.ManufacturingDate
                            objMainDoc.Lines.BatchNumbers.Notes = objremoteDoc.Lines.BatchNumbers.Notes
                            objMainDoc.Lines.BatchNumbers.Quantity = objremoteDoc.Lines.BatchNumbers.Quantity
                        Next
                    Next

                    If objMainDoc.Add <> 0 Then
                        sPath = strErrorLogPath
                        WriteErrorlog("Failed to create Credit Notes docuemnt :" & objMainCompany.GetLastErrorDescription, sPath)
                    Else
                        Dim strDocNum As String
                        objMainCompany.GetNewObjectCode(strDocNum)
                        'MsgBox("DOcument created :" & strDocNum)
                        sPath = strErrorLogPath
                        WriteErrorlog("Credit Notes Created Successfully. DocNum : " & strDocNum, sPath)
                        objremoteDoc.UserFields.Fields.Item("U_Export").Value = "Y"
                        objremoteDoc.Update()
                    End If
                End If
                objremoteRec.MoveNext()
            Next

        Catch ex As Exception
            If Not IsNothing(objremoteDoc) Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objremoteDoc)
                GC.Collect()
            End If
            If Not IsNothing(objMainDoc) Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objMainDoc)
                GC.Collect()
            End If
        End Try
    End Sub
#End Region

    Private Sub ExportProduction(ByVal aChoice As String)
        Dim conn As New SqlConnection()
        Dim objMainDoc, objremoteDoc As SAPbobsCOM.ProductionOrders
        Dim intRemoteDocEntry As Integer
        Dim strPath, strFilename, strMessage, strTableName, stMessagetext As String
        Dim strFileName1 As String
        Dim objremoteRec As SAPbobsCOM.Recordset
        strTableName = ""
        Select Case aChoice
            Case "PO"
                strTableName = "OWOR"
                stMessagetext = "Production Order Document"
                objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
        End Select
        If strTableName = "" Then
            Exit Sub
        End If
        objremoteRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objremoteRec.DoQuery("Select DocEntry from " & strTableName & " where isnull(U_Export,'N')='N'")
        For intRemoteLoop As Integer = 0 To objremoteRec.RecordCount - 1
            If objremoteDoc.GetByKey(Convert.ToInt32(objremoteRec.Fields.Item(0).Value)) Then
                If checkDuplicateInvoice(objremoteDoc.AbsoluteEntry, objremoteDoc.DocumentNumber, objMainCompany.CompanyName, strTableName) = True Then
                    intRemoteDocEntry = objremoteDoc.AbsoluteEntry
                    Select Case aChoice
                        Case "PO"
                            strTableName = "owor"
                            'objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                            objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                    End Select
                    objMainDoc.GetByKey(getbaseDocument(objremoteDoc.DocumentNumber, objremoteDoc.DocumentNumber, 100, objMainCompany.CompanyName))
                    objMainDoc.ClosingDate = objremoteDoc.ClosingDate
                    objMainDoc.CustomerCode = objremoteDoc.CustomerCode
                    objMainDoc.DistributionRule = objremoteDoc.DistributionRule
                    objMainDoc.DistributionRule2 = objremoteDoc.DistributionRule2
                    objMainDoc.DistributionRule3 = objremoteDoc.DistributionRule3
                    objMainDoc.DistributionRule4 = objremoteDoc.DistributionRule4
                    objMainDoc.DistributionRule5 = objremoteDoc.DistributionRule5
                    objMainDoc.DueDate = objremoteDoc.DueDate
                    objMainDoc.ItemNo = objremoteDoc.ItemNo
                    objMainDoc.PlannedQuantity = objremoteDoc.PlannedQuantity
                    objMainDoc.PostingDate = objremoteDoc.PostingDate
                    If objremoteDoc.ProductionOrderOrigin <> 0 Then
                        objMainDoc.ProductionOrderOrigin = objremoteDoc.ProductionOrderOrigin
                        objMainDoc.ProductionOrderOriginEntry = objremoteDoc.ProductionOrderOriginEntry
                    End If
                    objMainDoc.ProductionOrderType = objremoteDoc.ProductionOrderType
                    objMainDoc.ProductionOrderStatus = objremoteDoc.ProductionOrderStatus

                    objMainDoc.Project = objremoteDoc.Project
                    objMainDoc.Remarks = objremoteDoc.Remarks


                    objMainDoc.UserFields.Fields.Item("U_Import").Value = "Y"
                    objMainDoc.UserFields.Fields.Item("U_BaseEntry").Value = objremoteDoc.AbsoluteEntry
                    objMainDoc.UserFields.Fields.Item("U_BaseNum").Value = objremoteDoc.DocumentNumber
                    objMainDoc.UserFields.Fields.Item("U_Branch").Value = objMainCompany.CompanyName

                    If 1 = 1 Then
                        For intLoop As Integer = 0 To objremoteDoc.Lines.Count - 1
                            objMainDoc.Lines.SetCurrentLine(intLoop)
                            objremoteDoc.Lines.SetCurrentLine(intLoop)
                            objMainDoc.Lines.BaseQuantity = objremoteDoc.Lines.BaseQuantity
                            objMainDoc.Lines.DistributionRule = objremoteDoc.Lines.DistributionRule
                            objMainDoc.Lines.DistributionRule2 = objremoteDoc.Lines.DistributionRule2
                            objMainDoc.Lines.DistributionRule3 = objremoteDoc.Lines.DistributionRule3
                            objMainDoc.Lines.DistributionRule4 = objremoteDoc.Lines.DistributionRule4
                            objMainDoc.Lines.DistributionRule5 = objremoteDoc.Lines.DistributionRule5
                            objMainDoc.Lines.ItemNo = objremoteDoc.Lines.ItemNo
                            objMainDoc.Lines.LocationCode = objremoteDoc.Lines.LocationCode
                            objMainDoc.Lines.PlannedQuantity = objremoteDoc.Lines.PlannedQuantity
                            objMainDoc.Lines.ProductionOrderIssueType = objremoteDoc.Lines.ProductionOrderIssueType
                            objMainDoc.Lines.Project = objremoteDoc.Lines.Project
                            objMainDoc.Lines.Warehouse = objremoteDoc.Lines.Warehouse
                        Next
                    Else

                    End If

                    If objMainDoc.Update <> 0 Then
                        sPath = strErrorLogPath
                        WriteErrorlog("Failed to Export " & stMessagetext & " docuemnt :" & objremoteDoc.DocumentNumber & ": Error :" & objRemoteCompany.GetLastErrorDescription & "  : Branch : " & objRemoteCompany.CompanyName, sPath)
                    Else
                        Dim strDocNum As String
                        objRemoteCompany.GetNewObjectCode(strDocNum)
                        Dim oInvoice As SAPbobsCOM.ProductionOrders
                        Dim intInvoiceDocentry As String
                        oInvoice = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                        If oInvoice.GetByKey(Convert.ToInt64(strDocNum)) Then
                            strDocNum = oInvoice.DocumentNumber
                            intInvoiceDocentry = oInvoice.AbsoluteEntry
                        End If
                        objMainDoc.GetByKey(intInvoiceDocentry)
                        objMainDoc.ProductionOrderStatus = objremoteDoc.ProductionOrderStatus
                        objMainDoc.Update()
                        sPath = strErrorLogPath
                        WriteErrorlog(stMessagetext & " Updated Successfully. DocNum : " & strDocNum, sPath)
                        Dim otempRec As SAPbobsCOM.Recordset
                        otempRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        objremoteDoc.GetByKey(intRemoteDocEntry)
                        objremoteDoc.UserFields.Fields.Item("U_Export").Value = "Y"
                        objremoteDoc.Update()
                    End If
                Else
                    intRemoteDocEntry = objremoteDoc.AbsoluteEntry
                    Select Case aChoice
                        Case "PO"
                            strTableName = "owor"
                            'objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                            objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                    End Select
                    objMainDoc.ClosingDate = objremoteDoc.ClosingDate
                    objMainDoc.CustomerCode = objremoteDoc.CustomerCode
                    objMainDoc.DistributionRule = objremoteDoc.DistributionRule
                    objMainDoc.DistributionRule2 = objremoteDoc.DistributionRule2
                    objMainDoc.DistributionRule3 = objremoteDoc.DistributionRule3
                    objMainDoc.DistributionRule4 = objremoteDoc.DistributionRule4
                    objMainDoc.DistributionRule5 = objremoteDoc.DistributionRule5
                    objMainDoc.DueDate = objremoteDoc.DueDate
                    objMainDoc.ItemNo = objremoteDoc.ItemNo
                    objMainDoc.PlannedQuantity = objremoteDoc.PlannedQuantity
                    objMainDoc.PostingDate = objremoteDoc.PostingDate
                    '  MsgBox(objremoteDoc.ProductionOrderOrigin)
                    If objremoteDoc.ProductionOrderOrigin <> 0 Then




                        objMainDoc.ProductionOrderOrigin = objremoteDoc.ProductionOrderOrigin
                        objMainDoc.ProductionOrderOriginEntry = objremoteDoc.ProductionOrderOriginEntry
                    End If

                    objMainDoc.ProductionOrderType = objremoteDoc.ProductionOrderType
                    objMainDoc.Project = objremoteDoc.Project
                    objMainDoc.Remarks = objremoteDoc.Remarks


                    objMainDoc.UserFields.Fields.Item("U_Import").Value = "Y"
                    objMainDoc.UserFields.Fields.Item("U_BaseEntry").Value = objremoteDoc.AbsoluteEntry
                    objMainDoc.UserFields.Fields.Item("U_BaseNum").Value = objremoteDoc.DocumentNumber
                    objMainDoc.UserFields.Fields.Item("U_Branch").Value = objMainCompany.CompanyName

                    If 1 = 1 Then
                        For intLoop As Integer = 0 To objremoteDoc.Lines.Count - 1
                            If intLoop > 0 Then
                                objMainDoc.Lines.Add()
                                objMainDoc.Lines.SetCurrentLine(intLoop)
                            End If
                            objremoteDoc.Lines.SetCurrentLine(intLoop)
                            objMainDoc.Lines.BaseQuantity = objremoteDoc.Lines.BaseQuantity
                            objMainDoc.Lines.DistributionRule = objremoteDoc.Lines.DistributionRule
                            objMainDoc.Lines.DistributionRule2 = objremoteDoc.Lines.DistributionRule2
                            objMainDoc.Lines.DistributionRule3 = objremoteDoc.Lines.DistributionRule3
                            objMainDoc.Lines.DistributionRule4 = objremoteDoc.Lines.DistributionRule4
                            objMainDoc.Lines.DistributionRule5 = objremoteDoc.Lines.DistributionRule5
                            objMainDoc.Lines.ItemNo = objremoteDoc.Lines.ItemNo
                            objMainDoc.Lines.LocationCode = objremoteDoc.Lines.LocationCode
                            objMainDoc.Lines.PlannedQuantity = objremoteDoc.Lines.PlannedQuantity
                            objMainDoc.Lines.ProductionOrderIssueType = objremoteDoc.Lines.ProductionOrderIssueType
                            objMainDoc.Lines.Project = objremoteDoc.Lines.Project
                            objMainDoc.Lines.Warehouse = objremoteDoc.Lines.Warehouse

                        Next
                    Else

                    End If
                    If objMainDoc.Add <> 0 Then
                        sPath = strErrorLogPath
                        WriteErrorlog("Failed to Export " & stMessagetext & " docuemnt :" & objremoteDoc.DocumentNumber & ": Error :" & objRemoteCompany.GetLastErrorDescription & "  : Branch : " & objRemoteCompany.CompanyName, sPath)
                    Else
                        Dim strDocNum As String
                        objRemoteCompany.GetNewObjectCode(strDocNum)
                        Dim oInvoice As SAPbobsCOM.ProductionOrders
                        Dim intInvoiceDocentry As String
                        oInvoice = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders)
                        If oInvoice.GetByKey(Convert.ToInt64(strDocNum)) Then
                            strDocNum = oInvoice.DocumentNumber
                            intInvoiceDocentry = oInvoice.AbsoluteEntry
                        End If
                        objMainDoc.GetByKey(intInvoiceDocentry)
                        objMainDoc.ProductionOrderStatus = objremoteDoc.ProductionOrderStatus
                        objMainDoc.Update()

                        sPath = strErrorLogPath
                        WriteErrorlog(stMessagetext & " Created Successfully. DocNum : " & strDocNum, sPath)
                        Dim otempRec As SAPbobsCOM.Recordset
                        otempRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        objremoteDoc.GetByKey(intRemoteDocEntry)
                        objremoteDoc.UserFields.Fields.Item("U_Export").Value = "Y"
                        objremoteDoc.Update()
                    End If


                End If
            End If
            objremoteRec.MoveNext()
        Next
    End Sub



    Private Sub ExportInventoryDocuments(ByVal aChoice As String)
        Dim conn As New SqlConnection()
        Dim objMainDoc, objremoteDoc As SAPbobsCOM.Documents
        Dim intRemoteDocEntry As Integer
        Dim strPath, strFilename, strMessage, strTableName As String
        Dim strFileName1 As String
        Dim objremoteRec As SAPbobsCOM.Recordset
        Try
            strTableName = ""
            Select Case aChoice
                Case "GoodsReceipt"
                    strTableName = "OIGN"
                    objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)
                    objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)
                Case "GoodsIssue"
                    strTableName = "OIGE"
                    objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
                    objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)

            End Select
            If strTableName = "" Then
                Exit Sub
            End If
            objremoteRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objremoteRec.DoQuery("Select DocEntry from " & strTableName & " where  isnull(U_Export,'N')='N'")
            For intRemoteLoop As Integer = 0 To objremoteRec.RecordCount - 1
                If objremoteDoc.GetByKey(Convert.ToInt32(objremoteRec.Fields.Item(0).Value)) Then
                    If checkDuplicateInvoice(objremoteDoc.DocEntry, objremoteDoc.DocNum, objMainCompany.CompanyName, strTableName) = True Then
                        'objremoteDoc.GetByKey(intRemoteDocEntry)
                        objremoteDoc.UserFields.Fields.Item("U_Export").Value = "Y"
                        objremoteDoc.Update()
                    Else
                        intRemoteDocEntry = objremoteDoc.DocEntry
                        Select Case aChoice
                            Case "GoodsReceipt"
                                strTableName = "OIGN"
                                'objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)
                                objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)
                            Case "GoodsIssue"
                                strTableName = "OIGE"
                                '  objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
                                objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
                        End Select
                        objMainDoc.CardCode = objremoteDoc.CardCode
                        objMainDoc.DocDate = objremoteDoc.DocDate
                        objMainDoc.DocDueDate = objremoteDoc.DocDueDate
                        objMainDoc.NumAtCard = objremoteDoc.NumAtCard
                        objMainDoc.Comments = objremoteDoc.Comments
                        objMainDoc.DiscountPercent = objremoteDoc.DiscountPercent
                        objMainDoc.DocCurrency = objremoteDoc.DocCurrency
                        objMainDoc.ShipToCode = objremoteDoc.ShipToCode
                        objMainDoc.SalesPersonCode = objremoteDoc.SalesPersonCode
                        'objMainDoc.TaxDate = objremoteDoc.TaxDate
                        'objMainDoc.PaymentGroupCode = objremoteDoc.PaymentGroupCode
                        objMainDoc.PaymentMethod = objremoteDoc.PaymentMethod
                        objMainDoc.Address = objremoteDoc.Address
                        objMainDoc.Address2 = objremoteDoc.Address2
                        objMainDoc.AgentCode = objremoteDoc.AgentCode
                        objMainDoc.BPChannelCode = objremoteDoc.BPChannelCode
                        objMainDoc.BPChannelContact = objremoteDoc.BPChannelContact
                        objMainDoc.ContactPersonCode = objremoteDoc.ContactPersonCode
                        ' objMainDoc.DocRate = objremoteDoc.DocRate
                        objMainDoc.DocType = objremoteDoc.DocType
                        If objremoteDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES Then
                            objMainDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES
                            objMainDoc.RoundingDiffAmount = objremoteDoc.RoundingDiffAmount
                        Else
                            objMainDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tNO
                        End If
                        objMainDoc.UserFields.Fields.Item("U_Import").Value = "Y"
                        objMainDoc.UserFields.Fields.Item("U_BaseEntry").Value = objremoteDoc.DocEntry
                        objMainDoc.UserFields.Fields.Item("U_BaseNum").Value = objremoteDoc.DocNum
                        objMainDoc.UserFields.Fields.Item("U_Branch").Value = objMainCompany.CompanyName

                        If aChoice <> "PQ" Then
                            For intLoop As Integer = 0 To objremoteDoc.Lines.Count - 1
                                If intLoop > 0 Then
                                    objMainDoc.Lines.Add()
                                    objMainDoc.Lines.SetCurrentLine(intLoop)
                                End If
                                objremoteDoc.Lines.SetCurrentLine(intLoop)
                                If objremoteDoc.Lines.BaseType <> -1 Then
                                    objMainDoc.Lines.BaseType = objremoteDoc.Lines.BaseType
                                    objMainDoc.Lines.BaseEntry = getbaseDocument(objremoteDoc.Lines.BaseEntry, objremoteDoc.Lines.BaseEntry, objremoteDoc.Lines.BaseType, objMainCompany.CompanyName) 'objremoteDoc.Lines.BaseEntry
                                    If aChoice <> "GoodsReceipt" Then
                                        objMainDoc.Lines.BaseLine = objremoteDoc.Lines.BaseLine
                                    End If

                                    objMainDoc.Lines.Quantity = objremoteDoc.Lines.Quantity
                                    objMainDoc.Lines.WarehouseCode = objremoteDoc.Lines.WarehouseCode
                                Else
                                    objMainDoc.Lines.AccountCode = objremoteDoc.Lines.AccountCode
                                    objMainDoc.Lines.ItemDescription = objremoteDoc.Lines.ItemDescription
                                    objMainDoc.Lines.ItemCode = objremoteDoc.Lines.ItemCode
                                    objMainDoc.Lines.BarCode = objremoteDoc.Lines.BarCode
                                    ' objMainDoc.Lines.UnitPrice = objremoteDoc.Lines.UnitPrice
                                    objMainDoc.Lines.DiscountPercent = objremoteDoc.Lines.DiscountPercent
                                    objMainDoc.Lines.VatGroup = objremoteDoc.Lines.VatGroup
                                    objMainDoc.Lines.PriceAfterVAT = objremoteDoc.Lines.PriceAfterVAT
                                    objMainDoc.Lines.LineTotal = objremoteDoc.Lines.LineTotal
                                    objMainDoc.Lines.ProjectCode = objremoteDoc.Lines.ProjectCode
                                    objMainDoc.Lines.Quantity = objremoteDoc.Lines.Quantity
                                    objMainDoc.Lines.WarehouseCode = objremoteDoc.Lines.WarehouseCode
                                    objMainDoc.Lines.TaxCode = objremoteDoc.Lines.TaxCode
                                End If
                                If aChoice <> "PQ" Then
                                    For intSer As Integer = 0 To objremoteDoc.Lines.SerialNumbers.Count - 1
                                        If intSer > 0 Then
                                            objMainDoc.Lines.SerialNumbers.Add()
                                            objMainDoc.Lines.SerialNumbers.SetCurrentLine(intSer)
                                        End If
                                        objremoteDoc.Lines.SerialNumbers.SetCurrentLine(intSer)
                                        objMainDoc.Lines.SerialNumbers.BaseLineNumber = objremoteDoc.Lines.SerialNumbers.BaseLineNumber
                                        objMainDoc.Lines.SerialNumbers.ExpiryDate = objremoteDoc.Lines.SerialNumbers.ExpiryDate
                                        objMainDoc.Lines.SerialNumbers.InternalSerialNumber = objremoteDoc.Lines.SerialNumbers.InternalSerialNumber
                                        objMainDoc.Lines.SerialNumbers.ManufactureDate = objremoteDoc.Lines.SerialNumbers.ManufactureDate
                                        objMainDoc.Lines.SerialNumbers.ManufacturerSerialNumber = objremoteDoc.Lines.SerialNumbers.ManufacturerSerialNumber
                                        objMainDoc.Lines.SerialNumbers.Notes = objremoteDoc.Lines.SerialNumbers.Notes
                                        objMainDoc.Lines.SerialNumbers.ReceptionDate = objremoteDoc.Lines.SerialNumbers.ReceptionDate
                                    Next
                                End If

                                If aChoice <> "PQ" Then
                                    For intSer As Integer = 0 To objremoteDoc.Lines.BatchNumbers.Count - 1
                                        If intSer > 0 Then
                                            objMainDoc.Lines.BatchNumbers.Add()
                                            objMainDoc.Lines.BatchNumbers.SetCurrentLine(intSer)
                                        End If
                                        objremoteDoc.Lines.BatchNumbers.SetCurrentLine(intSer)
                                        objMainDoc.Lines.BatchNumbers.AddmisionDate = objremoteDoc.Lines.BatchNumbers.AddmisionDate
                                        objMainDoc.Lines.BatchNumbers.BaseLineNumber = objremoteDoc.Lines.BatchNumbers.BaseLineNumber
                                        objMainDoc.Lines.BatchNumbers.BatchNumber = objremoteDoc.Lines.BatchNumbers.BatchNumber
                                        objMainDoc.Lines.BatchNumbers.ExpiryDate = objremoteDoc.Lines.BatchNumbers.ExpiryDate
                                        objMainDoc.Lines.BatchNumbers.InternalSerialNumber = objremoteDoc.Lines.BatchNumbers.InternalSerialNumber
                                        objMainDoc.Lines.BatchNumbers.Location = objremoteDoc.Lines.BatchNumbers.Location
                                        objMainDoc.Lines.BatchNumbers.ManufacturingDate = objremoteDoc.Lines.BatchNumbers.ManufacturingDate
                                        objMainDoc.Lines.BatchNumbers.Notes = objremoteDoc.Lines.BatchNumbers.Notes
                                        objMainDoc.Lines.BatchNumbers.Quantity = objremoteDoc.Lines.BatchNumbers.Quantity
                                    Next
                                End If
                            Next
                        Else
                        End If


                        If objMainDoc.Add <> 0 Then
                            sPath = strErrorLogPath
                            WriteErrorlog("Failed to Export " & aChoice & " docuemnt :" & objremoteDoc.DocNum & ": Error :" & objRemoteCompany.GetLastErrorDescription & "  : Branch : " & objRemoteCompany.CompanyName, sPath)
                        Else
                            Dim strDocNum As String
                            objRemoteCompany.GetNewObjectCode(strDocNum)
                            Dim oInvoice As SAPbobsCOM.Documents
                            Dim intInvoiceDocentry As String
                            oInvoice = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                            If oInvoice.GetByKey(Convert.ToInt64(strDocNum)) Then
                                strDocNum = oInvoice.DocNum
                                intInvoiceDocentry = oInvoice.DocEntry
                            End If
                            sPath = strErrorLogPath
                            WriteErrorlog(aChoice & " Created Successfully. DocNum : " & strDocNum, sPath)
                            Dim otempRec As SAPbobsCOM.Recordset
                            otempRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            objremoteDoc.GetByKey(intRemoteDocEntry)
                            objremoteDoc.UserFields.Fields.Item("U_Export").Value = "Y"
                            objremoteDoc.Update()
                        End If
                    End If
                End If
                objremoteRec.MoveNext()
            Next
        Catch ex As Exception
            WriteErrorlog(ex.Message, sPath)
        End Try
    End Sub




    Private Sub ExportStockTransfer()
        Dim conn As New SqlConnection()
        Dim objMainDoc, objremoteDoc As SAPbobsCOM.StockTransfer
        Dim intRemoteDocEntry As Integer
        Dim strPath, strFilename, strMessage, strTableName, aChoice As String
        Dim strFileName1 As String
        Dim objremoteRec As SAPbobsCOM.Recordset
        strTableName = ""
        strTableName = "OWTR"
        objremoteDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
        objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
        aChoice = "Stock Transfer"
        If strTableName = "" Then
            Exit Sub
        End If
        Try


            objremoteRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            objremoteRec.DoQuery("Select DocEntry from " & strTableName & " where  isnull(U_Export,'N')='N'")
            For intRemoteLoop As Integer = 0 To objremoteRec.RecordCount - 1
                If objremoteDoc.GetByKey(Convert.ToInt32(objremoteRec.Fields.Item(0).Value)) Then
                    If checkDuplicateInvoice(objremoteDoc.DocEntry, objremoteDoc.DocNum, objMainCompany.CompanyName, strTableName) = True Then
                        'objremoteDoc.GetByKey(intRemoteDocEntry)
                        objremoteDoc.UserFields.Fields.Item("U_Export").Value = "Y"
                        objremoteDoc.Update()
                    Else
                        intRemoteDocEntry = objremoteDoc.DocEntry
                        objMainDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)

                        objMainDoc.CardCode = objremoteDoc.CardCode
                        objMainDoc.Address = objremoteDoc.Address
                        objMainDoc.CardName = objremoteDoc.CardName
                        objMainDoc.Comments = objremoteDoc.Comments
                        objMainDoc.ContactPerson = objremoteDoc.ContactPerson
                        objMainDoc.DocDate = objremoteDoc.DocDate
                        objMainDoc.FromWarehouse = objremoteDoc.FromWarehouse
                        objMainDoc.PriceList = objremoteDoc.PriceList
                        objMainDoc.Reference1 = objremoteDoc.Reference1
                        objMainDoc.Reference2 = objremoteDoc.Reference2
                        objMainDoc.SalesPersonCode = objremoteDoc.SalesPersonCode
                        objMainDoc.TaxDate = objremoteDoc.TaxDate

                        objMainDoc.UserFields.Fields.Item("U_Import").Value = "Y"
                        objMainDoc.UserFields.Fields.Item("U_BaseEntry").Value = objremoteDoc.DocEntry
                        objMainDoc.UserFields.Fields.Item("U_BaseNum").Value = objremoteDoc.DocNum
                        objMainDoc.UserFields.Fields.Item("U_Branch").Value = objMainCompany.CompanyName

                        If aChoice <> "PQ" Then
                            For intLoop As Integer = 0 To objremoteDoc.Lines.Count - 1
                                If intLoop > 0 Then
                                    objMainDoc.Lines.Add()
                                    objMainDoc.Lines.SetCurrentLine(intLoop)
                                End If
                                objremoteDoc.Lines.SetCurrentLine(intLoop)
                                objMainDoc.Lines.Currency = objremoteDoc.Lines.Currency
                                objMainDoc.Lines.DiscountPercent = objremoteDoc.Lines.DiscountPercent
                                objMainDoc.Lines.DistributionRule = objremoteDoc.Lines.DistributionRule
                                objMainDoc.Lines.DistributionRule2 = objremoteDoc.Lines.DistributionRule2
                                objMainDoc.Lines.DistributionRule3 = objremoteDoc.Lines.DistributionRule3
                                objMainDoc.Lines.DistributionRule4 = objremoteDoc.Lines.DistributionRule4
                                objMainDoc.Lines.DistributionRule5 = objremoteDoc.Lines.DistributionRule5

                                objMainDoc.Lines.Factor = objremoteDoc.Lines.Factor
                                objMainDoc.Lines.Factor2 = objremoteDoc.Lines.Factor2
                                objMainDoc.Lines.Factor3 = objremoteDoc.Lines.Factor3
                                objMainDoc.Lines.Factor4 = objremoteDoc.Lines.Factor4
                                objMainDoc.Lines.ItemCode = objremoteDoc.Lines.ItemCode
                                objMainDoc.Lines.ItemDescription = objremoteDoc.Lines.ItemDescription
                                objMainDoc.Lines.Price = objremoteDoc.Lines.Price
                                objMainDoc.Lines.MeasureUnit = objremoteDoc.Lines.MeasureUnit
                                objMainDoc.Lines.ProjectCode = objremoteDoc.Lines.ProjectCode
                                objMainDoc.Lines.Quantity = objremoteDoc.Lines.Quantity
                                objMainDoc.Lines.WarehouseCode = objremoteDoc.Lines.WarehouseCode
                                'objMainDoc.Lines.TaxCode = objremoteDoc.Lines.TaxCode
                                If aChoice <> "PQ" Then
                                    For intSer As Integer = 0 To objremoteDoc.Lines.SerialNumbers.Count - 1
                                        If intSer > 0 Then
                                            objMainDoc.Lines.SerialNumbers.Add()
                                            objMainDoc.Lines.SerialNumbers.SetCurrentLine(intSer)
                                        End If
                                        objremoteDoc.Lines.SerialNumbers.SetCurrentLine(intSer)
                                        objMainDoc.Lines.SerialNumbers.BaseLineNumber = objremoteDoc.Lines.SerialNumbers.BaseLineNumber
                                        '  objMainDoc.Lines.SerialNumbers.BatchID = objremoteDoc.Lines.SerialNumbers.BatchID
                                        objMainDoc.Lines.SerialNumbers.ExpiryDate = objremoteDoc.Lines.SerialNumbers.ExpiryDate
                                        objMainDoc.Lines.SerialNumbers.InternalSerialNumber = objremoteDoc.Lines.SerialNumbers.InternalSerialNumber
                                        ' objMainDoc.Lines.SerialNumbers.Location = objremoteDoc.Lines.SerialNumbers.Location
                                        objMainDoc.Lines.SerialNumbers.ManufactureDate = objremoteDoc.Lines.SerialNumbers.ManufactureDate
                                        objMainDoc.Lines.SerialNumbers.ManufacturerSerialNumber = objremoteDoc.Lines.SerialNumbers.ManufacturerSerialNumber
                                        objMainDoc.Lines.SerialNumbers.Notes = objremoteDoc.Lines.SerialNumbers.Notes
                                        objMainDoc.Lines.SerialNumbers.ReceptionDate = objremoteDoc.Lines.SerialNumbers.ReceptionDate
                                        ' objMainDoc.Lines.SerialNumbers.SystemSerialNumber = objremoteDoc.Lines.SerialNumbers.SystemSerialNumber
                                    Next
                                End If

                                If aChoice <> "PQ" Then
                                    For intSer As Integer = 0 To objremoteDoc.Lines.BatchNumbers.Count - 1
                                        If intSer > 0 Then
                                            objMainDoc.Lines.BatchNumbers.Add()
                                            objMainDoc.Lines.BatchNumbers.SetCurrentLine(intSer)
                                        End If
                                        objremoteDoc.Lines.BatchNumbers.SetCurrentLine(intSer)
                                        objMainDoc.Lines.BatchNumbers.AddmisionDate = objremoteDoc.Lines.BatchNumbers.AddmisionDate
                                        objMainDoc.Lines.BatchNumbers.BaseLineNumber = objremoteDoc.Lines.BatchNumbers.BaseLineNumber
                                        objMainDoc.Lines.BatchNumbers.BatchNumber = objremoteDoc.Lines.BatchNumbers.BatchNumber
                                        objMainDoc.Lines.BatchNumbers.ExpiryDate = objremoteDoc.Lines.BatchNumbers.ExpiryDate
                                        objMainDoc.Lines.BatchNumbers.InternalSerialNumber = objremoteDoc.Lines.BatchNumbers.InternalSerialNumber
                                        objMainDoc.Lines.BatchNumbers.Location = objremoteDoc.Lines.BatchNumbers.Location
                                        objMainDoc.Lines.BatchNumbers.ManufacturingDate = objremoteDoc.Lines.BatchNumbers.ManufacturingDate
                                        objMainDoc.Lines.BatchNumbers.Notes = objremoteDoc.Lines.BatchNumbers.Notes
                                        objMainDoc.Lines.BatchNumbers.Quantity = objremoteDoc.Lines.BatchNumbers.Quantity
                                    Next
                                End If
                            Next
                        Else

                        End If
                        If objMainDoc.Add <> 0 Then
                            sPath = strErrorLogPath
                            WriteErrorlog("Failed to Export Stock Transfer docuemnt :" & objremoteDoc.DocNum & ": Error :" & objRemoteCompany.GetLastErrorDescription & "  : Branch : " & objRemoteCompany.CompanyName, sPath)
                        Else
                            Dim strDocNum As String
                            objRemoteCompany.GetNewObjectCode(strDocNum)
                            Dim oInvoice As SAPbobsCOM.Documents
                            Dim intInvoiceDocentry As String
                            oInvoice = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
                            If oInvoice.GetByKey(Convert.ToInt64(strDocNum)) Then
                                strDocNum = oInvoice.DocNum
                                intInvoiceDocentry = oInvoice.DocEntry
                            End If
                            sPath = strErrorLogPath
                            WriteErrorlog("Stock Transfer Created Successfully. DocNum : " & strDocNum, sPath)
                            Dim otempRec As SAPbobsCOM.Recordset
                            otempRec = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                            objremoteDoc.GetByKey(intRemoteDocEntry)
                            objremoteDoc.UserFields.Fields.Item("U_Export").Value = "Y"
                            objremoteDoc.Update()
                        End If
                    End If
                End If
                objremoteRec.MoveNext()
            Next
        Catch ex As Exception
            WriteErrorlog(ex.Message, sPath)
        End Try
    End Sub

#Region "Get BaseEntry"

    Private Function GetBaseDocEntry_LandedCost(ByVal aDocEntry As Integer, ByVal strCHoice As String, ByVal strBranch As String) As Integer
        Dim oTempMainRs, oTempRemoteRs As SAPbobsCOM.Recordset
        Dim strQuery, strQuery1, aDocNumber As String
        oTempMainRs = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRemoteRs = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        If strCHoice = "GRPO" Then
            strQuery1 = "Select * from OPDN where DocEntry=" & aDocEntry
        Else
            strQuery1 = "Select * from ORIN where DocEntry=" & aDocEntry
        End If
        oTempRemoteRs.DoQuery(strQuery1)
        aDocNumber = oTempRemoteRs.Fields.Item("DocNum").Value

        If strCHoice = "GRPO" Then
            strQuery = "Select docEntry from OPDN where U_Branch='" & strBranch & "' and   U_BaseEntry=" & aDocEntry & " and U_BaseNum=" & aDocNumber
        Else
            strQuery = "Select docEntry from ORIN where U_Branch='" & strBranch & "'  and   U_BaseEntry=" & aDocEntry & " and U_BaseNum=" & aDocNumber
        End If
        oTempMainRs.DoQuery(strQuery)
        aDocNumber = oTempMainRs.Fields.Item(0).Value
        Return oTempMainRs.Fields.Item(0).Value
    End Function
    Private Function GetBaseDocEntry(ByVal aDocEntry As Integer, ByVal aDocNumber As Integer, ByVal strCHoice As String, ByVal strBranch As String) As Integer
        Dim oTempMainRs, oTempRemoteRs As SAPbobsCOM.Recordset
        Dim strQuery, strQuery1 As String
        oTempMainRs = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRemoteRs = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        If strCHoice = "Invoice" Then
            strQuery1 = "Select  * from OINV where DocEntry=" & aDocEntry
        ElseIf strCHoice = "JE" Then
            strQuery1 = "Select  * from OJDT where TransID=" & aDocEntry
        Else
            strQuery1 = "Select  * from ORIN where DocEntry=" & aDocEntry
        End If
        oTempRemoteRs.DoQuery(strQuery1)
        If strCHoice = "JE" Then
            aDocNumber = oTempRemoteRs.Fields.Item("Number").Value
        Else
            aDocNumber = oTempRemoteRs.Fields.Item("DocNum").Value
        End If


        If strCHoice = "Invoice" Then
            strQuery = "Select docEntry from OINV where U_Branch='" & strBranch & "' and   U_BaseEntry=" & aDocEntry & " and U_BaseNum=" & aDocNumber
        ElseIf strCHoice = "JE" Then
            strQuery = "Select TransID from OJDT where U_Branch='" & strBranch & "' and   U_BaseEntry=" & aDocEntry & " and U_BaseNum=" & aDocNumber
        Else
            strQuery = "Select docEntry from ORIN where U_Branch='" & strBranch & "'  and   U_BaseEntry=" & aDocEntry & " and U_BaseNum=" & aDocNumber
        End If
        oTempMainRs.DoQuery(strQuery)
        Return oTempMainRs.Fields.Item(0).Value
    End Function

    Private Function getbaseDocument_IncomingPayment(ByVal aDocEntry As Integer, ByVal aDocNum As Integer, ByVal aBasetype As Double, ByVal aBranch As String) As Integer
        Dim oTempMainRs, oTempRemoteRs As SAPbobsCOM.Recordset
        Dim strQuery, strQuery1, strTable As String
        oTempMainRs = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRemoteRs = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Select Case aBasetype
            Case 22
                strTable = "OPOR"
            Case 21
                strTable = "ORPD"
            Case 18
                strTable = "OPCH"
            Case 20
                strTable = "OPDN"
            Case 19
                strTable = "ORPC"
            Case 100
                strTable = "OWOR"
            Case 202
                strTable = "OWOR"
            Case 30 'Journal'
                strTable = "OJDT"
            Case 13 'AR Invoice
                strTable = "OINV"
            Case "14" 'AR Credit Note
                strTable = "ORIN"
            Case "16" 'Sales Return
                strTable = "ORDN"
            Case "15" 'Sales Return
                strTable = "ODLN"
            Case "24"
                strTable = "ORCT"
        End Select
        If aBasetype = "-1" Or strTable = "" Then
            Return 0
        End If
        If aBasetype = 24 Then
            strQuery = "Select BaseRef from OJDT where TransID=" & aDocEntry '& " and U_BaseNum=" & aDocEntry
            oTempMainRs.DoQuery(strQuery)
            If oTempMainRs.RecordCount > 0 Then
                strTable = "ORCT"
                aDocEntry = oTempMainRs.Fields.Item(0).Value
                strQuery = "Select TransId from OJDT where TransType=24 and   BaseRef=" & aDocEntry '& " and U_BaseNum=" & aDocEntry
                oTempRemoteRs.DoQuery(strQuery)
                If oTempRemoteRs.RecordCount > 0 Then
                    Return oTempRemoteRs.Fields.Item(0).Value
                Else
                    Return 0
                End If
            End If
        End If
        If strTable <> "OJDT" Then
            strQuery = "Select docEntry from " & strTable & " where U_Branch='" & aBranch & "' and   U_BaseEntry=" & aDocEntry '& " and U_BaseNum=" & aDocEntry
        Else
            strQuery = "Select TransId from " & strTable & " where U_Branch='" & aBranch & "' and   U_BaseEntry=" & aDocEntry '& " and U_BaseNum=" & aDocEntry
        End If

        oTempRemoteRs.DoQuery(strQuery)
        If oTempRemoteRs.RecordCount > 0 Then
            Return oTempRemoteRs.Fields.Item(0).Value
        Else
            Return 0
        End If


    End Function


#End Region

#Region "get base document"

    Private Function getbaseDocument_Lines(ByVal aDocEntry As Integer, ByVal aDocNum As Integer, ByVal aBasetype As Double, ByVal aBranch As String, ByVal aLine As Integer) As Integer
        Dim oTempMainRs, oTempRemoteRs As SAPbobsCOM.Recordset
        Dim strQuery, strQuery1, strTable As String
        oTempMainRs = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRemoteRs = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Select Case aBasetype
            Case 22
                strTable = "OPOR"
            Case 21
                strTable = "ORPD"
            Case 18
                strTable = "OPCH"
            Case 20
                strTable = "OPDN"
            Case 19
                strTable = "ORPC"
            Case 100
                strTable = "OWOR"
            Case 202
                strTable = "OWOR"
            Case 30
                strTable = "OJDT"
            Case 13
                strTable = "OINV"
            Case 46
                strTable = "OVPM"
        End Select
        If aBasetype = "-1" Or strTable = "" Then
            Return 0
        End If

        If aBasetype = 46 Then
            strQuery = "Select BaseRef from OJDT where TransID=" & aDocEntry '& " and U_BaseNum=" & aDocEntry
            oTempMainRs.DoQuery(strQuery)
            If oTempMainRs.RecordCount > 0 Then
                strTable = "OVPM"
                aDocEntry = oTempMainRs.Fields.Item(0).Value
                strQuery = "Select TransId from OJDT where TransType=46 and   BaseRef=" & aDocEntry '& " and U_BaseNum=" & aDocEntry
                oTempRemoteRs.DoQuery(strQuery)
                If oTempRemoteRs.RecordCount > 0 Then
                    Return oTempRemoteRs.Fields.Item(0).Value
                Else
                    Return 0
                End If
            End If
        End If

        If strTable <> "OJDT" Then
            strQuery = "Select docEntry from " & strTable & " where U_Branch='" & aBranch & "' and   U_BaseEntry=" & aDocEntry '& " and U_BaseNum=" & aDocEntry
        Else
            strQuery = "Select TransId from " & strTable & " where U_Branch='" & aBranch & "' and   U_BaseEntry=" & aDocEntry '& " and U_BaseNum=" & aDocEntry
        End If

        oTempRemoteRs.DoQuery(strQuery)
        If oTempRemoteRs.RecordCount > 0 Then
            Return oTempRemoteRs.Fields.Item(0).Value
        Else
            Return 0
        End If


    End Function

    Private Function getbaseDocument(ByVal aDocEntry As Integer, ByVal aDocNum As Integer, ByVal aBasetype As Double, ByVal aBranch As String) As Integer
        Dim oTempMainRs, oTempRemoteRs, oTempRemoteRS1 As SAPbobsCOM.Recordset
        Dim strQuery, strQuery1, strTable As String
        oTempMainRs = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRemoteRs = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRemoteRS1 = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim strRemoteDocNum As String
        Select Case aBasetype
            Case 22
                strTable = "OPOR"
            Case 21
                strTable = "ORPD"
            Case 18
                strTable = "OPCH"
            Case 20
                strTable = "OPDN"
            Case 19
                strTable = "ORPC"
            Case 100
                strTable = "OWOR"
            Case 202
                strTable = "OWOR"
            Case 30
                strTable = "OJDT"
            Case 13
                strTable = "OINV"
            Case 46
                strTable = "OVPM"
        End Select
        If aBasetype = "-1" Or strTable = "" Then
            Return 0
        End If



        If aBasetype = 46 Then
            strQuery = "Select BaseRef from OJDT where TransID=" & aDocEntry '& " and U_BaseNum=" & aDocEntry
            oTempMainRs.DoQuery(strQuery)
            If oTempMainRs.RecordCount > 0 Then
                strTable = "OVPM"
                aDocEntry = oTempMainRs.Fields.Item(0).Value
                strQuery = "Select TransId from OJDT where TransType=46 and   BaseRef=" & aDocEntry '& " and U_BaseNum=" & aDocEntry
                oTempRemoteRs.DoQuery(strQuery)
                If oTempRemoteRs.RecordCount > 0 Then
                    Return oTempRemoteRs.Fields.Item(0).Value
                Else
                    Return 0
                End If
            End If
        End If

        If strTable <> "OJDT" Then
            strQuery = "Select docEntry from " & strTable & " where U_Branch='" & aBranch & "' and   U_BaseEntry=" & aDocEntry '& " and U_BaseNum=" & aDocEntry
        Else
            strQuery = "Select TransId from " & strTable & " where U_Branch='" & aBranch & "' and   U_BaseEntry=" & aDocEntry '& " and U_BaseNum=" & aDocEntry
        End If
        oTempRemoteRs.DoQuery(strQuery)
        If oTempRemoteRs.RecordCount > 0 Then
            Return oTempRemoteRs.Fields.Item(0).Value
        Else
            Return 0
        End If


    End Function
#End Region
    Private Sub ImportCreditNotes1(ByVal aCompany As SAPbobsCOM.Company)
        Dim conn As New SqlConnection()
        Dim objMainDoc, objremoteDoc As SAPbobsCOM.Documents
        Dim strPath, strFilename, strMessage As String
        Dim strFileName1 As String
        Dim objremoteRec As SAPbobsCOM.Recordset
        objremoteRec = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        objremoteRec.DoQuery("Select DocEntry from ORIN where   isnull(U_Export,'N')='N'")
        For intRemoteLoop As Integer = 0 To objremoteRec.RecordCount - 1
            objremoteDoc = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
            If objremoteDoc.GetByKey(Convert.ToInt32(objremoteRec.Fields.Item(0).Value)) Then
                If checkDuplicateInvoice(objremoteDoc.DocEntry, objremoteDoc.DocNum, objRemoteCompany.CompanyName, "ORIN") = True Then
                    objremoteDoc.UserFields.Fields.Item("U_Export").Value = "Y"
                    objremoteDoc.Update()
                Else
                    objMainDoc = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                    objMainDoc.CardCode = objremoteDoc.CardCode
                    objMainDoc.DocDate = objremoteDoc.DocDate
                    objMainDoc.DocDueDate = objremoteDoc.DocDueDate
                    objMainDoc.NumAtCard = objremoteDoc.NumAtCard
                    objMainDoc.Comments = objremoteDoc.Comments
                    objMainDoc.DiscountPercent = objremoteDoc.DiscountPercent
                    objMainDoc.DocCurrency = objremoteDoc.DocCurrency
                    objMainDoc.ShipToCode = objremoteDoc.ShipToCode
                    objMainDoc.SalesPersonCode = objremoteDoc.SalesPersonCode
                    objMainDoc.TaxDate = objremoteDoc.TaxDate
                    objMainDoc.PaymentGroupCode = objremoteDoc.PaymentGroupCode
                    objMainDoc.Address = objremoteDoc.Address
                    objMainDoc.Address2 = objremoteDoc.Address2
                    objMainDoc.AgentCode = objremoteDoc.AgentCode
                    objMainDoc.BPChannelCode = objremoteDoc.BPChannelCode
                    objMainDoc.BPChannelContact = objremoteDoc.BPChannelContact
                    objMainDoc.ContactPersonCode = objremoteDoc.ContactPersonCode
                    objMainDoc.DocRate = objremoteDoc.DocRate
                    objMainDoc.DocType = objremoteDoc.DocType
                    '  objMainDoc.TransportationCode = objremoteDoc.TransportationCode
                    If objremoteDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES Then
                        objMainDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES
                        objMainDoc.RoundingDiffAmount = objremoteDoc.RoundingDiffAmount
                    Else
                        objMainDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tNO
                    End If
                    'Document_LinesClass' 
                    objMainDoc.UserFields.Fields.Item("U_Import").Value = "Y"
                    objMainDoc.UserFields.Fields.Item("U_BaseEntry").Value = objremoteDoc.DocEntry
                    objMainDoc.UserFields.Fields.Item("U_BaseNum").Value = (objremoteDoc.DocNum)
                    objMainDoc.UserFields.Fields.Item("U_Branch").Value = objRemoteCompany.CompanyName
                    For IntExp As Integer = 0 To objremoteDoc.Expenses.Count - 1
                        If objremoteDoc.Expenses.LineTotal > 0 Then
                            If IntExp > 0 Then
                                objMainDoc.Expenses.Add()
                                objMainDoc.Expenses.SetCurrentLine(IntExp)
                            End If
                            objremoteDoc.Expenses.SetCurrentLine(IntExp)
                            objMainDoc.Expenses.BaseDocEntry = objremoteDoc.Expenses.BaseDocEntry
                            objMainDoc.Expenses.BaseDocLine = objremoteDoc.Expenses.BaseDocLine
                            objMainDoc.Expenses.BaseDocType = objremoteDoc.Expenses.BaseDocType
                            objMainDoc.Expenses.DistributionMethod = objremoteDoc.Expenses.DistributionMethod
                            objMainDoc.Expenses.DistributionRule = objremoteDoc.Expenses.DistributionRule
                            objMainDoc.Expenses.ExpenseCode = objremoteDoc.Expenses.ExpenseCode
                            objMainDoc.Expenses.LastPurchasePrice = objremoteDoc.Expenses.LastPurchasePrice
                            objMainDoc.Expenses.LineTotal = objremoteDoc.Expenses.LineTotal
                            objMainDoc.Expenses.Remarks = objremoteDoc.Expenses.Remarks
                            objMainDoc.Expenses.TaxCode = objremoteDoc.Expenses.TaxCode
                            objMainDoc.Expenses.VatGroup = objremoteDoc.Expenses.VatGroup
                        End If
                    Next

                    For intLoop As Integer = 0 To objremoteDoc.Lines.Count - 1
                        If intLoop > 0 Then
                            objMainDoc.Lines.Add()
                            objMainDoc.Lines.SetCurrentLine(intLoop)
                        End If
                        objremoteDoc.Lines.SetCurrentLine(intLoop)
                        objMainDoc.Lines.AccountCode = objremoteDoc.Lines.AccountCode
                        objMainDoc.Lines.ItemDescription = objremoteDoc.Lines.ItemDescription
                        objMainDoc.Lines.ItemCode = objremoteDoc.Lines.ItemCode
                        objMainDoc.Lines.BarCode = objremoteDoc.Lines.BarCode
                        objMainDoc.Lines.UnitPrice = objremoteDoc.Lines.UnitPrice
                        objMainDoc.Lines.AccountCode = objremoteDoc.Lines.AccountCode
                        objMainDoc.Lines.DiscountPercent = objremoteDoc.Lines.DiscountPercent
                        objMainDoc.Lines.VatGroup = objremoteDoc.Lines.VatGroup
                        objMainDoc.Lines.PriceAfterVAT = objremoteDoc.Lines.PriceAfterVAT
                        objMainDoc.Lines.LineTotal = objremoteDoc.Lines.LineTotal
                        objMainDoc.Lines.ProjectCode = objremoteDoc.Lines.ProjectCode
                        objMainDoc.Lines.Quantity = objremoteDoc.Lines.Quantity
                        objMainDoc.Lines.WarehouseCode = objremoteDoc.Lines.WarehouseCode
                    Next

                    If objMainDoc.Add <> 0 Then
                        sPath = strErrorLogPath
                        WriteErrorlog("Failed to export Credit Notes docuemnt :" & objremoteDoc.DocNum & ": Error :" & objMainCompany.GetLastErrorDescription & "  : Branch : " & objRemoteCompany.CompanyName, sPath)
                    Else
                        Dim strDocNum As String
                        objMainCompany.GetNewObjectCode(strDocNum)
                        'MsgBox("DOcument created :" & strDocNum)
                        sPath = strErrorLogPath
                        Dim oInvoice As SAPbobsCOM.Documents
                        Dim intDocEntry As Integer
                        oInvoice = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes)
                        If oInvoice.GetByKey(Convert.ToInt64(strDocNum)) Then
                            strDocNum = oInvoice.DocNum
                            intDocEntry = oInvoice.DocEntry
                        End If
                        WriteErrorlog("Credit Notes Created Successfully. DocNum : " & strDocNum, sPath)
                        objremoteDoc.UserFields.Fields.Item("U_Export").Value = "Y"
                        objremoteDoc.Update()
                        Dim oItem As SAPbobsCOM.Documents
                        oItem = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                        blnDraft = False
                        If oItem.GetByKey(intDocEntry) Then
                            intDocEntry = oItem.DocEntry
                            If ValidateGoodsDocument(intDocEntry, "CreditNotes") = False Then
                                blnDraft = True
                            Else
                                blnDraft = False
                            End If
                            CreateGoodsDocument(intDocEntry, "CreditNotes")
                        End If
                    End If
                End If
            End If
            objremoteRec.MoveNext()
        Next
    End Sub

#Region "validate onHandqty"
    Private Function validateOnhand(ByVal aItemCode As String, ByVal aWhs As String, ByVal dblqty As Double) As Boolean
        Dim oTempRs As SAPbobsCOM.Recordset
        Dim dblOnHand As Double
        oTempRs = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRs.DoQuery("Select * from OITW where itemcode='" & aItemCode & "' and whscode='" & aWhs & "'")
        dblOnHand = 0
        If oTempRs.RecordCount > 0 Then
            dblOnHand = oTempRs.Fields.Item("OnHand").Value - oTempRs.Fields.Item("MinStock").Value
        End If
        If dblOnHand >= dblqty Then
            Return True
        Else
            Return False
        End If
    End Function
#End Region

#Region "Create Goods Issue /Goods Receipt Document"

    Public Sub CreateGoodsDocument(ByVal aDocEntry As Integer, ByVal strChoice As String)
        Dim oDocument As SAPbobsCOM.Documents
        Dim strSQL As String
        Dim strItem, strwhs, strRec, strDocumentNumber As String
        Dim dblBaseQty, dblParentQty As Double
        Dim blnDocument As Boolean = False
        Dim intLine As Integer = 0
        Dim oRecInvoice, oRecBOM, oTempRecordset As SAPbobsCOM.Recordset
        If strChoice = "Invoice" Then
            strSQL = "SELECT T1.[ItemCode], T1.[ItemName], T1.[TreeType],T0.[Quantity],T0.[WhsCode] FROM INV1 T0  INNER JOIN OITM T1 ON T0.ItemCode = T1.ItemCode INNER JOIN OITT T2 ON T1.ItemCode = T2.Code WHERE   T0.DocEntry=" & aDocEntry & " and  T1.[TreeType] ='P' and T2.[U_SaleBOM]='Y'"
            strRec = "Select * from OINV where Docentry=" & aDocEntry
            If blnDraft = False Then
                oDocument = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
            Else
                oDocument = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
                oDocument.DocObjectCode = SAPbobsCOM.BoObjectTypes.oInventoryGenExit
            End If
        Else
            strRec = "Select * from ORIN where Docentry=" & aDocEntry
            strSQL = "SELECT T1.[ItemCode], T1.[ItemName], T1.[TreeType],T0.[Quantity],T0.[WhsCode] FROM RIN1 T0  INNER JOIN OITM T1 ON T0.ItemCode = T1.ItemCode INNER JOIN OITT T2 ON T1.ItemCode = T2.Code WHERE T0.DocEntry=" & aDocEntry & " and  T1.[TreeType] ='P' and T2.[U_SaleBOM]='Y'"
            oDocument = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)
        End If
        oRecInvoice = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecBOM = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecInvoice.DoQuery(strSQL)
        oTempRecordset = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRecordset.DoQuery(strRec)
        strDocumentNumber = oTempRecordset.Fields.Item("DocNum").Value
        For intRow As Integer = 0 To oRecInvoice.RecordCount - 1
            strItem = oRecInvoice.Fields.Item(0).Value
            dblParentQty = oRecInvoice.Fields.Item(3).Value
            strwhs = oRecInvoice.Fields.Item(4).Value
            oDocument.Reference2 = strDocumentNumber
            oRecBOM.DoQuery("Select * from ITT1 where Father='" & strItem & "'")
            For intLoop As Integer = 0 To oRecBOM.RecordCount - 1
                dblBaseQty = oRecBOM.Fields.Item("Quantity").Value
                dblBaseQty = dblBaseQty * dblParentQty
                If intLine > 0 Then
                    oDocument.Lines.Add()
                    oDocument.Lines.SetCurrentLine(intLine)
                End If
                oDocument.Lines.ItemCode = oRecBOM.Fields.Item("Code").Value
                oDocument.Lines.Quantity = dblBaseQty
                oDocument.Lines.WarehouseCode = strwhs
                blnDocument = True
                intLine = intLine + 1
                oRecBOM.MoveNext()
            Next
            oRecInvoice.MoveNext()
        Next
        If blnDocument = True Then
            If oDocument.Add <> 0 Then
                sPath = strErrorLogPath
                WriteErrorlog("Failed to Goods Issue  / Goods receipt docuemnt :" & objMainCompany.GetLastErrorDescription & "  : Branch : " & objRemoteCompany.CompanyName, sPath)

            Else
                Dim strDocNum As String
                objMainCompany.GetNewObjectCode(strDocNum)
                sPath = strErrorLogPath
                'WriteErrorlog()
                If strChoice = "Invoice" And blnDraft = False Then
                    WriteErrorlog("Goods Issue Document " & strDocNum & " created  successfully", sPath)
                ElseIf strChoice = "Invoice" And blnDraft = True Then
                    WriteErrorlog("Goods Issue Draft Document " & strDocNum & "  created  successfully ", sPath)
                Else
                    WriteErrorlog("Goods receipt Document " & strDocNum & " created  successfully ", sPath)
                End If
            End If
        End If
    End Sub

    Public Function ValidateGoodsDocument(ByVal aintDocEntry As Integer, ByVal stChoice As String) As Boolean
        Dim strSQL, CompanyDecimalSeprator, CompanyThousandSeprator As String
        Dim oRecInvoice, oRecBOM, oRecTemp As SAPbobsCOM.Recordset
        oRecInvoice = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecTemp = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecBOM = objMainCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim strItem, strwhs, strQty As String
        Dim dblBaseQty, dblParentQty As Double
        Dim intLine As Integer = 0
        If stChoice = "Invoice" Then
            oRecTemp.DoQuery("Select ItemCode,quantity,whsCode from INV1 where docentry=" & aintDocEntry)
        Else
            oRecTemp.DoQuery("Select ItemCode,quantity,whsCode from RIN1 where docentry=" & aintDocEntry)
        End If
        For intRow As Integer = 0 To oRecTemp.RecordCount - 1
            strItem = oRecTemp.Fields.Item(0).Value
            strQty = oRecTemp.Fields.Item(1).Value
            With objMainCompany.GetCompanyService
                CompanyDecimalSeprator = .GetAdminInfo.DecimalSeparator
                CompanyThousandSeprator = .GetAdminInfo.ThousandsSeparator
            End With
            If CompanyDecimalSeprator <> "." Then
                strQty = strQty.Replace(".", CompanyDecimalSeprator)
                dblParentQty = Convert.ToDouble(strQty)
            Else
                dblParentQty = Convert.ToDouble(strQty)
            End If
            strSQL = "SELECT T1.[ItemCode], T1.[ItemName], T1.[TreeType] FROM OITM T1  INNER JOIN OITT T2 ON T1.ItemCode = T2.Code WHERE  t1.itemcode='" & strItem & "' and  T1.[TreeType] ='P' and T2.[U_SaleBOM]='Y'"
            oRecInvoice.DoQuery(strSQL)
            If oRecInvoice.RecordCount > 0 Then
                strwhs = oRecTemp.Fields.Item(2).Value
                oRecBOM.DoQuery("Select * from ITT1 where Father='" & strItem & "'")
                For intLoop As Integer = 0 To oRecBOM.RecordCount - 1
                    dblBaseQty = oRecBOM.Fields.Item("Quantity").Value
                    dblBaseQty = dblBaseQty * dblParentQty
                    strItem = oRecBOM.Fields.Item("Code").Value
                    If validateOnhand(strItem, strwhs, dblBaseQty) = False Then
                        Return False
                    End If
                    oRecBOM.MoveNext()
                Next
            End If
            oRecTemp.MoveNext()
        Next
        Return True
    End Function
#End Region

    Private Sub CopyFilestoCustomers(ByVal aFileName As String, ByVal aLogPath As String)
        Dim otemp As SAPbobsCOM.Recordset
        Dim strFilePath, strDesgfilename, strMessage As String
        otemp = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp.DoQuery("Select * from OCRD where cardtype='C' and U_PharmInt = 'Y'")
        strFilePath = strExportFilePaty
        For intRow As Integer = 0 To otemp.RecordCount - 1
            strFilePath = strFilePath & "\" & otemp.Fields.Item("CardCode").Value
            If Directory.Exists(strFilePath) Then
            Else
                Directory.CreateDirectory(strFilePath)
            End If
            strDesgfilename = strFilePath & "\PROMFLQ.mfp"
            If File.Exists(strDesgfilename) Then
                File.Delete(strDesgfilename)
            End If
            File.Copy(aFileName, strDesgfilename)
            strFilePath = strExportFilePaty
            strMessage = "Exported :  File name : " & strDesgfilename
            WriteErrorlog(strMessage, aLogPath)
            otemp.MoveNext()
        Next


    End Sub

#Region "Connect MainCompany"
    Private Function connectMainCompany() As Boolean
        sPath = strErrorLogPath
        Try
            If cmbMainSAPCompany.Text = "" Then
                '   MsgBox("Choose Company")
                Return False
            Else
                ' strMainSAPCompany = cmbMainSAPCompany.Text
            End If
            If MainSAPUID.Text = "" Then
                ' MsgBox("Enter UserName")
                Return False
            End If
            If MainSAPPWD.Text = "" Then
                '  MsgBox("Enter Password")
                Return False
            End If
            sPath = strErrorLogPath
            oCompany = New SAPbobsCOM.Company
            oCompany.Server = MainSQLServer.Text
            oCompany.LicenseServer = mainLicenseServer.Text ' "Compaq-PC:30000"
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
            ' oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005
            If strMainServertype = "2005" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005
            ElseIf strLocalServertype = "2008" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
            ElseIf strLocalServertype = "2012" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
            ElseIf strLocalServertype = "2014" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014
            End If
            oCompany.DbUserName = MainSQLUID.Text
            oCompany.DbPassword = MainSQLPWD.Text
            Dim oRecs As SAPbobsCOM.Recordset
            '    oRecs = oCompany.GetCompanyList()
            WriteErrorlog("Connection to Main SAP B1 started : " & strMainSAPCompany, sPath)
            ' MsgBox(oRecs.RecordCount)
            If oCompany.Connect <> 0 Then
                ' Return False
            Else
                'MsgBox("Connected")
            End If
            If oCompany.Connected Then
                'MsgBox("Company connected")
            End If
            oCompany.CompanyDB = strMainSAPCompany ' cmbMainSAPCompany.Text
            oCompany.UserName = MainSAPUID.Text
            oCompany.Password = MainSAPPWD.Text
            If oCompany.Connected = True Then
                'If (objPC.Initialise(oCompany)) Then
                'Else
                '    ' MsgBox("Error in Connection")
                'End If
                WriteiniFile()
                objMainCompany = oCompany
                Return True
            Else
                If oCompany.Connect <> 0 Then
                    'MsgBox("Connection to Main SAP B1 failed. Error Description :" & oCompany.GetLastErrorDescription)
                    WriteErrorlog("Connection to Main SAP B1 : " & strMainSAPCompany & "  failed. Error Description :" & oCompany.GetLastErrorDescription, sPath)
                    Return False
                Else
                    WriteiniFile()
                    'MsgBox("Main SAP B1 company Connected successfully")
                    WriteErrorlog("Main SAP B1 : " & strMainSAPCompany & " Company connected", sPath)
                    objMainCompany = oCompany
                    Return True
                End If
            End If
        Catch ex As Exception
            ' MsgBox(ex.Message)
            WriteErrorlog(ex.Message, sPath)
            Return False
        End Try
        Return True
    End Function

    Private Function connectMainCompanyCheck() As Boolean
        sPath = strErrorLogPath
        Try

            sPath = strErrorLogPath
            oCompany = New SAPbobsCOM.Company
            oCompany.Server = strMainSQLServer
            oCompany.LicenseServer = strMainLicenseServer ' "Compaq-PC:30000"
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
            ' oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005
            If strMainServertype = "2005" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005
            ElseIf strLocalServertype = "2008" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
            ElseIf strLocalServertype = "2012" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
            ElseIf strLocalServertype = "2014" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014
            End If
            oCompany.DbUserName = strMainSQLUID
            oCompany.DbPassword = strMainSQLPWD
            Dim oRecs As SAPbobsCOM.Recordset
            '    oRecs = oCompany.GetCompanyList()
            WriteErrorlog("Connection to Main SAP B1 started : " & strMainSAPCompany, sPath)
            ' MsgBox(oRecs.RecordCount)
            If oCompany.Connect <> 0 Then
                ' Return False
            Else
                'MsgBox("Connected")
            End If
            If oCompany.Connected Then
                'MsgBox("Company connected")
            End If
            oCompany.CompanyDB = strMainSAPCompany ' cmbMainSAPCompany.Text
            oCompany.UserName = strMainSAPUID
            oCompany.Password = strMainSAPPWD
            If oCompany.Connected = True Then
                'If (objPC.Initialise(oCompany)) Then
                'Else
                '    ' MsgBox("Error in Connection")
                'End If
                WriteiniFile()
                objMainCompany = oCompany
                Return True
            Else
                If oCompany.Connect <> 0 Then
                    'MsgBox("Connection to Main SAP B1 failed. Error Description :" & oCompany.GetLastErrorDescription)
                    WriteErrorlog("Connection to Main SAP B1 : " & strMainSAPCompany & "  failed. Error Description :" & oCompany.GetLastErrorDescription, sPath)
                    Return False
                Else
                    WriteiniFile()
                    'MsgBox("Main SAP B1 company Connected successfully")
                    WriteErrorlog("Main SAP B1 : " & strMainSAPCompany & " Company connected", sPath)
                    objMainCompany = oCompany
                    Return True
                End If
            End If
        Catch ex As Exception
            ' MsgBox(ex.Message)
            WriteErrorlog(ex.Message, sPath)
            Return False
        End Try
        Return True
    End Function

    Private Function connectLocalCompany() As Boolean
        sPath = strErrorLogPath
        Try
            If cmbCompanyDB.Text = "" Then
                '  MsgBox("Choose Company")
                Return False
            Else
                ' strSAPCompany = "" cmbCompanyDB.Text
            End If
            If txtServerName.Text = "" Then
                '  MsgBox("Enter UserName")
                Return False
            End If
            If txtServerPwd.Text = "" Then
                '  MsgBox("Enter Password")
                Return False
            End If
            oCompany = New SAPbobsCOM.Company
            oCompany.Server = txtServerName.Text
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
            If strLocalServertype = "2005" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005
            ElseIf strLocalServertype = "2008" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
            ElseIf strLocalServertype = "2012" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
            ElseIf strLocalServertype = "2014" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014
            End If

            oCompany.LicenseServer = LocalLicenseServer.Text ' "Compaq-PC:30000"
            oCompany.DbUserName = txtDBUserName.Text
            oCompany.DbPassword = txtServerPwd.Text
            oCompany.CompanyDB = strSAPCompany ' cmbCompanyDB.Text
            oCompany.UserName = txtSBOUserName.Text
            oCompany.Password = txtCompanyPWD.Text
            WriteErrorlog("Connection to Retail DB  : " & strSAPCompany & " started    ", sPath)
            If oCompany.Connected = True Then
                If (objPC.Initialise(oCompany)) Then
                Else
                    'MsgBox("Error in Connection")
                    WriteErrorlog("Error in Connection to Local Server", sPath)
                    Return False
                End If
                '  WriteiniFile_Folder(sPath)
                WriteiniFile()
                objRemoteCompany = oCompany
                Return True
            Else
                If oCompany.Connect <> 0 Then
                    WriteErrorlog("Connection to Retail DB  : " & strSAPCompany & " failed. : " & strSAPCompany & " : Error Description :" & oCompany.GetLastErrorDescription, sPath)
                    Return False
                Else
                    WriteiniFile()
                    WriteErrorlog("Retail DB  : " & strSAPCompany & " Connected successfully", sPath)
                    objRemoteCompany = oCompany
                    Return True
                End If
            End If
        Catch ex As Exception
            ' MsgBox(ex.Message)
            WriteErrorlog(ex.Message, sPath)
            Return False
        End Try
        Return True
    End Function

    Private Function connectLocalCompanyCheck() As Boolean
        sPath = strErrorLogPath
        Try

            oCompany = New SAPbobsCOM.Company
            oCompany.Server = strSQLServer
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
            If strLocalServertype = "2005" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005
            ElseIf strLocalServertype = "2008" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
            ElseIf strLocalServertype = "2012" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
            ElseIf strLocalServertype = "2014" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014
            End If

            oCompany.LicenseServer = strLocalLicensserver ' "Compaq-PC:30000"
            oCompany.DbUserName = strSQLUID
            oCompany.DbPassword = strSQLPWD
            oCompany.CompanyDB = strSAPCompany ' cmbCompanyDB.Text
            oCompany.UserName = strSAPUID
            oCompany.Password = strSAPPWD
            WriteErrorlog("Connection to Retail DB  : " & strSAPCompany & " started    ", sPath)
            If oCompany.Connected = True Then
                If (objPC.Initialise(oCompany)) Then
                Else
                    'MsgBox("Error in Connection")
                    WriteErrorlog("Error in Connection to Local Server", sPath)
                    Return False
                End If
                '  WriteiniFile_Folder(sPath)
                WriteiniFile()
                objRemoteCompany = oCompany
                Return True
            Else
                If oCompany.Connect <> 0 Then
                    WriteErrorlog("Connection to Retail DB  : " & strSAPCompany & " failed. : " & strSAPCompany & " : Error Description :" & oCompany.GetLastErrorDescription, sPath)
                    Return False
                Else
                    WriteiniFile()
                    WriteErrorlog("Retail DB  : " & strSAPCompany & " Connected successfully", sPath)
                    objRemoteCompany = oCompany
                    Return True
                End If
            End If
        Catch ex As Exception
            ' MsgBox(ex.Message)
            WriteErrorlog(ex.Message, sPath)
            Return False
        End Try
        Return True
    End Function
#End Region

    Private Sub btnConnect_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConnect.Click
        Try
            If connectLocalCompany() = False Then
                MsgBox("Error in connection to local server")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub btnBrowse_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        fldFolderBrowse.ShowDialog()
        txtDirectory.Text = fldFolderBrowse.SelectedPath
    End Sub

#End Region

    Private Sub cmbCompanyDB_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCompanyDB.SelectedIndexChanged

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            If txtServerName.Text = "" Then
                MsgBox("Enter Server")
                Exit Sub
            End If
            If txtDBUserName.Text = "" Then
                MsgBox("Enter UserName")
                Exit Sub
            End If
            If txtServerPwd.Text = "" Then
                MsgBox("Enter Password")
                Exit Sub
            End If
            objMainCompany = New SAPbobsCOM.Company
            objMainCompany.Server = MainSQLServer.Text
            objMainCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
            strMainServertype = cmbMainServertype.Text
            If strMainServertype = "2005" Then
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005
            Else
                oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
            End If
            objMainCompany.DbUserName = MainSQLUID.Text
            objMainCompany.DbPassword = MainSQLPWD.Text
            Dim oRecordSet As SAPbobsCOM.Recordset
            Dim lErrCode As Long
            Dim sErrMsg As String
            oRecordSet = objMainCompany.GetCompanyList
            objMainCompany.GetLastError(lErrCode, sErrMsg)
            If lErrCode <> 0 Then
                MsgBox(sErrMsg)
                Exit Sub
            Else
                cmbCompanyDB.Items.Clear()
                Do Until oRecordSet.EoF = True
                    cmbMainSAPCompany.Items.Add(oRecordSet.Fields.Item(0).Value)
                    oRecordSet.MoveNext()
                Loop
                btnConnect.Enabled = True
            End If
        Catch ex As Exception
            MsgBox("Connection to Server =" & txtServerName.Text & ".  Failed")
            Exit Sub
        End Try
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If connectMainCompany() = False Then
            MsgBox("Error in connection to main server")
        End If
    End Sub
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        End
    End Sub

#Region "Get Payment Document Number"
    Private Function GetIncomingDocument(ByVal aDocEntry As Integer) As Integer
        Dim intPaymentdocentry As Integer
        Dim oRecset As SAPbobsCOM.Recordset
        oRecset = objRemoteCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecset.DoQuery("Select DocNum from RCT2 where DocEntry=" & aDocEntry)
        If oRecset.RecordCount > 0 Then
            intPaymentdocentry = oRecset.Fields.Item(0).Value
        Else
            intPaymentdocentry = 0
        End If
        Return intPaymentdocentry
    End Function
#End Region

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        MsgBox("Export process started")
        If connectLocalCompany() Then
            If connectMainCompany() Then
                sPath = strErrorLogPath
                WriteErrorlog("Export Process Started", sPath)
                Me.Hide()
                Me.Hide()
                WriteErrorlog("Export Invoice Document ", sPath)
                ExportSalesOrer(oCompany)
                WriteErrorlog("Export Invoice+Payment Document ", sPath)
                ExportInvoicePayment(oCompany)
                WriteErrorlog("Export Credit Note Document ", sPath)
                ImportCreditNotes(oCompany)
                WriteErrorlog("Export Incoming Payment Document ", sPath)
                ExportIncomingPaymentDocument()
                WriteErrorlog("Import Master Data Started", sPath)
                ImportItemMaster()
                UpdateItemMaster()
                ImportBPMaster()
                UpdateBPMaster()
                WriteErrorlog("Import Master Data completed", sPath)
                WriteErrorlog("Export Stock Transfer Request started..", sPath)
                '  StocktransferRequest()
                WriteErrorlog("Export Stock Transfer Request Completed..", sPath)
                Me.Hide()
                Me.Hide()
                End
            End If
        End If
        MsgBox("Export process Compleated")
        End
    End Sub


    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        fldFolderBrowse.ShowDialog()
        txtDirectory.Text = fldFolderBrowse.SelectedPath
    End Sub
End Class
