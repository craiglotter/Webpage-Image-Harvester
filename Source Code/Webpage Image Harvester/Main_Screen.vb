Imports Microsoft.Win32
Imports System.Net
Imports System.IO
Imports System.Text
Public Class Main_Screen

    Inherits System.Windows.Forms.Form
    Dim download_item_queue As ArrayList
    Dim WithEvents Worker1 As Worker
    Private downloadstatus As Boolean
    Private downloadindex As Integer
    Private workerpaused As Boolean = False
    Public Delegate Sub WorkerhSnapShotTakenHandler(ByVal Result As String)
    Public Delegate Sub WorkerhStringMessageHandler(ByVal message As String, ByVal labelname As String)
    Public Delegate Sub WorkerhHandler(ByVal Result As String)
    Public Delegate Sub WorkerProgresshHandler(ByVal Result As String)
    Public Delegate Sub WorkerErrorhHandler(ByVal Message As Exception)


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Worker1 = New Worker
        AddHandler Worker1.WorkerStringMessage, AddressOf workerstringmessagehandler
        AddHandler Worker1.WorkerSnapShotTaken, AddressOf WorkerSnapShotTakenHandler
        AddHandler Worker1.WorkerComplete, AddressOf WorkerHandler
        AddHandler Worker1.WorkerProgress, AddressOf WorkerProgressHandler
        AddHandler Worker1.WorkerError, AddressOf WorkerErrorHandler
        download_item_queue = New ArrayList
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
    Friend WithEvents txtURL As System.Windows.Forms.TextBox
    Friend WithEvents txtFiles As System.Windows.Forms.CheckedListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtSavePath As System.Windows.Forms.TextBox
    Friend WithEvents txtProxy As System.Windows.Forms.TextBox
    Friend WithEvents txtAccept As System.Windows.Forms.ListBox
    Friend WithEvents txtStatus As System.Windows.Forms.Label
    Friend WithEvents Button_ListFiles As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents File_Type_Context_Menu As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents ProgressLabel As System.Windows.Forms.Label
    Friend WithEvents byteslabel As System.Windows.Forms.Label
    Friend WithEvents txtselectedfile As System.Windows.Forms.TextBox
    Friend WithEvents button_download As System.Windows.Forms.Button
    Friend WithEvents button_folderbrowse As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents txtParse As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtfilehandle3 As System.Windows.Forms.RadioButton
    Friend WithEvents txtfilehandle1 As System.Windows.Forms.RadioButton
    Friend WithEvents txtfilehandle2 As System.Windows.Forms.RadioButton
    Friend WithEvents txtLastDownload As System.Windows.Forms.TextBox
    Friend WithEvents Button_Pause2 As System.Windows.Forms.Button
    Friend WithEvents Button_Pause1 As System.Windows.Forms.Button
    Friend WithEvents Button_ExitThread As System.Windows.Forms.Button
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents fullerrors As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Main_Screen))
        Me.txtURL = New System.Windows.Forms.TextBox
        Me.button_download = New System.Windows.Forms.Button
        Me.txtFiles = New System.Windows.Forms.CheckedListBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtSavePath = New System.Windows.Forms.TextBox
        Me.button_folderbrowse = New System.Windows.Forms.Button
        Me.Button_ListFiles = New System.Windows.Forms.Button
        Me.txtProxy = New System.Windows.Forms.TextBox
        Me.txtAccept = New System.Windows.Forms.ListBox
        Me.File_Type_Context_Menu = New System.Windows.Forms.ContextMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.txtStatus = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.Label5 = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.ProgressLabel = New System.Windows.Forms.Label
        Me.byteslabel = New System.Windows.Forms.Label
        Me.txtselectedfile = New System.Windows.Forms.TextBox
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtfilehandle3 = New System.Windows.Forms.RadioButton
        Me.txtfilehandle1 = New System.Windows.Forms.RadioButton
        Me.txtfilehandle2 = New System.Windows.Forms.RadioButton
        Me.Label6 = New System.Windows.Forms.Label
        Me.Button_Pause2 = New System.Windows.Forms.Button
        Me.Button_Pause1 = New System.Windows.Forms.Button
        Me.fullerrors = New System.Windows.Forms.CheckBox
        Me.txtParse = New System.Windows.Forms.Label
        Me.txtLastDownload = New System.Windows.Forms.TextBox
        Me.Button_ExitThread = New System.Windows.Forms.Button
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtURL
        '
        Me.txtURL.BackColor = System.Drawing.Color.White
        Me.txtURL.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtURL.ForeColor = System.Drawing.Color.Black
        Me.txtURL.Location = New System.Drawing.Point(72, 8)
        Me.txtURL.Name = "txtURL"
        Me.txtURL.Size = New System.Drawing.Size(328, 20)
        Me.txtURL.TabIndex = 0
        Me.txtURL.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtURL, "The URL to be processed for HREF targets. (Multiple URLs can be delmited by ';')")
        '
        'button_download
        '
        Me.button_download.BackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(192, Byte), CType(0, Byte))
        Me.button_download.Enabled = False
        Me.button_download.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.button_download.ForeColor = System.Drawing.Color.White
        Me.button_download.Location = New System.Drawing.Point(528, 240)
        Me.button_download.Name = "button_download"
        Me.button_download.Size = New System.Drawing.Size(72, 23)
        Me.button_download.TabIndex = 4
        Me.button_download.Text = "Download"
        Me.ToolTip1.SetToolTip(Me.button_download, "Download the files contained in the File List")
        '
        'txtFiles
        '
        Me.txtFiles.BackColor = System.Drawing.Color.White
        Me.txtFiles.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFiles.CheckOnClick = True
        Me.txtFiles.ForeColor = System.Drawing.Color.Black
        Me.txtFiles.Location = New System.Drawing.Point(8, 72)
        Me.txtFiles.Name = "txtFiles"
        Me.txtFiles.Size = New System.Drawing.Size(392, 197)
        Me.txtFiles.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.txtFiles, "Accepted File types linked to in the inputted URL")
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 16)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "URL:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.ToolTip1.SetToolTip(Me.Label1, "The URL to be processed for HREF targets")
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 16)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Save Path:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.ToolTip1.SetToolTip(Me.Label2, "The folder to which all downloaded files are to be downloaded to")
        '
        'txtSavePath
        '
        Me.txtSavePath.BackColor = System.Drawing.Color.White
        Me.txtSavePath.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSavePath.ForeColor = System.Drawing.Color.Black
        Me.txtSavePath.Location = New System.Drawing.Point(72, 40)
        Me.txtSavePath.Name = "txtSavePath"
        Me.txtSavePath.Size = New System.Drawing.Size(264, 20)
        Me.txtSavePath.TabIndex = 1
        Me.txtSavePath.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtSavePath, "The folder to which all downloaded files are to be downloaded to")
        '
        'button_folderbrowse
        '
        Me.button_folderbrowse.BackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(192, Byte), CType(0, Byte))
        Me.button_folderbrowse.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.button_folderbrowse.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.button_folderbrowse.ForeColor = System.Drawing.Color.White
        Me.button_folderbrowse.Location = New System.Drawing.Point(344, 40)
        Me.button_folderbrowse.Name = "button_folderbrowse"
        Me.button_folderbrowse.Size = New System.Drawing.Size(56, 20)
        Me.button_folderbrowse.TabIndex = 2
        Me.button_folderbrowse.Text = "Browse"
        Me.ToolTip1.SetToolTip(Me.button_folderbrowse, "Select the folder to which all downloaded files are to be downloaded to")
        '
        'Button_ListFiles
        '
        Me.Button_ListFiles.BackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(192, Byte), CType(0, Byte))
        Me.Button_ListFiles.Enabled = False
        Me.Button_ListFiles.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button_ListFiles.ForeColor = System.Drawing.Color.White
        Me.Button_ListFiles.Location = New System.Drawing.Point(448, 240)
        Me.Button_ListFiles.Name = "Button_ListFiles"
        Me.Button_ListFiles.Size = New System.Drawing.Size(72, 23)
        Me.Button_ListFiles.TabIndex = 3
        Me.Button_ListFiles.Text = "List Files"
        Me.ToolTip1.SetToolTip(Me.Button_ListFiles, "Parse the inputted URL to locate all HREF tags within")
        '
        'txtProxy
        '
        Me.txtProxy.BackColor = System.Drawing.Color.White
        Me.txtProxy.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtProxy.ForeColor = System.Drawing.Color.Black
        Me.txtProxy.Location = New System.Drawing.Point(424, 208)
        Me.txtProxy.Name = "txtProxy"
        Me.txtProxy.Size = New System.Drawing.Size(192, 20)
        Me.txtProxy.TabIndex = 9
        Me.txtProxy.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtProxy, "Internet Connection Proxy (If necessary)")
        '
        'txtAccept
        '
        Me.txtAccept.BackColor = System.Drawing.Color.White
        Me.txtAccept.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtAccept.ContextMenu = Me.File_Type_Context_Menu
        Me.txtAccept.ForeColor = System.Drawing.Color.Black
        Me.txtAccept.Location = New System.Drawing.Point(424, 72)
        Me.txtAccept.Name = "txtAccept"
        Me.txtAccept.Size = New System.Drawing.Size(192, 106)
        Me.txtAccept.TabIndex = 8
        Me.ToolTip1.SetToolTip(Me.txtAccept, "The file extensions accepted by the Parser")
        '
        'File_Type_Context_Menu
        '
        Me.File_Type_Context_Menu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem2})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.Text = "Add File Type"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.Text = "Remove File Type"
        '
        'txtStatus
        '
        Me.txtStatus.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.txtStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStatus.ForeColor = System.Drawing.Color.Black
        Me.txtStatus.Location = New System.Drawing.Point(0, 366)
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.Size = New System.Drawing.Size(624, 23)
        Me.txtStatus.TabIndex = 14
        Me.txtStatus.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(424, 48)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(120, 16)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "File Types:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.ToolTip1.SetToolTip(Me.Label3, "The file extensions accepted by the Parser")
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(424, 184)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(120, 16)
        Me.Label4.TabIndex = 16
        Me.Label4.Text = "Proxy:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.ToolTip1.SetToolTip(Me.Label4, "Internet Connection Proxy (If necessary)")
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.SaddleBrown
        Me.Label5.Location = New System.Drawing.Point(6, 366)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(72, 23)
        Me.Label5.TabIndex = 17
        Me.Label5.Text = "Default Settings"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.ToolTip1.SetToolTip(Me.Label5, "Resets the Application Configuration Defaults")
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.ProgressLabel)
        Me.Panel1.Location = New System.Drawing.Point(8, 298)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(200, 16)
        Me.Panel1.TabIndex = 12
        Me.ToolTip1.SetToolTip(Me.Panel1, "Job Progress")
        '
        'ProgressLabel
        '
        Me.ProgressLabel.BackColor = System.Drawing.Color.Firebrick
        Me.ProgressLabel.Location = New System.Drawing.Point(0, 0)
        Me.ProgressLabel.Name = "ProgressLabel"
        Me.ProgressLabel.Size = New System.Drawing.Size(0, 16)
        Me.ProgressLabel.TabIndex = 0
        '
        'byteslabel
        '
        Me.byteslabel.Location = New System.Drawing.Point(216, 298)
        Me.byteslabel.Name = "byteslabel"
        Me.byteslabel.Size = New System.Drawing.Size(392, 16)
        Me.byteslabel.TabIndex = 19
        '
        'txtselectedfile
        '
        Me.txtselectedfile.BackColor = System.Drawing.Color.OliveDrab
        Me.txtselectedfile.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtselectedfile.ForeColor = System.Drawing.Color.White
        Me.txtselectedfile.Location = New System.Drawing.Point(8, 280)
        Me.txtselectedfile.Name = "txtselectedfile"
        Me.txtselectedfile.ReadOnly = True
        Me.txtselectedfile.Size = New System.Drawing.Size(608, 13)
        Me.txtselectedfile.TabIndex = 10
        Me.txtselectedfile.Text = ""
        '
        'txtfilehandle3
        '
        Me.txtfilehandle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtfilehandle3.ForeColor = System.Drawing.Color.Black
        Me.txtfilehandle3.Location = New System.Drawing.Point(560, 24)
        Me.txtfilehandle3.Name = "txtfilehandle3"
        Me.txtfilehandle3.Size = New System.Drawing.Size(56, 24)
        Me.txtfilehandle3.TabIndex = 7
        Me.txtfilehandle3.Text = "Ignore"
        Me.ToolTip1.SetToolTip(Me.txtfilehandle3, "Don't download the file")
        '
        'txtfilehandle1
        '
        Me.txtfilehandle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtfilehandle1.ForeColor = System.Drawing.Color.Black
        Me.txtfilehandle1.Location = New System.Drawing.Point(424, 24)
        Me.txtfilehandle1.Name = "txtfilehandle1"
        Me.txtfilehandle1.Size = New System.Drawing.Size(72, 24)
        Me.txtfilehandle1.TabIndex = 5
        Me.txtfilehandle1.Text = "Overwrite"
        Me.ToolTip1.SetToolTip(Me.txtfilehandle1, "Overwrite any existing file found")
        '
        'txtfilehandle2
        '
        Me.txtfilehandle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtfilehandle2.ForeColor = System.Drawing.Color.Black
        Me.txtfilehandle2.Location = New System.Drawing.Point(496, 24)
        Me.txtfilehandle2.Name = "txtfilehandle2"
        Me.txtfilehandle2.Size = New System.Drawing.Size(64, 24)
        Me.txtfilehandle2.TabIndex = 6
        Me.txtfilehandle2.Text = "Rename"
        Me.ToolTip1.SetToolTip(Me.txtfilehandle2, "Rename the downloaded file")
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(424, 8)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(164, 16)
        Me.Label6.TabIndex = 26
        Me.Label6.Text = "If Existing Files Encountered:"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.ToolTip1.SetToolTip(Me.Label6, "What to do if existing files are located in the save directory")
        '
        'Button_Pause2
        '
        Me.Button_Pause2.BackColor = System.Drawing.Color.Red
        Me.Button_Pause2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button_Pause2.ForeColor = System.Drawing.Color.White
        Me.Button_Pause2.Location = New System.Drawing.Point(528, 240)
        Me.Button_Pause2.Name = "Button_Pause2"
        Me.Button_Pause2.Size = New System.Drawing.Size(72, 23)
        Me.Button_Pause2.TabIndex = 27
        Me.Button_Pause2.Text = "Pause"
        Me.ToolTip1.SetToolTip(Me.Button_Pause2, "Download the files contained in the File List")
        Me.Button_Pause2.Visible = False
        '
        'Button_Pause1
        '
        Me.Button_Pause1.BackColor = System.Drawing.Color.Red
        Me.Button_Pause1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button_Pause1.ForeColor = System.Drawing.Color.White
        Me.Button_Pause1.Location = New System.Drawing.Point(448, 240)
        Me.Button_Pause1.Name = "Button_Pause1"
        Me.Button_Pause1.Size = New System.Drawing.Size(72, 23)
        Me.Button_Pause1.TabIndex = 28
        Me.Button_Pause1.Text = "Pause"
        Me.ToolTip1.SetToolTip(Me.Button_Pause1, "Download the files contained in the File List")
        Me.Button_Pause1.Visible = False
        '
        'fullerrors
        '
        Me.fullerrors.BackColor = System.Drawing.Color.OliveDrab
        Me.fullerrors.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.fullerrors.Enabled = False
        Me.fullerrors.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.fullerrors.ForeColor = System.Drawing.Color.White
        Me.fullerrors.Location = New System.Drawing.Point(608, 0)
        Me.fullerrors.Name = "fullerrors"
        Me.fullerrors.Size = New System.Drawing.Size(16, 16)
        Me.fullerrors.TabIndex = 21
        Me.ToolTip1.SetToolTip(Me.fullerrors, "Toggles between Full or Message Error Reporting")
        '
        'txtParse
        '
        Me.txtParse.ForeColor = System.Drawing.Color.Honeydew
        Me.txtParse.Location = New System.Drawing.Point(8, 344)
        Me.txtParse.Name = "txtParse"
        Me.txtParse.Size = New System.Drawing.Size(608, 16)
        Me.txtParse.TabIndex = 22
        Me.txtParse.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtLastDownload
        '
        Me.txtLastDownload.BackColor = System.Drawing.Color.OliveDrab
        Me.txtLastDownload.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtLastDownload.ForeColor = System.Drawing.Color.White
        Me.txtLastDownload.Location = New System.Drawing.Point(8, 322)
        Me.txtLastDownload.Name = "txtLastDownload"
        Me.txtLastDownload.ReadOnly = True
        Me.txtLastDownload.Size = New System.Drawing.Size(608, 13)
        Me.txtLastDownload.TabIndex = 11
        Me.txtLastDownload.Text = ""
        '
        'Button_ExitThread
        '
        Me.Button_ExitThread.BackColor = System.Drawing.Color.Red
        Me.Button_ExitThread.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button_ExitThread.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_ExitThread.ForeColor = System.Drawing.Color.White
        Me.Button_ExitThread.Location = New System.Drawing.Point(408, 240)
        Me.Button_ExitThread.Name = "Button_ExitThread"
        Me.Button_ExitThread.Size = New System.Drawing.Size(35, 23)
        Me.Button_ExitThread.TabIndex = 29
        Me.Button_ExitThread.Text = "EXIT"
        Me.Button_ExitThread.Visible = False
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.ContextMenu = Me.ContextMenu1
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "Webpage Image Harvester"
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem3, Me.MenuItem5, Me.MenuItem4})
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 0
        Me.MenuItem3.Text = "Show Webpage Image Harvester"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 1
        Me.MenuItem5.Text = "-"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 2
        Me.MenuItem4.Text = "Exit"
        '
        'Main_Screen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.OliveDrab
        Me.ClientSize = New System.Drawing.Size(624, 389)
        Me.Controls.Add(Me.Button_ExitThread)
        Me.Controls.Add(Me.txtLastDownload)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtfilehandle2)
        Me.Controls.Add(Me.txtfilehandle1)
        Me.Controls.Add(Me.txtfilehandle3)
        Me.Controls.Add(Me.txtParse)
        Me.Controls.Add(Me.fullerrors)
        Me.Controls.Add(Me.txtselectedfile)
        Me.Controls.Add(Me.byteslabel)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtStatus)
        Me.Controls.Add(Me.txtAccept)
        Me.Controls.Add(Me.txtProxy)
        Me.Controls.Add(Me.Button_ListFiles)
        Me.Controls.Add(Me.button_folderbrowse)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtSavePath)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtFiles)
        Me.Controls.Add(Me.button_download)
        Me.Controls.Add(Me.txtURL)
        Me.Controls.Add(Me.Button_Pause1)
        Me.Controls.Add(Me.Button_Pause2)
        Me.ForeColor = System.Drawing.Color.White
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(632, 423)
        Me.MinimumSize = New System.Drawing.Size(632, 423)
        Me.Name = "Main_Screen"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "Webpage Image Harvester (Build 20051110.1)"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message("The Application encountered the following problem: " & vbCrLf & identifier_msg & ":" & ex.ToString)
                Display_Message1.ShowDialog()
                Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & identifier_msg & ":" & ex.ToString)
                filewriter.Flush()
                filewriter.Close()
            End If
        Catch exc As Exception
            MsgBox("An error occurred in Webpage Image Harvester's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_ListFiles.Click

        Try
            workerpaused = False
            Label5.Hide()
            disable_controls()
            Button_ListFiles.Visible = False
            Button_ListFiles.Refresh()
            Button_Pause1.Visible = True
            Button_Pause1.Refresh()
            download_item_queue.Clear()
            txtFiles.Items.Clear()
            txtparse_message("")
            txtSelectedFile_Message("")
            txtLastDownload_Message("")




            Dim inc As System.Drawing.Size
            inc.Height = ProgressLabel.Size.Height
            inc.Width = 0
            ProgressLabel.BackColor = Color.Firebrick
            ProgressLabel.Size = inc
            ProgressLabel.Refresh()
            byteslabel.Text = ""
            byteslabel.Refresh()
            Dim acceptable As String

            Dim goforth As Boolean = True
            If Worker1.WorkerThread Is Nothing = False Then
                If (Worker1.WorkerThread.ThreadState.ToString.IndexOf("Aborted") > -1) Or (Worker1.WorkerThread.ThreadState.ToString.IndexOf("AbortRequested") > -1) Then
                    goforth = False
                End If
            End If
            If goforth = True Then
                Worker1.download_item_queue.Clear()
                Worker1.acceptable_items.Clear()
                For Each acceptable In txtAccept.Items
                    Worker1.acceptable_items.Add(acceptable)
                Next
                Worker1.proxy = txtProxy.Text
                Worker1.url = txtURL.Text
                Worker1.ChooseThreads(2)

            End If
        Catch ex As Exception
            
                Error_Handler(ex)

        End Try
    End Sub
    Private Sub txtparse_message(ByVal message As String)
        txtParse.Text = message
        txtParse.Refresh()
    End Sub

    Private Sub txtSelectedFile_Message(ByVal message As String)
        txtselectedfile.Text = message
        txtselectedfile.Refresh()
    End Sub
    Private Sub txtLastDownload_Message(ByVal message As String)
        txtLastDownload.Text = message
        txtLastDownload.Refresh()
    End Sub



    Private Sub txtFiles_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles txtFiles.ItemCheck
        Try
            Dim downitem As Download_Item = download_item_queue.Item(e.Index)
            downitem.filedownload = e.NewValue
            download_item_queue.Item(e.Index) = downitem
            If downitem.filedownload = True Then
                txtFiles.Items(e.Index) = txtFiles.Items(e.Index).replace("  [Ignore]", "") & "  [Scraped]"
            Else
                txtFiles.Items(e.Index) = txtFiles.Items(e.Index).replace("  [Scraped]", "") & "  [Ignore]"
            End If
            'txtFiles.SelectedIndex = e.Index
        Catch ex As Exception
          
                Error_Handler(ex)

        End Try
    End Sub

    Private Sub txtFiles_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFiles.SelectedIndexChanged
        Try
            If txtFiles.SelectedIndex > -1 Then
                Dim downitem As Download_Item = download_item_queue.Item(txtFiles.SelectedIndex)
                txtSelectedFile_Message(downitem.fileurl)
            End If
        Catch ex As Exception

            Error_Handler(ex)
           
        End Try
    End Sub




    Private Sub txtURL_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtURL.TextChanged
        Try
            If txtURL.Text.Length > 0 And txtSavePath.Text.Length > 0 Then
                Button_ListFiles.Enabled = True
            Else
                Button_ListFiles.Enabled = False
            End If
        Catch ex As Exception
           
                Error_Handler(ex)

        End Try
    End Sub

    Protected Sub txtStatus_Message(ByVal message As String)
        Try
            txtStatus.Text = message
            txtStatus.Refresh()
        Catch ex As Exception
           
                Error_Handler(ex)

        End Try
    End Sub

    Public Sub workerstringmessagehandler(ByVal message As String, ByVal labelname As String)
        Try
            Select Case labelname.ToLower
                Case "txtstatus"
                    txtStatus_Message(message)
                Case "txtparse"
                    txtparse_message(message)
                Case Else
                    txtStatus_Message(message)

            End Select
        Catch ex As Exception
            If fullerrors.Checked = False Then
                Error_Handler(ex)
            Else
                Error_Handler(ex)
            End If
        End Try
    End Sub

    Public Sub WorkerErrorHandler(ByVal Message As Exception)

        Try
            
                Error_Handler(Message)

        Catch ex As Exception
            If fullerrors.Checked = False Then
                Error_Handler(ex)
            Else
                Error_Handler(ex)
            End If
        End Try
    End Sub
    Public Sub WorkerSnapShotTakenHandler(ByVal Result As String)
        Try
            Dim inc As System.Drawing.Size
            inc.Height = ProgressLabel.Size.Height
            inc.Width = Panel1.Width
            ProgressLabel.BackColor = Color.SteelBlue()
            ProgressLabel.Size = inc
            ProgressLabel.Refresh()
            ' byteslabel.Text = "File List Processed"
            Dim i As Integer

            For i = 0 To Worker1.download_item_queue.Count - 1
                Dim downitem As Download_Item
                downitem = Worker1.download_item_queue(i)

                download_item_queue.Add(downitem)
                txtFiles.Items.Add(downitem.filename, True)

            Next

            If txtFiles.Items.Count > 1 Then
                txtparse_message(txtFiles.Items.Count & " files were scraped for download.")
            Else
                txtparse_message(txtFiles.Items.Count & " file was scraped for download.")
            End If

            enable_controls()
            Label5.Show()
            Button_Pause1.Visible = False
            Button_Pause1.Refresh()
            Button_ListFiles.Visible = True
            Button_ListFiles.Refresh()
        Catch ex As Exception
            If fullerrors.Checked = False Then
                Error_Handler(ex)
            Else
                Error_Handler(ex)
            End If
        End Try
    End Sub
    Public Sub WorkerHandler(ByVal Result As String)
        Try

            txtFiles.Items(downloadindex) = txtFiles.Items(downloadindex).remove(txtFiles.Items(downloadindex).indexof("["), txtFiles.Items(downloadindex).length - txtFiles.Items(downloadindex).indexof("[")) & "[" & Result & "]"
            txtFiles.SelectedIndex = (downloadindex)
            txtFiles.Refresh()
            If Not Result = "Ignored" Then
                txtLastDownload_Message(("Last File Handled: " & txtselectedfile.Text.Replace("Downloading: ", "")).Replace("Ignoring: ", ""))
            Else
                txtLastDownload_Message("")
            End If
            txtSelectedFile_Message("")
            txtStatus_Message("Handled: " & txtselectedfile.Text)

            downloadindex = downloadindex + 1
            retrieve_file(downloadindex)
        Catch ex As Exception
            If fullerrors.Checked = False Then
                Error_Handler(ex)
            Else
                Error_Handler(ex)
            End If
        End Try
    End Sub

    Public Sub WorkerProgressHandler(ByVal value As Integer, ByVal read As Long, ByVal toread As Long, ByVal unit As String, ByVal bytesread As Long, ByVal bytestoread As Long)
        Try
            Dim inc As System.Drawing.Size
            inc.Height = ProgressLabel.Size.Height
            If value = 200 And read < toread Then
                value = 198
            End If
            inc.Width = value

            ProgressLabel.Size = inc
            ProgressLabel.Refresh()
            If bytestoread = 0 Then
                byteslabel.Text = read & "/" & toread & " " & unit
            Else
                byteslabel.Text = read & "/" & toread & " " & unit & "   (" & bytesread & "/" & bytestoread & " bytes)"
            End If
            byteslabel.Refresh()
        Catch ex As Exception
            If fullerrors.Checked = False Then
                Error_Handler(ex)
            Else
                Error_Handler(ex)
            End If
        End Try
    End Sub

    Private Sub button_folderbrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button_folderbrowse.Click
        Try
            Dim result As DialogResult = FolderBrowserDialog1.ShowDialog
            If result = DialogResult.OK Or result = DialogResult.Yes Then
                txtSavePath.Text = FolderBrowserDialog1.SelectedPath
                If Not txtSavePath.Text.EndsWith("\") Then
                    txtSavePath.Text = txtSavePath.Text & "\"
                End If
            End If
        Catch ex As Exception
            If fullerrors.Checked = False Then
                Error_Handler(ex)
            Else
                Error_Handler(ex)
            End If
        End Try
    End Sub

    Public Sub Load_Registry_Values()
        Try


            Dim configflag As Boolean
            configflag = False
            Dim str As String
            Dim keyflag1 As Boolean = False
            Dim oReg As RegistryKey = Registry.LocalMachine
            Dim keys() As String = oReg.GetSubKeyNames()
            System.Array.Sort(keys)

            For Each str In keys
                If str.Equals("Software\Webpage Image Harvester") = True Then
                    keyflag1 = True
                    Exit For
                End If
            Next str

            If keyflag1 = False Then
                oReg.CreateSubKey("Software\Webpage Image Harvester")
            End If

            keyflag1 = False

            Dim oKey As RegistryKey = oReg.OpenSubKey("Software\Webpage Image Harvester", True)

            str = oKey.GetValue("savepath")
            If Not IsNothing(str) And Not (str = "") Then
                txtSavePath.Text = str
            Else
                configflag = True
                oKey.SetValue("savepath", "C:\")
                txtSavePath.Text = "C:\"
            End If

            str = oKey.GetValue("URL")
            If Not IsNothing(str) And Not (str = "") Then
                txtURL.Text = str
                'Else
                '    configflag = True
                '    oKey.SetValue("savepath", "C:\")
                '    txtSavePath.Text = "C:\"
            End If

            str = oKey.GetValue("proxy")
            If Not IsNothing(str) And Not (str = "") Then
                txtProxy.Text = str
                'Else
                '    configflag = True
                '    oKey.SetValue("proxy", "cache22.uct.ac.za:8080")
                '    txtProxy.Text = "cache22.uct.ac.za:8080"
            End If

            str = oKey.GetValue("acceptfile")
            If Not IsNothing(str) And Not (str = "") Then
                Dim a As String()
                Dim i As Integer
                a = str.Split(";")
                txtAccept.Items.Clear()
                For i = 0 To a.Length - 1
                    txtAccept.Items.Add(a(i))
                Next
            Else
                configflag = True
                oKey.SetValue("acceptfile", "jpg")
                txtAccept.Items.Clear()
                txtAccept.Items.Add("jpg")
            End If


            str = oKey.GetValue("filehandle")
            If Not IsNothing(str) And Not (str = "") Then
                Select Case CInt(str)
                    Case 1 : txtfilehandle1.Checked = True
                    Case 2 : txtfilehandle2.Checked = True
                    Case 3 : txtfilehandle3.Checked = True
                End Select
            Else
                configflag = True
                oKey.SetValue("filehandle", "2")
                txtfilehandle2.Checked = True
            End If

            oKey.Close()
            oReg.Close()



        Catch ex As Exception
            If fullerrors.Checked = False Then
                Error_Handler(ex)
            Else
                Error_Handler(ex)
            End If
        End Try
    End Sub

    Private Sub default_values()
        Try


            txtSavePath.Text = "C:\"
            'txtProxy.Text = "cache22.uct.ac.za:8080"
            txtAccept.Items.Clear()
            txtAccept.Items.Add("jpg")
            txtfilehandle2.Checked = True

        Catch ex As Exception
            If fullerrors.Checked = False Then
                Error_Handler(ex)
            Else
                Error_Handler(ex)
            End If
        End Try
    End Sub

    Private Sub Save_Registry_Values()
        Try
            Dim oReg As RegistryKey = Registry.LocalMachine
            Dim oKey As RegistryKey = oReg.OpenSubKey("Software\Webpage Image Harvester", True)

            oKey.SetValue("savepath", txtSavePath.Text)
            oKey.SetValue("URL", txtURL.Text)
            oKey.SetValue("proxy", txtProxy.Text)

            Dim a As String
            Dim i As Integer

            For i = 0 To txtAccept.Items.Count - 1
                a = a & txtAccept.Items(i).ToString() & ";"
            Next
            a = a.Remove(a.Length - 1, 1)
            oKey.SetValue("acceptfile", a)

            If txtfilehandle1.Checked = True Then
                oKey.SetValue("filehandle", "1")
            End If
            If txtfilehandle2.Checked = True Then
                oKey.SetValue("filehandle", "2")
            End If
            If txtfilehandle3.Checked = True Then
                oKey.SetValue("filehandle", "3")
            End If


            oKey.Close()
            oReg.Close()
        Catch ex As Exception
            If fullerrors.Checked = False Then
                Error_Handler(ex)
            Else
                Error_Handler(ex)
            End If
        End Try
    End Sub

    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click
        Try
            default_values()
        Catch ex As Exception
            If fullerrors.Checked = False Then
                Error_Handler(ex)
            Else
                Error_Handler(ex)
            End If
        End Try
    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click
        Try
            Save_Registry_Values()
            exit_threads()
            NotifyIcon1.Dispose()
            Application.Exit()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub show_application()
        Try
            Me.Opacity = 1
            NotifyIcon1.Visible = False

            Me.BringToFront()
            Me.Refresh()
            Me.ShowInTaskbar = True
            Me.WindowState = FormWindowState.Normal

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub NotifyIcon1_dblclick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.DoubleClick
        show_application()
    End Sub
    Private Sub NotifyIcon1_snglclick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.Click
        show_application()
    End Sub


    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        show_application()
    End Sub



    Private Sub monitor_program_resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        Try

            If Me.WindowState = FormWindowState.Minimized Then
                Me.ShowInTaskbar = False
                NotifyIcon1.Visible = True
                Me.Opacity = 0
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            Load_Registry_Values()
        Catch ex As Exception
            If fullerrors.Checked = False Then
                Error_Handler(ex)
            Else
                Error_Handler(ex)
            End If
        End Try
    End Sub

    Private Sub Main_Screen_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Try
            Save_Registry_Values()
        Catch ex As Exception
            If fullerrors.Checked = False Then
                Error_Handler(ex)
            Else
                Error_Handler(ex)
            End If
        End Try
    End Sub



    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        Try
            Dim inp As User_Input = New User_Input
            Dim result As DialogResult
            result = inp.ShowDialog
            If result = DialogResult.OK Then
                txtAccept.Items.Add(inp.TextBox1.Text.ToLower.Trim)
            End If
            inp.Dispose()
        Catch ex As Exception
            If fullerrors.Checked = False Then
                Error_Handler(ex)
            Else
                Error_Handler(ex)
            End If
        End Try
    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        Try
            If txtAccept.SelectedIndex > -1 Then
                txtAccept.Items.RemoveAt(txtAccept.SelectedIndex)
            End If
        Catch ex As Exception
            If fullerrors.Checked = False Then
                Error_Handler(ex)
            Else
                Error_Handler(ex)
            End If
        End Try
    End Sub

    Private Sub txtSavePath_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSavePath.TextChanged
        Try
            If txtURL.Text.Length > 0 And txtSavePath.Text.Length > 0 Then
                Button_ListFiles.Enabled = True
            Else
                Button_ListFiles.Enabled = False
            End If

        Catch ex As Exception
            If fullerrors.Checked = False Then
                Error_Handler(ex)
            Else
                Error_Handler(ex)
            End If
        End Try
    End Sub

    Private Sub retrieve_file(ByVal index As Integer)
        Try
            If downloadstatus = True Then
                ' MsgBox(index & "  " & download_item_queue.Count)
                If index < download_item_queue.Count Then
                    Dim inc As System.Drawing.Size
                    inc.Height = ProgressLabel.Size.Height
                    inc.Width = 0
                    ProgressLabel.BackColor = Color.Firebrick
                    ProgressLabel.Size = inc
                    ProgressLabel.Refresh()
                    byteslabel.Text = ""
                    byteslabel.Refresh()
                    'txtLastDownload_Message("")

                    Dim downitem As Download_Item
                    downitem = download_item_queue(index)
                    'For Each downitem In download_item_queue
                    ' If downloadstatus = False Then
                    ' MsgBox("attempting download - " & downitem.filedownload)
                    If downitem.filedownload = True Then
                        ' MsgBox("initiating")
                        txtSelectedFile_Message("Downloading: " & downitem.fileurl)
                        txtparse_message("Downloading file " & index + 1 & " of " & txtFiles.Items.Count & ".")
                        txtselectedfile.Refresh()
                        txtStatus_Message(txtselectedfile.Text)
                        Dim goforth As Boolean = True
                        If Worker1.WorkerThread Is Nothing = False Then
                            If (Worker1.WorkerThread.ThreadState.ToString.IndexOf("Aborted") > -1) Or (Worker1.WorkerThread.ThreadState.ToString.IndexOf("AbortRequested") > -1) Then
                                goforth = False
                            End If
                        End If
                        If goforth = True Then
                            Worker1.proxy = txtProxy.Text
                            Worker1.savepath = txtSavePath.Text
                            Worker1.download = downitem
                            If txtfilehandle1.Checked = True Then
                                Worker1.filehandle = 1
                            End If
                            If txtfilehandle2.Checked = True Then
                                Worker1.filehandle = 2
                            End If
                            If txtfilehandle3.Checked = True Then
                                Worker1.filehandle = 3
                            End If
                            Worker1.ChooseThreads(1)
                        End If
                    Else
                        txtSelectedFile_Message("Ignoring: " & downitem.fileurl)
                        txtparse_message("Ignoring file " & index + 1 & " of " & txtFiles.Items.Count & ".")
                        txtselectedfile.Refresh()
                        txtStatus_Message(txtselectedfile.Text)
                        WorkerHandler("Ignored")
                    End If
                Else
                    downloadstatus = False
                    Dim inc As System.Drawing.Size
                    inc.Height = ProgressLabel.Size.Height
                    inc.Width = Panel1.Width
                    ProgressLabel.BackColor = Color.SteelBlue()
                    ProgressLabel.Size = inc
                    ProgressLabel.Refresh()
                    byteslabel.Text = "File List Processed"
                    txtparse_message("Handled file " & index & " of " & txtFiles.Items.Count & ".")

                    byteslabel.Refresh()
                    txtStatus_Message(byteslabel.Text)
                    enable_controls()
                    Button_Pause2.Visible = False
                    Button_Pause2.Refresh()
                    button_download.Visible = True
                    button_download.Refresh()

                    Label5.Show()
                    txtURL.Select(0, 0)

                    downloadindex = 0
                End If
            End If
        Catch ex As Exception
            If fullerrors.Checked = False Then
                Error_Handler(ex)
            Else
                Error_Handler(ex)
            End If
        End Try
    End Sub 'Main
    Private Sub disable_controls()
        Try


            txtFiles.Enabled = False
            txtURL.Enabled = False
            txtSavePath.Enabled = False
            txtAccept.Enabled = False
            txtProxy.Enabled = False
            Button_ListFiles.Enabled = False
            button_download.Enabled = False
            button_folderbrowse.Enabled = False
            txtfilehandle1.Enabled = False
            txtfilehandle2.Enabled = False
            txtfilehandle3.Enabled = False
            Label5.Enabled = False
            'txtselectedfile.Enabled = False
            'txtLastDownload.Enabled = False
            Me.Refresh()
        Catch ex As Exception
            If fullerrors.Checked = False Then
                Error_Handler(ex)
            Else
                Error_Handler(ex)
            End If
        End Try
    End Sub
    Private Sub enable_controls()
        Try
            txtFiles.Enabled = True
            txtURL.Enabled = True
            txtSavePath.Enabled = True
            txtAccept.Enabled = True
            txtProxy.Enabled = True
            Button_ListFiles.Enabled = True
            button_download.Enabled = True
            button_folderbrowse.Enabled = True
            txtfilehandle1.Enabled = True
            txtfilehandle2.Enabled = True
            txtfilehandle3.Enabled = True
            Label5.Enabled = True
            ' txtselectedfile.Enabled = True
            'txtLastDownload.Enabled = True
            Me.Refresh()
        Catch ex As Exception
            If fullerrors.Checked = False Then
                Error_Handler(ex)
            Else
                Error_Handler(ex)
            End If
        End Try
    End Sub
    Private Sub button_download_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button_download.Click
        Try
            workerpaused = False
            Label5.Hide()
            downloadstatus = True
            If Not txtSavePath.Text.EndsWith("\") Then
                txtSavePath.Text = txtSavePath.Text & "\"
                txtSavePath.Refresh()
            End If
            disable_controls()
            button_download.Visible = False
            button_download.Refresh()
            Button_Pause2.Visible = True
            Button_Pause2.Refresh()
            downloadindex = 0
            txtLastDownload_Message("")
            retrieve_file(downloadindex)



        Catch ex As Exception
            If fullerrors.Checked = False Then
                Error_Handler(ex)
            Else
                Error_Handler(ex)
            End If
        End Try

    End Sub





    Private Sub Button_Pause_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_Pause2.Click, Button_Pause1.Click
        Try
            'MsgBox(Worker1.WorkerThread.ThreadState.ToString & " -> " & Worker1.WorkerThread.ThreadState.ToString.IndexOf(Worker1.WorkerThread.ThreadState.ToString))
            'Dim presshandled As Boolean = False
            'If presshandled = False Then
            'If Worker1.WorkerThread.ThreadState.ToString.IndexOf("Running") > -1 Or Worker1.WorkerThread.ThreadState.ToString.IndexOf("WaitSleepJoin") > -1 Then
            'MsgBox("entered block 1")
            If workerpaused = False Then
                Worker1.WorkerThread.Suspend()
                Button_Pause1.Text = "Resume"
                Button_Pause1.Refresh()
                Button_Pause2.Text = "Resume"
                Button_Pause2.Refresh()
                Button_ExitThread.Visible = True
                Button_ExitThread.Refresh()
                workerpaused = True
            Else
                'presshandled = True
                'End If
                'End If
                'MsgBox(Worker1.WorkerThread.ThreadState.ToString & " -> " & Worker1.WorkerThread.ThreadState.ToString.IndexOf(Worker1.WorkerThread.ThreadState.ToString))
                'If presshandled = False Then
                'If Worker1.WorkerThread.ThreadState.ToString.IndexOf("Suspended") > -1 Or Worker1.WorkerThread.ThreadState.ToString.IndexOf("SuspendRequested") > -1 Then
                'MsgBox("entered block 2")
                Worker1.WorkerThread.Resume()
                Button_Pause1.Text = "Pause"
                Button_Pause1.Refresh()
                Button_Pause2.Text = "Pause"
                Button_Pause2.Refresh()
                Button_ExitThread.Visible = False
                Button_ExitThread.Refresh()
                workerpaused = False
            End If
            'presshandled = True
            'End If
            'End If
            'MsgBox(Worker1.WorkerThread.ThreadState.ToString & " -> " & Worker1.WorkerThread.ThreadState.ToString.IndexOf(Worker1.WorkerThread.ThreadState.ToString))
        Catch ex As Exception
            If fullerrors.Checked = False Then
                Error_Handler(ex)
            Else
                Error_Handler(ex)
            End If
        End Try
    End Sub

    Private Sub exit_threads()
        Try

            If Worker1.WorkerThread.ThreadState.ToString.IndexOf("Suspended") > -1 Or Worker1.WorkerThread.ThreadState.ToString.IndexOf("SuspendRequested") > -1 Then
                Worker1.WorkerThread.Resume()
            End If
            If Worker1.WorkerThread.ThreadState.ToString.IndexOf("WaitSleepJoin") > -1 Then
                Worker1.WorkerThread.Interrupt()
            End If
            If Worker1.WorkerThread.ThreadState.ToString.IndexOf("AbortRequested") > -1 Then
                Worker1.WorkerThread.ResetAbort()
            End If

            txtStatus_Message("Aborting worker thread")
            'Worker1.WorkerThread.Priority() = Threading.ThreadPriority.Highest
            'Worker1.WorkerThread.MemoryBarrier()
            Worker1.WorkerThread.Abort()

            'Worker1.WorkerThread.Join()
            txtStatus_Message("Worker thread aborted")
            Button_ExitThread.Visible = False
            Button_ExitThread.Refresh()
            Button_Pause1.Visible = False
            Button_Pause1.Text = "Pause"
            Button_Pause2.Visible = False
            Button_Pause2.Text = "Pause"
            button_download.Visible = True
            Button_ListFiles.Visible = True
            enable_controls()
            Label5.Show()
            txtparse_message("")
            Worker1.Dispose()
            Worker1 = New Worker
            AddHandler Worker1.WorkerStringMessage, AddressOf workerstringmessagehandler
            AddHandler Worker1.WorkerSnapShotTaken, AddressOf WorkerSnapShotTakenHandler
            AddHandler Worker1.WorkerComplete, AddressOf WorkerHandler
            AddHandler Worker1.WorkerProgress, AddressOf WorkerProgressHandler
            AddHandler Worker1.WorkerError, AddressOf WorkerErrorHandler

        Catch ex As Exception
            If fullerrors.Checked = False Then
                Error_Handler(ex)
            Else
                Error_Handler(ex)
            End If
        End Try
    End Sub

    Private Sub Button_ExitThread_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button_ExitThread.Click
        exit_threads()
    End Sub

 

End Class
