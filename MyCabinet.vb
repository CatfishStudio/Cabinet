Public Class MyCabinet
    Inherits System.Windows.Forms.Form

#Region " Код, автоматически созданный конструктором форм Windows "

    Public Sub New()
        MyBase.New()

        'Этот вызов требуется конструктором форм Windows.
        InitializeComponent()

        'Добавьте код инициализации после вызова InitializeComponent()

    End Sub

    'Форма переопределяет метод Dispose для очистки списка компонентов.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Требуется конструктором форм Windows
    Private components As System.ComponentModel.IContainer

    'ПРИМЕЧАНИЕ: следующая процедура требуется для конструктора форм Windows.
    'Ее можно изменить в конструкторе форм Windows.  
    'Не изменяйте ее в редакторе исходного текста.
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents StatusBar1 As System.Windows.Forms.StatusBar
    Friend WithEvents StatusBarPanel1 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents StatusBarPanel2 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents TreeView1 As System.Windows.Forms.TreeView
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents StatusBarPanel3 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents OleDbSelectCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbConnection1 As System.Data.OleDb.OleDbConnection
    Friend WithEvents OleDbDataAdapter1 As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents DataSet21 As Cabinet.DataSet2
    Friend WithEvents DataSet31 As Cabinet.DataSet3
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents OleDbSelectCommand2 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand2 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand2 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand2 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDataAdapter2 As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents DataSet51 As Cabinet.DataSet5
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents ContextMenu2 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem9 As System.Windows.Forms.MenuItem
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents MenuItem10 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem11 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem12 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem13 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem14 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem15 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem16 As System.Windows.Forms.MenuItem
    Friend WithEvents ColorDialog1 As System.Windows.Forms.ColorDialog
    Friend WithEvents FontDialog1 As System.Windows.Forms.FontDialog
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(MyCabinet))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Panel6 = New System.Windows.Forms.Panel
        Me.Panel5 = New System.Windows.Forms.Panel
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.StatusBar1 = New System.Windows.Forms.StatusBar
        Me.StatusBarPanel1 = New System.Windows.Forms.StatusBarPanel
        Me.StatusBarPanel3 = New System.Windows.Forms.StatusBarPanel
        Me.StatusBarPanel2 = New System.Windows.Forms.StatusBarPanel
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.TreeView1 = New System.Windows.Forms.TreeView
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.MenuItem10 = New System.Windows.Forms.MenuItem
        Me.MenuItem11 = New System.Windows.Forms.MenuItem
        Me.MenuItem12 = New System.Windows.Forms.MenuItem
        Me.MenuItem13 = New System.Windows.Forms.MenuItem
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Splitter1 = New System.Windows.Forms.Splitter
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.Label15 = New System.Windows.Forms.Label
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox
        Me.ContextMenu2 = New System.Windows.Forms.ContextMenu
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.MenuItem7 = New System.Windows.Forms.MenuItem
        Me.MenuItem8 = New System.Windows.Forms.MenuItem
        Me.MenuItem9 = New System.Windows.Forms.MenuItem
        Me.MenuItem14 = New System.Windows.Forms.MenuItem
        Me.MenuItem15 = New System.Windows.Forms.MenuItem
        Me.MenuItem16 = New System.Windows.Forms.MenuItem
        Me.OleDbSelectCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbConnection1 = New System.Data.OleDb.OleDbConnection
        Me.OleDbInsertCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbDeleteCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbDataAdapter1 = New System.Data.OleDb.OleDbDataAdapter
        Me.DataSet21 = New Cabinet.DataSet2
        Me.DataSet31 = New Cabinet.DataSet3
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.OleDbSelectCommand2 = New System.Data.OleDb.OleDbCommand
        Me.OleDbInsertCommand2 = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand2 = New System.Data.OleDb.OleDbCommand
        Me.OleDbDeleteCommand2 = New System.Data.OleDb.OleDbCommand
        Me.OleDbDataAdapter2 = New System.Data.OleDb.OleDbDataAdapter
        Me.DataSet51 = New Cabinet.DataSet5
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ColorDialog1 = New System.Windows.Forms.ColorDialog
        Me.FontDialog1 = New System.Windows.Forms.FontDialog
        Me.Panel1.SuspendLayout()
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StatusBarPanel3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StatusBarPanel2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataSet21, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataSet31, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataSet51, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.Label8)
        Me.Panel1.Controls.Add(Me.Panel6)
        Me.Panel1.Controls.Add(Me.Panel5)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(592, 43)
        Me.Panel1.TabIndex = 2
        '
        'Label9
        '
        Me.Label9.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Label9.Image = CType(resources.GetObject("Label9.Image"), System.Drawing.Image)
        Me.Label9.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label9.Location = New System.Drawing.Point(344, 4)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(32, 32)
        Me.Label9.TabIndex = 10
        Me.ToolTip1.SetToolTip(Me.Label9, "О программе")
        '
        'Label8
        '
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Label8.Image = CType(resources.GetObject("Label8.Image"), System.Drawing.Image)
        Me.Label8.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label8.Location = New System.Drawing.Point(304, 4)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(32, 32)
        Me.Label8.TabIndex = 9
        Me.ToolTip1.SetToolTip(Me.Label8, "Настройки")
        '
        'Panel6
        '
        Me.Panel6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel6.Location = New System.Drawing.Point(280, 8)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(1, 24)
        Me.Panel6.TabIndex = 8
        '
        'Panel5
        '
        Me.Panel5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel5.Location = New System.Drawing.Point(136, 8)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(1, 24)
        Me.Panel5.TabIndex = 7
        '
        'Label7
        '
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Label7.Image = CType(resources.GetObject("Label7.Image"), System.Drawing.Image)
        Me.Label7.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label7.Location = New System.Drawing.Point(240, 4)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(32, 32)
        Me.Label7.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.Label7, "Удалить файл")
        '
        'Label6
        '
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Label6.Image = CType(resources.GetObject("Label6.Image"), System.Drawing.Image)
        Me.Label6.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label6.Location = New System.Drawing.Point(200, 4)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(32, 32)
        Me.Label6.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.Label6, "Открыть файл")
        '
        'Label5
        '
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Label5.Image = CType(resources.GetObject("Label5.Image"), System.Drawing.Image)
        Me.Label5.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label5.Location = New System.Drawing.Point(160, 4)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(32, 32)
        Me.Label5.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.Label5, "Создать новый файл")
        '
        'Label4
        '
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Label4.Image = CType(resources.GetObject("Label4.Image"), System.Drawing.Image)
        Me.Label4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label4.Location = New System.Drawing.Point(96, 4)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(32, 32)
        Me.Label4.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.Label4, "Удалить папку")
        '
        'Label3
        '
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Label3.Image = CType(resources.GetObject("Label3.Image"), System.Drawing.Image)
        Me.Label3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label3.Location = New System.Drawing.Point(56, 4)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(32, 32)
        Me.Label3.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.Label3, "Изменить имя папки")
        '
        'Label2
        '
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Label2.Image = CType(resources.GetObject("Label2.Image"), System.Drawing.Image)
        Me.Label2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label2.Location = New System.Drawing.Point(16, 4)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(32, 32)
        Me.Label2.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.Label2, "Создать новая папка")
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Image = CType(resources.GetObject("Label1.Image"), System.Drawing.Image)
        Me.Label1.Location = New System.Drawing.Point(408, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(184, 48)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "CABINET"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'StatusBar1
        '
        Me.StatusBar1.Location = New System.Drawing.Point(0, 406)
        Me.StatusBar1.Name = "StatusBar1"
        Me.StatusBar1.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.StatusBarPanel1, Me.StatusBarPanel3, Me.StatusBarPanel2})
        Me.StatusBar1.ShowPanels = True
        Me.StatusBar1.Size = New System.Drawing.Size(592, 16)
        Me.StatusBar1.TabIndex = 3
        Me.StatusBar1.Text = "StatusBar1"
        '
        'StatusBarPanel1
        '
        Me.StatusBarPanel1.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Contents
        Me.StatusBarPanel1.BorderStyle = System.Windows.Forms.StatusBarPanelBorderStyle.Raised
        Me.StatusBarPanel1.Text = "Copyright © 2012 Somov Evgen."
        Me.StatusBarPanel1.Width = 179
        '
        'StatusBarPanel3
        '
        Me.StatusBarPanel3.Text = "Дата:"
        '
        'StatusBarPanel2
        '
        Me.StatusBarPanel2.Text = "..."
        Me.StatusBarPanel2.Width = 800
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.White
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.TreeView1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Left
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(176, 406)
        Me.Panel2.TabIndex = 5
        '
        'TreeView1
        '
        Me.TreeView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TreeView1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TreeView1.ContextMenu = Me.ContextMenu1
        Me.TreeView1.ImageIndex = 2
        Me.TreeView1.ImageList = Me.ImageList1
        Me.TreeView1.Location = New System.Drawing.Point(0, 48)
        Me.TreeView1.Name = "TreeView1"
        Me.TreeView1.Size = New System.Drawing.Size(174, 355)
        Me.TreeView1.TabIndex = 2
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem2, Me.MenuItem3, Me.MenuItem4, Me.MenuItem5, Me.MenuItem10, Me.MenuItem11})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.Text = "Обновить."
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.Text = "-"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 2
        Me.MenuItem3.Text = "Открыть файл"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 3
        Me.MenuItem4.Text = "Создать файл"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 4
        Me.MenuItem5.Text = "Удалить файл"
        '
        'MenuItem10
        '
        Me.MenuItem10.Index = 5
        Me.MenuItem10.Text = "-"
        '
        'MenuItem11
        '
        Me.MenuItem11.Index = 6
        Me.MenuItem11.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem12, Me.MenuItem13})
        Me.MenuItem11.Text = "Сортировать:"
        '
        'MenuItem12
        '
        Me.MenuItem12.Checked = True
        Me.MenuItem12.Index = 0
        Me.MenuItem12.Text = "По алфавиту (возрастание)"
        '
        'MenuItem13
        '
        Me.MenuItem13.Index = 1
        Me.MenuItem13.Text = "По алфавиту (убывание)"
        '
        'ImageList1
        '
        Me.ImageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth32Bit
        Me.ImageList1.ImageSize = New System.Drawing.Size(24, 20)
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'Splitter1
        '
        Me.Splitter1.Location = New System.Drawing.Point(176, 0)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(2, 406)
        Me.Splitter1.TabIndex = 6
        Me.Splitter1.TabStop = False
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.White
        Me.Panel3.Controls.Add(Me.Label15)
        Me.Panel3.Controls.Add(Me.TextBox2)
        Me.Panel3.Controls.Add(Me.Panel4)
        Me.Panel3.Controls.Add(Me.GroupBox1)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(178, 0)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(414, 406)
        Me.Panel3.TabIndex = 8
        '
        'Label15
        '
        Me.Label15.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label15.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Label15.Image = CType(resources.GetObject("Label15.Image"), System.Drawing.Image)
        Me.Label15.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label15.Location = New System.Drawing.Point(343, 144)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(64, 24)
        Me.Label15.TabIndex = 4
        Me.Label15.Text = "Поиск."
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.Label15, "Поиск слов по тексту")
        '
        'TextBox2
        '
        Me.TextBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox2.Location = New System.Drawing.Point(7, 144)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(328, 20)
        Me.TextBox2.TabIndex = 3
        Me.TextBox2.Text = ""
        '
        'Panel4
        '
        Me.Panel4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel4.BackColor = System.Drawing.Color.White
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel4.Controls.Add(Me.Label16)
        Me.Panel4.Controls.Add(Me.Label14)
        Me.Panel4.Controls.Add(Me.Label13)
        Me.Panel4.Controls.Add(Me.Label12)
        Me.Panel4.Controls.Add(Me.ComboBox1)
        Me.Panel4.Controls.Add(Me.TextBox1)
        Me.Panel4.Controls.Add(Me.Label11)
        Me.Panel4.Controls.Add(Me.Label10)
        Me.Panel4.Location = New System.Drawing.Point(0, 42)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(414, 94)
        Me.Panel4.TabIndex = 1
        '
        'Label16
        '
        Me.Label16.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label16.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Label16.Image = CType(resources.GetObject("Label16.Image"), System.Drawing.Image)
        Me.Label16.Location = New System.Drawing.Point(392, 0)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(21, 20)
        Me.Label16.TabIndex = 14
        Me.ToolTip1.SetToolTip(Me.Label16, "Закрыть файл")
        '
        'Label14
        '
        Me.Label14.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Label14.Image = CType(resources.GetObject("Label14.Image"), System.Drawing.Image)
        Me.Label14.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label14.Location = New System.Drawing.Point(272, 64)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(113, 24)
        Me.Label14.TabIndex = 13
        Me.Label14.Text = "Открыть файл."
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.Label14, "Открыть внешний файл")
        '
        'Label13
        '
        Me.Label13.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Label13.Image = CType(resources.GetObject("Label13.Image"), System.Drawing.Image)
        Me.Label13.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label13.Location = New System.Drawing.Point(136, 64)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(126, 24)
        Me.Label13.TabIndex = 12
        Me.Label13.Text = "Сохранить в файл."
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.Label13, "Сохранить данные во внешний файл")
        '
        'Label12
        '
        Me.Label12.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Label12.Image = CType(resources.GetObject("Label12.Image"), System.Drawing.Image)
        Me.Label12.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Label12.Location = New System.Drawing.Point(8, 64)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(120, 24)
        Me.Label12.TabIndex = 11
        Me.Label12.Text = "Сохранить в базу."
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.Label12, "Сохранить данные в базу данных")
        '
        'ComboBox1
        '
        Me.ComboBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ComboBox1.Location = New System.Drawing.Point(120, 32)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(264, 21)
        Me.ComboBox1.TabIndex = 3
        '
        'TextBox1
        '
        Me.TextBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox1.Location = New System.Drawing.Point(80, 8)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(304, 20)
        Me.TextBox1.TabIndex = 2
        Me.TextBox1.Text = ""
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(8, 40)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(128, 23)
        Me.Label11.TabIndex = 1
        Me.Label11.Text = "Родительская папка:"
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(8, 16)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(100, 16)
        Me.Label10.TabIndex = 0
        Me.Label10.Text = "Имя файла:"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.BackColor = System.Drawing.Color.White
        Me.GroupBox1.Controls.Add(Me.RichTextBox1)
        Me.GroupBox1.Location = New System.Drawing.Point(0, 168)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(412, 240)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Содержание:"
        '
        'RichTextBox1
        '
        Me.RichTextBox1.AcceptsTab = True
        Me.RichTextBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RichTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.RichTextBox1.ContextMenu = Me.ContextMenu2
        Me.RichTextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.RichTextBox1.Location = New System.Drawing.Point(5, 16)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.ForcedBoth
        Me.RichTextBox1.Size = New System.Drawing.Size(404, 220)
        Me.RichTextBox1.TabIndex = 0
        Me.RichTextBox1.Text = ""
        Me.RichTextBox1.WordWrap = False
        '
        'ContextMenu2
        '
        Me.ContextMenu2.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem6, Me.MenuItem7, Me.MenuItem8, Me.MenuItem9, Me.MenuItem14, Me.MenuItem15, Me.MenuItem16})
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 0
        Me.MenuItem6.Text = "Вставить                         Ctrl+V"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 1
        Me.MenuItem7.Text = "Копировать                    Ctrl+C"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 2
        Me.MenuItem8.Text = "Вырезать                        Ctrl+X"
        '
        'MenuItem9
        '
        Me.MenuItem9.Index = 3
        Me.MenuItem9.Text = "Отменить                        Ctrl+Z"
        '
        'MenuItem14
        '
        Me.MenuItem14.Index = 4
        Me.MenuItem14.Text = "-"
        '
        'MenuItem15
        '
        Me.MenuItem15.Index = 5
        Me.MenuItem15.Text = "Цвет текста."
        '
        'MenuItem16
        '
        Me.MenuItem16.Index = 6
        Me.MenuItem16.Text = "Шрифт текста."
        '
        'OleDbSelectCommand1
        '
        Me.OleDbSelectCommand1.CommandText = "SELECT ДатаПоследнегоСохранения, ПапкаИдентификатор, СодержаниеФайла, Строка, Тип" & _
        "Объекта, ФайлВПапке, ФайлИдентификатор FROM Хранилище"
        Me.OleDbSelectCommand1.Connection = Me.OleDbConnection1
        '
        'OleDbConnection1
        '
        Me.OleDbConnection1.ConnectionString = "Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Registry Path=;Jet OLEDB:Database L" & _
        "ocking Mode=1;Jet OLEDB:Database Password=;Data Source=""C:\Программирование\VB N" & _
        "ET\CATFISH\Исходники\Cabinet\Base\Cabinet.mdb"";Password=;Jet OLEDB:Engine Type=5" & _
        ";Jet OLEDB:Global Bulk Transactions=1;Provider=""Microsoft.Jet.OLEDB.4.0"";Jet OLE" & _
        "DB:System database=;Jet OLEDB:SFP=False;Extended Properties=;Mode=ReadWrite;Jet " & _
        "OLEDB:New Database Password=;Jet OLEDB:Create System Database=False;Jet OLEDB:Do" & _
        "n't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;" & _
        "User ID=Admin;Jet OLEDB:Encrypt Database=False"
        '
        'OleDbInsertCommand1
        '
        Me.OleDbInsertCommand1.CommandText = "INSERT INTO Хранилище(ДатаПоследнегоСохранения, ПапкаИдентификатор, СодержаниеФай" & _
        "ла, ТипОбъекта, ФайлВПапке, ФайлИдентификатор) VALUES (?, ?, ?, ?, ?, ?)"
        Me.OleDbInsertCommand1.Connection = Me.OleDbConnection1
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ДатаПоследнегоСохранения", System.Data.OleDb.OleDbType.DBDate, 0, "ДатаПоследнегоСохранения"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ПапкаИдентификатор", System.Data.OleDb.OleDbType.VarWChar, 255, "ПапкаИдентификатор"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("СодержаниеФайла", System.Data.OleDb.OleDbType.VarWChar, 0, "СодержаниеФайла"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ТипОбъекта", System.Data.OleDb.OleDbType.VarWChar, 255, "ТипОбъекта"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ФайлВПапке", System.Data.OleDb.OleDbType.VarWChar, 255, "ФайлВПапке"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ФайлИдентификатор", System.Data.OleDb.OleDbType.VarWChar, 255, "ФайлИдентификатор"))
        '
        'OleDbUpdateCommand1
        '
        Me.OleDbUpdateCommand1.CommandText = "UPDATE Хранилище SET ДатаПоследнегоСохранения = ?, ПапкаИдентификатор = ?, Содерж" & _
        "аниеФайла = ?, ТипОбъекта = ?, ФайлВПапке = ?, ФайлИдентификатор = ? WHERE (Стро" & _
        "ка = ?) AND (ДатаПоследнегоСохранения = ? OR ? IS NULL AND ДатаПоследнегоСохране" & _
        "ния IS NULL) AND (ПапкаИдентификатор = ? OR ? IS NULL AND ПапкаИдентификатор IS " & _
        "NULL) AND (ТипОбъекта = ? OR ? IS NULL AND ТипОбъекта IS NULL) AND (ФайлВПапке =" & _
        " ? OR ? IS NULL AND ФайлВПапке IS NULL) AND (ФайлИдентификатор = ? OR ? IS NULL " & _
        "AND ФайлИдентификатор IS NULL)"
        Me.OleDbUpdateCommand1.Connection = Me.OleDbConnection1
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ДатаПоследнегоСохранения", System.Data.OleDb.OleDbType.DBDate, 0, "ДатаПоследнегоСохранения"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ПапкаИдентификатор", System.Data.OleDb.OleDbType.VarWChar, 255, "ПапкаИдентификатор"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("СодержаниеФайла", System.Data.OleDb.OleDbType.VarWChar, 0, "СодержаниеФайла"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ТипОбъекта", System.Data.OleDb.OleDbType.VarWChar, 255, "ТипОбъекта"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ФайлВПапке", System.Data.OleDb.OleDbType.VarWChar, 255, "ФайлВПапке"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ФайлИдентификатор", System.Data.OleDb.OleDbType.VarWChar, 255, "ФайлИдентификатор"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Строка", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Строка", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ДатаПоследнегоСохранения", System.Data.OleDb.OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ДатаПоследнегоСохранения", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ДатаПоследнегоСохранения1", System.Data.OleDb.OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ДатаПоследнегоСохранения", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ПапкаИдентификатор", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ПапкаИдентификатор", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ПапкаИдентификатор1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ПапкаИдентификатор", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ТипОбъекта", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ТипОбъекта", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ТипОбъекта1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ТипОбъекта", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФайлВПапке", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФайлВПапке", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФайлВПапке1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФайлВПапке", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФайлИдентификатор", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФайлИдентификатор", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФайлИдентификатор1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФайлИдентификатор", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbDeleteCommand1
        '
        Me.OleDbDeleteCommand1.CommandText = "DELETE FROM Хранилище WHERE (Строка = ?) AND (ДатаПоследнегоСохранения = ? OR ? I" & _
        "S NULL AND ДатаПоследнегоСохранения IS NULL) AND (ПапкаИдентификатор = ? OR ? IS" & _
        " NULL AND ПапкаИдентификатор IS NULL) AND (ТипОбъекта = ? OR ? IS NULL AND ТипОб" & _
        "ъекта IS NULL) AND (ФайлВПапке = ? OR ? IS NULL AND ФайлВПапке IS NULL) AND (Фай" & _
        "лИдентификатор = ? OR ? IS NULL AND ФайлИдентификатор IS NULL)"
        Me.OleDbDeleteCommand1.Connection = Me.OleDbConnection1
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Строка", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Строка", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ДатаПоследнегоСохранения", System.Data.OleDb.OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ДатаПоследнегоСохранения", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ДатаПоследнегоСохранения1", System.Data.OleDb.OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ДатаПоследнегоСохранения", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ПапкаИдентификатор", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ПапкаИдентификатор", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ПапкаИдентификатор1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ПапкаИдентификатор", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ТипОбъекта", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ТипОбъекта", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ТипОбъекта1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ТипОбъекта", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФайлВПапке", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФайлВПапке", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФайлВПапке1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФайлВПапке", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФайлИдентификатор", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФайлИдентификатор", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФайлИдентификатор1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФайлИдентификатор", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbDataAdapter1
        '
        Me.OleDbDataAdapter1.DeleteCommand = Me.OleDbDeleteCommand1
        Me.OleDbDataAdapter1.InsertCommand = Me.OleDbInsertCommand1
        Me.OleDbDataAdapter1.SelectCommand = Me.OleDbSelectCommand1
        Me.OleDbDataAdapter1.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "Хранилище", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("ДатаПоследнегоСохранения", "ДатаПоследнегоСохранения"), New System.Data.Common.DataColumnMapping("ПапкаИдентификатор", "ПапкаИдентификатор"), New System.Data.Common.DataColumnMapping("СодержаниеФайла", "СодержаниеФайла"), New System.Data.Common.DataColumnMapping("Строка", "Строка"), New System.Data.Common.DataColumnMapping("ТипОбъекта", "ТипОбъекта"), New System.Data.Common.DataColumnMapping("ФайлВПапке", "ФайлВПапке"), New System.Data.Common.DataColumnMapping("ФайлИдентификатор", "ФайлИдентификатор")})})
        Me.OleDbDataAdapter1.UpdateCommand = Me.OleDbUpdateCommand1
        '
        'DataSet21
        '
        Me.DataSet21.DataSetName = "DataSet2"
        Me.DataSet21.Locale = New System.Globalization.CultureInfo("ru-RU")
        '
        'DataSet31
        '
        Me.DataSet31.DataSetName = "DataSet3"
        Me.DataSet31.Locale = New System.Globalization.CultureInfo("ru-RU")
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.Filter = "*.txt|*.txt|*.rtf|*.rtf|*.*|*.*"
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.Filter = "*.txt|*.txt|*.rtf|*.rtf|*.*|*.*"
        '
        'OleDbSelectCommand2
        '
        Me.OleDbSelectCommand2.CommandText = "SELECT ДатаПоследнегоСохранения, ПапкаИдентификатор, СодержаниеФайла, Строка, Тип" & _
        "Объекта, ФайлВПапке, ФайлИдентификатор FROM Хранилище"
        Me.OleDbSelectCommand2.Connection = Me.OleDbConnection1
        '
        'OleDbInsertCommand2
        '
        Me.OleDbInsertCommand2.CommandText = "INSERT INTO Хранилище(ДатаПоследнегоСохранения, ПапкаИдентификатор, СодержаниеФай" & _
        "ла, ТипОбъекта, ФайлВПапке, ФайлИдентификатор) VALUES (?, ?, ?, ?, ?, ?)"
        Me.OleDbInsertCommand2.Connection = Me.OleDbConnection1
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("ДатаПоследнегоСохранения", System.Data.OleDb.OleDbType.DBDate, 0, "ДатаПоследнегоСохранения"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("ПапкаИдентификатор", System.Data.OleDb.OleDbType.VarWChar, 255, "ПапкаИдентификатор"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("СодержаниеФайла", System.Data.OleDb.OleDbType.VarWChar, 0, "СодержаниеФайла"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("ТипОбъекта", System.Data.OleDb.OleDbType.VarWChar, 255, "ТипОбъекта"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("ФайлВПапке", System.Data.OleDb.OleDbType.VarWChar, 255, "ФайлВПапке"))
        Me.OleDbInsertCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("ФайлИдентификатор", System.Data.OleDb.OleDbType.VarWChar, 255, "ФайлИдентификатор"))
        '
        'OleDbUpdateCommand2
        '
        Me.OleDbUpdateCommand2.CommandText = "UPDATE Хранилище SET ДатаПоследнегоСохранения = ?, ПапкаИдентификатор = ?, Содерж" & _
        "аниеФайла = ?, ТипОбъекта = ?, ФайлВПапке = ?, ФайлИдентификатор = ? WHERE (Стро" & _
        "ка = ?) AND (ДатаПоследнегоСохранения = ? OR ? IS NULL AND ДатаПоследнегоСохране" & _
        "ния IS NULL) AND (ПапкаИдентификатор = ? OR ? IS NULL AND ПапкаИдентификатор IS " & _
        "NULL) AND (ТипОбъекта = ? OR ? IS NULL AND ТипОбъекта IS NULL) AND (ФайлВПапке =" & _
        " ? OR ? IS NULL AND ФайлВПапке IS NULL) AND (ФайлИдентификатор = ? OR ? IS NULL " & _
        "AND ФайлИдентификатор IS NULL)"
        Me.OleDbUpdateCommand2.Connection = Me.OleDbConnection1
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("ДатаПоследнегоСохранения", System.Data.OleDb.OleDbType.DBDate, 0, "ДатаПоследнегоСохранения"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("ПапкаИдентификатор", System.Data.OleDb.OleDbType.VarWChar, 255, "ПапкаИдентификатор"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("СодержаниеФайла", System.Data.OleDb.OleDbType.VarWChar, 0, "СодержаниеФайла"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("ТипОбъекта", System.Data.OleDb.OleDbType.VarWChar, 255, "ТипОбъекта"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("ФайлВПапке", System.Data.OleDb.OleDbType.VarWChar, 255, "ФайлВПапке"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("ФайлИдентификатор", System.Data.OleDb.OleDbType.VarWChar, 255, "ФайлИдентификатор"))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Строка", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Строка", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ДатаПоследнегоСохранения", System.Data.OleDb.OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ДатаПоследнегоСохранения", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ДатаПоследнегоСохранения1", System.Data.OleDb.OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ДатаПоследнегоСохранения", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ПапкаИдентификатор", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ПапкаИдентификатор", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ПапкаИдентификатор1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ПапкаИдентификатор", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ТипОбъекта", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ТипОбъекта", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ТипОбъекта1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ТипОбъекта", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФайлВПапке", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФайлВПапке", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФайлВПапке1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФайлВПапке", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФайлИдентификатор", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФайлИдентификатор", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФайлИдентификатор1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФайлИдентификатор", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbDeleteCommand2
        '
        Me.OleDbDeleteCommand2.CommandText = "DELETE FROM Хранилище WHERE (Строка = ?) AND (ДатаПоследнегоСохранения = ? OR ? I" & _
        "S NULL AND ДатаПоследнегоСохранения IS NULL) AND (ПапкаИдентификатор = ? OR ? IS" & _
        " NULL AND ПапкаИдентификатор IS NULL) AND (ТипОбъекта = ? OR ? IS NULL AND ТипОб" & _
        "ъекта IS NULL) AND (ФайлВПапке = ? OR ? IS NULL AND ФайлВПапке IS NULL) AND (Фай" & _
        "лИдентификатор = ? OR ? IS NULL AND ФайлИдентификатор IS NULL)"
        Me.OleDbDeleteCommand2.Connection = Me.OleDbConnection1
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Строка", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Строка", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ДатаПоследнегоСохранения", System.Data.OleDb.OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ДатаПоследнегоСохранения", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ДатаПоследнегоСохранения1", System.Data.OleDb.OleDbType.DBDate, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ДатаПоследнегоСохранения", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ПапкаИдентификатор", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ПапкаИдентификатор", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ПапкаИдентификатор1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ПапкаИдентификатор", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ТипОбъекта", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ТипОбъекта", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ТипОбъекта1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ТипОбъекта", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФайлВПапке", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФайлВПапке", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФайлВПапке1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФайлВПапке", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФайлИдентификатор", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФайлИдентификатор", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand2.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФайлИдентификатор1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФайлИдентификатор", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbDataAdapter2
        '
        Me.OleDbDataAdapter2.DeleteCommand = Me.OleDbDeleteCommand2
        Me.OleDbDataAdapter2.InsertCommand = Me.OleDbInsertCommand2
        Me.OleDbDataAdapter2.SelectCommand = Me.OleDbSelectCommand2
        Me.OleDbDataAdapter2.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "Хранилище", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("ДатаПоследнегоСохранения", "ДатаПоследнегоСохранения"), New System.Data.Common.DataColumnMapping("ПапкаИдентификатор", "ПапкаИдентификатор"), New System.Data.Common.DataColumnMapping("СодержаниеФайла", "СодержаниеФайла"), New System.Data.Common.DataColumnMapping("Строка", "Строка"), New System.Data.Common.DataColumnMapping("ТипОбъекта", "ТипОбъекта"), New System.Data.Common.DataColumnMapping("ФайлВПапке", "ФайлВПапке"), New System.Data.Common.DataColumnMapping("ФайлИдентификатор", "ФайлИдентификатор")})})
        Me.OleDbDataAdapter2.UpdateCommand = Me.OleDbUpdateCommand2
        '
        'DataSet51
        '
        Me.DataSet51.DataSetName = "DataSet5"
        Me.DataSet51.Locale = New System.Globalization.CultureInfo("ru-RU")
        '
        'MyCabinet
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(592, 422)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Splitter1)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.StatusBar1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.Name = "MyCabinet"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Cabinet"
        Me.Panel1.ResumeLayout(False)
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StatusBarPanel3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StatusBarPanel2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.DataSet21, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataSet31, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataSet51, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Dim NewOrEdit As String
    Private Sub Cabinet_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            RichTextBox1.Font = ModuleCabinet.FontModule
            If (ModuleCabinet.ActionTopMost = True) Then Me.TopMost = True
            TREE_UPDATE()
            NewOrEdit = "NEW"
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Public Sub TREE_UPDATE()
        Try
            ComboBox1.Items.Clear()
            OleDbConnection1.ConnectionString = ModuleBase.StringConnection
            OleDbConnection1.Open()

            DataSet21.Clear()
            DataSet31.Clear()

            Dim CommandSQL As String

            If (MenuItem12.Checked = True) Then CommandSQL = "SELECT ДатаПоследнегоСохранения, ПапкаИдентификатор, СодержаниеФайла, Строка, ТипОбъекта, ФайлВПапке, ФайлИдентификатор FROM Хранилище ORDER BY ПапкаИдентификатор ASC"
            If (MenuItem13.Checked = True) Then CommandSQL = "SELECT ДатаПоследнегоСохранения, ПапкаИдентификатор, СодержаниеФайла, Строка, ТипОбъекта, ФайлВПапке, ФайлИдентификатор FROM Хранилище ORDER BY ПапкаИдентификатор DESC"

            OleDbDataAdapter1.SelectCommand.CommandText = CommandSQL
            OleDbDataAdapter1.Fill(DataSet21, "Хранилище")
            OleDbDataAdapter1.SelectCommand.CommandText = "SELECT ДатаПоследнегоСохранения, ПапкаИдентификатор, СодержаниеФайла, Строка, ТипОбъекта, ФайлВПапке, ФайлИдентификатор FROM Хранилище ORDER BY Строка ASC"
            OleDbDataAdapter1.Fill(DataSet31, "Хранилище")



            'загрузка дерева
            Dim NameGr As String
            Dim ActionGr, ActionEl As Integer
            Dim NextGr As Boolean = False
            TreeView1.Nodes.Clear()
            For i As Integer = 0 To DataSet21.Хранилище.Rows.Count - 1
                If (DataSet21.Хранилище.Item(i)("ТипОбъекта") = "Группа") Then
                    NameGr = DataSet21.Хранилище.Item(i)("ПапкаИдентификатор")
                    TreeView1.Nodes.Add(NameGr)
                    TreeView1.Nodes.Item(ActionGr).ImageIndex = 0
                    TreeView1.Nodes.Item(ActionGr).SelectedImageIndex = 1
                    ActionGr = ActionGr + 1
                    NextGr = True
                    ComboBox1.Items.Add(NameGr)
                End If

                If (NextGr = True) Then
                    ActionEl = 0
                    For iEl As Integer = 0 To DataSet31.Хранилище.Rows.Count - 1
                        If (DataSet31.Хранилище.Item(iEl)("ТипОбъекта") = "Элемент") Then
                            If (DataSet31.Хранилище.Item(iEl)("ФайлВПапке") = NameGr) And (DataSet31.Хранилище.Item(iEl)("ТипОбъекта") = "Элемент") Then
                                TreeView1.Nodes.Item(ActionGr - 1).Nodes.Add(DataSet31.Хранилище.Item(iEl)("ФайлИдентификатор"))
                                TreeView1.Nodes.Item(ActionGr - 1).Nodes.Item(ActionEl).ImageIndex = 2
                                TreeView1.Nodes.Item(ActionGr - 1).Nodes.Item(ActionEl).SelectedImageIndex = 2
                                ActionEl = ActionEl + 1
                            End If
                        End If
                    Next
                    NextGr = False
                End If
            Next
            OleDbConnection1.Close()
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub
    Private Sub Cabinet_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        OleDbConnection1.Close()
    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click
        Try
            Dim NewFolder As String
            Dim PoliticNewFolder As Boolean = False
            NewFolder = InputBox("Введите имя новой папки.", "Новая папка", "")
            If NewFolder <> "" Then
                'Проверка уникальности имени папки
                For i As Integer = 0 To DataSet21.Хранилище.Count - 1
                    If (DataSet21.Хранилище.Item(i)("ТипОбъекта") = "Группа") Then
                        If (DataSet21.Хранилище.Item(i)("ПапкаИдентификатор") = NewFolder) Then
                            PoliticNewFolder = True
                            MsgBox("Имя папки не уникально!!! Введите другое имя папки.", MsgBoxStyle.OKOnly, "Сообщение:")
                        End If
                    End If
                Next
                If (PoliticNewFolder = False) Then
                    OleDbConnection1.ConnectionString = ModuleBase.StringConnection
                    OleDbConnection1.Open()
                    Dim DT As DataTable = DataSet21.Хранилище
                    Dim row As DataRow
                    row = DT.NewRow
                    row("ТипОбъекта") = "Группа"
                    row("ПапкаИдентификатор") = NewFolder.ToString
                    row("СодержаниеФайла") = ""
                    row("ДатаПоследнегоСохранения") = Date.Today
                    row("ФайлИдентификатор") = ""
                    row("ФайлВПапке") = ""
                    DT.Rows.Add(row)
                    OleDbDataAdapter1.Update(DataSet21, "Хранилище")
                    OleDbConnection1.Close()
                    TREE_UPDATE()
                End If
            End If
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label2_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label2.MouseMove
        Try
            Label2.BorderStyle = BorderStyle.FixedSingle
            Label3.BorderStyle = BorderStyle.None
            Label4.BorderStyle = BorderStyle.None
            Label5.BorderStyle = BorderStyle.None
            Label6.BorderStyle = BorderStyle.None
            Label7.BorderStyle = BorderStyle.None
            Label8.BorderStyle = BorderStyle.None
            Label9.BorderStyle = BorderStyle.None
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label2_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label2.MouseLeave
        Try
            Label2.BorderStyle = BorderStyle.None
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click
        'Изменить
        Try
            For iTest As Integer = 0 To DataSet21.Хранилище.Count - 1
                If (DataSet21.Хранилище.Item(iTest)("ТипОбъекта") = "Группа") Then
                    If (DataSet21.Хранилище.Item(iTest)("ПапкаИдентификатор") = StatusBarPanel2.Text) Then
                        If (DataSet21.Хранилище.Item(iTest)("ТипОбъекта") = "Группа") Then
                            Dim NewFolder As String
                            Dim EditFolder As String
                            Dim PoliticNewFolder As Boolean = False
                            Dim FindEditFolder As Boolean = False
                            EditFolder = TreeView1.SelectedNode.Text
                            NewFolder = InputBox("Введите новое имя папки.", "Редактировать папку : [" + EditFolder + "]", EditFolder)
                            If NewFolder <> "" Then
                                'Проверка уникальности имени папки
                                For i As Integer = 0 To DataSet21.Хранилище.Count - 1
                                    If (DataSet21.Хранилище.Item(i)("ТипОбъекта") = "Группа") Then
                                        If (DataSet21.Хранилище.Item(i)("ПапкаИдентификатор") = NewFolder) Then
                                            PoliticNewFolder = True
                                            MsgBox("Имя папки не уникально!!! Введите другое имя папки.", MsgBoxStyle.OKOnly, "Сообщение:")
                                        End If
                                    End If
                                    If (DataSet21.Хранилище.Item(i)("ТипОбъекта") = "Группа") Then
                                        If (DataSet21.Хранилище.Item(i)("ПапкаИдентификатор") = EditFolder) Then
                                            FindEditFolder = True
                                        End If
                                    End If
                                Next
                                If (PoliticNewFolder = False) Then
                                    If (FindEditFolder = True) Then
                                        OleDbConnection1.Open()
                                        For iEdit As Integer = 0 To DataSet21.Хранилище.Count - 1
                                            If (DataSet21.Хранилище.Item(iEdit)("ТипОбъекта") = "Группа") Then
                                                If (DataSet21.Хранилище.Item(iEdit)("ТипОбъекта") = "Группа") And (DataSet21.Хранилище.Item(iEdit)("ПапкаИдентификатор") = EditFolder) Then
                                                    DataSet21.Хранилище.Item(iEdit)("ПапкаИдентификатор") = NewFolder
                                                End If
                                            End If
                                            If (DataSet21.Хранилище.Item(iEdit)("ТипОбъекта") = "Элемент") Then
                                                If (DataSet21.Хранилище.Item(iEdit)("ТипОбъекта") = "Элемент") And (DataSet21.Хранилище.Item(iEdit)("ФайлВПапке") = EditFolder) Then
                                                    DataSet21.Хранилище.Item(iEdit)("ФайлВПапке") = NewFolder
                                                End If
                                            End If
                                        Next
                                        OleDbDataAdapter1.Update(DataSet21, "Хранилище")
                                        OleDbConnection1.Close()
                                        TREE_UPDATE()
                                        MsgBox("Папка успешно переименованна.", MsgBoxStyle.OKOnly, "Сообщение:")
                                    Else
                                        MsgBox("Ошибка!!! " + "[" + EditFolder + "] не является папкой!", MsgBoxStyle.OKOnly, "Сообщение:")
                                    End If
                                End If
                            End If
                            Exit For
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label3_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label3.MouseMove
        Try
            Label3.BorderStyle = BorderStyle.FixedSingle
            Label2.BorderStyle = BorderStyle.None
            Label4.BorderStyle = BorderStyle.None
            Label5.BorderStyle = BorderStyle.None
            Label6.BorderStyle = BorderStyle.None
            Label7.BorderStyle = BorderStyle.None
            Label8.BorderStyle = BorderStyle.None
            Label9.BorderStyle = BorderStyle.None
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label3_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label3.MouseLeave
        Try
            Label3.BorderStyle = BorderStyle.None
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.Click
        'Удаление
        Try
            Dim DelFolder As String
            Dim FindEditFolder As Boolean = False
            DelFolder = TreeView1.SelectedNode.Text
            If (MsgBox("Вы увирены что хотите удалить папку [" + DelFolder + "] и её содержимое?", MsgBoxStyle.YesNo, "Вопрос: ") = MsgBoxResult.Yes) Then
                For i As Integer = 0 To DataSet21.Хранилище.Count - 1
                    If (DataSet21.Хранилище.Item(i)("ТипОбъекта") = "Группа") Then
                        If (DataSet21.Хранилище.Item(i)("ТипОбъекта") = "Группа") And (DataSet21.Хранилище.Item(i)("ПапкаИдентификатор") = DelFolder) Then
                            FindEditFolder = True
                        End If
                    End If
                Next
                If (FindEditFolder = True) Then
                    OleDbConnection1.Open()
                    For iDelEl As Integer = 0 To DataSet21.Хранилище.Count - 1
                        If (DataSet21.Хранилище.Item(iDelEl)("ТипОбъекта") = "Элемент") Then
                            If (DataSet21.Хранилище.Item(iDelEl)("ТипОбъекта") = "Элемент") And (DataSet21.Хранилище.Item(iDelEl)("ФайлВПапке") = DelFolder) Then
                                DataSet21.Хранилище.Item(iDelEl).Delete()
                            End If
                        End If
                    Next
                    For iDelFolder As Integer = 0 To DataSet31.Хранилище.Count - 1
                        If (DataSet31.Хранилище.Item(iDelFolder)("ТипОбъекта") = "Группа") Then
                            If (DataSet31.Хранилище.Item(iDelFolder)("ТипОбъекта") = "Группа") And (DataSet31.Хранилище.Item(iDelFolder)("ПапкаИдентификатор") = DelFolder) Then
                                DataSet31.Хранилище.Item(iDelFolder).Delete()
                            End If
                        End If
                    Next
                    OleDbDataAdapter1.Update(DataSet21, "Хранилище")
                    OleDbDataAdapter1.Update(DataSet31, "Хранилище")
                    OleDbConnection1.Close()
                    TREE_UPDATE()
                    MsgBox("Удаление прошло успешно.", MsgBoxStyle.OKOnly, "Сообщение:")
                Else
                    MsgBox("Ошибка!!! " + "[" + DelFolder + "] не является папкой!", MsgBoxStyle.OKOnly, "Сообщение:")
                End If
            End If
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label4_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label4.MouseMove
        Try
            Label4.BorderStyle = BorderStyle.FixedSingle
            Label2.BorderStyle = BorderStyle.None
            Label3.BorderStyle = BorderStyle.None
            Label5.BorderStyle = BorderStyle.None
            Label6.BorderStyle = BorderStyle.None
            Label7.BorderStyle = BorderStyle.None
            Label8.BorderStyle = BorderStyle.None
            Label9.BorderStyle = BorderStyle.None
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label4_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label4.MouseLeave
        Try
            Label4.BorderStyle = BorderStyle.None
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click
        NewOrEdit = "NEW"
        TextBox1.Text = ""
        ComboBox1.Text = StatusBarPanel2.Text
        RichTextBox1.Clear()
        DataSet51.Clear()
    End Sub

    Private Sub Label5_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label5.MouseMove
        Try
            Label5.BorderStyle = BorderStyle.FixedSingle
            Label2.BorderStyle = BorderStyle.None
            Label3.BorderStyle = BorderStyle.None
            Label4.BorderStyle = BorderStyle.None
            Label6.BorderStyle = BorderStyle.None
            Label7.BorderStyle = BorderStyle.None
            Label8.BorderStyle = BorderStyle.None
            Label9.BorderStyle = BorderStyle.None
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label5_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label5.MouseLeave
        Try
            Label5.BorderStyle = BorderStyle.None
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Dim ActionID As Integer
    Private Sub Label6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label6.Click
        Try
            NewOrEdit = "EDIT"
            TextBox1.Text = ""
            ComboBox1.Text = ""
            RichTextBox1.Clear()
            DataSet51.Clear()
            Dim ContinueEdit As Boolean = False
            For iFindFile As Integer = 0 To DataSet21.Хранилище.Count - 1
                If (DataSet21.Хранилище.Item(iFindFile)("ТипОбъекта") = "Элемент") Then
                    If (DataSet21.Хранилище.Item(iFindFile)("ФайлИдентификатор") = StatusBarPanel2.Text) Then
                        ContinueEdit = True
                        Exit For
                    End If
                End If
            Next
            If (ContinueEdit = True) Then
                OleDbConnection1.Open()
                Dim SQLComand As String
                SQLComand = "SELECT ДатаПоследнегоСохранения, ПапкаИдентификатор, СодержаниеФайла, Строка, ТипОбъекта, ФайлВПапке, ФайлИдентификатор FROM Хранилище WHERE (ФайлИдентификатор = '" + StatusBarPanel2.Text + "')"
                OleDbDataAdapter2.SelectCommand.CommandText = SQLComand
                OleDbDataAdapter2.Fill(DataSet51, "Хранилище")
                TextBox1.Text = DataSet51.Хранилище.Item(0)("ФайлИдентификатор")
                ComboBox1.Text = DataSet51.Хранилище.Item(0)("ФайлВПапке")
                RichTextBox1.Rtf = DataSet51.Хранилище.Item(0)("СодержаниеФайла")
                StatusBarPanel3.Text = DataSet51.Хранилище.Item(0)("ДатаПоследнегоСохранения")
                ActionID = DataSet51.Хранилище.Item(0)("Строка")
                OleDbConnection1.Close()
            End If
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label6_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label6.MouseMove
        Try
            Label6.BorderStyle = BorderStyle.FixedSingle
            Label2.BorderStyle = BorderStyle.None
            Label3.BorderStyle = BorderStyle.None
            Label4.BorderStyle = BorderStyle.None
            Label5.BorderStyle = BorderStyle.None
            Label7.BorderStyle = BorderStyle.None
            Label8.BorderStyle = BorderStyle.None
            Label9.BorderStyle = BorderStyle.None
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label6_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label6.MouseLeave
        Try
            Label6.BorderStyle = BorderStyle.None
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label7.Click
        'Удалить модуль
        Try
            Dim DelModule As Boolean = False
            If (MsgBox("Вы увирены что хотите удалить файл [" + StatusBarPanel2.Text + "]?", MsgBoxStyle.YesNo, "Вопрос: ") = MsgBoxResult.Yes) Then
                For i As Integer = 0 To DataSet21.Хранилище.Count - 1
                    If (DataSet21.Хранилище.Item(i)("ТипОбъекта") = "Элемент") Then
                        If (DataSet21.Хранилище.Item(i)("ТипОбъекта") = "Элемент") And (DataSet21.Хранилище.Item(i)("ФайлИдентификатор") = StatusBarPanel2.Text) Then
                            OleDbConnection1.Open()
                            DataSet21.Хранилище.Item(i).Delete()
                            OleDbDataAdapter1.Update(DataSet21, "Хранилище")
                            OleDbConnection1.Close()
                            DelModule = True
                            Exit For
                        End If
                    End If
                Next
                If (DelModule = True) Then
                    TREE_UPDATE()
                    NewOrEdit = "NEW"
                    TextBox1.Text = ""
                    ComboBox1.Text = ""
                    RichTextBox1.Clear()
                    DataSet51.Clear()
                    MsgBox("Удаление прошло успешно.", MsgBoxStyle.OKOnly, "Сообщение:")
                Else
                    MsgBox("Ошибка!!! " + "[" + StatusBarPanel2.Text + "] не является файлом!", MsgBoxStyle.OKOnly, "Сообщение:")
                End If
            End If
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label7_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label7.MouseMove
        Try
            Label7.BorderStyle = BorderStyle.FixedSingle
            Label2.BorderStyle = BorderStyle.None
            Label3.BorderStyle = BorderStyle.None
            Label4.BorderStyle = BorderStyle.None
            Label5.BorderStyle = BorderStyle.None
            Label6.BorderStyle = BorderStyle.None
            Label8.BorderStyle = BorderStyle.None
            Label9.BorderStyle = BorderStyle.None
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label7_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label7.MouseLeave
        Try
            Label7.BorderStyle = BorderStyle.None
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label8.Click
        Try
            If (ModuleCabinet.SettingsShow = False) Then
                ModuleCabinet.LoadSettingsProg()
                ModuleCabinet.SettingsProg.Show()
            End If
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label8_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label8.MouseMove
        Try
            Label8.BorderStyle = BorderStyle.FixedSingle
            Label2.BorderStyle = BorderStyle.None
            Label3.BorderStyle = BorderStyle.None
            Label4.BorderStyle = BorderStyle.None
            Label5.BorderStyle = BorderStyle.None
            Label6.BorderStyle = BorderStyle.None
            Label7.BorderStyle = BorderStyle.None
            Label9.BorderStyle = BorderStyle.None
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label8_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label8.MouseLeave
        Try
            Label8.BorderStyle = BorderStyle.None
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label9.Click
        Try
            If (ModuleCabinet.AboutShow = False) Then
                Dim Ab As New About
                Ab.Show()
            End If
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label9_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label9.MouseMove
        Try
            Label9.BorderStyle = BorderStyle.FixedSingle
            Label2.BorderStyle = BorderStyle.None
            Label3.BorderStyle = BorderStyle.None
            Label4.BorderStyle = BorderStyle.None
            Label5.BorderStyle = BorderStyle.None
            Label6.BorderStyle = BorderStyle.None
            Label7.BorderStyle = BorderStyle.None
            Label8.BorderStyle = BorderStyle.None
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label9_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label9.MouseLeave
        Try
            Label9.BorderStyle = BorderStyle.None
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub TreeView1_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterSelect
        Try
            StatusBarPanel2.Text = TreeView1.SelectedNode.Text
            For i As Integer = 0 To DataSet21.Хранилище.Count - 1
                If (DataSet21.Хранилище.Item(i)("ТипОбъекта") = "Группа") Then
                    If (DataSet21.Хранилище.Item(i)("ПапкаИдентификатор") = StatusBarPanel2.Text) Then
                        StatusBarPanel3.Text = "Дата: " & DataSet21.Хранилище.Item(i)("ДатаПоследнегоСохранения")
                        Exit For
                    End If
                End If
                If (DataSet21.Хранилище.Item(i)("ТипОбъекта") = "Элемент") Then
                    If (DataSet21.Хранилище.Item(i)("ФайлИдентификатор") = StatusBarPanel2.Text) Then
                        StatusBarPanel3.Text = "Дата: " & DataSet21.Хранилище.Item(i)("ДатаПоследнегоСохранения")
                        Exit For
                    End If
                End If
            Next
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label13.Click
        If (SaveFileDialog1.ShowDialog = DialogResult.OK) Then

            Dim FileNameCreate As String
            Dim iIndex, iStartDel, iEndDel As Integer
            iIndex = SaveFileDialog1.FileName.Length - 1
            For iFN As Integer = 0 To SaveFileDialog1.FileName.Length - 1
                If (SaveFileDialog1.FileName.Chars(iIndex) = ".") Then
                    iStartDel = 0
                    iEndDel = SaveFileDialog1.FileName.Length - iFN
                    FileNameCreate = SaveFileDialog1.FileName.Remove(iStartDel, iEndDel)
                    Exit For
                Else
                    iIndex = iIndex - 1
                End If
            Next
            If (FileNameCreate <> "rtf") Then
                RichTextBox1.SaveFile(SaveFileDialog1.FileName, RichTextBoxStreamType.PlainText)
            Else
                RichTextBox1.SaveFile(SaveFileDialog1.FileName, RichTextBoxStreamType.RichText)
            End If
            MsgBox("Сохранение прошло успешно.", MsgBoxStyle.OKOnly, "Сообщение:")
        End If
    End Sub

    Private Sub Label14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label14.Click
        If (OpenFileDialog1.ShowDialog = DialogResult.OK) Then
            Dim FileNameCreate As String
            Dim iIndex, iStartDel, iEndDel As Integer
            iIndex = OpenFileDialog1.FileName.Length - 1
            For iFN As Integer = 0 To OpenFileDialog1.FileName.Length - 1
                If (OpenFileDialog1.FileName.Chars(iIndex) = ".") Then
                    iStartDel = 0
                    iEndDel = OpenFileDialog1.FileName.Length - iFN
                    FileNameCreate = OpenFileDialog1.FileName.Remove(iStartDel, iEndDel)
                    Exit For
                Else
                    iIndex = iIndex - 1
                End If
            Next

            If (FileNameCreate <> "rtf") Then
                RichTextBox1.LoadFile(OpenFileDialog1.FileName, RichTextBoxStreamType.PlainText)
            Else
                RichTextBox1.LoadFile(OpenFileDialog1.FileName, RichTextBoxStreamType.RichText)
            End If

        End If
    End Sub

    Dim FindIndex As Integer = 0
    Dim FindText As String
    Dim FindLast As Integer = 0
    Private Sub Label15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label15.Click
        'поиск
        Try
            If (FindText <> TextBox2.Text) Then
                FindIndex = 0
                FindLast = 0
                FindText = TextBox2.Text
            End If
            If (RichTextBox1.Find(TextBox2.Text, FindIndex, RichTextBox1.TextLength - 1, RichTextBoxFinds.None)) Then
                RichTextBox1.Select()
                FindIndex = RichTextBox1.SelectionStart + RichTextBox1.SelectionLength
                If (FindLast = RichTextBox1.SelectionStart) Then
                    MsgBox("Поиск завершен.", MsgBoxStyle.OKOnly, "Сообщение:")
                    FindIndex = 0
                    FindLast = 0
                    FindText = TextBox2.Text
                Else
                    FindLast = RichTextBox1.SelectionStart
                End If
            End If
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label12.Click
        If (NewOrEdit = "NEW") Then
            SaveNewDATA()
        End If
        If (NewOrEdit = "EDIT") Then
            SaveEditDATA()
        End If
    End Sub

    Private Sub SaveNewDATA()
        Dim FolderFind As Boolean = False
        Dim NameModul As String
        Dim ModulePolitic As Boolean = True
        For iFolder As Integer = 0 To ComboBox1.Items.Count - 1
            If (ComboBox1.Text = ComboBox1.Items.Item(iFolder)) Then FolderFind = True
        Next
        If (FolderFind = True) Then
            If (TextBox1.Text = "") Then
                NameModul = InputBox("Введите имя файла.", "Новый файл:", "")
            Else
                NameModul = TextBox1.Text
            End If
            If (NameModul <> "") Then
                For iFindFile As Integer = 0 To DataSet21.Хранилище.Count - 1
                    If (DataSet21.Хранилище.Item(iFindFile)("ТипОбъекта") = "Элемент") Then
                        If (DataSet21.Хранилище.Item(iFindFile)("ФайлИдентификатор") = NameModul) Then
                            ModulePolitic = False
                            MsgBox("Ошибка: модуль с таким именем уже существует!!!.", MsgBoxStyle.OKOnly, "Сообщение:")
                            Label12.BorderStyle = BorderStyle.None
                        End If
                    End If
                Next
            Else
                ModulePolitic = False
                Label12.BorderStyle = BorderStyle.None
            End If
            'Сохраняем данные
            If (ModulePolitic = True) Then
                OleDbConnection1.Open()
                Dim DT As DataTable = DataSet21.Хранилище
                Dim row As DataRow
                row = DT.NewRow
                row("ТипОбъекта") = "Элемент"
                If (ComboBox1.Text <> "") Then
                    row("ФайлВПапке") = ComboBox1.Text
                Else
                    MsgBox("Вы не выбрали папку в которую нужно сохранить файл.", MsgBoxStyle.OKOnly, "Сообщение:")
                    Exit Sub
                End If
                row("ФайлИдентификатор") = NameModul
                row("СодержаниеФайла") = RichTextBox1.Rtf
                row("ДатаПоследнегоСохранения") = Date.Today.ToString
                DT.Rows.Add(row)
                OleDbDataAdapter1.Update(DataSet21, "Хранилище")
                OleDbConnection1.Close()
                TREE_UPDATE()
                MsgBox("Сохранение прошло успешно." + System.Environment.NewLine + "Файл будет автоматически закрыт.", MsgBoxStyle.OKOnly, "Сообщение:")
                NewOrEdit = "NEW"
                TextBox1.Text = ""
                ComboBox1.Text = ""
                RichTextBox1.Clear()
            End If
        Else
            MsgBox("Папка не найдена! Проверте имя папки в которую нужно сохранить модуль.", MsgBoxStyle.OKOnly, "Сообщение:")
            Label12.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub SaveEditDATA()
        Dim FolderFind As Boolean = False
        Dim NameModul As String
        Dim ModulePolitic As Boolean = True
        For iFolder As Integer = 0 To ComboBox1.Items.Count - 1
            If (ComboBox1.Text = ComboBox1.Items.Item(iFolder)) Then FolderFind = True
        Next
        If (FolderFind = True) Then
            If (TextBox1.Text = "") Then
                NameModul = InputBox("Введите имя Файла.", "Новые имя файла:", "")
            Else
                NameModul = TextBox1.Text
            End If
            If (NameModul <> "") Then 'проверка уникальности имени модуля
                For i As Integer = 0 To DataSet21.Хранилище.Count - 1
                    If (DataSet21.Хранилище.Item(i)("ТипОбъекта") = "Элемент") Then
                        If (DataSet21.Хранилище.Item(i)("ФайлИдентификатор") = NameModul) And (DataSet21.Хранилище.Item(i)("Строка") <> ActionID) Then
                            ModulePolitic = False
                            MsgBox("Ошибка: модуль с таким именем уже существует!!!.", MsgBoxStyle.OKOnly, "Сообщение:")
                            Label12.BorderStyle = BorderStyle.None
                        End If
                    End If
                Next
            Else
                ModulePolitic = False
                Label12.BorderStyle = BorderStyle.None
            End If
            'Сохраняем данные
            If (ModulePolitic = True) Then
                OleDbConnection1.Open()

                DataSet51.Хранилище.Item(0)("СодержаниеФайла") = RichTextBox1.Rtf
                DataSet51.Хранилище.Item(0)("ФайлИдентификатор") = NameModul
                If (ComboBox1.Text <> "") Then
                    DataSet51.Хранилище.Item(0)("ФайлВПапке") = ComboBox1.Text
                Else
                    MsgBox("Вы не выбрали папку в которую нужно сохранить файл.", MsgBoxStyle.OKOnly, "Сообщение:")
                    Exit Sub
                End If
                DataSet51.Хранилище.Item(0)("ДатаПоследнегоСохранения") = Date.Today.ToString

                OleDbDataAdapter2.Update(DataSet51, "Хранилище")
                OleDbConnection1.Close()
                TREE_UPDATE()
                MsgBox("Сохранение прошло успешно.", MsgBoxStyle.OKOnly, "Сообщение:")
            End If

        Else
            MsgBox("Папка не найдена! Проверте имя папки в которую нужно сохранить модуль.", MsgBoxStyle.OKOnly, "Сообщение:")
            Label12.BorderStyle = BorderStyle.None
        End If
    End Sub

    Private Sub Label12_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label12.MouseMove
        Try
            Label12.BorderStyle = BorderStyle.FixedSingle
            Label13.BorderStyle = BorderStyle.None
            Label14.BorderStyle = BorderStyle.None
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label12_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label12.MouseLeave
        Try
            Label12.BorderStyle = BorderStyle.None
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label13_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label13.MouseMove
        Try
            Label13.BorderStyle = BorderStyle.FixedSingle
            Label12.BorderStyle = BorderStyle.None
            Label14.BorderStyle = BorderStyle.None
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label13_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label13.MouseLeave
        Try
            Label13.BorderStyle = BorderStyle.None
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label14_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label14.MouseMove
        Try
            Label14.BorderStyle = BorderStyle.FixedSingle
            Label12.BorderStyle = BorderStyle.None
            Label13.BorderStyle = BorderStyle.None
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label14_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label14.MouseLeave
        Try
            Label14.BorderStyle = BorderStyle.None
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label15_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label15.MouseMove
        Try
            Label15.BorderStyle = BorderStyle.FixedSingle
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label15_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label15.MouseLeave
        Try
            Label15.BorderStyle = BorderStyle.None
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub Label16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label16.Click
        NewOrEdit = "NEW"
        TextBox1.Text = ""
        ComboBox1.Text = ""
        RichTextBox1.Clear()
        DataSet51.Clear()
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        TREE_UPDATE()
    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click
        NewOrEdit = "NEW"
        TextBox1.Text = ""
        ComboBox1.Text = StatusBarPanel2.Text
        RichTextBox1.Clear()
        DataSet51.Clear()
    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Try
            NewOrEdit = "EDIT"
            TextBox1.Text = ""
            ComboBox1.Text = ""
            RichTextBox1.Clear()
            DataSet51.Clear()
            Dim ContinueEdit As Boolean = False
            For iFindFile As Integer = 0 To DataSet21.Хранилище.Count - 1
                If (DataSet21.Хранилище.Item(iFindFile)("ТипОбъекта") = "Элемент") Then
                    If (DataSet21.Хранилище.Item(iFindFile)("ФайлИдентификатор") = StatusBarPanel2.Text) Then
                        ContinueEdit = True
                        Exit For
                    End If
                End If
            Next
            If (ContinueEdit = True) Then
                OleDbConnection1.Open()
                Dim SQLComand As String
                SQLComand = "SELECT ДатаПоследнегоСохранения, ПапкаИдентификатор, СодержаниеФайла, Строка, ТипОбъекта, ФайлВПапке, ФайлИдентификатор FROM Хранилище WHERE (ФайлИдентификатор = '" + StatusBarPanel2.Text + "')"
                OleDbDataAdapter2.SelectCommand.CommandText = SQLComand
                OleDbDataAdapter2.Fill(DataSet51, "Хранилище")
                TextBox1.Text = DataSet51.Хранилище.Item(0)("ФайлИдентификатор")
                ComboBox1.Text = DataSet51.Хранилище.Item(0)("ФайлВПапке")
                RichTextBox1.Rtf = DataSet51.Хранилище.Item(0)("СодержаниеФайла")
                StatusBarPanel3.Text = DataSet51.Хранилище.Item(0)("ДатаПоследнегоСохранения")
                ActionID = DataSet51.Хранилище.Item(0)("Строка")
                OleDbConnection1.Close()
            End If
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click
        'Удалить модуль
        Try
            Dim DelModule As Boolean = False
            If (MsgBox("Вы увирены что хотите удалить файл [" + StatusBarPanel2.Text + "]?", MsgBoxStyle.YesNo, "Вопрос: ") = MsgBoxResult.Yes) Then
                For i As Integer = 0 To DataSet21.Хранилище.Count - 1
                    If (DataSet21.Хранилище.Item(i)("ТипОбъекта") = "Элемент") Then
                        If (DataSet21.Хранилище.Item(i)("ТипОбъекта") = "Элемент") And (DataSet21.Хранилище.Item(i)("ФайлИдентификатор") = StatusBarPanel2.Text) Then
                            OleDbConnection1.Open()
                            DataSet21.Хранилище.Item(i).Delete()
                            OleDbDataAdapter1.Update(DataSet21, "Хранилище")
                            OleDbConnection1.Close()
                            DelModule = True
                            Exit For
                        End If
                    End If
                Next
                If (DelModule = True) Then
                    TREE_UPDATE()
                    NewOrEdit = "NEW"
                    TextBox1.Text = ""
                    ComboBox1.Text = ""
                    RichTextBox1.Clear()
                    DataSet51.Clear()
                    MsgBox("Удаление прошло успешно.", MsgBoxStyle.OKOnly, "Сообщение:")
                Else
                    MsgBox("Ошибка!!! " + "[" + StatusBarPanel2.Text + "] не является файлом!", MsgBoxStyle.OKOnly, "Сообщение:")
                End If
            End If
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub TreeView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles TreeView1.DoubleClick
        Try
            For i As Integer = 0 To DataSet21.Хранилище.Count - 1
                If (DataSet21.Хранилище.Item(i)("ТипОбъекта") = "Элемент") Then
                    If (DataSet21.Хранилище.Item(i)("ФайлИдентификатор") = StatusBarPanel2.Text) Then
                        NewOrEdit = "EDIT"
                        TextBox1.Text = ""
                        ComboBox1.Text = ""
                        RichTextBox1.Clear()
                        DataSet51.Clear()
                        Dim ContinueEdit As Boolean = False
                        For iFindFile As Integer = 0 To DataSet21.Хранилище.Count - 1
                            If (DataSet21.Хранилище.Item(iFindFile)("ТипОбъекта") = "Элемент") Then
                                If (DataSet21.Хранилище.Item(iFindFile)("ФайлИдентификатор") = StatusBarPanel2.Text) Then
                                    ContinueEdit = True
                                    Exit For
                                End If
                            End If
                        Next
                        If (ContinueEdit = True) Then
                            OleDbConnection1.Open()
                            Dim SQLComand As String
                            SQLComand = "SELECT ДатаПоследнегоСохранения, ПапкаИдентификатор, СодержаниеФайла, Строка, ТипОбъекта, ФайлВПапке, ФайлИдентификатор FROM Хранилище WHERE (ФайлИдентификатор = '" + StatusBarPanel2.Text + "')"
                            OleDbDataAdapter2.SelectCommand.CommandText = SQLComand
                            OleDbDataAdapter2.Fill(DataSet51, "Хранилище")
                            TextBox1.Text = DataSet51.Хранилище.Item(0)("ФайлИдентификатор")
                            ComboBox1.Text = DataSet51.Хранилище.Item(0)("ФайлВПапке")
                            RichTextBox1.Rtf = DataSet51.Хранилище.Item(0)("СодержаниеФайла")
                            StatusBarPanel3.Text = DataSet51.Хранилище.Item(0)("ДатаПоследнегоСохранения")
                            ActionID = DataSet51.Хранилище.Item(0)("Строка")
                            OleDbConnection1.Close()
                            Exit For
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub MenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem6.Click
        Try
            RichTextBox1.Paste()
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub MenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem7.Click
        Try
            RichTextBox1.Copy()
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub MenuItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem8.Click
        Try
            RichTextBox1.Cut()
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub MenuItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem9.Click
        Try
            RichTextBox1.Undo()
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub RichTextBox1_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkClickedEventArgs) Handles RichTextBox1.LinkClicked
        System.Diagnostics.Process.Start(e.LinkText)
    End Sub

    Private Sub TreeView1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TreeView1.KeyPress
        
    End Sub

    Private Sub TreeView1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TreeView1.KeyDown
        Try
            If (e.KeyData = Keys.Enter) Then
                For i As Integer = 0 To DataSet21.Хранилище.Count - 1
                    If (DataSet21.Хранилище.Item(i)("ТипОбъекта") = "Элемент") Then
                        If (DataSet21.Хранилище.Item(i)("ФайлИдентификатор") = StatusBarPanel2.Text) Then
                            NewOrEdit = "EDIT"
                            TextBox1.Text = ""
                            ComboBox1.Text = ""
                            RichTextBox1.Clear()
                            DataSet51.Clear()
                            Dim ContinueEdit As Boolean = False
                            For iFindFile As Integer = 0 To DataSet21.Хранилище.Count - 1
                                If (DataSet21.Хранилище.Item(iFindFile)("ТипОбъекта") = "Элемент") Then
                                    If (DataSet21.Хранилище.Item(iFindFile)("ФайлИдентификатор") = StatusBarPanel2.Text) Then
                                        ContinueEdit = True
                                        Exit For
                                    End If
                                End If
                            Next
                            If (ContinueEdit = True) Then
                                OleDbConnection1.Open()
                                Dim SQLComand As String
                                SQLComand = "SELECT ДатаПоследнегоСохранения, ПапкаИдентификатор, СодержаниеФайла, Строка, ТипОбъекта, ФайлВПапке, ФайлИдентификатор FROM Хранилище WHERE (ФайлИдентификатор = '" + StatusBarPanel2.Text + "')"
                                OleDbDataAdapter2.SelectCommand.CommandText = SQLComand
                                OleDbDataAdapter2.Fill(DataSet51, "Хранилище")
                                TextBox1.Text = DataSet51.Хранилище.Item(0)("ФайлИдентификатор")
                                ComboBox1.Text = DataSet51.Хранилище.Item(0)("ФайлВПапке")
                                RichTextBox1.Rtf = DataSet51.Хранилище.Item(0)("СодержаниеФайла")
                                StatusBarPanel3.Text = DataSet51.Хранилище.Item(0)("ДатаПоследнегоСохранения")
                                ActionID = DataSet51.Хранилище.Item(0)("Строка")
                                OleDbConnection1.Close()
                                Exit For
                            End If
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            'MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
    End Sub

    Private Sub TreeView1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TreeView1.MouseMove
        Label2.BorderStyle = BorderStyle.None
        Label3.BorderStyle = BorderStyle.None
        Label4.BorderStyle = BorderStyle.None
        Label5.BorderStyle = BorderStyle.None
        Label6.BorderStyle = BorderStyle.None
        Label7.BorderStyle = BorderStyle.None
        Label8.BorderStyle = BorderStyle.None
        Label9.BorderStyle = BorderStyle.None
        Label12.BorderStyle = BorderStyle.None
        Label13.BorderStyle = BorderStyle.None
        Label14.BorderStyle = BorderStyle.None
        Label15.BorderStyle = BorderStyle.None
    End Sub

    Private Sub RichTextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RichTextBox1.TextChanged
        Label2.BorderStyle = BorderStyle.None
        Label3.BorderStyle = BorderStyle.None
        Label4.BorderStyle = BorderStyle.None
        Label5.BorderStyle = BorderStyle.None
        Label6.BorderStyle = BorderStyle.None
        Label7.BorderStyle = BorderStyle.None
        Label8.BorderStyle = BorderStyle.None
        Label9.BorderStyle = BorderStyle.None
        Label12.BorderStyle = BorderStyle.None
        Label13.BorderStyle = BorderStyle.None
        Label14.BorderStyle = BorderStyle.None
        Label15.BorderStyle = BorderStyle.None
    End Sub

    Private Sub RichTextBox1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles RichTextBox1.MouseMove
        Label2.BorderStyle = BorderStyle.None
        Label3.BorderStyle = BorderStyle.None
        Label4.BorderStyle = BorderStyle.None
        Label5.BorderStyle = BorderStyle.None
        Label6.BorderStyle = BorderStyle.None
        Label7.BorderStyle = BorderStyle.None
        Label8.BorderStyle = BorderStyle.None
        Label9.BorderStyle = BorderStyle.None
        Label12.BorderStyle = BorderStyle.None
        Label13.BorderStyle = BorderStyle.None
        Label14.BorderStyle = BorderStyle.None
        Label15.BorderStyle = BorderStyle.None
    End Sub

    Private Sub MenuItem12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem12.Click
        MenuItem12.Checked = True
        MenuItem13.Checked = False
        TREE_UPDATE()
    End Sub

    Private Sub MenuItem13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem13.Click
        MenuItem13.Checked = True
        MenuItem12.Checked = False
        TREE_UPDATE()
    End Sub

    Private Sub MenuItem15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem15.Click
        If (ColorDialog1.ShowDialog = DialogResult.OK) Then
            RichTextBox1.SelectionColor = ColorDialog1.Color
        End If
    End Sub

    Private Sub MenuItem16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem16.Click
        If (FontDialog1.ShowDialog = DialogResult.OK) Then
            RichTextBox1.SelectionFont = FontDialog1.Font
        End If
    End Sub
End Class
