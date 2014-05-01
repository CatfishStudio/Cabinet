Imports System.IO
Public Class Start
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
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents OleDbSelectCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbInsertCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbUpdateCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbDeleteCommand1 As System.Data.OleDb.OleDbCommand
    Friend WithEvents OleDbConnection1 As System.Data.OleDb.OleDbConnection
    Friend WithEvents OleDbDataAdapter1 As System.Data.OleDb.OleDbDataAdapter
    Friend WithEvents DataSet11 As Cabinet.DataSet1
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Start))
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.OleDbSelectCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbConnection1 = New System.Data.OleDb.OleDbConnection
        Me.OleDbInsertCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbUpdateCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbDeleteCommand1 = New System.Data.OleDb.OleDbCommand
        Me.OleDbDataAdapter1 = New System.Data.OleDb.OleDbDataAdapter
        Me.DataSet11 = New Cabinet.DataSet1
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        CType(Me.DataSet11, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PictureBox1.Cursor = System.Windows.Forms.Cursors.WaitCursor
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(0, 0)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(352, 252)
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'OleDbSelectCommand1
        '
        Me.OleDbSelectCommand1.CommandText = "SELECT ИмяПользователя, ИмяШрифта, ИспользоватьПароль, ПарольПользователя, Поверх" & _
        "Окон, РазмерШрифта, Строка, ФорматированиеМодуля FROM Настройки"
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
        Me.OleDbInsertCommand1.CommandText = "INSERT INTO Настройки(ИмяПользователя, ИмяШрифта, ИспользоватьПароль, ПарольПольз" & _
        "ователя, ПоверхОкон, РазмерШрифта, ФорматированиеМодуля) VALUES (?, ?, ?, ?, ?, " & _
        "?, ?)"
        Me.OleDbInsertCommand1.Connection = Me.OleDbConnection1
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ИмяПользователя", System.Data.OleDb.OleDbType.VarWChar, 255, "ИмяПользователя"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ИмяШрифта", System.Data.OleDb.OleDbType.VarWChar, 255, "ИмяШрифта"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ИспользоватьПароль", System.Data.OleDb.OleDbType.Integer, 0, "ИспользоватьПароль"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ПарольПользователя", System.Data.OleDb.OleDbType.VarWChar, 255, "ПарольПользователя"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ПоверхОкон", System.Data.OleDb.OleDbType.Integer, 0, "ПоверхОкон"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("РазмерШрифта", System.Data.OleDb.OleDbType.VarWChar, 255, "РазмерШрифта"))
        Me.OleDbInsertCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ФорматированиеМодуля", System.Data.OleDb.OleDbType.Integer, 0, "ФорматированиеМодуля"))
        '
        'OleDbUpdateCommand1
        '
        Me.OleDbUpdateCommand1.CommandText = "UPDATE Настройки SET ИмяПользователя = ?, ИмяШрифта = ?, ИспользоватьПароль = ?, " & _
        "ПарольПользователя = ?, ПоверхОкон = ?, РазмерШрифта = ?, ФорматированиеМодуля =" & _
        " ? WHERE (Строка = ?) AND (ИмяПользователя = ? OR ? IS NULL AND ИмяПользователя " & _
        "IS NULL) AND (ИмяШрифта = ? OR ? IS NULL AND ИмяШрифта IS NULL) AND (Использоват" & _
        "ьПароль = ? OR ? IS NULL AND ИспользоватьПароль IS NULL) AND (ПарольПользователя" & _
        " = ? OR ? IS NULL AND ПарольПользователя IS NULL) AND (ПоверхОкон = ? OR ? IS NU" & _
        "LL AND ПоверхОкон IS NULL) AND (РазмерШрифта = ? OR ? IS NULL AND РазмерШрифта I" & _
        "S NULL) AND (ФорматированиеМодуля = ? OR ? IS NULL AND ФорматированиеМодуля IS N" & _
        "ULL)"
        Me.OleDbUpdateCommand1.Connection = Me.OleDbConnection1
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ИмяПользователя", System.Data.OleDb.OleDbType.VarWChar, 255, "ИмяПользователя"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ИмяШрифта", System.Data.OleDb.OleDbType.VarWChar, 255, "ИмяШрифта"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ИспользоватьПароль", System.Data.OleDb.OleDbType.Integer, 0, "ИспользоватьПароль"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ПарольПользователя", System.Data.OleDb.OleDbType.VarWChar, 255, "ПарольПользователя"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ПоверхОкон", System.Data.OleDb.OleDbType.Integer, 0, "ПоверхОкон"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("РазмерШрифта", System.Data.OleDb.OleDbType.VarWChar, 255, "РазмерШрифта"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("ФорматированиеМодуля", System.Data.OleDb.OleDbType.Integer, 0, "ФорматированиеМодуля"))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Строка", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Строка", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ИмяПользователя", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ИмяПользователя", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ИмяПользователя1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ИмяПользователя", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ИмяШрифта", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ИмяШрифта", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ИмяШрифта1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ИмяШрифта", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ИспользоватьПароль", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ИспользоватьПароль", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ИспользоватьПароль1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ИспользоватьПароль", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ПарольПользователя", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ПарольПользователя", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ПарольПользователя1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ПарольПользователя", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ПоверхОкон", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ПоверхОкон", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ПоверхОкон1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ПоверхОкон", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_РазмерШрифта", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "РазмерШрифта", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_РазмерШрифта1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "РазмерШрифта", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФорматированиеМодуля", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФорматированиеМодуля", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbUpdateCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФорматированиеМодуля1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФорматированиеМодуля", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbDeleteCommand1
        '
        Me.OleDbDeleteCommand1.CommandText = "DELETE FROM Настройки WHERE (Строка = ?) AND (ИмяПользователя = ? OR ? IS NULL AN" & _
        "D ИмяПользователя IS NULL) AND (ИмяШрифта = ? OR ? IS NULL AND ИмяШрифта IS NULL" & _
        ") AND (ИспользоватьПароль = ? OR ? IS NULL AND ИспользоватьПароль IS NULL) AND (" & _
        "ПарольПользователя = ? OR ? IS NULL AND ПарольПользователя IS NULL) AND (ПоверхО" & _
        "кон = ? OR ? IS NULL AND ПоверхОкон IS NULL) AND (РазмерШрифта = ? OR ? IS NULL " & _
        "AND РазмерШрифта IS NULL) AND (ФорматированиеМодуля = ? OR ? IS NULL AND Формати" & _
        "рованиеМодуля IS NULL)"
        Me.OleDbDeleteCommand1.Connection = Me.OleDbConnection1
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_Строка", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "Строка", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ИмяПользователя", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ИмяПользователя", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ИмяПользователя1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ИмяПользователя", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ИмяШрифта", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ИмяШрифта", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ИмяШрифта1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ИмяШрифта", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ИспользоватьПароль", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ИспользоватьПароль", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ИспользоватьПароль1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ИспользоватьПароль", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ПарольПользователя", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ПарольПользователя", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ПарольПользователя1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ПарольПользователя", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ПоверхОкон", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ПоверхОкон", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ПоверхОкон1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ПоверхОкон", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_РазмерШрифта", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "РазмерШрифта", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_РазмерШрифта1", System.Data.OleDb.OleDbType.VarWChar, 255, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "РазмерШрифта", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФорматированиеМодуля", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФорматированиеМодуля", System.Data.DataRowVersion.Original, Nothing))
        Me.OleDbDeleteCommand1.Parameters.Add(New System.Data.OleDb.OleDbParameter("Original_ФорматированиеМодуля1", System.Data.OleDb.OleDbType.Integer, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "ФорматированиеМодуля", System.Data.DataRowVersion.Original, Nothing))
        '
        'OleDbDataAdapter1
        '
        Me.OleDbDataAdapter1.DeleteCommand = Me.OleDbDeleteCommand1
        Me.OleDbDataAdapter1.InsertCommand = Me.OleDbInsertCommand1
        Me.OleDbDataAdapter1.SelectCommand = Me.OleDbSelectCommand1
        Me.OleDbDataAdapter1.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "Настройки", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("ИмяПользователя", "ИмяПользователя"), New System.Data.Common.DataColumnMapping("ИмяШрифта", "ИмяШрифта"), New System.Data.Common.DataColumnMapping("ИспользоватьПароль", "ИспользоватьПароль"), New System.Data.Common.DataColumnMapping("ПарольПользователя", "ПарольПользователя"), New System.Data.Common.DataColumnMapping("ПоверхОкон", "ПоверхОкон"), New System.Data.Common.DataColumnMapping("РазмерШрифта", "РазмерШрифта"), New System.Data.Common.DataColumnMapping("Строка", "Строка"), New System.Data.Common.DataColumnMapping("ФорматированиеМодуля", "ФорматированиеМодуля")})})
        Me.OleDbDataAdapter1.UpdateCommand = Me.OleDbUpdateCommand1
        '
        'DataSet11
        '
        Me.DataSet11.DataSetName = "DataSet1"
        Me.DataSet11.Locale = New System.Globalization.CultureInfo("ru-RU")
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.ContextMenu = Me.ContextMenu1
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "Cabinet"
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem2})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.Text = "Открыть личный кабинет."
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.Text = "Закрыть программу."
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(8, 208)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(336, 16)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Portions Copyright © Microsoft Corporation. All Rights Reserved."
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(8, 192)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(296, 16)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "© Somov Evgen, 2012"
        '
        'Start
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.ClientSize = New System.Drawing.Size(352, 252)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.PictureBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Start"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Cabinet"
        CType(Me.DataSet11, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Stop()
        Me.Visible = False
        NotifyIcon1.Visible = True
    End Sub

    Private Sub Start_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Dim INFO As DirectoryInfo = New DirectoryInfo("resource") 'получаем адрес программы
            ''Dim MyProgramDirectory As String
            Dim strName As String
            Dim nDel As Integer
            strName = INFO.FullName.ToString
            nDel = strName.Length - 8   'отнимаем четыре символа TASM
            ModuleBase.MyProgramDirectory = strName.Remove(nDel, 8) 'удаляем символы

            Me.Update()
            ModuleBase.PathBase = ModuleBase.MyProgramDirectory + "resource\cabinet.ccb"
            ModuleBase.StringConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ModuleBase.PathBase & ";Jet OLEDB:Database Password=WX0490134QL"
            OleDbConnection1.ConnectionString = ModuleBase.StringConnection
            OleDbConnection1.Open()
            OleDbDataAdapter1.Fill(DataSet11, "Настройки")
            Dim NameF As String = DataSet11.Настройки.Item(0)("ИмяШрифта")
            Dim SizeF As Integer = DataSet11.Настройки.Item(0)("РазмерШрифта")
            ModuleCabinet.FontModule = New Font(NameF, SizeF, FontStyle.Regular)
            If (DataSet11.Настройки.Item(0)("ФорматированиеМодуля") = 1) Then
                ModuleCabinet.ActionModuleFormat = True
            Else
                ModuleCabinet.ActionModuleFormat = False
            End If
            If (DataSet11.Настройки.Item(0)("ПоверхОкон") = 1) Then
                ModuleCabinet.ActionTopMost = True
            Else
                ModuleCabinet.ActionTopMost = False
            End If
            If (DataSet11.Настройки.Item(0)("ИспользоватьПароль") = 1) Then
                ModuleCabinet.ActionUserPass = True
                ModuleCabinet.UserName = DataSet11.Настройки.Item(0)("ИмяПользователя")
                ModuleCabinet.UserPass = DataSet11.Настройки.Item(0)("ПарольПользователя")
            Else
                ModuleCabinet.ActionUserPass = False
                ModuleCabinet.UserName = ""
                ModuleCabinet.UserPass = ""
            End If
            '...
            OleDbConnection1.Close()
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.OKOnly, "Произошла ошибка!")
        End Try
        Timer1.Start()
    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        End
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        If (ModuleCabinet.ActionUserPass = True) Then
            If (ModuleCabinet.UserLoad = False) Then
                ModuleCabinet.LoadTUser()
                ModuleCabinet.TUser.TextBox1.Text = ModuleCabinet.UserName
                ModuleCabinet.TUser.Show()
                ModuleCabinet.UserLoad = True
            Else
                ModuleCabinet.LoadProgramm()
                If (ModuleCabinet.ActionTopMost = True) Then ModuleCabinet.Programm.TopMost = True
                ModuleCabinet.Programm.RichTextBox1.Font = ModuleCabinet.FontModule
                ModuleCabinet.Programm.Show()
                ModuleCabinet.UserLoad = True
            End If
        Else
            ModuleCabinet.LoadProgramm()
            If (ModuleCabinet.ActionTopMost = True) Then ModuleCabinet.Programm.TopMost = True
            ModuleCabinet.Programm.RichTextBox1.Font = ModuleCabinet.FontModule
            ModuleCabinet.Programm.Show()
            ModuleCabinet.UserLoad = True
        End If
    End Sub

    Private Sub NotifyIcon1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.DoubleClick
        If (ModuleCabinet.ActionUserPass = True) Then
            If (ModuleCabinet.UserLoad = False) Then
                ModuleCabinet.LoadTUser()
                ModuleCabinet.TUser.TextBox1.Text = ModuleCabinet.UserName
                ModuleCabinet.TUser.Show()
                ModuleCabinet.UserLoad = True
            Else
                ModuleCabinet.LoadProgramm()
                If (ModuleCabinet.ActionTopMost = True) Then ModuleCabinet.Programm.TopMost = True
                ModuleCabinet.Programm.RichTextBox1.Font = ModuleCabinet.FontModule
                ModuleCabinet.Programm.Show()
                ModuleCabinet.UserLoad = True
            End If
        Else
            ModuleCabinet.LoadProgramm()
            If (ModuleCabinet.ActionTopMost = True) Then ModuleCabinet.Programm.TopMost = True
            ModuleCabinet.Programm.RichTextBox1.Font = ModuleCabinet.FontModule
            ModuleCabinet.Programm.Show()
            ModuleCabinet.UserLoad = True
        End If
    End Sub

    Private Sub NotifyIcon1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDown

    End Sub
End Class
