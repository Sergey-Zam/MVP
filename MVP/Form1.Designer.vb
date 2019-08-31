<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnGetExcelPath = New System.Windows.Forms.Button()
        Me.btnGetInventorPath = New System.Windows.Forms.Button()
        Me.btnGetData = New System.Windows.Forms.Button()
        Me.btnExportTable = New System.Windows.Forms.Button()
        Me.btnClearTable = New System.Windows.Forms.Button()
        Me.tbExcelDirectory = New System.Windows.Forms.TextBox()
        Me.tbInventorDirectory = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblCountOfRows = New System.Windows.Forms.Label()
        Me.dgvAspects = New System.Windows.Forms.DataGridView()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        CType(Me.dgvAspects, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnGetExcelPath
        '
        Me.btnGetExcelPath.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnGetExcelPath.Location = New System.Drawing.Point(697, 12)
        Me.btnGetExcelPath.Name = "btnGetExcelPath"
        Me.btnGetExcelPath.Size = New System.Drawing.Size(75, 23)
        Me.btnGetExcelPath.TabIndex = 0
        Me.btnGetExcelPath.Text = "Выбрать..."
        Me.btnGetExcelPath.UseVisualStyleBackColor = True
        '
        'btnGetInventorPath
        '
        Me.btnGetInventorPath.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnGetInventorPath.Location = New System.Drawing.Point(697, 41)
        Me.btnGetInventorPath.Name = "btnGetInventorPath"
        Me.btnGetInventorPath.Size = New System.Drawing.Size(75, 23)
        Me.btnGetInventorPath.TabIndex = 1
        Me.btnGetInventorPath.Text = "Выбрать..."
        Me.btnGetInventorPath.UseVisualStyleBackColor = True
        '
        'btnGetData
        '
        Me.btnGetData.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnGetData.Location = New System.Drawing.Point(12, 76)
        Me.btnGetData.Name = "btnGetData"
        Me.btnGetData.Size = New System.Drawing.Size(760, 23)
        Me.btnGetData.TabIndex = 2
        Me.btnGetData.Text = "Считать и сравнить данные"
        Me.btnGetData.UseVisualStyleBackColor = True
        '
        'btnExportTable
        '
        Me.btnExportTable.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExportTable.AutoSize = True
        Me.btnExportTable.Location = New System.Drawing.Point(545, 326)
        Me.btnExportTable.Name = "btnExportTable"
        Me.btnExportTable.Size = New System.Drawing.Size(114, 23)
        Me.btnExportTable.TabIndex = 3
        Me.btnExportTable.Text = "Экспорт таблицы..."
        Me.btnExportTable.UseVisualStyleBackColor = True
        '
        'btnClearTable
        '
        Me.btnClearTable.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClearTable.AutoSize = True
        Me.btnClearTable.Location = New System.Drawing.Point(665, 326)
        Me.btnClearTable.Name = "btnClearTable"
        Me.btnClearTable.Size = New System.Drawing.Size(107, 23)
        Me.btnClearTable.TabIndex = 4
        Me.btnClearTable.Text = "Очистить таблицу"
        Me.btnClearTable.UseVisualStyleBackColor = True
        '
        'tbExcelDirectory
        '
        Me.tbExcelDirectory.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbExcelDirectory.Location = New System.Drawing.Point(201, 14)
        Me.tbExcelDirectory.Name = "tbExcelDirectory"
        Me.tbExcelDirectory.ReadOnly = True
        Me.tbExcelDirectory.Size = New System.Drawing.Size(490, 20)
        Me.tbExcelDirectory.TabIndex = 5
        '
        'tbInventorDirectory
        '
        Me.tbInventorDirectory.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbInventorDirectory.Location = New System.Drawing.Point(201, 43)
        Me.tbInventorDirectory.Name = "tbInventorDirectory"
        Me.tbInventorDirectory.ReadOnly = True
        Me.tbInventorDirectory.Size = New System.Drawing.Size(490, 20)
        Me.tbInventorDirectory.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label1.Location = New System.Drawing.Point(9, 112)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(235, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Таблица полученных аспектов, всего:"
        '
        'lblCountOfRows
        '
        Me.lblCountOfRows.AutoSize = True
        Me.lblCountOfRows.Location = New System.Drawing.Point(250, 112)
        Me.lblCountOfRows.Name = "lblCountOfRows"
        Me.lblCountOfRows.Size = New System.Drawing.Size(13, 13)
        Me.lblCountOfRows.TabIndex = 8
        Me.lblCountOfRows.Text = "0"
        '
        'dgvAspects
        '
        Me.dgvAspects.AllowUserToAddRows = False
        Me.dgvAspects.AllowUserToDeleteRows = False
        Me.dgvAspects.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvAspects.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvAspects.Location = New System.Drawing.Point(12, 128)
        Me.dgvAspects.Name = "dgvAspects"
        Me.dgvAspects.ReadOnly = True
        Me.dgvAspects.RowHeadersVisible = False
        Me.dgvAspects.Size = New System.Drawing.Size(760, 192)
        Me.dgvAspects.TabIndex = 9
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 17)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(183, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Документ с аспектами (*.xlsx, *.xls)"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(66, 46)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(129, 13)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Документ сборки (*.iam)"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(784, 361)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.dgvAspects)
        Me.Controls.Add(Me.lblCountOfRows)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.tbInventorDirectory)
        Me.Controls.Add(Me.tbExcelDirectory)
        Me.Controls.Add(Me.btnClearTable)
        Me.Controls.Add(Me.btnExportTable)
        Me.Controls.Add(Me.btnGetData)
        Me.Controls.Add(Me.btnGetInventorPath)
        Me.Controls.Add(Me.btnGetExcelPath)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "MVP"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.dgvAspects, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnGetExcelPath As Button
    Friend WithEvents btnGetInventorPath As Button
    Friend WithEvents btnGetData As Button
    Friend WithEvents btnExportTable As Button
    Friend WithEvents btnClearTable As Button
    Friend WithEvents tbExcelDirectory As TextBox
    Friend WithEvents tbInventorDirectory As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents lblCountOfRows As Label
    Friend WithEvents dgvAspects As DataGridView
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
End Class
