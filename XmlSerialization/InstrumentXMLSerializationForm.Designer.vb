<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class InstrumentXMLSerializationForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.ExitButton = New System.Windows.Forms.Button()
        Me.SaveButton = New System.Windows.Forms.Button()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.AutoFillButton = New System.Windows.Forms.Button()
        Me.LoadButton = New System.Windows.Forms.Button()
        Me.SumdataGridView = New System.Windows.Forms.DataGridView()
        Me.TestName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Instrument = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Count = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ComType = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.Address = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.SumdataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.ExitButton)
        Me.SplitContainer1.Panel1.Controls.Add(Me.SaveButton)
        Me.SplitContainer1.Panel1.Controls.Add(Me.RichTextBox1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.AutoFillButton)
        Me.SplitContainer1.Panel1.Controls.Add(Me.LoadButton)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.SumdataGridView)
        Me.SplitContainer1.Size = New System.Drawing.Size(747, 621)
        Me.SplitContainer1.SplitterDistance = 104
        Me.SplitContainer1.TabIndex = 1
        '
        'ExitButton
        '
        Me.ExitButton.Location = New System.Drawing.Point(21, 399)
        Me.ExitButton.Name = "ExitButton"
        Me.ExitButton.Size = New System.Drawing.Size(75, 23)
        Me.ExitButton.TabIndex = 4
        Me.ExitButton.Text = "Close"
        Me.ExitButton.UseVisualStyleBackColor = True
        '
        'SaveButton
        '
        Me.SaveButton.Location = New System.Drawing.Point(21, 151)
        Me.SaveButton.Name = "SaveButton"
        Me.SaveButton.Size = New System.Drawing.Size(75, 23)
        Me.SaveButton.TabIndex = 3
        Me.SaveButton.Text = "Save"
        Me.SaveButton.UseVisualStyleBackColor = True
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Location = New System.Drawing.Point(15, 260)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.Size = New System.Drawing.Size(91, 84)
        Me.RichTextBox1.TabIndex = 2
        Me.RichTextBox1.Text = "Loaded from the last session"
        '
        'AutoFillButton
        '
        Me.AutoFillButton.Location = New System.Drawing.Point(21, 97)
        Me.AutoFillButton.Name = "AutoFillButton"
        Me.AutoFillButton.Size = New System.Drawing.Size(75, 23)
        Me.AutoFillButton.TabIndex = 0
        Me.AutoFillButton.Text = "Auto Fill"
        Me.AutoFillButton.UseVisualStyleBackColor = True
        '
        'LoadButton
        '
        Me.LoadButton.Location = New System.Drawing.Point(21, 39)
        Me.LoadButton.Name = "LoadButton"
        Me.LoadButton.Size = New System.Drawing.Size(75, 23)
        Me.LoadButton.TabIndex = 0
        Me.LoadButton.Text = "Load"
        Me.LoadButton.UseVisualStyleBackColor = True
        '
        'SumdataGridView
        '
        Me.SumdataGridView.AllowUserToAddRows = False
        Me.SumdataGridView.AllowUserToDeleteRows = False
        Me.SumdataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.SumdataGridView.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.TestName, Me.Instrument, Me.Count, Me.ComType, Me.Address})
        Me.SumdataGridView.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SumdataGridView.Location = New System.Drawing.Point(0, 0)
        Me.SumdataGridView.Name = "SumdataGridView"
        Me.SumdataGridView.Size = New System.Drawing.Size(639, 621)
        Me.SumdataGridView.TabIndex = 0
        '
        'TestName
        '
        Me.TestName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.TestName.HeaderText = "Test Name"
        Me.TestName.Name = "TestName"
        Me.TestName.ReadOnly = True
        Me.TestName.Width = 84
        '
        'Instrument
        '
        Me.Instrument.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.Instrument.HeaderText = "Instrument"
        Me.Instrument.Name = "Instrument"
        Me.Instrument.ReadOnly = True
        Me.Instrument.Width = 81
        '
        'Count
        '
        Me.Count.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.Count.HeaderText = "Count"
        Me.Count.Name = "Count"
        Me.Count.ReadOnly = True
        Me.Count.Width = 60
        '
        'ComType
        '
        Me.ComType.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.ComType.HeaderText = "ComType"
        Me.ComType.Items.AddRange(New Object() {"GPIB", "USB", "RS232", "USBX", "LAN"})
        Me.ComType.Name = "ComType"
        Me.ComType.Width = 58
        '
        'Address
        '
        Me.Address.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Address.HeaderText = "Address"
        Me.Address.Name = "Address"
        '
        'InstrumentXMLSerializationForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(747, 621)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Name = "InstrumentXMLSerializationForm"
        Me.Text = "Select Instrument ComType and Address"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.SumdataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents ExitButton As System.Windows.Forms.Button
    Friend WithEvents SaveButton As System.Windows.Forms.Button
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents AutoFillButton As System.Windows.Forms.Button
    Friend WithEvents LoadButton As System.Windows.Forms.Button
    Friend WithEvents SumdataGridView As System.Windows.Forms.DataGridView
    Friend WithEvents TestName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Instrument As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Count As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ComType As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Address As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
