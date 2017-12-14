<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ResponseView
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
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ResponseView))
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Splitter1 = New System.Windows.Forms.Splitter()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.FilterCombo3 = New TemplateDB.FilterCombo()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.FilterCombo5 = New TemplateDB.FilterCombo()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.FilterCombo4 = New TemplateDB.FilterCombo()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.FilterCombo30 = New TemplateDB.FilterCombo()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.FilterCombo90 = New TemplateDB.FilterCombo()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.FilterCombo1 = New TemplateDB.FilterCombo()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.FilterCombo2 = New TemplateDB.FilterCombo()
        Me.StaffQueryGrid = New System.Windows.Forms.DataGridView()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.BindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        CType(Me.StaffQueryGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer1.IsSplitterFixed = True
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label8)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Splitter1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.FilterCombo3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label7)
        Me.SplitContainer1.Panel1.Controls.Add(Me.FilterCombo5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.FilterCombo4)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.FilterCombo30)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel1.Controls.Add(Me.FilterCombo90)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.FilterCombo1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer1.Panel1.Controls.Add(Me.FilterCombo2)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.StaffQueryGrid)
        Me.SplitContainer1.Size = New System.Drawing.Size(1256, 484)
        Me.SplitContainer1.SplitterDistance = 35
        Me.SplitContainer1.TabIndex = 0
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(42, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(57, 20)
        Me.Label8.TabIndex = 39
        Me.Label8.Text = "Label8"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'Splitter1
        '
        Me.Splitter1.Dock = System.Windows.Forms.DockStyle.Right
        Me.Splitter1.Location = New System.Drawing.Point(99, 0)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(29, 35)
        Me.Splitter1.TabIndex = 38
        Me.Splitter1.TabStop = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(128, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(50, 20)
        Me.Label4.TabIndex = 31
        Me.Label4.Text = "Study"
        '
        'FilterCombo3
        '
        Me.FilterCombo3.Dock = System.Windows.Forms.DockStyle.Right
        Me.FilterCombo3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FilterCombo3.FormattingEnabled = True
        Me.FilterCombo3.Location = New System.Drawing.Point(178, 0)
        Me.FilterCombo3.Margin = New System.Windows.Forms.Padding(2)
        Me.FilterCombo3.Name = "FilterCombo3"
        Me.FilterCombo3.Size = New System.Drawing.Size(142, 25)
        Me.FilterCombo3.TabIndex = 32
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(320, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(57, 20)
        Me.Label7.TabIndex = 35
        Me.Label7.Text = "Cohort"
        '
        'FilterCombo5
        '
        Me.FilterCombo5.Dock = System.Windows.Forms.DockStyle.Right
        Me.FilterCombo5.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FilterCombo5.FormattingEnabled = True
        Me.FilterCombo5.Location = New System.Drawing.Point(377, 0)
        Me.FilterCombo5.Margin = New System.Windows.Forms.Padding(2)
        Me.FilterCombo5.Name = "FilterCombo5"
        Me.FilterCombo5.Size = New System.Drawing.Size(64, 25)
        Me.FilterCombo5.TabIndex = 36
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(441, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(81, 20)
        Me.Label6.TabIndex = 33
        Me.Label6.Text = "Raised By"
        '
        'FilterCombo4
        '
        Me.FilterCombo4.Dock = System.Windows.Forms.DockStyle.Right
        Me.FilterCombo4.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FilterCombo4.FormattingEnabled = True
        Me.FilterCombo4.Location = New System.Drawing.Point(522, 0)
        Me.FilterCombo4.Margin = New System.Windows.Forms.Padding(2)
        Me.FilterCombo4.Name = "FilterCombo4"
        Me.FilterCombo4.Size = New System.Drawing.Size(87, 25)
        Me.FilterCombo4.TabIndex = 34
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(609, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 20)
        Me.Label3.TabIndex = 23
        Me.Label3.Text = "Priority"
        '
        'FilterCombo30
        '
        Me.FilterCombo30.Dock = System.Windows.Forms.DockStyle.Right
        Me.FilterCombo30.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FilterCombo30.FormattingEnabled = True
        Me.FilterCombo30.Location = New System.Drawing.Point(665, 0)
        Me.FilterCombo30.Margin = New System.Windows.Forms.Padding(2)
        Me.FilterCombo30.Name = "FilterCombo30"
        Me.FilterCombo30.Size = New System.Drawing.Size(96, 25)
        Me.FilterCombo30.TabIndex = 24
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(761, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(63, 20)
        Me.Label5.TabIndex = 25
        Me.Label5.Text = "Subject"
        '
        'FilterCombo90
        '
        Me.FilterCombo90.Dock = System.Windows.Forms.DockStyle.Right
        Me.FilterCombo90.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FilterCombo90.FormattingEnabled = True
        Me.FilterCombo90.Location = New System.Drawing.Point(824, 0)
        Me.FilterCombo90.Margin = New System.Windows.Forms.Padding(2)
        Me.FilterCombo90.Name = "FilterCombo90"
        Me.FilterCombo90.Size = New System.Drawing.Size(122, 25)
        Me.FilterCombo90.TabIndex = 26
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(946, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(37, 20)
        Me.Label1.TabIndex = 27
        Me.Label1.Text = "Site"
        '
        'FilterCombo1
        '
        Me.FilterCombo1.Dock = System.Windows.Forms.DockStyle.Right
        Me.FilterCombo1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FilterCombo1.FormattingEnabled = True
        Me.FilterCombo1.Location = New System.Drawing.Point(983, 0)
        Me.FilterCombo1.Margin = New System.Windows.Forms.Padding(2)
        Me.FilterCombo1.Name = "FilterCombo1"
        Me.FilterCombo1.Size = New System.Drawing.Size(112, 25)
        Me.FilterCombo1.TabIndex = 28
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(1095, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(54, 20)
        Me.Label2.TabIndex = 29
        Me.Label2.Text = "Group"
        '
        'FilterCombo2
        '
        Me.FilterCombo2.Dock = System.Windows.Forms.DockStyle.Right
        Me.FilterCombo2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FilterCombo2.FormattingEnabled = True
        Me.FilterCombo2.Location = New System.Drawing.Point(1149, 0)
        Me.FilterCombo2.Margin = New System.Windows.Forms.Padding(2)
        Me.FilterCombo2.Name = "FilterCombo2"
        Me.FilterCombo2.Size = New System.Drawing.Size(107, 25)
        Me.FilterCombo2.TabIndex = 30
        '
        'StaffQueryGrid
        '
        Me.StaffQueryGrid.AllowUserToAddRows = False
        Me.StaffQueryGrid.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.Gainsboro
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.GradientActiveCaption
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.Black
        Me.StaffQueryGrid.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.StaffQueryGrid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.StaffQueryGrid.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable
        Me.StaffQueryGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.StaffQueryGrid.DefaultCellStyle = DataGridViewCellStyle2
        Me.StaffQueryGrid.Dock = System.Windows.Forms.DockStyle.Fill
        Me.StaffQueryGrid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.StaffQueryGrid.Location = New System.Drawing.Point(0, 0)
        Me.StaffQueryGrid.MultiSelect = False
        Me.StaffQueryGrid.Name = "StaffQueryGrid"
        Me.StaffQueryGrid.ReadOnly = True
        Me.StaffQueryGrid.RowHeadersVisible = False
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.White
        DataGridViewCellStyle3.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.GradientActiveCaption
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.StaffQueryGrid.RowsDefaultCellStyle = DataGridViewCellStyle3
        Me.StaffQueryGrid.RowTemplate.DefaultCellStyle.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.StaffQueryGrid.RowTemplate.Height = 40
        Me.StaffQueryGrid.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.StaffQueryGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.StaffQueryGrid.Size = New System.Drawing.Size(1256, 445)
        Me.StaffQueryGrid.TabIndex = 2
        Me.StaffQueryGrid.TabStop = False
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 60000
        '
        'Timer2
        '
        Me.Timer2.Enabled = True
        '
        'ResponseView
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1256, 484)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "ResponseView"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ResponseView"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        CType(Me.StaffQueryGrid, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SplitContainer1 As SplitContainer
    Friend WithEvents StaffQueryGrid As DataGridView
    Friend WithEvents Label3 As Label
    Friend WithEvents FilterCombo30 As TemplateDB.FilterCombo
    Friend WithEvents Label5 As Label
    Friend WithEvents FilterCombo90 As TemplateDB.FilterCombo
    Friend WithEvents Label1 As Label
    Friend WithEvents FilterCombo1 As TemplateDB.FilterCombo
    Friend WithEvents Label2 As Label
    Friend WithEvents FilterCombo2 As TemplateDB.FilterCombo
    Friend WithEvents Label4 As Label
    Friend WithEvents FilterCombo3 As TemplateDB.FilterCombo
    Friend WithEvents BindingSource1 As BindingSource
    Friend WithEvents Timer1 As Timer
    Friend WithEvents Label6 As Label
    Friend WithEvents FilterCombo4 As TemplateDB.FilterCombo
    Friend WithEvents Label7 As Label
    Friend WithEvents FilterCombo5 As TemplateDB.FilterCombo
    Friend WithEvents Timer2 As Timer
    Friend WithEvents Label8 As Label
    Friend WithEvents Splitter1 As Splitter
End Class
