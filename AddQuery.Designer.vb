<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AddQuery
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
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AddQuery))
        Me.SplitContainer3 = New System.Windows.Forms.SplitContainer()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.NewQueryGrid = New System.Windows.Forms.DataGridView()
        Me.BindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
        CType(Me.SplitContainer3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer3.Panel1.SuspendLayout()
        Me.SplitContainer3.Panel2.SuspendLayout()
        Me.SplitContainer3.SuspendLayout()
        CType(Me.NewQueryGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SplitContainer3
        '
        Me.SplitContainer3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer3.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer3.Name = "SplitContainer3"
        Me.SplitContainer3.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer3.Panel1
        '
        Me.SplitContainer3.Panel1.Controls.Add(Me.Button3)
        Me.SplitContainer3.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer3.Panel1.Controls.Add(Me.ComboBox2)
        '
        'SplitContainer3.Panel2
        '
        Me.SplitContainer3.Panel2.Controls.Add(Me.NewQueryGrid)
        Me.SplitContainer3.Size = New System.Drawing.Size(509, 331)
        Me.SplitContainer3.SplitterDistance = 25
        Me.SplitContainer3.TabIndex = 3
        '
        'Button3
        '
        Me.Button3.BackgroundImage = Global.QueryTool.My.Resources.Resources.SAVER
        Me.Button3.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Button3.Dock = System.Windows.Forms.DockStyle.Left
        Me.Button3.FlatAppearance.BorderSize = 0
        Me.Button3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button3.Location = New System.Drawing.Point(0, 0)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(130, 25)
        Me.Button3.TabIndex = 8
        Me.Button3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Dock = System.Windows.Forms.DockStyle.Right
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(294, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(54, 20)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Study:"
        '
        'ComboBox2
        '
        Me.ComboBox2.Dock = System.Windows.Forms.DockStyle.Right
        Me.ComboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(348, 0)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ComboBox2.Size = New System.Drawing.Size(161, 24)
        Me.ComboBox2.TabIndex = 0
        '
        'NewQueryGrid
        '
        Me.NewQueryGrid.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.Gainsboro
        Me.NewQueryGrid.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.NewQueryGrid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.NewQueryGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.NewQueryGrid.Dock = System.Windows.Forms.DockStyle.Fill
        Me.NewQueryGrid.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.NewQueryGrid.Location = New System.Drawing.Point(0, 0)
        Me.NewQueryGrid.Name = "NewQueryGrid"
        Me.NewQueryGrid.RowHeadersVisible = False
        Me.NewQueryGrid.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.NewQueryGrid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.NewQueryGrid.Size = New System.Drawing.Size(509, 302)
        Me.NewQueryGrid.TabIndex = 1
        '
        'AddQuery
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(509, 331)
        Me.Controls.Add(Me.SplitContainer3)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "AddQuery"
        Me.Text = "Query Tool"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.SplitContainer3.Panel1.ResumeLayout(False)
        Me.SplitContainer3.Panel1.PerformLayout()
        Me.SplitContainer3.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer3.ResumeLayout(False)
        CType(Me.NewQueryGrid, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.BindingSource1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer3 As System.Windows.Forms.SplitContainer
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Friend WithEvents NewQueryGrid As System.Windows.Forms.DataGridView
    Friend WithEvents BindingSource1 As System.Windows.Forms.BindingSource
End Class
