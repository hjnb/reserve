<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class searchForm
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
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

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.searchTextBox = New System.Windows.Forms.TextBox()
        Me.btnSearch = New System.Windows.Forms.Button()
        Me.searchDataGridView = New System.Windows.Forms.DataGridView()
        Me.startDateLabel = New System.Windows.Forms.Label()
        Me.endDateLabel = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.searchDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'searchTextBox
        '
        Me.searchTextBox.ImeMode = System.Windows.Forms.ImeMode.KatakanaHalf
        Me.searchTextBox.Location = New System.Drawing.Point(30, 12)
        Me.searchTextBox.Name = "searchTextBox"
        Me.searchTextBox.Size = New System.Drawing.Size(140, 19)
        Me.searchTextBox.TabIndex = 0
        '
        'btnSearch
        '
        Me.btnSearch.Location = New System.Drawing.Point(380, 10)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(75, 23)
        Me.btnSearch.TabIndex = 1
        Me.btnSearch.Text = "検索"
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'searchDataGridView
        '
        Me.searchDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.searchDataGridView.Location = New System.Drawing.Point(12, 51)
        Me.searchDataGridView.Name = "searchDataGridView"
        Me.searchDataGridView.RowTemplate.Height = 21
        Me.searchDataGridView.Size = New System.Drawing.Size(500, 168)
        Me.searchDataGridView.TabIndex = 2
        '
        'startDateLabel
        '
        Me.startDateLabel.AutoSize = True
        Me.startDateLabel.Location = New System.Drawing.Point(194, 15)
        Me.startDateLabel.Name = "startDateLabel"
        Me.startDateLabel.Size = New System.Drawing.Size(61, 12)
        Me.startDateLabel.TabIndex = 3
        Me.startDateLabel.Text = "H25/05/23"
        '
        'endDateLabel
        '
        Me.endDateLabel.AutoSize = True
        Me.endDateLabel.Location = New System.Drawing.Point(295, 15)
        Me.endDateLabel.Name = "endDateLabel"
        Me.endDateLabel.Size = New System.Drawing.Size(61, 12)
        Me.endDateLabel.TabIndex = 4
        Me.endDateLabel.Text = "H29/10/02"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(268, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(17, 12)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "～"
        '
        'searchForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(527, 233)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.endDateLabel)
        Me.Controls.Add(Me.startDateLabel)
        Me.Controls.Add(Me.searchDataGridView)
        Me.Controls.Add(Me.btnSearch)
        Me.Controls.Add(Me.searchTextBox)
        Me.KeyPreview = True
        Me.Name = "searchForm"
        Me.Text = "検索"
        CType(Me.searchDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents searchTextBox As System.Windows.Forms.TextBox
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents searchDataGridView As System.Windows.Forms.DataGridView
    Friend WithEvents startDateLabel As System.Windows.Forms.Label
    Friend WithEvents endDateLabel As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
