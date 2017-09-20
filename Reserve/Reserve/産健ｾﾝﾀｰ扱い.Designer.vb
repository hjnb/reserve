<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 産健ｾﾝﾀｰ扱い
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
        Me.sankenDataGridView = New System.Windows.Forms.DataGridView()
        Me.btnRegist = New System.Windows.Forms.Button()
        CType(Me.sankenDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'sankenDataGridView
        '
        Me.sankenDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.sankenDataGridView.Location = New System.Drawing.Point(22, 23)
        Me.sankenDataGridView.Name = "sankenDataGridView"
        Me.sankenDataGridView.RowTemplate.Height = 21
        Me.sankenDataGridView.Size = New System.Drawing.Size(265, 442)
        Me.sankenDataGridView.TabIndex = 0
        '
        'btnRegist
        '
        Me.btnRegist.Location = New System.Drawing.Point(112, 487)
        Me.btnRegist.Name = "btnRegist"
        Me.btnRegist.Size = New System.Drawing.Size(75, 40)
        Me.btnRegist.TabIndex = 1
        Me.btnRegist.Text = "登録"
        Me.btnRegist.UseVisualStyleBackColor = True
        '
        '産健ｾﾝﾀｰ扱い
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(313, 541)
        Me.Controls.Add(Me.btnRegist)
        Me.Controls.Add(Me.sankenDataGridView)
        Me.Name = "産健ｾﾝﾀｰ扱い"
        Me.Text = "産健ｾﾝﾀｰ扱い"
        CType(Me.sankenDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents sankenDataGridView As System.Windows.Forms.DataGridView
    Friend WithEvents btnRegist As System.Windows.Forms.Button
End Class
