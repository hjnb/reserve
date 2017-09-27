<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.予約データToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.産健センター扱いToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.備忘六ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.検索ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.印刷設定ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.終了ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.予約データToolStripMenuItem, Me.産健センター扱いToolStripMenuItem, Me.備忘六ToolStripMenuItem, Me.検索ToolStripMenuItem, Me.印刷設定ToolStripMenuItem, Me.終了ToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(554, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        '予約データToolStripMenuItem
        '
        Me.予約データToolStripMenuItem.Name = "予約データToolStripMenuItem"
        Me.予約データToolStripMenuItem.Size = New System.Drawing.Size(69, 20)
        Me.予約データToolStripMenuItem.Text = "予約データ"
        '
        '産健センター扱いToolStripMenuItem
        '
        Me.産健センター扱いToolStripMenuItem.Name = "産健センター扱いToolStripMenuItem"
        Me.産健センター扱いToolStripMenuItem.Size = New System.Drawing.Size(101, 20)
        Me.産健センター扱いToolStripMenuItem.Text = "産健センター扱い"
        '
        '備忘六ToolStripMenuItem
        '
        Me.備忘六ToolStripMenuItem.Name = "備忘六ToolStripMenuItem"
        Me.備忘六ToolStripMenuItem.Size = New System.Drawing.Size(55, 20)
        Me.備忘六ToolStripMenuItem.Text = "備忘録"
        '
        '検索ToolStripMenuItem
        '
        Me.検索ToolStripMenuItem.Name = "検索ToolStripMenuItem"
        Me.検索ToolStripMenuItem.Size = New System.Drawing.Size(43, 20)
        Me.検索ToolStripMenuItem.Text = "検索"
        '
        '印刷設定ToolStripMenuItem
        '
        Me.印刷設定ToolStripMenuItem.Name = "印刷設定ToolStripMenuItem"
        Me.印刷設定ToolStripMenuItem.Size = New System.Drawing.Size(67, 20)
        Me.印刷設定ToolStripMenuItem.Text = "印刷設定"
        '
        '終了ToolStripMenuItem
        '
        Me.終了ToolStripMenuItem.Name = "終了ToolStripMenuItem"
        Me.終了ToolStripMenuItem.Size = New System.Drawing.Size(43, 20)
        Me.終了ToolStripMenuItem.Text = "終了"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(554, 345)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents 予約データToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 産健センター扱いToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 備忘六ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 検索ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 印刷設定ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents 終了ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

End Class
