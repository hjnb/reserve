Imports System.Data.OleDb
Imports System.Runtime.InteropServices

Public Class 産健ｾﾝﾀｰ扱い

    Public DB1 As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\yoshi\Desktop\Reserve.mdb"

    Private Sub 産健ｾﾝﾀｰ扱い_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        Form1.f_sanken = Nothing
    End Sub

    Private Sub 産健ｾﾝﾀｰ扱い_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'リスト表示
        displaySankenList()
    End Sub

    Private Sub displaySankenList()
        Dim Cn As New OleDbConnection(DB1)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable

        SQLCm.CommandText = "SELECT distinct Ind, Sanken FROM RsvD ORDER BY Ind ASC"
        Adapter.Fill(Table)

        '▼値の表示
        sankenDataGridView.DataSource = Table

        '▼後処理

        Table.Dispose()
        Adapter.Dispose()
        SQLCm.Dispose()
        Cn.Dispose()

        'セルの編集不可
        sankenDataGridView.Columns("Ind").ReadOnly = True

        'DataGridViewでセル、行、列が複数選択されないようにする
        sankenDataGridView.MultiSelect = False

        For Each c As DataGridViewColumn In sankenDataGridView.Columns
            c.SortMode = DataGridViewColumnSortMode.NotSortable
        Next c

        sankenDataGridView.AllowUserToAddRows = False

        sankenDataGridView.Columns("Ind").HeaderText = "企業名"
        sankenDataGridView.Columns("Ind").Width = 165
        sankenDataGridView.Columns("Ind").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        
        sankenDataGridView.Columns("Sanken").HeaderText = "該当"
        sankenDataGridView.Columns("Sanken").Width = 40
        sankenDataGridView.Columns("Sanken").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

    End Sub

    Private Sub sankenDataGridView_CellPainting(ByVal sender As Object, _
        ByVal e As DataGridViewCellPaintingEventArgs) _
        Handles sankenDataGridView.CellPainting
        '列ヘッダーかどうか調べる
        If e.ColumnIndex < 0 And e.RowIndex >= 0 Then
            'セルを描画する
            e.Paint(e.ClipBounds, DataGridViewPaintParts.All)

            '行番号を描画する範囲を決定する
            'e.AdvancedBorderStyleやe.CellStyle.Paddingは無視しています
            Dim indexRect As Rectangle = e.CellBounds
            indexRect.Inflate(-2, -2)
            '行番号を描画する
            TextRenderer.DrawText(e.Graphics, _
                (e.RowIndex + 1).ToString(), _
                e.CellStyle.Font, _
                indexRect, _
                e.CellStyle.ForeColor, _
                TextFormatFlags.Right Or TextFormatFlags.VerticalCenter)
            '描画が完了したことを知らせる
            e.Handled = True
        End If
    End Sub
End Class