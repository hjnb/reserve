Imports System.Data.OleDb
Imports System.Runtime.InteropServices

Public Class 産健ｾﾝﾀｰ扱い

    Public DB1 As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\yoshi\Desktop\Reserve.mdb"
    'Public DB1 As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\Primergytx100s1\Reserve\Reserve.mdb"

    Public preData As New List(Of SankenData)

    Public Class SankenData
        Private sankenName As String

        Private sankenFlg As String

        Public Sub New(ByVal sankenName As String, ByVal sankenFlg As String)
            Me.sankenName = sankenName
            Me.sankenFlg = sankenFlg
        End Sub

        Public Function getSankenName() As String
            Return sankenName
        End Function

        Public Function getSankenFlg() As String
            Return sankenFlg
        End Function

    End Class

    Private Sub 産健ｾﾝﾀｰ扱い_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        Form1.f_sanken = Nothing
    End Sub

    Private Sub 産健ｾﾝﾀｰ扱い_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'ダブルバッファリングを有効
        予約データ.EnableDoubleBuffering(sankenDataGridView)

        '位置
        Me.Left = 800
        Me.Top = 50

        'リスト表示
        displaySankenList()

        '初期状態を保持
        Dim rowsCount As Integer = sankenDataGridView.Rows.Count
        For i = 0 To rowsCount - 1
            preData.Add(New SankenData(sankenDataGridView(0, i).Value, sankenDataGridView(1, i).Value))
        Next

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

        '重複表示の制御処理
        Dim rowsCount As Integer = sankenDataGridView.Rows.Count
        For i = 0 To rowsCount - 1
            If i <> 0 AndAlso sankenDataGridView(0, i).Value = sankenDataGridView(0, i - 1).Value Then
                sankenDataGridView.Rows.RemoveAt(i - 1)
                rowsCount -= 1
                i -= 1
            End If
            If i = rowsCount - 1 Then
                Exit For
            End If
        Next

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

    Private Sub btnRegist_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRegist.Click
        Dim updateIndexList As New List(Of Integer)
        Dim nowData As New List(Of SankenData)
        Dim rowsCount As Integer = sankenDataGridView.Rows.Count
        For i = 0 To rowsCount - 1
            nowData.Add(New SankenData(sankenDataGridView(0, i).Value, If(IsDBNull(sankenDataGridView(1, i).Value), "", sankenDataGridView(1, i).Value)))
            If nowData(i).getSankenName = preData(i).getSankenName AndAlso nowData(i).getSankenFlg <> preData(i).getSankenFlg Then
                '初期状態と比較し異なる場合は対象のインデックスをリストに追加
                updateIndexList.Add(i)
            End If
        Next

        If updateIndexList.Count = 0 Then
            MsgBox("変更がありません")
        Else
            '更新処理
            updateSankenData(nowData, updateIndexList)

            MsgBox(updateIndexList.Count & "件、変更しました")
            '再表示
            displaySankenList()

            '初期状態の更新
            preData.Clear()
            For i = 0 To rowsCount - 1
                preData.Add(New SankenData(nowData(i).getSankenName, nowData(i).getSankenFlg))
            Next
        End If

    End Sub

    Private Sub updateSankenData(ByVal nowSankenDataList As List(Of SankenData), ByVal updateIndexList As List(Of Integer))
        Dim Cn As New OleDbConnection(DB1)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand

        Cn.Open()
        For Each i As Integer In updateIndexList
            SQLCm.CommandText = "Update RsvD SET Sanken='" & nowSankenDataList(i).getSankenFlg & "' WHERE Ind='" & nowSankenDataList(i).getSankenName & "'"
            SQLCm.ExecuteNonQuery()
        Next
        SQLCm.Dispose()
        Cn.Close()
    End Sub
End Class