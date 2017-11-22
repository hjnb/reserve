Imports System.Data.OleDb
Imports System.Runtime.InteropServices

Public Class searchForm

    Public DB1 As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\yoshi\Desktop\Reserve.mdb"
    'Public DB1 As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\Primergytx100s1\Reserve\Reserve.mdb"

    Private Sub searchForm_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        Form1.f_search = Nothing
    End Sub

    Private Sub searchForm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            If e.Control = False Then
                Me.SelectNextControl(Me.ActiveControl, Not e.Shift, True, True, True)
            End If
        End If
    End Sub

    Private Sub searchForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim startDate As String = ""
        Dim endDate As String = ""

        Dim Cn As New OleDbConnection(DB1)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim reader As System.Data.OleDb.OleDbDataReader

        '位置
        Me.Left = 800
        Me.Top = 50

        '最新の予約日を取得
        SQLCm.CommandText = "SELECT TOP 1 * FROM RsvD ORDER BY Ymd DESC"
        Cn.Open()
        reader = SQLCm.ExecuteReader()

        While reader.Read() = True
            endDate = reader("Ymd")
        End While
        reader.Close()

        endDateLabel.Text = endDate

        '最古の予約日を取得
        SQLCm.CommandText = "SELECT TOP 1 * FROM RsvD ORDER BY Ymd"
        reader = SQLCm.ExecuteReader()

        While reader.Read() = True
            startDate = reader("Ymd")
        End While

        startDateLabel.Text = startDate

        Cn.Close()
        SQLCm.Dispose()
        Cn.Dispose()

    End Sub

    Private Sub searchDataGridView_CellPainting(ByVal sender As Object, _
        ByVal e As DataGridViewCellPaintingEventArgs) _
        Handles searchDataGridView.CellPainting
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

    Private Sub displaySearchList(ByVal searchText As String)
        Dim Cn As New OleDbConnection(DB1)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable

        SQLCm.CommandText = "SELECT Nam, Kana, Birth, Tok1, Futan, Ymd, Syu FROM RsvD WHERE Kana LIKE '%" & searchText & "%' ORDER BY Ymd DESC, Kana ASC"
        Adapter.Fill(Table)

        '▼値の表示
        searchDataGridView.DataSource = Table

        '▼後処理

        Table.Dispose()
        Adapter.Dispose()
        SQLCm.Dispose()
        Cn.Dispose()

        searchDataGridView.Columns("Nam").HeaderText = "氏名"
        searchDataGridView.Columns("Nam").Width = 80
        searchDataGridView.Columns("Nam").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        searchDataGridView.Columns("Kana").HeaderText = "カナ"
        searchDataGridView.Columns("Kana").Width = 80
        searchDataGridView.Columns("Kana").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        
        searchDataGridView.Columns("Birth").HeaderText = "生年月日"
        searchDataGridView.Columns("Birth").Width = 70
        searchDataGridView.Columns("Birth").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        searchDataGridView.Columns("Birth").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        searchDataGridView.Columns("Tok1").HeaderText = "保険"
        searchDataGridView.Columns("Tok1").Width = 40
        searchDataGridView.Columns("Tok1").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        searchDataGridView.Columns("Tok1").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        searchDataGridView.Columns("Futan").HeaderText = "窓口負担"
        searchDataGridView.Columns("Futan").Width = 60
        searchDataGridView.Columns("Futan").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        searchDataGridView.Columns("Futan").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight

        searchDataGridView.Columns("Ymd").HeaderText = "予約日"
        searchDataGridView.Columns("Ymd").Width = 70
        searchDataGridView.Columns("Ymd").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        searchDataGridView.Columns("Ymd").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        searchDataGridView.Columns("Syu").HeaderText = "種別"
        searchDataGridView.Columns("Syu").Width = 40
        searchDataGridView.Columns("Syu").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        searchDataGridView.Columns("Syu").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        'セルの編集不可
        searchDataGridView.ReadOnly = True

        'DataGridViewでセル、行、列が複数選択されないようにする
        searchDataGridView.MultiSelect = False

        'セルを選択すると行全体が選択されるようにする
        searchDataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        For Each c As DataGridViewColumn In searchDataGridView.Columns
            c.SortMode = DataGridViewColumnSortMode.NotSortable
        Next c

        searchDataGridView.AllowUserToAddRows = False
    End Sub

    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Dim searchStr As String = searchTextBox.Text
        If searchStr = "" Then
            MsgBox("検索文字列を入力してください")
        Else
            displaySearchList(searchStr)
            convertJapanCalender("Birth")
            convertJapanCalender("Ymd")
        End If

    End Sub

    Private Sub convertJapanCalender(columnName As String)
        '生年月日を和暦で表示
        ' JapaneseCalendarクラスのインスタンスを作る
        Dim calendarJp = New System.Globalization.JapaneseCalendar()
        Dim tmpStr As String

        Dim ci As New System.Globalization.CultureInfo("ja-JP", False)
        ci.DateTimeFormat.Calendar = New System.Globalization.JapaneseCalendar()
        Dim rowsCount As Integer = searchDataGridView.Rows.Count
        Dim dt As DateTime
        For i = 0 To rowsCount - 1
            If searchDataGridView(columnName, i).Value Is Nothing Then
                Continue For
            End If
            dt = searchDataGridView(columnName, i).Value
            tmpStr = dt.ToString("gyy/MM/dd", ci)
            If tmpStr.Substring(0, 2) = "平成" Then
                searchDataGridView(columnName, i).Value = tmpStr.Replace("平成", "H")
            ElseIf tmpStr.Substring(0, 2) = "昭和" Then
                searchDataGridView(columnName, i).Value = tmpStr.Replace("昭和", "S")
            ElseIf tmpStr.Substring(0, 2) = "大正" Then
                searchDataGridView(columnName, i).Value = tmpStr.Replace("大正", "T")
            ElseIf tmpStr.Substring(0, 2) = "明治" Then
                searchDataGridView(columnName, i).Value = tmpStr.Replace("明治", "M")
            End If
        Next
    End Sub
End Class