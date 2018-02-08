Public Class Form1

    Public f_yoyaku As 予約データ
    Public f_search As searchForm
    Public f_sanken As 産健ｾﾝﾀｰ扱い

    Public eraTable = New Dictionary(Of Integer, String)

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Maximized
        createEraTable()

        f_yoyaku = New 予約データ()
        f_yoyaku.Owner = Me
        f_yoyaku.Show()

    End Sub

    ''' <summary>
    ''' 和暦年号テーブル作成
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub createEraTable()
        Dim ci As New System.Globalization.CultureInfo("ja-JP", False)
        ci.DateTimeFormat.Calendar = New System.Globalization.JapaneseCalendar()
        For code = AscW("A"c) To AscW("Z"c)
            Dim e As String = ChrW(code)
            Dim eraIndex As Integer = ci.DateTimeFormat.GetEra(e)
            If (eraIndex > 0) Then
                eraTable.Add(eraIndex, e)
            End If
        Next
    End Sub

    Private Sub 終了ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 終了ToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub 予約データToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 予約データToolStripMenuItem.Click
        If f_yoyaku Is Nothing Then
            f_yoyaku = New 予約データ()
            f_yoyaku.Owner = Me
            f_yoyaku.Show()
        End If
    End Sub

    Private Sub 検索ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 検索ToolStripMenuItem.Click
        If f_search Is Nothing Then
            f_search = New searchForm(eraTable)
            f_search.Owner = Me
            f_search.Show()
        End If
    End Sub

    Private Sub 産健センター扱いToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles 産健センター扱いToolStripMenuItem.Click
        If f_sanken Is Nothing Then
            f_sanken = New 産健ｾﾝﾀｰ扱い()
            f_sanken.Owner = Me
            f_sanken.Show()
        End If
    End Sub
End Class
