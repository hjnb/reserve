Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core

Public Class 予約データ

    Public DB_reserve As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\yoshi\Desktop\Reserve.mdb"
    Public DB_diagnose As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\yoshi\Desktop\Diagnose.mdb"
    Public DB_health As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\yoshi\Desktop\Health3.mdb"

    Public initFlg As Boolean = True

    Private Sub 予約データ_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        Form1.f_yoyaku = Nothing
    End Sub

    Private Sub 予約データ_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        '一覧表示の初期設定
        initialSetting4DataGridView()

        'セルの編集不可
        DataGridView1.ReadOnly = True

        'DataGridView1でセル、行、列が複数選択されないようにする
        'DataGridView1.MultiSelect = False

        'セルを選択すると行全体が選択されるようにする
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        For Each c As DataGridViewColumn In DataGridView1.Columns
            c.SortMode = DataGridViewColumnSortMode.NotSortable
        Next c

        DataGridView1.AllowUserToAddRows = False

        TabControl1.SizeMode = TabSizeMode.Fixed
        TabControl1.ItemSize = New Size(65, 25)
        TabControl1.SelectedTab = referenceTabPage



    End Sub

    Private Sub displayDiagnose()
        referenceListBox.Items.Clear()

        Dim Cn As New OleDbConnection(DB_diagnose)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim reader As System.Data.OleDb.OleDbDataReader

        SQLCm.CommandText = "SELECT Ind FROM IndM ORDER BY Kana"
        Cn.Open()
        reader = SQLCm.ExecuteReader()

        While reader.Read() = True
            referenceListBox.Items.Add(reader("Ind"))
        End While
        reader.Close()
        Cn.Close()
        SQLCm.Dispose()
        Cn.Dispose()
    End Sub

    Private Sub displayHealth()
        referenceListBox.Items.Clear()

        Dim Cn As New OleDbConnection(DB_health)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim reader As System.Data.OleDb.OleDbDataReader

        SQLCm.CommandText = "SELECT Ind FROM IndM ORDER BY Kana"
        Cn.Open()
        reader = SQLCm.ExecuteReader()

        While reader.Read() = True
            referenceListBox.Items.Add(reader("Ind"))
        End While
        reader.Close()
        Cn.Close()
        SQLCm.Dispose()
        Cn.Dispose()
    End Sub

    Private Sub displaySankenCenter()
        referenceListBox.Items.Clear()

        Dim Cn As New OleDbConnection(DB_reserve)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim reader As System.Data.OleDb.OleDbDataReader

        SQLCm.CommandText = "SELECT distinct Ind FROM RsvD WHERE Sanken='*' ORDER BY Ind"
        Cn.Open()
        reader = SQLCm.ExecuteReader()

        While reader.Read() = True
            referenceListBox.Items.Add(reader("Ind"))
        End While
        reader.Close()
        Cn.Close()
        SQLCm.Dispose()
        Cn.Dispose()
    End Sub

    Private Sub DataGridView1_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        If DataGridView1.CurrentRow Is Nothing Then
            Return
        End If

        Dim type, companyName, name, kana, sex, birthDay, reserveDay, ampm, resultDay, post, memo1, memo2, windowPay As String
        Dim rowIndex As Integer = DataGridView1.CurrentRow.Index

        '選択した行の値を取得
        type = DataGridView1("Syu", rowIndex).Value
        companyName = DataGridView1("Ind", rowIndex).Value
        name = DataGridView1("Nam", rowIndex).Value
        kana = DataGridView1("Kana", rowIndex).Value
        sex = DataGridView1("Sex", rowIndex).Value
        birthDay = DataGridView1("Birth", rowIndex).Value
        reserveDay = DataGridView1("Ymd", rowIndex).Value
        ampm = DataGridView1("Apm", rowIndex).Value
        resultDay = DataGridView1("Ymd2", rowIndex).Value
        post = DataGridView1("Send", rowIndex).Value
        memo1 = DataGridView1("Memo1", rowIndex).Value
        memo2 = DataGridView1("Memo2", rowIndex).Value
        windowPay = DataGridView1("Futan", rowIndex).Value

        'テキストボックスへ反映
        typeBox.Text = type
        companyNameBox.Text = companyName
        nameBox.Text = name
        kanaBox.Text = kana
        sexBox.Text = sex
        birthDayBox.Text = birthDay
        resultDayBox.Text = resultDay

        'ampmの表示処理
        ampmBox.Text = ampm

        '予約日の表示処理(和暦で表示させる)
        reserveDateBox.Text = convertADToWareki(reserveDay.Substring(0, 4)) & reserveDay.Substring(4, 6)

        postBox.Text = post
        memo1Box.Text = memo1
        memo2Box.Text = memo2


        'タブ切り替え
        If type = "個人" Then
            TabControl1.SelectedTab = personalTabPage
        ElseIf type = "企業" Then
            TabControl1.SelectedTab = companyTabPage
        ElseIf type = "生活" Then
            TabControl1.SelectedTab = lifeStyleTabPage
        ElseIf type = "特定" Then
            TabControl1.SelectedTab = specificTabPage
        ElseIf type = "がん" Then
            TabControl1.SelectedTab = cancerTabPage
        End If

        'チェックボックス、ラジオボタンへの反映処理
        '個人
        '血液
        If DataGridView1("Kjn1", rowIndex).Value = 1 Then
            personalBlood.Checked = True
        Else
            personalBlood.Checked = False
        End If
        '心電図
        If DataGridView1("Kig2", rowIndex).Value = 1 Then
            personalElectro.Checked = True
        Else
            personalElectro.Checked = False
        End If
        '胸部XP
        If DataGridView1("Kig3", rowIndex).Value = 1 Then
            personalChestXP.Checked = True
        Else
            personalChestXP.Checked = False
        End If
        '超音波
        If DataGridView1("Kig4", rowIndex).Value = 1 Then
            personalUltrasonic.Checked = True
        Else
            personalUltrasonic.Checked = False
        End If
        '胃バリウム
        If DataGridView1("Kig5", rowIndex).Value = 1 Then
            personalStomachBa.Checked = True
        Else
            personalStomachBa.Checked = False
        End If
        '胃カメラ
        If DataGridView1("Kig6", rowIndex).Value = 1 Then
            personalStomachCamera.Checked = True
        Else
            personalStomachCamera.Checked = False
        End If
        '窓口負担
        personalWindowPay.Text = windowPay


        '企業
        '血液
        If DataGridView1("Kig1", rowIndex).Value = 1 Then
            companyBlood.Checked = True
        Else
            companyBlood.Checked = False
        End If
        '心電図
        If DataGridView1("Kig2", rowIndex).Value = 1 Then
            companyElectro.Checked = True
        Else
            companyElectro.Checked = False
        End If
        '胸部XP
        If DataGridView1("Kig3", rowIndex).Value = 1 Then
            companyChestXP.Checked = True
        Else
            companyChestXP.Checked = False
        End If
        '超音波
        If DataGridView1("Kig4", rowIndex).Value = 1 Then
            companyUltrasonic.Checked = True
        Else
            companyUltrasonic.Checked = False
        End If
        '胃バリウム
        If DataGridView1("Kig5", rowIndex).Value = 1 Then
            companyStomachBa.Checked = True
        Else
            companyStomachBa.Checked = False
        End If
        '胃カメラ
        If DataGridView1("Kig6", rowIndex).Value = 1 Then
            companyStomachCamera.Checked = True
        Else
            companyStomachCamera.Checked = False
        End If
        '窓口負担
        companyWindowPay.Text = windowPay


        '生活
        '胃バリウム
        If DataGridView1("Sei3", rowIndex).Value = 1 Then
            lifeStyleStomachBa.Checked = True
        Else
            lifeStyleStomachBa.Checked = False
        End If
        '胃カメラ
        If DataGridView1("Sei4", rowIndex).Value = 1 Then
            lifeStyleStomachCamera.Checked = True
        Else
            lifeStyleStomachCamera.Checked = False
        End If


        '特定
        '種別
        If DataGridView1("ToK1", rowIndex).Value <> "" Then
            insuranceTypeBox.Text = DataGridView1("ToK1", rowIndex).Value
        Else
            insuranceTypeBox.Text = ""
        End If
        '生化学
        If DataGridView1("ToK2", rowIndex).Value <> "" Then
            biochemistryBox.Text = DataGridView1("ToK2", rowIndex).Value
        Else
            biochemistryBox.Text = ""
        End If
        '血糖
        If DataGridView1("ToK3", rowIndex).Value <> "" Then
            bloodSugarBox.Text = DataGridView1("ToK3", rowIndex).Value
        Else
            bloodSugarBox.Text = ""
        End If
        '貧血
        If DataGridView1("ToK4", rowIndex).Value <> "" Then
            anemiaBox.Text = DataGridView1("ToK4", rowIndex).Value
        Else
            anemiaBox.Text = ""
        End If
        '心機能
        If DataGridView1("ToK5", rowIndex).Value <> "" Then
            cardiacBox.Text = DataGridView1("ToK5", rowIndex).Value
        Else
            cardiacBox.Text = ""
        End If
        '胃がんリスク
        If DataGridView1("ToK6", rowIndex).Value <> "" Then
            gastricCancerRiskBox.Text = DataGridView1("ToK6", rowIndex).Value
        Else
            gastricCancerRiskBox.Text = ""
        End If
        '糖尿病性腎症
        If DataGridView1("ToK7", rowIndex).Value <> "" Then
            gastricCancerRiskBox.Text = DataGridView1("ToK7", rowIndex).Value
        Else
            gastricCancerRiskBox.Text = ""
        End If
        '前立腺がん
        If DataGridView1("ToK8", rowIndex).Value <> "" Then
            prostateCancerBox.Text = DataGridView1("ToK8", rowIndex).Value
        Else
            prostateCancerBox.Text = ""
        End If
        '無料クーポン券
        If DataGridView1("ToK9", rowIndex).Value = 1 Then
            couponTicketBox.Checked = True
        Else
            couponTicketBox.Checked = False
        End If


        'がん
        '胃がん
        If DataGridView1("Gan1", rowIndex).Value = 1 Then
            gastricCancerBox.Checked = True
        Else
            gastricCancerBox.Checked = False
        End If
        '大腸がん
        If DataGridView1("Gan2", rowIndex).Value = 1 Then
            colorectalCancerBox.Checked = True
        Else
            colorectalCancerBox.Checked = False
        End If

        '参照

        




    End Sub

    Private Sub DataGridView1_CellPainting(ByVal sender As Object, _
        ByVal e As DataGridViewCellPaintingEventArgs) _
        Handles DataGridView1.CellPainting
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

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, _
        ByVal e As DataGridViewCellFormattingEventArgs) _
        Handles DataGridView1.CellFormatting

        If DataGridView1.Columns(e.ColumnIndex).Name = "Ymd" Then
            '予約日の表示設定とグループ化
            If e.RowIndex > 0 AndAlso DataGridView1(e.ColumnIndex, e.RowIndex - 1).Value = e.Value Then
                e.Value = ""
                e.FormattingApplied = True
            Else
                e.Value = Integer.Parse(e.Value.Substring(e.Value.ToString.Length - 2, 2))
            End If
        ElseIf DataGridView1.Columns(e.ColumnIndex).Name = "day" Then
            '曜日のグループ化
            If e.RowIndex > 0 AndAlso DataGridView1(e.ColumnIndex, e.RowIndex - 1).Value = e.Value Then
                e.Value = ""
                e.FormattingApplied = True
            End If
        ElseIf DataGridView1.Columns(e.ColumnIndex).Name = "Apm" Then
            '時間のグループ化
            If e.RowIndex > 0 AndAlso DataGridView1("day", e.RowIndex).Value = DataGridView1("day", e.RowIndex - 1).Value AndAlso DataGridView1(e.ColumnIndex, e.RowIndex - 1).Value = e.Value Then
                e.Value = ""
                e.FormattingApplied = True
            End If
        End If

    End Sub

    Private Sub convertJapanCalender()
        '生年月日を和暦で表示
        ' JapaneseCalendarクラスのインスタンスを作る
        Dim calendarJp = New System.Globalization.JapaneseCalendar()
        Dim tmpStr As String

        Dim ci As New System.Globalization.CultureInfo("ja-JP", False)
        ci.DateTimeFormat.Calendar = New System.Globalization.JapaneseCalendar()
        Dim rowsCount As Integer = DataGridView1.Rows.Count
        Dim dt As DateTime
        For i = 0 To rowsCount - 1
            If DataGridView1("Birth", i).Value Is Nothing Then
                Continue For
            End If
            dt = DataGridView1("Birth", i).Value
            tmpStr = dt.ToString("gyy/MM/dd", ci)
            If tmpStr.Substring(0, 2) = "平成" Then

                DataGridView1("Birth", i).Value = tmpStr.Replace("平成", "H")
            ElseIf tmpStr.Substring(0, 2) = "昭和" Then
                DataGridView1("Birth", i).Value = tmpStr.Replace("昭和", "S")
            ElseIf tmpStr.Substring(0, 2) = "大正" Then
                DataGridView1("Birth", i).Value = tmpStr.Replace("大正", "T")
            ElseIf tmpStr.Substring(0, 2) = "明治" Then
                DataGridView1("Birth", i).Value = tmpStr.Replace("明治", "M")
            End If
        Next
    End Sub

    Private Sub displayReserveList()
        Dim eraStr As String = eraBox.Text
        Dim monthStr As String = monthBox.Text
        Dim targetDateStr As String = convertWarekiToAD(eraStr) & "/" & monthStr

        Dim Cn As New OleDbConnection(DB_reserve)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable

        SQLCm.CommandText = "SELECT Ymd, Apm, Syu, Nam, Kana, Sex, Birth, Ind, Ymd2, Send, Memo1, Memo2, Futan, Kjn1, Kjn2, Kjn3, Kjn4, Kjn5, Kjn6, Kig1, Kig2, Kig3, Kig4, Kig5, Kig6, Sei1, Sei2, Sei3, Sei4, Tok1, Tok2, Tok3, Tok4, Tok5, Tok6, Tok7, Tok8, Tok9, Gan1, Gan2, Sanken FROM RsvD WHERE Ymd LIKE '%" & targetDateStr & "%' ORDER BY Ymd ASC, Apm ASC, Kana ASC"
        Adapter.Fill(Table)

        '▼値の表示
        DataGridView1.DataSource = Table

        '▼後処理

        Table.Dispose()
        Adapter.Dispose()
        SQLCm.Dispose()
        Cn.Dispose()

    End Sub

    Private Sub displayDayColumun()
        Dim rowsCount As Integer = DataGridView1.Rows.Count
        Dim year As Integer
        Dim month As Integer
        Dim day As Integer
        Dim dateTime As DateTime

        For i = 0 To rowsCount - 1
            If DataGridView1("Ymd", i).Value Is Nothing Then
                Continue For
            End If
            year = Integer.Parse(DataGridView1("Ymd", i).Value.ToString.Substring(0, 4))
            month = Integer.Parse(DataGridView1("Ymd", i).Value.ToString.Substring(5, 2))
            day = Integer.Parse(DataGridView1("Ymd", i).Value.ToString.Substring(8, 2))
            dateTime = New DateTime(year, month, day)

            DataGridView1("day", i).Value = dateTime.ToString("ddd")
        Next

    End Sub

    Private Sub displayAgeColumn()
        Dim rowsCount As Integer = DataGridView1.Rows.Count
        Dim birthStr As String
        Dim birthYear As String
        Dim birthMonth As String
        Dim birthDate As String
        Dim birthDateTime As DateTime
        Dim age As Integer
        Dim todayDateTime As DateTime = DateTime.Today

        For i = 0 To rowsCount - 1
            If DataGridView1("Birth", i).Value Is Nothing Then
                Continue For
            End If
            birthStr = DataGridView1("Birth", i).Value
            birthYear = convertWarekiToAD(birthStr.Substring(0, 3))
            birthMonth = birthStr.Substring(4, 2)
            birthDate = birthStr.Substring(7, 2)
            birthDateTime = New DateTime(birthYear, birthMonth, birthDate)
            age = GetAge(birthDateTime, todayDateTime)
            DataGridView1("age", i).Value = age
        Next

    End Sub

    Private Function GetAge(ByVal birthDate As DateTime, ByVal today As DateTime) As Integer
        Dim age As Integer = today.Year - birthDate.Year
        '誕生日がまだ来ていなければ、1引く
        If today.Month < birthDate.Month OrElse _
            (today.Month = birthDate.Month AndAlso _
                today.Day < birthDate.Day) Then
            age -= 1
        End If

        Return age
    End Function

    Private Function convertWarekiToAD(ByVal warekiStr As String) As String
        Dim ADStr As String = ""
        Dim initialStr As String = warekiStr.Substring(0, 1)
        Dim num As Integer = Integer.Parse(warekiStr.Substring(1, 2))

        If initialStr = "T" Then
            ADStr = 1911 + num
        ElseIf initialStr = "S" Then
            ADStr = 1925 + num
        ElseIf initialStr = "H" Then
            ADStr = 1988 + num
        End If

        Return ADStr
    End Function

    Public Shared Function convertADToWareki(ByVal adStr As String) As String
        Dim warekiStr As String = ""
        warekiStr = "H" & (Integer.Parse(adStr) - 1988)
        Return warekiStr
    End Function

    Private Sub addColumn()
        DataGridView1.Columns.Add("day", "曜日")
        DataGridView1.Columns("day").DisplayIndex = 1

        DataGridView1.Columns.Add("age", "年齢")
        DataGridView1.Columns("age").DisplayIndex = 8
    End Sub

    Private Sub initialSetting4DataGridView()

        '現在の年月を取得 
        Dim eraStr As String = convertADToWareki(DateTime.Today.ToString("yyyy"))
        Dim monthStr As String = DateTime.Today.ToString("MM")

        'コンボボックスに設定
        eraBox.Text = eraStr
        monthBox.Text = monthStr

        '一覧表示
        displayReserveList()

        '曜日、年齢の列を追加
        addColumn()

        '西暦を和暦表示に変更
        convertJapanCalender()

        '年齢の表示設定
        displayAgeColumn()

        '曜日の表示設定
        displayDayColumun()

        '列名、幅の設定
        '固定
        DataGridView1.Columns(3).Frozen = True

        DataGridView1.Columns(0).HeaderText = "予約日"
        DataGridView1.Columns(0).Width = 50
        DataGridView1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView1.Columns("day").HeaderText = "曜"
        DataGridView1.Columns("day").Width = 30
        DataGridView1.Columns("day").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("day").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView1.Columns(1).HeaderText = "AmPm"
        DataGridView1.Columns(1).Width = 50
        DataGridView1.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView1.Columns(2).HeaderText = "種別"
        DataGridView1.Columns(2).Width = 40
        DataGridView1.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView1.Columns(3).HeaderText = "氏名"
        DataGridView1.Columns(3).Width = 90

        DataGridView1.Columns(4).HeaderText = "カナ"
        DataGridView1.Columns(4).Width = 80

        DataGridView1.Columns(5).HeaderText = "性別"
        DataGridView1.Columns(5).Width = 35
        DataGridView1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView1.Columns(6).HeaderText = "生年月日"
        DataGridView1.Columns(6).Width = 80
        DataGridView1.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView1.Columns("age").HeaderText = "年齢"
        DataGridView1.Columns("age").Width = 40
        DataGridView1.Columns("age").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("age").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView1.Columns(7).HeaderText = "企業名"
        DataGridView1.Columns(7).Width = 130

        DataGridView1.Columns(8).HeaderText = "結果渡日"
        DataGridView1.Columns(8).Width = 60

        DataGridView1.Columns(9).HeaderText = "来院郵送"
        DataGridView1.Columns(9).Width = 60

        DataGridView1.Columns(10).HeaderText = "メモ1"
        DataGridView1.Columns(10).Width = 80

        DataGridView1.Columns(11).HeaderText = "メモ2"
        DataGridView1.Columns(11).Width = 80

        DataGridView1.Columns(12).HeaderText = "窓口負担"
        DataGridView1.Columns(12).Width = 60

        For i As Integer = 0 To 12
            DataGridView1.Columns(i).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Next

        '非表示の列
        DataGridView1.Columns("Kjn1").Visible = False
        DataGridView1.Columns("Kjn2").Visible = False
        DataGridView1.Columns("Kjn3").Visible = False
        DataGridView1.Columns("Kjn4").Visible = False
        DataGridView1.Columns("Kjn5").Visible = False
        DataGridView1.Columns("Kjn6").Visible = False
        DataGridView1.Columns("Kig1").Visible = False
        DataGridView1.Columns("Kig2").Visible = False
        DataGridView1.Columns("Kig3").Visible = False
        DataGridView1.Columns("Kig4").Visible = False
        DataGridView1.Columns("Kig5").Visible = False
        DataGridView1.Columns("Kig6").Visible = False
        DataGridView1.Columns("Sei1").Visible = False
        DataGridView1.Columns("Sei2").Visible = False
        DataGridView1.Columns("Sei3").Visible = False
        DataGridView1.Columns("Sei4").Visible = False
        DataGridView1.Columns("Tok1").Visible = False
        DataGridView1.Columns("Tok2").Visible = False
        DataGridView1.Columns("Tok3").Visible = False
        DataGridView1.Columns("Tok4").Visible = False
        DataGridView1.Columns("Tok5").Visible = False
        DataGridView1.Columns("Tok6").Visible = False
        DataGridView1.Columns("Tok7").Visible = False
        DataGridView1.Columns("Tok8").Visible = False
        DataGridView1.Columns("Tok9").Visible = False
        DataGridView1.Columns("Gan1").Visible = False
        DataGridView1.Columns("Gan2").Visible = False
        DataGridView1.Columns("Sanken").Visible = False

        initFlg = False

        '本日日付の行までスクロール
        Dim todayDate As String = DateTime.Today.ToString("yyyy/MM/dd")
        Dim rowsCount As Integer = DataGridView1.Rows.Count
        For i = 0 To rowsCount - 2
            If DataGridView1("Ymd", i).Value = todayDate Then
                DataGridView1.FirstDisplayedScrollingRowIndex = i
                Exit For
            End If
        Next

    End Sub

    Private Sub selectedClear()
        Dim rowsCount As Integer = DataGridView1.Rows.Count
        For i = 0 To rowsCount - 2
            DataGridView1.Rows.Item(i).Selected = False
        Next
    End Sub

    Private Sub eraBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles eraBox.TextChanged
        If initFlg = False Then
            '一覧表示
            displayReserveList()

            '西暦を和暦表示に変更
            convertJapanCalender()

            '年齢の表示設定
            displayAgeColumn()

            '曜日の表示設定
            displayDayColumun()

            '選択行のクリア
            selectedClear()
        End If
    End Sub

    Private Sub monthBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles monthBox.TextChanged
        If initFlg = False Then
            '一覧表示
            displayReserveList()

            '西暦を和暦表示に変更
            convertJapanCalender()

            '年齢の表示設定
            displayAgeColumn()

            '曜日の表示設定
            displayDayColumun()

            '選択行のクリア
            selectedClear()
        End If
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim selectedRowsCount As Integer = DataGridView1.SelectedRows.Count
        Dim selectedRowsIndexList As New ArrayList
        Dim rowsCount As Integer = DataGridView1.Rows.Count

        If selectedRowsCount = 0 Then
            For i = 0 To rowsCount - 1
                selectedRowsIndexList.Add(i)
            Next
        Else
            For Each row As DataGridViewRow In DataGridView1.SelectedRows
                selectedRowsIndexList.Add(row.Index)
            Next
            selectedRowsIndexList.Sort()
        End If

        Dim objExcel As Object
        Dim objWorkBooks As Object
        Dim objWorkBook As Object
        Dim oSheet As Object

        objExcel = CreateObject("Excel.Application")
        objWorkBooks = objExcel.Workbooks
        'objWorkBook = objWorkBooks.Open("\\PRIMERGYTX100S1\Hakojun\事務\さかもと\Patient -入院患者-\Patient.xls")
        objWorkBook = objWorkBooks.Open("C:\Users\yoshi\Desktop\Reserve.xls")
        oSheet = objWorkBook.Worksheets("予定表")

        '年月と時刻部分の書き込み
        Dim ymStr As String = eraBox.Text & " 年 " & monthBox.Text & " 月"
        Dim nowTime As DateTime = DateTime.Now
        oSheet.Range("G2").Value = ymStr
        oSheet.Range("I2").Value = nowTime.ToString

        'エクセルの**の文字列を削除(B4セルからAA4セルの文字列)
        Dim columnChar As Char = "B"
        While columnChar <> "["
            oSheet.Range(columnChar & 4).Value = ""
            columnChar = Chr((Convert.ToInt32(columnChar) + 1))
        End While
        oSheet.Range("AA" & 4).Value = ""

        'クリップボードにコピーする
        Dim xlRange As Excel.Range = oSheet.Cells.Range("B1:AA38")
        xlRange.Copy()

        If selectedRowsIndexList.Count = 0 Then
            MsgBox("データが存在しない年月です。")
            Return
        End If

        If selectedRowsIndexList.Count > 105 Then
            '4枚作成
            '指定位置にペーストする(2枚目)
            Dim xlPasteRange As Excel.Range = oSheet.Range("B40")
            oSheet.Paste(xlPasteRange)

            '行の高さ設定
            oSheet.Range("40:40").RowHeight = 6
            oSheet.Range("41:41").RowHeight = 24
            oSheet.Range("42:42").RowHeight = 54
            oSheet.Range("43:77").RowHeight = 15

            '指定位置にペーストする(3枚目)
            xlPasteRange = oSheet.Range("B79")
            oSheet.Paste(xlPasteRange)

            oSheet.Range("79:79").RowHeight = 6
            oSheet.Range("80:80").RowHeight = 24
            oSheet.Range("81:81").RowHeight = 54
            oSheet.Range("82:116").RowHeight = 15

            '指定位置にペーストする(4枚目)
            xlPasteRange = oSheet.Range("B118")
            oSheet.Paste(xlPasteRange)

            oSheet.Range("118:118").RowHeight = 6
            oSheet.Range("119:119").RowHeight = 24
            oSheet.Range("120:120").RowHeight = 54
            oSheet.Range("121:155").RowHeight = 15

        ElseIf selectedRowsIndexList.Count > 70 Then
            '3枚作成
            '指定位置にペーストする(2枚目)
            Dim xlPasteRange As Excel.Range = oSheet.Range("B40")
            oSheet.Paste(xlPasteRange)

            '行の高さ設定
            oSheet.Range("40:40").RowHeight = 6
            oSheet.Range("41:41").RowHeight = 24
            oSheet.Range("42:42").RowHeight = 54
            oSheet.Range("43:77").RowHeight = 15

            '指定位置にペーストする(3枚目)
            xlPasteRange = oSheet.Range("B79")
            oSheet.Paste(xlPasteRange)

            oSheet.Range("79:79").RowHeight = 6
            oSheet.Range("80:80").RowHeight = 24
            oSheet.Range("81:81").RowHeight = 54
            oSheet.Range("82:116").RowHeight = 15

        ElseIf selectedRowsIndexList.Count > 35 Then
            '2枚作成
            '指定位置にペーストする(2枚目)
            Dim xlPasteRange As Excel.Range = oSheet.Range("B40")
            oSheet.Paste(xlPasteRange)

            oSheet.Range("40:40").RowHeight = 6
            oSheet.Range("41:41").RowHeight = 24
            oSheet.Range("42:42").RowHeight = 54
            oSheet.Range("43:77").RowHeight = 15
        Else

            ''1枚作成
            'Dim type As String = ""
            'Dim border As Excel.Border = Nothing
            'Dim rowIndex As Integer = 4

            ''セルに書き込み
            'Dim excelIndex As Integer = 0
            'For Each i As Integer In selectedRowsIndexList
            '    oSheet.Range("B" & (rowIndex + excelIndex)).Value = excelIndex + 1 'No
            '    oSheet.Range("C" & (rowIndex + excelIndex)).Value = DataGridView1("Ymd", i).FormattedValue '予約日

            '    '予約日で区切りの罫線をいれる
            '    If i <> 0 AndAlso DataGridView1("Ymd", i).Value <> DataGridView1("Ymd", i - 1).Value Then
            '        border = oSheet.Range("B" & (rowIndex + excelIndex), "AA" & (rowIndex + excelIndex)).Borders(Excel.XlBordersIndex.xlEdgeTop)
            '        border.LineStyle = Excel.XlLineStyle.xlContinuous
            '        border.Weight = Excel.XlBorderWeight.xlThin
            '    End If

            '    oSheet.Range("D" & (rowIndex + excelIndex)).Value = DataGridView1("day", i).FormattedValue '曜日
            '    oSheet.Range("E" & (rowIndex + excelIndex)).Value = DataGridView1("Apm", i).FormattedValue '予約時間
            '    oSheet.Range("F" & (rowIndex + excelIndex)).Value = DataGridView1("Syu", i).FormattedValue '種別
            '    oSheet.Range("G" & (rowIndex + excelIndex)).Value = DataGridView1("Nam", i).FormattedValue '氏名
            '    oSheet.Range("H" & (rowIndex + excelIndex)).Value = DataGridView1("Kana", i).FormattedValue 'カナ
            '    oSheet.Range("I" & (rowIndex + excelIndex)).Value = DataGridView1("Sex", i).FormattedValue '性別
            '    oSheet.Range("J" & (rowIndex + excelIndex)).Value = DataGridView1("Birth", i).FormattedValue '生年月日
            '    oSheet.Range("K" & (rowIndex + excelIndex)).Value = DataGridView1("age", i).FormattedValue '年齢

            '    '企業名
            '    If DataGridView1("Ind", i).FormattedValue.ToString.Length > 10 Then
            '        oSheet.Range("L" & (rowIndex + excelIndex)).Value = DataGridView1("Ind", i).FormattedValue.ToString.Substring(0, 10)
            '    Else
            '        oSheet.Range("L" & (rowIndex + excelIndex)).Value = DataGridView1("Ind", i).FormattedValue
            '    End If

            '    oSheet.Range("M" & (rowIndex + excelIndex)).Value = DataGridView1("Ymd2", i).FormattedValue '結果渡日
            '    oSheet.Range("N" & (rowIndex + excelIndex)).Value = DataGridView1("Send", i).FormattedValue '来院郵送

            '    '窓口負担
            '    If DataGridView1("Futan", i).FormattedValue = 0 Then
            '        oSheet.Range("O" & (rowIndex + excelIndex)).Value = ""
            '    Else
            '        oSheet.Range("O" & (rowIndex + excelIndex)).Value = DataGridView1("Futan", i).FormattedValue
            '    End If

            '    oSheet.Range("P" & (rowIndex + excelIndex)).Value = DataGridView1("Memo1", i).FormattedValue 'メモ

            '    type = DataGridView1("Syu", i).FormattedValue
            '    If type = "個人" Then
            '        If DataGridView1("Kjn1", i).Value = 1 Then
            '            oSheet.Range("Q" & (rowIndex + excelIndex)).Value = 1 '血液
            '        Else
            '            oSheet.Range("Q" & (rowIndex + excelIndex)).Value = ""
            '        End If
            '        If DataGridView1("Kjn2", i).Value = 1 Then
            '            oSheet.Range("R" & (rowIndex + excelIndex)).Value = 1 '心電図
            '        Else
            '            oSheet.Range("R" & (rowIndex + excelIndex)).Value = ""
            '        End If
            '        If DataGridView1("Kjn3", i).Value = 1 Then
            '            oSheet.Range("S" & (rowIndex + excelIndex)).Value = 1 '胸部XP
            '        Else
            '            oSheet.Range("S" & (rowIndex + excelIndex)).Value = ""
            '        End If
            '        If DataGridView1("Kjn4", i).Value = 1 Then
            '            oSheet.Range("T" & (rowIndex + excelIndex)).Value = 1 '超音波
            '        Else
            '            oSheet.Range("T" & (rowIndex + excelIndex)).Value = ""
            '        End If
            '        If DataGridView1("Kjn5", i).Value = 1 Then
            '            oSheet.Range("U" & (rowIndex + excelIndex)).Value = 1 '胃Ba
            '        Else
            '            oSheet.Range("U" & (rowIndex + excelIndex)).Value = ""
            '        End If
            '        If DataGridView1("Kjn6", i).Value = 1 Then
            '            oSheet.Range("V" & (rowIndex + excelIndex)).Value = 1 '胃カメラ
            '        Else
            '            oSheet.Range("V" & (rowIndex + excelIndex)).Value = ""
            '        End If
            '    ElseIf type = "企業" Then
            '        If DataGridView1("Kig1", i).Value = 1 Then
            '            oSheet.Range("Q" & (rowIndex + excelIndex)).Value = 1 '血液
            '        Else
            '            oSheet.Range("Q" & (rowIndex + excelIndex)).Value = ""
            '        End If
            '        If DataGridView1("Kig2", i).Value = 1 Then
            '            oSheet.Range("R" & (rowIndex + excelIndex)).Value = 1 '心電図
            '        Else
            '            oSheet.Range("R" & (rowIndex + excelIndex)).Value = ""
            '        End If
            '        If DataGridView1("Kig3", i).Value = 1 Then
            '            oSheet.Range("S" & (rowIndex + excelIndex)).Value = 1 '胸部XP
            '        Else
            '            oSheet.Range("S" & (rowIndex + excelIndex)).Value = ""
            '        End If
            '        If DataGridView1("Kig4", i).Value = 1 Then
            '            oSheet.Range("T" & (rowIndex + excelIndex)).Value = 1 '超音波
            '        Else
            '            oSheet.Range("T" & (rowIndex + excelIndex)).Value = ""
            '        End If
            '        If DataGridView1("Kig5", i).Value = 1 Then
            '            oSheet.Range("U" & (rowIndex + excelIndex)).Value = 1 '胃Ba
            '        Else
            '            oSheet.Range("U" & (rowIndex + excelIndex)).Value = ""
            '        End If
            '        If DataGridView1("Kig6", i).Value = 1 Then
            '            oSheet.Range("V" & (rowIndex + excelIndex)).Value = 1 '胃カメラ
            '        Else
            '            oSheet.Range("V" & (rowIndex + excelIndex)).Value = ""
            '        End If
            '    ElseIf type = "生活" Then
            '        oSheet.Range("Q" & (rowIndex + excelIndex)).Value = 1 '血液
            '        oSheet.Range("R" & (rowIndex + excelIndex)).Value = 1 '心電図
            '        oSheet.Range("S" & (rowIndex + excelIndex)).Value = 1 '胸部XP

            '        If DataGridView1("Sei3", i).Value = 1 Then
            '            oSheet.Range("U" & (rowIndex + excelIndex)).Value = 1 '胃Ba
            '        Else
            '            oSheet.Range("U" & (rowIndex + excelIndex)).Value = ""
            '        End If
            '        If DataGridView1("Sei4", i).Value = 1 Then
            '            oSheet.Range("V" & (rowIndex + excelIndex)).Value = 1 '胃カメラ
            '        Else
            '            oSheet.Range("V" & (rowIndex + excelIndex)).Value = ""
            '        End If
            '    ElseIf type = "特定" Then
            '        oSheet.Range("W" & (rowIndex + excelIndex)).Value = DataGridView1("Tok1", i).Value '保険種別
            '        oSheet.Range("Y" & (rowIndex + excelIndex)).Value = 3 '採血数
            '    ElseIf type = "がん" Then
            '        If DataGridView1("Gan1", i).Value = 1 Then
            '            oSheet.Range("Z" & (rowIndex + excelIndex)).Value = 1 '胃がん
            '        Else
            '            oSheet.Range("Z" & (rowIndex + excelIndex)).Value = ""
            '        End If
            '        If DataGridView1("Gan2", i).Value = 1 Then
            '            oSheet.Range("AA" & (rowIndex + excelIndex)).Value = 1 '大腸がん
            '        Else
            '            oSheet.Range("AA" & (rowIndex + excelIndex)).Value = ""
            '        End If
            '    End If
            '    excelIndex += 1
            'Next
            writeReserveList(oSheet, selectedRowsIndexList)

        End If

        objExcel.DisplayAlerts = False
        ' エクセル表示
        objExcel.Visible = True

        '印刷
        oSheet.PrintPreview(1)

        ' EXCEL解放
        objExcel.Quit()
        Marshal.ReleaseComObject(oSheet)
        Marshal.ReleaseComObject(objWorkBook)
        Marshal.ReleaseComObject(objExcel)
        oSheet = Nothing

        objWorkBook = Nothing
        objExcel = Nothing
    End Sub

    Private Sub writeReserveList(ByVal oSheet As Object, ByVal selectedRowsIndexList As ArrayList)
        '1枚作成
        Dim type As String = ""
        Dim border As Excel.Border = Nothing
        Dim rowIndex As Integer = 4

        'セルに書き込み
        Dim excelIndex As Integer = 0
        For Each i As Integer In selectedRowsIndexList

            oSheet.Range("B" & (rowIndex + excelIndex)).Value = excelIndex + 1 'No
            oSheet.Range("C" & (rowIndex + excelIndex)).Value = DataGridView1("Ymd", i).FormattedValue '予約日

            '予約日で区切りの罫線をいれる
            If i <> 0 AndAlso DataGridView1("Ymd", i).Value <> DataGridView1("Ymd", i - 1).Value Then
                border = oSheet.Range("B" & (rowIndex + excelIndex), "AA" & (rowIndex + excelIndex)).Borders(Excel.XlBordersIndex.xlEdgeTop)
                border.LineStyle = Excel.XlLineStyle.xlContinuous
                border.Weight = Excel.XlBorderWeight.xlThin
            End If

            oSheet.Range("D" & (rowIndex + excelIndex)).Value = DataGridView1("day", i).FormattedValue '曜日
            oSheet.Range("E" & (rowIndex + excelIndex)).Value = DataGridView1("Apm", i).FormattedValue '予約時間
            oSheet.Range("F" & (rowIndex + excelIndex)).Value = DataGridView1("Syu", i).FormattedValue '種別
            oSheet.Range("G" & (rowIndex + excelIndex)).Value = DataGridView1("Nam", i).FormattedValue '氏名
            oSheet.Range("H" & (rowIndex + excelIndex)).Value = DataGridView1("Kana", i).FormattedValue 'カナ
            oSheet.Range("I" & (rowIndex + excelIndex)).Value = DataGridView1("Sex", i).FormattedValue '性別
            oSheet.Range("J" & (rowIndex + excelIndex)).Value = DataGridView1("Birth", i).FormattedValue '生年月日
            oSheet.Range("K" & (rowIndex + excelIndex)).Value = DataGridView1("age", i).FormattedValue '年齢

            '企業名
            If DataGridView1("Ind", i).FormattedValue.ToString.Length > 10 Then
                oSheet.Range("L" & (rowIndex + excelIndex)).Value = DataGridView1("Ind", i).FormattedValue.ToString.Substring(0, 10)
            Else
                oSheet.Range("L" & (rowIndex + excelIndex)).Value = DataGridView1("Ind", i).FormattedValue
            End If

            oSheet.Range("M" & (rowIndex + excelIndex)).Value = DataGridView1("Ymd2", i).FormattedValue '結果渡日
            oSheet.Range("N" & (rowIndex + excelIndex)).Value = DataGridView1("Send", i).FormattedValue '来院郵送

            '窓口負担
            If DataGridView1("Futan", i).FormattedValue = 0 Then
                oSheet.Range("O" & (rowIndex + excelIndex)).Value = ""
            Else
                oSheet.Range("O" & (rowIndex + excelIndex)).Value = DataGridView1("Futan", i).FormattedValue
            End If

            oSheet.Range("P" & (rowIndex + excelIndex)).Value = DataGridView1("Memo1", i).FormattedValue 'メモ

            type = DataGridView1("Syu", i).FormattedValue
            If type = "個人" Then
                If DataGridView1("Kjn1", i).Value = 1 Then
                    oSheet.Range("Q" & (rowIndex + excelIndex)).Value = 1 '血液
                Else
                    oSheet.Range("Q" & (rowIndex + excelIndex)).Value = ""
                End If
                If DataGridView1("Kjn2", i).Value = 1 Then
                    oSheet.Range("R" & (rowIndex + excelIndex)).Value = 1 '心電図
                Else
                    oSheet.Range("R" & (rowIndex + excelIndex)).Value = ""
                End If
                If DataGridView1("Kjn3", i).Value = 1 Then
                    oSheet.Range("S" & (rowIndex + excelIndex)).Value = 1 '胸部XP
                Else
                    oSheet.Range("S" & (rowIndex + excelIndex)).Value = ""
                End If
                If DataGridView1("Kjn4", i).Value = 1 Then
                    oSheet.Range("T" & (rowIndex + excelIndex)).Value = 1 '超音波
                Else
                    oSheet.Range("T" & (rowIndex + excelIndex)).Value = ""
                End If
                If DataGridView1("Kjn5", i).Value = 1 Then
                    oSheet.Range("U" & (rowIndex + excelIndex)).Value = 1 '胃Ba
                Else
                    oSheet.Range("U" & (rowIndex + excelIndex)).Value = ""
                End If
                If DataGridView1("Kjn6", i).Value = 1 Then
                    oSheet.Range("V" & (rowIndex + excelIndex)).Value = 1 '胃カメラ
                Else
                    oSheet.Range("V" & (rowIndex + excelIndex)).Value = ""
                End If
            ElseIf type = "企業" Then
                If DataGridView1("Kig1", i).Value = 1 Then
                    oSheet.Range("Q" & (rowIndex + excelIndex)).Value = 1 '血液
                Else
                    oSheet.Range("Q" & (rowIndex + excelIndex)).Value = ""
                End If
                If DataGridView1("Kig2", i).Value = 1 Then
                    oSheet.Range("R" & (rowIndex + excelIndex)).Value = 1 '心電図
                Else
                    oSheet.Range("R" & (rowIndex + excelIndex)).Value = ""
                End If
                If DataGridView1("Kig3", i).Value = 1 Then
                    oSheet.Range("S" & (rowIndex + excelIndex)).Value = 1 '胸部XP
                Else
                    oSheet.Range("S" & (rowIndex + excelIndex)).Value = ""
                End If
                If DataGridView1("Kig4", i).Value = 1 Then
                    oSheet.Range("T" & (rowIndex + excelIndex)).Value = 1 '超音波
                Else
                    oSheet.Range("T" & (rowIndex + excelIndex)).Value = ""
                End If
                If DataGridView1("Kig5", i).Value = 1 Then
                    oSheet.Range("U" & (rowIndex + excelIndex)).Value = 1 '胃Ba
                Else
                    oSheet.Range("U" & (rowIndex + excelIndex)).Value = ""
                End If
                If DataGridView1("Kig6", i).Value = 1 Then
                    oSheet.Range("V" & (rowIndex + excelIndex)).Value = 1 '胃カメラ
                Else
                    oSheet.Range("V" & (rowIndex + excelIndex)).Value = ""
                End If
            ElseIf type = "生活" Then
                oSheet.Range("Q" & (rowIndex + excelIndex)).Value = 1 '血液
                oSheet.Range("R" & (rowIndex + excelIndex)).Value = 1 '心電図
                oSheet.Range("S" & (rowIndex + excelIndex)).Value = 1 '胸部XP

                If DataGridView1("Sei3", i).Value = 1 Then
                    oSheet.Range("U" & (rowIndex + excelIndex)).Value = 1 '胃Ba
                Else
                    oSheet.Range("U" & (rowIndex + excelIndex)).Value = ""
                End If
                If DataGridView1("Sei4", i).Value = 1 Then
                    oSheet.Range("V" & (rowIndex + excelIndex)).Value = 1 '胃カメラ
                Else
                    oSheet.Range("V" & (rowIndex + excelIndex)).Value = ""
                End If
            ElseIf type = "特定" Then
                oSheet.Range("W" & (rowIndex + excelIndex)).Value = DataGridView1("Tok1", i).Value '保険種別
                oSheet.Range("Y" & (rowIndex + excelIndex)).Value = 3 '採血数
            ElseIf type = "がん" Then
                If DataGridView1("Gan1", i).Value = 1 Then
                    oSheet.Range("Z" & (rowIndex + excelIndex)).Value = 1 '胃がん
                Else
                    oSheet.Range("Z" & (rowIndex + excelIndex)).Value = ""
                End If
                If DataGridView1("Gan2", i).Value = 1 Then
                    oSheet.Range("AA" & (rowIndex + excelIndex)).Value = 1 '大腸がん
                Else
                    oSheet.Range("AA" & (rowIndex + excelIndex)).Value = ""
                End If
            End If
            excelIndex += 1
        Next

    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim deleteFlg As Boolean = False
        For i = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows.Item(i).Selected = True Then
                deleteFlg = True
                Exit For
            End If
        Next

        If deleteFlg = False Then
            MsgBox("選択されていません")
            Return
        End If

        Dim index As Integer = DataGridView1.CurrentRow.Index

        Dim birthDay As String = DataGridView1("Birth", index).Value
        Dim name As String = DataGridView1("Nam", index).Value

        birthDay = birthDay.Replace(birthDay.Substring(0, 3), convertWarekiToAD(birthDay.Substring(0, 3)))

        '削除処理
        Dim cn As New OleDbConnection(DB_reserve)
        Dim sqlcm As OleDbCommand = cn.CreateCommand
        Dim adapter As New OleDbDataAdapter(sqlcm)
        Dim table As New DataTable
        sqlcm.CommandText = "delete from RsvD where Nam='" & name & "' AND Birth='" & birthDay & "'"
        adapter.Fill(table)

        MsgBox("削除しました")

    End Sub

    Private Sub diagnoseButton_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles diagnoseButton.CheckedChanged
        If diagnoseButton.Checked = True Then
            inputClear()
            personListBox.Items.Clear()
            displayDiagnose()
        End If
    End Sub

    Private Sub HealthButton_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HealthButton.CheckedChanged
        If HealthButton.Checked = True Then
            inputClear()
            personListBox.Items.Clear()
            displayHealth()
        End If
    End Sub

    Private Sub sankenCenterButton_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles sankenCenterButton.CheckedChanged
        If sankenCenterButton.Checked = True Then
            inputClear()
            personListBox.Items.Clear()
            displaySankenCenter()
        End If
    End Sub

    Private Sub referenceListBox_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles referenceListBox.SelectedValueChanged
        'リストのクリア
        personListBox.Items.Clear()

        '選択された企業名
        Dim ind As String = referenceListBox.SelectedItem

        'DBの選択
        Dim DB As String = ""
        Dim Cn As OleDbConnection
        Dim SQLCm As OleDbCommand
        Dim reader As System.Data.OleDb.OleDbDataReader

        If diagnoseButton.Checked = True Then
            '健診
            Cn = New OleDbConnection(DB_diagnose)
            SQLCm = Cn.CreateCommand
            SQLCm.CommandText = "SELECT Nam, Kana FROM UsrM WHERE Ind='" & ind & " 'ORDER BY Kana"
        ElseIf HealthButton.Checked = True Then
            '生活習慣病
            Cn = New OleDbConnection(DB_health)
            SQLCm = Cn.CreateCommand
            SQLCm.CommandText = "SELECT Nam, Kana FROM UsrM WHERE Ind='" & ind & "' ORDER BY Kana"
        Else
            '産健ｾﾝﾀｰ
            Cn = New OleDbConnection(DB_reserve)
            SQLCm = Cn.CreateCommand
            SQLCm.CommandText = "SELECT distinct Nam, Kana FROM RsvD WHERE Ind='" & ind & "' ORDER BY Kana"
        End If

        Cn.Open()
        reader = SQLCm.ExecuteReader()
        While reader.Read() = True
            personListBox.Items.Add(reader("Nam"))
        End While
        reader.Close()
        Cn.Close()
        SQLCm.Dispose()
        Cn.Dispose()
    End Sub

    Private Sub personListBox_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles personListBox.SelectedValueChanged
        '入力エリアのクリア
        inputClear()

        '選択された企業名
        Dim selectedInd As String = referenceListBox.SelectedItem

        '選択された名前
        Dim selectedName As String = personListBox.SelectedItem

        'DBの選択
        Dim DB As String = ""
        Dim Cn As OleDbConnection
        Dim SQLCm As OleDbCommand
        Dim reader As System.Data.OleDb.OleDbDataReader

        If diagnoseButton.Checked = True Then
            '健診
            Cn = New OleDbConnection(DB_diagnose)
            SQLCm = Cn.CreateCommand
            SQLCm.CommandText = "SELECT Nam, Kana, Birth, Sex, Ind FROM UsrM WHERE Ind='" & selectedInd & "' AND Nam='" & selectedName & "'"
        ElseIf HealthButton.Checked = True Then
            '生活習慣病
            Cn = New OleDbConnection(DB_health)
            SQLCm = Cn.CreateCommand
            SQLCm.CommandText = "SELECT Nam, Kana, Birth, Sex, Ind FROM UsrM WHERE Ind='" & selectedInd & "' AND Nam='" & selectedName & "'"
        Else
            '産健ｾﾝﾀｰ
            Cn = New OleDbConnection(DB_reserve)
            SQLCm = Cn.CreateCommand
            SQLCm.CommandText = "SELECT Nam, Kana, Birth, Sex, Ind FROM RsvD WHERE Ind='" & selectedInd & "' AND Nam='" & selectedName & "'"
        End If

        Cn.Open()
        reader = SQLCm.ExecuteReader()
        While reader.Read() = True
            companyNameBox.Text = reader("Ind")
            nameBox.Text = reader("Nam")
            kanaBox.Text = reader("Kana")
            If reader("Sex") = "男" OrElse reader("Sex") = "女" Then
                sexBox.Text = reader("Sex")
            ElseIf reader("Sex") = 1 Then
                sexBox.Text = "男"
            ElseIf reader("Sex") = 2 Then
                sexBox.Text = "女"
            End If
            birthDayBox.Text = reader("Birth")
        End While
        reader.Close()
        Cn.Close()
        SQLCm.Dispose()
        Cn.Dispose()

    End Sub

    Private Sub inputClear()
        typeBox.Text = ""
        companyNameBox.Text = ""
        nameBox.Text = ""
        kanaBox.Text = ""
        sexBox.Text = ""
        birthDayBox.Text = ""
        reserveDateBox.Text = ""
        ampmBox.Text = ""
        resultDayBox.Text = ""
        postBox.Text = ""
        memo1Box.Text = ""
        memo2Box.Text = ""
    End Sub

    Private Sub btnSelectClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelectClear.Click
        Dim rowsCount As Integer = DataGridView1.Rows.Count
        For i = 0 To rowsCount - 1
            DataGridView1.Rows.Item(i).Selected = False
        Next
    End Sub
End Class