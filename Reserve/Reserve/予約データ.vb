Imports System.Data.OleDb
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Core
Imports System.Reflection

Public Class 予約データ

    Public initFlg As Boolean = True

    Private eraTable As Dictionary(Of Integer, String)

    Private ci As System.Globalization.CultureInfo

    Public Sub New(eraTable As Dictionary(Of Integer, String), ci As System.Globalization.CultureInfo)

        ' この呼び出しはデザイナーで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後で初期化を追加します。
        Me.eraTable = eraTable
        Me.ci = ci

    End Sub

    ''' <summary>
    ''' コントロールのDoubleBufferedプロパティをTrueにする
    ''' </summary>
    ''' <param name="control">対象のコントロール</param>
    Public Shared Sub EnableDoubleBuffering(control As Control)
        control.GetType().InvokeMember( _
            "DoubleBuffered", _
            BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.SetProperty, _
            Nothing, _
            control, _
            New Object() {True})
    End Sub

    Private Sub 予約データ_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        Form1.f_yoyaku = Nothing
    End Sub

    Private Sub 予約データ_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Enter Then
            If e.Control = False Then
                Me.SelectNextControl(Me.ActiveControl, Not e.Shift, True, True, True)
            End If
        End If
    End Sub

    Private Sub 予約データ_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'DoubleBufferedプロパティをTrueにする
        EnableDoubleBuffering(DataGridView1)

        '位置
        Me.Left = 10
        Me.Top = 50

        '一覧表示の初期設定
        initialSetting4DataGridView()

        'セルの編集不可
        DataGridView1.ReadOnly = True

        'セルを選択すると行全体が選択されるようにする
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        '予約日と種別と企業名以外の列をソート不可に設定
        For Each c As DataGridViewColumn In DataGridView1.Columns
            If c.Name = "Ymd" OrElse c.Name = "Syu" OrElse c.Name = "Kana" OrElse c.Name = "Ind" Then
                Continue For
            End If
            c.SortMode = DataGridViewColumnSortMode.NotSortable
        Next c

        DataGridView1.AllowUserToAddRows = False

        TabControl1.SizeMode = TabSizeMode.Fixed
        TabControl1.ItemSize = New Size(65, 25)
        TabControl1.SelectedTab = referenceTabPage

    End Sub

    Private Sub displayDiagnose()
        referenceListBox.Items.Clear()

        Dim Cn As New OleDbConnection(Form1.DB_diagnose)
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

        Dim Cn As New OleDbConnection(Form1.DB_health)
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

        Dim Cn As New OleDbConnection(Form1.DB_reserve)
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
        resultDayBox.Text = resultDay
        postBox.Text = post
        memo1Box.Text = memo1
        memo2Box.Text = memo2
        ampmBox.Text = ampm
        birthYmdBox.setADStr(birthDay)
        reserveYmdBox.setADStr(reserveDay)

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
        If DataGridView1("Kjn2", rowIndex).Value = 1 Then
            personalElectro.Checked = True
        Else
            personalElectro.Checked = False
        End If
        '胸部XP
        If DataGridView1("Kjn3", rowIndex).Value = 1 Then
            personalChestXP.Checked = True
        Else
            personalChestXP.Checked = False
        End If
        '超音波
        If DataGridView1("Kjn4", rowIndex).Value = 1 Then
            personalUltrasonic.Checked = True
        Else
            personalUltrasonic.Checked = False
        End If
        '胃バリウム
        If DataGridView1("Kjn5", rowIndex).Value = 1 Then
            personalStomachBa.Checked = True
        Else
            personalStomachBa.Checked = False
        End If
        '胃カメラ
        If DataGridView1("Kjn6", rowIndex).Value = 1 Then
            personalStomachCamera.Checked = True
        Else
            personalStomachCamera.Checked = False
        End If
        '窓口負担
        If windowPay <> 0 AndAlso DataGridView1("Syu", rowIndex).Value = "個人" Then
            personalWindowPay.Text = windowPay
        Else
            personalWindowPay.Text = ""
        End If

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
        If windowPay <> 0 AndAlso DataGridView1("Syu", rowIndex).Value = "企業" Then
            companyWindowPay.Text = windowPay
        Else
            companyWindowPay.Text = ""
        End If


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
        '窓口負担
        If windowPay <> 0 AndAlso DataGridView1("Syu", rowIndex).Value = "生活" Then
            lifeStyleWindowPay.Text = windowPay
        Else
            lifeStyleWindowPay.Text = ""
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
            diabetesBox.Text = DataGridView1("ToK7", rowIndex).Value
        Else
            diabetesBox.Text = ""
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
        '窓口負担
        If windowPay <> 0 AndAlso DataGridView1("Syu", rowIndex).Value = "特定" Then
            specificWindowPay.Text = windowPay
        Else
            specificWindowPay.Text = ""
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
        '窓口負担
        If windowPay <> 0 AndAlso DataGridView1("Syu", rowIndex).Value = "がん" Then
            cancerWindowPay.Text = windowPay
        Else
            cancerWindowPay.Text = ""
        End If

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
            '予約日の表示設定,グループ化
            If e.RowIndex > 0 AndAlso DataGridView1(e.ColumnIndex, e.RowIndex - 1).Value = e.Value Then
                e.Value = ""
                e.FormattingApplied = True
            Else
                e.Value = Integer.Parse(e.Value.Substring(e.Value.ToString.Length - 2, 2))
            End If
        ElseIf DataGridView1.Columns(e.ColumnIndex).Name = "day" Then
            '曜日の表示設定,グループ化
            If e.RowIndex > 0 AndAlso DataGridView1(e.ColumnIndex, e.RowIndex - 1).Value = e.Value Then
                e.Value = ""
                e.FormattingApplied = True
            Else
                If e.Value = 1 Then
                    e.Value = "日"
                ElseIf e.Value = 2 Then
                    e.Value = "月"
                ElseIf e.Value = 3 Then
                    e.Value = "火"
                ElseIf e.Value = 4 Then
                    e.Value = "水"
                ElseIf e.Value = 5 Then
                    e.Value = "木"
                ElseIf e.Value = 6 Then
                    e.Value = "金"
                ElseIf e.Value = 7 Then
                    e.Value = "土"
                End If
                e.FormattingApplied = True
            End If
        ElseIf DataGridView1.Columns(e.ColumnIndex).Name = "Apm" Then
            '時間のグループ化
            If e.RowIndex > 0 AndAlso DataGridView1("day", e.RowIndex).Value = DataGridView1("day", e.RowIndex - 1).Value AndAlso DataGridView1(e.ColumnIndex, e.RowIndex - 1).Value = e.Value Then
                e.Value = ""
                e.FormattingApplied = True
            End If
        ElseIf DataGridView1.Columns(e.ColumnIndex).Name = "Birth" Then
            '生年月日の和暦表示
            Dim dt As DateTime = e.Value
            Dim eraIndex As Integer = ci.DateTimeFormat.Calendar.GetEra(dt)
            e.Value = eraTable(eraIndex) & dt.ToString("yy/MM/dd", ci)
            e.FormattingApplied = True
        End If

    End Sub

    Private Sub displayReserveList()
        Dim eraStr As String = eraBox.Text
        Dim monthStr As String = monthBox.Text
        Dim targetDateStr As String = convertWarekiToAD(eraStr) & "/" & monthStr

        Dim Cn As New OleDbConnection(Form1.DB_reserve)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim Adapter As New OleDbDataAdapter(SQLCm)
        Dim Table As New DataTable

        SQLCm.CommandText = "SELECT Ymd, WeekDay(Ymd) as [day], Apm, Syu, Nam, Kana, Sex, Birth, Int((Format(NOW(),'YYYYMMDD')-Format(Birth, 'YYYYMMDD'))/10000) as age, Ind, Ymd2, Send, Memo1, Memo2, Futan, Kjn1, Kjn2, Kjn3, Kjn4, Kjn5, Kjn6, Kig1, Kig2, Kig3, Kig4, Kig5, Kig6, Sei1, Sei2, Sei3, Sei4, Tok1, Tok2, Tok3, Tok4, Tok5, Tok6, Tok7, Tok8, Tok9, Gan1, Gan2, Sanken FROM RsvD WHERE Ymd LIKE '%" & targetDateStr & "%' ORDER BY Ymd ASC, Apm ASC, Kana ASC"
        Adapter.Fill(Table)

        '▼値の表示
        DataGridView1.DataSource = Table

        '▼後処理

        Table.Dispose()
        Adapter.Dispose()
        SQLCm.Dispose()
        Cn.Dispose()

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

    Private Sub initialSetting4DataGridView()

        '現在の年月を取得 
        Dim dt As DateTime = DateTime.Today
        Dim eraIndex As Integer = ci.DateTimeFormat.Calendar.GetEra(dt)
        Dim eraStr As String = eraTable(eraIndex) & dt.ToString("yy", ci)
        Dim monthStr As String = dt.ToString("MM")

        'コンボボックスに設定
        eraBox.Text = eraStr
        monthBox.Text = monthStr

        '一覧表示
        displayReserveList()

        '列名、幅の設定
        '固定
        DataGridView1.Columns(3).Frozen = True

        DataGridView1.Columns("Ymd").HeaderText = "予約日"
        DataGridView1.Columns("Ymd").Width = 70
        DataGridView1.Columns("Ymd").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("Ymd").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView1.Columns("day").HeaderText = "曜"
        DataGridView1.Columns("day").Width = 30
        DataGridView1.Columns("day").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("day").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView1.Columns("Apm").HeaderText = "AmPm"
        DataGridView1.Columns("Apm").Width = 50
        DataGridView1.Columns("Apm").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView1.Columns("Syu").HeaderText = "種別"
        DataGridView1.Columns("Syu").Width = 40
        DataGridView1.Columns("Syu").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView1.Columns("Nam").HeaderText = "氏名"
        DataGridView1.Columns("Nam").Width = 90

        DataGridView1.Columns("Kana").HeaderText = "カナ"
        DataGridView1.Columns("Kana").Width = 80

        DataGridView1.Columns("Sex").HeaderText = "性別"
        DataGridView1.Columns("Sex").Width = 35
        DataGridView1.Columns("Sex").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView1.Columns("Birth").HeaderText = "生年月日"
        DataGridView1.Columns("Birth").Width = 80
        DataGridView1.Columns("Birth").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView1.Columns("age").HeaderText = "年齢"
        DataGridView1.Columns("age").Width = 40
        DataGridView1.Columns("age").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("age").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView1.Columns("Ind").HeaderText = "企業名"
        DataGridView1.Columns("Ind").Width = 125

        DataGridView1.Columns("Ymd2").HeaderText = "結果渡日"
        DataGridView1.Columns("Ymd2").Width = 80

        DataGridView1.Columns("Send").HeaderText = "来院郵送"
        DataGridView1.Columns("Send").Width = 80

        DataGridView1.Columns("Memo1").HeaderText = "メモ1"
        DataGridView1.Columns("Memo1").Width = 80

        DataGridView1.Columns("Memo2").HeaderText = "メモ2"
        DataGridView1.Columns("Memo2").Width = 80

        DataGridView1.Columns("Futan").HeaderText = "窓口負担"
        DataGridView1.Columns("Futan").Width = 80

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
            If DataGridView1("Ymd", i).Value >= todayDate Then
                DataGridView1.FirstDisplayedScrollingRowIndex = i
                Exit For
            End If
        Next

    End Sub

    Private Sub reloadDataGridView()

        '一覧表示
        displayReserveList()

        '列名、幅の設定
        '固定
        DataGridView1.Columns(3).Frozen = True

        DataGridView1.Columns("Ymd").HeaderText = "予約日"
        DataGridView1.Columns("Ymd").Width = 70
        DataGridView1.Columns("Ymd").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView1.Columns("day").HeaderText = "曜"
        DataGridView1.Columns("day").Width = 30
        DataGridView1.Columns("day").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("day").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView1.Columns("Apm").HeaderText = "AmPm"
        DataGridView1.Columns("Apm").Width = 50
        DataGridView1.Columns("Apm").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView1.Columns("Syu").HeaderText = "種別"
        DataGridView1.Columns("Syu").Width = 40
        DataGridView1.Columns("Syu").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView1.Columns("Nam").HeaderText = "氏名"
        DataGridView1.Columns("Nam").Width = 90

        DataGridView1.Columns("Kana").HeaderText = "カナ"
        DataGridView1.Columns("Kana").Width = 80

        DataGridView1.Columns("Sex").HeaderText = "性別"
        DataGridView1.Columns("Sex").Width = 35
        DataGridView1.Columns("Sex").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView1.Columns("Birth").HeaderText = "生年月日"
        DataGridView1.Columns("Birth").Width = 80
        DataGridView1.Columns("Birth").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView1.Columns("age").HeaderText = "年齢"
        DataGridView1.Columns("age").Width = 40
        DataGridView1.Columns("age").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        DataGridView1.Columns("age").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        DataGridView1.Columns("Ind").HeaderText = "企業名"
        DataGridView1.Columns("Ind").Width = 125

        DataGridView1.Columns("Ymd2").HeaderText = "結果渡日"
        DataGridView1.Columns("Ymd2").Width = 80

        DataGridView1.Columns("Send").HeaderText = "来院郵送"
        DataGridView1.Columns("Send").Width = 80

        DataGridView1.Columns("Memo1").HeaderText = "メモ1"
        DataGridView1.Columns("Memo1").Width = 80

        DataGridView1.Columns("Memo2").HeaderText = "メモ2"
        DataGridView1.Columns("Memo2").Width = 80

        DataGridView1.Columns("Futan").HeaderText = "窓口負担"
        DataGridView1.Columns("Futan").Width = 60

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

            '選択行のクリア
            selectedClear()
        End If
    End Sub

    Private Sub monthBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles monthBox.TextChanged
        If initFlg = False Then
            '一覧表示
            displayReserveList()

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
        objWorkBook = objWorkBooks.Open("\\PRIMERGYTX100S1\Reserve\Reserve.xls")
        'objWorkBook = objWorkBooks.Open("C:\Users\yoshi\Desktop\Reserve.xls")
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

            writeReserveList(oSheet, selectedRowsIndexList)

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

            writeReserveList(oSheet, selectedRowsIndexList)

        ElseIf selectedRowsIndexList.Count > 35 Then
            '2枚作成
            '指定位置にペーストする(2枚目)
            Dim xlPasteRange As Excel.Range = oSheet.Range("B40")
            oSheet.Paste(xlPasteRange)

            oSheet.Range("40:40").RowHeight = 6
            oSheet.Range("41:41").RowHeight = 24
            oSheet.Range("42:42").RowHeight = 54
            oSheet.Range("43:77").RowHeight = 15

            writeReserveList(oSheet, selectedRowsIndexList)
        Else
            writeReserveList(oSheet, selectedRowsIndexList)
        End If

        objExcel.DisplayAlerts = False

        '印刷
        If print.Checked = True Then
            oSheet.PrintOut()
        ElseIf printPreview.Checked = True Then
            objExcel.Visible = True
            oSheet.PrintPreview(1)
        End If

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
            If excelIndex > 104 Then
                rowIndex = 16
            ElseIf excelIndex > 69 Then
                rowIndex = 12
            ElseIf excelIndex > 34 Then
                rowIndex = 8
            End If

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
                If DataGridView1("Tok1", i).Value = "国保" Then
                    oSheet.Range("Y" & (rowIndex + excelIndex)).Value = 3 '採血数
                ElseIf DataGridView1("Tok1", i).Value = "社・家" OrElse DataGridView1("Tok1", i).Value = "共済" Then
                    oSheet.Range("Y" & (rowIndex + excelIndex)).Value = 2 '採血数
                End If

            ElseIf type = "がん" Then
                If DataGridView1("Gan1", i).Value = 1 Then
                    oSheet.Range("Z" & (rowIndex + excelIndex)).Value = 1 '胃がん
                    oSheet.Range("U" & (rowIndex + excelIndex)).Value = 1 '胃Ba
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
        Dim selectedRowsCount As Integer = 0
        For i = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows.Item(i).Selected = True Then
                selectedRowsCount = selectedRowsCount + 1
            End If
        Next

        If selectedRowsCount <> 1 Then
            MsgBox("削除対象の行を１つ選択してください")
            Return
        End If

        Dim index As Integer = DataGridView1.CurrentRow.Index

        Dim birthDay As String = DataGridView1("Birth", index).Value
        Dim name As String = DataGridView1("Nam", index).Value
        Dim reserveDay As String = DataGridView1("Ymd", index).Value

        '削除処理
        Dim cn As New OleDbConnection(Form1.DB_reserve)
        Dim sqlcm As OleDbCommand = cn.CreateCommand
        sqlcm.CommandText = "delete from RsvD where Nam='" & name & "' AND Birth='" & birthDay & "' AND Ymd='" & reserveDay & "'"
        cn.Open()
        sqlcm.ExecuteNonQuery()
        cn.Close()
        cn.Dispose()

        MsgBox("削除しました")
        inputClear()
        tabPageInputClear()
        TabControl1.SelectedTab = referenceTabPage
        reserveListViewReload()

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
            Cn = New OleDbConnection(Form1.DB_diagnose)
            SQLCm = Cn.CreateCommand
            SQLCm.CommandText = "SELECT Nam, Kana FROM UsrM WHERE Ind='" & ind & " 'ORDER BY Kana"
        ElseIf HealthButton.Checked = True Then
            '生活習慣病
            Cn = New OleDbConnection(Form1.DB_health)
            SQLCm = Cn.CreateCommand
            SQLCm.CommandText = "SELECT Nam, Kana FROM UsrM WHERE Ind='" & ind & "' ORDER BY Kana"
        Else
            '産健ｾﾝﾀｰ
            Cn = New OleDbConnection(Form1.DB_reserve)
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
            Cn = New OleDbConnection(Form1.DB_diagnose)
            SQLCm = Cn.CreateCommand
            SQLCm.CommandText = "SELECT Nam, Kana, Birth, Sex, Ind FROM UsrM WHERE Ind='" & selectedInd & "' AND Nam='" & selectedName & "'"
            If selectedInd = "個人健診" Then
                typeBox.Text = "個人"
            ElseIf selectedInd = "特定健診" Then
                typeBox.Text = "特定"
            ElseIf selectedInd = "がん検診" Then
                typeBox.Text = "がん"
            Else
                typeBox.Text = "企業"
            End If
        ElseIf HealthButton.Checked = True Then
            '生活習慣病
            Cn = New OleDbConnection(Form1.DB_health)
            SQLCm = Cn.CreateCommand
            SQLCm.CommandText = "SELECT Nam, Kana, Birth, Sex, Ind FROM UsrM WHERE Ind='" & selectedInd & "' AND Nam='" & selectedName & "'"
            typeBox.Text = "生活"
        Else
            '産健ｾﾝﾀｰ
            Cn = New OleDbConnection(Form1.DB_reserve)
            SQLCm = Cn.CreateCommand
            SQLCm.CommandText = "SELECT Nam, Kana, Birth, Sex, Ind FROM RsvD WHERE Ind='" & selectedInd & "' AND Nam='" & selectedName & "'"
            typeBox.Text = "企業"
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
            If HealthButton.Checked Then
                birthYmdBox.EraText = reader("Birth").Substring(0, 3)
                birthYmdBox.MonthText = reader("Birth").Substring(4, 2)
                birthYmdBox.DateText = reader("Birth").Substring(7, 2)
            Else
                birthYmdBox.setADStr(reader("Birth"))
            End If

        End While
        reader.Close()
        Cn.Close()
        SQLCm.Dispose()
        Cn.Dispose()

    End Sub

    Private Sub inputClear()
        '左の入力エリアのクリア
        typeBox.Text = ""
        companyNameBox.Text = ""
        nameBox.Text = ""
        kanaBox.Text = ""
        sexBox.Text = ""
        birthYmdBox.EraText = ""
        birthYmdBox.MonthText = ""
        birthYmdBox.DateText = ""
        reserveYmdBox.EraText = ""
        reserveYmdBox.MonthText = ""
        reserveYmdBox.DateText = ""
        ampmBox.Text = ""
        resultDayBox.Text = ""
        postBox.Text = ""
        memo1Box.Text = ""
        memo2Box.Text = ""
    End Sub

    Private Sub tabPageInputClear()

        '右の各タブの入力エリアのクリア
        '個人タブ
        personalBlood.Checked = False
        personalElectro.Checked = False
        personalChestXP.Checked = False
        personalUltrasonic.Checked = False
        personalStomachBa.Checked = False
        personalStomachCamera.Checked = False
        personalNone.Checked = False
        personalWindowPay.Text = ""

        '企業タブ
        companyBlood.Checked = False
        companyElectro.Checked = False
        companyChestXP.Checked = False
        companyUltrasonic.Checked = False
        companyStomachBa.Checked = False
        companyStomachCamera.Checked = False
        companyNone.Checked = False
        companyWindowPay.Text = ""

        '生活タブ
        lifeStyleStomachBa.Checked = False
        lifeStyleStomachCamera.Checked = False
        lifeStyleNone.Checked = False
        lifeStyleWindowPay.Text = ""

        '特定タブ
        insuranceTypeBox.Text = ""
        biochemistryBox.Text = ""
        bloodSugarBox.Text = ""
        anemiaBox.Text = ""
        couponTicketBox.Checked = False
        cardiacBox.Text = ""
        gastricCancerRiskBox.Text = ""
        diabetesBox.Text = ""
        prostateCancerBox.Text = ""
        specificWindowPay.Text = ""

        'がんタブ
        gastricCancerBox.Checked = False
        colorectalCancerBox.Checked = False
        cancerWindowPay.Text = ""

        '参照タブ
        If diagnoseButton.Checked = True Then
            HealthButton.Checked = True
            diagnoseButton.Checked = True
        Else
            diagnoseButton.Checked = True
        End If
    End Sub

    Private Sub btnSelectClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelectClear.Click
        Dim rowsCount As Integer = DataGridView1.Rows.Count
        For i = 0 To rowsCount - 1
            DataGridView1.Rows.Item(i).Selected = False
        Next
        tabPageInputClear()
        TabControl1.SelectedTab = referenceTabPage
    End Sub

    Private Sub btnInputClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInputClear.Click
        Dim rowsCount As Integer = DataGridView1.Rows.Count
        For i = 0 To rowsCount - 1
            DataGridView1.Rows.Item(i).Selected = False
        Next
        inputClear()
        tabPageInputClear()
        TabControl1.SelectedTab = referenceTabPage
    End Sub

    Private Sub reserveListViewReload()
        DataGridView1.DataSource = New DataTable
        DataGridView1.Columns.Clear()

        '一覧の再表示
        reloadDataGridView()

        'セルの編集不可
        DataGridView1.ReadOnly = True

        'セルを選択すると行全体が選択されるようにする
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        '予約日と種別と企業名以外の列をソート不可に設定
        For Each c As DataGridViewColumn In DataGridView1.Columns
            If c.Name = "Ymd" OrElse c.Name = "Syu" OrElse c.Name = "Kana" OrElse c.Name = "Ind" Then
                Continue For
            End If
            c.SortMode = DataGridViewColumnSortMode.NotSortable
        Next c

        DataGridView1.AllowUserToAddRows = False
    End Sub

    Private Sub btnUpMonth_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpMonth.Click
        Dim currentMonthStr As String = monthBox.Text
        Dim currentMonthInt As Integer = Integer.Parse(currentMonthStr)
        Dim uppedMonthStr As String = ""
        If currentMonthInt = 12 Then
            Dim eraStr As String = eraBox.Text.Substring(1, 2)
            If eraStr <> "30" Then
                eraBox.Text = eraBox.Items(eraBox.SelectedIndex + 1)
                monthBox.Text = "01"
            End If
        Else
            uppedMonthStr = If((currentMonthInt + 1) >= 10, (currentMonthInt + 1), "0" & (currentMonthInt + 1))
            monthBox.Text = uppedMonthStr
        End If
    End Sub

    Private Sub btnDownMonth_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDownMonth.Click
        Dim currentMonthStr As String = monthBox.Text
        Dim currentMonthInt As Integer = Integer.Parse(currentMonthStr)
        Dim downedMonthStr As String = ""
        If currentMonthInt = 1 Then
            Dim eraStr As String = eraBox.Text.Substring(1, 2)
            If eraStr <> "22" Then
                eraBox.Text = eraBox.Items(eraBox.SelectedIndex - 1)
                monthBox.Text = "12"
            End If
        Else
            downedMonthStr = If((currentMonthInt - 1) >= 10, (currentMonthInt - 1), "0" & (currentMonthInt - 1))
            monthBox.Text = downedMonthStr
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If lifeStyleStomachBa.Checked = True OrElse lifeStyleStomachCamera.Checked = True Then
            lifeStyleWindowPay.Text = "7038"
        Else
            lifeStyleWindowPay.Text = "3750"
        End If
    End Sub

    Private Sub insuranceTypeBox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles insuranceTypeBox.SelectedIndexChanged
        If insuranceTypeBox.Text = "国保" Then
            biochemistryBox.Text = "○"
            bloodSugarBox.Text = "○"
            anemiaBox.Text = "○"
        ElseIf insuranceTypeBox.Text = "社・家" OrElse insuranceTypeBox.Text = "共済" Then
            biochemistryBox.Text = "○"
            bloodSugarBox.Text = "○"
            anemiaBox.Text = "×"
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim totalPay As Integer = 0

        If insuranceTypeBox.Text = "社・家" Then
            totalPay = totalPay + 1160
        End If

        If cardiacBox.Text = "○" AndAlso couponTicketBox.Checked <> True Then
            totalPay = totalPay + 1550
        End If

        If gastricCancerRiskBox.Text = "○" AndAlso couponTicketBox.Checked <> True Then
            totalPay = totalPay + 3600
        End If

        If diabetesBox.Text = "○" AndAlso couponTicketBox.Checked <> True Then
            totalPay = totalPay + 1400
        End If

        If prostateCancerBox.Text = "○" AndAlso couponTicketBox.Checked <> True Then
            totalPay = totalPay + 1550
        End If

        specificWindowPay.Text = totalPay
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim totalPay As Integer = 0
        Dim age As Integer = 0

        '入力生年月日から年齢取得
        If birthYmdBox.EraText <> "" AndAlso birthYmdBox.MonthText <> "" AndAlso birthYmdBox.DateText <> "" Then
            Dim yyyy As Integer = convertWarekiToAD(birthYmdBox.EraText)
            Dim MM As Integer = Integer.Parse(birthYmdBox.MonthText)
            Dim dd As Integer = Integer.Parse(birthYmdBox.DateText)
            age = GetAge(New Date(yyyy, MM, dd), DateTime.Today)
        End If

        If gastricCancerBox.Checked = True AndAlso age < 70 Then
            totalPay = totalPay + 1000
        End If

        If colorectalCancerBox.Checked = True AndAlso age < 70 Then
            totalPay = totalPay + 1000
        End If

        cancerWindowPay.Text = totalPay
    End Sub

    Private Sub typeBox_TextChanged(sender As Object, e As System.EventArgs) Handles typeBox.TextChanged
        Dim selectedValue As String = typeBox.Text
        If selectedValue = "個人" Then
            TabControl1.SelectedTab = personalTabPage
        ElseIf selectedValue = "企業" Then
            TabControl1.SelectedTab = companyTabPage
        ElseIf selectedValue = "生活" Then
            TabControl1.SelectedTab = lifeStyleTabPage
        ElseIf selectedValue = "がん" Then
            TabControl1.SelectedTab = cancerTabPage
        ElseIf selectedValue = "特定" Then
            TabControl1.SelectedTab = specificTabPage
        End If
    End Sub

    Private Sub btnRegist_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRegist.Click
        '入力内容取得
        Dim type As String = typeBox.Text '種別
        Dim companyName As String = companyNameBox.Text '企業名
        Dim name As String = nameBox.Text '氏名
        Dim kana As String = kanaBox.Text 'カナ
        Dim sex As String = sexBox.Text '性別
        Dim ampm As String = ampmBox.Text 'AmPm
        Dim resultDay As String = resultDayBox.Text '結果渡日
        Dim post As String = postBox.Text '来院郵送
        Dim memo1 As String = memo1Box.Text 'メモ１
        Dim memo2 As String = memo2Box.Text 'メモ２

        '必須項目チェック
        If type = "" Then
            MsgBox("種別が未入力です")
            Return
        ElseIf companyName = "" Then
            MsgBox("企業名が未入力です")
            Return
        ElseIf name = "" Then
            MsgBox("氏名が未入力です")
            Return
        ElseIf sex = "" Then
            MsgBox("性別が未入力です")
            Return
        ElseIf birthYmdBox.EraText = "" OrElse birthYmdBox.MonthText = "" OrElse birthYmdBox.DateText = "" Then
            MsgBox("生年月日が未入力です")
            Return
        ElseIf reserveYmdBox.EraText = "" OrElse reserveYmdBox.MonthText = "" OrElse reserveYmdBox.DateText = "" Then
            MsgBox("予約日が未入力です")
            Return
        End If

        '予約日(yyyy/MM/dd)
        Dim reserveDay As String = reserveYmdBox.getADStr()
        '生年月日(yyyy/MM/dd)
        Dim birthDay As String = birthYmdBox.getADStr()

        'タブページ部分の入力内容
        Dim kjn(5) As Integer
        Dim kig(5) As Integer
        Dim sei(3) As Integer
        Dim tok(7) As String
        Dim tok9 As Integer = 0
        Dim gan(1) As Integer
        Dim sanken As String = ""
        Dim windowPay As Integer = 0

        If type = "個人" Then
            If personalBlood.Checked = True Then
                kjn(0) = 1
            End If

            If personalElectro.Checked = True Then
                kjn(1) = 1
            End If

            If personalChestXP.Checked = True Then
                kjn(2) = 1
            End If

            If personalUltrasonic.Checked = True Then
                kjn(3) = 1
            End If

            If personalStomachBa.Checked = True Then
                kjn(4) = 1
            End If

            If personalStomachCamera.Checked = True Then
                kjn(5) = 1
            End If

            windowPay = If(personalWindowPay.Text = "", 0, Integer.Parse(personalWindowPay.Text))
        ElseIf type = "企業" Then
            If companyBlood.Checked = True Then
                kig(0) = 1
            End If

            If companyElectro.Checked = True Then
                kig(1) = 1
            End If

            If companyChestXP.Checked = True Then
                kig(2) = 1
            End If

            If companyUltrasonic.Checked = True Then
                kig(3) = 1
            End If

            If companyStomachBa.Checked = True Then
                kig(4) = 1
            End If

            If companyStomachCamera.Checked = True Then
                kig(5) = 1
            End If

            windowPay = If(companyWindowPay.Text = "", 0, Integer.Parse(companyWindowPay.Text))
        ElseIf type = "生活" Then
            If lifeStyleStomachBa.Checked = True Then
                sei(2) = 1
            End If

            windowPay = If(lifeStyleWindowPay.Text = "", 0, Integer.Parse(lifeStyleWindowPay.Text))
        ElseIf type = "特定" Then
            tok(0) = insuranceTypeBox.Text
            tok(1) = biochemistryBox.Text
            tok(2) = bloodSugarBox.Text
            tok(3) = anemiaBox.Text
            tok(4) = cardiacBox.Text
            tok(5) = gastricCancerRiskBox.Text
            tok(6) = diabetesBox.Text
            tok(7) = prostateCancerBox.Text
            If couponTicketBox.Checked = True Then
                tok9 = 1
            End If

            windowPay = If(specificWindowPay.Text = "", 0, Integer.Parse(specificWindowPay.Text))

        ElseIf type = "がん" Then
            If gastricCancerBox.Checked = True Then
                gan(0) = 1
            End If

            If colorectalCancerBox.Checked = True Then
                gan(1) = 1
            End If

            windowPay = If(cancerWindowPay.Text = "", 0, Integer.Parse(cancerWindowPay.Text))
        End If


        Dim Cn As New OleDbConnection(Form1.DB_reserve)
        Dim SQLCm As OleDbCommand = Cn.CreateCommand
        Dim reader As System.Data.OleDb.OleDbDataReader
        Dim SQL As String = ""

        SQL = "select top 1 * from RsvD where Ymd='" & reserveDay & "' and Nam='" & name & "' and Birth='" & birthDay & "'"
        SQLCm.CommandText = SQL
        Cn.Open()
        reader = SQLCm.ExecuteReader()
        Dim changeFlg As Boolean = reader.Read()
        Cn.Close()
        If changeFlg = False Then
            '新規登録
            SQL = "INSERT INTO RsvD VALUES("
            SQL &= "'" & reserveDay & "', "
            SQL &= "'" & ampm & "', "
            SQL &= "'" & type & "', "
            SQL &= "'" & name & "', "
            SQL &= "'" & kana & "', "
            SQL &= "'" & sex & "',"
            SQL &= "'" & birthDay & "',"
            SQL &= "'" & companyName & "',"
            SQL &= "'" & resultDay & "',"
            SQL &= "'" & post & "',"
            SQL &= "'" & memo1 & "',"
            SQL &= "'" & memo2 & "',"
            SQL &= "" & kjn(0) & ","
            SQL &= "" & kjn(1) & ","
            SQL &= "" & kjn(2) & ","
            SQL &= "" & kjn(3) & ","
            SQL &= "" & kjn(4) & ","
            SQL &= "" & kjn(5) & ","
            SQL &= "" & kig(0) & ","
            SQL &= "" & kig(1) & ","
            SQL &= "" & kig(2) & ","
            SQL &= "" & kig(3) & ","
            SQL &= "" & kig(4) & ","
            SQL &= "" & kig(5) & ","
            SQL &= "" & sei(0) & ","
            SQL &= "" & sei(1) & ","
            SQL &= "" & sei(2) & ","
            SQL &= "" & sei(3) & ","
            SQL &= "'" & tok(0) & "',"
            SQL &= "'" & tok(1) & "',"
            SQL &= "'" & tok(2) & "',"
            SQL &= "'" & tok(3) & "',"
            SQL &= "'" & tok(4) & "',"
            SQL &= "'" & tok(5) & "',"
            SQL &= "'" & tok(6) & "',"
            SQL &= "'" & tok(7) & "',"
            SQL &= "" & tok9 & ","
            SQL &= "" & gan(0) & ","
            SQL &= "" & gan(1) & ","
            SQL &= "" & windowPay & ","
            SQL &= "'" & sanken & "'"
            SQL &= ")"

            SQLCm.CommandText = SQL
            Cn.Open()
            SQLCm.ExecuteNonQuery()

            Cn.Close()
            SQLCm.Dispose()
            Cn.Dispose()

            MsgBox("登録しました")
            btnSelectClear.PerformClick()
            reserveListViewReload()

        Else
            '更新処理
            SQL = "UPDATE RsvD SET "
            SQL &= "Ymd='" & reserveDay & "', "
            SQL &= "Apm='" & ampm & "', "
            SQL &= "Syu='" & type & "', "
            SQL &= "Nam='" & name & "', "
            SQL &= "Kana='" & kana & "', "
            SQL &= "Sex='" & sex & "',"
            SQL &= "Birth='" & birthDay & "',"
            SQL &= "Ind='" & companyName & "',"
            SQL &= "Ymd2='" & resultDay & "',"
            SQL &= "Send='" & post & "',"
            SQL &= "Memo1='" & memo1 & "',"
            SQL &= "Memo2='" & memo2 & "',"
            SQL &= "Kjn1=" & kjn(0) & ","
            SQL &= "Kjn2=" & kjn(1) & ","
            SQL &= "Kjn3=" & kjn(2) & ","
            SQL &= "Kjn4=" & kjn(3) & ","
            SQL &= "Kjn5=" & kjn(4) & ","
            SQL &= "Kjn6=" & kjn(5) & ","
            SQL &= "Kig1=" & kig(0) & ","
            SQL &= "Kig2=" & kig(1) & ","
            SQL &= "Kig3=" & kig(2) & ","
            SQL &= "Kig4=" & kig(3) & ","
            SQL &= "Kig5=" & kig(4) & ","
            SQL &= "Kig6=" & kig(5) & ","
            SQL &= "Sei1=" & sei(0) & ","
            SQL &= "Sei2=" & sei(1) & ","
            SQL &= "Sei3=" & sei(2) & ","
            SQL &= "Sei4=" & sei(3) & ","
            SQL &= "Tok1='" & tok(0) & "',"
            SQL &= "Tok2='" & tok(1) & "',"
            SQL &= "Tok3='" & tok(2) & "',"
            SQL &= "Tok4='" & tok(3) & "',"
            SQL &= "Tok5='" & tok(4) & "',"
            SQL &= "Tok6='" & tok(5) & "',"
            SQL &= "Tok7='" & tok(6) & "',"
            SQL &= "Tok8='" & tok(7) & "',"
            SQL &= "Tok9=" & tok9 & ","
            SQL &= "Gan1=" & gan(0) & ","
            SQL &= "Gan2=" & gan(1) & ","
            SQL &= "Futan=" & windowPay & ","
            SQL &= "Sanken='" & sanken & "'"
            SQL &= "WHERE "
            SQL &= "Ymd='" & reserveDay & "' AND Nam='" & name & "' AND Birth='" & birthDay & "'"

            SQLCm.CommandText = SQL
            Cn.Open()
            SQLCm.ExecuteNonQuery()

            Cn.Close()
            SQLCm.Dispose()
            Cn.Dispose()

            MsgBox("変更しました")
            btnSelectClear.PerformClick()
            reserveListViewReload()
        End If

    End Sub

End Class