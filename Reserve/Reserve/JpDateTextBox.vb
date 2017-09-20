Imports System.ComponentModel

''' <summary>
''' YYYYMMDD形式で入力し和暦形式で表示する
''' </summary>
''' <remarks></remarks>
Public Class JpDateTextBox
    Inherits TextBox

    Private _calendar As System.Globalization.JapaneseCalendar
    Private _culture As System.Globalization.CultureInfo

    ''' <summary>入力用の書式パターン</summary>
    Private Const _inputFormat As String = "yyyyMMdd"

    ''' <summary>表示用の書式パターン</summary>
    Private _displayFormat As String = "gg yy年MM月dd日 (ddd)"

    ''' <summary></summary>
    Private _value As Nullable(Of Date)

    ''' <summary>
    '''
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

        MyBase.New()

        ' 和暦表示用のカレンダー/カルチャーを初期化
        _calendar = New System.Globalization.JapaneseCalendar()
        _culture = New System.Globalization.CultureInfo("ja-JP")
        _culture.DateTimeFormat.Calendar = _calendar

        ' ハンドラの定義
        AddHandler Me.GotFocus, AddressOf HandlerFocusChanged
        AddHandler Me.LostFocus, AddressOf HandlerFocusChanged
        AddHandler Me.TextChanged, AddressOf HandlerValueChanged

    End Sub

    ''' <summary>
    ''' GotFocus/LostFocusのイベント発生時にハンドラでフォーマットの切り替えを行う
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub HandlerFocusChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        ' コントロールにフォーカスがあるときは入力用のフォーマットに、
        ' ないときは表示用のフォーマットに
        If Me.Focused Then
            If _value IsNot Nothing Then
                Me.Text = Format(_value, _inputFormat)
            End If
        Else
            If _value IsNot Nothing Then
                Me.Text = _value.Value.ToString(_displayFormat, _culture)
            End If
        End If

    End Sub

    ''' <summary>
    '''
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub HandlerValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        ' コントロールにフォーカスがあるときは表示の更新をパスする
        If Me.Focused Then
            Return
        End If

        ' コントロールにフォーカスがないときは表示の更新を行う
        If _value IsNot Nothing Then
            Me.Text = _value.Value.ToString(_displayFormat, _culture)
        End If

    End Sub

    ''' <summary>
    '''
    ''' </summary>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Overrides Sub OnKeyPress(ByVal e As System.Windows.Forms.KeyPressEventArgs)

        If (e.KeyChar >= "0"c And e.KeyChar <= "9"c) Then

            If Me.SelectionLength = 0 AndAlso Text.Length >= 8 Then
                e.Handled = True
                Return
            End If

            Return
        End If

        If e.KeyChar = ControlChars.Back _
        OrElse e.KeyChar = ControlChars.Tab Then
            Return
        End If

        ' Enterキーを押されたタイミングで日付として認識可能かパースを試みる
        If e.KeyChar = ControlChars.Cr Then
            _value = Parse(Me.Text)
            Return
        End If

        ' 0～9、Back、Tab、Enter以外の時はキー入力を無視
        e.Handled = True

    End Sub

    ''' <summary>
    '''
    ''' </summary>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Overrides Sub OnValidating(ByVal e As System.ComponentModel.CancelEventArgs)
        ' 日付として認識可能かパースを試みる
        _value = Parse(Me.Text)
        MyBase.OnValidating(e)
    End Sub


    ''' <summary>
    ''' 文字列のDate化
    ''' </summary>
    ''' <param name="s"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function Parse(ByVal s As String) As Nullable(Of Date)

        Try

            ' 8文字(YYYYMMDD)未満の場合、日付として認識しない
            If s.Length < 8 Then
                Return Nothing
            End If

            ' YYYYMMDDを前提としてパース
            Return DateTime.Parse(s.Substring(0, 4) & "/" & s.Substring(4, 2) & "/" & s.Substring(6, 2))

        Catch ex As Exception
            Return Nothing
        End Try

    End Function

#Region "プロパティ"

    ''' <summary>
    '''
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shadows Property Value As Nullable(Of Date)
        Get
            If _value Is Nothing Then
                Return Nothing
            End If
            Return _value
        End Get

        Set(ByVal value As Nullable(Of Date))
            _value = value
            Me.Text = If(value Is Nothing, "", _value.Value.ToString(_displayFormat, _culture))
        End Set
    End Property


    ''' <summary>
    ''' 表示書式
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shadows Property DisplayFormat As String
        Get
            Return _displayFormat
        End Get

        Set(ByVal value As String)
            _displayFormat = value
            Me.Text = If(_value Is Nothing, "", _value.Value.ToString(_displayFormat, _culture))
        End Set
    End Property

#End Region

End Class
