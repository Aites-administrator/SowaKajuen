Imports System.Data
Imports System.Windows.Forms
Imports System.Text
Imports System.Text.RegularExpressions
Imports Common
Imports Common.ClsCommonGlobalData
Imports Common.ClsFunction
Imports ClsAutoCommunication.ClsAutoCommunication
Imports ClsPrintingProcess.ClsPrintingProcess

Public Class FrmKeiryokiMasterOutput

  Private _service As KeiryokiMasterOutputService
  Private _dtTokuisaki As DataTable
  Private _dtShohin As DataTable
  Private _dtTanto As DataTable
  Private _sqlServer As New ClsSqlServer
  Private _masterImportCompleted As Boolean = False

  ''' <summary>
  ''' 画面起動時の初期処理
  ''' </summary>
  ''' <param name="sender">イベント発生元</param>
  ''' <param name="e">イベント引数</param>
  Private Sub FrmKeiryokiMasterOutput_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Try
      _service = New KeiryokiMasterOutputService()

      InitializeForm()
      InitializeGrid()
      SearchAllData()

    Catch ex As Exception
      ComWriteErrLog(ex, False)
    End Try
  End Sub

  ''' <summary>
  ''' フォームの初期設定
  ''' </summary>
  Private Sub InitializeForm()
    Me.Text = "計量器マスタ出力"
    Me.StartPosition = FormStartPosition.CenterScreen
    Me.KeyPreview = True
  End Sub

  ''' <summary>
  ''' 各マスタのグリッド初期設定
  ''' </summary>
  Private Sub InitializeGrid()
    InitializeTokuisakiGrid()
    InitializeShohinGrid()
    InitializeTantoGrid()
  End Sub

  ''' <summary>
  ''' 仕入先マスタグリッドの初期設定
  ''' </summary>
  Private Sub InitializeTokuisakiGrid()
    dgvTokuisaki.AutoGenerateColumns = False
    dgvTokuisaki.Columns.Clear()
    SetupGridCommon(dgvTokuisaki)

    AddTextColumn(dgvTokuisaki, "TokuiCD", "仕入先コード", 150)
    AddTextColumn(dgvTokuisaki, "TokuiNM1", "仕入先名", 220)
    AddTextColumn(dgvTokuisaki, "Add1", "住所１", 240)
    AddTextColumn(dgvTokuisaki, "Add2", "住所２", 180)
    AddTextColumn(dgvTokuisaki, "ZipCD", "郵便番号", 130)
    AddTextColumn(dgvTokuisaki, "TelNo", "電話番号", 140)
    AddTextColumn(dgvTokuisaki, "FaxNo", "FAX番号", 140)
    AddTextColumn(dgvTokuisaki, "TokuiInvoiceNumber", "事業者登録番号", 180)
  End Sub

  ''' <summary>
  ''' 商品マスタグリッドの初期設定
  ''' </summary>
  Private Sub InitializeShohinGrid()
    dgvShohin.AutoGenerateColumns = False
    dgvShohin.Columns.Clear()
    SetupGridCommon(dgvShohin)

    AddTextColumn(dgvShohin, "ShohinCD", "商品コード", 150)
    AddTextColumn(dgvShohin, "ShohinNM", "商品名", 280)
    AddTextColumn(dgvShohin, "Tani", "単位", 120)
    AddTextColumn(dgvShohin, "ZeikomiKBN", "税込区分", 120)
    AddTextColumn(dgvShohin, "HyojunKakaku", "標準価格", 150)
  End Sub

  ''' <summary>
  ''' 担当者マスタグリッドの初期設定
  ''' </summary>
  Private Sub InitializeTantoGrid()
    dgvTanto.AutoGenerateColumns = False
    dgvTanto.Columns.Clear()
    SetupGridCommon(dgvTanto)

    AddTextColumn(dgvTanto, "CODE", "担当者コード", 180)
    AddTextColumn(dgvTanto, "NAME", "担当者名", 280)
  End Sub

  ''' <summary>
  ''' グリッド共通設定
  ''' </summary>
  ''' <param name="grid">対象DataGridView</param>
  Private Sub SetupGridCommon(grid As DataGridView)
    grid.DefaultCellStyle.Font = New Font("Segoe UI", 15)
    grid.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 15)
    grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
    grid.ColumnHeadersHeight = 50
    grid.RowTemplate.Height = 50
    grid.AllowUserToAddRows = False
    grid.AllowUserToDeleteRows = False
    grid.RowHeadersVisible = False
    grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    grid.MultiSelect = False
    grid.EditMode = DataGridViewEditMode.EditOnEnter
  End Sub

  ''' <summary>
  ''' グリッドへテキスト列を追加
  ''' </summary>
  ''' <param name="grid">対象DataGridView</param>
  ''' <param name="dataPropertyName">DataTable列名</param>
  ''' <param name="headerText">グリッド列ヘッダ表示名</param>
  ''' <param name="width">列幅</param>
  Private Sub AddTextColumn(grid As DataGridView, dataPropertyName As String, headerText As String, width As Integer)
    Dim col As New DataGridViewTextBoxColumn()
    col.Name = dataPropertyName
    col.HeaderText = headerText
    col.DataPropertyName = dataPropertyName
    col.ReadOnly = False
    col.Width = width
    grid.Columns.Add(col)
  End Sub

  ''' <summary>
  ''' グリッド編集時の入力コントロール設定
  ''' </summary>
  Private Sub Grid_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) _
      Handles dgvTokuisaki.EditingControlShowing, dgvShohin.EditingControlShowing, dgvTanto.EditingControlShowing

    Dim textBox As DataGridViewTextBoxEditingControl = TryCast(e.Control, DataGridViewTextBoxEditingControl)
    If textBox Is Nothing Then Return

    RemoveHandler textBox.KeyPress, AddressOf GridEditingControl_KeyPress
    RemoveHandler textBox.TextChanged, AddressOf GridEditingControl_TextChanged

    Dim grid As DataGridView = TryCast(sender, DataGridView)
    If grid Is Nothing OrElse grid.CurrentCell Is Nothing Then Return

    Dim columnName As String = grid.CurrentCell.OwningColumn.Name
    Dim maxBytes As Integer = GetColumnMaxLength(grid, columnName)

    ' TextBox.MaxLengthは文字数制御なので、十分大きな値にしておき、
    ' 実際の上限はTextChangedでShift-JISのバイト数として制御する。
    textBox.MaxLength = Integer.MaxValue
    textBox.Tag = grid.Name & "|" & columnName

    textBox.Tag = grid.Name & "|" & grid.CurrentCell.OwningColumn.Name
    AddHandler textBox.KeyPress, AddressOf GridEditingControl_KeyPress
    AddHandler textBox.TextChanged, AddressOf GridEditingControl_TextChanged
  End Sub

  ''' <summary>
  ''' グリッド編集中のバイト数制御
  ''' </summary>
  ''' <param name="sender">イベント発生元</param>
  ''' <param name="e">イベント引数</param>
  Private Sub GridEditingControl_TextChanged(sender As Object, e As EventArgs)

    Dim textBox As DataGridViewTextBoxEditingControl =
        TryCast(sender, DataGridViewTextBoxEditingControl)
    If textBox Is Nothing OrElse textBox.Tag Is Nothing Then Return

    Dim tag As String = textBox.Tag.ToString()
    Dim separatorIndex As Integer = tag.IndexOf("|"c)
    If separatorIndex < 0 Then Return

    Dim gridName As String = tag.Substring(0, separatorIndex)
    Dim columnName As String = tag.Substring(separatorIndex + 1)

    Dim maxBytes As Integer = GetColumnMaxLengthByGridName(gridName, columnName)
    If maxBytes = Integer.MaxValue Then Return

    Dim value As String = If(textBox.Text, String.Empty)
    Dim enc As Encoding = Encoding.GetEncoding(932)

    If enc.GetByteCount(value) <= maxBytes Then Return

    Dim selectionStart As Integer = textBox.SelectionStart
    Dim selectionLength As Integer = textBox.SelectionLength

    Dim trimmed As String = TrimToByteLength(value, maxBytes)

    If trimmed = value Then Return

    textBox.Text = trimmed
    textBox.SelectionStart = Math.Min(selectionStart, textBox.Text.Length)
    textBox.SelectionLength = Math.Min(selectionLength,
                                       Math.Max(0, textBox.Text.Length - textBox.SelectionStart))
  End Sub

  ''' <summary>
  ''' グリッド・列に対応する最大バイト数を取得
  ''' </summary>
  ''' <param name="gridName">DataGridView名</param>
  ''' <param name="columnName">対象列名</param>
  ''' <returns>対象列の最大バイト数</returns>
  Private Function GetColumnMaxLengthByGridName(gridName As String, columnName As String) As Integer

    Select Case gridName
      Case "dgvTokuisaki"
        Select Case columnName
          Case "TokuiCD" : Return 6
          Case "TokuiNM1" : Return 100
          Case "Add1" : Return 40
          Case "Add2" : Return 30
          Case "ZipCD" : Return 8
          Case "TelNo" : Return 13
          Case "FaxNo" : Return 13
          Case "TokuiInvoiceNumber" : Return 14
        End Select

      Case "dgvShohin"
        Select Case columnName
          Case "ShohinCD" : Return 6
          Case "ShohinNM" : Return 36
          Case "Tani" : Return 4
          Case "ZeikomiKBN" : Return 1
          Case "HyojunKakaku" : Return 9
        End Select

      Case "dgvTanto"
        Select Case columnName
          Case "CODE" : Return 2
          Case "NAME" : Return 20
        End Select
    End Select

    Return Integer.MaxValue
  End Function

  ''' <summary>
  ''' 指定バイト数以内に文字列を切り詰める
  ''' </summary>
  ''' <param name="value">入力値</param>
  ''' <param name="maxBytes">最大バイト数</param>
  ''' <returns>指定バイト数以内に切り詰めた文字列</returns>
  Private Function TrimToByteLength(value As String, maxBytes As Integer) As String

    If String.IsNullOrEmpty(value) Then Return String.Empty

    Dim enc As Encoding = Encoding.GetEncoding(932)
    Dim sb As New StringBuilder()

    For Each c As Char In value
      Dim test As String = sb.ToString() & c

      If enc.GetByteCount(test) > maxBytes Then
        Exit For
      End If

      sb.Append(c)
    Next

    Return sb.ToString()
  End Function

  ''' <summary>
  ''' グリッド・列に対応する最大バイト数を取得
  ''' </summary>
  ''' <param name="grid">対象DataGridView</param>
  ''' <param name="columnName">対象列名</param>
  ''' <returns>対象列の最大バイト数</returns>
  Private Function GetColumnMaxLength(grid As DataGridView, columnName As String) As Integer
    Select Case grid.Name
      Case "dgvTokuisaki"
        Select Case columnName
          Case "TokuiCD" : Return 6
          Case "TokuiNM1" : Return 100
          Case "Add1" : Return 40
          Case "Add2" : Return 30
          Case "ZipCD" : Return 8
          Case "TelNo" : Return 13
          Case "FaxNo" : Return 13
          Case "TokuiInvoiceNumber" : Return 14
        End Select

      Case "dgvShohin"
        Select Case columnName
          Case "ShohinCD" : Return 6
          Case "ShohinNM" : Return 50
          Case "Tani" : Return 4
          Case "ZeikomiKBN" : Return 1
          Case "HyojunKakaku" : Return 9
        End Select

      Case "dgvTanto"
        Select Case columnName
          Case "CODE" : Return 2
          Case "NAME" : Return 20
        End Select
    End Select

    Return 32767
  End Function

  ''' <summary>
  ''' グリッド編集中の入力文字制御
  ''' </summary>
  ''' <param name="sender">イベント発生元</param>
  ''' <param name="e">イベント引数</param>
  Private Sub GridEditingControl_KeyPress(sender As Object, e As KeyPressEventArgs)
    If Char.IsControl(e.KeyChar) Then Return

    Dim textBox As DataGridViewTextBoxEditingControl = TryCast(sender, DataGridViewTextBoxEditingControl)
    If textBox Is Nothing OrElse textBox.Tag Is Nothing Then Return

    Dim tag As String = textBox.Tag.ToString()
    Dim separatorIndex As Integer = tag.IndexOf("|")
    If separatorIndex < 0 Then Return

    Dim gridName As String = tag.Substring(0, separatorIndex)
    Dim columnName As String = tag.Substring(separatorIndex + 1)

    Select Case gridName & "|" & columnName
      Case "dgvTokuisaki|TokuiCD", "dgvTokuisaki|ZipCD",
           "dgvShohin|ShohinCD", "dgvShohin|ZeikomiKBN", "dgvShohin|HyojunKakaku",
           "dgvTanto|CODE"
        If e.KeyChar < "0"c OrElse e.KeyChar > "9"c Then e.Handled = True

      Case "dgvTokuisaki|TelNo", "dgvTokuisaki|FaxNo"
        If (e.KeyChar < "0"c OrElse e.KeyChar > "9"c) AndAlso e.KeyChar <> "-"c Then
          e.Handled = True
        End If

      Case "dgvTokuisaki|TokuiInvoiceNumber"
        If Not ((e.KeyChar >= "0"c AndAlso e.KeyChar <= "9"c) OrElse
                (e.KeyChar >= "A"c AndAlso e.KeyChar <= "Z"c) OrElse
                (e.KeyChar >= "a"c AndAlso e.KeyChar <= "z"c)) Then
          e.Handled = True
        End If
    End Select
  End Sub

  ''' <summary>
  ''' グリッドセル確定前の入力チェック
  ''' </summary>
  Private Sub Grid_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) _
      Handles dgvTokuisaki.CellValidating, dgvShohin.CellValidating, dgvTanto.CellValidating

    Dim grid As DataGridView = TryCast(sender, DataGridView)
    If grid Is Nothing OrElse e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Return

    Dim columnName As String = grid.Columns(e.ColumnIndex).Name
    Dim value As String = If(e.FormattedValue, "").ToString()
    If value Is Nothing Then value = ""

    Dim maxLength As Integer = GetColumnMaxLength(grid, columnName)
    Dim byteLength As Integer = Encoding.GetEncoding(932).GetByteCount(value)

    If byteLength > maxLength Then
      MessageBox.Show(grid.Columns(columnName).HeaderText & "は" & maxLength.ToString() & "バイト以内で入力してください。",
                      "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning)
      e.Cancel = True
      Return
    End If

    Dim errorMessage As String = GetInvalidCharacterMessage(grid, columnName, value)
    If errorMessage <> "" Then
      MessageBox.Show(errorMessage, "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning)
      e.Cancel = True
      Return
    End If

    ' コード重複チェック
    Select Case grid.Name & "|" & columnName
      Case "dgvTokuisaki|TokuiCD"
        Dim tokuiCode As String = value.PadLeft(CUSTOMER_ZERO_LENGTH, "0"c)
        If IsDuplicateCode(grid, e.RowIndex, "TokuiCD", tokuiCode) Then
          MessageBox.Show("仕入先コード「" & tokuiCode & "」は既に登録されています。",
                          "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning)
          e.Cancel = True
          Return
        End If

      Case "dgvShohin|ShohinCD"
        Dim shohinCode As String = value.PadLeft(ITEM_ZERO_LENGTH, "0"c)
        If IsDuplicateCode(grid, e.RowIndex, "ShohinCD", shohinCode) Then
          MessageBox.Show("商品コード「" & shohinCode & "」は既に登録されています。",
                          "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning)
          e.Cancel = True
          Return
        End If

      Case "dgvTanto|CODE"
        Dim tantoCode As String = value.PadLeft(TANTO_ZERO_LENGTH, "0"c)
        If IsDuplicateCode(grid, e.RowIndex, "CODE", tantoCode) Then
          MessageBox.Show("担当者コード「" & tantoCode & "」は既に登録されています。",
                          "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning)
          e.Cancel = True
          Return
        End If
    End Select

    ' コードは入力確定時に0埋めする。
    If value <> "" Then
      Select Case grid.Name & "|" & columnName
        Case "dgvTokuisaki|TokuiCD"
          grid.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = value.PadLeft(CUSTOMER_ZERO_LENGTH, "0"c)
        Case "dgvShohin|ShohinCD"
          grid.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = value.PadLeft(ITEM_ZERO_LENGTH, "0"c)
        Case "dgvTanto|CODE"
          grid.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = value.PadLeft(TANTO_ZERO_LENGTH, "0"c)
      End Select
    End If
  End Sub

  ''' <summary>
  ''' グリッドセル確定後のコード0埋め処理
  ''' </summary>
  Private Sub Grid_CellValidated(sender As Object, e As DataGridViewCellEventArgs) _
    Handles dgvTokuisaki.CellValidated, dgvShohin.CellValidated, dgvTanto.CellValidated

    Dim grid As DataGridView = TryCast(sender, DataGridView)
    If grid Is Nothing OrElse e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Return

    Dim columnName As String = grid.Columns(e.ColumnIndex).Name

    Dim rowView As DataGridViewRow = grid.Rows(e.RowIndex)

    Dim value As String = If(rowView.Cells(e.ColumnIndex).Value, "").ToString().Trim()
    If value = "" Then Return

    Select Case grid.Name & "|" & columnName

      Case "dgvTokuisaki|TokuiCD"
        rowView.Cells(e.ColumnIndex).Value = value.PadLeft(6, "0"c)

      Case "dgvShohin|ShohinCD"
        rowView.Cells(e.ColumnIndex).Value = value.PadLeft(6, "0"c)

      Case "dgvTanto|CODE"
        rowView.Cells(e.ColumnIndex).Value = value.PadLeft(2, "0"c)

    End Select

  End Sub

  ''' <summary>
  ''' 入力値の文字種エラーメッセージを取得
  ''' </summary>
  ''' <param name="grid">対象DataGridView</param>
  ''' <param name="columnName">対象列名</param>
  ''' <param name="value">入力値</param>
  ''' <returns>入力内容に応じたエラーメッセージ。問題がない場合は空文字列</returns>
  Private Function GetInvalidCharacterMessage(grid As DataGridView, columnName As String, value As String) As String
    If value = "" Then Return ""

    Select Case grid.Name & "|" & columnName
      Case "dgvTokuisaki|TokuiCD", "dgvTokuisaki|ZipCD",
           "dgvShohin|ShohinCD", "dgvShohin|ZeikomiKBN", "dgvShohin|HyojunKakaku",
           "dgvTanto|CODE"
        If Not Regex.IsMatch(value, "^[0-9]+$") Then
          Return grid.Columns(columnName).HeaderText & "は数値のみ入力できます。"
        End If

      Case "dgvTokuisaki|TelNo", "dgvTokuisaki|FaxNo"
        If Not Regex.IsMatch(value, "^[0-9-]+$") Then
          Return grid.Columns(columnName).HeaderText & "は数値と-のみ入力できます。"
        End If

      Case "dgvTokuisaki|TokuiInvoiceNumber"
        If Not Regex.IsMatch(value, "^[A-Za-z0-9]+$") Then
          Return grid.Columns(columnName).HeaderText & "は半角英数字のみ入力できます。"
        End If
    End Select

    Return ""
  End Function

  ''' <summary>
  ''' 全マスタデータを取得して表示
  ''' </summary>
  Private Sub SearchAllData()
    dgvTokuisaki.EndEdit()
    dgvShohin.EndEdit()
    dgvTanto.EndEdit()

    _dtTokuisaki = _service.GetTokuisakiList(_sqlServer)
    _dtShohin = _service.GetShohinList(_sqlServer)
    _dtTanto = _service.GetTantoList(_sqlServer)

    dgvTokuisaki.DataSource = _dtTokuisaki
    dgvShohin.DataSource = _dtShohin
    dgvTanto.DataSource = _dtTanto
  End Sub

  ''' <summary>
  ''' 仕入先マスタを検索して表示
  ''' </summary>
  Private Sub SearchTokuisaki()
    txtTokuisakiCd.Text = If(String.IsNullOrWhiteSpace(txtTokuisakiCd.Text), "", txtTokuisakiCd.Text.PadLeft(CUSTOMER_ZERO_LENGTH, "0"c))
    dgvTokuisaki.EndEdit()
    ApplyFilter(_dtTokuisaki, txtTokuisakiCd.Text.Trim(), txtTokuisakiNm.Text.Trim(), "TokuiCD", "TokuiNM1", dgvTokuisaki)
  End Sub

  ''' <summary>
  ''' 商品マスタを検索して表示
  ''' </summary>
  Private Sub SearchShohin()
    txtShohinCd.Text = If(String.IsNullOrWhiteSpace(txtShohinCd.Text), "", txtShohinCd.Text.PadLeft(ITEM_ZERO_LENGTH, "0"c))
    dgvShohin.EndEdit()
    ApplyFilter(_dtShohin, txtShohinCd.Text.Trim(), txtShohinNm.Text.Trim(), "ShohinCD", "ShohinNM", dgvShohin)
  End Sub

  ''' <summary>
  ''' 担当者マスタを検索して表示
  ''' </summary>
  Private Sub SearchTanto()
    txtTantoCd.Text = If(String.IsNullOrWhiteSpace(txtTantoCd.Text), "", txtTantoCd.Text.PadLeft(TANTO_ZERO_LENGTH, "0"c))
    dgvTanto.EndEdit()
    ApplyFilter(_dtTanto, txtTantoCd.Text.Trim(), txtTantoNm.Text.Trim(), "CODE", "NAME", dgvTanto)
  End Sub

  ''' <summary>
  ''' 指定条件でグリッドデータを絞り込む
  ''' </summary>
  ''' <param name="dt">対象DataTable</param>
  ''' <param name="code">確認するコード</param>
  ''' <param name="name">nameの値</param>
  ''' <param name="codeColumn">codeColumnの値</param>
  ''' <param name="nameColumn">nameColumnの値</param>
  ''' <param name="grid">対象DataGridView</param>
  Private Sub ApplyFilter(dt As DataTable,
                           code As String,
                           name As String,
                           codeColumn As String,
                           nameColumn As String,
                           grid As DataGridView)
    If dt Is Nothing Then Return

    Dim filters As New List(Of String)
    If code <> "" Then
      filters.Add("[" & codeColumn & "] LIKE '%" & EscapeFilter(code) & "%'")
    End If
    If name <> "" Then
      filters.Add("[" & nameColumn & "] LIKE '%" & EscapeFilter(name) & "%'")
    End If

    Dim view As DataView = dt.DefaultView
    view.RowFilter = String.Join(" AND ", filters)
    grid.DataSource = view
  End Sub

  ''' <summary>
  ''' DataViewの検索条件用に文字列をエスケープ
  ''' </summary>
  ''' <param name="value">入力値</param>
  ''' <returns>DataViewの検索条件用にエスケープした文字列</returns>
  Private Function EscapeFilter(value As String) As String
    If value Is Nothing Then Return ""
    Return value.Replace("'", "''").Replace("[", "[[").Replace("]", "]]")
  End Function

  ''' <summary>
  ''' 現在選択中のマスタに行を追加
  ''' </summary>
  ''' <param name="sender">イベント発生元</param>
  ''' <param name="e">イベント引数</param>
  Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
    Try
      Dim grid As DataGridView = GetCurrentGrid()
      Dim dt As DataTable = GetCurrentDataTable()
      If dt Is Nothing Then Return

      ' 検索中でも追加行を確認・編集できるよう、追加時は全件表示に戻す。
      dt.DefaultView.RowFilter = String.Empty
      grid.DataSource = dt.DefaultView

      Dim row As DataRow = dt.NewRow()
      dt.Rows.Add(row)

      grid.DataSource = dt.DefaultView
      grid.ClearSelection()
      grid.CurrentCell = grid.Rows(grid.Rows.Count - 1).Cells(0)
      grid.CurrentRow.Selected = True
      grid.BeginEdit(True)

    Catch ex As Exception
      ComWriteErrLog(ex, False)
    End Try
  End Sub

  ''' <summary>
  ''' 現在選択中のマスタの行を削除
  ''' </summary>
  ''' <param name="sender">イベント発生元</param>
  ''' <param name="e">イベント引数</param>
  Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
    Try
      Dim grid As DataGridView = GetCurrentGrid()
      Dim dt As DataTable = GetCurrentDataTable()
      If dt Is Nothing OrElse grid.CurrentRow Is Nothing Then Return

      Dim rowView As DataRowView = TryCast(grid.CurrentRow.DataBoundItem, DataRowView)
      If rowView Is Nothing Then Return

      rowView.Row.Delete()
      grid.Refresh()

    Catch ex As Exception
      ComWriteErrLog(ex, False)
    End Try
  End Sub

  ''' <summary>
  ''' 計量器マスタCSVを出力してマスタ取込・PC側更新を実行
  ''' </summary>
  ''' <param name="sender">イベント発生元</param>
  ''' <param name="e">イベント引数</param>
  Private Sub btnOutput_Click(sender As Object, e As EventArgs) Handles btnOutput.Click
    Try
      dgvTokuisaki.EndEdit()
      dgvShohin.EndEdit()
      dgvTanto.EndEdit()

      If MessageBox.Show("計量器マスタCSVを出力します。" & vbCrLf & "よろしいでしょうか。",
                         "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) = DialogResult.Cancel Then
        Return
      End If

      ValidateCodes()

      Dim basePath As String = ReadSettingIniFile("FTP_DOWNLOAD_PATH", "VALUE")
      If String.IsNullOrWhiteSpace(basePath) Then
        Throw New Exception("FTP_DOWNLOAD_PATH が設定されていません。")
      End If

      Dim scaleDt As DataTable = _service.GetScaleList(_sqlServer, "")
      If scaleDt Is Nothing OrElse scaleDt.Rows.Count = 0 Then
        MessageBox.Show("登録されている計量器がありません。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Return
      End If

      LblMessage.Text = "CSV出力中・・・"
      _service.OutputAllCsv(_dtTokuisaki, _dtShohin, _dtTanto, scaleDt, basePath, _sqlServer)
      LblMessage.Text = "CSV出力完了"

      ' 出力したCSVを既存のマスタ取込で取り込む。
      _masterImportCompleted = False
      CallGetMaster()
      If Not _masterImportCompleted Then
        Return
      End If

      ' 計量器取込後、PC側で必要な項目をDBへ反映する。
      _service.UpdatePcFields(_dtTokuisaki, _dtShohin, _dtTanto, _sqlServer)

      MessageBox.Show("CSV出力およびマスタ取込が完了しました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information)

    Catch ex As Exception
      ComWriteErrLog(ex, False)
      MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Try
  End Sub

  ''' <summary>
  ''' 各マスタの入力内容を検証
  ''' </summary>
  Private Sub ValidateCodes()
    ValidateTokuisaki()
    ValidateShohin()
    ValidateTanto()
  End Sub

  ''' <summary>
  ''' 仕入先マスタの入力内容を検証
  ''' </summary>
  Private Sub ValidateTokuisaki()
    If _dtTokuisaki Is Nothing Then Return

    For Each row As DataRow In _dtTokuisaki.Rows
      If row.RowState = DataRowState.Deleted Then Continue For

      Dim cd As String = GetDataRowString(row, "TokuiCD")
      If cd = "" Then Throw New Exception("仕入先コードが未入力の行があります。")
      If Not Regex.IsMatch(cd, "^[0-9]{1,6}$") Then Throw New Exception("仕入先コードは6桁以内の数値で入力してください。")
      row("TokuiCD") = cd.PadLeft(6, "0"c)

      ValidateTextLength(row, "TokuiNM1", 100, "仕入先名")
      ValidateTextLength(row, "Add1", 40, "住所１")
      ValidateTextLength(row, "Add2", 30, "住所２")

      ValidatePattern(row, "ZipCD", 8, "^[0-9]+$", "郵便番号は8桁以内の数値で入力してください。")
      ValidatePattern(row, "TelNo", 13, "^[0-9-]+$", "電話番号は13桁以内の数値と-で入力してください。")
      ValidatePattern(row, "FaxNo", 13, "^[0-9-]+$", "FAX番号は13桁以内の数値と-で入力してください。")
      ValidatePattern(row, "TokuiInvoiceNumber", 14, "^[A-Za-z0-9]+$", "事業者登録番号は14桁以内の半角英数字で入力してください。")
    Next

    ValidateDuplicateCodes(_dtTokuisaki, "TokuiCD", "仕入先")
  End Sub

  ''' <summary>
  ''' 商品マスタの入力内容を検証
  ''' </summary>
  Private Sub ValidateShohin()
    If _dtShohin Is Nothing Then Return

    For Each row As DataRow In _dtShohin.Rows
      If row.RowState = DataRowState.Deleted Then Continue For

      Dim cd As String = GetDataRowString(row, "ShohinCD")
      If cd = "" Then Throw New Exception("商品コードが未入力の行があります。")
      If Not Regex.IsMatch(cd, "^[0-9]{1,6}$") Then Throw New Exception("商品コードは6桁以内の数値で入力してください。")
      row("ShohinCD") = cd.PadLeft(6, "0"c)

      ValidateTextLength(row, "ShohinNM", 36, "商品名")
      ValidateTextLength(row, "Tani", 4, "単位")
      ValidatePattern(row, "ZeikomiKBN", 1, "^[0-9]+$", "税込区分は1桁の数値で入力してください。")
      ValidatePattern(row, "HyojunKakaku", 9, "^[0-9]+$", "標準価格は9桁以内の数値で入力してください。")
    Next

    ValidateDuplicateCodes(_dtShohin, "ShohinCD", "商品")
  End Sub

  ''' <summary>
  ''' 担当者マスタの入力内容を検証
  ''' </summary>
  Private Sub ValidateTanto()
    If _dtTanto Is Nothing Then Return

    For Each row As DataRow In _dtTanto.Rows
      If row.RowState = DataRowState.Deleted Then Continue For

      Dim cd As String = GetDataRowString(row, "CODE")
      If cd = "" Then Throw New Exception("担当者コードが未入力の行があります。")
      If Not Regex.IsMatch(cd, "^[0-9]{1,2}$") Then Throw New Exception("担当者コードは2桁以内の数値で入力してください。")
      row("CODE") = cd.PadLeft(2, "0"c)

      ValidateTextLength(row, "NAME", 20, "担当者名")
    Next

    ValidateDuplicateCodes(_dtTanto, "CODE", "担当者")
  End Sub

  ''' <summary>
  ''' 指定コードがグリッド内で重複しているか確認
  ''' </summary>
  ''' <param name="grid">対象DataGridView</param>
  ''' <param name="rowIndex">対象行番号</param>
  ''' <param name="columnName">対象列名</param>
  ''' <param name="code">確認するコード</param>
  ''' <returns>コードが重複している場合はTrue、それ以外はFalse</returns>
  Private Function IsDuplicateCode(grid As DataGridView,
                                  rowIndex As Integer,
                                  columnName As String,
                                  code As String) As Boolean
    Try
      Dim dt As DataTable = GetCurrentDataTable()
      If dt Is Nothing Then Return False

      Dim currentRowView As DataRowView =
          TryCast(grid.Rows(rowIndex).DataBoundItem, DataRowView)
      If currentRowView Is Nothing Then Return False

      For Each row As DataRow In dt.Rows
        If row.RowState = DataRowState.Deleted Then Continue For

        ' 編集中の自分自身は比較対象外
        If Object.ReferenceEquals(row, currentRowView.Row) Then Continue For

        Dim existingCode As String = GetDataRowString(row, columnName)
        If existingCode = "" Then Continue For

        Dim normalizedCode As String = code

        Select Case grid.Name & "|" & columnName
          Case "dgvTokuisaki|TokuiCD"
            existingCode = existingCode.PadLeft(CUSTOMER_ZERO_LENGTH, "0"c)

          Case "dgvShohin|ShohinCD"
            existingCode = existingCode.PadLeft(ITEM_ZERO_LENGTH, "0"c)

          Case "dgvTanto|CODE"
            existingCode = existingCode.PadLeft(TANTO_ZERO_LENGTH, "0"c)
        End Select

        If existingCode = normalizedCode Then Return True
      Next

    Catch ex As Exception
      ComWriteErrLog(ex, False)
    End Try

    Return False
  End Function

  ''' <summary>
  ''' マスタ内のコード重複を検証
  ''' </summary>
  ''' <param name="dt">対象DataTable</param>
  ''' <param name="columnName">対象列名</param>
  ''' <param name="targetName">対象マスタ名</param>
  Private Sub ValidateDuplicateCodes(dt As DataTable,
                                     columnName As String,
                                     targetName As String)
    If dt Is Nothing Then Return

    Dim codeSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

    Try
      For Each row As DataRow In dt.Rows
        If row.RowState = DataRowState.Deleted Then Continue For

        Dim code As String = GetDataRowString(row, columnName)
        If code = "" Then Continue For

        Select Case targetName
          Case "仕入先"
            code = code.PadLeft(CUSTOMER_ZERO_LENGTH, "0"c)

          Case "商品"
            code = code.PadLeft(ITEM_ZERO_LENGTH, "0"c)

          Case "担当者"
            code = code.PadLeft(TANTO_ZERO_LENGTH, "0"c)
        End Select

        If Not codeSet.Add(code) Then
          Throw New Exception(targetName & "コード「" & code & "」が重複しています。")
        End If
      Next

    Catch ex As Exception
      Throw
    End Try
  End Sub

  ''' <summary>
  ''' 文字列のShift-JISバイト数を検証
  ''' </summary>
  ''' <param name="row">対象DataRow</param>
  ''' <param name="columnName">対象列名</param>
  ''' <param name="maxLength">最大バイト数</param>
  ''' <param name="displayName">表示名</param>
  Private Sub ValidateTextLength(row As DataRow, columnName As String, maxLength As Integer, displayName As String)
    Dim value As String = GetDataRowString(row, columnName)
    If Encoding.GetEncoding(932).GetByteCount(value) > maxLength Then
      Throw New Exception(displayName & "は" & maxLength.ToString() & "バイト以内で入力してください。")
    End If
  End Sub

  ''' <summary>
  ''' 文字種とShift-JISバイト数を検証
  ''' </summary>
  ''' <param name="row">対象DataRow</param>
  ''' <param name="columnName">対象列名</param>
  ''' <param name="maxLength">最大バイト数</param>
  ''' <param name="pattern">入力値を判定する正規表現</param>
  ''' <param name="errorMessage">入力エラーメッセージ</param>
  Private Sub ValidatePattern(row As DataRow, columnName As String, maxLength As Integer, pattern As String, errorMessage As String)
    Dim value As String = GetDataRowString(row, columnName)
    If value = "" Then Return
    If Encoding.GetEncoding(932).GetByteCount(value) > maxLength OrElse Not Regex.IsMatch(value, pattern) Then
      Throw New Exception(errorMessage)
    End If
  End Sub

  ''' <summary>
  ''' DataRowから文字列値を取得
  ''' </summary>
  ''' <param name="row">対象DataRow</param>
  ''' <param name="columnName">対象列名</param>
  ''' <returns>指定列の文字列値</returns>
  Private Function GetDataRowString(row As DataRow, columnName As String) As String
    If row Is Nothing OrElse Not row.Table.Columns.Contains(columnName) Then Return ""
    If IsDBNull(row(columnName)) OrElse row(columnName) Is Nothing Then Return ""
    Return row(columnName).ToString().Trim()
  End Function

  ''' <summary>
  ''' 既存のマスタ取込処理を呼び出す
  ''' </summary>
  Private Sub CallGetMaster()
    Dim dr As DialogResult
    Dim Concat_ScaleNumber As String = String.Empty
    Try
      dr = MessageBox.Show("マスタ取込を行います。" & vbCrLf & "よろしいでしょうか。", "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)
      If dr = DialogResult.Cancel Then
        _masterImportCompleted = False
        Exit Sub
      End If

      LblMessage.Text = "得意先マスタ取込中・・・"
      GetMaster("CustomerMasterDownload", Concat_ScaleNumber)
      LblMessage.Text = "得意先マスタ取込完了しました。 商品マスタ取込中・・・"

      GetMaster("ItemMasterDownload", Concat_ScaleNumber)
      LblMessage.Text = "商品マスタ取込完了しました。担当者取込中・・・"

      GetMaster("TantoshaMasterDownload", Concat_ScaleNumber)
      LblMessage.Text = "担当者マスタ取込完了"

      ''商品登録
      'PcaInsShohin()

      ''得意先登録
      'PcaInsTokuisaki()

      _masterImportCompleted = True
      MessageBox.Show("マスタ取込終了しました。" & vbCrLf & "処理結果をご確認下さい。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information)

    Catch ex As Exception
      ComWriteErrLog(ex, False)
      Throw
    End Try
  End Sub

  ''' <summary>
  ''' 現在選択中のタブに対応するグリッドを取得
  ''' </summary>
  ''' <returns>現在選択中のDataGridView</returns>
  Private Function GetCurrentGrid() As DataGridView
    If tabMaster.SelectedTab Is tabTokuisaki Then Return dgvTokuisaki
    If tabMaster.SelectedTab Is tabShohin Then Return dgvShohin
    Return dgvTanto
  End Function

  ''' <summary>
  ''' 現在選択中のタブに対応するDataTableを取得
  ''' </summary>
  ''' <returns>現在選択中のDataTable</returns>
  Private Function GetCurrentDataTable() As DataTable
    If tabMaster.SelectedTab Is tabTokuisaki Then Return _dtTokuisaki
    If tabMaster.SelectedTab Is tabShohin Then Return _dtShohin
    Return _dtTanto
  End Function

  ''' <summary>
  ''' ファンクションキーおよびEnterキーのキーボード操作を処理
  ''' </summary>
  ''' <param name="sender">イベント発生元</param>
  ''' <param name="e">イベント引数</param>
  Private Sub FrmKeiryokiMasterOutput_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
    Try
      If e.KeyCode = Keys.Enter Then
        If GetCurrentGrid().Focused OrElse GetCurrentGrid().IsCurrentCellInEditMode Then Return
        e.SuppressKeyPress = True

        If tabMaster.SelectedTab Is tabTokuisaki Then
          SearchTokuisaki()
        ElseIf tabMaster.SelectedTab Is tabShohin Then
          SearchShohin()
        Else
          SearchTanto()
        End If
      Else
        Select Case e.KeyCode
          Case Keys.F1
            e.SuppressKeyPress = True
            btnAdd.PerformClick()
          Case Keys.F3
            e.SuppressKeyPress = True
            btnDelete.PerformClick()
          Case Keys.F5
            e.SuppressKeyPress = True
            btnOutput.PerformClick()
          Case Keys.F8
            e.SuppressKeyPress = True
            btnBarcodePrint.PerformClick()
          Case Keys.F10
            e.SuppressKeyPress = True
            btnClose.PerformClick()
        End Select
      End If

    Catch ex As Exception
      ComWriteErrLog(ex, False)
    End Try
  End Sub

  ''' <summary>
  ''' 仕入先検索条件変更時の処理
  ''' </summary>
  ''' <param name="sender">イベント発生元</param>
  ''' <param name="e">イベント引数</param>
  Private Sub txtTokuisaki_TextChanged(sender As Object, e As EventArgs) Handles txtTokuisakiCd.TextChanged, txtTokuisakiNm.TextChanged
    ' 検索はEnterで実行
  End Sub

  ''' <summary>
  ''' 商品検索条件変更時の処理
  ''' </summary>
  ''' <param name="sender">イベント発生元</param>
  ''' <param name="e">イベント引数</param>
  Private Sub txtShohin_TextChanged(sender As Object, e As EventArgs) Handles txtShohinCd.TextChanged, txtShohinNm.TextChanged
    ' 検索はEnterで実行
  End Sub

  ''' <summary>
  ''' 担当者検索条件変更時の処理
  ''' </summary>
  ''' <param name="sender">イベント発生元</param>
  ''' <param name="e">イベント引数</param>
  Private Sub txtTanto_TextChanged(sender As Object, e As EventArgs) Handles txtTantoCd.TextChanged, txtTantoNm.TextChanged
    ' 検索はEnterで実行
  End Sub

  ''' <summary>
  ''' 検索用コード欄の数値入力制御
  ''' </summary>
  Private Sub TxtNumberOnly_KeyPress(sender As Object, e As KeyPressEventArgs) _
    Handles txtTokuisakiCd.KeyPress, txtShohinCd.KeyPress, txtTantoCd.KeyPress

    ' バックスペースは許可
    If e.KeyChar = ControlChars.Back Then
      Return
    End If

    ' 半角数字だけ許可（"0"～"9"）
    If e.KeyChar < "0"c OrElse e.KeyChar > "9"c Then
      e.Handled = True
    End If
  End Sub

  ''' <summary>
  ''' マスタバーコード印刷処理
  ''' </summary>
  ''' <param name="sender">イベント発生元</param>
  ''' <param name="e">イベント引数</param>
  Private Sub btnBarcodePrint_Click(sender As Object, e As EventArgs) Handles btnBarcodePrint.Click
    Dim ReportWkTable As String = "WK_MASTER"
    Dim ReportName As String = "R_MASTER"
    Dim ClsPrintingProcess As New ClsPrintingProcess.ClsPrintingProcess()

    Try

      ClsPrintingProcess.PrintProcess(ClsCommonGlobalData.PRINT_PREVIEW, ReportWkTable, ReportName)
    Catch ex As Exception
      ComWriteErrLog(ex, False)
    End Try
  End Sub

  ''' <summary>
  ''' 画面を終了
  ''' </summary>
  ''' <param name="sender">イベント発生元</param>
  ''' <param name="e">イベント引数</param>
  Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
    Me.Close()
  End Sub

  ''' <summary>
  ''' 画面終了時の後処理
  ''' </summary>
  ''' <param name="sender">イベント発生元</param>
  ''' <param name="e">イベント引数</param>
  Private Sub FrmKeiryokiMasterOutput_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
    Try
      _sqlServer.Dispose()
    Catch ex As Exception
      ComWriteErrLog(ex, False)
    End Try
  End Sub

End Class
