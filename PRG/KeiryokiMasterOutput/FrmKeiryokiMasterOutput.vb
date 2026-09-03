Imports System.Data
Imports System.Windows.Forms
Imports Common
Imports Common.ClsFunction
Imports ClsAutoCommunication.ClsAutoCommunication

Public Class FrmKeiryokiMasterOutput

  Private _service As KeiryokiMasterOutputService
  Private _dtTokuisaki As DataTable
  Private _dtShohin As DataTable
  Private _dtTanto As DataTable
  Private _sqlServer As New clsSqlServer
  Private _masterImportCompleted As Boolean = False

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

  Private Sub InitializeForm()
    Me.Text = "計量器マスタ出力"
    Me.StartPosition = FormStartPosition.CenterScreen
    Me.KeyPreview = True
  End Sub

  Private Sub InitializeGrid()
    InitializeTokuisakiGrid()
    InitializeShohinGrid()
    InitializeTantoGrid()
  End Sub

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

  Private Sub InitializeTantoGrid()
    dgvTanto.AutoGenerateColumns = False
    dgvTanto.Columns.Clear()
    SetupGridCommon(dgvTanto)

    AddTextColumn(dgvTanto, "CODE", "担当者コード", 180)
    AddTextColumn(dgvTanto, "NAME", "担当者名", 280)
  End Sub

  Private Sub SetupGridCommon(grid As DataGridView)
    grid.DefaultCellStyle.Font = New Font("Segoe UI", 15)
    grid.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 15)
    grid.ColumnHeadersHeight = 50
    grid.RowTemplate.Height = 50
    grid.AllowUserToAddRows = False
    grid.AllowUserToDeleteRows = False
    grid.RowHeadersVisible = False
    grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
    grid.MultiSelect = False
    grid.EditMode = DataGridViewEditMode.EditOnEnter
  End Sub

  Private Sub AddTextColumn(grid As DataGridView, dataPropertyName As String, headerText As String, width As Integer)
    Dim col As New DataGridViewTextBoxColumn()
    col.Name = dataPropertyName
    col.HeaderText = headerText
    col.DataPropertyName = dataPropertyName
    col.ReadOnly = False
    col.Width = width
    grid.Columns.Add(col)
  End Sub

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

  Private Sub SearchTokuisaki()
    dgvTokuisaki.EndEdit()
    ApplyFilter(_dtTokuisaki, txtTokuisakiCd.Text.Trim(), txtTokuisakiNm.Text.Trim(), "TokuiCD", "TokuiNM1", dgvTokuisaki)
  End Sub

  Private Sub SearchShohin()
    dgvShohin.EndEdit()
    ApplyFilter(_dtShohin, txtShohinCd.Text.Trim(), txtShohinNm.Text.Trim(), "ShohinCD", "ShohinNM", dgvShohin)
  End Sub

  Private Sub SearchTanto()
    dgvTanto.EndEdit()
    ApplyFilter(_dtTanto, txtTantoCd.Text.Trim(), txtTantoNm.Text.Trim(), "CODE", "NAME", dgvTanto)
  End Sub

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

  Private Function EscapeFilter(value As String) As String
    If value Is Nothing Then Return ""
    Return value.Replace("'", "''").Replace("[", "[[").Replace("]", "]]" )
  End Function

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

  Private Sub btnOutput_Click(sender As Object, e As EventArgs) Handles btnOutput.Click
    Try
      dgvTokuisaki.EndEdit()
      dgvShohin.EndEdit()
      dgvTanto.EndEdit()

      If MessageBox.Show("計量器マスタCSVを出力します。" & vbCrLf & "よろしいでしょうか。", _
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
      _service.UpdatePcFields(_dtTokuisaki, _dtShohin, _sqlServer)

      MessageBox.Show("CSV出力およびマスタ取込が完了しました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information)

    Catch ex As Exception
      ComWriteErrLog(ex, False)
      MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Try
  End Sub

  Private Sub ValidateCodes()
    ValidateCodeTable(_dtTokuisaki, "TokuiCD", "仕入先")
    ValidateCodeTable(_dtShohin, "ShohinCD", "商品")
    ValidateCodeTable(_dtTanto, "CODE", "担当者")
  End Sub

  Private Sub ValidateCodeTable(dt As DataTable, columnName As String, targetName As String)
    If dt Is Nothing Then Return

    For Each row As DataRow In dt.Rows
      If row.RowState = DataRowState.Deleted Then Continue For
      If String.IsNullOrWhiteSpace(If(row(columnName), "").ToString()) Then
        Throw New Exception(targetName & "コードが未入力の行があります。")
      End If
    Next
  End Sub

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
      LblMessage.Text = "商品マスタ取込完了"

      'GetMaster("ShippingMasterDownload", Concat_ScaleNumber)
      'LblMessage.Text = "発送先マスタ取込完了"

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

  Private Function GetCurrentGrid() As DataGridView
    If tabMaster.SelectedTab Is tabTokuisaki Then Return dgvTokuisaki
    If tabMaster.SelectedTab Is tabShohin Then Return dgvShohin
    Return dgvTanto
  End Function

  Private Function GetCurrentDataTable() As DataTable
    If tabMaster.SelectedTab Is tabTokuisaki Then Return _dtTokuisaki
    If tabMaster.SelectedTab Is tabShohin Then Return _dtShohin
    Return _dtTanto
  End Function

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
          Case Keys.F10
            e.SuppressKeyPress = True
            btnClose.PerformClick()
        End Select
      End If

    Catch ex As Exception
      ComWriteErrLog(ex, False)
    End Try
  End Sub

  Private Sub txtTokuisaki_TextChanged(sender As Object, e As EventArgs) Handles txtTokuisakiCd.TextChanged, txtTokuisakiNm.TextChanged
    ' 検索はEnterで実行
  End Sub

  Private Sub txtShohin_TextChanged(sender As Object, e As EventArgs) Handles txtShohinCd.TextChanged, txtShohinNm.TextChanged
    ' 検索はEnterで実行
  End Sub

  Private Sub txtTanto_TextChanged(sender As Object, e As EventArgs) Handles txtTantoCd.TextChanged, txtTantoNm.TextChanged
    ' 検索はEnterで実行
  End Sub

  Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
    Me.Close()
  End Sub

  Private Sub FrmKeiryokiMasterOutput_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
    Try
      _sqlServer.Dispose()
    Catch ex As Exception
      ComWriteErrLog(ex, False)
    End Try
  End Sub

End Class
