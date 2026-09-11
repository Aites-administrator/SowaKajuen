Imports System.Data
Imports System.IO
Imports System.Text
Imports Common
Imports Common.ClsFunction
Imports ClsAutoCommunication.ClsAutoCommunication

Public Class KeiryokiMasterOutputService

  Private Const SupplierFileName As String = "40FRE1"
  Private Const ItemFileName As String = "40MAS1"
  Private Const TantoFileName As String = "40OPTR1.CSV"
  Private Const SCALE_DIR As String = "SCALE"

  Private Const ItemHeader As String = "呼出コード,品番,風袋,風袋単位,上限値,上限値単位,基準値,基準値単位,下限値,下限値単位,設定値１,設定値１単位,設定値２,設定値２単位,設定値３,設定値３単位,ロケーション,サブコード１,サブコード２,在庫重量,在庫重量単位,補充発注重量,補充発注重量単位,小計目標値,小計目標値単位,小計目標回数,品名,加工日印字,加工時刻印字,加工時刻フラグ,加工時刻,有効日印字,有効時刻印字,有効時刻フラグ,有効日期間,有効時刻,加算フォーマット№,小計フォーマット№,風袋１№,風袋２№,風袋２の掛け算,読取りバー№,フリー１№,フリー２№,フリー３№,フリー４№,フリー５№,秤№,加算ラベル枚数,小計ラベル枚数,イメージ№,文字列2,文字列3,文字列4,文字列5,文字列6,文字列7,文字列8,文字列9,文字列10,風袋上限値,風袋上限値単位,風袋下限値,風袋下限値単位,風袋上下限値No.,内容量,内容量単位"

  ' 添付いただいた計量器CSVの1データ行を基本値として使用する。
  ' 呼出コード(列2)と品名(列64)だけを画面データで置換する。
  Private Const ItemDefaultRow As String = "{0},,0,g,0,g,0,g,0,g,0,g,0,g,0,g,0,,,0,g,0,g,0,g,0,""{1}"",1,2,2,0,1,2,2,1,0,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,,,,,,,,,,0,g,0,g,0,0,g"

  ''' <summary>
  ''' 仕入先マスタデータを取得
  ''' </summary>
  ''' <param name="prmSqlServer">SQL Server接続オブジェクト</param>
  ''' <returns>仕入先マスタのDataTable</returns>
  Public Function GetTokuisakiList(prmSqlServer As ClsSqlServer) As DataTable
    Dim raw As New DataTable()
    Dim sql As String = BuildTokuisakiSelectSql()
    Dim rtn As New DataTable
    Try
      prmSqlServer.GetResult(raw, sql)
      rtn = CreateStringTable(raw, {"TokuiCD", "TokuiNM1", "Add1", "Add2", "ZipCD", "TelNo", "FaxNo", "TokuiInvoiceNumber"})

    Catch ex As Exception
      ComWriteErrLog(ex)
    End Try
    Return rtn
  End Function

  ''' <summary>
  ''' 商品マスタデータを取得
  ''' </summary>
  ''' <param name="prmSqlServer">SQL Server接続オブジェクト</param>
  ''' <returns>商品マスタのDataTable</returns>
  Public Function GetShohinList(prmSqlServer As ClsSqlServer) As DataTable
    Dim raw As New DataTable()
    Dim sql As String = BuildShohinSelectSql()
    Dim rtn As New DataTable
    Try
      prmSqlServer.GetResult(raw, sql)
      rtn = CreateStringTable(raw, {"ShohinCD", "ShohinNM", "Tani", "ZeikomiKBN", "HyojunKakaku"})
    Catch ex As Exception
      ComWriteErrLog(ex)
    End Try
    Return rtn

  End Function

  ''' <summary>
  ''' 担当者マスタデータを取得
  ''' </summary>
  ''' <param name="prmSqlServer">SQL Server接続オブジェクト</param>
  ''' <returns>担当者マスタのDataTable</returns>
  Public Function GetTantoList(prmSqlServer As ClsSqlServer) As DataTable
    Dim raw As New DataTable()
    Dim sql As String = BuildTantoSelectSql()
    Dim rtn As New DataTable
    Try
      prmSqlServer.GetResult(raw, sql)
      rtn = CreateStringTable(raw, {"CODE", "NAME"})

    Catch ex As Exception
      ComWriteErrLog(ex)
    End Try
    Return rtn

  End Function

  ''' <summary>
  ''' 取得したDataTableから指定列の文字列DataTableを作成
  ''' </summary>
  ''' <param name="raw">元となるDataTable</param>
  ''' <returns>指定列で作成した文字列型のDataTable</returns>
  Private Function CreateStringTable(raw As DataTable, columnNames() As String) As DataTable
    Dim result As New DataTable()
    Try
      For Each name As String In columnNames
        result.Columns.Add(name, GetType(String))
      Next

      If raw Is Nothing Then Return result

      For Each sourceRow As DataRow In raw.Rows
        Dim row As DataRow = result.NewRow()
        For Each name As String In columnNames
          row(name) = If(sourceRow.Table.Columns.Contains(name) AndAlso Not IsDBNull(sourceRow(name)), sourceRow(name).ToString(), "")
        Next
        result.Rows.Add(row)
      Next

    Catch ex As Exception
      ComWriteErrLog(ex)
    End Try
    Return result
  End Function

  ''' <summary>
  ''' 計量器マスタデータを取得
  ''' </summary>
  ''' <param name="prmSqlServer">SQL Server接続オブジェクト</param>
  ''' <param name="Para_ScaleNumber">取得対象の計量器番号</param>
  ''' <returns>計量器マスタのDataTable</returns>
  Public Function GetScaleList(prmSqlServer As ClsSqlServer, Para_ScaleNumber As String) As DataTable
    Dim dt As New DataTable()
    Dim sql As String = GetMstScaleSelectSql(Para_ScaleNumber)
    Try
      prmSqlServer.GetResult(dt, sql)

    Catch ex As Exception
      ComWriteErrLog(ex)
    End Try
    Return dt
  End Function

  ''' <summary>
  ''' 全計量器向けのマスタCSVを出力
  ''' </summary>
  ''' <param name="dtTokuisaki">仕入先マスタDataTable</param>
  ''' <param name="dtShohin">商品マスタDataTable</param>
  ''' <param name="dtTanto">担当者マスタDataTable</param>
  ''' <param name="dtScale">計量器マスタDataTable</param>
  ''' <param name="basePath">CSV出力先の基準フォルダ</param>
  ''' <param name="prmSqlserver">SQL Server接続オブジェクト</param>
  Public Sub OutputAllCsv(dtTokuisaki As DataTable,
                           dtShohin As DataTable,
                           dtTanto As DataTable,
                           dtScale As DataTable,
                           basePath As String,
                          prmSqlserver As ClsSqlServer)
    Dim UnitNumberArray() As String = Nothing

    Try
      If String.IsNullOrWhiteSpace(basePath) Then Throw New ArgumentException("出力先が未設定です。")

      Dim supplierCsv As String = BuildSupplierCsv(dtTokuisaki)
      Dim itemCsv As String = BuildItemCsv(dtShohin)
      Dim tantoCsv As String = BuildTantoCsv(dtTanto)

      For Each scaleRow As DataRow In dtScale.Rows
        If scaleRow.RowState = DataRowState.Deleted Then Continue For

        Dim ipAddress As String = If(scaleRow.Table.Columns.Contains("ip_address") AndAlso Not IsDBNull(scaleRow("ip_address")), scaleRow("ip_address").ToString().Trim(), "")
        If ipAddress = "" Then Continue For

        Dim outputDir As String = Path.Combine(basePath, ipAddress)
        Dim outputScaleDir As String = Path.Combine(Path.Combine(basePath, ipAddress), SCALE_DIR)
        Directory.CreateDirectory(outputDir)
        Directory.CreateDirectory(outputScaleDir)

        WriteShiftJis(Path.Combine(outputDir, CreateDownloadFileName(SupplierFileName, ipAddress) & ".CSV"), supplierCsv)
        WriteShiftJis(Path.Combine(outputDir, CreateDownloadFileName(ItemFileName, ipAddress) & ".CSV"), itemCsv)

        WriteShiftJis(Path.Combine(outputScaleDir, CreateDownloadFileName(SupplierFileName, ipAddress) & ".CSV"), supplierCsv)
        WriteShiftJis(Path.Combine(outputScaleDir, CreateDownloadFileName(ItemFileName, ipAddress) & ".CSV"), itemCsv)

        WriteShiftJis(Path.Combine(outputDir, TantoFileName), tantoCsv)
        WriteShiftJis(Path.Combine(outputScaleDir, TantoFileName), tantoCsv)


      Next
    Catch ex As Exception
      ComWriteErrLog(ex)
    End Try

  End Sub

  ''' <summary>
  ''' PC側マスタの項目を更新
  ''' </summary>
  ''' <param name="dtTokuisaki">仕入先マスタDataTable</param>
  ''' <param name="dtShohin">商品マスタDataTable</param>
  ''' <param name="dtCode">担当者マスタDataTable</param>
  ''' <param name="prmSqlServer">SQL Server接続オブジェクト</param>
  Public Sub UpdatePcFields(dtTokuisaki As DataTable,
                            dtShohin As DataTable,
                            dtCode As DataTable,
                            prmSqlServer As ClsSqlServer)
    UpdateTokuisakiPcFields(dtTokuisaki, prmSqlServer)
    UpdateShohinPcFields(dtShohin, prmSqlServer)
    UpdateTantoshaPcFields(dtCode, prmSqlServer)
    UpdateTokuisakiShohinBaika(dtShohin, prmSqlServer)
  End Sub

  ''' <summary>
  ''' 仕入先マスタのPC側項目を更新
  ''' </summary>
  ''' <param name="dt">対象DataTable</param>
  ''' <param name="prmSqlServer">SQL Server接続オブジェクト</param>
  Private Sub UpdateTokuisakiPcFields(dt As DataTable, prmSqlServer As ClsSqlServer)
    Try
      If dt Is Nothing Then Return

      For Each row As DataRow In dt.Rows
        If row.RowState = DataRowState.Deleted Then Continue For

        Dim cd As String = GetString(row, "TokuiCD")
        If cd = "" Then Continue For

        Dim sql As New StringBuilder()
        sql.AppendLine("UPDATE MST_TOKUISAKI")
        sql.AppendLine("SET")
        sql.AppendLine("    TokuiNM1 = " & SqlValue(GetString(row, "TokuiNM1")) & ",")
        sql.AppendLine("    Add1 = " & SqlValue(GetString(row, "Add1")) & ",")
        sql.AppendLine("    Add2 = " & SqlValue(GetString(row, "Add2")) & ",")
        sql.AppendLine("    ZipCD = " & SqlValue(GetString(row, "ZipCD")) & ",")
        sql.AppendLine("    TelNo = " & SqlValue(GetString(row, "TelNo")) & ",")
        sql.AppendLine("    FaxNo = " & SqlValue(GetString(row, "FaxNo")) & ",")
        sql.AppendLine("    TokuiInvoiceNumber = " & SqlValue(GetString(row, "TokuiInvoiceNumber")))
        sql.AppendLine("WHERE TokuiCD = " & SqlValue(cd))

        prmSqlServer.Execute(sql.ToString())
      Next
    Catch ex As Exception
      ComWriteErrLog(ex)
    End Try

  End Sub

  ''' <summary>
  ''' 商品マスタのPC側項目を更新
  ''' </summary>
  ''' <param name="dt">対象DataTable</param>
  ''' <param name="prmSqlServer">SQL Server接続オブジェクト</param>
  Private Sub UpdateShohinPcFields(dt As DataTable, prmSqlServer As ClsSqlServer)
    Try
      If dt Is Nothing Then Return

      For Each row As DataRow In dt.Rows
        If row.RowState = DataRowState.Deleted Then Continue For

        Dim cd As String = GetString(row, "ShohinCD")
        If cd = "" Then Continue For

        Dim sql As New StringBuilder()
        sql.AppendLine("UPDATE MST_SHOHIN")
        sql.AppendLine("SET")
        sql.AppendLine("    ShohinNM = " & SqlValue(GetString(row, "ShohinNM")) & ",")
        sql.AppendLine("    Tani = " & SqlValue(GetString(row, "Tani")) & ",")
        sql.AppendLine("    ZeikomiKBN = " & NumericSqlValue(GetString(row, "ZeikomiKBN")) & ",")
        sql.AppendLine("    HyojunKakaku = " & NumericSqlValue(GetString(row, "HyojunKakaku")))
        sql.AppendLine("WHERE ShohinCD = " & SqlValue(cd))

        prmSqlServer.Execute(sql.ToString())
      Next
    Catch ex As Exception
      ComWriteErrLog(ex)
    End Try

  End Sub

  ''' <summary>
  ''' 得意先商品マスタの売価を更新
  ''' </summary>
  ''' <param name="dt">対象DataTable</param>
  ''' <param name="prmSqlServer">SQL Server接続オブジェクト</param>
  Private Sub UpdateTokuisakiShohinBaika(dt As DataTable, prmSqlServer As ClsSqlServer)
    Try
      If dt Is Nothing Then Return

      For Each row As DataRow In dt.Rows
        If row.RowState = DataRowState.Deleted Then Continue For

        Dim shohinCd As String = GetString(row, "ShohinCD")
        If shohinCd = "" Then Continue For

        Dim baika As String = GetString(row, "HyojunKakaku")

        Dim sql As New StringBuilder()
        sql.AppendLine("UPDATE MST_TOKUISAKI_SHOHIN")
        sql.AppendLine("SET")
        sql.AppendLine("    Baika = " & NumericSqlValue(baika))
        sql.AppendLine("WHERE TokuiCD = '000000'")
        sql.AppendLine("  AND ShohinCD = " & SqlValue(shohinCd))

        prmSqlServer.Execute(sql.ToString())
      Next
    Catch ex As Exception
      ComWriteErrLog(ex)
    End Try
  End Sub

  ''' <summary>
  ''' 担当者マスタのPC側項目を更新
  ''' </summary>
  ''' <param name="dt">対象DataTable</param>
  ''' <param name="prmSqlServer">SQL Server接続オブジェクト</param>
  Private Sub UpdateTantoshaPcFields(dt As DataTable, prmSqlServer As ClsSqlServer)
    Try
      If dt Is Nothing Then Return

      For Each row As DataRow In dt.Rows
        If row.RowState = DataRowState.Deleted Then Continue For

        Dim cd As String = GetString(row, "CODE")
        If cd = "" Then Continue For

        Dim sql As New StringBuilder()
        sql.AppendLine("UPDATE MST_TANTO")
        sql.AppendLine("SET")
        sql.AppendLine("    NAME = " & SqlValue(GetString(row, "NAME")))
        sql.AppendLine("WHERE CODE = " & SqlValue(cd))

        prmSqlServer.Execute(sql.ToString())
      Next
    Catch ex As Exception
      ComWriteErrLog(ex)
    End Try

  End Sub

  ''' <summary>
  ''' 仕入先マスタCSV文字列を作成
  ''' </summary>
  ''' <param name="dt">対象DataTable</param>
  ''' <returns>仕入先マスタのCSV文字列</returns>
  Private Function BuildSupplierCsv(dt As DataTable) As String
    Dim sb As New StringBuilder()
    Dim rtn As String = String.Empty

    Try
      sb.AppendLine("フリー１№,フリー１名称")

      For Each row As DataRow In dt.Rows
        If row.RowState = DataRowState.Deleted Then Continue For
        Dim cd As String = GetString(row, "TokuiCD")
        Dim nm As String = GetString(row, "TokuiNM1")
        sb.AppendLine(CsvValue(cd) & "," & CsvValue("<NS/>" & nm))
      Next

      rtn = sb.ToString().TrimEnd(ControlChars.Cr, ControlChars.Lf) & ControlChars.CrLf

    Catch ex As Exception
      ComWriteErrLog(ex)
    End Try
    Return rtn
  End Function

  ''' <summary>
  ''' 担当者マスタCSV文字列を作成
  ''' </summary>
  ''' <param name="dt">対象DataTable</param>
  ''' <returns>担当者マスタのCSV文字列</returns>
  Private Function BuildTantoCsv(dt As DataTable) As String
    Dim sb As New StringBuilder()
    Dim rtn As String = String.Empty

    Try
      sb.AppendLine("担当者№,担当者名称")

      For Each row As DataRow In dt.Rows
        If row.RowState = DataRowState.Deleted Then Continue For
        Dim cd As String = GetString(row, "CODE")
        Dim nm As String = GetString(row, "NAME")
        sb.AppendLine(CsvValue(cd) & "," & CsvValue("<NS/>" & nm))
      Next

      rtn = sb.ToString().TrimEnd(ControlChars.Cr, ControlChars.Lf) & ControlChars.CrLf


    Catch ex As Exception
      ComWriteErrLog(ex)
    End Try
    Return rtn
  End Function

  ''' <summary>
  ''' 商品マスタCSV文字列を作成
  ''' </summary>
  ''' <param name="dt">対象DataTable</param>
  ''' <returns>商品マスタのCSV文字列</returns>
  Private Function BuildItemCsv(dt As DataTable) As String
    Dim sb As New StringBuilder()
    Dim rtn As String = String.Empty

    Try
      sb.AppendLine(ItemHeader)

      For Each row As DataRow In dt.Rows
        If row.RowState = DataRowState.Deleted Then Continue For

        Dim cd As String = GetString(row, "ShohinCD")
        Dim nm As String = GetString(row, "ShohinNM")

        ' 添付サンプルと同じ85列構成を維持し、
        ' 呼出コードと品名だけ画面値に置換する。
        sb.AppendLine(String.Format(ItemDefaultRow, CsvRaw(cd), CsvRaw("<NL/>" & nm)))
      Next

      rtn = sb.ToString().TrimEnd(ControlChars.Cr, ControlChars.Lf) & ControlChars.CrLf
    Catch ex As Exception
      ComWriteErrLog(ex)
    End Try
    Return rtn

  End Function

  ''' <summary>
  ''' CSV出力用の文字列を整形
  ''' </summary>
  ''' <param name="value">入力値</param>
  ''' <returns>CSV出力用に整形した文字列</returns>
  Private Function CsvRaw(value As String) As String
    If value Is Nothing Then value = ""
    Return value.Replace("""", """")
  End Function

  ''' <summary>
  ''' CSV項目をダブルクォートで囲んで整形
  ''' </summary>
  ''' <param name="value">入力値</param>
  ''' <returns>CSV項目として整形した文字列</returns>
  Private Function CsvValue(value As String) As String
    If value Is Nothing Then value = ""
    Return """" & value.Replace("""", """") & """"
  End Function

  ''' <summary>
  ''' Shift-JIS形式でファイルを出力
  ''' </summary>
  ''' <param name="path">出力ファイルパス</param>
  ''' <param name="content">出力内容</param>
  Private Sub WriteShiftJis(path As String, content As String)
    File.WriteAllText(path, content, Encoding.GetEncoding(932))
  End Sub

  ''' <summary>
  ''' DataRowから文字列値を取得
  ''' </summary>
  ''' <param name="row">対象DataRow</param>
  ''' <param name="columnName">対象列名</param>
  ''' <returns>指定列の文字列値</returns>
  Private Function GetString(row As DataRow, columnName As String) As String
    If row Is Nothing OrElse Not row.Table.Columns.Contains(columnName) Then Return ""
    If IsDBNull(row(columnName)) OrElse row(columnName) Is Nothing Then Return ""
    Return row(columnName).ToString().Trim()
  End Function

  ''' <summary>
  ''' SQL文字列値を生成
  ''' </summary>
  ''' <param name="value">入力値</param>
  ''' <returns>SQLの文字列値</returns>
  Private Function SqlValue(value As String) As String
    If value Is Nothing Then Return "NULL"
    Return "'" & value.Replace("'", "''") & "'"
  End Function

  ''' <summary>
  ''' SQL数値値を生成
  ''' </summary>
  ''' <param name="value">入力値</param>
  ''' <returns>SQLの数値値</returns>
  Private Function NumericSqlValue(value As String) As String
    If String.IsNullOrWhiteSpace(value) Then Return "NULL"
    Dim number As Decimal
    If Decimal.TryParse(value, Globalization.NumberStyles.Number, Globalization.CultureInfo.InvariantCulture, number) Then
      Return number.ToString(Globalization.CultureInfo.InvariantCulture)
    End If
    Return "NULL"
  End Function

  ''' <summary>
  ''' 仕入先マスタ取得SQLを生成
  ''' </summary>
  ''' <returns>仕入先マスタ取得用SQL</returns>
  Private Function BuildTokuisakiSelectSql() As String
    Dim sb As New StringBuilder()
    sb.AppendLine("SELECT")
    sb.AppendLine("    TokuiCD,")
    sb.AppendLine("    TokuiNM1,")
    sb.AppendLine("    Add1,")
    sb.AppendLine("    Add2,")
    sb.AppendLine("    ZipCD,")
    sb.AppendLine("    TelNo,")
    sb.AppendLine("    FaxNo,")
    sb.AppendLine("    TokuiInvoiceNumber")
    sb.AppendLine("FROM MST_TOKUISAKI")
    sb.AppendLine("WHERE TokuiCd <> 0")
    sb.AppendLine("ORDER BY TokuiCD")
    Return sb.ToString()
  End Function

  ''' <summary>
  ''' 商品マスタ取得SQLを生成
  ''' </summary>
  ''' <returns>商品マスタ取得用SQL</returns>
  Private Function BuildShohinSelectSql() As String
    Dim sb As New StringBuilder()
    sb.AppendLine("SELECT")
    sb.AppendLine("    ShohinCD,")
    sb.AppendLine("    ShohinNM,")
    sb.AppendLine("    Tani,")
    sb.AppendLine("    ZeikomiKBN,")
    sb.AppendLine("    HyojunKakaku")
    sb.AppendLine("FROM MST_SHOHIN")
    sb.AppendLine("ORDER BY ShohinCD")
    Return sb.ToString()
  End Function

  ''' <summary>
  ''' 担当者マスタ取得SQLを生成
  ''' </summary>
  ''' <returns>担当者マスタ取得用SQL</returns>
  Private Function BuildTantoSelectSql() As String
    Dim sb As New StringBuilder()
    sb.AppendLine("SELECT")
    sb.AppendLine("    CODE,")
    sb.AppendLine("    NAME")
    sb.AppendLine("FROM MST_TANTO")
    sb.AppendLine("ORDER BY CODE")
    Return sb.ToString()
  End Function

  ''' <summary>
  ''' 計量器マスタ取得SQLを生成
  ''' </summary>
  ''' <param name="Para_ScaleNumber">取得対象の計量器番号</param>
  ''' <returns>計量器マスタ取得用SQL</returns>
  Private Function GetMstScaleSelectSql(Para_ScaleNumber As String) As String
    Dim sql As String = String.Empty
    sql &= " SELECT"
    sql &= "     unit_number,"
    sql &= "     ip_address"
    sql &= " FROM"
    sql &= "     MST_Scale"
    sql &= " WHERE"
    sql &= "     delete_flg = 0"
    If Para_ScaleNumber.Length <> 0 Then
      sql &= "     AND unit_number IN(" & Para_ScaleNumber & ")"
    End If
    sql &= " ORDER BY  "
    sql &= "     unit_number"
    Call WriteExecuteLog("Module_Upload", System.Reflection.MethodBase.GetCurrentMethod().Name, sql)
    Return sql
  End Function

End Class
