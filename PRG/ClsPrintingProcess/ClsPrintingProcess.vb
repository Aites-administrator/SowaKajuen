Imports Common
Imports Common.ClsFunction
Imports Common.ClsCommonGlobalData

Public Class ClsPrintingProcess
  ' ワークテーブル名

  ' １つ目のプロセスＩＤ
  Private Shared ryoProcesID_01 As System.Diagnostics.Process
  ' プロセスＩＤ
  Private Shared procesID As System.Diagnostics.Process
  ' 納品書用ワークテーブル
  Private tmpNouhinshoDT As New DataTable
  ' SQLサーバー操作オブジェクト
  Private _SqlServer As ClsSqlServer

  Private Const MAX_PRINT_COUNT As Integer = 6

  Private Const SOKUJI As Integer = 1      '即時発行
  Private Const TSUJYO As Integer = 0      '通常発行
  Public Sub New()
  End Sub


  Private ReadOnly Property SqlServer As ClsSqlServer
    Get
      If _SqlServer Is Nothing Then
        _SqlServer = New ClsSqlServer
      End If
      Return _SqlServer
    End Get
  End Property

  ''' <summary>
  ''' ACCESSファイルを開く
  ''' </summary>
  ''' <param name="printPreview">プレビューフラグ</param>
  ''' <param name="strReportName">レポートファイル名</param>
  ''' <param name="prmWaitFlag">待機フラグ</param>
  ''' <returns>
  ''' True :ファイルオープン成功
  ''' False:ファイルオープン失敗
  ''' </returns>
  Public Shared Function AccessRun(printPreview As Integer, strReportName As String, Optional prmWaitFlag As Boolean = False) As Boolean
    Try

      ' Threadオブジェクトを作成する
      Dim MultiProgram_run = New System.Threading.Thread(AddressOf DoRyoSomething01)
      ' １つ目のスレッドを開始する
      MultiProgram_run.Start(New prmReport(printPreview.ToString, strReportName))

      If (prmWaitFlag) Then
        MultiProgram_run.Join()
      End If

    Catch ex As Exception
      Call ComWriteErrLog(ex)
      Return False
    End Try

    Return True

  End Function

  ''' <summary>
  ''' １つ目の印刷スレッド
  ''' </summary>
  ''' <param name="arg"></param>
  Private Shared Sub DoRyoSomething01(arg As Object)

    Try
      Dim prm As prmReport = DirectCast(arg, prmReport)
      Dim myPath As String = System.IO.Path.Combine(My.Application.Info.DirectoryPath, ClsCommonGlobalData.REPORT_FILENAME)

      Dim strPrintPrwview As String
      If (prm.printPreview.Equals("1")) Then
        strPrintPrwview = "1"
      Else
        strPrintPrwview = "0"
      End If

      'ファイルを開く
      ryoProcesID_01 = System.Diagnostics.Process.Start(myPath, " /runtime /cmd " & strPrintPrwview & prm.strReportName)
      If ryoProcesID_01 IsNot Nothing Then
        '終了するまで待機する
        ryoProcesID_01.WaitForExit()
        ryoProcesID_01 = Nothing
      End If
    Catch ex As Exception
      Call ComWriteErrLog(ex, False)   ' Error出力（＋画面表示）
    End Try

  End Sub

  ''' <summary>
  ''' ACCESSファイルを開く
  ''' </summary>
  ''' <param name="printPreview">プレビューフラグ</param>
  ''' <param name="strReportName">レポートファイル名</param>
  ''' <returns>
  ''' True :ファイルオープン成功
  ''' False:ファイルオープン失敗
  ''' </returns>
  Public Shared Function ComAccessRun(printPreview As Integer, strReportName As String) As Boolean
    Try

      ' Threadオブジェクトを作成する
      Dim MultiProgram_run = New System.Threading.Thread(AddressOf DoSomething01)
      ' １つ目のスレッドを開始する
      MultiProgram_run.Start(New prmReport(printPreview.ToString, strReportName))

    Catch ex As Exception
      Call ComWriteErrLog(ex)
      Return False
    End Try

    Return True

  End Function


  ''' <summary>
  ''' 印刷スレッド
  ''' </summary>
  ''' <param name="arg"></param>
  Private Shared Sub DoSomething01(arg As Object)

    Dim prm As prmReport = DirectCast(arg, prmReport)
    Dim myPath As String = System.IO.Path.Combine(My.Application.Info.DirectoryPath, ClsCommonGlobalData.REPORT_FILENAME)

    Dim strPrintPrwview As String
    If (prm.printPreview.Equals("1")) Then
      strPrintPrwview = "1"
    Else
      strPrintPrwview = "0"
    End If

    'ファイルを開く
    procesID = System.Diagnostics.Process.Start(myPath, " /runtime /cmd " & strPrintPrwview & prm.strReportName)
    If procesID IsNot Nothing Then
      '終了するまで待機する
      procesID.WaitForExit()
      procesID = Nothing
    End If

  End Sub

  ''' <summary>
  ''' プロセスの終了
  ''' </summary>
  Public Shared Sub ProcessKill()

    If procesID IsNot Nothing Then
      ' 起動した１つ目のプロセスの終了
      procesID.Kill()
      procesID = Nothing
    End If

  End Sub

  ''' <summary>
  ''' プロセス状態確認
  ''' </summary>
  ''' <returns></returns>
  Public Shared Function ProcessStatus() As Boolean

    Dim ret As Boolean = False

    If procesID IsNot Nothing Then
      ret = True
    End If

    Return ret

  End Function

  Public Overloads Sub PrintProcess(prmPreview As Integer, prmTableName As String, prmReportName As String, Optional ByRef prmWhereList As Dictionary(Of String, String) = Nothing)
    Dim tmpDt As New DataTable
    Try
      '対象データ取得
      SqlServer.GetResult(tmpDt, SqlGetPrintData(prmWhereList))

      '印刷処理
      If tmpDt.Rows.Count = 0 Then
        Exit Sub
      End If
      If Not AccessPrint(prmPreview, prmTableName, prmReportName, tmpDt) Then
        Throw New Exception("印刷処理に失敗しました。")
      End If

      For Each tmpRow As DataRow In tmpDt.Rows
        SqlServer.Execute(SqlUpdPrintFlg(tmpRow, prmWhereList))
      Next
    Catch ex As Exception
      ComWriteErrLog(ex)
    End Try

  End Sub

  Public Overloads Sub PrintProcess(prmPreview As Integer, prmTableName As String, prmReportName As String)
    Dim tmpTantoDt As New DataTable
    Dim tmpTokuiDt As New DataTable
    Dim tmpShohinDt As New DataTable

    Try
      '対象データ取得
      SqlServer.GetResult(tmpTantoDt, SqlGetTantoData())
      SqlServer.GetResult(tmpTokuiDt, SqlGetTokuiData())
      SqlServer.GetResult(tmpShohinDt, SqlGetShohinData())

      '印刷処理
      If tmpTantoDt.Rows.Count = 0 Then
        Exit Sub
      End If
      If tmpTokuiDt.Rows.Count = 0 Then
        Exit Sub
      End If
      If tmpShohinDt.Rows.Count = 0 Then
        Exit Sub
      End If



      If Not AccessPrint(prmPreview, prmTableName, prmReportName, CreateMasterData(tmpTokuiDt, tmpShohinDt, tmpTantoDt)) Then
        Throw New Exception("印刷処理に失敗しました。")
      End If

    Catch ex As Exception
      ComWriteErrLog(ex)
    End Try

  End Sub

  Public Overloads Sub PrintProcess(prmPreview As Integer, prmTableName As String, prmReportName As String, Optional ByRef prmWhereList As Dictionary(Of String, List(Of String)) = Nothing)
    Dim tmpDt As New DataTable
    Try
      '対象データ取得
      SqlServer.GetResult(tmpDt, SqlGetPrintData(prmWhereList))

      '印刷処理
      If Not AccessPrint(prmPreview, prmTableName, prmReportName, tmpDt) Then
        Throw New Exception("印刷処理に失敗しました。")
      End If

      For Each tmpRow As DataRow In tmpDt.Rows
        SqlServer.Execute(SqlUpdPrintFlg(tmpRow, prmWhereList))
      Next
    Catch ex As Exception
      ComWriteErrLog(ex, False)
    End Try

  End Sub

  Private Function AccessPrint(prmPreview As Integer, prmTableName As String, prmReportName As String, tmpDt As DataTable) As Boolean
    Dim rtn As Boolean = True

    Try
      If (tmpDt.Rows.Count = 0) Then
        Throw New Exception
      End If

      'ワークテーブル作成
      UpdateReportNohinSet(tmpDt, prmTableName, True)

      '印刷処理
      AccessRun(prmPreview, prmReportName, True)

    Catch ex As Exception
      rtn = False
    End Try
    Return rtn
  End Function

  Private Function CreateMasterData(tmpTokuiDt As DataTable, tmpShohinDt As DataTable, tmpTantoDt As DataTable) As DataTable

    ' WK_MASTER の構造を作成
    Dim dt As New DataTable
    dt.Columns.Add("SIIRE_CD", GetType(String))
    dt.Columns.Add("SIIRE_NM", GetType(String))
    dt.Columns.Add("ITEM_CD", GetType(String))
    dt.Columns.Add("ITEM_NM", GetType(String))
    dt.Columns.Add("TANTO_CD", GetType(String))
    dt.Columns.Add("TANTO_NM", GetType(String))

    ' 仕入先 × 商品 × 担当者 の全組み合わせを作成
    For Each tokui As DataRow In tmpTokuiDt.Rows
      For Each shohin As DataRow In tmpShohinDt.Rows
        For Each tanto As DataRow In tmpTantoDt.Rows

          dt.Rows.Add(
                    tokui("TokuiCD"),
                    tokui("TokuiNM1"),
                    shohin("ShohinCD"),
                    shohin("ShohinNM"),
                    tanto("CODE"),
                    tanto("NAME")
                )

        Next
      Next
    Next

    Return dt
  End Function

  ''' <summary>
  ''' 量目表（セット）ワークテーブル削除と新規作成
  ''' </summary>
  ''' <returns>
  '''  True   -   成功
  '''  False  -   失敗
  ''' </returns>
  Private Function UpdateReportNohinSet(prmDt As DataTable, prmTableName As String, Optional prmMaster As Boolean = False) As Boolean

    Dim tmpDb As New ClsReport(ClsCommonGlobalData.REPORT_FILENAME)
    Dim dt As DateTime = DateTime.Parse(ComGetProcTime())
    Dim UpdEndFlg As Boolean = False

    ' 実行
    With tmpDb

      Try
        ' SQL文の作成
        .Execute("DELETE FROM " & prmTableName)

      Catch ex As Exception
        Call ComWriteErrLog(ex)
        Throw New Exception("量目表（セット）ワークテーブルの削除に失敗しました")

      End Try

      Try
        Dim sql As String
        Dim tmpDenNo As String = String.Empty

        ' トランザクション開始
        .TrnStart()

        ' データテーブルから追加SQL文を作成
        For Each row As DataRow In prmDt.Rows

          sql = SqlInsMaster(prmTableName, row)
          If String.IsNullOrWhiteSpace(sql) = False Then
            .Execute(sql)
          End If

        Next

        ' 更新成功
        .TrnCommit()

      Catch ex As Exception
        Call ComWriteErrLog(ex)
        .TrnRollBack()
        Throw New Exception("量目表（セット）ワークテーブルの書き込みに失敗しました")
      End Try

      .Dispose()

    End With


    Return True

  End Function

  ''' <summary>
  ''' 量目表（セット）ワークテーブル削除と新規作成
  ''' </summary>
  ''' <returns>
  '''  True   -   成功
  '''  False  -   失敗
  ''' </returns>
  Private Function UpdateReportNohinSet(prmDt As DataTable, prmTableName As String) As Boolean

    Dim tmpDb As New ClsReport(ClsCommonGlobalData.REPORT_FILENAME)
    Dim dt As DateTime = DateTime.Parse(ComGetProcTime())
    Dim UpdEndFlg As Boolean = False

    ' 実行
    With tmpDb

      Try
        ' SQL文の作成
        .Execute("DELETE FROM " & prmTableName)

      Catch ex As Exception
        Call ComWriteErrLog(ex)
        Throw New Exception("量目表（セット）ワークテーブルの削除に失敗しました")

      End Try

      Try
        Dim sql As String
        Dim tmpDenNo As String = String.Empty

        ' トランザクション開始
        .TrnStart()

        ' データテーブルから追加SQL文を作成
        For Each row As DataRow In prmDt.Rows
          If row("ZEIKOMI_TAX").ToString = "186" Then
            Dim a = 1
            a = 2
          End If


          'TODO 伝票ナンバーと行Noをチェック
          If tmpDenNo <> row("DenNo") Then

            tmpDenNo = row("DenNo")

            'TODO 変更があれば20でわった余り分を空行追加
            Dim rows() As DataRow = prmDt.Select("DenNo = '" & row("DenNo") & "'")
            Dim DenRowCount As Integer = rows.Length
            Dim EmptyRowCount As Integer = MAX_PRINT_COUNT - (DenRowCount Mod MAX_PRINT_COUNT)
            If DenRowCount Mod MAX_PRINT_COUNT <> 0 Then

              For i = 0 To EmptyRowCount - 1
                For j = 0 To 1
                  Dim tmpDt As DataTable = prmDt.Clone()
                  tmpDt.Rows.Clear()
                  Dim tmpRow As DataRow = tmpDt.NewRow

                  'TODO 伝票番号、行番号、得意先、発送先、ソート番号=1を入れて追加
                  tmpRow("DenNo") = row("DenNo")
                  tmpRow("GyoNo") = row("GyoNo")
                  tmpRow("TokuiCD") = row("TokuiCD")
                  tmpRow("TokuiNm") = row("TokuiNm")
                  tmpRow("TyokuCd") = row("TyokuCd")
                  tmpRow("TyokuNM") = row("TyokuNM")
                  tmpRow("SortNumber") = 1

                  If j = 0 Then
                    tmpRow("BARCODE_VALUE") = row("BARCODE_VALUE")
                  Else
                    tmpRow("BARCODE_VALUE") = ""
                  End If

                  sql = SqlInsNohin(prmTableName, tmpRow, dt)
                  If String.IsNullOrWhiteSpace(sql) = False Then
                    .Execute(sql)
                  End If

                Next
              Next
            End If

          End If

          '即時発行対象の得意先のみ単価を設定
          Dim tmpTokuiDt As New DataTable
          _SqlServer.GetResult(tmpTokuiDt, "SELECT * FROM M_TOKUISAKI_PRINT_CTRL WHERE TOKUISAKI_CD = '" & row("TokuiCD") & "' AND INSTANT_PRINT_FLG = 1 ")
          '即時発行ではない得意先のとき、単価を空にする
          If tmpTokuiDt.Rows.Count = 0 Then
            row("Tanka") = ""
          End If

          For i = 0 To 1
            If i = 0 Then
            Else
              row("BARCODE_VALUE") = ""
            End If
            sql = SqlInsNohin(prmTableName, row, dt)
            If String.IsNullOrWhiteSpace(sql) = False Then
              .Execute(sql)

            End If

          Next
        Next

        'TODO 必須やること
        Dim tmpShukeiDt As New DataTable
        Dim tmpReportDt As New DataTable
        '集計データ更新(バーコード金額、税別合計金額、税別消費税、消費税合計)
        '集計用データ取得処理
        GetShukeiData(tmpShukeiDt)
        '集計用データ作成処理(Accessのデータを取得)
        GetReportData(tmpDb, tmpReportDt)

        Dim query =
          From a In tmpShukeiDt.AsEnumerable()
          Join b In tmpReportDt.AsEnumerable()
          On a.Field(Of String)("DenNo") Equals b.Field(Of String)("DEN_NO")
          Select a

        Dim tmpReportShukeiDt As DataTable =
          If(query.Any(), query.CopyToDataTable(), tmpShukeiDt.Clone())
        '集計用データ更新処理
        For Each tmpRow In tmpReportShukeiDt.Rows
          .Execute(SqlUpdGroupData("WK_NOHIN ", tmpRow))
        Next

        ' 更新成功
        .TrnCommit()

      Catch ex As Exception
        Call ComWriteErrLog(ex)
        .TrnRollBack()
        Throw New Exception("量目表（セット）ワークテーブルの書き込みに失敗しました")
      End Try

      .Dispose()

    End With


    Return True

  End Function

  Private Sub GetShukeiData(ByRef prmDt As DataTable)

    Try
      SqlServer.GetResult(prmDt, SqlGetShukeiData())
    Catch ex As Exception
      ComWriteErrLog(ex)
    End Try

  End Sub

  Private Overloads Function SqlGetShukeiData() As String
    Dim sql As String = String.Empty

    sql &= " select DenNo "
    sql &= "	,	SUM(cast(UriageKin as int)) BARCODE_VALUE "
    sql &= "	,	SUM(cast(ISNULL(Sotozei,0) as int)) ZEINUKI_TAX "
    sql &= "	,	SUM(cast(ISNULL(Utizei,0) as int)) ZEIKOMI_TAX  "
    sql &= "	,	SUM(cast(ISNULL(Sotozei,0) as int)) + SUM(cast(ISNULL(Utizei,0) as int)) TOTAL_TAX  "
    sql &= " from TRN_JISSEKI "
    sql &= "group by DenNo "
    Return sql
  End Function

  Private Sub GetReportData(prmReport As ClsReport, ByRef prmDt As DataTable)

    Try
      prmReport.GetResult(prmDt, " SELECT * FROM WK_NOHIN ")
    Catch ex As Exception
      ComWriteErrLog(ex)
    End Try

  End Sub

  Private Overloads Function SqlGetTantoData() As String
    Dim sql As String = String.Empty

    sql &= " SELECT CODE "
    sql &= "     ,  NAME "
    sql &= " FROM MST_TANTO "
    sql &= " ORDER BY CODE "

    Return sql
  End Function

  Private Overloads Function SqlGetTokuiData() As String
    Dim sql As String = String.Empty

    sql &= " SELECT TokuiCD "
    sql &= "     ,  TokuiNM1"
    sql &= " FROM MST_TOKUISAKI "
    sql &= " ORDER BY TokuiCD "

    Return sql
  End Function

  Private Overloads Function SqlGetShohinData() As String
    Dim sql As String = String.Empty

    sql &= " SELECT ShohinCD "
    sql &= "    ,   ShohinNM "
    sql &= " FROM MST_SHOHIN "
    sql &= " ORDER BY ShohinCD "

    Return sql
  End Function


  Private Overloads Function SqlGetPrintData(Optional ByRef prmWhereList As Dictionary(Of String, String) = Nothing) As String
    Dim sql As String = String.Empty
    Dim ReportType As String = ReadSettingIniFile("REPORT_TYPE", "VALUE")

    If ReportType = SOKUJI Then
      If prmWhereList.ContainsKey("INSTANT_PRINT_FLG IS NOT ") Then
        prmWhereList.Remove("INSTANT_PRINT_FLG IS NOT ")
      End If
    End If

    sql &= " Select	trn_jisseki.NohinDay "
    sql &= "	,	Case When trn_jisseki.DenNO2 Is NULL Then trn_jisseki.DenNO Else trn_jisseki.DenNO2 + '*' END AS DenNo "
    sql &= "	,	RIGHT('00' + CAST(ISNULL(trn_jisseki.GyoNo2,trn_jisseki.GyoNo)AS VARCHAR(2)), 2)  GyoNo "
    sql &= "	,	trn_jisseki.TokuiCD "
    sql &= "	,	trn_jisseki.TokuiNm "
    sql &= "	,	trn_jisseki.ShohinCD "
    sql &= "	,	trn_jisseki.ShohinNM "
    sql &= "	,	CASE WHEN "
    sql &= "		trn_jisseki.ShohinCD >= 100 "
    sql &= "		then '' "
    sql &= "		else trn_jisseki.Iro  "
    sql &= "		end IRISU "
    'sql &= "	,	CASE WHEN "
    'sql &= "		trn_jisseki.ShohinCD >= 100 "
    'sql &= "		then '' "
    'sql &= "		else trn_jisseki.Suryo "
    'sql &= "		end Suryo "
    sql &= "	, trn_jisseki.Suryo "
    sql &= "	,	CASE WHEN "
    sql &= "		trn_jisseki.ShohinCD >= 100 "
    sql &= "		then '0' "
    sql &= "		else trn_jisseki.Tanka "
    sql &= "		end Tanka "
    sql &= "	,	trn_jisseki.UriageKin "
    sql &= "	,	CASE WHEN "
    sql &= "    trn_jisseki.ShohinCD >= 100 "
    sql &= "    then '10%' "
    sql &= "    else '8%' "   'プログラム側で8%対応1%対応を行う
    sql &= "    end Zeiritsu "
    sql &= "	,	CASE WHEN "
    sql &= "		trn_jisseki.ShohinCD >= 100 "
    sql &= "		then '' "
    sql &= "		else trn_jisseki.Biko "
    sql &= "		end Biko "
    sql &= "	,	trn_jisseki.TyokuCD "
    sql &= "	,	trn_jisseki.TyokuNM "
    sql &= "	,	0 AS SortNumber "
    '↓追加項目
    sql &= "  ,   trn_jisseki.UriageKin AS BARCODE_VALUE "
    sql &= "  ,   trn_jisseki.TokuiTel "
    sql &= "  ,   MST_TOKUISAKI.FaxNo "
    sql &= "  ,   trn_jisseki.TokuiAdd1 "
    sql &= "  ,   trn_jisseki.TokuiAdd2 "
    sql &= "  ,   trn_jisseki.TokuiZipCD "
    sql &= "  ,   MST_TOKUISAKI.TokuiInvoiceNumber AS TOKUI_INVOICE_NUMBER "   '事業者登録番号を得意先マスタに追加
    sql &= "  ,   UriageKin AS ZEIKOMI_KIN "   '集計項目
    sql &= "  ,   '' AS ZEINUKI_KIN "   '集計項目
    sql &= "  ,   CAST(ISNULL(Utizei, 0) AS INT) AS ZEIKOMI_TAX "
    sql &= "  ,   CAST(ISNULL(Sotozei, 0) AS INT) AS ZEINUKI_TAX "
    sql &= "  ,   CAST(ISNULL(Utizei, 0) AS INT)  + CAST(ISNULL(Sotozei, 0) AS INT)  AS TOTAL_TAX "
    sql &= "  ,   MST_TANTO.NAME AS TANTO_NM "
    sql &= "  ,   TRN_JISSEKI.Tani "
    sql &= "  ,   ' 8%※' AS TAX_8 "   'プログラム側で8%対応1%対応を行う
    sql &= "  ,   '' AS TAX_10 "
    '↑追加項目
    sql &= " FROM trn_jisseki "
    sql &= " LEFT JOIN MST_TOKUISAKI_SHOHIN "
    sql &= " ON MST_TOKUISAKI_SHOHIN.TokuiCD = trn_jisseki.TokuiCD "
    sql &= " AND MST_TOKUISAKI_SHOHIN.ShohinCD  = trn_jisseki.ShohinCD "
    sql &= " LEFT JOIN MST_TOKUISAKI_SHOHIN TOKUISAKI0 "
    sql &= " ON TOKUISAKI0.TokuiCD = 0 "
    sql &= " AND TOKUISAKI0.ShohinCD  =  trn_jisseki.ShohinCD "
    sql &= " LEFT JOIN M_TOKUISAKI_PRINT_CTRL "
    sql &= " ON M_TOKUISAKI_PRINT_CTRL.TOKUISAKI_CD = trn_jisseki.TokuiCD "
    sql &= "LEFT JOIN MST_TOKUISAKI "
    sql &= "ON MST_TOKUISAKI.TokuiCD = trn_jisseki.TokuiCD "
    sql &= "LEFT JOIN MST_TANTO "
    sql &= "ON MST_TANTO.CODE = trn_jisseki.UTantoCD "
    sql &= " WHERE 1=1 "
    For Each strValue As KeyValuePair(Of String, String) In prmWhereList
      sql &= " AND " & strValue.Key & " " & strValue.Value
    Next
    sql &= " ORDER BY trn_jisseki.DenNO,trn_jisseki.GyoNo "

    Return sql
  End Function

  Private Overloads Function SqlGetPrintData(Optional ByRef prmWhereList As Dictionary(Of String, List(Of String)) = Nothing) As String
    Dim sql As String = String.Empty
    Dim ReportType As String = ReadSettingIniFile("REPORT_TYPE", "VALUE")

    If ReportType = SOKUJI Then
      If prmWhereList.ContainsKey("INSTANT_PRINT_FLG IS NOT ") Then
        prmWhereList.Remove("INSTANT_PRINT_FLG IS NOT ")
      End If
    End If


    sql &= "SELECT	trn_jisseki.NohinDay "
    sql &= "	,	trn_jisseki.DenNO2 + '*' DenNo "
    sql &= "	,	RIGHT('00' + CAST(trn_jisseki.GyoNo2 AS VARCHAR(2)), 2) GyoNo "
    sql &= "	,	trn_jisseki.TokuiCD "
    sql &= "	,	trn_jisseki.TokuiNm "
    sql &= "	,	trn_jisseki.ShohinCD "
    sql &= "	,	trn_jisseki.ShohinNM "
    sql &= "	,	CASE WHEN "
    sql &= "		trn_jisseki.ShohinCD >= 100 "
    sql &= "		then '' "
    sql &= "		else trn_jisseki.Iro  "
    sql &= "		end IRISU "
    'sql &= "	,	CASE WHEN "
    'sql &= "		trn_jisseki.ShohinCD >= 100 "
    'sql &= "		then '' "
    'sql &= "		else trn_jisseki.Suryo "
    'sql &= "		end Suryo "
    sql &= "	, trn_jisseki.Suryo "
    sql &= "	,	CASE WHEN "
    sql &= "		trn_jisseki.ShohinCD >= 100 "
    sql &= "		then '0' "
    sql &= "		else trn_jisseki.Tanka "
    sql &= "		end Tanka "
    sql &= "	,	trn_jisseki.UriageKin "
    sql &= "  , CASE WHEN "
    sql &= "    trn_jisseki.ShohinCD >= 100 "
    sql &= "    then '10%' "
    sql &= "    else '8%' "   'プログラム側で8%対応1%対応を行う
    sql &= "    end Zeiritsu "
    sql &= "	,	CASE WHEN "
    sql &= "		trn_jisseki.ShohinCD >= 100 "
    sql &= "		then '' "
    sql &= "		else trn_jisseki.Biko "
    sql &= "		end Biko "
    sql &= "	,	trn_jisseki.TyokuCD "
    sql &= "	,	trn_jisseki.TyokuNM "
    sql &= "	,	0 AS SortNumber "
    '↓追加項目
    sql &= "  ,   trn_jisseki.UriageKin AS BARCODE_VALUE "
    sql &= "  ,   trn_jisseki.TokuiTel "
    sql &= "  ,   MST_TOKUISAKI.FaxNo "
    sql &= "  ,   trn_jisseki.TokuiAdd1 "
    sql &= "  ,   trn_jisseki.TokuiAdd2 "
    sql &= "  ,   trn_jisseki.TokuiZipCD "
    sql &= "  ,   MST_TOKUISAKI.TokuiInvoiceNumber AS TOKUI_INVOICE_NUMBER "   '事業者登録番号を得意先マスタに追加
    sql &= "  ,   '' AS ZEINUKI_KIN "   '集計項目
    sql &= "  ,   CAST(ISNULL(Utizei, 0) AS INT) AS ZEIKOMI_TAX "
    sql &= "  ,   CAST(ISNULL(Sotozei, 0) AS INT) AS ZEINUKI_TAX "
    sql &= "  ,   CAST(ISNULL(Utizei, 0) AS INT)  + CAST(ISNULL(Sotozei, 0) AS INT)  AS TOTAL_TAX "
    sql &= "  ,   MST_TANTO.NAME AS TANTO_NM "
    sql &= "  ,   TRN_JISSEKI.Tani "
    sql &= "  ,   ' 8%※' AS TAX_8 "   'プログラム側で8%対応1%対応を行う
    sql &= "  ,   '' AS TAX_10 "
    '↑追加項目
    sql &= "FROM trn_jisseki "
    sql &= "LEFT JOIN MST_TOKUISAKI_SHOHIN "
    sql &= "ON MST_TOKUISAKI_SHOHIN.TokuiCD = trn_jisseki.TokuiCD "
    sql &= "AND MST_TOKUISAKI_SHOHIN.ShohinCD  = trn_jisseki.ShohinCD "
    sql &= "LEFT JOIN MST_TOKUISAKI_SHOHIN TOKUISAKI0 "
    sql &= "ON TOKUISAKI0.TokuiCD = 0 "
    sql &= "AND TOKUISAKI0.ShohinCD  =  trn_jisseki.ShohinCD "
    sql &= "LEFT JOIN MST_TOKUISAKI "
    sql &= "ON MST_TOKUISAKI.TokuiCD = trn_jisseki.TokuiCD "
    sql &= "LEFT JOIN MST_TANTO "
    sql &= "ON MST_TANTO.CODE = trn_jisseki.UTantoCD "
    sql &= "WHERE 1=1 "
    If prmWhereList IsNot Nothing Then
      For Each tmpValue As KeyValuePair(Of String, List(Of String)) In prmWhereList

        Dim key As String = tmpValue.Key
        Dim vals As List(Of String) = tmpValue.Value

        If vals.Count = 1 Then
          sql &= $" AND {key} = '{vals(0)}'"
        Else
          Dim inList = String.Join(",", vals.Select(Function(v) $"'{v}'"))
          sql &= $" AND {key} IN ({inList})"
        End If

      Next
    End If

    sql &= "ORDER BY trn_jisseki.DenNO,trn_jisseki.GyoNo "

    Return sql
  End Function
  'Private Overloads Function SqlGetPrintData(Optional ByRef prmWhereList As Dictionary(Of String, String) = Nothing) As String
  '  Dim sql As String = String.Empty
  '  Dim ReportType As String = ReadSettingIniFile("REPORT_TYPE", "VALUE")

  '  sql &= " Select	trn_jisseki.NohinDay "
  '  sql &= "	,	Case When trn_jisseki.DenNO2 Is NULL Then trn_jisseki.DenNO Else trn_jisseki.DenNO2 + '*' END AS DenNo "
  '  sql &= "	,	RIGHT('00' + CAST(ISNULL(trn_jisseki.GyoNo2,trn_jisseki.GyoNo)AS VARCHAR(2)), 2)  GyoNo "
  '  sql &= "	,	trn_jisseki.TokuiCD "
  '  sql &= "	,	trn_jisseki.TokuiNm "
  '  sql &= "	,	trn_jisseki.ShohinCD "
  '  sql &= "	,	trn_jisseki.ShohinNM "
  '  sql &= "	,	CASE WHEN "
  '  sql &= "		trn_jisseki.ShohinCD >= 100 "
  '  sql &= "		then '' "
  '  sql &= "		else trn_jisseki.Iro  "
  '  sql &= "		end IRISU "
  '  'sql &= "	,	CASE WHEN "
  '  'sql &= "		trn_jisseki.ShohinCD >= 100 "
  '  'sql &= "		then '' "
  '  'sql &= "		else trn_jisseki.Suryo "
  '  'sql &= "		end Suryo "
  '  sql &= "	, trn_jisseki.Suryo "
  '  sql &= "	,	CASE WHEN "
  '  sql &= "		trn_jisseki.ShohinCD >= 100 "
  '  sql &= "		then '0' "
  '  sql &= "		else trn_jisseki.Tanka "
  '  sql &= "		end Tanka "
  '  sql &= "	,	trn_jisseki.UriageKin "
  '  sql &= "	,	'8%' Zeiritsu "
  '  sql &= "	,	CASE WHEN "
  '  sql &= "		trn_jisseki.ShohinCD >= 100 "
  '  sql &= "		then '' "
  '  sql &= "		else trn_jisseki.Biko "
  '  sql &= "		end Biko "
  '  sql &= "	,	trn_jisseki.TyokuCD "
  '  sql &= "	,	trn_jisseki.TyokuNM "
  '  sql &= "	,	0 AS SortNumber "
  '  sql &= " FROM trn_jisseki "
  '  sql &= " LEFT JOIN MST_TOKUISAKI_SHOHIN "
  '  sql &= " ON MST_TOKUISAKI_SHOHIN.TokuiCD = trn_jisseki.TokuiCD "
  '  sql &= " AND MST_TOKUISAKI_SHOHIN.ShohinCD  = trn_jisseki.ShohinCD "
  '  sql &= " LEFT JOIN MST_TOKUISAKI_SHOHIN TOKUISAKI0 "
  '  sql &= " ON TOKUISAKI0.TokuiCD = 0 "
  '  sql &= " AND TOKUISAKI0.ShohinCD  =  trn_jisseki.ShohinCD "
  '  sql &= " LEFT JOIN M_TOKUISAKI_PRINT_CTRL "
  '  sql &= " ON M_TOKUISAKI_PRINT_CTRL.TOKUISAKI_CD = trn_jisseki.TokuiCD "
  '  sql &= " WHERE 1=1 "
  '  For Each strValue As KeyValuePair(Of String, String) In prmWhereList
  '    sql &= " AND " & strValue.Key & " " & strValue.Value
  '  Next
  '  sql &= " ORDER BY trn_jisseki.DenNO,trn_jisseki.GyoNo "

  '  Return sql
  'End Function

  'Private Overloads Function SqlGetPrintData(Optional ByRef prmWhereList As Dictionary(Of String, List(Of String)) = Nothing) As String
  '  Dim sql As String = String.Empty
  '  Dim ReportType As String = ReadSettingIniFile("REPORT_TYPE", "VALUE")

  '  sql &= "SELECT	trn_jisseki.NohinDay "
  '  sql &= "	,	trn_jisseki.DenNO2 + '*' DenNo "
  '  sql &= "	,	RIGHT('00' + CAST(trn_jisseki.GyoNo2 AS VARCHAR(2)), 2) GyoNo "
  '  sql &= "	,	trn_jisseki.TokuiCD "
  '  sql &= "	,	trn_jisseki.TokuiNm "
  '  sql &= "	,	trn_jisseki.ShohinCD "
  '  sql &= "	,	trn_jisseki.ShohinNM "
  '  sql &= "	,	CASE WHEN "
  '  sql &= "		trn_jisseki.ShohinCD >= 100 "
  '  sql &= "		then '' "
  '  sql &= "		else trn_jisseki.Iro  "
  '  sql &= "		end IRISU "
  '  'sql &= "	,	CASE WHEN "
  '  'sql &= "		trn_jisseki.ShohinCD >= 100 "
  '  'sql &= "		then '' "
  '  'sql &= "		else trn_jisseki.Suryo "
  '  'sql &= "		end Suryo "
  '  sql &= "	, trn_jisseki.Suryo "
  '  sql &= "	,	CASE WHEN "
  '  sql &= "		trn_jisseki.ShohinCD >= 100 "
  '  sql &= "		then '0' "
  '  sql &= "		else trn_jisseki.Tanka "
  '  sql &= "		end Tanka "
  '  sql &= "	,	trn_jisseki.UriageKin "
  '  sql &= "	,	'8%' Zeiritsu "
  '  sql &= "	,	CASE WHEN "
  '  sql &= "		trn_jisseki.ShohinCD >= 100 "
  '  sql &= "		then '' "
  '  sql &= "		else trn_jisseki.Biko "
  '  sql &= "		end Biko "
  '  sql &= "	,	trn_jisseki.TyokuCD "
  '  sql &= "	,	trn_jisseki.TyokuNM "
  '  sql &= "	,	0 AS SortNumber "
  '  sql &= "FROM trn_jisseki "
  '  sql &= "LEFT JOIN MST_TOKUISAKI_SHOHIN "
  '  sql &= "ON MST_TOKUISAKI_SHOHIN.TokuiCD = trn_jisseki.TokuiCD "
  '  sql &= "AND MST_TOKUISAKI_SHOHIN.ShohinCD  = trn_jisseki.ShohinCD "
  '  sql &= "LEFT JOIN MST_TOKUISAKI_SHOHIN TOKUISAKI0 "
  '  sql &= "ON TOKUISAKI0.TokuiCD = 0 "
  '  sql &= "AND TOKUISAKI0.ShohinCD  =  trn_jisseki.ShohinCD "
  '  sql &= "WHERE 1=1 "
  '  If prmWhereList IsNot Nothing Then
  '    For Each tmpValue As KeyValuePair(Of String, List(Of String)) In prmWhereList

  '      Dim key As String = tmpValue.Key
  '      Dim vals As List(Of String) = tmpValue.Value

  '      If vals.Count = 1 Then
  '        sql &= $" AND {key} = '{vals(0)}'"
  '      Else
  '        Dim inList = String.Join(",", vals.Select(Function(v) $"'{v}'"))
  '        sql &= $" AND {key} IN ({inList})"
  '      End If

  '    Next
  '  End If

  '  sql &= "ORDER BY trn_jisseki.DenNO,trn_jisseki.GyoNo "

  '  Return sql
  'End Function

  ''' <summary>
  ''' 量目表テーブル追加SQL文作成
  ''' </summary>
  ''' <param name="tblName">テーブル名</param>
  ''' <param name="tmpRow">設定値</param>
  ''' <param name="dt">更新日付</param>
  ''' <returns>作成したSQL文</returns>
  Private Function SqlInsNohin(tblName As String,
                                    tmpRow As DataRow,
                                    dt As DateTime) As String

    Dim sql As String = String.Empty


    sql &= " INSERT INTO " & tblName
    sql &= "                   ( NOUHIN_DAY "                      '01:                       
    sql &= "                   , DEN_NO "                   '02:
    sql &= "                   , GYO_NO "                      '03:
    sql &= "                   , TOKUI_CD "                      '04:  
    sql &= "                   , TOKUI_NM "                      '05:
    sql &= "                   , SHOHIN_CD "                    '06:
    sql &= "                   , SHOHIN_NM "                      '07:
    sql &= "                   , IRISU "                      '08:
    sql &= "                   , SURYO "                      '09:
    sql &= "                   , TANKA "                      '10:
    sql &= "                   , URIAGE_KIN "                       '11:
    sql &= "                   , ZEIRITSU "                       '12:
    sql &= "                   , BIKO "                       '13:
    sql &= "                   , HASSO_CD "                       '14:
    sql &= "                   , HASSO_NM "                       '15:
    sql &= "                   , SORT_NUMBER "                       '16:
    sql &= "                   , KDATE "                       '17:
    sql &= "                   , BARCODE_VALUE "                       '18:
    sql &= "                   , TOKUI_TEL "                       '19:
    sql &= "                   , TOKUI_FAX "                       '20:
    sql &= "                   , TOKUI_ADD1 "                       '21:
    sql &= "                   , TOKUI_ADD2 "                       '22:
    sql &= "                   , TOKUI_ZIPCD "                       '23:
    sql &= "                   , TOKUI_INVOICE_NUMBER "                       '24:
    sql &= "                   , ZEIKOMI_KIN "                       '25:
    sql &= "                   , ZEINUKI_KIN "                       '26:
    sql &= "                   , ZEIKOMI_TAX "                       '27:
    sql &= "                   , ZEINUKI_TAX "                       '28:
    sql &= "                   , TOTAL_TAX "                       '29:
    sql &= "                   , TANTO_NM "                       '30:
    sql &= "                   , TANI "                       '31:
    sql &= "                   , TAX_8 "                       '32:
    sql &= "                   , TAX_10 "                       '33:
    sql &= ") VALUES("

    '納品日
    If String.IsNullOrWhiteSpace(tmpRow("NohinDay").ToString) Then
      sql &= "Null,"                                          '01:
    Else
      sql &= "'" & DateFormatChange(typDateFormat.FORMAT_DATE, tmpRow("NohinDay").ToString) & "'" & ","    '01:
    End If

    '伝票No
    If String.IsNullOrWhiteSpace(tmpRow("DenNO").ToString) Then
      sql &= "0,"                                          '02:
    Else
      sql &= "'" & tmpRow("DenNO").ToString & "'" & ","                   '02:
    End If

    '行No
    If String.IsNullOrWhiteSpace(tmpRow("GyoNo").ToString) Then
      sql &= "NULL,"                                          '03:
    Else
      sql &= "'" & tmpRow("GyoNo").ToString & "'" & ","       '03:
    End If

    '得意先コード
    If String.IsNullOrWhiteSpace(tmpRow("TokuiCD").ToString) Then
      sql &= "0,"                                          '04:
    Else
      sql &= "'" & tmpRow("TokuiCD").ToString & "'" & ","                   '04:
    End If

    '得意先名
    If String.IsNullOrWhiteSpace(tmpRow("TokuiNm").ToString) Then
      sql &= "NULL,"                                          '05:
    Else
      sql &= "'" & tmpRow("TokuiNm").ToString & "'" & ","     '05:
    End If

    '商品コード
    If String.IsNullOrWhiteSpace(tmpRow("ShohinCD").ToString) Then
      sql &= "0,"                                          '06:
    Else
      sql &= "'" & tmpRow("ShohinCD").ToString & "'" & ","                   '06:
    End If

    '商品名
    If String.IsNullOrWhiteSpace(tmpRow("ShohinNM").ToString) Then
      sql &= "NULL,"                                          '07:
    Else
      sql &= "'" & tmpRow("ShohinNM").ToString & "'" & ","                   '07:
    End If

    '入り数
    If String.IsNullOrWhiteSpace(tmpRow("Irisu").ToString) Then
      sql &= "0,"                                          '08:
    Else
      sql &= "'" & tmpRow("Irisu").ToString & "'" & ","       '08:
    End If

    '数量
    If String.IsNullOrWhiteSpace(tmpRow("Suryo").ToString) Then
      sql &= "0,"                                          '09:
    Else
      sql &= "'" & tmpRow("Suryo").ToString & "'" & ","       '09:
    End If

    '単価
    If String.IsNullOrWhiteSpace(tmpRow("Tanka").ToString) Then
      sql &= "0,"                                          '10:
    Else
      sql &= "'" & tmpRow("Tanka").ToString & "'" & ","                   '10:
    End If

    '売上金額
    If String.IsNullOrWhiteSpace(tmpRow("UriageKin").ToString) Then
      sql &= "0,"                                          '11:
    Else
      sql &= "'" & tmpRow("UriageKin").ToString & "'" & ","                   '11:
    End If

    '税率
    'TODO 本来なら税率を商品マスタから取得して設定する。
    sql &= "'" & tmpRow("Zeiritsu").ToString & "'" & ","                  '12:                        '12:


    '備考
    If String.IsNullOrWhiteSpace(tmpRow("Biko").ToString) Then
      sql &= "NULL,"                                          '13:
    Else
      sql &= "'" & tmpRow("Biko").ToString & "'" & ","                   '13:
    End If

    '直送先コード
    If String.IsNullOrWhiteSpace(tmpRow("TyokuCd").ToString) Then
      sql &= "0,"                                          '14:
    Else
      sql &= "'" & tmpRow("TyokuCd").ToString & "'" & ","                   '14:
    End If

    '直送先名
    If String.IsNullOrWhiteSpace(tmpRow("TyokuNM").ToString) Then
      sql &= "NULL,"                                          '15:
    Else
      sql &= "'" & tmpRow("TyokuNM").ToString & "'" & ","                   '15:
    End If
    'ソート番号
    If String.IsNullOrWhiteSpace(tmpRow("SortNumber").ToString) Then
      sql &= "0,"                                          '16:
    Else
      sql &= "'" & tmpRow("SortNumber").ToString & "'" & ","                   '16:
    End If

    '更新日
    sql &= "'" & dt.ToString & "',"       '16:

    'バーコード金額
    If String.IsNullOrWhiteSpace(tmpRow("BARCODE_VALUE").ToString) Then
      sql &= "NULL,"                                          '17:
    Else
      sql &= "'" & tmpRow("BARCODE_VALUE").ToString & "',"                     '17:
    End If

    '''TODO 値を設定する必要がある！

    '得意先電話番号
    If String.IsNullOrWhiteSpace(tmpRow("TokuiTel").ToString) Then
      sql &= "NULL,"                                          '17:
    Else
      sql &= "'" & tmpRow("TokuiTel").ToString & "',"                     '17:
    End If

    '得意先FAX
    If String.IsNullOrWhiteSpace(tmpRow("FaxNo").ToString) Then
      sql &= "NULL,"                                          '17:
    Else
      sql &= "'" & tmpRow("FaxNo").ToString & "',"                     '17:
    End If

    '住所１
    If String.IsNullOrWhiteSpace(tmpRow("TokuiAdd1").ToString) Then
      sql &= "NULL,"                                          '17:
    Else
      sql &= "'" & tmpRow("TokuiAdd1").ToString & "',"                     '17:
    End If

    '住所２
    If String.IsNullOrWhiteSpace(tmpRow("TokuiAdd2").ToString) Then
      sql &= "NULL,"                                          '17:
    Else
      sql &= "'" & tmpRow("TokuiAdd2").ToString & "',"                     '17:
    End If

    '郵便番号
    If String.IsNullOrWhiteSpace(tmpRow("TokuiZipCD").ToString) Then
      sql &= "NULL,"                                          '17:
    Else
      sql &= "'" & tmpRow("TokuiZipCD").ToString & "',"                     '17:
    End If

    '事業者登録番号
    If String.IsNullOrWhiteSpace(tmpRow("TOKUI_INVOICE_NUMBER").ToString) Then
      sql &= "NULL,"                                          '17:
    Else
      sql &= "'" & tmpRow("TOKUI_INVOICE_NUMBER").ToString & "',"                     '17:
    End If

    '税込み金額
    If String.IsNullOrWhiteSpace(tmpRow("ZEIKOMI_KIN").ToString) Then
      sql &= "NULL,"                                          '17:
    Else
      sql &= "'" & tmpRow("ZEIKOMI_KIN").ToString & "',"                     '17:
    End If

    '税抜き金額
    If String.IsNullOrWhiteSpace(tmpRow("ZEINUKI_KIN").ToString) Then
      sql &= "NULL,"                                          '17:
    Else
      sql &= "'" & tmpRow("ZEINUKI_KIN").ToString & "',"                     '17:
    End If

    '税込み消費税
    If String.IsNullOrWhiteSpace(tmpRow("ZEIKOMI_TAX").ToString) Then
      sql &= "'',"                                          '17:
    Else
      sql &= "'" & tmpRow("ZEIKOMI_TAX").ToString & "',"                     '17:
    End If

    '税抜き消費税
    If String.IsNullOrWhiteSpace(tmpRow("ZEINUKI_TAX").ToString) Then
      sql &= "NULL,"                                          '17:
    Else
      sql &= "'" & tmpRow("ZEINUKI_TAX").ToString & "',"                     '17:
    End If

    '消費税額
    If String.IsNullOrWhiteSpace(tmpRow("TOTAL_TAX").ToString) Then
      sql &= "NULL,"                                          '17:
    Else
      sql &= "'" & tmpRow("TOTAL_TAX").ToString & "',"                     '17:
    End If

    '担当者名
    If String.IsNullOrWhiteSpace(tmpRow("TANTO_NM").ToString) Then
      sql &= "NULL,"                                          '17:
    Else
      sql &= "'" & tmpRow("TANTO_NM").ToString & "',"                     '17:
    End If

    '単位
    If String.IsNullOrWhiteSpace(tmpRow("Tani").ToString) Then
      sql &= "NULL,"                                          '17:
    Else
      sql &= "'" & tmpRow("Tani").ToString & "',"                     '17:
    End If

    '8%消費税
    If String.IsNullOrWhiteSpace(tmpRow("TAX_8").ToString) Then
      sql &= "NULL,"                                          '17:
    Else
      sql &= "'" & tmpRow("TAX_8").ToString & "',"                     '17:
    End If

    '10%消費税
    If String.IsNullOrWhiteSpace(tmpRow("TAX_10").ToString) Then
      sql &= "NULL"                                          '17:
    Else
      sql &= "'" & tmpRow("TAX_10").ToString & "'"                     '17:
    End If

    sql &= " )"
    Console.WriteLine(sql)

    Return sql

  End Function

  ''' <summary>
  ''' WK_MASTER テーブル追加SQL文作成
  ''' </summary>
  ''' <param name="tblName">テーブル名</param>
  ''' <param name="tmpRow">設定値</param>
  ''' <returns>作成したSQL文</returns>
  Private Function SqlInsMaster(tblName As String,
                              tmpRow As DataRow) As String

    Dim sql As String = String.Empty

    sql &= " INSERT INTO " & tblName
    sql &= "                   ( SIIRE_CD "          '01:
    sql &= "                   , SIIRE_NM "          '02:
    sql &= "                   , ITEM_CD "           '03:
    sql &= "                   , ITEM_NM "           '04:
    sql &= "                   , TANTO_CD "          '05:
    sql &= "                   , TANTO_NM "          '06:
    sql &= ") VALUES("

    '仕入先コード
    If String.IsNullOrWhiteSpace(tmpRow("SIIRE_CD").ToString) Then
      sql &= "NULL,"                               '01:
    Else
      sql &= "'" & tmpRow("SIIRE_CD").ToString & "',"   '01:
    End If

    '仕入先名
    If String.IsNullOrWhiteSpace(tmpRow("SIIRE_NM").ToString) Then
      sql &= "NULL,"                               '02:
    Else
      sql &= "'" & tmpRow("SIIRE_NM").ToString & "',"   '02:
    End If

    '商品コード
    If String.IsNullOrWhiteSpace(tmpRow("ITEM_CD").ToString) Then
      sql &= "NULL,"                               '03:
    Else
      sql &= "'" & tmpRow("ITEM_CD").ToString & "',"    '03:
    End If

    '商品名
    If String.IsNullOrWhiteSpace(tmpRow("ITEM_NM").ToString) Then
      sql &= "NULL,"                               '04:
    Else
      sql &= "'" & tmpRow("ITEM_NM").ToString & "',"    '04:
    End If

    '担当者コード
    If String.IsNullOrWhiteSpace(tmpRow("TANTO_CD").ToString) Then
      sql &= "NULL,"                               '05:
    Else
      sql &= "'" & tmpRow("TANTO_CD").ToString & "',"   '05:
    End If

    '担当者名
    If String.IsNullOrWhiteSpace(tmpRow("TANTO_NM").ToString) Then
      sql &= "NULL"                                '06:
    Else
      sql &= "'" & tmpRow("TANTO_NM").ToString & "'"    '06:
    End If

    sql &= " )"

    Return sql

  End Function

  Private Overloads Function SqlUpdPrintFlg(prmDb As DataRow, Optional ByRef prmWhereList As Dictionary(Of String, String) = Nothing) As String
    Dim sql As String = String.Empty

    sql &= " Update TRN_JISSEKI "
    sql &= " SET TRN_JISSEKI.NohinPRTFLG = 1"
    sql &= " FROM TRN_JISSEKI "
    sql &= " LEFT JOIN M_TOKUISAKI_PRINT_CTRL "
    sql &= " ON M_TOKUISAKI_PRINT_CTRL.TOKUISAKI_CD = TRN_JISSEKI.TokuiCD "
    sql &= " WHERE 1 = 1 "
    If Not prmWhereList Is Nothing _
      AndAlso prmWhereList.Count > 0 Then
      For Each strValue As KeyValuePair(Of String, String) In prmWhereList
        sql &= "AND " & strValue.Key & " " & strValue.Value
      Next
    Else
      sql &= " AND DenNo = '" & prmDb.Item("DenNO").ToString & "'"
      sql &= " AND GyoNo = '" & prmDb.Item("GyoNO").ToString & "'"

    End If

    Return sql
  End Function

  Private Overloads Function SqlUpdPrintFlg(prmDb As DataRow, Optional ByRef prmWhereList As Dictionary(Of String, List(Of String)) = Nothing) As String
    Dim sql As String = String.Empty

    sql &= " Update TRN_JISSEKI "
    sql &= " SET NohinPRTFLG = 1"
    sql &= " WHERE 1 = 1 "
    If prmWhereList IsNot Nothing Then
      For Each tmpValue As KeyValuePair(Of String, List(Of String)) In prmWhereList

        Dim key As String = tmpValue.Key
        Dim vals As List(Of String) = tmpValue.Value

        If vals.Count = 1 Then
          sql &= $" AND {key} '{vals(0)}'"
        Else
          Dim inList = String.Join(",", vals.Select(Function(v) $"'{v}'"))
          sql &= $" AND {key} IN ({inList})"
        End If

      Next
    End If
    Return sql
  End Function

  Private Overloads Function SqlUpdGroupData(tblName As String, prmRow As DataRow, Optional ByRef prmWhereList As Dictionary(Of String, String) = Nothing) As String
    Dim sql As String = String.Empty

    sql &= " Update  " & tblName
    sql &= " SET  BARCODE_VALUE = " & prmRow.Item("BARCODE_VALUE").ToString
    sql &= " ,    ZEIKOMI_TAX = " & prmRow.Item("ZEIKOMI_TAX").ToString
    sql &= " ,    ZEINUKI_TAX = " & prmRow.Item("ZEINUKI_TAX").ToString
    sql &= " ,    TOTAL_TAX = " & prmRow.Item("TOTAL_TAX").ToString
    sql &= " WHERE 1 = 1 "
    sql &= " AND DEN_NO = '" & prmRow.Item("DenNO").ToString & "'"
    sql &= " AND BARCODE_VALUE <> ''"

    Return sql
  End Function

End Class
