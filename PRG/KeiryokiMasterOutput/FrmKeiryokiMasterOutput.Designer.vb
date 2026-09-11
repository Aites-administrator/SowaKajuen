Imports T.R.ZCommonCtrl
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmKeiryokiMasterOutput
  Inherits System.Windows.Forms.Form

  <System.Diagnostics.DebuggerNonUserCode()>
  Protected Overrides Sub Dispose(ByVal disposing As Boolean)
    Try
      If disposing AndAlso components IsNot Nothing Then
        components.Dispose()
      End If
    Finally
      MyBase.Dispose(disposing)
    End Try
  End Sub

  Private components As System.ComponentModel.IContainer

  <System.Diagnostics.DebuggerStepThrough()>
  Private Sub InitializeComponent()
    Me.tabMaster = New System.Windows.Forms.TabControl()
    Me.tabTokuisaki = New System.Windows.Forms.TabPage()
    Me.tabShohin = New System.Windows.Forms.TabPage()
    Me.tabTanto = New System.Windows.Forms.TabPage()
    Me.dgvTokuisaki = New System.Windows.Forms.DataGridView()
    Me.dgvShohin = New System.Windows.Forms.DataGridView()
    Me.dgvTanto = New System.Windows.Forms.DataGridView()
    Me.txtTokuisakiCd = New T.R.ZCommonCtrl.TxtBase()
    Me.txtTokuisakiNm = New T.R.ZCommonCtrl.TxtBase()
    Me.txtShohinCd = New T.R.ZCommonCtrl.TxtBase()
    Me.txtShohinNm = New T.R.ZCommonCtrl.TxtBase()
    Me.txtTantoCd = New T.R.ZCommonCtrl.TxtBase()
    Me.txtTantoNm = New T.R.ZCommonCtrl.TxtBase()
    Me.lblTokuisakiCd = New System.Windows.Forms.Label()
    Me.lblTokuisakiNm = New System.Windows.Forms.Label()
    Me.lblShohinCd = New System.Windows.Forms.Label()
    Me.lblShohinNm = New System.Windows.Forms.Label()
    Me.lblTantoCd = New System.Windows.Forms.Label()
    Me.lblTantoNm = New System.Windows.Forms.Label()
    Me.btnAdd = New T.R.ZCommonCtrl.BtnBase()
    Me.btnDelete = New T.R.ZCommonCtrl.BtnBase()
    Me.btnOutput = New T.R.ZCommonCtrl.BtnBase()
    Me.btnBarcodePrint = New T.R.ZCommonCtrl.BtnBase()
    Me.btnClose = New T.R.ZCommonCtrl.BtnBase()
    Me.LblMessage = New System.Windows.Forms.Label()
    Me.tabMaster.SuspendLayout()
    Me.tabTokuisaki.SuspendLayout()
    Me.tabShohin.SuspendLayout()
    Me.tabTanto.SuspendLayout()
    CType(Me.dgvTokuisaki, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.dgvShohin, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.dgvTanto, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SuspendLayout()
    '
    'tabMaster
    '
    Me.tabMaster.Controls.Add(Me.tabTokuisaki)
    Me.tabMaster.Controls.Add(Me.tabShohin)
    Me.tabMaster.Controls.Add(Me.tabTanto)
    Me.tabMaster.Font = New System.Drawing.Font("Segoe UI", 13.0!)
    Me.tabMaster.Location = New System.Drawing.Point(12, 12)
    Me.tabMaster.Name = "tabMaster"
    Me.tabMaster.SelectedIndex = 0
    Me.tabMaster.Size = New System.Drawing.Size(1010, 735)
    '
    'tabTokuisaki
    '
    Me.tabTokuisaki.Controls.Add(Me.lblTokuisakiCd)
    Me.tabTokuisaki.Controls.Add(Me.txtTokuisakiCd)
    Me.tabTokuisaki.Controls.Add(Me.lblTokuisakiNm)
    Me.tabTokuisaki.Controls.Add(Me.txtTokuisakiNm)
    Me.tabTokuisaki.Controls.Add(Me.dgvTokuisaki)
    Me.tabTokuisaki.Text = "仕入先"
    Me.tabTokuisaki.UseVisualStyleBackColor = True
    '
    'tabShohin
    '
    Me.tabShohin.Controls.Add(Me.lblShohinCd)
    Me.tabShohin.Controls.Add(Me.txtShohinCd)
    Me.tabShohin.Controls.Add(Me.lblShohinNm)
    Me.tabShohin.Controls.Add(Me.txtShohinNm)
    Me.tabShohin.Controls.Add(Me.dgvShohin)
    Me.tabShohin.Text = "商品"
    Me.tabShohin.UseVisualStyleBackColor = True
    '
    'tabTanto
    '
    Me.tabTanto.Controls.Add(Me.lblTantoCd)
    Me.tabTanto.Controls.Add(Me.txtTantoCd)
    Me.tabTanto.Controls.Add(Me.lblTantoNm)
    Me.tabTanto.Controls.Add(Me.txtTantoNm)
    Me.tabTanto.Controls.Add(Me.dgvTanto)
    Me.tabTanto.Text = "担当者"
    Me.tabTanto.UseVisualStyleBackColor = True
    '
    ' Search controls
    '
    Me.lblTokuisakiCd.AutoSize = True
    Me.lblTokuisakiCd.Font = New System.Drawing.Font("Segoe UI", 15.75!)
    Me.lblTokuisakiCd.Location = New System.Drawing.Point(12, 12)
    Me.lblTokuisakiCd.Text = "仕入先コード"
    Me.txtTokuisakiCd.Font = New System.Drawing.Font("Segoe UI", 15.75!)
    Me.txtTokuisakiCd.Location = New System.Drawing.Point(150, 10)
    Me.txtTokuisakiCd.Size = New System.Drawing.Size(180, 35)
    Me.txtTokuisakiNm.Font = New System.Drawing.Font("Segoe UI", 15.75!)
    Me.txtTokuisakiNm.Location = New System.Drawing.Point(500, 10)
    Me.txtTokuisakiNm.Size = New System.Drawing.Size(360, 35)
    Me.lblTokuisakiNm.AutoSize = True
    Me.lblTokuisakiNm.Font = New System.Drawing.Font("Segoe UI", 15.75!)
    Me.lblTokuisakiNm.Location = New System.Drawing.Point(355, 12)
    Me.lblTokuisakiNm.Text = "仕入先名"

    Me.lblShohinCd.AutoSize = True
    Me.lblShohinCd.Font = New System.Drawing.Font("Segoe UI", 15.75!)
    Me.lblShohinCd.Location = New System.Drawing.Point(12, 12)
    Me.lblShohinCd.Text = "商品コード"
    Me.txtShohinCd.Font = New System.Drawing.Font("Segoe UI", 15.75!)
    Me.txtShohinCd.Location = New System.Drawing.Point(150, 10)
    Me.txtShohinCd.Size = New System.Drawing.Size(180, 35)
    Me.lblShohinNm.AutoSize = True
    Me.lblShohinNm.Font = New System.Drawing.Font("Segoe UI", 15.75!)
    Me.lblShohinNm.Location = New System.Drawing.Point(355, 12)
    Me.lblShohinNm.Text = "商品名"
    Me.txtShohinNm.Font = New System.Drawing.Font("Segoe UI", 15.75!)
    Me.txtShohinNm.Location = New System.Drawing.Point(500, 10)
    Me.txtShohinNm.Size = New System.Drawing.Size(360, 35)

    Me.lblTantoCd.AutoSize = True
    Me.lblTantoCd.Font = New System.Drawing.Font("Segoe UI", 15.75!)
    Me.lblTantoCd.Location = New System.Drawing.Point(12, 12)
    Me.lblTantoCd.Text = "担当者コード"
    Me.txtTantoCd.Font = New System.Drawing.Font("Segoe UI", 15.75!)
    Me.txtTantoCd.Location = New System.Drawing.Point(150, 10)
    Me.txtTantoCd.Size = New System.Drawing.Size(180, 35)
    Me.lblTantoNm.AutoSize = True
    Me.lblTantoNm.Font = New System.Drawing.Font("Segoe UI", 15.75!)
    Me.lblTantoNm.Location = New System.Drawing.Point(355, 12)
    Me.lblTantoNm.Text = "担当者名"
    Me.txtTantoNm.Font = New System.Drawing.Font("Segoe UI", 15.75!)
    Me.txtTantoNm.Location = New System.Drawing.Point(500, 10)
    Me.txtTantoNm.Size = New System.Drawing.Size(360, 35)

    ' grids
    Me.dgvTokuisaki.Location = New System.Drawing.Point(12, 60)
    Me.dgvTokuisaki.Name = "dgvTokuisaki"
    Me.dgvTokuisaki.Size = New System.Drawing.Size(970, 620)
    Me.dgvShohin.Location = New System.Drawing.Point(12, 60)
    Me.dgvShohin.Name = "dgvShohin"
    Me.dgvShohin.Size = New System.Drawing.Size(970, 620)
    Me.dgvTanto.Location = New System.Drawing.Point(12, 60)
    Me.dgvTanto.Name = "dgvTanto"
    Me.dgvTanto.Size = New System.Drawing.Size(970, 620)

    ' buttons (right side)
    Me.btnAdd.Font = New System.Drawing.Font("Segoe UI", 14.25!)
    Me.btnAdd.Location = New System.Drawing.Point(1040, 40)
    Me.btnAdd.Name = "btnAdd"
    Me.btnAdd.Size = New System.Drawing.Size(135, 55)
    Me.btnAdd.Text = "F1:行追加"
    Me.btnAdd.UseVisualStyleBackColor = True

    Me.btnDelete.Font = New System.Drawing.Font("Segoe UI", 14.25!)
    Me.btnDelete.Location = New System.Drawing.Point(1040, 105)
    Me.btnDelete.Name = "btnDelete"
    Me.btnDelete.Size = New System.Drawing.Size(135, 55)
    Me.btnDelete.Text = "F3:削除"
    Me.btnDelete.UseVisualStyleBackColor = True

    Me.btnOutput.Font = New System.Drawing.Font("Segoe UI", 14.25!)
    Me.btnOutput.Location = New System.Drawing.Point(1040, 170)
    Me.btnOutput.Name = "btnOutput"
    Me.btnOutput.Size = New System.Drawing.Size(135, 55)
    Me.btnOutput.Text = "F5:CSV出力"
    Me.btnOutput.UseVisualStyleBackColor = True

    Me.btnBarcodePrint.Font = New System.Drawing.Font("Segoe UI", 12.0!)
    Me.btnBarcodePrint.Location = New System.Drawing.Point(1040, 235)
    Me.btnBarcodePrint.Name = "btnBarcodePrint"
    Me.btnBarcodePrint.Size = New System.Drawing.Size(135, 55)
    Me.btnBarcodePrint.Text = "F8:バーコード" & vbCrLf & "印刷"
    Me.btnBarcodePrint.UseVisualStyleBackColor = True

    Me.btnClose.Font = New System.Drawing.Font("Segoe UI", 14.25!)
    Me.btnClose.Location = New System.Drawing.Point(1040, 670)
    Me.btnClose.Name = "btnClose"
    Me.btnClose.Size = New System.Drawing.Size(135, 55)
    Me.btnClose.Text = "F10:終了"
    Me.btnClose.UseVisualStyleBackColor = True

    Me.LblMessage.AutoSize = True
    Me.LblMessage.Font = New System.Drawing.Font("Segoe UI", 12.0!)
    Me.LblMessage.Location = New System.Drawing.Point(12, 765)
    Me.LblMessage.Text = ""

    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(1200, 800)
    Me.Controls.Add(Me.LblMessage)
    Me.Controls.Add(Me.btnClose)
    Me.Controls.Add(Me.btnBarcodePrint)
    Me.Controls.Add(Me.btnOutput)
    Me.Controls.Add(Me.btnDelete)
    Me.Controls.Add(Me.btnAdd)
    Me.Controls.Add(Me.tabMaster)
    Me.Name = "FrmKeiryokiMasterOutput"
    Me.Text = "計量器マスタ出力"
    Me.tabMaster.ResumeLayout(False)
    Me.tabTokuisaki.ResumeLayout(False)
    Me.tabTokuisaki.PerformLayout()
    Me.tabShohin.ResumeLayout(False)
    Me.tabShohin.PerformLayout()
    Me.tabTanto.ResumeLayout(False)
    Me.tabTanto.PerformLayout()
    CType(Me.dgvTokuisaki, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.dgvShohin, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.dgvTanto, System.ComponentModel.ISupportInitialize).EndInit()
    Me.ResumeLayout(False)
    Me.PerformLayout()
  End Sub

  Friend WithEvents tabMaster As TabControl
  Friend WithEvents tabTokuisaki As TabPage
  Friend WithEvents tabShohin As TabPage
  Friend WithEvents tabTanto As TabPage
  Friend WithEvents dgvTokuisaki As DataGridView
  Friend WithEvents dgvShohin As DataGridView
  Friend WithEvents dgvTanto As DataGridView
  Friend WithEvents txtTokuisakiCd As T.R.ZCommonCtrl.TxtBase
  Friend WithEvents txtTokuisakiNm As T.R.ZCommonCtrl.TxtBase
  Friend WithEvents txtShohinCd As T.R.ZCommonCtrl.TxtBase
  Friend WithEvents txtShohinNm As T.R.ZCommonCtrl.TxtBase
  Friend WithEvents txtTantoCd As T.R.ZCommonCtrl.TxtBase
  Friend WithEvents txtTantoNm As T.R.ZCommonCtrl.TxtBase
  Friend WithEvents lblTokuisakiCd As Label
  Friend WithEvents lblTokuisakiNm As Label
  Friend WithEvents lblShohinCd As Label
  Friend WithEvents lblShohinNm As Label
  Friend WithEvents lblTantoCd As Label
  Friend WithEvents lblTantoNm As Label
  Friend WithEvents btnAdd As T.R.ZCommonCtrl.BtnBase
  Friend WithEvents btnDelete As T.R.ZCommonCtrl.BtnBase
  Friend WithEvents btnOutput As T.R.ZCommonCtrl.BtnBase
  Friend WithEvents btnBarcodePrint As T.R.ZCommonCtrl.BtnBase
  Friend WithEvents btnClose As T.R.ZCommonCtrl.BtnBase
  Friend WithEvents LblMessage As Label
End Class
