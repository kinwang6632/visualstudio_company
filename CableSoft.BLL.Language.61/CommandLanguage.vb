Public Class CommandLanguage
    Implements IDisposable

    Public Property DataTableNoData As String = "資料表無資料!!"
    Public Property ParamsError As String = "參數錯誤!!"
    Public Property LackField As String = "資料表缺少必要欄位[{0}]!!"
    Public Property MustField As String = "[{0}]為必要欄位需有值!!"
    Public Property TakeSetErr As String = "取出設定檔有誤!"
    Public Property CmdTimeOut As String = "連線逾時!"
    Public Property CmdOtherErr As String = "其它錯誤!"
    Public Property FieldNotExistsTable As String = "欄位{0}不存在於{1} Table中"
    Public Property NotFoundRealField As String = "實際欄位找不到 {0}!"
    Public Property NotFoundField As String = "{0}找不到{1}欄位"
    Public Property ConverDataTypeError As String = "{0} = {1} 無法轉換至實際欄位型態為{3}"
    Public Property SetFieldTypeError As String = "FieldType 設定有誤!--AutoSerialNo = {0}"
    Public Property TableNameIsNull As String = "傳入的TableName為NULL!"
    Public Property CmdIdIsNull As String = "CMDID為NULL!"
    Public Property NotFoundMasterCMD As String = "設定主檔找不到命令 {0} !"
    Public Property MasterCmdDouble As String = "設定主檔命令 {0} 有重複!"
    Public Property MasterNoOwner As String = "設定主檔命令 {0} 無設定Owner!"
    Public Property MasterCmdNoSucessCode As String = "設定主檔命令 {0} 無設定成功Code!"
    Public Property MasterCmdNoErrCode As String = "設定主檔命令 {0} 無設定失敗Code!"
    Public Property DetailNoData As String = "設定子檔無任何資料!"
    Public Property DetailSetError As String = "設定子檔 {0} 設定有誤!"
    Public Property DetailNotSetType As String = "設定子檔未設定FieldType = {0}的資料!"

#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                DataTableNoData = Nothing
                ParamsError = Nothing
                LackField = Nothing
                MustField = Nothing
                TakeSetErr = Nothing
                CmdTimeOut = Nothing
                CmdOtherErr = Nothing
                FieldNotExistsTable = Nothing
                NotFoundRealField = Nothing
                NotFoundField = Nothing
                ConverDataTypeError = Nothing
                SetFieldTypeError = Nothing
                TableNameIsNull = Nothing
                CmdIdIsNull = Nothing
                NotFoundMasterCMD = Nothing
                MasterCmdDouble = Nothing
                MasterNoOwner = Nothing
                MasterCmdNoSucessCode = Nothing
                MasterCmdNoErrCode = Nothing
                DetailNoData = Nothing
                DetailSetError = Nothing
                DetailNotSetType = Nothing
            End If

            ' TODO: 釋放 Unmanaged 資源 (Unmanaged 物件) 並覆寫下面的 Finalize()。
            ' TODO: 將大型欄位設定為 null。
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: 只有當上面的 Dispose(ByVal disposing As Boolean) 有可釋放 Unmanaged 資源的程式碼時，才覆寫 Finalize()。
    'Protected Overrides Sub Finalize()
    '    ' 請勿變更此程式碼。在上面的 Dispose(ByVal disposing As Boolean) 中輸入清除程式碼。
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' 由 Visual Basic 新增此程式碼以正確實作可處置的模式。
    Public Sub Dispose() Implements IDisposable.Dispose
        ' 請勿變更此程式碼。在以上的 Dispose 置入清除程式碼 (ByVal 視為布林值處置)。
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
