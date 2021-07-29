Public Class DynUpdateGrid2Language
    Implements IDisposable


    Public Property NoFundMaster As String = "找不到SysProgramId代號"
    Public Property NoPermission As String = "無權限操作此功能"
    Public Property NoSO1109AData As String = "{0} 代號找不到任何SO1109A資料!"
    Public Property NoSO1109BData As String = "{0} ProgramId 找不到任何SO1109B資料!"
    Public Property NoTableName As String = "SO1109A 尚未設定 TableName 欄位!"
    Public Property NoCondProgIdProp As String = "SO1109A 尚未設定 CondProgId 欄位!"
    Public Property NoSetSO1109BField As String = "SO1109B 尚未設定 FieldName 欄位!AutoSerialNo = {0} "
    Public Property NoSetFieldTypeOne As String = "SO1109B 尚未設定 FieldType =1 的資料!"
    Public Property NotFoundTable As String = "資料庫找不到 {0} Table!"
    Public Property NotFoundField As String = "{0} 找不到 {1} 欄位!AutoSerialNo = {2}"
    Public Property DetailFieldMustBe As String = "SourceTable 或 SourceField 尚未設定! AutoSerialNo = {0}"
    Public Property QueryDataError As String = "無任何動態產生資料!"
    Public Property NotFoundSourceField As String = "QueryData {0}.{1} 在SO1109B {2} 欄位找不到資料!"
    Public Property NotFoundPKValue As String = "Query Data 找不到任何PK值!"
    Public Property ConverDataTypeError As String = "{0} = {1} 無法轉換至實際欄位型態為{3}"
    Public Property GetNoWhere As String = "{0}取出Where條件有誤!"
    Public Property GetSeqNoError As String = "執行 {0} 語法錯誤!--{1}"

    Public Property CopyOK As String = "公司別:{0},複製成功!"
    Public Property CodeExists As String = "公司別:{0},資料已存在!"
    Public Property CopyErr As String = "公司別:{0},複製失敗!,錯誤原因:{1}"
    Public Property DataExists As String = "{0}重複!"
    Public Property AddClientInfo As String = "{0}新增"
    Public Property EditClientInfo As String = "{0}修改"
    Public Property DelClientInfo As String = "{0}刪除"
    Public Property CopyClientInfo As String = "{0}複製"
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
                NoFundMaster = Nothing
                NoPermission = Nothing
                NoSO1109AData = Nothing
                NoSO1109BData = Nothing
                NoTableName = Nothing
                NoCondProgIdProp = Nothing
                NoSetSO1109BField = Nothing
                NoSetFieldTypeOne = Nothing
                NotFoundTable = Nothing
                NotFoundField = Nothing
                DetailFieldMustBe = Nothing
                QueryDataError = Nothing
                NotFoundSourceField = Nothing
                NotFoundPKValue = Nothing
                ConverDataTypeError = Nothing
                GetNoWhere = Nothing
                GetSeqNoError = Nothing

                CopyOK = Nothing
                CodeExists = Nothing
                CopyErr = Nothing
                DataExists = Nothing
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
