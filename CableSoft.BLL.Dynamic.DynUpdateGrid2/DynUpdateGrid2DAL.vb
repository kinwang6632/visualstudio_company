Imports CableSoft.BLL.Utility
Public Class DynUpdateGrid2DAL
    Inherits DALBasic
    Implements IDisposable

    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Function QueryMaster()
        Return String.Format("SELECT * FROM SO1120A WHERE SysProgramId = {0}0", Sign)
    End Function
    Friend Function GetInsertSQL(ByVal TableName As String, ByVal FieldsAndValues As Dictionary(Of String, Object)) As String
        Dim aRet As String = Nothing
        Dim aFields As String = Nothing
        Dim aValues As String = Nothing
        For i As Int32 = 0 To FieldsAndValues.Keys.Count - 1

            If String.IsNullOrEmpty(aFields) Then
                aFields = FieldsAndValues.Keys(i).ToUpper
            Else
                aFields = aFields & "," & FieldsAndValues.Keys(i).ToUpper
            End If
            If i = 0 Then
                aValues = "{0}" & i
            Else
                aValues = aValues & ",{0}" & i
            End If
        Next
        aRet = "Select " & aFields & " From " & TableName & " Where 1=0"
        Return aRet
        aRet = String.Format("INSERT INTO {0} ( {1} ) VALUES ({2})", TableName, aFields, aValues)
        aRet = String.Format(aRet, Sign)
        Return aRet
    End Function
    Friend Function getSEQNo(ByVal SourceField As String) As String
        Return "SELECT " & SourceField & ".NEXTVAL FROM DUAL"
    End Function
    Friend Function QueryDetail()
        Return String.Format("SELECT * FROM SO1120B WHERE ProgramId = {0}0 Order by RowPos,ColumnPos  ", Sign)
    End Function
    Friend Function QurerySO1109B()
        Return String.Format("SELECT * FROM SO1109B WHERE PROGRAMID ={0}0", Sign)
    End Function
    Friend Function QuerySO1109A() As String
        Return String.Format("SELECT * FROM SO1109A WHERE SysProgramId = {0}0  AND STOPFLAG <> 1", Sign)
    End Function

    Friend Function QueryDynProgId() As String
        Return String.Format("SELECT PROGRAMID FROM SO1101A WHERE SysProgramId = {0}0", Sign)
    End Function
    Friend Function QuerySchema(ByVal TableName As String) As String
        Return String.Format("SELECT * FROM {0} WHERE 1=0 ", TableName)
    End Function
#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
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
