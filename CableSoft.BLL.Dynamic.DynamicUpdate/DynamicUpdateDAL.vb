Imports CableSoft.BLL.Utility
Public Class DynamicUpdateDAL
    Inherits DALBasic
    Implements IDisposable

    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Overridable Function getSEQNo(ByVal SourceField As String) As String
        Return "SELECT " & SourceField & ".NEXTVAL FROM DUAL"

    End Function

    Friend Function QuerySO1109A() As String
        Return String.Format("SELECT * FROM SO1109A WHERE " &
                             "SysProgramId = {0}0  And STOPFLAG <> 1 " &
                              CableSoft.BLL.Utility.Utility.GetDBType(MyBase.Provider, ""), Sign)
    End Function
    Friend Function QuerySO1109B() As String
        Return String.Format("SELECT * FROM SO1109B WHERE ProgramId = {0}0 " &
                              CableSoft.BLL.Utility.Utility.GetDBType(MyBase.Provider, ""), Sign)
    End Function
    Friend Function QueryDynProgId() As String
        Return String.Format("SELECT PROGRAMID FROM SO1101A WHERE SysProgramId = {0}0 " &
                              CableSoft.BLL.Utility.Utility.GetDBType(MyBase.Provider, ""), Sign)
    End Function
    Friend Function QuerySchema(ByVal TableName As String) As String
        Return String.Format("SELECT * FROM {0} WHERE 1=0 ", TableName)
    End Function
    Friend Function QueryCurrectData(ByVal TableName As String, ByVal WhereFieldsAndValues As Dictionary(Of String, Object)) As String
        Dim aRet As String = Nothing
        Dim aWhereFields As String = Nothing
        If WhereFieldsAndValues Is Nothing OrElse WhereFieldsAndValues.Keys.Count = 0 Then
            Return String.Format("SELECT * FROM {0} WHERE 1=0", TableName)
        End If
        For i As Int32 = 0 To WhereFieldsAndValues.Keys.Count - 1
            If String.IsNullOrEmpty(aWhereFields) Then
                aWhereFields = WhereFieldsAndValues.Keys(i).ToUpper & "={0}" & i.ToString
            Else
                aWhereFields = aWhereFields & " And " &
                    WhereFieldsAndValues.Keys(i).ToUpper & "={0}" & i.ToString
            End If
        Next
        aRet = String.Format("SELECT * FROM  {0} WHERE {1}", TableName, aWhereFields)
        aRet = String.Format(aRet, Sign)
        Return aRet
    End Function
    Friend Function GetCompCode(ByVal GroupId As String, ByVal strCD039 As String, ByVal strSO026 As String) As String
        If GroupId = "0" Then
            Return "Select A.CodeNo ,A.Description From " & strCD039 & " A Order By CodeNo"
        End If
        Return String.Format("Select distinct A.CodeNo ,A.Description " & _
                             " From " & strCD039 & " A," & strSO026 & " B  " & _
                             " Where Instr(',' ||B.CompStr|| ',' , ',' ||A.CodeNo|| ',') > 0 " & _
                             " And UserId = {0}0 Order By CodeNO", Sign)
    End Function
    Friend Function GetDelSQL(ByVal TableName As String, ByVal WhereFieldsAndValues As Dictionary(Of String, Object)) As String
        Dim aRet As String = Nothing
        Dim aWhereFields As String = Nothing
        For i As Int32 = 0 To WhereFieldsAndValues.Keys.Count - 1
            If String.IsNullOrEmpty(aWhereFields) Then
                aWhereFields = WhereFieldsAndValues.Keys(i).ToUpper & "={0}" & i.ToString
            Else
                aWhereFields = aWhereFields & " AND " & _
                    WhereFieldsAndValues.Keys(i).ToUpper & "={0}" & i.ToString
            End If
        Next
        aRet = String.Format("DELETE {0} WHERE {1}", TableName, aWhereFields)
        aRet = String.Format(aRet, Sign)
        Return aRet
    End Function
    Friend Function GetInsertAllDataSQL(ByVal TableName As String, ByVal tbSechema As DataTable) As String
        Dim ret As String = Nothing
        Dim aFields As String = Nothing
        Dim aValues As String = Nothing
        Dim i As Integer = 0

        For Each col As DataColumn In tbSechema.Columns
            If String.IsNullOrEmpty(aFields) Then
                aFields = col.ColumnName
                aValues = "{0}" & i
            Else
                aFields = aFields & "," & col.ColumnName
                aValues = aValues & ",{0}" & i
            End If
            i += 1
        Next
        ret = "Insert Into " & TableName & "(" & aFields & " ) Values (" & aValues & " ) "
        Return ret
    End Function
    
    Friend Function GetFindSQL(ByVal TableName As String, ByVal WhereFieldAndValues As Dictionary(Of String, Object)) As String
        Dim retString As String = Nothing
        Dim aFields As String = Nothing
        retString = String.Format("Select Count(*) From {0} Where 1=1 ", TableName)
        For i As Integer = 0 To WhereFieldAndValues.Keys.Count - 1
            retString = String.Format(retString & " And " & WhereFieldAndValues.Keys(i) & "={0}" & i, Sign)
        Next
        Return retString

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
        aRet = String.Format("INSERT INTO {0} ( {1} ) VALUES ({2})", TableName, aFields, aValues)
        aRet = String.Format(aRet, Sign)
        Return aRet
    End Function
    Friend Function GetUKSQL(ByVal editMode As EditMode, ByVal TableName As String, ByVal FieldName As String, ByVal WhereFieldsAndValues As Dictionary(Of String, Object)) As String
        Dim aSQL As String = Nothing
        Dim aWhereFields As String = Nothing
        If editMode = Utility.EditMode.Append Then
            aSQL = String.Format("Select Count(*) From " & TableName & " Where " & FieldName & "= {0}0", Sign)
        Else
            For i As Int32 = 0 To WhereFieldsAndValues.Keys.Count - 1
                If String.IsNullOrEmpty(aWhereFields) Then
                    aWhereFields = WhereFieldsAndValues.Keys(i).ToUpper & "<>{0}" & (i + 1).ToString
                Else
                    aWhereFields = aWhereFields & " AND " & _
                        WhereFieldsAndValues.Keys(i).ToUpper & "<>{0}" & (i + 1).ToString
                End If
                aSQL = String.Format("Select Count(*) From " & TableName & " Where " & FieldName & "= {0}0 " & _
                                     " And " & aWhereFields, Sign)
            Next
        End If
        Return aSQL
    End Function
    Friend Function GetPKSQL(ByVal TableName As String, ByVal WhereFieldsAndValues As Dictionary(Of String, Object)) As String
        Dim aRet As String = Nothing
        Dim aWhereFields As String = Nothing        'Dim aFields As String = Nothing

        For i As Int32 = 0 To WhereFieldsAndValues.Keys.Count - 1
            If String.IsNullOrEmpty(aWhereFields) Then
                aWhereFields = WhereFieldsAndValues.Keys(i).ToUpper & "={0}" & i.ToString
            Else
                aWhereFields = aWhereFields & " AND " & _
                    WhereFieldsAndValues.Keys(i).ToUpper & "={0}" & i.ToString
            End If
        Next
        aRet = String.Format("Select Count(*) From  {0} WHERE {1}", TableName, aWhereFields)
        aRet = String.Format(aRet, Sign)
        Return aRet
    End Function
    Friend Function GetUpdateSQL(ByVal TableName As String,
                                 ByVal FieldsAndValues As Dictionary(Of String, Object),
                                 ByVal WhereFieldsAndValues As Dictionary(Of String, Object)) As String
        Dim aRet As String = Nothing
        Dim aFields As String = Nothing
        Dim aWhereFields As String = Nothing
        For i As Int32 = 0 To FieldsAndValues.Keys.Count - 1
            If String.IsNullOrEmpty(aFields) Then
                aFields = FieldsAndValues.Keys(i).ToUpper & "={0}" & i.ToString
            Else
                aFields = aFields & "," & FieldsAndValues.Keys(i).ToUpper & "={0}" & i.ToString
            End If
        Next
        For i As Int32 = 0 To WhereFieldsAndValues.Keys.Count - 1
            If String.IsNullOrEmpty(aWhereFields) Then
                aWhereFields = WhereFieldsAndValues.Keys(i).ToUpper & "={0}" & (FieldsAndValues.Keys.Count + i).ToString
            Else
                aWhereFields = aWhereFields & " AND " & _
                    WhereFieldsAndValues.Keys(i).ToUpper & "={0}" & (FieldsAndValues.Keys.Count + i).ToString
            End If
        Next
        aRet = String.Format("UPDATE {0} SET {1} WHERE {2}", TableName, aFields, aWhereFields)
        aRet = String.Format(aRet, Sign)
        Return aRet
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
