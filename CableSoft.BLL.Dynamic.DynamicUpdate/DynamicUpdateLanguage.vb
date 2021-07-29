Public Class DynamicUpdateLanguage
    Public Shared Property NoSO1109AData As String = "{0} 代號找不到任何SO1109A資料！"
    Public Shared Property NoSO1109BData As String = "{0} ProgramId 找不到任何SO1109B資料！"
    Public Shared Property NoTableName As String = "SO1109A 尚未設定 TableName 欄位！"
    Public Shared Property NoCondProgIdProp As String = "SO1109A 尚未設定 CondProgId 欄位！"
    Public Shared Property NoSetSO1109BField As String = "SO1109B 尚未設定 FieldName 欄位！AutoSerialNo = {0} "
    Public Shared Property NoSetFieldTypeOne As String = "SO1109B 尚未設定 FieldType =1 的資料！"
    Public Shared Property NotFoundTable As String = "資料庫找不到 {0} Table！"
    Public Shared Property NotFoundField As String = "{0} 找不到 {1} 欄位！AutoSerialNo = {2}"
    Public Shared Property DetailFieldMustBe As String = "SourceTable 或 SourceField 尚未設定！ AutoSerialNo = {0}"
    Public Shared Property QueryDataError As String = "無任何動態產生資料！"
    Friend Shared Property NotFoundSourceField As String = "QueryData {0}.{1} 在SO1109B {2} 欄位找不到資料！"
    Friend Shared Property NotFoundPKValue As String = "Query Data 找不到任何PK值！"
    Friend Shared Property ConverDataTypeError As String = "{0} = {1} 無法轉換至實際欄位型態為{3}"
    Friend Shared Property GetNoWhere As String = "{0}取出Where條件有誤！"
    Friend Shared Property GetSeqNoError As String = "執行 {0} 語法錯誤！--{1}"

End Class
