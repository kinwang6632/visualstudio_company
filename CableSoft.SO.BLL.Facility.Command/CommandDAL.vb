Imports CableSoft.BLL.Utility
Public Class CommandDAL
    Inherits DALBasic
    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Function TakeMasterTable() As String
        Dim aRet As String = String.Format("SELECT * FROM SO1102A " &
                                           " WHERE UPPER(TABLENAME) = {0}0 AND UPPER(CMDID) = {0}1", Sign)
        Return aRet
    End Function
    Friend Overridable Function GetSeqNo(ByVal OwnerName As String, ByVal SourceField As String) As String
        Return "SELECT " & OwnerName & SourceField & ".NEXTVAL FROM DUAL"
    End Function
    Friend Function TakeDetailTable() As String
        Dim aRet As String = String.Format("SELECT * FROM SO1102B " & _
                                           " WHERE SEQNO = {0}0", Sign)
        Return aRet
    End Function
    Friend Function GetReturnDetailTable(ByVal strDTableName As String, ByVal strSeqNoFieldName As String) As String
        Return String.Format("SELECT * FROM {0} WHERE {1} = ", strDTableName, strSeqNoFieldName)
    End Function
    Friend Function GetChkDetailSchema(ByVal strChkTableName As String) As String
        Return "SELECT * FROM " & strChkTableName & " WHERE 1=0 "
    End Function
    Friend Function GetTakeInsertDetailSQL(ByVal strDTableName As String, ByVal strFields As String, ByVal strValues As String) As String
        Return String.Format("INSERT INTO {0} ( {1} ) VALUES ({2})", strDTableName, strFields, strValues)
    End Function
    Friend Function DeleteCMDData(ByVal TableName As String, ByVal SeqNoField As String)
        Return String.Format("DELETE " & TableName & " WHERE " & SeqNoField & "={0}0", Sign)
    End Function
    Friend Function WriteTimeOutError(ByVal TableName As String,
                                      ByVal CmdStatusField As String,
                                      ByVal ErrorCodeField As String,
                                      ByVal ErrMsgField As String,
                                      ByVal SeqNoField As String) As String
        Dim aRet As String = Nothing
        If ErrorCodeField <> ErrMsgField Then
            aRet = String.Format("UPDATE " & TableName & " SET " & _
              CmdStatusField & "='E'," & _
              ErrorCodeField & "={0}0," & _
              ErrMsgField & "={0}1" & _
              " WHERE " & SeqNoField & "={0}2", Sign)
        Else
            aRet = String.Format("UPDATE " & TableName & " SET " & _
              CmdStatusField & "='E'," & _
              ErrorCodeField & "={0}0" & _
              " WHERE " & SeqNoField & "={0}1", Sign)
        End If
      
        Return aRet
    End Function
    Friend Function QuertyStatus(ByVal TableName As String, ByVal SeqNoFieldName As String) As String
        Dim aRet As String = String.Format("SELECT * FROM {0} WHERE {1} ", TableName, SeqNoFieldName)
        aRet = String.Format(aRet & "={0}0", Sign)
        Return aRet
    End Function
    Friend Function GetMasterInsertSQL(ByVal TableName As String,
                                 ByVal Fields As String,
                                 ByVal Values As String) As String
        Dim aRet As String = Nothing
        aRet = String.Format("INSERT INTO {0} ( {1} ) VALUES ({2})", TableName, Fields, Values)
        aRet = String.Format(aRet, Sign)
        Return aRet
    End Function
    Friend Function GetDetailInsertSQL(ByVal TableName As String, ByVal Fields As String, ByVal Values As String) As String
        Dim aRet As String = Nothing
        aRet = String.Format("INSERT INTO {0} ( {1} ) VALUES ({2})", TableName, Fields, Values)
        aRet = String.Format(aRet, Sign)
        Return aRet
    End Function
    Friend Function GetSchemaTable(ByVal TableName As String) As String
        Return "SELECT * FROM " & TableName & " WHERE 1=0 "
    End Function

End Class
