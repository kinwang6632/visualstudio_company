Public Class ManualNoDALMultiDB
    Inherits ManualNoDAL

    Implements IDisposable
    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Overrides Function QuerySeqVal() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return "SELECT sf_getsequenceno('S_SO126_SEQ')"
            Case Else
                Return MyBase.QuerySeqVal
        End Select
    End Function
    Friend Overrides Function QueryData(ByVal ds As DataSet) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim tbWhere As DataTable = ds.Tables(0)
                Dim strWhere As String = " 1 = 1 "
                If Not DBNull.Value.Equals(tbWhere.Rows(0).Item("GETPAPERDATE1")) AndAlso Not String.IsNullOrEmpty(tbWhere.Rows(0).Item("GETPAPERDATE1")) Then
                    strWhere = strWhere & String.Format(" And GetPaperDate >= To_Date('{0}','yyyy/mm/dd') ", tbWhere.Rows(0).Item("GETPAPERDATE1"))
                End If
                If Not DBNull.Value.Equals(tbWhere.Rows(0).Item("GETPAPERDATE2")) AndAlso Not String.IsNullOrEmpty(tbWhere.Rows(0).Item("GETPAPERDATE2")) Then
                    strWhere = strWhere & String.Format(" And GetPaperDate <= To_Date('{0}','yyyy/mm/dd') ", tbWhere.Rows(0).Item("GETPAPERDATE2"))
                End If
                If Not DBNull.Value.Equals(tbWhere.Rows(0).Item("EMPNO")) AndAlso Not String.IsNullOrEmpty(tbWhere.Rows(0).Item("EMPNO")) Then
                    strWhere = strWhere & String.Format(" And EMPNO IN ({0})", tbWhere.Rows(0).Item("EMPNO"))
                End If
                If Not DBNull.Value.Equals(tbWhere.Rows(0).Item("PREFIX")) AndAlso Not String.IsNullOrEmpty(tbWhere.Rows(0).Item("PREFIX")) Then
                    strWhere = strWhere & String.Format(" And PREFIX IN ({0})", tbWhere.Rows(0).Item("PREFIX"))
                End If
                If Not DBNull.Value.Equals(tbWhere.Rows(0).Item("SEQNO")) AndAlso Not String.IsNullOrEmpty(tbWhere.Rows(0).Item("SEQNO")) Then
                    strWhere = strWhere & String.Format("And {0} between to_number(BEGINNUM) and  to_number(ENDNUM)", tbWhere.Rows(0).Item("SEQNO"))
                End If
                Return "Select A.*,A.CTID::text From SO126 A Where " & strWhere
            Case Else
                Return MyBase.QueryData(ds)
        End Select

    End Function
End Class
