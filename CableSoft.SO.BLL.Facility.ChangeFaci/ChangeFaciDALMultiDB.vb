Public Class ChangeFaciDALMultiDB
    Inherits ChangeFaciDAL
    Public Sub New()

    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Overrides Function GetAllChangeData(ByVal aSeqNos As String) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return String.Format("SELECT A.CTID::text,A.*,NVL(B.REFNO,0) FACIREFNO " &
                             " FROM SO004 A,CD022 B " &
                             " WHERE A.FACICODE = B.CODENO AND A.SEQNO IN ( {0}) ", aSeqNos)
            Case Else
                Return MyBase.GetAllChangeData(aSeqNos)
        End Select


    End Function
    Friend Overrides Function GetChangeData() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return String.Format("SELECT A.CTID::text,A.*,NVL(B.REFNO,0) FACIREFNO " &
                             " FROM SO004 A,CD022 B " &
                             " WHERE A.FACICODE = B.CODENO AND A.SEQNO = {0}0", Sign)
            Case Else
                Return MyBase.GetChangeData
        End Select

    End Function
    Friend Overloads Overrides Function GetChildFaci() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return String.Format("SELECT A.CTID::text,A.* FROM SO004 A " &
                             " WHERE CUSTID = {0}0 " &
                             " AND PRDATE IS NULL AND GETDATE IS NULL " &
                             " AND (FACISNO = {0}1 " &
                             " AND FaciCode In (Select CodeNo From CD022 Where RefNo = 4) OR STBSNO = {0}2)", Sign)
            Case Else
                Return MyBase.GetChildFaci
        End Select

    End Function
    Friend Overloads Overrides Function GetChildFaci(ByVal FilterDVR As Boolean) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                If FilterDVR Then
                    Return String.Format("SELECT A.CTID::text,A.* FROM SO004 A " &
                             " WHERE CUSTID = {0}0 " &
                             " AND PRDATE IS NULL AND GETDATE IS NULL " &
                             " AND (FACISNO = {0}1 AND FaciCode In (Select CodeNo From CD022 Where RefNo = 4))", Sign)
                Else
                    Return String.Format("SELECT A.CTID::text,A.* FROM SO004 A " &
                             " WHERE CUSTID = {0}0 " &
                             " AND PRDATE IS NULL AND GETDATE IS NULL " &
                             " AND (FACISNO = {0}1  " &
                             " AND FaciCode In (Select CodeNo From CD022 Where RefNo = 4) OR STBSNO = {0}2)", Sign)
                End If
            Case Else
                Return MyBase.GetChildFaci(FilterDVR)
        End Select

    End Function
End Class
