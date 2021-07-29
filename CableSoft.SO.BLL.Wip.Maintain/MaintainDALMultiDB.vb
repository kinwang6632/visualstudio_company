Public Class MaintainDALMultiDB
    Inherits MaintainDAL
    Implements IDisposable

    Public Sub New()

    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Overrides Function QuerySO008Log() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return String.Format("Select A.CTID::text,A.* From SO008  A where SNO = {0}0", Sign)
            Case Else
                Return MyBase.QuerySO008Log
        End Select

    End Function
    Friend Overrides Function GetSysDate() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return "Select now()"
            Case Else
                Return MyBase.GetSysDate
        End Select

    End Function
    Friend Overrides Function IsFixingArea(ByVal MduId As String, ByVal NodeNo As String, ByVal CircuitNo As String,
                                 ByVal AddrSort86 As String, ByVal Noe1 As String, ByVal Noe2 As String, ByVal Noe3 As String, ByVal Noe4 As String) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim aSQL = "SELECT 1   FROM SO022 A, SO023B B " &
                               " WHERE(B.Kind = 1) " &
                                " AND B.KeyData  In (" & MduId & " ) " &
                                " AND A.SNo = B.SNo  AND now() >= A.ErrorTime " &
                                " AND (A.FinTime IS NULL OR now() <= A.FinTime)  AND A.ReturnCode IS NULL " &
                                " AND A.ShowMalfunction = 1 " &
                                " Union All  " &
                            " Select 1   FROM SO022 A, SO023B B " &
                                " WHERE(B.Kind = 2) " &
                                " AND B.KeyData  In ( " & NodeNo & " ) " &
                                " AND A.SNo = B.SNo  AND now() >= A.ErrorTime " &
                                "  AND (A.FinTime IS NULL OR now() <= A.FinTime)  AND A.ReturnCode IS NULL " &
                                " AND A.ShowMalfunction = 1 " &
                                " Union All " &
                            " Select 1  FROM SO022 A, SO023B B " &
                            " WHERE B.Kind = 3 " &
                            " AND B.KeyData  In (" & CircuitNo & " ) " &
                            " AND A.SNo = B.SNo   AND now() >= A.ErrorTime " &
                            " AND (A.FinTime IS NULL OR now() <= A.FinTime)   AND A.ReturnCode IS NULL " &
                            " AND A.ShowMalfunction = 1 " &
                            " Union All " &
                             "Select 1  From SO023 A,SO022 B " &
                              " Where A.SNo=B.SNo " &
                              " And ('" & AddrSort86 & "'>=A.AddrSortA And '" & AddrSort86 & "'<=A.AddrSortB) " &
                              " And now()>=B.ErrorTime  And (B.FinTime Is Null Or now()<=B.FinTime) " &
                              " And B.ReturnCode Is Null  And B.ShowMalfunction=1" &
                              " And (A.Noe = 0 or A.Noe =  (Case when A.Alley2 is not null then " & Noe4 &
                                                                                   " When A.Alley is not null then " & Noe3 &
                                                                                    " When A.Lane is not null then " & Noe2 &
                                                                                    " else " & Noe1 & " End )) "

                Return aSQL
            Case Else
                Return MyBase.IsFixingArea(MduId, NodeNo, CircuitNo, AddrSort86, Noe1, Noe2, Noe3, Noe4)

        End Select

    End Function

End Class
