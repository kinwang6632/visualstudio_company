Public Class PRDALMultiDB
    Inherits PRDAL
    Implements IDisposable

    Public Sub New()

    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Overrides Function GetSO001() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return String.Format("Select A.CTID::text,A.* From SO001 A WHERE A.Custid={0}0 ", Sign)
            Case Else
                Return MyBase.GetSO001
        End Select

    End Function

    Friend Overrides Function GetSO002(ByVal ServiceType As String) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim strwhere As String = String.Empty
                If Not String.IsNullOrEmpty(ServiceType) Then
                    strwhere = String.Format("and A.ServiceType='{0}'", ServiceType)
                End If
                Return String.Format("Select A.CTID::text,A.* From SO002 A WHERE A.Custid={0}0 {1}", Sign, strwhere)
            Case Else
                Return MyBase.GetSO002(ServiceType)
        End Select

    End Function
End Class
