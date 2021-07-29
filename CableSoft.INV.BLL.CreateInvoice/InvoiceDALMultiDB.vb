Public Class InvoiceDALMultiDB
    Inherits InvoiceDAL
    Implements IDisposable
    Public Sub New()

    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Function QueryINV001ServiceType(ByVal serviceType As String) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return String.Format("select row_number() OVER () codeno,unnest Description from  unnest(ARRAY[{0}])", serviceType)
            Case Else
                Return "select rownum  codeno, column_value Description  " &
                    " From TABLE(SYS.ODCIVARCHAR2LIST(" & serviceType & ")) order by column_value"
        End Select

    End Function
End Class
