Public Class BatchCreateDALMultiDB
    Inherits BatchCreateDAL
    Implements IDisposable

    Public Sub New()

    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Overrides Function QueryCanCreateInv(ByVal invoicekind As Integer) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim result As String = Nothing
                Dim invoicekindStr As String = " - 1"
                If invoicekind = 2 Then invoicekindStr = "InvoiceKind"
                result = "Select SUM( COUNTS ) Count FROM  " &
                           " ( Select B.SEQ, CEIL( COUNT( B.SEQ ) / {0}0 ::decimal) As COUNTS  " &
                           "   FROM INV016 A, INV017 B   WHERE A.SEQ = B.SEQ  " &
                           "   And A.COMPID =  {0}1  And A.CHARGEDATE BETWEEN  " &
                           " {0}2  And {0}3  And A.BEASSIGNEDINVID = 'N'  " &
                           "   AND A.ISVALID = 'Y'   AND A.HOWTOCREATE = {0}4  " &
                           "   AND A.SHOULDBEASSIGNED = 'Y'  AND A.INVAMOUNT > 0  " &
                          "   AND A.TAXTYPE <> '0'  AND A.STOPFLAG = 0 " &
                          "   AND B.SHOULDBEASSIGNED = 'Y' And   (Nvl(InvoiceKind,0)  = {0}5 Or InvoiceKind = " & invoicekindStr & " )" &
                          "  GROUP BY B.SEQ  ) "
                Return String.Format(result, Sign)
            Case Else
                Return MyBase.QueryCanCreateInv(invoicekind)

        End Select

    End Function
    Friend Overrides Function QueryUnusualInv() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                'Return  String.Format("Select max(LogTime) LogTime From INV033 Where CompId= {0}0 And LOGTIME = {0}1", Sign)
                Return MyBase.QueryUnusualInv
            Case Else
                Return MyBase.QueryUnusualInv
        End Select

        ' Return String.Format("Select max(LogTime) LogTime From INV033 Where CompId= {0}0 And LOGTIME = {0}1", Sign)
        'Return "select max(LogTime) LogTime from inv033 where compid = '3' and rownum<=10"
    End Function
End Class
