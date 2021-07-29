Public Class EnterPayDALMultiDB
    Inherits EnterPayDAL
    Implements IDisposable
    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Overrides Function GetPara6(ByVal ServiceType As String) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                If Not String.IsNullOrEmpty(ServiceType) Then
                    Return String.Format("Select Nvl(Para6,0) Para6 From SO043 Where CompCode = {0}0 " &
                                              " And ServiceType = '" & ServiceType & "'  LIMIT  1 ", Sign)
                Else
                    Return String.Format("Select Nvl(Para6,0) Para6 From SO043 Where CompCode = {0}0 " &
                                             "   LIMIT  1 ", Sign)
                End If
            Case Else
                Return MyBase.GetPara6(ServiceType)
        End Select


    End Function
    Friend Overrides Function GetSO033Data(ByVal BillLen As Integer) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim result As String = Nothing
                Select Case BillLen
                    Case 11


                        result = String.Format("Select A.*,B.CustName,A.CTID::text ,C.CustStatusName,Nvl(C.CustStatusCode,0) CustStatusCode " &
                                    " From SO033 A  LEFT JOIN SO002 C ON A.SERVICETYPE = C.SERVICETYPE LEFT JOIN  SO001 B  " &
                                    " on A.CUSTID = B.CUSTID And A.COMPCODE = C.COMPCODE  " &
                                    " Where A.UCCode Not In (Select CodeNo From CD013 Where RefNo in (3,7,8) Or PayOk = 1) " &
                                    " AND A.UCCODE IS NOT NULL " &
                                    " AND A.COMPCODE = {0}0 " &
                                    " AND B.CUSTID = C.CUSTID " &
                                    " AND NVL(A.CancelFlag,0) = 0  " &
                                    " AND MediaBillNo = {0}1", Sign)
                        Return result
                    Case 12

                        result = String.Format("Select A.*,B.CustName,A.CTID::text ,C.CustStatusName,Nvl(C.CustStatusCode,0) CustStatusCode" &
                                     " From SO033 A  LEFT JOIN SO002 C ON A.SERVICETYPE = C.SERVICETYPE LEFT JOIN  SO001 B " &
                                     " ON A.CUSTID = B.CUSTID AND A.COMPCODE = B.COMPCODE " &
                                     " Where A.UCCode Not In (Select CodeNo From CD013 Where RefNo in (3,7,8) Or PayOk = 1) " &
                                     " AND A.UCCODE IS NOT NULL " &
                                     " AND B.CUSTID = C.CUSTID " &
                                     " AND A.COMPCODE = {0}0 " &
                                     " AND NVL(A.CancelFlag,0) = 0  " &
                                    " AND PrtSNo  = {0}1", Sign)
                        Return result
                    Case Else


                        result = String.Format("Select A.*,B.CustName,A.CTID::text,C.CustStatusName,Nvl(C.CustStatusCode,0) CustStatusCode " &
                                     " From SO033 A  LEFT JOIN SO002 C ON A.SERVICETYPE = C.SERVICETYPE LEFT JOIN SO001 B" &
                                     " ON A.CUSTID = B.CUSTID AND A.COMPCODE = B.COMPCODE " &
                                     " Where A.UCCode Not In (Select CodeNo From CD013 Where RefNo in (3,7,8) Or PayOk = 1) " &
                                     " AND A.UCCODE IS NOT NULL " &
                                     " AND B.CUSTID = C.CUSTID " &
                                     " AND NVL(A.CancelFlag,0) = 0  " &
                                     " AND A.COMPCODE = {0}0 " &
                                    " AND BillNo  = {0}1", Sign)
                        Return result
                End Select
            Case Else
                Return MyBase.GetSO033Data(BillLen)
        End Select


    End Function
    Friend Overrides Function GetChargeData() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return String.Format("Select A.*,A.CTID::text From SO033 A Where A.BillNo = {0}0 AND A.ITEM = {0}1", Sign)
            Case Else
                Return MyBase.GetChargeData
        End Select

    End Function
    Friend Overrides Function GetParameters() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return "Select * From (  " &
                        "   Select Nvl(A.TranDate,to_date('19900101','yyyymmdd')) TranDate " &
                        " ,Nvl(B.DayCut,0) DayCut From SO062 A,SO041 B Where A.Type = 1 Order By A.TranDate Desc) A  LIMIT  1"
            Case Else
                Return MyBase.GetParameters
        End Select


    End Function
    Friend Overrides Function GetDefaultUCCode() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return String.Format("Select CodeNo,Description From CD013 Where RefNo = {0}0 And StopFlag <> 1  LIMIT  1 ", Sign)
            Case Else
                Return MyBase.GetDefaultUCCode
        End Select

    End Function
    Friend Overrides Function GetDefaultCMCode() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return String.Format("Select CodeNo,Description From CD031 Where RefNo = 1 And StopFlag <> 1  LIMIT  1", Sign)
            Case Else
                Return MyBase.GetDefaultCMCode
        End Select

    End Function
    Friend Overrides Function GetDefaultPTCode() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return String.Format("Select CodeNo,Description From CD032 Where RefNo = 1 And StopFlag <> 1  LIMIT  1", Sign)
            Case Else
                Return MyBase.GetDefaultPTCode
        End Select

    End Function
    Friend Overrides Function GetCMAndPTData() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return String.Format("Select CMCode,CMName,PTCode,PTName " &
                             " From SO033 Where CustId = {0}0  " &
                             " And FaciSeqNo = {0}1 " &
                             " And CitemCode = {0}2 " &
                             "  LIMIT  1 ", Sign)
            Case Else
                Return MyBase.GetCMAndPTData
        End Select

    End Function
End Class
