Public Class PRDAL
    Inherits CableSoft.BLL.Utility.DALBasic
    Public Sub New()

    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub

    Friend Function GetPRCode() As String
        Return String.Format("Select CodeNo,Description,RefNo,WorkUnit From CD007 Where (ServiceType ={0}0 Or ServiceType is null) And StopFlag = 0", Sign)
    End Function

    Friend Function GetPRCodeByContactRefNo() As String
        Return String.Format("Select CodeNo,Description,RefNo,WorkUnit From CD007 Where (ServiceType ={0}0 Or ServiceType is null) And StopFlag = 0 And RefNo = {0}1", Sign)
    End Function

    Friend Function GetPRReasonCode() As String
        Return String.Format("Select CodeNo,Description,RefNo From CD014 Where (ServiceType = {0}0 Or ServiceType is null) And StopFlag = 0", Sign)
    End Function

    Friend Function GetPRReasonDescCode() As String
        Return String.Format("Select CodeNo,Description,RefNo From CD014A Where (ServiceType = {0}0 Or ServiceType is null) And StopFlag = 0 And Where CodeNo in (Select ReasonDescCode From CD014B Where ReasonCode = {0}1)", Sign)
    End Function

    Friend Function GetGroupCode() As String
        Return String.Format("Select * From CD003 A Where Exists (Select 1 From CD002CM003 B Where A.CodeNo = B.EmpNo And ServCode = {0}1 And Type = 3) And StopFlag = 0", Sign)
    End Function

    Friend Function GetWorkerEn() As String
        Return "Select * From CM003 Where StopFlag = 0"
    End Function

    Friend Function GetReturnCode() As String
        Return String.Format("Select CodeNo,Description,RefNo From CD015 Where StopFlag = 0 And (ServiceType is null or ServiceType ={0}0)", Sign)
    End Function

    Friend Function GetReturnDescCode() As String
        Return String.Format("Select CodeNo,Description,RefNo From CD072 Where StopFlag = 0 And (ServiceType is null or ServiceType ={0}0)", Sign)
    End Function

    Friend Function GetSignEn() As String
        Return "Select * From CM003 Where StopFlag = 0"
    End Function

    Friend Function GetSatiCode() As String
        Return String.Format("Select CodeNo,Description,RefNo From CD026 Where StopFlag = 0 And (ServiceType is null or ServiceType ={0}0)", Sign)
    End Function

    Friend Function GetCustomer() As String
        Return String.Format("SELECT A.*,B.ServArea,B.ClassName1,B.InstAddress,B.Tel1,Nvl(B.Balance,0) Balance From SO002 A,SO001 B " & _
                            " Where A.CustId = B.CustId And A.CustId = {0}0 And A.ServiceType={0}1", Sign)
    End Function

    Friend Function GetSO042() As String
        Return String.Format("SELECT * FROM SO042 WHERE SERVICETYPE={0}0", Sign)
    End Function

    Friend Function GetChangePRCode() As String
        Return String.Format("Select Count(*) From SO004 Where CustId ={0}0 And PRDate is null And FaciCode in (Select CodeNo From CD022 Where RefNo in ({0}1)", Sign)
    End Function

    Friend Function GetChangePRCode2() As String
        Return String.Format("Select CodeNo ,Description From CD007 Where RefNo in ({0}0) And (ServiceType = {0}1 Or ServiceType is null) Order By CodeNo", Sign)
    End Function

End Class
