Imports CableSoft.BLL.Utility

Public Class MaintainDAL
    Inherits DALBasic
    Implements IDisposable

    Public Sub New()

    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Function QueryCheckMFCode() As String
        Return String.Format("SELECT NVL(CheckMFCode,0) FROM SO042 " & _
                                               " WHERE SERVICETYPE={0}0", Sign)
    End Function
    Friend Function QueryCD006RefNo() As String
        Return String.Format("SELECT NVL(REFNO,0) FROM CD006 WHERE CODENO ={0}0",
                                                        Sign)
    End Function
    Friend Function GetCD002() As String
        Return String.Format("Select CODENO,DESCRIPTION FROM CD002 WHERE CODENO = {0}0 AND COMPCODE = {0}1 ", Sign)
    End Function
    Friend Function QueryPara6(ByVal ServiceType As String) As String
        Dim aWhere As String = String.Empty
        If Not String.IsNullOrEmpty(ServiceType) Then
            aWhere = String.Format(" AND SERVICETYPE='{0}'", ServiceType)
        End If
        Return String.Format("SELECT NVL(PARA6,0) FROM SO043 WHERE " & _
                                             "COMPCODE = {0}0 " & aWhere, Sign)
    End Function
    Friend Function QueryTranDate(ByVal ServiceType As String) As String
        Dim aWhere As String = String.Empty
        If Not String.IsNullOrEmpty(ServiceType) Then
            aWhere = String.Format(" AND SERVICETYPE='{0}'", ServiceType)
        End If
        Return String.Format("SELECT TRANDATE FROM SO062 WHERE COMPCODE={0}0 " & _
                                                             aWhere & " AND TYPE = 1 ORDER BY TRANDATE DESC ", Sign)
    End Function
    Friend Function QueryDayCut() As String
        Return String.Format("SELECT NVL(DAYCUT,0) FROM SO041 WHERE SYSID={0}0", Sign)
    End Function
    Friend Overridable Function QuerySO008Log() As String
        Return String.Format("Select A.rowid,A.* From SO008  A where SNO = {0}0", Sign)
    End Function
    Friend Function QueryServCode() As String
        Return String.Format("SELECT ServCode FROM SO001 WHERE CUSTID ={0}0", Sign)
    End Function
    Friend Function GetCustomer() As String
        Dim aRet As String
        'aRet = String.Format("SELECT A.*,B.ServArea," & _
        '                     "B.ClassName1,B.InstAddress," & _
        '                     "B.Tel1,Nvl(B.Balance,0) Balance From SO002 A,SO001 B " & _
        '                    " Where A.CustId = B.CustId " & _
        '                    " And A.CustId = {0}0 " & _
        '                    " And A.ServiceType={0}1", Sign)

        aRet = String.Format("SELECT A.WipCode1,A.WipCode2,A.WipCode3," & _
                             "B.ServArea,A.CustStatusCode," & _
                            "B.ClassName1,B.InstAddress," & _
                            "B.Tel1,Nvl(A.Balance,0) Balance From SO002 A,SO001 B " & _
                           " Where A.CustId = B.CustId " & _
                           " And A.CustId = {0}0 " & _
                           " And A.ServiceType={0}1", Sign)

        Return aRet
    End Function
    Friend Function GetTimePeriod() As String
        Dim aRet As String = String.Format("SELECT TimePeriod FROM SO011 " & _
                                         " WHERE ( SERVICETYPE ={0}0 OR SERVICETYPE IS NULL) " & _
                                         " AND TIMEPERIOD >= {0}1 ORDER BY TIMEPERIOD", Sign)
        Return aRet
    End Function
    Friend Function GetReserveDay() As String
        Return String.Format("SELECT NVL(ReserveDay,0) FROM CD006 " & _
                             " WHERE CODENO = {0}0", Sign)
    End Function
    Friend Function GetCD006() As String
        Return String.Format("SELECT * FROM CD006 WHERE CODENO = {0}0 ORDER BY CODENO ", Sign)
    End Function
    Friend Function GetReInstAddrNo() As String
        Dim aRet As String = String.Format("Select ReInstAddrNo From SO009 Where CustId = {0}0  " & _
                                         " And ServiceType = {0}1 " & _
                                         " And PRCode In (Select CodeNo From CD007 Where RefNo = 3 )  " & _
                                         " Order By AcceptTime Desc ", Sign)
        Return aRet
    End Function
    Friend Function GetSO014() As String
        Dim aRet As String = String.Format("Select AddrNo,Address,ServCode,StrtCode," & _
                                           " AreaCode,NodeNo,SalesCode,SalesName,ClctEn," & _
                                           " ClctName,CircuitNo From SO014 " & _
                                           " Where AddrNo = {0}0 AND COMPCODE = {0}1", Sign)
        Return aRet
    End Function
    Friend Function GetSO042() As String
        Return String.Format("SELECT * FROM SO042 WHERE SERVICETYPE={0}0", Sign)
    End Function
    Friend Function GetIsFixingAreaData() As String
        Return String.Format("Select MduId,NodeNo,CircuitNo,substr(AddrSort,0,86) AddrSort,Noe1,Noe2,Noe3,Noe4  " & _
                                " From SO014 Where AddrNo in (Select InstAddrNo From SO001 Where CustId = {0}0)", Sign)
    End Function
    Friend Function QueryFaciCode() As String
        Return String.Format("Select FaciCode From SO004 Where SeqNo= {0}0", Sign)
    End Function
    Friend Function QueryCD022RefNo() As String
        Return String.Format("Select Nvl(RefNo,0) From CD022 Where CodeNo = {0}0", Sign)
    End Function
    Friend Function QueryCD022Count() As String
        Return String.Format("Select Count(*) From CD022 Where CodeNo ={0}0  And RefNo in (3,4) ", Sign)
    End Function
    Friend Function QuerySTBFinTimeFlag() As String
        Return String.Format("SELECT NVL(STBFinTimeFlag,0) " & _
                                                                  " FROM SO042 WHERE COMPCODE={0}0 AND SERVICETYPE={0}1", Sign)
    End Function
    Friend Function QueryMustCallOk() As String
        Return String.Format("SELECT NVL(MustCallOk,0) FROM SO042 WHERE SERVICETYPE={0}0 " & _
                        " AND COMPCODE={0}1", Sign)
    End Function
    Friend Function GetCD046() As String
        Return "SELECT CodeNo,Description FROM CD046 ORDER BY CODENO "
    End Function
    Friend Function QueryCD022() As String
        Return String.Format("SELECT COUNT(1)  FROM CD022 " & _
                                             " WHERE CODENO = {0}0 AND REFNO = {0}1", Sign)
    End Function
    Friend Function QueryCanDelete() As String
        Return String.Format("SELECT COUNT(1) FROM SO008 WHERE SNO={0}0 AND CLSTIME IS NULL", Sign)
    End Function
    Friend Function QueryCanEdit() As String
        Return String.Format("SELECT COUNT(1)  FROM SO008 WHERE SNO={0}0 AND CLSTIME IS NULL", Sign)
    End Function
    Friend Function GetSO042Para() As String
        Return String.Format("SELECT NVL(MoreDay2,0) MoreDay2," & _
                                                                        "NVL(ModifyDateChange,0) ModifyDateChange FROM SO042 " & _
                                                                        " WHERE SERVICETYPE = {0}0", Sign)
    End Function
    Friend Overridable Function GetSysDate() As String
        Return "Select SysDate From Dual"
    End Function
    Friend Overridable Function IsFixingArea(ByVal MduId As String, ByVal NodeNo As String, ByVal CircuitNo As String,
                                 ByVal AddrSort86 As String, ByVal Noe1 As String, ByVal Noe2 As String, ByVal Noe3 As String, ByVal Noe4 As String) As String
        Dim aSQL = "SELECT 1   FROM SO022 A, SO023B B " &
                               " WHERE(B.Kind = 1) " &
                                " AND B.KeyData  In (" & MduId & " ) " &
                                " AND A.SNo = B.SNo  AND SYSDATE >= A.ErrorTime " &
                                " AND (A.FinTime IS NULL OR SYSDATE <= A.FinTime)  AND A.ReturnCode IS NULL " &
                                " AND A.ShowMalfunction = 1 " &
                                " Union All  " &
                            " Select 1   FROM SO022 A, SO023B B " &
                                " WHERE(B.Kind = 2) " &
                                " AND B.KeyData  In ( " & NodeNo & " ) " &
                                " AND A.SNo = B.SNo  AND SYSDATE >= A.ErrorTime " &
                                "  AND (A.FinTime IS NULL OR SYSDATE <= A.FinTime)  AND A.ReturnCode IS NULL " &
                                " AND A.ShowMalfunction = 1 " &
                                " Union All " &
                            " Select 1  FROM SO022 A, SO023B B " &
                            " WHERE B.Kind = 3 " &
                            " AND B.KeyData  In (" & CircuitNo & " ) " &
                            " AND A.SNo = B.SNo   AND SYSDATE >= A.ErrorTime " &
                            " AND (A.FinTime IS NULL OR SYSDATE <= A.FinTime)   AND A.ReturnCode IS NULL " &
                            " AND A.ShowMalfunction = 1 " &
                            " Union All " &
                             "Select 1  From SO023 A,SO022 B " &
                              " Where A.SNo=B.SNo " &
                              " And ('" & AddrSort86 & "'>=A.AddrSortA And '" & AddrSort86 & "'<=A.AddrSortB) " &
                              " And SysDate>=B.ErrorTime  And (B.FinTime Is Null Or SysDate<=B.FinTime) " &
                              " And B.ReturnCode Is Null  And B.ShowMalfunction=1" &
                              " And (A.Noe = 0 or A.Noe =  (Case when A.Alley2 is not null then " & Noe4 &
                                                                                   " When A.Alley is not null then " & Noe3 &
                                                                                    " When A.Lane is not null then " & Noe2 &
                                                                                    " else " & Noe1 & " End )) "

        Return aSQL
    End Function

    Friend Function GetSO001() As String
        'Return String.Format("SELECT A.* FROM SO001 A WHERE A.CUSTID = {0}0", Sign)
        Return String.Format("SELECT A.ServCode,A.Tel1,A.ServArea, " & _
                             "A.ClassName1,A.InstAddress,A.InstAddrNO,A.CustId,A.CustName " & _
                             " FROM SO001 A WHERE A.CUSTID = {0}0", Sign)
    End Function

    Friend Function GetMaitainCode() As String
        Dim aRet As String

        aRet = String.Format("Select CodeNo,Description,Nvl(RefNo,0) RefNo,WorkUnit," & _
                             "Nvl(GroupNo,0) GroupNo, Nvl(Resvdatebefore,0) Resvdatebefore,DefGroupCode, " & _
                             " (Select description From CD003 Where CodeNo = cd006.DefGroupCode ) DefGroupName " & _
                           " From CD006 Where (ServiceType = {0}0 Or ServiceType IS NULL) " & _
                           " And NVL(StopFlag,0) = 0 ORDER BY CODENO ", Sign)
        Return aRet
    End Function
    Friend Function GetSignEn() As String
        Return "Select * From CM003 Where Nvl(StopFlag,0) = 0"
    End Function
    Friend Function GetGroupCode2() As String
        Return "Select CodeNo,Description,0 Flag From CD003 Where NVL(StopFlag,0) = 0 "
    End Function
    Friend Function GetGroupCode() As String
        Dim aRet As String = String.Format("Select CODENO,Description,1 Flag From CD003 A " & _
                                         " Where Exists (Select * From CD002CM003 B " & _
                                         " Where A.CodeNo = B.EmpNo And ServCode = {0}0 " & _
                                         " And Type = 2) And Nvl(A.StopFlag,0) = 0 ORDER BY CODENO ", Sign)
        Return aRet
    End Function
    Friend Function GetWorkerEn() As String
        Return "Select EmpNo,EmpName From CM003 Where NVL(StopFlag,0) = 0 ORDER BY EMPNO"
    End Function
    Friend Function GetReturnCode() As String
        Dim aRet As String = String.Format("Select CodeNo,Description,RefNo " & _
                                           " From CD015 Where NVL(StopFlag,0) = 0 " & _
                                           " And (ServiceType is null or ServiceType = {0}0) ORDER BY CODENO", Sign)
        Return aRet
    End Function
    Friend Function GetReturnDescCode(ByVal ServiceType As String) As String
        Dim aRet As String = String.Format("Select CodeNo,Description,RefNo " & _
                                         " From CD072 Where Nvl(StopFlag,0) = 0 " & _
                                         " And instr(ServiceType,'" & ServiceType & "')>0  ORDER BY CODENO", Sign)
        Return aRet
    End Function
    Friend Function GetSatiCode() As String
        Dim aRet As String = String.Format("Select CodeNo,Description,RefNo " & _
                                           " From CD026 Where Nvl(StopFlag,0) = 0 " & _
                                           " And (ServiceType Is Null or ServiceType = {0}0) ORDER BY CODENO", Sign)
        Return aRet
    End Function
    Friend Function GetMFCode1() As String
        Dim aRet As String = String.Format("Select CodeNo,Description,RefNo From CD011 " & _
                                         " Where Nvl(StopFlag,0) = 0 " & _
                                         " And (ServiceType is null Or ServiceType = {0}0) ORDER BY CODENO", Sign)
        Return aRet
    End Function
    Friend Function GetMFCode2(ByVal SearchCount As Int32) As String
        Dim aRet As String = String.Empty
        Select Case SearchCount
            Case 0
                aRet = String.Format("Select  CodeNo,Description,RefNo " & _
                                   " From CD011B Where (ServiceType is null Or ServiceType = {0}0) " & _
                                   " And CodeNo In (Select CodeNo From CD011A Where MFCode = {0}1) " & _
                                   " And Nvl(StopFlag,0) = 0 ORDER BY CODENO", Sign)
            Case Else
                aRet = String.Format("Select CodeNo,Description,RefNo " & _
                                     " From CD011 Where Nvl(StopFlag,0) = 0 " & _
                                     " And (ServiceType is null Or ServiceType = {0}0) ORDER BY CODENO", Sign)
        End Select
        Return aRet
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
            End If

            ' TODO: 釋放 Unmanaged 資源 (Unmanaged 物件) 並覆寫下面的 Finalize()。
            ' TODO: 將大型欄位設定為 null。
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: 只有當上面的 Dispose(ByVal disposing As Boolean) 有可釋放 Unmanaged 資源的程式碼時，才覆寫 Finalize()。
    'Protected Overrides Sub Finalize()
    '    ' 請勿變更此程式碼。在上面的 Dispose(ByVal disposing As Boolean) 中輸入清除程式碼。
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' 由 Visual Basic 新增此程式碼以正確實作可處置的模式。
    Public Sub Dispose() Implements IDisposable.Dispose
        ' 請勿變更此程式碼。在以上的 Dispose 置入清除程式碼 (ByVal 視為布林值處置)。
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
