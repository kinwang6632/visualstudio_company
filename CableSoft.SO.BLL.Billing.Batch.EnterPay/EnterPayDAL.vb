Imports CableSoft.BLL.Utility
Public Class EnterPayDAL
    Inherits DALBasic
    Implements IDisposable
    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Function GetDayCut() As String
        Return String.Format("Select Nvl(DayCut,0) DayCut From SO041 Where CompCode = {0}0", Sign)
    End Function
    Friend Function GetTranDate(ByVal ServiceType As String) As String
        If Not String.IsNullOrEmpty(ServiceType) Then
            Return String.Format("Select tranDate From SO062 Where CompCode = {0}0 " & _
                                    " And ServiceType ='" & ServiceType & "' And Type = 1  And tranDate Is Not Null " & _
                                    " Order By TranDate Desc ", Sign)
        Else
            Return String.Format("Select tranDate From SO062 Where CompCode = {0}0 " & _
                                    " And Type = 1 And tranDate Is Not Null  Order By TranDate Desc ", Sign)
        End If
    End Function
    Friend Overridable Function GetPara6(ByVal ServiceType As String) As String
        If Not String.IsNullOrEmpty(ServiceType) Then
            Return String.Format("Select Nvl(Para6,0) Para6 From SO043 Where CompCode = {0}0 " &
                                      " And ServiceType = '" & ServiceType & "' And RowNum = 1 ", Sign)
        Else
            Return String.Format("Select Nvl(Para6,0) Para6 From SO043 Where CompCode = {0}0 " &
                                     " And RowNum = 1 ", Sign)
        End If
    End Function
    Friend Function GetCompCode(ByVal GroupId As String) As String
        If GroupId = "0" AndAlso 1 = 0 Then
            Return "Select A.CodeNo ,A.Description From CD039 A Order By CodeNo"
        Else
            Return String.Format("Select A.CodeNo,A.Description  " & _
                              " From CD039 A,SO026 B  " & _
                              " Where Instr(','||B.CompStr||',',','||A.CodeNo||',')>0 " & _
                             " And UserId = {0}0 Order By CodeNO", Sign)
        End If
    End Function
    Friend Function GetCMCode() As String
        Return "Select CodeNo,Description,RefNo From CD031 Where Nvl(StopFlag,0) = 0 Order By CodeNo"
    End Function
    Friend Function GetPTCode() As String
        Return "Select CodeNo,Description,RefNo From CD032 Where Nvl(StopFlag,0) = 0 Order By CodeNo "
    End Function
    Friend Function GetClctEn() As String
        Return "Select EmpNo,EmpName From CM003 Where Nvl(StopFlag,0) = 0 Order By EmpNo"
    End Function
    Friend Function GetSTCode() As String
        Return "Select CodeNo,Description From CD016 Where Nvl(StopFlag,0) = 0 And (RefNo is null Or RefNo <> 1 ) Order By CodeNo"
    End Function
    Friend Overridable Function GetParameters() As String
        Return "Select * From (  " &
            "   Select Nvl(A.TranDate,to_date('19900101','yyyymmdd')) TranDate " &
            " ,Nvl(B.DayCut,0) DayCut From SO062 A,SO041 B Where A.Type = 1 Order By A.TranDate Desc) Where rownum=1"
    End Function
    Friend Function IsUseManual() As String
        Return String.Format("Select Count(1) Cnt " & _
                             " From SO127 Where PaperNum = {0}0 " & _
                             " And CompCode = {0}1" & _
                             " And BillNo = {0}2", Sign)
    End Function
    Friend Function GetUseManualStatus() As String
        Return String.Format("Select Nvl(Status,0) Status " & _
                             " From SO127 Where PaperNum = {0}0 And CompCode = {0}1", Sign)
    End Function
    Friend Function ChkDupData(ByVal EntryType As Integer, ByVal BillLen As Integer) As String
        Dim aTableName As String = "SO074"
        If EntryType = 1 Then
            aTableName = "SO077"
        End If
        Select Case BillLen
            Case 11
                Return String.Format("SELECT COUNT(*) CNT FROM " & aTableName & _
                                     " WHERE MEDIABILLNO = {0}0", Sign)
            Case 12
                Return String.Format("SELECT COUNT(*) CNT FROM " & aTableName & _
                                     " WHERE PRTSNO = {0}0", Sign)
            Case Else
                Return String.Format("SELECT COUNT(*) CNT FROM " & aTableName & _
                                     " WHERE BILLNO = {0}0", Sign)
        End Select
    End Function
    Friend Function GetUseManual() As String
        Return String.Format("select Nvl(UseManual,0) UseManual from so043 where servicetype={0}0", Sign)
    End Function
    Friend Function GetEntryNoCount(ByVal EntryType As Integer) As String
        Dim tableName As String = "SO074"
        If EntryType = 1 Then
            tableName = "SO077"
        End If
        Return String.Format("SELECT NVL(MAX(EntryNo),0)   FROM " & tableName & _
                " WHERE EntryEn={0}0", Sign)
    End Function
    Friend Function GetSecondDiscount() As String
        Return String.Format("SELECT NVL(SecondDiscount,0) SecondDiscount FROM SO041 WHERE COMPCODE = {0}0 ", Sign)
    End Function
    Friend Function ChkHaveDiscount() As String
        Return String.Format("Select Nvl(SecendDiscount,0) SecendDiscount ,Nvl(RefNo,0) RefNo " & _
                             " From CD019 Where CodeNo = {0}0 " & _
                             " UNION ALL " & _
                             " SELECT 0 SecendDiscount,NVL(REFNO,0) REFNO " & _
                             " FROM CD019 WHERE ReturnCode = {0}1", Sign)
    End Function
    Friend Function DelTempData(ByVal EntryType As Integer) As String
        Dim tbName As String = "SO074"
        If EntryType = 1 Then
            tbName = "SO077"
        End If
        Return String.Format("Delete  from " & tbName & " Where BillNo={0}0 And Item = {0}1", Sign)
    End Function
    Friend Overridable Function GetChargeData() As String
        Return String.Format("Select A.*,RowId AS CTID From SO033 A Where A.BillNo = {0}0 AND A.ITEM = {0}1", Sign)
    End Function
    Friend Function CancelTempData(ByVal EntryType As Integer) As String
        Dim tableName As String = "SO074"
        Dim aSQL As String = Nothing
        If EntryType = 1 Then
            tableName = "SO077"
        End If
        aSQL = String.Format("Update " & tableName & " Set RealDate = {0}0,RealAmt = 0 ," & _
                            "STCode = Null,STName = Null, " & _
                            "CancelCode = {0}1,CancelName = {0}2,CancelFlag = 1 " & _
                            " Where BillNo = {0}3 And Item = {0}4", Sign)
        Return aSQL
    End Function
    Friend Function GetTempOK(ByVal EntryType As Integer) As String
        Dim tableName As String = "SO074"
        Dim aSQL As String = Nothing
        If EntryType = 1 Then
            tableName = "SO077"
        End If
        aSQL = "Select A.*,B.CustStatusName,Nvl(B.CustStatusCode,0) CustStatusCode From " & tableName & " A,SO002 B " & _
            " Where A.BillNo = {0}0 And A.Item = {0}1 " & _
            " And A.ServiceType = B.ServiceType " & _
            " And A.CustId = B.CustId " & _
            " And A.CompCode = B.CompCode  Order by EntryNo"
        Return String.Format(aSQL, Sign)
    End Function
    Friend Function QueryEnterData(ByVal EntryType As Integer) As String
        Dim tableName As String = "SO074"
        Dim aSQL As String = Nothing
        If EntryType = 1 Then
            tableName = "SO077"
        End If
        aSQL = "Select A.*,B.CustStatusName,Nvl(B.CustStatusCode,0) CustStatusCode From " & tableName & " A,SO002 B " & _
            " Where A.ServiceType = B.ServiceType " & _
            " And A.CustId = B.CustId " & _
            " And A.CompCode = B.CompCode " & _
            " And EntryEn = {0}0 Order by EntryNo"
        Return String.Format(aSQL, Sign)
    End Function
    Friend Function QueryCancelReason() As String
        Return String.Format("Select CodeNo, Description From CD051 Where StopFlag <> 1 And CompCode = {0}0", Sign)
    End Function
    Friend Overloads Function GetTempData(ByVal EntryType As Integer) As String
        Dim tableName As String = "SO074"
        If EntryType = 1 Then
            tableName = "SO077"
        End If
        Return String.Format("SELECT * FROM " & tableName & _
                " WHERE EntryEn={0}0", Sign)
    End Function
    Friend Overloads Function GetTempData(ByVal EntryType As Integer, ByVal BillNo As String, ByVal Item As Integer)
        Dim tableName As String = "SO074"
        Dim Result As String = Nothing
        If EntryType = 1 Then
            tableName = "SO077"
        End If
        Result = String.Format("Select * From {0} ", tableName)
        Result = String.Format("{0} Where BillNo = '{1}' And Item = {2}", Result, BillNo, Item)
        Return Result
    End Function

    Friend Function GetTempInfo(ByVal EntryType As Integer) As String
        Dim tableName As String = "SO074"
        If EntryType = 1 Then
            tableName = "SO077"
        End If
        Return String.Format("SELECT NVL(MAX(EntryNo),0) BillCount,NVL(SUM(RealAmt),0) AMTCOUNT FROM " & tableName & _
                " WHERE EntryEn={0}0", Sign)
    End Function
    Friend Function UpdDefUCCodeCharge() As String
        Return String.Format("Update SO033 Set UCCode = {0}0, " & _
                             "UCName = {0}1 " & _
                             " Where BillNo= {0}2 And Item = {0}3", Sign)
    End Function
    Friend Function CancelCharge() As String        
        Return String.Format("Update SO033 Set RealDate=Null,UCCode = {0}0,UCName = {0}1," & _
                             "CMCode={0}2,CMName={0}3,PTCode={0}4,PTName={0}5,ClctEn=OldClctEn,ClctName=OldClctName " & _
                             " Where BillNo={0}6 And Item = {0}7", Sign)
    End Function
    Friend Overridable Function GetDefaultUCCode() As String
        Return String.Format("Select CodeNo,Description From CD013 Where RefNo = {0}0 And StopFlag <> 1 And RowNum = 1", Sign)
    End Function
    Friend Overridable Function GetDefaultCMCode() As String
        Return String.Format("Select CodeNo,Description From CD031 Where RefNo = 1 And StopFlag <> 1 And RowNum = 1", Sign)
    End Function
    Friend Overridable Function GetDefaultPTCode() As String
        Return String.Format("Select CodeNo,Description From CD032 Where RefNo = 1 And StopFlag <> 1 And RowNum = 1", Sign)
    End Function
    Friend Overridable Function GetCMAndPTData() As String
        Return String.Format("Select CMCode,CMName,PTCode,PTName " &
                             " From SO033 Where CustId = {0}0  " &
                             " And FaciSeqNo = {0}1 " &
                             " And CitemCode = {0}2 " &
                             " And RowNum = 1 ", Sign)
    End Function
    Friend Function getTel1() As String
        Return String.Format("Select Tel1 From SO001 Where CustId= {0}0", Sign)
    End Function
    Friend Function UpdateSO127() As String
        Dim aSQL As String = String.Format("Update SO127  SET CustID = {0}0,CustName = {0}1," & _
                                            "CustTEL = {0}2,BillNo= {0}3,RealDate = {0}4,UPDTIME={0}5 " & _
                    " Where PaperNum = {0}6", Sign)
        Return aSQL
    End Function
    Friend Function UpdateChargeData() As String
        Dim aFieldValue As String = String.Empty
        Dim aRet As String = String.Empty
        Dim aFieldName As String = "CitemCode,CitemName,SHOULDDATE, " & _
                                            " REALDATE,SHOULDAMT,REALAMT, " & _
                         "REALPERIOD,REALSTARTDATE,REALSTOPDATE,CLCTEN," & _
                         "CLCTNAME,PTCODE,PTNAME,UPDTIME,NEWUPDTIME,UPDEN,CMCODE,CMNAME," & _
                        "MANUALNO,UCCode,UCName,STCODE,STNAME,Note,ServiceType," & _
                        "CancelFlag,CancelCode,CancelName," & _
                        "BankCode,BankName,AccountNo,AuthorizeNo,AdjustFlag," & _
                        "NextPeriod,NextAmt,InvSeqNo,FirstTime,ClctYM"
        Dim lstFieldName As List(Of String) = aFieldName.Split(",").ToList
        For i As Integer = 0 To lstFieldName.Count - 1
            If String.IsNullOrEmpty(aFieldValue) Then
                aFieldValue = String.Format("{0}={1}{2}", lstFieldName.Item(i), Sign, i.ToString)
            Else
                aFieldValue = String.Format("{0},{1}={2}{3}", aFieldValue, lstFieldName.Item(i), Sign, i.ToString)
            End If
        Next
        aRet = String.Format("Update SO033 Set {0} Where BillNo= {1}{2} And Item = {3}{4}",
                                            aFieldValue, Sign, lstFieldName.Count,
                                            Sign, lstFieldName.Count + 1)
        
        Return aRet
        '.Fields("CLCTNAME").Value = GetFieldValue(rsSource, "CLCTNAME")
        '.Fields("PTCODE").Value = GetFieldValue(rsSource, "PTCODE")
        '.Fields("PTNAME").Value = GetFieldValue(rsSource, "PTNAME")
        '.Fields("UPDTIME").Value = GetDTString(strUpdTime)
        '.Fields("UPDEN").Value = IIf(strUpdName <> "", strUpdName, GetFieldValue(rsSource, "EntryEn"))
        '.Fields("CMCODE").Value = GetFieldValue(rsSource, "CMCODE")
        '.Fields("CMNAME").Value = GetFieldValue(rsSource, "CMNAME")
        '.Fields("MANUALNO").Value = GetFieldValue(rsSource, "MANUALNO")
        '.Fields("UCCode").Value = Null
        '.Fields("UCName").Value = Null
        '.Fields("STCODE").Value = GetFieldValue(rsSource, "STCODE")
        '.Fields("STNAME").Value = GetFieldValue(rsSource, "STNAME")
        '.Fields("Note").Value = GetFieldValue(rsSource, "Note")
        '.Fields("ServiceType").Value = GetFieldValue(rsSource, "ServiceType")
        '.Fields("CancelFlag") = Val(rsSource("CancelFlag") & "")
        '.Fields("CancelCode") = NoZero(rsSource("CancelCode"))
        '.Fields("CancelName") = NoZero(rsSource("CancelName"))

        'On Error Resume Next
        ''SO077 沒這些欄位 95/02/06 Jacky
        '.Fields("BankCode") = NoZero(rsSource("BankCode"))
        '.Fields("BankName") = NoZero(rsSource("BankName"))
        '.Fields("AccountNo") = NoZero(rsSource("AccountNo"))
        '.Fields("AuthorizeNo") = NoZero(rsSource("AuthorizeNo"))
        '.Fields("AdjustFlag") = Val(rsSource("AdjustFlag") & "")
        '.Fields("NextPeriod") = NoZero(rsSource("NextPeriod"))
        '.Fields("NextAmt") = Val(rsSource("NextAmt") & "")
        '.Fields("InvSeqNo") = NoZero(rsSource("InvSeqNo") & "")

        ''如實收日期由無值變為有值
        'If IsNull(.Fields("RealDate").OriginalValue) And (Not IsNull(.Fields("RealDate"))) And IsNull(.Fields("FirstTime")) Then
        '    .Fields("FirstTime") = .Fields("UpdTime")
        'End If





    End Function
    Friend Function InsertTmpData(ByVal EntryType As Integer) As String
        Dim tableName As String = "SO074"
        Dim aRet As String = String.Empty
        Dim aFieldName As String = String.Empty
        Dim aFieldValue As String = String.Empty
        If EntryType = 1 Then
            tableName = "SO077"
        End If
        aFieldName = "BillNo,Item,Custid,CustName,CitemCode,CitemName, " & _
                            "MediaBillNo,PrtSNo,ShouldAmt,ShouldDate," & _
                            "RealAmt,ManualNo,RealDate,RealPeriod,RealStartDate,RealStopDate," & _
                            "EntryEn,Note,CMCode,CMName,ClctEn,ClctName," & _
                            "PTCode,PTName,RcdRowId,EntryNO,StCode,STName," & _
                            "ServiceType,CompCode,CancelFlag,CancelCode," & _
                            "CancelName,BankCode,BankName,AccountNo,AuthorizeNo," & _
                            "AdjustFlag,NextPeriod,NextAmt,FaciSeqNo,InvSeqNo," & _
                            "SUCCode,SUCName"
        For i As Int32 = 0 To aFieldName.Split(",").Count - 1
            If i = 0 Then
                aFieldValue = "{0}0"
            Else
                aFieldValue = aFieldValue & ",{0}" & i
            End If
        Next
        aRet = String.Format("Insert into {0} ( {1} ) Values ({2})", tableName, aFieldName, aFieldValue)
        aRet = String.Format(aRet, Sign)
        Return aRet
    End Function
    Friend Function IsPayOKOrCancel() As String
        Return String.Format("Select Count(1) cnt From SO033 " & _
                                " Where BillNo={0}0 " & _
                                " And UCCode Is Not NULL " & _
                                " And CancelFlag <> 1 " & _
                                " And UCCode Not In (Select CodeNo From CD013 Where RefNo In (7,8) )", Sign)
    End Function
    Friend Function chkREF3(ByVal BillLen As String) As String
        Select Case BillLen
            Case 11
                Return String.Format("Select count(*) from CD013 A Where ( RefNo in (3,7,8) or payOK = 1 ) And CodeNo in ( " & _
                                     "Select UCCode From SO033 Where MediaBillNo = {0}0 " & _
                                        " And CompCode = {0}1 " & _
                                        "And nvl(CancelFlag,0) = 0 And Uccode is not null )", Sign)
            Case 12
                Return String.Format("Select count(*) from CD013 A Where ( RefNo in (3,7,8) or payOK = 1 ) And CodeNo in ( " & _
                                   "Select UCCode From SO033 Where PrtSNo = {0}0 " & _
                                    " And CompCode = {0}1 " & _
                                      "And nvl(CancelFlag,0) = 0 And Uccode is not null )", Sign)
            Case Else
                Return String.Format("Select count(*) from CD013 A Where ( RefNo in (3,7,8) or payOK = 1 ) And CodeNo in ( " & _
                                   "Select UCCode From SO033 Where BillNo = {0}0 " & _
                                    " And CompCode = {0}1 " & _
                                      "And nvl(CancelFlag,0) = 0 And Uccode is not null )", Sign)
        End Select
    End Function
    Friend Overridable Function GetSO033Data(ByVal BillLen As Integer) As String
        Dim result As String = Nothing
        Select Case BillLen
            Case 11
                'Return String.Format("Select A.*,B.CustName,A.RowId,C.CustStatusName,Nvl(C.CustStatusCode,0) CustStatusCode " &
                '                     " From SO033 A ,SO001 B, SO002 C" &
                '                     " Where A.UCCode Not In (Select CodeNo From CD013 Where RefNo in (3,7,8) Or PayOk = 1) " &
                '                     " AND A.UCCODE IS NOT NULL " &
                '                     " AND A.CUSTID = B.CUSTID(+) " &
                '                     " AND A.COMPCODE = B.COMPCODE(+)1 " &
                '                     " AND A.COMPCODE = {0}0 " &
                '                     " AND B.CUSTID = C.CUSTID " &
                '                     " AND A.SERVICETYPE = C.SERVICETYPE(+) " &
                '                     " AND NVL(A.CancelFlag,0) = 0  " &
                '                     " AND MediaBillNo = {0}1", Sign)

                result = String.Format("Select A.*,B.CustName,A.RowId AS CTID ,C.CustStatusName,Nvl(C.CustStatusCode,0) CustStatusCode " &
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
                'result = String.Format("Select A.*,B.CustName,A.RowId,C.CustStatusName,Nvl(C.CustStatusCode,0) CustStatusCode" &
                '                     " From SO033 A ,SO001 B,SO002 C" &
                '                     " Where A.UCCode Not In (Select CodeNo From CD013 Where RefNo in (3,7,8) Or PayOk = 1) " &
                '                     " AND A.UCCODE IS NOT NULL " &
                '                     " AND A.CUSTID = B.CUSTID(+) " &
                '                     " AND A.COMPCODE = B.COMPCODE(+) " &
                '                     " AND B.CUSTID = C.CUSTID " &
                '                     " AND A.SERVICETYPE = C.SERVICETYPE(+) " &
                '                     " AND A.COMPCODE = {0}0 " &
                '                     " AND NVL(A.CancelFlag,0) = 0  " &
                '                    " AND PrtSNo  = {0}1", Sign)
                result = String.Format("Select A.*,B.CustName,A.RowId AS CTID ,C.CustStatusName,Nvl(C.CustStatusCode,0) CustStatusCode" &
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

                'result = String.Format("Select A.*,B.CustName,A.RowId,C.CustStatusName,Nvl(C.CustStatusCode,0) CustStatusCode " &
                '                     " From SO033 A ,SO001 B,SO002 C" &
                '                     " Where A.UCCode Not In (Select CodeNo From CD013 Where RefNo in (3,7,8) Or PayOk = 1) " &
                '                     " AND A.UCCODE IS NOT NULL " &
                '                     " AND A.CUSTID = B.CUSTID(+) " &
                '                     " AND A.COMPCODE = B.COMPCODE(+) " &
                '                     " AND B.CUSTID = C.CUSTID " &
                '                     " AND A.SERVICETYPE = C.SERVICETYPE(+) " &
                '                     " AND NVL(A.CancelFlag,0) = 0  " &
                '                     " AND A.COMPCODE = {0}0 " &
                '                    " AND BillNo  = {0}1", Sign)

                result = String.Format("Select A.*,B.CustName,A.RowId AS CTID,C.CustStatusName,Nvl(C.CustStatusCode,0) CustStatusCode " &
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
