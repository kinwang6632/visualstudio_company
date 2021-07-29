Imports CableSoft.BLL.Utility

Public Class CounterPayDAL
    Inherits CableSoft.BLL.Utility.DALBasic
    Implements IDisposable

    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)

    End Sub

#Region "SO3318A"

    ''' <summary>
    ''' 取得公司別
    ''' </summary>
    ''' <param name="GroupId">權限群組代碼</param>
    ''' <param name="LoginId">登入人員</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function GetCompCode(ByVal GroupId As String, ByVal LoginId As String) As String
        Dim strSQL As String
        'If GroupId = "0" Then
        '    strSQL = String.Format("Select A.CodeNo,A.Description,A.SOCompCode,A.UsePG From CD039 A Order By A.CodeNo ", Sign)
        'Else
        strSQL = String.Format("Select A.CodeNo,A.Description,A.SOCompCode,A.UsePG FROM CD039 A,SO026 B " & _
                                                    " Where Instr(','||B.CompStr||',' , ','||A.CodeNo||',')>0 And B.UserId='{0}' " & _
                                                    "Group by A.CodeNo,A.Description,A.SOCompCode,A.UsePG Order by A.CodeNo", LoginId)
        'End If
        Return strSQL
    End Function
    ''' <summary>
    ''' 取得可選收費方式
    ''' </summary>
    ''' <returns>查詢可選收費方式</returns>
    ''' <remarks></remarks>
    Friend Function GetCMCode() As String
        Dim strSQL As String = "Select CodeNo,Description,RefNo,Kind From CD031 Where StopFlag=0 Order by CodeNo"
        Return strSQL
    End Function
    ''' <summary>
    ''' 取得可選付款種類
    ''' </summary>
    ''' <returns>查詢可選付款種類</returns>
    ''' <remarks></remarks>
    Friend Function GetPTCode() As String
        Dim strSQL As String = "Select CodeNo,Description,RefNo From CD032 Where StopFlag=0 Order by CodeNo"
        Return strSQL
    End Function
    ''' <summary>
    ''' 取得可選收費人員
    ''' </summary>
    ''' <returns>查詢可選收費人員</returns>
    ''' <remarks></remarks>
    Friend Function GetClctEn() As String
        Dim strSQL As String = "Select EmpNo,EmpName From CM003 Where StopFlag=0 Order by EmpNo"
        Return strSQL
    End Function
    ''' <summary>
    ''' 取得發票開立設定檔 CD125
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function GetInvConSetting() As String
        Dim strSQL As String = String.Format("Select * From CD125 Where StopFlag=0 And SOCompCode={0}0 Order By ServiceType,Type desc", Sign)
        Return strSQL
    End Function
    ''' <summary>
    ''' 取得結帳日期SO062.TranDate
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function GetTranDate() As String
        Dim strSQL As String = String.Format("Select TranDate From SO062 Where CompCode={0}0 And Type=1 Order By TranDate Desc", Sign)
        Return strSQL
    End Function
    ''' <summary>
    ''' 取得客戶載具類別
    ''' </summary>
    ''' <returns>查詢客戶載具類別</returns>
    ''' <remarks></remarks>
    Friend Function GetCarrierType() As String
        Dim strSQL As String = "Select CodeNo,Description,RefNo From CD122 Where StopFlag=0 Order by CodeNo"
        Return strSQL
    End Function
    ''' <summary>
    ''' 取得信用卡別
    ''' </summary>
    ''' <returns>查詢信用卡別</returns>
    ''' <remarks></remarks>
    Friend Function GetCardCode() As String
        Dim strSQL As String = "Select CodeNo,Description,RefNo From CD037 Where StopFlag=0 Order by CodeNo"
        Return strSQL
    End Function

#End Region

#Region "SO3318B"
    '取得SO074A費用
    Friend Overridable Function GetChargeTmp(ByVal strWhere As String, intTotalAmt As Integer, strMaxEntryNo As String, intCountNo As Integer) As String
        Dim strField As String = ",C.SeqNo,C.GetDate,C.PrDate,C.InstDate,C.StopTime,C.ReInstTime,D.GuiNo,D.PreInvoice"
        strField = String.Format("{0},{1} as TotalAmt,'{2}' as MaxEntryNo,{3} as CountNo", strField, intTotalAmt, strMaxEntryNo, intCountNo)
        'Dim strSQL As String = String.Format("Select A.RowId,A.*,B.CustStatusName {0} " &
        '                                    "From SO074A A,SO002 B,SO004 C,SO033 D " &
        '                                    "Where A.CustId=B.CustId And A.FaciSeqNo=C.SeqNo(+) And A.CustId=C.CustId(+) " &
        '                                    "And A.ServiceType=B.ServiceType And A.Custid=D.Custid And A.BillNo=D.BillNo And A.Item=D.Item And {1} " &
        '                                    "And D.CancelFlag=0 Order by A.EntryNo Desc,A.BillNo,A.Item", strField, strWhere)
        Dim strSQL As String = String.Format("Select A.RowId CTID ,A.*,B.CustStatusName {0} " &
                                            "From SO074A A left join so004 C on A.FaciSeqNo=C.SeqNo AND A.CustId=C.CustId ,SO002 B,SO033 D " &
                                            "Where A.CustId=B.CustId " &
                                            "And A.ServiceType=B.ServiceType And A.Custid=D.Custid And A.BillNo=D.BillNo And A.Item=D.Item And {1} " &
                                            "And D.CancelFlag=0 Order by A.EntryNo Desc,A.BillNo,A.Item", strField, strWhere)
        Return strSQL
    End Function
    '畫面顯示總金額,總單數
    Friend Function GetShowCount(ByVal strWhere As String) As String
        Dim strSQL As String = String.Format("Select Nvl(Sum(A.RealAmt),0) TotalAmt,Max(EntryNo) MaxEntryNo,Count(distinct EntryNo || EntryEn) CountNo  " &
                                                                            "From SO074A A Where {0}", strWhere)
        Return strSQL
    End Function
    '單據編號取收費資料
    Friend Overridable Function GetRealCharge(ByVal CompCode As Integer, ByVal strWhere As String) As String
        Dim strSQL As String = String.Format("Select Rowid  ctid, A.*  From SO033 A Where A.CompCode={0} And {1} And A.UCCode is not null And A.CancelFlag=0", CompCode, strWhere)
        Return strSQL
    End Function
    Friend Function GetRealCharge2(ByVal CompCode As Integer, ByVal strWhere As String) As String
        'Dim strSQL As String = String.Format("Select A.*,B.RefNo,B.PayOK  From SO033 A,CD013 B Where A.UCCode=B.CodeNo
        ') And A.CompCode={0} And {1} And A.CancelFlag=0", CompCode, strWhere)
        Dim strSQL As String = String.Format("Select A.*,B.RefNo,B.PayOK  From SO033 A LEFT JOIN CD013 B ON A.UCCODE = B.CODENO  Where  A.CompCode={0} And {1} And A.CancelFlag=0", CompCode, strWhere)
        Return strSQL
    End Function

    '判別此單據是否已過帳！
    Friend Function GetTmpCharge(ByVal CompCode As Integer, ByVal strWhere As String) As String
        Dim strSQL As String = String.Format("Select A.*  From SO074A A Where A.CompCode={0} And {1} ", CompCode, strWhere)
        Return strSQL
    End Function
    '客戶狀態
    Friend Function GetCustStatus(ByVal Custid As Integer, ByVal ServiceType As String) As String
        Dim strSQL As String = String.Format("Select A.CustName,B.CustStatusCode,B.CustStatusName From SO001 A,SO002 B Where A.CustId=B.CustId And A.Custid={0} And B.ServiceType='{1}' ", Custid, ServiceType)
        Return strSQL
    End Function
    '取得SO074A Max EntryNo
    Friend Function GetMaxEntryNo(ByVal CompCode As Integer, ByVal strClctEn As String, ByVal strRealDate As String) As String
        Dim strSQL As String = String.Format("Select Nvl(Max(EntryNo),0) From SO074A Where CompCode={0} And ClctEn='{1}' " & _
                                                                            "And RealDate=To_Date('{2}','yyyy/MM/dd')", CompCode, strClctEn, strRealDate)
        Return strSQL
    End Function
    '取得CustName
    Friend Function GetCustName(ByVal CompCode As Integer, ByVal Custid As Integer) As String
        Dim strSQL As String = String.Format("Select CustName From SO001 Where CustId={0} And CompCode={1}", Custid, CompCode)
        Return strSQL
    End Function
    'CD013
    Friend Function GetUCCode(ByVal strWhere As String) As String
        Dim strSQL As String = String.Format("Select CodeNo,Description From CD013 Where StopFlag=0 And {0}", strWhere)
        Return strSQL
    End Function
    'SO033
    Friend Overridable Function GetSimple(ByVal CompCode As Integer, ByVal BillNo As String, ByVal Item As Integer) As String
        Dim strSQL As String = String.Format("Select A.RowID CTID ,0 as Type,A.*  From SO033 A Where A.CompCode={0} And A.BillNo='{1}' And A.Item={2}", CompCode, BillNo, Item)
        Return strSQL
    End Function
    'Delete SO074A
    Friend Function DeleteChargeTmp(ByVal strWhere As String) As String
        Dim strSQL As String
        strSQL = String.Format("Delete From SO074A A Where {0} ", strWhere)
        Return strSQL
    End Function
    ''' <summary>
    ''' 取得SO003原收費週期的預設值
    ''' </summary>
    ''' <param name="Custid"></param>
    ''' <param name="CitemCode"></param>
    ''' <param name="RcdRowId"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Overridable Function GetPeriodCycle(ByVal Custid As Integer, ByVal CitemCode As Integer, ByVal RcdRowId As String) As String
        Dim strSQL As String = String.Format("Select CMCode,CMName,PTCode,PTName FROM SO003 " &
                                                               "WHERE CustId={0} AND CitemCode={1} AND SeqNo in (Select SeqNo FROM SO033 WHERE Rowid='{2}')", Custid, CitemCode, RcdRowId)
        Return strSQL
    End Function
    ''' <summary>
    ''' 刪除登錄檔後,還原SO033收費設定資料
    ''' </summary>
    ''' <param name="SUCCode"></param>
    ''' <param name="SUCName"></param>
    ''' <param name="CMCode"></param>
    ''' <param name="CMName"></param>
    ''' <param name="PTCode"></param>
    ''' <param name="PTName"></param>
    ''' <param name="RcdRowid"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Overridable Function UpdRealCharge(ByVal SUCCode As String, ByVal SUCName As String, ByVal CMCode As String, ByVal CMName As String,
                                                                ByVal PTCode As String, ByVal PTName As String, ByVal RcdRowid As String) As String
        Dim strSQL As String = String.Format("Update SO033 Set RealDate= Null,RealAmt=0 " &
                                                                                        ",UCCode={0},UCName='{1}' " &
                                                                                        ",CMCode={2},CMName='{3}' " &
                                                                                        ",PTCode={4},PTName='{5}' " &
                                                                                        ",ClctEn=OldClctEn,ClctName=OldClctName" &
                                                                                        " Where RowId='{6}' And CancelFlag=0", Integer.Parse(SUCCode), SUCName, Integer.Parse(CMCode), CMName, Integer.Parse(PTCode), PTName, RcdRowid)
        Return strSQL
    End Function
    '取得SO074A
    Friend Function GetChargeTmp2(ByVal CompCode As Integer, ByVal strWhere As String) As String
        Dim strSQL As String = String.Format("Select A.* From SO074A A Where A.CompCode={0} And {1} Order By A.CustId,MediaBillNo,BillNo,Item", CompCode, strWhere)
        Return strSQL
    End Function
    '取得SO033
    Friend Overridable Function GetSimple2() As String
        'Dim strSQL As String = String.Format("Select A.RowID CTID,A.* From SO033 A Where A.CompCode={0}0 And A.Rowid={0}1", Sign)
        Dim strSQL As String = String.Format("Select  A.* From SO033 A Where A.CompCode={0}0 And A.BILLNO={0}1 AND ITEM = {0}2", Sign)
        Return strSQL
    End Function
    Friend Function GetCustomerData() As String
        Dim strSQL As String
        strSQL = String.Format("Select ClctAreaCode,AddrNo,StrtCode,MduId,ServCode,ClassCode,AreaCode,NodeNo From SO033  Where BillNo = {0}0", Sign)
        Return strSQL
    End Function
    Friend Function GetFaciSNo() As String
        Dim strSQL As String
        strSQL = String.Format("Select FaciSno From SO004 Where SeqNo = {0}0", Sign)
        Return strSQL
    End Function
    'SO043.Para41
    Friend Function GetSO043() As String
        Dim strSQL As String
        strSQL = String.Format("Select Servicetype,Para41 From SO043")
        Return strSQL
    End Function
    'SO003
    Friend Function GetPeriodData(ByVal CompCode As Integer, ByVal strWhere As String) As String
        Dim strSQL As String
        strSQL = String.Format("Select B.* From SO033 A,SO003 B Where A.CustId=B.CustId And A.CompCode=B.CompCode " &
                                                    "And A.ServiceType=B.ServiceType And A.CitemCode=B.CitemCode And A.FaciSeqNo=B.FaciSeqNo " &
                                                    "And A.CompCode={0} And {1} ", CompCode, strWhere)
        Return strSQL
    End Function
    '判斷是否有異業資料
    Friend Function GetDiffBill(ByVal CardBillNo As String) As String
        Dim strSQL As String = String.Format("Select * From SO315B Where CardBillNo='{0}' And CancelFlag=0", CardBillNo)
        Return strSQL
    End Function
    '判斷是否已入實收或作廢
    Friend Function ChkChargeData(ByVal strWhere As String) As String
        Dim strSQL As String = String.Format("Select * From SO033 {0} ", strWhere)
        Return strSQL
    End Function

#End Region

#Region "SO3318B1"

    Friend Function UpdCharge(ByVal TableName As String, PTCode As Integer, PTName As String, BillNo As String, Item As Integer) As String
        Dim strSQL As String = String.Format("Update {0} Set " & _
                                        " PTCode={1},PTName='{2}' " & _
                                        " Where BillNo='{3}' And Item={4} ", TableName, PTCode, PTName, BillNo, Item)
        Return strSQL
    End Function


#End Region

#Region "SO3318B2"

    Friend Function UpdCustData(ByVal TableName As String, ByVal Value As String, ByVal Where As String) As String
        Dim strSQL As String = String.Format("Update {0} Set {1} {2} ", TableName, Value, Where)
        Return strSQL
    End Function

    Friend Function ChkCarrierType(ByVal CarrierTypeCode As String) As String
        Dim strSQL As String = String.Format("Select RefNo From CD122 Where CodeNo='{0}' And RefNo=1 And StopFlag=0", CarrierTypeCode)
        Return strSQL
    End Function

#End Region

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
