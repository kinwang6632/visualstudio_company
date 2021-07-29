Imports CableSoft.BLL.Utility
Public Class CreditDAL
    Inherits DALBasic
    Implements IDisposable
    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Function QueryCitemName() As String
        Return String.Format("Select Description From CD019 Where CodeNo = {0}0", Sign)
    End Function
    Friend Function QuerySO041() As String
        Return "Select bbTranspriority From SO041"
    End Function
    Friend Function QueryAddrNo() As String
        Return String.Format("Select StrtCode,AreaCode,ClctEn,ClctName From SO014 " & _
                             " Where AddrNo = {0}0 And CompCode = {0}1", Sign)
    End Function
    Friend Function QueryCustId() As String
        Return String.Format("Select  InstAddrNO,MduId,ServCode,ClctAreaCode,ClassCode1 " & _
                                    " From SO001 Where CustId = {0}0 And CompCode ={0}1", Sign)
    End Function
    Friend Function QueryUCCode() As String
        Return "Select CodeNo,Description From CD013 Where RefNo = 3"
    End Function
    Friend Function InsSO033BBTrans() As String
        Return String.Format("Insert Into SO033BBTRANS (bbAccountID,SeqNo,transSavepoint," & _
                                            "transbonus,transdate,transStoreCode,Rate,Mallpoint,Mallbonus,UpdTime,UpdEn) " & _
                                            " Values ({0}0,{0}1,{0}2," & _
                                            " {0}3,sysdate,{0}4,{0}5,{0}6,{0}7,sysdate,{0}8)", Sign)
    End Function
    Friend Function GetTransseqno() As String

        Return "Select to_char(sysdate,'yyyymmdd') || Lpad( S_SO033BBTRANS.NEXTVAL,12,0) FROM DUAL"


    End Function
    Friend Function QueryCD129() As String
        Return String.Format("Select CodeNo,Description,Nvl(RefNo,0) RefNo,Nvl(Rate,1) Rate," & _
                                            "CitemCode,Condition " & _
                                            " From CD129 " & _
                                            " Where Nvl(StopFlag,0) = 0 And CodeNo = {0}0 And Property = 1 ", Sign)
    End Function
    Friend Function QuerySO193()
        Dim result As String = Nothing
        result = String.Format("Select Nvl(A.TotalSavePoint,0) TotalSavePoint,Nvl(A.TotalBonus,0) TotalBonus," & _
                               " B.bbAccountID,B.SeqNo,B.Savedate,B.SaveBillno," & _
                               "B.SavePlanCode,B.CitemCode,B.CitemName,Nvl(B.Savepoint,0) As Savepoint,Nvl(B.bonus,0) As bonus," & _
                                "B.bonusstopdate,Nvl(B.UsedSavepoint,0) As UsedSavepoint," & _
                                " Nvl(B.Usedbonus,0) As Usedbonus ,B.closeBillno," & _
                                " B.MinusType,B.CreditTypeCode,Nvl(B.StopFlag,0) As StopFlag, " & _
                               "B.RowId,C.Description SavePlanName  " & _
                        " From SO004J A,SO193 B, CD130 C " & _
                        " Where A.bbAccountID = B.bbAccountID And A.bbAccountID ={0}0 " & _
                        " And B.closeBillno is Null " & _
                        " And Nvl(B.StopFlag,0) = 0 " & _
                        " And Nvl(B.bonusstopdate,To_Date('29991231235959','YYYYMMDDHH24MISS') )>= " & _
                            " To_Date('" & Date.Now.ToString("yyyyMMddHHmmss") & "','YYYYMMDDHH24MISS')" & _
                        " And B.SavePlanCode = C.CodeNo(+)  Order by Nvl(B.Savedate,To_Date('19110101','yyyymmdd'))", Sign)
        Return result
    End Function
    Friend Function InsSO033() As String
        Dim result As String = Nothing
        result = String.Format("Insert Into SO033 (FaciSeqNo,FaciSNo,CustID,BillNo,Item,CitemCode," & _
                               "CitemName,ShouldDate,RealAmt,CMCode,CMName,PTCode,PTName, " & _
                               "CreateTime,CreateEn,CompCode,AddrNo,StrtCode,MduId,ServCode,ClctAreaCode," & _
                               "AreaCode,OldClctEn,OldClctName,ClctEn,ClctName,CancelFlag,ClassCode," & _
                               "UPDTIME,UPDEN,UCCode,UCName,ServiceType,Quantity,OldAmt,OldPeriod," & _
                               "RealPeriod,Amt,ShouldAmt ) Values ({0}0,{0}1,{0}2,{0}3,1,{0}4," & _
                               "{0}5,{0}6),{0}7,1,{0}8,1,{0}9," & _
                               "{0}10,{0}11,{0}12,{0}13,{0}14,{0}15,{0}16,{0}17," & _
                               "{0}18,{0}19,{0}20,{0}21,{0}22,0,{0}23, " & _
                               "{0}24,{0}25,{0}26,{0}27,{0}28,0,{0}29,0," & _
                               "0,{0}30,{0}31 )", Sign)
        Return result
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
