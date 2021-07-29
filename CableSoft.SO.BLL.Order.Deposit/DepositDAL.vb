
Public Class DepositDAL

    Inherits CableSoft.BLL.Utility.DALBasic
    Implements IDisposable
    Public Sub New()
    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    ''' <summary>
    ''' 取得查詢可選付款種類SQL[CD078A2,CD032] ; SQL參數:BPCode,CitemCode
    ''' </summary>
    ''' <returns>回傳查詢可選付款種類SQL</returns>
    ''' <remarks>SQL參數:BPCode,CitemCode</remarks>
    Friend Function QueryPTCode() As String
        '不需要ServiceTypeㄇ? => And (ServiceType={0}0 Or ServiceType Is Null)
        Return String.Format("Select B.CodeNo,B.Description" &
                             " From CD078A2 A,CD032 B" &
                             " Where B.StopFlag<>1" &
                             " And A.BPCode={0}0 And A.CitemCode={0}1" &
                             " And A.PTCode=B.CodeNo" &
                             " Order By B.CodeNo", Sign)
    End Function

    ''' <summary>
    ''' 取得選擇付款種類SQL[CD078A2,CD019] ; SQL參數:BPCode,CitemCode,PTCode
    ''' </summary>
    ''' <returns>回傳選擇付款種類SQL</returns>
    ''' <remarks>SQL參數:BPCode,CitemCode,PTCode</remarks>
    Friend Function ChoosePTCode() As String
        '不需要ServiceTypeㄇ? => And (ServiceType={0}0 Or ServiceType Is Null)
        Return String.Format("Select A.DepositAmt,B.CodeNo,B.Description" &
                             " From CD078A2 A,CD019 B" &
                             " Where B.StopFlag<>1" &
                             " And A.BPCode={0}0" &
                             " And A.CitemCode={0}1" &
                             " And A.PTCode={0}2" &
                             " And A.DepositCode=B.CodeNo" &
                             " Order By CodeNo", Sign)
    End Function

    ''' <summary>
    ''' 取得查詢可選收費方式SQL[CD031]
    ''' </summary>
    ''' <returns>回傳查詢可選收費方式SOL</returns>
    ''' <remarks></remarks>
    Friend Function QueryCMCode() As String
        '不需要ServiceTypeㄇ? => And (ServiceType={0}0 Or ServiceType Is Null)
        'Order By Description ??
        Return String.Format("Select CodeNo,Description,RefNo From CD031" &
                             " Where StopFlag<>1" &
                             " Order By CodeNo", Sign)
    End Function

    'Execute

#Region "IDisposable Support"
    Private _Disposed As Boolean


    ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not _Disposed Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
            End If

            ' TODO: 釋放 Unmanaged 資源 (Unmanaged 物件) 並覆寫下面的 Finalize()。
            ' TODO: 將大型欄位設定為 null。
        End If
        _Disposed = True
    End Sub

    ' TODO: 只有當上面的 Dispose(ByVal disposing As Boolean) 有可釋放 Unmanaged 資源的程式碼時，才覆寫 Finalize()。
    'Protected Overrides Sub Finalize()
    '    ' 請勿變更此程式碼。在上面的 Dispose(ByVal disposing As Boolean) 中輸入清除程式碼。
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        ' 請勿變更此程式碼。在以上的 Dispose 置入清除程式碼 (ByVal 視為布林值處置)。
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
