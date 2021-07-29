Public Class FacilityDAL
    Inherits CableSoft.BLL.Utility.DALBasic
    Implements IDisposable
    Public Sub New()
    End Sub

    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    ''' <summary>
    ''' 取得服務類別SQL[CD046]
    ''' </summary>
    ''' <returns>回傳取得服務類別SQL</returns>
    ''' <remarks></remarks>
    Friend Function GetServiceTypeCode() As String
        Return "Select CodeNo,Description,RefNo From CD046 Order By CodeNo"
    End Function

    ''' <summary>
    ''' 取得可選設備項目SQL[CD022] ; SQL參數:ServiceType
    ''' </summary>
    ''' <returns>回傳可選設備項目SQL</returns>
    ''' <remarks>SQL參數:ServiceType</remarks>
    Friend Function QueryFaciCode() As String
        Return String.Format("Select CodeNo,Description,RefNo,DefBuyCode,DefBuyName From CD022" &
                             " Where StopFlag<>1" &
                             " And (ServiceType={0}0 Or ServiceType Is Null)" &
                             " Order By CodeNo", Sign)
    End Function

    ''' <summary>
    ''' 取得可選買賣方式SQL[CD034] ; SQL參數:ServiceType
    ''' </summary>
    ''' <returns>回傳可選買賣方式SQL</returns>
    ''' <remarks>SQL參數:ServiceType</remarks>
    Friend Function QueryBuyCode() As String
        Return String.Format("Select CodeNo,Description,RefNo From CD034" &
                             " Where StopFlag<>1" &
                             " And (ServiceType={0}0 Or ServiceType Is Null)" &
                             " Order By CodeNo", Sign)


    End Function

    ''' <summary>
    ''' 取得可選派工類別SQL[CD005] ; SQL參數:ServiceType
    ''' </summary>
    ''' <returns>回傳可選派工類別SQL</returns>
    ''' <remarks>SQL參數:ServiceType</remarks>
    Friend Function QueryWipCode() As String
        Return String.Format("Select CodeNo,Description,RefNo From CD005" &
                             " Where StopFlag<>1" &
                             " And (ServiceType={0}0 Or ServiceType Is Null)" &
                             " And Nvl(RefNo,0)<>-99" &
                             " Order By CodeNo", Sign)
    End Function


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
