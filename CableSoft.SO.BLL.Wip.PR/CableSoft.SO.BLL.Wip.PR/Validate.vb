Imports System.Data.Common
Imports CableSoft.BLL.Utility

Public Class Validate
    Inherits BLLBasic
    Implements IDisposable
    Private _DAL As New PRDAL(Me.LoginInfo.Provider)

    Public Sub New()
    End Sub

    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo)
        MyBase.New(LoginInfo)
    End Sub
    ''' <summary>
    ''' 檢查停拆機類別是否正常可派(CheckCanPR)
    ''' </summary>
    ''' <param name="PRCode">停拆機類別</param>
    ''' <param name="PRName">停拆機類別名稱</param>
    ''' <param name="PRRefNo">停拆機參考號</param>
    ''' <param name="Interdepend">服務依存</param>
    ''' <param name="CustStatusCode">客戶狀態</param>
    ''' <param name="WipCode3">派工類別3</param>
    ''' <param name="CompCode">公司別</param>
    ''' <param name="ServiceType">服務別</param>
    ''' <returns>Boolean</returns>
    ''' <remarks></remarks>
    Public Function ChkCanResv(ByVal PRCode As Int32, ByVal PRName As String,
                               ByVal PRRefNo As Int32, ByVal Interdepend As Int32,
                               ByVal CustStatusCode As Int32, ByVal WipCode3 As String,
                               ByVal CompCode As Int32, ByVal ServiceType As String) As Boolean


    End Function
    ''' <summary>
    ''' 檢查預約時段是否可派工
    ''' </summary>
    ''' <param name="ServCode">服務區</param>
    ''' <param name="WipCode"></param>
    ''' <param name="MCode">裝機類別名稱</param>
    ''' <param name="ServiceType">服務別</param>
    ''' <param name="ResvTime">預約時間</param>
    ''' <param name="AcceptTime">受理時間</param>
    ''' <param name="OldResvTime">舊預約時間</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ChkCanResv(ByVal ServCode As String, ByVal WipCode As Int32,
                            ByVal MCode As String, ByVal ServiceType As String,
                            ByVal ResvTime As Date,
                            ByVal AcceptTime As Date, ByVal OldResvTime As Date,
                            ByVal Resvdatebefore As Int32) As RIAResult

        Dim obj As New CableSoft.SO.BLL.Wip.Utility.Validate(Me.LoginInfo, DAO)
        Try
            Return obj.ChkCanResv(BLL.Utility.InvoiceType.Maintain,
                                  WipCode, ServCode, MCode, ServiceType,
                                  ResvTime, AcceptTime, OldResvTime, Resvdatebefore)
        Finally
            obj.Dispose()
        End Try
    End Function



End Class
