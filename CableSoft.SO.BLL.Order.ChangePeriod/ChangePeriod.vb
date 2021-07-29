
Imports System.Data.Common
Imports CableSoft.BLL.Utility
Imports CableSoft.Utility.DataAccess

Imports Lang = CableSoft.SO.BLL.Order.ChangePeriod.ChangePeriodLanguage

Public Class ChangePeriod
    Inherits BLLBasic
    Implements IDisposable

    Private _DAL As New ChangePeriodDALMultiDB(LoginInfo.Provider)
    'Private _DAO As New DAO(LoginInfo.Provider, LoginInfo.ConnectionString)
    Private _GetSchema As Boolean
    Private _Disposed As Boolean ' 偵測多餘的呼叫

    Public Sub New()
    End Sub

    Public Sub New(ByVal LoginInfo As LoginInfo)
        MyBase.New(LoginInfo)
    End Sub

    Public Sub New(ByVal LoginInfo As LoginInfo, ByVal DBConnection As DbConnection)
        MyBase.New(LoginInfo, DBConnection)
    End Sub

    Public Sub New(ByVal LoginInfo As LoginInfo, ByVal DAO As DAO)
        MyBase.New(LoginInfo, DAO)
    End Sub

    Protected Overrides Sub Finalize()
        Try
            Dispose(False)
        Finally
            MyBase.Finalize()
        End Try
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overridable Sub Dispose(ByVal disposing As Boolean) ' IDisposable
        If Not _Disposed Then
            If disposing Then
                ' TODO: 釋放其他狀態 (Managed 物件)。
                Try
                    If _DAL IsNot Nothing Then
                        _DAL.Dispose()
                    End If
                    If (Me.MustDispose) AndAlso (Me.DAO IsNot Nothing) Then
                        DAO.Dispose()
                    End If
                Catch ex As Exception
                    Throw ex
                End Try
            End If
            ' TODO: 釋放您自己的狀態 (Unmanaged 物件) 或將大型欄位設定為 null。
        End If
        _Disposed = True
    End Sub

    ''' <summary>
    ''' 取得查詢可選繳別
    ''' </summary>
    ''' <param name="BPCode">優惠組合</param>
    ''' <param name="CitemCode">收費項目</param>
    ''' <returns>回傳取得查詢可選繳別</returns>
    ''' <remarks></remarks>
    Public Function QueryBPPeriod(ByVal BPCode As String,
                                  ByVal CitemCode As Integer) As DataTable
        QueryBPPeriod = Nothing
        Try
            If BPCode IsNot Nothing And CitemCode > 0 Then
                Return DAO.ExecQry(_DAL.QueryBPPeriod,
                                   New Object() {BPCode, CitemCode},
                                   _GetSchema)
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 取得查詢繳別細項
    ''' </summary>
    ''' <param name="StepNo">方案編號</param>
    ''' <returns>回傳取得查詢繳別細項</returns>
    ''' <remarks></remarks>
    Public Function QueryBPPeriodLevel(ByVal StepNo As Integer) As DataTable
        QueryBPPeriodLevel = Nothing
        Try
            Dim objArray(11) As Object
            For i As Int16 = 0 To 11
                objArray(i) = StepNo
            Next

            Return DAO.ExecQry(_DAL.QueryBPPeriodLevel,
                               objArray,
                               _GetSchema)

        Catch ex As Exception
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 執行下收切齊
    ''' </summary>
    ''' <param name="Period">期數</param>
    ''' <param name="Charge">訂單收費</param>
    ''' <returns>訂單收費</returns>
    ''' <remarks></remarks>
    Public Function Execute(ByVal Period As Integer,
                            ByVal Charge As DataTable) As DataTable
        Dim tCharge As DataTable = Nothing
        Try
            If Charge IsNot Nothing Then
                tCharge = Charge.Clone
                If Charge.Rows.Count > 0 Then
                    '(1)Loop Charge 
                    For Each dr As DataRow In Charge.Rows
                        If (dr.IsNull("Period") OrElse dr.Item("Period") = 0) OrElse
                                ((Not dr.IsNull("AssignProd")) AndAlso dr.Item("AssignProd") <> 0) Then
                            Dim tdr As DataRow = tCharge.NewRow
                            tdr.ItemArray = dr.ItemArray
                            tCharge.Rows.Add(tdr)
                        Else
                            If dr.IsNull("StepNo") OrElse dr.Item("StepNo") = 0 Then
                                Dim tdr As DataRow = tCharge.NewRow
                                tdr.ItemArray = dr.ItemArray
                                tdr.Item("Period") = Period
                                tdr.Item("Amount") = dr.Item("Amount") / dr.Item("Period") * Period
                                tCharge.Rows.Add(tdr)
                            Else
                                If tCharge.Select("StepNo = " & dr.Item("StepNo")).Count = 0 Then
                                    '(2)如CitemCode為週期(CD019.PeriodFlag = 1)則做以下動作:
                                    'A.其他欄位同原欄位。
                                    'B.先將同StepNo 的資料刪除。
                                    'C.BPCode = Select...
                                    'Period,CitemCode,BPCode,StepNo
                                    Dim chgBPCode As String = dr("BPCode").ToString
                                    Dim chgCitemCode As String = dr("CitemCode").ToString
                                    Dim chgPeriod As Int32 = Period
                                    Dim chgStepNo As Integer = Integer.Parse(dr("StepNo").ToString)
                                    Dim TotalPeriod As Integer = 0
                                    Dim dt As DataTable = DAO.ExecQry(_DAL.QueryBPCode, _
                                                                      New Object() {chgCitemCode, chgBPCode, chgPeriod, chgStepNo})
                                    '檢核如無該期數則將原本的再回填回Charge
                                    If dt.Rows.Count = 0 Then
                                        Dim drNoStepNos As DataRow() = Charge.Select("StepNo = " & dr.Item("StepNo"))
                                        For Each drNoStep As DataRow In drNoStepNos
                                            Dim tdr As DataRow = tCharge.NewRow
                                            tdr.ItemArray = drNoStep.ItemArray
                                            tCharge.Rows.Add(tdr)
                                        Next
                                    Else
                                        'D.	Loop BPCode(每Loop 一筆則新增一筆資料)
                                        '1.TotalPeriod = TotalPeriod + BPCode.Period
                                        '2.	Period = BPCode.Period , Amount = BPCode.MonthAmt * BPCode.Period
                                        '3.	當TotalPeriod >= “Period” 則結束Loop                                       
                                        Dim StepItem As Integer = 0
                                        Do While Period > TotalPeriod
                                            If StepItem > dt.Rows.Count - 1 Then
                                                Exit Do
                                            End If
                                            Dim drCitemStep As DataRow = dt.Rows(StepItem)
                                            Dim tdr As DataRow = tCharge.NewRow
                                            tdr.ItemArray = dr.ItemArray
                                            Dim aMon As Int32 = 0
                                            If Not DBNull.Value.Equals(drCitemStep.Item("Mon")) Then
                                                aMon = Int32.Parse(drCitemStep.Item("Mon"))
                                            End If
                                            '2014/10/22 Jacky 增加判斷如Mon <=0 則不再產生收費資料
                                            If aMon <= 0 Then
                                                Exit Do
                                            End If
                                            If aMon < Period - TotalPeriod Then
                                                tdr.Item("Period") = drCitemStep.Item("Mon")
                                            Else
                                                tdr.Item("Period") = Period - TotalPeriod
                                            End If
                                            If (Not DBNull.Value.Equals(tdr.Item("Period"))) AndAlso
                                                (Not DBNull.Value.Equals(drCitemStep.Item("MonthAmt"))) Then
                                                tdr.Item("Amount") = Decimal.Round(tdr.Item("Period") * drCitemStep.Item("MonthAmt"))
                                            Else
                                                tdr.Item("Amount") = 0
                                            End If

                                            tdr.Item("StepNo") = drCitemStep.Item("StepNo")
                                            tdr.Item("LinkKey") = drCitemStep.Item("LinkKey")
                                            If Not DBNull.Value.Equals(tdr.Item("Period")) Then
                                                TotalPeriod += Integer.Parse(tdr.Item("Period"))
                                            End If

                                            StepItem += 1
                                            '增加長繳別判斷 By Kin 2014/10/06 For Jacky
                                            'If lngLONGPAYflag = 1 Then
                                            '    If Not GetRS(rsTmp, "Select A.Period,B.MonthAmt From " & GetOwner & "CD078A1 A," & GetOwner & "CD078A B Where A.BPCode = B.BPCode And A.CitemCode = B.CitemCode And A.CitemCode = " & strCitemCode & " And A.Period = " & lngPeriod & " And A.BPCode = '" & strBPCode & "'") Then Exit Function
                                            '    If Not rsTmp.EOF Then
                                            '        strNextPeriod = rsTmp("Period") & ""
                                            '        strNextAmt = rsTmp("MonthAmt") * Val(strNextPeriod)
                                            '    End If
                                            Using dbr As DbDataReader = DAO.ExecDtRdr(_DAL.QueryCD078A1,
                                                                                      New Object() {chgBPCode, chgCitemCode})
                                                While dbr.Read
                                                    If DBNull.Value.Equals(tdr.Item("Period")) Then
                                                        tdr.Item("NextPeriod") = 0
                                                    Else
                                                        tdr.Item("NextPeriod") = tdr.Item("Period")
                                                    End If
                                                    tdr.Item("NextAmt") = dbr.Item("MonthAmt") * tdr.Item("NextPeriod")
                                                End While

                                                dbr.Close()
                                                dbr.Dispose()
                                            End Using
                                            tCharge.Rows.Add(tdr)
                                        Loop

                                    End If
                                End If
                            End If
                        End If
                    Next
                    tCharge.AcceptChanges()
                    Using ChooseProduct As New CableSoft.SO.BLL.Order.ChooseProduct.ChooseProduct(Me.LoginInfo, Me.DAO)
                        ChooseProduct.LongPayChangeNextPeriod(tCharge)
                    End Using
                End If
            End If
        Catch ex As Exception
            Throw
        End Try
        Return tCharge
    End Function

End Class
