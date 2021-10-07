Imports System.Data.Common
Imports CableSoft.BLL.Utility

Public Class ChangeFaci
    Inherits BLLBasic

    Implements IDisposable
    Private _DAL As New ChangeFaciDALMultiDB(Me.LoginInfo.Provider)
    Private Const fWip_Wip As String = "Wip"
    Private Const fWip_Facility As String = "Facility"
    Private Const fWip_PRFacility As String = "PRFacility"
    Private Const fWip_Charge As String = "Charge"
    Private Const fWip_ChangeFacility As String = "ChangeFacility"
    Private fNowDate As Date = Date.Now
    Private Language As New CableSoft.BLL.Language.SO61.ChangeFaciLanguage




    Private Enum ProcessType
        ChangeFaci = 0
        PRFaci = 1
        MoveFaci = 2
    End Enum
    Public Sub New()

    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo)
        MyBase.New(LoginInfo)
    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo, ByVal DAO As CableSoft.Utility.DataAccess.DAO)
        MyBase.New(LoginInfo, DAO)
    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo, ByVal DBConnection As System.Data.Common.DbConnection)
        MyBase.New(LoginInfo, DBConnection)
    End Sub
    ''' <summary>
    ''' 取得可指定變更設備
    ''' </summary>
    ''' <param name="CustId">客編</param>
    ''' <param name="ServiceType">服務別</param>
    ''' <param name="IncludePR">包含拆除</param>
    ''' <param name="IncludeDVR">包否DVR</param>
    ''' <returns>DataTable</returns>
    ''' <remarks></remarks>
    Public Function GetCanChangeFaci(ByVal CustId As Integer,
                                     ByVal ServiceType As String,
                                     ByVal IncludePR As Boolean,
                                     ByVal IncludeDVR As Boolean,
                                     ByVal IncludFilter As Boolean,
                                      ByVal WipType As Integer,
                                     ByVal wipData As DataSet) As DataSet

        'Return DAO.ExecQry(_DAL.GetCanChangeFaci(IncludePR, IncludeDVR),
        '                   New Object() {CustId, ServiceType})
        Dim dsRet As New DataSet
        Dim _ChooseFaci As New CableSoft.SO.BLL.Facility.Facility(Me.LoginInfo, Me.DAO)
        Dim dt As DataTable
        Try
            Using tbCanChooseFaciRefNo As DataTable = DAO.ExecQry(_DAL.QryCanChooseFaciRefNo(WipType, wipData))
                If tbCanChooseFaciRefNo.Rows.Count > 0 AndAlso Not DBNull.Value.Equals(tbCanChooseFaciRefNo.Rows(0).Item(0)) Then
                    dt = _ChooseFaci.QueryCanChooseFaci(CustId, ServiceType, IncludePR, IncludeDVR, IncludFilter, Nothing, tbCanChooseFaciRefNo.Rows(0).Item(0).ToString, False, Nothing)
                Else
                    dt = _ChooseFaci.QueryCanChooseFaci(CustId, ServiceType, IncludePR, IncludeDVR, IncludFilter)
                End If

            End Using
            Dim dtPriv As DataTable = GetPriv()
            dtPriv.TableName = "Priv"
            dsRet.Tables.Add(dt.Copy)
            dsRet.Tables.Add(dtPriv.Copy)
        Finally
            _ChooseFaci.Dispose()
        End Try
        Return dsRet
    End Function
    Public Function GetPriv() As DataTable
        Dim obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, Me.DAO)
        Dim dt As DataTable = obj.GetPriv(Me.LoginInfo.EntryId, "SO11128")
        Dim dtSO111381 As DataTable = obj.GetPriv(Me.LoginInfo.EntryId, "SO111381")
        Dim dtSO111382 As DataTable = obj.GetPriv(Me.LoginInfo.EntryId, "SO111382")

        Try

            Try

                For Each dr As DataRow In dtSO111381.Rows
                    dt.Rows.Add(dr.ItemArray)
                Next
                For Each dr As DataRow In dtSO111382.Rows
                    dt.Rows.Add(dr.ItemArray)
                Next

            Finally
                'dtSO1144.Dispose()
            End Try

            Return dt.Copy
        Finally
            If dt IsNot Nothing Then
                dt.Dispose()
                dt = Nothing
            End If
            If dtSO111381 IsNot Nothing Then
                dtSO111381.Dispose()
                dtSO111381 = Nothing
            End If
            If dtSO111382 IsNot Nothing Then
                dtSO111382.Dispose()
                dtSO111382 = Nothing
            End If

            If obj IsNot Nothing Then
                obj.Dispose()
                obj = Nothing
            End If

        End Try

    End Function

    Public Function GetFaciCode(ByVal SeqNos As String, ByVal CustId As Integer) As DataTable
        Return DAO.ExecQry(_DAL.GetFaciCode(SeqNos), New Object() {LoginInfo.CompCode, CustId})

    End Function
    Public Function GetCanChangeKind(ByVal WipType As Int32, WipRefNo As Int32, ReInstAcrossFlag As Boolean) As DataTable
        Dim obj As New CableSoft.SO.BLL.Wip.Utility.Utility(Me.LoginInfo, Me.DAO)
        Try
            Return obj.GetCanChangeKind(WipType, WipRefNo, ReInstAcrossFlag)

        Finally
            obj.Dispose()
        End Try
    End Function
    Public Function GetChangeFacility(ByVal Kind As Int32, WipRow As DataRow,
                                      FaciRow As DataRow, InChangeDataRow As DataRow,
                                      DeleteCitems As String,
                                      ChooseServiceIDs As String,
                                      ChangeFacility As DataTable) As RIAResult

        Dim obj As New CableSoft.SO.BLL.Wip.Utility.Utility(Me.LoginInfo, Me.DAO)
        Dim aResult As New RIAResult() With {.ResultBoolean = True, .ErrorCode = -1}
        Try
            aResult.ResultBoolean = obj.GetChangeFacility(Kind,
                                                      WipRow, FaciRow, InChangeDataRow,
                                                      DeleteCitems, ChooseServiceIDs, ChangeFacility)
        Finally
            obj.Dispose()
        End Try
        'FaciRow=nothing 
        'DeleteCitems=nothing 

        Return aResult
    End Function
    Public Function GetCPEMAC(ByVal CustId As Integer, ByVal FaciSeqNo As String) As DataSet
        Dim obj As New CableSoft.SO.BLL.Facility.CPEMAC.CPEMAC(Me.LoginInfo, Me.DAO)
        Try
            Return obj.GetCPEMAC(CustId, FaciSeqNo)

        Finally
            obj.Dispose()
        End Try
    End Function
    ''' <summary>
    ''' 取得週期性收費資料
    ''' </summary>
    ''' <param name="CustId">客編</param>
    ''' <param name="FaciSeqNo">設備流水號</param>
    ''' <param name="ServiceType">服務別</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetPeriodCharge(ByVal CustId As Integer, ByVal FaciSeqNo As String,
                                    ByVal ServiceType As String) As DataTable
        Return DAO.ExecQry(_DAL.GetPeriodCharge, New Object() {CustId, FaciSeqNo, ServiceType})

    End Function
    Public Function GetAllChangeData(ByVal ChangeFacility As DataTable) As DataTable
        Dim objWipUtility As New CableSoft.SO.BLL.Wip.Utility.Utility(Me.LoginInfo, DAO)
        Dim aSeqNos As String = "'X'"
        Dim dsWip As DataSet = objWipUtility.GetWipDetail("X", False, Utility.InvoiceType.Maintain)
        Dim dsRet As New DataSet()
        Dim dtRet As DataTable = dsWip.Tables(fWip_Facility).Copy
        dtRet.TableName = fWip_Facility
        dtRet.Rows.Clear()
        Try
            If ChangeFacility.Rows.Count > 0 Then
                For Each rw As DataRow In ChangeFacility.Rows
                    aSeqNos = String.Format(aSeqNos & ",'{0}'", rw("SeqNo").ToString)
                Next
                Dim dtFacility As DataTable = DAO.ExecQry(_DAL.GetAllChangeData(aSeqNos))
                For Each rw As DataRow In dtFacility.Rows
                    Dim rwAdd As DataRow = dtRet.NewRow
                    For Each df As DataColumn In dtRet.Columns
                        If (dtFacility.Columns.Contains(df.ColumnName)) AndAlso
                            (Not rw.IsNull(df.ColumnName)) Then
                            rwAdd.Item(df.ColumnName) = rw.Item(df.ColumnName)
                        End If
                    Next
                    dtRet.Rows.Add(rwAdd)
                Next
            End If
        Finally
            objWipUtility.Dispose()
        End Try
        dsRet.Tables.Add(dtRet)
        Return dsRet.Tables(0)
    End Function
    Public Function ChkDataOK(ByVal SeqNo As String, ByVal SNO As String) As RIAResult
        Dim aResult As New RIAResult() With {.ResultBoolean = False, .ErrorCode = -1}
        If Int32.Parse(DAO.ExecSclr(_DAL.ChkDataOK(), New Object() {SeqNo, SNO}).ToString) > 0 Then
            aResult.ResultBoolean = True
            aResult.ErrorCode = -99
            aResult.ErrorMessage = Language.DataDouble
        Else
            aResult.ResultBoolean = True
        End If
        Return aResult
    End Function
    Public Function GetDelete003Citem(ByVal ServiceIds As String) As String
        Dim aRet As String = Nothing
        If String.IsNullOrEmpty(ServiceIds) Then
            Return String.Empty
        End If
        Using tbSO003C As DataTable = DAO.ExecQry(_DAL.GetDelete003Citem(ServiceIds))
            For Each rw As DataRow In tbSO003C.Rows
                If (Not DBNull.Value.Equals(rw.Item("CitemCode"))) AndAlso
                    (Not String.IsNullOrEmpty(rw.Item("CitemCode"))) Then
                    If String.IsNullOrEmpty(aRet) Then
                        aRet = rw.Item("CitemCode")
                    Else
                        aRet = aRet & "," & rw.Item("CitemCode")
                    End If
                End If

            Next
        End Using
        Return aRet
    End Function
    Public Function GetChooseServiceIDs(ByVal CustId As Integer,
                                         ByVal FaciSeqNo As String) As String
        Dim aRet As String = Nothing
        Dim tb As DataTable = DAO.ExecQry(_DAL.GetChooseServiceID, New Object() {CustId, FaciSeqNo})
        For Each rw As DataRow In tb.Rows
            If (Not DBNull.Value.Equals(rw.Item("ServiceId"))) AndAlso
                (Not String.IsNullOrEmpty(rw.Item("ServiceId"))) Then
                If String.IsNullOrEmpty(aRet) Then
                    aRet = rw.Item("ServiceId")
                Else
                    aRet = aRet & "," & rw.Item("ServiceId")
                End If
            End If
        Next

        'Using rd As DbDataReader = DAO.ExecDtRdr(_DAL.GetChooseServiceID, New Object() {CustId, FaciSeqNo})
        '    While rd.Read
        '        If Not DBNull.Value.Equals(rd.Item("ServiceId")) Then
        '            If String.IsNullOrEmpty(aRet) Then
        '                aRet = rd.Item("ServiceId")
        '            Else
        '                aRet = aRet & "," & rd.Item("ServiceId")
        '            End If
        '        End If
        '    End While
        'End Using
        Return aRet
    End Function
    Public Function GetMovePRFaci(ByVal SNo As String,
                                  ByVal FaciSeqNo As String) As DataSet
        Return GetMovePRFaci(SNo, FaciSeqNo, Nothing)
    End Function
    Public Function GetMovePRFaci(ByVal SNo As String,
                                  ByVal FaciSeqNo As String, ByVal dsSourceWip As DataSet) As DataSet
        Dim objWipUtility As New CableSoft.SO.BLL.Wip.Utility.Utility(Me.LoginInfo, DAO)
        Dim objUtility As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
        Dim dsWip As DataSet = objWipUtility.GetWipDetail("X", False, BLL.Utility.InvoiceType.Maintain)
        Dim tbChangeData As DataTable = DAO.ExecQry(_DAL.GetChangeData, New Object() {FaciSeqNo})
        fNowDate = Date.Now
        Try
            '*******************************拆除設備*****************************
            '新增一筆指定更換設備資料
            Dim rwChangeFacility As DataRow = AddChangeFacility(dsWip, tbChangeData, 0)
            rwChangeFacility.Item("AcceptTime") = New Date(fNowDate.Year, fNowDate.Month,
                                              fNowDate.Day, fNowDate.Hour, fNowDate.Minute, 0)

            rwChangeFacility.Item("SNO") = SNo
            rwChangeFacility.Item("UpdTime") = CableSoft.BLL.Utility.DateTimeUtility.GetDTString(fNowDate)
            rwChangeFacility.Item("NewUpdTime") = fNowDate
            rwChangeFacility.Item("Kind") = Language.MoveDown
            rwChangeFacility.Item("ChooseServiceID") = GetChooseServiceIDs(Int32.Parse(tbChangeData.Rows(0).Item("CustId")),
                                                                           tbChangeData.Rows(0).Item("SeqNo"))

            If (DBNull.Value.Equals(rwChangeFacility.Item("ChooseServiceID"))) OrElse
                (String.IsNullOrEmpty(rwChangeFacility.Item("ChooseServiceID"))) Then
                rwChangeFacility.Item("ChooseServiceID") = DBNull.Value
            End If

            rwChangeFacility.Item("Delete003Citem") = GetDelete003Citem(rwChangeFacility.Item("ChooseServiceID").ToString)

            If (DBNull.Value.Equals(rwChangeFacility.Item("Delete003Citem"))) OrElse
                (String.IsNullOrEmpty(rwChangeFacility.Item("Delete003Citem"))) Then
                rwChangeFacility.Item("Delete003Citem") = DBNull.Value
            End If

            dsWip.Tables(fWip_ChangeFacility).Rows.Add(rwChangeFacility)
            '再新增一筆拆除設備資料
            Dim rwPRFacility As DataRow = AddPRFacility(dsWip, tbChangeData, 0)
            rwPRFacility.Item("PRSNo") = SNo
            rwPRFacility.Item("PRFLAG") = 1
            If dsSourceWip IsNot Nothing Then
                If dsSourceWip.Tables(fWip_PRFacility).AsEnumerable.Count(Function(ByVal rw3 As DataRow)
                                                                              Return rw3.Item("SEQNO") = rwPRFacility.Item("SEQNO")
                                                                          End Function) = 0 Then

                    dsWip.Tables(fWip_PRFacility).Rows.Add(rwPRFacility)
                End If
            Else
                dsWip.Tables(fWip_PRFacility).Rows.Add(rwPRFacility)
            End If


            '*****************************************************************

            ProcessIccPvr(dsWip, tbChangeData, SNo, tbChangeData.Rows(0).Item("CustId"), ProcessType.MoveFaci, Nothing, False, dsSourceWip)
        Finally
            If tbChangeData IsNot Nothing Then
                tbChangeData.Dispose()
            End If
            objWipUtility.Dispose()
            objUtility.Dispose()
        End Try
        Return dsWip
    End Function
    Public Function GetServiceIdAndCitemCode(ByVal CustId As Integer, ByVal FaciSeqNo As String) As DataSet
        Dim tbReturn As New DataTable("ServiceIdCitem")
        Dim dsReturn As New DataSet()
        tbReturn.Columns.Add("ChooseServiceID", GetType(String))
        tbReturn.Columns.Add("Delete003Citem", GetType(String))
        Dim rw As DataRow = tbReturn.NewRow
        rw.Item("ChooseServiceID") = GetChooseServiceIDs(CustId, FaciSeqNo)
        If (DBNull.Value.Equals(rw.Item("ChooseServiceID"))) OrElse
            (String.IsNullOrEmpty(rw.Item("ChooseServiceID"))) Then
            rw.Item("ChooseServiceID") = DBNull.Value
        End If
        rw.Item("Delete003Citem") = GetDelete003Citem(rw.Item("ChooseServiceID").ToString)
        If (DBNull.Value.Equals(rw.Item("Delete003Citem"))) OrElse
            (String.IsNullOrEmpty(rw.Item("Delete003Citem"))) Then
            rw.Item("Delete003Citem") = DBNull.Value
        End If

        tbReturn.Rows.Add(rw)
        dsReturn.Tables.Add(tbReturn)
        Return dsReturn
    End Function
    Public Function GetPRFaci(ByVal SNo As String,
                                  ByVal FaciSeqNo As String) As DataSet
        Return GetPRFaci(SNo, FaciSeqNo, Nothing)
    End Function
    Public Function GetPRFaci(ByVal SNo As String,
                                  ByVal FaciSeqNo As String, ByVal dsSourceWip As DataSet) As DataSet
        Dim objWipUtility As New CableSoft.SO.BLL.Wip.Utility.Utility(Me.LoginInfo, DAO)
        Dim objUtility As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
        Dim dsWip As DataSet = objWipUtility.GetWipDetail("X", False, BLL.Utility.InvoiceType.Maintain)
        Dim tbChangeData As DataTable = DAO.ExecQry(_DAL.GetChangeData, New Object() {FaciSeqNo})
        fNowDate = Date.Now
        Try
            '*******************************拆除設備*****************************
            '新增一筆指定更換設備資料
            Dim rwChangeFacility As DataRow = AddChangeFacility(dsWip, tbChangeData, 0)
            rwChangeFacility.Item("AcceptTime") = New Date(fNowDate.Year, fNowDate.Month,
                                              fNowDate.Day, fNowDate.Hour, fNowDate.Minute, 0)

            rwChangeFacility.Item("SNO") = SNo
            rwChangeFacility.Item("UpdTime") = CableSoft.BLL.Utility.DateTimeUtility.GetDTString(fNowDate)
            rwChangeFacility.Item("NewUpdTime") = fNowDate
            rwChangeFacility.Item("Kind") = Language.Down
            rwChangeFacility.Item("ChooseServiceID") = GetChooseServiceIDs(Int32.Parse(tbChangeData.Rows(0).Item("CustId")),
                                                                           tbChangeData.Rows(0).Item("SeqNo"))

            If (DBNull.Value.Equals(rwChangeFacility.Item("ChooseServiceID"))) OrElse
                (String.IsNullOrEmpty(rwChangeFacility.Item("ChooseServiceID"))) Then
                rwChangeFacility.Item("ChooseServiceID") = DBNull.Value
            End If

            rwChangeFacility.Item("Delete003Citem") = GetDelete003Citem(rwChangeFacility.Item("ChooseServiceID").ToString)
            If (DBNull.Value.Equals(rwChangeFacility.Item("Delete003Citem"))) OrElse
                (String.IsNullOrEmpty(rwChangeFacility.Item("Delete003Citem"))) Then
                rwChangeFacility.Item("Delete003Citem") = DBNull.Value
            End If
            If dsSourceWip IsNot Nothing Then
                If dsSourceWip.Tables(fWip_ChangeFacility).AsEnumerable.Count(Function(ByVal rw3 As DataRow)
                                                                                  Return rw3.Item("SEQNO") = rwChangeFacility.Item("SEQNO")
                                                                              End Function) = 0 Then

                    dsWip.Tables(fWip_ChangeFacility).Rows.Add(rwChangeFacility)
                End If
            Else
                dsWip.Tables(fWip_ChangeFacility).Rows.Add(rwChangeFacility)
            End If


            '再新增一筆拆除設備資料
            Dim rwPRFacility As DataRow = AddPRFacility(dsWip, tbChangeData, 0)
            rwPRFacility.Item("PRSNo") = SNo
            If dsSourceWip IsNot Nothing Then
                If dsSourceWip.Tables(fWip_PRFacility).AsEnumerable.Count(Function(ByVal rw3 As DataRow)
                                                                              Return rw3.Item("SEQNO") = rwPRFacility.Item("SEQNO")
                                                                          End Function) = 0 Then

                    dsWip.Tables(fWip_PRFacility).Rows.Add(rwPRFacility)
                End If
            Else
                dsWip.Tables(fWip_PRFacility).Rows.Add(rwPRFacility)
            End If

            'dsWip.Tables(fWip_PRFacility).Rows.Add(rwPRFacility)
            '*****************************************************************

            ProcessIccPvr(dsWip, tbChangeData, SNo, tbChangeData.Rows(0).Item("CustId"), ProcessType.PRFaci, Nothing, False, dsSourceWip)
        Finally
            If tbChangeData IsNot Nothing Then
                tbChangeData.Dispose()
            End If
            objWipUtility.Dispose()
            objUtility.Dispose()
        End Try
        Return dsWip
    End Function
    Public Overloads Function GetReInstFaci(ByVal SNo As String, ByVal FaciSeqNo As String, ByVal dsSourceWip As DataSet) As DataSet
        Return GetReInstFaci(SNo, FaciSeqNo, dsSourceWip, False)
    End Function
    Public Overloads Function GetReInstFaci(ByVal SNo As String, ByVal FaciSeqNo As String) As DataSet
        Return GetReInstFaci(SNo, FaciSeqNo, Nothing, False)
    End Function
    Public Overloads Function GetReInstFaci(ByVal SNo As String, ByVal FaciSeqNo As String, ByVal FilterDVR As Boolean) As DataSet
        Return GetReInstFaci(SNo, FaciSeqNo, Nothing, FilterDVR)

    End Function
    ''' <summary>
    ''' 取得指定更換設備資訊(更換設備)
    ''' </summary>
    ''' <param name="SNo">工單單號</param>
    ''' <param name="FaciSeqNo">設備流水號</param>    
    ''' <returns>DataSet</returns>
    ''' <remarks></remarks>
    Public Overloads Function GetReInstFaci(ByVal SNo As String,
                                  ByVal FaciSeqNo As String,
                                  ByVal dsSourceWip As DataSet,
                                  ByVal FilterDVR As Boolean) As DataSet
        Dim objWipUtility As New CableSoft.SO.BLL.Wip.Utility.Utility(Me.LoginInfo, DAO)
        Dim objUtility As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
        Dim dsWip As DataSet = objWipUtility.GetWipDetail("X", False, BLL.Utility.InvoiceType.Maintain)
        Dim tbChangeData As DataTable = DAO.ExecQry(_DAL.GetChangeData, New Object() {FaciSeqNo})
        fNowDate = Date.Now
        Try
            '*******************************更換設備*****************************
            '新增一筆指定更換設備資料
            Dim rwChangeFacility As DataRow = AddChangeFacility(dsWip, tbChangeData, 0)
            Dim aChooseServiceIDs As String = Nothing
            Dim aDelete003Citems As String = Nothing
            rwChangeFacility.Item("AcceptTime") = New Date(fNowDate.Year, fNowDate.Month,
                                              fNowDate.Day, fNowDate.Hour, fNowDate.Minute, 0)

            rwChangeFacility.Item("SNO") = SNo
            rwChangeFacility.Item("UpdTime") = CableSoft.BLL.Utility.DateTimeUtility.GetDTString(fNowDate)
            rwChangeFacility.Item("UpdTime") = fNowDate
            rwChangeFacility.Item("Kind") = Language.Change '更換
            rwChangeFacility.Item("ChooseServiceID") = GetChooseServiceIDs(Int32.Parse(tbChangeData.Rows(0).Item("CustId")),
                                                                           tbChangeData.Rows(0).Item("SeqNo"))

            If (DBNull.Value.Equals(rwChangeFacility.Item("ChooseServiceID"))) OrElse
                (String.IsNullOrEmpty(rwChangeFacility.Item("ChooseServiceID"))) Then
                rwChangeFacility.Item("ChooseServiceID") = DBNull.Value
            End If

            aChooseServiceIDs = rwChangeFacility.Item("ChooseServiceID").ToString
            'Cancel fill the file of Delete003Citem For Jacky By Kin 2018/05/11
            'rwChangeFacility.Item("Delete003Citem") = GetDelete003Citem(rwChangeFacility.Item("ChooseServiceID").ToString)
            rwChangeFacility.Item("Delete003Citem") = DBNull.Value
            If (DBNull.Value.Equals(rwChangeFacility.Item("Delete003Citem"))) OrElse
                (String.IsNullOrEmpty(rwChangeFacility.Item("Delete003Citem"))) Then
                rwChangeFacility.Item("Delete003Citem") = DBNull.Value
            End If
            aDelete003Citems = rwChangeFacility.Item("Delete003Citem").ToString
            If dsSourceWip IsNot Nothing Then
                If dsSourceWip.Tables(fWip_ChangeFacility).AsEnumerable.Count(Function(ByVal rw As DataRow)
                                                                                  Return rw.Item("SEQNO") = rwChangeFacility.Item("SeqNo")
                                                                              End Function) = 0 Then
                    dsWip.Tables(fWip_ChangeFacility).Rows.Add(rwChangeFacility)
                End If
            Else
                dsWip.Tables(fWip_ChangeFacility).Rows.Add(rwChangeFacility)
            End If

            '再新增一筆拆除設備資料
            Dim rwPRFacility As DataRow = AddPRFacility(dsWip, tbChangeData, 0)
            rwPRFacility.Item("PRSNo") = SNo
            If dsSourceWip IsNot Nothing Then
                If dsSourceWip.Tables(fWip_PRFacility).AsEnumerable.Count(Function(ByVal rw As DataRow)
                                                                              Return rw.Item("SEQNO") = rwPRFacility.Item("SeqNo")
                                                                          End Function) = 0 Then
                    dsWip.Tables(fWip_PRFacility).Rows.Add(rwPRFacility)
                End If
            Else
                dsWip.Tables(fWip_PRFacility).Rows.Add(rwPRFacility)
            End If



            '*****************************************************************
            '********************************換裝設備***************************
            '新增一筆指定換裝設備資料
            Dim rwChangeFacility2 As DataRow = AddChangeFacility(dsWip, tbChangeData, 0)
            Dim aSeqNo As String = objUtility.GetFaciSeqNo(False)
            rwChangeFacility2.Item("AcceptTime") = New Date(fNowDate.Year, fNowDate.Month,
                                              fNowDate.Day, fNowDate.Hour, fNowDate.Minute, 0)
            'rwChangeFacility2.Item("CUSTID") = CustId
            rwChangeFacility2.Item("SeqNo") = aSeqNo
            rwChangeFacility2.Item("SNO") = SNo
            If Not String.IsNullOrEmpty(aChooseServiceIDs) Then
                rwChangeFacility2.Item("ChooseServiceID") = aChooseServiceIDs
            End If
            If Not String.IsNullOrEmpty(aDelete003Citems) Then
                rwChangeFacility2.Item("Delete003Citem") = aDelete003Citems
            End If
            'Cancel fill the file of Delete003Citem For Jacky By Kin 2019/03/20
            rwChangeFacility2.Item("Delete003Citem") = DBNull.Value
            rwChangeFacility2.Item("UpdTime") = CableSoft.BLL.Utility.DateTimeUtility.GetDTString(fNowDate)
            rwChangeFacility2.Item("NewUpdTime") = fNowDate
            rwChangeFacility2.Item("Kind") = Language.ChangeInstall '換裝
            SetNullChangeFacility(rwChangeFacility2)
            If dsSourceWip IsNot Nothing Then
                If dsSourceWip.Tables(fWip_ChangeFacility).AsEnumerable.Count(Function(ByVal rw As DataRow)
                                                                                  Return rw.Item("SEQNO") = rwChangeFacility2.Item("SeqNo")
                                                                              End Function) = 0 Then
                    dsWip.Tables(fWip_ChangeFacility).Rows.Add(rwChangeFacility2)
                End If
            Else
                dsWip.Tables(fWip_ChangeFacility).Rows.Add(rwChangeFacility2)
            End If



            '再新增一筆換裝設備資料
            Dim rwFacility As DataRow = AddFacility(dsWip, tbChangeData, 0)
            SetNullPRFacility(rwFacility)
            With rwFacility
                .Item("SNO") = SNo
                .Item("ReSEQNo") = tbChangeData.Rows(0).Item("SeqNo")
                .Item("SeqNo") = aSeqNo
            End With
            If dsSourceWip IsNot Nothing Then
                If dsSourceWip.Tables(fWip_Facility).AsEnumerable.Count(Function(ByVal rw As DataRow)
                                                                            If dsSourceWip.Tables(fWip_Facility).Columns.Contains("ReSEQNO") Then
                                                                                If Not DBNull.Value.Equals("ReSEQNO") AndAlso Not DBNull.Value.Equals(tbChangeData.Rows(0).Item("SeqNo")) Then
                                                                                    Return rw.Item("ReSEQNO").ToString = tbChangeData.Rows(0).Item("SeqNo").ToString
                                                                                Else
                                                                                    Return False
                                                                                End If
                                                                            Else
                                                                                Return False
                                                                            End If


                                                                        End Function) = 0 Then
                    dsWip.Tables.Item(fWip_Facility).Rows.Add(rwFacility)
                End If
            Else
                dsWip.Tables.Item(fWip_Facility).Rows.Add(rwFacility)
            End If


            '*****************************************************************

            ProcessIccPvr(dsWip, tbChangeData, SNo, tbChangeData.Rows(0).Item("CustId"), ProcessType.ChangeFaci, aSeqNo, FilterDVR, dsSourceWip)

        Finally
            If tbChangeData IsNot Nothing Then
                tbChangeData.Dispose()
            End If
            objWipUtility.Dispose()
            objUtility.Dispose()
        End Try
        Return dsWip
    End Function
    ''' <summary>
    ''' 取得指定變更設備資訊
    ''' </summary>
    ''' <param name="Wip">工單資料</param>
    ''' <param name="Facility">設備資料</param>
    ''' <returns>DataTable</returns>
    ''' <remarks>ChangeFacility</remarks>
    Public Function GetChangeFacility(ByVal Wip As DataTable, ByVal Facility As DataTable) As DataTable
        Dim objWipUtility As New CableSoft.SO.BLL.Wip.Utility.Utility(Me.LoginInfo, DAO)
        Dim dsWip As DataSet = objWipUtility.GetWipDetail("X", False, BLL.Utility.InvoiceType.Maintain)
        Dim dsRet As New DataSet()
        Dim dtRet As DataTable = Nothing
        fNowDate = Date.Now
        Try
            Dim rwChangeFacility As DataRow = dsWip.Tables(fWip_ChangeFacility).NewRow
            With rwChangeFacility
                .Item("AcceptTime") = New Date(fNowDate.Year, fNowDate.Month,
                                              fNowDate.Day, fNowDate.Hour, fNowDate.Minute, 0)
                .Item("UpdTime") = CableSoft.BLL.Utility.DateTimeUtility.GetDTString(fNowDate)
                .Item("NewUpdTime") = fNowDate
                .Item("UpdEn") = Me.LoginInfo.EntryName
                .Item("Kind") = Language.AddNew  '新增
                .Item("SNo") = Wip.Rows(0).Item("SNo")
                .Item("CustId") = Wip.Rows(0).Item("CustId")
                .Item("FaciSNo") = Facility.Rows(0).Item("FaciSNo")
                .Item("SeqNo") = Facility.Rows(0).Item("SeqNo")
                .Item("NPromCode") = Facility.Rows(0).Item("PromCode")
                .Item("NPromName") = Facility.Rows(0).Item("PromName")
                .Item("NBPCode") = Facility.Rows(0).Item("BPCode")
                .Item("NBPName") = Facility.Rows(0).Item("BPName")
                .Item("NOrderNo") = Facility.Rows(0).Item("OrderNo")
                .Item("NIPAddress") = Facility.Rows(0).Item("IPAddress")
                .Item("NCMBaudRateNo") = Facility.Rows(0).Item("CMBaudRateNo")
                .Item("NCMBaudRate") = Facility.Rows(0).Item("CMBaudRate")
                .Item("NDynIPCount") = Facility.Rows(0).Item("DynIPCount")
                .Item("NFixIPCount") = Facility.Rows(0).Item("FixIPCount")
                .Item("NContNo") = Facility.Rows(0).Item("ContNo")
                .Item("NContractDate") = Facility.Rows(0).Item("NContractDate")
                .Item("NContractCust") = Facility.Rows(0).Item("ContractCust")
                .Item("NDeposit") = Facility.Rows(0).Item("Deposit")
                .Item("NDVRAuthSizeCode") = Facility.Rows(0).Item("DVRAuthSizeCode")
            End With

            dsWip.Tables(fWip_ChangeFacility).Rows.Add(rwChangeFacility)


        Finally
            objWipUtility.Dispose()
        End Try
        dtRet = dsWip.Tables(fWip_ChangeFacility).Copy
        dtRet.TableName = fWip_ChangeFacility
        dsRet.Tables.Add(dtRet)
        Return dsRet.Tables(0)
        'Return dsWip.Tables(fMaintain_ChangeFacility)
    End Function
    ''' <summary>
    ''' 取得可選速率
    ''' </summary>
    ''' <param name="Type">種類 0: 降速, 1:升速</param>
    ''' <param name="CMRateCode">現有速率代碼</param>
    ''' <returns>DataTable</returns>
    ''' <remarks></remarks>
    Public Function GetCMRateCode(ByVal Type As Integer, ByVal CMRateCode As Integer) As DataTable
        Return DAO.ExecQry(_DAL.GetCMRateCode(Type), New Object() {CMRateCode})
    End Function
    ''' <summary>
    ''' 取得可選容量
    ''' </summary>
    ''' <param name="Type">種類</param>
    ''' <param name="DVRSizeCode">現有容量代碼</param>
    ''' <returns>DataTable</returns>
    ''' <remarks></remarks>
    Public Function GetDVRSizeCode(ByVal Type As Integer, ByVal DVRSizeCode As Integer) As DataTable
        Return DAO.ExecQry(_DAL.GetDVRSizeCode(Type), New Object() {DVRSizeCode})
    End Function
    ''' <summary>
    ''' 取得可選IP數
    ''' </summary>
    ''' <param name="Type">種類</param>
    ''' <param name="IPCount">現有IP數量</param>
    ''' <returns>DataTable</returns>
    ''' <remarks></remarks>
    Public Function GetIPCount(ByVal Type As Integer, ByVal IPCount As Integer, ByVal ZeroIPCount As Boolean) As DataTable
        Dim dtReturn As DataTable = DAO.ExecQry(_DAL.GetIPCount(Type, ZeroIPCount), New Object() {IPCount})
        Return dtReturn
    End Function
    Public Overloads Function GetMoveFaci(ByVal SNo As String, ByVal FaciSeqNo As String, ByVal filterDVR As Boolean) As DataTable
        Return GetMaintainFaci(SNo, FaciSeqNo, filterDVR, Nothing, 1)
    End Function
    Public Overloads Function GetMoveFaci(ByVal SNo As String, ByVal FaciSeqNo As String, ByVal dsWipSource As DataSet, ByVal filterDVR As Boolean) As DataTable
        Return GetMaintainFaci(SNo, FaciSeqNo, filterDVR, dsWipSource, 1)
    End Function
    Public Overloads Function GetMaintainFaci(ByVal SNo As String, ByVal FaciSeqNo As String, ByVal dsWipSource As DataSet) As DataTable
        Return GetMaintainFaci(SNo, FaciSeqNo, False, dsWipSource, 0)
    End Function
    Public Overloads Function GetMaintainFaci(ByVal SNo As String,
                                              ByVal FaciSeqNo As String, ByVal filterDVR As Boolean) As DataTable
        Return GetMaintainFaci(SNo, FaciSeqNo, filterDVR, Nothing, 0)
    End Function
    Public Overloads Function GetMaintainFaci(ByVal SNo As String,
                                              ByVal FaciSeqNo As String) As DataTable
        Return GetMaintainFaci(SNo, FaciSeqNo, False, Nothing, 0)
    End Function
    ''' <summary>
    ''' 取得指定維修設備資訊
    ''' </summary>
    ''' <param name="SNo">工單單號</param>
    ''' <param name="FaciSeqNo">設備流水號</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Overloads Function GetMaintainFaci(ByVal SNo As String,
                                              ByVal FaciSeqNo As String,
                                              ByVal filterDVR As Boolean,
                                              ByVal dsWipSource As DataSet, ByVal KindType As Integer) As DataTable
        Dim objWipUtility As New CableSoft.SO.BLL.Wip.Utility.Utility(Me.LoginInfo, DAO)
        Dim dsWip As DataSet = objWipUtility.GetWipDetail("X", False, BLL.Utility.InvoiceType.Maintain)        
        Dim tbChangeData As DataTable = DAO.ExecQry(_DAL.GetChangeData, New Object() {FaciSeqNo})
        Dim tbChildFaci As DataTable = Nothing
        Dim dtRet As DataTable = Nothing
        Dim dsRet As New DataSet
        Dim KindTypeName As String = Nothing
        fNowDate = Date.Now
        KindTypeName = Language.Maintain   '維修
        If KindType <> 0 Then KindTypeName = Language.MoveFaci   '移機
        Try
            Dim rw As DataRow = AddChangeFacility(dsWip, tbChangeData, 0)
            With rw
                .Item("AcceptTime") = New Date(fNowDate.Year, fNowDate.Month,
                                             fNowDate.Day, fNowDate.Hour, fNowDate.Minute, 0)

                .Item("SNO") = SNo
                .Item("UpdTime") = CableSoft.BLL.Utility.DateTimeUtility.GetDTString(fNowDate)
                .Item("NewUpdTime") = fNowDate
                .Item("Kind") = KindTypeName


                .Item("ChooseServiceID") = GetChooseServiceIDs(Int32.Parse(tbChangeData.Rows(0).Item("CustId")),
                                                               tbChangeData.Rows(0).Item("SeqNo"))
                If (DBNull.Value.Equals(.Item("ChooseServiceID"))) OrElse
                    (String.IsNullOrEmpty(.Item("ChooseServiceID"))) Then
                    .Item("ChooseServiceID") = DBNull.Value
                End If
                'Cancel fill the file of Delete003Citem for jacky By Kin 2018/05/11
                '.Item("Delete003Citem") = GetDelete003Citem(.Item("ChooseServiceID").ToString)
                .Item("Delete003Citem") = DBNull.Value
                If (DBNull.Value.Equals(.Item("Delete003Citem"))) OrElse
                    (String.IsNullOrEmpty(.Item("Delete003Citem"))) Then
                    .Item("Delete003Citem") = DBNull.Value
                End If
            End With
            dsWip.Tables(fWip_ChangeFacility).Rows.Add(rw)

            If (Not tbChangeData.Rows(0).IsNull("FaciRefNo")) AndAlso
               (Int32.Parse(tbChangeData.Rows(0).Item("FaciRefNo")) = 3) Then
                If filterDVR Then
                    tbChildFaci = DAO.ExecQry(_DAL.GetChildFaci(True),
                                       New Object() {tbChangeData.Rows(0).Item("CustId"),
                                                     tbChangeData.Rows(0).Item("SmartCardNo")})
                Else
                    tbChildFaci = DAO.ExecQry(_DAL.GetChildFaci(),
                                       New Object() {tbChangeData.Rows(0).Item("CustId"),
                                                     tbChangeData.Rows(0).Item("SmartCardNo"),
                                                     tbChangeData.Rows(0).Item("SeqNo")})
                End If

                For i As Int32 = 0 To tbChildFaci.Rows.Count - 1
                    Dim rw2 As DataRow = AddChangeFacility(dsWip, tbChildFaci, i)
                    With rw2
                        .Item("AcceptTime") = New Date(fNowDate.Year, fNowDate.Month,
                                                     fNowDate.Day, fNowDate.Hour, fNowDate.Minute, 0)

                        .Item("SNO") = SNo
                        .Item("UpdTime") = CableSoft.BLL.Utility.DateTimeUtility.GetDTString(fNowDate)
                        .Item("NewUpdTime") = fNowDate
                        .Item("Kind") = KindTypeName
                    End With
                    If dsWipSource IsNot Nothing Then
                        If dsWipSource.Tables(fWip_ChangeFacility).AsEnumerable.Count(Function(ByVal rw3 As DataRow)
                                                                                          Return rw3.Item("SEQNO") = tbChildFaci.Rows(i).Item("SeqNo")
                                                                                      End Function) = 0 Then

                            dsWip.Tables(fWip_ChangeFacility).Rows.Add(rw2)
                        End If
                    Else
                        dsWip.Tables(fWip_ChangeFacility).Rows.Add(rw2)
                    End If


                Next
            End If


        Finally
            objWipUtility.Dispose()
            If tbChangeData IsNot Nothing Then
                tbChangeData.Dispose()
            End If
            If tbChildFaci IsNot Nothing Then
                tbChildFaci.Dispose()
            End If
        End Try
        dsRet.Tables.Add(dsWip.Tables(fWip_ChangeFacility).Copy)
        Return dsRet.Tables(0)

        'Return dsWip.Tables(fMaintain_ChangeFacility).Copy
        'Return dsWip.Tables(fMaintain_ChangeFacility)
    End Function
    Private Sub ProcessIccPvr(ByVal dsWip As DataSet,
                                    ByVal tbParentData As DataTable,
                                    ByVal SNo As String,
                                    ByVal CustId As Integer,
                                    ByVal ProcessValue As ProcessType,
                                    ByVal DVRStbNo As String,
                                    ByVal FilterDVR As Boolean,
                                    ByVal dsSource As DataSet)


        If (tbParentData.Rows(0).IsNull("FaciRefNo")) OrElse
            (Int32.Parse(tbParentData.Rows(0).Item("FaciRefNo")) <> 3) Then
            Exit Sub
        End If

        Dim tbChildFaci As DataTable
        If Not FilterDVR Then
            tbChildFaci = DAO.ExecQry(_DAL.GetChildFaci(FilterDVR),
                                               New Object() {CustId,
                                                              tbParentData.Rows(0).Item("SmartCardNo"),
                                                              tbParentData.Rows(0).Item("SeqNo")})
        Else
            tbChildFaci = DAO.ExecQry(_DAL.GetChildFaci(FilterDVR),
                                               New Object() {CustId,
                                                              tbParentData.Rows(0).Item("SmartCardNo")})
        End If

        Dim objUtility As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
        Try
            For i As Int32 = 0 To tbChildFaci.Rows.Count - 1
                '*******************************更換設備*****************************
                '新增一筆指定更換設備資料

                Dim rwChangeFacility As DataRow = AddChangeFacility(dsWip, tbChildFaci, i)
                rwChangeFacility.Item("AcceptTime") = New Date(fNowDate.Year, fNowDate.Month,
                                                  fNowDate.Day, fNowDate.Hour, fNowDate.Minute, 0)
                'rwChangeFacility.Item("CUSTID") = CustId
                rwChangeFacility.Item("SNO") = SNo
                rwChangeFacility.Item("UpdTime") = CableSoft.BLL.Utility.DateTimeUtility.GetDTString(fNowDate)
                rwChangeFacility.Item("NewUpdTime") = fNowDate
                Select Case ProcessValue
                    Case ProcessType.ChangeFaci
                        rwChangeFacility.Item("Kind") = Language.Change '更換
                    Case ProcessType.PRFaci
                        rwChangeFacility.Item("Kind") = Language.Down '拆除
                    Case ProcessType.MoveFaci
                        rwChangeFacility.Item("Kind") = Language.MoveDown '移拆
                End Select
                If dsSource IsNot Nothing Then
                    If dsSource.Tables(fWip_ChangeFacility).AsEnumerable.Count(Function(ByVal rw As DataRow)
                                                                                   Return rw.Item("SEQNO") = tbChildFaci.Rows(i).Item("SeqNo")
                                                                               End Function) = 0 Then

                        dsWip.Tables(fWip_ChangeFacility).Rows.Add(rwChangeFacility)
                    End If
                Else
                    dsWip.Tables(fWip_ChangeFacility).Rows.Add(rwChangeFacility)
                End If


                '再新增一筆拆除設備資料

                Dim rwPRFacility As DataRow = AddPRFacility(dsWip, tbChildFaci, i)
                rwPRFacility.Item("PRSNo") = SNo
                If ProcessValue = ProcessType.MoveFaci Then
                    rwPRFacility.Item("PRFlag") = 1
                End If
                If dsSource IsNot Nothing Then
                    If dsSource.Tables(fWip_PRFacility).AsEnumerable.Count(Function(ByVal rw As DataRow)
                                                                               Return rw.Item("SEQNO") = tbChildFaci.Rows(i).Item("SeqNo")
                                                                           End Function) = 0 Then
                        dsWip.Tables(fWip_PRFacility).Rows.Add(rwPRFacility)
                    End If
                Else
                    dsWip.Tables(fWip_PRFacility).Rows.Add(rwPRFacility)
                End If


                '*****************************************************************
                '********************************換裝設備***************************
                If ProcessValue = ProcessType.ChangeFaci Then

                    '新增一筆指定換裝設備資料
                    Dim rwChangeFacility2 As DataRow = AddChangeFacility(dsWip, tbChildFaci, i)
                    Dim aSeqNo As String = objUtility.GetFaciSeqNo(False)
                    rwChangeFacility2.Item("AcceptTime") = New Date(fNowDate.Year, fNowDate.Month,
                                                      fNowDate.Day, fNowDate.Hour, fNowDate.Minute, 0)
                    'rwChangeFacility2.Item("CUSTID") = CustId
                    rwChangeFacility2.Item("SeqNo") = aSeqNo
                    rwChangeFacility2.Item("SNO") = SNo
                    rwChangeFacility2.Item("UpdTime") = CableSoft.BLL.Utility.DateTimeUtility.GetDTString(fNowDate)
                    rwChangeFacility2.Item("NewUpdTime") = fNowDate
                    rwChangeFacility2.Item("Kind") = Language.ChangeInstall  '換裝
                    SetNullChangeFacility(rwChangeFacility2)
                    If dsSource IsNot Nothing Then
                        If dsSource.Tables(fWip_ChangeFacility).AsEnumerable.Count(Function(ByVal rw As DataRow)
                                                                                       Return rw.Item("SEQNO") = tbChildFaci.Rows(i).Item("SeqNo")
                                                                                   End Function) = 0 Then

                            dsWip.Tables(fWip_ChangeFacility).Rows.Add(rwChangeFacility2)
                        End If
                    Else
                        dsWip.Tables(fWip_ChangeFacility).Rows.Add(rwChangeFacility2)
                    End If


                    '再新增一筆”換裝”設備資料
                    Dim rwFacility As DataRow = AddFacility(dsWip, tbChildFaci, i)
                    SetNullPRFacility(rwFacility)
                    With rwFacility
                        .Item("SNO") = SNo
                        .Item("ReSEQNo") = tbChildFaci.Rows(i).Item("SeqNo")
                        .Item("SeqNo") = aSeqNo
                        If Not .IsNull("STBSNo") Then
                            .Item("STBSNo") = DVRStbNo
                            '.Item("STBSNo") = rwChangeFacility2.Item("SeqNo")
                            '.Item("STBSNo") = tbChildFaci.Rows(0).Item("SeqNo")
                            '.Item("STBSNo") = .Item("SeqNo")
                            '.Item("STBSNo") = tbParentData.Rows(0).Item("SEQNO")
                        End If
                    End With
                    If dsSource IsNot Nothing Then
                        If dsSource.Tables(fWip_Facility).AsEnumerable.Count(Function(ByVal rw As DataRow)
                                                                                 If dsSource.Tables(fWip_Facility).Columns.Contains("ReSEQNO") Then
                                                                                     If Not DBNull.Value.Equals(rw.Item("ReSEQNO")) AndAlso Not DBNull.Value.Equals(tbChildFaci.Rows(0).Item("SeqNo")) Then
                                                                                         Return rw.Item("ReSEQNO").ToString() = tbChildFaci.Rows(i).Item("SeqNo").ToString()
                                                                                     Else
                                                                                         Return False
                                                                                     End If
                                                                                 Else
                                                                                     Return False
                                                                                 End If



                                                                             End Function) = 0 Then
                            dsWip.Tables.Item(fWip_Facility).Rows.Add(rwFacility)
                        End If
                    Else
                        dsWip.Tables.Item(fWip_Facility).Rows.Add(rwFacility)
                    End If


                End If
                '*****************************************************************

            Next
        Finally
            If tbChildFaci IsNot Nothing Then
                tbChildFaci.Dispose()
                tbChildFaci = Nothing
            End If

            If objUtility IsNot Nothing Then
                objUtility.Dispose()
                objUtility = Nothing
            End If

        End Try

    End Sub
    Private Sub SetNullChangeFacility(ByVal rw As DataRow)
        With rw
            rw.Item("FaciSNo") = DBNull.Value
        End With

    End Sub
    Private Sub SetNullPRFacility(ByVal rw As DataRow)
        With rw
            .Item("FaciSNo") = DBNull.Value
            .Item("SmartCardNo") = DBNull.Value
            .Item("PRSNo") = DBNull.Value
            .Item("PrDate") = DBNull.Value
            .Item("PREn1") = DBNull.Value
            .Item("PrName1") = DBNull.Value
            .Item("PREn2") = DBNull.Value
            .Item("PrName2") = DBNull.Value
            .Item("InstDate") = DBNull.Value
            .Item("CMOPENDATE") = DBNull.Value
            .Item("CMCLOSEDATE") = DBNull.Value
            .Item("DISABLEACCOUNT") = DBNull.Value
            .Item("ENABLEACCOUNT") = DBNull.Value
            .Item("GetDate") = DBNull.Value
        End With

    End Sub
    Private Function AddChangeFacility(ByVal dsWip As DataSet, ByVal tbChangeData As DataTable, ByVal RowIndex As Int32) As DataRow
        Dim rw As DataRow = dsWip.Tables(fWip_ChangeFacility).NewRow
        rw.Item("FaciSNo") = tbChangeData.Rows(RowIndex).Item("FaciSNo")
        rw.Item("SeqNo") = tbChangeData.Rows(RowIndex).Item("SeqNo")
        rw.Item("UpdEn") = Me.LoginInfo.EntryName
        rw.Item("CustId") = tbChangeData.Rows(RowIndex).Item("CustId")
        Return rw
    End Function
    Private Function AddPRFacility(ByVal dsWip As DataSet, ByVal tbChangeData As DataTable, ByVal RowIndex As Int32) As DataRow
        Dim rw As DataRow = dsWip.Tables.Item(fWip_PRFacility).NewRow

        For Each df As DataColumn In dsWip.Tables.Item(fWip_PRFacility).Columns
            If tbChangeData.Columns.Contains(df.ColumnName) Then
                rw.Item(df.ColumnName) = tbChangeData.Rows(RowIndex).Item(df.ColumnName)
            End If
        Next
        Return rw
    End Function
    Private Function AddFacility(ByVal dsWip As DataSet, ByVal tbChangeData As DataTable, ByVal RowIndex As Int32) As DataRow
        Dim rw As DataRow = dsWip.Tables.Item(fWip_Facility).NewRow

        For Each df As DataColumn In dsWip.Tables.Item(fWip_Facility).Columns
            If tbChangeData.Columns.Contains(df.ColumnName) Then
                rw.Item(df.ColumnName) = tbChangeData.Rows(RowIndex).Item(df.ColumnName)
            End If
        Next
        Return rw
    End Function


#Region "IDisposable Support"
    Private disposedValue As Boolean ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                If (Me.MustDispose) AndAlso (Me.DAO IsNot Nothing) Then
                    DAO.Dispose()
                End If
                If Language IsNot Nothing Then
                    Language.Dispose()
                    Language = Nothing
                End If
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
