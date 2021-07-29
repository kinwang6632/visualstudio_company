Imports System.Data.Common
Imports CableSoft.Utility.DataAccess
Imports CableSoft.BLL.Utility
Imports BillingAPILanguage = CableSoft.BLL.Language.SO61.WipPRLanguage
'停拆移機工單改約－236
'	使用時機：異動未完工停拆移工單的預約時間。
'	使用限制：不適用於已結案、已線上回報、已方案變更及訂單單號之工單。
Public Class BillingAPI236
    Inherits CableSoft.BLL.Utility.BLLBasic
    Implements IDisposable, CableSoft.BLL.BillingAPI.IBillingAPI
    Private _DAL As New BillingAPI236DAL(Me.LoginInfo.Provider)
    'Private _DAL As New BillingAPI236DALMultiDB(Me.LoginInfo.Provider)
    Private Lang As New CableSoft.BLL.Language.SO61.BillingAPI236Language
    '回應碼	回應狀態	回應訊息
    '0	    成功	
    '-1	    失敗	    {參數}資料有誤!!
    '-101	失敗	    查無單號
    '-151	失敗	    該單號已結案, 不可改約
    '-152	失敗	    該單號已線上回報, 不可改約
    '-153	失敗	    部份費用已入暫收款或實收,不可改約
    '-154	失敗	    該單號已方案變更, 不可改約
    '-155	失敗	    己排件數大於預設件數，不允許預約
    '-103	失敗	    此為訂單產生之工單，請由訂單機制處理
    Public Function Execute(SeqNo As Integer, InData As System.Data.DataSet) As CableSoft.BLL.Utility.RIAResult Implements CableSoft.BLL.BillingAPI.IBillingAPI.Execute
        Dim result As RIAResult = Nothing
        Dim drMain As DataRow = InData.Tables("Main").Rows(0)
        Dim SNo As String = drMain("SNo")
        Dim ResvTime As DateTime = DateTime.Parse(InData.Tables("Main").Rows(0).Item("ResvTime"))
        '#8706
        'Me.LoginInfo.EntryName = InData.Tables("Main").Rows(0).Item("Upden")
        'Main	單	APIID	        命令識別碼	        V	236
        'Main	單	Compcode	    公司別	            V	預設公司別
        'Main	單	Caller	        呼叫來源	        V	自定名稱 Ex. IVR,CSR,WEB…
        'Main	單	Seqno	        來源識別碼	        V	自定編碼，編碼規則：17碼，YYYYMMDDHHMMSS+3碼流水號
        'Main	單	SNO	            工單單號	        V	
        'Main	單	ResvTime	    預約日期	        V	YYYY/MM/DD HH24:MI:SS
        'Main	單	NoteType	    備註型態		        0:覆蓋1:前加2:後加
        'Main	單	Note	        備註		
        'Main	單	WorkServCode	實際派工服務區代碼	V	'#7899 增加
        If InData.Tables("Main").Columns.Contains("HandleEn") Then
            Me.LoginInfo.EntryId = InData.Tables("Main").Rows(0).Item("HandleEn")
            Using tbCM003 As DataTable = DAO.ExecQry(_DAL.GetEmpName, New Object() {Me.LoginInfo.EntryId})
                Me.LoginInfo.EntryName = tbCM003.Rows(0).Item("EmpName")
            End Using
        End If
       
        Using bll As New SaveData(LoginInfo, DAO)
            Dim WipData As DataSet = Nothing
            Using PR As New PR(LoginInfo, DAO)
                WipData = PR.GetPRData(SNo)
                '檢核預約時間是否可以改約
                result = ChkCanResv(WipData, ResvTime, drMain)
                If result.ResultBoolean = False Then
                    result.ErrorCode = -155
                    Return result
                End If
            End Using
            WipData.Tables("Wip").Rows(0).Item("ResvTime") = ResvTime

            'If InData.Tables("Main").Columns.Contains("WorkServCode") Then
            '    If Not DBNull.Value.Equals(InData.Tables("Main").Rows(0).Item("WorkServCode")) Then
            '        WipData.Tables("Wip").Rows(0).Item("WorkServCode") = InData.Tables("Main").Rows(0).Item("WorkServCode")
            '    End If
            'End If

            '重新取得收費/設備資料
            result = ChangePRData(ResvTime, WipData)
            If result.ResultBoolean = False Then
                Return result
            End If
            '填寫工單資料 
            result = UpdateWipData(drMain, WipData)
            If result.ResultBoolean = False Then
                Return result
            End If

            '檢核工單是否可存檔
            Using Vali As New Validate(LoginInfo, DAO)
                result = Vali.ChkDataOk(EditMode.Edit, WipData)
                If result.ResultBoolean = False Then
                    Return result
                End If
            End Using
            '工單存檔
            result = bll.Save(EditMode.Edit, False, WipData, False)
            If result.ResultBoolean = False Then
                Return result
            End If
            '回傳資料
            result.ResultDataSet = GetReturnData(WipData)
            WipData.Dispose()
        End Using
        Return result
    End Function
    Private Function ChangePRData(ResvTime As DateTime, ByRef WipData As DataSet) As RIAResult
        '重新取得相關資料
        Using PR As New PR(LoginInfo, DAO)
            Dim WipRow As DataRow = WipData.Tables("Wip").Rows(0)
            Using RetData As DataSet = PR.GetNormalCalculateData(WipRow.Item("CustId"), WipRow.Item("ServiceType"), WipRow.Item("PRCode"), ResvTime, WipRow.Item("SNo"), True, Nothing, WipData)
                For Each Table As String In New String() {"Charge", "Facility", "ChangeFacility"}
                    WipData.Tables.Remove(Table)
                    WipData.Tables.Add(RetData.Tables(Table).Copy())
                Next
            End Using
        End Using
        Return New RIAResult With {.ResultBoolean = True}
    End Function
    Private Function UpdateWipData(ByVal drData As DataRow, ByRef WipData As DataSet) As RIAResult
        Dim result As New RIAResult With {.ResultBoolean = True}
        'API 測試報告 236 沒有沒有修改到實際派工服務區代碼 By Kin For Quintina 2019/11/18
        If drData.Table.Columns.Contains("WorkServCode") Then
            If Not String.IsNullOrEmpty(drData("WorkServCode").ToString) Then
                WipData.Tables("Wip").Rows(0)("WorkServCode") = drData.Item("WorkServCode")
            End If
        End If
        '#7355 2016.12.01 by Corey 因應TBC的需求，於API-236增加參數Note、NoteType。 
        '     需處理若NoteType=0，則覆蓋。NoteType=1，則前加。NoteType=2，則後加。 
        If drData.Table.Columns.Contains("NoteType") AndAlso drData.Table.Columns.Contains("Note") Then
            '該兩個欄位都存在 才需要填 NOTE。
            If drData.IsNull("NoteType") Then Return result '沒有指定0=覆蓋 1=補前面 2=補後面 規則
            If Not ",0,1,2,".Contains(String.Format(",{0},", drData("NoteType").ToString)) Then
                'NOTETYPE 沒有設定在 0,1,2內 回應設定錯誤
                Return New RIAResult With {.ResultBoolean = False, .ErrorCode = BillingAPILanguage.OtherErrorCode, .ErrorMessage = BillingAPILanguage.OtherErrorMessage & String.Format("({0})", Lang.NoteTypeErr)}
            End If
            Dim intNoteType As Integer = drData("NoteType")
            Dim strNewNote As String = WipData.Tables("Wip").Rows(0)("Note").ToString
            Select Case intNoteType
                Case 0
                    strNewNote = drData("Note")
                Case 1
                    strNewNote = String.Format("{0}{1}", drData("Note"), strNewNote)
                Case 2
                    strNewNote = String.Format("{0}{1}", strNewNote, drData("Note"))
            End Select
            WipData.Tables("Wip").Rows(0)("Note") = strNewNote
        End If
       
        Return result
    End Function
    Private Function ChkCanResv(WipData As System.Data.DataSet, ResvTime As DateTime, inRow As DataRow) As RIAResult
        Using bll As New Validate(LoginInfo, DAO)
            Dim result As RIAResult = Nothing
            '檢核預約時間是否可以改約
            Dim WipRow As DataRow = WipData.Tables("Wip").Rows(0)
            Using WorkCode As DataTable = DAO.ExecQry(_DAL.GetWorkCode(), New Object() {WipRow.Item("PRCode")})
                Dim MCode As Integer = CableSoft.BLL.Utility.Utility.ConvertDBNullToInteger(WorkCode.Rows(0).Item("GroupNo"))
                Dim Resvdatebefore As Integer = CableSoft.BLL.Utility.Utility.ConvertDBNullToInteger(WorkCode.Rows(0).Item("Resvdatebefore"))
                Dim WorkUnit As Decimal = CableSoft.BLL.Utility.Utility.ConvertDBNullToDecimal(WorkCode.Rows(0).Item("WorkUnit"))
                '2016.09.20 傳參順序錯誤調正
                'result = bll.ChkCanResv(WipRow.Item("PRCode"), WipRow.Item("WorkServCode"), MCode, WipRow.Item("ServiceType"), ResvTime, WipRow.Item("AcceptTime"), WipRow.Item("ResvTime"), Resvdatebefore, WorkUnit, True)

                'WipRow.Item("WorkServCode")
                '#8605 Accoding to API'WorkServCode as data By Kin 2020/05/14
                result = bll.ChkCanResv(inRow.Item("WorkServCode"), WipRow.Item("PRCode"), MCode, WipRow.Item("ServiceType"), ResvTime, WipRow.Item("AcceptTime"), WipRow.Item("ResvTime"), Resvdatebefore, WorkUnit, True)
                If result.ResultBoolean = False Then
                    Return result
                End If
            End Using
            Return New RIAResult With {.ResultBoolean = True}
        End Using
    End Function
    Private Function GetReturnData(WipData As DataSet) As DataSet
        Dim RetData As New DataSet With {.DataSetName = "DataSet"}
        Dim RetTable As New DataTable With {.TableName = "AMT"}
        RetTable.Columns.Add(New DataColumn With {.ColumnName = "Amount", .DataType = GetType(String)})
        RetTable.Rows.Add(RetTable.NewRow())
        Using GetAmt As DataTable = DAO.ExecQry(_DAL.GetSNoTotalAmount(), New Object() {WipData.Tables("Wip").Rows(0).Item("SNo")})
            RetTable.Rows(0).Item("Amount") = GetAmt.Rows(0).Item(0)
        End Using
        RetData.Tables.Add(RetTable)
        Return RetData
    End Function
#Region "IDisposable Support"
    Private disposedValue As Boolean
    Public Sub New()

    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo)
        MyBase.New(LoginInfo)

    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo, ByVal DBConnection As System.Data.Common.DbConnection)
        MyBase.New(LoginInfo, DBConnection)

    End Sub
    Public Sub New(ByVal LoginInfo As CableSoft.BLL.Utility.LoginInfo, ByVal DAO As CableSoft.Utility.DataAccess.DAO)
        MyBase.New(LoginInfo, DAO)

    End Sub
    ' 偵測多餘的呼叫

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: 處置 Managed 狀態 (Managed 物件)。
            End If
            Try
                If _DAL IsNot Nothing Then
                    _DAL.Dispose()
                End If
                If MyBase.MustDispose AndAlso DAO IsNot Nothing Then
                    DAO.Dispose()
                End If
                If Lang IsNot Nothing Then
                    Lang.Dispose()
                    Lang = Nothing
                End If
            Catch ex As Exception
            End Try
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

'填入設定檔資料
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('236', 'APIID', '命令識別碼', 0, 1, 'Main', -1, 1);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('236', 'Compcode', '公司別', 1, 1, 'Main', -1, 2);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('236', 'Caller', '呼叫來源', 0, 1, 'Main', -1, 3);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DTATABLENAME, ERRORCODE, ORD)
' Values
'   ('236', 'Seqno', '來源識別碼', 0, 1, 'Main', -1, 4);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('236', 'SNo', '工單單號', 0, 1, 'Main', -1, 5);
'Insert into SO1114B
'   (APIID, FIELDNAME, HEADNAME, DATATYPE, MUSTBE, DATATABLENAME, ERRORCODE, ORD)
' Values
'   ('236', 'ResvTime', '預約時間', 3, 1, 'Main', -1, 6);
'================================================================================================================================
'SO1114B.ResvTime 預約時間 ChkSQLQuery設定內容
'Select -101 ErrorCode,'查無單號' ErrorMsg From Dual
'Where (Select Count(*) From SO009 Where SNo = '[SNo]') = 0
'Union All
'Select -151 ErrorCode,'該單號已結案, 不可做改約' ErrorMsg From SO009
'Where SNo = '[SNo]' And (ReturnCode is not null Or FinTime is not null)
'Union All
'Select -152 ErrorCode,'該單號已線上回報, 不可改約' ErrorMsg From SO009
'Where SNo = '[SNo]' And (CallOkTime is not null)
'Union All
'Select -153 ErrorCode,'部份費用已入暫收款,不可改約' ErrorMsg From SO033
'Where BillNo = '[SNo]' And UCCode in (Select CodeNo From CD013 Where (RefNo in (3,7,8) Or PayOk = 1) And StopFlag = 0) And Rownum = 1
'Union All
'Select -153 ErrorCode,'部份費用已入實收,不可改約' ErrorMsg From SO033
'Where BillNo = '[SNo]' And UCCode is null And Rownum = 1
'Union All
'Select -154 ErrorCode,'該單號有結清資料, 不可改約' ErrorMsg From SO004D A,SO009 B
'Where A.SNo = B.SNo And A.SNo = '[SNo]' And A.Delete003Citem is not null 
'And PRCode In (Select CodeNo From CD007 Where RefNo in (1,2,5,6,10))
'Union All
'Select -103 ErrorCode,'此為訂單產生之工單, 請由訂單機制處理' ErrorMsg From SO009
'Where SNo = '[SNo]' And (OrderNo is not null)
