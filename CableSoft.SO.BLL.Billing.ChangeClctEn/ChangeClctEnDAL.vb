Imports CableSoft.BLL.Utility

Public Class ChangeClctEnDAL
    Inherits DALBasic
    Implements IDisposable

    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub

    Friend Function GetCompCode(ByVal GroupId As String) As String
        If GroupId = "0" AndAlso 1 = 0 Then
            Return "Select A.CodeNo ,A.Description From CD039 A Order By CodeNo"
        End If
        Return String.Format("Select distinct A.CodeNo ,A.Description " & _
                             " From CD039 A,SO026 B  " & _
                             " Where Instr(',' ||B.CompStr|| ',' , ',' ||A.CodeNo|| ',') > 0 And UserId = {0}0 Order By CodeNO", Sign)
    End Function
    Friend Function GetClctEn() As String
        'Dim aSQL As String = Nothing
        'aSQL = "select '' empno,'' empname,''  From CM003 where rownum = 1 " & _
        '    "union all " & _
        '    " select * from (" & _
        '    "Select empno,empname From CM003 Where Nvl(StopFlag,0) = 0 Order By EmpNo)"
        'Return aSQL
        Return "Select EmpNo as CODENO,EmpName as Description From CM003 Where Nvl(StopFlag,0) = 0 Order By EmpNo"
    End Function
    Friend Function GetServiceType() As String
        Return "Select * from CD046 Order by CodeNo"
    End Function
    Friend Function GetStrtCode() As String
        Return "Select CodeNo,Description,RefNo From CD017 Where Nvl(StopFlag,0) = 0 Order By CodeNo"
    End Function
    Friend Function GetMduId() As String
        Return "Select MduId CodeNo,Name Description  From SO017 Order by MduId"
    End Function
    Friend Function Execute(ByVal RefNo As Integer,
                            ByVal GroupByStr As Integer,
                            ByVal ModifyType As Integer,
                            ByVal tbPara As DataTable,
                             ByVal rw As DataRow) As String
        Dim aWhere As String = " Where UCCode not in (Select CodeNO From CD013 Where RefNo in (3,7,8) Or PayOK = 1) "
        Dim tbName As String = "SO032"
        Dim aField As String = Nothing
        Dim aSQL As String = Nothing
        With tbPara.Rows(0)
            If (Not DBNull.Value.Equals(.Item("ClctYM1"))) AndAlso
           (Not String.IsNullOrEmpty(.Item("ClctYM1"))) Then
                aWhere = aWhere & " AND  ClctYM >= " & .Item("ClctYM1")
            End If
            If (Not DBNull.Value.Equals(.Item("ClctYM2"))) AndAlso
                (Not String.IsNullOrEmpty(.Item("ClctYM2"))) Then
                aWhere = aWhere & " AND  ClctYM <= " & .Item("ClctYM1")
            End If
            If (Not DBNull.Value.Equals(.Item("ServiceType"))) AndAlso
                (Not String.IsNullOrEmpty(.Item("ServiceType"))) Then
                aWhere = String.Format("{0} AND SERVICETYPE IN ({1})", aWhere, .Item("ServiceType"))
            End If
            If (Not DBNull.Value.Equals(.Item("CreateTime1"))) AndAlso
                (Not String.IsNullOrEmpty(.Item("CreateTime1"))) Then
                aWhere = String.Format("{0} AND CreateTime >= TO_DATE('{1}','YYYYMMDDHH24MISS')",
                                     aWhere, .Item("CreateTime1").ToString.Replace("/", "").Replace(":", "").Replace(" ", ""))
            End If
            If (Not DBNull.Value.Equals(.Item("CreateTime2"))) AndAlso
                (Not String.IsNullOrEmpty(.Item("CreateTime2"))) Then
                aWhere = String.Format("{0} AND CreateTime <= TO_DATE('{1}','YYYYMMDDHH24MISS')",
                                     aWhere, .Item("CreateTime2").ToString.Replace("/", "").Replace(":", "").Replace(" ", ""))
            End If
            'If (Not DBNull.Value.Equals(.Item("ClctEn"))) AndAlso
            '    (Not String.IsNullOrEmpty(.Item("ClctEn").ToString)) Then
            '    aWhere = String.Format("{0} AND CLCTEN IN ({1}) ", aWhere, .Item("ClctEn"))
            'End If
            If (Not DBNull.Value.Equals(.Item("StrtCode"))) AndAlso
                (Not String.IsNullOrEmpty(.Item("StrtCode").ToString)) Then
                aWhere = String.Format("{0} AND StrtCode IN ({1})", aWhere, .Item("StrtCode"))
            End If
            If (Not DBNull.Value.Equals(.Item("MduId"))) AndAlso
                (Not String.IsNullOrEmpty(.Item("MduId").ToString)) Then
                aWhere = String.Format("{0} AND MduId IN ({1})", aWhere, .Item("MduId"))
            End If
            If (Not DBNull.Value.Equals(.Item("BillNo1"))) AndAlso
                (Not String.IsNullOrEmpty(.Item("BillNo1").ToString)) Then
                aWhere = String.Format("{0} AND BILLNO >= '{1}'", aWhere, .Item("BillNo1"))
            End If
            If (Not DBNull.Value.Equals(.Item("BillNo2"))) AndAlso
                (Not String.IsNullOrEmpty(.Item("BillNo2").ToString)) Then
                aWhere = String.Format("{0} AND BILLNO <= '{1}'", aWhere, .Item("BillNo2"))
            End If
            If (Not DBNull.Value.Equals(.Item("CustId"))) AndAlso
                (Not String.IsNullOrEmpty(.Item("CustId").ToString)) Then
                aWhere = String.Format("{0} AND CUSTID IN ({1})", aWhere, .Item("CUSTID"))
            End If

        End With
       
        Select Case GroupByStr
            Case 0
                aField = "StrtCode"
            Case 1
                aField = "MduId"
            Case 2
                aField = "BillNo"
            Case 3
                aField = "CustId"
        End Select
        If RefNo = 1 Then
            tbName = "SO033"
        End If
        If DBNull.Value.Equals(rw.Item("ClctEn")) OrElse (String.IsNullOrEmpty(rw.Item("ClctEn").ToString)) Then
            aWhere = aWhere & " AND ClctEn is Null "
        Else
            aWhere = String.Format(aWhere & " AND ClctEn='{0}' ", rw.Item("ClctEn"))
        End If
        Select Case ModifyType
            Case 0
                aSQL = String.Format("Update " & tbName & " SET ClctEn ={0}0,ClctName={0}1,OldClctEn={0}2, " & _
                                        " OldClctName = {0}3,UpdTime = {0}4 , UpdEn = {0}5, NewUpdTime = {0}6 " & _
                                        aWhere & " AND " & aField & "={0}7", Sign)
            Case 1
                aSQL = String.Format("Update " & tbName & " SET OldClctEn={0}0, " & _
                                        " OldClctName = {0}1,UpdTime = {0}2 , UpdEn = {0}3, NewUpdTime = {0}4 " & _
                                        aWhere & " AND " & aField & "={0}5", Sign)
            Case 2

                aSQL = String.Format("Update " & tbName & " SET ClctEn ={0}0,ClctName={0}1, " & _
                                        " UpdTime = {0}2 , UpdEn = {0}3, NewUpdTime = {0}4 " & _
                                        aWhere & " AND " & aField & "={0}5", Sign)
        End Select
        Return aSQL
        'Return String.Format("Update " & tbName & " SET ClctEn ={0}0,ClctName={0}1,OldClctEn={0}2, " & _
        '                                " OldClctName = {0}3,UpdTime = {0}4 , UpdEnd = {0}5 " & _
        '                                aWhere & " AND " & aField & "={0}6", Sign)

    End Function

    Friend Function Execute(ByVal rw As DataRow) As String
        Dim aWhere As String = Nothing
        If DBNull.Value.Equals(rw.Item("ClctEn")) OrElse (String.IsNullOrEmpty(rw.Item("ClctEn").ToString)) Then
            aWhere = " AND ClctEn is Null "
        Else
            aWhere = String.Format(" AND ClctEn='{0}' ", rw.Item("ClctEn"))
        End If

        Return String.Format("Update SO033 Set ClctEn = {0}0,ClctName={0}1 ,OldClctEn = {0}2, " & _
                                    " OldClctName= {0}3,UpdTime = {0}4, UpdEn = {0}5, NewUpdTime = {0}6 " & _
                                    " Where StrtCode = {0}7 " & aWhere, Sign)
    End Function
    Friend Function GetGroupData(ByVal RefNo As Integer, ByVal tbWhere As DataTable) As String
        Dim aWhere As String = " Where A.UCCode not in (Select CodeNO From CD013 Where RefNo in (3,7,8) Or PayOK = 1)"
        Dim aSQL As String = Nothing
        Dim tbName As String = Nothing
        tbName = "SO032"
        If RefNo = 1 Then
            tbName = "SO033"
        End If
        'tbPara.Columns.Add(New DataColumn("ClctYM1", GetType(String)))
        'tbPara.Columns.Add(New DataColumn("ClctYM2", GetType(String)))
        'tbPara.Columns.Add(New DataColumn("CreateTime1", GetType(String)))
        'tbPara.Columns.Add(New DataColumn("CreateTime2", GetType(String)))
        'tbPara.Columns.Add(New DataColumn("BillNo1", GetType(String)))
        'tbPara.Columns.Add(New DataColumn("BillNo2", GetType(String)))
        'tbPara.Columns.Add(New DataColumn("ServiceType", GetType(String)))
        'tbPara.Columns.Add(New DataColumn("MduId", GetType(String)))
        'tbPara.Columns.Add(New DataColumn("CustId", GetType(String)))
        'tbPara.Columns.Add(New DataColumn("ServiceType", GetType(String)))
        'tbPara.Columns.Add(New DataColumn("StrtCode", GetType(String)))
        'tbPara.Columns.Add(New DataColumn("ClctEn", GetType(String)))
        With tbWhere.Rows(0)
            If (Not DBNull.Value.Equals(.Item("ClctYM1"))) AndAlso
                (Not String.IsNullOrEmpty(.Item("ClctYM1"))) Then
                aWhere = aWhere & " AND  A.ClctYM >= " & .Item("ClctYM1")
            End If
            If (Not DBNull.Value.Equals(.Item("ClctYM2"))) AndAlso
                (Not String.IsNullOrEmpty(.Item("ClctYM2"))) Then
                aWhere = aWhere & " AND  A.ClctYM <= " & .Item("ClctYM1")
            End If
            If (Not DBNull.Value.Equals(.Item("ServiceType"))) AndAlso
                (Not String.IsNullOrEmpty(.Item("ServiceType"))) Then
                aWhere = String.Format("{0} AND A.SERVICETYPE IN ({1})", aWhere, .Item("ServiceType"))
            End If
            If (Not DBNull.Value.Equals(.Item("CreateTime1"))) AndAlso
                (Not String.IsNullOrEmpty(.Item("CreateTime1"))) Then
                aWhere = String.Format("{0} AND A.CreateTime >= TO_DATE('{1}','YYYYMMDDHH24MISS')",
                                     aWhere, .Item("CreateTime1").ToString.Replace("/", "").Replace(":", "").Replace(" ", ""))
            End If
            If (Not DBNull.Value.Equals(.Item("CreateTime2"))) AndAlso
                (Not String.IsNullOrEmpty(.Item("CreateTime2"))) Then
                aWhere = String.Format("{0} AND A.CreateTime <= TO_DATE('{1}','YYYYMMDDHH24MISS')",
                                     aWhere, .Item("CreateTime2").ToString.Replace("/", "").Replace(":", "").Replace(" ", ""))
            End If
            If (Not DBNull.Value.Equals(.Item("ClctEn"))) AndAlso
                (Not String.IsNullOrEmpty(.Item("ClctEn").ToString)) Then
                aWhere = String.Format("{0} AND A.CLCTEN IN ({1}) ", aWhere, .Item("ClctEn"))
            End If
            If (Not DBNull.Value.Equals(.Item("StrtCode"))) AndAlso
                (Not String.IsNullOrEmpty(.Item("StrtCode").ToString)) Then
                aWhere = String.Format("{0} AND A.StrtCode IN ({1})", aWhere, .Item("StrtCode"))
            End If
            If (Not DBNull.Value.Equals(.Item("MduId"))) AndAlso
                (Not String.IsNullOrEmpty(.Item("MduId").ToString)) Then
                aWhere = String.Format("{0} AND A.MduId IN ({1})", aWhere, .Item("MduId"))
            End If
            If (Not DBNull.Value.Equals(.Item("BillNo1"))) AndAlso
                (Not String.IsNullOrEmpty(.Item("BillNo1").ToString)) Then
                aWhere = String.Format("{0} AND A.BILLNO >= '{1}'", aWhere, .Item("BillNo1"))
            End If
            If (Not DBNull.Value.Equals(.Item("BillNo2"))) AndAlso
                (Not String.IsNullOrEmpty(.Item("BillNo2").ToString)) Then
                aWhere = String.Format("{0} AND A.BILLNO <= '{1}'", aWhere, .Item("BillNo2"))
            End If
            If (Not DBNull.Value.Equals(.Item("CustId"))) AndAlso
                (Not String.IsNullOrEmpty(.Item("CustId").ToString)) Then
                aWhere = String.Format("{0} AND A.CUSTID IN ({1})", aWhere, .Item("CUSTID"))
            End If
            Select Case Integer.Parse(.Item("GroupKind"))
                Case 0
                    aSQL = String.Format(" Select A.ClctEn,A.ClctName,A.StrtCode GroupCode,B.Description GroupName," & _
                        " Count(*) Count, '' NewClctEn, '' NewClctName " & _
                        " From {0} A Join CD017 B " & _
                        " On (A.StrtCode = B.CodeNo) {1} " & _
                        " Group by A.ClctEn,A.ClctName,A.StrtCode,B.Description " & _
                        " Order By A.ClctEn,A.StrtCode", tbName, aWhere)
                Case 1
                    aSQL = String.Format(" Select A.ClctEn,A.ClctName,A.MduId  GroupCode,B.Name GroupName," & _
                       " Count(*) Count, '' NewClctEn, '' NewClctName " & _
                       " From {0} A Join SO017 B " & _
                       " On (A.MduId = B.MduId) {1} " & _
                       " Group by A.ClctEn,A.ClctName,A.MduId,B.Name " & _
                       " Order By A.ClctEn,A.MduId", tbName, aWhere)
                Case 2
                    aSQL = String.Format(" Select A.ClctEn,A.ClctName,A.BillNo  GroupCode,'單據編號' GroupName," & _
                      " Count(*) Count, '' NewClctEn, '' NewClctName " & _
                      " From {0} A {1} " & _
                      " Group by A.ClctEn,A.ClctName,A.BillNo " & _
                      " Order By A.ClctEn,A.BillNo", tbName, aWhere)
                Case 3
                    aSQL = String.Format(" Select A.ClctEn,A.ClctName,A.CustId  GroupCode,B.CustName GroupName," & _
                      " Count(*) Count, '' NewClctEn, '' NewClctName " & _
                      " From {0} A Join SO001 B " & _
                      " On (A.CustId = B.CustId) {1} " & _
                      " Group by A.ClctEn,A.ClctName,A.CustId,B.CustName " & _
                      " Order By A.ClctEn,A.CustId", tbName, aWhere)
            End Select
        End With
        Return aSQL
    End Function

    Friend Function GetClctStrtGroupData(ByVal ClctEnStr As String, ByVal StrtCodeStr As String) As String
        Dim aWhere As String = ""
        If Not String.IsNullOrEmpty(ClctEnStr) Then
            aWhere = String.Format(" WHERE A.CLCTEN IN ( {0} ) ", ClctEnStr)
        End If
        If Not String.IsNullOrEmpty(StrtCodeStr) Then
            If Not String.IsNullOrEmpty(aWhere) Then
                aWhere = String.Format(aWhere & " AND  A.StrtCode In ( {0} ) ", StrtCodeStr)
            Else
                aWhere = String.Format(" Where A.StrtCode In ( {0} )", StrtCodeStr)
            End If
        End If

        Return String.Format("Select A.ClctEn,A.ClctName,A.StrtCode,B.Description StrtName," & _
                    "Count(*) Count ,'' NewClctEn,'' NewClctName " & _
                    " From SO033 A Join CD017 B On (A.StrtCode = B.CodeNo)  {0} AND  rownum<=10 " & _
                    " Group by A.ClctEn,A.ClctName,A.StrtCode,B.Description " & _
                    " Order By A.ClctEn,A.StrtCode ", aWhere)

        Return String.Format("Select A.ClctEn,A.ClctName,A.StrtCode,B.Description StrtName," & _
                             " Count(*) Count ,'' NewClctEn,'' NewClctName " & _
                             " From SO032 A Join CD017 B On (A.StrtCode = B.CodeNo) {0} " & _
                             " Group by A.ClctEn,A.ClctName,A.StrtCode,B.Description " & _
                             " Order By A.ClctEn,A.StrtCode", aWhere)
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
