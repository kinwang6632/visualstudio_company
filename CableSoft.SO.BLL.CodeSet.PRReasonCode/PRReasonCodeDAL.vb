Imports CableSoft.BLL.Utility

Public Class PRReasonCodeDAL
    Inherits DALBasic
    Implements IDisposable
    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Function GetCompCode(ByVal GroupId As String, ByVal strCD039 As String, ByVal strSO026 As String) As String
        If GroupId = "0" Then
            Return "Select A.CodeNo ,A.Description From " & strCD039 & " A Order By CodeNo"
        End If
        Return String.Format("Select distinct A.CodeNo ,A.Description " & _
                             " From " & strCD039 & " A," & strSO026 & " B  " & _
                             " Where Instr(',' ||B.CompStr|| ',' , ',' ||A.CodeNo|| ',') > 0 " & _
                             " And UserId = {0}0 Order By CodeNO", Sign)
    End Function
    Friend Function GetServiceType() As String
        Return String.Format("Select CodeNo,Description From CD046 Order by CodeNo")
    End Function
    Friend Function GetCD014A() As String
        Return "Select A.CodeNo,Description,B.ReasonCode,B.ReasonDescCode From CD014A A ,CD014B B Where A.CodeNo=B.ReasonDescCode  Order by A.CodeNo"
    End Function
    Friend Function GetMaxCode() As String
        Return "select max(codeno)+1 MaxCode From CD014"
    End Function
    Friend Function QueryCD014A(ByVal ServiceType As String) As String
        If String.IsNullOrEmpty(ServiceType) Then
            Return "Select CodeNo,Description From CD014A Where Nvl(StopFlag,0)=0 Order By CodeNo "
        Else
            Return String.Format("Select CodeNo,Description From CD014A " & _
                                            " Where Nvl(StopFlag,0)=0  And ( ServiceType ='{0}' Or ServiceType Is Null ) " & _
                                            " Order By CodeNo ", ServiceType)
        End If

    End Function
    Friend Function GetSO041() As String
        Return String.Format("Select  Nvl(AutoGetCode,0) AutoGetCode From SO041  Where CompCode = {0}0", Sign)
    End Function
    Friend Function GetCD014Sechema() As String
        Return "Select * From CD014 Where 1 = 0 "
    End Function
    Friend Function UpdateCD014() As String
        Return String.Format("Update CD014 Set Description={0}0, " & _
                                                "RefNo = {0}1,UpdTime={0}2,UpdEn={0}3, " & _
                                                "ServiceType={0}4,StopFlag={0}5 " & _
                                                " Where CodeNo = {0}6", Sign)
    End Function
    Friend Function InsertCD014() As String
        Return String.Format("Insert Into CD014 (CodeNo,Description," & _
                                              "RefNo,UpdTime,UpdEn,ServiceType,StopFlag ) " & _
                                              " Values ( {0}0, " & _
                                               " {0}1,{0}2,{0}3, " & _
                                               "{0}4,{0}5,{0}6 )", Sign)
        'Return String.Format("Insert into CD011 (CodeNo,Description" & _
        '                                    " ) " & _
        '                                     " Values ( {0}0, " & _
        '                                      " {0}1)", Sign)
        'Return "insert into CD011 (CodeNo,Description) Values ( 969,'TEST')"
    End Function
    Friend Function DeleteCD014() As String
        Return String.Format("Delete CD014 Where CodeNo = {0}0", Sign)
    End Function
    Friend Function DeleteCD014A() As String
        Dim result As String = Nothing

        result = "Delete CD014A Where Exists (Select * From CD014B  " & _
                        " Where CD014A.CodeNo =CD014B.ReasonDescCode And ReasonCode = {0}0 ) "
        Return String.Format(result, Sign)

    End Function
    Friend Function InsertCD014A() As String
        Return String.Format("Insert into CD014A (CodeNo,Description,ServiceType,UpdTime,UpdEn ) " & _
                                        " Values ({0}0,{0}1,{0}2,{0}3,{0}4)", Sign)
    End Function
    Friend Function DeleteCD014B() As String
        Dim result As String = Nothing
        'result = "Delete CD014B Where Exists (Select * From CD014B  " & _
        '                " Where CD014A.CodeNo =CD014B.ReasonDescCode And ReasonCode = {0}0 ) "
        result = "Delete CD014B Where ReasonCode = {0}0"
        Return String.Format(result, Sign)
    End Function
    Friend Function InsertCD014B() As String
        Dim result As String = Nothing
        result = String.Format("Insert Into CD014B (ReasonCode,ReasonDescCode ) Values " & _
                        "( {0}0,{0}1)", Sign)
        Return result
    End Function
    Friend Function QueryCD014Code() As String
        Return String.Format("Select * From CD014 Where CodeNo = {0}0", Sign)
    End Function
    Friend Function QueryCD014BCode() As String
        Return String.Format("Select * From CD014B Where ReasonCode = {0}0", Sign)
    End Function
    Friend Function QueryMasterExists() As String
        Return String.Format("Select Count(*) From CD014 Where CodeNo = {0}0", Sign)
    End Function
    Friend Function QueryDetailExists() As String
        Return String.Format("Select Count(*) From CD014B Where ReasonCode = {0}0 And ReasonDescCode = {0}1", Sign)
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
