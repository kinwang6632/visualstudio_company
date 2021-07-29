Imports CableSoft.BLL.Utility
Public Class MFCodeDAL
    Inherits DALBasic
    Implements IDisposable
    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Function GetCompCode(ByVal GroupId As String, ByVal strCD039 As String, ByVal strSO026 As String) As String
        If GroupId = "0" And 1 = 0 Then
            Return "Select A.CodeNo ,A.Description From " & strCD039 & " A Order By CodeNo"
        End If
        Return String.Format("Select distinct A.CodeNo ,A.Description " & _
                             " From " & strCD039 & " A," & strSO026 & " B  " & _
                             " Where Instr(',' ||B.CompStr|| ',' , ',' ||A.CodeNo|| ',') > 0 " & _
                             " And UserId = {0}0 Order By CodeNO", Sign)
    End Function
    Friend Function GetCD011Sechema() As String
        Return "Select * From CD011 Where 1 = 0 "
    End Function
    Friend Function UpdateCD011() As String
        Return String.Format("Update CD011 Set Description={0}0, " & _
                                                "RefNo = {0}1,UpdTime={0}2,UpdEn={0}3, " & _
                                                "ServiceType={0}4,StopFlag={0}5 " & _
                                                " Where CodeNo = {0}6", Sign)
    End Function
    Friend Function DeleteCD011A() As String
        Return String.Format("Delete CD011A Where MFCode= {0}0 ", Sign)
    End Function
    Friend Function InsertCD011A() As String
        Return String.Format("Insert into CD011A (MFCode,CodeNo,Description,ServiceType ) " & _
                                        " Values ({0}0,{0}1,{0}2,{0}3)", Sign)
    End Function

    Friend Function InsertCD011() As String
        Return String.Format("Insert Into CD011 (CodeNo,Description," & _
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
    Friend Function QueryCD011Code() As String
        Return String.Format("Select * From CD011 Where CodeNo = {0}0", Sign)
    End Function
    Friend Function QueryCD011ACode() As String
        Return String.Format("Select * From CD011A Where MFCode = {0}0", Sign)
    End Function
    Friend Function QueryMasterExists() As String
        Return String.Format("Select Count(*) From CD011 Where CodeNo = {0}0", Sign)
    End Function
    Friend Function QueryDetailExists() As String
        Return String.Format("Select Count(*) From CD011A Where MFCode = {0}0 And CodeNo = {0}1", Sign)
    End Function
    Friend Function DeleteCD011() As String
        Return String.Format("Delete CD011 Where CodeNo = {0}0", Sign)
    End Function
    Friend Function DeleteCD011ACode() As String
        Return String.Format("Delete CD011A Where MFCode = {0}0 And CodeNo = {0}1", Sign)
    End Function


    Friend Function QueryCD011B(ByVal ServiceType As String) As String
        If String.IsNullOrEmpty(ServiceType) Then
            Return "Select CodeNo,Description From CD011B Where Nvl(StopFlag,0)=0 Order By CodeNo "
        Else
            Return String.Format("Select CodeNo,Description From CD011B " & _
                                            " Where Nvl(StopFlag,0)=0  And ( ServiceType ='{0}' Or ServiceType Is Null ) " & _
                                            " Order By CodeNo ", ServiceType)
        End If

    End Function
    Friend Function GetMaxCode() As String
        Return "select max(codeno)+1 MaxCode From CD011"
    End Function
    Friend Function GetSO041() As String
        Return String.Format("Select  Nvl(AutoGetCode,0) AutoGetCode From SO041  Where CompCode = {0}0", Sign)
    End Function
    Friend Function GetServiceType() As String
        Return String.Format("Select CodeNo,Description From CD046 Order by CodeNo")
    End Function
    Friend Function GetCD011A() As String
        Return "Select MFCode,CodeNo,Description From CD011A Order by CodeNo"
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
