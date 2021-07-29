Imports CableSoft.BLL.Utility

Public MustInherit Class IntroMediaDAL
    Inherits DALBasic
    Implements IDisposable

    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub

    Friend Function GetWhere(ByVal MediaRefNo As Int32, ByVal Search1 As String,
                              ByVal Search2 As String) As String
        Dim aRet As String = String.Empty
        Dim strXtra2 As String = " 1=1 "
        Dim aKeyWord As String = String.Empty
        Select Case MediaRefNo
            Case 1
                aKeyWord = "CustName"
            Case 2
                aKeyWord = "EmpName"
            Case 3
                aKeyWord = "NameP"
            Case Else
                aKeyWord = "CustName"
        End Select

        Try
            If Not String.IsNullOrEmpty(Search1) Then
                If MediaRefNo <> 1 Then
                    If MediaRefNo = 2 Then
                        aRet = String.Format(" WHERE {0}  Like '%{1}%' AND STOPFLAG <> 1",
                                             aKeyWord, Search1)
                    Else
                        aRet = String.Format(" WHERE {0} Like '%{1}%' ", aKeyWord, Search1)
                    End If
                Else
                    aRet = String.Format(" WHERE {0} Like '{1}%'", aKeyWord, Search1)
                End If
            End If

            If Not String.IsNullOrEmpty(Search2) Then
                Select Case MediaRefNo
                    Case 1
                        strXtra2 = String.Format(" Tel1 = '{0}' OR " &
                                   " Tel2 = '{1}' OR " &
                                   " Tel3 = '{2}'", Search2, Search2, Search2)
                    Case 2
                        strXtra2 = String.Format(" EmpNo = '{0}' AND STOPFLAG <> 1", Search2)
                    Case 3
                        strXtra2 = String.Format(" TelH = '{0}' OR " &
                                   " TelO = '{1}' OR " &
                                   " TelM = '{2}' OR " &
                                   " TelB = '{3}'", Search2, Search2, Search2, Search2)
                End Select

                If Not String.IsNullOrEmpty(aRet) Then
                    aRet = aRet & " AND " & strXtra2
                Else
                    aRet = " WHERE " & strXtra2
                End If

            End If

            Return aRet
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Friend Function GetkeyCodeSearchSQL(ByVal MediaRefNo As Integer, ByVal searchWord As String) As String
        Dim aSQL As String = Nothing
        Select Case MediaRefNo
            Case 1
                aSQL = String.Format("Select CustName as Description ,CustId as CodeNo From SO001 Where CustId = {0}", searchWord)
            Case 2
                aSQL = String.Format("SELECT EmpName as Description, EmpNo as CodeNo FROM CM003 WHERE EmpNo = '{0}'", searchWord)
            Case 3
                aSQL = String.Format("SELECT NameP as Description, IntroID as CodeNo FROM SO013 WHERE IntroID = '{0}'", searchWord)
            Case Else
                aSQL = String.Format("Select CustName as Description ,CustId as CodeNo From SO001 Where CustId = {0}", searchWord)
        End Select
        Return aSQL
    End Function
    Friend Function GetIntroId(ByVal MediaRefNo As Integer) As String
        Dim aRet As String = String.Empty
        Select Case MediaRefNo
            Case 1
                aRet = String.Format("Select CustId as CodeNo ,CustName as Description From SO001 " &
                                     " Where CustId = {0}0", Sign)
            Case 2
                aRet = "Select EmpNo As CodeNo, EmpName As Description From CM003 " &
                    " Where Nvl(StopFlag,0) = 0"
            Case 3
                aRet = "Select NameP As Description, IntroID As CodeNo FROM SO013"
        End Select
        Return aRet
    End Function
    Friend Function GetIntroData(ByVal MediaRefNo As Integer) As String
        Dim aRet As String = String.Empty
        Select Case MediaRefNo
            Case 1
                aRet = "Select CustId as CodeNo ,CustName as Description From SO001 "

            Case 2
                aRet = "Select EmpNo As CodeNo, EmpName As Description From CM003 "

            Case 3
                aRet = "Select NameP As Description, IntroID As CodeNo FROM SO013 "
            Case Else
                aRet = "Select CustId as CodeNo ,CustName as Description From SO001 "
        End Select
        Return aRet
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
