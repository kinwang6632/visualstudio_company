Imports CableSoft.BLL.Utility
Imports System.Web
Imports System.Xml
Imports System.Data.Common
Public Class ManualNo
    Inherits BLLBasic
    Implements IDisposable
    Private _DAL As New ManualNoDALMultiDB(Me.LoginInfo.Provider)
    Private Language As New CableSoft.BLL.Language.SO61.ManualNoLanguage
    Private FNowDate As Date = Date.Now
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
    Public Function QueryCompCode() As DataTable
        Return DAO.ExecQry(_DAL.QueryCompCode,
                              New Object() {Me.LoginInfo.EntryId})
        
    End Function
    Private Function IsDataOK(ByVal dsSave As DataSet) As Boolean
        Dim result As Boolean = False
        result = Integer.Parse(DAO.ExecSclr(_DAL.chkDual, New Object() {dsSave.Tables(0).Rows(0).Item("BeginNum"),
                                                       dsSave.Tables(0).Rows(0).Item("EndNum"),
                                                       dsSave.Tables(0).Rows(0).Item("BeginNum"),
                                                       dsSave.Tables(0).Rows(0).Item("EndNum"), dsSave.Tables(0).Rows(0).Item("Prefix")})) = 0
        Return result
    End Function
    Private Function chkReUseOK(ByVal dsData As DataSet) As RIAResult
        Dim result As New RIAResult With {.ErrorMessage = Nothing, .ErrorCode = 0, .ResultBoolean = True}
        With dsData.Tables(0).Rows(0)
            If Integer.Parse(DAO.ExecSclr(_DAL.QueryExistData,
                                        New Object() {.Item("SEQ"), .Item("Prefix") & .Item("NewBeginNum"), LoginInfo.CompCode})) = 0 Then
                result.ResultBoolean = False
                result.ErrorCode = -1
                result.ErrorMessage = Language.NoFoundNo
                Return result
            End If

            If Integer.Parse(dsData.Tables(0).Rows(0).Item("NewBeginNum")) > Integer.Parse(dsData.Tables(0).Rows(0).Item("EndNum")) Then
                result.ResultBoolean = False
                result.ErrorCode = -2
                result.ErrorMessage = Language.ExceedBegin
                Return result
            End If

            Dim i As Integer = Integer.Parse(DAO.ExecSclr(_DAL.chkHadUse, New Object() {.Item("Prefix"), .Item("NewBeginNum"),
                                                                         .Item("Prefix"),
                                                                         .Item("EndNum"), .Item("SEQ"), LoginInfo.CompCode}))
            If i > 0 Then
                result.ResultBoolean = False
                result.ErrorCode = -3
                result.ErrorMessage = Language.hadUse
                Return result
            End If
        End With

        Return result
    End Function
    Public Function DeleteData(ByVal dsData As DataSet) As RIAResult
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        'Dim CSLog As CableSoft.SO.BLL.DataLog.DataLog = Nothing
        Dim blnAutoClose As Boolean = False
        Dim aRiaresult As New RIAResult() With {.ResultBoolean = True, .ErrorCode = 0, .ErrorMessage = Nothing}
        Dim tbSO008Log As DataTable = Nothing
        Dim isBeginTrans As Boolean = False
        aRiaresult.ResultBoolean = True
        aRiaresult.ErrorCode = 0
        If DAO.Transaction IsNot Nothing Then
            trans = DAO.Transaction
        Else
            cn.ConnectionString = Me.LoginInfo.ConnectionString
            cn.Open()
            trans = cn.BeginTransaction
            DAO.Transaction = trans
            blnAutoClose = True
        End If
        DAO.AutoCloseConn = False
        If blnAutoClose Then
            CableSoft.BLL.Utility.Utility.SetClientInfo(DAO, LoginInfo.EntryId, Language.DelClientInfo)
        End If
        Dim updNow As Date = Date.Now
        Try
            isBeginTrans = True
            With dsData.Tables(0).Rows(0)

                DAO.ExecNqry(_DAL.DeleteSO126, New Object() {
                             .Item("SEQ"), .Item("Prefix"), .Item("BeginNum"),
                             .Item("EndNum"), LoginInfo.CompCode
                             })
                DAO.ExecNqry(_DAL.DeleteSO127, New Object() {
                           .Item("SEQ"), .Item("Prefix") & .Item("BeginNum"),
                            .Item("Prefix") & .Item("EndNum"), LoginInfo.CompCode
                             })
            End With
            If blnAutoClose Then
                trans.Commit()
            End If

            aRiaresult.ResultBoolean = True
            dsData.Dispose()
            dsData = Nothing
        Catch ex As Exception
            If isBeginTrans Then
                trans.Rollback()
            End If
            aRiaresult.ErrorMessage = ex.ToString
            aRiaresult.ResultBoolean = False
            aRiaresult.ErrorCode = -99
        Finally
            If blnAutoClose Then
                CableSoft.BLL.Utility.Utility.ClearClientInfo(DAO)
                If trans IsNot Nothing Then
                    trans.Dispose()
                End If
                If cn IsNot Nothing Then
                    cn.Close()
                    cn.Dispose()
                End If
                If blnAutoClose Then
                    DAO.AutoCloseConn = True
                End If

            End If
        End Try

        Return aRiaresult
    End Function
    Public Function CanDelete(ByVal dsData As DataSet) As RIAResult
        Dim result As New RIAResult With {.ResultBoolean = True, .ErrorCode = 0, .ErrorMessage = Nothing}       
        Try
            With dsData.Tables(0).Rows(0)
                Dim i As Integer = Integer.Parse(DAO.ExecSclr(_DAL.CanDelete, New Object() {
                                                    .Item("SEQ"), .Item("Prefix") & .Item("BeginNum"),
                                                    .Item("Prefix") & .Item("EndNum"), LoginInfo.CompCode
                                                }))
                If i > 0 Then
                    result.ResultBoolean = False
                    result.ErrorCode = -1
                    result.ErrorMessage = Language.CannotDelete
                End If
            End With
        Catch ex As Exception
            result.ErrorMessage = ex.ToString
            result.ResultBoolean = False
            result.ErrorCode = -99
        Finally
            If dsData IsNot Nothing Then
                dsData.Dispose()
                dsData = Nothing
            End If
        End Try
        Return result
    End Function
    Public Function UpdNewManualNo(ByVal dsData As DataSet) As RIAResult
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        'Dim CSLog As CableSoft.SO.BLL.DataLog.DataLog = Nothing
        Dim blnAutoClose As Boolean = False
        Dim aRiaresult As New RIAResult()
        Dim tbSO008Log As DataTable = Nothing
        Dim isBeginTrans As Boolean = False
        aRiaresult.ResultBoolean = True
        aRiaresult.ErrorCode = 0
        If DAO.Transaction IsNot Nothing Then
            trans = DAO.Transaction
        Else
            cn.ConnectionString = Me.LoginInfo.ConnectionString
            cn.Open()
            trans = cn.BeginTransaction
            DAO.Transaction = trans
            blnAutoClose = True
        End If
        DAO.AutoCloseConn = False
        Dim updNow As Date = Date.Now
        CableSoft.BLL.Utility.Utility.SetClientInfo(Me.DAO, LoginInfo.EntryId, Language.ReUseClientInfo)
        Try
            isBeginTrans = True
            With dsData.Tables(0)
                For Each rw As DataRow In .Rows
                    If Not DBNull.Value.Equals(rw.Item("MANUALNO")) AndAlso Not String.IsNullOrEmpty(rw.Item("MANUALNO")) Then
                        DAO.ExecNqry(_DAL.ClearSO127, New Object() {
                                 LoginInfo.EntryName,
                                CableSoft.BLL.Utility.DateTimeUtility.GetDTString(updNow),
                                 rw.Item("MANUALNO"),
                                 LoginInfo.CompCode})
                    End If
                    

                    If Not DBNull.Value.Equals(rw.Item("NEWMANUALNO")) AndAlso Not String.IsNullOrEmpty(rw.Item("NEWMANUALNO")) Then
                        DAO.ExecNqry(_DAL.UpdSO127ManualNo, New Object() {
                                     rw.Item("CUSTID"), rw.Item("CUSTNAME"), rw.Item("CustTEL"),
                                     rw.Item("BILLNO"), rw.Item("RealDate"), LoginInfo.EntryName,
                                   CableSoft.BLL.Utility.DateTimeUtility.GetDTString(updNow),
                                     rw.Item("NEWMANUALNO"),
                                      LoginInfo.CompCode})
                        DAO.ExecNqry(_DAL.UpdBillManual, New Object() {rw.Item("NEWMANUALNO"),
                                                                       rw.Item("BillNo"), rw.Item("ITEM"), LoginInfo.CompCode})

                    End If
                Next
            End With            
            trans.Commit()
            aRiaresult.ResultBoolean = True
            dsData.Dispose()
            dsData = Nothing
        Catch ex As Exception
            If isBeginTrans Then
                trans.Rollback()
            End If
            aRiaresult.ErrorMessage = ex.ToString
            aRiaresult.ResultBoolean = False
            aRiaresult.ErrorCode = -99
        Finally
            If blnAutoClose Then
                CableSoft.BLL.Utility.Utility.ClearClientInfo(DAO)
                If trans IsNot Nothing Then
                    trans.Dispose()
                End If
                If cn IsNot Nothing Then
                    cn.Close()
                    cn.Dispose()
                End If
                If blnAutoClose Then
                    DAO.AutoCloseConn = True
                End If

            End If
        End Try

        Return aRiaresult
    End Function
    Public Function ChkNewManualNo(ByVal dsData As DataSet) As RIAResult

        Dim aRiaresult As New RIAResult() With {.ErrorMessage = Nothing, .ErrorCode = 0, .ResultBoolean = True}

        Try

            With dsData.Tables(0)
                For Each rw As DataRow In .Rows
                    If Not DBNull.Value.Equals(rw.Item("NEWMANUALNO")) AndAlso Not String.IsNullOrEmpty(rw.Item("NEWMANUALNO").ToString) Then
                        Using tbSO127 As DataTable = DAO.ExecQry(_DAL.QuerySingleSO127, New Object() {
                                                            rw.Item("NEWMANUALNO"), LoginInfo.CompCode})
                            If tbSO127.Rows.Count = 0 Then
                                aRiaresult.ResultBoolean = False
                                aRiaresult.ErrorCode = -1
                                aRiaresult.ErrorMessage = String.Format(Language.NoneSO127, rw.Item("NEWMANUALNO"))
                                Return aRiaresult
                            End If
                            If Not DBNull.Value.Equals(tbSO127.Rows(0).Item("BILLNO")) AndAlso Not String.IsNullOrEmpty(tbSO127.Rows(0).Item("BILLNO")) Then
                                aRiaresult.ResultBoolean = False
                                aRiaresult.ErrorCode = -2
                                aRiaresult.ErrorMessage = String.Format(Language.HasManualNo, rw.Item("NEWMANUALNO"), tbSO127.Rows(0).Item("BILLNO"))
                                Return aRiaresult
                            End If
                            If DBNull.Value.Equals(tbSO127.Rows(0).Item("Status")) OrElse Integer.Parse(tbSO127.Rows(0).Item("Status")) = 0 Then
                                aRiaresult.ResultBoolean = False
                                aRiaresult.ErrorCode = -3
                                aRiaresult.ErrorMessage = String.Format(Language.HadAbandon, rw.Item("NEWMANUALNO"))
                                Return aRiaresult
                            End If
                            tbSO127.Dispose()
                        End Using
                    End If

                Next

            End With


        Catch ex As Exception

            aRiaresult.ErrorMessage = ex.ToString
            aRiaresult.ResultBoolean = False
            aRiaresult.ErrorCode = -99
        Finally
            If dsData IsNot Nothing Then
                dsData.Dispose()
                dsData = Nothing
            End If
        End Try

        Return aRiaresult
    End Function
    Public Function abandonPaper(ByVal dsData As DataSet) As RIAResult
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        'Dim CSLog As CableSoft.SO.BLL.DataLog.DataLog = Nothing
        Dim blnAutoClose As Boolean = False
        Dim aRiaresult As New RIAResult()
        Dim tbSO008Log As DataTable = Nothing
        Dim isBeginTrans As Boolean = False
        aRiaresult.ResultBoolean = True
        aRiaresult.ErrorCode = 0
        If DAO.Transaction IsNot Nothing Then
            trans = DAO.Transaction
        Else
            cn.ConnectionString = Me.LoginInfo.ConnectionString
            cn.Open()
            trans = cn.BeginTransaction
            DAO.Transaction = trans
            blnAutoClose = True
        End If
        DAO.AutoCloseConn = False
        If blnAutoClose Then
            CableSoft.BLL.Utility.Utility.SetClientInfo(DAO, LoginInfo.EntryId, Language.VoidClientInfo)
        End If
        Dim updNow As Date = Date.Now
        Try
            With dsData.Tables(0)
                For Each rw As DataRow In .Rows
                    DAO.ExecNqry(_DAL.abandonPaper, New Object() {LoginInfo.EntryName,
                                    CableSoft.BLL.Utility.DateTimeUtility.GetDTString(updNow),
                                    rw.Item("PaperNum1"), rw.Item("PaperNum2"), rw.Item("SEQ"),
                                    LoginInfo.CompCode
                                 })
                Next
            End With
            trans.Commit()
            aRiaresult.ResultBoolean = True
            dsData.Dispose()
            dsData = Nothing
        Catch ex As Exception
            If isBeginTrans Then
                trans.Rollback()
            End If
            aRiaresult.ErrorMessage = ex.ToString
            aRiaresult.ResultBoolean = False
            aRiaresult.ErrorCode = -99
        Finally
            If blnAutoClose Then
                CableSoft.BLL.Utility.Utility.ClearClientInfo(DAO)
                If trans IsNot Nothing Then
                    trans.Dispose()
                End If
                If cn IsNot Nothing Then
                    cn.Close()
                    cn.Dispose()
                End If
                If blnAutoClose Then
                    DAO.AutoCloseConn = True
                End If

            End If
        End Try

        Return aRiaresult
    End Function
    Public Function ReUseSave(ByVal dsData As DataSet) As RIAResult
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        'Dim CSLog As CableSoft.SO.BLL.DataLog.DataLog = Nothing
        Dim blnAutoClose As Boolean = False
        Dim aRiaresult As New RIAResult()
        Dim tbSO008Log As DataTable = Nothing
        Dim isBeginTrans As Boolean = False
        aRiaresult.ResultBoolean = False
        aRiaresult.ErrorCode = -99
        If DAO.Transaction IsNot Nothing Then
            trans = DAO.Transaction
        Else
            cn.ConnectionString = Me.LoginInfo.ConnectionString
            cn.Open()
            trans = cn.BeginTransaction
            DAO.Transaction = trans
            blnAutoClose = True
        End If
        DAO.AutoCloseConn = False
        If blnAutoClose Then
            CableSoft.BLL.Utility.Utility.SetClientInfo(DAO, LoginInfo.EntryId, Language.ReUseClientInfo)
        End If
        Dim updNow As Date = Date.Now
        aRiaresult = chkReUseOK(dsData)
        If Not aRiaresult.ResultBoolean Then
            Return aRiaresult
        End If
        isBeginTrans = True
        Try
            Dim Seqval As Object = Nothing
            With dsData.Tables(0).Rows(0)
                If (Integer.Parse(.Item("NewBeginNum")) - 1) > Integer.Parse(.Item("BeginNum")) Then
                    Dim oEndNum As String = (Integer.Parse(.Item("NewBeginNum")) - 1).ToString.PadLeft(.Item("OBeginNum").ToString.Length, "0")
                    Dim oTotalPaperCount As Integer = Integer.Parse(.Item("NewBeginNum")) - Integer.Parse(.Item("BeginNum"))
                    DAO.ExecNqry(_DAL.UpdReUseSO126(False), New Object() {oEndNum,
                                                                  CableSoft.BLL.Utility.DateTimeUtility.GetDTString(updNow),
                                                                  LoginInfo.EntryName, oTotalPaperCount, .Item("SEQ"), LoginInfo.CompCode})
                    Seqval = DAO.ExecSclr(_DAL.QuerySeqVal)
                    Dim newTotalPaperCount As Integer = Integer.Parse(.Item("EndNum")) - Integer.Parse(.Item("NewBeginNum")) + 1
                    DAO.ExecNqry(_DAL.InsSO126, New Object() {LoginInfo.CompCode, Seqval, .Item("NewEmpNO"), .Item("NewEmpName"),
                                                               .Item("NewGetPaperDate"), .Item("NewBeginNum"), .Item("EndNum"), newTotalPaperCount,
                                                               LoginInfo.EntryName, CableSoft.BLL.Utility.DateTimeUtility.GetDTString(updNow),
                                                               .Item("Prefix"), DBNull.Value, DBNull.Value, DBNull.Value})

                    DAO.ExecNqry(_DAL.UpdReUseSO127, New Object() {Seqval, .Item("NewEmpNO"), .Item("NewEmpName"),
                                                                   .Item("NewGetPaperDate"), LoginInfo.EntryName, CableSoft.BLL.Utility.DateTimeUtility.GetDTString(updNow),
                                                                   .Item("SEQ"), .Item("Prefix"), .Item("NewBeginNum"), .Item("Prefix"),
                                                                   .Item("EndNum")})

                Else

                    DAO.ExecNqry(_DAL.UpdReUseSO126(True), New Object() {CableSoft.BLL.Utility.DateTimeUtility.GetDTString(updNow),
                                                                  LoginInfo.EntryName, .Item("NewEmpNO"), .Item("NewEmpName"),
                                                                         .Item("NewGetPaperDate"), .Item("SEQ"), LoginInfo.CompCode})
                    DAO.ExecNqry(_DAL.UpdReUseSO127, New Object() {.Item("SEQ"), .Item("NewEmpNO"), .Item("NewEmpName"),
                                                                   .Item("NewGetPaperDate"), LoginInfo.EntryName, CableSoft.BLL.Utility.DateTimeUtility.GetDTString(updNow),
                                                                   .Item("SEQ"), .Item("Prefix"), .Item("BeginNum"), .Item("Prefix"),
                                                                   .Item("EndNum")})
                End If

            End With
            If blnAutoClose Then
                trans.Commit()
            End If
        Catch ex As Exception
            aRiaresult.ErrorCode = -99
            aRiaresult.ErrorMessage = ex.ToString
            aRiaresult.ResultBoolean = False
            If blnAutoClose AndAlso isBeginTrans Then
                trans.Rollback()
            End If
        Finally
            If blnAutoClose Then
                CableSoft.BLL.Utility.Utility.ClearClientInfo(DAO)
                If trans IsNot Nothing Then
                    trans.Dispose()
                End If
                If cn IsNot Nothing Then
                    cn.Close()
                    cn.Dispose()
                End If
                If blnAutoClose Then
                    DAO.AutoCloseConn = True
                End If
                'If CSLog IsNot Nothing Then
                '    CSLog.Dispose()
                '    CSLog = Nothing
                'End If
            End If
        End Try
        
        Return aRiaresult
    End Function
    Public Function SaveData(ByVal editMode As EditMode, ByVal dsSave As DataSet) As RIAResult
        Dim cn As DbConnection = DAO.GetConn()
        Dim trans As DbTransaction = Nothing
        Dim CSLog As CableSoft.SO.BLL.DataLog.DataLog = Nothing
        Dim blnAutoClose As Boolean = False
        Dim aRiaresult As New RIAResult()
        Dim tbSO008Log As DataTable = Nothing
        Dim isBeginTrans As Boolean = False
        aRiaresult.ResultBoolean = False
        aRiaresult.ErrorCode = -99
        If DAO.Transaction IsNot Nothing Then
            trans = DAO.Transaction
        Else
            cn.ConnectionString = Me.LoginInfo.ConnectionString
            cn.Open()
            trans = cn.BeginTransaction
            DAO.Transaction = trans
            blnAutoClose = True
        End If
        DAO.AutoCloseConn = False
        Dim aAction As String = Nothing

        Select Case editMode
            Case CableSoft.BLL.Utility.EditMode.Append
                aAction = Language.AddClientInfo
            Case CableSoft.BLL.Utility.EditMode.Edit
                aAction = Language.EditClientInfo
            Case Else
                aAction = Language.EditClientInfo
        End Select
        If blnAutoClose Then
            CableSoft.BLL.Utility.Utility.SetClientInfo(Me.DAO, LoginInfo.EntryId, aAction)
        End If

        Dim updNow As Date = Date.Now
        Try
            Dim aRETURNDATE As Object = DBNull.Value
            Dim aCLEARDATE As Object = DBNull.Value
            Dim aNOTE As Object = DBNull.Value
            If Not DBNull.Value.Equals(dsSave.Tables(0).Rows(0).Item("RETURNDATE")) AndAlso Not String.IsNullOrEmpty(dsSave.Tables(0).Rows(0).Item("RETURNDATE")) Then
                aRETURNDATE = dsSave.Tables(0).Rows(0).Item("RETURNDATE")
            End If
            If Not DBNull.Value.Equals(dsSave.Tables(0).Rows(0).Item("CLEARDATE")) AndAlso Not String.IsNullOrEmpty(dsSave.Tables(0).Rows(0).Item("CLEARDATE")) Then
                aCLEARDATE = dsSave.Tables(0).Rows(0).Item("CLEARDATE")
            End If
            If Not DBNull.Value.Equals(dsSave.Tables(0).Rows(0).Item("NOTE")) AndAlso Not String.IsNullOrEmpty(dsSave.Tables(0).Rows(0).Item("NOTE")) Then
                aNOTE = dsSave.Tables(0).Rows(0).Item("NOTE")
            End If
            If editMode = CableSoft.BLL.Utility.EditMode.Edit Then
                isBeginTrans = True
                DAO.ExecNqry(_DAL.UpdSO126, New Object() {LoginInfo.EntryName,
                                                          CableSoft.BLL.Utility.DateTimeUtility.GetDTString(updNow),
                                                          aRETURNDATE, aCLEARDATE, aNOTE, dsSave.Tables(0).Rows(0).Item("SEQ"),
                                                          LoginInfo.CompCode})
                DAO.ExecNqry(_DAL.UpdSO127, New Object() {LoginInfo.EntryName,
                                                          CableSoft.BLL.Utility.DateTimeUtility.GetDTString(updNow),
                                                          dsSave.Tables(0).Rows(0).Item("SEQ")})
            Else
                If Not IsDataOK(dsSave) Then
                    aRiaresult.ResultBoolean = False
                    aRiaresult.ErrorCode = -1
                    aRiaresult.ErrorMessage = Language.DualError
                    Return aRiaresult
                End If
                isBeginTrans = True
                Dim Seqval As Object = DAO.ExecSclr(_DAL.QuerySeqVal)
                With dsSave.Tables(0).Rows(0)
                    DAO.ExecNqry(_DAL.InsSO126, New Object() {LoginInfo.CompCode, Seqval, .Item("EmpNO"), .Item("EmpName"),
                                                              .Item("GetPaperDate"), .Item("BeginNum"), .Item("EndNum"), .Item("TotalPaperCount"),
                                                              LoginInfo.EntryName, CableSoft.BLL.Utility.DateTimeUtility.GetDTString(updNow),
                                                              .Item("Prefix"), .Item("RETURNDATE"), .Item("CLEARDATE"), .Item("NOTE")})
                    For i As Integer = 0 To Integer.Parse(.Item("TotalPaperCount")) - 1
                        Dim aPaperNum As String = .Item("Prefix").ToString & _
                                        (Integer.Parse(.Item("BeginNum")) + i).ToString.PadLeft(.Item("BeginNum").ToString.Length, "0")
                        DAO.ExecNqry(_DAL.InsSO127, New Object() {LoginInfo.CompCode, Seqval, aPaperNum,
                                                                  .Item("EmpNO"), .Item("EmpName"), .Item("GetPaperDate"), 1, LoginInfo.EntryName,
                                                                  CableSoft.BLL.Utility.DateTimeUtility.GetDTString(updNow)})
                    Next
                End With


            End If
            If blnAutoClose Then
                trans.Commit()
            End If
            aRiaresult.ResultBoolean = True
            aRiaresult.ErrorCode = 0
            aRiaresult.ErrorMessage = Nothing
        Catch ex As Exception
            
            aRiaresult.ErrorCode = -99
            aRiaresult.ErrorMessage = ex.ToString
            aRiaresult.ResultBoolean = False
            If blnAutoClose AndAlso isBeginTrans Then
                trans.Rollback()
            End If
        Finally
            If blnAutoClose Then
                CableSoft.BLL.Utility.Utility.ClearClientInfo(Me.DAO)
                If trans IsNot Nothing Then
                    trans.Dispose()
                End If
                If cn IsNot Nothing Then
                    cn.Close()
                    cn.Dispose()
                End If
                If blnAutoClose Then
                    DAO.AutoCloseConn = True
                End If
                If CSLog IsNot Nothing Then
                    CSLog.Dispose()
                    CSLog = Nothing
                End If
            End If
        End Try
        Return aRiaresult
    End Function
    Public Function QueryEmployee() As DataTable
        Return DAO.ExecQry(_DAL.QueryEmployee)
    End Function
    
    Public Function QueryAllData() As DataSet
        Dim dsReturn As New DataSet
        Dim tbCompCode As DataTable = Nothing
        Dim tbEmployee As DataTable = Nothing
        Try

            tbCompCode = QueryCompCode.Copy
            tbCompCode.TableName = "CompCode"
            tbEmployee = QueryEmployee.Copy
            tbEmployee.TableName = "Employee"
            With dsReturn.Tables
                .Add(tbCompCode)
                .Add(tbEmployee)
            End With
        Catch ex As Exception
            Throw
        Finally
            If tbCompCode IsNot Nothing Then
                tbCompCode.Dispose()
                tbCompCode = Nothing
            End If
            If tbEmployee IsNot Nothing Then
                tbEmployee.Dispose()
                tbEmployee = Nothing
            End If
        End Try

        Return dsReturn
    End Function
    Public Function QueryBillData(ByVal dsData As DataSet) As DataSet
        Dim ds As New DataSet
        Try

            Using tbQuery As DataTable = DAO.ExecQry(_DAL.QueryBillData, New Object() {
                                                        dsData.Tables(0).Rows(0).Item("BillNo"), LoginInfo.CompCode})
                tbQuery.TableName = "Result"
                ds.Tables.Add(tbQuery.Copy)
                tbQuery.Dispose()
            End Using
            Return ds.Copy
        Catch ex As Exception
            Throw ex
        Finally
            If ds IsNot Nothing Then
                ds.Dispose()
                ds = Nothing
            End If

        End Try
    End Function
    Public Function QueryData(ByVal dsWhere As DataSet) As DataSet
        Dim ds As New DataSet
        Try

            Using tbQuery As DataTable = DAO.ExecQry(_DAL.QueryData(dsWhere))
                ds.Tables.Add(tbQuery.Copy)
                tbQuery.Dispose()
            End Using
            Return ds.Copy
        Catch ex As Exception
            Throw ex
        Finally
            If ds IsNot Nothing Then
                ds.Dispose()
                ds = Nothing
            End If
            
        End Try
    End Function
    Public Function QueryPaperNum(ByVal dsData As DataSet) As DataSet
        Dim dsResult As New DataSet
        Try
            With dsData.Tables(0).Rows(0)
                Dim tbResult As DataTable = DAO.ExecQry(_DAL.QueryPaperNum, New Object() {
                                                   .Item("PaperNum1"), .Item("PaperNum2")})
                tbResult.TableName = "Result"
                dsResult.Tables.Add(tbResult.Copy)

            End With
            Return dsResult
        Catch ex As Exception
            Throw ex
        Finally
            If dsResult IsNot Nothing Then
                dsResult.Dispose()
                dsResult = Nothing
            End If
        End Try

    End Function
    
    Public Function ChkAuthority(ByVal Mid As String) As RIAResult
        Dim result As New RIAResult() With {.ErrorCode = 0, .ErrorMessage = Nothing, .ResultBoolean = True}
        Try
            Using obj As New CableSoft.SO.BLL.Utility.Utility(Me.LoginInfo, DAO)
                result = obj.ChkPriv(LoginInfo.EntryId, Mid)
                obj.Dispose()
            End Using
            If Not result.ResultBoolean AndAlso String.IsNullOrEmpty(result.ErrorMessage) Then
                result.ErrorCode = -1
                result.ErrorMessage = Language.NoPermission
            End If


        Catch ex As Exception
            result.ErrorMessage = ex.ToString
            result.ResultBoolean = False
            result.ErrorCode = -2
        Finally

        End Try
        Return result

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
