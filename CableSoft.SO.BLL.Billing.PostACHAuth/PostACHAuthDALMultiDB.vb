Public Class PostACHAuthDALMultiDB
    Inherits PostACHAuthDAL

    Implements IDisposable
    Private Const ReturnOK As String = "RETURNOK"
    Private Const ReturnFail As String = "RETURNFAIL"
    Private Const CancelAuth As String = "CANCELAUTH"
    Private Const PostTerminal As String = "POSTTERMINAL"
    Private Const ResumeData As String = "RESUMEDATA"
    Private Const FetchOK As String = "FETCHOK"
    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Overrides Function FetchData(ByVal strWhere As String) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return "SELECT CTID::text,Custid,BankCode,AccountID,AccountNameId,CitemStr,ACHTNo,ACHSN,ACHCUSTID,MasterId," &
                        " (Select BankId From CD018 Where CD018.CodeNo = BankCode  LIMIT  1 ) BankID " &
                        " From SO106 " & strWhere
            Case Else
                Return MyBase.FetchData(strWhere)
        End Select

    End Function
    Friend Overrides Function getApplySO106SQL(ByVal ACHCUSTID As String, ByVal AccountID As String, ByVal AccountNameID As String, ByVal strWhere As String) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return "Select  CTID::text,SO106.* From SO106  Where ACHCUSTID='" & ACHCUSTID &
                                 "' And LPAD(AccountID,14,'0')='" & AccountID & "'" &
                                 " And RPAD(nvl(AccountNameID,'0'),10,'0')='" & AccountNameID & "'" & strWhere
            Case Else
                Return MyBase.getApplySO106SQL(ACHCUSTID, AccountID, AccountNameID, strWhere)
        End Select

    End Function
    Friend Overrides Function StopSO003() As String
        Dim Result As String = Nothing
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Result = String.Format("Update SO003 Set " &
                "BANKCODE=NULL" &
                ",BANKNAME=NULL" &
                ",ACCOUNTNO=NULL" &
                ",CMCode = {0}0 " &
                ",CMName = {0}1 " &
                ",PTCode = {0}2 " &
                ",PTName = {0}3 " &
                ",UpdEn = {0}4,UpdTime = {0}5,NewUpdTime = To_Date({0}6,'yyyymmddhh24miss') " &
                " Where CompCode = {0}7 " &
                " And CTID::text In ( " &
               "Select so003.CTID::text  from so003,( " &
                  " Select distinct custid,faciseqno,codeno from so003c,cd019 where SO003c.masterid = {0}8 " &
                 " and so003c.productcode = cd019.productcode ) b " &
                   " where so003.citemcode = b.codeno and so003.custid=b.custid and so003.faciseqno = b.faciseqno ) " &
                   " And ACCOUNTNO = {0}9", Sign)
            Case Else
                Result = MyBase.StopSO003
        End Select

        Return Result
    End Function
    Friend Overrides Function UpdateSO106AStatus(ByVal ReturnString As String) As String

        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Select Case ReturnString
                    Case FetchOK
                        Return String.Format("Update  SO106A Set AuthorizeStatus = null,UpdTime = {0}0,UpdEn = {0}1 Where CTID = {0}2::tid", Sign)
                    Case ReturnOK
                        Return String.Format("Update  SO106A Set AuthorizeStatus = 1,UpdTime = {0}0,UpdEn = {0}1 Where CTID = {0}2::tid", Sign)
                    Case CancelAuth
                        Return String.Format("Update SO106A Set AuthorizeStatus = 2,UpdTime = {0}0,UpdEn = {0}1 Where CTID = {0}2::tid", Sign)
                    Case ReturnFail
                        Return String.Format("Update SO106A Set AuthorizeStatus = 3, " &
                                                "UpdTime = {0}0," &
                                                "UpdEn = {0}1, " &
                                                "Notes = {0}2 " &
                                  " Where CTID  = {0}3::tid", Sign)
                End Select
            Case Else
                Return MyBase.UpdateSO106AStatus(ReturnString)
        End Select


    End Function
    Friend Overrides Function QuerySO003() As String
        Dim result As String = Nothing
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                result = String.Format("select * from  SO003 " &
                      " Where CompCode = {0}0  And CTID::text In ( " &
                          "Select so003.CTID::text  from so003,( " &
                         " Select distinct custid,faciseqno,codeno from so003c,cd019 where SO003c.masterid = {0}1 " &
                        " and so003c.productcode = cd019.productcode ) b " &
                          " where so003.citemcode = b.codeno and so003.custid=b.custid and so003.faciseqno = b.faciseqno ) ", Sign)
            Case Else
                result = MyBase.QuerySO003
        End Select

        Return result
    End Function
    Friend Overrides Function QueryUniqueSO106() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return String.Format("Select ctid::text,SO106.* From SO106  Where Masterid = {0}0", Sign)
            Case Else
                Return MyBase.QueryUniqueSO106
        End Select

    End Function
    Friend Overrides Function QuerySO106A(ByVal strInACHTNO As String, ByVal strType As String, ByVal rwCondition As DataRow) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Select Case strType
                    Case "1"

                        Return String.Format("Select CTID::text,SO106A.* From SO106A  " &
                            "Where MasterId={0}0 " &
                            " And ACHTNO In(" & strInACHTNO & ")" &
                            " And RecordType=0" &
                            " And AuthorizeStatus is null" &
                            " And StopFlag<>1", Sign)
                    Case "2"
                        Dim aWhere As String = " And 1 =1 "
                        If Not DBNull.Value.Equals(rwCondition("STOPDATE1")) AndAlso Not String.IsNullOrEmpty(rwCondition("STOPDATE1")) Then
                            aWhere = aWhere & " And StopDate>=To_Date('" & rwCondition("STOPDATE1") & "','YYYYMMDD')"
                        End If
                        If Not DBNull.Value.Equals(rwCondition("STOPDATE2")) AndAlso Not String.IsNullOrEmpty(rwCondition("STOPDATE2")) Then
                            aWhere = aWhere & " And StopDate<To_Date(('" & rwCondition("STOPDATE2") & "','YYYYMMDD'))::date+1"
                        End If
                        Return String.Format("Select CTID::text,SO106A.* From SO106A  " &
                            "Where MasterId= {0}0  " &
                            " And ACHTNO In(" & strInACHTNO & ")" &
                            " And RecordType=1" &
                            " And AuthorizeStatus is null" &
                            " And StopFlag=1" &
                            aWhere, Sign)
                    Case "3"
                        Return String.Format("Select CTID::text,SO106A.* From SO106A  " &
                            "Where MasterId= {0}0" &
                            " And ACHTNO In(" & strInACHTNO & ")" &
                            " And AuthorizeStatus = 1 " &
                            " And StopFlag <> 1 ", Sign)

                    Case "4"
                        Return String.Format("Select CTID::text,SO106A.* From SO106A " &
                            " Where MasterId={0}0 ", Sign)

                End Select
            Case Else
                Return MyBase.QuerySO106A(strInACHTNO, strType, rwCondition)
        End Select


    End Function
    Friend Overrides Function FetchStopSO106A(ByVal strInACHTNO As String, ByVal strInACHDesc As String) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return String.Format("Select CTID::text,SO106A.* From SO106A  " &
                               "Where MasterId= {0}0 " &
                               " And ACHTNO In(" & strInACHTNO & ")" &
                               " And ACHDesc In(" & strInACHDesc & ")" &
                               " And RecordType=1 And AuthorizeStatus=4", Sign)
            Case Else
                Return MyBase.FetchStopSO106A(strInACHTNO, strInACHDesc)
        End Select

    End Function
    Friend Overrides Function FetchApplySO106A(ByVal strInACHTNO As String, ByVal strInACHDesc As String) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return String.Format("Select CTID::text,SO106A.* From SO106A  " &
                               "Where MasterId= {0}0 " &
                               " And ACHTNO In(" & strInACHTNO & ")" &
                               " And ACHDesc In(" & strInACHDesc & ")" &
                               " And RecordType=0 And AuthorizeStatus=4 " &
                               " And StopFlag<>1 And StopDate is Null Order By ACHTNO", Sign)
            Case Else
                Return MyBase.FetchApplySO106A(strInACHTNO, strInACHDesc)
        End Select


    End Function
    Friend Overrides Function UpdateSO003() As String
        Dim result As String = Nothing
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                result = String.Format("Update SO003 Set BankCode={0}0," &
          " BANKNAME={0}1," &
          " ACCOUNTNO={0}2," &
          " PTCode={0}3," &
          " PTName={0}4," &
          " CMCode={0}5," &
          " CMName={0}6" &
          ",UpdEn = {0}7,UpdTime = {0}8,NewUpdTime = To_Date({0}9,'yyyymmddhh24miss')" &
          " Where CTID::text In ( " &
              "Select so003.CTID::text  from so003,( " &
                 " Select distinct custid,faciseqno,codeno from so003c,cd019 where SO003c.masterid = {0}10 " &
                " and so003c.productcode = cd019.productcode ) b " &
                  " where so003.citemcode = b.codeno and so003.custid=b.custid and so003.faciseqno = b.faciseqno ) ", Sign)
            Case Else
                result = MyBase.UpdateSO003
        End Select

        Return result
    End Function
    Friend Overrides Function QuerySO003SEQNO(ByVal CitemStr As String) As String
        Dim result As String = Nothing

        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                result = String.Format("select * from  SO003 " &
                       " Where CompCode = {0}0  And CTID::text In ( " &
                           "Select so003.CTID::text  from so003,( " &
                          " Select distinct custid,faciseqno,codeno from so003c,cd019 where SO003c.masterid = {0}1 " &
                         " and so003c.productcode = cd019.productcode ) b " &
                           " where so003.citemcode = b.codeno and so003.custid=b.custid and so003.faciseqno = b.faciseqno )  " &
                           " And SO003.CitemCode In (" & CitemStr & ")", Sign)
            Case Else
                result = MyBase.QuerySO003SEQNO(CitemStr)
        End Select

        Return result
    End Function
End Class
