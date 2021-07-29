Public Class ACHAuthInDALMultiDB
    Inherits ACHAuthInDAL
    Implements IDisposable
    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Overrides Function GetMasterIdSeq() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return "Select sf_getsequenceno('S_SO106_MasterId')  "
            Case Else
                Return MyBase.GetMasterIdSeq
        End Select

    End Function
    Friend Overrides Function QueryOracleDate() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return "Select now()"
            Case Else
                Return MyBase.QueryOracleDate
        End Select


    End Function
    Friend Overrides Function UpdateOldAuth(ByVal RowIds As String) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim result As String = "Update SO106  Set StopFlag=1," &
                     "Stopdate=SendDate,OldACH =1 Where CTID::text In (" & RowIds & ")"
                Return result
            Case Else
                Return MyBase.UpdateOldAuth(RowIds)
        End Select

    End Function
    Friend Overrides Function InsertSO106() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim result As String = String.Format("Insert Into SO106  " &
                                           "(AuthorizeStatus,SnactionDate,AcceptTime," &
                                            "PropDate,OldACH,ReAuthorizeStatus," &
                                            "UpdEn,UpdTime,MasterId,AcceptName,Proposer," &
                                            "ID,BankCode,BankName,CardCode,CardName," &
                                            "StopYM,AccountID,AccountName,AccountNameID," &
                                            "MediaCode,MediaName,IntroID,IntroName,Note," &
                                            "UpdateFlag,CompCode,CustId,CMCode,CMName," &
                                            "Alien,AccountAlien,AcceptEn,CVC2,CitemStr,CitemStr2," &
                                            "AddCitemAccount,PTCode,PTName,ACHCustId,ACHSN," &
                                            "ACHTNo,ACHTDESC,NewUpdTime) Values (" &
                                            "1,To_Date({0}0,'yyyymmdd'),{0}1," &
                                            "{0}2,1,{0}3," &
                                            "{0}4,{0}5,{0}6,{0}7,{0}8,{0}9,{0}10,{0}11,{0}12,{0}13,{0}14," &
                                            "{0}15,{0}16,{0}17,{0}18,{0}19,{0}20,{0}21,{0}22,{0}23,{0}24," &
                                            "{0}25,{0}26,{0}27,{0}28,{0}29,{0}30,{0}31,{0}32,{0}33,{0}34," &
                                            "{0}35,{0}36,{0}37,{0}38,{0}39,{0}40,now())", Sign)
                Return result
            Case Else
                Return MyBase.InsertSO106
        End Select

    End Function
    Friend Overrides Function InsertSO106A() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim result As String = String.Format("Insert Into SO106A" &
                                "(MasterRowID,ACHTNO,Notes,CitemCodeStr,CitemNameStr," &
                                " StopFlag,StopDate,UpdEn,UpdTime,CreateTime,CreateEn," &
                                " RecordType,AuthorizeStatus,AchDesc,MasterId) Values" &
                                "( {0}0," &
                                " {0}1," &
                                " {0}2," &
                                 "{0}3," &
                                 " {0}4," &
                                 " {0}5," &
                                 " {0}6," &
                                 " {0}7," &
                                 " {0}8," &
                                " now()," &
                                 " {0}9," &
                                " {0}10," &
                               " {0}11," &
                                " {0}12, " &
                                " {0}13 )", Sign)
                Return result
            Case Else
                Return MyBase.InsertSO106A
        End Select

    End Function
    Friend Overrides Function QuerySO003() As String

        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim result As String = Nothing
                result = String.Format("select * from  SO003 " &
            " Where CompCode = {0}0  And CTID::text In ( " &
                "Select so003.CTID::text  from so003,( " &
                   " Select distinct custid,faciseqno,codeno from so003c,cd019 where SO003c.masterid = {0}1 " &
                  " and so003c.productcode = cd019.productcode ) b " &
                    " where so003.citemcode = b.codeno and so003.custid=b.custid and so003.faciseqno = b.faciseqno ) ", Sign)
                Return result
            Case Else
                Return MyBase.QuerySO003
        End Select
    End Function
    Friend Overrides Function QuerySO106A(ByVal ReplyType As CableSoft.SO.BLL.Billing.ACHAuthIn.ACHAuthIn.AuthType,
                                ByVal InputAchNo As String, ByVal StopDate1 As String, ByVal StopDate2 As String) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Select Case ReplyType
                    Case ACHAuthIn.AuthType.Auth
                        Return String.Format("Select CTID::text,SO106A.* From SO106A  " &
                        " Where MasterId={0}0 " &
                        " And ACHTNO={0}1 " &
                        " And ACHTNO In(" & InputAchNo & ") " &
                        " And Nvl(RecordType,0)=0" &
                        " And AuthorizeStatus is null" &
                        " And StopFlag<>1", Sign)
                    Case ACHAuthIn.AuthType.CancelAuth
                        Return String.Format("Select CTID::text,SO106A.* From SO106A  " &
                        " Where MasterId={0}0 " &
                        " And ACHTNO={0}1 " &
                        " And ACHTNO In( " & InputAchNo & " )" &
                        " And RecordType=1" &
                        " And AuthorizeStatus is null" &
                        " And StopFlag=1" &
                        " And StopDate >= {0}2 " &
                        " And StopDate <= {0}3", Sign)
                    Case ACHAuthIn.AuthType.OldAuth
                        Return String.Format("Select CTID::text,SO106A.* From  SO106A  " &
                        " Where MasterId={0}0 " &
                        " And ACHTNO={0}1 " &
                        " And ACHTNO In(" & InputAchNo & ") " &
                        " And RecordType=0 " &
                        " And AuthorizeStatus=1" &
                        " And StopFlag<>1", Sign)
                    Case Else
                        Return "Select CTID:text,SO106A.* From  SO106A  Where 1 = 0 "
                End Select
            Case Else
                Return MyBase.QuerySO106A(ReplyType, InputAchNo, StopDate1, StopDate2)
        End Select

    End Function
    Friend Overrides Function QuerySO106AllData()
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return String.Format("Select CTID:text,SO106.* From SO106  Where MasterId = {0}0", Sign)
            Case Else
                Return MyBase.QuerySO106AllData
        End Select

    End Function
    Friend Overloads Overrides Function StopSO003() As String

        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim Result As String = Nothing
                Result = String.Format("Update SO003 Set " &
                 "BANKCODE=NULL" &
                 ",BANKNAME=NULL" &
                 ",ACCOUNTNO=NULL" &
                 ",InvSeqNO=NULL" &
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
                Return Result
                Return result
            Case Else
                Return MyBase.StopSO003
        End Select
    End Function
    Friend Overrides Function UpdateSO016A(ByVal ReplyType As CableSoft.SO.BLL.Billing.ACHAuthIn.ACHAuthIn.AuthType) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Select Case ReplyType
                    Case ACHAuthIn.AuthType.Auth
                        Return String.Format("Update SO106A Set AuthorizeStatus = 1,UpdTime = {0}0,UpdEn = {0}1 Where CTID = {0}2::tid", Sign)
                    Case ACHAuthIn.AuthType.CancelAuth
                        Return String.Format("Update SO106A Set AuthorizeStatus = 2,UpdTime = {0}0,UpdEn = {0}1 Where CTID = {0}2::tid", Sign)
                    Case ACHAuthIn.AuthType.ErrorType
                        Return String.Format("Update SO106A Set AuthorizeStatus = 3, " &
                                                "UpdTime = {0}0," &
                                                "UpdEn = {0}1, " &
                                                "Notes = {0}2 " &
                                  " Where CTID  = {0}3::tid", Sign)
                End Select
            Case Else
                Return MyBase.UpdateSO016A(ReplyType)
        End Select


    End Function
    Friend Overrides Function GetEmptySO106A() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return "Select CTID::text,SO106A.* from SO106A  Where 1=0"
            Case Else
                Return MyBase.GetEmptySO106A
        End Select

    End Function
    Friend Overrides Function UpdateSO003() As String

        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim result As String = Nothing
                result = String.Format("Update SO003 Set BankCode={0}0," &
                " BANKNAME={0}1," &
                " ACCOUNTNO={0}2," &
                " PTCode={0}3," &
                " PTName={0}4," &
                " CMCode={0}5," &
                " CMName={0}6" &
                ",UpdEn = {0}7,UpdTime = {0}8,NewUpdTime = To_Date({0}9,'yyyymmddhh24miss')" &
                " Where Ctid::text In ( " &
                    "Select so003.Ctid::text  from so003,( " &
                       " Select distinct custid,faciseqno,codeno from so003c,cd019 where SO003c.masterid = {0}10 " &
                      " and so003c.productcode = cd019.productcode ) b " &
                        " where so003.citemcode = b.codeno and so003.custid=b.custid and so003.faciseqno = b.faciseqno ) ", Sign)
                Return result
            Case Else
                Return MyBase.UpdateSO003
        End Select
    End Function
End Class
