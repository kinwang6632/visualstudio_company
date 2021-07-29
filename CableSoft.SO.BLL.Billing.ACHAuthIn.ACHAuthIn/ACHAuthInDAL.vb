Imports CableSoft.BLL.Utility
Public Class ACHAuthInDAL
    Inherits DALBasic
    Implements IDisposable
    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Function QueryCompCode(ByVal GroupId As String) As String
        If GroupId = "0" AndAlso 1 = 0 Then
            Return "Select A.CodeNo ,A.Description From CD039 A Order By CodeNo"
        Else
            Return String.Format("Select A.CodeNo,A.Description  " & _
                             " From CD039 A,SO026 B  " & _
                             " Where Instr(','||B.CompStr||',',','||A.CodeNo||',')>0 " & _
                            " And UserId = {0}0 Order By CodeNO", Sign)
        End If
    End Function
    Friend Function QueryFormatType() As String
        Return String.Format("Select Nvl(ACHCustID,0) ACHCustId From SO041 Where SysID={0}0", Sign)
    End Function
    Friend Function QueryBankId() As String
        Return String.Format("Select CodeNo,Description From CD018 " & _
                             " Where UPPER(PrgName) = 'ACHTRANREFER'  AND  COMPCODE ={0}0 " & _
                             " AND STOPFLAG <> 1", Sign)
    End Function
    Friend Function QueryBillHeadFmt() As String
        Return "Select ACHTDESC CodeNo,ACHTNO Description,CitemCodeStr From CD068 " & _
                    " Where ACHTNO Is Not Null And ACHTDesc Is not Null And ACHType=1 "
    End Function
    Friend Function QueryCitemCode() As String
        Return "Select CodeNo,Description From CD019 Where StopFlag <> 1"
    End Function
    Friend Function UpdateSO106(ByVal ReplyType As CableSoft.SO.BLL.Billing.ACHAuthIn.ACHAuthIn.AuthType) As String
        Dim Result As String = Nothing
        Select Case ReplyType
            Case ACHAuthIn.AuthType.Auth
                Result = String.Format("Update SO106  Set AuthorizeStatus=1, " &
                    " SnactionDate=To_Date({0}0,'YYYYMMDD'),ReAuthorizeStatus = {0}1," &
                    " UpdEn = {0}2,UpdTime = {0}3, " &
                    " NewUpdTime = To_Date({0}4,'yyyymmddhh24miss') " &
                    " Where ACHCustId = {0}5 And SendDate=To_Date({0}6,'YYYYMMDD') " &
                    " And LPAD(AccountID,14,'0')  = {0}7 " &
                    " And nvl(StopFlag,0) = 0 And SnactionDate is Null " &
                    " And InStr(',' || ACHTNO || ',' ,chr(39)|| {0}8 ||chr(39))>0 ", Sign)
            Case ACHAuthIn.AuthType.CancelAuth
                Result = String.Format("Update SO106  Set AuthorizeStatus=2, " &
                                     " ReAuthorizeStatus={0}0,UpdEn={0}1,UpdTime={0}2, " &
                                    " NewUpdTime = To_Date({0}3,'yyyymmddhh24miss') " &
                                     " Where  ACHCustId={0}4 And LPAD(AccountId,14,'0')= {0}5 ", Sign)


        End Select
        Return Result
    End Function
    Friend Function UpdateSO002A() As String
        Dim result As String = String.Format("UPDATE SO002A SET " & _
                 " BANKCODE={0}0" & _
                  ",BANKNAME={0}1" & _
                  ",ID=0,ACCOUNTNO={0}2 " & _
                  ",CHARGEADDRNO={0}3 " & _
                  ",CHARGEADDRESS={0}4 " & _
                  ",MAILADDRNO={0}5 " & _
                  ",MAILADDRESS={0}6 " & _
                  ",STOPFLAG=0,STOPDATE=NULL " & _
                  " WHERE CUSTID={0}7 " & _
                  " AND LPAD(AccountNo,14,'0')={0}8 " & _
                  " AND COMPCODE= {0}9", Sign)
        Return result
    End Function
    Friend Overridable Function GetMasterIdSeq() As String
        Return "Select S_SO106_MasterId.NEXTVAL FROM DUAL  "
    End Function
    Friend Overridable Function InsertSO106() As String
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
                                            "{0}35,{0}36,{0}37,{0}38,{0}39,{0}40,sysdate)", Sign)
        Return result
    End Function
    Friend Overridable Function InsertSO106A() As String
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
                                " sysdate," &
                                 " {0}9," &
                                " {0}10," &
                               " {0}11," &
                                " {0}12, " &
                                " {0}13 )", Sign)
        Return result
    End Function
    Friend Overridable Function UpdateOldAuth(ByVal RowIds As String) As String
        Dim result As String = "Update SO106  Set StopFlag=1," &
                     "Stopdate=SendDate,OldACH =1 Where RowId In(" & RowIds & ")"
        Return result
    End Function
    Friend Function InserSO002A() As String
        Dim result As String = String.Format("INSERT INTO SO002A " & _
                 "(CUSTID,COMPCODE,BANKCODE,BANKNAME,ID,ACCOUNTNO," & _
                 "CHARGEADDRNO,CHARGEADDRESS," & _
                 "MAILADDRNO,MAILADDRESS,CHARGETITLE," & _
                 "INVNO,INVTITLE,INVADDRESS,INVOICETYPE)" & _
                 " VALUES (" & _
                 "{0}0 ," & _
                " {0}1, " & _
                " {0}2, " & _
                " {0}3, " & _
                 0 & "," & _
                 "{0}4, " & _
                 "{0}5, " & _
                 "{0}6, " & _
                 "{0}7, " & _
                 "{0}8, " & _
                 "{0}9, " & _
                 "{0}10, " & _
                 "{0}11, " & _
                 "{0}12, " & _
                 "{0}13  )", Sign)
        Return result
    End Function
    Friend Function IsExistsSO002AD() As String
        Dim result As String = String.Format("Select Count(*) From  SO002AD " & _
                                " Where CustId={0}0 " & _
                                " And AccountNo={0}1 " & _
                                " And COMPCODE= {0}2 " & _
                                " AND INVSEQNO= {0}3 ", Sign)
        Return result
    End Function
    Friend Function InsertSO002AD() As String
        Dim result As String = String.Format("Insert Into SO002AD " & _
                    "(AccountNo,CompCode,CustId,InvSeqNo)" & _
                    " Values(" & _
                    " {0}0,{0}1,{0}2,{0}3 )", Sign)
        Return result
    End Function
    Friend Function IsExistsSO002A() As String
        Dim result As String = String.Format("select count(*) from so002a where " & _
                     "LPAD(AccountNo,14,'0')={0}0 " & _
                     " And CustId={0}1 " & _
                     " And CompCode = {0}2", Sign)
        Return result
    End Function
    'Friend Function QuerySO003CAndSO003(ByVal Masterid As String, ByVal ACHTNo As String) As String
    '    Dim lstACHTNo As List(Of String) = ACHTNo.Split(",").ToList
    '    Dim aWhere As String = Nothing
    '    For Each ACH As String In lstACHTNo
    '        If String.IsNullOrEmpty(aWhere) Then
    '            aWhere = String.Format("instr(ACHTNo,chr(39)||{0}||chr(39))>0", ACH)
    '        Else
    '            aWhere = String.Format("{0} Or instr(ACHTNo,chr(39)||{1}||chr(39))>0", aWhere, ACH)

    '        End If
    '    Next
    '    Dim result = "Select codeno from so003 where productcode in ( Select productcode from SO003C where masterid = " & Masterid & " )"

    'End Function
    Friend Function QueryUpdSO003Data(ByVal ACHTNo As String) As String
        Dim lstACHTNo As List(Of String) = ACHTNo.Split(",").ToList
        Dim aWhere As String = Nothing
        For Each ACH As String In lstACHTNo
            If String.IsNullOrEmpty(aWhere) Then
                aWhere = String.Format("instr(ACHTNo,chr(39)||{0}||chr(39))>0", ACH)
            Else
                aWhere = String.Format("{0} Or instr(ACHTNo,chr(39)||{1}||chr(39))>0", aWhere, ACH)

            End If
        Next
        'Dim result As String = String.Format("SELECT Citemstr,AccountID,BankCode,BankName," & _
        '                                     " PTCode,PTName,CMCode,CMName,InvSeqNo " & _
        '                                     " FROM  SO106 WHERE CUSTID={0}0 " & _
        '                                    " AND LPAD(AccountId,14,'0')={0}1 " & _
        '                                    " And ACHCustId={0}2" & _
        '                                    " AND NVL(StopFlag,0)=0 AND " & _
        '                                    " SendDate=To_Date({0}3,'YYYYMMDD') " & _
        '                                    " And instr(ACHTNo,chr(39)||{0}4||chr(39))>0", Sign)
        Dim result As String = String.Format("SELECT Citemstr,AccountID,BankCode,BankName," & _
                                             " PTCode,PTName,CMCode,CMName,InvSeqNo " & _
                                             " FROM  SO106 WHERE CUSTID={0}0 " & _
                                            " AND LPAD(AccountId,14,'0')={0}1 " & _
                                            " And ACHCustId={0}2" & _
                                            " AND NVL(StopFlag,0)=0 AND " & _
                                            " SendDate=To_Date({0}3,'YYYYMMDD') " & _
                                            " And (" & aWhere & ")", Sign)

        Return result
    End Function
    Friend Function GetCD008Where() As String
        Return "And Exists(Select CitemCode From SO003 B Where " &
                                                                                 "  1=1 " &
                                                                                    " And B.CompCode=SO106.CompCode " &
                                                                                     " And instr(','||SO106.Citemstr||',',','||Chr(39)||B.Seqno||Chr(39)||',')>0 " &
                                                                                     " And Exists(Select * From  CD068 C Where " &
                                                                                                         " instr(','||C.Citemcodestr||',',','||B.CitemCode||',')>0 " &
                                                                                                         " And C.BillHeadFmt In({0}) And C.ACHType=1 ))"


        'Return "And Exists(Select CitemCode From SO003 B Where " & _
        '                                                                         "B.Custid = A.Custid " & _
        '                                                                            " And B.CompCode=A.CompCode " & _
        '                                                                             " And instr(','||A.Citemstr||',',','||Chr(39)||B.Seqno||Chr(39)||',')>0 " & _
        '                                                                             " And Exists(Select * From  CD068 C Where " & _
        '                                                                                                 " instr(','||C.Citemcodestr||',',','||B.CitemCode||',')>0 " & _
        '                                                                                                 " And C.BillHeadFmt In({0}) And C.ACHType=1 ))"
    End Function
    Friend Overridable Function GetEmptySO106A() As String
        Return "Select rowid,SO106.* from SO106A  Where 1=0"
    End Function
    Friend Function UpdateACHSO003C(ByVal prdServiceId As String) As String
        Dim result As String = Nothing
        result = String.Format("Update SO003C Set CMCode={0}0," & _
           " CMName={0}1," & _
           " PTCode={0}2," & _
           " PTName={0}3," & _
           " UpdTime = {0}4, " & _
           " UpdEn = {0}5, " & _
           " NewUpdTime = To_Date({0}6,'yyyymmddhh24miss'), " & _
           " MasterId = {0}7 " & _
           " Where ServiceId In (" & prdServiceId & ")", Sign)
        Return result
    End Function

    Friend Function UpdateSO003C() As String
        Dim result As String = Nothing
        result = String.Format("Update SO003C Set CMCode={0}0," & _
            " CMName={0}1," & _
            " PTCode={0}2," & _
            " PTName={0}3," & _
            " UpdTime = {0}4, " & _
            " UpdEn = {0}5, " & _
            " NewUpdTime = To_Date({0}6,'yyyymmddhh24miss') " & _            
            " Where Masterid = {0}7 ", Sign)
        Return result
    End Function
    Friend Function UpdNonePeriod(ByVal CitemStrSeqNo As String) As String


        Dim aCustId = "Select distinct custid from so001  " & _
                                                   " where ID = {0}10 "
        Dim aSQL As String = String.Format("UPDATE SO003 SET BankCode={0}0, " &
                       "BANKNAME = {0}1,ACCOUNTNO = {0}2,PTCode = {0}3,PTName={0}4, " &
                       "CMCode = {0}5,CMName = {0}6,UpdEn = {0}7,UpdTime = {0}8,NewUpdTime = To_Date({0}9,'yyyymmddhh24miss')  " &
                       " WHERE 1=1 AND CUSTID In (" & aCustId & ") " & _
                       " And SEQNO IN (" & CitemStrSeqNo & ") " & _
                       " And CitemCode In (Select CodeNo From CD019 Where ProductCode is Null)", Sign)
        Return aSQL
    End Function
    Friend Overridable Overloads Function UpdateSO003() As String
        Dim result As String = Nothing
        result = String.Format("Update SO003 Set BankCode={0}0," &
            " BANKNAME={0}1," &
            " ACCOUNTNO={0}2," &
            " PTCode={0}3," &
            " PTName={0}4," &
            " CMCode={0}5," &
            " CMName={0}6" &
            ",UpdEn = {0}7,UpdTime = {0}8,NewUpdTime = To_Date({0}9,'yyyymmddhh24miss')" &
            " Where Rowid In ( " &
                "Select so003.rowid  from so003,( " &
                   " Select distinct custid,faciseqno,codeno from so003c,cd019 where SO003c.masterid = {0}10 " &
                  " and so003c.productcode = cd019.productcode ) b " &
                    " where so003.citemcode = b.codeno and so003.custid=b.custid and so003.faciseqno = b.faciseqno ) ", Sign)
        Return result
    End Function

    Friend Overloads Function UpdateSO003(ByVal CitemCodeStr As String) As String
        Dim result As String = String.Format("Update SO003 Set BankCode={0}0," & _
            " BANKNAME={0}1," & _
            " ACCOUNTNO={0}2," & _
            " PTCode={0}3," & _
            " PTName={0}4," & _
            " CMCode={0}5," & _
            " CMName={0}6," & _
            " InvSeqNo={0}7 " & _
            " WHERE CUSTID={0}8 " & _
            " And CitemCode In(" & CitemCodeStr & ")", Sign)
        Return result
    End Function
    Friend Function QueryUpdSO002AData() As String
        Dim result As String = String.Format(
                "Select DISTINCT CUSTID,BANKCODE,BANKNAME,AccountID," &
                            "AccountName,B.ChargeAddrNo,B.ChargeAddress,B.MailAddrNo," &
                        "B.MailAddress,C.InvNo,C.InvTitle,C.InvAddress,C.InvoiceType,InvSeqNo " &
                        " From SO106 ,So001 B,So002 C" &
             " Where CustId=B.Custid And Custid=C.Custid And MasterId = {0}0 " &
             " And LPAD(AccountID,14,'0') = {0}1" &
             " And Compcode={0}2 ", Sign)
        Return result
    End Function
    Friend Function QueryCD068() As String
        Return "Select BillHeadFmt || ACHTNO PKName, BillHeadFmt,ACHTNO,ACHTDesc,CitemCodeStr From CD068 " & _
            " Where ACHTNO Is Not Null And ACHTDesc Is not Null And ACHType=1"
    End Function
    Friend Function ClearSO106() As String
        Return String.Format("Update SO106 Set ACHTNO = {0}0,ACHTDesc = {0}1,CitemStr = {0}2 Where MasterId = {0}3 ", Sign)
    End Function
    Friend Function InsertSO004() As String
        Dim result As String = String.Format("UPDATE SO004 SET " & _
                                        " AccountNo = {0}0," & _
                                        " BankCode = {0}1," & _
                                        " BankName = {0}2," & _
                                        " ChPTCode = {0}3," & _
                                        " ChPTName = {0}4," & _
                                        " ChCMCode = {0}5," & _
                                        " ChCMName = {0}6 " & _
                                " Where MasterId = {0}7 ", Sign)
        Return result
    End Function
    Friend Overridable Function UpdateSO016A(ByVal ReplyType As CableSoft.SO.BLL.Billing.ACHAuthIn.ACHAuthIn.AuthType) As String
        Select Case ReplyType
            Case ACHAuthIn.AuthType.Auth
                Return String.Format("Update SO106A Set AuthorizeStatus = 1,UpdTime = {0}0,UpdEn = {0}1 Where RowId = {0}2", Sign)
            Case ACHAuthIn.AuthType.CancelAuth
                Return String.Format("Update SO106A Set AuthorizeStatus = 2,UpdTime = {0}0,UpdEn = {0}1 Where RowId = {0}2", Sign)
            Case ACHAuthIn.AuthType.ErrorType
                Return String.Format("Update SO106A Set AuthorizeStatus = 3, " &
                                                "UpdTime = {0}0," &
                                                "UpdEn = {0}1, " &
                                                "Notes = {0}2 " &
                                  " Where RowId  = {0}3", Sign)
        End Select

    End Function
    Friend Overridable Function QueryOracleDate() As String
        Return "Select SysDate From Dual"
    End Function
    Friend Function QuerySO106AErrNote() As String
        Return String.Format("Select * From SO106A " & _
                               " Where MasterId={0}0 " & _
                               " And AuthorizeStatus=3", Sign)
    End Function
    Friend Function UpdateSO106Note() As String
        Return String.Format("Update SO106 Set Note = Decode(Note,Null,Null,Note || chr(13) || chr(10)) || {0}0  " & _
                                             "  Where MasterId = {0}1", Sign)
    End Function
   
   
    Friend Function GetPTCode() As String
        Return "Select CodeNo,Description,RefNo From CD032 Where Nvl(StopFlag,0) = 0 ORDER BY CODENO"
    End Function
    Friend Function GetDefCMCode(ByVal LoginInfo As LoginInfo, ByVal ServiceType As String) As String
        Dim aOBJ As New CableSoft.SO.BLL.Utility.Charge(LoginInfo)
        Try
            Dim aCMCode As String = aOBJ.GetDefaultCMCode(ServiceType).ToString
            Return "SELECT " & aCMCode & " CODENO , Description FROM CD031 WHERE CODENO = " & aCMCode
        Finally
            aOBJ.Dispose()
            aOBJ = Nothing
        End Try

    End Function
    Friend Function StopSO003C() As String
        Dim result As String = Nothing
        result = String.Format("Update SO003C Set CMCode={0}0," & _
            " CMName={0}1," & _
            " PTCode={0}2," & _
            " PTName={0}3," & _
            " UpdTime = {0}4, " & _
            " UpdEn = {0}5, " & _
            " NewUpdTime = To_Date({0}6,'yyyymmddhh24miss') " & _
            " Where Masterid = {0}7 ", Sign)
        Return result
    End Function
    Friend Function StopNonePeriod(ByVal CitemStrSeqNo As String) As String
        Dim aCustId = "Select distinct custid from so001  " & _
                                                 " where ID = {0}8 "
        Dim aSQL As String = String.Format("UPDATE SO003 SET BankCode= Null, " &
                       "BANKNAME = Null ,ACCOUNTNO = Null ,CMCode = {0}0,CMName = {0}1 " & _
                       " ,PTCode = {0}2,PTName={0}3 " &
                       ",UpdEn = {0}4,UpdTime = {0}5,NewUpdTime = To_Date({0}6,'yyyymmddhh24miss')  " &
                       " WHERE 1=1  And ACCOUNTNO = {0}7 AND CUSTID In (" & aCustId & ") " & _
                       " And SEQNO IN (" & CitemStrSeqNo & ") " & _
                       " And CitemCode In (Select CodeNo From CD019 Where ProductCode is Null)", Sign)
        Return aSQL
    End Function
    Friend Overridable Overloads Function StopSO003() As String
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
                 " And RowId In ( " &
                "Select so003.rowid  from so003,( " &
                   " Select distinct custid,faciseqno,codeno from so003c,cd019 where SO003c.masterid = {0}8 " &
                  " and so003c.productcode = cd019.productcode ) b " &
                    " where so003.citemcode = b.codeno and so003.custid=b.custid and so003.faciseqno = b.faciseqno ) " &
                    " And ACCOUNTNO = {0}9", Sign)
        Return Result
    End Function
    Friend Overloads Function StopSO003(ByVal CitemCodeStr As String) As String
      

        Dim Result As String = String.Format("Update SO003 Set " & _
                 "BANKCODE=NULL" & _
                 ",BANKNAME=NULL" & _
                 ",ACCOUNTNO=NULL" & _
                 ",InvSeqNO=NULL" & _
                 " Where CompCode={0}0 " & _
                 " And CustId={0}1 " & _
                 " And CitemCode In(" & CitemCodeStr & ") " & _
                 " And ACCOUNTNO={0}2 ", Sign)
        Return Result
    End Function
    Friend Function StopSO106(ByVal StopAll As Boolean) As String
        Dim Result As String = Nothing
        If StopAll Then
            Result = String.Format("Update SO106  Set SendDate=Null,ACHCustId=ACHCustId,UpdEn={0}0 ,UpdTime={0}1, " &
                         " StopFlag=1,StopDate={0}2 ,ReAuthorizeStatus={0}3, " &
                         "Note= Note || decode(note,null,'',  chr(13)  || chr(10) ) || {0}4, " &
                         "NewUpdTime = To_Date({0}5,'yyyymmddhh24miss') " &
                         " Where LPAD(AccountId,14,'0')={0}6 " &
                         "  And ACHCustId={0}7 " &
                         " And SnactionDate is Null And nvl(StopFlag,0) = 0", Sign)
        Else
            Result = String.Format("Update SO106  Set SendDate=Null,ACHCustId=ACHCustId,UpdEn={0}0,UpdTime={0}1 ," &
                                 " ReAuthorizeStatus={0}2, " &
                                 "Note= Note || decode(note,null,'',  chr(13)  || chr(10) ) || {0}3, " &
                                 "NewUpdTime = To_Date({0}4,'yyyymmddhh24miss') " &
                                 " Where  LPAD(AccountId,14,'0')={0}5  " &
                                 "  And ACHCustId={0}6 " &
                                 " And SnactionDate is Null And nvl(StopFlag,0) = 0 ", Sign)

        End If
        Return Result
    End Function
    Friend Overridable Function QuerySO003() As String
        Dim result As String = Nothing
        result = String.Format("select * from  SO003 " &
            " Where CompCode = {0}0  And Rowid In ( " &
                "Select so003.rowid  from so003,( " &
                   " Select distinct custid,faciseqno,codeno from so003c,cd019 where SO003c.masterid = {0}1 " &
                  " and so003c.productcode = cd019.productcode ) b " &
                    " where so003.citemcode = b.codeno and so003.custid=b.custid and so003.faciseqno = b.faciseqno ) ", Sign)
        Return result
    End Function
    Friend Function QuerySO003(ByVal CitemCodeStr As String) As String        
        Return String.Format("Select * From  SO003 " & _
                    " Where CustId={0}0 " & _
                    " And CompCode={0}1 " & _
                    " And CitemCode IN(" & CitemCodeStr & ")", Sign)
    End Function
    Friend Function ChkSO106AAllFail() As String
        Return String.Format("Select Count(*) From SO106A" & _
                            " Where AchtNO<>{0}0 " & _
                            " And ACHDesc<>{0}1 " & _
                            " And MasterId={0}2 " & _
                            " And (AuthorizeStatus=1 or AuthorizeStatus is Null)", Sign)
    End Function
    Friend Function ChkCancelAuthStopAll() As String
        Dim result As String = String.Format("Select Count(Decode(AuthorizeStatus,Null,Null)) A," & _
                                    "Count(Decode(AuthorizeStatus,1,1)) B," & _
                                    "Count(Decode(AuthorizeStatus,2,2)) C," & _
                                    "Count(Decode(AuthorizeStatus,3,3)) D" & _
                            " From SO106A" & _
                            " Where MasterId={0}0", Sign)
        Return result
    End Function
    Friend Function UpdateCancelAuth() As String
        'Dim result As String = String.Format("Update SO106 A Set AuthorizeStatus=2,ReAuthorizeStatus='" & strReStatus & "',UpdEn='" & strUpdName & "',UpdTime='" & GetDTString(Now) & "' Where" & _
        '             " ACHCustId='" & GetACHCustID(strData) & "' And LPAD(AccountId,14,'0')='" & GetString(Mid(strData, 27, 14), 14, giRight, True) & "'")
    End Function
    Friend Overridable Function QuerySO106A(ByVal ReplyType As CableSoft.SO.BLL.Billing.ACHAuthIn.ACHAuthIn.AuthType,
                                ByVal InputAchNo As String, ByVal StopDate1 As String, ByVal StopDate2 As String) As String
        Select Case ReplyType
            Case ACHAuthIn.AuthType.Auth
                Return String.Format("Select RowId CTID,SO106A.* From SO106A  " &
                        " Where MasterId={0}0 " &
                        " And ACHTNO={0}1 " &
                        " And ACHTNO In(" & InputAchNo & ") " &
                        " And Nvl(RecordType,0)=0" &
                        " And AuthorizeStatus is null" &
                        " And StopFlag<>1", Sign)
            Case ACHAuthIn.AuthType.CancelAuth
                Return String.Format("Select RowId CTID,SO106A.* From SO106A  " &
                        " Where MasterId={0}0 " &
                        " And ACHTNO={0}1 " &
                        " And ACHTNO In( " & InputAchNo & " )" &
                        " And RecordType=1" &
                        " And AuthorizeStatus is null" &
                        " And StopFlag=1" &
                        " And StopDate >= {0}2 " &
                        " And StopDate <= {0}3", Sign)


                'IIf(Not String.IsNullOrEmpty(StopDate1), " And A.StopDate>=To_Date('" & StopDate1 & "','YYYYMMDDHH24MISS')", "") & _
                'IIf(Not String.IsNullOrEmpty(StopDate2), " And A.StopDate<=To_Date('" & StopDate2 & "','YYYYMMDDHH24MISS')", ""), Sign)
            Case ACHAuthIn.AuthType.OldAuth
                Return String.Format("Select RowId CTID ,SO106A.* From  SO106A  " &
                        " Where MasterId={0}0 " &
                        " And ACHTNO={0}1 " &
                        " And ACHTNO In(" & InputAchNo & ") " &
                        " And RecordType=0 " &
                        " And AuthorizeStatus=1" &
                        " And StopFlag<>1", Sign)
            Case Else
                Return "Select RowId CTID,SO106A.* From  SO106A  Where 1 = 0 "
        End Select
    End Function
    Friend Overridable Function QuerySO106AllData()
        Return String.Format("Select RowId CTID,SO106.* From SO106  Where MasterId = {0}0", Sign)
    End Function
    Friend Function QueryUpdOldAch() As String
        Dim result As String = String.Format("Select RowId CTID ,SO106.* From SO106  Where " &
                    " LPAD(AccountID,14,'0')={0}0" &
                    " And ACHCustId={0}1 " &
                    " And SnactionDate Is not null and Propdate is not null" &
                    " And OldACH=0 And nvl(StopFlag,0) = 0", Sign)
        Return result
    End Function
    Friend Function QueryUniqueSO106() As String
        Return String.Format("Select rowId CTID,SO106.* From SO106  Where Masterid = {0}0", Sign)
    End Function

    Friend Function QuerySO003C() As String
        Return String.Format("Select rowId CTID,SO003C.* From SO003C Where Masterid = {0}0", Sign)
    End Function
    Friend Function QueryCD019CodeNo() As String
        Dim result As String = Nothing
        result = String.Format("Select CodeNo From CD019 Where ProductCode In (" & _
            "Select ProductCode From SO003C Where MasterId= {0}0", Sign)

        Return result
    End Function
    Friend Function QuerySO106Data(ByVal ReplyType As CableSoft.SO.BLL.Billing.ACHAuthIn.ACHAuthIn.AuthType, ByVal StopDate1 As String, ByVal StopDate2 As String) As String
        Dim strSQL As String = Nothing
        strSQL = String.Format("Select RowId CTID,SO106.* From SO106  " &
                             " Where ACHCUSTID = {0}0 " &
                             " And LPAD(AccountID,14,'0') = {0}1  ", Sign)
        If ReplyType = ACHAuthIn.AuthType.CancelAuth Then            
            strSQL = strSQL & String.Format(" And {0}2 = {0}2 ", Sign)
           
          
            If Not String.IsNullOrEmpty(StopDate1) Then
                StopDate1 = StopDate1.Replace("/", "")
                If StopDate1.Length < "20141014235900".Length Then
                    StopDate1 = String.Format("{0}000000", StopDate1.Replace("/", ""))
                End If
                strSQL = strSQL & " And StopDate>=To_Date('" & StopDate1 & "','YYYYMMDDHH24MISS')"
            End If
            If Not String.IsNullOrEmpty(StopDate2) Then
                StopDate2 = StopDate2.Replace("/", "")
                If StopDate2.Length < "20141014235900".Length Then
                    StopDate2 = String.Format("{0}235959", StopDate2.Replace("/", ""))
                End If
                strSQL = strSQL & " And StopDate<=To_Date('" & StopDate2 & "','YYYYMMDDHH24MISS')"
            End If
        End If
        If (ReplyType = ACHAuthIn.AuthType.Auth) OrElse (ReplyType = ACHAuthIn.AuthType.OldAuth) Then
            strSQL = String.Format("{0} AND STOPFLAG <> 1  ", strSQL) & _
                String.Format(" And InStr(',' || ACHTNO || ',' ,chr(39)|| {0}2||chr(39))>0", Sign)
        End If
        If ReplyType = ACHAuthIn.AuthType.OldAuth Then
            strSQL = String.Format("{0}  And SnactionDate Is not null " & _
                                   " and Propdate is not null  " & _
                                   "And OldACH=0 And nvl(StopFlag,0) = 0", strSQL)
        End If
        Return strSQL
    End Function

    Friend Function QuerySO106Data(ByVal ReplyType As CableSoft.SO.BLL.Billing.ACHAuthIn.ACHAuthIn.AuthType) As String
        'Return QuerySO106Data(ReplyType, Nothing, Nothing)
        Dim strSQL As String = Nothing
        strSQL = String.Format("Select RowId CTID,SO106.* From SO106  " &
                             " Where ACHCUSTID = {0}0 " &
                             " And LPAD(AccountID,14,'0') = {0}1  ", Sign)
        If ReplyType = ACHAuthIn.AuthType.CancelAuth Then
            strSQL = strSQL & String.Format(" And {0}2 = {0}2 ", Sign)
        End If
        If (ReplyType = ACHAuthIn.AuthType.Auth) OrElse (ReplyType = ACHAuthIn.AuthType.OldAuth) Then
            strSQL = String.Format("{0} AND STOPFLAG <> 1  ", strSQL) & _
                String.Format(" And InStr(',' || ACHTNO || ',' ,chr(39)|| {0}2||chr(39))>0", Sign)
        End If
        If ReplyType = ACHAuthIn.AuthType.OldAuth Then
            strSQL = String.Format("{0}  And SnactionDate Is not null " & _
                                   " and Propdate is not null  " & _
                                   "And OldACH=0 And nvl(StopFlag,0) = 0", strSQL)
        End If
        Return strSQL
    End Function

    Friend Function chkAuthority(ByVal GroupField As String) As String
        Return String.Format("Select count(*) From SO029 Where Mid = {0}0 And  Group" & GroupField & "= 1", Sign)
    End Function
    Friend Function QueryActLength() As String
        Return String.Format("Select Nvl(ActLength,0) ActLength From CD018  Where BankId2 = {0}0", Sign)
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
