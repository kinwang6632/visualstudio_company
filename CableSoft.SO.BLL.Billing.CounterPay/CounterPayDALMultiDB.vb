Public Class CounterPayDALMultiDB
    Inherits CounterPayDAL
    Implements IDisposable

    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)

    End Sub
    Friend Overrides Function GetSimple(ByVal CompCode As Integer, ByVal BillNo As String, ByVal Item As Integer) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim strSQL As String = String.Format("Select A.CTID::text,0 as Type,A.*  From SO033 A Where A.CompCode={0} And A.BillNo='{1}' And A.Item={2}", CompCode, BillNo, Item)
                Return strSQL
            Case Else
                Return MyBase.GetSimple(CompCode, BillNo, Item)
        End Select

    End Function
    Friend Overrides Function GetPeriodCycle(ByVal Custid As Integer, ByVal CitemCode As Integer, ByVal RcdRowId As String) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim strSQL As String = String.Format("Select CMCode,CMName,PTCode,PTName FROM SO003 " &
                                                               "WHERE CustId={0} AND CitemCode={1} AND SeqNo in (Select SeqNo FROM SO033 WHERE CTID='{2}')", Custid, CitemCode, RcdRowId)
                Return strSQL
            Case Else
                Return MyBase.GetPeriodCycle(Custid, CitemCode, RcdRowId)
        End Select

    End Function
    Friend Overrides Function UpdRealCharge(ByVal SUCCode As String, ByVal SUCName As String, ByVal CMCode As String, ByVal CMName As String,
                                                                ByVal PTCode As String, ByVal PTName As String, ByVal RcdRowid As String) As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim strSQL As String = String.Format("Update SO033 Set RealDate= Null,RealAmt=0 " &
                                                                                        ",UCCode={0},UCName='{1}' " &
                                                                                        ",CMCode={2},CMName='{3}' " &
                                                                                        ",PTCode={4},PTName='{5}' " &
                                                                                        ",ClctEn=OldClctEn,ClctName=OldClctName" &
                                                                                        " Where CTID='{6}' And CancelFlag=0", Integer.Parse(SUCCode), SUCName, Integer.Parse(CMCode), CMName, Integer.Parse(PTCode), PTName, RcdRowid)
                Return strSQL
            Case Else
                Return MyBase.UpdRealCharge(SUCCode, SUCName, CMCode, CMName, PTCode, PTName, RcdRowid)
        End Select

    End Function
    Friend Overrides Function GetChargeTmp(ByVal strWhere As String, intTotalAmt As Integer, strMaxEntryNo As String, intCountNo As Integer) As String

        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim strField As String = ",C.SeqNo,C.GetDate,C.PrDate,C.InstDate,C.StopTime,C.ReInstTime,D.GuiNo,D.PreInvoice"
                strField = String.Format("{0},{1} as TotalAmt,'{2}' as MaxEntryNo,{3} as CountNo", strField, intTotalAmt, strMaxEntryNo, intCountNo)

                Dim strSQL As String = String.Format("Select A.CTID::text,A.*,B.CustStatusName {0} " &
                                            "From SO074A A left join so004 C on A.FaciSeqNo=C.SeqNo AND A.CustId=C.CustId ,SO002 B,SO033 D " &
                                            "Where A.CustId=B.CustId " &
                                            "And A.ServiceType=B.ServiceType And A.Custid=D.Custid And A.BillNo=D.BillNo And A.Item=D.Item And {1} " &
                                            "And D.CancelFlag=0 Order by A.EntryNo Desc,A.BillNo,A.Item", strField, strWhere)
                Return strSQL
            Case Else
                Return MyBase.GetChargeTmp(strWhere, intTotalAmt, strMaxEntryNo, intCountNo)
        End Select
    End Function
    Friend Overrides Function GetRealCharge(ByVal CompCode As Integer, ByVal strWhere As String) As String

        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim strSQL As String = String.Format("Select  A.ctid::text, A.*  From SO033 A Where A.CompCode={0} And {1} And A.UCCode is not null And A.CancelFlag=0", CompCode, strWhere)
                Return strSQL
            Case Else
                Return MyBase.GetRealCharge(CompCode, strWhere)
        End Select
    End Function
    Friend Overrides Function GetSimple2() As String

        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Dim strSQL As String = String.Format("Select A.CTID::text,A.* From SO033 A Where A.CompCode={0}0 And A.CTID={0}1", Sign)
                'Return strSQL
                Return MyBase.GetSimple2
            Case Else
                Return MyBase.GetSimple2
        End Select
    End Function
End Class
