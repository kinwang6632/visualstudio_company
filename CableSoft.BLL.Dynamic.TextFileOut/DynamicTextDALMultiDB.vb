Public Class DynamicTextDALMultiDB
    Inherits DynamicTextDAL
    Implements IDisposable

    Public Sub New()

    End Sub
    Public Sub New(ByVal Provider As String)
        MyBase.New(Provider)
    End Sub
    Friend Overrides Function InsertSO1119A() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return String.Format("Insert into SO1119A (SeqNo,ProgramId,EntryId,EntryName,Parameters,IsExec, " &
                        "ExecType,ParentSeqNo,AcceptTime,ResvTime,ExecProgramId,Caption,ProgItem,FilePath,FileName ) " &
                        " values (sf_getsequenceno('S_SO1108A'),{0}0,{0}1,{0}2", Sign)
            Case Else
                Return MyBase.InsertSO1119A
        End Select

    End Function
    Friend Overrides Function UpdLogData() As String
        Select Case CableSoft.BLL.Utility.Utility.GetDataBaseName(MyBase.Provider)
            Case CableSoft.BLL.Utility.Utility.DataBaseName.PostgreSql
                Return String.Format("UPDATE SO1108A SET EXECSTATUS={0}0 " &
                             ",EXECMESSAGE={0}1,FINISHTIME = now(),DOWNLOADFILENAME = {0}2,SQLQUERY = {0}3 " &
                        " WHERE SEQNO = {0}4", Sign)
            Case Else
                Return MyBase.UpdLogData
        End Select

    End Function
End Class
