Imports System.Data.OleDb
Module dbcon
    'Public connString As String = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Users\User\Documents\hris.accdb"
    Public conmed As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\pmis_server\IDPGLUdb2013\id_pglu.mdb")
    Public Myconnection As OleDbConnection
    Public idadapter As OleDbDataAdapter
    Public idcommand As OleDbCommand
    Public idreader As OleDbDataReader


    Public iddataset As DataSet
    Public idtables As DataTableCollection
    Public source As New BindingSource
    Public idresult As Integer
    Public idquery As String




End Module
