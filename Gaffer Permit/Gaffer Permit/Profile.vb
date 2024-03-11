Imports System.Data.OleDb
Public Class Profile
    Private Sub Profile_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Connect = New OleDbConnection(ConnectionString_MDB)
        DataGridView1.DataSource = GET_RECORDS("SELECT * FROM Main_Record", Connect)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim addsql As String
        addsql = "insert into Main_Record(Gaffer_ID,Full_Name,Permit_No,Amount_Paid,OR_No,Date_Issued)values('" & "ID1" & "','" & TextBox1.Text &
            "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & DateTimePicker1.Value & "')"
        execute(addsql)
        DataGridView1.DataSource = GET_RECORDS("SELECT * FROM Main_Record", Connect)
    End Sub

    Sub loadcostingform()

        Dim dt As New DataTable
        cmd = New OleDbCommand
        dt.Clear()
        dt = GET_RECORDS("Select * from table where chuchuchu", Connect)
        If dt.Rows.Count <= 0 Then
            MsgBox("No data found!")
            Exit Sub
        Else
            'rpt = New ReportDocument
            Dim rpt As New ReportDocument
            rpt.Load(Application.StartupPath & "/NameNgReportMO.rpt")
            rpt.SetDataSource(dt)
            CrystalReportViewer1.ReportSource = rpt
            CrystalReportViewer1.RefreshReport()
        End If

    End Sub
End Class
