Imports System
Imports System.Data
Imports System.Data.OleDb

Class Form1
    Private Shared Sub Main()
        Dim connectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "C:\Users\usilva.gti\Desktop\AXA\bd\Teste1.mdb;User Id=admin;Password=;"
        Dim queryString As String = "SELECT * from Tab1 " & "ORDER BY 3;"

        Using connection As OleDbConnection = New OleDbConnection(connectionString)
            Dim command As OleDbCommand = New OleDbCommand(queryString, connection)

            Try
                connection.Open()
                Dim reader As OleDbDataReader = command.ExecuteReader()

                While reader.Read()
                    'MessageBox.Show(vbTab & "{0}" & vbTab & "{1}" & vbTab & "{2}", reader(0), reader(1), reader(2))
                    'Console.WriteLine(vbTab & "{0}" & vbTab & "{1}" & vbTab & "{2}", reader(0), reader(1), reader(2))
                    Console.WriteLine(vbTab & "{0}" & ";" & "{1}" & ";" & "{2}" & ";" & "{3}", reader(0), reader(1), reader(2), reader(3))
                End While

                reader.Close()
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try

            Console.ReadLine()
        End Using
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Main()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Main()
    End Sub
End Class