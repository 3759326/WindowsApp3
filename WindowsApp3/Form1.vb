Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        SrvService("Pervasive.SQL (relational)", "Start")
        SrvService("Pervasive.SQL (transactional)", "Stop")
        MsgBox(SrvService("Pervasive.SQL (relational)", "Status"))
        My.Computer.Registry.SetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Pervasive Software\MicroKernel Server Engine\Version 10\Settings", "File Version", "0950")



    End Sub
    Function SrvService(srvName As String, srvAction As String) As String
        ServiceController1.ServiceName = srvName
        Select Case srvAction
            Case "Stop"
                Try
                    If ServiceController1.Status = 1 Then
                        MsgBox("Служба " & srvName & " уже остановленна")
                        Exit Function
                    End If
                    ServiceController1.Stop()
                    ServiceController1.Refresh()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try


            Case "Start"
                Try
                    If ServiceController1.Status = 4 Then
                        MsgBox("Служба " & srvName & " уже запущенна")
                        Exit Function
                    End If
                    ServiceController1.Start()
                    ServiceController1.Refresh()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            Case "Status"
                SrvService = ServiceController1.Status

        End Select
    End Function
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        'ListBox1.Items.AddRange(lv)
        'ListBox1.Items.Add("Harry Potter", 5.99)
    End Sub
End Class
