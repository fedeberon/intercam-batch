Public Class ConsoleOut

    Private _progressBarCounter As Integer = 0
    'Private _spinner As String = "⣾⣽⣻⢿⡿⣟⣯⣷"
    Private _spinner As String = "/-\|"

    Public Sub Print(ByVal mensaje As String, Optional ByVal insertarNuevaLinea As Boolean = True)
        Console.WriteLine(mensaje)
    End Sub

    Public Sub UpdateLastLine(ByVal mensaje As String)
        Console.CursorLeft = 0
        Console.WriteLine(mensaje)
        If Console.CursorTop > 0 Then
            Console.CursorTop -= 1
        End If
    End Sub

    Public Function ProgressBarStep() As String
        Dim resChar As String = _spinner.Substring(_progressBarCounter, 1)
        _progressBarCounter += 1
        If _progressBarCounter > 3 Then _progressBarCounter = 0

        Return resChar
    End Function
End Class
