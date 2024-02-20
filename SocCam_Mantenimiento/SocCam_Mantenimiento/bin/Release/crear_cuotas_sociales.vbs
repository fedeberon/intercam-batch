Const ForAppending = 8

Dim Mes
Mes = DatePart("m", Date)
Dim Anio
Anio = DatePart("yyyy", Date)

Dim Com
Com = "SocCam_Mantenimiento.exe --crear-cuotas-sociales ccb nf produccion p" & 1 & " a" & Anio & " --log C:\Users\Ema\Desktop"

' Registro: Mostrar el valor de la variable Com
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objLogFile = objFSO.CreateTextFile("C:\Users\Ema\Desktop\log.txt", True)
objLogFile.WriteLine "El valor de la variable Com es: " & Com
objLogFile.Close

Dim exitCode
Dim outputText  ' Variable para almacenar la salida del comando

' Llamar a la función EjecutarComando y pasar la variable outputText por referencia
exitCode = EjecutarComando(Com, outputText)

' Registro: Mostrar la salida del comando
Set objLogFile = objFSO.OpenTextFile("C:\Users\Ema\Desktop\log.txt", ForAppending, True)
objLogFile.WriteLine "La salida del comando es: " & outputText
objLogFile.Close

' Registro: Mostrar el código de salida del comando
Set objLogFile = objFSO.OpenTextFile("C:\Users\Ema\Desktop\log.txt", ForAppending, True)
objLogFile.WriteLine "El código de salida del comando es: " & exitCode
objLogFile.Close

' Mostrar la salida del comando en un MsgBox
MsgBox "La salida del comando es: " & outputText

Function EjecutarComando(comando, ByRef outputText)
    ' Ejecutar un comando de sistema y capturar su salida
    ' Argumentos:
    '   comando [String]: Comando a ejecutar
    '   outputText [String]: Variable para almacenar la salida del comando
    ' 
    ' Retorna:
    '   Código de salida del comando, generalmente 0 si terminó correctamente, otro número si falló
    Dim oShell, oExec
    Set oShell = WScript.CreateObject("WScript.Shell")
    Set oExec = oShell.Exec(comando)

    ' Leer la salida del comando línea por línea
    Do While Not oExec.StdOut.AtEndOfStream
        outputText = outputText & oExec.StdOut.ReadLine() & vbCrLf
    Loop

    ' Devolver el código de salida del comando
    EjecutarComando = oExec.ExitCode
End Function

' Mostrar el código de salida del comando en un MsgBox
MsgBox "El codigo de salida del comando es: " & exitCode