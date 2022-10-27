If WScript.Arguments.length = 0 Then 
   Set objShell = CreateObject("Shell.Application")
   'Pass a bogus argument, say [ uac]
       objShell.ShellExecute "wscript.exe", Chr(34) & _
         WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 1
    Else

    AtualizaIntegradorService("https://app.vidadesindico.com.br/linear/is.rar")
    
    MsgBox("fim")
 End If


Sub AtualizaIntegradorService(URLservice)
    
    dim caminho86 , caminho, caminhoTEMP
    Set objFolder = CreateObject("Scripting.FileSystemObject")
    Set shellTemporario = WScript.CreateObject("Wscript.shell")

    Retorno = shellTemporario.Run("net STOP ""Integrador Service""" , 0 , True)'Para Integrador Service(Se estiver rodando)

    caminho = shellTemporario.ExpandEnvironmentStrings("C:\Program Files\Integrador Service")
    caminho86 = shellTemporario.ExpandEnvironmentStrings("C:\Program Files (x86)\Integrador Service")
    caminhoTEMP = shellTemporario.ExpandEnvironmentStrings("%TEMP%\integradorUpdate\Integrador Service")

    
    If objFolder.FolderExists(caminho) = True Or objFolder.FolderExists(caminho86) = True Then

        
        If objFolder.FolderExists(caminho) = true  and objFolder.FolderExists(caminho86) = False Then
            caminho = shellTemporario.ExpandEnvironmentStrings("C:\Program Files\Integrador Service\*")
            objFolder.DeleteFolder(caminho)
            objFolder.DeleteFile(caminho)
        
        ElseIf objFolder.FolderExists(caminho86) = True Then

            caminho86 = shellTemporario.ExpandEnvironmentStrings("C:\Program Files (x86)\Integrador Service\*")
            objFolder.DeleteFolder(caminho86)
            objFolder.DeleteFile(caminho86)

        End IF
        
        
        Retorno = shellTemporario.Run("bitsadmin /transfer myDownloadJob /download /priority Foreground " & URLservice & " %temp%\integradorUpdate\is.rar" , 0 , true)
        
        'Estamos mudando a variaveis caminho e caminho 86 pois vamos extrair o rar que fizemos o download nessa mesma SUB
        caminho = "C:\Program Files\"
        caminho86 = "C:\Program Files (x86)\"
        

        IF objFolder.FolderExists(caminhoTEMP) = true Then
            caminhoTEMP = shellTemporario.ExpandEnvironmentStrings("%TEMP%\integradorUpdate\Integrador Service\*")
            objFolder.DeleteFolder(caminhoTEMP)
            objFolder.DeleteFile(caminhoTEMP)
            
        End IF

        If objFolder.FolderExists(caminho86) = True Then
            
            Retorno = shellTemporario.Run("rar x %TEMP%\integradorUpdate\is.rar  %TEMP%\integradorUpdate\", 0 , True)
            Retorno = shellTemporario.Run("xcopy /s ""%TEMP%\integradorUpdate\Integrador Service"" ""C:\Program Files (x86)\Integrador Service"" /y", 0 , True)
           
            'Retorno = shellTemporario.Run("rar x %TEMP%\integradorUpdate\isn.rar ""C:\Program Files (x86)", 0 , True)
        
        ElseIf objFolder.FolderExists(caminho) = True and objFolder.FolderExists(caminho86) = False Then
            

            
            Retorno = shellTemporario.Run("rar x %TEMP%\integradorUpdate\is.rar  %TEMP%\integradorUpdate\", 0 , True)
            retorno = shellTemporario.Run("xcopy /s ""%TEMP%\integradorUpdate\Integrador Service"" ""C:\Program Files\"" /y ", 0 , True)
            'Retorno = shellTemporario.Run("rar x %TEMP%\integradorUpdate\isn.rar ""C:\Program Files\", 0 , True)

        End If 
        
        Retorno = shellTemporario.Run("net START ""Integrador Service""" , 0 , True)

    End If 
    
 End Sub

