
 If WScript.Arguments.length = 0 Then 
    Set objShell = CreateObject("Shell.Application")
    'Pass a bogus argument, say [ uac]
        objShell.ShellExecute "wscript.exe", Chr(34) & _
          WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 1
    Else
        Set shell = CreateObject ("WScript.Shell" )

        verificaRar()
        a = MsgBox("verificou rar")
        donwloadIntegrador()
        
        MatarProcessos()

        FinalizarKeepAlive()
        
        DeletarVersaoAntiga()
        
        Extrair()
        
        IniciarIntegrador()
        
        IniciaKeepAlive()
End If

Sub HTTPDownload( myURL, myPath )
 
    Dim i, objFile, objFSO, objHTTP, strFile, strMsg    
    Const ForReading = 1, ForWriting = 2, ForAppending = 8
    
    Set objFSO = CreateObject( "Scripting.FileSystemObject" )   
    If objFSO.FolderExists( myPath ) Then
        strFile = objFSO.BuildPath( myPath, Mid( myURL, InStrRev( myURL, "/" ) + 1 ) )
    ElseIf objFSO.FolderExists( Left( myPath, InStrRev( myPath, "\" ) - 1 ) ) Then
        strFile = myPath
    Else
        WScript.Echo "ERRO: Pasta de destino nao encontrada."
        Exit Sub
    End If
    
    Set objFile = objFSO.OpenTextFile( strFile, ForWriting, True )
    Set objHTTP = CreateObject( "WinHttp.WinHttpRequest.5.1" )
    
    objHTTP.Open "GET", myURL, False
    objHTTP.Send

    For i = 1 To LenB( objHTTP.ResponseBody )
        objFile.Write Chr( AscB( MidB( objHTTP.ResponseBody, i, 1 ) ) )
    Next

    objFile.Close()

End Sub

Sub DonwloadNovaVersion()

    On Error Resume Next

    Set shellTemporario = WScript.CreateObject("Wscript.shell")
    Retorno = shellTemporario.Run("bitsadmin /transfer myDownloadJob /download /priority normal https://app.vidadesindico.com.br/linear/ilp.rar %temp%\integradorUpdate\ilpt.rar" , 0 , true)
    

    If err.Number <> 0 Then

        a = MsgBox("Problema na fase de download por favor desative o seu antivirus")
        WScript.Quit
    End If

End sub

Sub VerificaRar()
    
    On Error Resume Next
    Set shell = CreateObject ("WScript.Shell" )
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If not objFSO.FileExists("C:\Windows\System32\rar.exe") Then
    
        Set shellTemporario = WScript.CreateObject("Wscript.shell")
        Retorno = shellTemporario.Run("bitsadmin /transfer myDownloadJob /download /priority normal https://ca1.vidadesindico.com.br/download/rar.exe C:\Windows\System32\rar.exe", 0 ,True)
        

        If err.Number <> 0 Then

          a = MsgBox("Problema na fase de download(Rar) por favor desative o seu antivirus")
          WScript.Quit

        End If

    End If

End Sub

Sub donwloadIntegrador()
    
    Dim caminhoUpdate, shell, folder
    
    set shell = CreateObject("WScript.shell")
    set folder = CreateObject("Scripting.FileSystemObject")
    caminhoUpdate = shell.ExpandEnvironmentStrings("%TEMP%") & "\integradorUpdate"

    if folder.FolderExists(caminhoUpdate) Then 
        
        IF not folder.FileExists("%TEMP%\integradorUpdate\ilpt.rar") Then  

            
            DonwloadNovaVersion()
            
        ElseIf FileExists("%TEMP%\integradorUpdate\ilpt.rar") Then 

            folder.DeleteFile("%TEMP%\integradorUpdate\ilpt.rar")
            
            DonwloadNovaVersion()

        End If

    ElseIf not folder.FolderExists(caminhoUpdate) Then
         
         folder.CreateFolder(caminhoUpdate)
         donwloadIntegrador()
    End if  

    
End Sub

Sub DeletarVersaoAntiga()

    
    On Error Resume Next
    Dim objFSO, folder , caminho86 , caminho
    set objFSO=CreateObject("Scripting.FileSystemObject")
    Const DeleteReadOnly = TRUE

    caminho = "C:\Program Files"
    caminho86 = "C:\Program Files (x86)"

    If objFSO.FolderExists(caminho86) Then

        Folder="C:\Program files (x86)\Integrador Linear*"
        objFSO.DeleteFolder(folder)       
    
    ElseIf  objFSO.FolderExists(caminho) and objFSO.FolderExists(caminho86) <> True Then
    
        Folder="C:\Program files\Integrador Linear*"
        objFSO.DeleteFolder(folder)          
    
    End If
    
End Sub

Sub MatarProcessos()

    Set shellTemporario = WScript.CreateObject("Wscript.shell")
    Retorno = shellTemporario.Run("taskkill /f /im IntegradorVDSGuaritaHCS.exe", 0 ,True)
    
End sub

Sub IniciarIntegrador()
    WScript.Sleep(2000)
    Dim objFSO ,caminho86 , caminho, aux
    set objFSO = CreateObject("Scripting.FileSystemObject")
    
    caminho = "C:\Program Files"
    caminho86 = "C:\Program Files (x86)"


    If objFSO.FolderExists(caminho86) Then

        Shell.Run("""C:\Program Files (x86)\Integrador Linear\IntegradorVDSGuaritaHCS.exe")
        
    ElseIf  objFSO.FolderExists(caminho) and objFSO.FolderExists(caminho86) <> True Then
        Shell.Run("""C:\Program Files\Integrador Linear\IntegradorVDSGuaritaHCS.exe")

    End If

End Sub

Sub Extrair()
    
    Dim objFSO ,caminho86 , caminho
    set objFSO = CreateObject("Scripting.FileSystemObject")
    Set shellTemporario = WScript.CreateObject("Wscript.shell")
    
    
    caminho = "C:\Program Files"
    caminho86 = "C:\Program Files (x86)"

    If objFSO.FolderExists(caminho86) Then
        
        Retorno = shellTemporario.Run("rar x %TEMP%\integradorUpdate\ilpt.rar ""C:\Program Files (x86)", 0 , True) 
        
    ElseIf  objFSO.FolderExists(caminho) and objFSO.FolderExists(caminho86) <> True Then
        
        Retorno = shellTemporario.Run("rar x %TEMP%\integradorUpdate\ilpt.rar ""C:\Program Files" , 0 , True)
        
    End If

End Sub

Sub FinalizarKeepAlive()

    Dim objFolder, caminho
    caminho = "C:\Program Files\Integrador Linear\Integrador Tray\IntegradorTray.exe"
    caminho86 = "C:\Program Files (x86)\Integrador Linear\Integrador Tray\IntegradorTray.exe"
    Set objFolder = CreateObject( "Scripting.FileSystemObject" )   

    If objFolder.FileExists(caminho) = True or objFolder.FileExists(caminho86) = True Then

        Set shellTemporario = WScript.CreateObject("Wscript.shell")
        Retorno = shellTemporario.Run("taskkill /f /im IntegradorTray.exe", 0 ,True)
    
    End If

End Sub

Sub IniciaKeepAlive()

    Dim objFolder, caminho
    caminho = "C:\Program Files\Integrador Linear\Integrador Tray\IntegradorTray.exe"
    caminho86 = "C:\Program Files (x86)\Integrador Linear\Integrador Tray\IntegradorTray.exe"
    Set objFolder = CreateObject( "Scripting.FileSystemObject" )   
    Set shellTemporario = WScript.CreateObject("Wscript.shell")

    If objFolder.FileExists(caminho) = True Then

        Retorno = shellTemporario.Run("""" & caminho, 0 ,True)
    
    ElseIf objFolder.FileExists(caminho86) Then
        
        Retorno = shellTemporario.Run("""" & caminho86, 0 ,True)

    End If

End Sub
