'On Error Resume Next

 If WScript.Arguments.length = 0 Then
    Set objShell = CreateObject("Shell.Application")
    'Pass a bogus argument, say [ uac]
        objShell.ShellExecute "wscript.exe", Chr(34) & _
          WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 1
    Else
        Set shell = CreateObject ("WScript.Shell" )

        verificaRar()
        
        donwloadIntegrador()
        WScript.Sleep(20000)
        
        MatarProcessos()
        
        DeletarVersaoAntiga()
        
        Extrair()
        
        IniciarIntegrador()
        

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

    ' HTTPDownload("<URL VIDA DE SINDICO BAT>" , shell.ExpandEnvironmentStrings("C:\Windows\System32\"))
    ' shell.Run("C:\Windows\System32\")
    On Error Resume Next
    shell.Run("bitsadmin /transfer myDownloadJob /download /priority normal https://app.vidadesindico.com.br/linear/ilp.rar %temp%\integradorUpdate\ilpt.rar")

    If err.Number <> 0 Then

        a = MsgBox("Problema na fase de download por favor desative o seu antivirus")
        
    End If

End sub

Sub VerificaRar()
    
    On Error Resume Next
    Set shell = CreateObject ("WScript.Shell" )
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If not objFSO.FileExists("C:\Windows\System32\rar.exe") Then
        
        shell.Run("bitsadmin /transfer myDownloadJob /download /priority normal https://ca1.vidadesindico.com.br/download/rar.exe C:\Windows\System32\rar.exe")
        
        
        If err.Number <> 0 Then

          a = MsgBox("Problema na fase de download por favor desative o seu antivirus")
        
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
    Shell.Run("taskkill /f /im IntegradorVDSGuaritaHCS.exe")
    WScript.Sleep(2000)
End sub

Sub IniciarIntegrador()
    WScript.Sleep(2000)
    Dim objFSO ,caminho86 , caminho, aux
    set objFSO = CreateObject("Scripting.FileSystemObject")
    
    caminho = "C:\Program Files"
    caminho86 = "C:\Program Files (x86)"

    'shell.Run("""C:\Users\Public\Desktop\Integrador Linear.lnk")

    If objFSO.FolderExists(caminho86) Then

        Shell.Run("""C:\Program Files (x86)\Integrador Linear\IntegradorVDSGuaritaHCS.exe")
        
    ElseIf  objFSO.FolderExists(caminho) and objFSO.FolderExists(caminho86) <> True Then
        aux = """C:\Program Files\Integrador Linear\IntegradorVDSGuaritaHCS.exe"
        
    End If

        

End Sub

Sub Extrair()
    
    Dim objFSO ,caminho86 , caminho
    set objFSO = CreateObject("Scripting.FileSystemObject")
    
    caminho = "C:\Program Files"
    caminho86 = "C:\Program Files (x86)"

    If objFSO.FolderExists(caminho86) Then
        
        shell.Run("rar x %TEMP%\integradorUpdate\ilpt.rar ""C:\Program Files (x86)")
        
    ElseIf  objFSO.FolderExists(caminho) and objFSO.FolderExists(caminho86) <> True Then
        
        shell.Run("rar x %TEMP%\integradorUpdate\ilpt.rar ""C:\Program Files")
        
    End If

End Sub

