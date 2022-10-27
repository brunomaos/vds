Class ProgressBar
Private m_PercentComplete
Private m_CurrentStep
Private m_ProgressBar
Private m_Title
Private m_Text
Private m_StatusBarText

'Initialize defaults
Private Sub ProgessBar_Initialize
m_PercentComplete = 0
m_CurrentStep = 0
m_Title = "Progresso"
m_Text = ""
End Sub

Public Function SetTitle(pTitle)
m_Title = pTitle
End Function

Public Function SetText(pText)
m_Text = pText
End Function

Public Function Update(percentComplete)
m_PercentComplete = percentComplete
UpdateProgressBar()
End Function

Public Function Show()
Set m_ProgressBar = CreateObject("InternetExplorer.Application")
'in code, the colon acts as a line feed
m_ProgressBar.navigate2 "about:blank" : m_ProgressBar.width = 400 : m_ProgressBar.height = 100 : m_ProgressBar.toolbar = false : m_ProgressBar.menubar = false : m_ProgressBar.statusbar = false : m_ProgressBar.visible = True
m_ProgressBar.document.write "<body Scroll=no style='margin:0px;padding:0px;'><div style='text-align:center;'><span name='pc' id='pc'>0</span></div>"
m_ProgressBar.document.write "<div id='statusbar' name='statusbar' style='border:1px solid blue;line-height:10px;height:10px;color:green;'></div>"
m_ProgressBar.document.write "<div style='text-align:center'><span id='text' name='text'></span></div>"
End Function

Public Function Close()
m_ProgressBar.quit
End Function

Private Function UpdateProgressBar()
If m_PercentComplete = 0 Then
m_StatusBarText = ""
End If
For n = m_CurrentStep to m_PercentComplete - 1
m_StatusBarText = m_StatusBarText & "|"
m_ProgressBar.Document.GetElementById("statusbar").InnerHtml = m_StatusBarText
m_ProgressBar.Document.title = n & "% Completado : " & m_Title
m_ProgressBar.Document.GetElementById("pc").InnerHtml = n & "% Completado : " & m_Title
wscript.sleep 20
Next
m_ProgressBar.Document.GetElementById("statusbar").InnerHtml = m_StatusBarText
m_ProgressBar.Document.title = m_PercentComplete & "% Completado : " & m_Title
m_ProgressBar.Document.GetElementById("pc").InnerHtml = m_PercentComplete & "% Completado : " & m_Title
m_ProgressBar.Document.GetElementById("text").InnerHtml = m_Text
m_CurrentStep = m_PercentComplete
End Function

End Class

If WScript.Arguments.length = 0 Then 
   Set objShell = CreateObject("Shell.Application")
   'Pass a bogus argument, say [ uac]
       objShell.ShellExecute "wscript.exe", Chr(34) & _
         WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 1
    Else

        DesabilitarRealTime()
        
        Const URLservice = "https://app.vidadesindico.com.br/linear/is.rar"
        Set shell = CreateObject ("WScript.Shell" )
        Dim pb
        Dim percentComplete
        Set pb = New ProgressBar

        percentComplete = 0
        pb.SetTitle("Integrador Update")
        pb.SetText("Aguarde o integrador esta sendo Atualizado")
        pb.Show()

        
        pb.SetText("Verificando se existe rar instalado")
        pb.Update(percentComplete)
        
        verificaRar()
        
        percentComplete = percentComplete + 15
        pb.Update(percentComplete)
        
        pb.SetText("Fazendo download do integrador")
        pb.Update(percentComplete)
        
        donwloadIntegrador()
        
        percentComplete = percentComplete + 15
        pb.Update(percentComplete)
        
        pb.SetText("Fechando processos do integrador")
        pb.Update(percentComplete)
        
        MatarProcessos()

        percentComplete = percentComplete + 15
        pb.Update(percentComplete)

        pb.SetText("Finalizando KeepAlive")
        pb.Update(percentComplete)

        FinalizarKeepAlive()

        percentComplete = percentComplete + 10
        pb.Update(percentComplete)
        
        pb.SetText("Exluindo versão antiga")
        pb.Update(percentComplete)

        DeletarVersaoAntiga()

        percentComplete = percentComplete + 10
        pb.Update(percentComplete)
        
        pb.SetText("Extraindo Arquivos")
        pb.Update(percentComplete)
        
        Extrair()

        percentComplete = percentComplete + 15
        pb.Update(percentComplete)
        

        
        pb.SetText("Verificando Integrador Service - Aguarde")
        percentComplete = percentComplete + 2
        pb.Update(percentComplete)

        AtualizaIntegradorService(URLservice)

        pb.SetText("Iniciando integrador!")
        percentComplete = percentComplete + 3
        pb.Update(percentComplete)

        IniciarIntegrador()

        percentComplete = percentComplete + 5
        pb.Update(percentComplete)
        
        pb.SetText("Iniciando KeepAlive")
        percentComplete = percentComplete
        pb.Update(percentComplete)

        pb.Close()

        IniciaKeepAlive()

        percentComplete = 100
        pb.Update(percentComplete)
        
        HabilitarWindowsDefender()

        ExcluirRegraWindowsDefender()
        
        Wscript.quit
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
    Retorno = shellTemporario.Run("bitsadmin /transfer myDownloadJob /download /priority Foreground https://app.vidadesindico.com.br/linear/ilp.rar %temp%\integradorUpdate\ilpt.rar" , 0 , true)
    

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
        Retorno = shellTemporario.Run("bitsadmin /transfer myDownloadJob /download /priority Foreground https://ca1.vidadesindico.com.br/download/rar.exe C:\Windows\System32\rar.exe", 0 ,True)
        

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
    Set objFolder = CreateObject( "Scripting.FileSystemObject" )   
    Set shellTemporario = WScript.CreateObject("Wscript.shell")
    caminho86 = shellTemporario.ExpandEnvironmentStrings("C:\Program Files (x86)\Integrador Linear\Integrador Tray\IntegradorTray.exe")
    caminho = shellTemporario.ExpandEnvironmentStrings( "C:\Program Files\Integrador Linear\Integrador Tray\IntegradorTray.exe")

    If objFolder.FileExists(caminho) = True Then
         MsgBox(caminho)
        Retorno = shellTemporario.Run("""" & caminho, 0 )
    
    ElseIf objFolder.FileExists(caminho86) Then
         MsgBox(caminho86)
        Retorno = shellTemporario.Run("""" & caminho86, 0 )

    End If

End Sub

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

Sub DesabilitarRealTime()

    Set shellTemporario = WScript.CreateObject("Wscript.shell")
    Retorno = shellTemporario.Run("REG ADD ""HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection"" /v ""DisableRealtimeMonitoring"" /t REG_DWORD /d 1 /f")
    
    
End Sub

Sub HabilitarWindowsDefender()
    Set shellTemporario = WScript.CreateObject("Wscript.shell")
    Retorno = shellTemporario.Run("REG ADD ""HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection"" /v ""DisableRealtimeMonitoring"" /t REG_DWORD /d 1 /f" , 0 , True)
    
End Sub

Sub ExcluirRegraWindowsDefender()
    Set shellTemporario = WScript.CreateObject("Wscript.shell")
    Retorno = shellTemporario.Run("REG DELETE ""HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection"" /v ""DisableRealtimeMonitoring"" /f" , 0 , True)
    
End Sub

