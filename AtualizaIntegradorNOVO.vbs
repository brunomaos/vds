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
m_ProgressBar.navigate2 "about:blank" : m_ProgressBar.width = 400 : m_ProgressBar.height = 150 : m_ProgressBar.toolbar = false : m_ProgressBar.menubar = false : m_ProgressBar.statusbar = false : m_ProgressBar.visible = True
m_ProgressBar.document.write "<body Scroll=no style='margin:0px;padding:0px;'><div style='text-align:center;'><span name='pc' id='pc'>0</span></div>"
m_ProgressBar.document.write "<div id='statusbar' name='statusbar' style='border:1px solid green;line-height:10px;height:12px;color:green;'></div>"
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
wscript.sleep 10
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

        Const URLservice = "https://app.vidadesindico.com.br/linear/isn.rar"

        Set shell = CreateObject ("WScript.Shell" )
        Dim pb
        Dim percentComplete
        Set pb = New ProgressBar

        DesabilitarRealTime()

        percentComplete = 0
        pb.SetTitle("Integrador Update")
        pb.SetText("Aguarde o integrador esta sendo Atualizado")
        pb.Show()

        pb.SetText("Verificando se existe rar instalado")
        verificaRar()
        percentComplete = percentComplete + 10
        pb.Update(percentComplete)
        
        
        pb.SetText("Fazendo download do integrador")
        pb.Update(percentComplete)
        donwloadIntegrador()
        
        percentComplete = percentComplete + 10
        pb.Update(percentComplete)
        
        pb.SetText("Fechando processos do integrador")

        MatarProcessos()

        percentComplete = percentComplete + 10
        pb.Update(percentComplete)

        pb.SetText("Finalizando KeepAlive")

        FinalizarKeepAlive()
        
        percentComplete = percentComplete + 10
        pb.Update(percentComplete)

        pb.SetText("Exluindo versão antiga")

        DeletarVersaoAntiga()
        
        percentComplete = percentComplete + 10
        pb.Update(percentComplete)

        pb.SetText("Extraindo Arquivos")

        Extrair()

        percentComplete = percentComplete + 10
        pb.Update(percentComplete)
        
        pb.SetText("Instalando Fontes")

        InstalarFontes()

        percentComplete = percentComplete + 10
        pb.Update(percentComplete)

        pb.SetText("Criando icone na area de trabalho")

        CriarIcone()

        percentComplete = percentComplete + 10
        pb.Update(percentComplete)

        pb.SetText("Iniciando integrador")
        pb.Update(percentComplete)

        ConfigurarPortasFirewall()
        IniciarIntegrador()

        pb.SetText("Verificando Integrador Service")
        pb.Update(percentComplete)
        
        AtualizaIntegradorService(URLservice)

        percentComplete = percentComplete + 20
        pb.Update(percentComplete)

        IniciaKeepAlive()
        HabilitarWindowsDefender()
        ExcluirRegraWindowsDefender()

        pb.Close()
        wscript.quit

        'IniciaKeepAlive() -> Esta versão não possui Keep ALive ainda 
End If

'DOWNLOAD RAR / INTEGRADOR
Sub DonwloadNovaVersion()

    On Error Resume Next

    Set shellTemporario = WScript.CreateObject("Wscript.shell")
    
    Retorno = shellTemporario.Run("bitsadmin /transfer myDownloadJob /download /priority Foreground https://app.vidadesindico.com.br/linear/ilpn.rar %temp%\integradorUpdate\ilpt.rar" , 0 , true)    
    Retorno = shellTemporario.Run("bitsadmin /transfer myDownloadJob /download /priority Foreground https://app.vidadesindico.com.br/linear/f.rar %temp%\integradorUpdate\f.rar" , 0 , true)
    

    If err.Number <> 0 Then

        OpenDefender()
        MsgBox("Problema na fase de download(Rar) Por favor Desabilite seu antivirus")
        
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

            OpenDefender()
            MsgBox("Problema na fase de download(Rar) Por favor desabilite a protecao em tempo real do seu antivirus")
            WScript.Quit

        End If

    End If

End Sub

'ARQUIVOS DO INTEGRADOR
Sub donwloadIntegrador()
    
    Dim caminhoUpdate, shell, folder
    
    set shell = CreateObject("WScript.shell")
    set folder = CreateObject("Scripting.FileSystemObject")
    caminhoUpdate = shell.ExpandEnvironmentStrings("%TEMP%") & "\integradorUpdate"


    if folder.FolderExists(caminhoUpdate) Then 
        
        If folder.FileExists("%TEMP%\integradorUpdate\ilpt.rar") = true and folder.FileExists("%TEMP%\integradorUpdate\f.rar") = true Then 
            
            folder.DeleteFile("%TEMP%\integradorUpdate\ilpt.rar")
            folder.DeleteFile("%TEMP%\integradorUpdate\f.rar")
            
            DonwloadNovaVersion()
    
        ElseIF not folder.FileExists("%TEMP%\integradorUpdate\ilpt.rar") and not folder.FileExists("%TEMP%\integradorUpdate\f.rar") Then  

            DonwloadNovaVersion()

        End If

    ElseIf not folder.FolderExists(caminhoUpdate) Then
         
         folder.CreateFolder(caminhoUpdate)
         donwloadIntegrador()
    End if  

    
End Sub

Sub DeletarVersaoAntiga()

    
    On Error Resume Next
    Dim objFSO, Folder , caminho86 , caminho
    set objFSO=CreateObject("Scripting.FileSystemObject")
    Const DeleteReadOnly = TRUE

    caminho = "C:\Program Files"
    caminho86 = "C:\Program Files (x86)"

    If objFSO.FolderExists(caminho86) Then

        If objFolder.FolderExists(caminho86 & "\integrador Linear") = true or objFolder.FolderExists(caminho86 & "\integrador") = True Then

            Folder ="C:\Program files (x86)\Integrador Linear*"
            objFSO.DeleteFolder(Folder)       

            Folder="C:\Program files (x86)\Integrador*"
            objFSO.DeleteFolder(Folder)       
            
        End If
        
    ElseIf  objFSO.FolderExists(caminho) and objFSO.FolderExists(caminho86) <> True Then
    
        If objFolder.FolderExists(caminho & "\integrador Linear") = true or objFolder.FolderExists(caminho & "\integrador") = True Then

            Folder="C:\Program files\Integrador Linear*"
            objFSO.DeleteFolder(Folder)       
            
            Folder="C:\Program files\Integrador*"
            objFSO.DeleteFolder(Folder)       

        End If

    End If
    
End Sub

Sub MatarProcessos()


    Set shellTemporario = WScript.CreateObject("Wscript.shell")
    Retorno = shellTemporario.Run("taskkill /f /im IntegradorVDSGuaritaHCS.exe", 0 ,True)
    Retorno = shellTemporario.Run("taskkill /f /im Integrador.exe", 0 ,True)


    
End sub

Sub IniciarIntegrador()
    WScript.Sleep(2000)
    Dim objFSO ,caminho86 , caminho, aux
    set objFSO = CreateObject("Scripting.FileSystemObject")
    
    caminho = "C:\Program Files"
    caminho86 = "C:\Program Files (x86)"

        
    If objFSO.FolderExists(caminho86) Then

        Shell.Run("""C:\Program Files (x86)\Integrador\Integrador.exe")
        
    ElseIf  objFSO.FolderExists(caminho) and objFSO.FolderExists(caminho86) <> True Then
        Shell.Run("""C:\Program Files\Integrador\Integrador.exe")

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

    If objFSO.FolderExists("c:\tmp\fontes") = true Then
        objFSO.DeleteFolder("c:\tmp\fontes")

    End IF

    Retorno = shellTemporario.Run("rar x %TEMP%\integradorUpdate\f.rar ""C:\tmp" , 0 , True)
    
End Sub

'KEEPALIVE
Sub FinalizarKeepAlive()

    Dim objFolder, caminho

    caminhoNovo = "C:\Program Files\Integrador\Tray\IntegradorTray.exe"
    caminho86Novo = "C:\Program Files (x86)\Integrador\Tray\IntegradorTray.exe"

    caminho = "C:\Program Files\Integrador Linear\Integrador Tray\IntegradorTray.exe"
    caminho86 = "C:\Program Files (x86)\Integrador Linear\Integrador Tray\IntegradorTray.exe"
    Set objFolder = CreateObject( "Scripting.FileSystemObject" )   

    If objFolder.FileExists(caminho) = True or objFolder.FileExists(caminho86) = True or objFolder.FileExists(caminhoNovo) = True or objFolder.FileExists(caminho86Novo) = True Then

        Set shellTemporario = WScript.CreateObject("Wscript.shell")
        Retorno = shellTemporario.Run("taskkill /f /im IntegradorTray.exe", 0 ,True)
    
    End If

End Sub

Sub IniciaKeepAlive()

    Dim objFolder, caminho
    caminho = "C:\Program Files\Integrador\Tray\IntegradorTray.exe"
    caminho86 = "C:\Program Files (x86)\Integrador\Tray\IntegradorTray.exe"
    Set objFolder = CreateObject( "Scripting.FileSystemObject" )   
    Set shellTemporario = WScript.CreateObject("Wscript.shell")

    If objFolder.FileExists(caminho) = True Then

        Retorno = shellTemporario.Run("""" & caminho, 0 )
    
    ElseIf objFolder.FileExists(caminho86) Then
        
        Retorno = shellTemporario.Run("""" & caminho86, 0 )

    End If

End Sub


'ATALHOS INTEGRADOR
Sub CriarIcone

    Dim atalhoIntegradorAntigo, atalhoIntegradorNovo
    Set objShell = CreateObject("WScript.Shell")
    Set shellTemporario = WScript.CreateObject("Wscript.shell")
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    strDesktop = objShell.SpecialFolders("desktop")
    Set objLink = objShell.CreateShortcut(strDesktop & "\Integrador.lnk")
    
    atalhoDesktopAntigo =  objShell.ExpandEnvironmentStrings("%PUBLIC%\Desktop\Integrador Linear.lnk")
    atalhoDesktopNovo =  objShell.ExpandEnvironmentStrings("%PUBLIC%\Desktop\Integrador.lnk")
    atalhoIntegradorAntigo = objShell.ExpandEnvironmentStrings("%APPDATA%\Microsoft\Windows\start menu\programs\startup\IntegradorVDSGuaritaHCS.lnk")
    atalhoIntegradorNovo =  objShell.ExpandEnvironmentStrings("%APPDATA%\Microsoft\Windows\start menu\programs\startup\Integrador.lnk")
    
    If objFSO.FileExists(atalhoIntegradorAntigo) = True or objFSO.FileExists(atalhoIntegradorNovo) = True Then
        

        If objFSO.FileExists(atalhoIntegradorAntigo) = true Then
            
            objFSO.DeleteFile(atalhoIntegradorAntigo)

        ElseIf objFSO.FileExists(atalhoIntegradorNovo)  = true Then
            
            objFSO.DeleteFile(atalhoIntegradorNovo)
        end if
    End If
    
    If objFSO.FileExists(atalhoDesktopAntigo)  = True Then
            
            objFSO.DeleteFile(atalhoDesktopAntigo)

    ElseIf objFSO.FileExists(atalhoDesktopNovo)  = True Then
            
            objFSO.DeleteFile(atalhoDesktopNovo)

    End If

    
    

    If objFSO.FileExists("C:\Program Files (x86)\Integrador\Integrador.exe") = True and objFSO.FileExists("C:\Program Files\Integrador\Integrador.exe") = False Then

        objLink.IconLocation = "C:\Program Files (x86)\Integrador\Resources\logo.ico"
        objLink.TargetPath = "C:\Program Files (x86)\Integrador\Integrador.exe"
    
    ElseIf objFSO.FileExists("C:\Program Files (x86)\Integrador\Integrador.exe") = False And objFSO.FileExists("C:\Program Files\Integrador\Integrador.exe") = True Then
        
        objLink.IconLocation = "C:\Program Files\Integrador\Resources\logo.ico"
        objLink.TargetPath = "C:\Program Files\Integrador\Integrador.exe"

    End If


    objLink.WorkingDirectory = "%HOMEDRIVE%%HOMEPATH%"
    objLink.Description = "Atalho integrador"
    objLink.Save
    

    Retorno = shellTemporario.Run("xcopy integrador.lnk ""%APPDATA%\Microsoft\Windows\start menu\programs\startup"" /y", 0)
    

End Sub

'FONTES
Sub InstalarFontes()

    Dim caminhoFonte, caminhoVerificacao
    Set shellTemporario = WScript.CreateObject("Wscript.shell")
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    caminhoVerificacao = shellTemporario.ExpandEnvironmentStrings("%USERPROFILE%\AppData\Local\Microsoft\Windows\Fonts\MaterialIcons-Regular.ttf")
    
    If objFSO.FileExists(caminhoVerificacao) = False Then
        CollectFonts()
        Retorno = shellTemporario.Run("xcopy ""C:\tmp\Fontes\Material Icons\MaterialIcons-Regular.ttf"" ""C:\tmp\Fontes\Inter\"" /y", 0 ,True)
        caminhoFonte = shellTemporario.ExpandEnvironmentStrings("C:\tmp")
        
        IF objFSO.FolderExists(caminhoFonte) = False Then

            objFSO.CreateFolder(caminhoFonte)
            
        End If

        InstallFonts("C:\tmp\Fontes\Inter")
    End If
    
End Sub

Sub CollectFonts

    Const FONTS = &H14&
    Const ForAppending = 8
    Dim fso
    doexist = 0
    dontexist = 0

    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.Namespace(FONTS)
    set oShell = CreateObject("WScript.Shell") 

    strSystemRootDir = oshell.ExpandEnvironmentStrings("%systemroot%")
    strFontDir = strSystemRootDir & "\fonts\"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objDictionary = CreateObject("Scripting.Dictionary")
    objDictionary.CompareMode = TextMode
    Set f1 = FSO.createTextFile("c:\Windows\fontes_instaladas.txt", ForAppending)
 
    set colItems = objfolder.Items
     For each ObjItem in ColItems
        If LCase(Right(objItem.Name, 3)) = "ttf" or _
           LCase(Right(objItem.Name, 3)) = "otf" or _
           LCase(Right(objItem.Name, 3)) = "pfm" or _
           LCase(Right(objItem.Name, 3)) = "fon" Then
            If Not objDictionary.Exists(LCase(ObjItem.Name)) Then
                objDictionary.Add LCase(ObjItem.Name), LCase(ObjItem.Name)
            End If
        End If
    Next
    For each ObjItem in ObjDictionary
        f1.writeline ObjDictionary.Item(objItem)
    Next
End Sub
 
Sub InstallFonts(Folder)

    Const FONTS = &H14&
    Const ForAppending = 8
    Dim fso
    doexist = 0
    dontexist = 0

    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.Namespace(FONTS)
    set oShell = CreateObject("WScript.Shell") 

    strSystemRootDir = oshell.ExpandEnvironmentStrings("%systemroot%")
    strFontDir = strSystemRootDir & "\fonts\"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set objDictionary = CreateObject("Scripting.Dictionary")
    objDictionary.CompareMode = TextMode
    Set f1 = FSO.createTextFile("c:\Windows\fontes_instaladas.txt", ForAppending)

 Set FontFolder = fso.getfolder(Folder)
        For Each File in FontFolder.Files
             If LCase(fso.GetExtensionName(File))="ttf" or _
                LCase(fso.GetExtensionName(File))="otf" or _
                LCase(fso.GetExtensionName(File))="pfm" or _
                LCase(fso.GetExtensionName(File))="fon" Then
                If objDictionary.Exists(lcase(fso.GetFileName(File))) then
                    doexist = doexist + 1
                Else
                    objFolder.CopyHere FontFolder & "\" & fso.GetFileName(File)
                    dontexist = dontexist + 1
                end If
            End If
        Next
        For Each SubFolder in FontFolder.subFolders
            InstallFonts SubFolder
        Next
End Sub

'Integrador Service
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

'WINDOWS DEFENDER
Sub DesabilitarRealTime()

    Set shellTemporario = WScript.CreateObject("Wscript.shell")
    Retorno = shellTemporario.Run("REG ADD ""HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection"" /v ""DisableRealtimeMonitoring"" /t REG_DWORD /d 1 /f")
    
    
End Sub

Sub HabilitarWindowsDefender()
    Set shellTemporario = WScript.CreateObject("Wscript.shell")
    Retorno = shellTemporario.Run("REG ADD ""HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection"" /v ""DisableRealtimeMonitoring"" /t REG_DWORD /d 1 /f", 0 , true)
    
End Sub

Sub ExcluirRegraWindowsDefender()
    Set shellTemporario = WScript.CreateObject("Wscript.shell")
    Retorno = shellTemporario.Run("REG DELETE ""HKLM\SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection"" /v ""DisableRealtimeMonitoring"" /f" , 0 ,true)
    
End Sub

Sub OpenDefender()

    Dim tCommand

    tCommand = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%ProgramFiles%\Windows Defender\MSASCui.exe")
    If Not (CreateObject("Scripting.FileSystemObject").FileExists(tCommand)) Then tCommand = "windowsdefender://Threatsettings"
    Set shellTemporario = WScript.CreateObject("Wscript.shell")

    shellTemporario.Run(tCommand)

End Sub

'FIREWALL
Sub ConfigurarPortasFirewall()
    
    Set objFirewall = CreateObject("HNetCfg.FwMgr")
    Set objPolicy = objFirewall.LocalPolicy.CurrentProfile
    portas = Array(9000,9001,8080,4370,1600,80,5010,8282,8181,4370,9870,8383,8686,9762,9762,9100,9765,2845,10057,5432,5433)
    Set colPorts = objPolicy.GloballyOpenPorts

    For Each a in portas
       
        dim aux

        for Each objPorta in colPorts
            
            aux = True

            If objPorta.Port = a Then
                'Se a porta já existir define variavel auxiliar igual a true
                aux = False
                Exit for

            End If


        Next    
        'Se a variavel auxiliar existir manda criar as portas
        If aux = true Then
            
            CriarPortasDeEntrada(a)
            CriarPortasDeSaida(a)
            
        End if
        
    Next
    
End Sub

Sub CriarPortasDeSaida(MyPort)

    Dim CurrentProfiles
    
    ' Protocol
    Const NET_FW_IP_PROTOCOL_TCP = 6
    Const NET_FW_IP_PROTOCOL_UDP = 17

    'Direction
    Const NET_FW_RULE_DIR_IN = 1
    Const NET_FW_RULE_DIR_OUT = 2

    'Action
    Const NET_FW_ACTION_ALLOW = 1

    ' Create the FwPolicy2 object.
    Dim fwPolicy2
    Set fwPolicy2 = CreateObject("HNetCfg.FwPolicy2")

    ' Get the Rules object
    Dim RulesObject
    Set RulesObject = fwPolicy2.Rules

    CurrentProfiles = fwPolicy2.CurrentProfileTypes


    Set NovaPorta = CreateObject("HNetCfg.FWRule")
    NovaPorta.Name = "VDC portas de acesso " & MyPort
    NovaPorta.Description = "Esta porta é utilizada para conexão com os equipamentos de controle de acesso!"
    NovaPorta.Protocol = NET_FW_IP_PROTOCOL_TCP
    NovaPorta.LocalPorts = MyPort
    NovaPorta.Direction = NET_FW_RULE_DIR_OUT
    NovaPorta.Enabled = TRUE  
    NovaPorta.Profiles = CurrentProfiles
    NovaPorta.Action = NET_FW_ACTION_ALLOW
        
    RulesObject.Add NovaPorta
     
End Sub

Sub CriarPortasDeEntrada(MyPort)
    Set objFirewall = CreateObject("HNetCfg.FwMgr")
    Set objPolicy = objFirewall.LocalPolicy.CurrentProfile
    Set objPort = CreateObject("HNetCfg.FwOpenPort")
    
    objPort.Port = MyPort
    objPort.Name = "VDC portas de acesso " & MyPort
    'objPort.Description = "Esta porta é utilizada para conexão com os equipamentos de controle de acesso!"
    objPort.Enabled = TRUE

    Set colPorts = objPolicy.GloballyOpenPorts
    errReturn = colPorts.Add(objPort)
        
End Sub

