' ============================================
' Script para criar atalho do Finanças Pessoais
' com ícone personalizado na Área de Trabalho
' ============================================

Set WshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Obtém o caminho da área de trabalho
strDesktop = WshShell.SpecialFolders("Desktop")

' Obtém o caminho do script atual
strScriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)

' Caminho do arquivo HTML
strHTMLFile = strScriptPath & "\financas.html"

' Caminho do ícone
strIconFile = strScriptPath & "\icone-financas.ico"

' Verifica se os arquivos existem
If Not objFSO.FileExists(strHTMLFile) Then
    MsgBox "Arquivo 'financas.html' não encontrado!" & vbCrLf & _
           "Certifique-se de que este script está na mesma pasta do arquivo HTML.", _
           vbExclamation, "Erro"
    WScript.Quit
End If

If Not objFSO.FileExists(strIconFile) Then
    MsgBox "Arquivo 'icone-financas.ico' não encontrado!" & vbCrLf & _
           "Converta o SVG para ICO primeiro.", _
           vbExclamation, "Erro"
    WScript.Quit
End If

' Cria o atalho
Set oShortcut = WshShell.CreateShortcut(strDesktop & "\Finanças Pessoais.lnk")
oShortcut.TargetPath = strHTMLFile
oShortcut.Description = "Sistema de Finanças Pessoais - Controle Inteligente"
oShortcut.WorkingDirectory = strScriptPath
oShortcut.IconLocation = strIconFile
oShortcut.Save

MsgBox "✓ Atalho criado com sucesso na Área de Trabalho!" & vbCrLf & vbCrLf & _
       "Nome: Finanças Pessoais" & vbCrLf & _
       "Ícone: Personalizado (nuvem + gráfico)" & vbCrLf & _
       "Local: " & strDesktop, _
       vbInformation, "Sucesso"