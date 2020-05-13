#include<File.au3>
#include<Array.au3>
#include <INet.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>

Main()

Func Main()

   Login()
   Export_HCE()
   Salir()

EndFunc

Func Login()

   Send("#r")
   Sleep(1000)
   Send("C:\Vital\Bin\Vital.exe")
   Sleep(1000)
   Send("{ENTER}")
   WinWaitActive ("Conexión")
   Sleep(1000)
   WinActive ("Conexión","")
   Sleep(1000)
   MouseClick("left",707,373,2,10)
   Sleep(1000)
   Send("{ENTER}")
   Sleep(1000)

   If (WinActive ("Inicio de sesión") <> 0) Then
   Send("{ENTER}")
   EndIf

   MouseClick("left",1071,566,1,10)
   Sleep(1000)
   Send("{DEL}")
   Sleep(1000)
   WinWaitActive ("Conexión con SAHI en APOLO")
   Send("yeilet01")
   Sleep(1000)
   MouseClick("left",1018,610,1,10)
   Sleep(1000)
   Send("Azul")
   Send("{ENTER}")

EndFunc


Func Export_HCE()

   ;Ingreso al modulo
   WinWaitActive ("Sistema de Administración Hospitalaria Integrada")
   WinActivate("Sistema de Administración Hospitalaria Integrada")
   Sleep(1000)
   MouseClick("left",99,40,1,10)
   Sleep(1000)
   Send("Azul")
   Send("{ENTER}")
   Sleep(2000)

   If (WinWaitActive ("2")) Then
	  Send("{ENTER}")
   EndIf

   ;Inicio Exportar varias HC

   ;Cargo el csv
   Global $CSV[1]
   Global $intCount = 0
   Global $RuteArchive = "C:\RPA\RPA-DES\Fuentes\UsuariosHCE.csv"

   _FileReadToArray($RuteArchive, $CSV,Default,",")
   $intLineCount = _FileCountLines($RuteArchive) - 1

While $intCount <= $intLineCount

   $intCount = $intCount + 1
   $string = FileReadLine($RuteArchive,$intCount)
   $input = StringSplit($string,",",1)
   Global $NombreUser = $input[1]    ;Nombre del usuario
   Global $DocumentoUser = $input[2] ;Documento del usuario
   Global $EmailUser = $input[3]     ;Email del usuario

   ;Dentro del modulo
   MouseClick("left",579,49,1,10)
   Sleep(1000)
   WinWaitActive("Historia Clínica Electrónica") ;Ventana Historia CLinica
   Sleep(1000)
   MouseClick("left",297,48,1,10)
   Sleep(1000)
   WinWaitActive("Datos Generales")
   WinActive ("Datos Generales")
   MouseClick("left",974,389,1,10)
   Sleep(1000)
   MouseClick("left",782,418,2,10)
   Sleep(1000)
   Send("{DEL}")
   Sleep(1000)
   Send($DocumentoUser)
   Sleep(1000)
   Send("{ENTER}")
   Sleep(1000)
   MouseClick("left",836,461,2,10)
   Sleep(1000)
   WinWaitActive ("Historia Clínica Electrónica","Datos Generales")
   Sleep(1000)
   MouseClick("left",588,734,1,10)
   Sleep(4000)
   MouseClick("left",741,774,1,10)
   Sleep(1000)
   MouseClick("left",715,690,1,10)
   Sleep(1000)
   MouseClick("left",591,448,1,10) ;Autorizaciones de atencion
   Sleep(1000)
   MouseClick("left",591,640,1,10) ;Conciliacion de medicamentos
   Sleep(1000)
   MouseClickDrag("left",966,361,966,429,10)
   Sleep(1000)
   MouseClick("left",590,543,1,10) ; Controles
   Sleep(1000)
   MouseClick("left",590,528,1,10); Control sistema
   Sleep(1000)
   MouseClick("left",590,512,1,10) ;Control transfusion x1
   Sleep(1000)
   MouseClick("left",590,494,1,10) ;Control transfusion x2
   Sleep(1000)
   MouseClickDrag("left",966,429,966,497,10)
   Sleep(1000)
   MouseClick("left",590,528,1,10) ;Formulacion de hemoderivados
   Sleep(1000)
   MouseClick("left",590,543,1,10) ;Formulacion de medicamentos
   Sleep(1000)
   MouseClickDrag("left",966,497,966,572,10)
   Sleep(1000)
   MouseClick("left",590,367,1,10) ;Notas
   Sleep(1000)
   MouseClick("left",590,384,1,10) ;Notas administrativas
   Sleep(1000)
   MouseClick("left",590,399,1,10) ;Notas  de atencion
   Sleep(1000)
   MouseClickDrag("left",966,572,966,630,10)
   Sleep(1000)
   MouseClick("left",590,448,1,10) ;Resumen del servicio
   Sleep(1000)
   MouseClick("left",590,655,1,10) ;PDF
   Sleep(1000)
   MouseClick("left",940,690,1,10) ;Clic exportar HC PDF
   Sleep(1000)
   WinWaitActive("  Exportar HCE en PDF")
   MouseClick("left",721,294,1,10);Seleccionar toda la historia
   Sleep(1000)
   MouseClick("left",1180,304,1,10);Listo
   Sleep(1000)
   WinWaitActive("Historia Clínica - Datos Generales","Finalizó proceso exportación de HCE en PDF, resultados en el archivo:") ;Ventana de HC generada
   Sleep(1000)
   Send("{ENTER}")
   Sleep(2000)
   MouseClick("left",1335,297,1,10)
   Sleep(1000)
   WinActivate("Historia Clínica Electrónica")
   Sleep(1000)

   Compress_Archive()
   Sleep(1000)

   Global $SmtpServer = "WINPVTIDIRSYNC.husi.javeriana.edu.co"              ; Direccion del servidor SMTP
   Global $FromName = "BOT"                                                 ; Nombre de quien envia
   Global $FromAddress = "BotHCE@husi.org.co"                               ; Direccion de correo de quien envia
   Global $ToAddress = $EmailUser                                           ; Correo destino
   Global $Subject = "Historia Clinica HUSI"                                ; Asunto
   Global $Body = "Buen dia Adjunto encontrara la historia clinica solicitada del paciente:" & $NombreUser        ; Body

   $filepath1 = "C:\RPA\RPA-DES\Zipped"
   $search3 = _FileListToArrayRec ($filepath1 & "",$NombreUser & "*-*" & $DocumentoUser & "*.zip", $FLTAR_FILES, -1, $FLTAR_SORT, 2)
   Global $FileZipped = $search3[1]

   Global $AttachFiles = $FileZipped                                        ; Ruta del archivo a adjuntar
   Global $CcAddress = ""                                                   ; Copia de envio
   Global $BccAddress = ""
   Global $Importance = "Normal"                                            ; Prioridad del mensaje: "High", "Normal", "Low"
   Global $Username = "##########"                               ; Direccion del usuario que envia el correo
   Global $Password = "##########"                                     ; Contraseña del usuario que envia el correo
   Global $IPPort = 25                                                      ; Puerto usado para enviar el correo
   Global $ssl = 0                                                          ; SSL


   Global $oMyRet[2]
   Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")
   $rc = _INetSmtpMailCom($SmtpServer, $FromName, $FromAddress, $ToAddress, $Subject, $Body, $AttachFiles, $CcAddress, $BccAddress, $Importance, $Username, $Password, $IPPort, $ssl)
   If @error Then
      MsgBox(0, "Error sending message", "Error code:" & @error & "  Description:" & $rc)
   EndIf

   Delete()

WEnd
  ;Fin Exportar HC


EndFunc

Func Compress_Archive()

   $filepath2 = "\\Winpvtiflsrv01\HCE_EXPORT"
   $search1 = _FileListToArrayRec ($filepath2 & "",$NombreUser & "*-*" & $DocumentoUser & "*", $FLTAR_FOLDERS , -1, $FLTAR_SORT, 1)
   $search2 = _FileListToArrayRec ($filepath2 & "",$NombreUser & "*-*" & $DocumentoUser & "*", $FLTAR_FOLDERS , -1, $FLTAR_SORT, 2)


   $Size1 = UBound($search1)-1
   $String1 = $search1[$Size1]
   Global $FileName = $String1

   $Size2 = UBound($search2)-1
   $String2 = $search2[$Size2]
   Global $FileNameFullPath = $String2


   $sZip = "C:\RPA\RPA-DES\Zipped\" & $FileName & ".zip"
   $sFile = $FileNameFullPath
   $list = "C:\RPA\RPA-DES\Zipped\"& $FileName & ".zip"

   For $i = 0 to UBound($list, 1) - 1
      MsgBox(0, '[' & $i & ']', $list[$i],2)
   Next

      ;sZip
      If not StringLen(Chr(0)) Then Return SetError(1)
      Local $sHeader = Chr(80) & Chr(75) & Chr(5) & Chr(6), $hFile
      For $i = 1 to 18
        $sHeader &= Chr(0)
      Next
      $hFile = FileOpen($sZip, 2)
      FileWrite($hFile, $sHeader)
      FileClose($hFile)

      If not StringLen(Chr(0)) Then Return SetError(1)
      If not FileExists($sZip) or not FileExists($sFile) Then Return SetError(2)
      Local $oShell = ObjCreate('Shell.Application')
      If @error or not IsObj($oShell) Then Return SetError(3)
      Local $oFolder = $oShell.NameSpace($sZip)
      If @error or not IsObj($oFolder) Then Return SetError(4)
      $oFolder.CopyHere($sFile)
      Sleep(500)


      If not StringLen(Chr(0)) Then Return SetError(1)
      If not FileExists($sZip) Then Return SetError(2)
      Local $oShell = ObjCreate('Shell.Application')
      If @error or not IsObj($oShell) Then Return SetError(3)
      Local $oFolder = $oShell.NameSpace($sZip)
      If @error or not IsObj($oFolder) Then Return SetError(4)
      Local $oItems = $oFolder.Items()
      If @error or not IsObj($oItems) Then Return SetError(5)
      Local $i = 0
      For $o in $oItems
        $i += 1
      Next
      Local $aNames[$i + 1]
      $aNames[0] = $i
      $i = 0
      For $o in $oItems
        $i += 1
        $aNames[$i] = $oFolder.GetDetailsOf($o, 0)
      Next
      Return $aNames
EndFunc




Func _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject = "", $as_Body = "", $s_AttachFiles = "", $s_CcAddress = "", $s_BccAddress = "", $s_Importance="Normal", $s_Username = "", $s_Password = "", $IPPort = 25, $ssl = 0)
      Local $objEmail = ObjCreate("CDO.Message")
      $objEmail.From = '"' & $s_FromName & '" <' & $s_FromAddress & '>'
      $objEmail.To = $s_ToAddress
      Local $i_Error = 0
      Local $i_Error_desciption = ""
      If $s_CcAddress <> "" Then $objEmail.Cc = $s_CcAddress
      If $s_BccAddress <> "" Then $objEmail.Bcc = $s_BccAddress
	  $objEmail.Subject = $s_Subject
      If StringInStr($as_Body, "<") And StringInStr($as_Body, ">") Then
	     $objEmail.HTMLBody = $as_Body
      Else
	     $objEmail.Textbody = $as_Body & @CRLF
      EndIf
	  If $s_AttachFiles <> "" Then
        Local $S_Files2Attach = StringSplit($s_AttachFiles, ";")
        For $x = 1 To $S_Files2Attach[0]
            $S_Files2Attach[$x] = _PathFull($S_Files2Attach[$x])
            ConsoleWrite('@@ Debug(62) : $S_Files2Attach = ' & $S_Files2Attach & @LF & '>Error code: ' & @error & @LF) ;### Debug Console
            If FileExists($S_Files2Attach[$x]) Then
                $objEmail.AddAttachment ($S_Files2Attach[$x])
            Else
                ConsoleWrite('!> File not found to attach: ' & $S_Files2Attach[$x] & @LF)
                SetError(1)
                Return 0
            EndIf
        Next
    EndIf
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $s_SmtpServer
    If Number($IPPort) = 0 then $IPPort = 25
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $IPPort
    ;Authenticated SMTP
    If $s_Username <> "" Then
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
    EndIf
    If $ssl Then
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    EndIf
    ;Update settings
    $objEmail.Configuration.Fields.Update
    ; Set Email Importance
    Switch $s_Importance
        Case "High"
            $objEmail.Fields.Item ("urn:schemas:mailheader:Importance") = "High"
        Case "Normal"
            $objEmail.Fields.Item ("urn:schemas:mailheader:Importance") = "Normal"
        Case "Low"
            $objEmail.Fields.Item ("urn:schemas:mailheader:Importance") = "Low"
    EndSwitch
    $objEmail.Fields.Update
    ; Sent the Message
    $objEmail.Send
    If @error Then
        SetError(2)
        Return $oMyRet[1]
    EndIf
    $objEmail=""
   EndFunc   ;==>_INetSmtpMailCom
  ;
  ;
  ; Com Error Handler
Func MyErrFunc()
     $HexNumber = Hex($oMyError.number, 8)
     $oMyRet[0] = $HexNumber
     $oMyRet[1] = StringStripWS($oMyError.description, 3)
     ConsoleWrite("### COM Error !  Number: " & $HexNumber & "   ScriptLine: " & $oMyError.scriptline & "   Description:" & $oMyRet[1] & @LF)
     SetError(1); something to check for when this function returns
    Return
 EndFunc   ;==>MyErrFunc}

 Func Delete()

   FileDelete($FileZipped)

EndFunc

Func Salir()


EndFunc
