;=========================================================================================
;
;Descripción:     Exportación y compresión de historias clinicas electronicas en SAHI
;
;Autor:          Johan Sebastian Ciprian Anzola
;Fecha:          4 de junio de 2020
;Notas:          Requiere 7zip instalado
;Version:        1.0.0
;Cuenta SAHI:    BOTHCE02
;=========================================================================================

#include<File.au3>
#include<Array.au3>
#include <INet.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>

Main()                                           ;Ejecución del proceso secuencial

Func Main()

   Login()                                       ;Apertura y Autenticación en SAHI
   Export_HCE()                                  ;Exportación, Compresión y envio de correo electronico de la HCE
   Salir()                                       ;Cerrar SAHI

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
   MouseClick("left",696,389,2,10)
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
   WinWaitActive ("Conexión con HSI_PRI en TYCHO") ;Conexión con HSI_PRI en TYCHO
   Send("BOTHCE02")                                ;Usuario SAHI
   Sleep(1000)
   MouseClick("left",1018,610,1,10)
   Sleep(1000)
   Send("4fgrsX")                                 ;Contraseña SAHI
   Send("{ENTER}")

EndFunc


Func Export_HCE()

   ;Ingreso al modulo
   WinWaitActive ("Sistema de Administración Hospitalaria Integrada")
   WinActivate("Sistema de Administración Hospitalaria Integrada")
   Sleep(1000)
   MouseClick("left",15,37,1,10)                 ;Clic en el modulo de historia clinica
   Sleep(1000)
   Send("4fgrsX")
   Send("{ENTER}")
   Sleep(4000)

   ;Inicio Exportar varias HCE

   ;Cargar el csv
   Global $CSV[1]
   Global $intCount = 0
   Global $RuteArchive = "C:\Users\s-botoearhce01\RPA\RPA-PRO\Fuentes\UsuariosHCE - Produccion - BOT1.csv"    ;Ruta donde se encuentra el archivo CSV

   _FileReadToArray($RuteArchive, $CSV,Default,",")
   $intLineCount = _FileCountLines($RuteArchive) - 1

While $intCount <= $intLineCount

   $intCount = $intCount + 1
   $string = FileReadLine($RuteArchive,$intCount)
   $input = StringSplit($string,",",1)
   Global $DocumentoUser = $input[1]      ;Documento del usuario
   Global $NombreUser = $input[2]         ;Nombre del usuario
   Global $EmailUser = $input[3]          ;Email del usuario

   WriteLogI() ;Inicio de Exportación HCE, Log Inicio

   ;Dentro del modulo
   MouseClick("left",579,49,1,10)
   Sleep(1000)
   WinWaitActive("Historia Clínica Electrónica") ;Ventana Historia Clinica
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
   Send($DocumentoUser)                        ;Se escribe el documento del paciente en el cuadro de busqueda de SAHI, modulo de HCE
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
   MouseClick("left",715,690,1,10) ;Marcar todas las casillas HC
   Sleep(1000)
   MouseClick("left",940,690,1,10) ;Clic exportar HC PDF
   Sleep(1000)
   WinWaitActive("  Exportar HCE en PDF")
   MouseClick("left",721,294,1,10);Seleccionar toda la historia
   Sleep(1000)
   MouseClick("left",1180,304,1,10);Listo
   Sleep(1000)
   WinWaitActive("Historia Clínica - Datos Generales","Finalizó proceso exportación de HCE en PDF, resultados en el archivo:") ;Ventana de HCE generada
   Sleep(1000)
   Send("{ENTER}")
   Sleep(2000)
   MouseClick("left",1335,297,1,10)
   Sleep(1000)
   WinActivate("Historia Clínica Electrónica")
   Sleep(1000)

   ;Compresión Archivo
   Compress_Archive()
   Sleep(1000)
   MsgBox(0,"BOT-HCE", "Realizando envio de notificación", 4)

   Global $SmtpServer = "WINPVTIDIRSYNC.husi.javeriana.edu.co"              ; Direccion del servidor SMTP
   Global $FromName = "BOT"                                                 ; Nombre de quien envia
   Global $FromAddress = "historiasclinicas@husi.org.co"                    ; Direccion de correo de quien envia
   Global $ToAddress = $EmailUser                                           ; Correo destino
   Global $Subject = "Historia clínica paciente: " & $DocumentoUser         ; Asunto
   Global $Body = "Historia clinica exportada y comprimida..."              ; Body
   Global $AttachFiles = ""                                                 ; Ruta del archivo a adjuntar
   Global $CcAddress = ""                                                   ; Copia de envio
   Global $BccAddress = ""
   Global $Importance = "Normal"                                            ; Prioridad del mensaje: "High", "Normal", "Low"
   Global $Username = "s-botoearhce01@husi.org.co"                          ; Direccion del usuario que envia el correo
   Global $Password = "egn+a:FWzpUqE,|T^bg@t/861fxIy6H4"                    ; Contraseña del usuario que envia el correo
   Global $IPPort = 25                                                      ; Puerto usado para enviar el correo
   Global $ssl = 0                                                          ; SSL


   Global $oMyRet[2]
   Global $oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")
   $rc = _INetSmtpMailCom($SmtpServer, $FromName, $FromAddress, $ToAddress, $Subject, $Body, $AttachFiles, $CcAddress, $BccAddress, $Importance, $Username, $Password, $IPPort, $ssl)
   If @error Then
      MsgBox(0, "Error sending message", "Error code:" & @error & "  Description:" & $rc)
   EndIf

   WriteLogF() ;Fin de exportación HCE, Log Fin

WEnd
  ;Fin Exportar HCE


EndFunc

Func Compress_Archive()

   MsgBox(0,"BOT-HCE", "Realizando busqueda del archivo...", 3)
   $filepath2 = "\\Winpvtiflsrv01\HCE_EXPORT"
   $search1 = _FileListToArrayRec ($filepath2 & "",$NombreUser & "*" & $DocumentoUser & "*", $FLTAR_FOLDERS , 0, $FLTAR_SORT, 2)

   $Size1 = UBound($search1)-1
   $String1 = $search1[$Size1]
   Global $FileNameFullPath = $String1

   $cad1 = '"' & $FileNameFullPath & '"' ;cadena con la ruta de la HCE, va entre comillas porque tiene espacios la cadena
   $cad2 = "c:\Program Files\7-Zip\7z.exe" ; Comando para ejecutar 7zip desde CMD
   $cad3 = '"' & $cad2 & '" ' & "A " ; El comando anterior $cad2 con comillas por tema de espacios en la cadena
   $CMD =  $cad3 & $cad1 & '.zip ' & $cad1 ; Comando de ejecución completo
   Sleep(1000)

   Send("#r")
   Sleep(1000)
   Send("cmd")
   Sleep(1000)
   Send("{ENTER}")
   Sleep(1000)
   WinActivate("[CLASS:ConsoleWindowClass]","") ;Enfocar la ventana de CMD abierta previamente
   Sleep(1000)
   Send($CMD)
   Sleep(1000)
   Send("{ENTER}")
   Sleep(240000)                                ;Tiempo de espera mientras el archivo de comprime
   WinClose("[CLASS:ConsoleWindowClass]","")    ;Cerrar la ventana de CMD

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
    ;Autenticación SMTP
    If $s_Username <> "" Then
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
    EndIf
    If $ssl Then
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    EndIf
    ;Modificar las configuraciones
    $objEmail.Configuration.Fields.Update
    ;Modificar la importancia del correo
    Switch $s_Importance
        Case "High"
            $objEmail.Fields.Item ("urn:schemas:mailheader:Importance") = "High"
        Case "Normal"
            $objEmail.Fields.Item ("urn:schemas:mailheader:Importance") = "Normal"
        Case "Low"
            $objEmail.Fields.Item ("urn:schemas:mailheader:Importance") = "Low"
    EndSwitch
    $objEmail.Fields.Update
    ;Envio del mensaje
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


Func WriteLogI()

    Local $Logfile = @WorkingDir & "\" & @ScriptName & ".log"
    Local $errorFile = $Logfile
    Local $LogTime = @YEAR & "/" & @MON & "/" & @MDAY & " " &@HOUR & ":" & @MIN & ":" & @SEC
	Local $TextLog = "Exportando Historia Clinica No. Documento: " & $DocumentoUser
    Local $hFileOpen = FileOpen($errorFile, 9)
    FileWriteLine($hFileOpen, $LogTime & ";" & $TextLog & " " & @CRLF)
    FileClose($hFileOpen)

EndFunc

Func WriteLogF()

    Local $Logfile = @WorkingDir & "\" & @ScriptName & ".log"
    Local $errorFile = $Logfile
    Local $LogTime = @YEAR & "/" & @MON & "/" & @MDAY & " " &@HOUR & ":" & @MIN & ":" & @SEC
	Local $TextLog = "Historia clinica No. Documento:" & $DocumentoUser & " exportada correctamente"
    Local $hFileOpen = FileOpen($errorFile, 9)
    FileWriteLine($hFileOpen, $LogTime & ";" & $TextLog & " " & @CRLF)
    FileClose($hFileOpen)

EndFunc


Func Salir()

   MouseClick("left",1908,4,1,10)
   Sleep(1000)
   MouseClick("left",1034,583,1,10)
   Sleep(1000)
   $CMD =  "taskkill /im Vital.exe"
   Run('"' & @ComSpec & '" /k ' & $CMD, @SystemDir, @SW_HIDE)

EndFunc