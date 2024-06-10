#include <Constants.au3>

; MENSAJE DE WINDONS PARA CONFIRMAR LA EJECUCIÓN
Local $iAnswer = MsgBox(BitOR($MB_YESNO, $MB_SYSTEMMODAL), "Ejecución del Demonio -ETL", "Desea ejecutar el Demoio ETL. Click en Si / No?")

; CONFIRMACION DE LA NOTIFICACIÓN
If $iAnswer = 7 Then
	MsgBox($MB_SYSTEMMODAL, "Ejecución del Demonio -ETL", "OK.  Bye!")
	Exit
EndIf

; ABRIR EN NAVEGADOR EDGE - LINK DE EXCEL
Local $sURL = "https://drive.google.com/drive/u/1/folders/1pIFZ5HlhaAwMOxNPqLT_D_yUx_akFpft"
Run(@ComSpec & " /c start iexplore.exe " & $sURL)

; Esperar a que Google Drive se cargue
Sleep(6000)

; Buscar el archivo '01..-InformRefineria acion_Personal.xlsx'
Send("^f")
Sleep(500)
ClipPut("01..-Informacion_Personal.xlsx")
Send("^v")
Send("{ENTER}")
Sleep(1000)

; Seleccionar el archivo '01.-Informacion_Personal.xlsx' y Descargar
Send("{ESC 2}")
Sleep(500)
Send("{ENTER}")
Sleep(500)
Send("{TAB}")
Sleep(500)
Send("{RIGHT 1}")
Sleep(500)
Send("{ENTER}")
Sleep(500)

Local $visualStudioPath = "C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\IDE\devenv.exe"
Run($visualStudioPath)
Sleep(10000)
Send("{TAB 6}")
Sleep(1000)
Send("{ENTER}")
Sleep(5000)
; Buscar el archivo 'SISS' 
Send("^q")
Sleep(1000)
ClipPut("Integration Services Project   ")
Send("^v")
Sleep(6000)
Send("{ENTER}")
Sleep(5000)

; Colocar nombre al proyecto
Local $fechaActual = StringFormat("%02d-%02d-%04d", @MDAY, @MON, @YEAR)
Local $proyectoNombre = "ETL - Refineria (" & $fechaActual & ")"
Send($proyectoNombre)
Sleep(2000)
Send("{TAB 6}")
Sleep(500)
Send("{ENTER}")
Sleep(10000)

; Cerrar Ventana introduccion
Func CerrarVentanaIntro()
    MouseMove(283, 80, 0) ;
    MouseClick("left") ;
EndFunc
CerrarVentanaIntro()
Sleep(10000)

; COLOCACION DE COMPONENTES PARA LA EXTRACCION

; Inicializar AutoIt para mejorar la estabilidad
AutoItSetOption("WinTitleMatchMode", 2)
AutoItSetOption("MouseCoordMode", 2)
AutoItSetOption("PixelCoordMode", 2)
AutoItSetOption("SendKeyDelay", 5)
AutoItSetOption("SendKeyDownDelay", 5)

; FUNCION ARRASTRAR Y SOLTAR
Func ARRASTRARSOLTAR($startX, $startY, $endX, $endY)
    MouseMove($startX, $startY, 0)
    MouseDown("left")
    Sleep(800) 
    MouseMove($endX, $endY, 10)
    Sleep(500)
    MouseUp("left")
EndFunc

; FUNCIONES para buscar componentes
Func BuscarExcelSource()
    MouseMove(199, 113, 0) 
    MouseClick("left") 
	Sleep(1000)
	Send("^a")
	Sleep(500)
	Send("{DEL}")
    ClipPut("Excel Source")
    Send("^v")
EndFunc
Func BuscarOLEDBDestination()
    MouseMove(199, 113, 0) 
    MouseClick("left") 
	Sleep(1000)
	Send("^a")
	Sleep(500)
	Send("{DEL}")
    ClipPut("OLE DB Destination")
    Send("^v")
EndFunc

; UBICACION DE COMPONENTES
Local $DataFlowTaskInicioX = 105 
Local $DataFlowTaskInicioY = 157
Local $ExcelSourceInicioX = 107 
Local $ExcelSourceInicioY = 257
Local $OLEDestinatioInicioX = 124 
Local $OLEDestinatioInicioY = 281


; COLOCACION DE "Data Flow Task"
Local $DataFlowTaskFinX = 653 
Local $DataFlowTaskFinY = 277
ARRASTRARSOLTAR($DataFlowTaskInicioX, $DataFlowTaskInicioY, $DataFlowTaskFinX, $DataFlowTaskFinY)
Sleep(8000)
Send("{ENTER}")
Sleep(5000)

; PARTE INFORMACION PERSONAL
Func InformacionPersonal()
	BuscarExcelSource()
    Local $ExcelSourceFinX = 340 
    Local $ExcelSourceFinY = 179
    ARRASTRARSOLTAR($ExcelSourceInicioX, $ExcelSourceInicioY, $ExcelSourceFinX, $ExcelSourceFinY)
    Sleep(3000)

    BuscarOLEDBDestination()
    Local $OLEDestinatioFinX = 559 
    Local $OLEDestinatioFinY = 179
    ARRASTRARSOLTAR($OLEDestinatioInicioX, $OLEDestinatioInicioY, $OLEDestinatioFinX, $OLEDestinatioFinY)
    Sleep(3000)

    ; Enlazar
    MouseMove(417, 201, 0)
    MouseClick("left")
    Send("{F2}")
    ClipPut("Informacion_Personal")
    Send("^v")
    Sleep(1000)
    Local $ConectInicioX = 437
    Local $ConectInicioY = 227
    Local $ConectFinX = 614 
    Local $ConectFinY = 201
    ARRASTRARSOLTAR($ConectInicioX, $ConectInicioY, $ConectFinX, $ConectFinY)
    Sleep(1000)
EndFunc
InformacionPersonal()

; PARTE ESTUDIOS
Func Estudios()
	BuscarExcelSource()
    Local $ExcelSourceFinX = 340 
    Local $ExcelSourceFinY = 231
    ARRASTRARSOLTAR($ExcelSourceInicioX, $ExcelSourceInicioY, $ExcelSourceFinX, $ExcelSourceFinY)
    Sleep(1000)

    BuscarOLEDBDestination()
    Local $OLEDestinatioFinX = 559 
    Local $OLEDestinatioFinY = 231
    ARRASTRARSOLTAR($OLEDestinatioInicioX, $OLEDestinatioInicioY, $OLEDestinatioFinX, $OLEDestinatioFinY)
    Sleep(1000)

    ; Enlazar
    MouseMove(421, 258, 0)
    MouseClick("left")
    Send("{F2}")
    ClipPut("Estudios")
    Send("^v")
    Sleep(1000)
    Local $ConectInicioX = 407
    Local $ConectInicioY = 286
    Local $ConectFinX = 618 
    Local $ConectFinY = 250
    ARRASTRARSOLTAR($ConectInicioX, $ConectInicioY, $ConectFinX, $ConectFinY)
    Sleep(1000)
EndFunc
Estudios()

; PARTE UBCACION
Func Ubicacion()
	BuscarExcelSource()
    Local $ExcelSourceFinX = 340 
    Local $ExcelSourceFinY = 287
    ARRASTRARSOLTAR($ExcelSourceInicioX, $ExcelSourceInicioY, $ExcelSourceFinX, $ExcelSourceFinY)
    Sleep(1000)

    BuscarOLEDBDestination()
    Local $OLEDestinatioFinX = 559 
    Local $OLEDestinatioFinY = 287
    ARRASTRARSOLTAR($OLEDestinatioInicioX, $OLEDestinatioInicioY, $OLEDestinatioFinX, $OLEDestinatioFinY)
    Sleep(1000)

    ; Enlazar
    MouseMove(410, 315, 0)
    MouseClick("left")
    Send("{F2}")
    ClipPut("Ubicacion")
    Send("^v")
    Sleep(1000)
    Local $ConectInicioX = 410
    Local $ConectInicioY = 339
    Local $ConectFinX = 618
    Local $ConectFinY = 306
    ARRASTRARSOLTAR($ConectInicioX, $ConectInicioY, $ConectFinX, $ConectFinY)
    Sleep(1000)
EndFunc
Ubicacion()

; PARTE TIPDOC
Func TipDoc()
	BuscarExcelSource()
    Local $ExcelSourceFinX = 340 
    Local $ExcelSourceFinY = 343
    ARRASTRARSOLTAR($ExcelSourceInicioX, $ExcelSourceInicioY, $ExcelSourceFinX, $ExcelSourceFinY)
    Sleep(1000)

    BuscarOLEDBDestination()
    Local $OLEDestinatioFinX = 559 
    Local $OLEDestinatioFinY = 343
    ARRASTRARSOLTAR($OLEDestinatioInicioX, $OLEDestinatioInicioY, $OLEDestinatioFinX, $OLEDestinatioFinY)
    Sleep(1000)

    ; Enlazar
	MouseMove(421, 372, 0)
    MouseClick("left")
    Send("{F2}")
    ClipPut("Tip_Doc")
    Send("^v")
    Sleep(1000)
    Local $ConectInicioX = 406
    Local $ConectInicioY = 394
    Local $ConectFinX = 618
    Local $ConectFinY = 365
    ARRASTRARSOLTAR($ConectInicioX, $ConectInicioY, $ConectFinX, $ConectFinY)
    Sleep(1000)         
EndFunc
TipDoc()

Sleep(1000)
Func clickrandom()
    MouseMove(1423, 14, 0)
    MouseClick("left")
EndFunc
clickrandom()   

Sleep(3000)
Func clickrandom2()
    MouseMove(706, 840, 0)
    MouseClick("left")
EndFunc
clickrandom2()
Sleep(4000)

; CARGAR EXCEL
; FUNCIONES DE UBICACIONES
Func Entrar_InformacionPersonal()
    MouseMove(424, 202, 0)
    MouseClick("left")
EndFunc
Func Entrar_Estudios()
    MouseMove(424, 262, 0)
    MouseClick("left")
EndFunc
Func Entrar_Ubicacion()
    MouseMove(424, 321, 0)
    MouseClick("left")
EndFunc
Func Entrar_TipDoc()
    MouseMove(400, 364, 0)
    MouseClick("left")
EndFunc

Func Entrar_OLEDBInformacionPersonal()
    MouseMove(690, 205, 0)
    MouseClick("left")
EndFunc
Func Entrar_OLEDBEstudios()
    MouseMove(690, 259, 0)
    MouseClick("left")
EndFunc
Func Entrar_OLEDBUbicacionl()
    MouseMove(690, 321, 0)
    MouseClick("left")
EndFunc
Func Entrar_OLEDBTipDoc()
    MouseMove(690, 364, 0)
    MouseClick("left")
EndFunc

Func CargarExcel($value)
    Sleep(1000)
    Send("{ENTER}")
    Sleep(1000)
    Send("{TAB 2}")
    Sleep(1000)
    Send("{ENTER}")
    Sleep(1000)
    Send("{TAB 1}")
    Sleep(1000)
    Send("{ENTER}")
    Sleep(1000)					
    ; Descarga
    ClipPut("01.-Informacion_Personal.xlsx")
    Send("^v")
    Sleep(1000)
    Send("{ENTER}")
    Sleep(1000)
    Send("{TAB 3}")
    Sleep(1000)
    Send("{ENTER}")
    Sleep(1000)
    MouseMove(161, 300, 0)
    MouseClick("left")
    Sleep(1000)

    For $i = 1 To $value
        Send("{DOWN}")
        Sleep(100)
    Next
    Sleep(1000)
    Send("{ENTER}")
    Sleep(1000)
    Send("{TAB 2}")
    Sleep(1000)
    Send("{ENTER}")
    Sleep(2000)
EndFunc

Func Cargar_OLEDB($value,$value2)
	Sleep(1000)
	Send("{ENTER}")
	Sleep(1000)
	Send("{TAB 2}")
	Sleep(1000)
	Send("{ENTER}")
	Sleep(1000)
	Send("{TAB 2}")
	Sleep(1000)
	Send("{ENTER}")
	Sleep(1000)		
	Send("{ENTER}")
	Sleep(3000)					
	; SELECCIONAR OLE
    MouseMove(87, 16, 0)
    MouseClick("left")
	Sleep(1000)
	Send("{DOWN 6}")
	Sleep(1000)
	Send("{ENTER}")
	Sleep(1000)
	Send("{TAB 2}")
	Sleep(1000)
	ClipPut("KENNETH")
	Send("^v")
	Sleep(1000)
	Send("{TAB 4}") 
	Sleep(1000)
	ClipPut("BD_Repsol")
	Send("^v")
	Sleep(1000)
	Send("{TAB 2}") 
	Sleep(1000)
	Send("{ENTER}")
	Sleep(1000)
	Send("{ENTER}")
	Sleep(1000)
	; SELECCIONAR TABLA DE BD
	Send("{TAB 2}") 
	Sleep(1000)
	Send("{DOWN 2}")
	Sleep(2000)

    For $i = 1 To $value
        Send("{DOWN}")
        Sleep(100)
    Next
	Sleep(1000)
    For $i = 1 To $value2
        Send("{TAB}")
        Sleep(100)
    Next
	Sleep(2000)
	Send("{DOWN 1}")
	Sleep(2000)
	Send("{ENTER}")
    Sleep(1000)
EndFunc

; PARTE (INFORMACION PERSONAL)
Entrar_InformacionPersonal()
CargarExcel(0)
Entrar_OLEDBInformacionPersonal()
Cargar_OLEDB(1,6)

; PARTE (ESTUDIOS)
Sleep(2000)
Entrar_Estudios()
CargarExcel(2)
Entrar_OLEDBEstudios()
Cargar_OLEDB(2,12)

; PARTE (UBICACIÓN)
Sleep(2000)
Entrar_Ubicacion()      
CargarExcel(3)
Entrar_OLEDBUbicacionl()
Cargar_OLEDB(3,12)

; PARTE (TIPDOC)
Sleep(2000)
Entrar_TipDoc()
CargarExcel(4)
Sleep(2000)
Entrar_OLEDBTipDoc()
Cargar_OLEDB(4,12)

; CARGAR ELT
Func cargarEtl()
    Sleep(2000)
    MouseMove(492, 46, 0)
    MouseClick("left")
EndFunc
cargarEtl()