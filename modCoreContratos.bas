Attribute VB_Name = "modCoreContratos"
' ===== MÓDULO PRINCIPAL =====
Option Explicit

' ===========================
' Configuración / Constantes (GENÉRICAS PARA PORTAFOLIO)
' ===========================
Private Const RUTA_EXCEL_2025 As String = "C:\Ruta_Ejemplo\DB_Validacion\contrato_2025.xlsm"
Private Const RUTA_EXCEL_2026 As String = "C:\Ruta_Ejemplo\DB_Validacion\contrato_2026.xlsx"
Private Const RUTA_EXCEL As String = RUTA_EXCEL_2025
Private Const FILA_INICIO As Long = 7

' Columnas (hoja "contrato")
Private Const COL_NUM_CONTRATO     As Long = 2
Private Const COL_ID_CONTRATISTA   As Long = 7
Private Const COL_NOM_CONTRATISTA  As Long = 8
Private Const COL_ID_INTERVENTOR   As Long = 9
Private Const COL_NOM_INTERVENTOR  As Long = 10
Private Const COL_OBJETO           As Long = 17
Private Const COL_TIPO_MOVIMIENTO  As Long = 14 ' N
Private Const COL_CDP              As Long = 38
Private Const COL_FECHA_CDP        As Long = 39
Private Const COL_RP               As Long = 41
Private Const COL_FECHA_RP         As Long = 42
Private Const COL_FORMA_PAGO       As Long = 46 ' AT
Private Const COL_PLAZO            As Long = 47 ' AU

Public mostrarFormaPago As Boolean
Public mostrarPlazo As Boolean
Public filaExcel As Long
Public hojaExcel As Object

' Flag global: indica si el contrato actual TIENE adición real
Public EsAdicionContratoActual As Boolean
Public HuboCambiosDatosPago As Boolean

' ===========================
' Historial (bloc de notas)
' ===========================
Private Const LOG_FILENAME As String = "historial_firmas.txt"
Private Const SUPV_NO_IDENT As String = "N/N"   ' Cuando no se puede identificar
Public gUltimoInterventor As String            ' Último interventor leído
Public gAnioTrabajo As Long      ' Año del contrato cargado (2025/2026)
Public gAnioBusquedaCedula As Long  ' Año elegido en búsqueda por cédula

' ==============
' Lanzadores UI
' ==============
Public Sub UI_MostrarBusqueda()
    frmBuscarContrato.Show vbModeless
End Sub

' =======================
' Utilidades de limpieza
' =======================
Public Function LimpiarTexto(ByVal texto As String) As String: LimpiarTexto = Trim$(texto): End Function

Public Function LimpiarCedula(ByVal cedula As String) As String
    cedula = Replace(cedula, ".", "")
    cedula = Replace(cedula, " ", "")
    LimpiarCedula = cedula
End Function

' ===========================
' Año / Rutas (2025 / 2026)
' ===========================
Public Function ExisteArchivo(ByVal ruta As String) As Boolean
    ExisteArchivo = (Len(Dir$(ruta, vbNormal)) > 0)
End Function

Public Function RutaExcelPorAnio(ByVal anio As Long) As String
    If anio = 2026 Then
        RutaExcelPorAnio = RUTA_EXCEL_2026
    Else
        RutaExcelPorAnio = RUTA_EXCEL_2025
    End If
End Function

Public Function PreguntarAnioBusqueda() As Long
    Dim resp As String
    resp = InputBox( _
        "¿En qué año desea buscar el contrato?" & vbCrLf & vbCrLf & _
        "Escriba:" & vbCrLf & _
        "  1    2025" & vbCrLf & _
        "  2    2026", _
        "Seleccionar año de búsqueda")

    If resp = "" Then
        PreguntarAnioBusqueda = 0
        Exit Function
    End If

    Select Case Trim$(resp)
        Case "1": PreguntarAnioBusqueda = 2025
        Case "2": PreguntarAnioBusqueda = 2026
        Case Else
            MsgBox "Opción inválida. Operación cancelada.", vbExclamation
            PreguntarAnioBusqueda = 0
    End Select
End Function

Public Function DetectarAnioDesdeContrato(ByVal numeroContrato As String) As Long
    Dim s As String
    Dim partes() As String
    Dim i As Long

    s = Trim$(numeroContrato)
    If Len(s) = 0 Then DetectarAnioDesdeContrato = 0: Exit Function

    s = Replace(s, "-", "|")
    s = Replace(s, "/", "|")
    s = Replace(s, "_", "|")
    s = Replace(s, " ", "|")

    partes = Split(s, "|")

    For i = LBound(partes) To UBound(partes)
        Select Case partes(i)
            Case "2025": DetectarAnioDesdeContrato = 2025: Exit Function
            Case "2026": DetectarAnioDesdeContrato = 2026: Exit Function
        End Select
    Next i
    DetectarAnioDesdeContrato = 0
End Function

Public Function CedulaDesdeCelda(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Or IsEmpty(v) Then
        CedulaDesdeCelda = ""
        Exit Function
    End If

    If IsNumeric(v) Then
        s = Format$(CDbl(v), "0")
    Else
        s = CStr(v)
    End If
    CedulaDesdeCelda = LimpiarCedula(s)
End Function

' =======================
' Resto de funciones (Lógica del motor sin cambios)
' =======================
Private Function QuitarAcentos(ByVal s As String) As String
    s = Replace(s, "á", "a"): s = Replace(s, "Á", "A")
    s = Replace(s, "é", "e"): s = Replace(s, "É", "E")
    s = Replace(s, "í", "i"): s = Replace(s, "Í", "I")
    s = Replace(s, "ó", "o"): s = Replace(s, "Ó", "O")
    s = Replace(s, "ú", "u"): s = Replace(s, "Ú", "U")
    s = Replace(s, "ñ", "n"): s = Replace(s, "Ñ", "N")
    QuitarAcentos = s
End Function

Private Function NormalizarEspacios(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    NormalizarEspacios = t
End Function

' =============================================================
' Motor: fecha explícita "hasta el/al/a/el … de <mes> de|del <año>"
' =============================================================
Private Function TryParseHastaFecha(ByVal plazoTexto As String, ByRef fechaOut As Date) As Boolean
    On Error GoTo fallo
    Dim re As Object, m As Object
    Dim diaNum As Integer, mesTxt As String, anio As Integer, diaTxt As String
    Dim texto As String, mesTxtNorm As String

    texto = plazoTexto
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = _
        "(?:hasta\s+|al\s+|a\s+)?" & _
        "(?:el\s+)?" & _
        "(?:d[ií]a\s+)?" & _
        "(?:(\d{1,2})|\((\d{1,2})\)|([a-záéíóú\s]+?))\s+" & _
        "de\s+([a-záéíóú\.]+)\s+(?:de|del)\s+(\d{4})"

    If re.Test(texto) Then
        Set m = re.Execute(texto)(0)

        If m.SubMatches(0) <> "" Then
            diaNum = CInt(m.SubMatches(0))
        ElseIf m.SubMatches(1) <> "" Then
            diaNum = CInt(m.SubMatches(1))
        Else
            diaTxt = NormalizarEspacios(QuitarAcentos(LCase$(m.SubMatches(2))))
            diaNum = DiaTextoANumero(diaTxt)
        End If
        If diaNum <= 0 Or diaNum > 31 Then GoTo fallo

        mesTxt = m.SubMatches(3)
        mesTxtNorm = QuitarAcentos(LCase$(Trim$(mesTxt)))
        anio = CInt(m.SubMatches(4))

        fechaOut = DateSerial(anio, MesANumero(mesTxtNorm), diaNum)
        TryParseHastaFecha = True
        Exit Function
    End If
fallo:
    TryParseHastaFecha = False
End Function

' ================================
' Persona natural / persona jurídica
' ================================
Public Function EsPersonaJuridica(nombre As String) As Boolean
    Dim palabrasClave As Variant, i As Long
    palabrasClave = Array("S.A.S", "S.A", "FUNDACIÓN", "CORPORACIÓN", "ORGANIZACIÓN", _
                          "ASOCIACIÓN", "COOPERATIVA", "EMPRESA", "CONSORCIO", "UNIÓN TEMPORAL", _
                          "UNION TEMPORAL", "GRUPO", "SOCIEDAD", "SOCIAL", "INSTITUCIÓN", _
                          "COMPAÑÍA", "COMPAÑIA", "ENTIDAD", "AGENCIA", "ALIANZA", "ONG", _
                          "FIRMA", "CENTRO", "CONSTRUCTORA", "CLÍNICA", "HOSPITAL", "ESE", _
                          "IPS", "LTDA", "E.U", "E.S.P", "C.T.A")
    nombre = UCase$(nombre)
    For i = LBound(palabrasClave) To UBound(palabrasClave)
        If InStr(1, nombre, palabrasClave(i), vbTextCompare) > 0 Then EsPersonaJuridica = True: Exit Function
    Next
    EsPersonaJuridica = False
End Function

' ==================
' Fechas de vigencia
' ==================
Public Function CalcularFechaFinContrato(plazoTexto As String, fechaRP As Date, nombreContratista As String) As String
    Dim fechaHasta As Date
    Dim meses As Long, patronMeses As Object, mMeses As Object
    Dim fechaDestino As Date

    ' 1) Fecha explícita "hasta el ..."
    If TryParseHastaFecha(plazoTexto, fechaHasta) Then
        CalcularFechaFinContrato = Format$(fechaHasta, "dd/mm/yyyy")
        Exit Function
    End If

    ' 2) "por/de ... meses"
    Set patronMeses = CreateObject("VBScript.RegExp")
    patronMeses.IgnoreCase = True: patronMeses.Global = False
    patronMeses.Pattern = "(?:por|de)?\s*(?:(uno|dos|tres|cuatro|cinco|seis|siete|ocho|nueve|diez|once|doce)|\(?(\d{1,2})\)?)\s*mes(?:es)?"

    If patronMeses.Test(plazoTexto) Then
        Set mMeses = patronMeses.Execute(plazoTexto)(0)
        If mMeses.SubMatches(0) <> "" Then
            meses = ConvertirTextoANumero(mMeses.SubMatches(0))
        Else
            meses = CLng(mMeses.SubMatches(1))
        End If

        fechaDestino = DateAdd("m", meses, fechaRP)

        If EsPersonaJuridica(nombreContratista) Then
            CalcularFechaFinContrato = Format$(fechaDestino, "dd/mm/yyyy")
        Else
            If Day(fechaDestino) = Day(fechaRP) Then
                CalcularFechaFinContrato = Format$(DateAdd("d", -1, fechaDestino), "dd/mm/yyyy")
            Else
                CalcularFechaFinContrato = Format$(fechaDestino, "dd/mm/yyyy")
            End If
        End If
        Exit Function
    End If

    CalcularFechaFinContrato = "Plazo no reconocido"
End Function

Private Function ConvertirTextoANumero(texto As String) As Long
    Select Case LCase$(texto)
        Case "uno": ConvertirTextoANumero = 1
        Case "dos": ConvertirTextoANumero = 2
        Case "tres": ConvertirTextoANumero = 3
        Case "cuatro": ConvertirTextoANumero = 4
        Case "cinco": ConvertirTextoANumero = 5
        Case "seis": ConvertirTextoANumero = 6
        Case "siete": ConvertirTextoANumero = 7
        Case "ocho": ConvertirTextoANumero = 8
        Case "nueve": ConvertirTextoANumero = 9
        Case "diez": ConvertirTextoANumero = 10
        Case "once": ConvertirTextoANumero = 11
        Case "doce": ConvertirTextoANumero = 12
        Case Else: ConvertirTextoANumero = 0
    End Select
End Function

Private Function MesANumero(mesTexto As String) As Integer
    Dim s As String
    s = QuitarAcentos(LCase$(Trim$(mesTexto)))
    If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)

    Select Case s
        Case "enero", "ene": MesANumero = 1
        Case "febrero", "feb": MesANumero = 2
        Case "marzo", "mar": MesANumero = 3
        Case "abril", "abr": MesANumero = 4
        Case "mayo", "may": MesANumero = 5
        Case "junio", "jun": MesANumero = 6
        Case "julio", "jul": MesANumero = 7
        Case "agosto", "ago": MesANumero = 8
        Case "septiembre", "setiembre", "sept", "set", "sep": MesANumero = 9
        Case "octubre", "oct": MesANumero = 10
        Case "noviembre", "nov": MesANumero = 11
        Case "diciembre", "dic": MesANumero = 12
        Case Else: MesANumero = 1
    End Select
End Function

' ======================
' Búsquedas en la hoja
' ======================
Private Function UltimaFila(ws As Object, Optional ByVal col As Long = COL_NUM_CONTRATO, _
                            Optional ByVal filaInicio As Long = FILA_INICIO) As Long
    On Error Resume Next
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, col).End(-4162).row ' xlUp
    If lr < filaInicio Then lr = filaInicio
    UltimaFila = lr
End Function

' ===== Detección estricta de adición/prórroga =====
Private Function EsFilaAdicion(ByVal textoColN As String) As Boolean
    Dim s As String, re As Object
    s = UCase$(QuitarAcentos(NormalizarEspacios(CStr(textoColN))))
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True: re.Global = False
    re.Pattern = "(?:^|\b)(ADICION|PRORROGA)(?:\b|$)" ' evita ADICIONAL/ADICIONALMENTE
    EsFilaAdicion = re.Test(s)
End Function

Private Function TieneDatosAdicion(ws As Object, ByVal f As Long) As Boolean
    TieneDatosAdicion = (Trim$(ws.Cells(f, COL_CDP).value) <> "" Or _
                         Trim$(ws.Cells(f, COL_FECHA_CDP).value) <> "" Or _
                         Trim$(ws.Cells(f, COL_RP).value) <> "" Or _
                         Trim$(ws.Cells(f, COL_FECHA_RP).value) <> "" Or _
                         Trim$(ws.Cells(f, COL_FORMA_PAGO).value) <> "" Or _
                         Trim$(ws.Cells(f, COL_PLAZO).value) <> "")
End Function

Private Function FilaAdicionVigente(ws As Object, ByVal numeroContrato As String) As Long
    Dim lr As Long, f As Long, mejorFila As Long
    Dim mejorFecha As Date, fechaRP As Variant, esAdic As Boolean

    lr = UltimaFila(ws, COL_NUM_CONTRATO, FILA_INICIO)
    mejorFila = 0: mejorFecha = #1/1/1900#

    For f = FILA_INICIO To lr
        If CStr(ws.Cells(f, COL_NUM_CONTRATO).value) = numeroContrato Then
            esAdic = EsFilaAdicion(ws.Cells(f, COL_TIPO_MOVIMIENTO).value)
            If esAdic And TieneDatosAdicion(ws, f) Then
                fechaRP = ws.Cells(f, COL_FECHA_RP).value
                If IsDate(fechaRP) Then
                    If CDate(fechaRP) >= mejorFecha Then
                        mejorFecha = CDate(fechaRP)
                        mejorFila = f
                    End If
                ElseIf mejorFila = 0 Or Not IsDate(ws.Cells(mejorFila, COL_FECHA_RP).value) Then
                    mejorFila = f
                End If
            End If
        End If
    Next

    FilaAdicionVigente = mejorFila
End Function

' ===========================
' Fusión de "Forma de pago"
' ===========================
Private Function FusionarFormaPago(ByVal fpBase As String, ByVal fpAdic As String) As String
    Dim b As String, a As String
    b = Trim$(fpBase): a = Trim$(fpAdic)
    If a = "" Then
        FusionarFormaPago = b
    ElseIf FormaPagoEsCompleta(a) Then
        FusionarFormaPago = a
    ElseIf b = "" Then
        FusionarFormaPago = a
    Else
        FusionarFormaPago = b & vbCrLf & a
    End If
End Function

Private Function FormaPagoEsCompleta(ByVal texto As String) As Boolean
    Dim t As String: t = LCase$(texto)
    FormaPagoEsCompleta = (InStr(t, "forma de pago") > 0) _
                       Or (InStr(t, "cancelará el valor del contrato") > 0) _
                       Or (InStr(t, "pagos mensuales") > 0)
End Function

' ====================================================
' Fecha fin vigente considerando una adición/prórroga
' ====================================================
Private Function CalcularFechaFinVigente(ByVal plazoBase As String, ByVal plazoAdic As String, _
                                         ByVal fechaRPBaseTexto As String, ByVal nombreContratista As String) As String
    Dim finBaseStr As String, finBase As Date
    Dim reMeses As Object, m As Object
    Dim meses As Long
    Dim fechaHastaAdic As Date

    If IsDate(fechaRPBaseTexto) Then
        finBaseStr = CalcularFechaFinContrato(plazoBase, CDate(fechaRPBaseTexto), nombreContratista)
    Else
        finBaseStr = "Plazo no reconocido"
    End If

    If Trim$(plazoAdic) = "" Then
        CalcularFechaFinVigente = finBaseStr: Exit Function
    End If

    If TryParseHastaFecha(plazoAdic, fechaHastaAdic) Then
        CalcularFechaFinVigente = Format$(fechaHastaAdic, "dd/mm/yyyy")
        Exit Function
    End If

    Set reMeses = CreateObject("VBScript.RegExp")
    reMeses.IgnoreCase = True: reMeses.Global = False
    reMeses.Pattern = "(?:por|de)?\s*(?:(uno|dos|tres|cuatro|cinco|seis|siete|ocho|nueve|diez|once|doce)|\(?(\d{1,2})\)?)\s*meses?"

    If reMeses.Test(plazoAdic) Then
        Set m = reMeses.Execute(plazoAdic)(0)
        If m.SubMatches(0) <> "" Then meses = ConvertirTextoANumero(m.SubMatches(0)) Else meses = CLng(m.SubMatches(1))

        If IsDate(finBaseStr) Then
            finBase = CDate(finBaseStr)
            CalcularFechaFinVigente = Format$(DateAdd("m", meses, finBase), "dd/mm/yyyy")
        ElseIf IsDate(fechaRPBaseTexto) Then
            CalcularFechaFinVigente = Format$(DateAdd("m", meses, CDate(fechaRPBaseTexto)), "dd/mm/yyyy")
        Else
            CalcularFechaFinVigente = finBaseStr
        End If
        Exit Function
    End If

    CalcularFechaFinVigente = finBaseStr
End Function

' ==========================================================
' Construye el texto a mostrar en "Fecha Fin Contrato"
' ==========================================================
Public Function ConstruirTextoFechasFin( _
    ByVal plazoBase As String, _
    ByVal plazoAdic As String, _
    ByVal fechaRPBaseTexto As String, _
    ByVal nombreContratista As String _
) As String

    Dim finBase As String
    Dim finVigente As String
    Dim hayAdicion As Boolean

    hayAdicion = (Trim$(plazoAdic) <> "")

    If IsDate(fechaRPBaseTexto) Then
        finBase = CalcularFechaFinContrato(plazoBase, CDate(fechaRPBaseTexto), nombreContratista)
    Else
        finBase = "Plazo no reconocido"
    End If

    finVigente = CalcularFechaFinVigente(plazoBase, plazoAdic, fechaRPBaseTexto, nombreContratista)

    If Not hayAdicion Or finVigente = finBase Then
        ConstruirTextoFechasFin = finBase
    Else
        ConstruirTextoFechasFin = "BASE: " & finBase & vbCrLf & "ADICIÓN: " & finVigente
    End If
End Function

Private Function DiaTextoANumero(t As String) As Integer
    Dim s As String
    s = LCase$(Trim$(t))
    s = Replace(s, "á", "a"): s = Replace(s, "é", "e")
    s = Replace(s, "í", "i"): s = Replace(s, "ó", "o")
    s = Replace(s, "ú", "u")
    Select Case s
        Case "uno", "primero": DiaTextoANumero = 1
        Case "dos": DiaTextoANumero = 2
        Case "tres": DiaTextoANumero = 3
        Case "cuatro": DiaTextoANumero = 4
        Case "cinco": DiaTextoANumero = 5
        Case "seis": DiaTextoANumero = 6
        Case "siete": DiaTextoANumero = 7
        Case "ocho": DiaTextoANumero = 8
        Case "nueve": DiaTextoANumero = 9
        Case "diez": DiaTextoANumero = 10
        Case "once": DiaTextoANumero = 11
        Case "doce": DiaTextoANumero = 12
        Case "trece": DiaTextoANumero = 13
        Case "catorce": DiaTextoANumero = 14
        Case "quince": DiaTextoANumero = 15
        Case "dieciseis", "dieciséis": DiaTextoANumero = 16
        Case "diecisiete": DiaTextoANumero = 17
        Case "dieciocho": DiaTextoANumero = 18
        Case "diecinueve": DiaTextoANumero = 19
        Case "veinte": DiaTextoANumero = 20
        Case "veintiuno", "veintiun": DiaTextoANumero = 21
        Case "veintidos", "veintidós": DiaTextoANumero = 22
        Case "veintitres", "veintitrés": DiaTextoANumero = 23
        Case "veinticuatro": DiaTextoANumero = 24
        Case "veinticinco": DiaTextoANumero = 25
        Case "veintiseis", "veintiséis": DiaTextoANumero = 26
        Case "veintisiete": DiaTextoANumero = 27
        Case "veintiocho": DiaTextoANumero = 28
        Case "veintinueve": DiaTextoANumero = 29
        Case "treinta": DiaTextoANumero = 30
        Case "treinta y uno": DiaTextoANumero = 31
        Case Else: DiaTextoANumero = 0
    End Select
End Function

' ===========================
' Buscar por Nº de contrato
' ===========================
Public Function BuscarContratoPorNumero(ByVal numeroContrato As String) As Boolean
    Dim excelApp As Object, wb As Object, ws As Object
    Dim fila As Long, filaAdic As Long, lr As Long, encontrado As Boolean
    Dim respuesta As VbMsgBoxResult
    Dim huboCambios As Boolean: huboCambios = False

    Dim numeroContratoTexto As String, fechaFinContrato As String
    Dim identificacionContratista As String, nombreContratista As String
    Dim identificacionInterventor As String, nombreInterventor As String
    Dim objetoContrato As String

    Dim cdpBase As String, fCdpBase As String
    Dim rpBase As String, fRpBase As String
    Dim formaPagoBase As String, plazoBase As String

    Dim cdpAdic As String, fCdpAdic As String
    Dim rpAdic As String, fRpAdic As String
    Dim formaPagoAdic As String, plazoAdic As String

    Dim cdpMostrar As String, rpMostrar As String
    Dim fCdpMostrar As String, fRpMostrar As String
    Dim plazoMostrar As String, formaPagoFusion As String

    On Error GoTo CerrarExcel
    numeroContrato = LimpiarTexto(numeroContrato)

    Dim anioN As Long
    Dim rutaLibro As String

    anioN = DetectarAnioDesdeContrato(numeroContrato)
    If anioN <> 2025 And anioN <> 2026 Then anioN = 2025

    rutaLibro = RutaExcelPorAnio(anioN)
    If Not ExisteArchivo(rutaLibro) Then
        MsgBox "No se encontró el archivo Excel del año " & anioN & ":" & vbCrLf & rutaLibro, vbCritical
        BuscarContratoPorNumero = False
        Exit Function
    End If


    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    Set wb = excelApp.Workbooks.Open(rutaLibro)
    Set ws = wb.Sheets("contrato")

    encontrado = False
    lr = UltimaFila(ws, COL_NUM_CONTRATO, FILA_INICIO)

    For fila = FILA_INICIO To lr
        If CStr(ws.Cells(fila, COL_NUM_CONTRATO).value) = numeroContrato Then
            encontrado = True

            gAnioTrabajo = anioN
            numeroContratoTexto = ws.Cells(fila, COL_NUM_CONTRATO).value
            identificacionContratista = ws.Cells(fila, COL_ID_CONTRATISTA).value
            nombreContratista = ws.Cells(fila, COL_NOM_CONTRATISTA).value
            identificacionInterventor = ws.Cells(fila, COL_ID_INTERVENTOR).value
            nombreInterventor = ws.Cells(fila, COL_NOM_INTERVENTOR).value
            objetoContrato = ws.Cells(fila, COL_OBJETO).value

            gUltimoInterventor = nombreInterventor

            cdpBase = Trim$(ws.Cells(fila, COL_CDP).value)
            fCdpBase = Trim$(ws.Cells(fila, COL_FECHA_CDP).value)
            rpBase = Trim$(ws.Cells(fila, COL_RP).value)
            fRpBase = Trim$(ws.Cells(fila, COL_FECHA_RP).value)
            formaPagoBase = Trim$(ws.Cells(fila, COL_FORMA_PAGO).value)
            plazoBase = Trim$(ws.Cells(fila, COL_PLAZO).value)

            ' --- Completar BASE si falta ---
            If formaPagoBase = "" Or plazoBase = "" Then
                respuesta = MsgBox("Faltan los siguientes datos:" & vbCrLf & _
                    IIf(formaPagoBase = "", "- FORMA DE PAGO" & vbCrLf, "") & _
                    IIf(plazoBase = "", "- PLAZO" & vbCrLf, "") & vbCrLf & _
                    "¿Desea ingresarlos ahora?", vbYesNo + vbQuestion, "Datos incompletos (Base)")
                If respuesta = vbYes Then
                    mostrarFormaPago = (formaPagoBase = "")
                    mostrarPlazo = (plazoBase = "")
                    filaExcel = fila: Set hojaExcel = ws
                    With frmDatosPago
                        .EsAdicion = False
                        .numeroContratoActual = numeroContratoTexto
                        .cancelado = False
                        .PrepararFormulario
                        .Show
                    End With
                    If frmDatosPago.cancelado Then
                        BuscarContratoPorNumero = False
                        GoTo CerrarExcel
                    End If
                    ' Se guardó: refrescar y marcar cambios
                    formaPagoBase = Trim$(ws.Cells(fila, COL_FORMA_PAGO).value)
                    plazoBase = Trim$(ws.Cells(fila, COL_PLAZO).value)
                    huboCambios = True
                End If
            End If

            ' --- ADICIÓN vigente (si hay) ---
            filaAdic = FilaAdicionVigente(ws, numeroContratoTexto)
            EsAdicionContratoActual = (filaAdic > 0)

            If filaAdic > 0 Then
                cdpAdic = Trim$(ws.Cells(filaAdic, COL_CDP).value)
                fCdpAdic = Trim$(ws.Cells(filaAdic, COL_FECHA_CDP).value)
                rpAdic = Trim$(ws.Cells(filaAdic, COL_RP).value)
                fRpAdic = Trim$(ws.Cells(filaAdic, COL_FECHA_RP).value)
                formaPagoAdic = Trim$(ws.Cells(filaAdic, COL_FORMA_PAGO).value)
                plazoAdic = Trim$(ws.Cells(filaAdic, COL_PLAZO).value)

                If (formaPagoAdic = "" Or plazoAdic = "") Then
                    respuesta = MsgBox("Faltan los siguientes datos en la ADICIÓN:" & vbCrLf & _
                        IIf(formaPagoAdic = "", "- FORMA DE PAGO (Adición)" & vbCrLf, "") & _
                        IIf(plazoAdic = "", "- PLAZO (Adición)" & vbCrLf, "") & vbCrLf & _
                        "¿Desea ingresarlos ahora?", vbYesNo + vbQuestion, "Datos incompletos (Adición)")
                    If respuesta = vbYes Then
                        mostrarFormaPago = (formaPagoAdic = "")
                        mostrarPlazo = (plazoAdic = "")
                        filaExcel = filaAdic: Set hojaExcel = ws
                        With frmDatosPago
                            .EsAdicion = True
                            .numeroContratoActual = numeroContratoTexto
                            .cancelado = False
                            .PrepararFormulario
                            .Show
                        End With
                        If frmDatosPago.cancelado Then
                            BuscarContratoPorNumero = False
                            GoTo CerrarExcel
                        End If
                        formaPagoAdic = Trim$(ws.Cells(filaAdic, COL_FORMA_PAGO).value)
                        plazoAdic = Trim$(ws.Cells(filaAdic, COL_PLAZO).value)
                        huboCambios = True
                    End If
                End If
            Else
                cdpAdic = "": fCdpAdic = ""
                rpAdic = "": fRpAdic = ""
                formaPagoAdic = "": plazoAdic = ""
            End If

            cdpMostrar = cdpBase & IIf(cdpAdic <> "", " - " & cdpAdic, "")
            rpMostrar = rpBase & IIf(rpAdic <> "", " - " & rpAdic, "")
            fCdpMostrar = fCdpBase & IIf(fCdpAdic <> "", " - " & fCdpAdic, "")
            fRpMostrar = fRpBase & IIf(fRpAdic <> "", " - " & fRpAdic, "")
            plazoMostrar = "BASE: " & plazoBase & IIf(plazoAdic <> "", vbCrLf & vbCrLf & "ADICIÓN: " & plazoAdic, "")

            formaPagoFusion = FusionarFormaPago(formaPagoBase, formaPagoAdic)

            fechaFinContrato = ConstruirTextoFechasFin(plazoBase, plazoAdic, fRpBase, nombreContratista)

            With frmContrato
                .Label1.Caption = "Número de Contrato: " & numeroContratoTexto
                .Text_NumContrato.Text = numeroContratoTexto
                .Text_FechaFin.Text = fechaFinContrato

                .Text_IDContratista.Text = identificacionContratista
                .Text_NombreContratista.Text = nombreContratista
                .Text_IDInterventor.Text = identificacionInterventor
                .Text_NombreInterventor.Text = nombreInterventor
                .Text_Objeto.Text = objetoContrato

                .Text_CDP.Text = cdpMostrar
                .Text_FechaCDP.Text = fCdpMostrar
                .Text_RP.Text = rpMostrar
                .Text_FechaRP.Text = fRpMostrar
                .Text_FormaPago.Text = formaPagoFusion
                .Text_Plazo.Text = plazoMostrar
                .Show vbModeless
            End With

            Dim doc As Document
            Set doc = Application.ActiveDocument

            BuscarYResaltarCoincidencia doc, numeroContratoTexto, "Número de Contrato: "

            Dim finBase__ As String, finVigente__ As String, pSep As Long
            pSep = InStr(1, fechaFinContrato, vbCrLf, vbBinaryCompare)
            If pSep > 0 Then
                finBase__ = Trim$(Replace(Split(fechaFinContrato, vbCrLf)(0), "BASE:", ""))
                finVigente__ = Trim$(Replace(Split(fechaFinContrato, vbCrLf)(1), "ADICIÓN:", ""))
            Else
                finBase__ = fechaFinContrato
                finVigente__ = ""
            End If
            If Len(finBase__) > 0 Then BuscarYResaltarCoincidencia doc, finBase__, "Fecha Fin Contrato (Base): "
            If Len(finVigente__) > 0 And finVigente__ <> finBase__ Then _
                BuscarYResaltarCoincidencia doc, finVigente__, "Fecha Fin Contrato (Adición): "

            BuscarYResaltarCoincidencia doc, identificacionContratista, "Identificación Contratista: "
            BuscarYResaltarCoincidencia doc, nombreContratista, "Nombre Contratista: "
            BuscarYResaltarCoincidencia doc, identificacionInterventor, "Identificación Interventor: "
            BuscarYResaltarCoincidencia doc, nombreInterventor, "Nombre Interventor: "
            BuscarYResaltarCoincidencia doc, objetoContrato, "Objeto del Contrato: "
            BuscarYResaltarCoincidencia doc, cdpBase, "Código CDP: "
            If cdpAdic <> "" Then BuscarYResaltarCoincidencia doc, cdpAdic, "Código CDP (Adición): "
            BuscarYResaltarCoincidencia doc, fCdpBase, "Fecha Registro CDP: "
            If fCdpAdic <> "" Then BuscarYResaltarCoincidencia doc, fCdpAdic, "Fecha Registro CDP (Adición): "
            BuscarYResaltarCoincidencia doc, rpBase, "Código RP: "
            If rpAdic <> "" Then BuscarYResaltarCoincidencia doc, rpAdic, "Código RP (Adición): "
            BuscarYResaltarCoincidencia doc, fRpBase, "Fecha Registro RP: "
            If fRpAdic <> "" Then BuscarYResaltarCoincidencia doc, fRpAdic, "Fecha Registro RP (Adición): "
            If plazoBase <> "" Then BuscarYResaltarCoincidencia doc, plazoBase, "Plazo (Base): "
            If plazoAdic <> "" Then BuscarYResaltarCoincidencia doc, plazoAdic, "Plazo (Adición): "

            BuscarContratoPorNumero = True
            GoTo CerrarExcel
        End If
    Next

    UI_MostrarBusqueda
    BuscarContratoPorNumero = False

CerrarExcel:
    On Error Resume Next
    If Not excelApp Is Nothing Then excelApp.DisplayAlerts = False
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=huboCambios
    End If
    If Not excelApp Is Nothing Then
        excelApp.DisplayAlerts = True
        excelApp.Quit
    End If
    Set ws = Nothing: Set wb = Nothing: Set excelApp = Nothing
    On Error GoTo 0
End Function

' =========================
' Buscar por cédula (ID)
' =========================
Public Function BuscarContratoPorCedula(ByVal cedula As String) As Boolean
    Dim excelApp As Object, wb As Object, ws As Object
    Dim fila As Long, filaAdic As Long, lr As Long, encontrado As Boolean
    Dim respuesta As VbMsgBoxResult
    Dim huboCambios As Boolean: huboCambios = False

    Dim numeroContratoTexto As String, fechaFinContrato As String
    Dim identificacionContratista As String, nombreContratista As String
    Dim identificacionInterventor As String, nombreInterventor As String
    Dim objetoContrato As String

    Dim cdpBase As String, fCdpBase As String
    Dim rpBase As String, fRpBase As String
    Dim formaPagoBase As String, plazoBase As String

    Dim cdpAdic As String, fCdpAdic As String
    Dim rpAdic As String, fRpAdic As String
    Dim formaPagoAdic As String, plazoAdic As String

    Dim cdpMostrar As String, rpMostrar As String
    Dim fCdpMostrar As String, fRpMostrar As String
    Dim plazoMostrar As String, formaPagoFusion As String

    On Error GoTo CerrarExcel
    cedula = LimpiarCedula(cedula)

    Dim anioElegido As Long
    Dim rutaLibro As String

    anioElegido = PreguntarAnioBusqueda()
    If anioElegido = 0 Then
        BuscarContratoPorCedula = False
        Exit Function
    End If

    gAnioBusquedaCedula = anioElegido
    rutaLibro = RutaExcelPorAnio(anioElegido)
    If Not ExisteArchivo(rutaLibro) Then
        MsgBox "No se encontró el archivo Excel del año " & anioElegido & ":" & vbCrLf & rutaLibro, vbCritical
        BuscarContratoPorCedula = False
        Exit Function
    End If


    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = False
    Set wb = excelApp.Workbooks.Open(rutaLibro)
    Set ws = wb.Sheets("contrato")

    encontrado = False
    lr = UltimaFila(ws, COL_ID_CONTRATISTA, FILA_INICIO)

    For fila = FILA_INICIO To lr
        If CedulaDesdeCelda(ws.Cells(fila, COL_ID_CONTRATISTA).value) = cedula Then
            encontrado = True

            gAnioTrabajo = anioElegido
            numeroContratoTexto = ws.Cells(fila, COL_NUM_CONTRATO).value
            identificacionContratista = ws.Cells(fila, COL_ID_CONTRATISTA).value
            nombreContratista = ws.Cells(fila, COL_NOM_CONTRATISTA).value
            identificacionInterventor = ws.Cells(fila, COL_ID_INTERVENTOR).value
            nombreInterventor = ws.Cells(fila, COL_NOM_INTERVENTOR).value
            objetoContrato = ws.Cells(fila, COL_OBJETO).value

            gUltimoInterventor = nombreInterventor

            cdpBase = Trim$(ws.Cells(fila, COL_CDP).value)
            fCdpBase = Trim$(ws.Cells(fila, COL_FECHA_CDP).value)
            rpBase = Trim$(ws.Cells(fila, COL_RP).value)
            fRpBase = Trim$(ws.Cells(fila, COL_FECHA_RP).value)
            formaPagoBase = Trim$(ws.Cells(fila, COL_FORMA_PAGO).value)
            plazoBase = Trim$(ws.Cells(fila, COL_PLAZO).value)

            ' --- Completar BASE si falta ---
            If formaPagoBase = "" Or plazoBase = "" Then
                respuesta = MsgBox("Faltan los siguientes datos:" & vbCrLf & _
                    IIf(formaPagoBase = "", "- FORMA DE PAGO" & vbCrLf, "") & _
                    IIf(plazoBase = "", "- PLAZO" & vbCrLf, "") & vbCrLf & _
                    "¿Desea ingresarlos ahora?", vbYesNo + vbQuestion, "Datos incompletos (Base)")
                If respuesta = vbYes Then
                    mostrarFormaPago = (formaPagoBase = "")
                    mostrarPlazo = (plazoBase = "")
                    filaExcel = fila: Set hojaExcel = ws
                    With frmDatosPago
                        .EsAdicion = False
                        .numeroContratoActual = numeroContratoTexto
                        .cancelado = False
                        .PrepararFormulario
                        .Show
                    End With
                    If frmDatosPago.cancelado Then
                        BuscarContratoPorCedula = False
                        GoTo CerrarExcel
                    End If
                    formaPagoBase = Trim$(ws.Cells(fila, COL_FORMA_PAGO).value)
                    plazoBase = Trim$(ws.Cells(fila, COL_PLAZO).value)
                    huboCambios = True
                End If
            End If

            ' --- ADICIÓN vigente (si hay) ---
            filaAdic = FilaAdicionVigente(ws, numeroContratoTexto)
            EsAdicionContratoActual = (filaAdic > 0)

            If filaAdic > 0 Then
                cdpAdic = Trim$(ws.Cells(filaAdic, COL_CDP).value)
                fCdpAdic = Trim$(ws.Cells(filaAdic, COL_FECHA_CDP).value)
                rpAdic = Trim$(ws.Cells(filaAdic, COL_RP).value)
                fRpAdic = Trim$(ws.Cells(filaAdic, COL_FECHA_RP).value)
                formaPagoAdic = Trim$(ws.Cells(filaAdic, COL_FORMA_PAGO).value)
                plazoAdic = Trim$(ws.Cells(filaAdic, COL_PLAZO).value)

                If (formaPagoAdic = "" Or plazoAdic = "") Then
                    respuesta = MsgBox("Faltan los siguientes datos en la ADICIÓN:" & vbCrLf & _
                        IIf(formaPagoAdic = "", "- FORMA DE PAGO (Adición)" & vbCrLf, "") & _
                        IIf(plazoAdic = "", "- PLAZO (Adición)" & vbCrLf, "") & vbCrLf & _
                        "¿Desea ingresarlos ahora?", vbYesNo + vbQuestion, "Datos incompletos (Adición)")
                    If respuesta = vbYes Then
                        mostrarFormaPago = (formaPagoAdic = "")
                        mostrarPlazo = (plazoAdic = "")
                        filaExcel = filaAdic: Set hojaExcel = ws
                        With frmDatosPago
                            .EsAdicion = True
                            .numeroContratoActual = numeroContratoTexto
                            .cancelado = False
                            .PrepararFormulario
                            .Show
                        End With
                        If frmDatosPago.cancelado Then
                            BuscarContratoPorCedula = False
                            GoTo CerrarExcel
                        End If
                        formaPagoAdic = Trim$(ws.Cells(filaAdic, COL_FORMA_PAGO).value)
                        plazoAdic = Trim$(ws.Cells(filaAdic, COL_PLAZO).value)
                        huboCambios = True
                    End If
                End If
            Else
                cdpAdic = "": fCdpAdic = ""
                rpAdic = "": fRpAdic = ""
                formaPagoAdic = "": plazoAdic = ""
            End If

            cdpMostrar = cdpBase & IIf(cdpAdic <> "", " - " & cdpAdic, "")
            rpMostrar = rpBase & IIf(rpAdic <> "", " - " & rpAdic, "")
            fCdpMostrar = fCdpBase & IIf(fCdpAdic <> "", " - " & fCdpAdic, "")
            fRpMostrar = fRpBase & IIf(fRpAdic <> "", " - " & fRpAdic, "")
            plazoMostrar = "BASE: " & plazoBase & IIf(plazoAdic <> "", vbCrLf & vbCrLf & "ADICIÓN: " & plazoAdic, "")

            formaPagoFusion = FusionarFormaPago(formaPagoBase, formaPagoAdic)

            fechaFinContrato = ConstruirTextoFechasFin(plazoBase, plazoAdic, fRpBase, nombreContratista)

            With frmContrato
                .Label1.Caption = "Número de Contrato: " & numeroContratoTexto
                .Text_NumContrato.Text = numeroContratoTexto
                .Text_FechaFin.Text = fechaFinContrato

                .Text_IDContratista.Text = identificacionContratista
                .Text_NombreContratista.Text = nombreContratista
                .Text_IDInterventor.Text = identificacionInterventor
                .Text_NombreInterventor.Text = nombreInterventor
                .Text_Objeto.Text = objetoContrato

                .Text_CDP.Text = cdpMostrar
                .Text_FechaCDP.Text = fCdpMostrar
                .Text_RP.Text = rpMostrar
                .Text_FechaRP.Text = fRpMostrar
                .Text_FormaPago.Text = formaPagoFusion
                .Text_Plazo.Text = plazoMostrar
                .Show vbModeless
            End With

            Dim doc As Document
            Set doc = Application.ActiveDocument

            BuscarYResaltarCoincidencia doc, numeroContratoTexto, "Número de Contrato: "

            Dim finBase__ As String, finVigente__ As String, pSep As Long
            pSep = InStr(1, fechaFinContrato, vbCrLf, vbBinaryCompare)
            If pSep > 0 Then
                finBase__ = Trim$(Replace(Split(fechaFinContrato, vbCrLf)(0), "BASE:", ""))
                finVigente__ = Trim$(Replace(Split(fechaFinContrato, vbCrLf)(1), "ADICIÓN:", ""))
            Else
                finBase__ = fechaFinContrato
                finVigente__ = ""
            End If
            If Len(finBase__) > 0 Then BuscarYResaltarCoincidencia doc, finBase__, "Fecha Fin Contrato (Base): "
            If Len(finVigente__) > 0 And finVigente__ <> finBase__ Then _
                BuscarYResaltarCoincidencia doc, finVigente__, "Fecha Fin Contrato (Adición): "

            BuscarYResaltarCoincidencia doc, identificacionContratista, "Identificación Contratista: "
            BuscarYResaltarCoincidencia doc, nombreContratista, "Nombre Contratista: "
            BuscarYResaltarCoincidencia doc, identificacionInterventor, "Identificación Interventor: "
            BuscarYResaltarCoincidencia doc, nombreInterventor, "Nombre Interventor: "
            BuscarYResaltarCoincidencia doc, objetoContrato, "Objeto del Contrato: "
            BuscarYResaltarCoincidencia doc, cdpBase, "Código CDP: "
            If cdpAdic <> "" Then BuscarYResaltarCoincidencia doc, cdpAdic, "Código CDP (Adición): "
            BuscarYResaltarCoincidencia doc, fCdpBase, "Fecha Registro CDP: "
            If fCdpAdic <> "" Then BuscarYResaltarCoincidencia doc, fCdpAdic, "Fecha Registro CDP (Adición): "
            BuscarYResaltarCoincidencia doc, rpBase, "Código RP: "
            If rpAdic <> "" Then BuscarYResaltarCoincidencia doc, rpAdic, "Código RP (Adición): "
            BuscarYResaltarCoincidencia doc, fRpBase, "Fecha Registro RP: "
            If fRpAdic <> "" Then BuscarYResaltarCoincidencia doc, fRpAdic, "Fecha Registro RP (Adición): "
            If plazoBase <> "" Then BuscarYResaltarCoincidencia doc, plazoBase, "Plazo (Base): "
            If plazoAdic <> "" Then BuscarYResaltarCoincidencia doc, plazoAdic, "Plazo (Adición): "

            Unload frmBuscarContrato
            BuscarContratoPorCedula = True
            GoTo CerrarExcel
        End If
    Next

    BuscarContratoPorCedula = False

CerrarExcel:
    On Error Resume Next
    If Not excelApp Is Nothing Then excelApp.DisplayAlerts = False
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=huboCambios
    End If
    If Not excelApp Is Nothing Then
        excelApp.DisplayAlerts = True
        excelApp.Quit
    End If
    Set ws = Nothing: Set wb = Nothing: Set excelApp = Nothing
    On Error GoTo 0
End Function

' =========================
' Resaltado en documento
' =========================
Private Sub BuscarYResaltarCoincidencia(doc As Document, ByVal valor As String, ByVal etiqueta As String)
    On Error Resume Next
    If Len(Trim$(valor)) > 0 Then
        With doc.Content.Find
            .Text = valor
            .Replacement.Text = valor & " *"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = True
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            .Execute Replace:=wdReplaceAll
        End With
    End If
    On Error GoTo 0
End Sub
