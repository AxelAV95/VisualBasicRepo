# **Laboratorio: Automatización de Outlook con Excel VBA**

## **Objetivo**:
El objetivo de este laboratorio es aprender a interactuar con Outlook desde Excel usando VBA (Visual Basic for Applications) para:
1. **Detectar conflictos en el calendario de Outlook**.
2. **Obtener los últimos correos de Outlook**.
3. **Obtener eventos del calendario de Outlook**.

**Requisitos previos**:
1. **Microsoft Excel** instalado.
2. **Microsoft Outlook** instalado y configurado.
3. Conocimientos básicos de Excel y VBA.

---

## **Paso 1: Preparar el entorno de trabajo**

1. **Abrir Excel**.
2. Crear un nuevo libro de trabajo en Excel.
3. **Habilitar la pestaña "Desarrollador"** si no la tienes activada:
   - Ve a "Archivo" → "Opciones" → "Personalizar cinta de opciones".
   - Activa la casilla **Desarrollador**.
4. **Abrir el Editor VBA**:
   - Haz clic en la pestaña **Desarrollador** y selecciona **Visual Basic**.
   
---

## **Paso 2: Insertar las macros en el Editor de VBA**

1. Dentro del **Editor de VBA**, selecciona **Insertar** → **Módulo**. Esto creará un nuevo módulo.
2. Copia y pega cada una de las macros proporcionadas en los módulos de VBA. **Una por una**:
   
### Macro 1: Detectar Conflictos en el Calendario de Outlook

```vba
Sub DetectarConflictosCalendario()
    Dim OutlookApp As Object
    Dim Namespace As Object
    Dim Calendar As Object
    Dim Items As Object
    Dim Appt As Object, Appt2 As Object
    Dim FechaInicio As Date, FechaFin As Date
    Dim UltimaFila As Integer
    Dim ws As Worksheet
    Dim ConflictosDetectados As Boolean

    ' Inicializar la variable de conflictos detectados
    ConflictosDetectados = False

    ' Verificar si la hoja "Conflictos" existe, si no, crearla
    On Error Resume Next
    Set ws = Sheets("Conflictos")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = "Conflictos"
    End If

    ' Iniciar Outlook
    On Error Resume Next
    Set OutlookApp = GetObject(, "Outlook.Application")
    If OutlookApp Is Nothing Then Set OutlookApp = CreateObject("Outlook.Application")
    On Error GoTo 0

    ' Obtener el calendario predeterminado
    Set Namespace = OutlookApp.GetNamespace("MAPI")
    Set Calendar = Namespace.GetDefaultFolder(9) ' 9 = olFolderCalendar
    Set Items = Calendar.Items

    ' Ordenar eventos por fecha de inicio
    Items.Sort "[Start]"
    Items.IncludeRecurrences = True

    ' Definir el rango de fechas (hoy + 30 días)
    FechaInicio = Date
    FechaFin = Date + 30

    ' Limpiar la hoja y colocar encabezados
    With ws
        .Cells.Clear
        .Range("A1:E1").Value = Array("Asunto", "Inicio", "Fin", "Organizador", "Conflicto")

        ' Estilo para los encabezados
        With .Range("A1:E1")
            .Font.Bold = True
            .Font.Size = 12
            .Interior.Color = RGB(0, 112, 192) ' Color de fondo azul
            .Font.Color = RGB(255, 255, 255) ' Color de fuente blanco
            .HorizontalAlignment = xlCenter
        End With
    End With

    ' Variable para la fila en Excel
    UltimaFila = 2

    ' Comparar eventos uno por uno sin usar arreglos grandes
    For Each Appt In Items
        If Appt.Start >= FechaInicio And Appt.End <= FechaFin Then
            ' Comparar con las siguientes reuniones
            For Each Appt2 In Items
                If Appt2.Start >= FechaInicio And Appt2.End <= FechaFin Then
                    ' Evitar comparar una reunión consigo misma
                    If Appt.Start < Appt2.End And Appt.End > Appt2.Start And Appt.Subject <> Appt2.Subject Then
                        ' Registrar el conflicto en la hoja
                        With ws
                            .Cells(UltimaFila, 1).Value = Appt.Subject
                            .Cells(UltimaFila, 2).Value = Appt.Start
                            .Cells(UltimaFila, 3).Value = Appt.End
                            .Cells(UltimaFila, 4).Value = Appt.Organizer
                            .Cells(UltimaFila, 5).Value = "? Conflicto con: " & Appt2.Subject

                            ' Aplicar formato de fecha
                            .Cells(UltimaFila, 2).NumberFormat = "dd/mm/yyyy hh:mm AM/PM"
                            .Cells(UltimaFila, 3).NumberFormat = "dd/mm/yyyy hh:mm AM/PM"

                            ' Resaltar la fila de conflicto
                            .Range(.Cells(UltimaFila, 1), .Cells(UltimaFila, 5)).Interior.Color = RGB(255, 235, 156) ' Amarillo claro
                        End With
                        UltimaFila = UltimaFila + 1
                        ConflictosDetectados = True ' Marcar que se ha detectado un conflicto
                    End If
                End If
            Next Appt2
        End If
    Next Appt

    ' Ajustar el ancho de las columnas
    With ws
        .Columns("A:E").AutoFit ' Ajusta automáticamente el ancho de las columnas
        .Columns("B:C").ColumnWidth = 20 ' Asegurar que las fechas tengan suficiente espacio
    End With

    ' Aplicar bordes a la tabla
    With ws.Range("A1:E" & UltimaFila - 1)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With

    ' Liberar memoria
    Set Items = Nothing
    Set Calendar = Nothing
    Set Namespace = Nothing
    Set OutlookApp = Nothing

    ' Mostrar mensaje dependiendo de si se han detectado conflictos
    If ConflictosDetectados Then
        MsgBox "Conflictos detectados. Ver la hoja 'Conflictos'.", vbInformation, "Detección Completa"
    Else
        MsgBox "No se detectaron conflictos en el calendario.", vbInformation, "Sin Conflictos"
    End If
End Sub


```

### Macro 2: Obtener los Últimos 20 Correos de Outlook

```vba
Sub ObtenerUltimosCorreos()
    Sub ObtenerUltimosCorreos()
    Dim OutlookApp As Object
    Dim Namespace As Object
    Dim Inbox As Object
    Dim Items As Object
    Dim Mail As Object
    Dim i As Integer
    Dim ws As Worksheet
    
    ' Iniciar Outlook
    On Error Resume Next
    Set OutlookApp = GetObject(, "Outlook.Application")
    If OutlookApp Is Nothing Then Set OutlookApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    
    ' Obtener la Bandeja de Entrada
    Set Namespace = OutlookApp.GetNamespace("MAPI")
    Set Inbox = Namespace.GetDefaultFolder(6) ' 6 = olFolderInbox
    Set Items = Inbox.Items
    Items.Sort "[ReceivedTime]", True ' Ordenar por fecha descendente
    
    ' Verificar si la hoja "Correos" existe, si no, crearla
    On Error Resume Next
    Set ws = Sheets("Correos")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = "Correos"
    End If
    
    ' Limpiar hoja y establecer encabezados
    With ws
        .Cells.Clear
        .Range("A1:D1").Value = Array("Remitente", "Asunto", "Fecha", "Cuerpo (resumen)")

        ' Estilo para los encabezados
        With .Range("A1:D1")
            .Font.Bold = True
            .Font.Size = 12
            .Interior.Color = RGB(0, 112, 192) ' Fondo azul
            .Font.Color = RGB(255, 255, 255) ' Fuente blanca
            .HorizontalAlignment = xlCenter
        End With
    End With
    
    ' Obtener los últimos 20 correos
    i = 2
    For Each Mail In Items
        If i > 21 Then Exit For ' Solo 20 correos
        
        ' Verificar que el ítem sea un correo
        If Mail.Class = 43 Then ' 43 = olMailItem
            With ws
                .Cells(i, 1).Value = Mail.SenderName
                .Cells(i, 2).Value = Mail.Subject
                .Cells(i, 3).Value = Mail.ReceivedTime
                .Cells(i, 4).Value = Left(Mail.Body, 100) ' Solo primeros 100 caracteres

                ' Aplicar formato de fecha
                .Cells(i, 3).NumberFormat = "dd/mm/yyyy hh:mm AM/PM"

                ' Aplicar color de fondo a las filas alternadas para mayor legibilidad
                If i Mod 2 = 0 Then
                    .Range(.Cells(i, 1), .Cells(i, 4)).Interior.Color = RGB(240, 240, 240) ' Gris claro
                End If
            End With
            i = i + 1
        End If
    Next Mail
    
    ' Ajustar el ancho de las columnas
    With ws
        .Columns("A:D").AutoFit ' Ajusta automáticamente el ancho de las columnas
        .Columns("C").ColumnWidth = 20 ' Asegura que la columna de fecha tenga suficiente espacio
    End With
    
    ' Aplicar bordes a la tabla
    With ws.Range("A1:D" & i - 1)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With
    
    ' Liberar memoria
    Set Items = Nothing
    Set Inbox = Nothing
    Set Namespace = Nothing
    Set OutlookApp = Nothing

    MsgBox "Correos cargados con éxito.", vbInformation, "Importación Completada"
End Sub

```

### Macro 3: Obtener Eventos de Outlook

```vba
Sub ObtenerEventosOutlook()
    Dim OutlookApp As Object ' Usamos Object para evitar problemas de compatibilidad
    Dim Namespace As Object
    Dim Calendar As Object
    Dim Items As Object
    Dim Appt As Object
    Dim Filtro As String
    Dim i As Integer
    Dim FechaInicio As Date
    Dim FechaFin As Date
    Dim ws As Worksheet
    
    ' Iniciar Outlook
    On Error Resume Next
    Set OutlookApp = GetObject(, "Outlook.Application")
    If OutlookApp Is Nothing Then Set OutlookApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    
    ' Obtener el calendario principal
    Set Namespace = OutlookApp.GetNamespace("MAPI")
    Set Calendar = Namespace.GetDefaultFolder(9) ' 9 = olFolderCalendar

    ' Definir el rango de fechas (últimos 30 días hasta los próximos 180 días)
    FechaInicio = Date - 30
    FechaFin = Date + 180
    Filtro = "[Start] >= '" & Format(FechaInicio, "ddddd hh:mm AMPM") & "' AND [End] <= '" & Format(FechaFin, "ddddd hh:mm AMPM") & "'"

    ' Obtener los eventos
    Set Items = Calendar.Items
    Items.Sort "[Start]"
    Items.IncludeRecurrences = True
    Set Items = Items.Restrict(Filtro)

    ' Verificar si la hoja "Calendario" existe, si no, crearla
    On Error Resume Next
    Set ws = Sheets("Calendario")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = "Calendario"
    End If

    ' Limpiar hoja y definir encabezados
    With ws
        .Cells.Clear
        .Range("A1:H1").Value = Array("Asunto", "Inicio", "Fin", "Ubicación", "Descripción", "Categoría", "Estado", "Asistentes")

        ' Estilo para los encabezados
        With .Range("A1:H1")
            .Font.Bold = True
            .Font.Size = 12
            .Interior.Color = RGB(0, 112, 192) ' Fondo azul
            .Font.Color = RGB(255, 255, 255) ' Fuente blanca
            .HorizontalAlignment = xlCenter
        End With
    End With

    ' Recorrer eventos y escribir en Excel
    i = 2
    For Each Appt In Items
        On Error Resume Next
        ws.Cells(i, 1).Value = Appt.Subject
        ws.Cells(i, 2).Value = Appt.Start
        ws.Cells(i, 3).Value = Appt.End
        ws.Cells(i, 4).Value = Appt.Location
        ws.Cells(i, 5).Value = Appt.Body ' Descripción o detalles
        ws.Cells(i, 6).Value = Appt.Categories ' Categoría asignada
        ws.Cells(i, 7).Value = Appt.BusyStatus ' Estado del evento (Libre, Ocupado, etc.)
        
        ' Obtener los asistentes
        Dim Attendees As String
        Dim Recipients As Object
        Set Recipients = Appt.Recipients
        Attendees = ""
        If Not Recipients Is Nothing Then
            Dim j As Integer
            For j = 1 To Recipients.Count
                Attendees = Attendees & Recipients.Item(j).Name & "; "
            Next j
        End If
        ws.Cells(i, 8).Value = Attendees
        
        ' Aplicar formato de fecha
        ws.Cells(i, 2).NumberFormat = "dd/mm/yyyy hh:mm AM/PM"
        ws.Cells(i, 3).NumberFormat = "dd/mm/yyyy hh:mm AM/PM"
        
        ' Filas alternadas con color suave para mayor legibilidad
        If i Mod 2 = 0 Then
            ws.Range(ws.Cells(i, 1), ws.Cells(i, 8)).Interior.Color = RGB(240, 240, 240) ' Gris claro
        End If
        
        i = i + 1
        On Error GoTo 0
    Next Appt

    ' Ajustar el ancho de las columnas
    With ws
        .Columns("A:H").AutoFit ' Ajusta automáticamente el ancho de las columnas
        .Columns("B:C").ColumnWidth = 20 ' Asegura que las columnas de fecha tengan suficiente espacio
    End With

    ' Aplicar bordes a la tabla
    With ws.Range("A1:H" & i - 1)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With
    
    ' Liberar memoria
    Set Items = Nothing
    Set Calendar = Nothing
    Set Namespace = Nothing
    Set OutlookApp = Nothing

    MsgBox "Eventos cargados con éxito.", vbInformation, "Importación Completada"
End Sub


```

> **Nota**: Cada macro debe estar en su propio procedimiento dentro de un módulo. Si lo prefieres, puedes añadir todas las macros dentro del mismo módulo.

---

## **Paso 3: Configurar las hojas de Excel**

1. **Crear hojas de trabajo en el libro de Excel**:
   - Crea tres hojas en el libro de trabajo y nómbralas de la siguiente forma:
     - **Conflictos** (para la macro de conflictos).
     - **Correos** (para la macro de correos).
     - **Calendario** (para la macro de eventos de Outlook).

> **Nota**: Si la hoja **"Conflictos"** no existe, la macro la creará automáticamente. Lo mismo ocurre con las otras hojas.

---

## **Paso 4: Ejecutar las macros**

### Para ejecutar cada macro:

1. Ve a la pestaña **Desarrollador** en Excel y haz clic en **Macros**.
2. Selecciona la macro que deseas ejecutar y haz clic en **Ejecutar**.

---

### **Instrucciones para cada macro**:

1. **Detectar Conflictos en el Calendario de Outlook**:
   - **¿Qué hace?** Detecta eventos en el calendario de Outlook que se solapan (es decir, que tienen horarios que se superponen).
   - **¿Cómo ejecutarlo?**
     1. Asegúrate de que Outlook esté abierto y configurado correctamente.
     2. Ejecuta la macro **DetectarConflictosCalendario**.
     3. La macro verificará si hay eventos en conflicto dentro de los próximos 30 días.
     4. Los eventos en conflicto se mostrarán en la hoja **Conflictos** con detalles como el asunto, el inicio, el fin, el organizador y los conflictos encontrados.

2. **Obtener los Últimos 20 Correos de Outlook**:
   - **¿Qué hace?** Obtiene los últimos 20 correos recibidos en tu bandeja de entrada de Outlook.
   - **¿Cómo ejecutarlo?**
     1. Asegúrate de que Outlook esté abierto y configurado correctamente.
     2. Ejecuta la macro **ObtenerUltimosCorreos**.
     3. La macro recuperará los últimos 20 correos de tu bandeja de entrada.
     4. Los correos se mostrarán en la hoja **Correos** con detalles como remitente, asunto, fecha y un resumen del cuerpo del correo.

3. **Obtener Eventos de Outlook**:
   - **¿Qué hace?** Recupera los eventos de tu calendario de Outlook en el rango de los próximos 180 días.
   - **¿Cómo ejecutarlo?**
     1. Asegúrate de que Outlook esté abierto y configurado correctamente.
     2. Ejecuta la macro **ObtenerEventosOutlook**.
     3. Los eventos se mostrarán en la hoja **Calendario**, con detalles como el asunto, la fecha de inicio y fin, la ubicación, la descripción, la categoría, el estado y los asistentes.

---

## **Paso 5: Personalización y Estilo (Opcional)**

1. **Personaliza las hojas** según tus preferencias:
   - Puedes cambiar los colores y el formato de las celdas, ajustar el tamaño de las columnas, o agregar más campos si lo necesitas.
   
2. **Mejorar la experiencia de usuario**:
   - Puedes agregar botones en la hoja de Excel para ejecutar las macros más fácilmente. Para ello:
     - Ve a la pestaña **Desarrollador** → **Insertar** → **Botón**.
     - Asocia cada botón con una de las macros que has creado.
   
---

## **Paso 6: Resultado esperado**

1. **Hoja 'Conflictos'**:
   - Si hay conflictos, verás los detalles de los eventos que se solapan. Si no hay conflictos, recibirás un mensaje diciendo que no se encontraron conflictos.

2. **Hoja 'Correos'**:
   - Verás los últimos 20 correos recibidos en tu bandeja de entrada con detalles sobre el remitente, asunto, fecha y un resumen del cuerpo del correo.

3. **Hoja 'Calendario'**:
   - Verás los eventos que ocurren en los próximos 180 días con detalles sobre el asunto, inicio, fin, ubicación, descripción, categoría, estado y asistentes.

---

## **Paso 7: Solución de problemas**

1. **Error de "Memoria insuficiente"**:
   - Si encuentras este error, cierra otros programas o aplicaciones que consuman memoria, o reinicia Excel.
   
2. **Error de "Subíndice fuera del intervalo"**:
   - Asegúrate de que todas las hojas de trabajo estén correctamente nombradas y que Outlook esté correctamente configurado y abierto.

3. **Outlook no se conecta**:
   - Verifica que Outlook esté abierto antes de ejecutar las macros.

---

## **Conclusión**

Con este laboratorio has aprendido cómo usar Excel y VBA para interactuar con Outlook y automatizar tareas como la detección de conflictos en el calendario, la obtención de correos y la visualización de eventos. Este tipo de automatización puede ahorrarte tiempo y hacer tu trabajo mucho más eficiente.
