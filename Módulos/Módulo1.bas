Attribute VB_Name = "Módulo1"
Option Explicit

Sub CheckPDFGerado()
    Dim pastaSelecionada As String
    Dim arquivosPDF As Object
    Dim fso As Object, arquivo As Object
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim celula As Range
    Dim nomeArquivo As String
    Dim faltando As String
    Dim dialogo As FileDialog
    Dim pastaPadrao As String
    Dim PrimeiraData As Date
    Dim i As Long

    ' GARANTIA: Sempre aponta para a aba "Agenda"
    Set ws = ThisWorkbook.Sheets("Agenda")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set arquivosPDF = CreateObject("Scripting.Dictionary")

    ' Caminho padrão
    pastaPadrao = "C:\Users\sidney.oliveira\OneDrive\Linx Implantações\Linx DMS - Apollo"

    ' Abrir seletor de pasta
    Set dialogo = Application.FileDialog(msoFileDialogFolderPicker)
    dialogo.Title = "Selecione a pasta onde estão os arquivos PDF"
    dialogo.InitialFileName = pastaPadrao

    If dialogo.Show <> -1 Then
        MsgBox "Nenhuma pasta foi selecionada.", vbExclamation
        Exit Sub
    End If
    pastaSelecionada = dialogo.SelectedItems(1) & "\"

    ' Listar arquivos PDF
    For Each arquivo In fso.GetFolder(pastaSelecionada).Files
        If LCase(fso.GetExtensionName(arquivo.Name)) = "pdf" Then
            arquivosPDF.Add arquivo.Name, True
        End If
    Next arquivo

    ' Encontrar a primeira data válida na coluna D
    ultimaLinha = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    PrimeiraData = 0
    For i = 3 To ultimaLinha
        If IsDate(ws.Range("D" & i).Value) Then
            PrimeiraData = ws.Range("D" & i).Value
            Exit For
        End If
    Next i

    ultimaLinha = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    faltando = ""

    For Each celula In ws.Range("B3:B" & ultimaLinha)
        If IsDate(ws.Range("D" & celula.Row).Value) Then
            If PrimeiraData = 0 Or ws.Range("D" & celula.Row).Value >= PrimeiraData Then
                If Trim(celula.Value) <> "" Then
                    nomeArquivo = Trim(celula.Value) & ".pdf"
                    If Not arquivosPDF.exists(nomeArquivo) Then
                        faltando = faltando & " - " & celula.Value & vbNewLine
                    End If
                End If
            End If
        End If
    Next celula

    If faltando = "" Then
        MsgBox "Ok, todas as agendas (PDFs) foram geradas com sucesso!", vbInformation
    Else
        MsgBox "As agendas abaixo não foram criadas:" & vbNewLine & vbNewLine & faltando, vbExclamation
    End If
End Sub

Sub CriarReunioesOutlook()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim TotalEventos As Long
    Dim respostaUsuario As VbMsgBoxResult

    ' GARANTIA: Sempre aponta para a aba "Agenda"
    Set ws = ThisWorkbook.Sheets("Agenda")
    LastRow = ws.Cells(ws.Rows.Count, "Q").End(xlUp).Row
    TotalEventos = 0

    For i = 3 To LastRow
        If UCase(Trim(ws.Range("Q" & i).Value)) = "S" Then
            TotalEventos = TotalEventos + 1
        End If
    Next i
    
    If TotalEventos = 0 Then
        MsgBox "Nenhuma agenda marcada para criar (Coluna Q = S).", vbInformation
        Exit Sub
    End If

    respostaUsuario = MsgBox("Deseja criar as agenda(s)?" & vbNewLine & " - Total: " & TotalEventos & " evento(s)", vbYesNo + vbQuestion, "Confirmação")
    If respostaUsuario = vbNo Then Exit Sub

    Dim OutlookApp As Object, OutlookNamespace As Object
    Dim CalendarFolder As Object, OutlookMeeting As Object
    Dim FilteredItems As Object, ExistingAppointment As Object
    Dim Title As String, Subject As String
    Dim RequiredAttendees As String, OptionalAttendees As String
    Dim Attendee As Variant, Recipient As Object
    Dim MeetingDate As Date, StartTime As Date, EndTime As Date
    Dim DataHoraInicio As Date, DataHoraFim As Date
    Dim Conflict As Boolean, CreatedMeetings As Boolean
    Dim CountMeetings As Long, FilePath As String
    Dim Restriction As String

    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")
    Set CalendarFolder = OutlookNamespace.GetDefaultFolder(9)

    CreatedMeetings = False
    CountMeetings = 0
    FilePath = ThisWorkbook.Path & "\Manual Treinamento Microsoft Teams - Cliente.pdf"

    For i = 3 To LastRow
        If UCase(Trim(ws.Range("Q" & i).Value)) = "S" Then
            Title = ws.Range("O" & i).Value
            MeetingDate = ws.Range("D" & i).Value
            StartTime = ws.Range("E" & i).Value
            EndTime = ws.Range("F" & i).Value
            Subject = ws.Range("P" & i).Value
            RequiredAttendees = ws.Range("K" & i).Value
            OptionalAttendees = ws.Range("R" & i).Value
            
            DataHoraInicio = MeetingDate + StartTime
            DataHoraFim = MeetingDate + EndTime

            If Right(RequiredAttendees, 1) = ";" Then RequiredAttendees = Left(RequiredAttendees, Len(RequiredAttendees) - 1)
            If Right(OptionalAttendees, 1) = ";" Then OptionalAttendees = Left(OptionalAttendees, Len(OptionalAttendees) - 1)

            Restriction = "[Start] < '" & Format(DataHoraFim, "mm/dd/yyyy hh:nn AMPM") & "' AND [End] > '" & Format(DataHoraInicio, "mm/dd/yyyy hh:nn AMPM") & "'"
            Set FilteredItems = CalendarFolder.Items.Restrict(Restriction)
            
            Conflict = False
            For Each ExistingAppointment In FilteredItems
                Conflict = True
                Exit For
            Next ExistingAppointment

            If Conflict Then
                If MsgBox("Já existe um compromisso neste horário: " & ExistingAppointment.Subject & vbNewLine & "Deseja criar a reunião mesmo assim?", vbYesNo + vbExclamation, "Conflito de Horário") = vbNo Then
                    GoTo NextRecord
                End If
            End If

            Set OutlookMeeting = OutlookApp.CreateItem(1)
            With OutlookMeeting
                .Subject = Title
                .Start = DataHoraInicio
                .End = DataHoraFim
                .Body = Subject
                .Location = "Microsoft Teams Meeting"
                .MeetingStatus = 1
                .IsOnlineMeeting = True
                
                If RequiredAttendees <> "" Then
                    For Each Attendee In Split(RequiredAttendees, ";")
                        .Recipients.Add Trim(Attendee)
                    Next Attendee
                End If

                If OptionalAttendees <> "" Then
                    For Each Attendee In Split(OptionalAttendees, ";")
                        If Trim(Attendee) <> "" Then
                            Set Recipient = .Recipients.Add(Trim(Attendee))
                            Recipient.Type = 2
                        End If
                    Next Attendee
                End If

                On Error Resume Next
                .Attachments.Add FilePath
                On Error GoTo 0

                .Recipients.ResolveAll
                .Display
                Application.Wait (Now + TimeValue("00:00:02"))
                
                If .Recipients.Count > 0 Then
                    .Send
                Else
                    .Close 0
                End If
            End With

            CreatedMeetings = True
            CountMeetings = CountMeetings + 1
            Set OutlookMeeting = Nothing
        End If
NextRecord:
    Next i

    MsgBox "Processo concluído! Foram criadas " & CountMeetings & " agendas."
End Sub

Sub VerificarAgendasCriadas()
    Dim OutlookApp As Object, OutlookNamespace As Object, CalendarFolder As Object
    Dim CalendarItems As Object, FilteredItems As Object, Appointment As Object
    Dim ws As Worksheet
    Dim LastRow As Long, i As Long
    Dim Titulo As String, NaoEncontradas As String
    Dim DataHoraInicio As Date, PrimeiraData As Date
    Dim TotalEncontradas As Long, TotalNaoEncontradas As Long
    Dim Restriction As String, Encontrado As Boolean

    ' GARANTIA: Sempre aponta para a aba "Agenda"
    Set ws = ThisWorkbook.Sheets("Agenda")
    LastRow = ws.Cells(ws.Rows.Count, "O").End(xlUp).Row

    For i = 3 To LastRow
        If IsDate(ws.Range("D" & i).Value) Then
            PrimeiraData = ws.Range("D" & i).Value
            Exit For
        End If
    Next i
    If PrimeiraData = 0 Then Exit Sub

    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")
    Set CalendarFolder = OutlookNamespace.GetDefaultFolder(9)
    Set CalendarItems = CalendarFolder.Items
    CalendarItems.Sort "[Start]"
    CalendarItems.IncludeRecurrences = True

    Restriction = "[Start] >= '" & Format(PrimeiraData, "mm/dd/yyyy") & " 12:00 AM'"
    Set FilteredItems = CalendarItems.Restrict(Restriction)

    TotalEncontradas = 0: TotalNaoEncontradas = 0: NaoEncontradas = ""

    For i = 3 To LastRow
        If Trim(ws.Range("O" & i).Value) <> "" And IsDate(ws.Range("D" & i).Value) Then
            Titulo = Trim(ws.Range("O" & i).Value)
            DataHoraInicio = ws.Range("D" & i).Value + ws.Range("E" & i).Value
            Encontrado = False
            
            For Each Appointment In FilteredItems
                If Abs(DateDiff("n", Appointment.Start, DataHoraInicio)) <= 2 Then
                    If UCase(Trim(Appointment.Subject)) = UCase(Titulo) Then
                        Encontrado = True
                        Exit For
                    End If
                End If
            Next Appointment

            If Encontrado Then
                TotalEncontradas = TotalEncontradas + 1
            Else
                TotalNaoEncontradas = TotalNaoEncontradas + 1
                NaoEncontradas = NaoEncontradas & " - " & Titulo & " (" & Format(DataHoraInicio, "dd/mm/yyyy HH:mm") & ")" & vbNewLine
            End If
        End If
    Next i

    Dim mensagem As String
    mensagem = "Total encontradas no Outlook: " & TotalEncontradas & vbNewLine & "Total não encontradas: " & TotalNaoEncontradas
    If TotalNaoEncontradas > 0 Then mensagem = mensagem & vbNewLine & vbNewLine & "Faltantes:" & vbNewLine & NaoEncontradas
    MsgBox mensagem, vbInformation, "Resumo"
End Sub

Sub LinkReuniao()
    Dim ws As Worksheet
    Dim LastRow As Long, i As Long, TotalLinhas As Long
    Dim PrimeiraData As Date
    Dim OutlookApp As Object, OutlookNamespace As Object, CalendarFolder As Object
    Dim CalendarItems As Object, FilteredItems As Object, Appointment As Object
    Dim Titulo As String, corpo As String, LinkReuniao As String
    Dim DataHoraInicio As Date
    Dim inicioLink As Long, fimLink As Long
    Dim Encontrado As Boolean, Restriction As String

    ' GARANTIA: Sempre aponta para a aba "Agenda"
    Set ws = ThisWorkbook.Sheets("Agenda")
    LastRow = ws.Cells(ws.Rows.Count, "O").End(xlUp).Row
    TotalLinhas = 0

    For i = 3 To LastRow
        If IsDate(ws.Range("D" & i).Value) Then
            PrimeiraData = ws.Range("D" & i).Value
            Exit For
        End If
    Next i

    For i = 3 To LastRow
        If Trim(ws.Range("O" & i).Value) <> "" Then TotalLinhas = TotalLinhas + 1
    Next i

    If MsgBox("Deseja gerar o(s) link(s) da(s) reuniões?" & vbNewLine & "Total: " & TotalLinhas, vbYesNo + vbQuestion) = vbNo Then Exit Sub

    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")
    Set CalendarFolder = OutlookNamespace.GetDefaultFolder(9)
    Set CalendarItems = CalendarFolder.Items
    CalendarItems.Sort "[Start]"
    CalendarItems.IncludeRecurrences = True

    Restriction = "[Start] >= '" & Format(PrimeiraData, "mm/dd/yyyy") & " 12:00 AM'"
    Set FilteredItems = CalendarItems.Restrict(Restriction)

    For i = 3 To LastRow
        If Trim(ws.Range("O" & i).Value) <> "" And IsDate(ws.Range("D" & i).Value) Then
            Titulo = Trim(ws.Range("O" & i).Value)
            DataHoraInicio = ws.Range("D" & i).Value + ws.Range("E" & i).Value
            Encontrado = False
            LinkReuniao = ""

            For Each Appointment In FilteredItems
                If Abs(DateDiff("n", Appointment.Start, DataHoraInicio)) <= 2 Then
                    If UCase(Trim(Appointment.Subject)) = UCase(Titulo) Then
                        corpo = Appointment.Body
                        inicioLink = InStr(1, corpo, "https://teams.microsoft.com/l/meetup-join")
                        
                        If inicioLink > 0 Then
                            fimLink = InStr(inicioLink, corpo, vbCrLf)
                            If fimLink = 0 Then fimLink = InStr(inicioLink, corpo, " ")
                            If fimLink = 0 Then fimLink = InStr(inicioLink, corpo, ">")
                            If fimLink = 0 Then fimLink = Len(corpo) + 1
                            
                            LinkReuniao = Mid(corpo, inicioLink, fimLink - inicioLink)
                            LinkReuniao = Replace(Replace(LinkReuniao, ">", ""), """", "")
                            LinkReuniao = Trim(LinkReuniao)
                        End If
                        Encontrado = True
                        Exit For
                    End If
                End If
            Next Appointment

            If Encontrado And LinkReuniao <> "" Then
                ws.Range("V" & i).Value = LinkReuniao
            ElseIf Encontrado And LinkReuniao = "" Then
                ws.Range("V" & i).Value = "Link não gerado no Outlook (Aguarde e tente novamente)"
            Else
                ws.Range("V" & i).Value = "Agenda não encontrada"
            End If
        End If
    Next i

    MsgBox "Links processados com sucesso!", vbInformation
End Sub

Sub CancelarAgendasOutlook()
    Dim ws As Worksheet
    Dim LastRow As Long, i As Long, TotalParaCancelar As Long
    Dim OutlookApp As Object, OutlookNamespace As Object, CalendarFolder As Object
    Dim CalendarItems As Object, FilteredItems As Object, Appointment As Object
    Dim Titulo As String, DataHoraInicio As Date
    Dim Canceladas As Long, PrimeiraData As Date, Restriction As String

    ' GARANTIA: Sempre aponta para a aba "Agenda"
    Set ws = ThisWorkbook.Sheets("Agenda")
    LastRow = ws.Cells(ws.Rows.Count, "U").End(xlUp).Row
    TotalParaCancelar = 0

    For i = 3 To LastRow
        If UCase(Trim(ws.Range("U" & i).Value)) = "S" Then
            TotalParaCancelar = TotalParaCancelar + 1
            If PrimeiraData = 0 Then PrimeiraData = ws.Range("D" & i).Value
        End If
    Next i

    If TotalParaCancelar = 0 Then
        MsgBox "Nenhuma agenda marcada para cancelar (Coluna U = S).", vbInformation
        Exit Sub
    End If
    
    If MsgBox("Deseja remover " & TotalParaCancelar & " agenda(s)?", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")
    Set CalendarFolder = OutlookNamespace.GetDefaultFolder(9)
    
    Set CalendarItems = CalendarFolder.Items
    CalendarItems.Sort "[Start]"
    CalendarItems.IncludeRecurrences = True
    
    Restriction = "[Start] >= '" & Format(PrimeiraData, "mm/dd/yyyy") & " 12:00 AM'"
    Set FilteredItems = CalendarItems.Restrict(Restriction)

    Canceladas = 0

    For i = 3 To LastRow
        If UCase(Trim(ws.Range("U" & i).Value)) = "S" Then
            Titulo = Trim(ws.Range("O" & i).Value)
            DataHoraInicio = ws.Range("D" & i).Value + ws.Range("E" & i).Value

            For Each Appointment In FilteredItems
                If Abs(DateDiff("n", Appointment.Start, DataHoraInicio)) <= 2 Then
                    If UCase(Trim(Appointment.Subject)) = UCase(Titulo) Then
                        Appointment.Delete
                        Canceladas = Canceladas + 1
                        Exit For
                    End If
                End If
            Next Appointment
        End If
    Next i

    MsgBox "Processo concluído! Total canceladas: " & Canceladas, vbInformation
End Sub

