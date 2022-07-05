
Option Explicit
Dim getList         As Worksheet, inputList As Worksheet, v As Integer
Const SUPPLIER_RESULTS As String = "https://zakupki.gov.ru/epz/order/notice/ea44/view/supplier-results.html?regNumber="
Const ZAK_NOTICE    As String = "https://zakupki.gov.ru/epz/order/notice/ea44/view/documents.html?regNumber="
Const ZAK_COMMON    As String = "https://zakupki.gov.ru/epz/order/notice/ea44/view/common-info.html?regNumber="
Const ZAK_COMMON_CONTRACT As String = "https://zakupki.gov.ru/epz/contract/contractCard/common-info.html?reestrNumber="
Const CONTRACT_CONCLUSION_COMMON As String = "https://zakupki.gov.ru/epz/order/notice/rpec/common-info.html?regNumber="

Sub Scrapp()

    Dim response    As Integer, lastrow As Integer
    Dim regionUTCDict As Scripting.Dictionary
    
    response = MsgBox("Начать выполнение?", vbOKCancel, "Заполнение реестра")
    
    If response = vbCancel Then Exit Sub
    
    Call toggle_screen_upd
    
    Set getList = ActiveWorkbook.ActiveSheet
    Call NewList23
   
    Set inputList = ActiveWorkbook.Sheets(1)
    getList.Activate
   
    Call perenosNomera
    inputList.Activate
    lastrow = inputList.Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
   
    Call Scrap2022
    Call sposob
    Call tip
    Call predmet
    Call organizator
    Call ploschadka
    Call preference
    Call sum_of_preference
    
    Set regionUTCDict = RegionUTCDictionary()
    Call TimeOkonch(regionUTCDict)
    Call TimeProved
    Call Ssylka
    Call dataSeg
    Call plusNomer
    Call Formatirovan1
    
    Set regionUTCDict = Nothing
    
    Call toggle_screen_upd
    
    MsgBox "Выполнено"
End Sub

Private Sub utc(regionUTCDict As Scripting.Dictionary)
    Dim lastrow     As Integer, s As String, i As Integer, pos As Long
    lastrow = getList.Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    For i = 2 To lastrow
        If regionUTCDict.Exists(inputList.Range("Region")(i).Value) Then
            inputList.Range("UTC")(i).Value = regionUTCDict.Item(inputList.Range("Region")(i).Value)
        Else
            inputList.Range("UTC")(i).Value = -999
        End If
    Next i
End Sub
Private Sub TimeOkonch(regionUTCDict As Scripting.Dictionary)
    Dim lastrow     As Integer, i As Integer, t As Date
    Call utc(regionUTCDict)
    lastrow = getList.Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    For i = 2 To lastrow
        If Not IsEmpty(inputList.Range("TimeOkonch")(i).Value) And inputList.Range("UTC")(i).Value <> -999 Then
            t = DateAdd("h", (-(inputList.Range("UTC")(i) - 3)), inputList.Range("TimeOkonch")(i))
            If DateDiff("d", t, inputList.Range("DataOkonch")(i)) = 1 Then
                inputList.Range("DataOkonch")(i).Value = DateAdd("d", -1, inputList.Range("DataOkonch")(i))
                inputList.Range("TimeOkonch")(i).Value = DateAdd("d", -1, inputList.Range("TimeOkonch")(i))
            ElseIf DateDiff("d", t, inputList.Range("DataOkonch")(i)) = -1 Then
                inputList.Range("DataOkonch")(i).Value = DateAdd("d", 1, inputList.Range("DataOkonch")(i))
                inputList.Range("TimeOkonch")(i).Value = DateAdd("d", 1, inputList.Range("TimeOkonch")(i))
            End If
            inputList.Range("TimeOkonch")(i).Value = t
            inputList.Range("DataOkonch")(i).Value = t
        End If
    Next i
End Sub
Private Sub RemoveShitTime()
    Dim lastrow     As Integer, i As Integer
    lastrow = getList.Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    For i = 2 To lastrow
        If UCase(inputList.Range("TimeProved")(i)) Like UCase("*время*") Then
            inputList.Range("TimeProved")(i).Value = ""
        End If
    Next i
End Sub
Private Sub TimeProved()
    Dim lastrow     As Integer, i As Integer, t As Date
    Call RemoveShitTime
    lastrow = getList.Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    For i = 2 To lastrow
        If Not IsEmpty(inputList.Range("TimeProved")(i)) And inputList.Range("UTC")(i).Value <> -999 Then
            t = DateAdd("h", (-(inputList.Range("UTC")(i) - 3)), inputList.Range("TimeProved")(i))
            If DateDiff("d", t, inputList.Range("DataProved")(i)) = 1 Then
                inputList.Range("DataProved")(i).Value = DateAdd("d", -1, inputList.Range("DataProved")(i))
                inputList.Range("TimeProved")(i).Value = DateAdd("d", -1, inputList.Range("TimeProved")(i))
            ElseIf DateDiff("d", t, inputList.Range("DataProved")(i)) = -1 Then
                inputList.Range("DataProved")(i).Value = DateAdd("d", 1, inputList.Range("DataProved")(i))
                inputList.Range("TimeProved")(i).Value = DateAdd("d", 1, inputList.Range("TimeProved")(i))
            End If
            inputList.Range("TimeProved")(i).Value = t
            
            If inputList.Range("Sposob")(i).Value = "ЗКЭФ" Then
                inputList.Range("DataProved")(i).Value = inputList.Range("DataOkonch")(i).Value
                inputList.Range("TimeProved")(i).Value = inputList.Range("DataOkonch")(i).Value
            End If
        End If
    Next i
End Sub
Private Sub perenosNomera()
    Dim lastrow     As Integer, nomerCol As Integer, arr() As Variant, getRange As Range
    lastrow = getList.Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    nomerCol = getList.Cells.Find(What:="*реестр*номер*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).Column
    Set getRange = getList.Range(Cells(2, nomerCol), Cells(lastrow, nomerCol))
    arr = getRange.Resize(lastrow - 1, 1)
    inputList.Range("Nomer")(2).Resize(lastrow - 1, 1) = arr
End Sub
Private Sub dataSeg()
    Dim i           As Integer, lastrow As Integer
    lastrow = inputList.Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    For i = 2 To lastrow
        inputList.Range("Date")(i) = DateValue(Now)
    Next i
End Sub
Private Sub plusNomer()
    Dim i           As Integer, lastrow As Integer
    lastrow = inputList.Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    For i = 2 To lastrow
        inputList.Range("Nomer")(i) = "№" & inputList.Range("Nomer")(i)
    Next i
End Sub
Private Sub Ssylka()
    Dim i           As Integer, lastrow As Integer
    lastrow = inputList.Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    For i = 2 To lastrow
        If inputList.Range("Ploschadka")(i) = "РТС-тендер" Then
            inputList.Range("Ssylka")(i).Value = "http://www.rts-tender.ru/auctionsearch/ctl/procDetail/mid/691/procId/" & CStr(inputList.Range("Nomer")(i).Value) & "/etpName/fks"
        ElseIf inputList.Range("Ploschadka")(i) = "Росельторг" Then
            inputList.Range("Ssylka")(i).Value = "https://www.roseltorg.ru/trade/view/?id=" & CStr(inputList.Range("Nomer")(i).Value)
        End If
    Next i
End Sub
Private Sub ZKTime()
    Dim i           As Integer, lastrow As Integer, splitter As Variant, a As String
    lastrow = inputList.Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    For i = 2 To lastrow
        If inputList.Range("Sposob")(i) Like "?К" Or inputList.Range("Sposob")(i) = "ЗКЭФ" Then
            inputList.Range("TimeProved")(i).Value = inputList.Range("TimeOkonch")(i).Value
            inputList.Range("DataProved")(i).Value = inputList.Range("DataOkonch")(i).Value
            inputList.Range("OkonchRasm")(i).Value = inputList.Range("DataOkonch")(i).Value
        End If
    Next i
End Sub
Private Sub sposob()
    Dim repRange    As Range, lastrow As Integer, Col As Integer
    lastrow = inputList.Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    Col = inputList.Range("Sposob").Column
    Set repRange = inputList.Range(Cells(2, Col), Cells(lastrow, Col))
    With repRange
        .Replace What:="*Аукцион*", Replacement:="ЭА"
        .Replace What:="*Двухэтап*конкурс*элек*фор*", Replacement:="ДКЭФ"
        .Replace What:="*Двухэтап*конкурс*", Replacement:="ДК"
        .Replace What:="*Открытый конкурс*элек*фор*", Replacement:="ОКЭФ"
        .Replace What:="*Открытый конкурс*", Replacement:="ОК"
        .Replace What:="*Запрос котировок*элек*фор*", Replacement:="ЗКЭФ"
        .Replace What:="*Запрос котировок*", Replacement:="ЗК"
        .Replace What:="*Запрос предложений*электр*форм*", Replacement:="ЗПЭФ"
        .Replace What:="*Запрос предложений*", Replacement:="ЗП"
    End With
End Sub
Private Sub ploschadka()
    Dim repRange    As Range, lastrow As Integer, Col As Integer
    lastrow = inputList.Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    Col = inputList.Range("Ploschadka").Column
    Set repRange = inputList.Range(Cells(2, Col), Cells(lastrow, Col))
    With repRange
        .Replace What:="*РТС*", Replacement:="РТС-тендер"
        .Replace What:="*Сбербанк*", Replacement:="Сбер"
        .Replace What:="*ЕЭТП*", Replacement:="Росельторг"
        .Replace What:="*ММВБ*", Replacement:="ЭТП НЭП"
        .Replace What:="*Фабрикант*", Replacement:="Фабрикант"
        .Replace What:="*АГЗ РТ*", Replacement:="ЗаказРФ"
        .Replace What:="*Национальная электронная площадка*", Replacement:="ЭТП НЭП"
        .Replace What:="*ТЭК-Торг*", Replacement:="ТЭК-Торг"
        .Replace What:="*РАД*", Replacement:="РАД"
        .Replace What:="*Газпромбанк*", Replacement:="ЭТП ГПБ"
    End With
End Sub
Private Sub organizator()
    Dim repRange    As Range, lastrow As Integer, Col As Integer
    lastrow = inputList.Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    Col = inputList.Range("Organizator").Column
    Set repRange = inputList.Range(Cells(2, Col), Cells(lastrow, Col))
    With repRange
        .Replace What:="*фонд*соц*страх*", Replacement:="ФСС"
        .Replace What:="*соц*защ*", Replacement:="СЗ"
        .Replace What:="*соц*опек*", Replacement:="СЗ"
        .Replace What:="*соц*разв*", Replacement:="СЗ"
        .Replace What:="*соц*обслуж*", Replacement:="СЗ"
        .Replace What:="*здравоохран*", Replacement:="ЗДР"
        .Replace What:="*гос*заказ*", Replacement:=""
        .Replace What:="*закуп*", Replacement:=""
        .Replace What:="*имущ*зем*", Replacement:=""
    End With
End Sub
Private Function tip_zakazchik(s As String) As String
    tip_zakazchik = UCase(s)
    If tip_zakazchik Like "*ФОНД*СОЦ*СТРАХ*" Then tip_zakazchik = "ФСС": Exit Function
    If tip_zakazchik Like "*СОЦ*ЗАЩ*" Then tip_zakazchik = "СЗ": Exit Function
    If tip_zakazchik Like "*СОЦ*ОПЕК*" Then tip_zakazchik = "СЗ": Exit Function
    If tip_zakazchik Like "*СОЦ*РАЗВ*" Then tip_zakazchik = "СЗ": Exit Function
    If tip_zakazchik Like "*ЗДРАВООХРАН*" Then tip_zakazchik = "ЗДР": Exit Function
End Function
Private Function purchase_subject_abbr(s As String) As String
    purchase_subject_abbr = UCase(s)
    If purchase_subject_abbr Like "*ПОДГУЗ*" Then purchase_subject_abbr = "АБС": Exit Function
    If purchase_subject_abbr Like "*ПЕЛЕН*" Then purchase_subject_abbr = "АБС": Exit Function
    If purchase_subject_abbr Like "*АБСОРБ*" Then purchase_subject_abbr = "АБС": Exit Function
    If purchase_subject_abbr Like "*ПАМПЕРС*" Then purchase_subject_abbr = "АБС": Exit Function
    If purchase_subject_abbr Like "*ПРОКЛАД*" Then purchase_subject_abbr = "АБС": Exit Function
    If purchase_subject_abbr Like "*ВКЛАДЫШ*" Then purchase_subject_abbr = "АБС": Exit Function
    If purchase_subject_abbr Like "*КРЕС*КОЛЯС*" Then purchase_subject_abbr = "ИКК": Exit Function
    If purchase_subject_abbr Like "*КОЛЯС*" Then purchase_subject_abbr = "ИКК": Exit Function
    If purchase_subject_abbr Like "*КРЕС*СТУЛ*" Then purchase_subject_abbr = "ИКК": Exit Function
    If purchase_subject_abbr Like "*ДЦП*" Then purchase_subject_abbr = "ИКК": Exit Function
    If purchase_subject_abbr Like "*КАЛОПРИЕМ*" Then purchase_subject_abbr = "ССВ": Exit Function
    If purchase_subject_abbr Like "*УРОПР*" Then purchase_subject_abbr = "ССВ": Exit Function
    If purchase_subject_abbr Like "*МОЧЕПРИЕМ*" Then purchase_subject_abbr = "ССВ": Exit Function
    If purchase_subject_abbr Like "*СРЕДСТ*ФУНКЦ*ВЫД*" Then purchase_subject_abbr = "ССВ": Exit Function
    If purchase_subject_abbr Like "*СРЕДСТВ*УХОД*" Then purchase_subject_abbr = "ССВ": Exit Function
    If purchase_subject_abbr Like "*КАТЕТЕР*" Then purchase_subject_abbr = "ССВ": Exit Function
    If purchase_subject_abbr Like "*ПРОТЕЗ*" Then purchase_subject_abbr = "ПОИ": Exit Function
    If purchase_subject_abbr Like "*ОРТЕЗ*" Then purchase_subject_abbr = "ПОИ": Exit Function
    If purchase_subject_abbr Like "*ОБУВ*" Then purchase_subject_abbr = "ПОИ": Exit Function
    If purchase_subject_abbr Like "*БАНДАЖ*" Then purchase_subject_abbr = "ПОИ": Exit Function
    If purchase_subject_abbr Like "*КОРСЕТ*" Then purchase_subject_abbr = "ПОИ": Exit Function
    If purchase_subject_abbr Like "*ХОДУНК*" Then purchase_subject_abbr = "ПОИ": Exit Function
    If purchase_subject_abbr Like "*ТУТОР*" Then purchase_subject_abbr = "ПОИ": Exit Function
    If purchase_subject_abbr Like "*ТЕЛЕВИЗ*" Then purchase_subject_abbr = "др": Exit Function
    If purchase_subject_abbr Like "*ТЕЛЕФОН*" Then purchase_subject_abbr = "др": Exit Function
    If purchase_subject_abbr Like "*тонометр*" Then purchase_subject_abbr = "др": Exit Function
End Function
Private Sub preference()
    Dim repRange    As Range, lastrow As Integer, Col As Integer, i As Integer, s As String
    lastrow = inputList.Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    Col = inputList.Range("SMP").Column
    For i = 2 To lastrow
        If UCase(inputList.Range("Pref")(i)) Like UCase("*орг*инвалид*") Then s = s + ", ОИ"
        If UCase(inputList.Range("SMP")(i)) Like UCase("*субъект*мал*предпр*") Or UCase(inputList.Range("SMP")(i)) Like UCase("*ч. 3 ст. 30*") Then s = s + ", СМП"
        If UCase(inputList.Range("Pref")(i)) Like UCase("*прик*мин*126*") Then s = s + ", ПТРБ"
        inputList.Range("SMP")(i).Value = Replace(s, ", ", "", 1, 1)
        s = ""
    Next i
    Set repRange = inputList.Range(Cells(2, Col), Cells(lastrow, Col))
    With repRange
        .Replace What:="*прик*мин*126*", Replacement:="ПТРБ"
        .Replace What:="*субъект*мал*предпр*", Replacement:="СМП"
        .Replace What:="*орг*инвалид*", Replacement:="ОИ"
        .Replace What:="*исправ*сист*", Replacement:="ЗК"
        .Replace What:="*треб*", Replacement:=""
        .Replace What:="*не установ*", Replacement:=""
    End With
End Sub
Private Sub sum_of_preference()
    Dim lastrow     As Integer, Col As Integer, i As Integer, m As String, pref As String, k As Integer
    lastrow = inputList.Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    Col = inputList.Range("Pref").Column
    For i = 2 To lastrow
        m = InStr(1, Range("Pref")(i), "%")
        If m > 0 Then
            pref = Mid(Range("Pref")(i), m - 3, 4)
            Range("Pref")(i).Value = Trim(pref)
        Else
            Range("Pref")(i).Value = ""
        End If
    Next i
End Sub
Private Sub Scrap2020()
    Dim http        As New MSXML2.XMLHTTP60
    Dim doc         As New HTMLDocument, splitter() As String, i As Long, b, lastrow, pos As Integer, NMCK, currentNMCK, obespechK, obespechZ As Double, zakIsSet, aucTimeIsSet As Boolean
    lastrow = inputList.Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    For v = 2 To lastrow
        http.Open "GET", ZAK_COMMON & Replace(Range("Nomer")(v).Value, "№", ""), FALSE
        http.send
        doc.body.innerHTML = http.responseText
        aucTimeIsSet = FALSE
        b = 0
        zakIsSet = FALSE
        NMCK = 0
        currentNMCK = 0
        obespechK = 0
        obespechZ = 0
        Dim td, tdPlusOne As String
        For i = 0 To doc.getElementsByTagName("span").Length - 1
            td = Trim(doc.getElementsByTagName("span")(i).innerText)
            tdPlusOne = Trim(doc.getElementsByTagName("span")(i + 1).innerText)
            If UCase(td) Like UCase("Способ*опред*поставщ*подрядч*исполн*)") Then inputList.Range("Sposob")(v).Value = tdPlusOne
            If UCase(td) Like UCase("Наим*электр*площ*") Then inputList.Range("Ploschadka")(v).Value = tdPlusOne
            If UCase(td) Like UCase("Наим*объект*зак*") And zakIsSet = FALSE Then
                inputList.Range("Predmet")(v).Value = tdPlusOne
                inputList.Range("Tip")(v).Value = tdPlusOne
                zakIsSet = TRUE
            End If
            If UCase(td) Like UCase("орган*осущ*размещ*") Or UCase(td) Like UCase("размещен*осущ*") Then
                inputList.Range("Organizator")(v).Value = tdPlusOne
            End If
            If UCase(td) Like UCase("Почт*адрес*") Then inputList.Range("Region")(v).Value = tdPlusOne
            If UCase(td) Like UCase("дата*окон*подач*") Or UCase(td) Like UCase("дата*окон*прием*") Then
                splitter = Split(tdPlusOne, " ")
                On Error Resume Next
                If UBound(splitter) = 1 Then
                    inputList.Range("DataOkonch")(v).Value = DateValue(splitter(0))
                    inputList.Range("TimeOkonch")(v).Value = CDate(tdPlusOne)
                ElseIf UBound(splitter) >= 2 Then
                    inputList.Range("DataOkonch")(v).Value = DateValue(splitter(0))
                    inputList.Range("TimeOkonch")(v).Value = CDate(splitter(0) + " " + splitter(1))
                End If
            End If
            If UCase(td) Like UCase("дата*окон*рассм*") Then
                If IsDate(tdPlusOne) Then inputList.Range("OkonchRasm")(v).Value = DateValue(tdPlusOne)
            End If
            If UCase(td) Like UCase("дата*время*рассм*первы*заявок*") Then
                splitter = Split(tdPlusOne, " ")
                On Error Resume Next
                inputList.Range("OkonchRasm")(v).Value = DateValue(splitter(0))
            End If
            If UCase(td) Like UCase("дата*аукц*") Then
                If IsDate(tdPlusOne) Then
                    inputList.Range("DataProved")(v).Value = tdPlusOne
                End If
            End If
            If UCase(td) Like UCase("время*аукц*") And Not aucTimeIsSet Then
                inputList.Range("TimeProved")(v).Value = tdPlusOne
                aucTimeIsSet = TRUE
            End If
            If UCase(td) Like UCase("Преимущества*") Then
                inputList.Range("Pref")(v).Value = tdPlusOne
            End If
            If UCase(td) Like UCase("нач*цена*контракт*") Or UCase(td) Like UCase("макс*знач*цены*контракта") Then
                currentNMCK = Val(Replace(tdPlusOne, ",", "."))
                If NMCK = 0 Then
                    NMCK = Val(Replace(tdPlusOne, ",", "."))
                    inputList.Range("NMCK")(v).Value = NMCK
                End If
            End If
            If UCase(td) Like UCase("Ограничения и запреты") Then inputList.Range("SMP")(v).Value = tdPlusOne
            If UCase(td) Like UCase("Размер обеспечения заяв*") Then
                obespechZ = obespechZ + Val(Replace(tdPlusOne, ",", "."))
            End If
            If UCase(td) Like UCase("Размер обеспечения исполнения контракта") Then
                pos = InStr(1, tdPlusOne, "(", vbTextCompare)
                If pos <> 0 Then
                    obespechK = obespechK + Val(Replace(tdPlusOne, ",", "."))
                Else
                    pos = InStr(1, tdPlusOne, "%", vbTextCompare)
                    If pos <> 0 Then
                        obespechK = obespechK + Application.WorksheetFunction.Floor(currentNMCK * (Val(Replace(tdPlusOne, ",", ".")) / 100), 0.01)
                    End If
                End If
            End If
            If UCase(td) Like UCase("Срок*остав*товар*заверш*работ*граф*оказ*услу*") Then
                inputList.Range("Srok")(v).Value = tdPlusOne
            End If
        Next i
        inputList.Range("ObespechISP")(v).Value = obespechK
        inputList.Range("ObespechZayav")(v).Value = obespechZ
    Next v
    Set http = Nothing
    Set doc = Nothing
End Sub
Private Sub Scrap2022()
    Dim http        As New MSXML2.XMLHTTP60
    Dim elements    As IHTMLDOMChildrenCollection, search_block As IHTMLDOMChildrenCollection, pref_block As IHTMLDOMChildrenCollection
    Dim doc         As New HTMLDocument
    Dim block_title As String, finded_info As String, pref_title As String
    Dim i           As Long, k As Integer, lastrow As Integer, pos As Integer
    Dim NMCK        As Double, avans As Double, obespechK As Double, obespechZ As Double, x As Double
    Dim nmck_is_set As Boolean, term_is_set As Boolean
    Dim prefs       As IHTMLElement
    
    lastrow = inputList.Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    
    For v = 2 To lastrow
        
        Set doc = HTMLDoc(ZAK_COMMON & Replace(Range("Nomer")(v).Value, "№", ""))
        
        Set elements = doc.querySelectorAll("div.row.blockInfo")
        
        NMCK = 0
        nmck_is_set = FALSE
        obespechK = 0
        obespechZ = 0
        
        For i = 0 To elements.Length - 1
            
            block_title = UCase(elements.Item(i).querySelector(".blockInfo__title").innerText)
            
            If block_title Like "*ОБЩАЯ*ИНФ*ЗАКУПКЕ*" Then
                
                Set search_block = elements.Item(i).querySelectorAll("section")
                
                inputList.Range("Sposob")(v).Value = element_info("*СПОСОБ*ОПРЕД*ПОСТАВ*ПОДРЯД*ИСПОЛН*", search_block)
                inputList.Range("Ploschadka")(v).Value = element_info("*НАИМ*ЭЛЕКТР*ПЛОЩ*", search_block)
                finded_info = element_info("*НАИМ*ОБЪЕКТ*ЗАК*", search_block)
                inputList.Range("Predmet")(v).Value = finded_info
                inputList.Range("Tip")(v).Value = finded_info
                
            End If
            
            If block_title Like "*КОНТАКТН*ИНФОРМАЦИЯ*" Then
                
                Set search_block = elements.Item(i).querySelectorAll("section")
                
                inputList.Range("Organizator")(v).Value = element_info("*ОРГАН*ОСУЩ*РАЗМЕЩ*", search_block)
                inputList.Range("Region")(v).Value = RegionCustomerStr(inputList.Range("Organizator")(v).Value)
                If inputList.Range("Region")(v).Value = "" Then
                    inputList.Range("Region")(v).Value = RegionStr(element_info("*ПОЧТ*АДРЕС*", search_block))
                End If
                
            End If
            
            If block_title Like "*ИНФОРМАЦ*ПРОЦЕДУР*ЗАКУПК*" Then
                
                Set search_block = elements.Item(i).querySelectorAll("section")
                
                finded_info = Left(element_info("*Дата*время*оконч*подач*заявок*", search_block), 16)
                inputList.Range("DataOkonch")(v).Value = CDate(finded_info)
                inputList.Range("TimeOkonch")(v).Value = CDate(finded_info)
                inputList.Range("DataProved")(v).Value = CDate(finded_info)
                inputList.Range("TimeProved")(v).Value = DateAdd("h", 2, CDate(finded_info))
                
                finded_info = element_info("*Дата*подвед*итог*", search_block)
                inputList.Range("OkonchRasm")(v).Value = CDate(finded_info)
            End If
            
            If block_title Like "*НАЧАЛЬН*МАКСИМАЛЬН*КОНТРАКТ*" Then
                
                Set search_block = elements.Item(i).querySelectorAll("section")
                
                finded_info = element_info("*ЦЕН*КОНТРАКТ*", search_block)
                
                NMCK = Val(Replace(finded_info, ",", "."))
                If nmck_is_set = FALSE Then inputList.Range("NMCK")(v).Value = NMCK
                
                finded_info = element_info("*РАЗМЕР*АВАНС*", search_block)
                
                avans = Val(Replace(finded_info, ",", "."))
                If avans > 0 Then
                    inputList.Range("Srok")(v) = inputList.Range("Srok")(v) & "Аванс: " & CStr(avans) & " %" & vbCrLf
                End If
            End If
            
            If block_title Like "*ПРЕИМУЩЕСТВ*ТРЕБОВАН*УЧАСТНИК*" Then
                
                Set pref_block = elements.Item(i).querySelectorAll("section.blockInfo__section")
                
                For k = 0 To pref_block.Length - 1
                    
                    pref_title = UCase(pref_block.Item(k).querySelector(".section__title").innerText)
                    If pref_title Like "*ПРЕИМУЩЕСТВА*" Then
                        inputList.Range("SMP")(v).Value = pref_block.Item(k).innerText
                        inputList.Range("Pref")(v).Value = pref_block.Item(k).innerText
                    End If
                    
                    If pref_title Like "*ОГРАНИЧЕН*ЗАПРЕТ*" Then
                        inputList.Range("SMP")(v).Value = inputList.Range("SMP")(v).Value + pref_block.Item(k).innerText
                    End If
                    
                Next k
                
            End If
            
            If block_title Like "*УСЛОВИЯ*КОНТРАКТ*" And term_is_set = FALSE Then
                
                Set search_block = elements.Item(i).querySelectorAll("section")
                
                inputList.Range("Srok")(v) = inputList.Range("Srok")(v) + element_info("*СРОК*ИСПОЛНЕН*КОНТРАКТ*", search_block)
                
            End If
            
            If block_title Like "*ОБЕСПЕЧЕНИЕ*ЗАЯВ*" Then
                
                Set search_block = elements.Item(i).querySelectorAll("section")
                
                finded_info = element_info("*РАЗМЕР*ОБЕСПЕЧЕН*ЗАЯВК*", search_block)
                obespechZ = obespechZ + Val(Replace(finded_info, ",", "."))
                
            End If
            
            If block_title Like "*ОБЕСПЕЧЕН*ИСПОЛНЕН*КОНТРАКТ*" Then
                
                Set search_block = elements.Item(i).querySelectorAll("section")
                
                finded_info = element_info("*РАЗМЕР*ОБЕСПЕЧЕН*ИСПОЛН*КОНТРАКТ*", search_block)
                
                pos = InStr(1, finded_info, "(", vbTextCompare)
                If pos <> 0 Then
                    obespechK = obespechK + Val(Replace(finded_info, ",", "."))
                Else
                    pos = InStr(1, finded_info, "%", vbTextCompare)
                    If pos <> 0 Then
                        x = CDbl(Replace(finded_info, " %", ""))
                        obespechK = obespechK + Application.WorksheetFunction.Floor(NMCK * (x / 100), 0.01)
                    End If
                End If
                
            End If
            
        Next i
        
        inputList.Range("ObespechISP")(v).Value = obespechK
        inputList.Range("ObespechZayav")(v).Value = obespechZ
        
    Next v
    
End Sub
Sub Заполнить_Реестр_Контрактов()
    Dim http        As New MSXML2.XMLHTTP60
    Dim elements    As IHTMLDOMChildrenCollection
    Dim doc         As New HTMLDocument, doc2 As New HTMLDocument
    Dim a           As IHTMLElement, section_title As IHTMLElement, section_info As IHTMLElement
    Dim i           As Long, lastrow As Integer, pos As Integer, result_total As Integer, k As Long
    Dim number      As String, region As String, customer_type As String, contract_price As String, contract_abbr As String
    Dim posName     As Integer, posLineBreak As Integer, link As String, Text As String, customer_name As String
    Dim contract_supplier As String, contract_sign As String, contract_end_date As String, contract_number As String, contract_number_registry As String
    
    Call toggle_screen_upd
    
    contract_supplier = "Держатель_контракта"
    number = "Номер"
    region = "Регион"
    customer_type = "Заказчик"
    contract_abbr = "Предмет_контракта"
    contract_number = "Номер_ГК"
    contract_price = "Сумма_контракта"
    contract_sign = "Дата_заключения"
    contract_end_date = "Срок_действия_контракта"
    contract_number_registry = "Номер_ГК_реестровый"
    
    lastrow = Range(number).Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    
    For i = 2 To lastrow
        
        If IsEmpty(Range(contract_number)(i)) And Not Range(number)(i) = "прямой" Then
            
            Set doc = HTMLDoc("https://zakupki.gov.ru/epz/contract/search/results.html?orderNumber=" & Replace(Range(number)(i).Value, "№", ""))
            
            result_total = Val(Trim(doc.querySelector("div.search-results__total").innerText))
            If result_total = 1 Then
                Set a = doc.querySelector("div.registry-entry__header-mid__number > a")
                link = Replace(a.href, "about:", "https://zakupki.gov.ru")
                Range(contract_number_registry)(i).Value = Trim(Replace(a.innerText, "№", ""))
                Set doc = HTMLDoc(link)
                Set elements = doc.querySelectorAll("div.blockInfo__section > section.blockInfo__section.section")
                
                For k = 0 To elements.Length - 1
                    Set section_title = elements.Item(k).querySelector(".section__title")
                    Set section_info = elements.Item(k).querySelector(".section__info")
                    Text = section_title.innerText
                    
                    If section_title.innerText = "Полное наименование заказчика" Then
                        Range(customer_type)(i) = tip_zakazchik(section_info.innerText)
                        Set a = section_info.querySelector("a")
                        link = Replace(a.href, "about:", "https://zakupki.gov.ru")
                        Set doc2 = HTMLDoc(link)
                        Range(region)(i) = RegionStr(doc2.querySelector(".registry-entry__body-value").innerText)
                        Set doc2 = Nothing
                    End If
                    
                    If section_title.innerText = "Дата заключения контракта" Then
                        Range(contract_sign)(i) = CDate(Trim(section_info.innerText))
                    End If
                    
                    If section_title.innerText = "Номер контракта" Then
                        Range(contract_number)(i) = "№ " + section_info.innerText
                    End If
                    
                    If section_title.innerText = "Предмет контракта" Then
                        Range(contract_abbr)(i) = purchase_subject_abbr(section_info.innerText)
                    End If
                    
                    If section_title.innerText = "Цена контракта" Or section_title.innerText = "Максимальное значение цены контракта" Then
                        Range(contract_price)(i) = Val(Replace(section_info.innerText, ",", "."))
                    End If
                    
                    If section_title.innerText = "Дата окончания исполнения контракта" Then
                        Range(contract_end_date)(i) = CDate(Left(Trim(section_info.innerText), 10))
                    End If
                Next k
                
                Text = doc.querySelector("div.participantsInnerHtml td.tableBlock__col").innerText
                
                If IsEmpty(Range(contract_supplier)(i)) Then
                    posName = InStr(text, "(ООО")
                    posLineBreak = InStr(text, vbCr)
                    
                    If posLineBreak = 0 Then
                        posLineBreak = InStr(text, "Код по ОКПО")
                    End If
                    If posName > 0 Then
                        Range(contract_supplier)(i).Value = Trim(Mid(text, posName + 1, posLineBreak - posName - 3))
                    Else
                        Range(contract_supplier)(i).Value = Trim(Left(text, posLineBreak - 1))
                    End If
                End If
                
            ElseIf result_total > 1 Then
                Range(contract_number)(i).Value = "Больше одной записи"
                Range(contract_number_registry)(i).Value = "Больше одной записи"
            End If
            
        End If
        
    Next i
    
    Call toggle_screen_upd
    
End Sub
Private Function element_info(search_str As String, elems As IHTMLDOMChildrenCollection) As String
    Dim i           As Integer, section_title As String
    
    search_str = UCase(search_str)
    
    For i = 0 To elems.Length - 1
        
        section_title = UCase(elems.Item(i).querySelector(".section__title").innerText)
        
        If section_title Like search_str Then
            element_info = Trim(elems.Item(i).querySelector(".section__info").innerText)
            Exit Function
        End If
        
    Next i
    
End Function
Private Sub NewList23()
    Worksheets.Add Before:=Sheets(1)
    With Sheets(1)
        .Range("A:A").name = "Date": .Range("A:A").ColumnWidth = 9.3: .Range("A:A").NumberFormat = "m/d/yyyy"
        .Range("B:B").name = "Tip": .Range("B:B").ColumnWidth = 5.2: .Range("B:B").NumberFormat = "@"
        .Range("C:C").name = "Predmet": .Range("C:C").ColumnWidth = 28: .Range("C:C").NumberFormat = "@"
        .Range("D:D").name = "SMP": .Range("D:D").ColumnWidth = 7: .Range("D:D").NumberFormat = "@"
        .Range("E:E").name = "Pref": .Range("E:E").ColumnWidth = 5: .Range("E:E").NumberFormat = "0%"
        .Range("F:F").name = "DataOkonch": .Range("F:F").ColumnWidth = 9.3: .Range("F:F").NumberFormat = "m/d/yyyy"
        .Range("G:G").name = "TimeOkonch": .Range("G:G").ColumnWidth = 6.7: .Range("G:G").NumberFormat = "h:mm;@"
        .Range("H:H").name = "OkonchRasm": .Range("H:H").ColumnWidth = 10: .Range("H:H").NumberFormat = "m/d/yyyy"
        .Range("I:I").name = "DataProved": .Range("I:I").ColumnWidth = 11: .Range("I:I").NumberFormat = "m/d/yyyy"
        .Range("J:J").name = "TimeProved": .Range("J:J").ColumnWidth = 7.9: .Range("J:J").NumberFormat = "[$-F400]h:mm:ss AM/PM"
        .Range("K:K").name = "Region": .Range("K:K").ColumnWidth = 18.8: .Range("K:K").NumberFormat = "@"
        .Range("L:L").name = "Organizator": .Range("L:L").ColumnWidth = 7: .Range("L:L").NumberFormat = "@"
        .Range("M:M").name = "NMCK": .Range("M:M").ColumnWidth = 12: .Range("M:M").NumberFormat = "#,##0.00"
        .Range("N:N").name = "ObespechZayav": .Range("N:N").ColumnWidth = 10: .Range("N:N").NumberFormat = "#,##0.00"
        .Range("O:O").name = "ObespechISP": .Range("O:O").ColumnWidth = 12: .Range("O:O").NumberFormat = "#,##0.00"
        .Range("Q:Q").name = "Srok": .Range("Q:Q").ColumnWidth = 28: .Range("Q:Q").NumberFormat = "@"
        .Range("V:V").name = "Sposob": .Range("V:V").ColumnWidth = 9: .Range("V:V").NumberFormat = "@"
        .Range("W:W").name = "Nomer": .Range("W:W").ColumnWidth = 21.4: .Range("W:W").NumberFormat = "@"
        .Range("X:X").name = "Ploschadka": .Range("X:X").ColumnWidth = 11: .Range("X:X").NumberFormat = "@"
        .Range("Y:Y").name = "Ssylka": .Range("Y:Y").ColumnWidth = 15
        .Range("Z:Z").name = "UTC": .Range("Z:Z").ColumnWidth = 15: .Range("Z:Z").NumberFormat = "0"
        .Cells(1, 1).Value = "Дата": .Cells(1, 2).Value = "Тип": .Cells(1, 3).Value = "Предмет закупки"
        .Cells(1, 4).Value = "СМП/ПТРБ": .Cells(1, 6).Value = "Дата оконч. приема"
        .Cells(1, 7).Value = "Время оконч. приема": .Cells(1, 8).Value = "Окончание рассм.": .Cells(1, 9).Value = "Дата проведения"
        .Cells(1, 10).Value = "Время проведения": .Cells(1, 11).Value = "Регион": .Cells(1, 12).Value = "Организатор"
        .Cells(1, 13).Value = "Начальная цена": .Cells(1, 14).Value = "Обеспечение заявки": .Cells(1, 15).Value = "Обеспечение контракта"
        .Cells(1, 17).Value = "Срок поставки": .Cells(1, 22).Value = "Способ закупки": .Cells(1, 23).Value = "Реестровый номер"
        .Cells(1, 24).Value = "Площадка": .Cells(1, 25).Value = "Ссылка": .Cells(1, 5).Value = "Преф.": .Cells(1, 26).Value = "UTC"
        .Range("R:U").ColumnWidth = 1: .Range("P:P").ColumnWidth = 1
        .Range("R:U").EntireColumn.Hidden = True: .Range("P:P").EntireColumn.Hidden = TRUE
    End With
    
    With Sheets(1).Range(Cells(1, 1), Cells(1, 27))
        .WrapText = TRUE
        .AutoFilter
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 43
        .Font.name = "Cambria"
        .Font.Size = 10
        .Font.Bold = TRUE
    End With
End Sub
Private Sub Formatirovan1()
    Dim format_area As Range, lastrow As Integer, lastCol As Integer
    lastrow = Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    lastCol = Cells.Find(What:="*", SearchOrder:=xlColumns, SearchDirection:=xlPrevious, LookIn:=xlValues).Column
    Set format_area = Range(Cells(2, 1), Cells(lastrow, lastCol))
    format_area.Replace What:="" & Chr(10) & "", Replacement:=" ", SearchOrder:=xlByColumns
    With format_area
        .WrapText = TRUE
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .RowHeight = 30
        .Font.name = "Cambria"
        .Font.Size = 9
        '.Value = Application.Trim(.Value)
    End With
    With Range(Cells(1, 1), Cells(lastrow, lastCol))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    End With
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = TRUE
End Sub
Private Sub tip()
    Dim repRange    As Range, lastrow As Integer, Col As Integer
    lastrow = Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    Col = Range("Tip").Column
    Set repRange = Range(Cells(2, Col), Cells(lastrow, Col))
    With repRange
        .Replace What:="*подгуз*", Replacement:="АБС"
        .Replace What:="*пелен*", Replacement:="АБС"
        .Replace What:="*абсорб*", Replacement:="АБС"
        .Replace What:="*памперс*", Replacement:="АБС"
        .Replace What:="*проклад*", Replacement:="АБС"
        .Replace What:="*вкладыш*", Replacement:="АБС"
        .Replace What:="*крес*коляс*", Replacement:="ИКК"
        .Replace What:="*коляс*", Replacement:="ИКК"
        .Replace What:="*крес*стул*", Replacement:="ИКК"
        .Replace What:="*ДЦП*", Replacement:="ИКК"
        .Replace What:="*калоприем*", Replacement:="ССВ"
        .Replace What:="*уропр*", Replacement:="ССВ"
        .Replace What:="*мочеприем*", Replacement:="ССВ"
        .Replace What:="*средст*функц*выд*", Replacement:="ССВ"
        .Replace What:="*средств*уход*", Replacement:="ССВ"
        .Replace What:="*катетер*", Replacement:="ССВ"
        .Replace What:="*протез*", Replacement:="ПОИ"
        .Replace What:="*ортез*", Replacement:="ПОИ"
        .Replace What:="*обув*", Replacement:="ПОИ"
        .Replace What:="*бандаж*", Replacement:="ПОИ"
        .Replace What:="*корсет*", Replacement:="ПОИ"
        .Replace What:="*ходунк*", Replacement:="ПОИ"
        .Replace What:="*тутор*", Replacement:="ПОИ"
        .Replace What:="*телевиз*", Replacement:="др"
        .Replace What:="*телефон*", Replacement:="др"
        .Replace What:="*тонометр*", Replacement:="др"
    End With
End Sub
Private Sub predmet()
    Dim repRange    As Range, lastrow As Integer, Col As Integer
    lastrow = Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    Col = Range("Predmet").Column
    Set repRange = Range(Cells(2, Col), Cells(lastrow, Col))
    With repRange
        .Replace What:="*телевиз*", Replacement:="телевизоров с телетекстом"
        .Replace What:="*телефон*", Replacement:="телефонных устройств с текстовым выходом"
        .Replace What:="*тонометр*", Replacement:="тонометров с речевым выходом"
        .Replace What:="*крес*коляс* с ручным приводом*", Replacement:="кресел-колясок с ручным приводом"
        .Replace What:="*крес*коляс* с электро*", Replacement:="кресел-колясок с электроприводом"
        .Replace What:="*крес*коляс* различ*", Replacement:="кресел-колясок различной модификации"
        .Replace What:="*крес*стул* санитар*", Replacement:="кресел-стульев с санитарным оснащением"
    End With
End Sub
Private Sub region()
    Dim repRange    As Range, lastrow As Integer, Col As Integer, i As Integer
    lastrow = Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    Col = Range("Region").Column
    Set repRange = Range(Cells(2, Col), Cells(lastrow, Col))
    With repRange
        .Replace What:="*196191*Санкт-Петербург*Ленинский проспект*д.168*", Replacement:="Ленинградская область", MatchCase:=True
        .Replace What:="*г.*Москва*ул.*3-я*Хорошевская*д.12*", Replacement:="Московская область", MatchCase:=True
        .Replace What:="*Карел*Респ*", Replacement:="Республика Карелия", MatchCase:=True
        .Replace What:="*Петрозаводск,*", Replacement:="Республика Карелия", MatchCase:=True
        .Replace What:="*Коми*Респ*", Replacement:="Республика Коми", MatchCase:=True
        .Replace What:="*Сыктывкар,*", Replacement:="Республика Коми", MatchCase:=True
        .Replace What:="*Чечен*Респ*", Replacement:="Чеченская Республика", MatchCase:=True
        .Replace What:="*Грозный,*", Replacement:="Чеченская Республика", MatchCase:=True
        .Replace What:="*Чуваш*Респ*", Replacement:="Республика Чувашия", MatchCase:=True
        .Replace What:="*Элиста,*", Replacement:="Республика Чувашия", MatchCase:=True
        .Replace What:="*Чукотский*", Replacement:="Чукотский АО", MatchCase:=True
        .Replace What:="*Аныдырь,*", Replacement:="Чукотский АО", MatchCase:=True
        .Replace What:="*Удмурт*Респ*", Replacement:="Удмуртская Республика", MatchCase:=True
        .Replace What:="*Ижевск,*", Replacement:="Удмуртская Республика", MatchCase:=True
        .Replace What:="*Ингуш*Респ*", Replacement:="Республика Ингушетия", MatchCase:=True
        .Replace What:="*Кемеров*обл*", Replacement:="Кемеровская область", MatchCase:=True
        .Replace What:="*Кемерово,*", Replacement:="Кемеровская область", MatchCase:=True
        .Replace What:="*Дагест*Респ*", Replacement:="Республика Дагестан", MatchCase:=True
        .Replace What:="*Махачкала,*", Replacement:="Республика Дагестан", MatchCase:=True
        .Replace What:="*Крым*Респ*", Replacement:="Республика Крым", MatchCase:=True
        .Replace What:="*Симферополь,*", Replacement:="Республика Крым", MatchCase:=True
        .Replace What:="*Саха*Якути*", Replacement:="Республика Саха (Якутия)", MatchCase:=True
        .Replace What:="*Якутск,*", Replacement:="Республика Саха (Якутия)", MatchCase:=True
        .Replace What:="*Хакас*", Replacement:="Республика Хакасия", MatchCase:=True
        .Replace What:="*Абакан,*", Replacement:="Республика Хакасия", MatchCase:=True
        .Replace What:="*Ханты*Ман*", Replacement:="Ханты-Мансийский АО — Югра", MatchCase:=True
        .Replace What:="*Ханты-Мансийск,*", Replacement:="Ханты-Мансийский АО — Югра", MatchCase:=True
        .Replace What:="*Башкорт*Респ*", Replacement:="Республика Башкортостан", MatchCase:=True
        .Replace What:="*Уфа,*", Replacement:="Республика Башкортостан", MatchCase:=True
        .Replace What:="*Санкт*Пет*", Replacement:="Санкт-Петербург", MatchCase:=True
        .Replace What:="*Санкт-Петербург,*", Replacement:="Санкт-Петербург", MatchCase:=True
        .Replace What:="*Ямало*Нен*", Replacement:="Ямало-Ненецкий АО", MatchCase:=True
        .Replace What:="*Салехард,*", Replacement:="Ямало-Ненецкий АО", MatchCase:=True
        .Replace What:="*Татарстан*", Replacement:="Республика Татарстан", MatchCase:=True
        .Replace What:="*Казань,*", Replacement:="Республика Татарстан", MatchCase:=True
        .Replace What:="*Краснодар*край*", Replacement:="Краснодарский край", MatchCase:=True
        .Replace What:="*Краснодар,*", Replacement:="Краснодарский край", MatchCase:=True
        .Replace What:="*Челябинск*обл*", Replacement:="Челябинская область", MatchCase:=True
        .Replace What:="*Челябинск,*", Replacement:="Челябинская область", MatchCase:=True
        .Replace What:="*Осети*Респ*", Replacement:="Республика Северная Осетия - Алания", MatchCase:=True
        .Replace What:="*Владикавказ,*", Replacement:="Республика Северная Осетия - Алания", MatchCase:=True
        .Replace What:="*Ставропол*край*", Replacement:="Ставропольский край", MatchCase:=True
        .Replace What:="*Ставрополь,*", Replacement:="Ставропольский край", MatchCase:=True
        .Replace What:="*Кабард*Респ*", Replacement:="Республика Кабардино-Балкарская", MatchCase:=True
        .Replace What:="*Нальчик,*", Replacement:="Республика Кабардино-Балкарская", MatchCase:=True
        .Replace What:="*Астрахан*обл*", Replacement:="Астраханская область", MatchCase:=True
        .Replace What:="*Астрахань,*", Replacement:="Астраханская область", MatchCase:=True
        .Replace What:="*Адыг*Респ*", Replacement:="Республика Адыгея", MatchCase:=True
        .Replace What:="*Майкоп,*", Replacement:="Республика Адыгея", MatchCase:=True
        .Replace What:="*Карач*Респ*", Replacement:="Республика Карачаево-Черкесская", MatchCase:=True
        .Replace What:="*Черкесск,*", Replacement:="Республика Карачаево-Черкесская", MatchCase:=True
        .Replace What:="*Москв*", Replacement:="Москва", MatchCase:=True
        .Replace What:="*Москва,*", Replacement:="Москва", MatchCase:=True
        .Replace What:="*Ростов*обл*", Replacement:="Ростовская область", MatchCase:=True
        .Replace What:="*Ростов-На-Дону,*", Replacement:="Ростовская область", MatchCase:=True
        .Replace What:="*Алтай*Респ*", Replacement:="Республика Алтай", MatchCase:=True
        .Replace What:="*Горно-Алтайск,*", Replacement:="Республика Алтай", MatchCase:=True
        .Replace What:="*Алтайск*край*", Replacement:="Алтайский край", MatchCase:=True
        .Replace What:="*Барнаул,*", Replacement:="Алтайский край", MatchCase:=True
        .Replace What:="*Амурск*обл*", Replacement:="Амурская область", MatchCase:=True
        .Replace What:="*Благовещенск,*", Replacement:="Амурская область", MatchCase:=True
        .Replace What:="*Архангельск*обл*", Replacement:="Архангельская область", MatchCase:=True
        .Replace What:="*Архангельск,*", Replacement:="Архангельская область", MatchCase:=True
        .Replace What:="*Брянск*обл*", Replacement:="Брянская область", MatchCase:=True
        .Replace What:="*Брянск,*", Replacement:="Брянская область", MatchCase:=True
        .Replace What:="*Бурят*Респ*", Replacement:="Республика Бурятия", MatchCase:=True
        .Replace What:="*Улан-Удэ,*", Replacement:="Республика Бурятия", MatchCase:=True
        .Replace What:="*Владимир*обл*", Replacement:="Владимирская область", MatchCase:=True
        .Replace What:="*Владимир*Обл*", Replacement:="Владимирская область", MatchCase:=True
        .Replace What:="*Владимир,*", Replacement:="Владимирская область", MatchCase:=True
        .Replace What:="*Волгоград*обл*", Replacement:="Волгоградская область", MatchCase:=True
        .Replace What:="*Волгоград,*", Replacement:="Волгоградская область", MatchCase:=True
        .Replace What:="*Вологодск*обл*", Replacement:="Вологодская область", MatchCase:=True
        .Replace What:="*Вологда,*", Replacement:="Вологодская область", MatchCase:=True
        .Replace What:="*Воронеж*обл*", Replacement:="Воронежская область", MatchCase:=True
        .Replace What:="*Воронеж,*", Replacement:="Воронежская область", MatchCase:=True
        .Replace What:="*Еврейск*", Replacement:="Еврейская АО", MatchCase:=True
        .Replace What:="*Забайкальск*край*", Replacement:="Забайкальский край", MatchCase:=True
        .Replace What:="*Чита,*", Replacement:="Забайкальский край", MatchCase:=True
        .Replace What:="*Иванов*обл*", Replacement:="Ивановская область", MatchCase:=True
        .Replace What:="*Иваново,*", Replacement:="Ивановская область", MatchCase:=True
        .Replace What:="*Белгород*обл*", Replacement:="Белгородская область", MatchCase:=True
        .Replace What:="*Белгород,*", Replacement:="Белгородская область", MatchCase:=True
        .Replace What:="*Туль*обл*", Replacement:="Тульская область", MatchCase:=True
        .Replace What:="*Тула,*", Replacement:="Тульская область", MatchCase:=True
        .Replace What:="*Иркутск*обл*", Replacement:="Иркутская область", MatchCase:=True
        .Replace What:="*Иркутск,*", Replacement:="Иркутская область", MatchCase:=True
        .Replace What:="*Калинингр*обл*", Replacement:="Калининградская область", MatchCase:=True
        .Replace What:="*Калининград,*", Replacement:="Калининградская область", MatchCase:=True
        .Replace What:="*Калмык*Респ*", Replacement:="Республика Калмыкия", MatchCase:=True
        .Replace What:="*Элиста,*", Replacement:="Республика Калмыкия", MatchCase:=True
        .Replace What:="*Калужск*обл*", Replacement:="Калужская область", MatchCase:=True
        .Replace What:="*Калуга,*", Replacement:="Калужская область", MatchCase:=True
        .Replace What:="*Камчатск*край*", Replacement:="Камчатский край", MatchCase:=True
        .Replace What:="*Костром*обл*", Replacement:="Костромская область", MatchCase:=True
        .Replace What:="*Кострома,*", Replacement:="Костромская область", MatchCase:=True
        .Replace What:="*Красноярск*край*", Replacement:="Красноярский край", MatchCase:=True
        .Replace What:="*Красноярск,*", Replacement:="Красноярский край", MatchCase:=True
        .Replace What:="*Курган*обл*", Replacement:="Курганская область", MatchCase:=True
        .Replace What:="*Курган,*", Replacement:="Курганская область", MatchCase:=True
        .Replace What:="*Курск*обл*", Replacement:="Курская область", MatchCase:=True
        .Replace What:="*Курск,*", Replacement:="Курская область", MatchCase:=True
        .Replace What:="*Ленинград*обл*", Replacement:="Ленинградская область", MatchCase:=True
        .Replace What:="*Липецк*обл*", Replacement:="Липецкая область", MatchCase:=True
        .Replace What:="*Липецк,*", Replacement:="Липецкая область", MatchCase:=True
        .Replace What:="*Магаданск*обл*", Replacement:="Магаданская область", MatchCase:=True
        .Replace What:="*Магадан,*", Replacement:="Магаданская область", MatchCase:=True
        .Replace What:="*Марий*", Replacement:="Республика Марий Эл", MatchCase:=True
        .Replace What:="*Йошкар-Ола,*", Replacement:="Республика Марий Эл", MatchCase:=True
        .Replace What:="*Мордов*Респ*", Replacement:="Республика Мордовия", MatchCase:=True
        .Replace What:="*Саранск,*", Replacement:="Республика Мордовия", MatchCase:=True
        .Replace What:="*Московск*обл*", Replacement:="Московская область", MatchCase:=True
        .Replace What:="*Мурманск*обл*", Replacement:="Мурманская область", MatchCase:=True
        .Replace What:="*Мурманск,*", Replacement:="Мурманская область", MatchCase:=True
        .Replace What:="*Нижегород*обл*", Replacement:="Нижегородская область", MatchCase:=True
        .Replace What:="*Нижний Новгород,*", Replacement:="Нижегородская область", MatchCase:=True
        .Replace What:="*Новгородск*обл*", Replacement:="Новгородская область", MatchCase:=True
        .Replace What:="*Новосибир*обл*", Replacement:="Новосибирская область", MatchCase:=True
        .Replace What:="*Новосибирск,*", Replacement:="Новосибирская область", MatchCase:=True
        .Replace What:="*Омск*обл*", Replacement:="Омская область", MatchCase:=True
        .Replace What:="*Омск,*", Replacement:="Омская область", MatchCase:=True
        .Replace What:="*Оренбургск*обл*", Replacement:="Оренбургская область", MatchCase:=True
        .Replace What:="*Оренбург,*", Replacement:="Оренбургская область", MatchCase:=True
        .Replace What:="*Орлов*обл*", Replacement:="Орловская область", MatchCase:=True
        .Replace What:="*Орёл,*", Replacement:="Орловская область", MatchCase:=True
        .Replace What:="*Пензенск*обл*", Replacement:="Пензенская область", MatchCase:=True
        .Replace What:="*Пенза,*", Replacement:="Пензенская область", MatchCase:=True
        .Replace What:="*Пермск*край*", Replacement:="Пермский край", MatchCase:=True
        .Replace What:="*Пермь,*", Replacement:="Пермский край", MatchCase:=True
        .Replace What:="*Приморск*край*", Replacement:="Приморский край", MatchCase:=True
        .Replace What:="*Владивосток,*", Replacement:="Приморский край", MatchCase:=True
        .Replace What:="*Псков*обл*", Replacement:="Псковская область", MatchCase:=True
        .Replace What:="*Псков,*", Replacement:="Псковская область", MatchCase:=True
        .Replace What:="*Рязан*обл*", Replacement:="Рязанская область", MatchCase:=True
        .Replace What:="*Рязань,*", Replacement:="Рязанская область", MatchCase:=True
        .Replace What:="*Самарск*обл*", Replacement:="Самарская область", MatchCase:=True
        .Replace What:="*Самара,*", Replacement:="Самарская область", MatchCase:=True
        .Replace What:="*Саратов*обл*", Replacement:="Саратовская область", MatchCase:=True
        .Replace What:="*Саратов,*", Replacement:="Саратовская область", MatchCase:=True
        .Replace What:="*Свердловск*обл*", Replacement:="Свердловская область", MatchCase:=True
        .Replace What:="*Екатеринбург,*", Replacement:="Свердловская область", MatchCase:=True
        .Replace What:="*Смоленск*обл*", Replacement:="Смоленская область", MatchCase:=True
        .Replace What:="*Смоленск,*", Replacement:="Смоленская область", MatchCase:=True
        .Replace What:="*Тамбов*обл*", Replacement:="Тамбовская область", MatchCase:=True
        .Replace What:="*Тамбов,*", Replacement:="Тамбовская область", MatchCase:=True
        .Replace What:="*Тверск*обл*", Replacement:="Тверская область", MatchCase:=True
        .Replace What:="*Тверь,*", Replacement:="Тверская область", MatchCase:=True
        .Replace What:="*Томск*обл*", Replacement:="Томская область", MatchCase:=True
        .Replace What:="*Томск,*", Replacement:="Томская область", MatchCase:=True
        .Replace What:="*Тыва*Респ*", Replacement:="Республика Тыва", MatchCase:=True
        .Replace What:="*Тюменск*обл*", Replacement:="Тюменская область", MatchCase:=True
        .Replace What:="*Тюмень,*", Replacement:="Тюменская область", MatchCase:=True
        .Replace What:="*Ульяновск*обл*", Replacement:="Ульяновская область", MatchCase:=True
        .Replace What:="*Ульяновск,*", Replacement:="Ульяновская область", MatchCase:=True
        .Replace What:="*Хабаровск*край*", Replacement:="Хабаровский край", MatchCase:=True
        .Replace What:="*Хабаровск,*", Replacement:="Хабаровский край", MatchCase:=True
        .Replace What:="*Ярославск*обл*", Replacement:="Ярославская область", MatchCase:=True
        .Replace What:="*Ярославль,*", Replacement:="Ярославская область", MatchCase:=True
        .Replace What:="*Киров*обл*", Replacement:="Кировская область", MatchCase:=True
        .Replace What:="*Киров,*", Replacement:="Кировская область", MatchCase:=True
        .Replace What:="*Сахалинск*обл*", Replacement:="Сахалинская область", MatchCase:=True
        .Replace What:="*Южно-Сахалинск,*", Replacement:="Сахалинская область", MatchCase:=True
        .Replace What:="*Севастопол*", Replacement:="Севастополь", MatchCase:=True
        .Replace What:="*Севастополь,*", Replacement:="Севастополь", MatchCase:=True
        .Replace What:="*Велик*Новгород*", Replacement:="Новгородская область", MatchCase:=True
    End With
End Sub
Sub Аналитика()
    Dim http        As New MSXML2.XMLHTTP60
    Dim doc         As HTMLDocument
    Dim i           As Long, v As Long
    Dim firstRow    As Integer, firstRow2 As Integer, lastrow As Integer, posName As Integer, posLineBreak As Integer, posINN As Integer, result_total As Integer
    Dim link        As String, org As String, Text As String, WorkSheetName As String, contract_number As String
    Dim number      As String, customer_type As String, placement_date As String, max_price As String, customer_inn As String, contract_summ As String, contract_sign As String
    Dim contract_end_date As String, contract_supplier As String, contract_supplier_inn As String, region As String
    Dim contract_fact_payed As String, contract_fact_done As String, contract_status As String
    Dim isSet       As Boolean
    Dim elements    As IHTMLDOMChildrenCollection
    Dim a           As IHTMLElement
    
    Call toggle_screen_upd
    
    WorkSheetName = ActiveSheet.name
    
    If WorkSheetName = "Взрослые" Then
        number = "Номер"
        region = "Регион"
        customer_type = "Заказчик_Тип"
        placement_date = "Размещено"
        max_price = "НМЦК"
        customer_inn = "Заказчик_ИНН"
        contract_summ = "ГК_Сумма"
        contract_sign = "Дата_заключения_ГК"
        contract_end_date = "Срок_действия"
        contract_supplier = "Победитель"
        contract_supplier_inn = "Победитель_ИНН"
        contract_fact_payed = "ГК_Фактически_оплачено"
        contract_fact_done = "ГК_Фактически_исполнено"
        contract_status = "ГК_Статус_контракта"
    End If
    If WorkSheetName = "Пеленки" Then
        number = "П_Номер"
        region = "П_Регион"
        customer_type = "П_Заказчик_Тип"
        placement_date = "П_Размещено"
        max_price = "П_НМЦК"
        customer_inn = "П_Заказчик_ИНН"
        contract_summ = "П_ГК_Сумма"
        contract_sign = "П_Дата_заключения_ГК"
        contract_end_date = "П_Срок_действия"
        contract_supplier = "П_Победитель"
        contract_supplier_inn = "П_Победитель_ИНН"
        contract_fact_payed = "П_ГК_Фактически_оплачено"
        contract_fact_done = "П_ГК_Фактически_исполнено"
        contract_status = "П_ГК_Статус_контракта"
    End If
    If WorkSheetName = "Детские" Then
        number = "Д_Номер"
        region = "Д_Регион"
        customer_type = "Д_Заказчик_Тип"
        placement_date = "Д_Размещено"
        max_price = "Д_НМЦК"
        customer_inn = "Д_Заказчик_ИНН"
        contract_summ = "Д_ГК_Сумма"
        contract_sign = "Д_Дата_заключения_ГК"
        contract_end_date = "Д_Срок_действия"
        contract_supplier = "Д_Победитель"
        contract_supplier_inn = "Д_Победитель_ИНН"
        contract_fact_payed = "Д_ГК_Фактически_оплачено"
        contract_fact_done = "Д_ГК_Фактически_исполнено"
        contract_status = "Д_ГК_Статус_контракта"
    End If
    
    lastrow = Range(number).Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    firstRow = Range(Range(number), Range(contract_end_date)).Find(What:="", SearchOrder:=xlRows, SearchDirection:=xlNext, LookIn:=xlValues).row
    firstRow2 = Range(Range(contract_supplier), Range(contract_status)).Find(What:="", SearchOrder:=xlRows, SearchDirection:=xlNext, LookIn:=xlValues).row
    
    If firstRow > firstRow2 Then
        firstRow = firstRow2
    End If
    firstRow2 = Range(contract_status).Find(What:="Исполнение", SearchOrder:=xlRows, SearchDirection:=xlNext, LookIn:=xlValues).row
    If firstRow > firstRow2 Then
        firstRow = firstRow2
    End If
    
    For i = firstRow To lastrow
        If IsEmpty(Range(region)(i)) Or IsEmpty(Range(customer_type)(i)) Or IsEmpty(Range(customer_inn)(i)) Or IsEmpty(Range(placement_date)(i)) Or IsEmpty(Range(max_price)(i)) Then
            
            Set doc = HTMLDoc(ZAK_COMMON & Replace(Range(number)(i).Value, "№", ""))
            
            Set a = doc.querySelector(".cardMainInfo__content > a")
            link = a.href
            org = a.innerText
            
            If IsEmpty(Range(max_price)(i)) Then
                Range(max_price)(i).Value = Val(Replace(doc.querySelector(".cardMainInfo__content.cost").innerText, ",", "."))
            End If
            
            Range(placement_date)(i).Value = CDate(Trim(doc.querySelector("div.date span.cardMainInfo__content").innerText))
            
            If IsEmpty(Range(customer_inn)(i)) Then
                
                Set doc = HTMLDoc(link)
                
                If IsEmpty(Range(region)(i)) Then
                    Range(region)(i).Value = RegionStr(doc.querySelector("div.registry-entry__body-value").innerText)
                End If
                
                If IsEmpty(Range(customer_type)(i)) Then
                    Range(customer_type)(i).Value = tip_zakazchik(org)
                End If
                
                isSet = FALSE
                Set elements = doc.querySelectorAll("div.col-md-auto")
                For v = 0 To elements.Length - 1
                    If isSet Then Exit For
                    
                    If UCase(elements.Item(v).innerText) Like Trim(UCase("*ИНН*")) Then
                        Range(customer_inn)(i).Value = Trim(elements.Item(v).querySelector(".registry-entry__body-value").innerText)
                        isSet = TRUE
                    End If
                    
                Next v
                
            End If
        End If
        If IsEmpty(Range(contract_summ)(i)) Or IsEmpty(Range(contract_sign)(i)) Or IsEmpty(Range(contract_end_date)(i)) _
           Or IsEmpty(Range(contract_supplier)(i)) Or IsEmpty(Range(contract_supplier_inn)(i)) Or IsEmpty(Range(contract_fact_payed)(i)) _
           Or IsEmpty(Range(contract_fact_done)(i)) Or IsEmpty(Range(contract_status)(i)) Or Range(contract_status)(i) = "Исполнение" Then
        
        Set doc = HTMLDoc("https://zakupki.gov.ru/epz/contract/search/results.html?orderNumber=" & Replace(Range(number)(i).Value, "№", ""))
        
        result_total = Val(Trim(doc.querySelector("div.search-results__total").innerText))
        
        If result_total > 1 Then
            Range(contract_supplier_inn)(i).Value = "Больше одной записи"
            Range(contract_status)(i).Value = "Больше одной записи"
            Range(contract_fact_payed)(i).Value = 0
            Range(contract_fact_done)(i).Value = 0
        Else
            Set a = doc.querySelector("div.registry-entry__header-mid__number > a")
            If Not a Is Nothing Then
                
                link = Replace(a.href, "about:", "https://zakupki.gov.ru")
                contract_number = Trim(Replace(a.innerText, "№", ""))
                
                If IsEmpty(Range(contract_summ)(i)) Then
                    Range(contract_summ)(i).Value = Val(Replace(doc.querySelector(".price-block__value").innerText, ",", "."))
                End If
                
                Set elements = doc.querySelectorAll("div.data-block__value")
                If IsEmpty(Range(contract_sign)(i)) Then
                    Range(contract_sign)(i).Value = CDate(Trim(elements.Item(0).innerText))
                End If
                
                If IsEmpty(Range(contract_end_date)(i)) Then
                    Range(contract_end_date)(i).Value = CDate(Trim(elements.Item(1).innerText))
                End If
                
                If IsEmpty(Range(contract_supplier)(i)) Or IsEmpty(Range(contract_supplier_inn)(i)) _
                   Or Range(contract_supplier_inn)(i) = "Нет контракта" Then
                Set doc = HTMLDoc(link)
                Text = doc.querySelector("div.participantsInnerHtml td.tableBlock__col").innerText
                
                If IsEmpty(Range(contract_supplier)(i)) Then
                    posName = InStr(text, "(ООО")
                    posLineBreak = InStr(text, vbCr)
                    
                    If posLineBreak = 0 Then
                        posLineBreak = InStr(text, "Код по ОКПО")
                    End If
                    If posName > 0 Then
                        Range(contract_supplier)(i).Value = Trim(Mid(text, posName + 1, posLineBreak - posName - 3))
                    Else
                        Range(contract_supplier)(i).Value = Trim(Left(text, posLineBreak - 1))
                    End If
                End If
                
                posINN = InStr(text, "ИНН")
                Text = Mid(text, posINN + 5, 12)
                
                If InStr(text, " ") = 0 Then
                    Range(contract_supplier_inn)(i).Value = Trim(Replace(text, vbCr, ""))
                Else
                    Range(contract_supplier_inn)(i).Value = Trim(Left(text, InStr(text, " ")))
                End If
                
            End If
            
            If IsEmpty(Range(contract_fact_payed)(i)) Or IsEmpty(Range(contract_fact_done)(i)) _
               Or IsEmpty(Range(contract_status)(i)) Or Range(contract_status)(i) = "Исполнение" Then
            
            Set doc = HTMLDoc("https://zakupki.gov.ru/epz/contract/contractCard/process-info.html?reestrNumber=" & contract_number)
            Range(contract_status)(i).Value = Trim(doc.querySelector("span.cardMainInfo__state").innerText)
            
            Set elements = doc.querySelectorAll("section.blockInfo__section > span")
            isSet = FALSE
            
            For v = 0 To elements.Length - 1
                If isSet Then Exit For
                
                If UCase(elements.Item(v).innerText) Like Trim(UCase("*исполнен*поставщик*обязательств*")) Then
                    Range(contract_fact_done)(i).Value = Val(Replace(elements.Item(v + 1).innerText, ",", "."))
                End If
                
                If UCase(elements.Item(v).innerText) Like Trim(UCase("*Фактичес*оплачено*")) Then
                    Range(contract_fact_payed)(i).Value = Val(Replace(elements.Item(v + 1).innerText, ",", "."))
                    If IsEmpty(Range(contract_fact_done)(i).Value) Then Range(contract_fact_done)(i).Value = 0
                    isSet = TRUE
                End If
                
            Next v
            
        End If
        
    Else
        
        If IsEmpty(Range(contract_supplier_inn)(i)) Then
            Range(contract_supplier_inn)(i).Value = "Нет контракта"
        End If
        
    End If
    
End If

End If

Next i

Set doc = Nothing
Set http = Nothing

Call toggle_screen_upd
End Sub
Function region_time_diff(s As String) As Integer
    Dim pos         As Integer
    region_time_diff = 3
    pos = InStr(s, "+")
    If pos = 0 Then pos = InStr(s, "-")
    If pos = 0 Then Exit Function
    region_time_diff = region_time_diff + Val(Mid(s, pos + 1))
End Function
Private Sub set_red_color(address As String)
    If Not IsEmpty(Range(address)) Then Range(address).Font.Color = RGB(182, 25, 25)
End Sub
Sub Реестр_Проверить_Изменения22()
    Dim doc         As HTMLDocument
    Dim i           As Long, v As Long, counter As Integer, date_current As Long, date_received As Long, time_received As Date, time_current As Date
    Dim received_date As Date
    Dim Text        As String, text2 As String
    Dim zak_number  As Range, visible_range As Range
    Dim lastrow     As Integer, hour_diff As Integer, utc_diff As Integer, total_rows As Integer
    Dim isSet       As Boolean, is_changed As Boolean
    Dim elements    As IHTMLDOMChildrenCollection
    Dim utc_dict    As Scripting.Dictionary
    Dim ws          As Worksheet
    Const label     As String = "Реестр Проверка Изменений "
    
    Set ws = Worksheets(ThisWorkbook.Sheets(1).name)
    Worksheets(ws.name).AutoFilterMode = FALSE
    Worksheets(ws.name).Range("A1").AutoFilter Field:=Worksheets(ws.name).Range("Статус")(1).Column, Criteria1:=Array("допущены", "заявлены", "идем", "расчет"), Operator:=xlFilterValues
    
    lastrow = Range("Статус").Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    Set visible_range = Range(Range("Номер")(2).address & ":" & Range("Номер")(lastrow).address).SpecialCells(xlCellTypeVisible)
    Set utc_dict = RegionUTCDictionary()
    total_rows = visible_range.count
    counter = 1
    
    For Each zak_number In visible_range
        
        i = zak_number.row
        isSet = FALSE
        Set doc = HTMLDoc(ZAK_COMMON & Replace(zak_number.Value, "№", ""))
        Set elements = doc.querySelectorAll(".blockInfo__section")
        
        For v = 0 To elements.Length - 1
            If isSet Then Exit For
            
            Text = UCase(elements.Item(v).innerText)
            is_changed = FALSE
            If Text Like "*ДАТА И ВРЕМЯ*ОКОНЧ*СРОК*ПОДАЧИ*" Then
                text2 = Range("Регион")(i).Value
                utc_diff = utc_dict.Item(Range("Регион")(i).Value)
                
                If utc_diff <> 0 Then
                    hour_diff = 3 - utc_diff
                Else
                    hour_diff = 3 - region_time_diff(Trim(elements.Item(v).querySelector("span:nth-child(2)").innerText))
                End If
                
                received_date = DateAdd("h", hour_diff, CDate(Left(Trim(elements.Item(v).querySelector("span:nth-child(2)").innerText), 16)))
                
                If changed_date(Range("Дата_окончания_подачи_заявок")(i), CLng(received_date)) Or changed_time(Range("Дата_окончания_подачи_заявок")(i), received_date) Then
                    Call set_red_color(Range("Дата_окончания_подачи_заявок")(i).address)
                    Range("Дата_окончания_подачи_заявок")(i).Value = received_date
                End If
                
                If changed_date(Range("Время_окончания_подачи_заявок")(i), CLng(received_date)) Or changed_time(Range("Время_окончания_подачи_заявок")(i), received_date) Then
                    Call set_red_color(Range("Время_окончания_подачи_заявок")(i).address)
                    Range("Время_окончания_подачи_заявок")(i).Value = received_date
                End If
                
                is_changed = changed_date(Range("Время_проведения_аукциона_конкурса")(i), CLng(received_date)) _
                             Or changed_time(Range("Время_проведения_аукциона_конкурса")(i), DateAdd("h", 2, received_date))
                If is_changed = TRUE And Range("Форма_проведения")(i).Value = "ЭА" Or Range("Форма_проведения")(i).Value = "ОКЭФ" Then
                    Call set_red_color(Range("Время_проведения_аукциона_конкурса")(i).address)
                    Call set_red_color(Range("Дата_проведения_аукциона_конкурса")(i).address)
                    Range("Дата_проведения_аукциона_конкурса")(i).Value = DateAdd("h", 2, Range("Время_окончания_подачи_заявок")(i).Value)
                    Range("Время_проведения_аукциона_конкурса")(i).Value = DateAdd("h", 2, Range("Время_окончания_подачи_заявок")(i).Value)
                End If
                
                If Range("Форма_проведения")(i).Value = "ЗКЭФ" And (changed_date(Range("Время_проведения_аукциона_конкурса")(i), CLng(received_date)) _
                   Or changed_time(Range("Время_проведения_аукциона_конкурса")(i), received_date)) Then
                Call set_red_color(Range("Время_проведения_аукциона_конкурса")(i).address)
                Call set_red_color(Range("Дата_проведения_аукциона_конкурса")(i).address)
                Range("Дата_проведения_аукциона_конкурса")(i).Value = Range("Время_окончания_подачи_заявок")(i).Value
                Range("Время_проведения_аукциона_конкурса")(i).Value = Range("Время_окончания_подачи_заявок")(i).Value
            End If
            
        End If
        
        If Text Like "*ДАТА*ПОДВЕД*ИТОГОВ*ОПРЕД*" Or Text Like "*ДАТА*ОКОНЧ*СРОК*РАССМ*ЗАЯВОК*" Or Text Like "*ДАТА*РАССМОТР*ОЦЕНК*ПЕРВ*ЧАСТЕЙ*" Then
            Text = Left(Trim(elements.Item(v).querySelector("span:nth-child(2)").innerText), 10)
            If changed_date(Range("Дата_окончания_срока_рассмотрения_заявок")(i), CDate(text)) Then
                Call set_red_color(Range("Дата_окончания_срока_рассмотрения_заявок")(i).address)
                Range("Дата_окончания_срока_рассмотрения_заявок")(i).Value = CDate(text)
            End If
            isSet = TRUE
        End If
        
    Next v
    Call show_status(counter, total_rows, label)
    counter = counter + 1
Next zak_number
End Sub
Private Sub Реестр_Проверить_Изменения()
    Dim doc         As HTMLDocument
    Dim i           As Long, v As Long, counter As Integer, date_current As Long, date_received As Long, time_received As Date, time_current As Date
    Dim Text        As String, address As String, reg As String
    Dim zak_number  As Range, visible_range As Range
    Dim lastrow     As Integer, hour_diff As Integer, utc_diff As Integer, total_rows As Integer
    Dim isSet       As Boolean
    Dim elements    As IHTMLDOMChildrenCollection
    Dim utc_dict    As Scripting.Dictionary
    Dim ws          As Worksheet
    Const label     As String = "Реестр Проверка Изменений "
    
    Set ws = Worksheets(ThisWorkbook.Sheets(1).name)
    Worksheets(ws.name).AutoFilterMode = FALSE
    Worksheets(ws.name).Range("A1").AutoFilter Field:=Worksheets(ws.name).Range("Статус")(1).Column, Criteria1:=Array("допущены", "заявлены", "идем", "расчет"), Operator:=xlFilterValues
    
    lastrow = Range("Статус").Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    Set visible_range = Range(Range("Номер")(2).address & ":" & Range("Номер")(lastrow).address).SpecialCells(xlCellTypeVisible)
    Set utc_dict = RegionUTCDictionary()
    total_rows = visible_range.count
    counter = 1
    
    For Each zak_number In visible_range
        
        i = zak_number.row
        isSet = FALSE
        Set doc = HTMLDoc(ZAK_COMMON & Replace(zak_number.Value, "№", ""))
        Set elements = doc.querySelectorAll(".blockInfo__section")
        
        For v = 0 To elements.Length - 1
            If isSet Then Exit For
            Text = UCase(elements.Item(v).innerText)
            If Text Like "*ПОЧТОВЫЙ АДРЕС*" Then
                address = Trim(elements.Item(v).querySelector("span:nth-child(2)").innerText)
            End If
            If Text Like "*ДАТА И ВРЕМЯ*ОКОНЧ*СРОК*ПОДАЧИ*" Then
                reg = RegionStr(address)
                utc_diff = utc_dict.Item(reg)
                If utc_diff <> 0 Then
                    hour_diff = 3 - utc_diff
                Else
                    hour_diff = 3 - region_time_diff(Trim(elements.Item(v).querySelector("span:nth-child(2)").innerText))
                End If
                Text = Left(Trim(elements.Item(v).querySelector("span:nth-child(2)").innerText), 16)
                If changed_date(Range("Дата_окончания_подачи_заявок")(i), CDate(text)) Then
                    Call set_red_color(Range("Дата_окончания_подачи_заявок")(i).address)
                    Range("Дата_окончания_подачи_заявок")(i).Value = CDate(text)
                End If
                time_received = DateAdd("h", hour_diff, TimeValue(text))
                time_current = TimeValue(CDate(Range("Время_окончания_подачи_заявок")(i)))
                If changed_time(time_received, time_current) Then
                    Call set_red_color(Range("Время_окончания_подачи_заявок")(i).address)
                    Range("Время_окончания_подачи_заявок")(i).Value = Range("Дата_окончания_подачи_заявок")(i).Value + time_received
                End If
            End If
            If Text Like "*ДАТА*ОКОНЧ*СРОК*РАССМ*ЗАЯВОК*" Or Text Like "*ДАТА*РАССМОТР*ОЦЕНК*ПЕРВ*ЧАСТЕЙ*" Then
                Text = Left(Trim(elements.Item(v).querySelector("span:nth-child(2)").innerText), 10)
                If changed_date(Range("Дата_окончания_срока_рассмотрения_заявок")(i), CDate(text)) Then
                    Call set_red_color(Range("Дата_окончания_срока_рассмотрения_заявок")(i).address)
                    Range("Дата_окончания_срока_рассмотрения_заявок")(i).Value = CDate(text)
                End If
            End If
            If Text Like "*ДАТА*ПРОВЕДЕНИЯ АУКЦИОНА*" Or Text Like "*ДАТА*ПОДАЧИ*ОКОНЧ*ПРЕДЛ*" Then
                Text = Left(Trim(elements.Item(v).querySelector("span:nth-child(2)").innerText), 10)
                If changed_date(Range("Дата_проведения_аукциона_конкурса")(i), CDate(text)) Then
                    Call set_red_color(Range("Дата_проведения_аукциона_конкурса")(i).address)
                    Range("Дата_проведения_аукциона_конкурса")(i).Value = CDate(text)
                End If
                If Range("Форма_проведения")(i) = "ОКЭФ" Then isSet = TRUE
            End If
            If Text Like "*ВРЕМЯ*ПРОВЕДЕНИЯ АУКЦИОНА*" Then
                Text = Left(Trim(elements.Item(v).querySelector("span:nth-child(2)").innerText), 5)
                time_received = DateAdd("h", hour_diff, TimeValue(text))
                time_current = TimeValue(CDate(Range("Время_проведения_аукциона_конкурса")(i)))
                If changed_time(time_received, time_current) Then
                    Call set_red_color(Range("Время_проведения_аукциона_конкурса")(i).address)
                    Range("Время_проведения_аукциона_конкурса")(i).Value = Range("Дата_проведения_аукциона_конкурса")(i).Value + time_received
                End If
                isSet = TRUE
            End If
        Next v
        Call show_status(counter, total_rows, label)
        counter = counter + 1
    Next zak_number
End Sub
Function changed_date(a As Long, b As Long) As Boolean
    changed_date = FALSE
    If a <> b Then changed_date = TRUE
End Function
Function changed_time(time_received As Date, time_current As Date) As Boolean
    changed_time = FALSE
    If Hour(time_received) <> Hour(time_current) Or Minute(time_received) <> Minute(time_current) Or Second(time_received) <> Second(time_current) Then
        changed_time = TRUE
    End If
End Function
Function bigger(x   As Integer, y As Integer) As Integer
    bigger = y
    If x > y Then bigger = x
End Function
Function smaller(x  As Integer, y As Integer) As Integer
    smaller = y
    If x < y Then smaller = x
End Function
Function reduced_name(name As String) As String
    reduced_name = UCase(Trim(Replace(name, "-.", "")))
    If reduced_name Like "*ОБЩЕСТВО*ОГРАНИЧ*ОТВЕТСТВ*" Then reduced_name = Replace(reduced_name, "ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ", "ООО")
    If reduced_name Like "*ЗАКРЫТ*АКЦИОНЕР*ОБЩЕСТ*" Then reduced_name = Replace(reduced_name, "ЗАКРЫТОЕ АКЦИОНЕРНОЕ ОБЩЕСТВО", "ЗАО"): Exit Function
    If reduced_name Like "*АКЦИОНЕР*ОБЩЕСТ*" Then reduced_name = Replace(reduced_name, "АКЦИОНЕРНОЕ ОБЩЕСТВО", "АО"): Exit Function
    If reduced_name Like "*ИНДИВИДУАЛЬН*ПРЕДПРИНИМ*" Then reduced_name = Replace(reduced_name, "ИНДИВИДУАЛЬНЫЙ ПРЕДПРИНИМАТЕЛЬ", "ИП"): Exit Function
    If reduced_name Like "*ОБЩЕСТВ*ИНВАЛИД*" Then
        reduced_name = Replace(reduced_name, "ОБЩЕСТВЕННАЯ ОРГАНИЗАЦИЯ ИНВАЛИДОВ", "ООИ")
        reduced_name = Replace(reduced_name, "ОБЛАСТНАЯ ОРГАНИЗАЦИЯ ОБЩЕРОССИЙСКОЙ ОБЩЕСТВЕННОЙ ОРГАНИЗАЦИИ", "ОООИ")
        reduced_name = Replace(reduced_name, "РЕГИОНАЛЬНОЕ ОТДЕЛЕНИЕ", "РО")
        reduced_name = Replace(reduced_name, "ВСЕРОССИЙСКОЕ ОБЩЕСТВО ИНВАЛИДОВ", "ВОИ")
        Exit Function
    End If
    If reduced_name Like "*ФЕДЕРАЛЬН*ГОСУДАРСТВ*УНИТАР*ПРЕДП*" Then
        reduced_name = Replace(reduced_name, "ФЕДЕРАЛЬНОГО ГОСУДАРСТВЕННОГО УНИТАРНОГО ПРЕДПРИЯТИЯ", "ФГУП")
        reduced_name = Replace(reduced_name, "ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ УНИТАРНОЕ ПРЕДПРИЯТИЕ", "ФГУП")
        reduced_name = Replace(reduced_name, "ПРОТЕЗНО-ОРТОПЕДИЧЕСКОЕ ПРЕДПРИЯТИЕ", "ПрОП")
        reduced_name = Replace(reduced_name, "МИНИСТЕРСТВА ТРУДА И СОЦИАЛЬНОЙ ЗАЩИТЫ РОССИЙСКОЙ ФЕДЕРАЦИИ", "МИНТРУДА И СЗ РФ")
        Exit Function
    End If
    If reduced_name Like "*ПРОТЕЗНО*ОРТОПЕД*ПРЕДПР*" Then reduced_name = Replace(reduced_name, "ПРОТЕЗНО-ОРТОПЕДИЧЕСКОЕ ПРЕДПРИЯТИЕ", "ПрОП"): Exit Function
    If reduced_name Like "*РЕГИОНАЛ*ОТДЕЛЕН*" Then reduced_name = Replace(reduced_name, "РЕГИОНАЛЬНОЕ ОТДЕЛЕНИЕ", "РО")
End Function
Function href_to_protocol(doc As HTMLDocument) As String
    Dim v           As Long
    Dim Text        As String
    Dim elements    As IHTMLDOMChildrenCollection
    
    Set elements = doc.querySelectorAll(".section__value.docName a")
    
    For v = 0 To elements.Length - 1
        Text = Trim(elements.Item(v).innerText)
        If Not Text Like "*об отмене*" And Text Like "*подвед*итогов*" Or Text Like "*рассмотрен*заяв*запрос*котир*" Then
            href_to_protocol = Replace(elements.Item(v).href, "about:", "https://zakupki.gov.ru")
            href_to_protocol = Replace(href_to_protocol, "main-info", "bid-list")
            href_to_protocol = Replace(href_to_protocol, "change-info", "bid-list")
        End If
    Next v
    
End Function
Function is_canceled(zak_number As String) As Boolean
    Dim v           As Long
    Dim doc         As HTMLDocument
    Dim elements    As IHTMLDOMChildrenCollection
    
    is_canceled = FALSE
    Set doc = HTMLDoc(ZAK_NOTICE & zak_number)
    Set elements = doc.querySelectorAll(".section__value.docName > span")
    
    For v = 0 To elements.Length - 1
        If elements.Item(v).innerText Like "*Извещ*отмен*опред*" Then is_canceled = True: Exit Function
    Next v
    
End Function
Function is_no_bids(zak_number As String) As Boolean
    Dim v           As Long, bids As Integer, disqualified_bids As Integer
    Dim i           As Long, Text As String
    Dim doc         As HTMLDocument
    Dim elements    As IHTMLDOMChildrenCollection
    
    is_no_bids = FALSE
    Set doc = HTMLDoc(ZAK_NOTICE & zak_number)
    Text = href_to_protocol(doc)
    
    If Text = "" Then Exit Function
    
    Set doc = HTMLDoc(text)
    Set elements = doc.querySelectorAll("section.blockInfo__section")
    
    For i = 0 To elements.Length - 1
        
        Text = Trim(elements.Item(i).querySelector("span.section__title").innerText)
        If Text = "Подано заявок" Then
            Text = Trim(elements.Item(i).querySelector("span.section__info").innerText)
            bids = Val(text)
            If bids = 0 Then
                is_no_bids = TRUE
                Exit Function
            End If
            
            Text = Mid(text, InStr(text, "отклонено: ") + 11)
            disqualified_bids = Val(text)
            If bids = disqualified_bids Then is_no_bids = TRUE
            
            Exit Function
        End If
        
    Next i
    
End Function
Private Sub Реестр_Победители()
    Dim doc         As HTMLDocument
    Dim i           As Long, v As Long
    Dim lastrow     As Integer, range_number As Integer, zkef_bids_count As Integer
    Dim total_rows  As Integer, counter As Integer
    Dim visible_range As Range, zak As Range
    Dim participant As String, admission As String, cr As String, zak_number As String, Text As String
    Dim no_bids     As Boolean, has_anybody As Boolean
    Dim protocol_date As IHTMLElement, table As IHTMLElement
    Dim elements    As IHTMLDOMChildrenCollection
    Dim bids        As Scripting.Dictionary
    Dim ws          As Worksheet
    Dim wb          As Workbook
    Const label     As String = "Реестр Итоги Закупок "
    
    Set ws = Worksheets(ThisWorkbook.Sheets(1).name)
    Worksheets(ws.name).AutoFilterMode = FALSE
    Worksheets(ws.name).Range("A1").AutoFilter Field:=Worksheets(ws.name).Range("Победитель")(1).Column, Criteria1:="="
    Worksheets(ws.name).Range("A1").AutoFilter Field:=Worksheets(ws.name).Range("Дата_окончания_подачи_заявок")(1).Column, Criteria1:="<" & CLng(Now())
    
    lastrow = Range("Номер").Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    Set visible_range = Range(Range("Номер")(2).address & ":" & Range("Номер")(lastrow).address).SpecialCells(xlCellTypeVisible)
    total_rows = visible_range.count
    counter = 1
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = FALSE
    Application.EnableEvents = FALSE
    
    For Each zak In visible_range
        
        i = zak.row
        no_bids = FALSE
        has_anybody = FALSE
        If IsEmpty(Range("Победитель")(i)) Or Range("Победитель")(i).Value = "" Then
            zak_number = Replace(Range("Номер")(i), "№", "")
            Set doc = HTMLDoc(SUPPLIER_RESULTS & zak_number)
            Set protocol_date = doc.querySelector("section.blockInfo__section section.blockInfo__section:last-child span:last-child")
            If Not protocol_date Is Nothing Then
                Set elements = doc.querySelectorAll("td.tableBlock__col")
                For v = 0 To elements.Length - 1
                    Text = UCase(elements.Item(v).innerText)
                    If Text Like "*1*-*ПОБЕДИТЕЛЬ*" Then
                        Range("Сумма_выигранного_лота")(i).Value = Val(Replace(elements.Item(v + 1).innerText, ",", "."))
                        has_anybody = TRUE
                        Exit For
                    End If
                    If Text Like "*ПОДАН*ТОЛЬКО*ОДН*ПРИЗНАНА СООТВЕТСТВУЮЩ*" Or _
                       Text Like "*ПОДАН*ТОЛЬКО*ОДН*ЗАЯВКА СООТВЕТСТ*" Or _
                       Text Like "*ПОДАН*ЕДИНСТВ*ПРЕДЛОЖ*О ЕЕ СООТВЕТСТВИИ*" Or _
                       Text Like "*РЕШЕНИЕ*СООТВЕ*ТОЛЬКО*ОДНОЙ*" Then
                    
                    Range("Сумма_выигранного_лота")(i).Value = Val(Replace(elements.Item(v + 1).innerText, ",", "."))
                    Range("Победитель")(i).Value = reduced_name(Trim(elements.Item(v - 1).innerText))
                    has_anybody = TRUE
                    no_bids = TRUE
                    Exit For
                End If
            Next v
            If has_anybody = FALSE Then
                Range("Победитель")(i).Value = "не состоялся"
                no_bids = TRUE
            End If
        ElseIf Range("Форма_проведения")(i) = "ЗКЭФ" Then
            no_bids = TRUE
            If is_no_bids(zak_number) Then Range("Победитель")(i).Value = "не состоялся"
        Else
            no_bids = TRUE
            If is_canceled(zak_number) Then Range("Победитель")(i).Value = "отменен"
        End If
        If no_bids = FALSE Then
            Set doc = HTMLDoc(ZAK_NOTICE & zak_number)
            Set doc = HTMLDoc(href_to_protocol(doc))
            Set table = doc.querySelector("table")
            Set elements = table.querySelectorAll("tr.table__row:not(:empty), tr.tableBlock__row")
            Set bids = New Scripting.Dictionary
            zkef_bids_count = 0
            For v = 1 To elements.Length - 1
                If Range("Форма_проведения")(i) = "ЭА" Then
                    admission = Trim(elements.Item(v).querySelector("td:nth-child(4)").innerText)
                    range_number = Val(elements.Item(v).querySelector("td:nth-child(5)").innerText)
                Else
                    admission = Trim(elements.Item(v).querySelector("td:nth-child(5)").innerText)
                    range_number = Val(elements.Item(v).querySelector("td:nth-child(6)").innerText)
                End If
                If admission = "Допущена" Or admission = "Соответствует требованиям" Then
                    participant = reduced_name(elements.Item(v).querySelector("td:nth-child(3)").innerText)
                    If Range("Форма_проведения")(i) = "ЗКЭФ" And range_number = 0 Then
                        If Trim(elements.Item(v).querySelector("td:nth-child(6)").innerText) = "Победитель" Then
                            range_number = 1
                        Else
                            range_number = zkef_bids_count + 2
                            zkef_bids_count = zkef_bids_count + 1
                        End If
                    End If
                    bids.Add range_number, participant
                End If
            Next v
            
            Range("Победитель")(i).Value = bids.Item(1)
            cr = ""
            For v = 2 To bids.count
                If v > 5 Then Exit For
                If v > 2 Then cr = vbCrLf
                If Range("Форма_проведения")(i) <> "ЗКЭФ" Then
                    Range("Участники")(i).Value = Range("Участники")(i).Value + cr + CStr(v) + " - " + bids.Item(v)
                Else
                    Range("Участники")(i).Value = Range("Участники")(i).Value + cr + bids.Item(v)
                End If
            Next v
            
        End If
    End If
    Call show_status(counter, total_rows, label)
    counter = counter + 1
Next zak

Worksheets(ws.name).Range("A1").AutoFilter Field:=Worksheets(ws.name).Range("Победитель")(1).Column
Worksheets(ws.name).Range("A1").AutoFilter Field:=Worksheets(ws.name).Range("Дата_окончания_подачи_заявок")(1).Column

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = TRUE
Application.EnableEvents = TRUE
End Sub
Sub Реестр_Победители22()
    Dim doc         As HTMLDocument
    Dim i           As Long, v As Long
    Dim lastrow     As Integer, range_number As Integer
    Dim total_rows  As Integer, counter As Integer
    Dim visible_range As Range, zak As Range
    Dim zak_number  As String, Text As String, contractConclusionNumber As String, block_title As String
    Dim protocol_date As IHTMLElement
    Dim elements    As IHTMLDOMChildrenCollection, search_block As IHTMLDOMChildrenCollection
    Dim ws          As Worksheet
    Dim wb          As Workbook
    Const label     As String = "Реестр Итоги Закупок "
    
    Set ws = Worksheets(ThisWorkbook.Sheets(1).name)
    Worksheets(ws.name).AutoFilterMode = FALSE
    Worksheets(ws.name).Range("A1").AutoFilter Field:=Worksheets(ws.name).Range("Победитель")(1).Column, Criteria1:="="
    Worksheets(ws.name).Range("A1").AutoFilter Field:=Worksheets(ws.name).Range("Дата_окончания_подачи_заявок")(1).Column, Criteria1:="<" & CLng(Now())
    
    lastrow = Range("Номер").Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).row
    Set visible_range = Range(Range("Номер")(2).address & ":" & Range("Номер")(lastrow).address).SpecialCells(xlCellTypeVisible)
    total_rows = visible_range.count
    counter = 1
    
    '    Call toggle_screen_upd
    
    For Each zak In visible_range
        
        i = zak.row
        
        If IsEmpty(Range("Победитель")(i)) Or Range("Победитель")(i).Value = "" Then
            
            zak_number = Replace(Range("Номер")(i), "№", "")
            
            If is_canceled(zak_number) Then
                Range("Победитель")(i).Value = "отменен"
            Else
                Set doc = HTMLDoc(SUPPLIER_RESULTS & zak_number)
                Set protocol_date = doc.querySelector("section.blockInfo__section section.blockInfo__section:last-child span:last-child")
                
                If Not protocol_date Is Nothing Then
                    
                    If is_no_bids(zak_number) Then
                        Range("Победитель")(i).Value = "не состоялся"
                    Else
                        Set elements = doc.querySelectorAll("td.tableBlock__col")
                        For v = 0 To elements.Length - 2
                            Text = UCase(Trim(elements.Item(v).innerText))
                            If Text Like "*1*-*ПОБЕДИТЕЛЬ*" Or Text = "1" And elements.Item(v + 1).innerText <> "1" Then
                                Range("Сумма_выигранного_лота")(i).Value = Val(Replace(elements.Item(v + 1).innerText, ",", "."))
                                Exit For
                            ElseIf Text Like "*ПОДАН*ТОЛЬКО*ОДН*ПРИЗНАНА СООТВЕТСТВУЮЩ*" Or _
                                   Text Like "*ПОДАН*ТОЛЬКО*ОДН*ЗАЯВКА СООТВЕТСТ*" Or _
                                   Text Like "*ПОДАН*ЕДИНСТВ*ПРЕДЛОЖ*О ЕЕ СООТВЕТСТВИИ*" Or _
                                   Text Like "*РЕШЕНИЕ*СООТВЕ*ТОЛЬКО*ОДНОЙ*" Then
                            
                            Range("Сумма_выигранного_лота")(i).Value = Val(Replace(elements.Item(v + 1).innerText, ",", "."))
                            Exit For
                        ElseIf Text = "1" And elements.Item(v + 1).innerText = "1" Then
                            Range("Сумма_выигранного_лота")(i).Value = Val(Replace(elements.Item(v + 2).innerText, ",", "."))
                            Exit For
                            '                            ElseIf IsNumeric(text) And elements.Item(v + 1).innerText = "" And IsNumeric(elements.Item(v + 2).innerText) Then
                            '                                Range("Сумма_выигранного_лота")(i).Value = Val(Replace(elements.Item(v + 2).innerText, ",", "."))
                            '                                Exit For
                        End If
                    Next v
                    
                    contractConclusionNumber = zak_number & "0001"
                    
                    Set doc = HTMLDoc(CONTRACT_CONCLUSION_COMMON & contractConclusionNumber)
                    
                    Set elements = doc.querySelectorAll("div.row.blockInfo")
                    For v = 1 To elements.Length - 1
                        
                        block_title = UCase(elements.Item(v).querySelector("h2.blockInfo__title").innerText)
                        
                        If block_title Like "*ИНФОРМ*ПОСТАВЩИК*" Then
                            
                            Set search_block = elements.Item(v).querySelectorAll("section")
                            If element_info("ВИД", search_block) = "Физическое лицо" Then
                                Range("Победитель")(i).Value = reduced_name(element_info("*ФАМИЛИЯ*ИМЯ*ОТЧЕСТ*", search_block))
                                Exit For
                            End If
                            
                            Range("Победитель")(i).Value = reduced_name(element_info("*ПОЛНОЕ*НАИМ*ПОСТАВЩ*", search_block))
                            Exit For
                        End If
                        
                    Next v
                End If
            End If
        End If
        
    End If
    Call show_status(counter, total_rows, label)
    counter = counter + 1
Next zak

Worksheets(ws.name).Range("A1").AutoFilter Field:=Worksheets(ws.name).Range("Победитель")(1).Column
Worksheets(ws.name).Range("A1").AutoFilter Field:=Worksheets(ws.name).Range("Дата_окончания_подачи_заявок")(1).Column

'  Call toggle_screen_upd
End Sub
Private Function RegionStr(s As String) As String
    RegionStr = UCase(s)
    If RegionStr Like "*КАРЕЛ*РЕСП*" Then RegionStr = "Республика Карелия": Exit Function
    If RegionStr Like "*ПЕТРОЗАВОДСК,*" Then RegionStr = "Республика Карелия": Exit Function
    If RegionStr Like "*КОМИ*РЕСП*" Then RegionStr = "Республика Коми": Exit Function
    If RegionStr Like "*СЫКТЫВКАР,*" Then RegionStr = "Республика Коми": Exit Function
    If RegionStr Like "*ЧЕЧЕН*РЕСП*" Then RegionStr = "Чеченская Республика": Exit Function
    If RegionStr Like "*ГРОЗНЫЙ,*" Then RegionStr = "Чеченская Республика": Exit Function
    If RegionStr Like "*ЧУВАШ*РЕСП*" Then RegionStr = "Республика Чувашия": Exit Function
    If RegionStr Like "*ЭЛИСТА,*" Then RegionStr = "Республика Чувашия": Exit Function
    If RegionStr Like "*ЧУКОТСКИЙ*" Then RegionStr = "Чукотский АО": Exit Function
    If RegionStr Like "*АНЫДЫРЬ,*" Then RegionStr = "Чукотский АО": Exit Function
    If RegionStr Like "*УДМУРТ*РЕСП*" Then RegionStr = "Удмуртская Республика": Exit Function
    If RegionStr Like "*ИЖЕВСК,*" Then RegionStr = "Удмуртская Республика": Exit Function
    If RegionStr Like "*ИНГУШ*РЕСП*" Then RegionStr = "Республика Ингушетия": Exit Function
    If RegionStr Like "*КЕМЕРОВ*ОБЛ*" Then RegionStr = "Кемеровская область": Exit Function
    If RegionStr Like "*КЕМЕРОВО,*" Then RegionStr = "Кемеровская область": Exit Function
    If RegionStr Like "*ДАГЕСТ*РЕСП*" Then RegionStr = "Республика Дагестан": Exit Function
    If RegionStr Like "*МАХАЧКАЛА,*" Then RegionStr = "Республика Дагестан": Exit Function
    If RegionStr Like "*КРЫМ*РЕСП*" Then RegionStr = "Республика Крым": Exit Function
    If RegionStr Like "*КРЫМ,*" Then RegionStr = "Республика Крым": Exit Function
    If RegionStr Like "*СИМФЕРОПОЛЬ,*" Then RegionStr = "Республика Крым": Exit Function
    If RegionStr Like "*САХА*ЯКУТИ*" Then RegionStr = "Республика Саха (Якутия)": Exit Function
    If RegionStr Like "*ЯКУТСК,*" Then RegionStr = "Республика Саха (Якутия)": Exit Function
    If RegionStr Like "*ХАКАС*" Then RegionStr = "Республика Хакасия": Exit Function
    If RegionStr Like "*АБАКАН,*" Then RegionStr = "Республика Хакасия": Exit Function
    If RegionStr Like "*ХАНТЫ*МАН*" Then RegionStr = "Ханты-Мансийский АО — Югра": Exit Function
    If RegionStr Like "*ХАНТЫ-МАНСИЙСК,*" Then RegionStr = "Ханты-Мансийский АО — Югра": Exit Function
    If RegionStr Like "*БАШКОРТ*РЕСП*" Then RegionStr = "Республика Башкортостан": Exit Function
    If RegionStr Like "*УФА,*" Then RegionStr = "Республика Башкортостан": Exit Function
    If RegionStr Like "*САНКТ*ПЕТ*" Then RegionStr = "Санкт-Петербург": Exit Function
    If RegionStr Like "*САНКТ-ПЕТЕРБУРГ,*" Then RegionStr = "Санкт-Петербург": Exit Function
    If RegionStr Like "*ЯМАЛО*НЕН*" Then RegionStr = "Ямало-Ненецкий АО": Exit Function
    If RegionStr Like "*САЛЕХАРД,*" Then RegionStr = "Ямало-Ненецкий АО": Exit Function
    If RegionStr Like "*ТАТАРСТАН*" Then RegionStr = "Республика Татарстан": Exit Function
    If RegionStr Like "*КАЗАНЬ,*" Then RegionStr = "Республика Татарстан": Exit Function
    If RegionStr Like "*КРАСНОДАР*КРАЙ*" Then RegionStr = "Краснодарский край": Exit Function
    If RegionStr Like "*КРАСНОДАР,*" Then RegionStr = "Краснодарский край": Exit Function
    If RegionStr Like "*Г.КРАСНОДАР*" Then RegionStr = "Краснодарский край": Exit Function
    If RegionStr Like "*Г. КРАСНОДАР*" Then RegionStr = "Краснодарский край": Exit Function
    If RegionStr Like "*ЧЕЛЯБИНСК*ОБЛ*" Then RegionStr = "Челябинская область": Exit Function
    If RegionStr Like "*ЧЕЛЯБИНСК,*" Then RegionStr = "Челябинская область": Exit Function
    If RegionStr Like "*ОСЕТИ*РЕСП*" Then RegionStr = "Республика Северная Осетия - Алания": Exit Function
    If RegionStr Like "*ВЛАДИКАВКАЗ,*" Then RegionStr = "Республика Северная Осетия - Алания": Exit Function
    If RegionStr Like "*СТАВРОПОЛ*КРАЙ*" Then RegionStr = "Ставропольский край": Exit Function
    If RegionStr Like "*СТАВРОПОЛЬ,*" Then RegionStr = "Ставропольский край": Exit Function
    If RegionStr Like "*КАБАРД*РЕСП*" Then RegionStr = "Республика Кабардино-Балкарская": Exit Function
    If RegionStr Like "*НАЛЬЧИК,*" Then RegionStr = "Республика Кабардино-Балкарская": Exit Function
    If RegionStr Like "*АСТРАХАН*ОБЛ*" Then RegionStr = "Астраханская область": Exit Function
    If RegionStr Like "*АСТРАХАНЬ,*" Then RegionStr = "Астраханская область": Exit Function
    If RegionStr Like "*АДЫГ*РЕСП*" Then RegionStr = "Республика Адыгея": Exit Function
    If RegionStr Like "*МАЙКОП,*" Then RegionStr = "Республика Адыгея": Exit Function
    If RegionStr Like "*КАРАЧ*РЕСП*" Then RegionStr = "Республика Карачаево-Черкесская": Exit Function
    If RegionStr Like "*ЧЕРКЕССК,*" Then RegionStr = "Республика Карачаево-Черкесская": Exit Function
    If RegionStr Like "*МОСКВ*" Then RegionStr = "Москва": Exit Function
    If RegionStr Like "*МОСКВА,*" Then RegionStr = "Москва": Exit Function
    If RegionStr Like "*РОСТОВ*ОБЛ*" Then RegionStr = "Ростовская область": Exit Function
    If RegionStr Like "*РОСТОВ-НА-ДОНУ,*" Then RegionStr = "Ростовская область": Exit Function
    If RegionStr Like "*АЛТАЙ*РЕСП*" Then RegionStr = "Республика Алтай": Exit Function
    If RegionStr Like "*ГОРНО-АЛТАЙСК,*" Then RegionStr = "Республика Алтай": Exit Function
    If RegionStr Like "*АЛТАЙСК*КРАЙ*" Then RegionStr = "Алтайский край": Exit Function
    If RegionStr Like "*БАРНАУЛ,*" Then RegionStr = "Алтайский край": Exit Function
    If RegionStr Like "*АМУРСК*ОБЛ*" Then RegionStr = "Амурская область": Exit Function
    If RegionStr Like "*БЛАГОВЕЩЕНСК,*" Then RegionStr = "Амурская область": Exit Function
    If RegionStr Like "*АРХАНГЕЛЬСК*ОБЛ*" Then RegionStr = "Архангельская область": Exit Function
    If RegionStr Like "*АРХАНГЕЛЬСК,*" Then RegionStr = "Архангельская область": Exit Function
    If RegionStr Like "*БРЯНСК*ОБЛ*" Then RegionStr = "Брянская область": Exit Function
    If RegionStr Like "*БРЯНСК,*" Then RegionStr = "Брянская область": Exit Function
    If RegionStr Like "*БУРЯТ*РЕСП*" Then RegionStr = "Республика Бурятия": Exit Function
    If RegionStr Like "*УЛАН-УДЭ,*" Then RegionStr = "Республика Бурятия": Exit Function
    If RegionStr Like "*ВЛАДИМИР*ОБЛ*" Then RegionStr = "Владимирская область": Exit Function
    If RegionStr Like "*ВЛАДИМИР*ОБЛ*" Then RegionStr = "Владимирская область": Exit Function
    If RegionStr Like "*ВЛАДИМИР,*" Then RegionStr = "Владимирская область": Exit Function
    If RegionStr Like "*ВОЛГОГРАД*ОБЛ*" Then RegionStr = "Волгоградская область": Exit Function
    If RegionStr Like "*ВОЛГОГРАД,*" Then RegionStr = "Волгоградская область": Exit Function
    If RegionStr Like "*ВОЛОГОДСК*ОБЛ*" Then RegionStr = "Вологодская область": Exit Function
    If RegionStr Like "*ВОЛОГДА,*" Then RegionStr = "Вологодская область": Exit Function
    If RegionStr Like "*ВОРОНЕЖ*ОБЛ*" Then RegionStr = "Воронежская область": Exit Function
    If RegionStr Like "*ВОРОНЕЖ,*" Then RegionStr = "Воронежская область": Exit Function
    If RegionStr Like "*ЕВРЕЙСК*" Then RegionStr = "Еврейская АО": Exit Function
    If RegionStr Like "*ЗАБАЙКАЛЬСК*КРАЙ*" Then RegionStr = "Забайкальский край": Exit Function
    If RegionStr Like "*ЧИТА,*" Then RegionStr = "Забайкальский край": Exit Function
    If RegionStr Like "*ИВАНОВ*ОБЛ*" Then RegionStr = "Ивановская область": Exit Function
    If RegionStr Like "*ИВАНОВО,*" Then RegionStr = "Ивановская область": Exit Function
    If RegionStr Like "*БЕЛГОРОД*ОБЛ*" Then RegionStr = "Белгородская область": Exit Function
    If RegionStr Like "*БЕЛГОРОД,*" Then RegionStr = "Белгородская область": Exit Function
    If RegionStr Like "*ТУЛЬ*ОБЛ*" Then RegionStr = "Тульская область": Exit Function
    If RegionStr Like "*ТУЛА,*" Then RegionStr = "Тульская область": Exit Function
    If RegionStr Like "*ИРКУТСК*ОБЛ*" Then RegionStr = "Иркутская область": Exit Function
    If RegionStr Like "*ИРКУТСК,*" Then RegionStr = "Иркутская область": Exit Function
    If RegionStr Like "*КАЛИНИНГР*ОБЛ*" Then RegionStr = "Калининградская область": Exit Function
    If RegionStr Like "*КАЛИНИНГРАД,*" Then RegionStr = "Калининградская область": Exit Function
    If RegionStr Like "*КАЛМЫК*РЕСП*" Then RegionStr = "Республика Калмыкия": Exit Function
    If RegionStr Like "*ЭЛИСТА,*" Then RegionStr = "Республика Калмыкия": Exit Function
    If RegionStr Like "*КАЛУЖСК*ОБЛ*" Then RegionStr = "Калужская область": Exit Function
    If RegionStr Like "*КАЛУГА,*" Then RegionStr = "Калужская область": Exit Function
    If RegionStr Like "*КАМЧАТСК*КРАЙ*" Then RegionStr = "Камчатский край": Exit Function
    If RegionStr Like "*КОСТРОМ*ОБЛ*" Then RegionStr = "Костромская область": Exit Function
    If RegionStr Like "*КОСТРОМА,*" Then RegionStr = "Костромская область": Exit Function
    If RegionStr Like "*КРАСНОЯРСК*КРАЙ*" Then RegionStr = "Красноярский край": Exit Function
    If RegionStr Like "*КРАСНОЯРСК,*" Then RegionStr = "Красноярский край": Exit Function
    If RegionStr Like "*КУРГАН*ОБЛ*" Then RegionStr = "Курганская область": Exit Function
    If RegionStr Like "*КУРГАН,*" Then RegionStr = "Курганская область": Exit Function
    If RegionStr Like "*КУРСК*ОБЛ*" Then RegionStr = "Курская область": Exit Function
    If RegionStr Like "*КУРСК,*" Then RegionStr = "Курская область": Exit Function
    If RegionStr Like "*ЛЕНИНГРАД*ОБЛ*" Then RegionStr = "Ленинградская область": Exit Function
    If RegionStr Like "*ЛИПЕЦК*ОБЛ*" Then RegionStr = "Липецкая область": Exit Function
    If RegionStr Like "*ЛИПЕЦК,*" Then RegionStr = "Липецкая область": Exit Function
    If RegionStr Like "*МАГАДАНСК*ОБЛ*" Then RegionStr = "Магаданская область": Exit Function
    If RegionStr Like "*МАГАДАН,*" Then RegionStr = "Магаданская область": Exit Function
    If RegionStr Like "*МАРИЙ*" Then RegionStr = "Республика Марий Эл": Exit Function
    If RegionStr Like "*ЙОШКАР-ОЛА,*" Then RegionStr = "Республика Марий Эл": Exit Function
    If RegionStr Like "*МОРДОВ*РЕСП*" Then RegionStr = "Республика Мордовия": Exit Function
    If RegionStr Like "*САРАНСК,*" Then RegionStr = "Республика Мордовия": Exit Function
    If RegionStr Like "*МОСКОВСК*ОБЛ*" Then RegionStr = "Московская область": Exit Function
    If RegionStr Like "*МУРМАНСК*ОБЛ*" Then RegionStr = "Мурманская область": Exit Function
    If RegionStr Like "*МУРМАНСК,*" Then RegionStr = "Мурманская область": Exit Function
    If RegionStr Like "*НИЖЕГОРОД*ОБЛ*" Then RegionStr = "Нижегородская область": Exit Function
    If RegionStr Like "*НИЖНИЙ НОВГОРОД,*" Then RegionStr = "Нижегородская область": Exit Function
    If RegionStr Like "*НОВГОРОДСК*ОБЛ*" Then RegionStr = "Новгородская область": Exit Function
    If RegionStr Like "*НОВОСИБИР*ОБЛ*" Then RegionStr = "Новосибирская область": Exit Function
    If RegionStr Like "*НОВОСИБИРСК,*" Then RegionStr = "Новосибирская область": Exit Function
    If RegionStr Like "*ТОМСК*ОБЛ*" Then RegionStr = "Томская область": Exit Function
    If RegionStr Like "*ТОМСК,*" Then RegionStr = "Томская область": Exit Function
    If RegionStr Like "*ОМСК*ОБЛ*" Then RegionStr = "Омская область": Exit Function
    If RegionStr Like "*ОМСК,*" Then RegionStr = "Омская область": Exit Function
    If RegionStr Like "*ОРЕНБУРГСК*ОБЛ*" Then RegionStr = "Оренбургская область": Exit Function
    If RegionStr Like "*ОРЕНБУРГ,*" Then RegionStr = "Оренбургская область": Exit Function
    If RegionStr Like "*ОРЛОВ*ОБЛ*" Then RegionStr = "Орловская область": Exit Function
    If RegionStr Like "*ОРЁЛ,*" Then RegionStr = "Орловская область": Exit Function
    If RegionStr Like "*ПЕНЗЕНСК*ОБЛ*" Then RegionStr = "Пензенская область": Exit Function
    If RegionStr Like "*ПЕНЗА,*" Then RegionStr = "Пензенская область": Exit Function
    If RegionStr Like "*ПЕРМСК*КРАЙ*" Then RegionStr = "Пермский край": Exit Function
    If RegionStr Like "*ПЕРМЬ,*" Then RegionStr = "Пермский край": Exit Function
    If RegionStr Like "*ПРИМОРСК*КРАЙ*" Then RegionStr = "Приморский край": Exit Function
    If RegionStr Like "*ВЛАДИВОСТОК,*" Then RegionStr = "Приморский край": Exit Function
    If RegionStr Like "*ПСКОВ*ОБЛ*" Then RegionStr = "Псковская область": Exit Function
    If RegionStr Like "*ПСКОВ,*" Then RegionStr = "Псковская область": Exit Function
    If RegionStr Like "*РЯЗАН*ОБЛ*" Then RegionStr = "Рязанская область": Exit Function
    If RegionStr Like "*РЯЗАНЬ,*" Then RegionStr = "Рязанская область": Exit Function
    If RegionStr Like "*САМАРСК*ОБЛ*" Then RegionStr = "Самарская область": Exit Function
    If RegionStr Like "*САМАРА,*" Then RegionStr = "Самарская область": Exit Function
    If RegionStr Like "*САРАТОВ*ОБЛ*" Then RegionStr = "Саратовская область": Exit Function
    If RegionStr Like "*САРАТОВ,*" Then RegionStr = "Саратовская область": Exit Function
    If RegionStr Like "*СВЕРДЛОВСК*ОБЛ*" Then RegionStr = "Свердловская область": Exit Function
    If RegionStr Like "*ЕКАТЕРИНБУРГ,*" Then RegionStr = "Свердловская область": Exit Function
    If RegionStr Like "*СМОЛЕНСК*ОБЛ*" Then RegionStr = "Смоленская область": Exit Function
    If RegionStr Like "*СМОЛЕНСК,*" Then RegionStr = "Смоленская область": Exit Function
    If RegionStr Like "*ТАМБОВ*ОБЛ*" Then RegionStr = "Тамбовская область": Exit Function
    If RegionStr Like "*ТАМБОВ,*" Then RegionStr = "Тамбовская область": Exit Function
    If RegionStr Like "*ТВЕРСК*ОБЛ*" Then RegionStr = "Тверская область": Exit Function
    If RegionStr Like "*ТВЕРЬ,*" Then RegionStr = "Тверская область": Exit Function
    If RegionStr Like "*ТЫВА*РЕСП*" Then RegionStr = "Республика Тыва": Exit Function
    If RegionStr Like "*ТЮМЕНСК*ОБЛ*" Then RegionStr = "Тюменская область": Exit Function
    If RegionStr Like "*ТЮМЕНЬ,*" Then RegionStr = "Тюменская область": Exit Function
    If RegionStr Like "*УЛЬЯНОВСК*ОБЛ*" Then RegionStr = "Ульяновская область": Exit Function
    If RegionStr Like "*УЛЬЯНОВСК,*" Then RegionStr = "Ульяновская область": Exit Function
    If RegionStr Like "*ХАБАРОВСК*КРАЙ*" Then RegionStr = "Хабаровский край": Exit Function
    If RegionStr Like "*ХАБАРОВСК,*" Then RegionStr = "Хабаровский край": Exit Function
    If RegionStr Like "*ЯРОСЛАВСК*ОБЛ*" Then RegionStr = "Ярославская область": Exit Function
    If RegionStr Like "*ЯРОСЛАВЛЬ,*" Then RegionStr = "Ярославская область": Exit Function
    If RegionStr Like "*КИРОВ*ОБЛ*" Then RegionStr = "Кировская область": Exit Function
    If RegionStr Like "*КИРОВ,*" Then RegionStr = "Кировская область": Exit Function
    If RegionStr Like "*САХАЛИНСК*ОБЛ*" Then RegionStr = "Сахалинская область": Exit Function
    If RegionStr Like "*ЮЖНО-САХАЛИНСК,*" Then RegionStr = "Сахалинская область": Exit Function
    If RegionStr Like "*СЕВАСТОПОЛ*" Then RegionStr = "Севастополь": Exit Function
    If RegionStr Like "*СЕВАСТОПОЛЬ,*" Then RegionStr = "Севастополь": Exit Function
    If RegionStr Like "*ВЕЛИК*НОВГОРОД*" Then RegionStr = "Новгородская область": Exit Function
End Function
Private Function RegionCustomerStr(s As String) As String
    RegionCustomerStr = UCase(s)
    If RegionCustomerStr Like "*РЕСП*КАРЕЛИЯ*" Then RegionCustomerStr = "Республика Карелия": Exit Function
    If RegionCustomerStr Like "*ПЕТРОЗАВОДСК*" Then RegionCustomerStr = "Республика Карелия": Exit Function
    If RegionCustomerStr Like "*РЕСП*КОМИ*" Then RegionCustomerStr = "Республика Коми": Exit Function
    If RegionCustomerStr Like "*СЫКТЫВКАР*" Then RegionCustomerStr = "Республика Коми": Exit Function
    If RegionCustomerStr Like "*ЧЕЧЕН*РЕСП*" Then RegionCustomerStr = "Чеченская Республика": Exit Function
    If RegionCustomerStr Like "*РЕСП*ЧЕЧН*" Then RegionCustomerStr = "Чеченская Республика": Exit Function
    If RegionCustomerStr Like "*ЧУВАШ*РЕСП*" Then RegionCustomerStr = "Республика Чувашия": Exit Function
    If RegionCustomerStr Like "*РЕСП*ЧУВАШ*" Then RegionCustomerStr = "Республика Чувашия": Exit Function
    If RegionCustomerStr Like "*ЧУКОТСК*" Then RegionCustomerStr = "Чукотский АО": Exit Function
    If RegionCustomerStr Like "*УДМУРТ*РЕСП*" Then RegionCustomerStr = "Удмуртская Республика": Exit Function
    If RegionCustomerStr Like "*РЕСП*УДМУРТ*" Then RegionCustomerStr = "Удмуртская Республика": Exit Function
    If RegionCustomerStr Like "*ИНГУШ*РЕСП*" Then RegionCustomerStr = "Республика Ингушетия": Exit Function
    If RegionCustomerStr Like "*РЕСП*ИНГУШЕТ*" Then RegionCustomerStr = "Республика Ингушетия": Exit Function
    If RegionCustomerStr Like "*КЕМЕРОВ*ОБЛ*" Then RegionCustomerStr = "Кемеровская область": Exit Function
    If RegionCustomerStr Like "*КЕМЕРОВ*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Кемеровская область": Exit Function
    If RegionCustomerStr Like "*ДАГЕСТ*РЕСП*" Then RegionCustomerStr = "Республика Дагестан": Exit Function
    If RegionCustomerStr Like "*РЕСП*ДАГЕСТ*" Then RegionCustomerStr = "Республика Дагестан": Exit Function
    If RegionCustomerStr Like "*КРЫМ*РЕСП*" Then RegionCustomerStr = "Республика Крым": Exit Function
    If RegionCustomerStr Like "*РЕСП*КРЫМ*" Then RegionCustomerStr = "Республика Крым": Exit Function
    If RegionCustomerStr Like "*САХА*ЯКУТИ*" Then RegionCustomerStr = "Республика Саха (Якутия)": Exit Function
    If RegionCustomerStr Like "*РЕСП*ХАКАС*" Then RegionCustomerStr = "Республика Хакасия": Exit Function
    If RegionCustomerStr Like "*ХАНТЫ-МАНСИЙСК*" Then RegionCustomerStr = "Ханты-Мансийский АО — Югра": Exit Function
    If RegionCustomerStr Like "*БАШКОРТ*РЕСП*" Then RegionCustomerStr = "Республика Башкортостан": Exit Function
    If RegionCustomerStr Like "*РЕСП*БАШК*" Then RegionCustomerStr = "Республика Башкортостан": Exit Function
    If RegionCustomerStr Like "*САНКТ*ПЕТ*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Санкт-Петербург": Exit Function
    If RegionCustomerStr Like "*ЯМАЛО*НЕН*" Then RegionCustomerStr = "Ямало-Ненецкий АО": Exit Function
    If RegionCustomerStr Like "*РЕСП*ТАТАРСТАН*" Then RegionCustomerStr = "Республика Татарстан": Exit Function
    If RegionCustomerStr Like "*КРАСНОДАР*КРАЙ*" Then RegionCustomerStr = "Краснодарский край": Exit Function
    If RegionCustomerStr Like "*КРАСНОДАР*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Краснодарский край": Exit Function
    If RegionCustomerStr Like "*ЧЕЛЯБИНСК*ОБЛ*" Then RegionCustomerStr = "Челябинская область": Exit Function
    If RegionCustomerStr Like "*ЧЕЛЯБИНСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Челябинская область": Exit Function
    If RegionCustomerStr Like "*ОСЕТИ*РЕСП*" Then RegionCustomerStr = "Республика Северная Осетия - Алания": Exit Function
    If RegionCustomerStr Like "*РЕСП*ОСЕТИ*" Then RegionCustomerStr = "Республика Северная Осетия - Алания": Exit Function
    If RegionCustomerStr Like "*СТАВРОПОЛ*КРАЙ*" Then RegionCustomerStr = "Ставропольский край": Exit Function
    If RegionCustomerStr Like "*СТАВРОПОЛЬ*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Ставропольский край": Exit Function
    If RegionCustomerStr Like "*КАБАРД*РЕСП*" Then RegionCustomerStr = "Республика Кабардино-Балкарская": Exit Function
    If RegionCustomerStr Like "*РЕСП*КАБАРД*" Then RegionCustomerStr = "Республика Кабардино-Балкарская": Exit Function
    If RegionCustomerStr Like "*АСТРАХАН*ОБЛ*" Then RegionCustomerStr = "Астраханская область": Exit Function
    If RegionCustomerStr Like "*АСТРАХАН*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Астраханская область": Exit Function
    If RegionCustomerStr Like "*АДЫГ*РЕСП*" Then RegionCustomerStr = "Республика Адыгея": Exit Function
    If RegionCustomerStr Like "*РЕСП*АДЫГ*" Then RegionCustomerStr = "Республика Адыгея": Exit Function
    If RegionCustomerStr Like "*КАРАЧ*РЕСП*" Then RegionCustomerStr = "Республика Карачаево-Черкесская": Exit Function
    If RegionCustomerStr Like "*РЕСП*КАРАЧ*" Then RegionCustomerStr = "Республика Карачаево-Черкесская": Exit Function
    If RegionCustomerStr Like "*МОСКВЫ*" Then RegionCustomerStr = "Москва": Exit Function
    If RegionCustomerStr Like "*МОСКВА*" Then RegionCustomerStr = "Москва": Exit Function
    If RegionCustomerStr Like "*РОСТОВ*ОБЛ*" Then RegionCustomerStr = "Ростовская область": Exit Function
    If RegionCustomerStr Like "*РОСТОВ*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Ростовская область": Exit Function
    If RegionCustomerStr Like "*АЛТАЙ*РЕСП*" Then RegionCustomerStr = "Республика Алтай": Exit Function
    If RegionCustomerStr Like "*РЕСП*АЛТАЙ*" Then RegionCustomerStr = "Республика Алтай": Exit Function
    If RegionCustomerStr Like "*АЛТАЙСК*КРАЙ*" Then RegionCustomerStr = "Алтайский край": Exit Function
    If RegionCustomerStr Like "*АЛТАЙСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Алтайский край": Exit Function
    If RegionCustomerStr Like "*АМУРСК*ОБЛ*" Then RegionCustomerStr = "Амурская область": Exit Function
    If RegionCustomerStr Like "*АМУРСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Амурская область": Exit Function
    If RegionCustomerStr Like "*АРХАНГЕЛЬСК*ОБЛ*" Then RegionCustomerStr = "Архангельская область": Exit Function
    If RegionCustomerStr Like "*АРХАНГЕЛЬСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Архангельская область": Exit Function
    If RegionCustomerStr Like "*БРЯНСК*ОБЛ*" Then RegionCustomerStr = "Брянская область": Exit Function
    If RegionCustomerStr Like "*БРЯНСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Брянская область": Exit Function
    If RegionCustomerStr Like "*БУРЯТ*РЕСП*" Then RegionCustomerStr = "Республика Бурятия": Exit Function
    If RegionCustomerStr Like "*РЕСП*БУРЯТ*" Then RegionCustomerStr = "Республика Бурятия": Exit Function
    If RegionCustomerStr Like "*УЛАН-УДЭ*" Then RegionCustomerStr = "Республика Бурятия": Exit Function
    If RegionCustomerStr Like "*ВЛАДИМИР*ОБЛ*" Then RegionCustomerStr = "Владимирская область": Exit Function
    If RegionCustomerStr Like "*ВЛАДИМИР*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Владимирская область": Exit Function
    If RegionCustomerStr Like "*ВОЛГОГРАД*ОБЛ*" Then RegionCustomerStr = "Волгоградская область": Exit Function
    If RegionCustomerStr Like "*ВОЛГОГРАД*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Волгоградская область": Exit Function
    If RegionCustomerStr Like "*ВОЛОГОДСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Вологодская область": Exit Function
    If RegionCustomerStr Like "*ВОРОНЕЖ*ОБЛ*" Then RegionCustomerStr = "Воронежская область": Exit Function
    If RegionCustomerStr Like "*ВОРОНЕЖ*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Воронежская область": Exit Function
    If RegionCustomerStr Like "*ЕВРЕЙСК*" Then RegionCustomerStr = "Еврейская АО": Exit Function
    If RegionCustomerStr Like "*ЗАБАЙКАЛЬСК*КРАЙ*" Then RegionCustomerStr = "Забайкальский край": Exit Function
    If RegionCustomerStr Like "*ЗАБАЙКАЛЬСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Забайкальский край": Exit Function
    If RegionCustomerStr Like "*ИВАНОВ*ОБЛ*" Then RegionCustomerStr = "Ивановская область": Exit Function
    If RegionCustomerStr Like "*ИВАНОВСКОЕ*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Ивановская область": Exit Function
    If RegionCustomerStr Like "*БЕЛГОРОД*ОБЛ*" Then RegionCustomerStr = "Белгородская область": Exit Function
    If RegionCustomerStr Like "*БЕЛГОРОД*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Белгородская область": Exit Function
    If RegionCustomerStr Like "*ТУЛЬ*ОБЛ*" Then RegionCustomerStr = "Тульская область": Exit Function
    If RegionCustomerStr Like "*ТУЛЬСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Тульская область": Exit Function
    If RegionCustomerStr Like "*ИРКУТСК*ОБЛ*" Then RegionCustomerStr = "Иркутская область": Exit Function
    If RegionCustomerStr Like "*ИРКУТСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Иркутская область": Exit Function
    If RegionCustomerStr Like "*КАЛИНИНГР*ОБЛ*" Then RegionCustomerStr = "Калининградская область": Exit Function
    If RegionCustomerStr Like "*КАЛИНИНГР*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Калининградская область": Exit Function
    If RegionCustomerStr Like "*КАЛМЫК*РЕСП*" Then RegionCustomerStr = "Республика Калмыкия": Exit Function
    If RegionCustomerStr Like "*КАЛУЖСК*ОБЛ*" Then RegionCustomerStr = "Калужская область": Exit Function
    If RegionCustomerStr Like "*КАЛУЖСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Калужская область": Exit Function
    If RegionCustomerStr Like "*КАМЧАТСК*КРАЙ*" Then RegionCustomerStr = "Камчатский край": Exit Function
    If RegionCustomerStr Like "*КАМЧАТСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Камчатский край": Exit Function
    If RegionCustomerStr Like "*КОСТРОМ*ОБЛ*" Then RegionCustomerStr = "Костромская область": Exit Function
    If RegionCustomerStr Like "*КОСТРОМ*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Костромская область": Exit Function
    If RegionCustomerStr Like "*КРАСНОЯРСК*КРАЙ*" Then RegionCustomerStr = "Красноярский край": Exit Function
    If RegionCustomerStr Like "*КРАСНОЯРСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Красноярский край": Exit Function
    If RegionCustomerStr Like "*КУРГАН*ОБЛ*" Then RegionCustomerStr = "Курганская область": Exit Function
    If RegionCustomerStr Like "*КУРГАН*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Курганская область": Exit Function
    If RegionCustomerStr Like "*КУРСК*ОБЛ*" Then RegionCustomerStr = "Курская область": Exit Function
    If RegionCustomerStr Like "*КУРСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Курская область": Exit Function
    If RegionCustomerStr Like "*ЛЕНИНГРАД*ОБЛ*" Then RegionCustomerStr = "Ленинградская область": Exit Function
    If RegionCustomerStr Like "*ЛЕНИНГРАД*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Ленинградская область": Exit Function
    If RegionCustomerStr Like "*ЛИПЕЦК*ОБЛ*" Then RegionCustomerStr = "Липецкая область": Exit Function
    If RegionCustomerStr Like "*ЛИПЕЦК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Липецкая область": Exit Function
    If RegionCustomerStr Like "*МАГАДАНСК*ОБЛ*" Then RegionCustomerStr = "Магаданская область": Exit Function
    If RegionCustomerStr Like "*МАГАДАНСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Магаданская область": Exit Function
    If RegionCustomerStr Like "*МАРИЙ ЭЛ*" Then RegionCustomerStr = "Республика Марий Эл": Exit Function
    If RegionCustomerStr Like "*ЙОШКАР-ОЛА*" Then RegionCustomerStr = "Республика Марий Эл": Exit Function
    If RegionCustomerStr Like "*МОРДОВ*РЕСП*" Then RegionCustomerStr = "Республика Мордовия": Exit Function
    If RegionCustomerStr Like "*РЕСП*МОРДОВ*" Then RegionCustomerStr = "Республика Мордовия": Exit Function
    If RegionCustomerStr Like "*МОСКОВСК*ОБЛ*" Then RegionCustomerStr = "Московская область": Exit Function
    If RegionCustomerStr Like "*МОСКОВСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Московская область": Exit Function
    If RegionCustomerStr Like "*МУРМАНСК*ОБЛ*" Then RegionCustomerStr = "Мурманская область": Exit Function
    If RegionCustomerStr Like "*МУРМАНСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Мурманская область": Exit Function
    If RegionCustomerStr Like "*НИЖЕГОРОД*ОБЛ*" Then RegionCustomerStr = "Нижегородская область": Exit Function
    If RegionCustomerStr Like "*НИЖЕГОРОД*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Нижегородская область": Exit Function
    If RegionCustomerStr Like "*НОВГОРОДСК*ОБЛ*" Then RegionCustomerStr = "Новгородская область": Exit Function
    If RegionCustomerStr Like "*НОВГОРОДСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Новгородская область": Exit Function
    If RegionCustomerStr Like "*НОВОСИБИР*ОБЛ*" Then RegionCustomerStr = "Новосибирская область": Exit Function
    If RegionCustomerStr Like "*НОВОСИБИР*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Новосибирская область": Exit Function
    If RegionCustomerStr Like "*ТОМСК*ОБЛ*" Then RegionCustomerStr = "Томская область": Exit Function
    If RegionCustomerStr Like "*ТОМСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Томская область": Exit Function
    If RegionCustomerStr Like "*ОМСК*ОБЛ*" Then RegionCustomerStr = "Омская область": Exit Function
    If RegionCustomerStr Like "*ОМСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Омская область": Exit Function
    If RegionCustomerStr Like "*ОРЕНБУРГСК*ОБЛ*" Then RegionCustomerStr = "Оренбургская область": Exit Function
    If RegionCustomerStr Like "*ОРЕНБУРГСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Оренбургская область": Exit Function
    If RegionCustomerStr Like "*ОРЛОВ*ОБЛ*" Then RegionCustomerStr = "Орловская область": Exit Function
    If RegionCustomerStr Like "*ОРЛОВ*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Орловская область": Exit Function
    If RegionCustomerStr Like "*ПЕНЗЕНСК*ОБЛ*" Then RegionCustomerStr = "Пензенская область": Exit Function
    If RegionCustomerStr Like "*ПЕНЗЕНСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Пензенская область": Exit Function
    If RegionCustomerStr Like "*ПЕРМСК*КРАЙ*" Then RegionCustomerStr = "Пермский край": Exit Function
    If RegionCustomerStr Like "*ПЕРМСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Пермский область": Exit Function
    If RegionCustomerStr Like "*ПРИМОРСК*КРАЙ*" Then RegionCustomerStr = "Приморский край": Exit Function
    If RegionCustomerStr Like "*ПРИМОРСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Приморский край": Exit Function
    If RegionCustomerStr Like "*ПСКОВ*ОБЛ*" Then RegionCustomerStr = "Псковская область": Exit Function
    If RegionCustomerStr Like "*ПСКОВ*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Псковская область": Exit Function
    If RegionCustomerStr Like "*РЯЗАН*ОБЛ*" Then RegionCustomerStr = "Рязанская область": Exit Function
    If RegionCustomerStr Like "*РЯЗАН*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Рязанская область": Exit Function
    If RegionCustomerStr Like "*САМАРСК*ОБЛ*" Then RegionCustomerStr = "Самарская область": Exit Function
    If RegionCustomerStr Like "*САМАРСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Самарская область": Exit Function
    If RegionCustomerStr Like "*САРАТОВ*ОБЛ*" Then RegionCustomerStr = "Саратовская область": Exit Function
    If RegionCustomerStr Like "*САРАТОВ*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Саратовская область": Exit Function
    If RegionCustomerStr Like "*СВЕРДЛОВСК*ОБЛ*" Then RegionCustomerStr = "Свердловская область": Exit Function
    If RegionCustomerStr Like "*СВЕРДЛОВСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Свердловская область": Exit Function
    If RegionCustomerStr Like "*СМОЛЕНСК*ОБЛ*" Then RegionCustomerStr = "Смоленская область": Exit Function
    If RegionCustomerStr Like "*СМОЛЕНСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Смоленская область": Exit Function
    If RegionCustomerStr Like "*ТАМБОВ*ОБЛ*" Then RegionCustomerStr = "Тамбовская область": Exit Function
    If RegionCustomerStr Like "*ТАМБОВ*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Тамбовская область": Exit Function
    If RegionCustomerStr Like "*ТВЕРСК*ОБЛ*" Then RegionCustomerStr = "Тверская область": Exit Function
    If RegionCustomerStr Like "*ТВЕРСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Тверская область": Exit Function
    If RegionCustomerStr Like "*ТЫВА*РЕСП*" Then RegionCustomerStr = "Республика Тыва": Exit Function
    If RegionCustomerStr Like "*РЕСП*ТЫВА*" Then RegionCustomerStr = "Республика Тыва": Exit Function
    If RegionCustomerStr Like "*ТЮМЕНСК*ОБЛ*" Then RegionCustomerStr = "Тюменская область": Exit Function
    If RegionCustomerStr Like "*ТЮМЕНСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Тюменская область": Exit Function
    If RegionCustomerStr Like "*УЛЬЯНОВСК*ОБЛ*" Then RegionCustomerStr = "Ульяновская область": Exit Function
    If RegionCustomerStr Like "*УЛЬЯНОВСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Ульяновская область": Exit Function
    If RegionCustomerStr Like "*ХАБАРОВСК*КРАЙ*" Then RegionCustomerStr = "Хабаровский край": Exit Function
    If RegionCustomerStr Like "*ХАБАРОВСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Хабаровский край": Exit Function
    If RegionCustomerStr Like "*ЯРОСЛАВСК*ОБЛ*" Then RegionCustomerStr = "Ярославская область": Exit Function
    If RegionCustomerStr Like "*ЯРОСЛАВСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Ярославская область": Exit Function
    If RegionCustomerStr Like "*КИРОВ*ОБЛ*" Then RegionCustomerStr = "Кировская область": Exit Function
    If RegionCustomerStr Like "*КИРОВ*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Кировская область": Exit Function
    If RegionCustomerStr Like "*САХАЛИНСК*ОБЛ*" Then RegionCustomerStr = "Сахалинская область": Exit Function
    If RegionCustomerStr Like "*САХАЛИНСК*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Сахалинская область": Exit Function
    If RegionCustomerStr Like "*СЕВАСТОПОЛЬ*РЕГ*ОТДЕЛЕН*" Then RegionCustomerStr = "Севастополь": Exit Function
    RegionCustomerStr = ""
End Function
Function RegionUTCDictionary() As Scripting.Dictionary
    Set RegionUTCDictionary = New Scripting.Dictionary
    RegionUTCDictionary.Add "Республика Карелия", 3
    RegionUTCDictionary.Add "Республика Коми", 3
    RegionUTCDictionary.Add "Чеченская Республика", 3
    RegionUTCDictionary.Add "Республика Чувашия", 3
    RegionUTCDictionary.Add "Чукотский АО", 12
    RegionUTCDictionary.Add "Удмуртская Республика", 4
    RegionUTCDictionary.Add "Республика Ингушетия", 3
    RegionUTCDictionary.Add "Кемеровская область", 7
    RegionUTCDictionary.Add "Республика Дагестан", 3
    RegionUTCDictionary.Add "Республика Крым", 3
    RegionUTCDictionary.Add "Республика Саха (Якутия)", 9
    RegionUTCDictionary.Add "Республика Хакасия", 7
    RegionUTCDictionary.Add "Ханты-Мансийский АО — Югра", 5
    RegionUTCDictionary.Add "Республика Башкортостан", 5
    RegionUTCDictionary.Add "Санкт-Петербург", 3
    RegionUTCDictionary.Add "Ямало-Ненецкий АО", 3
    RegionUTCDictionary.Add "Республика Татарстан", 3
    RegionUTCDictionary.Add "Краснодарский край", 3
    RegionUTCDictionary.Add "Челябинская область", 5
    RegionUTCDictionary.Add "Республика Северная Осетия - Алания", 3
    RegionUTCDictionary.Add "Ставропольский край", 3
    RegionUTCDictionary.Add "Республика Кабардино-Балкарская", 3
    RegionUTCDictionary.Add "Астраханская область", 4
    RegionUTCDictionary.Add "Республика Адыгея", 3
    RegionUTCDictionary.Add "Республика Карачаево-Черкесская", 3
    RegionUTCDictionary.Add "Москва", 3
    RegionUTCDictionary.Add "Ростовская область", 3
    RegionUTCDictionary.Add "Республика Алтай", 7
    RegionUTCDictionary.Add "Алтайский край", 7
    RegionUTCDictionary.Add "Амурская область", 9
    RegionUTCDictionary.Add "Архангельская область", 3
    RegionUTCDictionary.Add "Брянская область", 3
    RegionUTCDictionary.Add "Республика Бурятия", 8
    RegionUTCDictionary.Add "Владимирская область", 3
    RegionUTCDictionary.Add "Волгоградская область", 3
    RegionUTCDictionary.Add "Вологодская область", 3
    RegionUTCDictionary.Add "Воронежская область", 3
    RegionUTCDictionary.Add "Еврейская АО", 10
    RegionUTCDictionary.Add "Забайкальский край", 9
    RegionUTCDictionary.Add "Ивановская область", 3
    RegionUTCDictionary.Add "Белгородская область", 3
    RegionUTCDictionary.Add "Тульская область", 3
    RegionUTCDictionary.Add "Иркутская область", 8
    RegionUTCDictionary.Add "Калининградская область", 2
    RegionUTCDictionary.Add "Республика Калмыкия", 3
    RegionUTCDictionary.Add "Калужская область", 3
    RegionUTCDictionary.Add "Камчатский край", 12
    RegionUTCDictionary.Add "Костромская область", 3
    RegionUTCDictionary.Add "Красноярский край", 7
    RegionUTCDictionary.Add "Курганская область", 5
    RegionUTCDictionary.Add "Курская область", 3
    RegionUTCDictionary.Add "Ленинградская область", 3
    RegionUTCDictionary.Add "Липецкая область", 3
    RegionUTCDictionary.Add "Магаданская область", 11
    RegionUTCDictionary.Add "Республика Марий Эл", 3
    RegionUTCDictionary.Add "Республика Мордовия", 3
    RegionUTCDictionary.Add "Московская область", 3
    RegionUTCDictionary.Add "Мурманская область", 3
    RegionUTCDictionary.Add "Нижегородская область", 3
    RegionUTCDictionary.Add "Новгородская область", 3
    RegionUTCDictionary.Add "Новосибирская область", 7
    RegionUTCDictionary.Add "Омская область", 6
    RegionUTCDictionary.Add "Оренбургская область", 5
    RegionUTCDictionary.Add "Орловская область", 3
    RegionUTCDictionary.Add "Пензенская область", 3
    RegionUTCDictionary.Add "Пермский край", 5
    RegionUTCDictionary.Add "Приморский край", 10
    RegionUTCDictionary.Add "Псковская область", 3
    RegionUTCDictionary.Add "Рязанская область", 3
    RegionUTCDictionary.Add "Самарская область", 4
    RegionUTCDictionary.Add "Саратовская область", 4
    RegionUTCDictionary.Add "Свердловская область", 5
    RegionUTCDictionary.Add "Смоленская область", 3
    RegionUTCDictionary.Add "Тамбовская область", 3
    RegionUTCDictionary.Add "Тверская область", 3
    RegionUTCDictionary.Add "Томская область", 7
    RegionUTCDictionary.Add "Республика Тыва", 7
    RegionUTCDictionary.Add "Тюменская область", 5
    RegionUTCDictionary.Add "Ульяновская область", 4
    RegionUTCDictionary.Add "Хабаровский край", 10
    RegionUTCDictionary.Add "Ярославская область", 3
    RegionUTCDictionary.Add "Кировская область", 3
    RegionUTCDictionary.Add "Сахалинская область", 11
    RegionUTCDictionary.Add "Севастополь", 3
End Function
Private Sub show_status(current As Integer, total_rows As Integer, topic As String)
    Dim NumberOfBars As Integer, CurrentStatus As Integer
    Dim pctDone     As Integer
    
    NumberOfBars = 50
    CurrentStatus = Int((current / total_rows) * NumberOfBars)
    pctDone = Round(CurrentStatus / NumberOfBars * 100, 0)
    Application.StatusBar = topic & " [" & String(CurrentStatus, "|") & _
                            Space(NumberOfBars - CurrentStatus) & "]" & _
                            " " & pctDone & "% Завершено"
    
    If current = total_rows Then Application.StatusBar = ""
    
End Sub
Private Function HTMLDoc(url As String) As HTMLDocument
    
    Dim http        As New MSXML2.XMLHTTP60
    http.Open "GET", url, FALSE
    http.send
    Set HTMLDoc = New HTMLDocument
    HTMLDoc.body.innerHTML = http.responseText
    
End Function
Private Sub toggle_screen_upd()
    
    If Application.Calculation = xlCalculationAutomatic Then
        Application.Calculation = xlCalculationManual
    Else
        Application.Calculation = xlCalculationAutomatic
    End If
    
    If Application.ScreenUpdating Then
        Application.ScreenUpdating = FALSE
    Else
        Application.ScreenUpdating = TRUE
    End If
    
    If Application.EnableEvents Then
        Application.EnableEvents = FALSE
    Else
        Application.EnableEvents = TRUE
    End If
    
    If Application.DisplayStatusBar = FALSE Then Application.DisplayStatusBar = TRUE
    
End Sub