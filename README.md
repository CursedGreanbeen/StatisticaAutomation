## Запуск макросов

### 1. Макрос для STATISTICA (извлечение таблиц в Excel)

**Файл `ExcelExp (working).svb`**  
Обходит все элементы (таблицы) в открытом файле `.stw` и выгружает их в книгу Excel:  
- таблицы с именем, содержащим «Mann-Whitney», объединяет на отдельный лист;  
- «Descriptive Statistics...» и «Сравнение групп...» – на лист `RES_NPAR`;  
- остальные таблицы – на отдельные листы с очисткой имени.

**Как запустить:**  

1. Откройте STATISTICA.  
2. Откройте нужный файл `.stw`.  
3. Откройте редактор макросов: `Tools → Macro → New` (если его ещё нет).  
4. Вставьте полный текст макроса в модуль `STB` или в новый модуль. Убедитесь, что в начале файла есть строки:  
   `'#Uses "*STB.SVX"`  
   `'#Uses "*GRAPHICS.SVX"`  
5. Закройте редактор.  
6. Запустите макрос: `Tools → Macro → Run` (или `Alt+F8`), выберите `Main` → `Run`.  
7. В появившемся окне введите путь к папке (например, `C:\Work\MacroLab\`) и имя файла `.stw` (без расширения).  
8. Макрос создаст файл `.xlsx` с обработанными данными в той же папке.

**Примечания:**  
- Если возникает ошибка, освободите рабочую память компьютера и запустите заново.

---

### 2. Макрос для Excel (сортировка и транспонирование групп)

**Назначение:**  
Из таблицы вида:  
- столбец A – название группы (повторяется только в первой строке группы);  
- столбец B – номер мыши (не используется);  
- столбец C – числовые значения ДИМ.  

Макрос строит ниже:  
- **первый блок** – группы в столбцах (0,1,2...), значения по возрастанию вниз;  
- **второй блок** – транспонирование: группы в строках (0,1,2...), значения X1..Xn по возрастанию.

#### Как применить макрос в новом Excel-файле (без готового шаблона)

**Шаг 1. Сохраните файл как книгу с поддержкой макросов**  
- Откройте Excel, создайте новый документ или откройте свой файл с данными.  
- Выберите `Файл → Сохранить как` → укажите тип **«Книга Excel с поддержкой макросов (*.xlsm)»**.  

**Шаг 2. Откройте редактор VBA**  
- Нажмите `Alt+F11` (или `Разработчик → Visual Basic`).  

**Шаг 3. Вставьте модуль**  
- В редакторе: `Insert → Module`. Появится пустая область «Module1».  

**Шаг 4. Скопируйте код макроса**  
- Возьмите полный текст макроса `ProcessGroups` (приведён ниже в этом README).  
- Вставьте его в модуль.  

**Шаг 5. Закройте редактор и сохраните файл**  
- `Ctrl+S` (или `Файл → Сохранить`).  

**Шаг 6. Подготовьте данные на листе**  
- Данные должны начинаться с **строки 3** (строки 1–2 могут содержать объединённый заголовок).  
- Столбец A – название группы. Первая строка группы содержит текст, ниже (до следующей группы) – пусто.  
- Столбец B – необязательные идентификаторы (могут быть любые значения).  
- Столбец C – числовые значения (разделитель – запятая или точка).  

Пример:

|     A      |  B  |   C   |
|------------|-----|-------|
| Группа     | №   | ДИМ   |
| (пусто)    | (пусто) | с    |
| Контроль   | 0_1 | 188,48 |
|            | 0_2 | 230,87 |
| Миансерин  | 1_1 | 187,91 |
| ...        | ... | ...   |

**Шаг 7. Запустите макрос**  
- Вернитесь в Excel, активируйте лист с данными.  
- Нажмите `Alt+F8` → выберите `ProcessGroups` → `Выполнить`.  

**Результат:** под исходной таблицей появятся два блока с обработанными данными.

---

### Полный код макроса для Excel (скопировать целиком)

```vba
Sub ProcessGroups()

    Dim ws As Worksheet
    Set ws = ActiveSheet
        
    ' 1. Получение массива из таблицы на активном листе
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row

    Dim dataRange As Range
    Set dataRange = ws.Range("A3:C" & lastRow) ' здесь жестко вшиты столбцы A, B, C
    
    Dim dataArr As Variant
    dataArr = dataRange.Value
    
    ' 2. Создание словаря, где ключ - название группы, значение - список ДИМ
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim groupName As String
    Dim currentGroup As String
    Dim dimValue As Double
    
    currentGroup = ""
    
    For i = 1 To UBound(dataArr, 1)
        groupName = Trim(dataArr(i, 1))
        If groupName <> "" Then
            currentGroup = groupName
        End If
        
        If currentGroup = "" Then
            GoTo NextRow
        End If
        
        Dim rawVal As Variant
        rawVal = dataArr(i, 3)
        
        If IsEmpty(rawVal) Then GoTo NextRow
        If Not IsNumeric(rawVal) Then
            Dim strVal As String
            strVal = Replace(CStr(rawVal), ",", ".")
            If Not IsNumeric(strVal) Then GoTo NextRow
            dimValue = CDbl(strVal)
        Else
            dimValue = CDbl(rawVal)
        End If
        
        If Not dict.exists(currentGroup) Then
            dict.Add currentGroup, New Collection
        End If
        
        dict(currentGroup).Add dimValue
NextRow:
    Next i
    
    ' Временный вывод словаря для отладки
    Dim grp As Variant
    Dim val As Variant
    For Each grp In dict.Keys
        Debug.Print "Группа: " & grp
        For Each val In dict(grp)
            Debug.Print "  " & val
        Next val
    Next grp
    

    ' 3. Сортировка значений внутри каждой группы
    Dim sortedDict As Object
    Set sortedDict = CreateObject("Scripting.Dictionary")
    
    Dim grpKey As Variant
    Dim colVals As Collection
    Dim arrTemp() As Double
    Dim j As Long, k As Long
    Dim maxN As Long
    maxN = 0
    
    For Each grpKey In dict.Keys
        Set colVals = dict(grpKey)
        ReDim arrTemp(1 To colVals.Count)
        For j = 1 To colVals.Count
            arrTemp(j) = colVals(j)
        Next j
        
        Dim temp As Double
        For j = 1 To colVals.Count - 1
            For k = j + 1 To colVals.Count
                If arrTemp(j) > arrTemp(k) Then
                    temp = arrTemp(j)
                    arrTemp(j) = arrTemp(k)
                    arrTemp(k) = temp
                End If
            Next k
        Next j

        sortedDict.Add grpKey, arrTemp
        
        If colVals.Count > maxN Then maxN = colVals.Count
    Next grpKey
    
    ' Вывод отсортированных значений
    For Each grpKey In sortedDict.Keys
        Debug.Print "Sorted " & grpKey
        arrTemp = sortedDict(grpKey)
        For j = 1 To UBound(arrTemp)
            Debug.Print "  " & arrTemp(j)
        Next j
    Next grpKey
    
    Debug.Print "maxN = " & maxN
    
    ' 4. Вывод первого блока (сортировка по столбцам)
    Dim outputRow As Long
    outputRow = lastRow + 3   ' две пустые строки после исходных данных
    
    ' Заголовок блока
    ws.Cells(outputRow, 1).Value = "Группы (сортировка значений по возрастанию)"
    outputRow = outputRow + 1
    
    ' Шапка первого блока
    ws.Cells(outputRow, 1).Value = "№ группы"   ' или можно оставить пустым
    ' Идентификаторы групп (0,1,2...) по горизонтали
    Dim colOffset As Integer
    colOffset = 0
    For Each grpKey In sortedDict.Keys
        ws.Cells(outputRow, 2 + colOffset).Value = colOffset   ' 0,1,2...
        colOffset = colOffset + 1
    Next grpKey
    outputRow = outputRow + 1
    
    ' Вывод значений: строки для каждой позиции (1..maxN)
    Dim rowPos As Long
    For rowPos = 1 To maxN
        colOffset = 0
        For Each grpKey In sortedDict.Keys
            arrTemp = sortedDict(grpKey)
            If rowPos <= UBound(arrTemp) Then
                ws.Cells(outputRow + rowPos - 1, 2 + colOffset).Value = arrTemp(rowPos)
            Else
                ' остается пустая ячейка
            End If
            colOffset = colOffset + 1
        Next grpKey
    Next rowPos
    
    ' 5. Вывод второго блока (транспонированный диапазон)
    outputRow = outputRow + maxN + 1
    ' Заголовок второго блока
    ws.Cells(outputRow, 1).Value = "Транспонированный диапазон (группы по строкам)"
    outputRow = outputRow + 1
    
    ' Заголовки столбцов: № группы, X1, X2, ..., XmaxN
    ws.Cells(outputRow, 1).Value = "№ группы"
    For j = 1 To maxN
        ws.Cells(outputRow, 1 + j).Value = "X" & j
    Next j
    outputRow = outputRow + 1
    
        groupIdx = 0
    For Each grpKey In sortedDict.Keys
        arrTemp = sortedDict(grpKey)   ' отсортированный массив
        ws.Cells(outputRow + groupIdx, 1).Value = groupIdx   ' номер группы
        For j = 1 To UBound(arrTemp)
            ws.Cells(outputRow + groupIdx, 1 + j).Value = arrTemp(j)
        Next j
        ' Ячейки с j > UBound(arrTemp) останутся пустыми
        groupIdx = groupIdx + 1
    Next grpKey

End Sub
```
