Option Explicit On
Imports Inventor
Imports Microsoft.Office.Interop

Public Class Form1
    'структура: исходные данные из excel
    Private Structure AspectData
        'поля, заполняемые через Excel:
        Public text As String 'текст аспекта
        Public valueFromExcel As String 'значение аспекта из Excel
        Public weight As Double 'вес (значимость) аспекта
        Public tolerance As Double 'допустимое отклонение аспекта
        Public interpretation As String 'интрпретация аспекта
        Public comment As String 'комментарий к аспекту

        'поля, заполняемые через Inventor
        Public valueFromInventor As String 'значение аспекта из Inventor
        Public delta As Double 'имеющееся отклонение аспекта

        'доп. техн. поля, не отображаются в таблице
        Public isCorrect As Integer 'правильно ли значение аспекта. 0-нет, 1-полностью правильно, 2-правильно в пределах допустимого отклонения
    End Structure

    'структура имя параметра-значение параметра
    Private Structure PartParameter
        Public name As String
        Public value As String
    End Structure

    'глобальные переменные
    Dim _invApp As Application = Nothing 'приложение Inventor
    Dim _isAppAutoStarted As Boolean = False 'был ли данный сеанс Inventor создан программой
    Dim _openExcelFileDialog As New OpenFileDialog 'диалог выбора файла эксель
    Dim _openInventorFileDialog As New OpenFileDialog 'диалог выбора файла инвентор
    Dim _saveExcelFileDialog As New SaveFileDialog 'диалог сохранения файла эксель
    Dim _conn As OleDb.OleDbConnection 'подключение к источнику данных
    Dim _listAspects As New List(Of AspectData)() 'список для хранения всех данных: из excel, из inventor
    Dim _excelCellsRead As String = "B2:G2" 'какие столбцы считывать из excel. 2-номер строки начала чтения. считывание идет до последней заполненной ячейки
    Dim _countOfExcelСolumns = 6 'количество столбцов, берущих данные из Excel
    Dim _counterForInventorAspects = 0 'счетчик, увеличивающийся при занесении записей из Inventor в _listAspects

    'переменные для сообщений
    Dim _msgLoadInventor As CustomMessage
    Dim _msgGetDataFromExcel As CustomMessage
    Dim _msgGetDataFromInventor As CustomMessage
    Dim _msgDrawDgv As CustomMessage

    'функция запускается, как только форма загружена.
    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        'объявление переменных для сообщений
        _msgLoadInventor = New CustomMessage("Поиск/запуск Inventor")
        _msgGetDataFromExcel = New CustomMessage("Получение данных из Excel")
        _msgGetDataFromInventor = New CustomMessage("Получение данных из Inventor")
        _msgDrawDgv = New CustomMessage("Оформление таблицы и вывода")

        _msgLoadInventor.Show() 'долгий процесс - показать сообщение

        'найти текущий сеанс Inventor (если Inventor не запущен - запустить)
        Try
            'пытаемся получить ссылку на запущенный Inventor
            _invApp = Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application")
        Catch ex As Exception
            'если не удалось получить ссылку (например, Inventor не запущен), то код ниже попытается создать новый сеанс Inventor.
            Try
                Dim invAppType As Type = Type.GetTypeFromProgID("Inventor.Application")
                _invApp = Activator.CreateInstance(invAppType)
                _invApp.Visible = True
                _isAppAutoStarted = True
            Catch ex2 As Exception
                MsgBox(ex2.ToString())
                MsgBox("Не удалось ни найти, ни создать сеанс Inventor")
            End Try
        End Try

        'добавить столбцы к dgvAspects
        dgvAspects.ColumnCount = 8
        'и задать им заголовки и ширину
        Dim standartWindth = dgvAspects.Width / 8
        dgvAspects.Columns(0).HeaderText = "Аспект"
        dgvAspects.Columns(0).Width = standartWindth * 2
        dgvAspects.Columns(1).HeaderText = "Значение (из Excel)"
        dgvAspects.Columns(1).Width = standartWindth * 1.25
        dgvAspects.Columns(2).HeaderText = "Вес аспекта"
        dgvAspects.Columns(2).Width = standartWindth * 0.5
        dgvAspects.Columns(3).HeaderText = "Допустимое отклонение, точность (%)"
        dgvAspects.Columns(3).Width = standartWindth
        dgvAspects.Columns(4).HeaderText = "Интрпретация"
        dgvAspects.Columns(4).Width = standartWindth * 0.5
        dgvAspects.Columns(5).HeaderText = "Комментарий"
        dgvAspects.Columns(5).Width = standartWindth * 0.5
        dgvAspects.Columns(6).HeaderText = "Значение (из Inventor)"
        dgvAspects.Columns(6).Width = standartWindth * 1.25
        dgvAspects.Columns(7).HeaderText = "Имеющееся отклонение (%)"
        dgvAspects.Columns(7).Width = standartWindth

        _msgLoadInventor.Hide() 'закрыть сообщение
    End Sub

    'функция по нажатию кнопки выбора файла эксель
    Private Sub btnGetExcelPath_Click(sender As Object, e As EventArgs) Handles btnGetExcelPath.Click
        'выбрать файл excel
        Dim fullName As String = ""
        Try
            '_openFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            _openExcelFileDialog.RestoreDirectory = True
            _openExcelFileDialog.Title = "Open Excel File"
            _openExcelFileDialog.Filter = "Excel Files(2007)|*.xlsx|Excel Files(2003)|*.xls"

            If _openExcelFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
                Dim fi As New IO.FileInfo(_openExcelFileDialog.FileName)
                fullName = fi.FullName
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            _conn.Close()
        End Try
        tbExcelDirectory.Text = fullName

        'проверка, может ли теперь быть доступна кнопка "считать данные"
        If (Not String.IsNullOrEmpty(tbExcelDirectory.Text) And Not String.IsNullOrEmpty(tbInventorDirectory.Text)) Then
            btnGetData.Enabled = True
        Else
            btnGetData.Enabled = False
        End If
    End Sub

    'функция по нажатию кнопки выбора файла из инвентор
    Private Sub btnGetInventorPath_Click(sender As Object, e As EventArgs) Handles btnGetInventorPath.Click
        'выбрать фаил сборки
        Dim fullName As String = ""
        Try
            '_openFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            _openInventorFileDialog.RestoreDirectory = True
            _openInventorFileDialog.Title = "Open Assembly File"
            _openInventorFileDialog.Filter = "Файл сборки|*.iam"

            If _openInventorFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
                Dim fi As New IO.FileInfo(_openInventorFileDialog.FileName)
                fullName = fi.FullName
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            _conn.Close()
        End Try
        tbInventorDirectory.Text = fullName

        'проверка, может ли теперь быть доступна кнопка "считать данные"
        If (Not String.IsNullOrEmpty(tbExcelDirectory.Text) And Not String.IsNullOrEmpty(tbInventorDirectory.Text)) Then
            btnGetData.Enabled = True
        Else
            btnGetData.Enabled = False
        End If
    End Sub

    'функция по нажатию кнопки получить данные
    Private Sub btnGetData_Click(sender As Object, e As EventArgs) Handles btnGetData.Click
        '1. Очистка _listAspects и dgvAspects
        _listAspects.Clear() 'перед заполнением _listAspects надо очистить
        dgvAspects.Rows.Clear() 'перед заполнением dgv надо очистить
        lblCountOfRows.Text = dgvAspects.RowCount() 'обновить текст количества строк

        '2. Получение данных из эксель, их запись в _listAspects
        _msgGetDataFromExcel.Show() 'долгий процесс - показать сообщение
        Dim exl As New Excel.Application
        Dim exlSheet As Excel.Worksheet
        Try
            exl.Workbooks.Open(tbExcelDirectory.Text) 'открыть документ
            exlSheet = exl.Workbooks(1).Worksheets(1) 'Переходим к первому листу

            Dim a(,) As Object
            a = exlSheet.Range(_excelCellsRead, exlSheet.Range(_excelCellsRead).End(Excel.XlDirection.xlDown)).Value 'Теперь вспомогательный массив a содержит таблицу из excel

            'закрыть документ excel - больше не нужен
            exl.Quit()
            exlSheet = Nothing
            exl = Nothing

            Dim countOfA As Integer = a.Length 'количество всех элементов массива a
            Dim countOfRowsInA As Integer = countOfA / _countOfExcelСolumns 'количество строк в a

            'проход по всем строкам массива a, что бы переписать их в _listAspects. записаны будут только не пустые значения
            For i As Integer = 1 To countOfRowsInA
                If a(i, 1) IsNot Nothing Then 'если поле текста аспекта существует
                    If a(i, 1) IsNot "" Then 'если поле текста аспекта не пустое
                        Dim ed As AspectData = Nothing

                        ed.text = a(i, 1)
                        ed.valueFromExcel = a(i, 2)

                        Try
                            ed.weight = Convert.ToDouble(a(i, 3))
                        Catch ex As Exception
                            ed.weight = 0
                        End Try

                        Try
                            ed.tolerance = Convert.ToDouble(a(i, 4))
                        Catch ex As Exception
                            ed.tolerance = 0
                        End Try

                        ed.interpretation = a(i, 5)
                        ed.comment = a(i, 6)

                        'все значения из excel добавлены, но нельзя оставлять оставшиеся поля (inventor-поля и техн. поля со значением Nothing)
                        ed.valueFromInventor = ""
                        ed.delta = 0
                        ed.isCorrect = 0

                        _listAspects.Add(ed)
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox("Не удалось открыть выбранный документ эксель")
        End Try
        _msgGetDataFromExcel.Hide() 'закрыть сообщение

        '3. Получение данных из инвентор, их запись в _listAspects
        _msgGetDataFromInventor.Show() 'долгий процесс - показать сообщение
        If _listAspects.Count > 0 Then 'получение данных из Inventor имеет смысл/возможно только при наличии считанного списка аспектов 
            'открыть существующий документ сборки по указанному пути
            Dim asmDoc As Document = _invApp.Documents.Open(tbInventorDirectory.Text)

            'продолжать работу с Inventor можно, если: открыт хотя бы 1 документ, и тип открытого документа - Assembly
            If (_invApp.Documents.Count > 0) And (_invApp.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject) Then
                'получение свойств из Inventor
                _counterForInventorAspects = 0 'нужно сбросить счетчик аспектов в изначальное состояние

                '---assembly 's---
                'Получить данные самой сборки
                getAsm(asmDoc)

                '---parts and drawings---
                'пройти по всем деталям, которые есть в сборке
                For i As Integer = asmDoc.AllReferencedDocuments.Count To 1 Step -1
                    'для каждой детали получить данные
                    getPart(asmDoc.AllReferencedDocuments(i))

                    'для каждой детали определить, есть ли для нее чертеж. если есть, получить данные чертежа
                    Dim drawingFullFileName As String = findDrawingFullFileNameForDocument(asmDoc.AllReferencedDocuments(i)) 'найти, если возможно, путь к чертежу детали
                    'если путь к чертежу найден, инициализировать переменную чертежа и открыть чертеж
                    If Not String.IsNullOrEmpty(drawingFullFileName) Then
                        Dim drawingDoc As Document = _invApp.Documents.Open(drawingFullFileName) 'открыть чертеж
                        getDrawing(drawingDoc) 'получить данные чертежа
                    End If
                Next
            Else
                MsgBox("Ошибка: не удалось открыть документ детали. Проверьте наличие документа с расширением .ipt по выбранному пути")
            End If

            'ТЕСТОВАЯ функция эскизов, потом удалить
            'For Each pd As PartDocument In asmDoc.AllReferencedDocuments
            '    getInfoAboutSketches(pd)
            'Next
        Else
            MsgBox("Данные из Excel не были получены. Получение данных из Inventor невозможно.")
        End If
        _msgGetDataFromInventor.Hide() 'закрыть сообщение

        '4. Запись _listAspects в dgvAspects
        'пройти весь _listAspects, и каждый его элемент (содержащий 8 предназначенных для отображения в таблице полей) переписать в dgv
        _msgDrawDgv.Show() 'долгий процесс - показать сообщение
        For Each d As AspectData In _listAspects
            dgvAspects.Rows.Add(d.text, d.valueFromExcel, d.weight, d.tolerance, d.interpretation, d.comment, d.valueFromInventor, d.delta)
        Next
        lblCountOfRows.Text = dgvAspects.RowCount() 'обновить текст количества строк  

        '5 Отметить в dgvAspects правильность ответов цветом, подвести итоги
        summarize()
        _msgDrawDgv.Hide() 'закрыть сообщение
    End Sub

    'функция по нажатию кнопки "очистить таблицу"
    Private Sub btnClearTable_Click(sender As Object, e As EventArgs) Handles btnClearTable.Click
        Dim result As Integer = MessageBox.Show("Вы действительно хотите очистить таблицу?", "Подтверждение действия", MessageBoxButtons.OKCancel)
        If result = DialogResult.Cancel Then
            'отмена: ничего не делать
        ElseIf result = DialogResult.OK Then
            'да: действие подтверждено
            dgvAspects.Rows.Clear()
            _listAspects.Clear()
            lblCountOfRows.Text = dgvAspects.RowCount() 'обновить текст количества строк
        End If
    End Sub

    'функция по нажатию кнопки "экспорт таблицы"
    Private Sub btnExportTable_Click(sender As Object, e As EventArgs) Handles btnExportTable.Click
        'проверка, имеются ли данные в dgv
        If (dgvAspects.RowCount() = 0) Then
            MsgBox("Таблица пуста, экспорт невозможен")
            Return 'выход из функции обработчика кнопки
        End If

        'выбрать место сохрания файла excel
        '_saveExcelFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
        _saveExcelFileDialog.RestoreDirectory = True
        _saveExcelFileDialog.Filter = "Excel Files(2007)|*.xlsx|Excel Files(2003)|*.xls"
        _saveExcelFileDialog.Title = "Save Excel File"
        _saveExcelFileDialog.ShowDialog()

        'если получена директория
        If (Not String.IsNullOrEmpty(_saveExcelFileDialog.FileName)) Then
            'Создание dataset для экспорта
            Dim dset As New DataSet
            dset.Tables.Add() 'добавить таблицу

            'добавление столбцов в эту таблицу
            For i As Integer = 0 To dgvAspects.ColumnCount - 1
                dset.Tables(0).Columns.Add(dgvAspects.Columns(i).HeaderText)
            Next

            'добавление строк в эту таблицу
            Dim dr1 As DataRow
            For i As Integer = 0 To dgvAspects.RowCount - 1
                dr1 = dset.Tables(0).NewRow
                For j As Integer = 0 To dgvAspects.Columns.Count - 1
                    dr1(j) = dgvAspects.Rows(i).Cells(j).Value
                Next
                dset.Tables(0).Rows.Add(dr1)
            Next

            Dim exl As New Excel.Application
            Dim exlBook As Excel.Workbook
            Dim exlSheet As Excel.Worksheet

            exlBook = exl.Workbooks.Add()
            exlSheet = exlBook.ActiveSheet()

            Dim dt As DataTable = dset.Tables(0)
            Dim dc As DataColumn
            Dim dr As DataRow
            Dim colIndex As Integer = 0
            Dim rowIndex As Integer = 0

            For Each dc In dt.Columns
                colIndex = colIndex + 1
                exl.Cells(1, colIndex) = dc.ColumnName
            Next

            For Each dr In dt.Rows
                rowIndex = rowIndex + 1
                colIndex = 0
                For Each dc In dt.Columns
                    colIndex = colIndex + 1
                    exl.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
                Next
            Next

            exlSheet.Columns.AutoFit()
            Dim strFileName As String = _saveExcelFileDialog.FileName
            Dim blnFileOpen As Boolean = False
            Try
                Dim fileTemp As System.IO.FileStream = System.IO.File.OpenWrite(strFileName)
                fileTemp.Close()
            Catch ex As Exception
                blnFileOpen = False
            End Try

            If System.IO.File.Exists(strFileName) Then
                System.IO.File.Delete(strFileName)
            End If

            exlBook.SaveAs(strFileName)
            exl.Workbooks.Open(strFileName)
            exl.Visible = True
        End If
    End Sub

    'функция по закрытию формы
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Dim result As Integer = MessageBox.Show("Вы действительно хотите выйти? Если Inventor был запущен программно, он будет закрыт", "Подтверждение действия", MessageBoxButtons.OKCancel)
        If result = DialogResult.Cancel Then
            'отмена: не закрывать форму
            e.Cancel = True
        ElseIf result = DialogResult.OK Then
            'подтверждение: закрыть форму
            e.Cancel = False
        End If
    End Sub

    'функция запускается при случившемся закрытии формы.
    Private Sub Form1_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        ' Закроем сеанс Inventor, если он был создан при создании формы
        If _isAppAutoStarted Then
            _invApp.Quit()
        End If

        _isAppAutoStarted = Nothing 'очистить переменную
    End Sub


    'Вспомогательные функции
    'вспомогательная функция: пройти по _listAspects и по dgv, отметить цветом правильность аспектов, подсчитать итоги
    Private Sub summarize()
        If (dgvAspects.RowCount() = 0) Then
            MsgBox("Таблица пуста. Необходимо считать данные из эксель и Inventor")
            Return 'выход из функции обработчика кнопки
        End If

        Dim correct As Boolean = True
        Dim errors As Integer = 0
        Dim total_points As Double = 0
        Dim style_wrong As New DataGridViewCellStyle
        style_wrong.BackColor = Drawing.Color.LightCoral
        Dim style_full_right As New DataGridViewCellStyle
        style_full_right.BackColor = Drawing.Color.LightGreen
        Dim style_right As New DataGridViewCellStyle
        style_right.BackColor = Drawing.Color.DarkSeaGreen

        For i = 0 To (_listAspects.Count - 1)
            Select Case _listAspects(i).isCorrect
                Case 0
                    'ответ не верный
                    correct = False
                    errors += 1
                    dgvAspects.Rows(i).DefaultCellStyle = style_wrong
                Case 1
                    'если значения точно совпадают, ответ верный
                    dgvAspects.Rows(i).DefaultCellStyle = style_full_right
                    total_points += _listAspects(i).weight
                Case 2
                    'ответ верный, в пределах отклонения (но не точный)
                    dgvAspects.Rows(i).DefaultCellStyle = style_right
                    total_points += _listAspects(i).weight
            End Select
        Next

        If (correct = True) Then
            MsgBox("Не найдено ни одной ошибки" & vbCrLf & "Всего набрано баллов: " & total_points)
        Else
            MsgBox("Найдено " & errors & " ошибок" & vbCrLf & "Всего набрано баллов: " & total_points)
        End If
    End Sub

    'вспомогательная функция найти чертеж к документу: сборке (assembly) или детали (part). если чертеж не найден, возвращает пустую строку: ""
    Private Function findDrawingFullFileNameForDocument(ByVal doc As Document) As String
        Try
            Dim fullFilename As String = doc.FullFileName

            'переменная drawingFilename будет хранить полное имя чертежа для сборки / детали
            Dim drawingFilename As String = ""

            ' Extract the path from the full filename.
            Dim path As String = Microsoft.VisualBasic.Left$(fullFilename, InStrRev(fullFilename, "\"))

            ' Extract the filename from the full filename.
            Dim filename As String = Microsoft.VisualBasic.Right$(fullFilename, Len(fullFilename) - InStrRev(fullFilename, "\"))

            ' Replace the extension with "dwg"
            filename = Microsoft.VisualBasic.Left$(filename, InStrRev(filename, ".")) & "dwg"
            ' Find if the drawing exists.
            drawingFilename = _invApp.DesignProjectManager.ResolveFile(path, filename)

            ' Check the result.
            If drawingFilename = "" Then
                ' Try again with idw extension.
                filename = Microsoft.VisualBasic.Left$(filename, InStrRev(filename, ".")) & "idw"
                ' Find if the drawing exists.
                drawingFilename = _invApp.DesignProjectManager.ResolveFile(path, filename)
            End If

            ' Return result. Если не найден чертеж, вернет пустую строку ""
            Return drawingFilename
        Catch ex As Exception
            MsgBox("Ошибка: невозможно найти чертеж для документа" & vbCrLf & ex.ToString)
            Return ""
        End Try
    End Function

    'вспомогательная функция: проверить видимость 2d эскизов и объектов вспомогательной геометрии (плоскости, оси, точки). true - они все невидимы, false - есть как минимум 1 видимый объект
    Private Function isOriginsInvisible(ByVal oDoc As Document) As Boolean
        Dim isInvisible As Boolean = True

        ' получть все 2d эскизы детали и проверить их видимость
        Dim oSketches As PlanarSketches = oDoc.ComponentDefinition.Sketches
        For Each oSketch In oSketches
            If oSketch.Visible = True Then
                isInvisible = False
                Return isInvisible 'выход из всей функции 
            End If
        Next

        'look at the WorkPlanes collection (все плоскости документа)
        For Each oWorkPlane In oDoc.ComponentDefinition.WorkPlanes
            If oWorkPlane.Visible = True Then
                isInvisible = False
                Return isInvisible 'выход из всей функции 
            End If
        Next

        'look at the WorkAxes collection (все оси документа)
        For Each oWorkAxe In oDoc.ComponentDefinition.WorkAxes
            If oWorkAxe.Visible = True Then
                isInvisible = False
                Return isInvisible 'выход из всей функции 
            End If
        Next

        'look at the WorkPoints collection (все точки документа)
        For Each oWorkPoint In oDoc.ComponentDefinition.WorkPoints
            If oWorkPoint.Visible = True Then
                isInvisible = False
                Return isInvisible 'выход из всей функции 
            End If
        Next

        'look at the WorkSurfaces collection (все поверхности(?) документа)
        'For Each oWorkSurface In oDoc.ComponentDefinition.WorkSurfaces
        '    If oWorkSurface.Visible = False Then
        '        MsgBox(oWorkSurface.Name & " Visible false: ok")
        '    Else
        '        MsgBox(oWorkSurface.Name & "Visible true: not ok")
        '    End If
        'Next        
        Return isInvisible
    End Function

    'вспомогательная функция: получить таблицу свойств детали (доступ к таблице параметров)
    Private Function getParametersFromPart(ByVal partDoc As Document) As List(Of PartParameter)
        Dim listOfParameters As New List(Of PartParameter)() 'список параметров документа

        Dim allParams As Parameters = partDoc.ComponentDefinition.Parameters
        If allParams.Count > 0 Then
            For Each param As Parameter In allParams
                Dim partParameter As PartParameter
                partParameter.name = param.Name
                partParameter.value = (param.ModelValue * 10).ToString
                listOfParameters.Add(partParameter)
            Next
        End If

        Return listOfParameters
    End Function

    'вспомогательная функция: вернуть из структуры типа PartParameter значение по имени
    Private Function findValueInPartParamListByName(ByVal name As String, ByVal list As List(Of PartParameter)) As String
        Dim value As String = ""

        For Each elem As PartParameter In list
            If elem.name = name Then
                'взять значение по модулю
                If CDec(elem.value) < 0 Then
                    elem.value = Math.Abs(CDec(elem.value))
                End If

                value = elem.value.ToString
                Exit For
            End If
        Next

        Return value
    End Function

    'вспомогательная функция: 
    '1)получить значения inventor-столбцов
    '2)получить значение столбца isCorrect: определить, верно ли значение
    '3)добавить в структуру типа AspectData значение с полученными столбцами
    Private Sub addInventorValuesInAspectsList(ByVal value As String)
        Dim aspect As AspectData = Nothing
        aspect = _listAspects(_counterForInventorAspects) 'получить текущее значение счетчика записей

        'столбец valueFromInventor
        aspect.valueFromInventor = value
        'столбец delta. подсчет дельты (формула?)
        Try
            If aspect.valueFromExcel > aspect.valueFromInventor Then
                aspect.delta = (Math.Abs(aspect.valueFromExcel - aspect.valueFromInventor) / aspect.valueFromExcel) * 100
            ElseIf aspect.valueFromExcel < aspect.valueFromInventor Then
                aspect.delta = (Math.Abs(aspect.valueFromExcel - aspect.valueFromInventor) / aspect.valueFromInventor) * 100
            Else
                aspect.delta = 0
            End If
        Catch ex As Exception
            aspect.delta = 0
        End Try
        'столбец isCorrect
        If (aspect.valueFromExcel = aspect.valueFromInventor) Then
            'если значения точно совпадают, ответ верный
            aspect.isCorrect = 1
        Else
            Dim valueFromInventor, valueFromExcel As Double
            'значения не совпадают, необходимо проверить точность (если возможно)
            If Double.TryParse(aspect.valueFromInventor, valueFromInventor) And Double.TryParse(aspect.valueFromExcel, valueFromExcel) Then
                'если допустимое отклонение больше, чем текущее отклонение, ответ верный, в пределах отклонения                
                If (aspect.tolerance > aspect.delta) Then
                    'ответ верный, в пределах отклонения (но не точный)
                    aspect.isCorrect = 2
                Else
                    'ответ не верный
                    aspect.isCorrect = 0
                End If
            Else
                'значение - не число, нет смысла проверять точность, ответ неверный
                aspect.isCorrect = 0
            End If
        End If

        'новое значение aspect, включающее данные из Inventor, нужно вставить вместо старого значения
        _listAspects(_counterForInventorAspects) = aspect
        'увеличить счетчик считанных свойств из Inventor
        _counterForInventorAspects += 1
    End Sub

    'вспомогательная функция: получить параметры резьбы
    Private Function getThreadsParams(ByVal partDoc As Document) As String
        Dim resultString As String = ""

        Dim fc As Face
        For Each fc In partDoc.ComponentDefinition.SurfaceBodies.Item(1).Faces
            If fc.SurfaceType = Inventor.SurfaceTypeEnum.kCylinderSurface Or fc.SurfaceType = Inventor.SurfaceTypeEnum.kConeSurface Then
                If Not fc.ThreadInfos Is Nothing Then
                    If fc.ThreadInfos.Count > 0 Then
                        Dim thread As ThreadInfo
                        For Each thread In fc.ThreadInfos
                            resultString = "" ' !пока берется последняя резьба, старые рез-ы очищ.

                            Dim threadDesignation As String = thread.ThreadDesignation 'designation (пример: М10х1.5)
                            threadDesignation = Replace(threadDesignation, "M", "М") 'заменить английскую букву M на русскую
                            threadDesignation = Replace(threadDesignation, "x", "х") 'заменить английскую букву x на русскую
                            threadDesignation = Replace(threadDesignation, ".", ",") 'заменить точку на запятую
                            resultString &= threadDesignation
                            resultString &= "-"

                            If TypeOf thread Is StandardThreadInfo Then
                                resultString &= thread.Class 'class (пример: 6H)
                            End If
                        Next
                    End If
                End If
            End If
        Next

        Return resultString
    End Function

    'Функции получения данных из сборки, деталей, чертежей
    'вспомогательная функция: получение параметров сборки
    Private Sub getAsm(ByVal asmDoc As Document)
        ' Get the PropertySets object.
        Dim oPropSets As PropertySets = asmDoc.PropertySets

        Dim oPropSetISI As PropertySet = oPropSets.Item("Inventor Summary Information") ' Get the Inventor Summary Information property set.
        Dim oPropSetIDSI As PropertySet = oPropSets.Item("Inventor Document Summary Information") ' Get the Inventor Document Summary Information property set.
        Dim oPropSetDTP As PropertySet = oPropSets.Item("Design Tracking Properties") ' Get the Design Tracking Properties property set.

        '"Автор"
        addInventorValuesInAspectsList(oPropSetISI.Item("Author").Value)

        '"Имя документа"
        addInventorValuesInAspectsList(oPropSetDTP.Item("Part Number").Value)

        '"Название сборки"
        addInventorValuesInAspectsList(oPropSetDTP.Item("Description").Value)

        '"Материал"
        addInventorValuesInAspectsList(oPropSetDTP.Item("Material").Value)

        '"Дата создания фаила"
        Dim filespec As String = asmDoc.File.FullFileName 'получить полное имя фаила
        Dim fs, f
        fs = CreateObject("Scripting.FileSystemObject")
        f = fs.GetFile(filespec)
        addInventorValuesInAspectsList(f.DateCreated.ToString)

        '"Дата изменения фаила"
        addInventorValuesInAspectsList(f.DateLastModified.ToString)

        '"Количество деталей в сборке"
        addInventorValuesInAspectsList(asmDoc.AllReferencedDocuments.Count.ToString)

        '"Все детали сборки закреплены (0 степеней свободы)"
        Dim occ As ComponentOccurrence
        Dim result As String = True
        For Each occ In asmDoc.ComponentDefinition.Occurrences 'occ - свойства part document (1..n) В assembly, их (документов) перебор
            result = occ.Grounded 'true - да, деталь закреплена, false - нет, деталь не закреплена
            If result = False Then
                Exit For
            End If
        Next
        addInventorValuesInAspectsList(result)

        '"Масса сборки"
        Dim massProps As MassProperties = asmDoc.ComponentDefinition.MassProperties
        addInventorValuesInAspectsList(massProps.Mass)
        'дополнит. способ (к MassProperties):
        'Dim massProps As MassProperties = asmDoc.ComponentDefinition.MassProperties
        'Dim uom As UnitsOfMeasure = asmDoc.UnitsOfMeasure
        'Dim defaultLength As String = uom.GetStringFromType(uom.LengthUnits)
        'MsgBox(uom.GetStringFromValue(massProps.Volume, defaultLength & "^3"))

        '"Площадь сборки"
        addInventorValuesInAspectsList(massProps.Area * 100)

        '"Объем сборки"
        addInventorValuesInAspectsList(massProps.Volume * 1000)
    End Sub

    'вспомогательная функция: получение параметров детали
    Private Sub getPart(ByVal partDoc As Document)
        ' Get the PropertySets object.
        Dim oPropSets As PropertySets = partDoc.PropertySets

        Dim oPropSetISI As PropertySet = oPropSets.Item("Inventor Summary Information") ' Get the Inventor Summary Information property set.
        Dim oPropSetIDSI As PropertySet = oPropSets.Item("Inventor Document Summary Information") ' Get the Inventor Document Summary Information property set.
        Dim oPropSetDTP As PropertySet = oPropSets.Item("Design Tracking Properties") ' Get the Design Tracking Properties property set.

        '"Имя документа"
        'или partDoc.DisplayName
        addInventorValuesInAspectsList(oPropSetDTP.Item("Part Number").Value)

        '"Название детали"
        addInventorValuesInAspectsList(oPropSetDTP.Item("Description").Value)

        '"Материал"
        addInventorValuesInAspectsList(oPropSetDTP.Item("Material").Value)

        '"Дата создания фаила"
        Dim filespec As String = partDoc.File.FullFileName 'получить полное имя фаила
        Dim fs, f
        fs = CreateObject("Scripting.FileSystemObject")
        f = fs.GetFile(filespec)
        addInventorValuesInAspectsList(f.DateCreated.ToString)

        '"Дата изменения фаила"
        addInventorValuesInAspectsList(f.DateLastModified.ToString)

        '"Деталь твердотельная (не поверхности)"
        Dim SrfBods As SurfaceBodies = partDoc.ComponentDefinition.SurfaceBodies
        Dim b As Boolean
        For Each SrfBod In SrfBods
            b = SrfBod.IsSolid '? значение последнего surface body ?
        Next
        addInventorValuesInAspectsList(b)

        '"Деталь состоит из одного твердого тела"
        Dim countOfSolidBody As Integer = 0
        Dim oCompDef As ComponentDefinition = partDoc.ComponentDefinition
        For Each SurfaceBody In oCompDef.SurfaceBodies
            countOfSolidBody += 1
        Next
        If countOfSolidBody = 1 Then
            addInventorValuesInAspectsList(True) 'true - да, из одного
        Else
            addInventorValuesInAspectsList(False) 'false - нет, не из одного
        End If

        '"Все эскизы (2D и 3D) и объекты вспомогательной геометрии (плоскости, оси, точки) невидимы"
        addInventorValuesInAspectsList(isOriginsInvisible(partDoc)) 'записать true - да, невидимый; false - видимый

        '"Все эскизы детали должны быть полностью определены"
        Dim isOk As Boolean = True
        Dim partDef As PartComponentDefinition = partDoc.ComponentDefinition
        'пройти по всем эскизам детали
        For Each sketch As Sketch In partDef.Sketches
            'является ли эскиз полностью определенным? если нет, то записывем ошибку
            If sketch.ConstraintStatus <> ConstraintStatusEnum.kFullyConstrainedConstraintStatus Then
                isOk = False
                Exit For
            End If
        Next
        addInventorValuesInAspectsList(isOk) 'записать true - все эскизы детали полностью определены; false - хотя бы один эскиз детали не полностью определен

        '"Масса детали"
        Dim massProps As MassProperties = partDoc.ComponentDefinition.MassProperties
        addInventorValuesInAspectsList(massProps.Mass)

        '"Площадь детали"
        addInventorValuesInAspectsList(massProps.Area * 100)

        '"Объем детали"
        addInventorValuesInAspectsList(massProps.Volume * 1000)
    End Sub

    'вспомогательная функция: получение параметров чертежа
    Private Sub getDrawing(ByVal drawingDoc As Document)
        Dim oSheet As Sheet = drawingDoc.Sheets.Item(1) 'лист чертежа
        Dim oView As DrawingView = oSheet.DrawingViews.Item(1) 'вид листа     

        '"Выбор формата листа"
        Dim result As String = "EMPTY VALUE"
        If oSheet.Size = DrawingSheetSizeEnum.kA4DrawingSheetSize Then
            result = "А4"
        ElseIf oSheet.Size = DrawingSheetSizeEnum.kA3DrawingSheetSize Then
            result = "А3"
        Else
            result = "Другой формат"
        End If
        addInventorValuesInAspectsList(result)

        '"Выбор масштаба главного вида"
        result = "EMPTY VALUE"
        Dim oPropSets As PropertySets = drawingDoc.PropertySets
        Dim oPropSetGOST As PropertySet = oPropSets.Item("Свойства ГОСТ")
        result = oPropSetGOST.Item("Масштаб").Value
        addInventorValuesInAspectsList(result)

        '"Заполнение основной надписи"
        Dim author As String = Nothing
        Dim designation As String = Nothing
        Dim header As String = Nothing

        Dim oTitleBlock As TitleBlock = oSheet.TitleBlock
        For Each tb As TextBox In oTitleBlock.Definition.Sketch.TextBoxes
            If tb.Text = "<АВТОР>" Then
                author = oTitleBlock.GetResultText(tb)
            End If
            If tb.Text = "<ОБОЗНАЧЕНИЕ>" Then
                designation = oTitleBlock.GetResultText(tb)
            End If
            If tb.Text = "<ЗАГОЛОВОК>" Then
                header = oTitleBlock.GetResultText(tb)
            End If
        Next
        'если одна из строк пустая - ошибка, основная надпись не заполнена
        If (String.IsNullOrEmpty(author) Or String.IsNullOrEmpty(designation) Or String.IsNullOrEmpty(header)) Then
            addInventorValuesInAspectsList(False)
        Else
            addInventorValuesInAspectsList(True)
        End If
    End Sub

    'ДОПОЛНИТЕЛЬНАЯ ТЕСТОВАЯ функция получения информации об эскизах
    Private Sub getInfoAboutSketches(ByVal partDoc As Document)
        Dim finalString As String = ""
        finalString &= "----------" & vbCrLf
        finalString &= "Имя детали: " & partDoc.DisplayName & vbCrLf
        finalString &= "Всего содержит эскизов: " & partDoc.ComponentDefinition.Sketches.Count & vbCrLf
        finalString &= "----------" & vbCrLf & vbCrLf

        For Each oSketch As Sketch In partDoc.ComponentDefinition.Sketches
            finalString &= "Имя эскиза: " & oSketch.Name.ToString & vbCrLf
            finalString &= " ConstraintStatus: " & oSketch.ConstraintStatus.ToString & vbCrLf
            'finalString &= " Color: " & oSketch.Color.ToString & vbCrLf
            finalString &= " LineType: " & oSketch.LineType.ToString & vbCrLf
            finalString &= " LineWeight: " & oSketch.LineWeight.ToString & vbCrLf
            finalString &= " Type: " & oSketch.Type.ToString & vbCrLf
            finalString &= " Visible: " & oSketch.Visible.ToString & vbCrLf

            finalString &= "  Всего AttributeSets в текущем эскизе: " & oSketch.AttributeSets.Count & vbCrLf
            For Each oAttributeSet As AttributeSet In oSketch.AttributeSets
                finalString &= "   Name: " & oAttributeSet.Name & vbCrLf
            Next

            finalString &= "  Всего DimensionConstraints в текущем эскизе: " & oSketch.DimensionConstraints.Count & vbCrLf
            For Each oDimensionConstraint As DimensionConstraint In oSketch.DimensionConstraints
                Select Case oDimensionConstraint.Type
                    Case ObjectTypeEnum.kArcLengthDimConstraintObject
                        finalString &= "   ArcLengthDimConstraint" & vbCrLf
                    Case ObjectTypeEnum.kDiameterDimConstraintObject
                        finalString &= "   DiameterDimConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kEllipseRadiusDimConstraintObject
                        finalString &= "   EllipseRadiusDimConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kOffsetDimConstraintObject
                        finalString &= "   OffsetDimConstraint" & vbCrLf
                    Case ObjectTypeEnum.kRadiusDimConstraintObject
                        finalString &= "   RadiusDimConstraint" & vbCrLf
                    Case ObjectTypeEnum.kTangentDistanceDimConstraintObject
                        finalString &= "   kTangentDistanceDimConstraint" & vbCrLf
                    Case ObjectTypeEnum.kThreePointAngleDimConstraintObject
                        finalString &= "   ThreePointAngleDimConstraint" & vbCrLf
                    Case ObjectTypeEnum.kTwoLineAngleDimConstraintObject
                        finalString &= "   TwoLineAngleDimConstraint" & vbCrLf
                    Case ObjectTypeEnum.kTwoPointDistanceDimConstraintObject
                        finalString &= "   TwoPointDistanceDimConstraint" & vbCrLf
                    Case Else
                        finalString &= "   Неизвестно" & vbCrLf
                End Select
            Next

            finalString &= "  Всего GeometricConstraints в текущем эскизе: " & oSketch.GeometricConstraints.Count & vbCrLf
            For Each oGeometricConstraint As GeometricConstraint In oSketch.GeometricConstraints
                Select Case oGeometricConstraint.Type
                    Case ObjectTypeEnum.kCoincidentConstraintObject
                        finalString &= "   CoincidentConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kCollinearConstraintObject
                        finalString &= "   CollinearConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kConcentricConstraintObject
                        finalString &= "   ConcentricConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kEqualLengthConstraintObject
                        finalString &= "   EqualLengthConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kEqualRadiusConstraintObject
                        finalString &= "   EqualRadiusConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kGroundConstraintObject
                        finalString &= "   GroundConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kHorizontalAlignConstraintObject
                        finalString &= "   HorizontalAlignConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kHorizontalConstraintObject
                        finalString &= "   HorizontalConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kMidpointConstraintObject
                        finalString &= "   MidpointConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kOffsetConstraintObject
                        finalString &= "   OffsetConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kParallelConstraintObject
                        finalString &= "   ParallelConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kPatternConstraintObject
                        finalString &= "   PatternConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kPerpendicularConstraintObject
                        finalString &= "   PerpendicularConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kSmoothConstraintObject
                        finalString &= "   SmoothConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kSplineFitPointConstraintObject
                        finalString &= "   SplineFitPointConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kSymmetryConstraintObject
                        finalString &= "   SymmetryConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kTangentSketchConstraintObject
                        finalString &= "   TangentSketchConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kVerticalAlignConstraintObject
                        finalString &= "   VerticalAlignConstraintObject" & vbCrLf
                    Case ObjectTypeEnum.kVerticalConstraintObject
                        finalString &= "   VerticalConstraintObject" & vbCrLf
                    Case Else
                        finalString &= "   Неизвестно" & vbCrLf
                End Select
            Next

            finalString &= "  Всего SketchArcs в текущем эскизе: " & oSketch.SketchArcs.Count & vbCrLf
            For Each oSketchArc As SketchArc In oSketch.SketchArcs
                finalString &= "   Length: " & oSketchArc.Length.ToString & vbCrLf
            Next

            finalString &= "  Всего SketchCircles в текущем эскизе: " & oSketch.SketchCircles.Count & vbCrLf
            For Each oSketchCircle As SketchCircle In oSketch.SketchCircles
                finalString &= "   Area: " & oSketchCircle.Area.ToString & vbCrLf
            Next

            '... добавить остальные массивы SketchФигура ...

            finalString &= "  Всего TextBoxes в текущем эскизе: " & oSketch.TextBoxes.Count & vbCrLf
            For Each oTextBox As TextBox In oSketch.TextBoxes
                finalString &= "   Text: " & oTextBox.Text & vbCrLf
            Next

            'Profiles (продолжить здесь)
            finalString &= "  Всего Profiles в текущем эскизе: " & oSketch.Profiles.Count & vbCrLf
            For Each oProfile As Profile In oSketch.Profiles
                finalString &= "   Count (the number of items in this collection): " & oProfile.Count & vbCrLf

                finalString &= "  Всего AttributeSets в текущем Profile: " & oProfile.AttributeSets.Count & vbCrLf
                For Each oAttributeSet As AttributeSet In oProfile.AttributeSets
                    finalString &= "   Name: " & oAttributeSet.Name & vbCrLf
                Next

                'profilepath
                finalString &= "  ProfilePath"
                Dim oProfilePath As ProfilePath
                For Each oProfilePath In oProfile

                    Dim oProfileEntity As ProfileEntity
                    For Each oProfileEntity In oProfilePath
                        finalString &= "   Type: " & oProfileEntity.Type & vbCrLf
                    Next

                    'Dim oTextBox As TextBox
                    'For Each oTextBox In oProfilePath
                    '    finalString &= "   Text: " & oTextBox.Text & vbCrLf
                    'Next
                Next

                'region properties
                finalString &= "   Region properties" & vbCrLf
                finalString &= "   Accuracy: " & oProfile.RegionProperties.Accuracy.ToString & vbCrLf
                finalString &= "   Area: " & oProfile.RegionProperties.Area.ToString & vbCrLf
                finalString &= "   Centroid: " & oProfile.RegionProperties.Centroid.ToString & vbCrLf
                finalString &= "   Perimeter: " & oProfile.RegionProperties.Perimeter.ToString & vbCrLf
                finalString &= "   RotationAngle: " & oProfile.RegionProperties.RotationAngle.ToString & vbCrLf
                finalString &= "   Type: " & oProfile.RegionProperties.Type.ToString & vbCrLf

                'wires
                finalString &= "  Всего Wires в текущем Profile: " & oProfile.Wires.Count & vbCrLf
                For Each oWire As Wire In oProfile.Wires
                    finalString &= "   Type: " & oWire.Type.ToString & vbCrLf
                Next
            Next
        Next

        My.Computer.FileSystem.WriteAllText("C:\Users\Сергей\Desktop\sketches_info.txt", finalString, True)
    End Sub


    'НЕАКТУЛЬНАЯ вспомогательная функция: заменить в структуре типа AspectData значение, получаемое из Inventor (2 последних столбца)
    'неактуальная потому, что имена аспектов не уникальные
    'Private Sub changeValueFromInventorInAspectDataList(ByVal text As String, ByVal value As String)
    '    For Each aspect As AspectData In _listAspects
    '        If aspect.text = text Then
    '            aspect.valueFromInventor = value

    '            'подсчет дельты
    '            Try
    '                aspect.delta = Math.Abs((aspect.valueFromExcel - aspect.valueFromInventor) / aspect.valueFromExcel)
    '            Catch ex As Exception
    '                aspect.delta = 0
    '            End Try

    '            'увеличить счетчик считанных свойств из Inventor
    '            Dim c As Integer = CInt(lblCountOfAssembly.Text)
    '            c += 1
    '            lblCountOfAssembly.Text = c

    '            Exit For
    '        End If
    '    Next
    'End Sub

End Class
