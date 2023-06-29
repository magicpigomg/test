Sub ЗагрузкаНепрочитанныхСообщений()
    Dim OutlookApp As Outlook.Application
    Dim OutlookNamespace As Outlook.Namespace
    Dim OutlookFolder As Outlook.MAPIFolder
    Dim InboxFolder As Outlook.MAPIFolder
    Dim EmailItem As Outlook.MailItem
    Dim ExcelApp As Excel.Application
    Dim ExcelWorkbook As Excel.Workbook
    Dim ExcelWorksheet As Excel.Worksheet
    Dim RowCount As Integer

    ' Установка ссылок на Outlook и Excel
    Set OutlookApp = New Outlook.Application
    Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")
    Set InboxFolder = OutlookNamespace.GetDefaultFolder(olFolderInbox)
    Set OutlookFolder = InboxFolder.Folders("Название папки") ' Замените "Название папки" на нужное имя папки в почте

    ' Создание нового экземпляра Excel
    Set ExcelApp = New Excel.Application
    Set ExcelWorkbook = ExcelApp.Workbooks.Add
    Set ExcelWorksheet = ExcelWorkbook.Sheets(1)

    ' Настройка заголовков столбцов в Excel
    ExcelWorksheet.Cells(1, 1).Value = "Отправитель"
    ExcelWorksheet.Cells(1, 2).Value = "Тема"
    ExcelWorksheet.Cells(1, 3).Value = "Дата"

    ' Заполнение таблицы данными из непрочитанных сообщений
    RowCount = 2 ' начинаем с 2 строки, чтобы пропустить заголовки столбцов
    For Each EmailItem In OutlookFolder.Items
        If EmailItem.UnRead Then
            ExcelWorksheet.Cells(RowCount, 1).Value = EmailItem.SenderEmailAddress
            ExcelWorksheet.Cells(RowCount, 2).Value = EmailItem.Subject
            ExcelWorksheet.Cells(RowCount, 3).Value = EmailItem.ReceivedTime
            RowCount = RowCount + 1
        End If
    Next EmailItem

    ' Отображение Excel и сохранение файла
    ExcelApp.Visible = True
    ExcelWorkbook.SaveAs "C:\Путь\К\Файлу.xlsx" ' Замените "C:\Путь\К\Файлу.xlsx" на путь, по которому хотите сохранить файл

    ' Освобождение ресурсов
    Set ExcelWorksheet = Nothing
    Set ExcelWorkbook = Nothing
    Set ExcelApp = Nothing
    Set EmailItem = Nothing
    Set OutlookFolder = Nothing
    Set InboxFolder = Nothing
    Set OutlookNamespace = Nothing
    Set OutlookApp = Nothing
End Sub


Примечания:

Убедитесь, что у вас есть ссылка на библиотеку Microsoft Outlook в вашем проекте VBA (Инструменты -> Ссылки -> Microsoft Outlook Object Library).
Замените "Название папки" на фактическое имя папки в почте Outlook, в которую вы хотите загрузить непрочитанные сообщения.
Замените "C:\Путь\К\Файлу.xlsx" на путь, по которому вы хотите сохранить файл Excel со списком непрочитанных сообщений.
После добавления этого кода в модуль VBA вы можете связать его с кнопкой в Excel, чтобы при нажатии на кнопку макрос выполнялся и загружал непрочитанные сообщения в Excel.
