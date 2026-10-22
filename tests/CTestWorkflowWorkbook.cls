VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestWorkflowWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: прокси приватной книги для наблюдения границ Excel-workflow
' ==========================================================================
' Передаёт таблицы и листы настоящей книги без изменения их структуры.
' Обязательные COM-имена Application/Names/Worksheets/Name/Path сохранены;
' счётчики проверяют порядок Calculate и чтения входов без ветвей в рабочем коде.

Private mWorkbook As Object ' Только приватная книга, выбранная native harness.
Private mApplication As CTestExcelApplicationState ' Делегат состояния и счётчик явного пересчёта.
Public TestNameReads As Long ' Число обращений к коллекции имён за действие.
Public TestNamesBeforeCalculate As Long ' Чтения имён до актуализации входов запрещены для Form/Calculate.
Public TestNamesFault As Long ' Отказ обязательной входной таблицы после успешного Calculate.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Связывает прокси с настоящей приватной книгой; владение Excel остаётся у harness.
Public Sub TestInitialize(ByVal workbook As Object, ByVal app As CTestExcelApplicationState)
    Set mWorkbook = workbook: Set mApplication = app
    TestNameReads = 0: TestNamesBeforeCalculate = 0: TestNamesFault = 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает наблюдаемый контракт приложения вместо глобального Application.
Public Property Get Application() As Object
    Set Application = mApplication
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Передаёт реальные имена и фиксирует нарушение порядка чтения до Calculate.
Public Property Get Names(Optional ByVal index As Variant) As Object
    TestNameReads = TestNameReads + 1
    If mApplication.TestCalculateCalls = 0 Then TestNamesBeforeCalculate = TestNamesBeforeCalculate + 1
    If TestNamesFault <> 0 Then Err.Raise TestNamesFault, "CTestWorkflowWorkbook.Names", "Имитирован недоступный именованный диапазон."
    If IsMissing(index) Then Set Names = mWorkbook.Names Else Set Names = mWorkbook.Names(index)
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает существующие листы без создания копий данных и без сохранения книги.
Public Property Get Worksheets(Optional ByVal index As Variant) As Object
    If IsMissing(index) Then Set Worksheets = mWorkbook.Worksheets Else Set Worksheets = mWorkbook.Worksheets(index)
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает имя приватной книги для штатной диагностики запуска.
Public Property Get Name() As String
    Name = mWorkbook.Name
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает исходный каталог только для штатного отключённого в тестах отчёта.
Public Property Get Path() As String
    Path = mWorkbook.Path
End Property
