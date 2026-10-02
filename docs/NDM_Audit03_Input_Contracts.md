# Audit03: Защитные Контракты Ввода

Дата: 2026-10-02. F07/K02 review, отдельный directed gate.
Статус: reproducer подготовлен, runtime до/после еще не выполнен.
Никакая расчетная методика либо нормативная формула здесь не меняется.

## Проверяемые Входы

| Контракт | Public Call Path | Ожидаемая Реакция |
| --- | --- | --- |
| Boolean | Config Range -> SettingsReader.LoadFromRange -> GetBoolean | Все утвержденные Yes/No aliases читаются; явное Maybe/TODO/пустое значение не становится False/default. |
| Целое число | Config Range -> GetLong | Точные Long границы принимаются; дробь не округляется молча, overflow не превращается в runtime 6. |
| Таблица Config | SettingsReader.LoadFromRange / AppendSettingsRange | Nothing, scalar Value2 или меньше обязательных колонок отвергаются адресной ошибкой структуры до UBound/indexing. |
| Таблица профилей | ProfileCatalog.LoadFromRange | Scalar Value2 не вызывает runtime 13; ошибка структуры имеет понятную причину. |
| Геометрия сетки | CFiberMeshBuilder.BuildMesh | Nothing проверяется до geometry.IsValid. |
| Размер сетки | BuildMesh -> axis/product -> ReDim | Непредставимое число ячеек отвергается до Long conversion, умножения и выделения памяти. |
| Boundary subdivisions | BuildMesh | Недопустимое значение не заменяется 1 без сообщения. |

Отсутствующий необязательный key и явно неверное введенное значение различаются.
Default-параметр публичного getter-а нужен для оговоренных optional/test/старых
схем; он не разрешает исправлять неверный явный ввод молча. Требование о каждом
required key проверяется отдельно в его reader/consumer, без массовой подмены
всех настроек одним новым правилом.

## Тест И Изоляция

`modTestWorkbookInterface.RunAudit03InputContractTests` вызывает тот же набор,
который входит в полный UI suite. Рабочая книга - независимая test fixture.
Временный лист `__Audit03Input` удаляется даже после ошибки; Config, validation
и пользовательские формулы не меняются. В отчете выводятся фактический error
code и reason, а не только флаг наличия исключения.

План: тест-only import в отрицательную копию текущего green v29; сохранить
negative log; затем минимальные guards существующих owners, новый positive
directed, оба full modes и source/book/export gate. До этого результата здесь
нет отметки PASS.
