# Карта Удаления Замененной Архитектуры

## Дополнение 13.10.2026: расчёт защитного слоя

`CSectionPropertiesCalculator.ConcreteCoverFromPointAlongDirection` удалён:
это была обёртка без собственных вычислений, с единственным вызовом из
`modTestGeometry.TestConcreteCoverUsesLocalContour`. Тест теперь напрямую
инициализирует `CSectionGeometryQuery` и вызывает его одноимённый метод.
Потребитель расчёта трещин уже использовал общий query. Опубликованный
декларативный API содержит метод query, который сохранён. Контроль 40 мм
локального защитного слоя против 340 мм глобальной проекции и допуски
не изменены. Полная приёмка текущего этапа ещё выполняется.

## Фактическое Состояние 12.10.2026

**VERIFIED_PUBLISHED.** Из source, проверочных книг и основного output
post-retirement книги удалены все 12 специализированных geometry/rebar/
annotation classes, ISectionShape, modGeometryPreparation и
modRebarLayoutGeometry (15 компонентов). CFakeMultiOpeningGeometry удалён
ранее. Актуальные владельцы ответственности указаны в исторической таблице
ниже; все её ACTIVE/REVIEW/IN_PROGRESS заменены этим подтверждённым состоянием.

Все test constructors перенесены. modTestGeometryFixtures форматирует
in-memory ввод действующих рецептов, описывает явные общие fixtures и
переносит готовые данные; собственной расстановки стержней/регистрации типа
в нём нет. Прежний Circle/RectSet/Rounded/Hollow builder нигде не вызывается.
Прежние registry factories и type-specific reader methods удалены ранее.
Поиск src/tests/tools не дал активных ссылок на заменённые components/helpers.
История, неизменяемый baseline и миграционные отчёты сохраняются.

Скрытый constructor performance runner заменён фиксированными Outer/Opening
с настоящими дугами через существующий query. Source-contract проверяет
общий prepared region и неизменяемое definition. Import lists/README/AGENTS
актуализированы. Custom не создавался, в baseline отсутствовал.

PrepareImportedSection и private helpers перенесены из CSectionAnnotations
в CSectionAnnotationBuilder. Контейнер хранит данные; общий builder добавляет
точные/приблизительные размеры и ALL-подпись. PrepareSectionSnapshot вызывает
его один раз. Формат, численные вычисления и единицы сохранены. Два символа
диаметра в комментариях записаны словами; runtime Unicode через ChrW сохранён.

Книга: `docs/regression/GeometryDeclarative_2026-10-12/RetiredArchitecture/RC_Section_NDM.xlsm`.
Native compile PASS. Read-only сверка с WorkbookRetirement: PASS, 15 ожидаемых
удалений, 0 additions, 140 компонентов (99 classes, 36 standard, 5 document),
unexpected/sourceMismatch=0; семь approved source bodies совпадают.

| Направленный набор после удаления | PASS | FAIL |
|---|---:|---:|
| Геометрия | 765 | 0 |
| Аналитические paths/anchors/RoundedPolygon/offset/input | 439 | 0 |
| Rules/Definition | 202 | 0 |
| Общий query | 1319 | 0 |
| Владение/round-trip контуров | 520 | 0 |
| Сохранённые CAD-контуры | 231 | 0 |
| Единицы импортированного снимка | 228 | 0 |
| SectionSolver | 2834 | 0 |
| Numerical baseline | 39 | 0 |
| Итого | 6577 | 0 |

Это post-retirement этап, не полная D/E/G/R/A/I приёмка. Основная книга и
VBA_All_Code.txt ещё прежние. Обязательные остатки: полный публичный API,
Config setbacks, новый пример, полный гайд в книге, адресная итоговая приёмка
и окончательный output/dump. Старый VBProject основного output удаляется при
его подготовке; его существование явно учитывается как OUTPUT_PENDING.

## Исторический Журнал До Retirement

Следующий текст сохранён как хронология промежуточных этапов; его состояния
ACTIVE/IN_PROGRESS не описывают текущий source/native staging.

Общий source/native staging-маршрут переключен на рецепты. Старые классы
еще остаются в численных тестах и прежней основной книге; удаление пока
не подтверждено. Готовность отдельных новых рецептов указана ниже.
Полезная общая математика переносится до удаления. Baseline остается только
в отдельном неизменяемом архиве/истории, не в production fallback.

| Старый компонент | Новый владелец | Состояние |
|---|---|---|
| CGeometryCircle | CSectionRecipeCircle + общие примитивы | ACTIVE |
| CGeometryRectSet | CSectionRecipeRectSet + общая сборка смежных областей | ACTIVE |
| CGeometryRoundedRectangle | CSectionRecipeRoundedRectangle + RoundedPolygon/paths | ACTIVE |
| CGeometryHollowRectangle | CSectionRecipeHollowRectangle + общие Outer/Opening | ACTIVE |
| CCircleRebarLayoutBuilder | Общие distribution/dependencies | ACTIVE |
| CRectSetRebarLayoutBuilder | Общие paths/trims/publication | ACTIVE |
| CRoundedRectRebarLayoutBuilder | Аналитические paths/distribution | ACTIVE |
| CHollowRectRebarLayoutBuilder | Общий ProjectFrom/constraints | ACTIVE |
| CCircleAnnotationBuilder | Общие annotation rules | ACTIVE |
| CRectSetAnnotationBuilder | Общие annotation rules | ACTIVE |
| CRoundedRectAnnotationBuilder | Общие annotation rules | ACTIVE |
| CHollowRectAnnotationBuilder | Общие annotation rules | ACTIVE |
| ISectionShape | ISectionRecipe/CSectionDefinition | ACTIVE |
| modGeometryPreparation.BuildConcreteGeometry | Единый общий build | ACTIVE |
| Registry CreateShape/CreateGeometry/CreateRebars/BuildAnnotations/BuildGeneratedModel | Recipe registry/единый build | REMOVED_FROM_SOURCE_AND_NATIVE_STAGING |
| Reader Append*RectSet/RoundedRectangle/HollowRectangle и scan *Geometry | Рецепт через generic table source | REMOVED_FROM_SOURCE_AND_NATIVE_STAGING |
| CFakeMultiOpeningGeometry | Definition fixture с двумя/пятью Opening | REMOVED_FROM_SOURCE_AND_NATIVE_STAGING |
| Старые численные test constructors | Явные общие контуры/правила и реальные рецепты | IN_PROGRESS |
| modTestShapeArchitecture.SourceSnapshot через CreateGeometry/CreateRebars | Готовый build snapshot, прежние числовые эталоны | REMOVED_FROM_SOURCE_AND_NATIVE_STAGING |
| modRebarLayoutGeometry helpers | Сохранить единственную общую операцию либо удалить замененную | REVIEW |
| Custom-код/UI/имена | Не создаются; в исходном inventory/Config/Results и исходниках не найдены | NOT_PRESENT_IN_BASELINE |

Для каждой строки нужны: потребители исходников/tests/runners, новый метод,
поэлементный regression, удаление файла, отсутствие ссылок и VBProject manifest.
Не считать удаление файла достаточным доказательством очистки книги.

Локальный этап Circle: новый рецепт и общий сборщик проходят 28 fixtures,
SP35/SP63 и native preview. Три старых Circle-компонента пока ACTIVE в
production и старом baseline harness; это не завершенный retirement.
Перед итоговой приемкой удалить их файлы, потребителей и VBProject-записи.

Custom: поиск по src/tests/tools/AGENTS/README для CustomSection/rngCustom/
CGeometryCustom/CSectionRecipeCustom и по полным исходным NativeInventory,
Config/Results JSON не дал совпадений. Пользовательские данные не удалялись.
Повторить проверку полного итогового VBProject/UI при финальной сборке.

Локальный RectSet: три стандартные формы и 18 вариантов проходят общий
recipe/build, SP35/SP63 и native preview. Старые три RectSet-компонента
пока ACTIVE только в прежнем production/baseline-маршруте. Новый рецепт
их не использует. Удаление выполняется после оставшихся семей.

Локальный RoundedRectangle: 24 варианта, независимый аналитический oracle,
SP35/SP63 и Excel preview проверены. Новый рецепт не использует прежние
geometry/rebar/annotation classes. Три старых компонента пока ACTIVE только
в прежнем production/baseline route. Их нельзя удалить до переключения
потребителей. HollowRectangle уже имеет отдельный development-рецепт;
его крайние формы еще проверяются.

Локальный HollowRectangle: новый рецепт не вызывает прежние три компонента
и RoundedRectangle-геометрию. Общий Projected/Dependent executor обрабатывает
нормали и связанные ряды; общий query строит дуги и отверстие. Старые три
Hollow-компонента остаются ACTIVE в production и baseline harness до общего
переключения. Это еще не завершенный retirement.

Итоговый счетчик: NativeInventory baseline содержит 98 Type=2, 34 Type=1
и 5 Type=100. Verify-GeometryBaselineCode теперь сохраняет отдельные counts
и поименные AddedComponents/RemovedComponents; итоговые числа следует взять
из проверки конечной книги, а не суммировать временные milestone-коммиты.

RetirementFixtures: 11 test modules перенесены, ссылки на заменяемые contracts
в них отсутствуют. Оставшиеся test consumers: modTestGeometry (199 matches)
и modTestWorkbookInterface (20); самих специализированных классов это еще
не удаляет. Fake class удален из source и проверочной книги. Native compile,
520 ownership assertions после удаления и строгая сверка source/VBProject
проходят. В основной прежней книге fake пока остается до конечной сборки.

Update-PrivateWorkbookCode/Verify-GeometryBaselineCode принимают явный
RetiredModules: source должен уже отсутствовать, document/designer modules
запрещены, факт удаления проверяется по baseline/native inventory. Обновление
разрешено только в docs/regression; основная книга защищена хешем.
Проверка отказа при существующем source CSectionModel: PASS, книга неизменна.

## Итоговая опубликованная книга 12.10.2026

Read-only native readback `PublishedSourceEquality.json`: 147 компонентов,
142 source + 5 document MATCH. Ни одного из 15 retired компонентов нет;
контрольная трапеция и измерительный harness в production отсутствуют.
`CFakeMultiOpeningGeometry`, удалённый более ранним этапом, также отсутствует.
Основной output обновлён, source и actual dump согласованы. Ранние записи
о pending/native staging выше сохранены как история; текущий статус —
VERIFIED_PUBLISHED. Подробности FINAL_REPORT.md/FINAL_ACCEPTANCE_EVIDENCE.json.
