# Карта Удаления Замененной Архитектуры

Старые производственные узлы сейчас ACTIVE: общий маршрут еще не переключен
и удаление не подтверждено. Готовность отдельных новых рецептов указана ниже.
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
| Registry CreateShape/CreateGeometry/CreateRebars/BuildAnnotations/BuildGeneratedModel | Recipe registry/единый build | ACTIVE |
| Reader Append*RectSet/RoundedRectangle/HollowRectangle и scan *Geometry | Рецепт через generic table source | ACTIVE |
| CFakeMultiOpeningGeometry и старые test constructors | Definition fixtures с прежними инвариантами | ACTIVE |
| modTestShapeArchitecture.SourceSnapshot через CreateGeometry/CreateRebars | Готовый build snapshot, прежние числовые эталоны | ACTIVE |
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
