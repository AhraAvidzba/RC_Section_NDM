# Карта Удаления Замененной Архитектуры

Все строки сейчас ACTIVE: замена еще не выполнена и удаление не подтверждено.
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
