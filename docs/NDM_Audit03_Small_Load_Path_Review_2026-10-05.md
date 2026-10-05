# Audit03: Маленькие Ненулевые Компоненты Пути

Контрольная точка: `fb364ab`. Новых классов нет, основная книга неизменна.

## Доказанный Дефект

Широкая матрица v318 CircleSym/Light/Off завершила все 588 случаев:
`24160/57`. В 19 Formation случаях были validInput и два noInternalFailure
assertions с отказом. Причина: `CLoadPathMath` не находил опорную компоненту
у ненулевого пути; Formation возвращал rsInternalError/INTERNAL_ERROR.
Текущий CrackedState при этом был успешно получен.

Formation передавал `NEAR_ZERO_FORCE=1e-6` как общий порог Base, включая
момент 1e-6 Н*мм. Сам helper по умолчанию также отбрасывал Base <=1e-9.
Это отдельные размерные пороги, не связанные с заданным пользователем
инженерным фильтром Calculation.ZeroMomentPerDepth, который в fixture равен 0.

Усиленный existing Solver test на frozen production v319: `81/20`.
Проверяются оба знака N/Mx/My от 1e6 до 1e-12 с одними допусками solver-а,
опорная компонента, точная lambda=2, две нулевые независимые невязки,
восстановительный обычный путь и ошибка полностью нулевой Base.

## Исправление И Границы

Production CLoadPathMath по умолчанию считает нулем только Base=0;
явный optional-порог программного API сохраняется. Formation больше
не передает порог классификации осевой силы в математику всех компонентов.
Допуски равновесия, предельные критерии, порядок Auto и чистые формулы
не изменяются. Справка разделяет точность solver-а и инженерное обнуление.

Positive v320: отдельный helper `125/0`, полный Solver `1492/0`,
Capacity `3562/0`, Crack `2155/0`; все runner-ы завершены с exit=0,
source unchanged=True. Общие численные assertions из полного Off v313
совпали точно: Solver 232, Capacity 601, Crack 196; missing=0, differences=0.
Новые проверки добавлены к прежним; их expected и tolerance не ослаблялись.

Read-only actual VBE export содержит 113 компонентов. Book SHA-256:
`CACB7F261AA3117BDD7C09340A41E65AD95E43F575A619F29789E9DA06EC5871`;
export SHA-256:
`0B059C5F06057AFE7613BBCB93F3DF2FF3F7968BD1CA80437653D4C543FC51B6`.
SourceContracts завершен `108/108`, failed=0. Первый broad v320 блок
CircleSym/Light/Off: 588 случаев, `24293/0`, Results save-reopen=True,
source unchanged=True; прежние 19 Formation internal failures не повторились.
Остальные блоки 52-block broad matrix выполняются; полный итог пока не принят.

Не утверждается, что при любой маленькой нагрузке численный поиск достигнет
критерия до технической границы: SEARCH_BOUND_REACHED отличается от ошибочного
INTERNAL_ERROR. Audit03 и выпускная приемка остаются незавершенными.
