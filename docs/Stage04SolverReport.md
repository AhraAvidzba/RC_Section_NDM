# Отчет этапа 4: нелинейный решатель

## Статус

Этап 4 реализован в границах численного решателя. Реализованы общий `CSectionSolver`, касательная матрица `3 x 3`, полный метод Ньютона, line search, демпфирование через коэффициент шага, пошаговое приложение нагрузки, ограничение приращений, нелинейное замещение бетона арматурой и диагностический журнал.

Поиск предельного множителя, удерживающего момента, коэффициента запаса, коэффициента использования и расчет ширины раскрытия трещин не реализовывались.

## Созданные файлы

- `src/Solver/CSectionSolver.cls`
- `src/Excel/CSystemSettingsReader.cls`
- `tests/modTestSectionSolver.bas`
- `tools/build_workbook/Run-SectionSolverTests.ps1`

## Измененные файлы

- `tools/build_workbook/Build-Workbook.ps1`
- `src/Materials/CConcreteBilinearMaterial.cls`
- `src/Materials/CConcreteTrilinearMaterial.cls`
- `docs/NDM_PROJECT_SPEC.md`
- `docs/Architecture.md`
- `docs/ImplementationPlan.md`
- `docs/ValidationPlan.md`
- `docs/Stage04Report.md`

## Временные параметры материалов

Для разработки и тестирования solver используется статус:

```text
PROVISIONAL_FOR_SOLVER_TESTING
```

Параметры находятся на листе `System` в `rngSystemSettings`, видимы пользователю и допускают ручное изменение.

Принятые временные значения:

- бетон B30: `Eb = 32500 МПа`, `Rb = 15.5 МПа`, `Rbt = 1.10 МПа`, `Eps1 = -0.0015`, `EpsU = -0.0035`;
- обычная ненапрягаемая A400: `Es = 200000 МПа`, `Rs = 350 МПа`, `Rsc = 350 МПа`, `EpsY = 0.00175`, `EpsU = 0.025`.

Расчет с такими параметрами не является окончательно верифицированным нормативным расчетом.

## Реализованная численная схема

- неизвестные: `epsilon0`, `kappaX`, `kappaY`;
- равновесие: `N`, `Mx`, `My`;
- деформации: `epsilon(x, y) = epsilon0 + kappaX * y + kappaY * x`;
- касательная матрица строится по текущим касательным модулям материалов;
- система `3 x 3` решается через `CLinearSystem3x3`;
- нагрузка прикладывается ступенями;
- шаг Ньютона ограничивается по `epsilon0` и кривизнам;
- line search уменьшает шаг, если невязка не улучшается;
- арматура учитывается через эффективное замещение:

```text
EffectiveSteelStress = SteelStress - ConcreteStressAtSameStrain
EffectiveTangent = SteelTangent - ConcreteTangentAtSameStrain
```

## Проверки

Команды:

```powershell
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Build-Workbook.ps1
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Validate-Workbook.ps1
powershell -ExecutionPolicy Bypass -File tools/build_workbook/Run-SectionSolverTests.ps1
```

Результат тестов этапа 4:

```text
TOTAL_SECTION_SOLVER: passed=34; failed=0
```

Проверено:

- чтение временных параметров из `rngSystemSettings`;
- совпадение `CSectionSolver` с линейным решателем на линейных материалах;
- линейный расчет с обычной ненапрягаемой арматурой;
- центральное сжатие на временной двухлинейной диаграмме бетона;
- работа временной двухлинейной диаграммы бетона совместно с обычной A400;
- пошаговое приложение нагрузки;
- ограничение приращений;
- наличие диагностического статуса `PROVISIONAL_FOR_SOLVER_TESTING`.

## Известные ограничения

- Трехлинейные диаграммы доступны как параметризованные классы, но не являются основным тестовым режимом этапа 4.
- Нормативные вопросы по формулам диаграмм остаются открытыми в `docs/OpenNormativeQuestions.md`.
- Текущие нагрузки в тестах выбраны для проверки сходимости и равновесия, а не для поиска предельного состояния.
- Результаты не должны использоваться как инженерная нормативная проверка до завершения нормативной трассировки и верификации.

## Следующий этап

Следующий этап - реализация поиска несущей способности и расчетов `Mx`/`My` только после сохранения текущих тестов этапа 4 как регрессионных.
