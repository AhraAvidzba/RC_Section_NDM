# NDM: этап5b — числовые границы таблиц и сохранения

До исправления отдельная нативная проба подтвердила52отказа ожиданий на6реальныхадаптерах при обоих десятичных знаках Excel: произвольные внутренние пробелы трактовались как группировка разрядов. В частности, «1 2» становилось12 в таблице рецепта/нагрузках/геометрииResults/необязательномчислеJSON/таблицеСП35, а «0 2» становилось2 в точности схемы. Первичный Baseline01_No.txt сохранён полностью; это диагностический FAIL старого поведения, не ослабление ожиданий.

Общая полная грамматика TryReadStrictDouble применена только к подтверждённым текстовым границам CSectionInputTable.Number, CLoadCombinationReader.ReadOptionalLoadDouble, modSectionSnapshot.ReadGeometryNumber, CCalculationProfile.ParseVisualizationPrecision, CStabilityCalculator.SP35TablePositiveNumber и modSectionOptionalData.ReadSectionOptionalNumber. Пустые optional-значения, диапазоны, целочисленные ограничения, адреса и единицы сохранены. ИнтерполяцияСП35 получает ту же локальную копию таблицы с уже проверенными Double; формулы/порядок интерполяции не меняются. Инвариантное необязательное поле по-прежнему сохраняется с точкой/E; Val оставлен после полной проверки для сохранения прежнего округления корректных файлов.

Новые RunLocaleNumberTests/RunLocaleBoundaryTests включены в RunBatchCalculationTests, с агрегацией их реальных passed/failed. В No и Yes прошли Locale184/0, Boundary148/0, SP35Diagnostic105/0, GeneralConfig152/0, Snapshot417/0 иbaseline39/0. Проверены точка/запятая/E, повреждённые группы/смешанные разделители и запись/повторное открытие собственной xlsx с другим десятичным знаком Excel. Бинарные/Excel-снимки и старые26/33-колоночные форматы проверены существующим Snapshot417. Baseline точен относительноэтапа5a, кромеelapsedSec.

Точная книга нативно readonly сверена со172исходниками+6документами; AllCode из неё. Заменён толькоvbaProject.bin,106остальных частей побайтно сохранены. Новых классов/модулей0. Проверки схемы/экспорта и иных границ аудита продолжаются; общего полногоPASS нет. Глобальная локальWindows не менялась, фактическая вторая локальruntimeVBA пока не подтверждена.

Book SHA256: `9413B97A4E1C8A2414A3834FF89DFA2CF182CAF9967A8E11484E18EDA4BFDFA4`. AllCode SHA256: `5C9F79CE06237D0D41F205884DB457B59DD89E60A1ECB67F6438E39B37C6252B`.

Доказательства: docs/regression/OOPLayers_2026-10-26_Run01/Stage5b/Acceptance.json,ConfirmedBaseline.json,ProfileAcceptance.json,PackagePreservation.json,SourceEquality.json,ChangedSources.json,Logs.zip.
