# Независимое ревью Feature 012 — XLSX Template Engine (текущее состояние M1)

**Вердикт: Needs revision.** Статический XLSX-срез работает и в основном соответствует плану M1, но перед закрытием M1 обязательна доработка: повреждённый пакет роняет CLI и даёт HTTP 500 вместо ограниченной диагностики (High). Дополнительно есть несколько Medium-дефектов корректности и диагностики, а также открытый M1.5, по которому нужно решение Product Owner.

## 1. Объём и базовая линия

1. Проверен весь рабочий дифф относительно HEAD: 31 отслеживаемый файл (staged, unstaged, новые) и 5 untracked; дополнительно — неизменённые `MergeEngine`, `TemplateAnalyzer`, `DataResolver`, CLI/API, DI и DOCX-путь на предмет регрессий.
2. Базовая линия: Feature Spec, Architecture Assessment, Implementation Plan (в частности требования к `Analyze`/`Merge` на стр. 56, `MergeDiagnostic.Code` на стр. 35, M1.1–M1.8), `implementation-progress.md` (как заявление, а не доказательство), `AGENTS.md`, `docs/docx-template-compatibility.md`.
3. Рабочее дерево не изменялось: `git status --porcelain` до и после ревью — одинаковые 36 записей (31 tracked + 5 untracked). Пробы выполнялись только во временном каталоге `/private/tmp/feature012-review`.

## 2. Дефекты

### High

1. **F1. Повреждённый XLSX не превращается в `TemplateException`: CLI падает, API отдаёт 500, часть ошибок `Analyze` не возвращает JSON.** `XlsxTemplateHandler.Merge` (`MonoReport/Services/Xlsx/XlsxTemplateHandler.cs:36-66`) не оборачивает открытие пакета, чтение XML и рендеринг; `XlsxPackageReader.Load` (`MonoReport/Services/Xlsx/XlsxPackageReader.cs:127-132`) и `XlsxDocumentRenderer.Load` (`MonoReport/Services/Xlsx/XlsxDocumentRenderer.cs:207-212`) бросают `XmlException`. В CLI этот тип не перехвачен (`MonoReport.Cli/Services/CliApplication.cs:110-115`), в API срабатывает общий `catch (Exception)` → 500 (`MonoReport.Api/Program.cs:49-53`).
   Доказательство: worksheet XML с незакрытым элементом → merge завершился с `exit=134` и `Unhandled exception. System.Xml.XmlException: ... elements are not closed`; отсутствие relationship листа → merge `exit=1` с сырым `Specified argument was out of the range of valid values. (Parameter 'id')`, а `--analyze` — `exit=1` без JSON и сообщением `Template analysis could not be completed.` (исключение выходит из-за слишком узкого фильтра в `XlsxTemplateHandler.cs:30-33` и отсутствия перехвата в `TemplateAnalyzer.cs:34-41`).
   Это прямо противоречит плану (стр. 56: `Analyze` маппит повреждённые пакеты в invalid-результат, `Merge` бросает `TemplateException`), обещанию OpenAPI «invalid nonempty package content returns 200 with valid=false» и контракту 422. Частичный вывод при этом не создаётся — это хорошо, но поведение недопустимо для повреждённых/недокачанных файлов.
   Требуется: единый маппинг ошибок пакета/XML для `Merge`, расширение фильтра `Analyze` (`ArgumentException`, `InvalidOperationException` и родственные ошибки SDK), тесты на усечённый XML, отсутствующий relationship/part и соответствующие CLI/HTTP-статусы.

### Medium

2. **F2. Не-объектный корень JSON молча очищает все поля.** План (стр. 56) требует для XLSX корень-объект. Проверки нет: `XlsxTemplateHandler.cs:47` создаёт `DataScope` из любого `RootElement`, `DataResolver.Resolve` возвращает `Missing` для не-объекта (`MonoReport/Services/DataResolver.cs:42`), а политика по умолчанию — `EmptyString` (`MonoReport/Models/MergeOptions.cs:9`). Воспроизведено: `[{"Name":"Ada"}]` → `exit=0`, в выходном листе `<c r="A1" />` (пустая строка ячейки), без warning/error. Для DOCX массивы остаются допустимыми, поэтому нужна именно XLSX-специфичная проверка корня с кодовой диагностикой.

3. **F3. `Analyze` объявляет `valid:true` для шаблона, который `Merge` отвергает.** `Analyze` вызывает только `Plan` (`XlsxTemplateHandler.cs:16-24`), а проверки неподдерживаемых структур находятся в cropper/renderer (`MonoReport/Services/Xlsx/XlsxSheetCropper.cs:39-48`, `MonoReport/Services/Xlsx/XlsxDocumentRenderer.cs:46-60`). Воспроизведено: шаблон с `{Name}` и формулой `SUM` → `{"errors":[],"warnings":[],"requirements":[],"occurrences":[],"valid":true}`, а генерация — `exit=1` с `template.xlsx_unsupported_feature`. Это дефект предпроверки сам по себе; пустые `requirements`/`occurrences` — отдельно известный открытый M1.5 (см. раздел 4) и не должен смешиваться с parity-дефектом.

4. **F4. Опубликованная диагностика теряет конкретную причину и рецепт.** `TemplateException.ToDiagnostic` подменяет сообщение на фиксированный `SafeMessage` по коду (`MonoReport/Models/TemplateException.cs:62,79-95`); у ошибок cropper/renderer нет `Explanation`, поэтому, например, «Formulas are not supported yet» или имя конкретной структуры до пользователя не доходят — остаётся generic «The XLSX feature cannot be preserved safely...» плюс location. Тест `Feature012FoundationTests.cs:245` проверяет внутренний `error.Message`, а не публикуемый текст, поэтому дефект не виден. Спецификация (раздел 6.10) требует конкретного remedy; нужно либо передавать `Explanation`, либо не затирать сообщение для XLSX-кодов.

5. **F5. Фонетические прогоны (`rPh`) попадают в значение поля.** `ReadCell` конкатенирует все `Text`-потомки shared/inline string (`MonoReport/Services/Xlsx/XlsxPackageReader.cs:82,88`), включая `<rPh><t>`. Воспроизведено: shared string `東京 {Name}` с `rPh` «トウキョウ» → в выходной ячейке `東京 Adaトウキョウ`. Для валидных японских шаблонов это порча данных; `rPh` нужно исключать из текста (и учитывать при определении rich text).

6. **F6. Инвентарь возможностей собран, но нигде не используется — M1.7 «reject affected metadata» закрыт неполно.** `Features`/`WorkbookFeatures` заполняются (`XlsxPackageReader.cs:58-60,67-69`), но не читаются никем; отклонение опирается только на фиксированный список имён элементов (`XlsxSheetCropper.cs:20-25`). Части, связанные только relationship (например, pivot tables у worksheet, query tables) и workbook-level `externalReferences`, в этот список не входят: их relationship/part переживают кроп, хотя исходные строки/ячейки удалены. С реальным pivot-fixture я это не воспроизводил, но код однозначно показывает неиспользуемый инвентарь и неполный охват; нужен либо запрет по инвентарю, либо явная фиксация, что такие шаблоны вне поддержки, с тестом.

7. **F7. XML-недопустимые символы из данных дают сырую ошибку/500.** `WriteValue` создаёт `XElement` из значения (`XlsxDocumentRenderer.cs:103`); строка с `\u0001` вызывает `ArgumentException` вне маппинга: CLI печатает `' ', hexadecimal value 0x01, is an invalid character.` без кода и поля, API уйдёт в 500. Нужна валидация текста с кодовой ошибкой (`generation.invalid_data`) и указанием ячейки/поля; частичный файл при этом не создаётся.

### Low

8. **F8. `MergeDiagnostic.Code` не добавлен**, хотя план (стр. 35) прямо требует опциональный nullable `Code` (`MonoReport/Models/MergeDiagnostic.cs:3-9`). Сейчас это не видно пользователю (warning’и о missing values ещё не выводятся CLI/HTTP и запланированы на M5), но модель придётся менять перед M5 и стабильные коды предупреждений отсутствуют.

9. **F9. Учёт ресурсов не доведён.** `XlsxOperationContext.OutputCells` никогда не присваивается (`MonoReport/Models/Xlsx/XlsxOperationContext.cs:8-9`), `XlsxGeometryPlanner.Dimensions` вызывается только из тестов (`MonoReport/Services/Xlsx/XlsxGeometryPlanner.cs:12`). Потолки официально отложены, но план требует, чтобы accounting стартовал уже в M1; без `OutputCells` активация лимитов в M2+ потребует переделки.

10. **F10. Нарушения AGENTS.md по коду.** `XlsxNumberSafety` — internal static-класс с сервисной логикой (`MonoReport/Services/Xlsx/XlsxNumberSafety.cs:10-26`), что противоречит правилу «services as instances… constants-only classes must not contain service logic»; `XlsxDocumentRenderer` использует primary constructor без null-валидации (`MonoReport/Services/Xlsx/XlsxDocumentRenderer.cs:10-11`), тогда как правило предписывает `?? throw`.

11. **F11. Тихая обработка view/columns.** `CropViews` удаляет `<pane>` без диагностики и переносит только `activeCell`, оставляя stale `sqref`/`pane`-атрибуты у `selection` (`MonoReport/Services/Xlsx/XlsxSheetCropper.cs:142-159`); `<col>` с нечисловыми `min`/`max` молча удаляется (`:92-96`). Данные не теряются, но поведение не отражено ни warning’ом, ни документированным ограничением.

## 3. Пробелы в тестах

1. Нет тестов на повреждённый пакет в `Merge` (усечённый worksheet XML, отсутствующий relationship/part), на CLI exit-код и на HTTP-статус — поэтому F1 не пойман; существующий `WordFamily_Binary_excel_and_malformed_packages_reject` проверяет только детектор.
2. Нет теста на не-объектный корень JSON (F2); happy-path тесты используют только объект.
3. Нет parity-теста «`Analyze.Valid == true` ⇒ `Merge` не падает с unsupported» (F3); HTTP/CLI тесты анализа проверяют только `valid` и пустые `errors`.
4. `StaticCrop_Rejects_affected_metadata` проверяет внутренний `error.Message` вместо публикуемого текста, поэтому потеря детализации (F4) не обнаруживается.
5. Нет фикстуры с `rPh`/фонетикой и rich-run вариациями shared strings (F5).
6. Нет теста на relationship-only metadata (pivot/query table, external references) и на неиспользуемый inventory (F6).
7. Нет теста на XML-недопустимые символы в данных и на отсутствие частичного вывода при такой ошибке (F7).
8. Нет `Feature012FoundationTests.Analyzer_` и `analyzer-scopes.template.xlsx`, требуемых M1.5; отсутствует тест на вывод warning’ов CLI (это M5, но фиксирует F8).

## 4. Готовность M1 и последовательность

1. M1.1–M1.4, M1.6–M1.8: заявленные возможности подтверждаются кодом и тестами для статического среза — типы/стили, shared/inline strings, кроп по PrintArea, переписывание PrintArea, отклонение merge за границей, удаление внешнего контента, DI-композиция, CLI/HTTP happy path, неизменность DOCX.
2. M1.5 открыт: `Analyze` всегда возвращает пустые `requirements`/`occurrences`, fixture отсутствует. Проекция marker-scope действительно зависит от грамматики M2.4, но проекция **корневых** полей (например, `{Name}`) от M2.4 не зависит и сейчас не реализована. Нужно решение Owner: (1) явно отложить M1.5 с переупорядочиванием, (2) реализовать сейчас статическую часть и оставить marker-scope к M2.4, либо (3) перенести M2.4 раньше M1.5.
3. M2 нельзя начинать по плану, пока M1.5 не закрыт или Owner не утвердил изменение порядка. F1 и F2 при этом стоит исправить до границы M1: это фундаментальные контракты `Merge`/`Analyze`, а не функциональность следующих вех.
4. Внешняя приёмка не закрыта: ручная проверка Microsoft Excel (E1: открытие без repair, пересчёт, печать/внешний вид, даты) не выполнялась; calculated-value evidence для M2 отсутствует.

## 5. Выполненные проверки

1. `dotnet test MonoReport.Tests/MonoReport.Tests.csproj -c Release --no-build` → 365/365 passed.
2. `dotnet test MonoReport.Tests/MonoReport.Tests.csproj -c Release --no-build --filter "FullyQualifiedName~Feature012"` → 54/54 passed.
3. `dotnet test MonoReport.Cli.Tests/MonoReport.Cli.Tests.csproj -c Release --no-build` → 15/15 passed.
4. `dotnet test MonoReport.Api.Tests/MonoReport.Api.Tests.csproj -c Release --no-build` → 27/27 passed.
5. `dotnet test MonoReport.Api.Tests/MonoReport.Api.Tests.csproj -c Release --no-build --filter "FullyQualifiedName~Real_kestrel_enforces_whole_request_bound"` → 1/1 passed.
6. `git status --porcelain` → 36 записей (31 tracked + 5 untracked) до и после ревью; `git diff --stat HEAD -- MonoReport.Api.Tests/Services/ApiHostTests.cs` → пусто.
7. Репро-пробы (CLI из текущей Release-сборки, временные файлы вне репозитория): валидный шаблон → `exit=0`, корректные `Ada`/тип/стиль/PrintArea; массивный корень → `exit=0` с пустой ячейкой; усечённый worksheet XML → `exit=134` (`XmlException`); отсутствующий sheet relationship → merge `exit=1` с `(Parameter 'id')`, analyze `exit=1` без JSON; формула в PrintArea → analyze `valid:true`, merge `exit=1` с generic-сообщением; shared string с `rPh` → в значении `東京 Adaトウキョウ`; `\u0001` в данных → `exit=1` с сырым XML-сообщением.
8. Сборки, использованные для тестов и проб, актуальнее исходников (проверено по mtime; `dotnet build MonoReport.Cli -c Release` — up-to-date, 0 warnings/0 errors).

## 6. Классификация падающего ApiHost-теста

`ApiHostTests.Real_kestrel_enforces_whole_request_bound_including_streamed_requests_and_recovers` — **не регрессия Feature 012**. Файл не изменён относительно HEAD (пустой diff), тест прошёл 11 раз подряд, включая текущий запуск; в предыдущем ревью Feature 003 зафиксировано такое же падение (`ResponseEnded`, строка 44). Это среда-зависимый/флейки-тест реального Kestrel; его нужно вынести в отдельную задачу стабилизации, а не блокировать M1.

## 7. Что не выполнялось

1. Ручная проверка в Microsoft Excel (E1) — не запускалась; остаётся обязательной для финальной приёмки.
2. HTTP-запрос с повреждённым XLSX не отправлялся: вывод о 500 основан на коде `Program.cs:49-53`; CLI-падение воспроизведено эмпирически.
3. Полная сборка solution и browser automation не запускались; `--no-build` использовался с проверенными актуальными бинарями.