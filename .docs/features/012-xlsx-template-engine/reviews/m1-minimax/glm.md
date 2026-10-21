# Независимый review Feature 012 (XLSX Template Engine) — M1

## Вердикт

**Approved with minor findings** — для текущего сос­тояния M1 (M1.1–M1.4, M1.6, M1.7, M1.8) при условии, что замечание **1-УВ (High)** и несоответствие чекпоинту M1.7 (вопрос 2-СР) устраняются до границы M1/M2, как это и предусмотрено планом. Все фактические тестовые прогоны из чистой сборки №1 повторили заявленные результаты (365 core / 27 API / 15 CLI, все PASS, ноль warning'ов компиляции).

Общая оценка: архитектура чистая (сервисы-инстансы с внешними зависимостями, per-operation context, immutable crop-план, единая точка дискриминации пакетов), DOCX-контракты не тронуты, «fail explicitly instead of plausible incomplete output» в целом соблюдён. Ниже — действия, которые необходимо закрыть, часть до старта M2.

---

## Findings

### High

**1-ВЫС. Merge для XLSX не требует JSON-объекта в корне: неявно-пустой отчёт при массиве/скаляре**
- Где: `MonoReport/Services/Xlsx/XlsxTemplateHandler.cs:36–66` (в Merge нет проверки `data.RootElement.ValueKind == JsonValueKind.Object`; `DataScope` строится из корня как угодно, строка 47).
- Требование: план §3.1.6 — «XLSX requires JSON object data per its specification; do not apply that new root restriction to DOCX».
- Воспроизведение (выполнено в scratch-проекте вне репозитория): шаблон с ячейкой `{Name}`, вход `[{"Name":"Ada"}]`, `MergeOptions` по умолчанию → **Merge завершается успешно, Format=Xlsx, ячейка A1 пустая, 0 warnings**; при `MissingValues=Warning` — 1 warning, при `Error`/`Warning` поведение все равно не ошибочного пакета, а пустого значения поля. По умолчанию (`EmptyString`) пользователь получает валидно выглядящий пустой DOCX-подобный отчёт (xlsx) без диагностики — «тихо потерянные данные».
- Минимальное исправление: в начале `XlsxTemplateHandler.Merge` проверить `data.RootElement.ValueKind is not JsonValueKind.Object` → `TemplateException` с кодом `generation.invalid_data` (или отдельный XLSX-код), не затрагивая DOCX-путь.

### Medium

**2-СР. M1.7 выполнен не полностью относительно собственных критериев: класс `Feature012DiagnosticsTests.Static_` не создан, 100-error граница для XLSX не проверена**
- Где: план, M1.7 (строка 159): фокус-группа `<Feature012FoundationTests.StaticCrop_>` **и `<Feature012DiagnosticsTests.Static_>`** — класса не существует (grep по всему репозиторию: только упоминания в плане). Соответственно, контрольная граница `99/ровно 100/101+` с наблюдением «дальнейшая оценка прекращена» (план §7.4) не выполнена: механизм `XlsxOperationContext.LimitReached/Report` (`MonoReport/Models/Xlsx/XlsxOperationContext.cs:10–20`) существует, но не покрыт тестами на XLSX-пути. Progress-отчёт (запись 9) не упоминает этот пробел — заявленное «M1.7 completed» оспоримо.
- Исправление: создать `Feature012DiagnosticsTests.Static_` с тестом 100-error границы (например, ×101 ячеек со скалярным полем на отсутствующем значении при `MissingValues=Error`), ассертами code/location/DataPath, «no 101st evaluation» и «no output on failure».

**3-СР. Несогласованность чтения/записи PrintArea: workbook-scoped PrintArea с незакавыченной формой принимается reader'ом и отклоняется на этапе рендера**
- Где: чтение — `MonoReport/Services/Xlsx/XlsxPackageReader.cs:95–102` (регулярка `plain`-группа допускает `Report!$B$3:$D$5` без `localSheetId`); запись — `MonoReport/Services/Xlsx/XlsxDocumentRenderer.cs:165–173` (match только по `localSheetId` **или** `'" + name + "'!"` — кавыченной форме).
- Воспроизведение (выполнено): workbook-scoped `Report!$B$3:$D$5` → `Merge` падает с `code=template.xlsx_invalid_print_area, msg="The PrintArea of the worksheet is missing from the workbook."` — при этом reader именно эту область успешно распарсил. Форма без кавычек — это сама форма из примеров ECMA-376 (`Sheet1!$A$1:$D$5`), т.е. валидная и поддерживаемая область отвергается по ложной причине локального несоответствия.
- Исправление: в `UpdatePrintAreas` повторить ту же дизъюнкцию, что и в reader'е (match по plain-имени листа), либо консистентно отвергать глобально-именованные области там, где reader их уже не намерен поддерживать (см. также 4-НИЗ №7).

**4-СР. Проверка числовой точности пропускает часть «небезопасных» десятичных значений JSON**
- Где: `MonoReport/Services/Xlsx/XlsxNumberSafety.cs:14–26` — проверяется G15 round-trip **уже преобразованного double**, а не значащие цифры исходной десятичной записи.
- Воспроизведение (выполнено): `{"Number":10000000000000000555}` → **принято**, в `<v>` пишется сырой текст; Excel хранит/показывает 1E+20 → идентификатор молча округлился (555 потеряны). `1.0000000000000000…3` (46 значащих цифр) → принято и переписано как 1.0. При этом `123456789012345678` и `9007199254740993` корректно отвергаются — детекция недетерминирована относительно самого условия спецификации §3.2.5 («reject unsafe conversion rather than silently truncate»).
- Исправление: дополнить проверку подсчётом значащих цифр сырой десятичной записи (≤15), не меняя проверку диапазона/underflow; пополнить негативный `[Theory]` случаями-«поглотителями».

**5-СР. Замороженные панели (frozen panes) удаляются молча; selection `sqref`/`topLeftCell` не трансформируются**
- Где: `MonoReport/Services/Xlsx/XlsxSheetCropper.cs:142–160` (`CropViews`: `pane` удаляется **без** `context.Report`; активная ячейка сдвигается или удаляется, `sqref` не трогается).
- Воспроизведение (выполнено): лист с `<pane TopLeftCell="C4" ... State="frozen">` → Merge успешен, 0 warnings; заморозка молча исчезла. Это явное ужесточение политики плана §4.5.4/spec §3.1.3 нарушает собственный принцип чекпоинта M1.7 («coordinate-bearing worksheet structures … are rejected with an actionable message instead of being silently carried over»), зафиксированный в progress-записи 9. Панель тоже содержит координату (`topLeftCell`), т.е. относится к «coordinate-bearing structures».
- Исправление (M1-граница): репортировать `XlsxUnsupportedFeature` при непустом `pane` (как сделано для `rowBreaks`/`hyperlinks`), либо, если致远 решено поддерживать, сдвигать `topLeftCell`/`sqref` по плану кропа; в обоих случаях убрать молчаливое удаление.

**6-СР. Вторая полная копия пакета при XLSX Merge (нарушение плана §3.1.5)**
- Где: `MonoReport/Services/MergeEngine.cs:42–47` — первый буфер (`template → output`); `XlsxTemplateHandler.Merge` (`MonoReport/Services/Xlsx/XlsxTemplateHandler.cs:58–61`) — второй `MemoryStream` + `template.CopyTo(output)` поверх уже «private buffered package»; книга открывается дважды (read-only из копии 1, read-write копии 2).
- План §3.1.5 явно запретил: «Keep one buffered package per operation where possible; **do not add a second format-specific input copy**». Пик потребления память и удвоенный I/O ровно на том пути, где M5/M6 понадобятся замеры ресурсов (§8.1); в M2 renderer/geometry переиспользуются без изменений — лучше исправить до M2.
- Исправление: рендер без дублирования — либо открыть `output` (копия 1) редактируемо после закрытия read-only хэндла (план уже требует «close all SDK handles before publishing bytes»), либо передавать в handler готовый байтовый буфер `output` (одна копирование на весь Merge, включая Word-путь).

**7-СР. `extLst` Coordinate-feature не отвергается: координатно-зависимые расширения переносятся молча**
- Где: `MonoReport/Services/Xlsx/XlsxSheetCropper.cs:20–25` — список `CoordinateFeatures` не содержит `extLst`. При этом `<extLst>` на уровне worksheet — стандартное место x14 conditionalFormatting/sparklines/dataValidations с координатами.
- Влияние: после кропа/смещения расширенные структуры останутся со старыми диапазонами — «silently carrying over affected metadata» (в отличие от прямых элементов CF/Validation, которые отвергаются). Schema-валидация это не ловит (аналогично probe «unsafe-shift.xlsx» из architecture-assessment §2).
- Исправление: добавить `extLst` в список M1-отвержений (простое, в границах текущей политики «reject with a remedy»), либо отложить до M4 с явной записью как позицию Fidelity — но сейчас он не только не поддерживается, он ещё и не отвергается.

### Low

**8-НИЗ. Текст `SafeMessage` и Problem Details при невалидном пакете — «DOCX-only» формулировка для всех форматов**
- Где: `MonoReport/Models/TemplateException.cs:81` — `TemplateInvalidPackage → "The content is not a readable DOCX package."` При этом TemplateAnalyzer/`PackageFormatDetector` теперь применяют этот код и к не-Word входам (например, плохой XLSX), и в HTTP-Problem он будет продолжать говорить о DOCX. Код совместим (тот же код, тесты V8 проходят), текст вводит в заблуждение. Минимум: нейтральная формулировка («Word or XLSX package»).
- Примечание: `TemplateAnalyzer.Analyze` (`MonoReport/Services/TemplateAnalyzer.cs:38–41`) безопасен — `error.Errors` никогда не пуст (обе цепочки конструкторов заполняют хотя бы одним элементом).

**9-НИЗ. Phonetic runs (`rPh`) удваивают текст при чтении shared/inline строк**
- Где: `MonoReport/Services/Xlsx/XlsxPackageReader.cs:82,88` — `Descendants<Text>()`/`Descendants(S+"t")` захватывают и `<rPh>`-фрагмент. Для CJK с чтением (phonetic info) значение поля будет с дублированием. Фрагмент редкий; достаточно отфильтровать по `si/rPh`/`is/rPh` при адаптации.

**10-НИЗ. Warning `MergeDiagnostic` не имеет поля `Code` (отступление от плана §3.2.3) и семантически сомнительно использует `FieldCode`**
- Где: `MonoReport/Models/MergeDiagnostic.cs` — позиционный набор (Message, Field, FieldCode, Location, Part, DataPath), `Code` отсутствует; в `XlsxTemplatePlanner.cs:76–78` в `FieldCode` кладётся исходник токена `{Name}` вместо кода («нет кода» vs «текстовый исходник»). Условие плана «nullable non-positional warning `Code`» может быть не реализовано — но это не блокер M1, т.к. коды нужны в M2+ и Query-формат API не уходил из контракта. Записать как осознанное решение (или добавить поле, закрыть план) в M2.
- Также: warning заказывается per‑cell, что фиксирует упорядочение и повторение — сейчас это дислокация (dislocation) локальная, но для путьов сортировки, совместимости с частью血流 (§3.2.4) достаточно.

**11-НИЗ. Неверная рабочая книга: два `<sheet>` на один `WorksheetPart` не отвергаются в reader'е**
- Где: `MonoReport/Services/Xlsx/XlsxPackageReader.cs:29–65` — unique SheetId/name не покрывает дубликат `r:id` на один part. Быть при дубликате с неравными областями сработает, но со странной диагностикой («The source cell is missing from the worksheet.») вместо требуемого «reject malformed/inconsistent main parts safely» (план §3.1.4). Минимум: проверка уникальности целевого part URI при чтении.

**12-НИЗ. Внешние шаблонные ячейки (с `‑ Поле`), попавшие под удаление КРАЕМ, отбрасываются молча**
- Где: `XlsxTemplatePlanner.Plan` парсит только ячейки **внутри** PrintArea (`XlsxTemplatePlanner.cs:21`); ячейка `{Name}` вне PrintArea обрезается (§3.1.2 — пользователь не должен), но нет никакого warning «все содержимое вне PrintArea отброшено» / «шаблонное поле вне PrintArea отброшено» (§8.3 «non-fatal anomalies should produce warnings even on success»). Неконфликтная легкая минимальная мера: warning-подобное сообщение «N ячеек вне PrintArea отброшено» бездиагностика вталкивания — уменьшить тихую потерю содержимого.

---

## Пробелы покрытия (не требуют сейчас фиксации, но должны не потеряться)

1. **Паритет по типам входных потоков**: `Streams_Current_position_nonseekable_and_caller_ownership` (Feature012CompatibilityTests:173–192) покрывает только DOCX; для XLSX-путей нет проверки current-position/nonseekable (M1.2 заявлен как общий).
2. **Мульти-лист кроп**: все Static/StaticCrop-тесты одно-листовые. Лист 2/3 с merge-ами через границу, координатными features и PrintArea как таковые не проверены.
3. **100-error граница XLSX** *(см. 2-СР)* и **observing no further evaluation** — отсутствуют.
4. **extLst / обращённые comments part / orphaned parts** после кропа не покрыты (и связаны с 7-СР).
5. **`Analyze` XLSX с невалидной структурой внутри worksheet** (bad `c`-атрибуты, 手机 `t="d"`/`e`) — часть его ловится, часть молча проходит; стабийных негативных тестов на malformed worksheet XML почти нет.
6. **Известное ограничение M1.6** (прогресс, запись 6): «две ячейки с одной shared-строкой рендерятся одинаково» — тест с названием «adapted per cell» (Feature012FoundationTests:187–205) не проверяет именно разхождение; при переходе в M2 нужно сравнить повторенные значения разlight.
7. М1.5 отсутствует ввиду известной зависимости (см. ниже).

## Готовность M1

- **Завершено и проверено**: M1.1 (Word‑family/XLSX-дискриминация по контенту, main-part type сохраняется), M1.2 (лег affinity/внутренние перезагрузки DI в TryAdd, потоки), M1.3 (инвентарь + финальная прямоугольная PrintArea + проверки измерений), M1.4 (нейтральный `DataScope`/per-cell parse-once), M1.6 (native values/защита чисел — с замечанием 4‑СР), M1.7 (кроп + shared-string списание — с замечаниями 3-СР/5-СР/7-СР), M1.8 (CLI/HTTP/MIME/имена/OpenAPI/безопасные коды).
- **Отстающие чекпоинты/условия**:
  1. **M1.5** (открыт; sequencing): зависимость от распознavezания маркеров M2.4 реальна и не сидится «блоковом» образом — M1.5 требует язык границ, чтобы проектировать requirements/occurrences по ним (план M1.5 §5.1). Пустой result «valid/без requirements» — M1.5 — а предсказанная частичная проекция потребует дважды пересобрать логику за M2.4. Наиболее здоровая последовательность: переместить (или renumber) M1.5 к M2.4 — реализовать проекцию сразу после распознavezания контролов и закрыть её вместе с fixture `analyzer-scopes.template.xlsx`. rudeRecorded product decision (progress entry 8) требует PO-фиксации, не молчаливого сдвига; при этом текущий «valid/empty» на XLSX с маркерами отдельно не генерирует «fail explicitly»-ошибку про нераспознанные маркеры — это стоит включить в закрытие M1.5/M2.4.
  2. Один **High** (1-ВЫС) и 2‑СР (несделанный класс диагностик) до M2.
  3. Расчетное/Excel-ограничение (E1) ожидание — затем «неинтерактивно» (принято, в плане).
- **M2 можно начинать безопасно** после снятия 1‑ВЫС/2‑СР и 6‑СР (второй копии пакета), поскольку архитектура M2 сущ-но построена (XlsxTemplatePlan/Исходные идентичности/кроп-план/контекст/ошибки/безопасная нумерация/сообщение организованы).

## вопрос об ApiHostTests.Real_kestrel… (запрошенный проверкой)

- Классификация по имеющимся следам, без перепроверки на максимум-надёжность: **не регрессия Feature 012**.
  1. Feature 012 не менял host/request-bound обработку: `DocumentEndpoints` получил только выбор MIME/имени после `Merge`; `ApiOperationTransformer`/`ApiProblemFactory` — метаданные OpenAPI/тексты; host-конфиг/`MonoReport__MaxRequestBytes` не затронуты.
  2. Тест уже падал **до** Feature 012: `.docs/features/003-http-api/review.md`, пункт 48 (финальный прогон Feature 003, единственный сбой — тот же тест, `HttpClient` → `ResponseEnded`); файлы `/private/tmp/feature003-finding42-evidence/` интерстально удалены, доступен только запись в ревью.
  3. Три освобожденных прогона 2026-10-21 (готовые результаты ниже) — все Passed.
- Последняя оценка: тест выполняет настоящий Kestrel + отказ-в-размер-ограничение (aborted streamed requests), где сбой "Connection reset by peer" вместо 413/`ResponseEnded` — известная race между abort'ом Kestrel и HTTP-клиентом; вероятно environment-dependent flake, не детерминированный дефект. Финально решать это — только с повторяющимся наблюдением; classify as «not a Feature 012 regression» данный микровердикт и пересчитывается при следующем возникновении.

## Верификация (команды и результаты, 2026‑10‑21)

1. `dotnet build MonoReport.slnx -c Release` → Build succeeded, 0 warnings, 0 errors (5 s).
2. `dotnet test MonoReport.Tests/MonoReport.Tests.csproj -c Release --no-build --filter 'FullyQualifiedName~Feature012FoundationTests|FullyQualifiedName~Feature012CompatibilityTests|FullyQualifiedName~TemplateLanguageParserTests|FullyQualifiedName~UnifiedTemplateTests|FullyQualifiedName~Feature010IntegrationTests|FullyQualifiedName~Feature011'` → **Passed 145/145** (повторяет заявленный v1v2.trx).
3. Полные пакеты (RealHost исключён фильтром проекта): `MonoReport.Tests` → **365/365 PASS**; `MonoReport.Api.Tests` → **27/27 PASS**; `MonoReport.Cli.Tests` → **15/15 PASS** (совпадает с объявлением progress, п. 8).
4. `dotnet test MonoReport.Cli.Tests/MonoReport.Cli.Tests.csproj -c Release --no-build --filter 'FullyQualifiedName~Feature012CliTests|FullyQualifiedName~CliTests|FullyQualifiedName~CliProcessTests'` → **15/15 PASS** (V6).
5. `dotnet test MonoReport.Api.Tests/MonoReport.Api.Tests.csproj -c Release --no-build --filter 'FullyQualifiedName~Feature012HttpTests|FullyQualifiedName~HttpApiTests|FullyQualifiedName~ValidationContractParityTests|FullyQualifiedName~OpenApiContractTests'` → **19/19 PASS** (V7+V8).
6. `dotnet test MonoReport.Api.Tests/... --filter FullyQualifiedName~ApiHostTests` × 3 → **3× Passed (1/1)**.
7. **Не протестировано/не выполнено** (по ограничениям review'а): любые смоделированные XLSXы с `extLst`/phonetic-строками/орфанными parts (пробелы в тестах), ручной Excel, балансовые ресурсные замеры, ре-calculation; вывод и условия вопроса «не протестировано» приведены выше.

**Гигиена review'а**: только read-only. Репозиторий оставлен в исходном состоянии (`git status` до/после идентичен); scratch-эксперименты выполнены в `/private/tmp/feature012-review-scratch/` вне репозитория, ничего не staged/committed; никаких правок файлов, документации или процента сделано.