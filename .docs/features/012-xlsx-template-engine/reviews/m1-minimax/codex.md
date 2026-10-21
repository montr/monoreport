**Вердикт: Needs revision.**

Проверены staged, unstaged и все пять untracked-файлов, а также связанные DOCX, DI, CLI и HTTP пути. Найдены воспроизводимые дефекты статического среза M1. Отсутствие возможностей, запланированных на M2–M6, само по себе дефектом не считалось.

**Findings — по серьёзности**

1. **High — проверка точности пропускает небезопасные числа.**  
   [XlsxNumberSafety.cs:16](/Users/blackish/Projects/monoreport/MonoReport/Services/Xlsx/XlsxNumberSafety.cs:16)  
   `10000000000000001` и `1.00000000000000001` успешно записываются как Excel numbers. Проверка сравнивает два уже округлённых `double`, поэтому потерянные при первом преобразовании цифры не обнаруживаются. Это нарушает требование отказа вместо потери точности. Дополнительно корректный `0e10` ошибочно считается underflow: `IsZero` учитывает цифры экспоненты.  
   **Минимальное исправление:** нормализовать исходную десятичную мантиссу и экспоненту до преобразования в `double`; проверять сохранность исходного значения и определять ноль по мантиссе.

2. **High — неподдерживаемые control-маркеры молча удаляются.**  
   [XlsxTemplatePlanner.cs:21](/Users/blackish/Projects/monoreport/MonoReport/Services/Xlsx/XlsxTemplatePlanner.cs:21)  
   Planner рассматривает только клетки внутри PrintArea. В воспроизведении `$l(10).Items` в A3, `{Name}` в B3 и два элемента `Items` дают успешный результат с одной пустой клеткой: маркер удалён, повторение проигнорировано. Analyze также возвращает `valid=true`.  
   Реализация повторений относится к M2; однако M1 прямо требует отвергать ещё неподдерживаемые конструкции.  
   **Минимальное исправление:** до crop обнаруживать управляющие маркеры и возвращать явную ошибку до появления соответствующей реализации.

3. **High — crop создаёт невалидный worksheet XML.**  
   [XlsxSheetCropper.cs:100](/Users/blackish/Projects/monoreport/MonoReport/Services/Xlsx/XlsxSheetCropper.cs:100)  
   Валидный шаблон с PrintArea B3:D5 и единственным описанием ширины колонки A после генерации содержит `<cols/>`. `OpenXmlValidator` сообщает отсутствие обязательного `col`; Merge при этом успешно возвращает документ. Это обычный случай форматирования внешней управляющей колонки.  
   **Минимальное исправление:** удалять контейнер `cols`, когда после обрезки не осталось дочерних колонок.

4. **Medium — повреждённый worksheet relationship выходит необработанным исключением.**  
   [XlsxPackageReader.cs:35](/Users/blackish/Projects/monoreport/MonoReport/Services/Xlsx/XlsxPackageReader.cs:35), [XlsxTemplateHandler.cs:30](/Users/blackish/Projects/monoreport/MonoReport/Services/Xlsx/XlsxTemplateHandler.cs:30)  
   Замена `Sheet.Id` на отсутствующий `rId404` вызывает `ArgumentOutOfRangeException` и в Analyze, и в Merge. Analyze не возвращает invalid result; HTTP-путь попадает в общий обработчик 500.  
   **Минимальное исправление:** проверять наличие relationship до `GetPartById` и нормализовать известные ошибки повреждённого пакета в существующий контракт диагностики.

5. **Medium — frozen panes удаляются без отказа, selection остаётся несогласованным.**  
   [XlsxSheetCropper.cs:146](/Users/blackish/Projects/monoreport/MonoReport/Services/Xlsx/XlsxSheetCropper.cs:146)  
   Валидный источник с frozen pane и selection B3 после crop содержит `activeCell="A1"`, но прежний `sqref="B3"` и `pane="bottomRight"`; сам pane удалён. Генерация успешна. Это потеря настройки и неверная трансляция координат, хотя schema validation проходит.  
   **Минимальное исправление для M1:** явно отвергать неподдерживаемые panes. Для обрабатываемых selections согласованно преобразовывать `activeCell`, `sqref` и pane-контекст.

6. **Medium — после удаления внешней формулы сохраняется устаревший calcChain.**  
   [XlsxDocumentRenderer.cs:69](/Users/blackish/Projects/monoreport/MonoReport/Services/Xlsx/XlsxDocumentRenderer.cs:69)  
   Валидный источник: формула A1 вне PrintArea, соответствующий calculation-chain entry и поле B3 внутри области. После crop формулы нет, A1 содержит текст `Ada`, но chain по-прежнему указывает на A1 как на формулу. SDK schema validation этого нарушения не обнаруживает. Проверка repair prompt в Excel не выполнялась.  
   **Минимальное исправление:** для текущего статического среза удалить obsolete calculation-chain part вместе с relationship либо отвергать такой пакет. Реализацию формул M2 для этого переносить вперёд не требуется.

7. **Medium — Analyze не выполняет существующие проверки допустимости шаблона.**  
   [XlsxTemplateHandler.cs:22](/Users/blackish/Projects/monoreport/MonoReport/Services/Xlsx/XlsxTemplateHandler.cs:22), [XlsxDocumentRenderer.cs:46](/Users/blackish/Projects/monoreport/MonoReport/Services/Xlsx/XlsxDocumentRenderer.cs:46)  
   Шаблоны с формулой, hyperlink и неподдерживаемым defined name получают `valid=true`, но Generate отвергает их независимо от данных. Проверки находятся только в renderer/cropper.  
   Это отдельная проблема preflight, **не известная незавершённость requirements/occurrences M1.5**.  
   **Минимальное исправление:** перенести data-independent проверки в общий этап планирования, используемый обеими операциями.

8. **Medium — числовая ошибка раскрывает значение JSON в диагностике.**  
   [XlsxNumberSafety.cs:35](/Users/blackish/Projects/monoreport/MonoReport/Services/Xlsx/XlsxNumberSafety.cs:35)  
   Небезопасное значение `1234567890123456789` передаётся как `TemplateException.Field`. Оно оказывается в `Errors[].Field`, CLI stderr и HTTP Problem Details. План запрещает включать raw JSON values. При этом полезная рекомендация передать значение текстом отсутствует в сериализованном `Errors[].Message`.  
   **Минимальное исправление:** указывать шаблонное выражение и `DataPath`, сохраняя безопасное объяснение причины и способ исправления без значения данных.

9. **Medium — XLSX принимает JSON root другого типа.**  
   [XlsxTemplateHandler.cs:47](/Users/blackish/Projects/monoreport/MonoReport/Services/Xlsx/XlsxTemplateHandler.cs:47)  
   Для обычного шаблона с `{Name}` вход `[]` успешно генерирует пустую клетку. Спецификация требует JSON object и отдельно запрещает распространять это новое ограничение на DOCX.  
   **Минимальное исправление:** проверять `RootElement.ValueKind == Object` внутри XLSX handler до binding.

10. **Medium — обязательный учёт ресурсов M1 практически отсутствует.**  
    [XlsxOperationContext.cs:8](/Users/blackish/Projects/monoreport/MonoReport/Models/Xlsx/XlsxOperationContext.cs:8), [XlsxPackageReader.cs:57](/Users/blackish/Projects/monoreport/MonoReport/Services/Xlsx/XlsxPackageReader.cs:57)  
    Увеличивается только `SourceCells`; `OutputCells` нигде не заполняется. Фактические распакованные байты частей не учитываются. Это не соответствует обязательному учёту из [плана:285](/Users/blackish/Projects/monoreport/.docs/features/012-xlsx-template-engine/implementation-plan.md:285).  
    **Минимальное исправление:** начать учитывать фактические чтения частей и создаваемые клетки в operation context. Численные rejection limits R1 остаются отдельным отложенным решением; вводить их сейчас не требуется.

11. **Low — reader и renderer по-разному разрешают PrintArea.**  
    [XlsxPackageReader.cs:97](/Users/blackish/Projects/monoreport/MonoReport/Services/Xlsx/XlsxPackageReader.cs:97), [XlsxDocumentRenderer.cs:167](/Users/blackish/Projects/monoreport/MonoReport/Services/Xlsx/XlsxDocumentRenderer.cs:167)  
    Workbook-scoped `_xlnm.Print_Area` со значением `Report!$B$3:$D$5` принимается reader и Analyze, но renderer ищет только вариант с кавычками и сообщает, что PrintArea отсутствует. Источник проходит SDK schema validation.  
    **Минимальное исправление:** передавать идентичность разрешённого defined name из inventory, используя единое разрешение scope и имени листа.

**Пробелы покрытия**

1. Текущие тесты не обнаруживают перечисленные случаи точности, игнорирования маркеров и повреждения XML после crop.
2. XLSX не имеет проверок границ 99/100/101 ошибок с доказательством остановки дальнейшего вычисления. Общая константа сама по себе этого не устанавливает.
3. `Streams_` проверяет DOCX, но не XLSX на nonseekable/current-position stream. Не установлена изоляция смешанных DOCX/XLSX операций при повторном и конкурентном использовании.
4. Для crop недостаточно проверены реальные widths/heights/hidden states, print/view settings, relationships и workbook-level ссылки. Проверка style ID и schema validity не покрывает их семантику.
5. Не установлено, что XLSX-ошибки Core, CLI и HTTP сохраняют одинаковые безопасные и actionable детали.

**Сбой ApiHostTests**

Вне sandbox точный тест прошёл **11 раз из 11**. Начальный запуск внутри sandbox был прерван VSTest на `SocketException (13): Permission denied` до выполнения теста; это другой сбой.

Сам тест, `Program.cs` и `MultipartRequestReader.cs` относительно HEAD не изменены. Проверка в [ApiHostTests.cs:45](/Users/blackish/Projects/monoreport/MonoReport.Api.Tests/Services/ApiHostTests.cs:45) предполагает получение HTTP response при превышении лимита. Microsoft отдельно документирует возможность reset connection при превышении `MaxRequestBodySize`. [Документация Microsoft](https://learn.microsoft.com/en-us/aspnet/core/mvc/models/file-uploads?view=aspnetcore-10.0#connection-failure)

Вывод: **есть основания считать тест потенциально нестабильным на транспортном уровне; регрессия Feature 012 не установлена**. Причина конкретного IDE-эпизода остаётся неподтверждённой. Перед изменением assertions стоит сохранять endpoint, `streamed` и вывод дочернего Kestrel при исключении. Простое подавление любого `HttpRequestException` ослабит проверку.

**Готовность M1**

1. Базовые dispatch четырёх Word-форматов, публичные конструкторы, DI, scalar binding и статические CLI/HTTP happy paths реализованы и подтверждены focused checks.
2. M1.6 и M1.7 нельзя считать закрытыми: воспроизведены небезопасные числа, неверный успешный результат и невалидный XML. Для inventory/diagnostics также остаются указанные нарушения.
3. M1.5 остаётся известным открытым checkpoint. Зависимость scoped projection от распознавания маркеров обоснована; полный nested grammar для начальной проекции не нужен. Возможны два согласованных порядка: вынести минимальное распознавание перед M1.5 либо перенести scoped projection после M2.4, сохранив полную проверку fixture в M3.14. Такое изменение последовательности требует решения Product Owner.
4. **Сейчас milestone gate M1 не закрыт.** M2 можно начинать после исправления нарушений M1 и согласования последовательности M1.5. Формулы, повторения, rich-text обработка, multipart и финальная Excel-проверка остаются в предусмотренных последующих этапах.

**Верификация**

Из корня репозитория выполнены следующие команды:

```text
dotnet test MonoReport.Tests/MonoReport.Tests.csproj -c Release --no-restore --filter 'FullyQualifiedName~Feature012|FullyQualifiedName~UnifiedTemplateTests|FullyQualifiedName~DependencyInjectionTests' --logger 'trx;LogFileName=core-focused.trx' --results-directory /private/tmp/feature012-review-evidence
```

Результат: **73 passed**.

```text
dotnet test MonoReport.Api.Tests/MonoReport.Api.Tests.csproj -c Release --no-restore --filter 'FullyQualifiedName~Feature012|FullyQualifiedName~OpenApiContractTests|FullyQualifiedName~ValidationContractParityTests|FullyQualifiedName~HttpApiTests' --logger 'trx;LogFileName=api-focused.trx' --results-directory /private/tmp/feature012-review-evidence
```

Результат: **19 passed**.

```text
dotnet test MonoReport.Cli.Tests/MonoReport.Cli.Tests.csproj -c Release --no-restore --filter 'FullyQualifiedName~Feature012|FullyQualifiedName~CliTests|FullyQualifiedName~CliProcessTests' --logger 'trx;LogFileName=cli-focused.trx' --results-directory /private/tmp/feature012-review-evidence
```

Результат: **15 passed**.

```text
dotnet test MonoReport.Tests/MonoReport.Tests.csproj -c Release --no-build --no-restore --filter 'FullyQualifiedName~Feature010IntegrationTests|FullyQualifiedName~Feature011' --logger 'trx;LogFileName=docx-regressions.trx' --results-directory /private/tmp/feature012-review-evidence
```

Результат: **59 passed**.

```text
dotnet test MonoReport.Api.Tests/MonoReport.Api.Tests.csproj -c Release --no-build --no-restore --filter FullyQualifiedName~ApiHostTests.Real_kestrel_enforces_whole_request_bound_including_streamed_requests_and_recovers --logger 'trx;LogFileName=realhost-current.trx' --results-directory /private/tmp/feature012-review-evidence
```

Результат вне sandbox: **1 passed**; затем десять повторов той же проверки с файлами `realhost-current-1.trx` … `realhost-current-10.trx`: **10 passed**.

Дополнительно выполнен временный runner с 16 probes через публичные API и `OpenXmlValidator`:

```text
dotnet run --project /private/tmp/feature012-review-probes/ReviewProbes.csproj --configuration Release --no-restore > /private/tmp/feature012-review-evidence/probes.log
```

Воспроизведения и XML доступны в [probes.log](/private/tmp/feature012-review-evidence/probes.log). Полный suite повторно не запускался; Excel, браузерные и тяжёлые capacity checks не выполнялись.

Хеши всех 345 проверенных файлов, включая Git index, совпали с исходным снимком. Новых неигнорируемых файлов в репозитории нет; исходники, документация и тесты не изменялись.