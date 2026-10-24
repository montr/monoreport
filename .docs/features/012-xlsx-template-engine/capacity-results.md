# Feature 012 — Capacity measurement history

1. Append each future series with its own date, source revision, scenario definitions, conditions, raw values and findings. Preserve earlier series and partial observations; record corrections explicitly rather than replacing historical results. Compare matching workload shapes and measurement scopes before interpreting differences. Missing metrics are unavailable, never zero unless an actual zero was observed.
2. This document records bounded observations under [Implementation Plan §8.1](implementation-plan.md#81-bounded-resource-measurements). It establishes no numerical rejection policy, SLA or deployment adequacy. R1 ceilings remain unapproved/inactive. Excel E1/A14 remain unexecuted and deferred until after UI implementation. Schema checks do not establish Excel calculation, visual or print acceptance.

## Series 1 — 2026-10-24

### Conditions and method

1. Source commit: `387b4e32e46087085bd6878f8380a776202e88e3`. The working tree and index were clean on entry. Production/test source did not change during this series. Final solution Release build passed with zero warnings/errors and normal V9 passed 1,180 tests, excluding RealHost. Measurements ran sequentially after V9 finished; no sustained concurrency or large-workbook stress test ran.
2. Machine: Apple M4, model Mac16,13, 10 logical processors, 34,359,738,368 physical RAM bytes. OS: macOS 27.0.1, Arm64. SDK: .NET 10.0.102; MSBuild 18.0.7; runtime: .NET 10.0.2. Configuration: Release for both production assemblies and the temporary runner. Date is the Owner's Europe/Moscow calendar date; individual run timestamps were not retained.
3. Temporary runner: `/private/tmp/feature012-m6-evidence/runner/Program.cs`, SHA-256 `304aa12bcf7e8eeb6c7052eb8db1d218f24303abe2dcc132933e7ae14f9758ec`. It directly references existing built Core/Open XML/DI assemblies, links the maintained XlsxFixture/XlsxNestedFixture sources and uses the existing friend-assembly identity only to observe internal stages. No production instrumentation, runtime dependency, benchmark framework, test-suite addition or repository harness was introduced. The temporary project restored framework assets with cleared package sources and built in Release with zero warnings/errors after correcting a runner-only Parse/Read method-name compile error.
4. Each invocation used a fresh dotnet process. Before the measured operation it built the fixture/JSON and service provider, performed one public MergeEngine warmup with the same template shape and 10 main-table logical rows (2 for deep), then forced collection, waited for finalizers and collected again. Each final case has one measured repetition, not a mean or percentile. Baseline 100/1,000/5,000 had one earlier observation each, retained below. No repeat was needed for a product failure; final baseline reruns obtained OS memory counters after the sandbox prevented macOS time from reading kern.clockrate.
5. Total elapsed time uses Stopwatch from JSON parsing through output byte-array extraction, including in-memory input copying, open/read/plan, preflight, binding, scheduling, geometry, value preparation, formula finalization, rendering and package disposal/ZIP compression. Stage times use nested Stopwatch readings. Total includes orchestration/instrumentation overhead, so stages need not sum exactly. Fixture construction, provider setup, warmup, disk writes, ZIP inspection, schema validation and the second public comparison merge are outside the measured total. This stage run uses the same production services/order as XlsxTemplateHandler; public dispatch and warning sorting are outside the staged total, so this is not end-to-end CLI/HTTP latency.
6. Allocated bytes are the process-wide delta of GC.GetTotalAllocatedBytes(true), including measurement/sampler allocations. This is cumulative managed allocation, not retained or peak heap memory. WorkingSet64 is sampled every 5 ms with a final synchronous reading; short transient peaks may be missed. The reported sampled maximum includes the resident baseline after warmup and is not a baseline-subtracted allocation measure. Process.PeakWorkingSet64 returned the raw value 0 on this platform for every final run; it is unavailable as a useful lifetime peak, not evidence of zero memory usage.
7. Final invocations were wrapped in `/usr/bin/time -l` with approved access to OS counters. Its maximum resident set size and peak memory footprint are recorded in bytes. These cover the whole process, including warmup, fixture construction, post-measurement SDK validation and a second public Core merge while stage objects may remain live. They are not interchangeable with the measured-operation sampled peak. Whole-process real seconds include these extra activities and startup. Other system load/temperature and isolated native peak allocation were not measured.
8. Counts are read from actual input JSON and output plans; styled cells are output placements whose source has an explicit style index, including blank styled cells. Populated cells mean emitted authored placements, including blank formatting placeholders, not a dense dimension product or only nonempty values. Output formula/reference-token counts exclude suppressed output formulas; all fixtures here have one parsed reference token per formula. Source/reference counts come from prepared source plans. Actual part bytes are counted during XML stream reads. Source/output uncompressed bytes are independently summed by reading every ZIP entry to EOF, not by trusting declared ZIP sizes.
9. Every final output was reopened and passed OpenXmlValidator. A second public MergeEngine.Merge produced exactly the same uncompressed bytes for every ZIP entry and the same entry count; input/output ZIP timestamps/compressed lengths may differ. All six cases returned zero warnings. These checks do not calculate formulas. Raw metrics, generated packages, input JSON, runner source and available time logs are under `/private/tmp/feature012-m6-evidence/`; the permanent numeric history below remains usable if that temporary directory is removed.

### Workloads

1. Base-100, base-1000 and base-5000 reuse XlsxNestedFixture.Create(false): two companies × seven criteria (14 records per logical row), parent-only zero-width columns, two-row child bodies, two parent header/footer groups, authored merged names and heading merges, a zero-height separator, a second three-record Other table, grand totals and a static cross-sheet Summary total. Logical-row count includes the two parent records; children comprise N−2. Each record has a unique contiguous row path, level 10 for indices 0 and N/2 and level 20 otherwise. Tuple values use Price=i+tuple+1, Quantity=i+1, Delta=1; ordinary Name is Row i. No sparse matrix, empty-product descriptor or warning-volume workload is measured here; those behaviors retain automated correctness evidence.
2. Formula-1000 changes only repeated E11 from Delta binding to scalar formula E10*2, retaining the same input JSON and report geometry. This increases output formula density while preserving the two tables and cross-sheet total.
3. Complex-1000 retains base geometry/JSON/formulas and varies styles/part complexity together: 16 fonts, 256 cell formats, 256 shared strings each with a 512-character payload plus its prefix, authored body style indices set to 255, and static B16 using shared-string index 255. This is one bounded style/part variant, not exhaustive metadata coverage or an isolated style-only benchmark.
4. Deep-100 is a lighter three-level 2×3×4 header tree: root source D:F, branch E:F, leaf F, with root and branch own columns. The retained source is B5:F9, header rows 5/6/7 and a two-row level-40 body at 8/9; F8 binds Price and F9 is F8*2. Its 100 logical rows have 24 tuple records each. The final 34 physical columns comprise two static columns plus 32 expanded columns. It has one worksheet, no second table or cross-sheet total, and is reported separately from base scaling.

### Final measured values

1. All sizes and memory values below are raw bytes, elapsed values are milliseconds unless marked seconds, and counts are integers. No value in these tables is an estimate.

| Case | Logical rows | Matrix records | Descriptor records | Other records | Output dimensions | Emitted cells | Styled cells | Formulas | Reference tokens |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| base-100 | 100 | 1400 | 0 | 3 | Report 208×19; Summary 1×1 | 3444 | 1568 | 71 | 71 |
| base-1000 | 1000 | 14000 | 0 | 3 | Report 2008×19; Summary 1×1 | 34044 | 15968 | 71 | 71 |
| base-5000 | 5000 | 70000 | 0 | 3 | Report 10008×19; Summary 1×1 | 170044 | 79968 | 71 | 71 |
| formula-1000 | 1000 | 14000 | 0 | 3 | Report 2008×19; Summary 1×1 | 34044 | 15968 | 14043 | 14043 |
| complex-1000 | 1000 | 14000 | 0 | 3 | Report 2008×19; Summary 1×1 | 34044 | 15968 | 71 | 71 |
| deep-100 | 100 | 2400 | 0 | 0 | Report 203×34 | 4932 | 0 | 2400 | 2400 |

| Case | Source cells | Source formulas | Source reference tokens | Warnings |
| --- | --- | --- | --- | --- |
| base-100 | 34 | 4 | 4 | 0 |
| base-1000 | 34 | 4 | 4 | 0 |
| base-5000 | 34 | 4 | 4 | 0 |
| formula-1000 | 34 | 5 | 5 | 0 |
| complex-1000 | 34 | 4 | 4 | 0 |
| deep-100 | 17 | 1 | 1 | 0 |

| Case | Source XLSX | JSON UTF-8 | Actual XML part reads | Source all ZIP entries, read | Output XLSX | Output all ZIP entries, read |
| --- | --- | --- | --- | --- | --- | --- |
| base-100 | 2993 | 114783 | 4821 | 6458 | 17023 | 175112 |
| base-1000 | 2993 | 1198502 | 4821 | 6458 | 140426 | 1737028 |
| base-5000 | 2996 | 6238502 | 4821 | 6458 | 699458 | 8819637 |
| formula-1000 | 3000 | 1198502 | 4771 | 6408 | 176664 | 1729342 |
| complex-1000 | 5102 | 1198502 | 156476 | 158409 | 141551 | 1778561 |
| deep-100 | 1916 | 139188 | 2139 | 3212 | 25007 | 198583 |

| Case | Total ms | Managed allocated bytes | WS before bytes | Sampled peak WS bytes | Whole-process max RSS bytes | Whole-process peak footprint bytes | Whole-process real seconds |
| --- | --- | --- | --- | --- | --- | --- | --- |
| base-100 | 24.3055 | 11821072 | 94289920 | 101859328 | 108527616 | 76432248 | 0.36 |
| base-1000 | 304.1451 | 130788296 | 109527040 | 146292736 | 200376320 | 168133640 | 1.33 |
| base-5000 | 1481.2828 | 1061792016 | 135806976 | 331956224 | 564707328 | 532628848 | 4.60 |
| formula-1000 | 786.8124 | 176950016 | 109412352 | 148750336 | 199376896 | 167166936 | 2.04 |
| complex-1000 | 312.5723 | 131099096 | 109740032 | 148160512 | 196329472 | 164152256 | 1.40 |
| deep-100 | 49.3548 | 22974136 | 92700672 | 102776832 | 114524160 | 82543552 | 0.41 |

2. Stage timings retain the observed decimal precision; read-plan includes source XML reading and immutable source/template planning, geometry includes cell/merge placement, and close includes package disposal/compression.

| Stage, ms | base-100 | base-1000 | base-5000 | formula-1000 | complex-1000 | deep-100 |
| --- | --- | --- | --- | --- | --- | --- |
| json | 0.8508 | 8.7904 | 45.0526 | 9.3455 | 9.371 | 1.1468 |
| open | 0.1349 | 0.2096 | 0.2249 | 0.2249 | 0.1842 | 0.1595 |
| read-plan | 0.5356 | 0.5962 | 0.6653 | 0.6287 | 1.2139 | 0.5242 |
| preflight | 0.0856 | 0.0825 | 0.0878 | 0.0852 | 0.0918 | 0.0531 |
| bind | 0.4905 | 11.8043 | 28.0099 | 11.6765 | 11.8197 | 0.771 |
| schedule | 0.0875 | 0.3444 | 7.4384 | 0.3154 | 0.333 | 0.1303 |
| geometry | 7.6652 | 106.2071 | 642.8951 | 104.8783 | 106.9462 | 7.8887 |
| values | 5.947 | 75.275 | 337.7938 | 65.9195 | 81.1041 | 8.9151 |
| formulas | 2.7892 | 42.148 | 174.7277 | 533.9875 | 40.2769 | 24.3034 |
| render | 4.9692 | 53.6292 | 219.5161 | 53.3756 | 55.7625 | 4.2746 |
| close | 0.6087 | 4.8327 | 24.6159 | 6.1583 | 5.2622 | 1.0155 |
| copy | 0.0007 | 0.0118 | 0.0471 | 0.0063 | 0.0139 | 0.0011 |

### Earlier observations retained from the same session

1. These three operations completed the measured production pipeline, schema validation and public package comparison successfully. Initial base-100 was invoked directly with dotnet. Initial base-1000/base-5000 were wrapped in sandboxed macOS time: the child operations succeeded, but time then exited 1 because kern.clockrate access was denied. Their whole-process max RSS/footprint is unavailable. The approved retry supplied final OS counter evidence above; these earlier observations remain distinct.

| Observation | Logical rows | Total ms | Output XLSX bytes | Sampled peak WS bytes | Allocated bytes | OS peak RSS |
| --- | --- | --- | --- | --- | --- | --- |
| Initial base-100 | 100 | 25.9245 | 17024 | 101564416 | 11876512 | unavailable |
| Initial base-1000 | 1000 | 304.8907 | 140430 | 141737984 | unavailable | unavailable |
| Initial base-5000 | 5000 | 1472.7395 | 699456 | 265502720 | unavailable | unavailable |

2. Initial base-100 additionally retained: 1,400 matrix records, 0 descriptors, 3 Other records; Report 208×19 and Summary 1×1; 3,444 emitted cells, 1,568 styled cells, 71 formulas, 71 reference tokens; 34 source cells, 4 source formulas and 4 source reference tokens; 0 warnings; source XLSX 2,995 bytes, JSON 114,783 bytes, actual XML reads 4,821 bytes, source entries 6,458 bytes, output entries 175,112 bytes; pre-operation WS 94,208,000 bytes; Process.PeakWorkingSet64 raw 0. Its stage milliseconds were json 0.9425, open 0.1726, read-plan 0.6402, preflight 0.0942, bind 0.5417, schedule 0.1145, geometry 8.2333, values 6.1285, formulas 3.0209, render 5.2119, close 0.6687, copy 0.0005.
3. The initial base-1000/base-5000 per-stage/allocation/input-size JSON was superseded in temporary files by their final observations before permanent-history authorization. Those earlier detailed values were not captured in the tool output and are unavailable; they are not reconstructed from the final runs. Sandboxed whole-process real seconds were 1.33 and 4.67 respectively before the time-counter error. Future runs must use unique series/observation filenames to retain every raw capture, not reuse those temporary metric filenames.

### Findings and limits

1. Observed: base geometry cost grows from 7.6652 to 106.2071 to 642.8951 ms across the three final sizes. At 5,000 logical rows it is the largest measured stage, followed by value binding (337.7938 ms), rendering (219.5161 ms) and formula finalization (174.7277 ms). This identifies where time is spent; stage timing does not establish which method or source scan causes it. No profiler was run and no hotspot optimization was performed.
2. Observed: formula-1000 finalization takes 533.9875 ms versus base-1000's 42.148 ms, with 14,043 versus 71 output formulas. Formula finalization is the dominant stage in that dense variant and in the lighter deep case (24.3034 of 49.3548 ms). Complex-1000 performs 156,476 actual XML-read bytes versus 4,821 for base-1000, while total elapsed is 312.5723 versus 304.1451 ms. One repetition cannot distinguish a small effect from timing noise.
3. Observed: base-5000 emits 170,044 cells in 1,481.2828 ms with a 699,458-byte compressed workbook; managed allocation is 1,061,792,016 bytes, sampled operation WS peak 331,956,224 bytes, and whole-process max RSS 564,707,328 bytes under the broader validation/comparison scope. Cumulative allocation is materially larger than resident memory and compressed output. These quantities measure different things and are not interchangeable resource budgets.
4. Interpretation: the bounded workloads complete on this 32 GiB development machine with schema/public parity checks. They do not establish safe deployment concurrency, worst-case memory, warning-volume behavior, a universal row limit, streaming need or large-workbook throughput. Representative Owner workloads/deployment constraints remain unavailable; R1 is still a deferred Owner decision before numerical ceilings are activated. No limit is proposed or enabled, and no optimization, heavy stress run, external host, browser or spreadsheet application was used.

### Fixture provenance

1. The following SHA-256 values identify the final inputs/outputs actually measured. ZIP container bytes can vary with timestamps; compare uncompressed content and workload definitions when reproducing rather than treating a compressed hash/length difference alone as a semantic regression. Temporary fixtures are evidence, not a new public authoring convention.

| Case | Template SHA-256 | JSON SHA-256 | Generated SHA-256 |
| --- | --- | --- | --- |
| base-100 | 56670c0db9469b7c4dad2a32d33cb5c44f520ec48206b1bebc79675e883464b0 | 4c0bdeaa08d4223432318c65675ba155fb669c9bded247a89f9e39ac120aada7 | 6d80389d8b02068fa3d7f21d593bab615e32c21a243d10afda4500ddabbf9817 |
| base-1000 | b881c7410dc21c67a227514fce0ca31620f213ae7c53b0575862e70ea564ae03 | 1fe30d2a4aad21ecc3a812b4aaae76cf9a08ccd6f5a80b9161f4fdb6c800b52b | 986ba9e1152e4b173e30e847f2852edc954169697b0224425759021340c7d68a |
| base-5000 | a6a805d0172abd32e30b5531d6a05165ec653c8efce67825059f559dd25bc325 | 24c7f4a8c7e296c236b206a7525d5f0d53ed785e0ae6ddcb222404d79c22a817 | e804e6d2a24681ab30f4550b29d03196004446128964333f2285cb616a37a25d |
| formula-1000 | 67c63315a89ad5128cbabdb2d1a00cdbfa2d0f03a294bf374207b48e52123aa2 | 1fe30d2a4aad21ecc3a812b4aaae76cf9a08ccd6f5a80b9161f4fdb6c800b52b | 77db19cffa06e799b45873690ea16292efe5ae41082fe4411b4e933a9e52dac3 |
| complex-1000 | 2d514b870fad3e692a9a5baf98331205738b073d7ca8c7f8d91f099d0f77c177 | 1fe30d2a4aad21ecc3a812b4aaae76cf9a08ccd6f5a80b9161f4fdb6c800b52b | b65b49ed5461b551eefb04298bd93a62081ed39edc6d21b36089a3640efc2f8c |
| deep-100 | 257a038e0ba2eb227fba49925ce7a1ebbce549e2d5badfe0aa6af35bd42e23ab | 6c508314d2a00f32f799e7bda68722969c0437ddd77e1972f339d570bcef5b26 | 0d67142a2bf4d78f921779dff88f7befc9ca9ecf254818916e4c8c84a0cb941d |
