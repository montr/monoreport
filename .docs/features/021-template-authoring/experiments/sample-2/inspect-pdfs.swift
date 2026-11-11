// Read existing PDF pagination/text only. No conversion, drawing or raster rendering.
import Foundation
import PDFKit

let directory = URL(fileURLWithPath: CommandLine.arguments[1], isDirectory: true)
let filenames = ["skill-sample-2", "original.result", "skill-sample-2.result",
                 "expanded.result", "empty-inner.result", "empty-outer.result"]
var results: [[String: Any]] = []
var sourcePages: [String] = []
for name in filenames {
    let url = directory.appendingPathComponent(name + ".pdf")
    guard let document = PDFDocument(url: url) else { fatalError("Cannot read \(url.path)") }
    let pages = (0..<document.pageCount).map { document.page(at: $0)?.string ?? "" }
    try pages.joined(separator: "\n\n--- PAGE BREAK ---\n\n")
        .write(to: directory.appendingPathComponent(name + ".txt"), atomically: true, encoding: .utf8)
    if name == "skill-sample-2" { sourcePages = pages }
    var entry: [String: Any] = ["file": name + ".pdf", "pageCount": document.pageCount,
                               "pageTextLengths": pages.map { $0.count },
                               "method": "PDFKit metadata and text extraction; no rendering"]
    if name == "original.result" { entry["pageTextEqualsSource"] = pages == sourcePages }
    results.append(entry)
}
let data = try JSONSerialization.data(withJSONObject: results, options: [.prettyPrinted, .sortedKeys])
try data.write(to: directory.appendingPathComponent("pdf-inspection.json"))
print(String(data: data, encoding: .utf8)!)
