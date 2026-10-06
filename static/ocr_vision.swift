// OCR via le framework Vision d'Apple — intégré à macOS, sans installation.
// Usage : swift ocr_vision.swift <image> [langues]   ex. "fr-FR,ar-SA,en-US"
import Foundation
import Vision
import AppKit

guard CommandLine.arguments.count > 1 else { FileHandle.standardError.write("usage: ocr_vision <image>\n".data(using: .utf8)!); exit(2) }
let path = CommandLine.arguments[1]
let langs = CommandLine.arguments.count > 2
    ? CommandLine.arguments[2].split(separator: ",").map(String.init)
    : ["fr-FR", "en-US"]

guard let img = NSImage(contentsOfFile: path),
      let tiff = img.tiffRepresentation,
      let bmp = NSBitmapImageRep(data: tiff),
      let cg = bmp.cgImage else {
    FileHandle.standardError.write("image illisible\n".data(using: .utf8)!); exit(1)
}

let req = VNRecognizeTextRequest { (request, _) in
    guard let obs = request.results as? [VNRecognizedTextObservation] else { return }
    for o in obs {
        if let top = o.topCandidates(1).first { print(top.string) }
    }
}
req.recognitionLevel = .accurate
req.usesLanguageCorrection = true
req.recognitionLanguages = langs

let handler = VNImageRequestHandler(cgImage: cg, options: [:])
do { try handler.perform([req]) } catch { exit(1) }
