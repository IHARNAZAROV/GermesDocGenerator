const fs = require("fs");

const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const mammoth = require("mammoth");

/**
 * Remove <w:proofErr> elements that Word inserts as spell/grammar markers.
 * These elements can fragment placeholder tags (e.g. {{currentDate}}) across
 * multiple XML runs, causing Docxtemplater to throw "Multi error".
 */
function cleanProofErrors(zip) {
    const candidates = Object.keys(zip.files).filter((f) =>
        f.startsWith("word/") && f.endsWith(".xml")
    );
    candidates.forEach((file) => {
        let xml = zip.files[file].asText();
        if (!xml.includes("w:proofErr")) return;
        xml = xml.replace(/<w:proofErr[^/]*\/>/g, "");
        zip.file(file, xml);
    });
}

// Парсер с поддержкой точечной нотации: {{seller.fullName}}, {{deal.number}} и т.д.
function dotParser(tag) {
    return {
        get(scope) {
            return tag.split('.').reduce((obj, key) => {
                if (obj == null) return '';
                return obj[key] !== undefined ? obj[key] : '';
            }, scope);
        }
    };
}

/**
 * Внутренняя функция: читает шаблон, заполняет данными, возвращает готовый doc.
 * Используется и generateWord, и previewWord.
 */
async function _renderTemplate(templatePath, data) {
    const content = await fs.promises.readFile(templatePath, "binary");
    const zip = new PizZip(content);
    cleanProofErrors(zip);

    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
        delimiters: { start: '{{', end: '}}' },
        parser: dotParser,
    });

    doc.render(data);
    return doc;
}

// ============================================================
//  generateWord — заполняет шаблон и сохраняет .docx на диск
// ============================================================
async function generateWord(templatePath, outputPath, data) {
    try {
        const doc = await _renderTemplate(templatePath, data);

        const buffer = doc.getZip().generate({
            type: "nodebuffer",
            compression: "DEFLATE"
        });

        await fs.promises.writeFile(outputPath, buffer);

        return { success: true, path: outputPath };

    } catch (error) {
        return { success: false, error: _extractErrorMessage(error) };
    }
}

// ============================================================
//  previewWord — рендерит шаблон и конвертирует результат в HTML
//  через mammoth. Таблицы, жирный текст, заголовки сохраняются.
// ============================================================
async function previewWord(templatePath, data) {
    try {
        const doc = await _renderTemplate(templatePath, data);

        const buffer = doc.getZip().generate({ type: "nodebuffer" });
        const { value: html } = await mammoth.convertToHtml({ buffer });

        return { success: true, html };

    } catch (error) {
        return { success: false, error: _extractErrorMessage(error) };
    }
}

// Docxtemplater wraps multiple issues in a "Multi error" — extract them
function _extractErrorMessage(error) {
    if (error.properties && error.properties.errors) {
        return error.properties.errors
            .map((e) => e.message || String(e))
            .join("; ");
    }
    return error.message;
}

module.exports = { generateWord, previewWord };
