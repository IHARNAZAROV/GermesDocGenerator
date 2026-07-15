const fs = require("fs");
const path = require("path");

const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

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

async function generateWord(templatePath, outputPath, data) {

    try {

        const content = fs.readFileSync(templatePath, "binary");

        const zip = new PizZip(content);

        cleanProofErrors(zip);

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

        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
            delimiters: { start: '{{', end: '}}' },
            parser: dotParser,
        });

        doc.render(data);

        const buffer = doc.getZip().generate({
            type: "nodebuffer",
            compression: "DEFLATE"
        });

        fs.writeFileSync(outputPath, buffer);

        return {
            success: true,
            path: outputPath
        };

    } catch (error) {

        // Docxtemplater wraps multiple issues in a "Multi error" — extract them
        let message = error.message;
        if (error.properties && error.properties.errors) {
            message = error.properties.errors
                .map((e) => e.message || String(e))
                .join("; ");
        }

        return {
            success: false,
            error: message
        };

    }

}

module.exports = {
    generateWord
};
