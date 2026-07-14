const fs = require("fs");
const path = require("path");

const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

async function generateWord(templatePath, outputPath, data) {

    try {

        const content = fs.readFileSync(templatePath, "binary");

        const zip = new PizZip(content);

        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true
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

        return {
            success: false,
            error: error.message
        };

    }

}

module.exports = {
    generateWord
};