const fs = require('fs');
const path = require('path');
const mammoth = require('mammoth');
const wordextractor = require('word-extractor');
const { getDataReforming, getDocBuffer } = require('./Utils');

const INPUT_DIR = `${__dirname}/input/`;
const OUTPUT_DIR = `${__dirname}/output/`;
const TEMPLATE_FILE_DIR = `${__dirname}/template/template.docx`;
const OUTPUT_FILENAME_PRIFIX = process.env.OUTPUT_FILENAME_PRIFIX;
const OUTPUT_FILENAME_SUFFIX = process.env.OUTPUT_FILENAME_SUFFIX;
const TEMPLATE_CONTENTS = fs.readFileSync(TEMPLATE_FILE_DIR, 'binary');

// Read each input file
fs.readdirSync(INPUT_DIR).forEach((fileName) => {
    if (fileName === '.gitignore')
        return;
    console.log(`Input file name : ${fileName}`);
    if (fileName.match('.docx')) {
        mammoth.extractRawText({path: INPUT_DIR + fileName}).then((result) => {
            let rawText = result.value;
            let data = getDataReforming(rawText);
            let buffer = getDocBuffer(TEMPLATE_CONTENTS, data);
            let outputFileName = `${OUTPUT_FILENAME_PRIFIX}${data.name}${OUTPUT_FILENAME_SUFFIX}.docx`;
            fs.writeFileSync(path.resolve(OUTPUT_DIR, outputFileName), buffer);
        });
    } else {
        let extractor = new wordextractor();
        extractor.extract(INPUT_DIR + fileName).then((result) => {
            let rawText = result.getBody();
            let data = getDataReforming(rawText);
            let buffer = getDocBuffer(TEMPLATE_CONTENTS, data);
            let outputFileName = `${OUTPUT_FILENAME_PRIFIX}${data.name}${OUTPUT_FILENAME_SUFFIX}.docx`;
            fs.writeFileSync(path.resolve(OUTPUT_DIR, outputFileName), buffer);
        });
    }
});