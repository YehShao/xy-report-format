const fs = require('fs');
const path = require('path');
const jszip = require('jszip');
const docxtemplater = require('docxtemplater');
const mammoth = require('mammoth');
const { getDataStructure, getAuthorInfo } = require('./Utils');

const INPUT_DIR = `${__dirname}/input/`;
const OUTPUT_DIR = `${__dirname}/output/`;
const TEMPLATE_FILE_DIR = `${__dirname}/template/template.docx`;
const OUTPUT_FILENAME_PRIFIX = process.env.OUTPUT_FILENAME_PRIFIX;
const OUTPUT_FILENAME_SUFFIX = process.env.OUTPUT_FILENAME_SUFFIX;

// Read each input file
fs.readdirSync(INPUT_DIR).forEach((fileName) => {
    console.log(`Input file name : ${fileName}`);
    mammoth.extractRawText({path: INPUT_DIR + fileName}).then((result) => {
        let rawText = result.value; // The raw text
        let rawTextArray = rawText.split('\n').filter((value) => {
            return (value !=='' && !value.match('作者簡介'));
        });
        console.log(rawTextArray);

        let data = getDataStructure();
        rawTextArray.forEach((value, index) => {
            switch(index) {
                case 0:
                    data.title = value;
                    break;
                case 1:
                    data.name = value;
                    data.author = getAuthorInfo(value);
                    break;
                case 2:
                    if (value.length !== 6 || !value.match('士林'))
                        data.organization = '台北士林分會';
                    else
                        data.organization = value;
                    break;
                case 3:
                    data.date = value;
                    break;
                default:
                    data.contents.push({content: value});
                    break;
            }
        });
        console.log(data);

        // Load the template.docx file as a binary
        let templateContent = fs.readFileSync(TEMPLATE_FILE_DIR, 'binary');
        let zip = new jszip(templateContent);
        let doc = new docxtemplater();
        doc.loadZip(zip);
        doc.setOptions({paragraphLoop: true});
        // Set the template variables
        doc.setData(data);
        try {
            doc.render();
        } catch (error) {
            let e = {
                message: error.message,
                name: error.name,
                stack: error.stack,
                properties: error.properties
            };
            console.log(JSON.stringify({error: e}));
            // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
            throw error;
        }

        // Nodejs buffer, you can either write it to a file or do anything else with it.
        let buffer = doc.getZip().generate({type: 'nodebuffer'});
        let filename = `${OUTPUT_FILENAME_PRIFIX}${data.name}${OUTPUT_FILENAME_SUFFIX}.docx`;
        fs.writeFileSync(path.resolve(OUTPUT_DIR, filename), buffer);
    }).done();
});