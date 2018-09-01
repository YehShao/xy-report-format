const fs = require('fs');
const jszip = require('jszip');
const docxtemplater = require('docxtemplater');

function getDataStructure() {
    return {
        fileName: '檔名',
        title: '標題',
        name: '姓名',
        organization: '台北士林分會',
        date: '日期',
        author: '作者',
        contents: []
    };
};

exports.getDefaltDataWithName = (name) => {
    var newData = getDataStructure();
    newData.name = name;
    return newData;
}

function regex(value) {
    return value.replace(/\s/g, '');
};

exports.getReformingData = (rawText, file) => {
    let rawTextArray = rawText.split('\n').filter((value) => {
        return (value !== '');
    });
    console.log(rawTextArray);
    let data = getDataStructure();
    data.fileName = file;
    rawTextArray.forEach((value, index) => {
        value = regex(value);
        switch(index) {
            case 0:
                data.title = value;
                break;
            case 1:
                data.name = value;
                break;
            case 2:
                if (value.length !== 6 || !value.match('士林')) {
                    data.organization = '台北士林分會';
                } else {
                    data.organization = value;
                }
                break;
            case 3:
                data.date = value;
                break;
            default:
                if (value.match('作者簡介') && value.match(data.name)) {
                    data.author = value.slice(value.indexOf(data.name));
                } else {
                    data.contents.push({content: value});
                }
                break;
        }
    });
    console.log(data);

    return data;
};

exports.getDocBuffer = (templateContent, data) => {
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

    return buffer;
};