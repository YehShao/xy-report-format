const jszip = require('jszip');
const docxtemplater = require('docxtemplater');
const Cryptr = require('cryptr');
const cryptr = new Cryptr(process.env.SECRET_KEY);
const { AUTHOR_INFO }= require('./Constants');

function getDataStructure() {
    return {
        title: '',
        name: '',
        organization: '',
        date: '',
        author: '',
        contents: []
    };
};

function getAuthorInfo(name) {
    let info = '';
    AUTHOR_INFO.forEach((object) => {
        if (cryptr.decrypt(object.name) === name) {
            info = cryptr.decrypt(object.value);
        }
    })
    console.log('info = ' + info);
    return info;
};

exports.getDataReforming = (rawText) => {
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

    return data;
}

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
}