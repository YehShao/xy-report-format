const Cryptr = require('cryptr');
const cryptr = new Cryptr(process.env.SECRET_KEY);
const { AUTHOR_INFO }= require('./Constants');

exports.getDataStructure = function() {
    return {
        title: '',
        name: '',
        organization: '',
        date: '',
        author: '',
        contents:[]
    };
};

exports.getAuthorInfo = function(name) {
    let info = '';
    AUTHOR_INFO.forEach((object) => {
        if (cryptr.decrypt(object.name) === name) {
            info = cryptr.decrypt(object.value);
        }
    })
    console.log('info = ' + info);
    return info;
};