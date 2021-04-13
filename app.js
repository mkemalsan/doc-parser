const express = require('express')
const app = express()
const port = process.env.PORT || 3000

const fs = require('fs')
const path = require('path')

const PizZip = require('pizzip')
const Docxtemplater = require('docxtemplater')

const expressions = require('angular-expressions')
const assign = require("lodash/assign")

const { exec } = require("child_process");


// Init middleware
app.use(express.json({limit: '50mb'}));
app.use(express.urlencoded({extended: true, limit: '50mb'}));


// Hi!
app.get('/', (req, res) => {
  res.send('Hello World!')
})


// POST Request to fill data in placeholders of template
// REQUEST JSON schema:
// {
//      "document": "String",       Value of string is expected to be a base64 encoded Word document
//      "data": "String"            Value of string is expected to be a base64 encoded JSON Object
// }
app.post('/', (req, res) => {

    var document = Buffer.from(req.body.document, 'base64')
    var data = Buffer.from(req.body.data, 'base64')
    var templateData = {}

    
    data = data.toString('utf-8')
    data = JSON.parse(data)

    data.forEach(formDataObject => {
        templateData = { ...templateData, ...formDataObject }
    })

    var zip = new PizZip(document)
    var doc
    
    try {
        doc = new Docxtemplater(zip, {nullGetter() { return ''; }, parser:angularParser})
    } catch(error) {
        // Catch compilation errors (errors caused by the compilation of the template : misplaced tags)
        errorHandler(error)
    }

    //set the templateVariables
    doc.setData(templateData)

    try {
        // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
        doc.render()
    }
    catch (error) {
        // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
        errorHandler(error)
    }

    var buf = doc.getZip().generate({type: 'base64'})

    // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
    // buf = Buffer.from(document, 'base64')
    // fs.writeFileSync(path.resolve(__dirname, 'tmp/' + '456.docx'), buf)
	
    var JSONresponse = {}

    JSONresponse.data 		= req.body.data
    JSONresponse.document 	= buf
    // JSONresponse.pdf		= pdf

    res.send(JSONresponse)

})



app.get('/test/', (req, res) => {
    res.send('test')
})

// POST Request to fill data in placeholders of template
// REQUEST JSON schema:
// {
//      "document": "String",       Value of string is expected to be a base64 encoded Word document
//      "data": "String"            Value of string is expected to be a base64 encoded JSON Object
// }
app.post('/test/', (req, res) => {
    console.log(req.headers)

    var document = Buffer.from(req.body.document, 'base64')

    var data = Buffer.from(req.body.data, 'base64')
    var templateData = {}
    data = data.toString('utf-8')
    data = JSON.parse(data)

    data.forEach(formDataObject => {
        templateData = { ...templateData, ...formDataObject }
    })
    var zip = new PizZip(document)
    var doc
    
    try {
        doc = new Docxtemplater(zip, {nullGetter() { return ''; }, parser:angularParser})
    } catch(error) {
        // Catch compilation errors (errors caused by the compilation of the template : misplaced tags)
        errorHandler(error)
    }

    //set the templateVariables
    doc.setData(templateData)

    try {
        // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
        doc.render()
    }
    catch (error) {
        // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
        errorHandler(error)
    }

    var buf = doc.getZip().generate({type: 'nodebuffer'})
    var bufBase64 = doc.getZip().generate({type: 'base64'})

    // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
    var dest = path.resolve(__dirname, 'tmp/' + 'tmp.docx')
    fs.writeFileSync(dest, buf)


    var root = "/var/www/vhosts/deggroep.nl/api.deggroep.nl",
    soffice,
    tmp


    switch (process.platform) {
        case 'darwin':
            soffice = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
            tmp     = "/Users/mkemalsan/Documents/GitHub/doc-parser/tmp/"
            break;
        case 'linux' :
            soffice = root + "/bin/squashfs-root/opt/libreoffice7.1/program/soffice"
            tmp     = root + "/tmp/"
            break;
    }



    
    
    var JSONresponse = {}

    JSONresponse.data       = req.body.data
    JSONresponse.document   = bufBase64
    // JSONresponse.pdf        = "pdf"



    fs.writeFileSync("Log" + Date.now(), "before")

    exec(`${soffice} --convert-to pdf --outdir ${tmp} ${tmp}tmp.docx`, (error, stdout, stderr) => {
        if (error) {
            console.log(`error: ${error.message}`);
            return;
        }
        if (stderr) {
            console.log(`stderr: ${stderr}`);
            return;
        }
        console.log(`stdout: ${stdout}`);

        setTimeout(function(){console.log('waited 3000ms')}, 3000);
        fs.writeFileSync("Log" + Date.now(), "during")
        JSONresponse.pdf        = fs.readFileSync(tmp + "tmp.pdf", {encoding: 'base64'})

    });
    
    fs.writeFileSync("Log" + Date.now(), "after")

    res.send(JSONresponse)

})

app.listen(port, () => {
  console.log(`Listening at port:${port}`)
})

// ANGULAR STYLE EXPRESSIONS AND CUSTOM FILTERS
expressions.filters.valuta = function(input) {
    if(!input) return input;
    return parseFloat(input).toLocaleString("en-EN", {style: "currency", currency: "EUR", minimumFractionDigits: 2}).replace(',',',/').replace('.','./').replace(',/','.').replace('./',',').replace(',00', ',-')
}
expressions.filters.datum = function(input) {
    if(!input) return input;
    return input.split("-").reverse().join("-");
}

expressions.filters.keuze = function(input) {
    if(!input) return input;
    return input.Value
}

function angularParser(tag) {
    if (tag === '.') {
        return {
            get: function(s){ return s;}
        };
    }
    const expr = expressions.compile(
        tag.replace(/(’|‘)/g, "'").replace(/(“|”)/g, '"')
    );
    return {
        get: function(scope, context) {
            let obj = {};
            const scopeList = context.scopeList;
            const num = context.num;
            for (let i = 0, len = num + 1; i < len; i++) {
                obj = assign(obj, scopeList[i]);
            }
            return expr(scope, obj);
        }
    };
}