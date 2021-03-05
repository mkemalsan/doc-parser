const express = require('express')
const app = express()
const port = process.env.PORT || 3000

const https = require('https')
const fs = require('fs')
const path = require('path')

const PizZip = require('pizzip')
const Docxtemplater = require('docxtemplater')

const spauth = require('node-sp-auth')
const request = require('request-promise')

const url = "https://hn594a44314c984.sharepoint.com/"
const clientId = "fa78755c-5c36-437e-83ec-cffb77c4fe84"
const clientSecret = "dwcheLsZzbfLcvekh4+lfC5Ef3wjeFvxSXJ3amldfe4="

var token = ""

const sharepointSite = "/sites/testwaarschuwingsbeleidtest2"

var file = ""

// Please make sure to use angular-expressions 1.1.2 or later
// More detail at https://github.com/open-xml-templating/docxtemplater/issues/589
var expressions = require('angular-expressions');
var assign = require("lodash/assign");
// define your filter functions here, for example, to be able to write {clientname | lower}
expressions.filters.lower = function(input) {
    // This condition should be used to make sure that if your input is
    // undefined, your output will be undefined as well and will not
    // throw an error
    if(!input) return input;
    return input.toLowerCase();
}

expressions.filters.valuta = function(input) {
    // This condition should be used to make sure that if your input is
    // undefined, your output will be undefined as well and will not
    // throw an error
    if(!input) return input;
    return parseInt(input).toLocaleString("nl-NL", {style: "currency", currency: "EUR", minimumFractionDigits: 2});
}
expressions.filters.datum = function(input) {
    // This condition should be used to make sure that if your input is
    // undefined, your output will be undefined as well and will not
    // throw an error
    if(!input) return input;
    return input.split("-").reverse().join("-");
}



// var date = "03-11-2014";
// var newdate = date.split("-").reverse().join("-");

// let bla = "2500"
// console.log(parseInt(bla).toLocaleString("nl-NL", {style: "currency", currency: "EUR", minimumFractionDigits: 2}))

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
// new Docxtemplater(zip, {parser:angularParser});

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
//      "document": "String",       String is expected to be a URI to the template
//      "data": "String"            String is expected to be a base64 encoded JSON Object
// }
app.post('/', (req, res) => {

    var document = req.body.document;
    var data = req.body.data;


    data = Buffer.from(data, 'base64');
    data = data.toString('utf-8');
    data = JSON.parse(data)


    spauth
    .getAuth('https://hn594a44314c984.sharepoint.com/sites/testwaarschuwingsbeleidtest2/', {
    clientId: clientId,
    clientSecret: clientSecret
    })
    .then(data => {
        // console.log(data.headers['Authorization'])
        token = data.headers['Authorization']
        var headers = data.headers;
        headers['Accept'] = 'application/json;odata=verbose';


    });
    getSPDocument(document, data)
  
    


    res.send(req.body)

})

app.listen(port, () => {
  console.log(`Listening at port:${port}`)
})

















function getSPDocument(documentURI, formData){
    // var sharepointSite = sharepointSite
    var documentURI = documentURI
    var docName = path.basename(documentURI)
    var dirName = path.dirname(documentURI)

    var formData = formData
    // console.log(formData)
    var newData = {}


    // Prepare formData data
    // Expected input:
    // [{data},{data}]
    formData.forEach(formDataObject => {
        newData = { ...newData, ...formDataObject }
    })
    // console.log(newData)





    spauth
    .getAuth('https://hn594a44314c984.sharepoint.com/sites/testwaarschuwingsbeleidtest2/', {
    username: 'kemal@deggroep.nl',
    password: '41Mijnkoplokop43106',

    })
    .then(data => {
        var headers = data.headers;
        headers['Accept'] = 'application/json;odata=verbose';
        headers['Content-type'] = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        request
        .get({
          url: "https://hn594a44314c984.sharepoint.com/sites/testwaarschuwingsbeleidtest2/_api/web/GetFileByServerRelativeUrl('" + sharepointSite + documentURI + "')/$value",
          headers: headers,
          encoding: "binary",
          json: true
        })
        .then(response => {
            const buffer = Buffer.from(response, 'binary');
            fs.writeFileSync('./tmp/' + docName, buffer);
     


            // The error object contains additional information when logged with JSON.stringify (it contains a properties object containing all suberrors).
            function replaceErrors(key, value) {
                if (value instanceof Error) {
                    return Object.getOwnPropertyNames(value).reduce(function(error, key) {
                        error[key] = value[key];
                        return error;
                    }, {});
                }
                return value;
            }

            function errorHandler(error) {
                console.log(JSON.stringify({error: error}, replaceErrors));

                if (error.properties && error.properties.errors instanceof Array) {
                    const errorMessages = error.properties.errors.map(function (error) {
                        return error.properties.explanation;
                    }).join("\n");
                    // console.log('errorMessages', errorMessages);
                    // errorMessages is a humanly readable message looking like this :
                    // 'The tag beginning with "foobar" is unopened'
                }
                throw error;
            }

            //Load the docx file as a binary
            var content = fs
                .readFileSync(path.resolve(__dirname, 'tmp/' + docName), 'binary');

            var zip = new PizZip(content);
            var doc;
            try {
                // doc = new Docxtemplater(zip, {parser:angularParser});
                // doc = new Docxtemplater(zip, {nullGetter() { return ''; }});
                doc = new Docxtemplater(zip, {nullGetter() { return ''; }, parser:angularParser});
            } catch(error) {
                // Catch compilation errors (errors caused by the compilation of the template : misplaced tags)
                errorHandler(error);
            }

            //set the templateVariables
            doc.setData(newData);

            try {
                // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
                doc.render()
            }
            catch (error) {
                // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
                errorHandler(error);
            }

            var buf = doc.getZip()
                         .generate({type: 'nodebuffer'});

            file = buf

            // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
            fs.writeFileSync(path.resolve(__dirname, 'tmp/' + docName), buf);


        })

        .finally(()=> {
            postSPDocument(documentURI)
        })
    })

}




function postSPDocument(documentURI){
    var documentURI = documentURI
    var docName = path.basename(documentURI)
    var dirName = path.dirname(documentURI)


    // FORM DIGEST VALUE
    spauth
    .getAuth('https://hn594a44314c984.sharepoint.com/sites/testwaarschuwingsbeleidtest2/', {
    username: 'kemal@deggroep.nl',
    password: '41Mijnkoplokop43106',

    })
    .then(data => {
        var headers = data.headers;

        headers['Authorization'] = token
        headers['Accept'] = 'application/json;odata=verbose';
        headers['Content-type'] = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        
        // console.log(headers)
        request
        .post({
            url: "https://hn594a44314c984.sharepoint.com/sites/testwaarschuwingsbeleidtest2/_api/contextinfo",
            headers: headers
        })
        .then(response => {
            response = JSON.parse(response)

            var postData = fs.readFileSync(path.resolve(__dirname, 'tmp/' + docName), 'binary');

            headers = {
                'Authorization': headers['Authorization'],
                'X-RequestDigest' : response.d.GetContextWebInformation.FormDigestValue,
                'Accept': "application/json;odata=verbose",
                'Content-Type': "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                'Content-length': Buffer.byteLength(file, 'utf8')
            }

            var uri =  encodeURI(sharepointSite + "/_api/web/GetFolderByServerRelativeUrl('" + sharepointSite + dirName + "')/Files/add(url='OUTPUT" + docName + "',overwrite=true)")

            var options = {
              'method': 'POST',
              'hostname': 'hn594a44314c984.sharepoint.com',
              'path': uri,
              'headers': headers,
              'maxRedirects': 20
            };

            var req = https.request(options, function (res) {
              var chunks = [];

              res.on("data", function (chunk) {
                chunks.push(chunk);
              });

              res.on("end", function (chunk) {
                var body = Buffer.concat(chunks);
                // console.log(body.toString());
              });

              res.on("error", function (error) {
                console.error(error);
              });
            });
            // console.log(postData);

            req.write(file);
            req.end();
        })
    })

}









