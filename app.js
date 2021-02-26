const express = require('express')
const app = express()
const port = process.env.PORT || 3000




const http = require('http'); // or 'https' for https:// URLs
const fs = require('fs');

const path = require('path');

// const PizZip = require('pizzip');
// const Docxtemplater = require('docxtemplater');


// const spauth = require('node-sp-auth');
// const request = require('request-promise');

// const url = 'https://hn594a44314c984.sharepoint.com/';  
// const clientId = "fa78755c-5c36-437e-83ec-cffb77c4fe84";  
// const clientSecret = "dwcheLsZzbfLcvekh4+lfC5Ef3wjeFvxSXJ3amldfe4=";


// var docURI = "/Gedeelde documenten/AOK-001 - 001 Kemal San - 2021-02-26T22_21_57.8233425Z.docx";
// var docName = path.basename(docURI)
// var siteSP = "/sites/testwaarschuwingsbeleidtest2"

// spauth
//   .getAuth('https://hn594a44314c984.sharepoint.com/sites/testwaarschuwingsbeleidtest2/', {
//     username: 'kemal@deggroep.nl',
//     password: '41Mijnkoplokop43106',

//   })
//   .then(data => {
//     var headers = data.headers;
//     headers['Accept'] = 'application/json;odata=verbose';
//     headers['Content-type'] = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
//     request.get({
//       url: "https://hn594a44314c984.sharepoint.com/sites/testwaarschuwingsbeleidtest2/_api/web/GetFileByServerRelativeUrl('" + siteSP + docURI + "')/$value",
//       headers: headers,
//       encoding: "binary",
//       json: true
//     }).then(response => {
//         const buffer = Buffer.from(response, 'binary');
//         fs.writeFileSync('./tmp/' + docName, buffer);
//     });
//   });







// // The error object contains additional information when logged with JSON.stringify (it contains a properties object containing all suberrors).
// function replaceErrors(key, value) {
//     if (value instanceof Error) {
//         return Object.getOwnPropertyNames(value).reduce(function(error, key) {
//             error[key] = value[key];
//             return error;
//         }, {});
//     }
//     return value;
// }

// function errorHandler(error) {
//     console.log(JSON.stringify({error: error}, replaceErrors));

//     if (error.properties && error.properties.errors instanceof Array) {
//         const errorMessages = error.properties.errors.map(function (error) {
//             return error.properties.explanation;
//         }).join("\n");
//         console.log('errorMessages', errorMessages);
//         // errorMessages is a humanly readable message looking like this :
//         // 'The tag beginning with "foobar" is unopened'
//     }
//     throw error;
// }

// //Load the docx file as a binary
// var content = fs
//     .readFileSync(path.resolve(__dirname, 'input.docx'), 'binary');

// var zip = new PizZip(content);
// var doc;
// try {
//     doc = new Docxtemplater(zip);
// } catch(error) {
//     // Catch compilation errors (errors caused by the compilation of the template : misplaced tags)
//     errorHandler(error);
// }

// //set the templateVariables
// doc.setData({
//     first_name: 'Kemal',
//     last_name: 'Doe',
//     phone: '0652455478',
//     description: 'New Website'
// });

// try {
//     // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
//     doc.render()
// }
// catch (error) {
//     // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
//     errorHandler(error);
// }

// var buf = doc.getZip()
//              .generate({type: 'nodebuffer'});

// // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
// fs.writeFileSync(path.resolve(__dirname, 'output.docx'), buf);






app.use(express.json({limit: '50mb'}));
app.use(express.urlencoded({
  extended: true, limit: '50mb'}));


app.get('/', (req, res) => {
  res.send('Hello World!')
})


app.post('/', (req, res) => {






    res.send(req.body)

})

app.listen(port, () => {
  console.log(`Example app listening at http://localhost:${port}`)
})




