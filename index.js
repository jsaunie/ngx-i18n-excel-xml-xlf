const XLSX = require('xlsx'),
    fs = require('fs'),
    xml2js = require('xml2js');

console.log('Written with ❤️ by Jean Saunie!');

const wb = XLSX.readFile('./trad-excel.xlsx'); // Choose the workbook to work with
const ws = wb.Sheets['trad-be-u']; // Choose which sheet to work with
const data = XLSX.utils.sheet_to_json(ws); // Extract the data of sheet

// Define the structure of translation based on the excel sheet
class Trad {
    constructor(row) {
        this.fr = row.fr;
        this.en = row.en;
    }
}

// Map the excel data in Map object to allow retrieve data easily with the id of translation
const Trans = new Map();
data.map(record => {
    const trad = new Trad(record);
    Trans.set(record.id, trad);
});

// Initiate the parser with useful options
const parser = new xml2js.Parser({
    explicitArray: false,
    explicitRoot: false,
    trim: true,
});

const fileName = '/messages.xml',
    locales = ['fr', 'en'], // Use to create multiple translation
    defaultLocale = 'fr'; // Use to set the source-language

// Create translate file for each language
locales.forEach((lang) => {

    // Read XML File
    fs.readFile(__dirname + fileName, (err, data) => {
        if (err) throw err;

        parser.parseString(data, (err, result) => {
            if (err) throw err;

            // Set file attributes
            result.file['$']['source-language'] = defaultLocale;
            result.file['$']['target-language'] = lang;

            result.file.body['trans-unit'] = result.file.body['trans-unit'].map((trans) => {
                if (!!trans['$']) {
                    const id = trans['$'].id;

                    if (Trans.has(id)) {
                        const trad = Trans.get(id);

                        // Check if note is undefined, an array or an object before update the description content
                        if (!!trans.note) {
                            if (!!trans.note['$']) {
                                if (trans.note['$'].from === 'description') trans.note._ = trad[lang];
                            } else {
                                trans.note = trans.note.map(note => {
                                    if (!!note['$']) if (note['$'].from === 'description') note._ = trad[lang];
                                    return note;
                                });
                            }
                        }

                        // Check if note is undefined before update the translate content
                        if (!!trans.target) {
                            trans.target._ = trad[lang];

                            // Attr "state" is used by ngx-i18nsupport package
                            // https://github.com/martinroob/ngx-i18nsupport
                            if (!!trans.target['$']) trans.target['$'].state = 'final';
                        }
                    }
                }
                return trans;
            });

            // Initiate the Builder with useful options
            const builder = new xml2js.Builder({
                rootName: 'xliff',
                attrKey: 'attr',
                charKey: 'content',
            });

            // Build the XML with JSON Object
            const xml = builder.buildObject(result);

            // Create File in project folder
            fs.writeFile(__dirname + '/messages.' + lang + '.xlf', xml, () => {
                console.log('File created succesfully : messages.' + lang + '.xlf');
            });
        });

    });
});
