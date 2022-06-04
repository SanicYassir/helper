const XLSX = require('xlsx');
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");
const path = require("path");

const workbook = XLSX.readFile('./srv.xlsx');

let worksheets = {};

const table=[];

for(const sheetname of workbook.SheetNames){
    worksheets[sheetname]=XLSX.utils.sheet_to_json(workbook.Sheets[sheetname],{
        raw: false,
    });
}

const sheet  = worksheets["Planning Surv S2.CPI"];


/*Editable*/
const Semestre = 'SEMESTRE 4';
const Year = 'Année Universitaire 2021-2022';
const controlType = 'CONTRÔLE TERMINAL'




let Dates,Locaux,Heure,Module,Coordinateur,Responsables,Surveillants;

for(var i =0 ; i<sheet.length;i++){
  if(i===0){
      Dates =sheet[0]["Dates"];
      Locaux = sheet[0]["Locaux"];
      Heure=sheet[0]["Heure"];
      Module=sheet[0]["Module"];
      Coordinateur=sheet[0]["Coordianteur"];
      Responsables=sheet[0]["Responsables"];
      Surveillants=sheet[0]["Surveillants"];
  }

    if(sheet[i]["Dates"] !== Dates && sheet[i]["Dates"]!==undefined){
        Dates = sheet[i]["Dates"];
    }

    if(sheet[i]["Locaux"] !== Locaux && sheet[i]["Locaux"]!==undefined){
        Locaux = sheet[i]["Locaux"];
    }

    if(sheet[i]["Heure"] !== Heure && sheet[i]["Heure"]!==undefined){
        Heure = sheet[i]["Heure"];
    }

    if(sheet[i]["Module"] !== Module && sheet[i]["Module"]!==undefined){
        Module = sheet[i]["Module"];
    }

    if(sheet[i]["Coordinateur"] !== Coordinateur && sheet[i]["Coordinateur"]!==undefined){
        Coordinateur = sheet[i]["Coordinateur"];
    }

    if(sheet[i]["Responsables"] !== Responsables && sheet[i]["Responsables"]!==undefined){
        Responsables = sheet[i]["Responsables"];
    }

    if(sheet[i]["Surveillants"] !== Surveillants && sheet[i]["Surveillants"]!==undefined){
        Surveillants = sheet[i]["Surveillants"];
    }

    table.push({
        Semestre,
        Year,
        controlType,
        Dates,
        Locaux,
        Heure:Heure.split('-')[0],
        Module,
        Coordinateur,
        Responsables,
        Surveillants
    })
}


                                /* Generate PVS */
const dir = `./PVS/${Semestre}`;

if (!fs.existsSync(dir)){
    fs.mkdirSync(dir);
}

// Load the docx file as binary content


// Creating docs 
table.forEach(async (item,index)=>{

    const content = fs.readFileSync(
        path.resolve(__dirname, "PV template.docx"),
        "binary"
    );
    
    
    const zip = new PizZip(content);
    
    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
    });

    doc.render(item);
    
    const buf = await doc.getZip().generate({
        type: "nodebuffer",
        compression: "DEFLATE",
    });
    fs.writeFileSync(path.resolve(dir, `${index}.docx`), buf);
})



