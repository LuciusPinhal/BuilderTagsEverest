let toBeReplaced;
let selectedSheet;
let sourceWorkbook;
let destinationWorkbook;

let prefixAtribute03;
let variableAtribute03;
let dustPrefixAtribute03;
let variableAtributes03;
let InputAtribute03Value= [];
let InputAtribute03 = [];

let AtributeInput;
let AtributeInput2;
let AtributeInput3;

let InputColumnsAtributteOne01 = []; 
let InputColumnsAtributteOne02 = []; 
let InputColumnsAtributteOne03 = []; 
let InputColumnsAtributteOne04 = []; 
let InputColumnsAtributteOne05 = []; 

let InputColumnsAtributteTwo01 = []; 
let InputColumnsAtributteTwo02 = []; 
let InputColumnsAtributteTwo03 = []; 
let InputColumnsAtributteTwo04 = []; 
let InputColumnsAtributteTwo05 = []; 

let InputColumnsAtributteThree01 = []; 
let InputColumnsAtributteThree02 = []; 
let InputColumnsAtributteThree03 = []; 
let InputColumnsAtributteThree04 = []; 
let InputColumnsAtributteThree05 = []; 

//Propriedade dos Atributos 01

let propertiesAtributteOne01 = [];
let propertiesAtributteOne01prefix= [];
let propertiesAtributteOne01dustPrefix= [];
let propertiesAtributteOne01variable= [];

let propertiesAtributteOne02 = [];
let propertiesAtributteOne02prefix= [];
let propertiesAtributteOne02dustPrefix= [];
let propertiesAtributteOne02variable= [];

let propertiesAtributteOne03 =[];
let propertiesAtributteOne03prefix= [];
let propertiesAtributteOne03dustPrefix= [];
let propertiesAtributteOne03variable= [];

let propertiesAtributteOne04 = [];
let propertiesAtributteOne04prefix= [];
let propertiesAtributteOne04dustPrefix= [];
let propertiesAtributteOne04variable= [];

let propertiesAtributteOne05 = [];
let propertiesAtributteOne05prefix= [];
let propertiesAtributteOne05dustPrefix= [];
let propertiesAtributteOne05variable= [];

//Propriedade dos Atributos 02
let propertiesAtributteTwo01= [];
let propertiesAtributteTwo01prefix= [];
let propertiesAtributteTwo01dustPrefix= [];
let propertiesAtributteTwo01variable= [];

let propertiesAtributteTwo02 =[];
let propertiesAtributteTwo02prefix= [];
let propertiesAtributteTwo02dustPrefix= [];
let propertiesAtributteTwo02variable= [];

let propertiesAtributteTwo03=[];
let propertiesAtributteTwo03prefix= [];
let propertiesAtributteTwo03dustPrefix= [];
let propertiesAtributteTwo03variable= [];

let propertiesAtributteTwo04 = [];
let propertiesAtributteTwo04prefix= [];
let propertiesAtributteTwo04dustPrefix= [];
let propertiesAtributteTwo04variable= [];

let propertiesAtributteTwo05 = [];
let propertiesAtributteTwo05prefix= [];
let propertiesAtributteTwo05dustPrefix= [];
let propertiesAtributteTwo05variable= [];

//Propriedade dos Atributos 03

let propertiesAtributteThree01 =[];
let propertiesAtributteThree01prefix= [];
let propertiesAtributteThree01dustPrefix= [];
let propertiesAtributteThree01variable= [];

let propertiesAtributteThree02 = [];
let propertiesAtributteThree02prefix= [];
let propertiesAtributteThree02dustPrefix= [];
let propertiesAtributteThree02variable= [];

let propertiesAtributteThree03 = [];
let propertiesAtributteThree03prefix= [];
let propertiesAtributteThree03dustPrefix= [];
let propertiesAtributteThree03variable= [];

let propertiesAtributteThree04 =[];
let propertiesAtributteThree04prefix= [];
let propertiesAtributteThree04dustPrefix= [];
let propertiesAtributteThree04variable= [];

let propertiesAtributteThree05 = [];
let propertiesAtributteThree05prefix= [];
let propertiesAtributteThree05dustPrefix= [];
let propertiesAtributteThree05variable= [];

let Colums = [];
let ColumnsAtributteOne01 = [];
let ColumnsAtributteOne02 = [];
let ColumnsAtributteOne03 = [];
let ColumnsAtributteOne04 = [];
let ColumnsAtributteOne05 = [];

let Gadgets = [];
let ColumnsAtributteTwo01 = [];
let ColumnsAtributteTwo02 = [];
let ColumnsAtributteTwo03 = [];
let ColumnsAtributteTwo04 = [];
let ColumnsAtributteTwo05 = [];

let Atribute03 = [];
let ColumnsAtributteThree01 = [];
let ColumnsAtributteThree02 = [];
let ColumnsAtributteThree03 = [];
let ColumnsAtributteThree04 = [];
let ColumnsAtributteThree05 = [];

let InputColumns = [];
let InputGadgets = [];
let getResultInput= [];
let InputGadgetsValue = [];
let InputColumnsValue = [];
let destinationData = [];
let selectedSheetData = []; 
let checked = false;
let control = [];

async function handleSourceFile(e) {
    const file = e.target.files[0];
    const reader = new FileReader();
  
    reader.onload = function (e) {
      const data = new Uint8Array(e.target.result);
      sourceWorkbook = XLSX.read(data, { type: 'array' });

      console.warn("🍷🗿 >> sourceWorkbook:", sourceWorkbook);
  
      const sheetCheckboxes = document.getElementById('sheetCheckboxes');
      sheetCheckboxes.innerHTML = '';
  
      const sheetNames = sourceWorkbook.SheetNames;
      const selectedSheets = [];
        
      let count = 0;
      sheetNames.forEach(function (sheetName) {
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.name = 'selectedSheets';
        checkbox.value = sheetName;
        checkbox.id = 'checkbox_' + sheetName + count;

        const labelText = document.createElement('p');
        labelText.innerText = sheetName;
  
        checkbox.addEventListener('change', function () {
          if (this.checked) {
            selectedSheets.push(sheetName);
          } else {
            const index = selectedSheets.indexOf(sheetName);
            if (index > -1) {
              selectedSheets.splice(index, 1);
            }
          }
        });
        
        const label = document.createElement('label');
        label.htmlFor = 'checkbox_' + sheetName + count;
        label.classList.add('checkboxContainer');
        
        label.appendChild(checkbox);
        label.appendChild(labelText);
        sheetCheckboxes.appendChild(label);
        sheetCheckboxes.appendChild(document.createElement('br'));

        count++
      });
        let control = [];
        const saveButton = document.getElementById('saveButton');
        saveButton.addEventListener('click', function () {
        selectedSheetData = []; // Array para armazenar os valores das células das planilhas selecionadas
        
        selectedSheets.forEach(function (sheetName) {
            const sheet = sourceWorkbook.Sheets[sheetName];
            const sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            selectedSheetData.push({ sheetName, data: sheetData });
        });
        checked = true;
        checkSheetSelection();
        updateTable();
        // Agora você tem um array selectedSheetData que contém os valores das células das planilhas selecionadas
        console.log(selectedSheetData);
        });

    };
    const Button = document.getElementById('sheetCheckboxes');
    Button.addEventListener('click', function () {
        if(JSON.stringify(control) != JSON.stringify(selectedSheetData)){
            control = selectedSheetData;
            checked = true;
        }else{
            checked = false;
        }
        checkSheetSelection();
    });
    
    reader.readAsArrayBuffer(file);

    const Attribute = document.getElementById('buttonSave');
    const Attribute1 = document.getElementById('saveButton');
    
    Attribute.style.display = 'flex';
    Attribute1.style.display = 'block'; 
}
  
window.addEventListener('DOMContentLoaded', function () {
    document.getElementById('sourceFileInput').addEventListener('change', handleSourceFile);
});
  
function handleDestinationFile(e) {
    const file = e.target.files[0];
    const reader = new FileReader();
  
    reader.onload = function (e) {
      const data = new Uint8Array(e.target.result);
      destinationWorkbook = XLSX.read(data, { type: 'array' });
      checkSheetSelection();
  
      // Chamar a função table() após o carregamento completo do arquivo de destino
      reader.onloadend = function () {
        const sourceSheet = sourceWorkbook.Sheets[selectedSheet];
        table(sourceSheet);
      };
    };
    reader.readAsArrayBuffer(file);
}
  
function table() {
    InputColumns = []
    // Extrair os valores das células A4 a E9
    //o 3 seria A4 e o 8 seria o E9
    const cellValues = [];

    for (let row = 3; row <= 24; row++) {
        const rowValues = [];
        for (let col = 0; col < 5; col++) {
            const cellValue = selectedSheetData[0].data[row][col] || "";
            rowValues.push(cellValue);
        }
        cellValues.push(rowValues);
    }

    // Concatenar os valores das keys da tabela
    if(cellValues.length >=0){
        //console.warn("🔭🐢 >> cellValues:", cellValues);
        //Atributo 01
        prefixColumns = cellValues[1][2];
        variableColumns = cellValues[1][3];
        dustPrefixColumns = cellValues[1][4];

        //propriedades do atributo 1 
        propertiesAtributteOne01prefix = cellValues[5][2];
        propertiesAtributteOne01variable = cellValues[5][3];
        propertiesAtributteOne01dustPrefix = cellValues[5][4];

        propertiesAtributteOne01 = propertiesAtributteOne01prefix + ' '+ propertiesAtributteOne01variable+propertiesAtributteOne01dustPrefix;

        propertiesAtributteOne02prefix = cellValues[6][2];
        propertiesAtributteOne02dustPrefix = cellValues[6][4];
        propertiesAtributteOne02variable = cellValues[6][3];

        propertiesAtributteOne02 = propertiesAtributteOne02prefix + ' '+ propertiesAtributteOne02variable+ propertiesAtributteOne02dustPrefix;
        
        propertiesAtributteOne03prefix = cellValues[7][2];
        propertiesAtributteOne03dustPrefix = cellValues[7][4];
        propertiesAtributteOne03variable = cellValues[7][3];

        propertiesAtributteOne03 = propertiesAtributteOne03prefix+ ' '+propertiesAtributteOne03variable+ propertiesAtributteOne03dustPrefix;

        propertiesAtributteOne04prefix = cellValues[8][2];
        propertiesAtributteOne04dustPrefix = cellValues[8][4];
        propertiesAtributteOne04variable = cellValues[8][3];

        propertiesAtributteOne04 = propertiesAtributteOne04prefix+ ' '+propertiesAtributteOne04variable+propertiesAtributteOne04dustPrefix;

        propertiesAtributteOne05prefix = cellValues[9][2];
        propertiesAtributteOne05dustPrefix = cellValues[9][4];
        propertiesAtributteOne05variable = cellValues[9][3];

        propertiesAtributteOne05 =  propertiesAtributteOne05prefix+ ' '+  propertiesAtributteOne05variable+propertiesAtributteOne05dustPrefix;

        console.warn("🍷🗿 >> propertiesAtributteOne05:", propertiesAtributteOne05);
        
        //Atributo 02
        prefixGadgets = cellValues[2][2];
        variableGadgets = cellValues[2][3];
        dustPrefixGadgets = cellValues[2][4];

        //propriedades do atributo 2
        propertiesAtributteTwo01prefix = cellValues[11][2];
        propertiesAtributteTwo01variable = cellValues[11][3];
        propertiesAtributteTwo01dustPrefix = cellValues[11][4];

        propertiesAtributteTwo01 = propertiesAtributteTwo01prefix + ' '+ propertiesAtributteTwo01variable +  propertiesAtributteTwo01dustPrefix;

        propertiesAtributteTwo02prefix = cellValues[12][2];
        propertiesAtributteTwo02dustPrefix = cellValues[12][4];
        propertiesAtributteTwo02variable = cellValues[12][3];

        propertiesAtributteTwo02 =propertiesAtributteTwo02prefix + ' '+ propertiesAtributteTwo02variable + propertiesAtributteTwo02dustPrefix;

        propertiesAtributteTwo03prefix = cellValues[13][2];
        propertiesAtributteTwo03dustPrefix = cellValues[13][4];
        propertiesAtributteTwo03variable = cellValues[13][3];

        propertiesAtributteTwo03 = propertiesAtributteTwo03prefix+ ' '+ propertiesAtributteTwo03variable + propertiesAtributteTwo03dustPrefix;

        propertiesAtributteTwo04prefix = cellValues[14][2];
        propertiesAtributteTwo04dustPrefix = cellValues[14][4];
        propertiesAtributteTwo04variable = cellValues[14][3];

        propertiesAtributteTwo04 = propertiesAtributteTwo04prefix + ' '+ propertiesAtributteTwo04variable + propertiesAtributteTwo04dustPrefix;

        propertiesAtributteTwo05prefix = cellValues[15][2];
        propertiesAtributteTwo05dustPrefix = cellValues[15][4];
        propertiesAtributteTwo05variable = cellValues[15][3];

        propertiesAtributteTwo05 = propertiesAtributteTwo05prefix + ' '+ propertiesAtributteTwo05variable+ propertiesAtributteTwo05dustPrefix;

        console.warn("🍷🗿 >> propertiesAtributteTwo05:", propertiesAtributteTwo05);

        //atributo 03
        prefixAtribute03 = cellValues[3][2];
        variableAtribute03 = cellValues[3][3];
        dustPrefixAtribute03 = cellValues[3][4];

        //propriedades do atributo 3
        propertiesAtributteThree01prefix = cellValues[17][2];
        propertiesAtributteThree01variable = cellValues[17][3];
        propertiesAtributteThree01dustPrefix = cellValues[17][4];

        propertiesAtributteThree01 = propertiesAtributteThree01prefix+ ' '+ propertiesAtributteThree01variable+ propertiesAtributteThree01dustPrefix;

        propertiesAtributteThree02prefix = cellValues[18][2];
        propertiesAtributteThree02dustPrefix = cellValues[18][4];
        propertiesAtributteThree02variable = cellValues[18][3];

        propertiesAtributteThree02 = propertiesAtributteThree02prefix + ' ' + propertiesAtributteThree02variable + propertiesAtributteThree02dustPrefix;

        propertiesAtributteThree03prefix = cellValues[19][2];
        propertiesAtributteThree03dustPrefix = cellValues[19][4];
        propertiesAtributteThree03variable = cellValues[19][3];

        propertiesAtributteThree03 = propertiesAtributteThree03prefix + ' '+ propertiesAtributteThree03variable+ propertiesAtributteThree03dustPrefix;

        propertiesAtributteThree04prefix = cellValues[20][2];
        propertiesAtributteThree04dustPrefix = cellValues[20][4];
        propertiesAtributteThree04variable = cellValues[20][3];

        propertiesAtributteThree04 = propertiesAtributteThree04prefix + ' '+ propertiesAtributteThree04variable+ propertiesAtributteThree04dustPrefix;

        propertiesAtributteThree05prefix = cellValues[21][2];
        propertiesAtributteThree05dustPrefix = cellValues[21][4];
        propertiesAtributteThree05variable = cellValues[21][3];

        propertiesAtributteThree05 = propertiesAtributteThree05prefix + ' '+ propertiesAtributteThree05variable+ propertiesAtributteThree05dustPrefix;

        console.warn("🍷🗿 >> propertiesAtributteThree05:", propertiesAtributteThree05);

    }
  
    //Pegando Atributo 1 e 2
    AtributeInput = cellValues[1][1];
    AtributeInput2 = cellValues[2][1];
    AtributeInput3 = cellValues[3][1];
    
    if(prefixGadgets !== ""){
        const Attribute = document.getElementById('inputContainerAttribute');
        Attribute.style.display = 'block';
        const Attribute2 = document.getElementById('attributeOne');
        Attribute2.textContent = AtributeInput2;
        const AttributeValue = document.getElementById('attributeOneValue');
        AttributeValue.textContent =  'Valor a ser substituido: "' + prefixGadgets+ ' ' + variableGadgets+'"';
    }else{ 
        InputGadgetsValue = [''];
        InputGadgets = [''];
        document.getElementById("replaceGadgets").value = "";
        processCommaSeparatedValues();
        const Attribute = document.getElementById('inputContainerAttribute');
        Attribute.style.display = 'none';
    }

    if(prefixColumns !== ""){
        const Attribute = document.getElementById('inputContainerAttributeTwo');
        Attribute.style.display = 'block';
        const Attribute2 = document.getElementById('attributeTwo');
        Attribute2.textContent = AtributeInput;
        const AttributeValue = document.getElementById('attributeTwoValue');
        AttributeValue.textContent =  'Valor a ser substituido: "' + prefixColumns+ ' ' + variableColumns+'"';
        
    }else{
        InputColumnsValue = [''];
        InputColumns = [''];
        document.getElementById("replaceTo").value = "";
        const Attribute = document.getElementById('inputContainerAttributeTwo');
        Attribute.style.display = 'none';
    }

    if(prefixAtribute03 !== ""){
        const Attribute = document.getElementById('inputContainerAttributeThree');
        Attribute.style.display = 'block';
        const Attribute2 = document.getElementById('attributeThree');
        Attribute2.textContent = AtributeInput3;
        const AttributeValue = document.getElementById('attributeThreeValue');
        AttributeValue.textContent =  'Valor a ser substituido: "' + prefixAtribute03 + ' '+ variableAtribute03 +'"';
        
    }else{
        InputAtribute03  = [''];
        InputAtribute03Value = [''];
        Atribute03  = [''];
        document.getElementById("replaceAtribute03").value = "";
        const Attribute = document.getElementById('inputContainerAttributeThree');
        Attribute.style.display = 'none';
    }

    // Variaveis Mocadas do Excel 
    // *** colocar + ' '+ da um espaços na variavel***

    toBeReplaced = prefixColumns + variableColumns + dustPrefixColumns;
    toBeReplacedAtributte = prefixGadgets + variableGadgets + dustPrefixGadgets;
    variableAtributes03 = prefixAtribute03 + variableAtribute03 + dustPrefixAtribute03;

}

function processCommaSeparatedValues() {
    const inputField = document.getElementById('replaceGadgets');
    const inputreplaceTo = document.getElementById('replaceTo');
    const replaceAtribute03 = document.getElementById('replaceAtribute03');
    //const replaceOption = document.getElementById('replaceOption');

    inputField.addEventListener('keyup', function(event) {
        const value = inputField.value;
        const values = value.split(',');

        InputGadgetsValue = values.map(val => val.replace(/\s/g, ''));
        concatenateInput();
    });

    inputreplaceTo.addEventListener('keyup', function(event) {
        const value = inputreplaceTo.value;
        const values = value.split(',');

        InputColumnsValue = values.map(val => val.replace(/\s/g, ''));
        concatenateInput();
    });

    replaceAtribute03.addEventListener('keyup', function(event) {
        const value = replaceAtribute03.value;
        const values = value.split(',');

        InputAtribute03Value = values.map(val => val.replace(/\s/g, ''));
        concatenateInput();
    });

    concatenateInput();
}

function concatenateInput() {
    if(InputColumnsValue.length >= 1){
        InputColumns = InputColumnsValue.map(val => prefixColumns + val + dustPrefixColumns);
        InputColumnsAtributteOne01 = InputColumnsValue.map(val =>  propertiesAtributteOne01prefix + ' '+ val +propertiesAtributteOne01dustPrefix);
        InputColumnsAtributteOne02 = InputColumnsValue.map(val =>  propertiesAtributteOne02prefix + ' '+ val +propertiesAtributteOne02dustPrefix); 
        InputColumnsAtributteOne03 = InputColumnsValue.map(val =>  propertiesAtributteOne03prefix + ' '+ val +propertiesAtributteOne03dustPrefix);
        InputColumnsAtributteOne04 = InputColumnsValue.map(val =>  propertiesAtributteOne04prefix + ' '+ val +propertiesAtributteOne04dustPrefix);
        InputColumnsAtributteOne05 = InputColumnsValue.map(val =>  propertiesAtributteOne05prefix + ' '+ val +propertiesAtributteOne05dustPrefix);
    }else{
        InputColumns = [''];
        //InputColumns = InputColumnsValue.length >= 1 ? InputColumnsValue.map(val => prefixColumns + val + dustPrefixColumns) : [''];    
        //console.warn("🔭🐢 >> InputColumns:", InputColumns);
    }
    if(InputGadgetsValue.length >= 1){
        InputGadgets = InputGadgetsValue.map(val => prefixGadgets + val + dustPrefixGadgets);
        InputColumnsAtributteTwo01 = InputGadgetsValue.map(val => propertiesAtributteTwo01prefix +' '+ val + propertiesAtributteTwo01dustPrefix);
        InputColumnsAtributteTwo02 = InputGadgetsValue.map(val => propertiesAtributteTwo02prefix + ' '+ val + propertiesAtributteTwo02dustPrefix);
        InputColumnsAtributteTwo03 = InputGadgetsValue.map(val => propertiesAtributteTwo03prefix + ' '+ val + propertiesAtributteTwo03dustPrefix);
        InputColumnsAtributteTwo04 = InputGadgetsValue.map(val => propertiesAtributteTwo04prefix + ' '+ val + propertiesAtributteTwo04dustPrefix);
        InputColumnsAtributteTwo05 = InputGadgetsValue.map(val => propertiesAtributteTwo05prefix + ' '+ val + propertiesAtributteTwo05dustPrefix);
    } else { 
        InputGadgets = ['']; 
        //InputGadgets = InputGadgetsValue.length >= 1 ? InputGadgetsValue.map(val => prefixGadgets + val + dustPrefixGadgets) : ['']; 
    }
    if(prefixAtribute03 !== ""){
        InputAtribute03 = InputAtribute03Value.map(val => prefixAtribute03 + val + dustPrefixAtribute03); 
        InputColumnsAtributteThree01 = InputAtribute03Value.map(val => propertiesAtributteThree01prefix +' '+ val + propertiesAtributteThree01dustPrefix);
        InputColumnsAtributteThree02 = InputAtribute03Value.map(val => propertiesAtributteThree02prefix + ' '+ val + propertiesAtributteThree02dustPrefix);
        InputColumnsAtributteThree03 = InputAtribute03Value.map(val => propertiesAtributteThree03prefix + ' '+ val + propertiesAtributteThree03dustPrefix);
        InputColumnsAtributteThree04 = InputAtribute03Value.map(val => propertiesAtributteThree04prefix + ' '+ val + propertiesAtributteThree04dustPrefix);
        InputColumnsAtributteThree05 = InputAtribute03Value.map(val => propertiesAtributteThree05prefix + ' '+ val + propertiesAtributteThree05dustPrefix);
    }else{
        InputAtribute03 = ['']; 
        // InputAtribute03 = InputAtribute03Value.length >= 1 ? InputAtribute03Value.map(val => prefixAtribute03 + val + dustPrefixAtribute03) : [''];  
    }
}

function checkSheetSelection() {
    const copyBtn = document.getElementById('copyBtn');
    const downloadBtn = document.getElementById('downloadBtn');
    //sourceWorkbook && destinationWorkbook &&
    if (selectedSheetData && checked) {
        copyBtn.disabled = false;
        downloadBtn.disabled = false;
    } else {
        copyBtn.disabled = true;
        downloadBtn.disabled = true;
    }
}

async function updateTable() {
    const sourceSheet = selectedSheetData[0].data;
    await table(sourceSheet, getResultInput);
    processCommaSeparatedValues();
    destinationData.length = 0;
}

function ResultInput() {
    getResultInput = [];

    Colums =[];
    ColumnsAtributteOne01 =[];
    ColumnsAtributteOne02 =[];
    ColumnsAtributteOne03 =[];
    ColumnsAtributteOne04 =[];
    ColumnsAtributteOne05 =[];

    Gadgets = [];
    ColumnsAtributteTwo01 =[];
    ColumnsAtributteTwo02 =[];
    ColumnsAtributteTwo04 =[];
    ColumnsAtributteTwo05 =[];
    
    Atribute03 = [];
    ColumnsAtributteThree01 = [];
    ColumnsAtributteThree02 = [];
    ColumnsAtributteThree03 = [];
    ColumnsAtributteThree04 = [];
    ColumnsAtributteThree05 = [];


    for (let i = 0; i < InputGadgets.length; i++) {
        // Verifica se o array InputGadgets está vazio
        if (InputGadgets.length === 0) {
            continue; // Pula a iteração atual e passa para a próxima
        }

        for (let j = 0; j < InputColumns.length; j++) {
            if (InputColumns.length === 0) {
            continue; // Pula a iteração atual e passa para a próxima
            }

            for (let l = 0; l < InputAtribute03.length; l++) {

                if(InputColumns.length >=1){
                    const variableJ = InputColumns[j];
                    Colums.push(variableJ);

                    const atribute01 = InputColumnsAtributteOne01[j];

                    ColumnsAtributteOne01.push(atribute01);

                    const atribute02 = InputColumnsAtributteOne02[j];
                    ColumnsAtributteOne02.push(atribute02);

                    const atribute03 = InputColumnsAtributteOne03[j];
                    ColumnsAtributteOne03.push(atribute03);

                    const atribute04 = InputColumnsAtributteOne04[j];
                    ColumnsAtributteOne04.push(atribute04);

                    const atribute05 = InputColumnsAtributteOne05[j];
                    ColumnsAtributteOne05.push(atribute05);
                    
                }
                if(InputGadgets.length >=1){
                    const variableA = InputGadgets[i];
                    Gadgets.push(variableA);

                    const atribute01 = InputColumnsAtributteTwo01[i];
                    ColumnsAtributteTwo01.push(atribute01);

                    const atribute02 = InputColumnsAtributteTwo02[i];
                    ColumnsAtributteTwo02.push(atribute02);

                    const atribute03 = InputColumnsAtributteTwo03[i];
                    ColumnsAtributteTwo03.push(atribute03);

                    const atribute04 = InputColumnsAtributteTwo04[i];
                    ColumnsAtributteTwo04.push(atribute04);

                    const atribute05 = InputColumnsAtributteTwo05[i];
                    ColumnsAtributteTwo05.push(atribute05);
                }
                if(InputAtribute03.length >=1){
                    const variableB = InputAtribute03[l];
                    Atribute03.push(variableB);

                    const atribute01 = InputColumnsAtributteThree01[i];
                    ColumnsAtributteThree01.push(atribute01);

                    const atribute02 = InputColumnsAtributteThree02[i];
                    ColumnsAtributteThree02.push(atribute02);

                    const atribute03 = InputColumnsAtributteThree03[i];
                    ColumnsAtributteThree03.push(atribute03);

                    const atribute04 = InputColumnsAtributteThree04[i];
                    ColumnsAtributteThree04.push(atribute04);

                    const atribute05 = InputColumnsAtributteThree05[i];
                    ColumnsAtributteThree05.push(atribute05);
                }else{

                }
               
                getResultInput++;
            }
        }
    }
    console.log("Total combinations:", getResultInput);

}

async function copySheet() {
   // const sourceSheet = selectedSheetData[0].data;
    await updateTable();
    getResultInput = [];
    await ResultInput();
  
    for (let i = 0; i < getResultInput; i++) {
        destinationData = []; // mover a declaração para o início do loop
        // Limpar os dados antigos antes de adicionar novos dados
        destinationData.length = 0;
      
        // Percorrer cada célula da planilha de origem
        for (let i = 0; i < selectedSheetData.length; i++) {
            const sourceSheet = selectedSheetData[i].data;
            const sheetData = []; // Array para armazenar os dados da planilha atual
          
            for (let row = 0; row < sourceSheet.length; row++) {
              const rowData = [];
              for (let col = 0; col < sourceSheet[row].length; col++) {
                const cellValue = sourceSheet[row][col] ? sourceSheet[row][col].toString().trim() : '';
                rowData.push(cellValue);
              }
              sheetData.push(rowData);
            }
          
            destinationData.push(sheetData); // Adiciona os dados da planilha atual ao array de destino
        }
        for(let x = 0; x < destinationData.length; x++){
            let unitDestination = destinationData[x]
            // Substituir os valores na nova matriz, se corresponderem
            for (let row = 0; row < unitDestination.length; row++) {
                for (let col = 0; col < unitDestination[row].length; col++) {
                let replacedValue = unitDestination[row][col].toString();
            
                // Substituir o valor se corresponder a "toBeReplaced"
                if (prefixColumns !== "") {
                    if(replacedValue.includes(toBeReplaced) ){
                        replacedValue = replacedValue.replace(new RegExp(toBeReplaced, 'g'), Colums[i]);
                    }
        
                    if(propertiesAtributteOne01prefix !== ""){
                        replacedValue = replacedValue.replace(new RegExp(propertiesAtributteOne01, 'g'), ColumnsAtributteOne01[i]);
                    }
                    if(propertiesAtributteOne02prefix !== ""){
                        replacedValue = replacedValue.replace(new RegExp(propertiesAtributteOne02, 'g'), ColumnsAtributteOne02[i]);
                    }
                    if(propertiesAtributteOne03prefix !== ""){
                        replacedValue = replacedValue.replace(new RegExp(propertiesAtributteOne03, 'g'), ColumnsAtributteOne03[i]);
                    }
                    if(propertiesAtributteOne04prefix !== ""){
                        replacedValue = replacedValue.replace(new RegExp(propertiesAtributteOne04, 'g'), ColumnsAtributteOne04[i]);
                    }
                    if(propertiesAtributteOne05prefix !== ""){
                        replacedValue = replacedValue.replace(new RegExp(propertiesAtributteOne05, 'g'), ColumnsAtributteOne05[i]);
                    }
                }
            
                // Substituir o valor se corresponder a "toBeReplacedAtributte"
                if (prefixGadgets !== "") {
            
                    if(replacedValue.includes(toBeReplacedAtributte)){
                        replacedValue = replacedValue.replace(new RegExp(toBeReplacedAtributte, 'g'), Gadgets[i]);
                    }
        
                    if(propertiesAtributteTwo01prefix !== ""){
                        replacedValue = replacedValue.replace(new RegExp(propertiesAtributteTwo01, 'g'), ColumnsAtributteTwo01[i]);
                    }
                    if(propertiesAtributteTwo02prefix !== ""){
                        replacedValue = replacedValue.replace(new RegExp(propertiesAtributteTwo02, 'g'), ColumnsAtributteTwo02[i]);
                    }
                    if(propertiesAtributteTwo03prefix !== ""){
                        replacedValue = replacedValue.replace(new RegExp(propertiesAtributteTwo03, 'g'), ColumnsAtributteTwo03[i]);
                    }
                    if(propertiesAtributteTwo04prefix !== ""){
                        replacedValue = replacedValue.replace(new RegExp(propertiesAtributteTwo04, 'g'), ColumnsAtributteTwo04[i]);
                    }
                    if(propertiesAtributteTwo05prefix !== ""){
                        replacedValue = replacedValue.replace(new RegExp(propertiesAtributteTwo05, 'g'), ColumnsAtributteTwo05[i]);
                    }
                }
            
                // Substituir o valor se corresponder a "variableAtributes03"
        
                if (prefixAtribute03 !== "") {
        
                    if(replacedValue.includes(variableAtributes03)){
                        replacedValue = replacedValue.replace(new RegExp(variableAtributes03, 'g'), Atribute03[i]);
                    }
                    if(propertiesAtributteThree01prefix !== ""){
                        replacedValue = replacedValue.replace(new RegExp(propertiesAtributteThree01, 'g'), ColumnsAtributteThree01[i]);
                    }
                    if(propertiesAtributteThree02prefix !== ""){
                        replacedValue = replacedValue.replace(new RegExp(propertiesAtributteThree02, 'g'), ColumnsAtributteThree02[i]);
                    }
                    if(propertiesAtributteThree03prefix !== ""){
                        replacedValue = replacedValue.replace(new RegExp(propertiesAtributteThree03, 'g'), ColumnsAtributteThree03[i]);
                    }
                    if(propertiesAtributteThree04prefix !== ""){
                        replacedValue = replacedValue.replace(new RegExp(propertiesAtributteThree04, 'g'), ColumnsAtributteThree04[i]);
                    }
                    if(propertiesAtributteThree05prefix !== ""){
                        replacedValue = replacedValue.replace(new RegExp(propertiesAtributteThree05, 'g'), ColumnsAtributteThree05[i]);
                    }
                
                }
            
                unitDestination[row][col] = replacedValue;
        
                }
            }
            
            const sheetName = unitDestination['G5'] ? unitDestination['G5'].v : 'Nome da Planilha(G5)';
        
            let uniqueSheetName = sheetName;
            let sheetIndex = 1;
        
            while (destinationWorkbook.SheetNames.includes(uniqueSheetName)) {
                uniqueSheetName = sheetName + ' (' + sheetIndex + ')';
                sheetIndex++;
            }
            const destinationSheet = XLSX.utils.aoa_to_sheet(unitDestination);
        
            // Obter o valor da célula G5 da planilha copiada
            const newSheetName = destinationSheet['G5'] ? destinationSheet['G5'].v : 'Sheet';
            // Truncar o nome da planilha, se exceder o limite de 31 caracteres
            const truncatedSheetName = newSheetName.substring(0, 31);
        
            uniqueSheetName = truncatedSheetName;
            sheetIndex = 1;
        
            while (destinationWorkbook.SheetNames.includes(uniqueSheetName)) {
                uniqueSheetName = truncatedSheetName + ' (' + sheetIndex + ')';
                sheetIndex++;
            }
        
            destinationWorkbook.SheetNames.push(uniqueSheetName);
            destinationWorkbook.Sheets[uniqueSheetName] = destinationSheet;
        
            // Atualizar o valor no HTML
            document.getElementById('output').innerHTML += `
                <table>
                <tr>
                    <td>Planilha copiada</td>
                    <td>${uniqueSheetName}</td>
                </tr>
                <!-- Adicione mais linhas conforme necessário -->
                </table>
            `;
        
            // Habilitar o botão de download após a cópia
            document.getElementById('downloadBtn').disabled = false;
                      
        }        
    }  
}

function downloadModifiedWorkbook() {
    const data = XLSX.write(destinationWorkbook, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([data], { type: 'application/octet-stream' });
    const url = window.URL.createObjectURL(blob);

    const a = document.createElement('a');
    document.body.appendChild(a);
    a.style = 'display: none';
    a.href = url;
    a.download = 'planilha_modificada.xlsx';
    a.click();
    setTimeout(function () {
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
    }, 0);
}

processCommaSeparatedValues();

const destinationFileInput = document.getElementById('destinationFile');
const copyBtn = document.getElementById('copyBtn');
copyBtn.addEventListener('click', copySheet);
const downloadBtn = document.getElementById('downloadBtn');
destinationFileInput.addEventListener('change', handleDestinationFile);
downloadBtn.addEventListener('click', downloadModifiedWorkbook);


