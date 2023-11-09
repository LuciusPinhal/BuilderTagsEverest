let modifiedData = {}; // Armazena os valores modificados de G5, L5, M5 e N5 em cada planilha
let checked = false;
let checked2 = false;
let sourceWorkbook = null;
let sheetSequence = {}; // Armazena a sequência das planilhas

function checkSheetSelection() {

    console.warn("🔭🐢 >>  checked2:",  checked2);
    const copyBtn = document.getElementById('copyBtn');
    const downloadBtn = document.getElementById('downloadBtn');
    //sourceWorkbook && destinationWorkbook &&
    if (modifiedData) {
        copyBtn.disabled = false;
        downloadBtn.disabled = true;
   
        if(checked && checked2){
            downloadBtn.disabled = false;
        }
    } else {
        copyBtn.disabled = true;
        downloadBtn.disabled = true;
    }
}

async function handleSourceFile(e) {
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        sourceWorkbook = XLSX.read(data, { type: 'array' });
        var oversizedCells = [];

        sourceWorkbook.SheetNames.forEach(function(sheetName) {
            var sheet = sourceWorkbook.Sheets[sheetName];
            XLSX.utils.sheet_to_json(sheet).forEach(function(row) {
                Object.values(row).forEach(function(cellValue) {
                    if (typeof cellValue === 'string' && cellValue.length > 32767) {
                        oversizedCells.push({ sheet: sheetName, value: cellValue });
                    }
                });
            });
        });

        if (oversizedCells.length > 0) {
            var message = "As seguintes células excederam 32767 caracteres:\n\n";
            oversizedCells.forEach(function(cell) {
                message += "Planilha: " + cell.sheet +"\n\n"
            });
            alert(message);
        } 
        
        const sheetTableBody = document.querySelector('#sheetTable tbody');
        sheetTableBody.innerHTML = '';

        const sheetTableHead = document.querySelector('#sheetTable thead');
        sheetTableHead.innerHTML = '';

        // Cabeçalho da tabela
        const columnTitles = ['Sequência', 'Nome da Planilha', 'Comando', 'Sequencia', 'Ativo'];

        // Criação da linha de cabeçalho
        const headerRow = document.createElement('tr');
        columnTitles.forEach((titleText) => {
            const titleCell = document.createElement('th');
            titleCell.textContent = titleText;
            headerRow.appendChild(titleCell);
        });

        // Adicionar a linha de cabeçalho à tabela
        sheetTableHead.appendChild(headerRow);

        // Array com os títulos das colunas L5, M5 e N5
        const cellAddresses = ['G5', 'L5', 'M5', 'N5'];

        sourceWorkbook.SheetNames.forEach(function (sheetName, index) {
            const row = document.createElement('tr');

            const sequenceCell = document.createElement('td');
            const sequenceInput = document.createElement('input');
            sequenceInput.type = 'number';
            sequenceInput.value = index + 1;
            sequenceInput.addEventListener('input', function () {
                // Aqui você pode armazenar a sequência definida pelo usuário ou fazer alguma ação com ela
                checked2 = false;
                checkSheetSelection();
                console.log('Sequência definida para', sheetName, ':', sequenceInput.value);
            });
            sequenceCell.appendChild(sequenceInput);
            row.appendChild(sequenceCell);

            const sheetData = sourceWorkbook.Sheets[sheetName];

            cellAddresses.forEach((address) => {
                const cell = sheetData[address];
                const inputValue = cell && cell.v ? (cell.t === 'n' ? cell.v : sheetData['!sheetjs'] ? cell.w : cell.v) : '';

                const cellValueCell = document.createElement('td');
                const inputElement = document.createElement('input');
                inputElement.type = 'text';
                inputElement.value = inputValue;
                inputElement.addEventListener('input', function(){
                  checked2 = false;
                  checkSheetSelection();
               })

                // Desabilita o input somente para os labels G5
                if (address === 'G5') {
                    inputElement.disabled = true;
                }

                // Aqui você pode adicionar um evento de listener para capturar a alteração do valor
                inputElement.addEventListener('input', function () {
                    // Aqui você pode atualizar o valor do input em relação ao valor da célula
                    const newCellValue = inputElement.value;
                    if (cell && cell.v) {
                        if (cell.t === 'n') {
                            cell.v = parseFloat(newCellValue);
                        } else if (cell.t === 's') {
                            cell.v = newCellValue;
                        }
                    }

                    // Armazenar o valor modificado no objeto modifiedData
                    if (!modifiedData[sheetName]) {
                        modifiedData[sheetName] = {};
                    }
                    modifiedData[sheetName][address] = newCellValue;
                });

                cellValueCell.appendChild(inputElement);
                row.appendChild(cellValueCell);
            });

            sheetTableBody.appendChild(row);
        });

        checkSheetSelection();
    };
  
    reader.readAsArrayBuffer(file);
}

function copyValuesToWorksheets() {
  // Aguarda a conclusão da função changeSheetSequence()

    sourceWorkbook.SheetNames.forEach(function (sheetName) {
        const sheetData = sourceWorkbook.Sheets[sheetName];

        // Verifica se há valores modificados para esta planilha
        if (modifiedData[sheetName]) {
            Object.keys(modifiedData[sheetName]).forEach((address) => {
                const modifiedValue = modifiedData[sheetName][address];
                const cell = sheetData[address];

                // Substitui o valor modificado na célula, independentemente do valor original
                if (cell) {
                    if (!isNaN(modifiedValue) && !isNaN(parseFloat(modifiedValue))) {
                        cell.t = 'n'; // Define o tipo da célula como número
                        cell.v = parseFloat(modifiedValue);
                        cell.w = modifiedValue; // Atualiza também o valor exibido na célula (opcional)
                    } else {
                        cell.t = 's'; // Define o tipo da célula como string
                        cell.v = modifiedValue;
                        cell.w = modifiedValue; // Atualiza também o valor exibido na célula (opcional)
                    }
                } else {
                    // Se a célula não existe, cria uma nova célula para o endereço
                    const newCell = !isNaN(modifiedValue) && !isNaN(parseFloat(modifiedValue))
                        ? { t: 'n', v: parseFloat(modifiedValue) }
                        : { t: 's', v: modifiedValue };

                    sheetData[address] = newCell;
                }
            });
        }
    });

    // Salvar a planilha modificada em algum lugar ou fazer outras ações necessárias
    console.log('Planilha modificada:', sourceWorkbook);
    checked2 = true;
    checkSheetSelection();
}

async function changeSheetSequence() {
  const rows = Array.from(document.querySelectorAll('#sheetTable tbody tr'));

  sheetSequence = {}; // Limpa a sequência atual antes de atualizá-la novamente

  rows.forEach((row, index) => {
      const sheetName = sourceWorkbook.SheetNames[index];
      const sequence = parseInt(row.querySelector('input[type="number"]').value);

      // Armazena a sequência de cada planilha no objeto sheetSequence
      sheetSequence[sheetName] = sequence;
  });

  // Ordena as planilhas com base na sequência atualizada
  sourceWorkbook.SheetNames.sort((sheetNameA, sheetNameB) => {
      const sequenceA = sheetSequence[sheetNameA];
      const sequenceB = sheetSequence[sheetNameB];
      return sequenceA - sequenceB;
  });

  // Atualiza a tabela com as novas sequências
  const sheetTableBody = document.querySelector('#sheetTable tbody');
  sheetTableBody.innerHTML = ''; // Limpa o conteúdo da tabela

  sourceWorkbook.SheetNames.forEach((sheetName, index) => {
      const row = rows.find((row) => {
          const sheetInput = row.querySelector('input[type="number"]');
          return parseInt(sheetInput.value) === index + 1;
      });

      if (row) {
          // Atualiza o valor do campo de sequência (input) para corresponder à nova sequência após a reordenação
          const sequenceInput = row.querySelector('input[type="number"]');
          sequenceInput.value = index + 1;

          sheetTableBody.appendChild(row); // Adiciona a linha (row) da planilha reordenada ao tbody da tabela
      }
  });

  // Atualiza a variável checked após alterar a sequência
  checked = true;
  checkSheetSelection();

  const loading = document.getElementById('loadingBody');
  loading.style.display = 'flex'; 

  setTimeout(() => {
    copyValuesToWorksheets();
    loading.style.display = 'none';

  }, 3000);
}



// Função para baixar a planilha modificada
function downloadModifiedWorkbook() {
    // Verifica se a planilha foi modificada
    if (!Object.keys(modifiedData).length) {
        alert('Não há planilha modificada para baixar.');
        return;
    }

    // Cria o arquivo XLSX da planilha modificada
    const wbout = XLSX.write(sourceWorkbook, { bookType: 'xlsx', type: 'array' });

    // Converte o XLSX para um objeto Blob
    const blob = new Blob([wbout], { type: 'application/octet-stream' });

    // Cria um URL temporário para o Blob
    const url = URL.createObjectURL(blob);

    // Cria um elemento de link e simula um clique para iniciar o download
    const a = document.createElement('a');
    a.href = url;
    a.download = 'planilha_modificada.xlsx'; // Nome do arquivo para download (pode ser personalizado)
    a.click();

    // Libera o URL temporário criado
    URL.revokeObjectURL(url);
}

// Eventos de clique e alteração

document.getElementById('sourceFile').addEventListener('change', handleSourceFile);

document.getElementById('copyBtn').addEventListener('click', changeSheetSequence);

document.getElementById('downloadBtn').addEventListener('click', downloadModifiedWorkbook);