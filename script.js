let tabelaGlobal = null;  

function changeTheme() {  
    const theme = document.getElementById('themeSelector').value;  
    document.body.classList.toggle('dark-mode', theme === 'dark');  
}  

function lerPlanilha() {  
    const fileInput = document.getElementById('fileInput');  
    const resultadoDiv = document.getElementById('resultado');  
    resultadoDiv.innerHTML = '';  
    document.getElementById('countContainer').textContent = 'Total de resultados encontrados: 0';  

    const file = fileInput.files[0];  
    if (!file) {  
        alert('Por favor, selecione um arquivo.');  
        return;  
    }  
    
    const reader = new FileReader();  
    reader.onload = function(event) {  
        const data = new Uint8Array(event.target.result);  
        const workbook = XLSX.read(data, { type: 'array' });  
        const sheetName = workbook.SheetNames[0];  
        const worksheet = workbook.Sheets[sheetName];  
        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });  
        
        const tabela = document.createElement('table');  
        tabela.className = 'table table-striped';  
        tabelaGlobal = tabela;  

        // Criar cabeçalho da tabela  
        const thead = document.createElement('thead');  
        const headerRow = document.createElement('tr');  
        json[0].forEach(coluna => {  
            const th = document.createElement('th');  
            th.textContent = coluna.trim();  
            headerRow.appendChild(th);  
        });  
        thead.appendChild(headerRow);  
        tabela.appendChild(thead);  

        // Criar corpo da tabela  
        const tbody = document.createElement('tbody');  
        json.slice(1).forEach((linha) => {  
            const tr = document.createElement('tr');  
            linha.forEach((valor) => {  
                const td = document.createElement('td');  
                td.textContent = valor;  
                const commentInput = document.createElement('input');  
                commentInput.type = 'text';  
                commentInput.placeholder = 'Adicionar comentário';  
                commentInput.className = 'comment form-control';  
                td.appendChild(commentInput);  
                tr.appendChild(td);  
            });  
            tbody.appendChild(tr);  
        });  
        tabela.appendChild(tbody);  
        resultadoDiv.appendChild(tabela);  
    };  
    reader.readAsArrayBuffer(file);  
}  

function filtrarTabela() {  
    const filterValue = document.getElementById('filterInput').value.toLowerCase();  
    const rows = tabelaGlobal.getElementsByTagName('tr');  
    let count = 0;  

    for (let i = 1; i < rows.length; i++) {  
        const cells = rows[i].getElementsByTagName('td');  
        let rowVisible = false;  

        for (let j = 0; j < cells.length; j++) {  
            const cellValue = cells[j].textContent.toLowerCase();  
            if (cellValue.includes(filterValue) && filterValue) {  
                rowVisible = true;  
                count++;  
                const highlightedText = cellValue.replace(new RegExp(`(${filterValue})`, 'gi'), '<span class="highlight">$1</span>');  
                cells[j].innerHTML = highlightedText;  
            } else {  
                cells[j].innerHTML = cells[j].textContent; // Reset cell content if not matching  
            }  
        }  
        rows[i].style.display = rowVisible ? '' : 'none';  
    }  
    document.getElementById('countContainer').textContent = `Total de resultados encontrados: ${count}`;  
}