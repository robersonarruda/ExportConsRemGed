// ==UserScript==
// @name         Exporta Consulta Rematriculas Ged
// @version      1.1.6
// @description  Exporta em excel os dados exibidos na consulta realizada no painel rematrícula do GED.
// @author       Roberson Arruda
// @match		  http://*.seduc.mt.gov.br/ged/hwmgedpainelrematricula.aspx*
// @match		  https://*.seduc.mt.gov.br/ged/hwmgedpainelrematricula.aspx*
// @homepage      https://github.com/robersonarruda/ExportConsRemGed/blob/main/ExportConsRemGed.user.js
// @downloadURL   https://github.com/robersonarruda/ExportConsRemGed/raw/main/ExportConsRemGed.user.js
// @updateURL     https://github.com/robersonarruda/ExportConsRemGed/raw/main/ExportConsRemGed.user.js
// @copyright  2024, Roberson Arruda (robersonarruda@outlook.com)
// ==/UserScript==


(function() {
    'use strict';

    // Estilos CSS
    const style = document.createElement('style');
    style.innerHTML = `
    .button {
    border: 2px solid #04AA6D;
    transition-duration: 0.4s;
    background-color:#fff;
    }
    .button:hover {
    background-color: #04AA6D; /* Green */
    color: white;
    }`;
    document.head.appendChild(style);

    // Cria um botão e adiciona ao DOM
    const exportButton = document.createElement('button');
    exportButton.innerText = 'Exportar Resultados da Consulta(Excel)';
    exportButton.style.position = 'fixed';
    exportButton.style.top = '20px';
    exportButton.style.right = '30px';
    exportButton.style.height = "35px"
    exportButton.style.zIndex = 1000; // Garante que o botão fique acima de outros elementos
    exportButton.classList.add('button');
    document.body.appendChild(exportButton);

    //Variáveis globais
    let abortar = false;

    exportButton.addEventListener('click', function() {
        // EXPORTAR PARA EXCEL
        var script = document.createElement('script');
        script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.5/xlsx.full.min.js';
        document.head.appendChild(script);

        script.onload = async function () {
            // Variável para armazenar todos os dados das páginas
            let allData = [];
            let isFirstPage = true;

            // Defina quais colunas remover
            const RemoverColuna = [
                "Grid_ Ged Rem Ger Amb Cod Mask",
                "Grid_ Ged Rem Ger Tur Cod",
                "Aluno",
                "Grid_ Ged Rem Ger Mat Cod",
                "Grid_ Ged Rem Ger Mat Dsc Mtz Cpl",
                "Grid_ Ged Rem Ger Ano Let Cod",
                "Grid_ Ged Rem Ger Lot Cod",
                "Grid_ Ged Rem Ger Amb Cod",
                "Grid_ Ger Tur Fin Mat",
                "Grid_ Ger Tur Dta Ini",
                "Grid_ Ger Tur Dta Fim",
                "Grid_ Ged Rem Ger Trn Cod",
                "Turno",
                "Justificativa",
                "Grid_ Ged Rem Ori Ger Ano Let Cod",
                "Grid_ Ged Rem Ori Ger Mat Cod",
                "Ged Mat Ano Let Fin",
                "Is Liberar Matricula",
                "Is Liberar Estorno",
                "Grid_ Ged Rem Alu Id"
            ];
            // Defina quais cabeçalhos renomear
            const RenomearCabecalho = {
                "Grid_ Ged Rem Ged Alu Nom": "Nome do Aluno",
                "Grid_ Ged Alu Rec Ate Edu Esp": "Recebe Atendimento Educ Especial",
                "Grid_ Ged Rem Ged Alu Cod": "Cod Aluno"
            };

            // Função para aguardar o carregamento da página
            function waitForLoad(previousContent) {
                return new Promise((resolve) => {
                    const checkInterval = setInterval(() => {
                        const ajaxNotification = document.getElementById('gx_ajax_notification');
                        const newContent = document.getElementById('span_vGRID_GEDREMALUIDENTIFICADOR_0001').innerHTML;

                        if (ajaxNotification.style.display === 'none' && newContent !== previousContent) {
                            clearInterval(checkInterval);
                            resolve(newContent); // Retorna o novo conteúdo para comparação na próxima página
                        }
                    }, 100);
                });
            }

            // Função para coletar dados da tabela atual e adicionar ao allData
            function collectTableData() {
                const table = document.getElementById('GridrematriculaContainerTbl');
                const ws = XLSX.utils.table_to_sheet(table);

                const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

                // Extrair cabeçalhos da primeira linha
                const headerRow = rows[0];

                // Remover colunas indesejadas e renomear cabeçalhos
                const filteredHeaders = headerRow.filter(header => !RemoverColuna.includes(header));
                const headerMap = headerRow.map(header => RenomearCabecalho[header] || header);

                // No caso da primeira página, adiciona o cabeçalho
                if (isFirstPage) {
                    // Atualiza cabeçalhos com renomeação
                    rows[0] = headerMap;

                    // Remover colunas indesejadas do cabeçalho
                    const columnIndicesToKeep = headerRow.map((header, index) =>
                        !RemoverColuna.includes(header) ? index : -1
                    ).filter(index => index !== -1);

                    // Filtra os dados conforme os índices das colunas a manter
                    rows.forEach((row, rowIndex) => {
                        rows[rowIndex] = columnIndicesToKeep.map(index => row[index]);
                    });

                    allData.push(...rows); // Adiciona a primeira página inteira (cabeçalho + dados)
                } else {
                    // Remove o cabeçalho das páginas subsequentes
                    rows.shift(); // Remove a linha de cabeçalho para páginas subsequentes
                    const columnIndicesToKeep = headerRow.map((header, index) =>
                        !RemoverColuna.includes(header) ? index : -1
                    ).filter(index => index !== -1);

                    // Filtra os dados conforme os índices das colunas a manter
                    rows.forEach((row, rowIndex) => {
                        rows[rowIndex] = columnIndicesToKeep.map(index => row[index]);
                    });

                    allData.push(...rows); // Adiciona os dados das páginas subsequentes
                }

                isFirstPage = false; // Marca que o cabeçalho já foi adicionado
            }

            // Função principal para iterar pelas páginas e coletar dados
            async function fetchAllPages() {
                const select = document.getElementById('vPAG');
                let previousContent = "";

                const initialElement = document.getElementById('span_vGRID_GEDREMALUIDENTIFICADOR_0001');
                if (!initialElement) {
                    alert("Falha ao exportar: Defina os parâmetros de resultados e clique no botão de consultar para exibi-los antes de clicar em exportar");
                    abortar = true;
                    return; // Aborta a operação
                }

                collectTableData(); // Coleta dados da primeira página
                previousContent = initialElement.innerHTML;

                for (let i = 2; i <= select.options.length; i++) {
                    select.value = i; // Seleciona a página
                    const changeEvent = new Event('change');
                    select.dispatchEvent(changeEvent); // Dispara o evento de mudança

                    previousContent = await waitForLoad(previousContent);
                    collectTableData(); // Coleta os dados da tabela atual
                }
            }

            // Executa a coleta de todas as páginas e exporta para Excel
            await fetchAllPages();

            // Cria a planilha e adiciona os dados
            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet(allData);
            XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

            // Salva o arquivo Excel
            if(!abortar){
                XLSX.writeFile(wb, 'Consulta Rematriculas.xlsx');
            }
        };
    });
})();
