document.addEventListener('DOMContentLoaded', () => {
    // Estado global da aplicação
    let data = {};
    let services = [];
    let editIndex = -1;
    const configKeys = [
        { id: 'contratoSap', label: 'Contrato SAP' },
        { id: 'materialServico', label: 'Material/Serviço' },
        { id: 'pep', label: 'PEP' },
        { id: 'centroSap', label: 'Centro SAP' },
        { id: 'municipioNf', label: 'Município Geração NF' },
        { id: 'cnpj', label: 'CNPJ' }
    ];

    // Funções Utilitárias
    const parseCurrency = (str) => parseFloat((str || '0').replace(/[R$\s.]/g, '').replace(',', '.') || 0);
    const formatCurrency = (num) => `R$ ${num.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
    const sanitizeSheetName = (name) => name.replace(/[:\\/?*[\]]/g, '').substring(0, 31);
    const showAlert = (message, type = 'success', duration = 4000) => {
        const container = document.getElementById('export-status') || document.body;
        const alertDiv = document.createElement('div');
        alertDiv.className = `alert ${type}`;
        alertDiv.textContent = message;
        container.prepend(alertDiv);
        setTimeout(() => { alertDiv.remove(); }, duration);
    };

    // Gerenciamento de Dados (LocalStorage)
    const saveToLocalStorage = () => localStorage.setItem('contratoData', JSON.stringify(data));
    const loadFromLocalStorage = () => {
        const savedData = localStorage.getItem('contratoData');
        const defaultConfig = {
            cities: [], quantidades: {}, servicos: "Fiscalização com instalação de nova ligação,Fiscalização com religação de abastecimento,Fiscalização com remanejamento de cavalete,Fiscalização com subst./instal. HD,Fiscalização sem acesso ao cavalete,Fiscalização sem irregularidade,Fisclização com fraude - retirada bypass/lig clandestina,Fisclização com fraude - supressão ligação,Fisclização com fraude - suspensão ramal,Sondagem,Recomposição Pavimento",
            opexMetro: "28.3,161.89,28.3,26.9,28.33,185.15,180.42,168.08,99.43,124.01,71.25", capexMetro: "285.22,161.89,225.94,17.91,26.9,28.33,185.15,180.42,168.08,99.43,124.01",
            opexInterior: "28.3,161.89,28.3,26.9,28.33,185.15,180.42,168.08,99.43,124.01,71.25", capexInterior: "285.22,161.89,225.94,17.91,26.9,28.33,185.15,180.42,168.08,99.43,124.01",
        };
        configKeys.forEach(k => defaultConfig[k.id] = {});
        data = savedData ? { ...defaultConfig, ...JSON.parse(savedData) } : defaultConfig;
        services = data.servicos.split(',');
        
        Object.keys(data).forEach(key => {
            const elementId = key.replace(/[A-Z]/g, letter => `-${letter.toLowerCase()}`);
            const element = document.getElementById(elementId);
            if (element) element.value = data[key];
        });
        
        updateAllRenders();
    };

    // Funções de Renderização e Atualização da UI
    const updateBalance = () => {
        data.valorContrato = document.getElementById('valor-contrato').value;
        data.valorAditivo = document.getElementById('valor-aditivo').value;
        data.vlrUtilizado = document.getElementById('vlr-utilizado').value;
        const valorContrato = parseCurrency(data.valorContrato);
        const valorAditivo = parseCurrency(data.valorAditivo);
        const vlrUtilizado = parseCurrency(data.vlrUtilizado);
        const total = valorContrato + valorAditivo;
        const saldo = total - vlrUtilizado;
        document.getElementById('total-contratos').innerText = `${formatCurrency(total)} - Valor Total do Contrato`;
        document.getElementById('saldo-disponivel').innerText = `${formatCurrency(saldo)} - Saldo Disponível`;
    };

    const renderCities = () => {
        const tbody = document.getElementById('cidades-table').querySelector('tbody');
        tbody.innerHTML = '';
        data.cities.forEach((city, i) => {
            const tr = tbody.insertRow();
            tr.innerHTML = `<td>${city.nome}</td><td>${city.regional}</td><td>${city.centro}</td>
                            <td><button class="action-btn edit-btn" data-index="${i}">Editar</button>
                                <button class="action-btn delete-btn" data-index="${i}">Remover</button></td>`;
        });
    };

    const renderQuantidades = () => {
        const form = document.getElementById('quantidades-form');
        if (data.cities.length === 0) {
            form.innerHTML = '<p>Cadastre cidades para definir quantidades.</p>';
            return;
        }
        let html = '<table><thead><tr><th>Cidade</th>' + services.map(s => `<th title="${s}">${s.substring(0, 15)}${s.length > 15 ? '...' : ''}</th>`).join('') + '</tr></thead><tbody>';
        data.cities.forEach(city => {
            html += `<tr><td>${city.nome}</td>` + services.map((_, i) => {
                const val = (data.quantidades[city.nome] && data.quantidades[city.nome][i]) || '';
                return `<td><input type="number" min="0" class="quant-input" data-city="${city.nome}" data-index="${i}" value="${val}"></td>`;
            }).join('') + '</tr>';
        });
        form.innerHTML = html + '</tbody></table>';
    };

    const renderConfiguracoes = () => {
        const form = document.getElementById('configuracoes-form');
        if (data.cities.length === 0) {
            form.innerHTML = '<p>Cadastre cidades para definir as configurações.</p>';
            return;
        }
        let html = '<table><thead><tr><th>Cidade</th>' + configKeys.map(k => `<th>${k.label}</th>`).join('') + '</tr></thead><tbody>';
        data.cities.forEach(city => {
            html += `<tr><td>${city.nome}</td>` + configKeys.map(key => {
                const val = data[key.id][city.nome] || '';
                return `<td><input type="text" class="config-input" data-city="${city.nome}" data-key="${key.id}" value="${val}"></td>`;
            }).join('') + '</tr>';
        });
        form.innerHTML = html + '</tbody></table>';
    };

    const updateAllRenders = () => {
        updateBalance();
        renderCities();
        renderQuantidades();
        renderConfiguracoes();
    };

    // Manipuladores de Eventos
    const setupEventListeners = () => {
        document.getElementById('salvar-dados').addEventListener('click', () => {
            const formIds = ['nomeContratado', 'cnpjContratado', 'enderecoContratado', 'municipioContratado', 'contatoContratado', 'telefoneContratado', 'emailContratado', 'contratoFisico', 'contratoSapGlobal', 'codFornecedor', 'valorContrato', 'valorAditivo', 'prazoContrato', 'execInicio', 'execTermino', 'vigencia', 'vigTermino', 'vlrUtilizado', 'nomeCliente', 'cnpjCliente', 'enderecoCliente', 'municipioCliente', 'contatoCliente', 'telefoneCliente', 'emailCliente', 'objContrato'];
            formIds.forEach(id => {
                const elementId = id.replace(/[A-Z]/g, letter => `-${letter.toLowerCase()}`);
                data[id] = document.getElementById(elementId).value;
            });
            saveToLocalStorage();
            updateBalance();
            showAlert('Dados do contrato salvos com sucesso!');
        });

        document.getElementById('salvar-precos').addEventListener('click', () => {
            data.servicos = document.getElementById('servicos-disponiveis').value;
            services = data.servicos.split(',');
            ['opexMetro', 'capexMetro', 'opexInterior', 'capexInterior'].forEach(id => {
                const elementId = id.replace(/[A-Z]/g, letter => `-${letter.toLowerCase()}`);
                data[id] = document.getElementById(elementId).value;
            });
            saveToLocalStorage();
            updateAllRenders();
            showAlert('Preços salvos com sucesso!');
        });

        document.getElementById('add-cidade').addEventListener('click', () => {
            const nome = document.getElementById('nome-cidade').value.trim();
            const regional = document.getElementById('regional-cidade').value;
            const centro = document.getElementById('centro-cidade').value.trim();
            
            if (!nome || !regional || !centro) {
                return showAlert('Preencha todos os campos da cidade.', 'error');
            }

            if (editIndex === -1) {
                if (data.cities.some(c => c.nome.toLowerCase() === nome.toLowerCase())) {
                    return showAlert(`A cidade "${nome}" já existe.`, 'error');
                }
                data.cities.push({ nome, regional, centro });
            } else {
                const oldName = data.cities[editIndex].nome;
                data.cities[editIndex] = { nome, regional, centro };
                if (oldName.toLowerCase() !== nome.toLowerCase()) {
                    ['quantidades', ...configKeys.map(k => k.id)].forEach(key => {
                        if (data[key] && data[key][oldName]) {
                            data[key][nome] = data[key][oldName];
                            delete data[key][oldName];
                        }
                    });
                }
                editIndex = -1;
                document.getElementById('add-cidade').innerText = 'Adicionar Cidade';
            }
            
            document.getElementById('nome-cidade').value = '';
            document.getElementById('regional-cidade').value = '';
            document.getElementById('centro-cidade').value = '';
            saveToLocalStorage();
            updateAllRenders();
            showAlert('Cidade salva com sucesso!');
        });

        document.getElementById('import-btn').addEventListener('click', () => {
            const lines = document.getElementById('import-cidades').value.split('\n');
            let imported = 0;
            lines.forEach(line => {
                const parts = line.split(',').map(p => p.trim());
                if (parts.length === 3 && parts[0] && ['Metropolitana', 'Interior'].includes(parts[1]) && parts[2]) {
                    if (!data.cities.some(c => c.nome.toLowerCase() === parts[0].toLowerCase())) {
                        data.cities.push({ nome: parts[0], regional: parts[1], centro: parts[2] });
                        imported++;
                    }
                }
            });
            if (imported > 0) {
                saveToLocalStorage();
                updateAllRenders();
                document.getElementById('import-cidades').value = '';
                showAlert(`${imported} cidades importadas com sucesso!`);
            } else {
                showAlert('Nenhuma cidade nova ou válida encontrada para importação.', 'error');
            }
        });

        document.body.addEventListener('click', (e) => {
            if (e.target.classList.contains('edit-btn')) {
                const index = parseInt(e.target.dataset.index);
                const city = data.cities[index];
                document.getElementById('nome-cidade').value = city.nome;
                document.getElementById('regional-cidade').value = city.regional;
                document.getElementById('centro-cidade').value = city.centro;
                editIndex = index;
                document.getElementById('add-cidade').innerText = 'Salvar Alterações';
                document.getElementById('nome-cidade').focus();
            }
            if (e.target.classList.contains('delete-btn')) {
                const index = parseInt(e.target.dataset.index);
                const cityName = data.cities[index].nome;
                if (confirm(`Tem certeza que deseja remover a cidade ${cityName}?`)) {
                    data.cities.splice(index, 1);
                    ['quantidades', ...configKeys.map(k => k.id)].forEach(key => {
                        if (data[key]) delete data[key][cityName];
                    });
                    saveToLocalStorage();
                    updateAllRenders();
                    showAlert('Cidade removida com sucesso!');
                }
            }
        });

        document.body.addEventListener('input', (e) => {
            const { city, index, key } = e.target.dataset;
            if (e.target.classList.contains('quant-input')) {
                if (!data.quantidades[city]) data.quantidades[city] = Array(services.length).fill('');
                data.quantidades[city][parseInt(index)] = e.target.value;
            }
            if (e.target.classList.contains('config-input')) {
                if (!data[key]) data[key] = {};
                data[key][city] = e.target.value;
            }
            if (e.target.closest('.section')) { // Salva no LocalStorage ao digitar em qualquer campo
                saveToLocalStorage();
            }
        });

        ['valor-contrato', 'valor-aditivo', 'vlr-utilizado'].forEach(id => {
            document.getElementById(id).addEventListener('input', (e) => {
                let value = e.target.value.replace(/\D/g, '');
                value = (value / 100).toFixed(2).replace('.', ',');
                e.target.value = value.replace(/(\d)(?=(\d{3})+(,.*)?$)/g, '$1.');
                updateBalance();
                saveToLocalStorage();
            });
        });

        document.getElementById('calcular-btn').addEventListener('click', () => {
            let totalOpex = 0, totalCapex = 0;
            data.cities.forEach(city => {
                const regional = city.regional.toLowerCase();
                const opexPrices = (regional === 'metropolitana' ? data.opexMetro : data.opexInterior).split(',').map(p => parseFloat(p) || 0);
                const capexPrices = (regional === 'metropolitana' ? data.capexMetro : data.capexInterior).split(',').map(p => parseFloat(p) || 0);
                const quants = data.quantidades[city.nome] || [];
                
                services.forEach((_, i) => {
                    const q = parseFloat(quants[i] || 0);
                    totalOpex += q * (opexPrices[i] || 0);
                    totalCapex += q * (capexPrices[i] || 0);
                });
            });
            
            document.getElementById('resultados').innerHTML = `
                <div class="summary">
                    <p>Total OPEX: ${formatCurrency(totalOpex)}</p>
                    <p>Total CAPEX: ${formatCurrency(totalCapex)}</p>
                    <p>Total Geral: ${formatCurrency(totalOpex + totalCapex)}</p>
                </div>`;
        });

        document.getElementById('export-json').addEventListener('click', () => {
            const a = document.createElement('a');
            a.href = 'data:text/json;charset=utf-8,' + encodeURIComponent(JSON.stringify(data, null, 2));
            a.download = 'contratoData.json';
            a.click();
            showAlert('Arquivo JSON exportado com sucesso!');
        });

        document.getElementById('export-excel').addEventListener('click', () => {
            try {
                const wb = XLSX.utils.book_new();
                
                // Aba de Resumo Executivo
                const valorContrato = parseCurrency(data.valorContrato);
                const valorAditivo = parseCurrency(data.valorAditivo);
                const vlrUtilizado = parseCurrency(data.vlrUtilizado);
                const totalContrato = valorContrato + valorAditivo;
                const saldoDisponivel = totalContrato - vlrUtilizado;
                
                let totalGeralOpex = 0, totalGeralCapex = 0;
                data.cities.forEach(city => {
                    const regional = city.regional.toLowerCase();
                    const opexPrices = (regional === 'metropolitana' ? data.opexMetro : data.opexInterior).split(',').map(p => parseFloat(p) || 0);
                    const capexPrices = (regional === 'metropolitana' ? data.capexMetro : data.capexInterior).split(',').map(p => parseFloat(p) || 0);
                    const quants = data.quantidades[city.nome] || [];
                    services.forEach((_, i) => {
                        const q = parseFloat(quants[i] || 0);
                        totalGeralOpex += q * (opexPrices[i] || 0);
                        totalGeralCapex += q * (capexPrices[i] || 0);
                    });
                });

                const resumoData = [
                    ['RESUMO EXECUTIVO DO CONTRATO'],
                    ['Valor do Contrato', valorContrato],
                    ['Valor Aditivo', valorAditivo],
                    ['Valor Total do Contrato', totalContrato],
                    ['Valor Utilizado', vlrUtilizado],
                    ['Saldo Disponível', saldoDisponivel],
                    [],
                    ['TOTAIS DOS SERVIÇOS CALCULADOS'],
                    ['Total OPEX', totalGeralOpex],
                    ['Total CAPEX', totalGeralCapex],
                    ['Total Geral (OPEX + CAPEX)', totalGeralOpex + totalGeralCapex],
                ];
                const resumoWs = XLSX.utils.aoa_to_sheet(resumoData);
                XLSX.utils.book_append_sheet(wb, resumoWs, "Resumo Executivo");

                // Aba de Cálculos Detalhados
                const calcHeader = ['Cidade', 'Regional', 'Serviço', 'Quantidade', 'Preço OPEX', 'Valor OPEX', 'Preço CAPEX', 'Valor CAPEX'];
                const calcData = [calcHeader];
                data.cities.forEach(city => {
                    const regional = city.regional.toLowerCase();
                    const opexPrices = (regional === 'metropolitana' ? data.opexMetro : data.opexInterior).split(',').map(p => parseFloat(p) || 0);
                    const capexPrices = (regional === 'metropolitana' ? data.capexMetro : data.capexInterior).split(',').map(p => parseFloat(p) || 0);
                    const quants = data.quantidades[city.nome] || [];
                    
                    services.forEach((serv, i) => {
                        const q = parseFloat(quants[i] || 0);
                        const pOpex = opexPrices[i] || 0;
                        const vOpex = q * pOpex;
                        const pCapex = capexPrices[i] || 0;
                        const vCapex = q * pCapex;
                        if (q > 0) { // Adiciona apenas linhas com quantidade
                           calcData.push([city.nome, city.regional, serv, q, pOpex, vOpex, pCapex, vCapex]);
                        }
                    });
                });
                const calcWs = XLSX.utils.aoa_to_sheet(calcData);
                XLSX.utils.book_append_sheet(wb, calcWs, "Cálculos Detalhados");

                // Abas de Dados
                const dataSheets = [
                    { title: "Detalhes Contrato", data: Object.entries(data).filter(([k]) => !['cities', 'quantidades', ...configKeys.map(c => c.id)].includes(k)) },
                    { title: "Cidades", data: [['Nome', 'Regional', 'Centro'], ...data.cities.map(c => [c.nome, c.regional, c.centro])] },
                    { title: "Quantidades", data: [['Cidade', ...services], ...data.cities.map(c => [c.nome, ...services.map((_, i) => (data.quantidades[c.nome] || [])[i] || 0)])] },
                    ...configKeys.map(k => ({
                        title: k.label,
                        data: [['Cidade', k.label], ...data.cities.map(c => [c.nome, data[k.id][c.nome] || ''])]
                    }))
                ];
                dataSheets.forEach(sheetInfo => {
                    const ws = XLSX.utils.aoa_to_sheet(sheetInfo.data);
                    XLSX.utils.book_append_sheet(wb, ws, sanitizeSheetName(sheetInfo.title));
                });

                const date = new Date().toISOString().split('T')[0];
                XLSX.writeFile(wb, `Relatorio_Contrato_${date}.xlsx`);
                showAlert('Arquivo Excel exportado com sucesso!');
            } catch (error) {
                console.error('Erro ao exportar para Excel:', error);
                showAlert(`Erro ao exportar para Excel: ${error.message}`, 'error');
            }
        });
    };

    // Inicialização
    loadFromLocalStorage();
    setupEventListeners();
});
