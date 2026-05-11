// =============================================
// Dashboard Pernoite Boa Vista - AVSI Brasil
// script.js
// =============================================

let charts = {};
const CAPACIDADE_OPERACIONAL = 400;
const DIAS_BASE_ATIVA = 5;

// Registrar plugin datalabels globalmente
Chart.register(ChartDataLabels);

const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const errorMessage = document.getElementById('errorMessage');

uploadArea.addEventListener('click', () => fileInput.click());

uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('dragover');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file) processFile(file);
});

fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) processFile(file);
});

function showError(message) {
    errorMessage.textContent = '⚠️ ' + message;
    errorMessage.classList.add('show');
    setTimeout(() => {
        errorMessage.classList.remove('show');
    }, 5000);
}

function processFile(file) {
    document.getElementById('loadingIndicator').classList.remove('hidden');
    document.getElementById('dashboard').classList.add('hidden');
    document.getElementById('statusBadge').textContent = 'Processando...';

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});

            console.log('Abas encontradas:', workbook.SheetNames);

            const requiredSheets = ['Individuos', 'Dias_F1', 'Dias_F4'];
            const missingSheets = requiredSheets.filter(sheet => !workbook.SheetNames.includes(sheet));

            if (missingSheets.length > 0) {
                throw new Error(`Abas não encontradas: ${missingSheets.join(', ')}`);
            }

            // Salvar workbook globalmente para reprocessamento
            workbookGlobal = workbook;

            const processedData = extractData(workbook);
            renderDashboard(processedData);

            document.getElementById('loadingIndicator').classList.add('hidden');
            document.getElementById('dashboard').classList.remove('hidden');
            document.getElementById('statusBadge').textContent = '✅ Dados carregados';
            document.getElementById('statusBadge').style.background = 'linear-gradient(135deg, #48bb78 0%, #38a169 100%)';
        } catch (error) {
            console.error('Erro ao processar arquivo:', error);
            showError('Erro ao processar o arquivo: ' + error.message);
            document.getElementById('loadingIndicator').classList.add('hidden');
            document.getElementById('statusBadge').textContent = '❌ Erro';
            document.getElementById('statusBadge').style.background = 'linear-gradient(135deg, #f56565 0%, #e53e3e 100%)';
        }
    };
    reader.onerror = function() {
        showError('Erro ao ler o arquivo. Tente novamente.');
        document.getElementById('loadingIndicator').classList.add('hidden');
    };
    reader.readAsArrayBuffer(file);
}

function extractData(workbook) {
    const mainSheetName = workbook.SheetNames[0];
    const mainSheet = XLSX.utils.sheet_to_json(workbook.Sheets[mainSheetName]);
    const individuosSheet = XLSX.utils.sheet_to_json(workbook.Sheets['Individuos']);
    const diasF1Sheet = XLSX.utils.sheet_to_json(workbook.Sheets['Dias_F1']);
    const diasF4Sheet = XLSX.utils.sheet_to_json(workbook.Sheets['Dias_F4']);

    console.log('Dados carregados:', {
        main: mainSheet.length,
        individuos: individuosSheet.length,
        diasF1: diasF1Sheet.length,
        diasF4: diasF4Sheet.length
    });

    // Detectar nome da coluna de nacionalidade (pode variar)
    const colunasPossiveis = Object.keys(individuosSheet[0] || {});
    const colunaNacionalidade = colunasPossiveis.find(col =>
        col.toLowerCase().includes('nacional') ||
        col.toLowerCase().includes('país') ||
        col.toLowerCase().includes('pais')
    ) || '__field_7__';

    console.log('Coluna de nacionalidade detectada:', colunaNacionalidade);

    const todasDatas = [
        ...diasF1Sheet.map(d => parseExcelDate(d.Entrada)),
        ...diasF4Sheet.map(d => parseExcelDate(d.Entrada))
    ].filter(d => d !== null);

    // dataReferencia: usa Data Final do filtro de período (se definida), ou última data do Excel
    let dataReferencia;
    const { inicio: filtroInicio, fim: filtroFim } = getFiltrosPeriodo();

    if (filtroFim) {
        dataReferencia = new Date(filtroFim);
        console.log('📆 Usando Data Final do filtro de período:', filtroFim);
    } else {
        dataReferencia = todasDatas.length > 0
            ? new Date(Math.max(...todasDatas.map(d => d.getTime())))
            : new Date();
        console.log('✅ Usando última data do Excel (recomendado)');
    }

    dataReferencia.setHours(0, 0, 0, 0);
    console.log('📅 Data de Referência:', dataReferencia.toLocaleDateString('pt-BR'));

    const inicioMes = new Date(dataReferencia.getFullYear(), dataReferencia.getMonth(), 1);
    const fimMes = new Date(dataReferencia.getFullYear(), dataReferencia.getMonth() + 1, 0);

    const janelaDias = 10;
    const limite10Dias = new Date(dataReferencia.getTime() - (janelaDias * 24 * 60 * 60 * 1000));

    console.log('📊 Parâmetros:');
    console.log('   Dias base ativa:', DIAS_BASE_ATIVA);
    console.log('   Janela média recente:', janelaDias, 'dias');
    console.log('   Capacidade:', CAPACIDADE_OPERACIONAL);
    console.log('   Mês de referência:', dataReferencia.toLocaleDateString('pt-BR', {month: 'long', year: 'numeric'}));
    console.log('📊 Limite 10 dias (média):', limite10Dias.toLocaleDateString('pt-BR'));
    console.log('📊 Regra base ativa: gap < ' + DIAS_BASE_ATIVA + ' dias');

    const todasEntradas = [
        ...diasF1Sheet.map(d => ({
            ...d,
            tipo: 'Pernoite',
            data: parseExcelDate(d.Entrada)
        })),
        ...diasF4Sheet.map(d => ({
            ...d,
            tipo: 'Alimentação',
            data: parseExcelDate(d.Entrada)
        }))
    ].filter(e => {
        if (!e.data) return false;
        if (e.data > dataReferencia) return false;

        // Filtro de período — usa variáveis já resolvidas acima
        // Se inicio > fim, filtro inválido — ignorar período
        if (filtroInicio && filtroFim && filtroInicio > filtroFim) return true;

        if (filtroInicio && e.data < filtroInicio) return false;
        if (filtroFim    && e.data > filtroFim)    return false;

        return true;
    });

    console.log('Total de entradas processadas:', todasEntradas.length);

    const entradasPorPessoa = {};
    todasEntradas.forEach(entrada => {
        const pessoaId = entrada._parent_index;
        if (!entradasPorPessoa[pessoaId]) {
            entradasPorPessoa[pessoaId] = [];
        }
        entradasPorPessoa[pessoaId].push(entrada);
    });

    const baseAtiva = new Set();
    const ativosNoMes = new Set();
    const entradasNoMes = new Set();

    Object.entries(entradasPorPessoa).forEach(([pessoaId, entradas]) => {
        const ultimaEntrada = entradas.sort((a, b) => b.data - a.data)[0];

        const gapMs = dataReferencia.getTime() - ultimaEntrada.data.getTime();
        const gapDias = Math.floor(gapMs / (24 * 60 * 60 * 1000));

        if (gapDias < DIAS_BASE_ATIVA) {
            baseAtiva.add(parseInt(pessoaId));
        }

        const temEntradaNoMes = entradas.some(e => e.data >= inicioMes);
        if (temEntradaNoMes) {
            ativosNoMes.add(parseInt(pessoaId));

            const primeiraEntrada = entradas.sort((a, b) => a.data - b.data)[0];
            if (primeiraEntrada.data >= inicioMes) {
                entradasNoMes.add(parseInt(pessoaId));
            }
        }
    });

    console.log('Base ativa calculada:', baseAtiva.size);

    // DEBUG: Distribuição por GAP
    console.log('\n' + '='.repeat(80));
    console.log('🔍 DEBUG: DISTRIBUIÇÃO POR GAP (dias sem vir)');
    console.log('='.repeat(80));

    const gapDistribution = {};
    const pessoasPorGap = {};

    Object.entries(entradasPorPessoa).forEach(([pessoaId, entradas]) => {
        const ultimaEntrada = entradas.sort((a, b) => b.data - a.data)[0];
        const gapMs = dataReferencia.getTime() - ultimaEntrada.data.getTime();
        const gapDias = Math.floor(gapMs / (24 * 60 * 60 * 1000));

        if (gapDias >= 0 && gapDias <= 15) {
            gapDistribution[gapDias] = (gapDistribution[gapDias] || 0) + 1;
            if (!pessoasPorGap[gapDias]) pessoasPorGap[gapDias] = [];
            pessoasPorGap[gapDias].push({
                id: pessoaId,
                ultima: ultimaEntrada.data.toISOString().split('T')[0],
                gap: gapDias
            });
        }
    });

    console.log('Gap | Pessoas | Status na Lógica (gap < ' + DIAS_BASE_ATIVA + ')');
    console.log('-'.repeat(60));

    for (let gap = 0; gap <= 15; gap++) {
        const count = gapDistribution[gap] || 0;
        const status = gap < DIAS_BASE_ATIVA ? '✅ ATIVO' : '❌ INATIVO';
        if (count > 0) {
            console.log(`${gap.toString().padStart(3)} | ${count.toString().padStart(7)} | ${status}`);
        }
    }

    console.log('\n📊 RESUMO:');
    console.log(`   Total pessoas únicas: ${Object.keys(entradasPorPessoa).length}`);
    console.log(`   Base ativa (gap < ${DIAS_BASE_ATIVA}): ${baseAtiva.size}`);
    console.log(`   Data referência: ${dataReferencia.toISOString().split('T')[0]}`);
    console.log(`   Dias base ativa: ${DIAS_BASE_ATIVA}`);
    console.log('='.repeat(80) + '\n');

    const individuosAtivos = individuosSheet.filter(ind =>
        baseAtiva.has(ind._parent_index)
    );

    const registrosAtivos = mainSheet.filter(reg =>
        baseAtiva.has(reg._index)
    );

    // Criar mapa: _parent_index → quantidade de pessoas
    const pessoasPorFamilia = {};
    individuosSheet.forEach(ind => {
        const parent = ind._parent_index;
        pessoasPorFamilia[parent] = (pessoasPorFamilia[parent] || 0) + 1;
    });

    // Contar pessoas por dia
    const entradasPorDia = {};
    todasEntradas.forEach(entrada => {
        const diaKey = entrada.data.toISOString().split('T')[0];
        if (!entradasPorDia[diaKey]) {
            entradasPorDia[diaKey] = { familias: new Set(), pessoas: 0 };
        }

        const parent = entrada._parent_index;
        if (!entradasPorDia[diaKey].familias.has(parent)) {
            entradasPorDia[diaKey].familias.add(parent);
            entradasPorDia[diaKey].pessoas += (pessoasPorFamilia[parent] || 1);
        }
    });

    const diasComDados = Object.values(entradasPorDia);
    const mediaDiaria = diasComDados.length > 0
        ? diasComDados.reduce((sum, obj) => sum + obj.pessoas, 0) / diasComDados.length
        : 0;

    const entradas10Dias = todasEntradas.filter(e => e.data >= limite10Dias);
    const diasUnicos10 = {};
    entradas10Dias.forEach(entrada => {
        const diaKey = entrada.data.toISOString().split('T')[0];
        if (!diasUnicos10[diaKey]) {
            diasUnicos10[diaKey] = { familias: new Set(), pessoas: 0 };
        }
        const parent = entrada._parent_index;
        if (!diasUnicos10[diaKey].familias.has(parent)) {
            diasUnicos10[diaKey].familias.add(parent);
            diasUnicos10[diaKey].pessoas += (pessoasPorFamilia[parent] || 1);
        }
    });

    const media10Dias = Object.keys(diasUnicos10).length > 0
        ? Object.values(diasUnicos10).reduce((sum, obj) => sum + obj.pessoas, 0) / Object.keys(diasUnicos10).length
        : 0;

    const utilizacaoEspaco = {};
    registrosAtivos.forEach(reg => {
        const util = reg['Utilização do espaço'] || 'Não especificado';
        utilizacaoEspaco[util] = (utilizacaoEspaco[util] || 0) + 1;
    });

    const nacionalidadeData = {};
    individuosAtivos.forEach(ind => {
        const nac = ind[colunaNacionalidade] || 'Não informado';
        nacionalidadeData[nac] = (nacionalidadeData[nac] || 0) + 1;
    });

    const venezuelanosAtivos = nacionalidadeData['Venezuela'] || nacionalidadeData['Venezuelana'] ||
                               nacionalidadeData['venezuelana'] || nacionalidadeData['venezuela'] || 0;

    const dataRefKey = dataReferencia.toISOString().split('T')[0];
    const pessoasDiaReferencia = entradasPorDia[dataRefKey] ? entradasPorDia[dataRefKey].pessoas : 0;

    // Tempo de permanência
    const tempoPermanencia = {};
    Object.keys(entradasPorPessoa).forEach(pessoaId => {
        const entradas = entradasPorPessoa[pessoaId];
        if (entradas.length === 0) return;

        const primeiraEntrada = new Date(Math.min(...entradas.map(e => e.data.getTime())));
        const ultimaEntrada = new Date(Math.max(...entradas.map(e => e.data.getTime())));
        const diasPermanencia = Math.floor((ultimaEntrada - primeiraEntrada) / (24*60*60*1000));

        let faixa;
        if (diasPermanencia < 7) faixa = '< 7 dias';
        else if (diasPermanencia < 30) faixa = '7-30 dias';
        else if (diasPermanencia < 60) faixa = '30-60 dias';
        else if (diasPermanencia < 90) faixa = '60-90 dias';
        else faixa = '> 90 dias';

        tempoPermanencia[faixa] = (tempoPermanencia[faixa] || 0) + 1;
    });

    // Evolução da base ativa (6 meses)
    const evolucaoMensal = {};
    for (let i = 5; i >= 0; i--) {
        const mes = new Date(dataReferencia.getFullYear(), dataReferencia.getMonth() - i, 1);
        const mesKey = mes.toISOString().slice(0, 7);
        evolucaoMensal[mesKey] = 0;
    }

    Object.keys(evolucaoMensal).forEach(mesKey => {
        const [ano, mes] = mesKey.split('-').map(Number);
        const fimMesLocal = new Date(ano, mes, 0);
        fimMesLocal.setHours(0, 0, 0, 0);

        const limite5DiasMes = new Date(fimMesLocal.getTime() - (DIAS_BASE_ATIVA * 24*60*60*1000));

        let ativosMes = 0;
        Object.keys(entradasPorPessoa).forEach(pessoaId => {
            const entradas = entradasPorPessoa[pessoaId];
            if (entradas.length === 0) return;

            const ultimaEntradaMes = new Date(Math.max(...entradas
                .filter(e => e.data <= fimMesLocal)
                .map(e => e.data.getTime())));

            if (ultimaEntradaMes >= limite5DiasMes && ultimaEntradaMes <= fimMesLocal) {
                ativosMes++;
            }
        });

        evolucaoMensal[mesKey] = ativosMes;
    });

    // Crianças ativas (0-17 anos)
    const criancasAtivas = individuosAtivos.filter(ind => {
        const idade = ind.idade_anos || 0;
        return idade >= 0 && idade <= 17;
    }).length;

    return {
        dataReferencia,
        totalHistorico: individuosSheet.length,
        baseAtiva: individuosAtivos.length,
        baseFamilias: baseAtiva.size,
        refDay: pessoasDiaReferencia,
        ativosNoMes: ativosNoMes.size,
        entradasNoMes: entradasNoMes.size,
        mediaDiaria: Math.round(mediaDiaria),
        media10Dias: Math.round(media10Dias),
        individuosAtivos,
        registrosAtivos,
        mainSheet,
        utilizacaoEspaco,
        todasEntradas,
        entradasPorDia,
        nacionalidadeData,
        venezuelanosAtivos,
        tempoPermanencia,
        evolucaoMensal,
        criancasAtivas
    };
}

function parseExcelDate(excelDate) {
    if (!excelDate) return null;

    if (typeof excelDate === 'number') {
        const date = new Date((excelDate - 25569) * 86400 * 1000);
        return date;
    }

    const date = new Date(excelDate);
    return isNaN(date.getTime()) ? null : date;
}

let workbookGlobal = null;

function destacarDataManual() {
    const input = document.getElementById('dataManual');
    if (input.value && input.value.trim() !== '') {
        input.style.border = '2px solid #f6ad55';
        input.style.background = 'rgba(246, 173, 85, 0.1)';
    } else {
        input.style.border = '2px solid #cbd5e0';
        input.style.background = 'white';
    }
}

// =============================================
// FILTRO DE PERÍODO
// =============================================

function getFiltrosPeriodo() {
    const inicioVal = document.getElementById('filtroDataInicio')?.value;
    const fimVal    = document.getElementById('filtroDataFim')?.value;

    const inicio = inicioVal ? new Date(inicioVal + 'T00:00:00') : null;
    const fim    = fimVal    ? new Date(fimVal    + 'T23:59:59') : null;

    return { inicio, fim };
}

function onFiltroDataChange() {
    const { inicio, fim } = getFiltrosPeriodo();
    const infoEl = document.getElementById('filtroPeriodoInfo');

    // Highlight dos campos
    const elInicio = document.getElementById('filtroDataInicio');
    const elFim    = document.getElementById('filtroDataFim');

    [elInicio, elFim].forEach(el => {
        if (el.value) {
            el.style.border = '2px solid #f6ad55';
            el.style.background = 'rgba(246, 173, 85, 0.08)';
        } else {
            el.style.border = '2px solid #cbd5e0';
            el.style.background = 'white';
        }
    });

    // Validação: início > fim
    if (inicio && fim && inicio > fim) {
        infoEl.innerHTML = '⚠️ <strong style="color:#e53e3e">Data inicial não pode ser maior que a data final.</strong>';
        return;
    }

    // Mensagem informativa com novo design
    if (!inicio && !fim) {
        infoEl.innerHTML = '';
    } else if (inicio && !fim) {
        infoEl.innerHTML = `<span style="display:inline-flex;align-items:center;gap:8px;background:rgba(0,104,71,0.08);border:1px solid rgba(0,104,71,0.2);padding:6px 14px;border-radius:8px;font-size:13px;font-weight:600;color:#006847;">
            📆 Exibindo a partir de <strong>${inicio.toLocaleDateString('pt-BR')}</strong>
        </span>`;
    } else if (!inicio && fim) {
        infoEl.innerHTML = `<span style="display:inline-flex;align-items:center;gap:8px;background:rgba(0,104,71,0.08);border:1px solid rgba(0,104,71,0.2);padding:6px 14px;border-radius:8px;font-size:13px;font-weight:600;color:#006847;">
            📆 Exibindo até <strong>${fim.toLocaleDateString('pt-BR')}</strong>
        </span>`;
    } else {
        infoEl.innerHTML = `<span style="display:inline-flex;align-items:center;gap:8px;background:rgba(0,104,71,0.08);border:1px solid rgba(0,104,71,0.2);padding:6px 14px;border-radius:8px;font-size:13px;font-weight:600;color:#006847;">
            📆 Período ativo: <strong>${inicio.toLocaleDateString('pt-BR')}</strong> → <strong>${fim.toLocaleDateString('pt-BR')}</strong>
        </span>`;
    }

    // Reprocessar automaticamente se houver dados carregados
    if (workbookGlobal) {
        const processedData = extractData(workbookGlobal);
        renderDashboard(processedData);
    }
}

function limparFiltrosPeriodo() {
    const elInicio = document.getElementById('filtroDataInicio');
    const elFim    = document.getElementById('filtroDataFim');

    elInicio.value = '';
    elFim.value    = '';

    [elInicio, elFim].forEach(el => {
        el.style.border = '2px solid #cbd5e0';
        el.style.background = 'white';
    });

    document.getElementById('filtroPeriodoInfo').textContent = '';

    if (workbookGlobal) {
        const processedData = extractData(workbookGlobal);
        renderDashboard(processedData);
    }
}

function renderDashboard(data) {
    renderStats(data);
    renderCharts(data);
    renderInsights(data);
    updateLastUpdated(data.dataReferencia);
    rankingData = data;
}

function reprocessarDados() {
    if (!workbookGlobal) {
        alert('⚠️ Faça upload de um arquivo Excel primeiro!');
        return;
    }

    console.log('🔄 Reprocessando com novos parâmetros...');
    const processedData = extractData(workbookGlobal);
    renderDashboard(processedData);
    console.log('✅ Dados reprocessados com sucesso!');
}

function renderStats(data) {
    const statsGrid = document.getElementById('statsGrid');
    const capacidadePercent = (data.baseAtiva / CAPACIDADE_OPERACIONAL * 100).toFixed(1);

    let capacidadeClass = 'positive';
    if (capacidadePercent > 80) capacidadeClass = 'warning';
    if (capacidadePercent > 95) capacidadeClass = 'critical';

    const percVenezuelanos = (data.venezuelanosAtivos / data.baseAtiva * 100).toFixed(1);

    const stats = [
        {
            icon: '📊',
            label: 'Total Histórico',
            value: data.totalHistorico.toLocaleString('pt-BR'),
            detail: 'Pessoas registradas no sistema',
        },
        {
            icon: '👥',
            label: 'Base Ativa',
            value: data.baseAtiva.toLocaleString('pt-BR'),
            detail: `Não ficaram ${DIAS_BASE_ATIVA} dias completos sem vir`,
            trend: { text: `${capacidadePercent}% ocupação`, class: capacidadeClass },
            capacity: capacidadePercent
        },
        {
            icon: '🇻🇪',
            label: 'Venezuelanos Ativos',
            value: data.venezuelanosAtivos.toLocaleString('pt-BR'),
            detail: 'Elegíveis para Operação Acolhida',
            trend: { text: `${percVenezuelanos}% da base`, class: 'warning' }
        },
        {
            icon: '👨‍👩‍👧‍👦',
            label: 'Famílias Ativas',
            value: data.baseFamilias.toLocaleString('pt-BR'),
            detail: 'Núcleos familiares atendidos',
            trend: { text: `${(data.baseAtiva / data.baseFamilias).toFixed(1)} pessoas/família`, class: 'neutral' }
        },
        {
            icon: '👶',
            label: 'Crianças Ativas',
            value: data.criancasAtivas.toLocaleString('pt-BR'),
            detail: 'Crianças e adolescentes (0-17 anos)',
            trend: { text: `${(data.criancasAtivas / data.baseAtiva * 100).toFixed(1)}% da base`, class: 'warning' }
        },
        {
            icon: '📅',
            label: 'Ativos no Mês',
            value: data.ativosNoMes.toLocaleString('pt-BR'),
            detail: 'Pessoas com presença no mês atual',
        },
        {
            icon: '🆕',
            label: 'Novos no Mês',
            value: data.entradasNoMes.toLocaleString('pt-BR'),
            detail: 'Primeira entrada no mês',
        },
        {
            icon: '📈',
            label: 'Média Diária',
            value: data.mediaDiaria.toLocaleString('pt-BR'),
            detail: 'Pessoas por dia (período completo)',
        },
        {
            icon: '⏱️',
            label: 'Média 10 Dias',
            value: data.media10Dias.toLocaleString('pt-BR'),
            detail: 'Pessoas por dia (últimos 10 dias)',
        },
        {
            icon: '🏠',
            label: 'Capacidade',
            value: CAPACIDADE_OPERACIONAL.toLocaleString('pt-BR'),
            detail: 'Capacidade operacional máxima'
        },
        {
            icon: '🛏️',
            label: 'Pernoite',
            value: (data.utilizacaoEspaco['Pernoite'] || 0).toLocaleString('pt-BR'),
            detail: `${((data.utilizacaoEspaco['Pernoite'] || 0) / data.baseAtiva * 100).toFixed(1)}% da base ativa`,
        },
        {
            icon: '🍽️',
            label: 'Alimentação',
            value: (data.utilizacaoEspaco['Alimentação'] || 0).toLocaleString('pt-BR'),
            detail: `${((data.utilizacaoEspaco['Alimentação'] || 0) / data.baseAtiva * 100).toFixed(1)}% da base ativa`,
        }
    ];

    statsGrid.innerHTML = stats.map(stat => `
        <div class="stat-card">
            <div class="stat-icon">${stat.icon}</div>
            <div class="stat-label">${stat.label}</div>
            <div class="stat-value">${stat.value}</div>
            <div class="stat-detail">${stat.detail}</div>
            ${stat.trend ? `<span class="stat-trend ${stat.trend.class}">${stat.trend.text}</span>` : ''}
            ${stat.capacity ? `
                <div class="capacity-indicator">
                    <div class="capacity-fill ${capacidadeClass}" style="width: ${Math.min(stat.capacity, 100)}%"></div>
                </div>
            ` : ''}
        </div>
    `).join('');
}

function renderCharts(data) {
    const chartsGrid = document.getElementById('chartsGrid');

    const sexoData = data.individuosAtivos.reduce((acc, ind) => {
        acc[ind.Sexo] = (acc[ind.Sexo] || 0) + 1;
        return acc;
    }, {});

    const faixasEtarias = {
        '0-12': 0, '13-17': 0, '18-30': 0, '31-45': 0, '46-60': 0, '60+': 0
    };
    data.individuosAtivos.forEach(ind => {
        const idade = ind.idade_anos;
        if (idade <= 12) faixasEtarias['0-12']++;
        else if (idade <= 17) faixasEtarias['13-17']++;
        else if (idade <= 30) faixasEtarias['18-30']++;
        else if (idade <= 45) faixasEtarias['31-45']++;
        else if (idade <= 60) faixasEtarias['46-60']++;
        else faixasEtarias['60+']++;
    });

    const perfisData = data.registrosAtivos.reduce((acc, reg) => {
        const perfil = reg.Perfil || 'Não especificado';
        acc[perfil] = (acc[perfil] || 0) + 1;
        return acc;
    }, {});
    const topPerfis = Object.entries(perfisData).sort((a, b) => b[1] - a[1]).slice(0, 6);

    const dataRef = data.dataReferencia || new Date();
    const ultimos30Dias = Array.from({length: 30}, (_, i) => {
        const ms = dataRef.getTime() - ((29 - i) * 24 * 60 * 60 * 1000);
        const d = new Date(ms);
        return d.toISOString().split('T')[0];
    });
    const ocupacaoPorDia = ultimos30Dias.map(dia => {
        return data.entradasPorDia[dia] ? data.entradasPorDia[dia].pessoas : 0;
    });

    const topNacionalidades = Object.entries(data.nacionalidadeData)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 8);

    chartsGrid.innerHTML = `
        <div class="chart-card">
            <h3>🚻 Distribuição por Sexo</h3>
            <div class="chart-container"><canvas id="sexoChart"></canvas></div>
        </div>
        <div class="chart-card">
            <h3>🎂 Distribuição por Faixa Etária</h3>
            <div class="chart-container"><canvas id="idadeChart"></canvas></div>
        </div>
        <div class="chart-card">
            <h3>🌎 Distribuição por Nacionalidade</h3>
            <div class="chart-container"><canvas id="nacionalidadeChart"></canvas></div>
        </div>
        <div class="chart-card">
            <h3>👤 Top 6 Perfis mais Comuns</h3>
            <div class="chart-container"><canvas id="perfisChart"></canvas></div>
        </div>
        <div class="chart-card">
            <h3>⏱️ Tempo de Permanência</h3>
            <div class="chart-container"><canvas id="permanenciaChart"></canvas></div>
        </div>
        <div class="chart-card">
            <h3>📈 Evolução da Base Ativa (6 meses)</h3>
            <div class="chart-container"><canvas id="evolucaoChart"></canvas></div>
        </div>
        <div class="chart-card" style="grid-column: 1 / -1;">
            <h3>📊 Tendência de Ocupação (30 dias)</h3>
            <div class="chart-container"><canvas id="tendenciaChart"></canvas></div>
        </div>
    `;

    Object.values(charts).forEach(chart => chart.destroy());
    charts = {};

    const datalabelsConfig = {
        color: '#1a202c',
        font: { weight: 'bold', size: 14 },
        formatter: (value) => value.toLocaleString('pt-BR')
    };

    charts.sexo = new Chart(document.getElementById('sexoChart'), {
        type: 'doughnut',
        data: {
            labels: Object.keys(sexoData),
            datasets: [{
                data: Object.values(sexoData),
                backgroundColor: ['#006847', '#9ACD32', '#FFD700', '#00a854'],
                borderWidth: 0
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom', labels: { padding: 20, font: { size: 13, weight: '600' } } },
                datalabels: {
                    color: '#fff',
                    font: { weight: 'bold', size: 16 },
                    formatter: (value, context) => {
                        const total = context.dataset.data.reduce((a, b) => a + b, 0);
                        const percent = (value / total * 100).toFixed(1);
                        return `${value}\n(${percent}%)`;
                    }
                }
            }
        }
    });

    charts.idade = new Chart(document.getElementById('idadeChart'), {
        type: 'bar',
        data: {
            labels: Object.keys(faixasEtarias),
            datasets: [{
                label: 'Pessoas',
                data: Object.values(faixasEtarias),
                backgroundColor: '#764ba2',
                borderRadius: 8,
                borderSkipped: false,
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                datalabels: { anchor: 'end', align: 'top', ...datalabelsConfig }
            },
            scales: {
                y: { beginAtZero: true, grid: { color: '#f1f5f9' } },
                x: { grid: { display: false } }
            }
        }
    });

    const nacionalidadeCores = topNacionalidades.map(([pais]) => {
        const paisLower = pais.toLowerCase();
        if (paisLower.includes('venezuela')) return '#f59e0b';
        return '#667eea';
    });

    charts.nacionalidade = new Chart(document.getElementById('nacionalidadeChart'), {
        type: 'bar',
        data: {
            labels: topNacionalidades.map(p => p[0].length > 20 ? p[0].substring(0, 20) + '...' : p[0]),
            datasets: [{
                label: 'Quantidade',
                data: topNacionalidades.map(p => p[1]),
                backgroundColor: nacionalidadeCores,
                borderRadius: 8,
                borderSkipped: false,
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                datalabels: { anchor: 'end', align: 'right', ...datalabelsConfig }
            },
            scales: {
                x: { beginAtZero: true, grid: { color: '#f1f5f9' } },
                y: { grid: { display: false } }
            }
        }
    });

    charts.perfis = new Chart(document.getElementById('perfisChart'), {
        type: 'bar',
        data: {
            labels: topPerfis.map(p => {
                const label = p[0];
                return label.length > 25 ? label.substring(0, 25) + '...' : label;
            }),
            datasets: [{
                label: 'Quantidade',
                data: topPerfis.map(p => p[1]),
                backgroundColor: '#006847',
                borderRadius: 8,
                borderSkipped: false,
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                datalabels: { anchor: 'end', align: 'right', ...datalabelsConfig }
            },
            scales: {
                x: { beginAtZero: true, grid: { color: '#f1f5f9' } },
                y: { grid: { display: false } }
            }
        }
    });

    charts.tendencia = new Chart(document.getElementById('tendenciaChart'), {
        type: 'line',
        data: {
            labels: ultimos30Dias.map(d => {
                const [y, m, dia] = d.split('-');
                return `${dia}/${m}`;
            }),
            datasets: [{
                label: 'Pessoas/dia',
                data: ocupacaoPorDia,
                borderColor: '#006847',
                backgroundColor: 'rgba(0, 104, 71, 0.1)',
                fill: true,
                tension: 0.4,
                borderWidth: 3,
                pointRadius: 4,
                pointBackgroundColor: '#006847',
                pointBorderColor: '#fff',
                pointBorderWidth: 2,
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                datalabels: {
                    display: true,
                    align: 'top',
                    anchor: 'end',
                    color: '#1a202c',
                    font: { weight: 'bold', size: 11 },
                    formatter: (value) => value > 0 ? value : ''
                }
            },
            scales: {
                y: { beginAtZero: true, grid: { color: '#f1f5f9' } },
                x: { grid: { display: false } }
            }
        }
    });

    const faixasPermanencia = ['< 7 dias', '7-30 dias', '30-60 dias', '60-90 dias', '> 90 dias'];
    const valoresPermanencia = faixasPermanencia.map(f => data.tempoPermanencia[f] || 0);

    charts.permanencia = new Chart(document.getElementById('permanenciaChart'), {
        type: 'bar',
        data: {
            labels: faixasPermanencia,
            datasets: [{
                label: 'Pessoas',
                data: valoresPermanencia,
                backgroundColor: ['#9ACD32', '#7fb83e', '#006847', '#00704a', '#004d33'],
                borderRadius: 8,
                borderSkipped: false,
            }]
        },
        plugins: [ChartDataLabels],
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                datalabels: {
                    display: true,
                    align: 'end',
                    anchor: 'end',
                    color: '#1a202c',
                    font: { weight: 'bold', size: 11 },
                    formatter: (value) => value > 0 ? value : ''
                }
            },
            scales: {
                y: { beginAtZero: true, grid: { color: '#f1f5f9' } },
                x: { grid: { display: false } }
            }
        }
    });

    const mesesLabels = Object.keys(data.evolucaoMensal).map(mes => {
        const [ano, m] = mes.split('-');
        const mesesNomes = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
        return mesesNomes[parseInt(m) - 1] + '/' + ano.slice(2);
    });
    const valoresEvolucao = Object.values(data.evolucaoMensal);

    charts.evolucao = new Chart(document.getElementById('evolucaoChart'), {
        type: 'line',
        data: {
            labels: mesesLabels,
            datasets: [{
                label: 'Base Ativa',
                data: valoresEvolucao,
                borderColor: '#FFD700',
                backgroundColor: 'rgba(255, 215, 0, 0.1)',
                fill: true,
                tension: 0.4,
                borderWidth: 3,
                pointRadius: 5,
                pointBackgroundColor: '#FFD700',
                pointBorderColor: '#fff',
                pointBorderWidth: 2,
            }]
        },
        plugins: [ChartDataLabels],
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false },
                datalabels: {
                    display: true,
                    align: 'top',
                    anchor: 'end',
                    color: '#1a202c',
                    font: { weight: 'bold', size: 11 },
                    formatter: (value) => value > 0 ? value : ''
                }
            },
            scales: {
                y: { beginAtZero: true, grid: { color: '#f1f5f9' } },
                x: { grid: { display: false } }
            }
        }
    });
}

// ===== FUNÇÕES DO RANKING =====
let rankingData = null;

function toggleRanking() {
    const container = document.getElementById('rankingContainer');
    const btn = document.querySelector('.ranking-toggle-btn');

    if (container.classList.contains('show')) {
        container.classList.remove('show');
        btn.textContent = '📈 Ver Ranking 50+ Ativos';
    } else {
        container.classList.add('show');
        btn.textContent = '❌ Fechar Ranking';
        if (rankingData) {
            renderRanking(rankingData);
        }
    }
}

function calculateRanking(data) {
    const pessoasRanking = [];

    const entradasPorPessoa = {};
    data.todasEntradas.forEach(entrada => {
        const pessoaId = entrada._parent_index;
        if (!entradasPorPessoa[pessoaId]) {
            entradasPorPessoa[pessoaId] = [];
        }
        entradasPorPessoa[pessoaId].push(entrada);
    });

    data.individuosAtivos.forEach(individuo => {
        const pessoaId = individuo._parent_index;
        const entradas = entradasPorPessoa[pessoaId] || [];

        const diasUnicos = new Set();
        entradas.forEach(e => {
            const diaKey = e.data.toISOString().split('T')[0];
            diasUnicos.add(diaKey);
        });

        pessoasRanking.push({
            nome: individuo['Nome completo'] || 'Não informado',
            grupo: individuo['Relação'] || individuo['Perfil'] || 'N/A',
            idade: individuo.idade_anos || 0,
            nacionalidade: individuo.Nacionalidade || 'Não informada',
            diasAtivo: diasUnicos.size
        });
    });

    pessoasRanking.sort((a, b) => b.diasAtivo - a.diasAtivo);
    return pessoasRanking.slice(0, 50);
}

function renderRanking(data) {
    const rankingBody = document.getElementById('rankingBody');
    const top50 = calculateRanking(data);

    rankingBody.innerHTML = top50.map((pessoa, index) => {
        const rank = index + 1;
        const isTop3 = rank <= 3;
        const rankClass = isTop3 ? 'rank-badge top3' : 'rank-badge';

        let grupoClass = '';
        const grupoUpper = pessoa.grupo.toUpperCase();
        if (grupoUpper.includes('HD') || grupoUpper.includes('HOMEM')) grupoClass = 'HD';
        else if (grupoUpper.includes('MD') || grupoUpper.includes('MULHER')) grupoClass = 'MD';
        else if (grupoUpper.includes('FMC') || grupoUpper.includes('FILHO') || grupoUpper.includes('CRIAN')) grupoClass = 'FMC';

        return `
            <tr>
                <td><span class="${rankClass}">${rank}º</span></td>
                <td style="font-weight: 600; color: #1a202c;">${pessoa.nome}</td>
                <td><span class="grupo-badge ${grupoClass}">${pessoa.grupo}</span></td>
                <td>${pessoa.idade} anos</td>
                <td>${pessoa.nacionalidade}</td>
                <td style="font-weight: 700; color: #006847;">${pessoa.diasAtivo} dias</td>
            </tr>
        `;
    }).join('');

    console.log('🏆 Ranking Top 50 gerado:', top50.length, 'pessoas');
}

function renderInsights(data) {
    const insightsSection = document.getElementById('insightsSection');

    const idadeMedia = data.individuosAtivos.reduce((sum, ind) => sum + (ind.idade_anos || 0), 0) / data.individuosAtivos.length;

    const faixasCount = {
        '0-12': 0, '13-17': 0, '18-30': 0, '31-45': 0, '46-60': 0, '60+': 0
    };
    data.individuosAtivos.forEach(ind => {
        const idade = ind.idade_anos;
        if (idade <= 12) faixasCount['0-12']++;
        else if (idade <= 17) faixasCount['13-17']++;
        else if (idade <= 30) faixasCount['18-30']++;
        else if (idade <= 45) faixasCount['31-45']++;
        else if (idade <= 60) faixasCount['46-60']++;
        else faixasCount['60+']++;
    });
    const faixaPredominante = Object.entries(faixasCount).sort((a, b) => b[1] - a[1])[0];

    const homensDesacomp = data.registrosAtivos.filter(r => r.Perfil === 'HD - Homem Desacompanhado').length;
    const mulheresDesacomp = data.registrosAtivos.filter(r => r.Perfil === 'MD - Mulher Desacompanhada').length;
    const familiasCriancas = data.registrosAtivos.filter(r =>
        r.Perfil && (r.Perfil.includes('Criança') || r.Perfil.includes('FMC') || r.Perfil.includes('FPC'))
    ).length;

    const percHomensDesacomp = (homensDesacomp / data.baseAtiva * 100).toFixed(1);
    const percMulheresDesacomp = (mulheresDesacomp / data.baseAtiva * 100).toFixed(1);
    const percFamiliasCriancas = (familiasCriancas / data.baseAtiva * 100).toFixed(1);

    const taxaOcupacao = (data.baseAtiva / CAPACIDADE_OPERACIONAL * 100).toFixed(1);
    let statusCapacidade = 'normal';
    let messageCapacidade = 'dentro dos parâmetros normais';
    if (taxaOcupacao > 80) {
        statusCapacidade = 'alerta';
        messageCapacidade = 'próximo ao limite da capacidade';
    }
    if (taxaOcupacao > 95) {
        statusCapacidade = 'crítico';
        messageCapacidade = 'em capacidade crítica';
    }

    const pernoitePercent = ((data.utilizacaoEspaco['Pernoite'] || 0) / data.baseAtiva * 100).toFixed(1);
    const alimentacaoPercent = ((data.utilizacaoEspaco['Alimentação'] || 0) / data.baseAtiva * 100).toFixed(1);

    const percVenezuelanos = (data.venezuelanosAtivos / data.baseAtiva * 100).toFixed(1);
    const naoVenezuelanos = data.baseAtiva - data.venezuelanosAtivos;
    const percNaoVenezuelanos = (naoVenezuelanos / data.baseAtiva * 100).toFixed(1);

    insightsSection.innerHTML = `
        <h2>💡 Insights e Análises Estratégicas</h2>

        <div class="insight-item">
            <h4>🏠 Análise de Capacidade Operacional</h4>
            <p>O abrigo está operando com <strong>${taxaOcupacao}% de sua capacidade</strong> (${data.baseAtiva} de ${CAPACIDADE_OPERACIONAL} pessoas),
            ${messageCapacidade}. ${statusCapacidade === 'crítico' ? '⚠️ <strong>Atenção:</strong> considere medidas para redistribuição ou expansão temporária da capacidade.' :
            statusCapacidade === 'alerta' ? '⚠️ <strong>Monitoramento necessário:</strong> acompanhar de perto a entrada de novas pessoas.' : '✅ Capacidade adequada para operação normal.'}</p>
        </div>

        <div class="insight-item">
            <h4>🇻🇪 Análise de Elegibilidade para Abrigamento - Operação Acolhida</h4>
            <div class="venezuelan-highlight">
                <h5>🎯 INSIGHT CRÍTICO: Foco em Venezuelanos</h5>
                <p>Apenas <strong>cidadãos venezuelanos</strong> são elegíveis para abrigamento através da <strong>Operação Acolhida</strong>.
                Atualmente, <strong>${data.venezuelanosAtivos} pessoas (${percVenezuelanos}%)</strong> da base ativa são venezuelanas e
                <strong>elegíveis para encaminhamento ao programa de interiorização e abrigamento</strong>.</p>
            </div>
            <p style="margin-top: 15px;">Os <strong>${naoVenezuelanos} indivíduos (${percNaoVenezuelanos}%)</strong> de outras nacionalidades
            <strong>NÃO são elegíveis</strong> para a Operação Acolhida e necessitam de <strong>estratégias alternativas</strong> de apoio:
            parcerias com órgãos municipais, estaduais, organizações não-governamentais e programas de integração local.</p>
            <p style="margin-top: 10px;"><strong>📋 Recomendação:</strong> Priorizar mapeamento e encaminhamento ativo dos ${data.venezuelanosAtivos} venezuelanos
            para processos de documentação, interiorização e abrigamento permanente através da rede da Operação Acolhida.</p>
        </div>

        <div class="insight-item">
            <h4>👥 Perfil Demográfico e Etário</h4>
            <p>A idade média da base ativa é <strong>${idadeMedia.toFixed(1)} anos</strong>, com predominância da faixa etária
            <strong>${faixaPredominante[0]} anos</strong> (${faixaPredominante[1]} pessoas, <strong>${(faixaPredominante[1]/data.baseAtiva*100).toFixed(1)}%</strong>).
            Isso indica ${idadeMedia < 35 ? 'uma população relativamente jovem, demandando programas de capacitação profissional e inserção no mercado de trabalho' :
            'uma população com idade mais avançada, requerendo atenção especial à saúde e condições de acessibilidade'}.</p>
        </div>

        <div class="insight-item">
            <h4>👨‍👩‍👧‍👦 Composição e Perfis Familiares</h4>
            <div class="profile-grid">
                <div class="profile-item">
                    <div class="profile-label">👨 Homens Desacompanhados</div>
                    <div class="profile-value">${homensDesacomp}</div>
                    <div class="profile-percent">${percHomensDesacomp}%</div>
                </div>
                <div class="profile-item">
                    <div class="profile-label">👩 Mulheres Desacompanhadas</div>
                    <div class="profile-value">${mulheresDesacomp}</div>
                    <div class="profile-percent">${percMulheresDesacomp}%</div>
                </div>
                <div class="profile-item">
                    <div class="profile-label">👨‍👩‍👧 Famílias com Crianças</div>
                    <div class="profile-value">${familiasCriancas}</div>
                    <div class="profile-percent">${percFamiliasCriancas}%</div>
                </div>
            </div>
            <p style="margin-top: 20px;">O perfil predominante é de <strong>pessoas desacompanhadas (${(parseFloat(percHomensDesacomp) + parseFloat(percMulheresDesacomp)).toFixed(1)}%)</strong>,
            sugerindo ${familiasCriancas > 0 ? 'a necessidade de estratégias diferenciadas: programas de reinserção individual e suporte familiar específico' :
            'foco em programas de reinserção social e autonomia individual'}.</p>
        </div>

        <div class="insight-item">
            <h4>🛏️🍽️ Padrões de Utilização dos Serviços</h4>
            <p><strong>${pernoitePercent}%</strong> da base ativa utiliza o serviço de <strong>pernoite completo</strong>, enquanto
            <strong>${alimentacaoPercent}%</strong> buscam apenas <strong>alimentação</strong>.
            ${parseFloat(alimentacaoPercent) > 25 ?
            '⚠️ <strong>Insight:</strong> Demanda significativa por alimentação sem pernoite indica vulnerabilidade alimentar na comunidade local. Considere expandir programas de segurança alimentar.' :
            '✅ <strong>Insight:</strong> Maioria utiliza serviço completo, indicando situação de vulnerabilidade habitacional. Priorize programas de moradia.'}</p>
        </div>

        <div class="insight-item">
            <h4>📈 Tendência e Dinâmica de Frequência</h4>
            <p>A média de <strong>${data.media10Dias} pessoas/dia</strong> nos últimos 10 dias
            ${data.media10Dias > data.mediaDiaria ?
            `está <strong>${((data.media10Dias / data.mediaDiaria - 1) * 100).toFixed(1)}% acima</strong> da média histórica (${data.mediaDiaria}), indicando <strong>tendência de crescimento</strong> na demanda. 📊 <strong>Recomendação:</strong> Preparar infraestrutura para possível aumento sustentado.` :
            data.media10Dias < data.mediaDiaria ?
            `está <strong>${((1 - data.media10Dias / data.mediaDiaria) * 100).toFixed(1)}% abaixo</strong> da média histórica (${data.mediaDiaria}), indicando <strong>estabilização ou redução temporária</strong>. Momento propício para otimização de recursos.` :
            `está <strong>estável</strong> em relação à média histórica (${data.mediaDiaria}), indicando demanda consistente.`}</p>
        </div>

        <div class="insight-item">
            <h4>🎯 Estratégias Recomendadas para Abrigamento</h4>
            <p><strong>Para venezuelanos (${percVenezuelanos}% da base):</strong></p>
            <ul style="margin: 10px 0 10px 20px; line-height: 1.8;">
                <li>✅ Articulação imediata com pontos focais da <strong>Operação Acolhida</strong></li>
                <li>📋 Regularização documental (CPF, Carteira de Trabalho, RG)</li>
                <li>🚌 Processos de interiorização para outros estados</li>
                <li>🏘️ Encaminhamento para abrigos da rede federal</li>
            </ul>
            <p style="margin-top: 10px;"><strong>Para outras nacionalidades (${percNaoVenezuelanos}%):</strong></p>
            <ul style="margin: 10px 0 0 20px; line-height: 1.8;">
                <li>🤝 Parcerias com Secretarias Municipais de Assistência Social</li>
                <li>🏛️ Articulação com Sistema Único de Assistência Social (SUAS)</li>
                <li>🌐 ONGs e organizações internacionais (ACNUR, OIM, Cáritas)</li>
                <li>💼 Programas de qualificação profissional e empregabilidade</li>
            </ul>
        </div>
    `;
}

function updateLastUpdated(dataReferencia) {
    const now = new Date();
    const processado = now.toLocaleString('pt-BR', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
    });

    const ultimaEntrada = dataReferencia ? dataReferencia.toLocaleDateString('pt-BR') : 'N/A';

    document.getElementById('lastUpdated').innerHTML = `
        📅 <strong>Data de Referência dos Dados:</strong> ${ultimaEntrada} (última entrada no Excel)<br>
        🕐 <strong>Dashboard processado em:</strong> ${processado}
    `;
}
