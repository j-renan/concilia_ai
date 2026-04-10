document.addEventListener('DOMContentLoaded', () => {
    // Elements - Steps
    const stepUpload = document.getElementById('step-upload');
    const stepMapping = document.getElementById('step-mapping');
    const stepResults = document.getElementById('step-results');

    // Elements - Files
    const dropCredito = document.getElementById('drop-credito');
    const dropFrete = document.getElementById('drop-frete');
    const fileCredito = document.getElementById('file-credito');
    const fileFrete = document.getElementById('file-frete');
    const btnNextMapping = document.getElementById('btn-next-mapping');

    // Elements - Mapping
    const btnBackUpload = document.getElementById('btn-back-upload');
    const btnProcess = document.getElementById('btn-process');
    const selects = document.querySelectorAll('.header-select');

    // Elements - Results
    const btnRestart = document.getElementById('btn-restart');
    const tableBody = document.querySelector('#table-divergencias tbody');

    let headers = { credito: [], frete: [] };

    // --- Step 1: Upload Logic ---

    const setupDropZone = (zone, input) => {
        zone.addEventListener('click', () => input.click());
        zone.addEventListener('dragover', (e) => {
            e.preventDefault();
            zone.classList.add('hover');
        });
        zone.addEventListener('dragleave', () => zone.classList.remove('hover'));
        zone.addEventListener('drop', (e) => {
            e.preventDefault();
            zone.classList.remove('hover');
            if (e.dataTransfer.files.length) {
                input.files = e.dataTransfer.files;
                updateFileInfo(zone, input.files[0].name);
            }
        });
        input.addEventListener('change', () => {
            if (input.files.length) {
                updateFileInfo(zone, input.files[0].name);
            }
        });
    };

    const updateFileInfo = (zone, name) => {
        zone.classList.add('has-file');
        zone.querySelector('.file-name').textContent = name;
        checkUploadComplete();
    };

    const checkUploadComplete = () => {
        btnNextMapping.disabled = !(fileCredito.files.length && fileFrete.files.length);
    };

    setupDropZone(dropCredito, fileCredito);
    setupDropZone(dropFrete, fileFrete);

    btnNextMapping.addEventListener('click', async () => {
        const formData = new FormData();
        formData.append('file_credito', fileCredito.files[0]);
        formData.append('file_frete', fileFrete.files[0]);

        btnNextMapping.disabled = true;
        btnNextMapping.innerHTML = '<span class="loading">Enviando...</span>';

        try {
            const response = await fetch('/upload', { method: 'POST', body: formData });
            const data = await response.json();

            if (data.error) throw new Error(data.error);

            headers.credito = data.headers_credito;
            headers.frete = data.headers_frete;

            populateSelects();
            showStep('mapping');
        } catch (err) {
            alert('Erro no upload: ' + err.message);
        } finally {
            btnNextMapping.disabled = false;
            btnNextMapping.innerHTML = 'Próximo Passo <i data-lucide="chevron-right"></i>';
            lucide.createIcons();
        }
    });

    // --- Step 2: Mapping Logic ---

    const populateSelects = () => {
        const populate = (id, options) => {
            const select = document.getElementById(id);
            select.innerHTML = '<option value="">Selecione uma coluna...</option>' +
                options.map(h => `<option value="${h}">${h}</option>`).join('');
        };

        populate('map-credito-historico', headers.credito);
        populate('map-credito-valor', headers.credito);
        populate('map-frete-documento', headers.frete);
        populate('map-frete-valor', headers.frete);
        populate('map-frete-destinatario', headers.frete);

        // Auto-select common names if they exist
        autoMap();
    };

    const autoMap = () => {
        const normalize = (str) => str.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");

        const trySelect = (id, patterns) => {
            const select = document.getElementById(id);
            const normalizedOptions = [...select.options].map(o => normalize(o.value));

            for (let p of patterns) {
                const normP = normalize(p);
                const idx = normalizedOptions.findIndex(o => o.includes(normP));
                if (idx !== -1) {
                    select.selectedIndex = idx;
                    break;
                }
            }
        };

        trySelect('map-credito-historico', ['histórico', 'memo', 'descrição']);
        trySelect('map-credito-valor', ['crédito', 'valor', 'total', 'vl']);
        trySelect('map-frete-documento', ['documento', 'cte', 'número', 'docto']);
        trySelect('map-frete-valor', ['valor frete', 'valor', 'frete', 'vl']);
        trySelect('map-frete-destinatario', ['destinatário', 'cliente', 'nome']);
    };

    btnBackUpload.addEventListener('click', () => showStep('upload'));

    btnProcess.addEventListener('click', async () => {
        const mapping = {
            credito: {
                historico: document.getElementById('map-credito-historico').value,
                valor: document.getElementById('map-credito-valor').value
            },
            frete: {
                documento: document.getElementById('map-frete-documento').value,
                valor: document.getElementById('map-frete-valor').value,
                destinatario: document.getElementById('map-frete-destinatario').value
            }
        };

        if (!mapping.credito.historico || !mapping.credito.valor || !mapping.frete.documento || !mapping.frete.valor) {
            alert('Por favor, preencha todos os campos obrigatórios de mapeamento.');
            return;
        }

        btnProcess.disabled = true;
        btnProcess.textContent = 'Processando...';

        try {
            const response = await fetch('/process', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ mapping })
            });
            const data = await response.json();

            if (data.error) throw new Error(data.error);

            renderResults(data);
            showStep('results');
        } catch (err) {
            alert('Erro no processamento: ' + err.message);
        } finally {
            btnProcess.disabled = false;
            btnProcess.textContent = 'Iniciar Conciliação';
        }
    });

    // --- Step 3: Results Logic ---

    const renderResults = (data) => {
        document.getElementById('stat-total-divergencias').textContent = data.summary.total_divergencias;
        document.getElementById('stat-valor-total').textContent = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(data.summary.valor_total_divergencia);
        document.getElementById('stat-credito-total-faltantes').textContent = data.summary.total_credito_sem_frete;
        document.getElementById('stat-frete-total-faltantes').textContent = data.summary.total_frete_sem_credito;

        tableBody.innerHTML = '';
        data.divergencias.forEach(row => {
            const tr = document.createElement('tr');
            // Garantir que valores nulos sejam tratados como 0 para o formatador
            const valCredito = row['Valor Crédito'] || 0;
            const valFrete = row['Valor Frete'] || 0;
            const diferenca = row.Diferença || 0;

            tr.innerHTML = `
                <td><strong>${row.Documento}</strong></td>
                <td>${row.Destinatário || '-'}</td>
                <td>${new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(valCredito)}</td>
                <td>${new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(valFrete)}</td>
                <td style="color: var(--error); font-weight: 600;">${new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(diferenca)}</td>
                <td style="font-size: 0.8rem; color: var(--text-muted);">${row.Observação}</td>
            `;
            tableBody.appendChild(tr);
        });

        if (data.divergencias.length === 0) {
            tableBody.innerHTML = '<tr><td colspan="6" style="text-align: center; padding: 2rem;">Nenhuma divergência encontrada! ✅</td></tr>';
        }
    };

    btnRestart.addEventListener('click', () => {
        // Reset inputs
        fileCredito.value = '';
        fileFrete.value = '';
        dropCredito.classList.remove('has-file');
        dropFrete.classList.remove('has-file');
        dropCredito.querySelector('.file-name').textContent = 'Nenhum arquivo selecionado';
        dropFrete.querySelector('.file-name').textContent = 'Nenhum arquivo selecionado';
        btnNextMapping.disabled = true;
        showStep('upload');
    });

    // --- Helpers ---

    const showStep = (step) => {
        [stepUpload, stepMapping, stepResults].forEach(s => s.classList.remove('active'));
        if (step === 'upload') stepUpload.classList.add('active');
        if (step === 'mapping') stepMapping.classList.add('active');
        if (step === 'results') stepResults.classList.add('active');
        window.scrollTo({ top: 0, behavior: 'smooth' });
    };
});
