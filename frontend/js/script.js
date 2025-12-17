function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

let selectedFile = null;
let totalLines = 0;
let batchSize = 100; // valor inicial
const BATCH_MAX = 100;
let selectedOutputDirHandle = null; // File System Access API handle

// Inicialização dependente do DOM
document.addEventListener('DOMContentLoaded', function () {
    try {
        const dropArea = document.getElementById('dropArea');
        const fileInput = document.getElementById('fileInput');

        if (dropArea) {
            ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
                dropArea.addEventListener(eventName, preventDefaults, false);
            });

            ['dragenter', 'dragover'].forEach(eventName => {
                dropArea.addEventListener(eventName, () => {
                    dropArea.classList.add('drag-over');
                }, false);
            });

            ['dragleave', 'drop'].forEach(eventName => {
                dropArea.addEventListener(eventName, () => {
                    dropArea.classList.remove('drag-over');
                }, false);
            });

            dropArea.addEventListener('drop', (e) => {
                const files = e.dataTransfer.files;
                if (files.length > 0) {
                    handleFile(files[0]);
                }
            });
        }

        if (fileInput) {
            fileInput.addEventListener('change', (e) => {
                if (e.target.files.length > 0) {
                    handleFile(e.target.files[0]);
                }
            });
        }

        const btnProcess = document.getElementById('btnProcess');
        if (btnProcess) {
            btnProcess.addEventListener('click', processFiles);
        }

        // elementos novos: batch input, action, company
        const batchInput = document.getElementById('batchSize');
        if (batchInput) {
            // inicializa valor
            batchInput.value = Math.min(BATCH_MAX, Math.max(1, batchInput.value || 100));
            batchSize = Number(batchInput.value);
            batchInput.addEventListener('change', () => {
                let v = Number(batchInput.value) || 1;
                v = Math.max(1, Math.min(BATCH_MAX, v));
                batchInput.value = v;
                batchSize = v;
                updatePrediction();
            });
        }

        const actionSelect = document.getElementById('actionSelect');
        const companyInput = document.getElementById('companyInput');
        const outputPickBtn = document.getElementById('outputPickBtn');
        const outputPickName = document.getElementById('outputPickName');
        // botão para escolher pasta (File System Access API)
        if (outputPickBtn) {
            outputPickBtn.addEventListener('click', async () => {
                try {
                    if (window.showDirectoryPicker) {
                        const dir = await window.showDirectoryPicker();
                        selectedOutputDirHandle = dir;
                        // mostra nome da pasta
                        try { outputPickName.textContent = dir.name || '(selecionada)'; } catch (e) { outputPickName.textContent = '(selecionada)'; }
                        showDiagnostics('Pasta escolhida: ' + (dir.name || '(selecionada)'));
                    } else {
                        alert('Seu navegador não suporta seleção de pastas (use Chrome/Edge).');
                    }
                } catch (err) {
                    console.warn('Cancelado seleção de pasta', err);
                }
            });
        }
        // atualiza diagnóstico quando mudar
        if (actionSelect) actionSelect.addEventListener('change', () => showDiagnostics('Ação: ' + actionSelect.value));
        if (companyInput) companyInput.addEventListener('input', () => showDiagnostics('Empresa: ' + companyInput.value));
    } catch (err) {
        console.error('Erro na inicialização do script:', err);
    }
});

function handleFile(file) {
    if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
        selectedFile = file;
        document.getElementById('fileName').textContent = `📄 ${file.name}`;
        totalLines = 6600;
        // atualiza batchSize a partir do input caso exista
        const batchInput = document.getElementById('batchSize');
        if (batchInput) {
            batchSize = Math.max(1, Math.min(BATCH_MAX, Number(batchInput.value) || 100));
            batchInput.value = batchSize;
        }
        // limpa o campo empresa sempre que um novo arquivo for adicionado
        const companyInputEl = document.getElementById('companyInput');
        if (companyInputEl) companyInputEl.value = '';
        updatePrediction();
        document.getElementById('btnProcess').disabled = false;
    } else {
        alert('Por favor, selecione um arquivo Excel válido (.xlsx ou .xls)');
    }
}

function adjustBatch(amount) {
    const batchInput = document.getElementById('batchSize');
    let current = Number(batchInput ? batchInput.value : batchSize) || batchSize;
    current = current + amount;
    current = Math.max(1, Math.min(BATCH_MAX, current));
    if (batchInput) batchInput.value = current;
    batchSize = current;
    updatePrediction();
}

function updatePrediction() {
    if (totalLines > 0) {
        // garante que batchSize reflita o input
        const batchInput = document.getElementById('batchSize');
        if (batchInput) batchSize = Math.max(1, Math.min(BATCH_MAX, Number(batchInput.value) || batchSize));
        const fileCount = Math.ceil(totalLines / batchSize);
        const actionSelect = document.getElementById('actionSelect');
        const companyInput = document.getElementById('companyInput');
        const actionLabel = actionSelect ? actionSelect.value : '';
        const companyLabel = companyInput ? companyInput.value : '';
        document.getElementById('predictionText').innerHTML =
            `<strong>Previsão:</strong> Serão gerados aproximadamente <strong>${fileCount}</strong> arquivo(s)` +
            (actionLabel || companyLabel ? `<br><small>Ação: ${actionLabel} · Empresa: ${companyLabel}</small>` : '');
    }
}


async function processFiles() {
    const btn = document.getElementById('btnProcess');
    const progressContainer = document.getElementById('progressContainer');
    const progressBar = document.getElementById('progressBar');
    const progressText = document.getElementById('progressText');
    const successMessage = document.getElementById('successMessage');

    // Validações básicas
    if (!selectedFile) { alert('Selecione um arquivo!'); return; }
    const companyInput = document.getElementById('companyInput');
    if (!companyInput || !companyInput.value.trim()) { alert('Informe o nome da empresa!'); return; }

    btn.disabled = true;
    progressContainer.classList.add('active');

    // Preparar FormData
    const formData = new FormData();
    formData.append('file', selectedFile);
    const actionSelect = document.getElementById('actionSelect');
    formData.append('action', actionSelect ? actionSelect.value : 'criar');
    formData.append('company', companyInput.value.trim());
    const batchInput = document.getElementById('batchSize');
    formData.append('batchSize', batchInput ? batchInput.value : String(batchSize));
    // não enviamos outputBase — frontend grava localmente usando File System Access API

    // UI de envio
    progressBar.style.width = '5%';
    progressBar.textContent = 'Enviando...';
    progressText.textContent = 'Enviando arquivo e processando no servidor...';

    try {
        // se a página foi aberta via file://, fetch para '/api/processar' falhará
        if (location.protocol === 'file:') {
            alert('Erro: abra o aplicativo via servidor HTTP (ex: http://127.0.0.1:5000), não abra o arquivo index.html diretamente.');
            btn.disabled = false;
            return;
        }

        const apiUrl = (window.location.origin ? window.location.origin : '') + '/api/processar';
        const resp = await fetch(apiUrl, {
            method: 'POST',
            body: formData
        });

        const result = await resp.json();

        if (resp.ok && result.success) {
            progressBar.style.width = '100%';
            progressBar.textContent = '100%';
            progressBar.style.background = 'linear-gradient(90deg, #2ecc71 0%, #27ae60 100%)';
            progressText.textContent = 'Processamento concluído!';
            const count = result.total_files || (result.total_files === 0 ? 0 : 0);
            showSuccess(count);
            showDiagnostics(`Concluído: ${count} arquivo(s) em ${result.output_folder}`);

            // Se o backend retornou os conteúdos (base64), oferece salvar na pasta escolhida
            if (result.files_data && result.files_data.length) {
                try {
                    if (window.showDirectoryPicker) {
                        const dirHandle = selectedOutputDirHandle || await window.showDirectoryPicker();
                        for (const f of result.files_data) {
                            const name = f.name;
                            const b64 = f.content_b64;
                            if (!b64) continue;
                            const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
                            const fileHandle = await dirHandle.getFileHandle(name, { create: true });
                            const writable = await fileHandle.createWritable();
                            await writable.write(bytes);
                            await writable.close();
                        }
                        showDiagnostics('Arquivos salvos na pasta escolhida.');
                    } else {
                        // fallback: baixar cada arquivo individualmente
                        for (const f of result.files_data) {
                            if (!f.content_b64) continue;
                            const bytes = Uint8Array.from(atob(f.content_b64), c => c.charCodeAt(0));
                            const blob = new Blob([bytes]);
                            const url = URL.createObjectURL(blob);
                            const a = document.createElement('a');
                            a.href = url;
                            a.download = f.name;
                            document.body.appendChild(a);
                            a.click();
                            a.remove();
                            URL.revokeObjectURL(url);
                        }
                        showDiagnostics('Arquivos foram baixados (um por vez).');
                    }
                } catch (errSave) {
                    console.error('Erro ao salvar arquivos localmente:', errSave);
                    showDiagnostics('Erro ao salvar arquivos localmente: ' + (errSave.message || errSave));
                }
            }
        } else {
            throw new Error(result.error || 'Erro desconhecido do servidor');
        }

    } catch (err) {
        console.error('Erro:', err);
        progressText.textContent = 'Erro no processamento';
        progressBar.style.background = '#e74c3c';
        alert('Falha no processamento: ' + (err.message || err));
        showDiagnostics('Erro: ' + (err.message || err));
    } finally {
        btn.disabled = false;
    }
}

function showSuccess(fileCount) {
    const successMessage = document.getElementById('successMessage');
    document.getElementById('filesCreated').textContent = fileCount;
    successMessage.classList.add('active');

    document.getElementById('progressBar').style.background =
        'linear-gradient(90deg, #2ecc71 0%, #27ae60 100%)';
}

function openOutputFolder() {
    alert('Em produção, isso abriria a pasta: CREFITECH');
}