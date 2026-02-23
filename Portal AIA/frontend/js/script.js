function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

let selectedFile = null;
let totalLines = 0;
let batchSize = 100; // valor inicial
const BATCH_MAX = 100;
let selectedOutputDirHandle = null; // File System Access API handle
let detectedMappingLocal = null;
let lastOutputFolder = null;

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
        const outputFormat = document.getElementById('outputFormat');
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
        if (outputFormat) outputFormat.addEventListener('change', () => showDiagnostics('Formato: ' + outputFormat.value));

        // botão copiar link de rede
        const copyLinkBtn = document.getElementById('copyNetworkLinkBtn');
        if (copyLinkBtn) {
            copyLinkBtn.addEventListener('click', copyNetworkLink);
        }
    } catch (err) {
        console.error('Erro na inicialização do script:', err);
    }
});

function handleFile(file) {
    if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
        selectedFile = file;
        document.getElementById('fileName').textContent = `📄 ${file.name}`;
        // tenta ler o arquivo no cliente para contar linhas (usa SheetJS se disponível)
        const batchInput = document.getElementById('batchSize');
        if (batchInput) {
            batchSize = Math.max(1, Math.min(BATCH_MAX, Number(batchInput.value) || 100));
            batchInput.value = batchSize;
        }

        if (window.XLSX) {
            const reader = new FileReader();
            reader.onload = function (e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const wb = window.XLSX.read(data, { type: 'array' });
                    const firstSheet = wb.SheetNames && wb.SheetNames[0];
                    if (firstSheet) {
                        const ws = wb.Sheets[firstSheet];
                        const rows = window.XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
                        let count = 0;
                        if (rows.length === 0) count = 0;
                        else {
                            // detecta se existe header pela presença de letras na primeira linha
                            const firstRow = rows[0] || [];
                            const hasHeader = firstRow.some(cell => /[A-Za-zÀ-ú]/.test(String(cell)));
                            count = hasHeader ? Math.max(0, rows.length - 1) : rows.length;
                                // se houver header, tenta detectar mapeamento local de colunas
                                if (hasHeader) {
                                    const header_row = firstRow.map(h => String(h).trim());
                                    detectedMappingLocal = detectMappingFromHeaders(header_row);
                                    showLocalMapping(detectedMappingLocal);
                                } else {
                                    detectedMappingLocal = null;
                                    showLocalMapping(null);
                                }
                        }
                        totalLines = count;
                        showDiagnostics('Linhas detectadas: ' + totalLines + ' (sheet: ' + firstSheet + ')');
                        updatePrediction();
                        document.getElementById('btnProcess').disabled = false;
                        return;
                    }
                } catch (err) {
                    console.warn('Erro ao ler Excel localmente:', err);
                }
                // fallback se leitura falhar
                totalLines = 0;
                updatePrediction();
            };
            reader.onerror = function () {
                totalLines = 0;
                updatePrediction();
            };
            reader.readAsArrayBuffer(file);
        } else {
            // SheetJS não disponível: definir 0 e instruir a usar servidor para previsão
            totalLines = 0;
            showDiagnostics('SheetJS não disponível — previsão só após upload.');
            updatePrediction();
        }
        // limpa o campo empresa sempre que um novo arquivo for adicionado
        const companyInputEl = document.getElementById('companyInput');
        if (companyInputEl) companyInputEl.value = '';
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
    const predictionEl = document.getElementById('predictionText');
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
        return;
    }

    // quando não há contagem disponível
    if (selectedFile) {
        if (predictionEl) predictionEl.textContent = 'Lendo arquivo para previsão...';
    } else {
        if (predictionEl) predictionEl.textContent = 'Selecione um arquivo para ver a previsão';
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
    const outputFormatEl = document.getElementById('outputFormat');
    if (!outputFormatEl || !outputFormatEl.value) { alert('Selecione o formato de envio (obrigatório).'); btn.disabled = false; return; }

    // Se o formato for 'lista', garantimos que o usuário tenha escolhido pasta e concedido permissão
    if (String(outputFormatEl.value).toLowerCase() === 'lista') {
        // se o browser suportar showDirectoryPicker e não houver handle, solicitar agora (é user-gesture)
        if (window.showDirectoryPicker && !selectedOutputDirHandle) {
            const want = confirm('Formato "Lista" selecionado. Deseja escolher uma pasta para salvar os arquivos localmente agora? (recomendado)');
            if (want) {
                try {
                    selectedOutputDirHandle = await window.showDirectoryPicker();
                    const nameEl = document.getElementById('outputPickName');
                    try { if (nameEl) nameEl.textContent = selectedOutputDirHandle.name || '(selecionada)'; } catch(e) {}
                    showDiagnostics('Pasta escolhida: ' + (selectedOutputDirHandle.name || '(selecionada)'));
                } catch (errChoose) {
                    // usuário cancelou — permitimos continuar, mas avisamos que fará download
                    showDiagnostics('Seleção de pasta cancelada. Os arquivos serão baixados como fallback.');
                }
            }
        }

        // se já existe handle, tente garantir permissão de escrita (query/request)
        if (selectedOutputDirHandle && typeof selectedOutputDirHandle.queryPermission === 'function') {
            try {
                const q = await selectedOutputDirHandle.queryPermission({ mode: 'readwrite' });
                if (q !== 'granted') {
                    const r = await selectedOutputDirHandle.requestPermission({ mode: 'readwrite' });
                    if (r !== 'granted') {
                        alert('Permissão de escrita na pasta não concedida. Os arquivos serão baixados como fallback.');
                    } else {
                        showDiagnostics('Permissão concedida para pasta: ' + (selectedOutputDirHandle.name || '(selecionada)'));
                    }
                }
            } catch (permErr) {
                console.warn('Erro ao verificar/perguntar permissão:', permErr);
            }
        }
    }

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
    formData.append('output_format', outputFormatEl.value);
    // se mapeamento editável presente, anexar seleção explícita
    try {
        const selNum = document.getElementById('map_numero');
        const selCnpj = document.getElementById('map_cnpj');
        const selAcao = document.getElementById('map_acao');
        if (selNum && selNum.value) formData.append('numero_col', selNum.value);
        if (selCnpj && selCnpj.value) formData.append('cnpj_col', selCnpj.value);
        if (selAcao && selAcao.value) formData.append('acao_col', selAcao.value);
    } catch (e) { /* ignore */ }
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
            if (result.output_folder) lastOutputFolder = result.output_folder;

            // mostrar mapeamento retornado pelo backend, se presente
            if (result.column_mapping) {
                try {
                    const mapping = result.column_mapping;
                    const mapBox = document.getElementById('mappingBox');
                    const mapList = document.getElementById('mappingList');
                    const items = [];
                    if (mapping.numero) items.push(`<span class="mapping-item">numero → ${mapping.numero}</span>`);
                    if (mapping.cnpj) items.push(`<span class="mapping-item">cnpj → ${mapping.cnpj}</span>`);
                    if (mapping.acao) items.push(`<span class="mapping-item">acao → ${mapping.acao}</span>`);
                    if (mapList) mapList.innerHTML = items.join(' ');
                    if (mapBox) mapBox.style.display = items.length ? 'block' : 'none';
                } catch (e) { console.warn('Erro mostrando mapeamento:', e); }
            }

            // mostrar preview retornado pelo backend (opcional)
            if (result.preview && result.preview.length) {
                try {
                    const pv = result.preview;
                    const previewText = pv.map(r => `${r.numero} | ${r.acao} | ${r.cnpj}`).join('\n');
                    showDiagnostics('Preview (primeiras linhas):\n' + previewText);
                    alert('Preview (primeiras linhas):\n' + previewText);
                } catch (e) { }
            }

            // Se o backend retornou os conteúdos (base64), oferece salvar na pasta escolhida
            if (result.files_data && result.files_data.length) {
                try {
                    // Preferimos salvar diretamente na pasta escolhida, mas só se o handle existir e tiver permissão
                    if (window.showDirectoryPicker && selectedOutputDirHandle) {
                        let canWrite = false;
                        try {
                            const opts = { mode: 'readwrite' };
                            if (typeof selectedOutputDirHandle.queryPermission === 'function') {
                                const q = await selectedOutputDirHandle.queryPermission(opts);
                                if (q === 'granted') canWrite = true;
                                else {
                                    try {
                                        const r = await selectedOutputDirHandle.requestPermission(opts);
                                        if (r === 'granted') canWrite = true;
                                    } catch (e) {
                                        canWrite = false;
                                    }
                                }
                            }
                        } catch (e) {
                            canWrite = false;
                        }

                        if (canWrite) {
                            for (const f of result.files_data) {
                                const name = f.name;
                                const b64 = f.content_b64;
                                if (!b64) continue;
                                const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
                                try {
                                    const fileHandle = await selectedOutputDirHandle.getFileHandle(name, { create: true });
                                    const writable = await fileHandle.createWritable();
                                    await writable.write(bytes);
                                    await writable.close();
                                } catch (e) {
                                    console.warn('Falha ao escrever arquivo na pasta escolhida, fallback para download:', e);
                                    const blob = new Blob([bytes]);
                                    const url = URL.createObjectURL(blob);
                                    const a = document.createElement('a');
                                    a.href = url;
                                    a.download = name;
                                    document.body.appendChild(a);
                                    a.click();
                                    a.remove();
                                    URL.revokeObjectURL(url);
                                }
                            }
                            showDiagnostics('Arquivos salvos na pasta escolhida.');
                        } else {
                            // sem permissão para gravar — informar e usar fallback para downloads
                            showDiagnostics('Sem permissão para salvar na pasta escolhida. Faça a seleção da pasta antes de processar, conceda permissão, ou os arquivos serão baixados.');
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
                        }
                    } else {
                        // fallback: baixar cada arquivo individualmente
                        showDiagnostics('Nenhuma pasta escolhida para salvar localmente. Os arquivos serão baixados.');
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

            // adiciona botão para baixar ZIP de todos os arquivos (via backend)
            try {
                const zipBtnId = 'downloadZipBtn';
                let zipBtn = document.getElementById(zipBtnId);
                if (!zipBtn) {
                    zipBtn = document.createElement('button');
                    zipBtn.id = zipBtnId;
                    zipBtn.className = 'btn-copy-link';
                    zipBtn.textContent = '📦 Baixar ZIP dos arquivos';
                    zipBtn.style.marginLeft = '8px';
                    const parent = document.querySelector('.action-section');
                    if (parent) parent.appendChild(zipBtn);
                    zipBtn.addEventListener('click', async () => {
                        try {
                            const payload = { folder: result.output_folder };
                            const r = await fetch('/api/download_zip', {
                                method: 'POST',
                                headers: { 'Content-Type': 'application/json' },
                                body: JSON.stringify(payload)
                            });
                            if (!r.ok) {
                                const j = await r.json().catch(() => ({}));
                                throw new Error(j.error || 'Falha ao gerar ZIP');
                            }
                            const blob = await r.blob();
                            const url = URL.createObjectURL(blob);
                            const a = document.createElement('a');
                            a.href = url;
                            // tenta extrair filename do header, fallback
                            const cd = r.headers.get('Content-Disposition') || '';
                            const m = /filename="?([^";]+)"?/.exec(cd);
                            a.download = (m && m[1]) || ('files.zip');
                            document.body.appendChild(a);
                            a.click();
                            a.remove();
                            URL.revokeObjectURL(url);
                        } catch (e) {
                            alert('Erro ao baixar ZIP: ' + (e.message || e));
                        }
                    });
                }
            } catch (e) { /* ignore */ }
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
    // Tenta abrir / expor a pasta de saída para o usuário.
    // Comportamentos possíveis:
    // - Se o usuário escolheu uma pasta via File System Access, mostramos instruções e relembramos o nome.
    // - Se o backend retornou um caminho `lastOutputFolder`, copiamos para a área de transferência e tentamos abrir via file:// (pode ser bloqueado pelo navegador).
    (async () => {
        try {
            if (selectedOutputDirHandle) {
                const name = selectedOutputDirHandle.name || '(pasta selecionada)';
                try { await navigator.clipboard.writeText(name); } catch (e) { /* ignore */ }
                alert(`A pasta escolhida pelo navegador é: ${name}\nDica: use o gerenciador de arquivos para navegar até essa pasta.`);
                return;
            }

            if (lastOutputFolder) {
                // copiar para clipboard
                try {
                    await navigator.clipboard.writeText(lastOutputFolder);
                } catch (err) {
                    // não crítico
                }

                // tenta abrir via file:// (pode ser bloqueado pelo navegador)
                const path = lastOutputFolder.replace(/\\/g, '/');
                const fileUrl = 'file:///' + encodeURI(path);
                let opened = false;
                try {
                    const w = window.open(fileUrl, '_blank');
                    if (w) opened = true;
                } catch (e) {
                    opened = false;
                }

                const msg = `Pasta de saída: ${lastOutputFolder}\n(O caminho foi copiado para a área de transferência.)\n` +
                    (opened ? 'O navegador tentou abrir a pasta usando file:// (pode ser bloqueado).' : 'Se o navegador não abrir a pasta, cole o caminho no Windows Explorer (Barra de Endereço) ou use `explorer.exe "<caminho>"` no PowerShell.');
                alert(msg);
                return;
            }

            alert('Nenhuma pasta de saída conhecida. Execute um processamento ou escolha uma pasta antes.');
        } catch (err) {
            console.error('Erro ao tentar abrir pasta:', err);
            alert('Não foi possível abrir a pasta automaticamente. Caminho copiado se disponível.');
        }
    })();
}

function normalizeHeader(s) {
    if (!s) return '';
    return String(s).toLowerCase().normalize('NFD').replace(/\p{Diacritic}/gu, '').replace(/[^a-z0-9]/g, '');
}

function detectMappingFromHeaders(headers) {
    // headers: array of header strings
    const norms = headers.map(h => normalizeHeader(h));
    const mapping = { numero: null, cnpj: null, acao: null };
    const numeroAlts = ['numero','num','did','id','numeroid','msisdn','telefone','tel','phone','celular','mobile'];
    const cnpjAlts = ['cnpj','cpfcnpj','cpf','taxid','taxidnumber','documento'];
    const acaoAlts = ['acao','action','operacao','operacao'];
    norms.forEach((n, idx) => {
        if (!mapping.numero && numeroAlts.includes(n)) mapping.numero = headers[idx];
        if (!mapping.cnpj && cnpjAlts.includes(n)) mapping.cnpj = headers[idx];
        if (!mapping.acao && acaoAlts.includes(n)) mapping.acao = headers[idx];
    });
    return mapping;
}

function showLocalMapping(mapping, headers) {
    const box = document.getElementById('mappingBox');
    const list = document.getElementById('mappingList');
    if (!box || !list) return;
    if (!mapping || !headers || !headers.length) {
        box.style.display = 'none';
        list.innerHTML = '';
        return;
    }

    // cria selects editáveis para cada campo
    function buildSelect(id, label, selected) {
        const opts = [''].concat(headers);
        let html = `<label style="font-size:13px;margin-right:6px">${label}</label>`;
        html += `<select id="${id}" style="padding:6px;border-radius:6px;border:1px solid #ddd;margin-right:8px">`;
        for (const o of opts) {
            const safe = o || '';
            const sel = (safe === selected) ? 'selected' : '';
            html += `<option value="${safe}" ${sel}>${safe || '(não identificado)'}</option>`;
        }
        html += `</select>`;
        return html;
    }

    const htmlParts = [];
    htmlParts.push(buildSelect('map_numero', 'Coluna número/telefone:', mapping.numero || ''));
    htmlParts.push(buildSelect('map_cnpj', 'Coluna CPF/CNPJ:', mapping.cnpj || ''));
    htmlParts.push(buildSelect('map_acao', 'Coluna Ação (opcional):', mapping.acao || ''));

    htmlParts.push('<div style="margin-top:8px;font-size:12px;color:#666">Altere os mapeamentos acima, se necessário. As alterações serão aplicadas ao processar.</div>');

    list.innerHTML = htmlParts.join(' ');
    box.style.display = 'block';
}

async function copyNetworkLink() {
    try {
        const resp = await fetch('/api/hostinfo');
        if (resp.ok) {
            const j = await resp.json();
            if (j.success && j.ips && j.ips.length) {
                const port = j.port || window.location.port || 80;
                const urls = j.ips.map(ip => `http://${ip}:${port}`);
                const primary = urls[0];
                try {
                    await navigator.clipboard.writeText(primary);
                    showDiagnostics('Link copiado: ' + primary);
                    alert('Link copiado para a área de transferência:\n' + primary + '\n\nOutros IPs detectados:\n' + urls.join('\n'));
                    return;
                } catch (e) {
                    // fallback
                }
                // fallback copy via input
                const tmp = document.createElement('input');
                tmp.value = primary;
                document.body.appendChild(tmp);
                tmp.select();
                document.execCommand('copy');
                tmp.remove();
                showDiagnostics('Link copiado: ' + primary);
                alert('Link copiado: ' + primary);
                return;
            }
        }
    } catch (e) {
        console.warn('Erro obtendo hostinfo', e);
    }

    // fallback: usar origin
    try {
        const origin = window.location.origin;
        await navigator.clipboard.writeText(origin);
        showDiagnostics('Link copiado (origin): ' + origin);
        alert('Link copiado: ' + origin + '\nSe for localhost/127.0.0.1 substitua pelo IP do servidor na rede.');
    } catch (e) {
        alert('Não foi possível copiar automaticamente. Endereço sugerido:\n' + window.location.origin);
    }
}