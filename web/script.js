let dadosEnf = [], dadosUti = [], dadosUpa = [];
let editorEnf = [], editorUti = [], editorUpa = [];
let filaImpressao = []; // AGORA É UMA LISTA GLOBAL QUE NUNCA ZERA SOZINHA
let setorAtual = 'ENF';
let abaEditorAtual = 'ENF';

// Inicialização
window.addEventListener('pywebviewready', carregarDados);
setTimeout(() => { if (dadosEnf.length === 0) carregarDados(); }, 1500);

async function recarregarDadosComFeedback() {
    await carregarDados();
    alert("✅ Dados atualizados!");
}

function escaparTexto(texto) {
    if (texto === null || texto === undefined) return "";
    return String(texto).replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

// --- NAVEGAÇÃO ---
function mudarAba(aba) {
    document.querySelectorAll('.tab-content').forEach(d => d.style.display = 'none');
    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));

    let btnIndex = 0;
    if (aba === 'uti') btnIndex = 1;
    if (aba === 'upa') btnIndex = 2;
    if (aba === 'editor') btnIndex = 3;
    document.querySelectorAll('.nav-btn')[btnIndex].classList.add('active');

    if (aba === 'editor') {
        document.getElementById('tab-editor').style.display = 'block';
        document.getElementById('painelBusca').style.display = 'none';
        abaEditorAtual = 'ENF';
        renderizarEditor();
    } else {
        if (aba === 'enf') setorAtual = 'ENF';
        else if (aba === 'uti') setorAtual = 'UTI';
        else setorAtual = 'UPA';

        document.getElementById('tab-dashboard').style.display = 'block';

        let titulos = { 'ENF': 'Enfermarias', 'UTI': 'UTI - HRMSS', 'UPA': 'UPA - Urgência' };
        document.getElementById('tituloSetor').innerText = titulos[setorAtual];

        document.getElementById('painelBusca').style.display = 'block';

        // Renderiza a lista do setor selecionado
        renderizarLista(getDadosAtuais());

        // IMPORTANTE: NÃO LIMPAMOS MAIS A FILA AQUI! 
        // filaImpressao = []; <--- ISSO FOI REMOVIDO PARA PERMITIR MISTURAR

        atualizarFila(); // Apenas redesenha a fila que já existe
    }
}

function getDadosAtuais() {
    if (setorAtual === 'ENF') return dadosEnf;
    if (setorAtual === 'UTI') return dadosUti;
    return dadosUpa;
}

// --- CARREGAMENTO DE DADOS ---
async function carregarDados() {
    if (window.pywebview && window.pywebview.api) {
        try {
            let res = await pywebview.api.carregar_dados_excel();
            if (res.sucesso) {
                dadosEnf = res.dados_enf; dadosUti = res.dados_uti; dadosUpa = res.dados_upa;
                editorEnf = res.editor_enf; editorUti = res.editor_uti; editorUpa = res.editor_upa;

                if (document.getElementById('tab-editor').style.display === 'block') renderizarEditor();
                else renderizarLista(getDadosAtuais());
            } else alert(res.erro);
        } catch (e) { console.error(e); }
    }
}

// --- LISTA LATERAL (BUSCA) ---
function renderizarLista(lista) {
    const div = document.getElementById("listaPacientes");
    div.innerHTML = "";
    if (!lista || lista.length === 0) { div.innerHTML = "<p style='text-align:center;padding:20px;color:#888'>Vazio.</p>"; return; }

    lista.forEach(p => {
        let item = document.createElement("div");
        item.className = "patient-item";
        item.onclick = () => adicionarFila(p);

        let sub = (setorAtual === 'ENF') ? p['ENFERMARIA'] : (setorAtual === 'UTI' ? 'UTI' : 'UPA');
        let dieta = p['DIETA'] ? p['DIETA'] : '---';

        item.innerHTML = `<h4>${p['LEITO']} - ${p['NOME DO PACIENTE']}</h4><p>${sub} | ${dieta}</p>`;
        div.appendChild(item);
    });
}

function filtrarLista() {
    let termo = document.getElementById("inputBusca").value.toLowerCase();
    let listaBase = getDadosAtuais();
    let filtrados = listaBase.filter(p => {
        let nome = String(p['NOME DO PACIENTE']).toLowerCase();
        let leito = String(p['LEITO']).toLowerCase();
        return nome.includes(termo) || leito.includes(termo);
    });
    renderizarLista(filtrados);
}

// --- EDITOR DE PLANILHA ---
function renderizarEditor() {
    const container = document.getElementById("editorControls");
    const cls = (aba) => abaEditorAtual === aba ? 'btn-primary' : 'btn-secondary';

    container.innerHTML = `
        <button class="btn ${cls('ENF')}" onclick="trocarEditor('ENF')">Enf (${editorEnf.length})</button>
        <button class="btn ${cls('UTI')}" onclick="trocarEditor('UTI')">UTI (${editorUti.length})</button>
        <button class="btn ${cls('UPA')}" onclick="trocarEditor('UPA')">UPA (${editorUpa.length})</button>
    `;

    let dados;
    if (abaEditorAtual === 'ENF') dados = editorEnf;
    else if (abaEditorAtual === 'UTI') dados = editorUti;
    else dados = editorUpa;

    const tbody = document.getElementById("corpoTabelaEditor");
    const thead = document.querySelector("#tabelaEditor thead");
    tbody.innerHTML = ""; thead.innerHTML = "";

    let trHead = document.createElement("tr");
    if (abaEditorAtual === 'ENF') {
        trHead.innerHTML = `<th>ENFERMARIA</th><th>LEITO</th><th>NOME DO PACIENTE</th><th>DIETA</th><th>OBSERVAÇÕES</th><th style="width:50px">X</th>`;
    } else {
        trHead.innerHTML = `<th>LEITO</th><th>NOME DO PACIENTE</th><th>DIETA</th><th>OBSERVAÇÕES</th><th style="width:50px">X</th>`;
    }
    thead.appendChild(trHead);
    dados.forEach(row => criarLinhaEditor(tbody, row));
}

function criarLinhaEditor(tbody, row = {}) {
    let tr = document.createElement("tr");
    let html = "";
    const val = (k) => escaparTexto(row[k]);

    if (abaEditorAtual === 'ENF') html += `<td><input class="edit-enf" value="${val('ENFERMARIA')}"></td>`;

    html += `
        <td><input class="edit-leito" value="${val('LEITO')}"></td>
        <td><input class="edit-nome" value="${val('NOME DO PACIENTE')}"></td>
        <td><input class="edit-dieta" value="${val('DIETA')}"></td>
        <td><input class="edit-obs" value="${val('OBSERVAÇÕES')}"></td>
        <td style="text-align:center">
            <button class="btn-remove" tabindex="-1" onclick="this.closest('tr').remove()" title="Excluir">
                <span class="material-icons" style="font-size:18px">delete</span>
            </button>
        </td>
    `;
    tr.innerHTML = html;
    tbody.appendChild(tr);
}

function trocarEditor(tipo) {
    salvarEstadoTemporario();
    abaEditorAtual = tipo;
    renderizarEditor();
}

function salvarEstadoTemporario() {
    const linhas = document.querySelectorAll("#corpoTabelaEditor tr");
    let novosDados = [];
    linhas.forEach(tr => {
        let obj = {};
        if (abaEditorAtual === 'ENF') obj['ENFERMARIA'] = tr.querySelector(".edit-enf").value;
        obj['LEITO'] = tr.querySelector(".edit-leito").value;
        obj['NOME DO PACIENTE'] = tr.querySelector(".edit-nome").value;
        obj['DIETA'] = tr.querySelector(".edit-dieta").value;
        obj['OBSERVAÇÕES'] = tr.querySelector(".edit-obs").value;
        if (obj['LEITO'] || obj['NOME DO PACIENTE']) novosDados.push(obj);
    });
    if (abaEditorAtual === 'ENF') editorEnf = novosDados;
    else if (abaEditorAtual === 'UTI') editorUti = novosDados;
    else editorUpa = novosDados;
}

function adicionarLinhaVazia() {
    criarLinhaEditor(document.getElementById("corpoTabelaEditor"));
    document.getElementById("corpoTabelaEditor").lastElementChild.scrollIntoView({ behavior: 'smooth' });
}

async function salvarExcel() {
    salvarEstadoTemporario();
    if (confirm("Salvar alterações em TODAS as planilhas?")) {
        let res = await pywebview.api.salvar_dados_excel(editorEnf, editorUti, editorUpa);
        if (res.sucesso) { alert("✅ Salvo com sucesso!"); carregarDados(); }
        else alert("❌ " + res.msg);
    }
}

// --- FUNÇÕES DE FILA (ATUALIZADAS E GLOBAIS) ---

function adicionarFila(p) {
    // Evita duplicatas exatas na fila (Opcional, mas recomendado)
    // Se quiser permitir duplicatas, remova as 3 linhas abaixo
    let duplicado = filaImpressao.some(item => item['LEITO'] === p['LEITO'] && item['NOME DO PACIENTE'] === p['NOME DO PACIENTE']);
    if (duplicado) return;

    filaImpressao.push(p);
    atualizarFila();
}

function limparFila() {
    filaImpressao = [];
    atualizarFila();
}

function adicionarTodos() {
    let lista = getDadosAtuais();
    // Adiciona apenas quem ainda não está na fila para não duplicar
    lista.forEach(p => adicionarFila(p));
    atualizarFila();
}

// Nova função para remover 1 item específico
function removerDaFila(index) {
    filaImpressao.splice(index, 1);
    atualizarFila();
}

function atualizarFila() {
    document.getElementById("contadorFila").innerText = filaImpressao.length + " etiquetas";
    const ul = document.getElementById("listaFila");
    ul.innerHTML = "";

    if (filaImpressao.length === 0) {
        ul.innerHTML = '<li class="empty-msg" style="justify-content: center; color: #777;">Fila vazia.</li>';
    } else {
        filaImpressao.forEach((p, i) => {
            let nome = p['NOME DO PACIENTE'] || 'Sem Nome';
            let leito = p['LEITO'] || '?';

            // Item da lista com botão "X" para remover
            ul.innerHTML += `
                <li>
                    <span>✅ <b>${leito}</b> - ${nome}</span>
                    <button class="btn-remove-queue" onclick="removerDaFila(${i})" title="Remover este paciente">
                        <span class="material-icons">close</span>
                    </button>
                </li>
            `;
        });
    }
}

async function imprimirFila() {
    if (filaImpressao.length === 0) { alert("Fila vazia."); return; }
    let msg = await pywebview.api.imprimir_etiquetas(filaImpressao);
    if (msg !== "Cancelado.") alert(msg);
}

// --- RELATÓRIOS ---
async function gerarRelatorioSimples() {
    let msg;
    if (setorAtual === 'ENF') msg = await pywebview.api.gerar_relatorio_enf('simples');
    else if (setorAtual === 'UTI') msg = await pywebview.api.gerar_relatorio_uti('simples');
    else msg = await pywebview.api.gerar_relatorio_upa('simples');
    if (msg !== "Cancelado.") alert(msg);
}

async function gerarMapaGeral() {
    let msg;
    if (setorAtual === 'ENF') msg = await pywebview.api.gerar_relatorio_enf('geral');
    else if (setorAtual === 'UTI') msg = await pywebview.api.gerar_relatorio_uti('geral');
    else msg = await pywebview.api.gerar_relatorio_upa('geral');
    if (msg !== "Cancelado.") alert(msg);
}