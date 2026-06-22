const readerDiv = document.getElementById("reader");
const formContainer = document.getElementById("formContainer");

let currentUser = null; // Stores the currently scanned employee
let funcionarioIdentificado = false; // New state variable to prevent re-scanning employee QR
let html5QrCodeScanner = null; // Global instance for the QR code scanner

function dataHoje() {
    return new Date().toISOString().slice(0, 10);
}

function iniciarLeitor() {
    // Ensure previous scanner is stopped if it exists and is running
    if (html5QrCodeScanner) {
        html5QrCodeScanner.stop().catch(err => console.warn("Erro ao parar scanner existente:", err));
    }

    html5QrCodeScanner = new Html5Qrcode("reader");

    html5QrCodeScanner.start(
        { facingMode: "environment" },
        { fps: 10, qrbox: 250 },
        async (decodedText) => {
            try {
                // Lógica de Processamento de QR Codes Padronizados
                if (decodedText.includes(':')) {
                    const [tipo, id] = decodedText.split(':');

                    if (tipo === 'FUNCIONARIO') {
                        if (!funcionarioIdentificado) {
                            handleFuncionarioScan(id);
                        } else {
                            console.log("Funcionário já identificado, ignorando novo scan de funcionário.");
                        }
                    } else if (currentUser && funcionarioIdentificado) {
                        // Só processa outros QRs se já tiver um funcionário identificado
                        handleSecondaryScan(tipo, id);
                    }
                    return;
                }

                // Fallback para JSON antigo se necessário
                const obj = JSON.parse(decodedText);
                if (obj.id && !funcionarioIdentificado) {
                    handleFuncionarioScan(obj.id, obj.nome);
                } else if (obj.id && funcionarioIdentificado) {
                    console.log("Funcionário já identificado (JSON), ignorando novo scan de funcionário.");
                }
            } catch (e) {
                console.error("QR inválido:", e);
            }
        },
    ).catch(err => console.error("Erro ao iniciar câmera:", err));
}

function getArrayGlobal(name) {
    return (typeof window[name] !== 'undefined' && Array.isArray(window[name])) ? window[name] : [];
}

async function handleFuncionarioScan(id, nomeManual = null) {
    try {
        console.log("handleFuncionarioScan id=", id, "nomeManual=", nomeManual);
        currentUser = { id: id, nome: nomeManual };

        const checkResp = await fetch("/verificar", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ id: id })
        });
        const status = await checkResp.json();
        console.log("/verificar retorno=", status);

        if (!currentUser.nome) currentUser.nome = status.nome;

        if (status.aberto) {
            // TELA DE FINALIZAÇÃO
            formContainer.innerHTML = `
                <div style="text-align: left;">
                    <p style="color: #3C1276; font-size: 1.2rem;">Atividade em Andamento</p>
                    <p><strong>Funcionário:</strong> ${currentUser.nome}</p>
                    <p><strong>Área:</strong> ${status.area}</p>
                    <p><strong>Projeto:</strong> ${status.projeto}</p>
                    <p><strong>Atividade:</strong> ${status.atividade}</p>
                    <p><strong>Início:</strong> ${status.inicio} (${status.data})</p>
                    <form id="registroForm">
                        <input type="hidden" id="tipoAcao" value="finalizar">
                        <button type="submit" style="width: 100%; background-color: #ff6b6b;">Finalizar Atividade</button>
                    </form>
                </div>
            `;

            if (html5QrCodeScanner) {
                html5QrCodeScanner.stop().catch(() => {});
                html5QrCodeScanner = null;
                funcionarioIdentificado = false;
            }
        } else {
            // TELA DE INÍCIO
            const AREAS_SAFE = getArrayGlobal('AREAS');
            const PROJETOS_SAFE = getArrayGlobal('PROJETOS');

            // Atividades vêm do backend (não do window), para evitar ficar vazio
            let ATIVIDADES_SAFE = [];
            try {
                const respAt = await fetch('/api/atividades');
                const dataAt = await respAt.json();
                ATIVIDADES_SAFE = Array.isArray(dataAt) ? dataAt.map(x => x.nome) : [];
            } catch (e) {
                console.warn('Falha ao carregar /api/atividades. Usando fallback window.A? (se existir).', e);
                ATIVIDADES_SAFE = getArrayGlobal('ATIVIDADES');
            }

            formContainer.innerHTML = `
                <p>Olá, <strong>${currentUser.nome}</strong>. Escaneie Área/Projeto/Atividade ou selecione abaixo.</p>
                <form id="registroForm" novalidate>
                    <input type="hidden" id="tipoAcao" value="iniciar">
                    <label for="area">Área:</label>
                    <select id="area" name="area" required>
                        <option value="" disabled selected>Selecionar Área</option>
                        ${AREAS_SAFE.map(area => `<option value="${area}">${area}</option>`).join('')}
                    </select>

                    <label for="projeto">Projeto:</label>
                    <select id="projeto" name="projeto" required>
                        <option value="" disabled selected>Selecionar Projeto</option>
                        ${PROJETOS_SAFE.map(proj => `<option value="${proj}">${proj}</option>`).join('')}
                    </select>

                    <label for="atividade">Atividade:</label>
                    <select id="atividade" name="atividade" required>
                        <option value="" disabled selected>Selecionar Atividade</option>
                        ${ATIVIDADES_SAFE.map(at => `<option value="${at}">${at}</option>`).join('')}
                    </select>

                    <button type="submit" disabled>Iniciar Atividade</button>
                </form>
            `;
        }

        const form = document.getElementById("registroForm");
        funcionarioIdentificado = true;
        form.addEventListener("submit", enviarFormulario);

        if (!status.aberto) {
            ['area', 'projeto', 'atividade'].forEach(id => {
                document.getElementById(id).addEventListener('change', validarFormulario);
            });
            validarFormulario();
        }

        formContainer.style.display = "block";
    } catch (e) {
        console.error(e);
    }
}

function handleSecondaryScan(tipo, id) {
    if (tipo === 'AREA') selecionarNoCampo('area', id);
    if (tipo === 'PROJETO') selecionarNoCampo('projeto', id);
    if (tipo === 'ATIVIDADE') selecionarNoCampo('atividade', id);
}

async function selecionarNoCampo(campoId, id) {
    const resp = await fetch(`/admin/get_name_by_id?tipo=${campoId}&id=${id}`);
    const data = await resp.json();
    if (data.nome) {
        const select = document.getElementById(campoId);
        select.value = data.nome;
        validarFormulario();

        select.style.borderColor = '#4CBDE0';
        setTimeout(() => select.style.borderColor = '#657788', 1000);
    }
}

function validarFormulario() {
    const area = document.getElementById("area").value;
    const projeto = document.getElementById("projeto").value;
    const atividade = document.getElementById("atividade").value;
    const btn = document.querySelector("#registroForm button");
    btn.disabled = !(area && projeto && atividade);
}

async function enviarFormulario(e) {
    e.preventDefault();

    const tipoAcao = document.getElementById("tipoAcao").value;
    const payload = {
        id: currentUser.id,
        nome: currentUser.nome,
        tipoAcao: tipoAcao,
        area: document.getElementById("area") ? document.getElementById("area").value : null,
        projeto: document.getElementById("projeto") ? document.getElementById("projeto").value : null,
        atividade: document.getElementById("atividade") ? document.getElementById("atividade").value : null
    };

    try {
        const resp = await fetch("/registrar", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload)
        });

        const j = await resp.json();

        if (j.status === "ok") {
            alert(j.message);
        } else {
            alert(j.message || "Erro ao registrar.");
        }
    } catch (err) {
        alert("Erro na requisição: " + err);
    }

    formContainer.style.display = "none";
    currentUser = null;
    funcionarioIdentificado = false;
    iniciarLeitor();
}

window.onload = () => {
    iniciarLeitor();
};

// Admin login functionality
const adminBtn = document.getElementById('adminBtn');
if (adminBtn) {
    adminBtn.addEventListener('click', () => {
        document.getElementById('loginModal').style.display = 'flex';
    });
}

const cancelBtn = document.getElementById('cancelBtn');
if (cancelBtn) {
    cancelBtn.addEventListener('click', () => {
        document.getElementById('loginModal').style.display = 'none';
    });
}

const registerBtn = document.getElementById('registerBtn');
if (registerBtn) {
    registerBtn.addEventListener('click', () => {
        document.getElementById('registerModal').style.display = 'flex';
    });
}

document.getElementById('publicRegisterForm')?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const username = document.getElementById('regUser').value;
    const email = document.getElementById('regEmail').value;
    const password = document.getElementById('regPass').value;

    const response = await fetch('/register_user', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username, password, email })
    });

    const result = await response.json();
    if (result.status === 'ok') {
        alert('Solicitação enviada com sucesso! Aguarde a validação do administrador.');
        document.getElementById('registerModal').style.display = 'none';
    } else {
        alert(result.message);
    }
});

document.getElementById('loginForm')?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const username = document.getElementById('username').value;
    const password = document.getElementById('password').value;

    const response = await fetch('/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username, password }),
        credentials: 'include'
    });


    const result = await response.json();
    if (result.status === 'ok') {
        window.location.href = '/admin';
    } else {
        alert(result.message || 'Credenciais inválidas');
    }
});

