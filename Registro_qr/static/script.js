const readerDiv = document.getElementById("reader");
const formContainer = document.getElementById("formContainer");

let currentUser = null;

function dataHoje() {
    return new Date().toISOString().slice(0, 10);
}

function iniciarLeitor() {
    const html5QrCode = new Html5Qrcode("reader");

    html5QrCode.start(
        { facingMode: "environment" },
        { fps: 10, qrbox: 250 },
        async (decodedText) => {
            try {
                const obj = JSON.parse(decodedText);
                if (!obj.id || !obj.nome) throw "QR inválido";
                currentUser = obj;

                // Verifica status no servidor
                const checkResp = await fetch("/verificar", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ id: obj.id, nome: obj.nome }) // ID já é int do banco, não precisa de parseInt aqui
                });
                const status = await checkResp.json();

                let form; // Declara a variável form aqui para ser acessível em ambos os blocos
                if (status.aberto) {
                    // TELA DE FINALIZAÇÃO
                    formContainer.innerHTML = `
                        <div style="text-align: left;">
                            <p style="color: #3C1276; font-size: 1.2rem;">Atividade em Andamento</p>
                            <p><strong>Funcionário:</strong> ${obj.nome}</p>
                            <p><strong>Área:</strong> ${status.area}</p>
                            <p><strong>Projeto:</strong> ${status.projeto}</p>
                            <p><strong>Início:</strong> ${status.inicio} (${status.data})</p>
                            <p style="background: #e6edf3; padding: 10px; border-radius: 8px;">
                                ⏱️ <strong>Tempo decorrido:</strong><br>${status.tempo_decorrido}
                            </p>
                            <form id="registroForm">
                                <input type="hidden" id="tipoAcao" value="finalizar">
                                <button type="submit" style="width: 100%; background-color: #ff6b6b;">Finalizar Atividade</button>
                            </form>
                        </div>
                    `;
                    form = document.getElementById("registroForm");
                } else {
                    // TELA DE INÍCIO
                    formContainer.innerHTML = `
                        <p>Olá, ${obj.nome}. Selecione o projeto para iniciar.</p>
                        <form id="registroForm" novalidate>
                            <input type="hidden" id="tipoAcao" value="iniciar">
                            <label for="area">Área:</label>
                            <select id="area" name="area" required>
                                <option value="" disabled selected>Selecionar Área</option>
                                ${AREAS.map(area => `<option value="${area}">${area}</option>`).join('')}
                            </select>

                            <label for="projeto">Projeto:</label>
                            <select id="projeto" name="projeto" required>
                                <option value="" disabled selected>Selecionar Projeto</option>
                                ${PROJETOS.map(proj => `<option value="${proj}">${proj}</option>`).join('')}
                            </select>
                            <button type="submit" disabled>Iniciar Atividade</button>
                        </form>
                    `;
                    form = document.getElementById("registroForm");
                    const areaSelect = document.getElementById("area");
                    const projetoSelect = document.getElementById("projeto");
                    const submitButton = form.querySelector("button[type='submit']");

                    function validarFormulario() {
                        const valido = areaSelect.value !== "" && projetoSelect.value !== "";
                        submitButton.disabled = !valido;
                    }

                    areaSelect.addEventListener("change", validarFormulario);
                    projetoSelect.addEventListener("change", validarFormulario);
                    validarFormulario(); // Validação inicial
                }

                // Adiciona o event listener para o submit do formulário, independentemente da tela
                form.addEventListener("submit", enviarFormulario);

                formContainer.style.display = "block";
                html5QrCode.stop();
            } catch (e) {
                console.error("QR inválido:", e);
            }
        },
        (error) => { /* Ignora erros */ }
    ).catch(err => {
        console.error("Erro ao iniciar câmera:", err);
    });
}

async function enviarFormulario(e) {
    e.preventDefault();

    const tipoAcao = document.getElementById("tipoAcao").value;
    const payload = {
        id: currentUser.id, // ID já é int do banco, não precisa de parseInt aqui
        nome: currentUser.nome,
        tipoAcao: tipoAcao,
        area: document.getElementById("area") ? document.getElementById("area").value : null,
        projeto: document.getElementById("projeto") ? document.getElementById("projeto").value : null
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
    iniciarLeitor();
}

window.onload = () => {
    iniciarLeitor();
};

// Admin login functionality
document.getElementById('adminBtn').addEventListener('click', () => {
    document.getElementById('loginModal').style.display = 'flex';
});

document.getElementById('cancelBtn').addEventListener('click', () => {
    document.getElementById('loginModal').style.display = 'none';
});

document.getElementById('registerBtn').addEventListener('click', () => {
    document.getElementById('registerModal').style.display = 'flex';
});

document.getElementById('publicRegisterForm').addEventListener('submit', async (e) => {
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

document.getElementById('loginForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const username = document.getElementById('username').value;
    const password = document.getElementById('password').value;
    const response = await fetch('/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username, password })
    });
    const result = await response.json();
    if (result.status === 'ok') {
        window.location.href = '/admin';
    } else {
        alert(result.message || 'Credenciais inválidas');
    }
});
