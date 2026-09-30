let registros = [];
let clientes = [];
let grafico = null;
let paginaAtual = 1;
let registrosExcluidos = 0;
const ITENS_POR_PAGINA = 10;


const arquivoPlanilha = document.getElementById("arquivoPlanilha");

if (arquivoPlanilha) {

    arquivoPlanilha.addEventListener("change", async function (event) {

        const arquivos = [...event.target.files];

        event.target.value = "";

        if (arquivos.length === 0) {
            return;
        }

        const todos = [];
        const excluidos = [];
        const falhas = [];

        for (const arquivo of arquivos) {

            if (ehAtendimentoTelWhats(arquivo.name)) {
                continue;
            }

            try {

                const buffer = await lerArquivo(arquivo);
                const resultado = lerPlanilha(buffer, arquivo.name);

                if (resultado.itens.length === 0 && resultado.excluidos.length === 0) {
                    falhas.push(arquivo.name);
                    continue;
                }

                todos.push(...resultado.itens);
                excluidos.push(...resultado.excluidos);

            } catch (erro) {

                console.error(erro);
                falhas.push(arquivo.name);

            }

        }

        if (todos.length === 0 && excluidos.length === 0) {

            alert(
                "Não encontrei atendimentos nessa planilha. Use o Excel de Tempos de Atendimento exportado pelo TomTicket."
            );

            return;

        }

        registros = deduplicar(todos).filter(function (item) {
            return !ehAtendimentoTelWhats(item.cliente);
        });

        const protocolosIncluidos = new Set(registros.map(function (item) {
            return item.protocolo;
        }));

        registrosExcluidos = deduplicar(excluidos).filter(function (item) {
            return !protocolosIncluidos.has(item.protocolo);
        }).length;

        paginaAtual = 1;
        atualizarDashboard();

        if (falhas.length > 0) {

            alert(
                "Estes arquivos não entraram no dashboard:\n" +
                falhas.join("\n")
            );

        }

    });

}


function lerArquivo(arquivo) {

    return new Promise(function (resolve, reject) {

        const leitor = new FileReader();

        leitor.onload = function () {
            resolve(leitor.result);
        };

        leitor.onerror = function () {
            reject(new Error("Não foi possível ler o arquivo."));
        };

        leitor.readAsArrayBuffer(arquivo);

    });

}


function lerPlanilha(buffer, nomeArquivo) {

    if (typeof XLSX === "undefined") {
        throw new Error("A biblioteca de planilhas não carregou.");
    }

    const livro = XLSX.read(buffer, { type: "array" });
    const nomeAba = livro.SheetNames.find(function (nome) {
        return !ehAtendimentoTelWhats(nome);
    });

    if (!nomeAba) {
        return { itens: [], excluidos: [] };
    }

    const aba = livro.Sheets[nomeAba];

    const linhas = XLSX.utils.sheet_to_json(aba, {
        header: 1,
        raw: true,
        defval: ""
    });

    const periodo =
        extrairPeriodo(linhas) ||
        periodoDoNome(nomeArquivo);

    let indiceCabecalho = -1;

    for (let i = 0; i < linhas.length; i++) {

        const cabecalho = linhas[i].map(function (celula) {
            return normalizar(celula);
        });

        if (
            cabecalho.includes("protocolo") &&
            cabecalho.includes("cliente")
        ) {
            indiceCabecalho = i;
            break;
        }

    }

    if (indiceCabecalho === -1) {
        return { itens: [], excluidos: [] };
    }

    const cabecalho = linhas[indiceCabecalho].map(function (celula) {
        return normalizar(celula);
    });

    const colunas = {
        protocolo: cabecalho.indexOf("protocolo"),
        cliente: cabecalho.indexOf("cliente"),
        espera: cabecalho.findIndex(function (nome) {
            return nome.includes("espera");
        }),
        conversa: cabecalho.findIndex(function (nome) {
            return nome.includes("conversa");
        })
    };

    const itens = [];
    const excluidos = [];

    for (let i = indiceCabecalho + 1; i < linhas.length; i++) {

        const linha = linhas[i];
        const protocolo = textoCelula(linha[colunas.protocolo]);

        if (protocolo === "") {
            continue;
        }

        if (normalizar(protocolo).includes("tempomedio")) {
            continue;
        }

        const cliente = textoCelula(linha[colunas.cliente]);

        if (normalizar(cliente).includes("tempomedio")) {
            continue;
        }

        if (ehAtendimentoTelWhats(cliente)) {
            excluidos.push({
                protocolo: protocolo,
                cliente: cliente
            });
            continue;
        }

        const espera = linha[colunas.espera];
        const conversa = linha[colunas.conversa];

        itens.push({
            protocolo: protocolo,
            cliente: cliente,
            espera: tempoParaSegundos(espera),
            conversa: tempoParaSegundos(conversa),
            temEspera: celulaPreenchida(espera),
            temConversa: celulaPreenchida(conversa),
            inicio: periodo ? periodo.inicio : "",
            fim: periodo ? periodo.fim : ""
        });

    }

    return {
        itens: itens,
        excluidos: excluidos
    };

}


function deduplicar(itens) {

    const vistos = new Set();
    const unicos = [];

    itens.forEach(function (item) {

        if (vistos.has(item.protocolo)) {
            return;
        }

        vistos.add(item.protocolo);
        unicos.push(item);

    });

    return unicos;

}


function agruparClientes(itens) {

    const mapa = new Map();

    itens.forEach(function (item) {

        if (ehAtendimentoTelWhats(item.cliente)) {
            return;
        }

        const nome =
            item.cliente !== ""
                ? item.cliente
                : "Não informado";

        const chave = normalizar(nome);

        if (!mapa.has(chave)) {

            mapa.set(chave, {
                cliente: nome,
                atendimentos: 0,
                conversa: 0,
                espera: 0,
                comConversa: 0,
                comEspera: 0
            });

        }

        const atual = mapa.get(chave);

        atual.atendimentos += 1;

        if (item.temConversa) {
            atual.conversa += item.conversa;
            atual.comConversa += 1;
        }

        if (item.temEspera) {
            atual.espera += item.espera;
            atual.comEspera += 1;
        }

    });

    return [...mapa.values()];

}


function atualizarDashboard() {

    clientes = agruparClientes(registros);
    clientes.sort(compararMais);

    atualizarCards();
    atualizarResumo();
    atualizarTabela();
    atualizarGrafico();

}


function atualizarCards() {

    definirTexto("totalAtendimentos", registros.length);
    definirTexto("totalClientes", clientes.length);

    if (clientes.length === 0) {

        definirTexto("clienteMais", "—");
        definirTexto("clienteMaisDetalhe", "Nenhum cliente na planilha");
        definirTexto("clienteMenos", "—");
        definirTexto("clienteMenosDetalhe", "Nenhum cliente na planilha");

        return;

    }

    const mais = clientes[0];
    const menos = [...clientes].sort(compararMenos)[0];

    definirTexto("clienteMais", mais.cliente);
    definirTexto(
        "clienteMaisDetalhe",
        descreverCliente(mais, clientes, "mais")
    );

    definirTexto("clienteMenos", menos.cliente);
    definirTexto(
        "clienteMenosDetalhe",
        descreverCliente(menos, clientes, "menos")
    );

}


function descreverCliente(cliente, lista, tipo) {

    const texto =
        plural(cliente.atendimentos, "atendimento", "atendimentos") +
        " · " +
        formatarDuracao(cliente.conversa) +
        " de conversa";

    if (lista.length === 1) {
        return texto + " · único cliente no período";
    }

    const empates = lista.filter(function (item) {
        return item.atendimentos === cliente.atendimentos;
    });

    if (empates.length > 1 && tipo === "mais") {
        return texto + " · maior conversa entre " + empates.length;
    }

    if (empates.length > 1 && tipo === "menos") {
        return texto + " · menor conversa entre " + empates.length;
    }

    return texto;

}


function atualizarResumo() {

    definirTexto("mediaEspera", mediaDuracao("espera", "temEspera"));
    definirTexto("medianaEspera", medianaDuracao("espera", "temEspera"));
    definirTexto("p90Espera", percentil90Duracao("espera", "temEspera"));
    definirTexto("mediaConversa", mediaDuracao("conversa", "temConversa"));
    definirTexto("medianaConversa", medianaDuracao("conversa", "temConversa"));
    definirTexto("p90Conversa", percentil90Duracao("conversa", "temConversa"));

    const periodo = document.getElementById("periodo");

    if (periodo) {
        periodo.textContent = textoPeriodo();
    }

    const notaExclusoes = document.getElementById("notaExclusoes");

    if (notaExclusoes) {
        const total = registros.length + registrosExcluidos;
        notaExclusoes.hidden = registrosExcluidos === 0;
        notaExclusoes.textContent =
            plural(registrosExcluidos, "registro", "registros") +
            " de ATENDIMENTO - TEL - WHATS excluídos · " +
            registros.length + " de " + total + " analisados";
    }

}


function mediaDuracao(campo, flag) {

    const comValor = registros.filter(function (item) {
        return item[flag];
    });

    if (comValor.length === 0) {
        return "-";
    }

    const soma = comValor.reduce(function (total, item) {
        return total + item[campo];
    }, 0);

    return formatarDuracao(soma / comValor.length);

}


function duracoesOrdenadas(campo, flag) {

    return registros
        .filter(function (item) {
            return item[flag];
        })
        .map(function (item) {
            return item[campo];
        })
        .sort(function (a, b) {
            return a - b;
        });

}


function medianaDuracao(campo, flag) {

    const valores = duracoesOrdenadas(campo, flag);

    if (valores.length === 0) {
        return "-";
    }

    const meio = Math.floor(valores.length / 2);
    const mediana = valores.length % 2 === 0
        ? (valores[meio - 1] + valores[meio]) / 2
        : valores[meio];

    return formatarDuracao(mediana);

}


function percentil90Duracao(campo, flag) {

    const valores = duracoesOrdenadas(campo, flag);

    if (valores.length === 0) {
        return "-";
    }

    const indice = Math.ceil(valores.length * 0.9) - 1;

    return formatarDuracao(valores[indice]);

}


function textoPeriodo() {

    const inicios = registros
        .map(function (item) {
            return item.inicio;
        })
        .filter(Boolean);

    const fins = registros
        .map(function (item) {
            return item.fim;
        })
        .filter(Boolean);

    if (inicios.length === 0) {
        return plural(registros.length, "atendimento importado", "atendimentos importados");
    }

    inicios.sort(function (a, b) {
        return converterData(a) - converterData(b);
    });

    fins.sort(function (a, b) {
        return converterData(a) - converterData(b);
    });

    return inicios[0] + " até " + fins[fins.length - 1];

}


function clientesFiltrados() {

    const campo = document.getElementById("pesquisa");
    const termo = campo ? campo.value.toLowerCase().trim() : "";

    if (termo === "") {
        return clientes;
    }

    return clientes.filter(function (item) {
        return item.cliente.toLowerCase().includes(termo);
    });

}


function atualizarTabela() {

    const tabela = document.getElementById("tabela");

    if (!tabela) {
        return;
    }

    tabela.innerHTML = "";

    if (clientes.length === 0) {

        tabela.innerHTML = `
            <tr>
                <td colspan="5" class="vazio">
                    Nenhum cliente encontrado.
                </td>
            </tr>
        `;

        atualizarPaginacao([]);
        return;

    }

    const filtrados = clientesFiltrados();

    if (filtrados.length === 0) {

        tabela.innerHTML = `
            <tr>
                <td colspan="5" class="vazio">
                    Nenhum cliente corresponde à pesquisa.
                </td>
            </tr>
        `;

        atualizarPaginacao([]);
        return;

    }

    const totalPaginas = Math.max(1, Math.ceil(filtrados.length / ITENS_POR_PAGINA));

    if (paginaAtual > totalPaginas) {
        paginaAtual = totalPaginas;
    }

    if (paginaAtual < 1) {
        paginaAtual = 1;
    }

    const inicio = (paginaAtual - 1) * ITENS_POR_PAGINA;
    const pagina = filtrados.slice(inicio, inicio + ITENS_POR_PAGINA);
    const maior = clientes[0].atendimentos;

    pagina.forEach(function (item) {

        const tr = document.createElement("tr");
        const porcentagem =
            maior > 0
                ? (item.atendimentos / maior) * 100
                : 0;

        tr.innerHTML = `
            <td>
                <div class="cliente">
                    <span class="avatar">${iniciais(item.cliente)}</span>
                    <strong>${escapar(item.cliente)}</strong>
                </div>
            </td>
            <td class="num">${item.atendimentos}</td>
            <td>${formatarDuracao(item.conversa)}</td>
            <td>${formatarDuracao(item.espera)}</td>
            <td>
                <div class="trilho">
                    <div class="barra" style="width: ${porcentagem}%"></div>
                </div>
            </td>
        `;

        tabela.appendChild(tr);

    });

    atualizarPaginacao(filtrados);

}


function atualizarPaginacao(filtrados) {

    const barra = document.getElementById("paginacao");
    const info = document.getElementById("paginacaoInfo");
    const paginas = document.getElementById("paginacaoPaginas");
    const anterior = document.getElementById("paginaAnterior");
    const proxima = document.getElementById("paginaProxima");

    if (!barra || !info || !paginas || !anterior || !proxima) {
        return;
    }

    if (filtrados.length === 0) {
        barra.hidden = true;
        paginas.innerHTML = "";
        return;
    }

    const totalPaginas = Math.ceil(filtrados.length / ITENS_POR_PAGINA);
    const inicio = (paginaAtual - 1) * ITENS_POR_PAGINA + 1;
    const fim = Math.min(paginaAtual * ITENS_POR_PAGINA, filtrados.length);

    barra.hidden = false;
    info.textContent =
        "Mostrando " + inicio + "–" + fim + " de " + filtrados.length;

    anterior.disabled = paginaAtual <= 1;
    proxima.disabled = paginaAtual >= totalPaginas;

    paginas.innerHTML = "";

    numerosPagina(paginaAtual, totalPaginas).forEach(function (numero) {

        if (numero === "…") {

            const reticencias = document.createElement("span");
            reticencias.className = "paginacao-reticencias";
            reticencias.textContent = "…";
            paginas.appendChild(reticencias);
            return;

        }

        const botao = document.createElement("button");
        botao.type = "button";
        botao.textContent = String(numero);

        if (numero === paginaAtual) {
            botao.setAttribute("aria-current", "page");
        }

        botao.addEventListener("click", function () {
            paginaAtual = numero;
            atualizarTabela();
        });

        paginas.appendChild(botao);

    });

}


function numerosPagina(atual, total) {

    if (total <= 7) {

        return Array.from({ length: total }, function (_, i) {
            return i + 1;
        });

    }

    const itens = [1];

    if (atual > 3) {
        itens.push("…");
    }

    const inicio = Math.max(2, atual - 1);
    const fim = Math.min(total - 1, atual + 1);

    for (let i = inicio; i <= fim; i++) {
        itens.push(i);
    }

    if (atual < total - 2) {
        itens.push("…");
    }

    itens.push(total);

    return itens;

}


function atualizarGrafico() {

    const canvas = document.getElementById("grafico");

    if (!canvas || typeof Chart === "undefined") {
        return;
    }

    const topo = clientes.slice(0, 10);

    if (grafico) {
        grafico.destroy();
    }

    grafico = new Chart(canvas, {
        type: "bar",
        data: {
            labels: topo.map(function (item) {
                return item.cliente;
            }),
            datasets: [
                {
                    label: "Atendimentos",
                    data: topo.map(function (item) {
                        return item.atendimentos;
                    }),
                    backgroundColor: "#7817ff",
                    borderRadius: 8,
                    borderSkipped: false,
                    maxBarThickness: 18
                }
            ]
        },
        options: {
            indexAxis: "y",
            responsive: true,
            maintainAspectRatio: false,
            onHover: function (evento) {
                if (evento && evento.native && evento.native.target) {
                    evento.native.target.style.cursor = "default";
                }
            },
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    backgroundColor: "#202936",
                    borderColor: "#3b4c68",
                    borderWidth: 1,
                    titleColor: "#ffffff",
                    bodyColor: "#dce2eb",
                    padding: 12,
                    cornerRadius: 10,
                    displayColors: false
                }
            },
            scales: {
                x: {
                    beginAtZero: true,
                    ticks: {
                        color: "#9ba8bc",
                        precision: 0,
                        font: {
                            family: "Inter, Arial, sans-serif",
                            size: 11
                        }
                    },
                    grid: {
                        color: "#303b4d"
                    },
                    border: {
                        display: false
                    }
                },
                y: {
                    ticks: {
                        color: "#dce2eb",
                        font: {
                            family: "Inter, Arial, sans-serif",
                            size: 12
                        }
                    },
                    grid: {
                        display: false
                    },
                    border: {
                        display: false
                    }
                }
            }
        }
    });

}


function compararMais(a, b) {

    if (b.atendimentos !== a.atendimentos) {
        return b.atendimentos - a.atendimentos;
    }

    return b.conversa - a.conversa;

}


function compararMenos(a, b) {

    if (a.atendimentos !== b.atendimentos) {
        return a.atendimentos - b.atendimentos;
    }

    return a.conversa - b.conversa;

}


function extrairPeriodo(linhas) {

    const texto = linhas
        .flat()
        .map(function (celula) {
            return String(celula);
        })
        .join(" ");

    const encontrado = texto.match(
        /Per[ií]odo:\s*(\d{2}\/\d{2}\/\d{4})\s*-\s*(\d{2}\/\d{2}\/\d{4})/i
    );

    if (!encontrado) {
        return null;
    }

    return {
        inicio: encontrado[1],
        fim: encontrado[2]
    };

}


function periodoDoNome(nome) {

    const encontrado = String(nome).match(
        /(\d{2})-(\d{2})-(\d{4})-(\d{2})-(\d{2})-(\d{4})/
    );

    if (!encontrado) {
        return null;
    }

    return {
        inicio: encontrado[1] + "/" + encontrado[2] + "/" + encontrado[3],
        fim: encontrado[4] + "/" + encontrado[5] + "/" + encontrado[6]
    };

}


function tempoParaSegundos(valor) {

    if (!celulaPreenchida(valor)) {
        return 0;
    }

    if (typeof valor === "number") {
        return valor * 86400;
    }

    const texto = String(valor).trim();
    const partes = texto.split(":");

    if (partes.length >= 2) {

        const horas = Number(partes[0]) || 0;
        const minutos = Number(partes[1]) || 0;
        const segundos = Number(partes[2]) || 0;

        return (horas * 3600) + (minutos * 60) + segundos;

    }

    const numero = Number(texto.replace(",", "."));

    if (Number.isNaN(numero)) {
        return 0;
    }

    return numero * 86400;

}


function formatarDuracao(segundos) {

    const total = Math.max(0, Math.round(segundos));
    const dias = Math.floor(total / 86400);
    const horas = Math.floor((total % 86400) / 3600);
    const minutos = Math.floor((total % 3600) / 60);
    const resto = total % 60;

    if (dias > 0) {
        return dias + "d " + horas + "h " + minutos + "min";
    }

    if (horas > 0) {
        return horas + "h " + minutos + "min";
    }

    if (minutos > 0) {
        return minutos + "min " + resto + "s";
    }

    return resto + "s";

}


function celulaPreenchida(valor) {

    return valor !== "" && valor !== null && valor !== undefined;

}


function textoCelula(valor) {

    if (!celulaPreenchida(valor)) {
        return "";
    }

    return String(valor).trim();

}


function normalizar(valor) {

    return String(valor)
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLowerCase()
        .replace(/[^a-z0-9]/g, "");

}


function ehAtendimentoTelWhats(valor) {

    const n = normalizar(valor);

    if (n.includes("atendimentotelwhats")) {
        return true;
    }

    const tokens = String(valor)
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLowerCase()
        .split(/[^a-z0-9]+/)
        .filter(Boolean);

    const temAtendimento = tokens.some(function (token) {
        return token === "atendimento" || token.indexOf("atendimento") === 0;
    });

    const temTel = tokens.some(function (token) {
        return token === "tel" || token === "telefone";
    });

    const temWhats = tokens.some(function (token) {
        return token.indexOf("whats") === 0;
    });

    return temAtendimento && temTel && temWhats;

}


function converterData(data) {

    const partes = String(data).split("/");

    if (partes.length !== 3) {
        return new Date(data);
    }

    return new Date(
        Number(partes[2]),
        Number(partes[1]) - 1,
        Number(partes[0])
    );

}


function plural(quantidade, singular, pluralTexto) {

    if (quantidade === 1) {
        return "1 " + singular;
    }

    return quantidade + " " + pluralTexto;

}


function iniciais(nome) {

    const partes = String(nome)
        .trim()
        .split(/\s+/)
        .filter(Boolean);

    if (partes.length === 0) {
        return "?";
    }

    if (partes.length === 1) {
        return partes[0].slice(0, 2).toUpperCase();
    }

    return (partes[0][0] + partes[partes.length - 1][0]).toUpperCase();

}


function escapar(texto) {

    return String(texto)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");

}


function definirTexto(id, valor) {

    const elemento = document.getElementById(id);

    if (elemento) {
        elemento.textContent = valor;
    }

}


const pesquisa = document.getElementById("pesquisa");

if (pesquisa) {

    pesquisa.addEventListener("input", function () {
        paginaAtual = 1;
        atualizarTabela();
    });

}

const paginaAnterior = document.getElementById("paginaAnterior");
const paginaProxima = document.getElementById("paginaProxima");

if (paginaAnterior) {

    paginaAnterior.addEventListener("click", function () {
        paginaAtual -= 1;
        atualizarTabela();
    });

}

if (paginaProxima) {

    paginaProxima.addEventListener("click", function () {
        paginaAtual += 1;
        atualizarTabela();
    });

}
