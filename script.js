let atendimentos = [];
let dadosFiltrados = [];
let grafico = null;


// ==========================================
// IMPORTAR XML
// ==========================================

const arquivoXML = document.getElementById("arquivoXML");

if (arquivoXML) {

    arquivoXML.addEventListener("change", function (event) {

        const arquivo = event.target.files[0];

        if (!arquivo) {
            return;
        }

        const leitor = new FileReader();

        leitor.onload = function (evento) {

            const textoXML = evento.target.result;

            lerXML(textoXML);

        };

        leitor.onerror = function () {

            alert("Não foi possível ler o arquivo XML.");

        };

        leitor.readAsText(arquivo);

    });

}


// ==========================================
// LER XML
// ==========================================

function lerXML(texto) {

    const parser = new DOMParser();

    const xml = parser.parseFromString(
        texto,
        "application/xml"
    );


    // Verifica se o XML possui erro

    const erro = xml.querySelector("parsererror");

    if (erro) {

        alert("O XML está inválido.");

        console.error("Erro no XML:", erro.textContent);

        return;

    }


    // ==========================================
    // PROCURA OS REGISTROS
    // ==========================================

    const elementos =
        xml.querySelectorAll("atendimento");


    atendimentos = [];


    elementos.forEach(function (elemento) {

        // ------------------------------------------
        // DATA
        // ------------------------------------------

        const elementoData =
            elemento.querySelector("data");


        // ------------------------------------------
        // TOTAL
        // ------------------------------------------

        const elementoTotal =
            elemento.querySelector("total");


        const data =
            elementoData
                ? elementoData.textContent.trim()
                : "";


        const textoTotal =
            elementoTotal
                ? elementoTotal.textContent.trim()
                : "";


        // ------------------------------------------
        // CONVERTE O TOTAL
        // ------------------------------------------

        const total =
            converterTotal(textoTotal);


        // ------------------------------------------
        // ADICIONA O REGISTRO
        // ------------------------------------------

        if (data !== "") {

            atendimentos.push({

                data: data,

                total: total

            });

        }

    });


    // ==========================================
    // VERIFICA SE ENCONTROU OS DADOS
    // ==========================================

    if (atendimentos.length === 0) {

        alert(
            "Nenhum registro encontrado no XML."
        );

        console.log("XML:", xml);

        return;

    }


    // ==========================================
    // MOSTRA NO CONSOLE PARA CONFERÊNCIA
    // ==========================================

    console.log(
        "Registros encontrados:",
        atendimentos
    );


    // ==========================================
    // COPIA OS DADOS
    // ==========================================

    dadosFiltrados =
        [...atendimentos];


    // ==========================================
    // ATUALIZA DASHBOARD
    // ==========================================

    atualizarDashboard();

}


// ==========================================
// CONVERTER TOTAL DA PLANILHA
// ==========================================

function converterTotal(valor) {

    if (
        valor === null ||
        valor === undefined ||
        valor === ""
    ) {

        return 0;

    }


    let texto =
        String(valor).trim();


    // Remove espaços

    texto =
        texto.replace(/\s/g, "");


    // ------------------------------------------
    // CASO:
    // 1.234,56
    // ------------------------------------------

    if (
        texto.includes(".") &&
        texto.includes(",")
    ) {

        texto =
            texto
                .replace(/\./g, "")
                .replace(",", ".");

    }

    // ------------------------------------------
    // CASO:
    // 2,5
    // ------------------------------------------

    else if (texto.includes(",")) {

        texto =
            texto.replace(",", ".");

    }


    const numero =
        Number(texto);


    if (Number.isNaN(numero)) {

        console.warn(
            "Valor de Total inválido:",
            valor
        );

        return 0;

    }


    return numero;

}


// ==========================================
// ATUALIZAR DASHBOARD
// ==========================================

function atualizarDashboard() {

    atualizarCards();

    atualizarResumo();

    atualizarTabela();

    atualizarGrafico();

}


// ==========================================
// CARDS
// ==========================================

function atualizarCards() {

    // Quantidade de linhas da planilha

    const totalRegistros =
        dadosFiltrados.length;


    // Soma da coluna TOTAL

    const totalAtendimentos =
        dadosFiltrados.reduce(
            function (soma, item) {

                return soma + item.total;

            },
            0
        );


    // Média por dia

    const media =
        totalRegistros > 0
            ? totalAtendimentos / totalRegistros
            : 0;


    // Maior valor da coluna TOTAL

    const maior =
        totalRegistros > 0
            ? Math.max(
                ...dadosFiltrados.map(
                    item => item.total
                )
            )
            : 0;


    // ------------------------------------------
    // TOTAL DE REGISTROS
    // ------------------------------------------

    const totalRegistrosElemento =
        document.getElementById(
            "totalRegistros"
        );


    if (totalRegistrosElemento) {

        totalRegistrosElemento.textContent =
            totalRegistros;

    }


    // ------------------------------------------
    // TOTAL DE ATENDIMENTOS
    // ------------------------------------------

    const totalAtendimentosElemento =
        document.getElementById(
            "totalAtendimentos"
        );


    if (totalAtendimentosElemento) {

        totalAtendimentosElemento.textContent =
            totalAtendimentos;

    }


    // ------------------------------------------
    // MÉDIA
    // ------------------------------------------

    const mediaElemento =
        document.getElementById(
            "mediaDia"
        );


    if (mediaElemento) {

        mediaElemento.textContent =
            media.toFixed(1);

    }


    // ------------------------------------------
    // MAIOR ATENDIMENTO
    // ------------------------------------------

    const maiorElemento =
        document.getElementById(
            "maiorAtendimento"
        );


    if (maiorElemento) {

        maiorElemento.textContent =
            maior;

    }

}


// ==========================================
// RESUMO
// ==========================================

function atualizarResumo() {

    if (dadosFiltrados.length === 0) {

        return;

    }


    // Cria uma cópia para ordenar

    const ordenados =
        [...dadosFiltrados].sort(
            function (a, b) {

                return (
                    converterData(a.data) -
                    converterData(b.data)
                );

            }
        );


    const primeiro =
        ordenados[0];


    const ultimo =
        ordenados[ordenados.length - 1];


    // Procura o maior atendimento

    const maior =
        [...dadosFiltrados].sort(
            function (a, b) {

                return b.total - a.total;

            }
        )[0];


    // ------------------------------------------
    // PRIMEIRA DATA
    // ------------------------------------------

    const primeiraData =
        document.getElementById(
            "primeiraData"
        );


    if (primeiraData) {

        primeiraData.textContent =
            primeiro.data;

    }


    // ------------------------------------------
    // ÚLTIMA DATA
    // ------------------------------------------

    const ultimaData =
        document.getElementById(
            "ultimaData"
        );


    if (ultimaData) {

        ultimaData.textContent =
            ultimo.data;

    }


    // ------------------------------------------
    // DIA COM MAIS ATENDIMENTOS
    // ------------------------------------------

    const diaMaior =
        document.getElementById(
            "diaMaior"
        );


    if (diaMaior) {

        diaMaior.textContent =
            `${maior.data} — ${maior.total} atendimentos`;

    }


    // ------------------------------------------
    // PERÍODO
    // ------------------------------------------

    const periodo =
        document.getElementById(
            "periodo"
        );


    if (periodo) {

        periodo.textContent =
            `${primeiro.data} até ${ultimo.data}`;

    }

}


// ==========================================
// TABELA
// ==========================================

function atualizarTabela() {

    const tabela =
        document.getElementById(
            "tabela"
        );


    if (!tabela) {

        return;

    }


    tabela.innerHTML = "";


    if (dadosFiltrados.length === 0) {

        tabela.innerHTML = `
            <tr>
                <td colspan="3">
                    Nenhum dado encontrado.
                </td>
            </tr>
        `;

        return;

    }


    // Maior valor para calcular a barra

    const maior =
        Math.max(
            ...dadosFiltrados.map(
                item => item.total
            )
        );


    dadosFiltrados.forEach(
        function (item) {

            const tr =
                document.createElement("tr");


            const porcentagem =
                maior > 0
                    ? (item.total / maior) * 100
                    : 0;


            tr.innerHTML = `

                <td>
                    <strong>
                        ${item.data}
                    </strong>
                </td>

                <td>
                    ${item.total}
                </td>

                <td>

                    <div
                        class="barra"
                        style="width: ${porcentagem}%"
                    ></div>

                </td>

            `;


            tabela.appendChild(tr);

        }
    );

}


// ==========================================
// GRÁFICO
// ==========================================

function atualizarGrafico() {

    const canvas =
        document.getElementById(
            "grafico"
        );


    if (!canvas) {

        return;

    }


    const labels =
        dadosFiltrados.map(
            item => item.data
        );


    const valores =
        dadosFiltrados.map(
            item => item.total
        );


    // Remove gráfico anterior

    if (grafico) {

        grafico.destroy();

    }


    // Cria gráfico novo

    grafico =
        new Chart(
            canvas,
            {

                type: "bar",

                data: {

                    labels: labels,

                    datasets: [

                        {

                            label:
                                "Atendimentos",

                            data: valores

                        }

                    ]

                },

                options: {

                    responsive: true,

                    maintainAspectRatio: false,

                    scales: {

                        y: {

                            beginAtZero: true,

                            ticks: {

                                precision: 0

                            }

                        }

                    }

                }

            }
        );

}


// ==========================================
// CONVERTER DATA
// ==========================================

function converterData(data) {

    if (!data) {

        return new Date(NaN);

    }


    const texto =
        String(data).trim();


    // ------------------------------------------
    // FORMATO DA PLANILHA:
    // DD/MM/YYYY
    // ------------------------------------------

    const partes =
        texto.split("/");


    if (partes.length === 3) {

        return new Date(

            Number(partes[2]),

            Number(partes[1]) - 1,

            Number(partes[0])

        );

    }


    // ------------------------------------------
    // FORMATO YYYY-MM-DD
    // ------------------------------------------

    if (
        /^\d{4}-\d{2}-\d{2}$/.test(texto)
    ) {

        const partesISO =
            texto.split("-");


        return new Date(

            Number(partesISO[0]),

            Number(partesISO[1]) - 1,

            Number(partesISO[2])

        );

    }


    return new Date(texto);

}


// ==========================================
// FILTRO POR DATA
// ==========================================

const btnFiltrar =
    document.getElementById(
        "btnFiltrar"
    );


if (btnFiltrar) {

    btnFiltrar.addEventListener(
        "click",
        aplicarFiltro
    );

}


function aplicarFiltro() {

    const campoInicial =
        document.getElementById(
            "dataInicial"
        );


    const campoFinal =
        document.getElementById(
            "dataFinal"
        );


    const inicial =
        campoInicial
            ? campoInicial.value
            : "";


    const final =
        campoFinal
            ? campoFinal.value
            : "";


    dadosFiltrados =
        atendimentos.filter(
            function (item) {

                const data =
                    converterData(item.data);


                const inicio =
                    inicial
                        ? new Date(
                            inicial +
                            "T00:00:00"
                        )
                        : null;


                const fim =
                    final
                        ? new Date(
                            final +
                            "T23:59:59"
                        )
                        : null;


                if (
                    inicio &&
                    data < inicio
                ) {

                    return false;

                }


                if (
                    fim &&
                    data > fim
                ) {

                    return false;

                }


                return true;

            }
        );


    atualizarDashboard();

}


// ==========================================
// LIMPAR FILTRO
// ==========================================

const btnLimpar =
    document.getElementById(
        "btnLimpar"
    );


if (btnLimpar) {

    btnLimpar.addEventListener(
        "click",
        function () {

            const campoInicial =
                document.getElementById(
                    "dataInicial"
                );


            const campoFinal =
                document.getElementById(
                    "dataFinal"
                );


            if (campoInicial) {

                campoInicial.value = "";

            }


            if (campoFinal) {

                campoFinal.value = "";

            }


            dadosFiltrados =
                [...atendimentos];


            atualizarDashboard();

        }
    );

}


// ==========================================
// PESQUISA
// ==========================================

const pesquisa =
    document.getElementById(
        "pesquisa"
    );


if (pesquisa) {

    pesquisa.addEventListener(
        "input",
        function () {

            const termo =
                this.value
                    .toLowerCase()
                    .trim();


            const linhas =
                document.querySelectorAll(
                    "#tabela tr"
                );


            linhas.forEach(
                function (linha) {

                    const texto =
                        linha.textContent
                            .toLowerCase();


                    linha.style.display =
                        texto.includes(termo)
                            ? ""
                            : "none";

                }
            );

        }
    );

}