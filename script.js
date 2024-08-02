let dados = [];

document.getElementById('inputArquivo').addEventListener('change', handleFile, false);

function handleFile(event) {
    const reader = new FileReader();
    const file = event.target.files[0];
    
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        dados = XLSX.utils.sheet_to_json(firstSheet, {header: 1});

        const headers = dados[0];
        dados = dados.slice(1).map(row => {
            let obj = {};
            headers.forEach((header, index) => {
                obj[header] = row[index];
            });
            return obj;
        });

        console.log(dados);
    };

    reader.readAsArrayBuffer(file);
}

function verificarCTE() {
    let divResultado = document.getElementById('resultado');
    let notasResultado = document.getElementById('notasResultado');
    let notasCargaResultado = document.getElementById('notasCargaResultado');
    let notasRepetidasResultado = document.getElementById('notasRepetidasResultado');
    divResultado.value = '';
    notasResultado.value = '';
    notasCargaResultado.value = '';
    notasRepetidasResultado.value = '';

    if (dados.length === 0) {
        divResultado.value = 'Nenhum dado carregado ou formato de planilha incorreto.';
        return;
    }

    let ctesDivergentes = [];
    let notasDivergentes = {};
    let notasRepetidas = {};
    let ctesPorCarga = {};
    let ctesDuplicadosPorCarga = {};

    const ctesCod = {};
    const ctesCarga = {};
    const notasCteMap = {};

    dados.forEach((item) => {
        const Cte = item['CTE'];
        const Cod = item['Cod'];
        const Carga = item['Carga'];
        const NotaFiscal = item['Nota Fiscal'];

        if (Cte && Cod) {
            if (!ctesCod[Cte]) {
                ctesCod[Cte] = new Set();
            }
            ctesCod[Cte].add(Cod);
        } else {
            console.error("Linha com dados faltantes:", item);
        }

        if (Cte && Carga) {
            if (!ctesCarga[Cte]) {
                ctesCarga[Cte] = {};
            }
            if (!ctesCarga[Cte][Carga]) {
                ctesCarga[Cte][Carga] = new Set();
            }
            ctesCarga[Cte][Carga].add(NotaFiscal);

            // Agrupar CTEs por carga
            if (!ctesPorCarga[Carga]) {
                ctesPorCarga[Carga] = new Set();
            }
            ctesPorCarga[Carga].add(Cte);
        }

        if (NotaFiscal) {
            if (!notasCteMap[NotaFiscal]) {
                notasCteMap[NotaFiscal] = [];
            }
            notasCteMap[NotaFiscal].push(Cte);
        }
    });

    for (let cte in ctesCod) {
        const cods = Array.from(ctesCod[cte]);
        if (cods.length > 1) {
            ctesDivergentes.push(cte);
        }
    }

    for (let cte in ctesCarga) {
        const cargas = Object.keys(ctesCarga[cte]);
        if (cargas.length > 1) {
            notasDivergentes[cte] = cargas.map(carga => {
                return `Carga ${carga}: ${Array.from(ctesCarga[cte][carga]).join(', ')}`;
            }).join('; ');
        }
    }

    // Mapeia os CTEs com notas fiscais repetidas
    const notasPorCarga = {};
    Object.entries(ctesCarga).forEach(([cte, cargas]) => {
        Object.entries(cargas).forEach(([carga, notas]) => {
            if (!notasPorCarga[carga]) {
                notasPorCarga[carga] = { ctes: new Set(), notas: new Set() };
            }
            notasPorCarga[carga].ctes.add(cte);
            notas.forEach(nota => {
                notasPorCarga[carga].notas.add(nota);
            });
        });
    });

    Object.entries(notasPorCarga).forEach(([carga, { ctes, notas }]) => {
        const notasArray = Array.from(notas);
        notasArray.forEach(nota => {
            if (notasCteMap[nota] && notasCteMap[nota].length > 1) {
                const sortedCtes = [...new Set(notasCteMap[nota])].sort((a, b) => a - b);
                const menorCte = sortedCtes[0];
                const outrosCtes = sortedCtes.slice(1);

                if (!ctesDuplicadosPorCarga[carga]) {
                    ctesDuplicadosPorCarga[carga] = {};
                }
                ctesDuplicadosPorCarga[carga][menorCte] = outrosCtes;
            }
        });
    });

    if (ctesDivergentes.length > 0) {
        divResultado.value = ctesDivergentes.join(', ');
    } else {
        divResultado.value = 'Nenhum Cte com valores diferentes de Cod encontrado.';
    }

    if (Object.keys(notasDivergentes).length > 0) {
        notasResultado.value = Object.entries(notasDivergentes).map(([cte, notas]) => {
            return `Cte ${cte}:\n${notas}`;
        }).join('\n\n');
    } else {
        notasResultado.value = 'Nenhum Cte com cargas diferentes encontrado.';
    }

    if (Object.keys(ctesPorCarga).length > 0) {
        notasCargaResultado.value = Object.entries(ctesPorCarga).map(([carga, { ctes }]) => {
            const ctesPorCargaTexto = `Carga ${carga}:\n${Array.from(ctes).join(', ')}`;
            const ctesDuplicadosPorCargaTexto = ctesDuplicadosPorCarga[carga] 
                ? Object.entries(ctesDuplicadosPorCarga[carga]).map(([cteMenor, outrosCtes]) => {
                    return `Carga ${carga}: ${cteMenor} (CTE Menor)\nCTEs com notas fiscais repetidas: ${outrosCtes.join(', ')}`;
                }).join('\n')
                : '';
            return `${ctesPorCargaTexto}${ctesDuplicadosPorCargaTexto}`;
        }).join('\n\n');
    } else {
        notasCargaResultado.value = 'Nenhum Cte encontrado por carga.';
    }

    if (Object.keys(notasRepetidas).length > 0) {
        notasRepetidasResultado.value = Object.entries(notasRepetidas).map(([nota, ctes]) => {
            return `Nota Fiscal ${nota}: ${ctes.join(', ')} (Primeiro Cte: ${ctes[0]})`;
        }).join('\n\n');
    } else {
        notasRepetidasResultado.value = 'Nenhuma nota fiscal repetida encontrada.';
    }

    divResultado.select();
    divResultado.setSelectionRange(0, 99999);
    notasResultado.select();
    notasResultado.setSelectionRange(0, 99999);
    notasCargaResultado.select();
    notasCargaResultado.setSelectionRange(0, 99999);
    notasRepetidasResultado.select();
    notasRepetidasResultado.setSelectionRange(0, 99999);

    document.execCommand("copy");
}
