function validarCampo(id) {
    let campo = document.getElementById(id);

    if (!campo.value) {
        campo.classList.add("erro");
        campo.classList.remove("ok");
        return false;
    }

    campo.classList.remove("erro");
    campo.classList.add("ok");
    return true;
}

function processar() {

    let fileInput = document.getElementById("arquivo");
    let file = fileInput.files[0];

    if (!file) {
        alert("Selecione um arquivo!");
        return;
    }




    let header = {
    codigoLayout: document.getElementById("codigoLayout").value,
    dataGeracao: document.getElementById("dataGeracao").value,
    sequencial: document.getElementById("sequencial").value,
    anoReferencia: document.getElementById("anoReferencia").value,
    ugResponsavel: document.getElementById("ugResponsavel").value,
    cpfResponsavel: document.getElementById("cpfResponsavel").value
};
if (!header.ugResponsavel || !header.cpfResponsavel) {
    alert("Preencha os dados do header!");
    return;
}


    let layout = header.codigoLayout;

    



    let reader = new FileReader();

    reader.onload = function(e) {
        let data = new Uint8Array(e.target.result);
        let workbook = XLSX.read(data, { type: 'array' });

        let primeiraAba = workbook.SheetNames[0];
        let sheet = workbook.Sheets[primeiraAba];

        let json = XLSX.utils.sheet_to_json(sheet, {
    raw: false,
    dateNF: "yyyy-mm-dd"
});

       if (layout === "CH001") {
    gerarXML_CH001(json, header);
    } else if (layout === "DH001") {
    gerarXML_DH(json, header); 
    } else {
    alert("Layout não suportado!");
}
    };

    reader.readAsArrayBuffer(file);

    if (!validarCampo("ugResponsavel") || 
    !validarCampo("cpfResponsavel") )
     {

    document.getElementById("status").innerText = "Preencha os campos obrigatórios";
    return;

    document.getElementById("status").innerText = "⏳ Gerando XML...";

    document.getElementById("status").innerText = "✅ XML gerado com sucesso!";
}
}

function gerarXML_CH001(dados, header) {

    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sb:arquivo xmlns:ns2="http://services.tabelas.siafi.tesouro.fazenda.gov.br/" 
            xmlns:sb="http://www.tesouro.gov.br/siafi/submissao">

    <sb:header>
        <sb:codigoLayout>CH001</sb:codigoLayout>
        <sb:dataGeracao>${formatarDataBR(header.dataGeracao)}</sb:dataGeracao>
        <sb:sequencialGeracao>${header.sequencial}</sb:sequencialGeracao>
        <sb:anoReferencia>${header.anoReferencia}</sb:anoReferencia>
        <sb:ugResponsavel>${header.ugResponsavel}</sb:ugResponsavel>
        <sb:cpfResponsavel>${limparNumero(header.cpfResponsavel)}</sb:cpfResponsavel>
    </sb:header>

    <sb:detalhes>
`;

    let contador = 0;

    dados.forEach((linha, index) => {
        try {

            let credor = limparNumero(linha["Credor"]);
            let chave = linha["Chave Pix"];
            let tipo = (linha["Tipo Chave"] || "CPF").toUpperCase();

            if (!credor) throw new Error("Credor obrigatório");
            if (!chave) throw new Error("Chave Pix obrigatória");

            const tiposValidos = ["CPF", "CNPJ", "EMAIL", "TELEFONE", "EVP"];
            if (!tiposValidos.includes(tipo)) {
                throw new Error("Tipo de chave inválido");
            }

            contador++;

            xml += `
        <sb:detalhe>
            <ns2:tabAlterarChavesPix>
                <paramAlterarChavesPix>

                    <credor>${credor}</credor>

                    <chavePix>
                        <tipoChave>${tipo}</tipoChave>
                        <chave>${chave}</chave>
                    </chavePix>

                </paramAlterarChavesPix>
            </ns2:tabAlterarChavesPix>
        </sb:detalhe>
`;

        } catch (erro) {
            throw new Error(`Erro na linha ${index + 2}: ${erro.message}`);
        }
    });

    xml += `
    </sb:detalhes>

    <sb:trailler>
        <sb:quantidadeDetalhe>${contador}</sb:quantidadeDetalhe>
    </sb:trailler>

</sb:arquivo>`;

    baixarXML(xml);
}
    

function gerarPredocOB(linha, cpf) {

    let tipoOB = linha["Tipo OB"] || "OBPIX";

    if (tipoOB === "OBC") {

        // 🔒 Validação mínima
        if (!linha["Banco"] || !linha["Agência"] || !linha["Conta"]) {
            throw new Error("OBC exige banco, agência e conta!");
        }

        return `
        <predocOB>
            <codTipoOB>OBC</codTipoOB>
            <codCredorDevedor>${cpf}</codCredorDevedor>

            <numDomiBancFavo>
                <banco>${linha["Banco"]}</banco>
                <agencia>${linha["Agência"]}</agencia>
                <conta>${linha["Conta"]}</conta>
            </numDomiBancFavo>

            <numDomiBancPgto>
                <conta>UNICA</conta>
            </numDomiBancPgto>

            <txtProcesso>${linha["Processo"] || ""}</txtProcesso>
        </predocOB>
        `;
    }

    if (tipoOB === "OBPIX") {

        if (!linha["Credor DH"]) {
    throw new Error("OBPIX exige chave Pix (Credor DH)!");
}

        return `
        <predocOB>
            <codTipoOB>OBPIX</codTipoOB>
            <codCredorDevedor>${cpf}</codCredorDevedor>

            <txtChavePix>${cpf}</txtChavePix>

            <numDomiBancPgto>
                <banco>002</banco>
                <conta>PAGINST</conta>
            </numDomiBancPgto>

            <txtProcesso>${linha["Processo"] || ""}</txtProcesso>
        </predocOB>
        `;
    }

    throw new Error("Tipo OB inválido: " + tipoOB);
}


    function gerarXML_DH(dados, header)  {

    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sb:arquivo xmlns:ns2="http://services.docHabil.cpr.siafi.tesouro.fazenda.gov.br/" xmlns:sb="http://www.tesouro.gov.br/siafi/submissao">

    <sb:header>
        <sb:codigoLayout>${header.codigoLayout}</sb:codigoLayout>
        <sb:dataGeracao>${formatarDataBR(header.dataGeracao)}</sb:dataGeracao>
        <sb:sequencialGeracao>${header.sequencial}</sb:sequencialGeracao>
        <sb:anoReferencia>${header.anoReferencia}</sb:anoReferencia>
        <sb:ugResponsavel>${header.ugResponsavel}</sb:ugResponsavel>
        <sb:cpfResponsavel>${limparNumero(header.cpfResponsavel)}</sb:cpfResponsavel>
    </sb:header>

    <sb:detalhes>
`;

    let contador = 0;

    dados.forEach((linha, index) => {
    try {

        if (!linha["Valor DH"]) return;

        contador++;

        let valor = formatarValor(linha["Valor DH"]);
        let cpf = limparNumero(linha["Credor DH"]);

        xml += `
        <sb:detalhe>
            <ns2:CprDhCadastrar>
                <codUgEmit>${linha["UG Emitente"]}</codUgEmit>
                <anoDH>${linha["Ano DH"]}</anoDH>
                <codTipoDH>${linha["Tipo DH"]}</codTipoDH>

                <dadosBasicos>
                    <dtEmis>${formatarDataISO(linha["Data Emissão"])}</dtEmis>
                    <dtVenc>${formatarDataISO(linha["Data Vencimento"])}</dtVenc>
                    <codUgPgto>${linha["UG Emitente"]}</codUgPgto>
                    <vlr>${valor}</vlr>
                    <txtObser>${linha["Observação Dados Básicos"] || ""}</txtObser>
                    <txtProcesso>${linha["Processo"] || ""}</txtProcesso>
                    <dtAteste>${formatarDataISO(linha["Data Ateste"])}</dtAteste>
                    <codCredorDevedor>${cpf}</codCredorDevedor>
                    <dtPgtoReceb>${formatarDataISO(linha["Data Pagamento"])}</dtPgtoReceb>

                    <docOrigem>
                        <codIdentEmit>${linha["UG Emitente"]}</codIdentEmit>
                        <dtEmis>${formatarDataISO(linha["Data Ateste"])}</dtEmis>
                        <numDocOrigem>${linha["Número Doc Origem"] || ""}</numDocOrigem>
                        <vlr>${valor}</vlr>
                    </docOrigem>
                </dadosBasicos>

                <pco>
    <numSeqItem>1</numSeqItem>
    <codSit>${linha["Situação"] || "DSP061"}</codSit>
    <codUgEmpe>${linha["UG Emitente"]}</codUgEmpe>

    <pcoItem>
        <numSeqItem>1</numSeqItem>
        <numEmpe>${linha["Empenho"] || ""}</numEmpe>
        <codSubItemEmpe>${linha["Subitem"] || "01"}</codSubItemEmpe>
        <vlr>${valor}</vlr>
        <numClassA>${linha["Conta VPD"] || ""}</numClassA>
    </pcoItem>
</pco>

<centroCusto>
    <numSeqItem>1</numSeqItem>
    <codCentroCusto>${linha["Centro de Custo"] || "CC-GENERICO"}</codCentroCusto>
    <mesReferencia>${linha["Mês"] || "01"}</mesReferencia>
    <anoReferencia>${linha["Ano"] || header.anoReferencia}</anoReferencia>
    <codUgBenef>${linha["UG Emitente"]}</codUgBenef>
    <codSIORG>${linha["SIORG"]}</codSIORG>

    <relPcoItem>
        <numSeqPai>1</numSeqPai>
        <numSeqItem>1</numSeqItem>
        <vlr>${valor}</vlr>
    </relPcoItem>
</centroCusto>

                <dadosPgto>
                    <codCredorDevedor>${cpf}</codCredorDevedor>
                    <vlr>${valor}</vlr>
                    <predoc>
                        <txtObser>${linha["Observação Pré-Doc OB"] || ""}</txtObser>
                        ${gerarPredocOB(linha, cpf)}
                         
                    </predoc>
                </dadosPgto>

            </ns2:CprDhCadastrar>
        </sb:detalhe>
`;
} catch (erro) {
        alert(`Erro na linha ${index + 2}: ${erro.message}`);
    }
    });

    xml += `
    </sb:detalhes>

    <sb:trailler>
        <sb:quantidadeDetalhe>${contador}</sb:quantidadeDetalhe>
    </sb:trailler>

</sb:arquivo>`;

    baixarXML(xml);
}
function limparNumero(valor) {
    if (!valor) return "";
    return valor.toString().replace(/\D/g, "");
}
function formatarValor(valor) {
    return parseFloat(valor || 0).toString();
}    
function formatarDataISO(data) {

    if (!data) return "";

    // 🔹 Caso seja número (Excel)
    if (typeof data === "number") {
        let excelEpoch = new Date(1899, 11, 30);
        let result = new Date(excelEpoch.getTime() + data * 86400000);
        return result.toISOString().split("T")[0];
    }

    // 🔹 Caso seja string/data normal
    let d = new Date(data);

    if (isNaN(d)) return "";

    return d.toISOString().split("T")[0];
}
function formatarDataBR(data) {

    if (!data) return "";

    // Evita problema de fuso
    let partes = data.split("-"); // yyyy-mm-dd

    return `${partes[2]}/${partes[1]}/${partes[0]}`;
}

function baixarXML(conteudo) {

    let blob = new Blob([conteudo], { type: "application/xml" });

    let link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "siafi.xml";

    link.click();
}

function baixarModelo() {
    let link = document.createElement("a");
    link.href = "modelo_OBPIX.xlsx";
    link.download = "modelo_OBPIX.xlsx";
    link.click();
}


let dropzone = document.getElementById("dropzone");
let fileInput = document.getElementById("arquivo");

dropzone.addEventListener("click", () => fileInput.click());

dropzone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropzone.style.borderColor = "#0c326f";
});

dropzone.addEventListener("dragleave", () => {
    dropzone.style.borderColor = "#aaa";
});

dropzone.addEventListener("drop", (e) => {
    e.preventDefault();
    fileInput.files = e.dataTransfer.files;
    dropzone.innerText = e.dataTransfer.files[0].name;
});