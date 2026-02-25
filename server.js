const express = require('express');
const cors = require('cors');
const { Document, Packer, Paragraph, AlignmentType, HeadingLevel, TextRun } = require("docx");
const path = require('path');

const app = express();
app.use(cors());
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

app.post('/gerar-contrato', async (req, res) => {
    try {
        const d = req.body;
        const safe = (v) => v ? v : "__________";

        // Lógica de Datas
        const inicio = d.data_inicio ? new Date(d.data_inicio) : null;
        let dInicio = "", dFim = "";
        if (inicio && !isNaN(inicio)) {
            const fim = new Date(inicio);
            fim.setMonth(fim.getMonth() + parseInt(d.prazo || 0));
            dInicio = inicio.toLocaleDateString('pt-BR');
            dFim = fim.toLocaleDateString('pt-BR');
        }
        const dContrato = d.data_contrato ? new Date(d.data_contrato).toLocaleDateString('pt-BR') : new Date().toLocaleDateString('pt-BR');

        // Helpers de Formatação Word
        const titulo = (t) => new Paragraph({
            children: [new TextRun({ text: t, bold: true, size: 24 })],
            spacing: { before: 350, after: 150 }
        });
        const texto = (t) => new Paragraph({
            text: t, alignment: AlignmentType.JUSTIFIED,
            spacing: { after: 220 }, lineSpacing: { line: 360 }
        });

        const content = [
            new Paragraph({ text: "CONTRATO DE LOCAÇÃO DE IMÓVEL RESIDENCIAL", heading: HeadingLevel.HEADING_1, alignment: AlignmentType.CENTER, spacing: { after: 500 } }),

            titulo("CLÁUSULA 1ª – DAS PARTES"),
            texto(`LOCADOR(A): ${safe(d.nome_locador)}, CPF nº ${safe(d.cpf_locador)}${d.rg_locador ? ', RG nº ' + d.rg_locador + ' ' + safe(d.orgao_locador) : ''}, residente e domiciliado(a) em ${safe(d.endereco_locador)}${d.num_res_locador ? ', nº ' + d.num_res_locador : ''}${d.tel_locador ? ', Contato: ' + d.tel_locador : ''}.`),
            texto(`LOCATÁRIO(A): ${safe(d.nome_locatario)}, CPF nº ${safe(d.cpf_locatario)}${d.rg_locatario ? ', RG nº ' + d.rg_locatario + ' ' + safe(d.orgao_locatario) : ''}, residente e domiciliado(a) em ${safe(d.endereco_locatario)}${d.num_res_locatario ? ', nº ' + d.num_res_locatario : ''}${d.tel_locatario ? ', Contato: ' + d.tel_locatario : ''}.`),

            titulo("CLÁUSULA 2ª – DO OBJETO"),
            texto(`O objeto da presente locação é o imóvel residencial localizado em: ${safe(d.endereco_imovel)}. O LOCATÁRIO declara ter vistoriado o imóvel e o aceita nas condições em que se encontra.`),

            titulo("CLÁUSULA 3ª – DO PRAZO"),
            texto(`A locação terá o prazo determinado de ${safe(d.prazo)} meses, com início em ${dInicio} e término em ${dFim}, data em que o LOCATÁRIO se obriga a restituir o imóvel livre de pessoas e bens.`),

            titulo("CLÁUSULA 4ª – DO VALOR E PAGAMENTO"),
            texto(`O aluguel mensal é fixado em R$ ${safe(d.valor)} (${safe(d.valor_extenso)}). O pagamento deverá ser efetuado até o dia ${safe(d.dia_pagamento)} de cada mês, através de ${safe(d.forma_pagamento)}, sob pena de multa de 10% por atraso e juros de mora.`),

            titulo("CLÁUSULA 5ª – DOS ENCARGOS E TRIBUTOS"),
            texto("Caberá ao LOCATÁRIO, além do aluguel, o pagamento pontual das contas de energia elétrica, água, esgoto, taxa de lixo e condomínio, se houver, bem como o IPTU proporcional ao período de ocupação."),

            titulo("CLÁUSULA 6ª – DA CONSERVAÇÃO E MANUTENÇÃO"),
            texto("O LOCATÁRIO obriga-se a manter o imóvel em perfeito estado de conservação e higiene. Fica proibida qualquer alteração estrutural, pintura ou reforma sem consentimento prévio e por escrito do LOCADOR. Danos causados deverão ser reparados pelo LOCATÁRIO."),

            titulo("CLÁUSULA 7ª – DA DESTINAÇÃO"),
            texto("O imóvel destina-se exclusivamente ao uso RESIDENCIAL do LOCATÁRIO e sua família. É expressamente proibida a sublocação, cessão ou empréstimo do imóvel, seja total ou parcial, sem autorização prévia."),

            titulo("CLÁUSULA 8ª – DO SILÊNCIO E NORMAS SOCIAIS"),
            texto("O LOCATÁRIO compromete-se a respeitar as leis de silêncio e as normas de vizinhança, respondendo por quaisquer multas decorrentes de condutas inadequadas durante a vigência do contrato."),

            titulo("CLÁUSULA 9ª – DAS VISTORIAS"),
            texto("O LOCADOR poderá realizar vistorias periódicas no imóvel, mediante aviso prévio de 24 horas, para verificar o estado de conservação do bem locado."),
        ];

        // Condicionais
        if (d.chk_reajuste === "on") {
            content.push(titulo("CLÁUSULA 10ª – DO REAJUSTE ANUAL"), texto(`Caso o contrato seja prorrogado, o valor será reajustado anualmente pelo índice ${safe(d.indice_reajuste)}.`));
        }
        if (d.chk_garantia === "on") {
            content.push(titulo("CLÁUSULA 11ª – DA GARANTIA"), texto(`Como garantia das obrigações, o LOCATÁRIO entrega a título de ${safe(d.tipo_garantia)} o valor de R$ ${safe(d.valor_caucao)}.`));
        }
        if (d.chk_multa === "on") {
            content.push(titulo("CLÁUSULA 12ª – DA MULTA RESCISÓRIA"), texto(`A parte que rescindir o contrato antes do prazo pagará multa equivalente a ${safe(d.multa_alugueis)} (${safe(d.multa_extenso)}) aluguéis, de forma proporcional.`));
        }

        content.push(
            titulo("CLÁUSULA 13ª – DA DEVOLUÇÃO DO IMÓVEL"),
            texto("Ao término do contrato, o imóvel deverá ser devolvido pintado e limpo, com todas as contas quitadas e chaves entregues formalmente ao LOCADOR."),

            titulo("CLÁUSULA 14ª – DO FORO"),
            texto(`Para dirimir quaisquer questões fundadas neste contrato, as partes elegem o Foro da Comarca de ${safe(d.foro)}, com renúncia a qualquer outro.`),

            new Paragraph({ text: `${safe(d.foro)}, ${dContrato}.`, spacing: { before: 600, after: 600 } }),
            new Paragraph({ text: "________________________________________", alignment: AlignmentType.CENTER }),
            new Paragraph({ text: `LOCADOR(A): ${safe(d.nome_locador)}`, alignment: AlignmentType.CENTER, spacing: { after: 500 } }),
            new Paragraph({ text: "________________________________________", alignment: AlignmentType.CENTER }),
            new Paragraph({ text: `LOCATÁRIO(A): ${safe(d.nome_locatario)}`, alignment: AlignmentType.CENTER })
        );

        const doc = new Document({ sections: [{ children: content }] });
        const buffer = await Packer.toBuffer(doc);
        const nomeArquivo = `Contrato_${safe(d.nome_locatario).replace(/\s+/g, '_')}.docx`;

        res.set({ "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "Content-Disposition": `attachment; filename=${nomeArquivo}`, "Content-Length": buffer.length });
        res.end(buffer);
    } catch (e) { console.error(e); res.status(500).send("Erro no servidor."); }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Servidor rodando na porta ${PORT}`);
});