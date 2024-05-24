using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ConsoleToWord;

public class GerarBriefingFiscoCloud
{
    public static void Gerar()
    {
        // Caminho onde o arquivo será salvo
        var filePath = Path.Combine(Directory.GetCurrentDirectory(), "Briefing_FiscoCloud.docx");

        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }

        using var wordDocument =
            WordprocessingDocument.Create(filePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        // Adiciona uma parte principal de documento
        var mainPart = wordDocument.AddMainDocumentPart();
        mainPart.Document = new Document();
        var body = new Body();

        // Adiciona título
        body.Append(Utilities.CreateParagraph("Briefing para Descrição dos Serviços da Startup", true, 24));

        // Seção 1: Informações da Empresa
        body.Append(Utilities.CreateParagraph("1. Informações da Empresa", true, 20));
        body.Append(Utilities.CreateParagraph("Nome da Empresa: FiscoCloud Serviços em TI LTd."));
        body.Append(Utilities.CreateParagraph("Fundadores: Ademério Eduardo Moreira"));
        body.Append(Utilities.CreateParagraph("Ano de Fundação: 2024"));
        body.Append(Utilities.CreateParagraph("Localização: Brasil"));
        body.Append(Utilities.CreateParagraph(
            "Visão: Ser o serviço relativo a tributos no Brasil. Atuando como gateway de autorização, emissão e armazenamento de documentos."));
        body.Append(Utilities.CreateParagraph(
            "Missão: Fornecer tecnologia inovadora, provendo armazenamento, análise, simplificando documentos fiscais e não fiscais."));
        body.Append(Utilities.CreateParagraph("Valores: Inovação, Transparência e Responsabilidade social."));

        // Seção 2: Descrição dos Serviços Oferecidos
        body.Append(Utilities.CreateParagraph("2. Serviços Oferecidos", true, 20));
        body.Append(Utilities.CreateParagraph(
            "Solução como gateway de autorização de documentos fiscais NFE, NFCE, NFSE, CTE, DFE, dentre outros."));
        body.Append(Utilities.CreateParagraph("Solução de análise através de IA de tributos de mercadorias."));
        body.Append(Utilities.CreateParagraph("Solução de armazenamento de documentos fiscais e não fiscais."));

        // Seção 3: Análise de Mercado e Competidores
        body.Append(Utilities.CreateParagraph("3. Análise de Mercado e Competidores", true, 20));
        body.Append(Utilities.CreateParagraph(
            "A empresa enfrentará concorrência de outras empresas que oferecem serviços semelhantes, porém, a FiscoCloud se diferenciará pela sua abordagem inovadora, transparência e parcerias estratégicas com órgãos emissores, escritórios de contabilidade e SoftwareHouses."));

        // Seção 4: Estratégia de Marketing
        body.Append(Utilities.CreateParagraph("4. Estratégia de Marketing", true, 20));
        body.Append(Utilities.CreateParagraph("E-mail marketing."));
        body.Append(Utilities.CreateParagraph("Tráfego pago."));
        body.Append(Utilities.CreateParagraph("Escritórios contábeis."));
        body.Append(Utilities.CreateParagraph("Software houses."));

        // Seção 5: Estrutura e Equipe
        body.Append(Utilities.CreateParagraph("5. Estrutura e Equipe", true, 20));
        body.Append(Utilities.CreateParagraph("CEO - Ademério Eduardo Moreira"));

        // Seção 6: Metas e Objetivos
        body.Append(Utilities.CreateParagraph("6. Metas e Objetivos", true, 20));
        body.Append(Utilities.CreateParagraph(
            "A curto prazo prover Autorização, Emissão e armazenamento de documentos fiscais NFE, NFCE e Análise usando IA de tributos de mercadorias."));
        body.Append(Utilities.CreateParagraph(
            "A Média prazo autorização, emissão e armazenamento de documentos fiscais NFSE, CTE, DFE e outros."));
        body.Append(
            Utilities.CreateParagraph("A longo prazo, Solução de armazenamento, emissão de quaisquer tipos de documentos."));

        // Seção 7: Informações Adicionais
        body.Append(Utilities.CreateParagraph("7. Informações Adicionais", true, 20));
        body.Append(Utilities.CreateParagraph(
            "Parceira com os mais diversos orgãos emissores bem como escritórios de contabilidade e SoftwareHouse com approach junto aos profissionais da área contábil."));

        // Adiciona o corpo ao documento
        mainPart.Document.Append(body);
        mainPart.Document.Save();
    }
}