namespace ConsoleToWord;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;

public class ComFormatacao{
    public static void  GerarComFormatacao(string filePath)
    {
        // Caminho onde o arquivo será salvo

        // Criação do documento
        using WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        // Adiciona uma parte principal de documento
        MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
        mainPart.Document = new Document();
        Body body = new Body();

        // Adiciona título
        body.Append(Utilities.CreateParagraph("Briefing para Descrição dos Serviços da Startup", true, 24, JustificationValues.Center));

        // Seção 1: Informações da Empresa
        body.Append(Utilities.CreateParagraph("1. Informações da Empresa", true, 20, JustificationValues.Left));
        // ...

        // Seção 2: Descrição dos Serviços
        body.Append(Utilities.CreateParagraph("2. Descrição dos Serviços", true, 20, JustificationValues.Left));
        // ...

        // Seção 3: Mercado e Competidores
        body.Append(Utilities.CreateParagraph("3. Mercado e Competidores", true, 20, JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Análise de Mercado: Atualmente, rotarianos usam grupos separados e o MyRotary para obter informações, mas o contato direto é principalmente através de conferências anuais. A Merro oferece uma solução integrada para facilitar o contato direto entre rotarianos, filtrado por classificação e avenida de serviços.", justification: JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Principais Competidores: Grupos no WhatsApp, MyRotary, conferências anuais.", justification: JustificationValues.Left));
        // Diferenciais Competitivos em negrito
        body.Append(Utilities.CreateParagraphWithBold("Diferenciais Competitivos:", "Contato direto e filtrado, plataforma integrada com várias funcionalidades para facilitar a vida dos rotarianos.", JustificationValues.Left));

        // Seção 4: Estratégia de Marketing
        // ...

        // Seção 5: Estrutura e Equipe
        // ...

        // Seção 6: Metas e Objetivos
        // ...

        // Seção 7: Informações Adicionais
        // ...

        // Seção 8: Contato
        // ...

        // Adiciona uma seção para anexos
        // ...

        // Adiciona o corpo ao documento
        mainPart.Document.Append(body);
        mainPart.Document.Save();
    }

    // Método auxiliar para criar parágrafos
    
}
