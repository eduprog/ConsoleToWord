using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ConsoleToWord;

public class GerarBriefingMerro
{
    public static void Gerar()
    {
        // Caminho onde o arquivo será salvo
        var filePath = Path.Combine(Directory.GetCurrentDirectory(), "Briefing_Merro_Startup_v2.docx");

        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }

        // Criação do documento
        using WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        // Adiciona uma parte principal de documento
        MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
        mainPart.Document = new Document();
        Body body = new Body();

        // Adiciona título
        body.Append(Utilities.CreateParagraph("Briefing para Descrição dos Serviços da Startup", true, 22, JustificationValues.Center));

        // Seção 1: Informações da Empresa
        body.Append(Utilities.CreateParagraph("1. Informações da Empresa", true, 16, JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Nome da Empresa: Merro - Mercado Rotariano Co.", justification: JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Fundador: Ademério Eduardo Moreira", justification: JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Ano de Fundação: 1624", justification: JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Localização: Brasil", justification: JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Visão: Ser a rede social dos rotarianos pelo mundo.", justification: JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Missão: Fornecer tecnologia inovadora, provendo ainda mais o companheirismo pelo mundo.", justification: JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Valores: Companheirismo, Inovação, Transparência e Responsabilidade social.", justification: JustificationValues.Left));

        // Seção 2: Descrição dos Serviços
        body.Append(Utilities.CreateParagraph("2. Descrição dos Serviços", true, 16, JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("2.1. Solução para Oferecimento de Produtos e Serviços", true, 16, JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Nome do Serviço: Merro Marketplace", justification: JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Descrição: Plataforma onde rotarianos podem oferecer produtos e serviços aos companheiros, facilitando a troca e a colaboração entre membros.", justification: JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Público-alvo: Rotarianos de todo o mundo.", justification: JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Benefícios: Facilita o acesso a produtos e serviços de confiança, promovendo o comércio dentro da comunidade rotariana.", justification: JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Características Principais: Cadastro de produtos e serviços, perfis detalhados dos rotarianos, sistema de busca filtrada por classificação e avenida de serviços.", justification: JustificationValues.Left));

        body.Append(Utilities.CreateParagraph("2.2. Troca de Mensagens e Quadro de Avisos", true, 16, JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Nome do Serviço: Merro Connect", justification: JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Descrição: Ferramenta de comunicação que permite aos rotarianos trocar mensagens diretamente, além de um quadro de avisos para facilitar o funcionamento dos clubes.", justification: JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Público-alvo: Rotarianos e clubes rotarianos.", justification: JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Benefícios: Melhora a comunicação interna, facilita a organização de eventos e atividades do clube.", justification: JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Características Principais: Mensagens diretas, quadros de avisos, notificações em tempo real.", justification: JustificationValues.Left));

        // Seção 3: Mercado e Competidores
        body.Append(Utilities.CreateParagraph("3. Mercado e Competidores", true, 16, JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Análise de Mercado: Atualmente, rotarianos usam grupos separados e o MyRotary para obter informações, mas o contato direto é principalmente através de conferências anuais. A Merro oferece uma solução integrada para facilitar o contato direto entre rotarianos, filtrado por classificação e avenida de serviços.", justification: JustificationValues.Distribute));
        body.Append(Utilities.CreateParagraph("Principais Competidores: Grupos no WhatsApp, MyRotary, conferências anuais.", justification: JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Diferenciais Competitivos: Contato direto e filtrado, plataforma integrada com várias funcionalidades para facilitar a vida dos rotarianos.", justification: JustificationValues.Left));

        // Seção 4: Estratégia de Marketing
        body.Append(Utilities.CreateParagraph("4. Estratégia de Marketing", true, 16, JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Posicionamento de Marca: A plataforma essencial para rotarianos que desejam estreitar laços e colaborar de maneira mais eficiente.", justification: JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Canais de Comunicação: Divulgação através do MyRotary, grupos de amizade no WhatsApp, e-mail, redes sociais.", justification: JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Estratégias de Aquisição de Clientes: Campanhas de marketing digital, parcerias com clubes rotarianos, divulgação em conferências e eventos rotarianos.", justification: JustificationValues.Distribute));

        // Seção 5: Estrutura e Equipe
        body.Append(Utilities.CreateParagraph("5. Estrutura e Equipe", true, 16, JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Organograma: CEO (Ademério Eduardo Moreira) e um futuro CTO.", justification: JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Principais Membros da Equipe:", justification: JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Ademério Eduardo Moreira (CEO): Responsável pela visão estratégica e liderança da startup.", justification: JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Experiência e Competências: A equipe ainda está em formação, com planos para adicionar um CTO para definir as tecnologias a serem utilizadas.", justification: JustificationValues.Left));

        // Seção 6: Metas e Objetivos
        body.Append(Utilities.CreateParagraph("6. Metas e Objetivos", true, 16, JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Curto Prazo: Lançar um site e um aplicativo para cadastro de rotarianos e seus produtos e serviços. Facilitar a busca por serviços de rotarianos em diversas localidades.", justification: JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Médio Prazo: Expandir a rede para todo o território nacional e incluir funcionalidades de venda diretamente pelo aplicativo.", justification: JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Longo Prazo: Levar a plataforma para outros países, alcançando o Rotary Internacional.", justification: JustificationValues.Left));

        // Seção 7: Informações Adicionais
        body.Append(Utilities.CreateParagraph("7. Informações Adicionais", true, 16, JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Parcerias: Parcerias com o MyRotary e clubes do distrito para divulgação.", justification: JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Investimentos: Inicialmente, não haverá aporte monetário; os recursos serão oferecidos pelos profissionais envolvidos.", justification: JustificationValues.Left));

        // Seção 8: Contato
        body.Append(Utilities.CreateParagraph("8. Contato", true, 16, JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Nome: Ademério Eduardo Moreira (CEO)", justification: JustificationValues.Left));

        // Adiciona uma seção para anexos
        body.Append(Utilities.CreateParagraph("Anexos", true, 16, JustificationValues.Left));
        body.Append(Utilities.CreateParagraph("Inclua quaisquer anexos relevantes, como apresentações, gráficos ou estudos de caso.", justification: JustificationValues.Left));

        // Adiciona o corpo ao documento
        mainPart.Document.Append(body);
        mainPart.Document.Save();
    }
}