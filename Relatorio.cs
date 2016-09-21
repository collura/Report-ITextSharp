using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
//ESSAS SAO AS BIBLIOTECAS QUE DEVEREMOS ADICIONAR PARA CRIAR PDFS
using System.IO;// A BIBLIOTECA DE ENTRADA E SAIDA DE ARQUIVOS
using iTextSharp;//E A BIBLIOTECA ITEXTSHARP E SUAS EXTENÇÕES
using iTextSharp.text;//ESTENSAO 1 (TEXT)
using iTextSharp.text.pdf;//ESTENSAO 2 (PDF)
using System.Data.SQLite;

namespace Pdf
{


    public class GerarPdf
    {
     
        //Utilizando as Fontes do Windows
        int totalfonts = FontFactory.RegisterDirectory("C:\\WINDOWS\\Fonts");

        //Criando Fontes Personalizadas
        static String t;
        static Font titulo = FontFactory.GetFont("Arial", 28, Color.DARK_GRAY);
        static Font page = FontFactory.GetFont("Verdana", 10, Font.BOLDITALIC, Color.GRAY); // new Color(125, 88, 15));

        static SQLiteDataReader reader;



        //Emite relatorio de Efetivo
        public static void relatorioEfetivo(String batalhao, String oficial, String agenteInteligencia) {


            reader = Crud.relatorioEfetivo(batalhao, oficial, agenteInteligencia);

            Document doc = new Document(PageSize.A4); //criando e estipulando o tipo da folha usada           
            doc.SetMargins(19, 20, 10, 50); //estibulando o espaçamento das margens que queremos           
            doc.AddTitle ("Relatorio de Efetivo");
            doc.AddAuthor("Agência Regional -  CPA / M-11");
            // doc.AddProducer("Producer por Betto");
            doc.AddKeywords("IText");
            doc.AddCreationDate(); //adicionando as configuracoes
            //Caminho onde será savo o arquivo + o nomo do arquivo sempre seguido por .pdf
            String caminho = @"E:\" + "xxx.pdf";

            //Objeto para escrever no documento. com os parametros acima
            PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(caminho, FileMode.Create));

            //Abre a sessão para ediçao do documento 
            doc.Open();          

           //criando a variavel para a configuração do paragrafo          
            // **  CASO DE DOCUMENTO SEM TABELAS **

            //estipulando o alinhamneto
            // titulo.Alignment = Element.ALIGN_CENTER;
            //Estipulando a fonte
            //  titulo.Font = new Font(Font.NORMAL, 7, (int)System.Drawing.FontStyle.Bold);
            //Estipilando o Título
            // titulo.Add("Dados Cadastrais");  
       
       
            //Instancia para configuração do parágrafo que exibirá o numero da pagina       
            Paragraph pag = new Paragraph("Pag " + (doc.PageNumber + 1),page);            
            doc.Add(pag);

       
            //doc.AddHeader("Pag", "Pagina"); Nome e valor de uma propriedade

            //Instancia para configuração do parágrafo que exibirá a tabela                      
            Paragraph paragrafo = new Paragraph("", new Font(titulo));

            //Criação de uma tabela 
            PdfPTable table = new PdfPTable(6);    //QUANTIDADE DE COLUNAS DA TABELA    

            //Colunas da tabela com larguras relativas- 1/25 and 8/25 ...etc
            float[] widths = new float[] { 2.5f, 2f, 4.5f, 3f, 3.5f, 4.3f };
            //Define as larguras das colunas usando os parametros acima
            table.SetWidths(widths);
            //Definindo o alinhamento da tabela  - 1 Centro, 2 Direita, 0 Esquerda
            table.HorizontalAlignment = 1;

            // 'true' assume valor fixo para a tabela
            //'false' calcula e divide o espaço disponível                   
            table.LockedWidth = true;
            table.TotalWidth = 550; // ** no caso de  true assumira este valor ** //
            //Margim Superios da Tabela
            table.SpacingBefore = 10;
            //Margim Inferior da Tabela
            table.SpacingAfter = 30;

            //Criação do Objeto Célula da tabela para o Titulo  

            
            if (batalhao.Equals("CPA") && oficial.Equals("true"))
            {
                t = "Efetivo de Oficiais - CPA / M-11";
            }
            else if (batalhao.Equals("CPA") && oficial.Equals("false"))
            {
                t = "Efetivo de Praças - CPA / M-11";
            }

            PdfPCell cell = new PdfPCell(new Phrase(t + "\n\n", titulo));          

             //Adicionando Valor a celula     
            doc.Add(paragrafo);  // Adiciona a celula e pula para a proxima linha

            //Divide a Célula em 4 - 
            //Terá efeito Somente a partir da adição das próximas células        
            cell.Colspan = 6;
            cell.PaddingTop = 10;
            cell.PaddingLeft = 10;
            cell.Border = 0;
            cell.HorizontalAlignment = 1;

            //Adiciona na tabela as celulas criadas ate aqui
            table.AddCell(cell);


            //Adicionando as celulas que serão o cabeçalho da tabela
            //Quatro celulas pois o Colspan foi defido com 4 colunas, logo cada celula ocupara 
            //uma coluna a partir da esquerda
            table.AddCell(new Phrase("Post / Grad"));
            table.AddCell(new Phrase("Re"));
            table.AddCell(new Phrase("Nome de Guerra"));
            table.AddCell(new Phrase("Fone Res"));
            table.AddCell(new Phrase("Fone Cel"));
            table.AddCell(new Phrase("Email"));
            //Adicionando dados do banco de dados à tabela
           
            while (reader.Read())
            {
                table.AddCell(new PdfPCell(new Phrase(Dal.reader["postoGrad"].ToString())));
                table.AddCell(new PdfPCell(new Phrase(Dal.reader["re"].ToString())));
                table.AddCell(new PdfPCell(new Phrase(Dal.reader["nomeGuerra"].ToString())));
                table.AddCell(new PdfPCell(new Phrase(Dal.reader["foneRes"].ToString())));
                table.AddCell(new PdfPCell(new Phrase(Dal.reader["foneCel1"].ToString())));
                table.AddCell(new PdfPCell(new Phrase(Dal.reader["emailFuncional"].ToString())));
            }

            Crud.bdClose();
            //Adicionando a Tabela ao documento              
            doc.Add(table);
           
            //fechando documento para que seja salva as alteraçoes.
            doc.Close();

            //Exibindo o Documento salvo 
            System.Diagnostics.Process.Start(caminho);
        }
    }
}
