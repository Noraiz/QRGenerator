using AODL.Document.Content.Draw;
using AODL.Document.Content.Tables;
using AODL.Document.Content.Text;
using AODL.Document.Styles;
using AODL.Document.Styles.Properties;
using AODL.Document.TextDocuments;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace QRGenerator
{
    class OdtMaker
    {

        public static void GenerateOdt(List<QRModel> listQr)
        {
            int Columns = 11;
            int Rows = (listQr.Count / Columns) + 1;
            Console.WriteLine(Rows);
            TextDocument document = new TextDocument();
            document.New();
            document.FontList.Add(FontFamilies.Arial);
            //Create a table for a text document using the TableBuilder
            Table table = TableBuilder.CreateTextDocumentTable(
                document,
                "table1",
                "table1",
                Rows,
                Columns,
                16.99,
                false,
                true);

            int index = 0;
            //Fill the cells
            foreach (Row row in table.RowCollection)
            {
                foreach (Cell cell in row.CellCollection)
                {
                    Paragraph paragraph = ParagraphBuilder.CreateStandardTextParagraph(document);
                    var p2 = ParagraphBuilder.CreateParagraphWithCustomStyle(document, "centralized");
                    if (index >= listQr.Count) {
                        Console.WriteLine("In" + index);
                        paragraph.TextContent.Add(new SimpleText(document, "\n")); 
                    }
                    else
                    {
                       
                        var qr = listQr.ElementAt(index);
                        Console.WriteLine(qr.Phrase);
                        //Create a standard paragraph
                        Frame frame = new Frame(document, "frame1", qr.Phrase, qr.Phrase + ".png");
                        frame.SvgHeight = "0.40in";
                        frame.SvgWidth = "0.41in";
                        //Add some simple text

                        var label = new FormatedText(document,"Label",qr.Phrase);

                        label.TextStyle.TextProperties.FontSize = "5pt";
                        label.TextStyle.TextProperties.FontName = FontFamilies.Arial;
                        label.TextStyle.TextProperties.Position = "center";
                        //label.Style = new ParagraphStyle(document, "central");
                        //frame.Style = new ParagraphStyle(document, "central");

                        //((ParagraphStyle) label.Style).ParagraphProperties.Alignment = TextAlignments.center.ToString();

                        //((ParagraphStyle) frame.Style).ParagraphProperties.Alignment = TextAlignments.center.ToString();


                        //paragraph.ParagraphStyle = new ParagraphStyle(document, "central");
                        //paragraph.ParagraphStyle.ParagraphProperties.Alignment = TextAlignments.center.ToString();

                        ParagraphStyle ps1 = new ParagraphStyle(document, "style1");
                        ps1.Family = "paragraph";
                        ps1.TextProperties.FontName = FontFamilies.Arial;
                        ps1.TextProperties.FontSize = "5pt";
                        ps1.ParagraphProperties.Alignment = TextAlignments.center.ToString();
                        p2.ParagraphStyle = ps1;

                        paragraph.Content.Add(frame);
                        p2.TextContent.Add(label);

                    }
                   
                    //((CellStyle)cell.Style).CellProperties.Padding = "0.2cm";
                    cell.Content.Add(paragraph);
                    cell.Content.Add(p2);
                    index++;
                }
            }
            //Merge some cells. Notice this is only available in text documents!
            //table.RowCollection[1].MergeCells(document, 1, 2, true);
            
            //Add table to the document
            document.Content.Add(table);

            //Save the document
            document.SaveTo("QRFile.odt");

        }

    }
}
