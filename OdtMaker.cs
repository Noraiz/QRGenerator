using AODL.Document.Content.Draw;
using AODL.Document.Content.Tables;
using AODL.Document.Content.Text;
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
            int Rows = (listQr.Count / 8) + 1;
            Console.WriteLine(Rows);
            int Columns = 8;
            TextDocument document = new TextDocument();
            document.New();
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
                    AODL.Document.Content.Text.Paragraph paragraph = ParagraphBuilder.CreateStandardTextParagraph(document);
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
                        frame.SvgHeight = "0.6043in";
                        frame.SvgWidth = "0.6043in";
                        paragraph.Content.Add(frame);
                        //Add some simple text
                        paragraph.TextContent.Add(new SimpleText(document, "\n" + qr.Phrase));
                    }
                    cell.Content.Add(paragraph);
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
