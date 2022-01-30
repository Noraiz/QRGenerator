using Independentsoft.Office.Odf;
using Independentsoft.Office.Odf.Drawing;
using Independentsoft.Office.Odf.Styles;
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

            TextDocument doc = new TextDocument();

            ParagraphStyle fontStyle = new ParagraphStyle("P100");
            fontStyle.TextProperties.Font = "Arial";
            fontStyle.TextProperties.FontSize = new Size(5,Unit.Point);
            fontStyle.ParagraphProperties.TextAlignment = TextAlignment.Center;
            doc.AutomaticStyles.Styles.Add(fontStyle);


            ParagraphStyle imgStyle = new ParagraphStyle("P200");
            imgStyle.ParagraphProperties.TextAlignment = TextAlignment.Center;
            doc.AutomaticStyles.Styles.Add(imgStyle);

            Table table = new Table();
            for (int i = 0; i < Rows; i++)
            {
                Row row = new Row();
                for (int j = 0; j < Columns; j++)
                {
                int index = (i * Columns) + j;
                    
                    Cell cell;
                    if (index >= listQr.Count)
                    {
                        Console.WriteLine("In" + index);
                        cell = new Cell("");
                    }
                    else
                    {
                        var qr = listQr.ElementAt(index);
                        cell = GetCellValue(qr.Phrase);
                    }
                        row.Cells.Add(cell);
                }
                table.Rows.Add(row);

            }

            doc.Body.Add(table);
            doc.Save("QRFile.odt",true);
        }

        public static Cell GetCellValue(String phrase)
        {
            Cell cell = new Cell();

            double width, height;
            width = height = 1.058;

            Image image1 = new Image(phrase+".png");
            Frame frame1 = new Frame();
            frame1.Style = "fr1";
            frame1.Width = new Size(width, Unit.Centimeter);
            frame1.Height = new Size(height, Unit.Centimeter);
            frame1.Add(image1);

            Paragraph qrImage = new Paragraph();
            qrImage.Add(frame1);
            qrImage.Style = "P200";

            Paragraph label = new Paragraph();
            label.Add(phrase);
            label.Style = "P100";

            cell.Content.Add(qrImage);
            cell.Content.Add(label);
            return cell;
        }
    }
}
