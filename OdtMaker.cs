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
        private static TextDocument doc;
        public static void GenerateOdt(List<QRModel> listQr)
        {
            int Columns = 11;
            int Rows = (listQr.Count / Columns) + 1;

            doc = new TextDocument();

            PageLayout pl = new PageLayout();
            pl.Name = "Layout";
            pl.PageLayoutProperties.TopMargin = new Size
            {
                Value = 0.32,
                Unit = Unit.Inch
            };
            pl.PageLayoutProperties.LeftMargin = new Size
            {
                Value = 0.32,
                Unit = Unit.Inch
            };
            pl.PageLayoutProperties.RightMargin = new Size
            {
                Value = 0.39,
                Unit = Unit.Inch
            };

            doc.CommonStyles.AutomaticStyles = new AutomaticStyles();
            doc.CommonStyles.AutomaticStyles.PageLayouts.Add(pl);

            Font arial = new Font();
            arial.Name = "Arial";
            arial.Family = "Arial";
            arial.GenericFontFamily = GenericFontFamily.Swiss;
            arial.Pitch = FontPitch.Variable;

            doc.Fonts.Add(arial);

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

            MasterPage masterPage1 = new MasterPage();
            masterPage1.Name = "Standard";
            masterPage1.PageLayout = "Layout";

            doc.CommonStyles.MasterStyles = new MasterStyles();
            doc.CommonStyles.MasterStyles.MasterPages.Add(masterPage1);

            doc.Body.Add(table);
            doc.Save("QRFile.odt",true);
        }

        public static Cell GetCellValue(String phrase)
        {
            Cell cell = new Cell();

            CellStyle cellStyle = new CellStyle("cs");
            
            cellStyle.CellProperties.TopPadding = new Size { Unit = Unit.Inch, Value = 0.142 };
            cellStyle.CellProperties.BottomPadding = new Size { Unit = Unit.Inch, Value = 0.142 };
            doc.AutomaticStyles.Styles.Add(cellStyle);
            double width, height;
            width = height = 0.42;

            Image image1 = new Image(phrase+".png");
            Frame frame1 = new Frame();
            frame1.Style = "fr1";
            frame1.Width = new Size(width, Unit.Inch);
            frame1.Height = new Size(height, Unit.Inch);
            frame1.Add(image1);

            Paragraph qrImage = new Paragraph();
            qrImage.Add(frame1);
            qrImage.Style = "P200";

            Paragraph label = new Paragraph();
            label.Add(phrase);
            label.Style = "P100";
            cell.Style = "cs";
            cell.Content.Add(qrImage);
            cell.Content.Add(label);
            return cell;
        }
    }
}
