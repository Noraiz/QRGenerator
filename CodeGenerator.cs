using QRCodeEncoderLibrary;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QRGenerator
{
    class CodeGenerator
    {
        public static void GenerateQR(string phrase, string path)
        {
            // create QR Code encoder object
            QREncoder Encoder = new QREncoder();
            Encoder.ErrorCorrection = ErrorCorrection.Q;
            Encoder.ModuleSize = 21;
            Encoder.QuietZone = 84;

            // encode input text string 
            // Note: there are 4 Encode methods in total
            Encoder.Encode(phrase);

            // save the barcode to PNG file
            // This method DOES NOT use Bitmap class and is suitable for net-core and net-standard
            // It produces files significantly smaller than SaveQRCodeToFile.
            Encoder.SaveQRCodeToPngFile(phrase + ".png");
        }
        
        
    }
}
