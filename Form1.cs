using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Windows.Media.Ocr;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            var engine = OcrEngine.TryCreateFromLanguage(new Windows.Globalization.Language("en-US"));
            //string filePath = TestData.GetFilePath("testimage.png");
            var file =   Windows.Storage.StorageFile.GetFileFromPathAsync(@"C:\Cases\3333.jpg").GetAwaiter().GetResult();  
            var stream =   file.OpenAsync(Windows.Storage.FileAccessMode.Read).GetAwaiter().GetResult(); 
            var decoder =   Windows.Graphics.Imaging.BitmapDecoder.CreateAsync(stream).GetAwaiter().GetResult(); 
            var softwareBitmap =   decoder.GetSoftwareBitmapAsync().GetAwaiter().GetResult(); 
            var ocrResult =  engine.RecognizeAsync(softwareBitmap).GetAwaiter().GetResult();
            List<OcrLine> lines = ocrResult.Lines.ToList();

            StringBuilder sb = new StringBuilder();


            foreach (var line in lines)
            {
                sb.Append(line.Text);

                sb.Append("\n\n");



            }

              Console.WriteLine(sb.ToString());
            //CreateDocument(words);


        }


        
    }
}
