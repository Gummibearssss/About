using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using static System.Net.Mime.MediaTypeNames;
using Application = Microsoft.Office.Interop.PowerPoint.Application;

namespace InvitesExportToPPT
{
    class Program
    {
        const float _pixConvCm = 28.34645669291339F;

        static void Main(string[] args)
        {
            //Console.WriteLine("Hello World!");

            Application pptApplication = new Application();

            Microsoft.Office.Interop.PowerPoint.Slides slides;
            Microsoft.Office.Interop.PowerPoint._Slide slide;

            // Create the Presentation File
            Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoTrue);
            pptPresentation.Final = false;
            pptPresentation.PageSetup.NotesOrientation = MsoOrientation.msoOrientationVertical;
            pptPresentation.PageSetup.SlideOrientation = MsoOrientation.msoOrientationVertical;
            pptPresentation.PageSetup.SlideHeight = _pixConvCm * 29.7F;
            pptPresentation.PageSetup.SlideWidth = _pixConvCm * 20.999F;

            Microsoft.Office.Interop.PowerPoint.CustomLayout customLayout = pptPresentation.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutOrgchart];

            // Create new Slide
            slides = pptPresentation.Slides;

            int pixIndex = 0;

            for (int i = 0; i < 25; i++)
            {
                slide = slides.AddSlide(1, customLayout);

                pixIndex++;

                slide.Shapes.AddPicture(@"C:\Users\Kenny.Lim\Downloads\D&K\" + pixIndex.ToString() + ".png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue,
                    (float)_pixConvCm * 0.99F, (float)_pixConvCm * 0.8F, (float)_pixConvCm * 12.03F, (float)_pixConvCm * 6.78F);

                pixIndex++;

                slide.Shapes.AddPicture(@"C:\Users\Kenny.Lim\Downloads\D&K\" + pixIndex.ToString() + ".png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue,
                    (float)_pixConvCm * 0.99F, (float)_pixConvCm * 7.83F, (float)_pixConvCm * 12.03F, (float)_pixConvCm * 6.78F);

                pixIndex++;

                slide.Shapes.AddPicture(@"C:\Users\Kenny.Lim\Downloads\D&K\" + pixIndex.ToString() + ".png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue,
                    (float)_pixConvCm * 0.99F, (float)_pixConvCm * 14.86F, (float)_pixConvCm * 12.03F, (float)_pixConvCm * 6.78F);

                pixIndex++;

                slide.Shapes.AddPicture(@"C:\Users\Kenny.Lim\Downloads\D&K\" + pixIndex.ToString() + ".png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue,
                    (float)_pixConvCm * 0.99F, (float)_pixConvCm * 21.93F, (float)_pixConvCm * 12.03F, (float)_pixConvCm * 6.78F);

                pixIndex++;

                slide.Shapes.AddPicture(@"C:\Users\Kenny.Lim\Downloads\D&K\" + pixIndex.ToString() + ".png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue,
                    (float)_pixConvCm * 10.66F, (float)_pixConvCm * 3.43F, (float)_pixConvCm * 12.03F, (float)_pixConvCm * 6.78F).Rotation = 270;

                pixIndex++;

                slide.Shapes.AddPicture(@"C:\Users\Kenny.Lim\Downloads\D&K\" + pixIndex.ToString() + ".png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue,
                    (float)_pixConvCm * 10.66F, (float)_pixConvCm * 17.49F, (float)_pixConvCm * 12.03F, (float)_pixConvCm * 6.78F).Rotation = 270;
            }
            //slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = "This demo is created by FPPT using C# - Download free templates from http://FPPT.com";

            Console.WriteLine(pixIndex.ToString());

            pptPresentation.SaveAs(@"C:\Users\Kenny.Lim\Desktop\Wedding\VS_ExportToPPTInvites\Invites.pptx", Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
            //pptPresentation.Close();
            //pptApplication.Quit();
        }
    }
}
