using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using OpenCvSharp;
using Spire.Xls;
using TesseractOCR;
using TesseractOCR.Enums;

namespace AutomatedVerificationProgram
{
    internal class Ocr
    {
        public class FileTimeInfo
        {
            public string fileName;
            public DateTime fileCreateTime;
        }

        public static string ScreenCapture()
        {
            if (!Directory.Exists("./screen")) Directory.CreateDirectory("./screen");
            Bitmap myImage = new Bitmap(Screen.PrimaryScreen.Bounds.Size.Width, Screen.PrimaryScreen.Bounds.Size.Height);
            using (Graphics g = Graphics.FromImage(myImage))
            {
                g.CopyFromScreen(new System.Drawing.Point(0, 0), new System.Drawing.Point(0, 0), new System.Drawing.Size(Screen.PrimaryScreen.Bounds.Size.Width, Screen.PrimaryScreen.Bounds.Size.Height));
            }

            string imagePath = Path.Combine("./screen", "screenWebsite.png");
            myImage.Save(imagePath);
            return imagePath;
        }

        internal class ImageProcessor
        {
            public string CutImage(string imgPath, int x1, int x2, int y1, int y2)
            {
                using (Mat src = new Mat(imgPath))
                using (Mat grayImg = new Mat())
                using (Mat thrImg = new Mat())
                using (Mat addImg = new Mat())
                {
                    Mat re = src[y1, y2, x1, x2].Clone();

                    int col = re.Width;
                    int rows = re.Height;

                    Cv2.Resize(re, re, new OpenCvSharp.Size(5 * col, 5 * rows), 0, 0, InterpolationFlags.Cubic);
                    Cv2.CvtColor(re, grayImg, ColorConversionCodes.BGR2GRAY);
                    Cv2.Threshold(grayImg, thrImg, 150, 255, ThresholdTypes.Tozero);

                    Mat element1 = Cv2.GetStructuringElement(MorphShapes.Rect, new OpenCvSharp.Size(5, 5));
                    Cv2.MorphologyEx(thrImg, addImg, MorphTypes.Open, element1);
                    Cv2.ImWrite("./ocr_screen/att.png", addImg);
                    return "./ocr_screen/att.png";
                }
            }
        }

        public static int[] FindScenario(int[] tentativeScenariosResult, int x1, int x2, int y1, int y2)
        {
            string imagePath = ScreenCapture();
            ImageProcessor processor = new ImageProcessor();
            string processedImage = processor.CutImage(imagePath, x1, x2, y1, y2);

            using (var engine = new Engine(@"./tessdata", Language.English, EngineMode.TesseractAndLstm))
            using (var imgOcr = TesseractOCR.Pix.Image.LoadFromFile(processedImage))
            {
                var response = engine.Process(imgOcr);
                string ocrResult = response.Text;

                string classesPath = "classes.xls";
                Workbook classesExcel = new Workbook();
                classesExcel.LoadFromFile(classesPath);
                Worksheet classesExcelSheet = classesExcel.Worksheets[0];
                int rowCount = classesExcelSheet.Rows.Length;
                string homogeneousType = "";
                string classesMetal = "";
                string ScenariosResult = "";

                int[] fimalscenarioResult = new int[2];
                for (int i = 0; i < 2; i++)
                {
                    fimalscenarioResult[i] = -1;
                }

                for (int i = 1; i < rowCount; i++)
                {
                    homogeneousType = classesExcelSheet.Range["A" + i].Value;
                    classesMetal = classesExcelSheet.Range["B" + i].Value;
                    ScenariosResult = classesExcelSheet.Range["D" + i].Value;
                    if (homogeneousType == ocrResult)
                    {
                        for (int j = 0; j < tentativeScenariosResult.Length; j++)
                        {
                            if (tentativeScenariosResult[j] == 1)
                            {
                                if (j == 0 && ScenariosResult.Contains("1a"))
                                {
                                    fimalscenarioResult[0] = 0;
                                    if (classesMetal.Contains("金屬"))
                                    {
                                        fimalscenarioResult[1] = 0;
                                    }
                                    else
                                    {
                                        fimalscenarioResult[1] = 1;
                                    }
                                    break;
                                }
                                else if (j == 1 && ScenariosResult.Contains("1b"))
                                {
                                    fimalscenarioResult[0] = 1;
                                    if (classesMetal.Contains("金屬"))
                                    {
                                        fimalscenarioResult[1] = 0;
                                    }
                                    else
                                    {
                                        fimalscenarioResult[1] = 1;
                                    }
                                    break;
                                }
                                else if (j == 2 && ScenariosResult.Contains("2"))
                                {
                                    fimalscenarioResult[0] = 2;
                                    if (classesMetal.Contains("金屬"))
                                    {
                                        fimalscenarioResult[1] = 0;
                                    }
                                    else
                                    {
                                        fimalscenarioResult[1] = 1;
                                    }
                                    break;
                                }
                                else if (j == 3 && ScenariosResult.Contains("3a"))
                                {
                                    fimalscenarioResult[0] = 3;
                                    if (classesMetal.Contains("金屬"))
                                    {
                                        fimalscenarioResult[1] = 0;
                                    }
                                    else
                                    {
                                        fimalscenarioResult[1] = 1;
                                    }
                                    break;
                                }
                                else if (j == 4 && ScenariosResult.Contains("3b"))
                                {
                                    fimalscenarioResult[0] = 4;
                                    if (classesMetal.Contains("金屬"))
                                    {
                                        fimalscenarioResult[1] = 0;
                                    }
                                    else
                                    {
                                        fimalscenarioResult[1] = 1;
                                    }
                                    break;
                                }

                            }

                        }
                        return fimalscenarioResult;
                    }
                }
                return fimalscenarioResult;

            }
        }

            public static string OcrString(int x1, int x2, int y1, int y2)
            {
                string imagePath = ScreenCapture();
                ImageProcessor processor = new ImageProcessor();
                string processedImage = processor.CutImage(imagePath, x1, x2, y1, y2);

                using (var engine = new Engine(@"./tessdata", Language.English, EngineMode.TesseractAndLstm))
                using (var imgOcr = TesseractOCR.Pix.Image.LoadFromFile(processedImage))
                {
                    var response = engine.Process(imgOcr);
                    string ocrResult = response.Text;
                    return ocrResult;
                }
            } 
    }
}
