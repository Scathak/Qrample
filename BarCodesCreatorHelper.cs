using BarcodeStandard;
using Microsoft.Office.Interop.Excel;
using SkiaSharp;
using System;
using System.IO;
using System.Windows.Forms;

namespace Qrample
{
    internal class BarCodesCreatorHelper
    {
        public void GenerateBarCodeForSelectedRange()
        {
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            Range selectedRange = Globals.ThisAddIn.Application.Selection;

            if (selectedRange.Rows.Count == 1 || selectedRange.Columns.Count == 1)
            {
                var selected =  Globals.ThisAddIn.myUserControl2.comboBox1Selected;
                BarcodeStandard.Type codeType;
                if (!Enum.TryParse(selected, out codeType))
                {
                    MessageBox.Show($"Invalid code type name: {selected}",
                        "Validation error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }
                GenerateBarcode(activeSheet, selectedRange, codeType);
            }
            else
            {
                MessageBox.Show("Please select cells from a single row OR a single column \n to generate the QR code.",
                    "Validation error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        public void GenerateBarcode(Worksheet ws, Range usedRange, BarcodeStandard.Type barcodeFormat)
        {
            // Concatenate all cell values in the selected row/column
            string barcodeText = "";
            foreach (Range cell in usedRange.Cells)
            {
                barcodeText += cell.Value2?.ToString() + " ";
            }

            if (string.IsNullOrEmpty(barcodeText))
            {
                MessageBox.Show("Empty data to generate barcode.");
                return;
            }
            SKImage barcodeImage = null;
            Barcode barcode = new Barcode();
            try
            {
                barcodeImage = barcode.Encode(barcodeFormat, barcodeText.Trim(), SKColors.Black, SKColors.White, 200, 80);
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message, 
                    "Validation error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }
            var encoddedimage = barcodeImage.Encode();

            using (var memoryStream = new MemoryStream())
            {
                encoddedimage.SaveTo(memoryStream);
                memoryStream.Position = 0;
                using (var skBitmap = SKBitmap.Decode(memoryStream))
                {
                    using (var skImage = SKImage.FromBitmap(skBitmap))
                    {
                        SaveAndInsertBarcodeImage(ws, usedRange, skImage);
                    }
                }
            }
        }

        private void SaveAndInsertBarcodeImage(Worksheet ws, Range usedRange, SKImage skImage)
        {
            // Save the SkiaSharp image to a temporary file
            string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".png");

            using (var data = skImage.Encode(SKEncodedImageFormat.Png, 100))
            {
                using (var stream = File.OpenWrite(tempFilePath))
                {
                    data.SaveTo(stream);
                }
            }

            // Insert the image into the cell next to the selected cell
            Range nextCell = usedRange.Offset[0, usedRange.Columns.Count];
            nextCell.Select();

            // Insert the image from the file
            Pictures pictures = ws.Pictures();
            Picture picture = pictures.Insert(tempFilePath);
            picture.Left = nextCell.Left;
            picture.Top = nextCell.Top;
            picture.Width = nextCell.Width;
            picture.Height = nextCell.Height;

            // Optionally, delete the temporary file after the image is inserted
            File.Delete(tempFilePath);
        }
    }
}
