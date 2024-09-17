using Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.IO;
using QRCoder;
using System.Windows.Forms;

namespace Qrample
{
    internal class QRCodesCreatorHelper
    {
        private bool InRows = true;

        public void GenerateQRCodeForSelectedRange()
        {
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            Range selectedRange = Globals.ThisAddIn.Application.Selection;

            if (selectedRange.Rows.Count == 1 || selectedRange.Columns.Count == 1)
            {
                InRows &= !(selectedRange.Columns.Count == 1);

                string qrText = ConcatenateCellValues(selectedRange);

                if (Globals.ThisAddIn.myUserControl2.CheckBox1State)
                {
                    GenerateQRCode(activeSheet, selectedRange, qrText, isPicture: true);
                }

                if (Globals.ThisAddIn.myUserControl2.CheckBox2State)
                {
                    GenerateQRCode(activeSheet, selectedRange, qrText, isPicture: false);
                }

            }
            else
            {
                MessageBox.Show("Please select cells from a single row OR a single column \n to generate the QR code.",
                    "Validation error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private string ConcatenateCellValues(Range selectedRange)
        {
            string qrText = "";
            foreach (Range cell in selectedRange.Cells)
            {
                qrText += cell.Value2?.ToString() + " ";
            }
            return qrText.Trim();
        }

        private long GetPixelsPerModuleSize()
        {
            long pixelsPerModuleSize;
            if (!long.TryParse(Globals.ThisAddIn.myUserControl2.TextBox1Text, out pixelsPerModuleSize))
            {
                pixelsPerModuleSize = 10;
            }
            return pixelsPerModuleSize;
        }

        private Range GetNextCell(Range selectedRange)
        {
            if (InRows)
            {
                return selectedRange.Offset[0, selectedRange.Columns.Count];
            }
            else
            {
                return selectedRange.Offset[selectedRange.Rows.Count, 0];
            }
        }

        private void GenerateQRCode(Worksheet ws, Range selectedRange, string qrText, bool isPicture)
        {
            if (string.IsNullOrEmpty(qrText)) return;
            QRCodeGenerator qrGenerator = new QRCodeGenerator();
            QRCodeData qrCodeData = qrGenerator.CreateQrCode(qrText, QRCodeGenerator.ECCLevel.Q);
            SvgQRCode qrCode = new SvgQRCode(qrCodeData);
            string qrCodeSvg = qrCode.GetGraphic((int)GetPixelsPerModuleSize());
            Range nextCell = GetNextCell(selectedRange);

            if (isPicture)
            {
                InsertQRPicture(ws, qrCodeSvg, nextCell);
            }
            else
            {
                InsertQRVector(ws, qrCodeSvg, nextCell);
            }
        }

        private void InsertQRPicture(Worksheet ws, string qrCodeSvg, Range nextCell)
        {
            Bitmap bitmap;
            using (var ms = new MemoryStream())
            {
                byte[] svgBytes = System.Text.Encoding.UTF8.GetBytes(qrCodeSvg);
                ms.Write(svgBytes, 0, svgBytes.Length);
                ms.Position = 0;
                var svgDocument = Svg.SvgDocument.Open<Svg.SvgDocument>(ms);
                bitmap = svgDocument.Draw();
            }
            // Set the bitmap to the clipboard and paste it
            Clipboard.SetImage(bitmap);
            nextCell.Select();
            ws.Paste();
            Pictures pictures = ws.Pictures();
            pictures.Top = nextCell.Top;
            pictures.Left = nextCell.Left;
            pictures.Width = nextCell.Width;
            pictures.Height = nextCell.Height;
            Clipboard.Clear(); // Optionally, clear the clipboard
        }

        private void InsertQRVector(Worksheet ws, string qrCodeSvg, Range nextCell)
        {
            // Save the SVG to a temporary file
            string tempFilePath = Path.GetTempFileName() + ".svg";
            File.WriteAllText(tempFilePath, qrCodeSvg);

            // Insert the SVG into the cell next to the selected cell
            Pictures pictures = ws.Pictures();
            Picture picture = pictures.Insert(tempFilePath);
            picture.Top = nextCell.Top;
            picture.Left = nextCell.Left;
            picture.Width = nextCell.Width;
            picture.Height = nextCell.Height;
            File.Delete(tempFilePath);// Optionally, delete the temporary file
        }
    }
}
