using AForge.Video.DirectShow;
using AForge.Video;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Drawing;
using System.Windows.Forms;
using System.Net.Http;
using System.Runtime.Remoting.Contexts;
using OpenCvSharp;
using System.Diagnostics;
using OpenCvSharp.Extensions;
using ZXing.Common;
using ZXing;
using System.IO;
using System.Windows;


namespace Qrample
{
    public class AllCodesReader
    {
        private System.Windows.Forms.UserControl _userControl;
        private FilterInfoCollection videoDevices = null;
        private string previousResult = string.Empty;
        public string IPCameraAddress = string.Empty;
        public int USBCameraAddress = 0;
        private Task _previewTask;
        private CancellationTokenSource _cancellationTokenSource;
        private const int _readBarcodeEveryNFrame = 5;
        private int _currentBarcodeReadFrameCount = 0;
        private System.Drawing.Bitmap _lastFrame;
        //private readonly OpenCVQRCodeReader _qrCodeReader;
        public event EventHandler OnQRCodeRead;
        private BarcodeReader _reader;
        private bool IsItUSBCamera = true;

        public AllCodesReader(System.Windows.Forms.UserControl userControl)
        {
            _userControl = userControl;
            _reader = new BarcodeReader
            {
                AutoRotate = true,
                Options = new DecodingOptions { TryHarder = true }
            };
        }
        public void PopulateCameras(System.Windows.Forms.ComboBox comboBox)
        {
            if (comboBox == null) return;
            videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);
            var CamNum = 0;
            foreach (FilterInfo device in videoDevices)
            {
                comboBox.Items.Add(CamNum + ". Camera: " + device.Name);
                CamNum++;
            }
        }
        public void startSelectedCamera(string cameraName)
        {

            if (cameraName.Contains("http://"))
            {
                IPCameraAddress = cameraName.Remove(0, 2).Trim();
                IsItUSBCamera = false;
                StartCamera();
            }
            else
            {
                IsItUSBCamera = true;
                StartCamera();
            }
        }

        private async void StartCamera()
        {
            // Never run two parallel tasks for the webcam streaming
            if (_previewTask != null && !_previewTask.IsCompleted) return;

            var initializationSemaphore = new SemaphoreSlim(0, 1);
            _cancellationTokenSource = new CancellationTokenSource();
            _previewTask = Task.Run(async () =>
            {
                try
                {
                    var videoCapture = new VideoCapture();
                    bool resultOfOpen = false;
                    
                    if(IsItUSBCamera) { resultOfOpen = videoCapture.Open(USBCameraAddress); }
                    else { resultOfOpen = videoCapture.Open(IPCameraAddress); }
                    if (!resultOfOpen)
                    {
                        throw new ApplicationException("Cannot connect to camera");
                    }
                    using (var frame = new Mat())
                    {
                        while (!_cancellationTokenSource.IsCancellationRequested)
                        {
                            videoCapture.Read(frame);

                            if (!frame.Empty())
                            {

                                // Try read the barcode every n frames to reduce latency
                                if (_currentBarcodeReadFrameCount % _readBarcodeEveryNFrame == 0)
                                {
                                    try
                                    {
                                        var bitmap = BitmapConverter.ToBitmap(frame);
                                        using (var cloneBitmap = (Bitmap)bitmap.Clone())
                                        {
                                            var qrCodeData = _reader.Decode(cloneBitmap);
                                            if (qrCodeData != null)
                                            {
                                                InsertDecodedQR(qrCodeData.Text);
                                            }
                                        }
                                        VideoPreviewPlay(bitmap);
                                    }
                                    catch (Exception ex)
                                    {
                                        //Debug.WriteLine(ex);
                                    }
                                }
                                _currentBarcodeReadFrameCount += 1 % _readBarcodeEveryNFrame;

                                // Releases the lock on first not empty frame
                                if (initializationSemaphore != null)
                                    initializationSemaphore.Release();
                            }

                            // 30 FPS
                            await Task.Delay(33);
                        }
                    }
                    videoCapture?.Dispose();
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show("Error connecting to IP camera. Check address. " + ex.Message,
                        "Network error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
                finally {
                    if (initializationSemaphore != null)
                        initializationSemaphore.Release();
                }
            }, _cancellationTokenSource.Token);

            await initializationSemaphore.WaitAsync();
            initializationSemaphore.Dispose();
            initializationSemaphore = null;

            if (_previewTask.IsFaulted)
            {
                // To let the exceptions exit
                await _previewTask;
            }

        }
        private void VideoPreviewPlay(Bitmap bitmap)
        {
            // Use Invoke to safely update the PictureBox on the UI thread
            if (Globals.ThisAddIn.myUserControl1.checkBox5Checked)
            {
                if (Globals.ThisAddIn.myUserControl1.pictureBox.InvokeRequired)
                {
                    Globals.ThisAddIn.myUserControl1.pictureBox.Invoke(new Action(() => UpdatePictureBox(bitmap)));
                }
                else
                {
                    UpdatePictureBox(bitmap);
                }
            }
            else
            {
                Globals.ThisAddIn.myUserControl1.pictureBox.Image?.Dispose();
                Globals.ThisAddIn.myUserControl1.pictureBox.Image = null;
            }
        }
        private void UpdatePictureBox(Bitmap bitmap)
        {
            try
            {
                // Dispose the old image and assign the new one
                Globals.ThisAddIn.myUserControl1.pictureBox.Image?.Dispose();
                Globals.ThisAddIn.myUserControl1.pictureBox.Image = bitmap;
            }
            catch (Exception ex) {
                System.Windows.Forms.MessageBox.Show("Error to show preview. " + ex.Message,
                    "Preview error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }
        public async void StopCamera()
        {
            if(_cancellationTokenSource == null) return;
            // If "Dispose" gets called before Stop
            if (_cancellationTokenSource.IsCancellationRequested)
                return;

            if (!_previewTask.IsCompleted)
            {
                _cancellationTokenSource.Cancel();
                
                // Wait for it, to avoid conflicts with read/write of _lastFrame
                await _previewTask;
            }
            _previewTask?.Dispose();
        }
        public void InsertDecodedQR(string codeText)
        {
            if (string.IsNullOrEmpty(codeText)) return;

            if (!Globals.ThisAddIn.myUserControl1.checkBox3Checked || codeText != previousResult)
            {
                _userControl.Invoke(new Action(() =>
                {
                    System.Windows.Forms.Clipboard.SetText(codeText);
                    if (Globals.ThisAddIn.myUserControl1.checkBox4Checked)
                    {
                        var sheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet.Paste();
                        // TODO active worksheet focus detect 
                    }
                    else
                    {
                        SendKeys.SendWait("^{v}");

                        //TODO maybe outside Invoke()
                    }
                    previousResult = codeText;
                    if (Globals.ThisAddIn.myUserControl1.checkBox1Checked) SendKeys.Send("{ENTER}");
                    if (Globals.ThisAddIn.myUserControl1.checkBox2Checked) SendKeys.Send("{TAB}");
                    System.Windows.Forms.Clipboard.Clear();

                }));
                PlayInsertionSound();
            }
        }
        private void PlayInsertionSound()
        {
            if (Globals.ThisAddIn.myUserControl1.checkBox6Checked) new EventSoundPlayer("chimes.wav").StartPlaySound();
        }
        public async Task<bool> CheckIPCameraAsync(string url)
        {
            if (string.IsNullOrEmpty(url)) return false;

            try
            {
                using (var client = new HttpClient())
                {
                    // Send a request to the IP camera
                    HttpResponseMessage response = await client.GetAsync(url.Remove(0, 2).Trim());

                    // If the status code is OK (200), we assume it's an IP camera
                    return response.IsSuccessStatusCode;
                }
            }
            catch (Exception)
            {
                System.Windows.Forms.MessageBox.Show("Error sending HTTP request to IP camera",
                    "Network error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return false;
            }
        }
    }
}
