using System;
using System.IO;
using System.Threading;
using System.Windows.Media;

namespace Qrample
{
    internal class EventSoundPlayer
    {
        private readonly string _fileName;
        public EventSoundPlayer(string fileName)
        {
            _fileName = fileName;
        }
        public void StartPlaySound()
        {
            var thread1 = new Thread(new ThreadStart(EventSound));
            thread1.Start();
        }
        private void EventSound()
        {
            var uri = new Uri(Path.Combine(
                //AppDomain.CurrentDomain.BaseDirectory
                Environment.GetFolderPath(Environment.SpecialFolder.Windows)
                , "Media"
                , _fileName));
            var player = new MediaPlayer();
            player.MediaEnded += delegate {
                player.Close();
            };
            player.Open(uri);
            player.Play();
        }
    }
}
