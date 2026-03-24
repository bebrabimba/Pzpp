// using Microsoft.Win32;
using System.Reflection.Metadata;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Forms.Integration;
using System.Windows.Forms;
using AxWMPLib;
using MediaPlayer;
using WMPLauncher;


namespace Media
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// AxWMPLib.AxWindowsMediaPlayer axWindowsMediaPlayer1;
    public partial class MainWindow : Window
    {
        private System.Windows.Forms.Timer timer;
        AxWMPLib.AxWindowsMediaPlayer axWindowsMediaPlayer1;
        public MainWindow()
        {
            InitializeComponent();
            axWindowsMediaPlayer1 = new AxWMPLib.AxWindowsMediaPlayer();
            axWindowsMediaPlayer1.BeginInit();
            winFormsHost.Child = axWindowsMediaPlayer1;
            axWindowsMediaPlayer1.EndInit();
            timer = new System.Windows.Forms.Timer();
            timer.Interval = 500; // Co 0.5 sekundy
            timer.Tick += Timer_Tick;
            timer.Start();
        }

        private void volumeBarScroll(object sender, System.Windows.Input.MouseEventArgs e)
        {
            int volume = (int)trackBar.Value;
            value.Content = volume.ToString();
            axWindowsMediaPlayer1.settings.volume = volume;
        }

        string[] paths, files;

        bool mouse = false;
        double tempPosition = 0;

        private void button_close_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            this.Close();
        }

        private void button_minimize_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

        private void move(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                this.DragMove();
            }
        }

        private void list(object sender, SelectionChangedEventArgs e)
        {
            if (listBox1.SelectedIndex >= 0 && axWindowsMediaPlayer1 != null && axWindowsMediaPlayer1.Created)
            {
                axWindowsMediaPlayer1.URL = paths[listBox1.SelectedIndex];
            }
        }


        private void open(object sender, MouseButtonEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = true;

            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (files == null) files = new string[0];
                if (paths == null) paths = new string[0];

                // Połączenie starej i nowej listy plików
                files = files.Concat(ofd.SafeFileNames).ToArray();
                paths = paths.Concat(ofd.FileNames).ToArray();

                foreach (var file in ofd.SafeFileNames)
                {
                    listBox1.Items.Add(file);
                }

                // Automatyczne odtwarzanie pierwszego pliku, jeśli żaden nie jest wybrany
                if (listBox1.Items.Count > 0 && listBox1.SelectedIndex == -1)
                {
                    listBox1.SelectedIndex = 0;
                    axWindowsMediaPlayer1.URL = paths[listBox1.SelectedIndex];
                    axWindowsMediaPlayer1.Ctlcontrols.play();
                }
            }
        }

        private void DragEnter(object sender, System.Windows.DragEventArgs e)
        {

        }


        private void DragDrop(object sender, System.Windows.DragEventArgs e)
        {
            string[] droppedFiles = (string[])e.Data.GetData(System.Windows.Forms.DataFormats.FileDrop);

            if (droppedFiles != null && droppedFiles.Length > 0)
            {
                if (files == null) files = new string[0];
                if (paths == null) paths = new string[0];

                // Połączenie starej i nowej listy plików
                files = files.Concat(droppedFiles.Select(System.IO.Path.GetFileName)).ToArray();
                paths = paths.Concat(droppedFiles).ToArray();

                foreach (var file in droppedFiles)
                {
                    listBox1.Items.Add(System.IO.Path.GetFileName(file));
                }

                // Automatyczne odtwarzanie pierwszego pliku, jeśli żaden nie jest wybrany
                if (listBox1.Items.Count > 0 && listBox1.SelectedIndex == -1)
                {
                    listBox1.SelectedIndex = 0;
                    axWindowsMediaPlayer1.URL = paths[listBox1.SelectedIndex];
                    axWindowsMediaPlayer1.Ctlcontrols.play();
                }

            }
        }



        private void playСlick(object sender, MouseButtonEventArgs e)
        {
            axWindowsMediaPlayer1.Ctlcontrols.play();
        }

        private void pauseСlick(object sender, MouseButtonEventArgs e)
        {
            axWindowsMediaPlayer1.Ctlcontrols.pause();
        }


        private void stopСlick(object sender, MouseButtonEventArgs e)
        {
            axWindowsMediaPlayer1.Ctlcontrols.stop();
        }


        private void nextClick(object sender, MouseButtonEventArgs e)
        {
            if (listBox1.SelectedIndex < listBox1.Items.Count - 1)
                listBox1.SelectedIndex = listBox1.SelectedIndex + 1;
        }


        private void previousClick(object sender, MouseButtonEventArgs e)
        {
            if (listBox1.SelectedIndex > 0)
                listBox1.SelectedIndex = listBox1.SelectedIndex - 1;
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            if (axWindowsMediaPlayer1 != null && axWindowsMediaPlayer1.playState == WMPLib.WMPPlayState.wmppsPlaying)
            {
                double currentPosition = axWindowsMediaPlayer1.Ctlcontrols.currentPosition;
                double duration = axWindowsMediaPlayer1.currentMedia?.duration ?? 0;

                // Aktualizacja labeli
                time.Content = FormatTime(currentPosition);
                alltime.Content = FormatTime(duration);

                // Ustawienie wartości paska postępu
                ProgressBar.Maximum = duration;
                ProgressBar.Value = currentPosition;
            }
            else if (axWindowsMediaPlayer1 != null && axWindowsMediaPlayer1.playState == WMPLib.WMPPlayState.wmppsStopped)
            {
                // Zerujemy pasek, jeśli piosenka się zatrzyma
                ProgressBar.Value = 0;
            }
        }

        private string FormatTime(double seconds)
        {
            int mins = (int)seconds / 60;
            int secs = (int)seconds % 60;
            return $"{mins:D2}:{secs:D2}";
        }

        private void ProgressBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (axWindowsMediaPlayer1.currentMedia == null) return;

            double mouseX = e.GetPosition(ProgressBar).X;
            double ratio = mouseX / ProgressBar.ActualWidth;
            double newPosition = ratio * axWindowsMediaPlayer1.currentMedia.duration;

            axWindowsMediaPlayer1.Ctlcontrols.currentPosition = newPosition;
        }
    }
}