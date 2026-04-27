using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Forms.Integration;
using System.Windows.Media.Imaging;
using AxWMPLib;
using WMPLib;
using WinForms = System.Windows.Forms;

namespace Media
{
    public partial class MainWindow : Window
    {
        private WinForms.Timer timer;
        private AxWindowsMediaPlayer axWindowsMediaPlayer1;
        public ObservableCollection<SongMetadata> QueueItems { get; } = new();
        private string[] paths = Array.Empty<string>();
        private string[] files = Array.Empty<string>();
        private readonly Dictionary<string, SongMetadata> metadataCache = new();
        private bool isDraggingVolume;
        private bool suppressSelectionChanged;
        private string sortColumn = "Number";
        private bool sortAscending = true;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            ConfigureQueueView();

            axWindowsMediaPlayer1 = new AxWindowsMediaPlayer();
            axWindowsMediaPlayer1.BeginInit();
            winFormsHost.Child = axWindowsMediaPlayer1;
            axWindowsMediaPlayer1.EndInit();
            axWindowsMediaPlayer1.PlayStateChange += AxWindowsMediaPlayer1_PlayStateChange;
            axWindowsMediaPlayer1.settings.volume = (int)trackBar.Value;

            timer = new WinForms.Timer();
            timer.Interval = 500;
            timer.Tick += Timer_Tick;
            timer.Start();

            UpdatePlayPauseButtons(false);
        }

        private void button_close_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) => Close();

        private void button_minimize_MouseLeftButtonDown(object sender, MouseButtonEventArgs e) => WindowState = WindowState.Minimized;

        private void move(object sender, System.Windows.Input.MouseEventArgs e)
        {
            // Nie przechwytuj przeciągania kontrolek wewnątrz okna (np. scrollbarów i sliderów).
            // DragMove zostaje tylko dla górnego paska okna.
            if (e.LeftButton == MouseButtonState.Pressed && sender is Grid)
                DragMove();
        }

        private void open(object sender, MouseButtonEventArgs e)
        {
            WinForms.OpenFileDialog ofd = new WinForms.OpenFileDialog();
            ofd.Multiselect = true;
            ofd.Filter = "Pliki audio|*.mp3;*.wav;*.wma;*.aac;*.m4a;*.flac|Wszystkie pliki|*.*";

            if (ofd.ShowDialog() == WinForms.DialogResult.OK)
                AddFiles(ofd.FileNames);
        }

        private void DragEnter(object sender, System.Windows.DragEventArgs e)
        {
            if (e.Data.GetDataPresent(System.Windows.DataFormats.FileDrop))
                e.Effects = System.Windows.DragDropEffects.Copy;
        }

        private void DragDrop(object sender, System.Windows.DragEventArgs e)
        {
            string[] droppedFiles = (string[])e.Data.GetData(System.Windows.DataFormats.FileDrop);
            if (droppedFiles != null && droppedFiles.Length > 0)
                AddFiles(droppedFiles);
        }

        private async void AddFiles(string[] newPaths)
        {
            // Nie dodawaj ponownie tych samych plików.
            var existingPaths = new HashSet<string>(QueueItems.Select(x => Path.GetFullPath(x.FilePath)), StringComparer.OrdinalIgnoreCase);
            var uniquePaths = new List<string>();

            foreach (string path in newPaths.Where(File.Exists))
            {
                string fullPath = Path.GetFullPath(path);
                if (existingPaths.Add(fullPath))
                    uniquePaths.Add(fullPath);
            }

            if (uniquePaths.Count == 0)
                return;

            foreach (string file in uniquePaths)
                QueueItems.Add(CreateFallbackSong(file));

            SyncArraysFromQueue();

            // Najpierw metadane pierwszego nowego pliku, potem wybór/odtworzenie.
            if (listBox1.SelectedItem == null)
            {
                string firstPath = uniquePaths[0];
                await LoadMetadataAsync(firstPath);
                SelectSongByPath(firstPath);
            }

            ApplyCurrentSortPreservingSelection();

            foreach (string file in uniquePaths)
                _ = LoadMetadataAsync(file);
        }

        private void list(object sender, SelectionChangedEventArgs e)
        {
            if (suppressSelectionChanged) return;

            if (listBox1.SelectedIndex >= 0)
                PlaySelectedSong();
        }

        private async void PlaySelectedSong()
        {
            if (listBox1.SelectedItem is not SongMetadata selected)
                return;

            string filePath = selected.FilePath;

            if (!metadataCache.ContainsKey(filePath))
                await LoadMetadataAsync(filePath, refreshPlayer: true);
            else
                ApplyMetadata(filePath, metadataCache[filePath]);

            axWindowsMediaPlayer1.URL = filePath;
            axWindowsMediaPlayer1.Ctlcontrols.play();
            UpdatePlayPauseButtons(true);
        }

        private async Task LoadMetadataAsync(string filePath, bool refreshPlayer = false)
        {
            if (metadataCache.ContainsKey(filePath))
            {
                if (refreshPlayer) ApplyMetadata(filePath, metadataCache[filePath]);
                return;
            }

            SongMetadata metadata = await Task.Run(() => ReadMetadata(filePath));
            metadataCache[filePath] = metadata;

            Dispatcher.Invoke(() =>
            {
                int index = QueueItems.ToList().FindIndex(x => string.Equals(x.FilePath, filePath, StringComparison.OrdinalIgnoreCase));
                if (index >= 0)
                {
                    metadata.FilePath = filePath;
                    QueueItems[index] = metadata;
                }

                SyncArraysFromQueue();
                ApplyCurrentSortPreservingSelection();

                if (refreshPlayer && listBox1.SelectedItem is SongMetadata selected &&
                    string.Equals(selected.FilePath, filePath, StringComparison.OrdinalIgnoreCase))
                    ApplyMetadata(filePath, metadata);
            });
        }

        private SongMetadata ReadMetadata(string filePath)
        {
            string fallback = Path.GetFileNameWithoutExtension(filePath);

            try
            {
                using var tagFile = TagLib.File.Create(filePath);
                string title = string.IsNullOrWhiteSpace(tagFile.Tag.Title) ? fallback : tagFile.Tag.Title.Trim();
                string artist = tagFile.Tag.FirstPerformer;
                if (string.IsNullOrWhiteSpace(artist)) artist = tagFile.Tag.FirstAlbumArtist;
                if (string.IsNullOrWhiteSpace(artist)) artist = "Nieznany autor";

                BitmapImage? cover = null;
                if (tagFile.Tag.Pictures != null && tagFile.Tag.Pictures.Length > 0)
                    cover = CreateBitmapFromBytes(tagFile.Tag.Pictures[0].Data.Data);

                string album = string.IsNullOrWhiteSpace(tagFile.Tag.Album) ? "-" : tagFile.Tag.Album.Trim();
                string year = tagFile.Tag.Year > 0 ? tagFile.Tag.Year.ToString() : "-";
                string number = tagFile.Tag.Track > 0 ? tagFile.Tag.Track.ToString() : "-";
                double duration = tagFile.Properties.Duration.TotalSeconds;

                return new SongMetadata
                {
                    Number = number,
                    FilePath = filePath,
                    FolderPath = Path.GetDirectoryName(filePath) ?? string.Empty,
                    FolderDisplay = GetFolderDisplay(filePath),
                    Title = title,
                    Artist = artist.Trim(),
                    Album = album,
                    Year = year,
                    Duration = duration,
                    DurationText = duration > 0 ? FormatTime(duration) : "-",
                    Cover = cover,
                    ListTitle = string.IsNullOrWhiteSpace(artist) || artist == "Nieznany autor" ? title : $"{title} - {artist}"
                };
            }
            catch
            {
                return CreateFallbackSong(filePath);
            }
        }

        private SongMetadata CreateFallbackSong(string filePath)
        {
            return new SongMetadata
            {
                Number = "-",
                FilePath = filePath,
                FolderPath = Path.GetDirectoryName(filePath) ?? string.Empty,
                FolderDisplay = GetFolderDisplay(filePath),
                Title = Path.GetFileNameWithoutExtension(filePath),
                Artist = "Ładowanie...",
                Album = "-",
                Year = "-",
                Duration = 0,
                DurationText = "-",
                Cover = null,
                ListTitle = Path.GetFileName(filePath)
            };
        }

        private BitmapImage? CreateBitmapFromBytes(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0) return null;

            using var stream = new MemoryStream(bytes);
            var image = new BitmapImage();
            image.BeginInit();
            image.CacheOption = BitmapCacheOption.OnLoad;
            image.StreamSource = stream;
            image.EndInit();
            image.Freeze();
            return image;
        }

        private void ShowFallbackMetadata(string filePath)
        {
            NowTitle.Text = Path.GetFileNameWithoutExtension(filePath);
            NowArtist.Text = "Ładowanie metadanych...";
            alltime.Text = "00:00";
            AlbumArtImage.Source = null;
            AlbumArtImage.Visibility = Visibility.Collapsed;
            AlbumArtPlaceholder.Visibility = Visibility.Visible;
        }

        private void ApplyMetadata(string filePath, SongMetadata metadata)
        {
            NowTitle.Text = metadata.Title;
            NowArtist.Text = metadata.Artist;

            if (metadata.Duration > 0)
                alltime.Text = FormatTime(metadata.Duration);

            if (metadata.Cover != null)
            {
                AlbumArtImage.Source = metadata.Cover;
                AlbumArtImage.Visibility = Visibility.Visible;
                AlbumArtPlaceholder.Visibility = Visibility.Collapsed;
            }
            else
            {
                AlbumArtImage.Source = null;
                AlbumArtImage.Visibility = Visibility.Collapsed;
                AlbumArtPlaceholder.Visibility = Visibility.Visible;
            }
        }

        private void UpdateNowPlayingPanel()
        {
            if (listBox1.SelectedItem is not SongMetadata selected)
            {
                NowTitle.Text = "Wybierz utwór";
                NowArtist.Text = "Artysta";
                alltime.Text = "00:00";
                return;
            }

            string filePath = selected.FilePath;
            if (metadataCache.TryGetValue(filePath, out SongMetadata metadata))
            {
                ApplyMetadata(filePath, metadata);
                return;
            }

            ShowFallbackMetadata(filePath);
            _ = LoadMetadataAsync(filePath, refreshPlayer: true);
        }

        private void playСlick(object sender, MouseButtonEventArgs e)
        {
            if (axWindowsMediaPlayer1.currentMedia == null && listBox1.SelectedIndex >= 0)
                PlaySelectedSong();
            else
                axWindowsMediaPlayer1.Ctlcontrols.play();

            UpdatePlayPauseButtons(true);
        }

        private void pauseСlick(object sender, MouseButtonEventArgs e)
        {
            axWindowsMediaPlayer1.Ctlcontrols.pause();
            UpdatePlayPauseButtons(false);
        }

        private void nextClick(object sender, MouseButtonEventArgs e)
        {
            if (listBox1.SelectedIndex < listBox1.Items.Count - 1)
                listBox1.SelectedIndex++;
        }

        private void previousClick(object sender, MouseButtonEventArgs e)
        {
            if (listBox1.SelectedIndex > 0)
                listBox1.SelectedIndex--;
        }

        private void AxWindowsMediaPlayer1_PlayStateChange(object sender, _WMPOCXEvents_PlayStateChangeEvent e)
        {
            bool isPlaying = e.newState == (int)WMPPlayState.wmppsPlaying;
            Dispatcher.Invoke(() =>
            {
                UpdatePlayPauseButtons(isPlaying);
                UpdateNowPlayingPanel();
            });
        }

        private void UpdatePlayPauseButtons(bool isPlaying)
        {
            play.Visibility = isPlaying ? Visibility.Collapsed : Visibility.Visible;
            pause.Visibility = isPlaying ? Visibility.Visible : Visibility.Collapsed;
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            if (axWindowsMediaPlayer1?.currentMedia == null)
                return;

            double currentPosition = axWindowsMediaPlayer1.Ctlcontrols.currentPosition;
            double duration = axWindowsMediaPlayer1.currentMedia.duration;

            time.Text = FormatTime(currentPosition);
            alltime.Text = duration > 0 ? FormatTime(duration) : alltime.Text;

            ProgressBar.Maximum = duration > 0 ? duration : 1;
            ProgressBar.Value = Math.Min(currentPosition, ProgressBar.Maximum);

            if (axWindowsMediaPlayer1.playState == WMPPlayState.wmppsStopped)
            {
                ProgressBar.Value = 0;
                time.Text = "00:00";
                UpdatePlayPauseButtons(false);
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

        private void trackBar_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            isDraggingVolume = true;
            trackBar.CaptureMouse();
            SetVolumeByMousePosition(e);
            e.Handled = true;
        }

        private void trackBar_PreviewMouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (!isDraggingVolume) return;
            SetVolumeByMousePosition(e);
            e.Handled = true;
        }

        private void trackBar_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (!isDraggingVolume) return;
            SetVolumeByMousePosition(e);
            isDraggingVolume = false;
            trackBar.ReleaseMouseCapture();
            e.Handled = true;
        }

        private void SetVolumeByMousePosition(System.Windows.Input.MouseEventArgs e)
        {
            double x = e.GetPosition(trackBar).X;
            double ratio = Math.Max(0, Math.Min(1, x / trackBar.ActualWidth));
            trackBar.Value = Math.Round(trackBar.Minimum + ratio * (trackBar.Maximum - trackBar.Minimum));
        }

        private void trackBar_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e) => SetVolumeFromSlider();

        private void SetVolumeFromSlider()
        {
            if (axWindowsMediaPlayer1 == null || value == null) return;

            int volume = (int)Math.Round(trackBar.Value);
            value.Content = volume.ToString();
            axWindowsMediaPlayer1.settings.volume = volume;
        }


        private void RemoveButton_Click(object sender, RoutedEventArgs e)
        {
            if (listBox1.SelectedItem is not SongMetadata selected) return;

            int selectedViewIndex = listBox1.SelectedIndex;
            bool removingCurrent = axWindowsMediaPlayer1.currentMedia != null &&
                                   string.Equals(axWindowsMediaPlayer1.URL, selected.FilePath, StringComparison.OrdinalIgnoreCase);

            QueueItems.Remove(selected);
            metadataCache.Remove(selected.FilePath);
            SyncArraysFromQueue();

            if (QueueItems.Count == 0)
            {
                axWindowsMediaPlayer1.Ctlcontrols.stop();
                axWindowsMediaPlayer1.URL = string.Empty;
                NowTitle.Text = "Wybierz utwór";
                NowArtist.Text = "Artysta";
                AlbumArtImage.Source = null;
                AlbumArtImage.Visibility = Visibility.Collapsed;
                AlbumArtPlaceholder.Visibility = Visibility.Visible;
                ProgressBar.Value = 0;
                time.Text = "00:00";
                alltime.Text = "00:00";
                UpdatePlayPauseButtons(false);
                return;
            }

            listBox1.SelectedIndex = Math.Min(selectedViewIndex, listBox1.Items.Count - 1);
            if (removingCurrent)
                PlaySelectedSong();
        }

        private void HeaderSort_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is not TextBlock header || header.Tag is not string column) return;

            if (sortColumn == column)
                sortAscending = !sortAscending;
            else
            {
                sortColumn = column;
                sortAscending = true;
            }

            ApplyCurrentSortPreservingSelection();
        }

        private void ApplyCurrentSortPreservingSelection()
        {
            string? selectedPath = (listBox1.SelectedItem as SongMetadata)?.FilePath;

            var view = CollectionViewSource.GetDefaultView(QueueItems);
            if (view == null) return;

            suppressSelectionChanged = true;
            view.SortDescriptions.Clear();
            view.GroupDescriptions.Clear();
            view.GroupDescriptions.Add(new PropertyGroupDescription(nameof(SongMetadata.FolderDisplay)));
            view.SortDescriptions.Add(new SortDescription(nameof(SongMetadata.FolderDisplay), ListSortDirection.Ascending));
            view.SortDescriptions.Add(new SortDescription(SortPropertyName(sortColumn), sortAscending ? ListSortDirection.Ascending : ListSortDirection.Descending));
            view.Refresh();
            SyncArraysFromQueue();

            if (!string.IsNullOrWhiteSpace(selectedPath))
                SelectSongByPath(selectedPath);
            suppressSelectionChanged = false;
        }

        private void ConfigureQueueView()
        {
            var view = CollectionViewSource.GetDefaultView(QueueItems);
            view.GroupDescriptions.Add(new PropertyGroupDescription(nameof(SongMetadata.FolderDisplay)));
            view.SortDescriptions.Add(new SortDescription(nameof(SongMetadata.FolderDisplay), ListSortDirection.Ascending));
            view.SortDescriptions.Add(new SortDescription(nameof(SongMetadata.NumberValue), ListSortDirection.Ascending));
        }

        private string SortPropertyName(string column)
        {
            return column switch
            {
                "Number" => nameof(SongMetadata.NumberValue),
                "Year" => nameof(SongMetadata.YearValue),
                "Duration" => nameof(SongMetadata.Duration),
                "Album" => nameof(SongMetadata.Album),
                "Artist" => nameof(SongMetadata.Artist),
                _ => nameof(SongMetadata.Title),
            };
        }

        private void SyncArraysFromQueue()
        {
            paths = QueueItems.Select(x => x.FilePath).ToArray();
            files = QueueItems.Select(x => Path.GetFileName(x.FilePath)).ToArray();
        }

        private void SelectSongByPath(string filePath)
        {
            foreach (var item in listBox1.Items)
            {
                if (item is SongMetadata song && string.Equals(song.FilePath, filePath, StringComparison.OrdinalIgnoreCase))
                {
                    listBox1.SelectedItem = song;
                    listBox1.ScrollIntoView(song);
                    break;
                }
            }
        }

        private int CompareSongs(SongMetadata a, SongMetadata b, string column)
        {
            return column switch
            {
                "Number" => CompareNullableNumbers(a.Number, b.Number),
                "Year" => CompareNullableNumbers(a.Year, b.Year),
                "Duration" => a.Duration.CompareTo(b.Duration),
                "Album" => string.Compare(a.Album, b.Album, StringComparison.CurrentCultureIgnoreCase),
                "Artist" => string.Compare(a.Artist, b.Artist, StringComparison.CurrentCultureIgnoreCase),
                _ => string.Compare(a.Title, b.Title, StringComparison.CurrentCultureIgnoreCase),
            };
        }

        private int CompareNullableNumbers(string left, string right)
        {
            bool leftOk = int.TryParse(left, out int leftNumber);
            bool rightOk = int.TryParse(right, out int rightNumber);

            if (leftOk && rightOk) return leftNumber.CompareTo(rightNumber);
            if (leftOk) return -1;
            if (rightOk) return 1;
            return string.Compare(left, right, StringComparison.CurrentCultureIgnoreCase);
        }

        private string GetFolderDisplay(string filePath)
        {
            string folder = Path.GetDirectoryName(filePath) ?? string.Empty;
            string folderName = Path.GetFileName(folder);
            return string.IsNullOrWhiteSpace(folderName) ? folder : folderName;
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e) { }

        private void Category_Click(object sender, RoutedEventArgs e) { }

        public class SongMetadata
        {
            public string Number { get; set; } = "-";
            public string FilePath { get; set; } = string.Empty;
            public string FolderPath { get; set; } = string.Empty;
            public string FolderDisplay { get; set; } = string.Empty;
            public string Title { get; set; } = string.Empty;
            public string Album { get; set; } = "-";
            public string Artist { get; set; } = string.Empty;
            public string Year { get; set; } = "-";
            public string DurationText { get; set; } = "-";
            public string ListTitle { get; set; } = string.Empty;
            public double Duration { get; set; }
            public BitmapImage? Cover { get; set; }
            public int NumberValue => int.TryParse(Number, out int n) ? n : int.MaxValue;
            public int YearValue => int.TryParse(Year, out int y) ? y : int.MaxValue;
        }
    }
}
