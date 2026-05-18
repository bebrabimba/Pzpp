using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Forms.Integration;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
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
        public ObservableCollection<LibraryEntry> LibraryItems { get; } = new();
        private string libraryRootPath = string.Empty;
        private string currentLibraryPath = string.Empty;
        private readonly HashSet<string> expandedLibraryFolders = new(StringComparer.OrdinalIgnoreCase);
        private bool suppressLibrarySelectionChanged;
        private bool suppressSearchTextChanged;
        private DispatcherTimer librarySearchTimer;
        private readonly Dictionary<string, SongMetadata> metadataCache = new(StringComparer.OrdinalIgnoreCase);
        private bool isDraggingVolume;
        private bool suppressSelectionChanged;
        private string sortColumn = "Number";
        private bool sortAscending = true;
        private Point libraryDragStartPoint;
        private const string LibraryAudioDragFormat = "Media.LibraryAudioPaths";

        private static readonly string[] SupportedAudioExtensions =
        {
            ".mp3", ".wav", ".wma", ".aac", ".m4a", ".flac"
        };

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            listBox1.SelectionMode = SelectionMode.Extended;
            LibraryList.SelectionMode = SelectionMode.Extended;

            librarySearchTimer = new DispatcherTimer();
            librarySearchTimer.Interval = TimeSpan.FromSeconds(1);
            librarySearchTimer.Tick += LibrarySearchTimer_Tick;
            LoadLibraryBrowser();
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

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
                return;
            }

            if (e.LeftButton != MouseButtonState.Pressed)
                return;

            if (WindowState == WindowState.Maximized)
            {
                double ratio = ActualWidth > 0 ? e.GetPosition(this).X / ActualWidth : 0.5;
                WindowState = WindowState.Normal;
                Left = WinForms.Cursor.Position.X - (RestoreBounds.Width * ratio);
                Top = WinForms.Cursor.Position.Y - 18;
            }

            DragMove();
        }

        private void move(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
                DragMove();
        }

        private void LoadLibraryBrowser()
        {
            libraryRootPath = ResolveLibraryRootPath();
            currentLibraryPath = libraryRootPath;
            expandedLibraryFolders.Clear();
            LoadLibraryFolder(currentLibraryPath);
        }

        private string ResolveLibraryRootPath()
        {
            string start = AppContext.BaseDirectory;
            DirectoryInfo? dir = new DirectoryInfo(start);

            while (dir != null)
            {
                string candidate = Path.Combine(dir.FullName, "Library");
                if (Directory.Exists(candidate))
                    return candidate;

                if (File.Exists(Path.Combine(dir.FullName, "Media.csproj")))
                    return candidate;

                dir = dir.Parent;
            }

            return @"C:\xampp\htdocs\Pzpp\Media\Library";
        }

        private void LoadLibraryFolder(string folderPath)
        {
            suppressLibrarySelectionChanged = true;
            LibraryItems.Clear();

            if (!Directory.Exists(folderPath))
            {
                LibraryItems.Add(new LibraryEntry
                {
                    DisplayName = "Nie znaleziono folderu Library",
                    FullPath = folderPath,
                    IsFolder = false,
                    Icon = "⚠",
                    Level = 0
                });
                foreach (var item in LibraryItems)
                    item.SetSearchHighlight(string.Empty);

                LibraryBackButton.Visibility = Visibility.Collapsed;
                suppressLibrarySelectionChanged = false;
                return;
            }

            currentLibraryPath = folderPath;
            LibraryBackButton.Visibility = Visibility.Collapsed;

            string query = GetLibrarySearchQuery();
            if (string.IsNullOrWhiteSpace(query))
            {
                // Normalny widok: foldery są zawsze widoczne, a zawartość pojawia się pod rozwiniętym folderem.
                AddLibraryDirectoryEntries(folderPath, 0, includeCurrentFolderRow: false);
            }
            else
            {
                // Widok wyszukiwania: pokazujemy tylko pasujące foldery/pliki oraz ich ścieżkę nadrzędną.
                // Foldery z wynikami są automatycznie otwarte, a reszta biblioteki jest ukryta.
                AddLibrarySearchEntries(folderPath, 0, includeCurrentFolderRow: false, query);
            }

            foreach (var item in LibraryItems)
                item.SetSearchHighlight(query);

            suppressLibrarySelectionChanged = false;
        }

        private string GetLibrarySearchQuery()
        {
            if (SearchBox == null) return string.Empty;
            string text = SearchBox.Text?.Trim() ?? string.Empty;
            return string.Equals(text, "Szukaj...", StringComparison.OrdinalIgnoreCase) ? string.Empty : text;
        }

        private bool IsLibrarySearchActive() => !string.IsNullOrWhiteSpace(GetLibrarySearchQuery());

        private bool TextMatchesSearch(string? value, string query)
        {
            return !string.IsNullOrWhiteSpace(value) &&
                   value.IndexOf(query, StringComparison.CurrentCultureIgnoreCase) >= 0;
        }

        private void AddLibraryDirectoryEntries(string folderPath, int level, bool includeCurrentFolderRow)
        {
            if (includeCurrentFolderRow)
            {
                bool isExpanded = expandedLibraryFolders.Contains(folderPath);
                LibraryItems.Add(new LibraryEntry
                {
                    DisplayName = Path.GetFileName(folderPath),
                    FullPath = folderPath,
                    IsFolder = true,
                    IsExpanded = isExpanded,
                    Icon = isExpanded ? "📂" : "📁",
                    Level = level,
                    IsSearchMatch = false
                });

                if (!isExpanded)
                    return;

                level++;
            }

            foreach (string dir in Directory.EnumerateDirectories(folderPath).OrderBy(Path.GetFileName))
                AddLibraryDirectoryEntries(dir, level, includeCurrentFolderRow: true);

            foreach (string file in Directory.EnumerateFiles(folderPath).Where(IsSupportedAudioFile).OrderBy(Path.GetFileName))
                AddLibraryFileEntry(file, level);
        }

        private bool AddLibrarySearchEntries(string folderPath, int level, bool includeCurrentFolderRow, string query)
        {
            var buffer = new List<LibraryEntry>();
            bool found = AddLibrarySearchEntriesToBuffer(folderPath, level, query, buffer, includeRoot: includeCurrentFolderRow);
            if (!includeCurrentFolderRow)
            {
                foreach (var item in buffer)
                    LibraryItems.Add(item);
            }
            return found;
        }

        private bool AddLibrarySearchEntriesToBuffer(string folderPath, int level, string query, List<LibraryEntry> buffer, bool includeRoot = true)
        {
            bool folderNameMatches = TextMatchesSearch(Path.GetFileName(folderPath), query);
            var children = new List<LibraryEntry>();

            if (folderNameMatches)
            {
                AddFullDirectoryTreeToBuffer(folderPath, level + (includeRoot ? 1 : 0), children);
            }
            else
            {
                foreach (string dir in Directory.EnumerateDirectories(folderPath).OrderBy(Path.GetFileName))
                    AddLibrarySearchEntriesToBuffer(dir, level + (includeRoot ? 1 : 0), query, children, includeRoot: true);

                foreach (string file in Directory.EnumerateFiles(folderPath).Where(IsSupportedAudioFile).OrderBy(Path.GetFileName))
                {
                    if (TextMatchesSearch(Path.GetFileName(file), query))
                        children.Add(CreateLibraryFileEntry(file, level + (includeRoot ? 1 : 0), query));
                }
            }

            if (children.Count == 0 && !folderNameMatches)
                return false;

            if (includeRoot)
            {
                buffer.Add(new LibraryEntry
                {
                    DisplayName = Path.GetFileName(folderPath),
                    FullPath = folderPath,
                    IsFolder = true,
                    IsExpanded = true,
                    Icon = "📂",
                    Level = level,
                    IsSearchMatch = folderNameMatches
                });
            }

            buffer.AddRange(children);
            return true;
        }

        private void AddFullDirectoryTreeToBuffer(string folderPath, int level, List<LibraryEntry> buffer)
        {
            foreach (string dir in Directory.EnumerateDirectories(folderPath).OrderBy(Path.GetFileName))
            {
                buffer.Add(new LibraryEntry
                {
                    DisplayName = Path.GetFileName(dir),
                    FullPath = dir,
                    IsFolder = true,
                    IsExpanded = true,
                    Icon = "📂",
                    Level = level
                });
                AddFullDirectoryTreeToBuffer(dir, level + 1, buffer);
            }

            foreach (string file in Directory.EnumerateFiles(folderPath).Where(IsSupportedAudioFile).OrderBy(Path.GetFileName))
                buffer.Add(CreateLibraryFileEntry(file, level));
        }

        private void AddLibraryFileEntry(string file, int level) => LibraryItems.Add(CreateLibraryFileEntry(file, level));

        private LibraryEntry CreateLibraryFileEntry(string file, int level, string? searchQuery = null)
        {
            return new LibraryEntry
            {
                DisplayName = Path.GetFileName(file),
                FullPath = file,
                IsFolder = false,
                Icon = "♪",
                Level = level,
                IsSearchMatch = !string.IsNullOrWhiteSpace(searchQuery) && TextMatchesSearch(Path.GetFileName(file), searchQuery)
            };
        }

        private void RefreshLibrary_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(currentLibraryPath))
                LoadLibraryBrowser();
            else
                LoadLibraryFolder(currentLibraryPath);
        }

        private void LibraryBackButton_Click(object sender, RoutedEventArgs e)
        {
            // Przycisk zostawiony tylko po to, żeby XAML dalej się kompilował.
            LoadLibraryBrowser();
        }

        private void LibraryList_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // Nie przechwytuj myszy, gdy użytkownik klika/ciągnie pasek przewijania.
            // Dzięki temu scrollbar w Library działa 1:1 tak jak w kolejce.
            if (IsFromScrollBar(e.OriginalSource))
                return;

            libraryDragStartPoint = e.GetPosition(LibraryList);

            var itemContainer = FindAncestor<ListBoxItem>((DependencyObject)e.OriginalSource);
            if (itemContainer?.DataContext is not LibraryEntry entry)
                return;

            // Foldery można teraz zaznaczać tak samo jak utwory: zwykły klik, Ctrl+klik i Shift+klik.
            // Rozwijanie / zwijanie folderu odbywa się przez podwójne kliknięcie.
            if (entry.IsFolder && e.ClickCount == 2)
            {
                ToggleLibraryFolder(entry.FullPath);
                e.Handled = true;
                return;
            }

            // Gdy kilka elementów jest już zaznaczonych, kliknięcie jednego z nich w celu przeciągnięcia
            // nie może kasować całego zaznaczenia. Dotyczy to zarówno utworów, jak i folderów.
            // Ctrl/Shift zostawiamy systemowi, żeby działało jak w Windows.
            if (itemContainer.IsSelected && Keyboard.Modifiers == ModifierKeys.None)
                e.Handled = true;
        }

        private void LibraryList_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            // Jeżeli ruch pochodzi z paska przewijania, nie uruchamiaj Drag&Drop.
            if (IsFromScrollBar(e.OriginalSource))
                return;

            if (e.LeftButton != MouseButtonState.Pressed)
                return;

            Point currentPoint = e.GetPosition(LibraryList);
            if (Math.Abs(currentPoint.X - libraryDragStartPoint.X) < SystemParameters.MinimumHorizontalDragDistance &&
                Math.Abs(currentPoint.Y - libraryDragStartPoint.Y) < SystemParameters.MinimumVerticalDragDistance)
                return;

            var selectedPaths = LibraryList.SelectedItems
                .OfType<LibraryEntry>()
                .Where(x =>
                    (!x.IsFolder && File.Exists(x.FullPath) && IsSupportedAudioFile(x.FullPath)) ||
                    (x.IsFolder && Directory.Exists(x.FullPath)))
                .Select(x => x.FullPath)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();

            if (selectedPaths.Length == 0)
                return;

            var data = new DataObject();
            data.SetData(LibraryAudioDragFormat, selectedPaths);
            System.Windows.DragDrop.DoDragDrop(LibraryList, data, System.Windows.DragDropEffects.Copy);
        }

        private void LibraryList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Zaznaczanie w Library działa jak w Windows: zwykły klik, Ctrl+klik oraz Shift+klik.
            // Do kolejki nic nie trafia automatycznie — dodawanie odbywa się tylko przez przeciągnięcie zaznaczonych utworów.
            if (suppressLibrarySelectionChanged) return;
        }

        private void ToggleLibraryFolder(string folderPath)
        {
            if (string.IsNullOrWhiteSpace(folderPath)) return;

            // Podczas wyszukiwania foldery są automatycznie rozwinięte, więc podwójny klik nie powinien ich ukrywać.
            if (IsLibrarySearchActive())
                return;

            var selectedPaths = LibraryList.SelectedItems
                .OfType<LibraryEntry>()
                .Select(x => x.FullPath)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            if (expandedLibraryFolders.Contains(folderPath))
                expandedLibraryFolders.Remove(folderPath);
            else
                expandedLibraryFolders.Add(folderPath);

            LoadLibraryFolder(currentLibraryPath);

            // Po przeładowaniu listy zachowujemy zaznaczenia folderów i plików, które dalej są widoczne.
            if (selectedPaths.Count > 0)
                SelectLibraryEntriesByPaths(selectedPaths);
        }

        private void SelectLibraryFilesByPaths(IEnumerable<string> filePaths) => SelectLibraryEntriesByPaths(filePaths);

        private void SelectLibraryEntriesByPaths(IEnumerable<string> paths)
        {
            var set = paths.ToHashSet(StringComparer.OrdinalIgnoreCase);
            if (set.Count == 0) return;

            suppressLibrarySelectionChanged = true;
            LibraryList.SelectedItems.Clear();

            foreach (var item in LibraryList.Items)
            {
                if (item is LibraryEntry entry && set.Contains(entry.FullPath))
                    LibraryList.SelectedItems.Add(entry);
            }

            suppressLibrarySelectionChanged = false;
        }

        private static bool IsFromScrollBar(object originalSource)
        {
            if (originalSource is not DependencyObject dependencyObject)
                return false;

            return FindAncestor<ScrollBar>(dependencyObject) != null ||
                   FindAncestor<Thumb>(dependencyObject) != null;
        }

        private static T? FindAncestor<T>(DependencyObject? current) where T : DependencyObject
        {
            while (current != null)
            {
                if (current is T match)
                    return match;

                current = GetSafeParent(current);
            }

            return null;
        }

        private static DependencyObject? GetSafeParent(DependencyObject current)
        {
            // e.OriginalSource może być np. Run z TextBlocka. Run nie jest Visual ani Visual3D,
            // więc VisualTreeHelper.GetParent(current) rzuca InvalidOperationException.
            // Dlatego najpierw obsługujemy elementy tekstowe, a VisualTreeHelper wywołujemy
            // tylko dla prawdziwych elementów drzewa wizualnego.
            if (current is FrameworkContentElement frameworkContentElement)
                return frameworkContentElement.Parent;

            if (current is FrameworkElement frameworkElement && frameworkElement.Parent != null)
                return frameworkElement.Parent;

            if (current is System.Windows.Media.Visual || current is System.Windows.Media.Media3D.Visual3D)
                return System.Windows.Media.VisualTreeHelper.GetParent(current);

            return null;
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
            e.Effects = e.Data.GetDataPresent(LibraryAudioDragFormat)
                ? System.Windows.DragDropEffects.Copy
                : System.Windows.DragDropEffects.None;

            e.Handled = true;
        }

        private void DragDrop(object sender, System.Windows.DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(LibraryAudioDragFormat))
                return;

            string[] droppedFiles = (string[])e.Data.GetData(LibraryAudioDragFormat);
            if (droppedFiles != null && droppedFiles.Length > 0)
                AddFiles(droppedFiles);

            e.Handled = true;
        }

        private async void AddFiles(string[] newPaths)
        {
            var candidatePaths = ExpandAudioPaths(newPaths);
            var existingPaths = new HashSet<string>(QueueItems.Select(x => Path.GetFullPath(x.FilePath)), StringComparer.OrdinalIgnoreCase);
            var uniquePaths = new List<string>();

            foreach (string path in candidatePaths)
            {
                string fullPath = Path.GetFullPath(path);
                if (existingPaths.Add(fullPath))
                    uniquePaths.Add(fullPath);
            }

            if (uniquePaths.Count == 0)
                return;

            bool shouldAutoPlay = listBox1.SelectedItem == null && string.IsNullOrWhiteSpace(axWindowsMediaPlayer1.URL);
            string firstPath = uniquePaths[0];

            foreach (string file in uniquePaths)
                QueueItems.Add(CreateFallbackSong(file));

            if (shouldAutoPlay)
                await LoadMetadataAsync(firstPath);

            ApplyCurrentSortPreservingSelection();

            var addedFolders = uniquePaths
                .Select(x => Path.GetDirectoryName(x) ?? string.Empty)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            var pathsFromAddedFolders = QueueItems
                .Where(x => addedFolders.Contains(x.FolderPath))
                .Select(x => x.FilePath)
                .ToList();

            SelectSongsByPaths(pathsFromAddedFolders, firstPath);

            if (shouldAutoPlay)
                PlaySelectedSong();

            foreach (string file in uniquePaths)
                _ = LoadMetadataAsync(file);
        }

        private List<string> ExpandAudioPaths(IEnumerable<string> inputPaths)
        {
            var result = new List<string>();

            foreach (string path in inputPaths)
            {
                if (File.Exists(path) && IsSupportedAudioFile(path))
                {
                    result.Add(path);
                }
                else if (Directory.Exists(path))
                {
                    // Po przeciągnięciu folderu dodajemy wszystkie obsługiwane pliki audio z tego folderu
                    // oraz z jego podfolderów. W kolejce dalej będą pogrupowane według faktycznego folderu pliku.
                    result.AddRange(Directory.EnumerateFiles(path, "*.*", SearchOption.AllDirectories)
                        .Where(IsSupportedAudioFile));
                }
            }

            return result;
        }

        private bool IsSupportedAudioFile(string path)
        {
            string extension = Path.GetExtension(path);
            return SupportedAudioExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase);
        }

        private void list(object sender, SelectionChangedEventArgs e)
        {
            if (suppressSelectionChanged) return;

            if (listBox1.SelectedItem is SongMetadata)
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
                    QueueItems[index] = metadata;

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
            if (axWindowsMediaPlayer1.currentMedia == null && listBox1.SelectedItem != null)
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

        private void QueueGroup_Expanded(object sender, RoutedEventArgs e)
        {
            if (sender is not Expander expander || expander.DataContext is not CollectionViewGroup group)
                return;

            string folderDisplay = group.Name?.ToString() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(folderDisplay))
                return;

            var folderSongs = QueueItems
                .Where(x => string.Equals(x.FolderDisplay, folderDisplay, StringComparison.OrdinalIgnoreCase))
                .Select(x => x.FilePath)
                .ToList();

            if (folderSongs.Count > 0)
                SelectSongsByPaths(folderSongs, folderSongs[0]);
        }

        private void RemoveButton_Click(object sender, RoutedEventArgs e)
        {
            RemoveSelectedSongs();
        }

        private void listBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete)
            {
                RemoveSelectedSongs();
                e.Handled = true;
            }
        }

        private void RemoveSelectedSongs()
        {
            var selectedSongs = listBox1.SelectedItems.Cast<SongMetadata>().ToList();
            if (selectedSongs.Count == 0 && listBox1.SelectedItem is SongMetadata selected)
                selectedSongs.Add(selected);

            if (selectedSongs.Count == 0)
                return;

            int selectedViewIndex = listBox1.SelectedIndex;
            string currentPath = axWindowsMediaPlayer1.URL ?? string.Empty;
            bool removingCurrent = selectedSongs.Any(x => string.Equals(x.FilePath, currentPath, StringComparison.OrdinalIgnoreCase));

            suppressSelectionChanged = true;
            foreach (SongMetadata song in selectedSongs)
            {
                QueueItems.Remove(song);
                metadataCache.Remove(song.FilePath);
            }
            suppressSelectionChanged = false;

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

            int nextIndex = Math.Max(0, Math.Min(selectedViewIndex, listBox1.Items.Count - 1));
            listBox1.SelectedIndex = nextIndex;
            listBox1.ScrollIntoView(listBox1.SelectedItem);

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
            var selectedPaths = listBox1.SelectedItems.Cast<SongMetadata>()
                .Select(x => x.FilePath)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            string? currentSelectedPath = (listBox1.SelectedItem as SongMetadata)?.FilePath;

            var view = CollectionViewSource.GetDefaultView(QueueItems);
            if (view == null) return;

            suppressSelectionChanged = true;
            view.SortDescriptions.Clear();
            view.GroupDescriptions.Clear();
            view.GroupDescriptions.Add(new PropertyGroupDescription(nameof(SongMetadata.FolderDisplay)));
            view.SortDescriptions.Add(new SortDescription(nameof(SongMetadata.FolderDisplay), ListSortDirection.Ascending));
            view.SortDescriptions.Add(new SortDescription(SortPropertyName(sortColumn), sortAscending ? ListSortDirection.Ascending : ListSortDirection.Descending));
            view.Refresh();
            suppressSelectionChanged = false;

            if (selectedPaths.Count > 0)
                SelectSongsByPaths(selectedPaths, currentSelectedPath);
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
                "Album" => nameof(SongMetadata.Album),
                "Artist" => nameof(SongMetadata.Artist),
                _ => nameof(SongMetadata.Title),
            };
        }

        private void SelectSongsByPaths(IEnumerable<string> filePaths, string? primaryPath = null)
        {
            var set = filePaths.ToHashSet(StringComparer.OrdinalIgnoreCase);
            if (set.Count == 0) return;

            suppressSelectionChanged = true;
            listBox1.SelectedItems.Clear();

            SongMetadata? primary = null;
            var songsToSelect = new List<SongMetadata>();

            foreach (var item in listBox1.Items)
            {
                if (item is SongMetadata song && set.Contains(song.FilePath))
                {
                    songsToSelect.Add(song);
                    if (primary == null || string.Equals(song.FilePath, primaryPath, StringComparison.OrdinalIgnoreCase))
                        primary = song;
                }
            }

            if (primary != null)
                listBox1.SelectedItem = primary;

            foreach (SongMetadata song in songsToSelect)
            {
                if (!listBox1.SelectedItems.Contains(song))
                    listBox1.SelectedItems.Add(song);
            }

            if (primary != null)
                listBox1.ScrollIntoView(primary);

            suppressSelectionChanged = false;
        }

        private string GetFolderDisplay(string filePath)
        {
            string folder = Path.GetDirectoryName(filePath) ?? string.Empty;
            string folderName = Path.GetFileName(folder);
            return string.IsNullOrWhiteSpace(folderName) ? folder : folderName;
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (suppressSearchTextChanged || string.IsNullOrWhiteSpace(currentLibraryPath)) return;

            // Debounce: nie przebudowuj całej biblioteki po każdym znaku od razu,
            // tylko poczekaj 1 sekundę od ostatniego znaku, aż użytkownik przestanie pisać.
            librarySearchTimer.Stop();
            librarySearchTimer.Start();
        }

        private void LibrarySearchTimer_Tick(object? sender, EventArgs e)
        {
            librarySearchTimer.Stop();
            ApplyLibrarySearch();
        }

        private void ApplyLibrarySearch()
        {
            if (string.IsNullOrWhiteSpace(currentLibraryPath)) return;

            var selectedPaths = LibraryList?.SelectedItems
                .OfType<LibraryEntry>()
                .Select(x => x.FullPath)
                .ToHashSet(StringComparer.OrdinalIgnoreCase) ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            LoadLibraryFolder(currentLibraryPath);

            if (selectedPaths.Count > 0)
                SelectLibraryEntriesByPaths(selectedPaths);
        }

        private void SearchBox_GotFocus(object sender, RoutedEventArgs e)
        {
            if (SearchBox.Text == "Szukaj...")
            {
                suppressSearchTextChanged = true;
                SearchBox.Text = string.Empty;
                suppressSearchTextChanged = false;
            }
        }

        private void SearchBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(SearchBox.Text))
            {
                suppressSearchTextChanged = true;
                SearchBox.Text = "Szukaj...";
                suppressSearchTextChanged = false;
                librarySearchTimer.Stop();
                LoadLibraryFolder(currentLibraryPath);
            }
        }

        private void Category_Click(object sender, RoutedEventArgs e) { }

        public class LibraryEntry
        {
            public string DisplayName { get; set; } = string.Empty;
            public string FullPath { get; set; } = string.Empty;
            public bool IsFolder { get; set; }
            public bool IsExpanded { get; set; }
            public string Icon { get; set; } = string.Empty;
            public int Level { get; set; }
            public bool IsSearchMatch { get; set; }
            public string DisplayPrefix { get; set; } = string.Empty;
            public string DisplayMatch { get; set; } = string.Empty;
            public string DisplaySuffix { get; set; } = string.Empty;
            public Thickness IndentMargin => new Thickness(Level * 16, 0, 0, 0);

            public void SetSearchHighlight(string? query)
            {
                DisplayPrefix = DisplayName;
                DisplayMatch = string.Empty;
                DisplaySuffix = string.Empty;

                if (string.IsNullOrWhiteSpace(query))
                    return;

                int index = DisplayName.IndexOf(query.Trim(), StringComparison.CurrentCultureIgnoreCase);
                if (index < 0)
                    return;

                DisplayPrefix = DisplayName[..index];
                DisplayMatch = DisplayName.Substring(index, query.Trim().Length);
                DisplaySuffix = DisplayName[(index + query.Trim().Length)..];
            }
        }

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
