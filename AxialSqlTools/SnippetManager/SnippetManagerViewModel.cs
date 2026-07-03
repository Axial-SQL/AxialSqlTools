using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;

namespace AxialSqlTools
{
    public class SnippetManagerViewModel : INotifyPropertyChanged
    {
        private ObservableCollection<SnippetItem> _snippets;
        private SnippetItem _selectedSnippet;
        private string _editPrefix;
        private string _editDescription;
        private string _editBody;
        private bool _isEditing;
        private bool _isEditMode;

        private bool _useSnippets;
        private SettingsManager.SnippetReplaceKey _replaceKey;
        private string _cursorMarker;

        public SnippetManagerViewModel()
        {
            _snippets = new ObservableCollection<SnippetItem>();

            NewCommand = new RelayCommand(OnNew);
            SaveCommand = new RelayCommand(OnSave);
            EditCommand = new RelayCommand(OnEdit, () => SelectedSnippet != null && !IsEditMode);
            DeleteCommand = new RelayCommand(OnDelete);
            DuplicateCommand = new RelayCommand(OnDuplicate);
            ImportLegacyCommand = new RelayCommand(OnImportLegacy);
            CancelCommand = new RelayCommand(OnCancel);
            SaveSettingsCommand = new RelayCommand(OnSaveSettings);

            LoadSettings();
            LoadSnippets();
        }

        public ObservableCollection<SnippetItem> Snippets
        {
            get => _snippets;
            set { _snippets = value; OnPropertyChanged(); }
        }

        public SnippetItem SelectedSnippet
        {
            get => _selectedSnippet;
            set
            {
                _selectedSnippet = value;
                OnPropertyChanged();
                CommandManager.InvalidateRequerySuggested();
                if (_selectedSnippet != null && !_isEditMode)
                {
                    EditPrefix = _selectedSnippet.Prefix;
                    EditDescription = _selectedSnippet.Description;
                    EditBody = _selectedSnippet.Body;
                    IsEditing = true;
                    IsEditMode = false;
                }
            }
        }

        public string EditPrefix
        {
            get => _editPrefix;
            set { _editPrefix = value; OnPropertyChanged(); }
        }

        public string EditDescription
        {
            get => _editDescription;
            set { _editDescription = value; OnPropertyChanged(); }
        }

        public string EditBody
        {
            get => _editBody;
            set { _editBody = value; OnPropertyChanged(); }
        }

        public bool IsEditing
        {
            get => _isEditing;
            set { _isEditing = value; OnPropertyChanged(); }
        }

        public bool IsEditMode
        {
            get => _isEditMode;
            set
            {
                _isEditMode = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsEditorReadOnly));
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public bool IsEditorReadOnly => !IsEditMode;

        public bool UseSnippets
        {
            get => _useSnippets;
            set { _useSnippets = value; OnPropertyChanged(); }
        }

        public SettingsManager.SnippetReplaceKey ReplaceKey
        {
            get => _replaceKey;
            set { _replaceKey = value; OnPropertyChanged(); }
        }

        public string CursorMarker
        {
            get => _cursorMarker;
            set { _cursorMarker = value; OnPropertyChanged(); }
        }

        public ICommand NewCommand { get; }
        public ICommand SaveCommand { get; }
        public ICommand EditCommand { get; }
        public ICommand DeleteCommand { get; }
        public ICommand DuplicateCommand { get; }
        public ICommand ImportLegacyCommand { get; }
        public ICommand CancelCommand { get; }
        public ICommand SaveSettingsCommand { get; }

        private void LoadSnippets()
        {
            SnippetService.ReloadSnippets();
            var items = SnippetService.GetAllSnippets();
            _snippets.Clear();
            foreach (var item in items)
            {
                _snippets.Add(item);
            }
        }

        private void LoadSettings()
        {
            var settings = SettingsManager.GetSnippetSettings();
            _useSnippets = settings.useSnippets;
            _replaceKey = settings.replaceKey;
            _cursorMarker = settings.cursorMarker;
        }

        private void OnNew()
        {
            _selectedSnippet = null;
            EditPrefix = string.Empty;
            EditDescription = string.Empty;
            EditBody = string.Empty;
            IsEditing = true;
            IsEditMode = true;
            OnPropertyChanged(nameof(SelectedSnippet));
        }

        private void OnEdit()
        {
            if (_selectedSnippet != null)
                IsEditMode = true;
        }

        private void OnSave()
        {
            if (string.IsNullOrWhiteSpace(EditPrefix))
            {
                MessageBox.Show("Prefix is required.", "Snippet Manager", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string savedId;
            if (_selectedSnippet == null)
            {
                var newSnippet = new SnippetItem(EditPrefix.Trim(), EditDescription ?? string.Empty, EditBody ?? string.Empty);
                SnippetService.AddSnippet(newSnippet);
                savedId = newSnippet.Id;
            }
            else
            {
                _selectedSnippet.Prefix = EditPrefix.Trim();
                _selectedSnippet.Description = EditDescription ?? string.Empty;
                _selectedSnippet.Body = EditBody ?? string.Empty;
                SnippetService.UpdateSnippet(_selectedSnippet);
                savedId = _selectedSnippet.Id;
            }

            LoadSnippets();
            SelectedSnippet = _snippets.FirstOrDefault(s => s.Id == savedId);
            IsEditMode = false;
            IsEditing = SelectedSnippet != null;
        }

        private void OnDelete()
        {
            if (_selectedSnippet == null)
                return;

            var result = MessageBox.Show(
                $"Delete snippet '{_selectedSnippet.Prefix}'?",
                "Confirm Delete",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                SnippetService.DeleteSnippet(_selectedSnippet.Id);
                LoadSnippets();
                IsEditing = false;
                _selectedSnippet = null;
                IsEditMode = false;
                OnPropertyChanged(nameof(SelectedSnippet));
            }
        }

        private void OnDuplicate()
        {
            if (_selectedSnippet == null)
                return;

            var dup = new SnippetItem(
                _selectedSnippet.Prefix + "_copy",
                _selectedSnippet.Description,
                _selectedSnippet.Body);

            SnippetService.AddSnippet(dup);
            LoadSnippets();
            SelectedSnippet = _snippets.FirstOrDefault(s => s.Id == dup.Id);
        }

        private void OnImportLegacy()
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                dialog.Description = "Select legacy snippets folder (.sql files)";
                dialog.ShowNewFolderButton = false;

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    SnippetService.ImportFromLegacyFolder(dialog.SelectedPath);
                    LoadSnippets();
                    MessageBox.Show("Import completed.", "Snippet Manager", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
        }

        private void OnCancel()
        {
            IsEditMode = false;

            if (_selectedSnippet != null)
            {
                EditPrefix = _selectedSnippet.Prefix;
                EditDescription = _selectedSnippet.Description;
                EditBody = _selectedSnippet.Body;
            }
            else
            {
                IsEditing = false;
                EditPrefix = string.Empty;
                EditDescription = string.Empty;
                EditBody = string.Empty;
            }
        }

        private void OnSaveSettings()
        {
            var settings = SettingsManager.GetSnippetSettings();
            settings.useSnippets = UseSnippets;
            settings.replaceKey = ReplaceKey;
            settings.cursorMarker = CursorMarker;
            SettingsManager.SaveSnippetSettings(settings);
            MessageBox.Show("Settings saved.", "Snippet Manager", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
