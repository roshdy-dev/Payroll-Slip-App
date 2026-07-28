using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.VisualTree;
using PayrollSlipApp.ViewModels;

namespace PayrollSlipApp.Views;

public partial class MainWindow : Window
{
    private MainViewModel? _vm;

    public MainWindow()
    {
        InitializeComponent();
        _vm = new MainViewModel();
        DataContext = _vm;

        // Wire up drag-and-drop on the drop zone
        var dropZone = this.FindControl<Border>("DropZone");
        if (dropZone != null)
        {
            AddHandler(DragDrop.DropEvent, OnDrop);
            AddHandler(DragDrop.DragOverEvent, OnDragOver);
        }
    }

    private void OnDragOver(object? sender, DragEventArgs e)
    {
        // Only accept file drops
        if (e.Data.Contains(DataFormats.Files))
        {
            e.DragEffects = DragDropEffects.Copy;
            e.Handled = true;
        }
    }

    private void OnDrop(object? sender, DragEventArgs e)
    {
        if (e.Data.Contains(DataFormats.Files))
        {
            var files = e.Data.GetFiles();
            if (files != null)
            {
                foreach (var file in files)
                {
                    var path = file.Path.LocalPath;
                    _vm?.OnFileDropped(path);
                    break; // Only take the first file
                }
            }
        }
    }
}
