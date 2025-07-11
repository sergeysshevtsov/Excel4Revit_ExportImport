using Autodesk.Revit.UI;
using Excel4Revit_ExportImport.Views.ViewExport.Models;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Input;

namespace Excel4Revit_ExportImport.Views.ViewExport;
public class WindowExportDataContext : INotifyPropertyChanged
{
    private readonly List<Element> elements;

    public event PropertyChangedEventHandler PropertyChanged;
    protected void OnPropertyChanged(string propName) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propName));

    public WindowExportDataContext(List<Element> elements, List<ParameterData> parametersData)
    {
        Items = new ObservableCollection<ParameterData>(parametersData);
        this.elements = elements;
    }

    public ICommand CloseCommand => new RelayCommand<Window>(
        window =>
        {
            window?.Close();
        });

    public ICommand ExportCommand => new RelayCommand<Window>(
        window =>
        {
            IsExported = true;
            var selectedParameters = Items.AsEnumerable().Where(x => x.IsChecked).ToList();
            if (selectedParameters.Count == 0)
            {
                TaskDialog.Show("Excel4Revit_ExportImport", "No parameters selected for export.");
                return;
            }

            ExcelUtils.ExcelFile.Export(elements, selectedParameters);
            window?.Close();
        });

    public ObservableCollection<ParameterData> Items { get; set; }

    private bool _isAllSelected;


    public bool IsAllSelected
    {
        get => _isAllSelected;
        set
        {
            if (_isAllSelected != value)
            {
                _isAllSelected = value;
                OnPropertyChanged(nameof(IsAllSelected));

                foreach (var item in Items)
                    item.IsChecked = value;
            }
        }
    }

    public bool IsExported { get; set; }
}
