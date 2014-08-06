using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace HOK.DivaAutomation.Exporter
{
    public enum LightingTest
    {
        Illuminance=0,
        LEED=1,
        NECHPS=2,
        MACHPS=3,
        DaylightFactor=4,
        RadiationMap=5,
        DaylightAutonomy=6,
        ContinuousDaylightAutonomy=7,
        DaylightAvailability=8,
        UDI=9,
        IlluminanceFromDA=10,
    }

    public enum AnalysisBase
    {
       None=0,
       Rooms=1,
       Areas=2
    }

    /// <summary>
    /// Interaction logic for DwgExporterWindow.xaml
    /// </summary>
    public partial class DwgExporterWindow : Window
    {
        private UIApplication m_app;
        private Document m_doc;

        private LightingTest selectedLightingTest = LightingTest.Illuminance;
        private AnalysisBase selectedAnalysisBase = AnalysisBase.None;
        private ElementId selectedViewId = ElementId.InvalidElementId;

        public LightingTest SelectedLightingTest { get { return selectedLightingTest; } set { selectedLightingTest = value; } }
        public AnalysisBase SelectedAnalysisBase { get { return selectedAnalysisBase; } set { selectedAnalysisBase = value; } }
        public ElementId SelectedViewId { get { return selectedViewId; } set { selectedViewId = value; } }

        public DwgExporterWindow(UIApplication uiapp)
        {
            try
            {
                m_app = uiapp;
                m_doc = m_app.ActiveUIDocument.Document;

                InitializeComponent();
                AddComboBoxTests();
                cmbBoxTest.SelectedIndex = 0;
                AddComboBoxViews();

            }
            catch (Exception ex)
            {
                string message = ex.Message;
            }
        }

        private void AddComboBoxTests()
        {
            try
            {
                foreach (int value in Enum.GetValues(typeof(LightingTest)))
                {
                    ComboBoxItem item = new ComboBoxItem();
                    switch (value)
                    {
                        case 0:
                            item.Content = "Illuminance Values";
                            item.Tag = (LightingTest)value;
                            break;
                        case 1:
                            item.Content = "LEED IEQ 8.1";
                            item.Tag = (LightingTest)value;
                            break;
                        case 2:
                            item.Content = "NECHPS IEQ P2";
                            item.Tag = (LightingTest)value;
                            break;
                        case 3:
                            item.Content = "MACHPS IEQ C2";
                            item.Tag = (LightingTest)value;
                            break;
                        case 4:
                            item.Content = "Daylight Factor";
                            item.Tag = (LightingTest)value;
                            break;
                        case 5:
                            item.Content = "Radiation Map";
                            item.Tag = (LightingTest)value;
                            break;
                        case 6:
                            item.Content = "Daylight Autonomy";
                            item.Tag = (LightingTest)value;
                            break;
                        case 7:
                            item.Content = "Continuous Daylight Autonomy";
                            item.Tag = (LightingTest)value;
                            break;
                        case 8:
                            item.Content = "Daylight Availability";
                            item.Tag = (LightingTest)value;
                            break;
                        case 9:
                            item.Content = "Useful Daylight Illuminance (UDI)";
                            item.Tag = (LightingTest)value;
                            break;
                        case 10:
                            item.Content = "Illuminance from DA";
                            item.Tag = (LightingTest)value;
                            break;
                    }
                    cmbBoxTest.Items.Add(item);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to add combo box items for lighting tests.\n"+ex.Message, "Add Combo Box Items", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void AddComboBoxViews()
        {
            try
            {
                FilteredElementCollector collector = new FilteredElementCollector(m_doc);
                List<Element> views = collector.OfClass(typeof(View3D)).WhereElementIsNotElementType().ToElements().ToList();
                views = views.OrderBy(o => o.Name).ToList();

                int index = 0;
                for (int i = 0; i < views.Count; i++)
                {
                    Element view = views[i];
                    ComboBoxItem item = new ComboBoxItem();
                    item.Content = view.Name;
                    item.Tag = view.Id;
                    cmbBoxView.Items.Add(item);
                    if (view.Name == "{3D}")
                    {
                        index = i;
                    }
                }

                cmbBoxView.SelectedIndex = index;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to add combobox items for 3D Views.\n"+ex.Message, "Add Combo Box Items", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        
        private void bttnRooms_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                selectedAnalysisBase = AnalysisBase.Rooms;
                ComboBoxItem selectedItem = cmbBoxTest.SelectedItem as ComboBoxItem;
                selectedLightingTest = (LightingTest)selectedItem.Tag;

                ComboBoxItem selectedView = cmbBoxView.SelectedItem as ComboBoxItem;
                selectedViewId = (ElementId)selectedView.Tag;

                this.DialogResult = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to select rooms.\n"+ex.Message, "Select Rooms", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void bttnAreas_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                selectedAnalysisBase = AnalysisBase.Rooms;
                ComboBoxItem selectedItem = cmbBoxTest.SelectedItem as ComboBoxItem;
                selectedLightingTest = (LightingTest)selectedItem.Tag;

                ComboBoxItem selectedView = cmbBoxView.SelectedItem as ComboBoxItem;
                selectedViewId = (ElementId)selectedView.Tag;

                this.DialogResult = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to select areas.\n"+ex.Message, "Select Areas", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void bttnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
