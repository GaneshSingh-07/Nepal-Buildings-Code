// Syntax Example for importing namespaces: using System.Runtime.InteropServices;
using ETABSv1;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Media;

// IMPORTANT: Ensure this namespace matches x:Class="NBC105_2025_Calculator.MainWindow" in MainWindow.xaml
namespace NBC
{
    /// <summary>
    /// Represents a City / Municipality entry from NBC 105:2025 Annex C (Table C-1).
    /// </summary>
    public class CityZoneOption
    {
        // District name from Annex C (e.g., "Siraha", "Kathmandu", "Kaski")
        // Syntax Example: public string District { get; set; } = string.Empty;
        public string District { get; set; } = string.Empty;

        // Local unit / Municipality name from Annex C (e.g., "Dhangadhimai Nagarpalika")
        public string Municipality { get; set; } = string.Empty;

        // Peak Ground Acceleration (PGA) seismic zoning factor Z from Annex C
        public double ZValue { get; set; }

        // Formats display text in the searchable dropdown
        public override string ToString() => $"{Municipality} ({District}) — Z = {ZValue:F2}";
    }

    /// <summary>
    /// Represents a Site Subsoil Category from NBC 105:2025 Clause 4.1.2 (Table 4-1) & Clause 4.1.3 (Table 4-2).
    /// </summary>
    public class SoilTypeOption
    {
        // Soil type code ("A", "B", "C", or "D")
        public string Code { get; set; } = string.Empty;

        // Display label for the ComboBox
        public string DisplayName { get; set; } = string.Empty;

        // Lower limit of constant acceleration branch Ta (0.0 for Equivalent Static Method per Table 4-1 Note 1)
        public double Ta { get; set; }

        // Upper limit of constant acceleration branch Tc (sec)
        public double Tc { get; set; }

        // Beginning of constant displacement range Td (sec)
        public double Td { get; set; }

        // Normalized peak spectral acceleration factor Alpha
        public double Alpha { get; set; }

        public override string ToString() => DisplayName;
    }

    /// <summary>
    /// Represents a Building Importance Class from NBC 105:2025 Clause 4.1.5 (Table 4-4).
    /// </summary>
    public class ImportanceOption
    {
        // Description of the Importance Class
        public string DisplayName { get; set; } = string.Empty;

        // Importance Factor (I) value
        public double IValue { get; set; }

        public override string ToString() => DisplayName;
    }

    /// <summary>
    /// Represents an Empirical Period coefficient (kt) category from NBC 105:2025 Clause 5.1.2 (Eq. 5.1(2)).
    /// </summary>
    public class EmpiricalKtOption
    {
        // Description of the structural framing type
        public string DisplayName { get; set; } = string.Empty;

        // Base kt value (0.075, 0.085, or 0.050)
        public double KtBase { get; set; }

        // True if kt = 0.075 / sqrt(Aw) for concrete shear walls
        public bool IsConcreteShearWall { get; set; }

        public override string ToString() => DisplayName;
    }

    /// <summary>
    /// Represents a row in NBC 105:2025 Table 5-2 with interactive EQX and EQY selection checkboxes.
    /// </summary>
    public class StructuralSystemRow : INotifyPropertyChanged
    {
        // Event notifying the DataGrid when EQX or EQY selection changes
        public event PropertyChangedEventHandler? PropertyChanged;

        // Helper method to trigger PropertyChanged
        protected void OnPropertyChanged([CallerMemberName] string? propName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propName));
        }

        // Serial number in Table 5-2 (1 to 21)
        public int SerialNo { get; set; }

        // System category in Table 5-2 (Moment Resisting Frame, Braced Frame, Structural Wall, Dual System)
        public string Category { get; set; } = string.Empty;

        // Structural system name in Table 5-2
        public string SystemName { get; set; } = string.Empty;

        // ULS Ductility Factor (R_mu) from Table 5-2
        public double R_mu { get; set; }

        // ULS Overstrength Factor (Omega_u) from Table 5-2
        public double Omega_u { get; set; }

        // SLS Overstrength Factor (Omega_s) from Table 5-2
        public double Omega_s { get; set; }

        // Callback action invoked when the user checks EQX or EQY on this row
        public Action<StructuralSystemRow, bool>? OnSelectionChanged { get; set; }

        private bool _isSelectedX;
        // Bound to the "EQX" CheckBox column in the Table 5-2 DataGrid
        public bool IsSelectedX
        {
            get => _isSelectedX;
            set
            {
                if (_isSelectedX != value)
                {
                    _isSelectedX = value;
                    OnPropertyChanged();
                    // If checked, notify parent ViewModel to make this row the active X-direction system
                    if (value) OnSelectionChanged?.Invoke(this, true);
                }
            }
        }

        private bool _isSelectedY;
        // Bound to the "EQY" CheckBox column in the Table 5-2 DataGrid
        public bool IsSelectedY
        {
            get => _isSelectedY;
            set
            {
                if (_isSelectedY != value)
                {
                    _isSelectedY = value;
                    OnPropertyChanged();
                    // If checked, notify parent ViewModel to make this row the active Y-direction system
                    if (value) OnSelectionChanged?.Invoke(this, false);
                }
            }
        }

        // Internal method allowing the ViewModel to uncheck other rows without triggering recursive callbacks
        public void SetSelectedSilent(bool isX, bool state)
        {
            if (isX)
            {
                _isSelectedX = state;
                OnPropertyChanged(nameof(IsSelectedX));
            }
            else
            {
                _isSelectedY = state;
                OnPropertyChanged(nameof(IsSelectedY));
            }
        }
    }

    /// <summary>
    /// Native Win32 COM helper to connect to a running ETABS instance in both .NET Framework and .NET 6/7/8+.
    /// </summary>
    public static class EtabsComHelper
    {
        // Imports CLSIDFromProgID from ole32.dll to convert "CSI.ETABS.API.ETABSObject" into its COM GUID
        // Syntax Example: [DllImport("ole32.dll")] private static extern int CLSIDFromProgID(...);
        [DllImport("ole32.dll", PreserveSig = false)]
        private static extern void CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string progId, out Guid clsid);

        // Imports GetActiveObject from oleaut32.dll to fetch the running ETABS instance from the Running Object Table
        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void GetActiveObject(ref Guid rclsid, IntPtr reserved, [MarshalAs(UnmanagedType.Interface)] out object ppunk);

        /// <summary>
        /// Retrieves the active COM object for the given ProgID (e.g., "CSI.ETABS.API.ETABSObject").
        /// </summary>
        public static object GetActiveEtabsObject(string progId = "CSI.ETABS.API.ETABSObject")
        {
            CLSIDFromProgID(progId, out Guid clsid);
            GetActiveObject(ref clsid, IntPtr.Zero, out object activeObj);
            return activeObj;
        }
    }

    /// <summary>
    /// Main MVVM ViewModel holding all NBC 105:2025 data tables, user inputs, computed coefficients, and ETABS automation logic.
    /// </summary>
    public class MainViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        // ==================== COLLECTIONS FOR UI CONTROLS ====================
        private readonly List<CityZoneOption> _allCities;
        public ObservableCollection<CityZoneOption> FilteredCities { get; } = new ObservableCollection<CityZoneOption>();
        public List<SoilTypeOption> SoilTypes { get; }
        public List<ImportanceOption> ImportanceClasses { get; }
        public List<string> RegularityOptions { get; }
        public List<EmpiricalKtOption> KtOptionsX { get; }
        public List<EmpiricalKtOption> KtOptionsY { get; }
        public ObservableCollection<StructuralSystemRow> Table5_2Rows { get; } = new ObservableCollection<StructuralSystemRow>();

        // ==================== BACKING FIELDS FOR INPUTS ====================
        private string _citySearchText = string.Empty;
        private CityZoneOption? _selectedCity;
        private string _manualZText = "0.35";
        private SoilTypeOption _selectedSoilType;
        private ImportanceOption _selectedImportance;
        private int _selectedRegularityIndex = 0;
        private string _heightText = "12.0";
        private string _eccentricityText = "0.05";
        private EmpiricalKtOption _selectedKtX;
        private EmpiricalKtOption _selectedKtY;
        private string _awXText = "2.5";
        private string _awYText = "2.5";
        private bool _useRayleigh = false;
        private string _rayleighXText = "0.55";
        private string _rayleighYText = "0.55";
        private StructuralSystemRow _selectedSystemX;
        private StructuralSystemRow _selectedSystemY;

        // ==================== COMPUTED NUMERICAL VALUES FOR ETABS ====================
        // Stores the latest calculated base shear coefficients Cd(T1) and exponents k for ETABS export
        public double LatestCd_EQX_SLS { get; private set; }
        public double LatestCd_EQY_SLS { get; private set; }
        public double LatestCd_EQX_ULS { get; private set; }
        public double LatestCd_EQY_ULS { get; private set; }
        public double LatestKx { get; private set; } = 1.0;
        public double LatestKy { get; private set; } = 1.0;
        public double LatestEccentricity { get; private set; } = 0.05;

        // ==================== BACKING FIELDS FOR OUTPUT DISPLAY ====================
        private string _eqxSlsResult = "0.0000";
        private string _eqySlsResult = "0.0000";
        private string _eqxUlsResult = "0.0000";
        private string _eqyUlsResult = "0.0000";
        private string _selectedSystemXSummary = string.Empty;
        private string _selectedSystemYSummary = string.Empty;
        private string _breakdownText = string.Empty;
        private string _applicabilityMessage = string.Empty;
        private Brush _applicabilityBackground = new SolidColorBrush(Color.FromRgb(220, 252, 231));
        private Brush _applicabilityBorder = new SolidColorBrush(Color.FromRgb(134, 239, 172));
        private Brush _applicabilityForeground = new SolidColorBrush(Color.FromRgb(22, 101, 52));
        private string _etabsStatusMessage = "Ready: Open your model in ETABS and click 'Create ETABS Load'.";
        private Brush _etabsStatusBrush = new SolidColorBrush(Color.FromRgb(71, 85, 105));

        // ==================== BOUND PROPERTIES ====================

        // Search text that filters the Annex C City/Municipality dropdown in real time
        public string CitySearchText
        {
            get => _citySearchText;
            set
            {
                _citySearchText = value;
                OnPropertyChanged();
                UpdateFilteredCities();
            }
        }

        // Selected city from Annex C; automatically updates ManualZText with that city's PGA (Z)
        public CityZoneOption? SelectedCity
        {
            get => _selectedCity;
            set
            {
                _selectedCity = value;
                OnPropertyChanged();
                if (value != null)
                {
                    // Auto-populate the manual Z box with the selected city's Z factor
                    ManualZText = value.ZValue.ToString("F2");
                }
            }
        }

        // Manual Seismic Zoning Factor (Z) text box — user can type any custom Z or let city selection fill it
        public string ManualZText
        {
            get => _manualZText;
            set { _manualZText = value; OnPropertyChanged(); Recalculate(); }
        }

        public SoilTypeOption SelectedSoilType
        {
            get => _selectedSoilType;
            set { _selectedSoilType = value; OnPropertyChanged(); Recalculate(); }
        }

        public ImportanceOption SelectedImportance
        {
            get => _selectedImportance;
            set { _selectedImportance = value; OnPropertyChanged(); Recalculate(); }
        }

        public int SelectedRegularityIndex
        {
            get => _selectedRegularityIndex;
            set { _selectedRegularityIndex = value; OnPropertyChanged(); Recalculate(); }
        }

        public string HeightText
        {
            get => _heightText;
            set { _heightText = value; OnPropertyChanged(); Recalculate(); }
        }

        // Accidental eccentricity ratio (e.g., 0.05 for ±0.05b per Clause 5.6)
        public string EccentricityText
        {
            get => _eccentricityText;
            set { _eccentricityText = value; OnPropertyChanged(); Recalculate(); }
        }

        public EmpiricalKtOption SelectedKtX
        {
            get => _selectedKtX;
            set
            {
                _selectedKtX = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(AwXVisibility));
                Recalculate();
            }
        }

        public EmpiricalKtOption SelectedKtY
        {
            get => _selectedKtY;
            set
            {
                _selectedKtY = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(AwYVisibility));
                Recalculate();
            }
        }

        public Visibility AwXVisibility => (_selectedKtX != null && _selectedKtX.IsConcreteShearWall) ? Visibility.Visible : Visibility.Collapsed;
        public Visibility AwYVisibility => (_selectedKtY != null && _selectedKtY.IsConcreteShearWall) ? Visibility.Visible : Visibility.Collapsed;

        public string AwXText
        {
            get => _awXText;
            set { _awXText = value; OnPropertyChanged(); Recalculate(); }
        }

        public string AwYText
        {
            get => _awYText;
            set { _awYText = value; OnPropertyChanged(); Recalculate(); }
        }

        public bool UseRayleigh
        {
            get => _useRayleigh;
            set { _useRayleigh = value; OnPropertyChanged(); Recalculate(); }
        }

        public string RayleighXText
        {
            get => _rayleighXText;
            set { _rayleighXText = value; OnPropertyChanged(); Recalculate(); }
        }

        public string RayleighYText
        {
            get => _rayleighYText;
            set { _rayleighYText = value; OnPropertyChanged(); Recalculate(); }
        }

        public string SelectedSystemXSummary
        {
            get => _selectedSystemXSummary;
            private set { _selectedSystemXSummary = value; OnPropertyChanged(); }
        }

        public string SelectedSystemYSummary
        {
            get => _selectedSystemYSummary;
            private set { _selectedSystemYSummary = value; OnPropertyChanged(); }
        }

        public string EqxSlsResult
        {
            get => _eqxSlsResult;
            private set { _eqxSlsResult = value; OnPropertyChanged(); }
        }

        public string EqySlsResult
        {
            get => _eqySlsResult;
            private set { _eqySlsResult = value; OnPropertyChanged(); }
        }

        public string EqxUlsResult
        {
            get => _eqxUlsResult;
            private set { _eqxUlsResult = value; OnPropertyChanged(); }
        }

        public string EqyUlsResult
        {
            get => _eqyUlsResult;
            private set { _eqyUlsResult = value; OnPropertyChanged(); }
        }

        public string BreakdownText
        {
            get => _breakdownText;
            private set { _breakdownText = value; OnPropertyChanged(); }
        }

        public string ApplicabilityMessage
        {
            get => _applicabilityMessage;
            private set { _applicabilityMessage = value; OnPropertyChanged(); }
        }

        public Brush ApplicabilityBackground
        {
            get => _applicabilityBackground;
            private set { _applicabilityBackground = value; OnPropertyChanged(); }
        }

        public Brush ApplicabilityBorder
        {
            get => _applicabilityBorder;
            private set { _applicabilityBorder = value; OnPropertyChanged(); }
        }

        public Brush ApplicabilityForeground
        {
            get => _applicabilityForeground;
            private set { _applicabilityForeground = value; OnPropertyChanged(); }
        }

        public string EtabsStatusMessage
        {
            get => _etabsStatusMessage;
            set { _etabsStatusMessage = value; OnPropertyChanged(); }
        }

        public Brush EtabsStatusBrush
        {
            get => _etabsStatusBrush;
            set { _etabsStatusBrush = value; OnPropertyChanged(); }
        }

        public MainViewModel()
        {
            // 1. Load searchable municipalities from NBC 105:2025 Annex C (Table C-1)
            _allCities = new List<CityZoneOption>
            {
                // Kathmandu Valley
                new CityZoneOption { District = "Kathmandu", Municipality = "Kathmandu Mahanagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Kathmandu", Municipality = "Budhanilakantha Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Kathmandu", Municipality = "Chandragiri Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Kathmandu", Municipality = "Gokarneshwor Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Kathmandu", Municipality = "Kageshwori Manahora Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Kathmandu", Municipality = "Kirtipur Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Kathmandu", Municipality = "Nagarjun Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Kathmandu", Municipality = "Tarakeshwor Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Kathmandu", Municipality = "Tokha Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Lalitpur", Municipality = "Lalitpur Mahanagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Lalitpur", Municipality = "Godawari Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Lalitpur", Municipality = "Mahalaxmi Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Bhaktapur", Municipality = "Bhaktapur Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Bhaktapur", Municipality = "Madhyapur Thimi Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Bhaktapur", Municipality = "Suryabinayak Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Bhaktapur", Municipality = "Changunarayan Nagarpalika", ZValue = 0.35 },

                // Madhesh Province (Siraha, Dhanusha, Saptari, Mahottari, Sarlahi, Rautahat, Bara, Parsa)
                new CityZoneOption { District = "Siraha", Municipality = "Dhangadhimai Nagarpalika", ZValue = 0.40 },
                new CityZoneOption { District = "Siraha", Municipality = "Lahan Nagarpalika", ZValue = 0.40 },
                new CityZoneOption { District = "Siraha", Municipality = "Golbazar Nagarpalika", ZValue = 0.40 },
                new CityZoneOption { District = "Siraha", Municipality = "Siraha Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Siraha", Municipality = "Mirchaiya Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Siraha", Municipality = "Kalyanpur Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Siraha", Municipality = "Karjanha Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Siraha", Municipality = "Sukhipur Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Dhanusha", Municipality = "Janakpur Upamahanagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Dhanusha", Municipality = "Chhireshwornath Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Dhanusha", Municipality = "Mithila Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Dhanusha", Municipality = "Nagarain Nagarpalika", ZValue = 0.30 },
                new CityZoneOption { District = "Saptari", Municipality = "Rajbiraj Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Saptari", Municipality = "Khadak Nagarpalika", ZValue = 0.40 },
                new CityZoneOption { District = "Saptari", Municipality = "Surunga Nagarpalika", ZValue = 0.40 },
                new CityZoneOption { District = "Saptari", Municipality = "Kanchanrup Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Mahottari", Municipality = "Bardibas Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Mahottari", Municipality = "Gaushala Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Mahottari", Municipality = "Jaleswor Nagarpalika", ZValue = 0.30 },
                new CityZoneOption { District = "Sarlahi", Municipality = "Lalbandi Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Sarlahi", Municipality = "Hariwan Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Sarlahi", Municipality = "Malangawa Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Rautahat", Municipality = "Chandrapur Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Rautahat", Municipality = "Gaur Nagarpalika", ZValue = 0.30 },
                new CityZoneOption { District = "Bara", Municipality = "Jitpur Simara Upamahanagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Bara", Municipality = "Kalaiya Upamahanagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Bara", Municipality = "Nijgadh Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Bara", Municipality = "Simraungadh Nagarpalika", ZValue = 0.30 },
                new CityZoneOption { District = "Parsa", Municipality = "Birgunj Mahanagarpalika", ZValue = 0.35 },

                // Koshi, Bagmati, Gandaki, Lumbini, Karnali, Sudurpashchim Major Cities
                new CityZoneOption { District = "Morang", Municipality = "Biratnagar Mahanagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Morang", Municipality = "Sundarharaicha Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Morang", Municipality = "Rangeli Nagarpalika", ZValue = 0.30 },
                new CityZoneOption { District = "Sunsari", Municipality = "Dharan Upamahanagarpalika", ZValue = 0.40 },
                new CityZoneOption { District = "Sunsari", Municipality = "Itahari Upamahanagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Sunsari", Municipality = "Inaruwa Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Jhapa", Municipality = "Birtamod Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Jhapa", Municipality = "Damak Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Jhapa", Municipality = "Mechinagar Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Jhapa", Municipality = "Bhadrapur Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Jhapa", Municipality = "Gauriganj Gaunpalika", ZValue = 0.30 },
                new CityZoneOption { District = "Illam", Municipality = "Illam Nagarpalika", ZValue = 0.40 },
                new CityZoneOption { District = "Dhankuta", Municipality = "Dhankuta Nagarpalika", ZValue = 0.40 },
                new CityZoneOption { District = "Udayapur", Municipality = "Triyuga Nagarpalika", ZValue = 0.40 },
                new CityZoneOption { District = "Udayapur", Municipality = "Katari Nagarpalika", ZValue = 0.40 },
                new CityZoneOption { District = "Taplejung", Municipality = "Phaktanglung Gaunpalika", ZValue = 0.25 },
                new CityZoneOption { District = "Taplejung", Municipality = "Phungling Nagarpalika", ZValue = 0.30 },
                new CityZoneOption { District = "Chitawan", Municipality = "Bharatpur Mahanagarpalika", ZValue = 0.40 },
                new CityZoneOption { District = "Chitawan", Municipality = "Ratnanagar Nagarpalika", ZValue = 0.40 },
                new CityZoneOption { District = "Makwanpur", Municipality = "Hetauda Upamahanagarpalika", ZValue = 0.40 },
                new CityZoneOption { District = "Sindhuli", Municipality = "Kamalamai Nagarpalika", ZValue = 0.40 },
                new CityZoneOption { District = "Kabhrepalanchok", Municipality = "Dhulikhel Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Kabhrepalanchok", Municipality = "Banepa Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Nuwakot", Municipality = "Bidur Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Dolakha", Municipality = "Bhimeshwor Nagarpalika", ZValue = 0.30 },
                new CityZoneOption { District = "Dolakha", Municipality = "Jiri Nagarpalika", ZValue = 0.30 },
                new CityZoneOption { District = "Kaski", Municipality = "Pokhara Lekhnath Mahanagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Gorkha", Municipality = "Gorkha Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Tanahu", Municipality = "Byas Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Baglung", Municipality = "Baglung Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Myagdi", Municipality = "Beni Nagarpalika", ZValue = 0.30 },
                new CityZoneOption { District = "Lamjung", Municipality = "Besishahar Nagarpalika", ZValue = 0.30 },
                new CityZoneOption { District = "Nawalparasi_E", Municipality = "Kawasoti Nagarpalika", ZValue = 0.40 },
                new CityZoneOption { District = "Nawalparasi_E", Municipality = "Gaidakot Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Rupandehi", Municipality = "Butwal Upamahanagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Rupandehi", Municipality = "Siddharthanagar Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Rupandehi", Municipality = "Tillotama Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Palpa", Municipality = "Tansen Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Dang", Municipality = "Ghorahi Upamahanagarpalika", ZValue = 0.40 },
                new CityZoneOption { District = "Dang", Municipality = "Tulsipur Upamahanagarpalika", ZValue = 0.40 },
                new CityZoneOption { District = "Banke", Municipality = "Nepalgunj Upamahanagarpalika", ZValue = 0.40 },
                new CityZoneOption { District = "Banke", Municipality = "Kohalpur Nagarpalika", ZValue = 0.40 },
                new CityZoneOption { District = "Bardiya", Municipality = "Gulariya Nagarpalika", ZValue = 0.40 },
                new CityZoneOption { District = "Surkhet", Municipality = "Birendranagar Nagarpalika", ZValue = 0.40 },
                new CityZoneOption { District = "Jumla", Municipality = "Chandannath Nagarpalika", ZValue = 0.30 },
                new CityZoneOption { District = "Dolpa", Municipality = "Dolpo Buddha Gaunpalika", ZValue = 0.25 },
                new CityZoneOption { District = "Dolpa", Municipality = "Shey Phoksundo Gaunpalika", ZValue = 0.25 },
                new CityZoneOption { District = "Kailali", Municipality = "Dhangadhi Upamahanagarpalika", ZValue = 0.40 },
                new CityZoneOption { District = "Kailali", Municipality = "Tikapur Nagarpalika", ZValue = 0.35 },
                new CityZoneOption { District = "Kanchanpur", Municipality = "Bhimdatta Nagarpalika", ZValue = 0.35 }
            };
            UpdateFilteredCities();
            _selectedCity = FilteredCities.FirstOrDefault();

            // 2. Populate Site Subsoil Categories (Table 4-1 & Table 4-2)
            SoilTypes = new List<SoilTypeOption>
            {
                new SoilTypeOption { Code = "A", DisplayName = "Soil Type A — Rock / Rock-like (Vs,30 > 800 m/s) [Tc=0.5s, Td=4.0s, α=2.5]", Ta = 0.0, Tc = 0.5, Td = 4.0, Alpha = 2.5 },
                new SoilTypeOption { Code = "B", DisplayName = "Soil Type B — Very Dense Sand / Stiff Clay (350 < Vs,30 ≤ 800 m/s) [Tc=0.7s, Td=4.0s, α=2.5]", Ta = 0.0, Tc = 0.7, Td = 4.0, Alpha = 2.5 },
                new SoilTypeOption { Code = "C", DisplayName = "Soil Type C — Medium-Dense Sand / Stiff Clay (150 < Vs,30 ≤ 350 m/s) [Tc=1.0s, Td=4.0s, α=2.5]", Ta = 0.0, Tc = 1.0, Td = 4.0, Alpha = 2.5 },
                new SoilTypeOption { Code = "D", DisplayName = "Soil Type D — Loose / Soft Soil (Vs,30 ≤ 150 m/s or Table 4-3 Kathmandu Valley) [Tc=2.0s, Td=5.0s, α=2.25]", Ta = 0.0, Tc = 2.0, Td = 5.0, Alpha = 2.25 }
            };
            _selectedSoilType = SoilTypes[2]; // Default Type C

            // 3. Populate Importance Classes (Table 4-4)
            ImportanceClasses = new List<ImportanceOption>
            {
                new ImportanceOption { DisplayName = "Class I (I = 1.00) — Ordinary Structures (Residences, Offices, Commercial < 500 persons)", IValue = 1.00 },
                new ImportanceOption { DisplayName = "Class II (I = 1.25) — Schools, Colleges, Malls, Temples, Assembly > 500 persons", IValue = 1.25 },
                new ImportanceOption { DisplayName = "Class II as Disaster Shelter (I = 1.50) — Class II facility used as shelter (Table 4-4 Note 2)", IValue = 1.50 },
                new ImportanceOption { DisplayName = "Class III (I = 1.50) — Hospitals, Fire Stations, Power Stations, Critical Lifelines", IValue = 1.50 }
            };
            _selectedImportance = ImportanceClasses[0];

            // 4. Populate Building Regularity Options (Clause 3.2.1 & Clause 5.4)
            RegularityOptions = new List<string>
            {
                "Regular Structure (Meets Clause 5.4 criteria)",
                "Irregular Structure (Vertical or Plan Irregularity per Clause 5.4)"
            };

            // 5. Populate Empirical Period Coefficient (kt) options (Clause 5.1.2)
            KtOptionsX = new List<EmpiricalKtOption>
            {
                new EmpiricalKtOption { DisplayName = "Moment Resisting Concrete Frame (kt = 0.075)", KtBase = 0.075, IsConcreteShearWall = false },
                new EmpiricalKtOption { DisplayName = "Moment Resisting Structural Steel Frame (kt = 0.085)", KtBase = 0.085, IsConcreteShearWall = false },
                new EmpiricalKtOption { DisplayName = "Eccentrically Braced Structural Steel Frame (kt = 0.075)", KtBase = 0.075, IsConcreteShearWall = false },
                new EmpiricalKtOption { DisplayName = "Building with Concrete Shear Walls (kt = 0.075 / √Aw)", KtBase = 0.075, IsConcreteShearWall = true },
                new EmpiricalKtOption { DisplayName = "All Other Structural Systems (kt = 0.050)", KtBase = 0.050, IsConcreteShearWall = false }
            };
            KtOptionsY = new List<EmpiricalKtOption>(KtOptionsX);
            _selectedKtX = KtOptionsX[0];
            _selectedKtY = KtOptionsY[0];

            // 6. Populate the Entire 21-Row NBC 105:2025 Table 5-2
            var systems = new List<StructuralSystemRow>
            {
                new StructuralSystemRow { SerialNo = 1, Category = "Moment Resisting Frame", SystemName = "Steel Moment Resisting Frame", R_mu = 4.0, Omega_u = 1.5, Omega_s = 1.25 },
                new StructuralSystemRow { SerialNo = 2, Category = "Moment Resisting Frame", SystemName = "Reinforced Concrete Moment Resisting Frame", R_mu = 4.0, Omega_u = 1.5, Omega_s = 1.25 },
                new StructuralSystemRow { SerialNo = 3, Category = "Moment Resisting Frame", SystemName = "Steel + RC Composite Moment Resisting Frame", R_mu = 4.0, Omega_u = 1.5, Omega_s = 1.25 },
                new StructuralSystemRow { SerialNo = 4, Category = "Braced Frame System", SystemName = "Steel Eccentrically Braced Frame", R_mu = 4.0, Omega_u = 1.5, Omega_s = 1.25 },
                new StructuralSystemRow { SerialNo = 5, Category = "Braced Frame System", SystemName = "Steel + RC Composite Eccentrically Braced Frame", R_mu = 4.0, Omega_u = 1.5, Omega_s = 1.25 },
                new StructuralSystemRow { SerialNo = 6, Category = "Braced Frame System", SystemName = "Steel Concentric Braced Frame", R_mu = 3.0, Omega_u = 1.3, Omega_s = 1.15 },
                new StructuralSystemRow { SerialNo = 7, Category = "Braced Frame System", SystemName = "Steel + RC Composite Concentric Braced Frame", R_mu = 3.0, Omega_u = 1.3, Omega_s = 1.15 },
                new StructuralSystemRow { SerialNo = 8, Category = "Braced Frame System", SystemName = "Steel Buckling Restraint Braces", R_mu = 4.0, Omega_u = 1.5, Omega_s = 1.25 },
                new StructuralSystemRow { SerialNo = 9, Category = "Structural Wall System", SystemName = "RC Shear Wall", R_mu = 3.0, Omega_u = 1.3, Omega_s = 1.15 },
                new StructuralSystemRow { SerialNo = 10, Category = "Structural Wall System", SystemName = "Steel + RC Composite Shear Wall", R_mu = 3.0, Omega_u = 1.3, Omega_s = 1.15 },
                new StructuralSystemRow { SerialNo = 11, Category = "Structural Wall System", SystemName = "Reinforced Masonry Shear Wall", R_mu = 2.5, Omega_u = 1.2, Omega_s = 1.10 },
                new StructuralSystemRow { SerialNo = 12, Category = "Structural Wall System", SystemName = "Confined Masonry Wall", R_mu = 2.5, Omega_u = 1.2, Omega_s = 1.10 },
                new StructuralSystemRow { SerialNo = 13, Category = "Structural Wall System", SystemName = "Unreinforced Masonry Wall with Horizontal Bands & Vertical Bars", R_mu = 2.0, Omega_u = 1.2, Omega_s = 1.10 },
                new StructuralSystemRow { SerialNo = 14, Category = "Dual System", SystemName = "Steel Eccentrically Braced Frame", R_mu = 4.0, Omega_u = 1.5, Omega_s = 1.25 },
                new StructuralSystemRow { SerialNo = 15, Category = "Dual System", SystemName = "Steel + RC Composite Eccentrically Braced Frame", R_mu = 4.0, Omega_u = 1.5, Omega_s = 1.25 },
                new StructuralSystemRow { SerialNo = 16, Category = "Dual System", SystemName = "Steel Concentric Braced Frame", R_mu = 3.5, Omega_u = 1.4, Omega_s = 1.20 },
                new StructuralSystemRow { SerialNo = 17, Category = "Dual System", SystemName = "Steel + RC Composite Concentric Braced Frame", R_mu = 3.5, Omega_u = 1.4, Omega_s = 1.20 },
                new StructuralSystemRow { SerialNo = 18, Category = "Dual System", SystemName = "Steel Buckling Restraint Braces", R_mu = 4.0, Omega_u = 1.5, Omega_s = 1.25 },
                new StructuralSystemRow { SerialNo = 19, Category = "Dual System", SystemName = "RC Shear Wall", R_mu = 3.5, Omega_u = 1.4, Omega_s = 1.20 },
                new StructuralSystemRow { SerialNo = 20, Category = "Dual System", SystemName = "Steel + RC Composite Shear Wall", R_mu = 3.5, Omega_u = 1.4, Omega_s = 1.20 },
                new StructuralSystemRow { SerialNo = 21, Category = "Dual System", SystemName = "Reinforced Masonry Shear Wall", R_mu = 2.5, Omega_u = 1.2, Omega_s = 1.10 }
            };

            foreach (var row in systems)
            {
                row.OnSelectionChanged = HandleTable5_2SelectionChanged;
                Table5_2Rows.Add(row);
            }

            // Default to S.No. 2 (Reinforced Concrete Moment Resisting Frame) for both X and Y
            _selectedSystemX = Table5_2Rows[1];
            _selectedSystemY = Table5_2Rows[1];
            _selectedSystemX.SetSelectedSilent(true, true);
            _selectedSystemY.SetSelectedSilent(false, true);

            Recalculate();
        }

        /// <summary>
        /// Filters the Annex C City/Municipality dropdown whenever the user types in the Search City box.
        /// </summary>
        private void UpdateFilteredCities()
        {
            FilteredCities.Clear();
            string query = (_citySearchText ?? string.Empty).Trim();

            var matches = string.IsNullOrEmpty(query)
                ? _allCities
                : _allCities.Where(c =>
                    c.Municipality.Contains(query, StringComparison.OrdinalIgnoreCase) ||
                    c.District.Contains(query, StringComparison.OrdinalIgnoreCase)).ToList();

            foreach (var city in matches)
            {
                FilteredCities.Add(city);
            }

            // Automatically select the first matching city so Z updates immediately while searching
            if (FilteredCities.Count > 0)
            {
                SelectedCity = FilteredCities[0];
            }
        }

        /// <summary>
        /// Ensures radio-button behavior across rows when the user checks 'EQX' or 'EQY' in the Table 5-2 DataGrid.
        /// </summary>
        private void HandleTable5_2SelectionChanged(StructuralSystemRow selectedRow, bool isXDirection)
        {
            foreach (var row in Table5_2Rows)
            {
                if (row != selectedRow)
                {
                    row.SetSelectedSilent(isXDirection, false);
                }
            }

            if (isXDirection)
                _selectedSystemX = selectedRow;
            else
                _selectedSystemY = selectedRow;

            Recalculate();
        }

        /// <summary>
        /// Computes Spectral Shape Factor Ch(T) for the Equivalent Static Method as per Clause 4.1.2 (Eq. 4.1(2) & Table 4-1).
        /// </summary>
        private double CalculateSpectralShapeFactorESM(double T, SoilTypeOption soil)
        {
            if (T < soil.Tc)
                return soil.Alpha;
            if (T < soil.Td)
                return (soil.Alpha * soil.Tc) / T;
            return (soil.Alpha * soil.Tc * soil.Td) / Math.Pow(T, 2);
        }

        /// <summary>
        /// Computes vertical distribution exponent k as per Clause 6.3 (Eq. 6.3(1)).
        /// </summary>
        private double CalculateExponentK(double T)
        {
            if (T <= 0.5) return 1.0;
            if (T >= 2.5) return 2.0;
            return 1.0 + (T - 0.5) / 2.0;
        }

        /// <summary>
        /// Performs the NBC 105:2025 Equivalent Static Method calculation and updates all UI properties.
        /// </summary>
        public void Recalculate()
        {
            if (_selectedSoilType == null || _selectedImportance == null ||
                _selectedKtX == null || _selectedKtY == null || _selectedSystemX == null || _selectedSystemY == null)
                return;

            // Update active Table 5-2 selection summary labels
            SelectedSystemXSummary = $"Active EQX System (#{_selectedSystemX.SerialNo}): {_selectedSystemX.SystemName}  [Rμ = {_selectedSystemX.R_mu}, Ωu = {_selectedSystemX.Omega_u}, Ωs = {_selectedSystemX.Omega_s}]";
            SelectedSystemYSummary = $"Active EQY System (#{_selectedSystemY.SerialNo}): {_selectedSystemY.SystemName}  [Rμ = {_selectedSystemY.R_mu}, Ωu = {_selectedSystemY.Omega_u}, Ωs = {_selectedSystemY.Omega_s}]";

            // Validate Manual / City Seismic Zone Factor Z
            if (!double.TryParse(_manualZText, out double Z) || Z <= 0)
            {
                BreakdownText = "Please enter a valid positive Seismic Zoning Factor Z (e.g., 0.25, 0.30, 0.35, 0.40).";
                return;
            }

            // Validate Building Height H (m)
            if (!double.TryParse(_heightText, out double H) || H <= 0)
            {
                BreakdownText = "Please enter a valid positive Building Height H (m).";
                return;
            }

            // Validate Accidental Eccentricity Ratio (e.g., 0.05)
            if (!double.TryParse(_eccentricityText, out double ecc) || ecc < 0 || ecc >= 1.0)
            {
                BreakdownText = "Please enter a valid Accidental Eccentricity ratio between 0.00 and 0.99 (Clause 5.6 specifies 0.05).";
                return;
            }
            LatestEccentricity = ecc;

            // 1. Empirical coefficient kt for X and Y (Clause 5.1.2)
            double ktX = _selectedKtX.KtBase;
            if (_selectedKtX.IsConcreteShearWall)
            {
                if (!double.TryParse(_awXText, out double awX) || awX <= 0)
                {
                    BreakdownText = "Please enter a valid positive Effective Shear Wall Area Aw (m²) for X-Direction.";
                    return;
                }
                ktX = 0.075 / Math.Sqrt(awX);
            }

            double ktY = _selectedKtY.KtBase;
            if (_selectedKtY.IsConcreteShearWall)
            {
                if (!double.TryParse(_awYText, out double awY) || awY <= 0)
                {
                    BreakdownText = "Please enter a valid positive Effective Shear Wall Area Aw (m²) for Y-Direction.";
                    return;
                }
                ktY = 0.075 / Math.Sqrt(awY);
            }

            // 2. Amplified Approximate Period T1 = 1.25 * kt * H^(0.75) (Clauses 5.1.2 & 5.1.3)
            double tAmpX = 1.25 * ktX * Math.Pow(H, 0.75);
            double tAmpY = 1.25 * ktY * Math.Pow(H, 0.75);

            // 3. Adopt lesser of Amplified Empirical Period and Rayleigh Period if checked (Clause 5.1)
            double T1x = tAmpX;
            double T1y = tAmpY;
            if (_useRayleigh)
            {
                if (double.TryParse(_rayleighXText, out double tRayX) && tRayX > 0)
                    T1x = Math.Min(tAmpX, tRayX);
                if (double.TryParse(_rayleighYText, out double tRayY) && tRayY > 0)
                    T1y = Math.Min(tAmpY, tRayY);
            }

            // 4. Check ESM Applicability (Clause 3.2.1)
            bool isRegular = _selectedRegularityIndex == 0;
            bool esmValidX = (H <= 15.0) || (T1x < 0.5) || (isRegular && H < 40.0);
            bool esmValidY = (H <= 15.0) || (T1y < 0.5) || (isRegular && H < 40.0);

            if (esmValidX && esmValidY)
            {
                ApplicabilityBackground = new SolidColorBrush(Color.FromRgb(220, 252, 231));
                ApplicabilityBorder = new SolidColorBrush(Color.FromRgb(134, 239, 172));
                ApplicabilityForeground = new SolidColorBrush(Color.FromRgb(22, 101, 52));
                ApplicabilityMessage = "✓ Clause 3.2.1 Check: Equivalent Static Method (ESM) is VALID for both SLS and ULS.";
            }
            else
            {
                ApplicabilityBackground = new SolidColorBrush(Color.FromRgb(254, 243, 199));
                ApplicabilityBorder = new SolidColorBrush(Color.FromRgb(252, 211, 77));
                ApplicabilityForeground = new SolidColorBrush(Color.FromRgb(146, 64, 14));
                ApplicabilityMessage = "⚠ Clause 3.2.1 Warning: ESM is valid for SLS, but ULS requires Modal Response Spectrum Method (H > 15m, T1 ≥ 0.5s, and Irregular or H ≥ 40m).";
            }

            // 5. Spectral Shape Factor Ch(T1) (Clause 4.1.2)
            double ChX = CalculateSpectralShapeFactorESM(T1x, _selectedSoilType);
            double ChY = CalculateSpectralShapeFactorESM(T1y, _selectedSoilType);

            // 6. Elastic Site Spectra for ULS (Clause 4.1.1) and SLS (Clause 4.2)
            double I = _selectedImportance.IValue;
            double C_uls_X = ChX * Z * I;
            double C_uls_Y = ChY * Z * I;
            double Cs_sls_X = 0.20 * C_uls_X;
            double Cs_sls_Y = 0.20 * C_uls_Y;

            // 7. Horizontal Base Shear Coefficients Cd(T1) (Clauses 6.1.1 & 6.1.2)
            LatestCd_EQX_SLS = Cs_sls_X / _selectedSystemX.Omega_s;
            LatestCd_EQY_SLS = Cs_sls_Y / _selectedSystemY.Omega_s;
            LatestCd_EQX_ULS = C_uls_X / (_selectedSystemX.R_mu * _selectedSystemX.Omega_u);
            LatestCd_EQY_ULS = C_uls_Y / (_selectedSystemY.R_mu * _selectedSystemY.Omega_u);

            LatestKx = CalculateExponentK(T1x);
            LatestKy = CalculateExponentK(T1y);

            // 8. Update UI Display Strings
            EqxSlsResult = LatestCd_EQX_SLS.ToString("F4");
            EqySlsResult = LatestCd_EQY_SLS.ToString("F4");
            EqxUlsResult = LatestCd_EQX_ULS.ToString("F4");
            EqyUlsResult = LatestCd_EQY_ULS.ToString("F4");

            // 9. Build Detailed Breakdown Log
            var sb = new StringBuilder();
            sb.AppendLine("=== SITE & SEISMIC HAZARD (SECTION 4 & CLAUSE 5.6) ===");
            sb.AppendLine($"• Seismic Zoning Factor, Z (Cl. 4.1.4)    = {Z:F3}");
            sb.AppendLine($"• Importance Factor, I (Cl. 4.1.5)        = {I:F2}");
            sb.AppendLine($"• Subsoil Category (Cl. 4.1.3)            = Type {_selectedSoilType.Code} (Ta=0, Tc={_selectedSoilType.Tc}s, Td={_selectedSoilType.Td}s, α={_selectedSoilType.Alpha})");
            sb.AppendLine($"• Accidental Eccentricity Ratio (Cl. 5.6) = ±{LatestEccentricity:F3} b");
            sb.AppendLine();
            sb.AppendLine("=== X-DIRECTION (EQX) PARAMETERS ===");
            sb.AppendLine($"• Selected System (Table 5-2 #{_selectedSystemX.SerialNo})          = {_selectedSystemX.SystemName}");
            sb.AppendLine($"• Adopted Design Period, T1x (Cl. 5.1)    = {T1x:F4} sec (kx = {LatestKx:F3})");
            sb.AppendLine($"• Spectral Shape Factor, Ch(T1x) (4.1.2)  = {ChX:F4}");
            sb.AppendLine($"• ULS Elastic Spectrum, C(T1x) (Cl. 4.1.1)= {C_uls_X:F4}");
            sb.AppendLine($"• SLS Elastic Spectrum, Cs(T1x) (Cl. 4.2) = {Cs_sls_X:F4}");
            sb.AppendLine($"• Factors: Rμ={_selectedSystemX.R_mu}, Ωu={_selectedSystemX.Omega_u}, Ωs={_selectedSystemX.Omega_s}");
            sb.AppendLine($"  -> Cd(EQX SLS) = {Cs_sls_X:F4} / {_selectedSystemX.Omega_s}       = {LatestCd_EQX_SLS:F4}");
            sb.AppendLine($"  -> Cd(EQX ULS) = {C_uls_X:F4} / ({_selectedSystemX.R_mu}*{_selectedSystemX.Omega_u}) = {LatestCd_EQX_ULS:F4}");
            sb.AppendLine();
            sb.AppendLine("=== Y-DIRECTION (EQY) PARAMETERS ===");
            sb.AppendLine($"• Selected System (Table 5-2 #{_selectedSystemY.SerialNo})          = {_selectedSystemY.SystemName}");
            sb.AppendLine($"• Adopted Design Period, T1y (Cl. 5.1)    = {T1y:F4} sec (ky = {LatestKy:F3})");
            sb.AppendLine($"• Spectral Shape Factor, Ch(T1y) (4.1.2)  = {ChY:F4}");
            sb.AppendLine($"• ULS Elastic Spectrum, C(T1y) (Cl. 4.1.1)= {C_uls_Y:F4}");
            sb.AppendLine($"• SLS Elastic Spectrum, Cs(T1y) (Cl. 4.2) = {Cs_sls_Y:F4}");
            sb.AppendLine($"• Factors: Rμ={_selectedSystemY.R_mu}, Ωu={_selectedSystemY.Omega_u}, Ωs={_selectedSystemY.Omega_s}");
            sb.AppendLine($"  -> Cd(EQY SLS) = {Cs_sls_Y:F4} / {_selectedSystemY.Omega_s}       = {LatestCd_EQY_SLS:F4}");
            sb.AppendLine($"  -> Cd(EQY ULS) = {C_uls_Y:F4} / ({_selectedSystemY.R_mu}*{_selectedSystemY.Omega_u}) = {LatestCd_EQY_ULS:F4}");

            BreakdownText = sb.ToString();
        }

        /// <summary>
        /// Connects to the running ETABS instance and configures EQX ULS, EQY ULS, EQX SLS, and EQY SLS
        /// with "User Coefficient" using cDatabaseTables.
        /// Applies X, X+E, and X-E simultaneously for X-direction loads, and Y, Y+E, Y-E for Y-direction loads.
        /// </summary>
        public void PushLoadsToEtabs()
        {
            try
            {
                // 1. Instantiate the ETABSv1 Helper to attach to the active ETABS process
                // Syntax Example: cHelper myHelper = new Helper();
                // myHelper connects the WPF application to the running ETABS COM server
                cHelper myHelper = new Helper();

                // 2. Get the active ETABS application object from the Running Object Table
                // Syntax Example: cOAPI myETABSObject = myHelper.GetObject("CSI.ETABS.API.ETABSObject");
                cOAPI myETABSObject = myHelper.GetObject("CSI.ETABS.API.ETABSObject");
                if (myETABSObject == null)
                {
                    EtabsStatusBrush = new SolidColorBrush(Color.FromRgb(185, 28, 28));
                    EtabsStatusMessage = "✗ Could not find an active ETABS instance. Please open ETABS first.";
                    MessageBox.Show("No running ETABS instance was found.\n\nPlease open ETABS with an active structural model first.", "ETABS Not Found", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // 3. Retrieve the SapModel interface
                // Syntax Example: cSapModel mySapModel = myETABSObject.SapModel;
                cSapModel mySapModel = myETABSObject.SapModel;
                if (mySapModel == null)
                {
                    EtabsStatusBrush = new SolidColorBrush(Color.FromRgb(185, 28, 28));
                    EtabsStatusMessage = "✗ Connected to ETABS, but SapModel is null. Please open a model.";
                    return;
                }

                // 4. Unlock the model so load patterns and database tables can be modified
                // Syntax Example: mySapModel.SetModelIsLocked(false);
                mySapModel.SetModelIsLocked(false);

                // Define the 4 NBC 105:2025 Equivalent Static load patterns to create
                // IsXDir = true for EQX, false for EQY
                var loadsToCreate = new[]
                {
            new { Name = "EQX ULS", IsXDir = true,  C = LatestCd_EQX_ULS, K = LatestKx },
            new { Name = "EQY ULS", IsXDir = false, C = LatestCd_EQY_ULS, K = LatestKy },
            new { Name = "EQX SLS", IsXDir = true,  C = LatestCd_EQX_SLS, K = LatestKx },
            new { Name = "EQY SLS", IsXDir = false, C = LatestCd_EQY_SLS, K = LatestKy }
        };

                // 5. Add the 4 Quake Load Patterns via LoadPatterns.Add
                // Syntax Example: mySapModel.LoadPatterns.Add("EQX ULS", eLoadPatternType.Quake, 0.0, true);
                foreach (var load in loadsToCreate)
                {
                    mySapModel.LoadPatterns.Add(load.Name, eLoadPatternType.Quake, 0.0, true);
                }

                // =========================================================================
                // STEP A: Set AutoLoad = "User Coefficient" in "Load Pattern Definitions"
                // =========================================================================
                string lpTableKey = "Load Pattern Definitions";
                int lpTableVersion = 0;
                string[] lpFields = Array.Empty<string>();
                int lpNumRecords = 0;
                string[] lpTableData = Array.Empty<string>();

                int retGetLp = mySapModel.DatabaseTables.GetTableForEditingArray(
                    lpTableKey, "", ref lpTableVersion, ref lpFields, ref lpNumRecords, ref lpTableData);

                if (retGetLp == 0 && lpFields != null && lpFields.Length > 0)
                {
                    int colCount = lpFields.Length;
                    int idxName = Array.IndexOf(lpFields, "Name");
                    if (idxName < 0) idxName = Array.IndexOf(lpFields, "LoadPat");
                    int idxAutoLoad = Array.IndexOf(lpFields, "AutoLoad");

                    if (idxName >= 0 && idxAutoLoad >= 0)
                    {
                        for (int r = 0; r < lpNumRecords; r++)
                        {
                            string patName = lpTableData[r * colCount + idxName];
                            if (loadsToCreate.Any(l => string.Equals(l.Name, patName, StringComparison.OrdinalIgnoreCase)))
                            {
                                lpTableData[r * colCount + idxAutoLoad] = "User Coefficient";
                            }
                        }

                        mySapModel.DatabaseTables.SetTableForEditingArray(
                            lpTableKey, ref lpTableVersion, ref lpFields, lpNumRecords, ref lpTableData);

                        int fatalErr = 0, errMsg = 0, warnMsg = 0, infoMsg = 0;
                        string importLog = "";
                        mySapModel.DatabaseTables.ApplyEditedTables(true, ref fatalErr, ref errMsg, ref warnMsg, ref infoMsg, ref importLog);
                    }
                }

                // =========================================================================
                // STEP B: Populate "Load Pattern Definitions - Auto Seismic - User Coefficient"
                // =========================================================================
                string ucTableKey = "Load Pattern Definitions - Auto Seismic - User Coefficient";

                string[] ucFields = new string[]
                {
            "Name",
            "XDir", "XDirPlusE", "XDirMinusE",
            "YDir", "YDirPlusE", "YDirMinusE",
            "EccRatio", "C", "K"
                };

                var ucDataList = new List<string>();

                // Check if user entered a non-zero accidental eccentricity[cite: 199]
                bool hasEcc = LatestEccentricity > 0.0;
                string eccStr = LatestEccentricity.ToString("F4", System.Globalization.CultureInfo.InvariantCulture);

                foreach (var load in loadsToCreate)
                {
                    string cStr = Math.Round(load.C, 6).ToString(System.Globalization.CultureInfo.InvariantCulture);
                    string kStr = Math.Round(load.K, 4).ToString(System.Globalization.CultureInfo.InvariantCulture);

                    // X-direction logic: pure X is always "Yes" if IsXDir is true. PlusE and MinusE follow hasEcc.
                    string xDir = load.IsXDir ? "Yes" : "No";
                    string xDirPlus = (load.IsXDir && hasEcc) ? "Yes" : "No";
                    string xDirMinus = (load.IsXDir && hasEcc) ? "Yes" : "No";

                    // Y-direction logic: pure Y is always "Yes" if IsXDir is false. PlusE and MinusE follow hasEcc.
                    string yDir = !load.IsXDir ? "Yes" : "No";
                    string yDirPlus = (!load.IsXDir && hasEcc) ? "Yes" : "No";
                    string yDirMinus = (!load.IsXDir && hasEcc) ? "Yes" : "No";

                    // Add all column values to the 1D flat array required by ETABS DatabaseTables
                    ucDataList.Add(load.Name);
                    ucDataList.Add(xDir);
                    ucDataList.Add(xDirPlus);
                    ucDataList.Add(xDirMinus);
                    ucDataList.Add(yDir);
                    ucDataList.Add(yDirPlus);
                    ucDataList.Add(yDirMinus);
                    ucDataList.Add(eccStr);
                    ucDataList.Add(cStr);
                    ucDataList.Add(kStr);
                }

                int ucTableVersion = 1;
                int ucNumRecords = loadsToCreate.Length;
                string[] ucTableData = ucDataList.ToArray();

                // Stage the User Coefficient table in ETABS
                // Syntax Example: mySapModel.DatabaseTables.SetTableForEditingArray(ucTableKey, ref ucTableVersion, ref ucFields, ucNumRecords, ref ucTableData);
                int retSetUc = mySapModel.DatabaseTables.SetTableForEditingArray(
                    ucTableKey, ref ucTableVersion, ref ucFields, ucNumRecords, ref ucTableData);

                int numFatalErrors = 0;
                int numErrorMsgs = 0;
                int numWarnMsgs = 0;
                int numInfoMsgs = 0;
                string applyLog = string.Empty;

                // Apply all edited tables to the active ETABS model
                int retApply = mySapModel.DatabaseTables.ApplyEditedTables(
                    true, ref numFatalErrors, ref numErrorMsgs, ref numWarnMsgs, ref numInfoMsgs, ref applyLog);

                // Refresh the ETABS GUI window
                mySapModel.View.RefreshView(0, false);

                if (retSetUc == 0 && retApply == 0 && numFatalErrors == 0 && numErrorMsgs == 0)
                {
                    EtabsStatusBrush = new SolidColorBrush(Color.FromRgb(22, 101, 52));
                    EtabsStatusMessage = $"✓ Successfully created 'EQX ULS', 'EQY ULS', 'EQX SLS', and 'EQY SLS' in ETABS (Ecc = {LatestEccentricity:F3})!";
                    MessageBox.Show(
                        $"Successfully created/updated all 4 Seismic Load Patterns & Load Cases in ETABS:\n\n" +
                        $"• EQX ULS: X, X+E, X-E (C = {LatestCd_EQX_ULS:F4}, k = {LatestKx:F3})\n" +
                        $"• EQY ULS: Y, Y+E, Y-E (C = {LatestCd_EQY_ULS:F4}, k = {LatestKy:F3})\n" +
                        $"• EQX SLS: X, X+E, X-E (C = {LatestCd_EQX_SLS:F4}, k = {LatestKx:F3})\n" +
                        $"• EQY SLS: Y, Y+E, Y-E (C = {LatestCd_EQY_SLS:F4}, k = {LatestKy:F3})",
                        "ETABS Load Creation Successful",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
                else
                {
                    EtabsStatusBrush = new SolidColorBrush(Color.FromRgb(180, 83, 9));
                    EtabsStatusMessage = $"⚠ Applied with messages (Errors: {numErrorMsgs}, Warnings: {numWarnMsgs}). Details: {applyLog}";
                    MessageBox.Show(
                        $"ETABS Database Table Import Log:\n\n{applyLog}",
                        "ETABS Table Import Log",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                EtabsStatusBrush = new SolidColorBrush(Color.FromRgb(185, 28, 28));
                EtabsStatusMessage = $"✗ ETABS API Error: {ex.Message}";
                MessageBox.Show($"Error communicating with ETABS:\n\n{ex.Message}", "ETABS API Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    /// <summary>
    /// Interaction logic for MainWindow.xaml.
    /// </summary>
    public partial class MainWindow : Window
    {
        // Holds reference to the MainViewModel instance bound to the Window's DataContext
        private readonly MainViewModel _viewModel;

        public MainWindow()
        {
            InitializeComponent();
            _viewModel = new MainViewModel();
            DataContext = _viewModel;
        }

        // Event handler for the "Create ETABS Load" button click
        private void CreateEtabsLoad_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.PushLoadsToEtabs();
        }
    }
}