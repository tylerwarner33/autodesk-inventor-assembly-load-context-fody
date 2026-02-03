using Inventor;
using System.Reflection;
using System.Runtime.InteropServices;

namespace IsolatedInventorAddin;

/// <summary>
///	This is the primary <see cref="AddinServer"/> class that implements the <see cref="ApplicationAddInServer"/> interface that all Inventor add-ins are required to implement. 
///	The communication between Inventor and the add-in is via the methods on this interface.
/// </summary>
[Guid("054BFF0D-FA7E-43FB-AA61-302D0C34D9CA")]
public class AddinServer : ApplicationAddInServer
{
	private Inventor.Application? _inventorApplication;
	private UserInterfaceEvents? _userInterfaceEvents;
	private UserInterfaceEventsSink_OnResetRibbonInterfaceEventHandler? _userInterfaceEventsSink_OnResetRibbonInterfaceEventDelegate;

	private RibbonPanel? _assemblyLoadContextPanel;
	private SerilogPackageVersionButton? _packageVersionButton;
	private const string _packageVersionButtonInternalName = $"{nameof(Inventor)}:{nameof(IsolatedInventorAddin)}.Fody:{nameof(SerilogPackageVersionButton)}";

	private readonly string[] _ribbonNames = ["ZeroDoc"];

	public static GuidAttribute? PluginClientGuid => (GuidAttribute?)System.Attribute.GetCustomAttribute(typeof(AddinServer), typeof(GuidAttribute));

	public static string PluginClientId => "{" + PluginClientGuid?.Value + "}";

	public const string AppPrefix = $"{nameof(IsolatedInventorAddin)}";

	public static string? AppVersion => Assembly.GetExecutingAssembly().GetName().Version?.ToString(3);

	public static string? AppPrefixWithVersion => $"{AppPrefix} {AppVersion}";

	public void Activate(ApplicationAddInSite applicationAddInSite, bool firstTime)
	{
		try
		{
			_inventorApplication = applicationAddInSite.Application;
			_userInterfaceEvents = _inventorApplication.UserInterfaceManager.UserInterfaceEvents;

			_userInterfaceEventsSink_OnResetRibbonInterfaceEventDelegate = new UserInterfaceEventsSink_OnResetRibbonInterfaceEventHandler(UserInterfaceEvents_OnResetRibbonInterface);
			_userInterfaceEvents.OnResetRibbonInterface += _userInterfaceEventsSink_OnResetRibbonInterfaceEventDelegate;

			_packageVersionButton = new SerilogPackageVersionButton(applicationAddInSite.Application);

			if (firstTime)
				AddRibbonCustomization();
		}
		catch (Exception ex)
		{
			MessageBox.Show($"Unable to customize the ribbon.\n\n{ex}");
		}
	}

	public void Deactivate()
	{
		try
		{
			RemoveRibbonCustomization();

			_userInterfaceEvents!.OnResetRibbonInterface -= _userInterfaceEventsSink_OnResetRibbonInterfaceEventDelegate;
			_userInterfaceEventsSink_OnResetRibbonInterfaceEventDelegate = null;

			_userInterfaceEvents = null;

			if (_inventorApplication is not null)
			{
				Marshal.ReleaseComObject(_inventorApplication);
				_inventorApplication = null;
			}

			GC.WaitForPendingFinalizers();
			GC.Collect();
		}
		catch (Exception ex)
		{
			MessageBox.Show(ex.ToString());
		}
	}

	private void UserInterfaceEvents_OnResetRibbonInterface(NameValueMap context)
	{
		try
		{
			RemoveRibbonCustomization();
			AddRibbonCustomization();
		}
		catch (Exception ex)
		{
			MessageBox.Show(ex.ToString());
		}
	}

	private void AddRibbonCustomization()
	{
		UserInterfaceManager userInterfaceManager = _inventorApplication!.UserInterfaceManager;

		if (userInterfaceManager.InterfaceStyle != InterfaceStyleEnum.kRibbonInterface) return;

		foreach (string ribbonName in _ribbonNames)
		{
			try
			{
				Ribbon ribbon = userInterfaceManager.Ribbons[ribbonName];
				RibbonTab toolsRibbonTab = ribbon.RibbonTabs["id_TabTools"];

				_assemblyLoadContextPanel = toolsRibbonTab.RibbonPanels.Add("AssemblyLoadContext", _packageVersionButtonInternalName, PluginClientId);
				_assemblyLoadContextPanel.CommandControls.AddButton(_packageVersionButton!.ButtonDefinition, true, true);
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}
		}
	}

	private void RemoveRibbonCustomization()
	{
		if (_inventorApplication is null) return;

		foreach (var ribbonName in _ribbonNames)
		{
			Ribbon ribbon = _inventorApplication!.UserInterfaceManager.Ribbons[ribbonName];
			RibbonTab toolsRibbonTab = ribbon.RibbonTabs["id_TabTools"];

			_packageVersionButton?.Dispose();
			_assemblyLoadContextPanel?.Delete();

			if (_inventorApplication is not null)
			{
				Marshal.ReleaseComObject(_inventorApplication);
			}

			GC.Collect();
			GC.WaitForPendingFinalizers();
		}
	}

	public object? Automation { get; set; } = null;

	[Obsolete("Deprecated in the Inventor API. Required for legacy compatibility.")]
	public void ExecuteCommand(int commandId) { }
}
