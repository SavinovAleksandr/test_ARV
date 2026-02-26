using System;
using System.CodeDom.Compiler;
using System.Diagnostics;
using System.Reflection;
using System.Windows;
using Microsoft.VisualBasic.ApplicationServices;

namespace arv_mu_computation;

public class Application : Application
{
	internal AssemblyInfo Info
	{
		[DebuggerHidden]
		get
		{
			//IL_0006: Unknown result type (might be due to invalid IL or missing references)
			//IL_000c: Expected O, but got Unknown
			return new AssemblyInfo(Assembly.GetExecutingAssembly());
		}
	}

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void InitializeComponent()
	{
		((Application)this).StartupUri = new Uri("MainWindow.xaml", UriKind.Relative);
	}

	[STAThread]
	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public static void Main()
	{
		Application application = new Application();
		application.InitializeComponent();
		((Application)application).Run();
	}
}
