using System;
using System.CodeDom.Compiler;
using System.Diagnostics;
using System.Reflection;
using System.Windows;
using Microsoft.VisualBasic.ApplicationServices;

namespace arv_mu_computation;

public class Application : System.Windows.Application
{
	internal AssemblyInfo Info
	{
		[DebuggerHidden]
		get
		{
			return new AssemblyInfo(Assembly.GetExecutingAssembly());
		}
	}

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void InitializeComponent()
	{
		// Стартовое окно создаётся в Main() (UI загружается из mainwindow.baml).
	}

	[STAThread]
	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public static void Main()
	{
		var app = new Application();
		app.InitializeComponent();
		var mainWindow = new MainWindow();
		app.MainWindow = mainWindow;
		mainWindow.Show();
		app.Run();
	}
}
