using System;
using System.CodeDom.Compiler;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Threading;
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

	private static void WriteCrashLog(string message, Exception ex)
	{
		try
		{
			string dir = AppDomain.CurrentDomain.BaseDirectory;
			string path = Path.Combine(dir, "Ошибка_запуска.txt");
			string text = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss}\r\n{message}\r\n{(ex != null ? ex.ToString() : "")}\r\n";
			File.AppendAllText(path, text, System.Text.Encoding.UTF8);
		}
		catch { }
	}

	[STAThread]
	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public static void Main()
	{
		AppDomain.CurrentDomain.UnhandledException += (_, e) =>
		{
			WriteCrashLog("UnhandledException", e.ExceptionObject as Exception);
		};
		try
		{
			var app = new Application();
			app.DispatcherUnhandledException += (_, e) =>
			{
				WriteCrashLog("DispatcherUnhandledException", e.Exception);
				e.Handled = true;
			};
			app.InitializeComponent();
			var mainWindow = new MainWindow();
			app.MainWindow = mainWindow;
			mainWindow.Show();
			app.Run();
		}
		catch (Exception ex)
		{
			WriteCrashLog("Main()", ex);
			throw;
		}
	}
}
