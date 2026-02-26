using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using ASTRALib;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using Microsoft.VisualBasic.Devices;
using ForLoopControl = Microsoft.VisualBasic.CompilerServices.ObjectFlowControl.ForLoopControl;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using OxyPlot.Wpf;
using arv_mu_computation.My;
using arv_mu_computation.My.Resources;
using HorizontalAlignment = System.Windows.HorizontalAlignment;
using Brushes = System.Windows.Media.Brushes;
using Brush = System.Windows.Media.Brush;
using FontWeights = System.Windows.FontWeights;
using UIElement = System.Windows.UIElement;
using Color = System.Windows.Media.Color;
using Colors = System.Windows.Media.Colors;
using Font = Microsoft.Office.Interop.Word.Font;

namespace arv_mu_computation;

[DesignerGenerated]
public class MainWindow : System.Windows.Window, IComponentConnector
{
	public delegate void updPrBar(DependencyProperty dp, object value);

	[CompilerGenerated]
	internal sealed class _Closure_0024__10_002D0
	{
		public rgmsInfo _0024VB_0024Local_rg;

		public _Closure_0024__10_002D0(_Closure_0024__10_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_rg = arg0._0024VB_0024Local_rg;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__4(rgmsInfo x)
		{
			return Operators.CompareString(x.name, _0024VB_0024Local_rg.name, false) == 0;
		}
	}

	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ _0024I;

		public static Func<rgmsInfo, bool> _0024I10_002D0;

		public static Func<rgmsInfo, bool> _0024I10_002D1;

		public static Func<int, bool> _0024I10_002D2;

		public static Func<rgmsInfo, bool> _0024I10_002D3;

		public static Func<int, bool> _0024I10_002D5;

		public static Func<datevalue, double> _0024I11_002D3;

		public static Func<datevalue, double> _0024I11_002D5;

		public static Func<datevalue, double> _0024I11_002D7;

		public static Func<datevalue, double> _0024I11_002D9;

		public static Func<datevalue, double> _0024I11_002D11;

		public static Func<datevalue, double> _0024I11_002D14;

		public static Func<datevalue, double> _0024I11_002D17;

		public static Func<datevalue, double> _0024I14_002D1;

		public static Func<datevalue, double> _0024I14_002D5;

		public static Func<datevalue, double> _0024I14_002D7;

		public static Func<datevalue, double> _0024I14_002D9;

		public static Func<datevalue, double> _0024I14_002D11;

		static _Closure_0024__()
		{
			_0024I = new _Closure_0024__();
		}

		[SpecialName]
		internal bool _Lambda_0024__10_002D0(rgmsInfo x)
		{
			return x.sName != null;
		}

		[SpecialName]
		internal bool _Lambda_0024__10_002D1(rgmsInfo x)
		{
			return x.sName != null;
		}

		[SpecialName]
		internal bool _Lambda_0024__10_002D2(int x)
		{
			return x != 0;
		}

		[SpecialName]
		internal bool _Lambda_0024__10_002D3(rgmsInfo x)
		{
			return x.sName != null;
		}

		[SpecialName]
		internal bool _Lambda_0024__10_002D5(int x)
		{
			return x != 0;
		}

		[SpecialName]
		internal double _Lambda_0024__11_002D3(datevalue n)
		{
			return n.v;
		}

		[SpecialName]
		internal double _Lambda_0024__11_002D5(datevalue n)
		{
			return n.v;
		}

		[SpecialName]
		internal double _Lambda_0024__11_002D7(datevalue n)
		{
			return n.v;
		}

		[SpecialName]
		internal double _Lambda_0024__11_002D9(datevalue n)
		{
			return n.v;
		}

		[SpecialName]
		internal double _Lambda_0024__11_002D11(datevalue n)
		{
			return n.v;
		}

		[SpecialName]
		internal double _Lambda_0024__11_002D14(datevalue n)
		{
			return n.v;
		}

		[SpecialName]
		internal double _Lambda_0024__11_002D17(datevalue n)
		{
			return n.v;
		}

		[SpecialName]
		internal double _Lambda_0024__14_002D1(datevalue x)
		{
			return x.t;
		}

		[SpecialName]
		internal double _Lambda_0024__14_002D5(datevalue x)
		{
			return x.v;
		}

		[SpecialName]
		internal double _Lambda_0024__14_002D7(datevalue x)
		{
			return x.v;
		}

		[SpecialName]
		internal double _Lambda_0024__14_002D9(datevalue x)
		{
			return x.v;
		}

		[SpecialName]
		internal double _Lambda_0024__14_002D11(datevalue x)
		{
			return x.v;
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__11_002D0
	{
		public grfInfo _0024VB_0024Local_g;

		public _Closure_0024__11_002D1 _0024VB_0024NonLocal__0024VB_0024Closure_2;

		public _Closure_0024__11_002D0(_Closure_0024__11_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_g = arg0._0024VB_0024Local_g;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__0(grfData n)
		{
			return Operators.CompareString(n.name, _0024VB_0024Local_g.name, false) == 0;
		}

		[SpecialName]
		internal bool _Lambda_0024__1(grfData n)
		{
			return Operators.CompareString(n.name, _0024VB_0024Local_g.name, false) == 0;
		}

		[SpecialName]
		internal bool _Lambda_0024__19(grfData n)
		{
			return Operators.CompareString(n.name, _0024VB_0024Local_g.name, false) == 0;
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__11_002D1
	{
		public double _0024VB_0024Local_sPoint;

		public Func<datevalue, bool> _0024I2;

		public Func<datevalue, bool> _0024I4;

		public Func<datevalue, bool> _0024I6;

		public Func<datevalue, bool> _0024I8;

		public Func<datevalue, bool> _0024I10;

		public _Closure_0024__11_002D1(_Closure_0024__11_002D1 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_sPoint = arg0._0024VB_0024Local_sPoint;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__2(datevalue n)
		{
			return n.t > _0024VB_0024Local_sPoint + 15.0;
		}

		[SpecialName]
		internal bool _Lambda_0024__4(datevalue n)
		{
			return n.t > _0024VB_0024Local_sPoint + 15.0;
		}

		[SpecialName]
		internal bool _Lambda_0024__6(datevalue n)
		{
			return (n.t < _0024VB_0024Local_sPoint) & (n.t > _0024VB_0024Local_sPoint - 2.0);
		}

		[SpecialName]
		internal bool _Lambda_0024__8(datevalue n)
		{
			return n.t > _0024VB_0024Local_sPoint;
		}

		[SpecialName]
		internal bool _Lambda_0024__10(datevalue n)
		{
			return n.t > _0024VB_0024Local_sPoint + 15.0;
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__11_002D2
	{
		public double _0024VB_0024Local_Max1;

		public double _0024VB_0024Local_tMax1;

		public double _0024VB_0024Local_Min1;

		public double _0024VB_0024Local_tMin1;

		public double _0024VB_0024Local_Max2;

		public _Closure_0024__11_002D0 _0024VB_0024NonLocal__0024VB_0024Closure_3;

		public _Closure_0024__11_002D2(_Closure_0024__11_002D2 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_Max1 = arg0._0024VB_0024Local_Max1;
				_0024VB_0024Local_tMax1 = arg0._0024VB_0024Local_tMax1;
				_0024VB_0024Local_Min1 = arg0._0024VB_0024Local_Min1;
				_0024VB_0024Local_tMin1 = arg0._0024VB_0024Local_tMin1;
				_0024VB_0024Local_Max2 = arg0._0024VB_0024Local_Max2;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__12(datevalue n)
		{
			return (n.t > _0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_sPoint + 15.0) & (n.v == _0024VB_0024Local_Max1);
		}

		[SpecialName]
		internal bool _Lambda_0024__13(datevalue n)
		{
			return n.t > _0024VB_0024Local_tMax1;
		}

		[SpecialName]
		internal bool _Lambda_0024__15(datevalue n)
		{
			return (n.t > _0024VB_0024Local_tMax1) & (n.v == _0024VB_0024Local_Min1);
		}

		[SpecialName]
		internal bool _Lambda_0024__16(datevalue n)
		{
			return n.t > _0024VB_0024Local_tMin1;
		}

		[SpecialName]
		internal bool _Lambda_0024__18(datevalue n)
		{
			return (n.t > _0024VB_0024Local_tMin1) & (n.v == _0024VB_0024Local_Max2);
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__12_002D0
	{
		public grfInfo _0024VB_0024Local_row;

		public _Closure_0024__12_002D0(_Closure_0024__12_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_row = arg0._0024VB_0024Local_row;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__0(grfData n)
		{
			return Operators.CompareString(n.name, _0024VB_0024Local_row.name, false) == 0;
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__13_002D0
	{
		public int _0024VB_0024Local_ut;

		public _Closure_0024__13_002D0(_Closure_0024__13_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_ut = arg0._0024VB_0024Local_ut;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__0(utInfo x)
		{
			return x.id == _0024VB_0024Local_ut;
		}

		[SpecialName]
		internal bool _Lambda_0024__1(utInfo x)
		{
			return x.id == _0024VB_0024Local_ut;
		}
	}

	public compModel fInf;

	private bool _contentLoaded;

	[field: AccessedThroughProperty("rgms")]
	internal virtual GroupBox rgms
	{
		get; [MethodImpl(MethodImplOptions.Synchronized)]
		set;
	}

	[field: AccessedThroughProperty("rgms_grid")]
	internal virtual Grid rgms_grid
	{
		get; [MethodImpl(MethodImplOptions.Synchronized)]
		set;
	}

	[field: AccessedThroughProperty("uts")]
	internal virtual GroupBox uts
	{
		get; [MethodImpl(MethodImplOptions.Synchronized)]
		set;
	}

	[field: AccessedThroughProperty("ut_grid")]
	internal virtual Grid ut_grid
	{
		get; [MethodImpl(MethodImplOptions.Synchronized)]
		set;
	}

	[field: AccessedThroughProperty("schs")]
	internal virtual GroupBox schs
	{
		get; [MethodImpl(MethodImplOptions.Synchronized)]
		set;
	}

	[field: AccessedThroughProperty("sch_grid")]
	internal virtual Grid sch_grid
	{
		get; [MethodImpl(MethodImplOptions.Synchronized)]
		set;
	}

	[field: AccessedThroughProperty("task")]
	internal virtual TextBox task
	{
		get; [MethodImpl(MethodImplOptions.Synchronized)]
		set;
	}

	[field: AccessedThroughProperty("tvz")]
	internal virtual TextBox tvz
	{
		get; [MethodImpl(MethodImplOptions.Synchronized)]
		set;
	}

	[field: AccessedThroughProperty("grf_save")]
	internal virtual System.Windows.Controls.CheckBox grf_save
	{
		get; [MethodImpl(MethodImplOptions.Synchronized)]
		set;
	}

	[field: AccessedThroughProperty("rst_save")]
	internal virtual System.Windows.Controls.CheckBox rst_save
	{
		get; [MethodImpl(MethodImplOptions.Synchronized)]
		set;
	}

	[field: AccessedThroughProperty("csv_save")]
	internal virtual System.Windows.Controls.CheckBox csv_save
	{
		get; [MethodImpl(MethodImplOptions.Synchronized)]
		set;
	}

	[field: AccessedThroughProperty("start")]
	internal virtual Button start
	{
		get; [MethodImpl(MethodImplOptions.Synchronized)]
		set;
	}

	[field: AccessedThroughProperty("prBar")]
	internal virtual ProgressBar prBar
	{
		get; [MethodImpl(MethodImplOptions.Synchronized)]
		set;
	}

	public MainWindow()
	{
		fInf = new compModel();
		InitializeComponent();
	}

	private void files_drop(object sender, DragEventArgs e)
	{
		if (e.Data.GetDataPresent(DataFormats.FileDrop))
		{
			fInf.files = (string[])e.Data.GetData(DataFormats.FileDrop);
			rgms_add_to_main_win(fInf.regims);
			ut2_add_to_main_win(fInf.ut2s);
			sch_add_to_main_win(fInf.schs);
			tvz.Text = Path.GetFileNameWithoutExtension(fInf.scns);
			task.Text = Path.GetFileNameWithoutExtension(fInf.xlxs.name);
		}
	}

	private void files_PreviewDragOver(object sender, DragEventArgs e)
	{
		if (e.Data.GetDataPresent(DataFormats.FileDrop))
		{
			e.Effects = (DragDropEffects)1;
			((RoutedEventArgs)e).Handled = true;
		}
		else
		{
			e.Effects = (DragDropEffects)0;
			((RoutedEventArgs)e).Handled = false;
		}
	}

	/// <summary>Обход лицензии: проверка license.lic отключена.</summary>
	private void license_check()
	{
		// Bypass: оригинальная проверка license.lic отключена.
	}

	private void Window_Loaded(object sender, RoutedEventArgs e)
	{
		//IL_00ec: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f2: Invalid comparison between Unknown and I4
		((FrameworkElement)this).DataContext = fInf;
		if (!(!File.Exists($"{FileSystem.CurDir()}\\EPPlus.dll") & !File.Exists($"{FileSystem.CurDir()}\\OxyPlot.dll") & !File.Exists($"{FileSystem.CurDir()}\\OxyPlot.Wpf.dll")))
		{
			return;
		}
		try
		{
			File.WriteAllBytes($"{FileSystem.CurDir()}\\EPPlus.dll", arv_mu_computation.My.Resources.Resources.EPPlus);
			File.WriteAllBytes($"{FileSystem.CurDir()}\\OxyPlot.dll", arv_mu_computation.My.Resources.Resources.OxyPlot);
			File.WriteAllBytes($"{FileSystem.CurDir()}\\OxyPlot.Wpf.dll", arv_mu_computation.My.Resources.Resources.OxyPlot_Wpf);
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			if ((int)Interaction.MsgBox((object)$"Не удалось записать необходимые библиотеки в каталог программы!{Environment.NewLine}Разместите файлы EPPlus.dll, OxyPlot.dll и OxyPlot.Wpf.dll в каталоге программы {FileSystem.CurDir()}\r\n{Environment.NewLine}Контактные данные для получения консультации:{Environment.NewLine}И.Г. Выборных (651) 26-53", (MsgBoxStyle)16, (object)"Ошибка") == 1)
			{
				Application.Current.Shutdown();
			}
			ProjectData.ClearProjectError();
		}
	}

	private void rgms_add_to_main_win(List<string> r)
	{
		//IL_0044: Unknown result type (might be due to invalid IL or missing references)
		//IL_0049: Unknown result type (might be due to invalid IL or missing references)
		//IL_0054: Unknown result type (might be due to invalid IL or missing references)
		//IL_0064: Expected O, but got Unknown
		//IL_0065: Unknown result type (might be due to invalid IL or missing references)
		//IL_006c: Expected O, but got Unknown
		//IL_008f: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ae: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c8: Unknown result type (might be due to invalid IL or missing references)
		rgms_grid.ColumnDefinitions.Clear();
		rgms_grid.RowDefinitions.Clear();
		int num = 0;
		foreach (string item in r)
		{
			rgms_grid.RowDefinitions.Add(new RowDefinition
			{
				Height = new GridLength(1.0, (GridUnitType)0)
			});
			TextBox val = new TextBox();
			val.Text = $"{Path.GetFileNameWithoutExtension(item)}";
			((Control)val).BorderThickness = new Thickness(0.0);
			((Control)val).HorizontalContentAlignment = (HorizontalAlignment)0;
			((FrameworkElement)val).Margin = new Thickness(2.0);
			((Control)val).Background = (Brush)(object)Brushes.LightGray;
			((Control)val).FontWeight = FontWeights.Medium;
			TextBox val2 = val;
			((Panel)rgms_grid).Children.Add((UIElement)(object)val2);
			Grid.SetRow((UIElement)(object)val2, num);
			num = checked(num + 1);
		}
	}

	private void ut2_add_to_main_win(List<string> r)
	{
		//IL_006d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0072: Unknown result type (might be due to invalid IL or missing references)
		//IL_007d: Unknown result type (might be due to invalid IL or missing references)
		//IL_008d: Expected O, but got Unknown
		//IL_008e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0095: Expected O, but got Unknown
		//IL_00b8: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00f1: Unknown result type (might be due to invalid IL or missing references)
		ut_grid.ColumnDefinitions.Clear();
		ut_grid.RowDefinitions.Clear();
		if (r.Count != 0)
		{
			Grid.SetColumnSpan((UIElement)(object)rgms, 1);
			((UIElement)uts).Visibility = (Visibility)0;
		}
		int num = 0;
		foreach (string item in r)
		{
			ut_grid.RowDefinitions.Add(new RowDefinition
			{
				Height = new GridLength(1.0, (GridUnitType)0)
			});
			TextBox val = new TextBox();
			val.Text = $"{Path.GetFileNameWithoutExtension(item)}";
			((Control)val).BorderThickness = new Thickness(0.0);
			((Control)val).HorizontalContentAlignment = (HorizontalAlignment)0;
			((FrameworkElement)val).Margin = new Thickness(2.0);
			((Control)val).Background = (Brush)(object)Brushes.LightGray;
			((Control)val).FontWeight = FontWeights.Medium;
			TextBox val2 = val;
			((Panel)ut_grid).Children.Add((UIElement)(object)val2);
			Grid.SetRow((UIElement)(object)val2, num);
			num = checked(num + 1);
		}
	}

	private void sch_add_to_main_win(string r)
	{
		//IL_005d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0062: Unknown result type (might be due to invalid IL or missing references)
		//IL_006d: Unknown result type (might be due to invalid IL or missing references)
		//IL_007d: Expected O, but got Unknown
		//IL_007e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0084: Expected O, but got Unknown
		//IL_00a5: Unknown result type (might be due to invalid IL or missing references)
		//IL_00c2: Unknown result type (might be due to invalid IL or missing references)
		//IL_00da: Unknown result type (might be due to invalid IL or missing references)
		sch_grid.ColumnDefinitions.Clear();
		sch_grid.RowDefinitions.Clear();
		if (Operators.CompareString(r, "", false) != 0)
		{
			Grid.SetColumnSpan((UIElement)(object)rgms, 1);
			((UIElement)schs).Visibility = (Visibility)0;
		}
		sch_grid.RowDefinitions.Add(new RowDefinition
		{
			Height = new GridLength(1.0, (GridUnitType)0)
		});
		TextBox val = new TextBox();
		val.Text = $"{Path.GetFileNameWithoutExtension(r)}";
		((Control)val).BorderThickness = new Thickness(0.0);
		((Control)val).HorizontalContentAlignment = (HorizontalAlignment)0;
		((FrameworkElement)val).Margin = new Thickness(2.0);
		((Control)val).Background = (Brush)(object)Brushes.LightGray;
		((Control)val).FontWeight = FontWeights.Medium;
		TextBox val2 = val;
		((Panel)sch_grid).Children.Add((UIElement)(object)val2);
		Grid.SetRow((UIElement)(object)val2, 0);
	}

	private void start_Click(object sender, RoutedEventArgs e)
	{
		//IL_18a6: Unknown result type (might be due to invalid IL or missing references)
		//IL_18ab: Unknown result type (might be due to invalid IL or missing references)
		//IL_18ad: Unknown result type (might be due to invalid IL or missing references)
		//IL_18b0: Invalid comparison between Unknown and I4
		//IL_02b3: Unknown result type (might be due to invalid IL or missing references)
		//IL_02b8: Unknown result type (might be due to invalid IL or missing references)
		//IL_02b9: Unknown result type (might be due to invalid IL or missing references)
		//IL_02bb: Invalid comparison between Unknown and I4
		//IL_04ad: Unknown result type (might be due to invalid IL or missing references)
		//IL_04b7: Expected O, but got Unknown
		//IL_16ad: Unknown result type (might be due to invalid IL or missing references)
		//IL_16b2: Unknown result type (might be due to invalid IL or missing references)
		license_check();
		((UIElement)start).IsEnabled = false;
		((FrameworkElement)this).Cursor = Cursors.Wait;
		foreach (rgmsInfo item in fInf.xlxs.rgms.Where([SpecialName] (rgmsInfo x) => x.sName != null))
		{
			Rastr rastr = (Rastr)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("EFC5E4AD-A3DD-11D3-B73F-00500454CF3F")));
			string text = $"{Environment.GetFolderPath(Environment.SpecialFolder.Personal)}\\RastrWin3\\SHABLON\\динамика.rst";
			string shabl = Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "\\RastrWin3\\SHABLON\\сценарий.scn";
			rastr.Load(RG_KOD.RG_REPL, fInf.scns, shabl);
			rastr.Load(RG_KOD.RG_REPL, item.sName, text);
			table table = (table)rastr.Tables.Item("com_dynamics");
			col col = (col)table.Cols.Item("MaxResultFiles");
			col col2 = (col)table.Cols.Item("SnapMaxCount");
			col col3 = (col)table.Cols.Item("SnapAutoLoad");
			col col4 = (col)table.Cols.Item("ResultFlowDirection");
			col col5 = (col)table.Cols.Item("Tras");
			col.Calc(Conversions.ToString(1));
			col2.Calc(Conversions.ToString(1));
			((ICol)col3).set_Z(0, (object)true);
			((ICol)col4).set_Z(0, (object)0);
			table table2 = (table)rastr.Tables.Item("DFWAutoActionScn");
			col col6 = (col)table2.Cols.Item("TimeStart");
			double num = Conversions.ToDouble(((ICol)col6).get_Z(0));
			if (num == 0.0)
			{
				table table3 = (table)rastr.Tables.Item("DFWAutoLogicScn");
				col col7 = (col)table3.Cols.Item("Delay");
				if (table3.Count != 0)
				{
					num = Conversions.ToDouble(((ICol)col7).get_Z(0));
				}
			}
			if (Operators.ConditionalCompareObjectLess(((ICol)col5).get_Z(0), (object)(num + 25.0), false))
			{
				((ICol)col5).set_Z(0, (object)(num + 25.0));
			}
			rastr.Save(item.sName, text);
		}
		MsgBoxResult val = Interaction.MsgBox((object)"Начать расчет?", (MsgBoxStyle)1, (object)"Начало расчета");
		if ((int)val == 1)
		{
			Microsoft.Office.Interop.Word.Application application = (Microsoft.Office.Interop.Word.Application)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("000209FF-0000-0000-C000-000000000046")));
			Documents documents = application.Documents;
			object Template = RuntimeHelpers.GetObjectValue(Missing.Value);
			object NewTemplate = RuntimeHelpers.GetObjectValue(Missing.Value);
			object DocumentType = RuntimeHelpers.GetObjectValue(Missing.Value);
			object Visible = RuntimeHelpers.GetObjectValue(Missing.Value);
			Document document = documents.Add(ref Template, ref NewTemplate, ref DocumentType, ref Visible);
			ExcelPackage excelPackage = new ExcelPackage();
			ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets.Add($"{Path.GetFileNameWithoutExtension(fInf.xlxs.name)}");
			try
			{
				((RangeBase)prBar).Minimum = 0.0;
				int num2 = default(int);
				checked
				{
					foreach (rgmsInfo item2 in fInf.xlxs.rgms.Where([SpecialName] (rgmsInfo x) => x.sName != null))
					{
						num2 += item2.rId.Where((_Closure_0024__._0024I10_002D2 == null) ? (_Closure_0024__._0024I10_002D2 = [SpecialName] (int x) => x != 0) : _Closure_0024__._0024I10_002D2).Count() + 1;
					}
					((RangeBase)prBar).Maximum = num2;
					((RangeBase)prBar).Value = 0.0;
					Rastr rastr2 = (Rastr)Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("EFC5E4AD-A3DD-11D3-B73F-00500454CF3F")));
					string text2 = $"{Environment.GetFolderPath(Environment.SpecialFolder.Personal)}\\RastrWin3\\SHABLON\\динамика.rst";
					excelWorksheet.Cells[$"A{1}:XFD{1048576}"].Style.Font.SetFromFont(new System.Drawing.Font("Times New Roman", 10f));
					excelWorksheet.Cells[$"A{1}:XFD{1048576}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
					excelWorksheet.Cells[$"A{1}:XFD{1048576}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
					excelWorksheet.Cells[$"A{1}:XFD{1048576}"].Style.WrapText = true;
					excelWorksheet.Cells[$"A{1}"].Value = "Схема сети";
					excelWorksheet.Column(1).Width = 50.0;
					excelWorksheet.Cells[$"A{1}:A{2}"].Merge = true;
					excelWorksheet.Cells[$"B{1}"].Value = "Регистрируемый параметр";
					excelWorksheet.Column(2).Width = 50.0;
					excelWorksheet.Cells[$"B{1}:B{2}"].Merge = true;
					excelWorksheet.Cells[$"C{1}"].Value = "Степень демпфирования";
					excelWorksheet.Column(3).Width = 15.0;
					excelWorksheet.Cells[$"C{1}:C{2}"].Merge = true;
					excelWorksheet.Cells[$"D{1}"].Value = "Результат проверки";
					excelWorksheet.Cells[$"D{2}"].Value = "Демпфирование переходного процесса";
					excelWorksheet.Cells[$"E{2}"].Value = "Эффективность PSS";
					excelWorksheet.Column(4).Width = 20.0;
					excelWorksheet.Column(5).Width = 15.0;
					excelWorksheet.Cells[$"D{1}:E{1}"].Merge = true;
					excelWorksheet.Cells[$"F{1}"].Value = "Примечание";
					excelWorksheet.Column(6).Width = 20.0;
					excelWorksheet.Cells[$"F{1}:F{2}"].Merge = true;
					excelWorksheet.Cells[$"A{1}:F{2}"].Style.Font.Bold = true;
					excelWorksheet.Cells[$"A{1}:F{2}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
					excelWorksheet.Cells[$"A{1}:F{2}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.WhiteSmoke);
					int num3 = 3;
					grfSize grfSize = new grfSize();
					updPrBar updPrBar = ((DependencyObject)prBar).SetValue;
					double num4 = 0.0;
					using (IEnumerator<rgmsInfo> enumerator3 = fInf.xlxs.rgms.Where([SpecialName] (rgmsInfo x) => x.sName != null).GetEnumerator())
					{
						_Closure_0024__10_002D0 closure_0024__10_002D = default(_Closure_0024__10_002D0);
						while (enumerator3.MoveNext())
						{
							closure_0024__10_002D = new _Closure_0024__10_002D0(closure_0024__10_002D);
							closure_0024__10_002D._0024VB_0024Local_rg = enumerator3.Current;
							string text3;
							if (!Directory.Exists($"{FileSystem.CurDir()}\\{closure_0024__10_002D._0024VB_0024Local_rg.name}"))
							{
								text3 = Directory.CreateDirectory($"{FileSystem.CurDir()}\\{closure_0024__10_002D._0024VB_0024Local_rg.name}").FullName;
							}
							else
							{
								text3 = $"{FileSystem.CurDir()}\\{closure_0024__10_002D._0024VB_0024Local_rg.name}";
								string[] files = Directory.GetFiles(text3);
								foreach (string path in files)
								{
									File.Delete(path);
								}
							}
							int num6 = 0;
							excelWorksheet.Cells[$"A{num3}"].Value = $"{closure_0024__10_002D._0024VB_0024Local_rg.name}";
							excelWorksheet.Cells[$"A{num3}:F{num3}"].Merge = true;
							excelWorksheet.Cells[$"A{num3}:F{num3}"].Style.Font.Italic = true;
							excelWorksheet.Cells[$"A{num3}:F{num3}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
							excelWorksheet.Cells[$"A{num3}:F{num3}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGoldenrodYellow);
							excelWorksheet.Cells[$"A{num3 + 1}"].Value = "Нормальная схема";
							rastr2.Load(RG_KOD.RG_REPL, closure_0024__10_002D._0024VB_0024Local_rg.sName, text2);
							if (rastr2.rgm("") == RastrRetCode.AST_OK)
							{
								grfSize = get_pss_quality(rastr2, excelWorksheet, num3 + 1, closure_0024__10_002D._0024VB_0024Local_rg.num, num6, text3, application, "Нормальная схема");
								excelWorksheet.Cells[$"A{num3 + 1}:A{num3 + grfSize.grfAm}"].Merge = true;
								num3 = grfSize.j;
								num4 += 1.0;
								num6++;
								((DispatcherObject)this).Dispatcher.Invoke((Delegate)updPrBar, (DispatcherPriority)4, new object[2]
								{
									RangeBase.ValueProperty,
									num4
								});
								foreach (remsInfo rem in fInf.xlxs.rems)
								{
									if (!closure_0024__10_002D._0024VB_0024Local_rg.rId.Contains(rem.id))
									{
										continue;
									}
									excelWorksheet.Cells[$"A{num3}"].Value = $"{rem.name}";
									rastr2.Load(RG_KOD.RG_REPL, closure_0024__10_002D._0024VB_0024Local_rg.sName, text2);
									table table4 = (table)rastr2.Tables.Item("vetv");
									col col8 = (col)table4.Cols.Item("sta");
									table table5 = (table)rastr2.Tables.Item("node");
									col col9 = (col)table5.Cols.Item("sta");
									foreach (elId element in rem.elements)
									{
										if (element.iq != 0)
										{
											table4.SetSel($"ip={element.ip}&iq={element.iq}&np={element.np}");
											col8.Calc(Conversions.ToString(1));
										}
										else
										{
											table5.SetSel($"ny={element.ip}");
											col9.Calc(Conversions.ToString(1));
										}
									}
									if ((rastr2.rgm("") == RastrRetCode.AST_OK) & (rem.schId != -1) & (rem.mpdValue != -1.0) & (rem.utId != -1))
									{
										get_mpd(rastr2, rem.schId, rem.mpdValue, rem.utId);
										grfSize = get_pss_quality(rastr2, excelWorksheet, num3, closure_0024__10_002D._0024VB_0024Local_rg.num, num6, text3, application, rem.name);
										excelWorksheet.Cells[$"A{num3}:A{num3 + grfSize.grfAm - 1}"].Merge = true;
										num3 = grfSize.j;
										num6++;
									}
									else if ((rastr2.rgm("") != RastrRetCode.AST_OK) & (rem.schId != -1) & (rem.mpdValue != -1.0) & (rem.utId != -1))
									{
										get_mpd(rastr2, rem.schId, rem.mpdValue, rem.utId);
										grfSize = get_pss_quality(rastr2, excelWorksheet, num3, closure_0024__10_002D._0024VB_0024Local_rg.num, num6, text3, application, rem.name);
										excelWorksheet.Cells[$"A{num3}:A{num3 + grfSize.grfAm - 1}"].Merge = true;
										num3 = grfSize.j;
										num6++;
									}
									else if ((rastr2.rgm("") == RastrRetCode.AST_OK) & ((rem.schId == -1) | (rem.mpdValue == -1.0) | (rem.utId == -1)))
									{
										grfSize = get_pss_quality(rastr2, excelWorksheet, num3, closure_0024__10_002D._0024VB_0024Local_rg.num, num6, text3, application, rem.name);
										excelWorksheet.Cells[$"A{num3}:A{num3 + grfSize.grfAm - 1}"].Merge = true;
										num3 = grfSize.j;
										num6++;
									}
									else if ((rastr2.rgm("") != RastrRetCode.AST_OK) & ((rem.schId == -1) | (rem.mpdValue == -1.0) | (rem.utId == -1)))
									{
										excelWorksheet.Cells[$"B{num3}"].Value = "Режим не балансируется в ремонтной схеме";
										excelWorksheet.Cells[$"B{num3}"].Style.Font.Bold = true;
										excelWorksheet.Cells[$"B{num3}"].Style.Font.Color.SetColor(System.Drawing.Color.Red);
										excelWorksheet.Cells[$"B{num3}:F{num3}"].Merge = true;
										excelWorksheet.Cells[$"A{num3}:F{num3}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
										excelWorksheet.Cells[$"A{num3}:F{num3}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
										num3++;
									}
									bool? isChecked = ((ToggleButton)rst_save).IsChecked;
									isChecked = isChecked;
									if (isChecked == true)
									{
										rastr2.Save($"{text3}\\{rem.name}.rst", text2);
									}
									num4 += 1.0;
									((DispatcherObject)this).Dispatcher.Invoke((Delegate)updPrBar, (DispatcherPriority)4, new object[2]
									{
										RangeBase.ValueProperty,
										num4
									});
								}
							}
							else
							{
								excelWorksheet.Cells[$"B{num3 + 1}"].Value = "Режим не балансируется в нормальной схеме";
								excelWorksheet.Cells[$"B{num3 + 1}"].Style.Font.Bold = true;
								excelWorksheet.Cells[$"B{num3 + 1}"].Style.Font.Color.SetColor(System.Drawing.Color.Red);
								excelWorksheet.Cells[$"B{num3 + 1}:F{num3 + 1}"].Merge = true;
								excelWorksheet.Cells[$"A{num3 + 1}:F{num3 + 1}"].Style.Fill.PatternType = ExcelFillStyle.Solid;
								excelWorksheet.Cells[$"A{num3 + 1}:F{num3 + 1}"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
								num4 += (double)(fInf.xlxs.rgms.Where(closure_0024__10_002D._Lambda_0024__4).FirstOrDefault().rId.Where([SpecialName] (int x) => x != 0).Count() + 1);
								((DispatcherObject)this).Dispatcher.Invoke((Delegate)updPrBar, (DispatcherPriority)4, new object[2]
								{
									RangeBase.ValueProperty,
									num4
								});
								num3 += 2;
							}
						}
					}
					excelWorksheet.Cells[$"A{1}:F{num3 - 1}"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
					excelWorksheet.Cells[$"A{1}:F{num3 - 1}"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
					excelWorksheet.Cells[$"A{1}:F{num3 - 1}"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
					excelWorksheet.Cells[$"A{1}:F{num3 - 1}"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
					excelPackage.SaveAs(new FileInfo($"{FileSystem.CurDir()}\\Приложение. Результат в табличном виде.xlsx"));
					application.Selection.WholeStory();
					application.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
					application.Selection.Font.Name = "Times New Roman";
					application.Selection.Font.Size = 12f;
					Visible = $"{FileSystem.CurDir()}\\Приложение. Результат в графическом виде.docx";
					DocumentType = RuntimeHelpers.GetObjectValue(Missing.Value);
					NewTemplate = RuntimeHelpers.GetObjectValue(Missing.Value);
					Template = RuntimeHelpers.GetObjectValue(Missing.Value);
					object AddToRecentFiles = RuntimeHelpers.GetObjectValue(Missing.Value);
					object WritePassword = RuntimeHelpers.GetObjectValue(Missing.Value);
					object ReadOnlyRecommended = RuntimeHelpers.GetObjectValue(Missing.Value);
					object EmbedTrueTypeFonts = RuntimeHelpers.GetObjectValue(Missing.Value);
					object SaveNativePictureFormat = RuntimeHelpers.GetObjectValue(Missing.Value);
					object SaveFormsData = RuntimeHelpers.GetObjectValue(Missing.Value);
					object SaveAsAOCELetter = RuntimeHelpers.GetObjectValue(Missing.Value);
					object Encoding = RuntimeHelpers.GetObjectValue(Missing.Value);
					object InsertLineBreaks = RuntimeHelpers.GetObjectValue(Missing.Value);
					object AllowSubstitutions = RuntimeHelpers.GetObjectValue(Missing.Value);
					object LineEnding = RuntimeHelpers.GetObjectValue(Missing.Value);
					object AddBiDiMarks = RuntimeHelpers.GetObjectValue(Missing.Value);
					object CompatibilityMode = RuntimeHelpers.GetObjectValue(Missing.Value);
					document.SaveAs2(ref Visible, ref DocumentType, ref NewTemplate, ref Template, ref AddToRecentFiles, ref WritePassword, ref ReadOnlyRecommended, ref EmbedTrueTypeFonts, ref SaveNativePictureFormat, ref SaveFormsData, ref SaveAsAOCELetter, ref Encoding, ref InsertLineBreaks, ref AllowSubstitutions, ref LineEnding, ref AddBiDiMarks, ref CompatibilityMode);
					CompatibilityMode = RuntimeHelpers.GetObjectValue(Missing.Value);
					AddBiDiMarks = RuntimeHelpers.GetObjectValue(Missing.Value);
					LineEnding = RuntimeHelpers.GetObjectValue(Missing.Value);
					document.Close(ref CompatibilityMode, ref AddBiDiMarks, ref LineEnding);
					LineEnding = RuntimeHelpers.GetObjectValue(Missing.Value);
					AddBiDiMarks = RuntimeHelpers.GetObjectValue(Missing.Value);
					CompatibilityMode = RuntimeHelpers.GetObjectValue(Missing.Value);
					application.Quit(ref LineEnding, ref AddBiDiMarks, ref CompatibilityMode);
					MsgBoxResult val2 = Interaction.MsgBox((object)$"Расчет завершен.{Environment.NewLine}Результаты сохранены в файлы:{Environment.NewLine}{FileSystem.CurDir()}\\Приложение. Результат в табличном виде.xlsx{Environment.NewLine}{FileSystem.CurDir()}\\Приложение. Результат в графическом виде.docx", (MsgBoxStyle)64, (object)"Завершение расчета");
				}
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				object CompatibilityMode = $"{FileSystem.CurDir()}\\Приложение. Результат в графическом виде.docx";
				object AddBiDiMarks = RuntimeHelpers.GetObjectValue(Missing.Value);
				object LineEnding = RuntimeHelpers.GetObjectValue(Missing.Value);
				object AllowSubstitutions = RuntimeHelpers.GetObjectValue(Missing.Value);
				object InsertLineBreaks = RuntimeHelpers.GetObjectValue(Missing.Value);
				object Encoding = RuntimeHelpers.GetObjectValue(Missing.Value);
				object SaveAsAOCELetter = RuntimeHelpers.GetObjectValue(Missing.Value);
				object SaveFormsData = RuntimeHelpers.GetObjectValue(Missing.Value);
				object SaveNativePictureFormat = RuntimeHelpers.GetObjectValue(Missing.Value);
				object EmbedTrueTypeFonts = RuntimeHelpers.GetObjectValue(Missing.Value);
				object ReadOnlyRecommended = RuntimeHelpers.GetObjectValue(Missing.Value);
				object WritePassword = RuntimeHelpers.GetObjectValue(Missing.Value);
				object AddToRecentFiles = RuntimeHelpers.GetObjectValue(Missing.Value);
				Template = RuntimeHelpers.GetObjectValue(Missing.Value);
				NewTemplate = RuntimeHelpers.GetObjectValue(Missing.Value);
				DocumentType = RuntimeHelpers.GetObjectValue(Missing.Value);
				Visible = RuntimeHelpers.GetObjectValue(Missing.Value);
				document.SaveAs2(ref CompatibilityMode, ref AddBiDiMarks, ref LineEnding, ref AllowSubstitutions, ref InsertLineBreaks, ref Encoding, ref SaveAsAOCELetter, ref SaveFormsData, ref SaveNativePictureFormat, ref EmbedTrueTypeFonts, ref ReadOnlyRecommended, ref WritePassword, ref AddToRecentFiles, ref Template, ref NewTemplate, ref DocumentType, ref Visible);
				application.Selection.WholeStory();
				application.Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
				application.Selection.Font.Name = "Times New Roman";
				application.Selection.Font.Size = 12f;
				Visible = RuntimeHelpers.GetObjectValue(Missing.Value);
				DocumentType = RuntimeHelpers.GetObjectValue(Missing.Value);
				NewTemplate = RuntimeHelpers.GetObjectValue(Missing.Value);
				document.Close(ref Visible, ref DocumentType, ref NewTemplate);
				NewTemplate = RuntimeHelpers.GetObjectValue(Missing.Value);
				DocumentType = RuntimeHelpers.GetObjectValue(Missing.Value);
				Visible = RuntimeHelpers.GetObjectValue(Missing.Value);
				application.Quit(ref NewTemplate, ref DocumentType, ref Visible);
				excelPackage.SaveAs(new FileInfo($"{FileSystem.CurDir()}\\Приложение. Результат в табличном виде.xlsx"));
				MsgBoxResult val3 = Interaction.MsgBox((object)$"Ошибка работы алгоритма Модуля{Environment.NewLine}Для получения консультации обратитесь:{Environment.NewLine}И.Г. Выборных (651) 26-53", (MsgBoxStyle)16, (object)"Ошибка");
				if (((0 - (((int)val3 == 1) ? 1 : 0)) | 2) != 0)
				{
					Application.Current.Shutdown();
				}
				ProjectData.ClearProjectError();
			}
		}
		((FrameworkElement)this).Cursor = Cursors.Arrow;
		((UIElement)start).IsEnabled = true;
	}

	private grfSize get_pss_quality(Rastr r, ExcelWorksheet x, int i, int rgNum, int remNum, string path, Microsoft.Office.Interop.Word.Application doc, string xname)
	{
		//IL_0a01: Unknown result type (might be due to invalid IL or missing references)
		//IL_0a0b: Expected O, but got Unknown
		_Closure_0024__11_002D1 closure_0024__11_002D = null;
		closure_0024__11_002D = new _Closure_0024__11_002D1(closure_0024__11_002D);
		grfSize grfSize = new grfSize();
		int num = 1;
		table table = (table)r.Tables.Item("com_dynamics");
		col col = (col)table.Cols.Item("Tras");
		List<grfData> list = new List<grfData>();
		list = DynCalc(r);
		grfSize.grfAm = list.Count;
		closure_0024__11_002D._0024VB_0024Local_sPoint = list.FirstOrDefault().sPoint;
		checked
		{
			bool? isChecked;
			using (List<grfInfo>.Enumerator enumerator = fInf.xlxs.grf.GetEnumerator())
			{
				_Closure_0024__11_002D0 closure_0024__11_002D2 = default(_Closure_0024__11_002D0);
				_Closure_0024__11_002D2 closure_0024__11_002D3 = default(_Closure_0024__11_002D2);
				while (enumerator.MoveNext())
				{
					closure_0024__11_002D2 = new _Closure_0024__11_002D0(closure_0024__11_002D2);
					closure_0024__11_002D2._0024VB_0024NonLocal__0024VB_0024Closure_2 = closure_0024__11_002D;
					closure_0024__11_002D2._0024VB_0024Local_g = enumerator.Current;
					if (!list.Exists(closure_0024__11_002D2._Lambda_0024__0))
					{
						continue;
					}
					x.Cells[$"B{i}"].Value = closure_0024__11_002D2._0024VB_0024Local_g.name;
					Collection<datevalue> collection = new Collection<datevalue>();
					collection = list.Where(closure_0024__11_002D2._Lambda_0024__1).FirstOrDefault().pssOn;
					double num2 = default(double);
					num2 = collection.Where([SpecialName] (datevalue n) => n.t > closure_0024__11_002D2._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_sPoint + 15.0).Max([SpecialName] (datevalue n) => n.v) - collection.Where([SpecialName] (datevalue n) => n.t > closure_0024__11_002D2._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_sPoint + 15.0).Min([SpecialName] (datevalue n) => n.v);
					double num3 = default(double);
					double num4 = default(double);
					num4 = collection.Where([SpecialName] (datevalue n) => (n.t < closure_0024__11_002D2._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_sPoint) & (n.t > closure_0024__11_002D2._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_sPoint - 2.0)).Average([SpecialName] (datevalue n) => n.v);
					num3 = collection.Where([SpecialName] (datevalue n) => n.t > closure_0024__11_002D2._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_sPoint).Max([SpecialName] (datevalue n) => n.v) - num4;
					x.Cells[$"C{i}"].Value = num2 / num3;
					x.Cells[$"C{i}"].Style.Numberformat.Format = "0.00000";
					if (num2 / num3 > 0.01)
					{
						closure_0024__11_002D3 = new _Closure_0024__11_002D2(closure_0024__11_002D3);
						closure_0024__11_002D3._0024VB_0024NonLocal__0024VB_0024Closure_3 = closure_0024__11_002D2;
						closure_0024__11_002D3._0024VB_0024Local_Max1 = collection.Where([SpecialName] (datevalue n) => n.t > closure_0024__11_002D3._0024VB_0024NonLocal__0024VB_0024Closure_3._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_sPoint + 15.0).Max([SpecialName] (datevalue n) => n.v);
						closure_0024__11_002D3._0024VB_0024Local_tMax1 = collection.Where(closure_0024__11_002D3._Lambda_0024__12).FirstOrDefault().t;
						closure_0024__11_002D3._0024VB_0024Local_Min1 = collection.Where(closure_0024__11_002D3._Lambda_0024__13).Min([SpecialName] (datevalue n) => n.v);
						closure_0024__11_002D3._0024VB_0024Local_tMin1 = collection.Where(closure_0024__11_002D3._Lambda_0024__15).FirstOrDefault().t;
						closure_0024__11_002D3._0024VB_0024Local_Max2 = collection.Where(closure_0024__11_002D3._Lambda_0024__16).Max([SpecialName] (datevalue n) => n.v);
						double t = collection.Where(closure_0024__11_002D3._Lambda_0024__18).FirstOrDefault().t;
						if ((1.0 / (closure_0024__11_002D3._0024VB_0024Local_tMax1 - t) >= 0.6) & (1.0 / (closure_0024__11_002D3._0024VB_0024Local_tMax1 - t) <= 2.0))
						{
							x.Cells[$"D{i}"].Value = "Не демпфируется";
							x.Cells[$"D{i}"].Style.Font.Bold = true;
							x.Cells[$"D{i}"].Style.Font.Color.SetColor(System.Drawing.Color.Red);
						}
						else
						{
							x.Cells[$"D{i}"].Value = "Демпфируется*";
						}
					}
					else
					{
						x.Cells[$"D{i}"].Value = "Демпфируется";
					}
					Collection<datevalue> collection2 = new Collection<datevalue>();
					collection2 = list.Where(closure_0024__11_002D2._Lambda_0024__19).FirstOrDefault().pssOff;
					double num5 = default(double);
					double num6 = default(double);
					foreach (datevalue item in collection)
					{
						if ((item.t > closure_0024__11_002D2._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_sPoint + 0.06) & (item.t < closure_0024__11_002D2._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_sPoint + 25.0))
						{
							num5 += Math.Pow(item.v - num4, 2.0) * item.t;
						}
					}
					foreach (datevalue item2 in collection2)
					{
						if ((item2.t > closure_0024__11_002D2._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_sPoint + 0.06) & (item2.t < closure_0024__11_002D2._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_sPoint + 25.0))
						{
							num6 += Math.Pow(item2.v - num4, 2.0) * item2.t;
						}
					}
					x.Cells[$"E{i}"].Value = ((num5 < num6) ? "Эффективен" : "Не эффективен");
					if (num5 >= num6)
					{
						x.Cells[$"E{i}"].Style.Font.Bold = true;
						x.Cells[$"E{i}"].Style.Font.Color.SetColor(System.Drawing.Color.Red);
					}
					isChecked = ((ToggleButton)grf_save).IsChecked;
					isChecked = isChecked;
					if (isChecked == true)
					{
						x.Cells[$"F{i}"].Value = $"Рисунок {rgNum}.{remNum}.{num}";
						get_grf(collection, collection2, closure_0024__11_002D2._0024VB_0024Local_g.name, closure_0024__11_002D2._0024VB_0024NonLocal__0024VB_0024Closure_2._0024VB_0024Local_sPoint, Conversions.ToDouble(((ICol)col).get_Z(0)), num4, $"Рисунок {rgNum}.{remNum}.{num}", path, doc);
					}
					i++;
					num++;
				}
			}
			isChecked = ((ToggleButton)csv_save).IsChecked;
			isChecked = isChecked;
			if (isChecked == true)
			{
				ExcelPackage excelPackage = new ExcelPackage();
				foreach (grfData item3 in list)
				{
					ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets.Add($"{item3.name}");
					excelWorksheet.Cells[$"A{1}:XFD{1048576}"].Style.Font.SetFromFont(new System.Drawing.Font("Times New Roman", 10f));
					excelWorksheet.Cells[$"A{1}:XFD{1048576}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
					excelWorksheet.Cells[$"A{1}:XFD{1048576}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
					excelWorksheet.Cells[$"A{1}:XFD{1048576}"].Style.WrapText = false;
					int num7 = 2;
					int num8 = 2;
					excelWorksheet.Cells["A1"].Value = "PSS введен";
					excelWorksheet.Cells["A1:B1"].Merge = true;
					excelWorksheet.Cells["C1"].Value = "PSS выведен";
					excelWorksheet.Cells["C1:D1"].Merge = true;
					foreach (datevalue item4 in item3.pssOn)
					{
						excelWorksheet.Cells[$"A{num7}"].Value = item4.t;
						excelWorksheet.Cells[$"B{num7}"].Value = item4.v;
						num7++;
					}
					foreach (datevalue item5 in item3.pssOff)
					{
						excelWorksheet.Cells[$"C{num8}"].Value = item5.t;
						excelWorksheet.Cells[$"D{num8}"].Value = item5.v;
						num8++;
					}
				}
				excelPackage.SaveAs(new FileInfo($"{path}\\{xname}.xlsx"));
			}
			grfSize.j = i;
			return grfSize;
		}
	}

	private List<grfData> DynCalc(Rastr r)
	{
		List<grfData> list = new List<grfData>();
		string shabl = Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "\\RastrWin3\\SHABLON\\автоматика.dfw";
		string shabl2 = Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "\\RastrWin3\\SHABLON\\сценарий.scn";
		r.NewFile(shabl);
		r.Load(RG_KOD.RG_REPL, fInf.scns, shabl2);
		FWDynamic fWDynamic = r.FWDynamic();
		fWDynamic.Run();
		table table = (table)r.Tables.Item("DFWAutoActionScn");
		col col = (col)table.Cols.Item("TimeStart");
		double num = Conversions.ToDouble(((ICol)col).get_Z(0));
		if (num == 0.0)
		{
			table table2 = (table)r.Tables.Item("DFWAutoLogicScn");
			col col2 = (col)table2.Cols.Item("Delay");
			if (table2.Count != 0)
			{
				num = Conversions.ToDouble(((ICol)col2).get_Z(0));
			}
		}
		object obj = default(object);
		object obj2 = default(object);
		object obj3 = default(object);
		object obj4 = default(object);
		foreach (grfInfo item in fInf.xlxs.grf)
		{
			Collection<datevalue> collection = new Collection<datevalue>();
			object objectValue = RuntimeHelpers.GetObjectValue(new object());
			int num2 = 1;
			table table3 = (table)r.Tables.Item($"{item.table}");
			col col3 = (col)table3.Cols.Item($"{item.param}");
			col col4 = (col)table3.Cols.Item("sta");
			foreach (int item2 in item.id)
			{
				if (Conversions.ToBoolean(Operators.AndObject((object)(Operators.CompareString(item.table, "vetv", false) == 0), Operators.CompareObjectLess(((ICol)col3).get_Z(item2), (object)0, false))))
				{
					num2 = -1;
				}
				if (!Operators.ConditionalCompareObjectEqual(((ICol)col4).get_Z(item2), (object)false, false))
				{
					continue;
				}
				objectValue = RuntimeHelpers.GetObjectValue(r.GetChainedGraphSnapshot($"{item.table}", $"{item.param}", item2, 0));
				if (collection.Count == 0)
				{
					if (ForLoopControl.ForLoopInitObj(obj, (object)0, NewLateBinding.LateGet(objectValue, (Type)null, "getUpperBound", new object[1] { 0 }, (string[])null, (Type[])null, (bool[])null), (object)1, ref obj2, ref obj))
					{
						do
						{
							collection.Add(new datevalue
							{
								t = Conversions.ToDouble(NewLateBinding.LateIndexGet(objectValue, new object[2] { obj, 1 }, (string[])null)),
								v = Conversions.ToDouble(Operators.MultiplyObject(NewLateBinding.LateIndexGet(objectValue, new object[2] { obj, 0 }, (string[])null), (object)num2))
							});
						}
						while (ForLoopControl.ForNextCheckObj(obj, obj2, ref obj));
					}
				}
				else if (ForLoopControl.ForLoopInitObj(obj3, (object)0, NewLateBinding.LateGet(objectValue, (Type)null, "getUpperBound", new object[1] { 0 }, (string[])null, (Type[])null, (bool[])null), (object)1, ref obj4, ref obj3))
				{
					do
					{
						collection[Conversions.ToInteger(obj3)].v = Conversions.ToDouble(Operators.AddObject((object)collection[Conversions.ToInteger(obj3)].v, Operators.MultiplyObject(NewLateBinding.LateIndexGet(objectValue, new object[2] { obj3, 0 }, (string[])null), (object)num2)));
					}
					while (ForLoopControl.ForNextCheckObj(obj3, obj4, ref obj3));
				}
			}
			if (collection.Count != 0)
			{
				list.Add(new grfData
				{
					name = item.name,
					pssOn = collection,
					sPoint = num
				});
			}
		}
		foreach (pssInfo item3 in fInf.xlxs.arv)
		{
			table table4 = (table)r.Tables.Item($"{item3.table}");
			col col5 = (col)table4.Cols.Item($"{item3.param}");
			((ICol)col5).set_Z(item3.id, (object)item3.val);
		}
		r.Load(RG_KOD.RG_REPL, fInf.scns, shabl2);
		fWDynamic.Run();
		table = (table)r.Tables.Item("DFWAutoActionScn");
		col = (col)table.Cols.Item("TimeStart");
		using (List<grfInfo>.Enumerator enumerator4 = fInf.xlxs.grf.GetEnumerator())
		{
			_Closure_0024__12_002D0 closure_0024__12_002D = default(_Closure_0024__12_002D0);
			object obj5 = default(object);
			object obj6 = default(object);
			object obj7 = default(object);
			object obj8 = default(object);
			while (enumerator4.MoveNext())
			{
				closure_0024__12_002D = new _Closure_0024__12_002D0(closure_0024__12_002D);
				closure_0024__12_002D._0024VB_0024Local_row = enumerator4.Current;
				Collection<datevalue> collection2 = new Collection<datevalue>();
				object objectValue2 = RuntimeHelpers.GetObjectValue(new object());
				int num3 = 1;
				table table5 = (table)r.Tables.Item($"{closure_0024__12_002D._0024VB_0024Local_row.table}");
				col col6 = (col)table5.Cols.Item($"{closure_0024__12_002D._0024VB_0024Local_row.param}");
				col col7 = (col)table5.Cols.Item("sta");
				foreach (int item4 in closure_0024__12_002D._0024VB_0024Local_row.id)
				{
					if (Conversions.ToBoolean(Operators.AndObject((object)(Operators.CompareString(closure_0024__12_002D._0024VB_0024Local_row.table, "vetv", false) == 0), Operators.CompareObjectLess(((ICol)col6).get_Z(item4), (object)0, false))))
					{
						num3 = -1;
					}
					if (!Operators.ConditionalCompareObjectEqual(((ICol)col7).get_Z(item4), (object)false, false))
					{
						continue;
					}
					objectValue2 = RuntimeHelpers.GetObjectValue(r.GetChainedGraphSnapshot($"{closure_0024__12_002D._0024VB_0024Local_row.table}", $"{closure_0024__12_002D._0024VB_0024Local_row.param}", item4, 0));
					if (collection2.Count == 0)
					{
						if (ForLoopControl.ForLoopInitObj(obj5, (object)0, NewLateBinding.LateGet(objectValue2, (Type)null, "getUpperBound", new object[1] { 0 }, (string[])null, (Type[])null, (bool[])null), (object)1, ref obj6, ref obj5))
						{
							do
							{
								collection2.Add(new datevalue
								{
									t = Conversions.ToDouble(NewLateBinding.LateIndexGet(objectValue2, new object[2] { obj5, 1 }, (string[])null)),
									v = Conversions.ToDouble(Operators.MultiplyObject(NewLateBinding.LateIndexGet(objectValue2, new object[2] { obj5, 0 }, (string[])null), (object)num3))
								});
							}
							while (ForLoopControl.ForNextCheckObj(obj5, obj6, ref obj5));
						}
					}
					else if (ForLoopControl.ForLoopInitObj(obj7, (object)0, NewLateBinding.LateGet(objectValue2, (Type)null, "getUpperBound", new object[1] { 0 }, (string[])null, (Type[])null, (bool[])null), (object)1, ref obj8, ref obj7))
					{
						do
						{
							collection2[Conversions.ToInteger(obj7)].v = Conversions.ToDouble(Operators.AddObject((object)collection2[Conversions.ToInteger(obj7)].v, Operators.MultiplyObject(NewLateBinding.LateIndexGet(objectValue2, new object[2] { obj7, 0 }, (string[])null), (object)num3)));
						}
						while (ForLoopControl.ForNextCheckObj(obj7, obj8, ref obj7));
					}
				}
				if (collection2.Count != 0)
				{
					list.Where(closure_0024__12_002D._Lambda_0024__0).FirstOrDefault().pssOff = collection2;
				}
			}
		}
		return list;
	}

	private void get_mpd(Rastr r, int s, double val, int ut)
	{
		_Closure_0024__13_002D0 arg = default(_Closure_0024__13_002D0);
		_Closure_0024__13_002D0 CS_0024_003C_003E8__locals3 = new _Closure_0024__13_002D0(arg);
		CS_0024_003C_003E8__locals3._0024VB_0024Local_ut = ut;
		string shabl = $"{Environment.GetFolderPath(Environment.SpecialFolder.Personal)}\\RastrWin3\\SHABLON\\сечения.sch";
		string shabl2 = $"{Environment.GetFolderPath(Environment.SpecialFolder.Personal)}\\RastrWin3\\SHABLON\\траектория утяжеления.ut2";
		int num = default(int);
		int num2 = default(int);
		r.Load(RG_KOD.RG_REPL, fInf.xlxs.ut.Where([SpecialName] (utInfo x) => x.id == CS_0024_003C_003E8__locals3._0024VB_0024Local_ut).FirstOrDefault().sName, shabl2);
		for (num = (int)r.rgm(""); num != 0; num = (int)r.step_ut("z"))
		{
		}
		r.rgm("");
		r.Load(RG_KOD.RG_REPL, fInf.schs, shabl);
		r.Load(RG_KOD.RG_REPL, fInf.xlxs.ut.Where([SpecialName] (utInfo x) => x.id == CS_0024_003C_003E8__locals3._0024VB_0024Local_ut).FirstOrDefault().sName, shabl2);
		table table = (table)r.Tables.Item("ut_common");
		col col = (col)table.Cols.Item("kfc");
		table table2 = (table)r.Tables.Item("sechen");
		col col2 = (col)table2.Cols.Item("psech");
		table2.SetSel($"ns={s}");
		num2 = ((ITable)table2).get_FindNextSel(-1);
		if (Operators.ConditionalCompareObjectGreater(((ICol)col2).get_Z(num2), (object)val, false))
		{
			double num3 = default(double);
			double num4 = default(double);
			double num5 = default(double);
			num4 = Conversions.ToDouble(((ICol)col2).get_Z(num2));
			num3 = num4;
			while (Math.Abs(val - num3) > 1.0)
			{
				num5 = Conversions.ToDouble(((ICol)col).get_Z(0));
				object obj = r.step_ut("z");
				r.rgm("");
				num3 = Conversions.ToDouble(((ICol)col2).get_Z(num2));
				((ICol)col).set_Z(0, (object)((num3 - val) / (num4 - num3)));
				num4 = num3;
			}
		}
	}

	private void get_grf(Collection<datevalue> psson, Collection<datevalue> pssoff, string name, double s, double e, double sPss, string pngName, string path, Microsoft.Office.Interop.Word.Application doc)
	{
		//IL_05ff: Unknown result type (might be due to invalid IL or missing references)
		//IL_0606: Expected O, but got Unknown
		//IL_0636: Unknown result type (might be due to invalid IL or missing references)
		//IL_0666: Unknown result type (might be due to invalid IL or missing references)
		//IL_06a9: Unknown result type (might be due to invalid IL or missing references)
		//IL_06ae: Unknown result type (might be due to invalid IL or missing references)
		//IL_06b5: Expected O, but got Unknown
		//IL_06bf: Unknown result type (might be due to invalid IL or missing references)
		//IL_06c6: Expected O, but got Unknown
		double num = default(double);
		double num2 = default(double);
		double num3 = Math.Min(sPss * 1.005, sPss + 0.5);
		double num4 = default(double);
		double num5 = Math.Min(sPss * 0.001, 0.5);
		num4 = psson.Where([SpecialName] (datevalue x) => Math.Abs(x.v - sPss) > num5).Max([SpecialName] (datevalue x) => x.t);
		num = psson.Where([SpecialName] (datevalue x) => (x.v < num3) & (x.t > s + 0.06)).FirstOrDefault().t;
		num2 = pssoff.Where([SpecialName] (datevalue x) => (x.v < num3) & (x.t > s + 0.06)).FirstOrDefault().t;
		double num6 = Math.Min(num, num2);
		double absoluteMinimum = Math.Min(psson.Where([SpecialName] (datevalue x) => x.t > num6).Min([SpecialName] (datevalue x) => x.v), pssoff.Where([SpecialName] (datevalue x) => x.t > num6).Min([SpecialName] (datevalue x) => x.v));
		double absoluteMaximum = Math.Max(psson.Where([SpecialName] (datevalue x) => x.t > num6).Max([SpecialName] (datevalue x) => x.v), pssoff.Where([SpecialName] (datevalue x) => x.t > num6).Max([SpecialName] (datevalue x) => x.v));
		PlotModel plotModel = new PlotModel
		{
			LegendPosition = LegendPosition.BottomCenter,
			LegendPlacement = LegendPlacement.Outside,
			LegendBackground = OxyColors.White,
			LegendBorder = OxyColors.Black,
			Title = name,
			TitleFontSize = 24.0,
			TitleHorizontalAlignment = TitleHorizontalAlignment.CenteredWithinPlotArea
		};
		PlotModel plotModel2 = new PlotModel
		{
			LegendPosition = LegendPosition.BottomCenter,
			LegendPlacement = LegendPlacement.Outside,
			LegendBackground = OxyColors.White,
			LegendBorder = OxyColors.Black
		};
		plotModel.Axes.Add(new OxyPlot.Axes.LinearAxis
		{
			Position = AxisPosition.Bottom,
			MajorGridlineStyle = LineStyle.Solid,
			AbsoluteMinimum = s - 1.0,
			AbsoluteMaximum = e,
			Title = "Время, с",
			AxisTitleDistance = 10.0,
			TitleFontSize = 18.0
		});
		plotModel2.Axes.Add(new OxyPlot.Axes.LinearAxis
		{
			Position = AxisPosition.Bottom,
			MajorGridlineStyle = LineStyle.Solid,
			AbsoluteMinimum = num6,
			AbsoluteMaximum = Math.Min(num4, Math.Min(s + 25.0, e)),
			Title = "Время, с",
			AxisTitleDistance = 10.0,
			TitleFontSize = 18.0
		});
		plotModel.Axes.Add(new OxyPlot.Axes.LinearAxis
		{
			Position = AxisPosition.Left,
			MajorGridlineStyle = LineStyle.Solid,
			Title = "Переток, МВт",
			AxisTitleDistance = 10.0,
			TitleFontSize = 18.0
		});
		plotModel2.Axes.Add(new OxyPlot.Axes.LinearAxis
		{
			Position = AxisPosition.Left,
			MajorGridlineStyle = LineStyle.Solid,
			Title = "Переток, МВт",
			AxisTitleDistance = 10.0,
			TitleFontSize = 18.0,
			AbsoluteMinimum = absoluteMinimum,
			AbsoluteMaximum = absoluteMaximum
		});
		plotModel.Series.Add(new OxyPlot.Series.LineSeries
		{
			StrokeThickness = 2.0,
			Color = OxyColors.Red,
			ItemsSource = psson,
			DataFieldX = "t",
			DataFieldY = "v"
		});
		plotModel.Series.Add(new OxyPlot.Series.LineSeries
		{
			StrokeThickness = 2.0,
			Color = OxyColors.Blue,
			ItemsSource = pssoff,
			DataFieldX = "t",
			DataFieldY = "v"
		});
		plotModel2.Series.Add(new OxyPlot.Series.LineSeries
		{
			StrokeThickness = 2.0,
			Color = OxyColors.Red,
			Title = "Каналы стабилизации введены [PSS введен]",
			FontSize = 18.0,
			ItemsSource = psson,
			DataFieldX = "t",
			DataFieldY = "v"
		});
		plotModel2.Series.Add(new OxyPlot.Series.LineSeries
		{
			StrokeThickness = 2.0,
			Color = OxyColors.Blue,
			Title = "Каналы стабилизации выведены [PSS выведен]",
			FontSize = 18.0,
			ItemsSource = pssoff,
			DataFieldX = "t",
			DataFieldY = "v"
		});
		PngExporter pngExporter = new PngExporter
		{
			Width = 1200,
			Height = 800,
			Background = OxyColors.White
		};
		BitmapSource val = pngExporter.ExportToBitmap(plotModel);
		BitmapSource val2 = pngExporter.ExportToBitmap(plotModel2);
		BitmapFrame val3 = BitmapFrame.Create(val);
		BitmapFrame val4 = BitmapFrame.Create(val2);
		DrawingVisual val5 = new DrawingVisual();
		DrawingContext val6 = val5.RenderOpen();
		try
		{
			val6.DrawImage((ImageSource)(object)val3, new Rect(0.0, 0.0, (double)((BitmapSource)val3).PixelWidth, (double)((BitmapSource)val3).PixelHeight));
			val6.DrawImage((ImageSource)(object)val4, new Rect(0.0, (double)((BitmapSource)val3).PixelHeight, (double)((BitmapSource)val4).PixelWidth, (double)((BitmapSource)val4).PixelHeight));
		}
		finally
		{
			((IDisposable)val6)?.Dispose();
		}
		RenderTargetBitmap val7 = new RenderTargetBitmap(((BitmapSource)val3).PixelWidth, checked(((BitmapSource)val3).PixelHeight + ((BitmapSource)val4).PixelHeight), 96.0, 96.0, PixelFormats.Pbgra32);
		val7.Render((Visual)(object)val5);
		PngBitmapEncoder val8 = new PngBitmapEncoder();
		((BitmapEncoder)val8).Frames.Add(BitmapFrame.Create((BitmapSource)(object)val7));
		using (Stream stream = File.Create($"{path}\\{pngName}.png"))
		{
			((BitmapEncoder)val8).Save(stream);
		}
		Clipboard.SetImage((BitmapSource)(object)val7);
		doc.Selection.Paste();
		doc.Selection.TypeParagraph();
		doc.Selection.TypeText($"{pngName}");
		doc.Selection.TypeParagraph();
	}

	private void rgms_MouseDoubleClick(object sender, MouseButtonEventArgs e)
	{
		fInf.regims = new List<string>();
		fInf.files = new string[0];
		((Panel)rgms_grid).Children.Clear();
		if ((fInf.ut2s.Count == 0) & (Operators.CompareString(fInf.schs, "", false) == 0))
		{
			Grid.SetColumnSpan((UIElement)(object)rgms, 2);
			((UIElement)uts).Visibility = (Visibility)1;
			((UIElement)schs).Visibility = (Visibility)1;
		}
	}

	private void uts_MouseDoubleClick(object sender, MouseButtonEventArgs e)
	{
		fInf.ut2s = new List<string>();
		fInf.files = new string[0];
		((Panel)ut_grid).Children.Clear();
		if ((fInf.ut2s.Count == 0) & (Operators.CompareString(fInf.schs, "", false) == 0))
		{
			Grid.SetColumnSpan((UIElement)(object)rgms, 2);
			((UIElement)uts).Visibility = (Visibility)1;
			((UIElement)schs).Visibility = (Visibility)1;
		}
	}

	private void schs_MouseDoubleClick(object sender, MouseButtonEventArgs e)
	{
		fInf.schs = "";
		fInf.files = new string[0];
		((Panel)sch_grid).Children.Clear();
		if ((fInf.ut2s.Count == 0) & (Operators.CompareString(fInf.schs, "", false) == 0))
		{
			Grid.SetColumnSpan((UIElement)(object)rgms, 2);
			((UIElement)uts).Visibility = (Visibility)1;
			((UIElement)schs).Visibility = (Visibility)1;
		}
	}

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	public void InitializeComponent()
	{
		if (!_contentLoaded)
		{
			_contentLoaded = true;
			Uri uri = new Uri("/Модуль оценки АРВ;component/mainwindow.baml", UriKind.Relative);
			System.Windows.Application.LoadComponent(this, uri);
		}
	}

	[DebuggerNonUserCode]
	[GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
	[EditorBrowsable(EditorBrowsableState.Never)]
	void IComponentConnector.Connect(int connectionId, object target)
	{
		//IL_0016: Unknown result type (might be due to invalid IL or missing references)
		//IL_0020: Expected O, but got Unknown
		//IL_002e: Unknown result type (might be due to invalid IL or missing references)
		//IL_0038: Expected O, but got Unknown
		//IL_0046: Unknown result type (might be due to invalid IL or missing references)
		//IL_0050: Expected O, but got Unknown
		//IL_0061: Unknown result type (might be due to invalid IL or missing references)
		//IL_006b: Expected O, but got Unknown
		//IL_0079: Unknown result type (might be due to invalid IL or missing references)
		//IL_0083: Expected O, but got Unknown
		//IL_0094: Unknown result type (might be due to invalid IL or missing references)
		//IL_009e: Expected O, but got Unknown
		//IL_00af: Unknown result type (might be due to invalid IL or missing references)
		//IL_00b9: Expected O, but got Unknown
		//IL_00c7: Unknown result type (might be due to invalid IL or missing references)
		//IL_00d1: Expected O, but got Unknown
		//IL_00e4: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ee: Expected O, but got Unknown
		//IL_0101: Unknown result type (might be due to invalid IL or missing references)
		//IL_010b: Expected O, but got Unknown
		//IL_0119: Unknown result type (might be due to invalid IL or missing references)
		//IL_0123: Expected O, but got Unknown
		//IL_0136: Unknown result type (might be due to invalid IL or missing references)
		//IL_0140: Expected O, but got Unknown
		//IL_0153: Unknown result type (might be due to invalid IL or missing references)
		//IL_015d: Expected O, but got Unknown
		//IL_0171: Unknown result type (might be due to invalid IL or missing references)
		//IL_017b: Expected O, but got Unknown
		//IL_018f: Unknown result type (might be due to invalid IL or missing references)
		//IL_0199: Expected O, but got Unknown
		//IL_01ad: Unknown result type (might be due to invalid IL or missing references)
		//IL_01b7: Expected O, but got Unknown
		//IL_01c8: Unknown result type (might be due to invalid IL or missing references)
		//IL_01d2: Expected O, but got Unknown
		//IL_01e3: Unknown result type (might be due to invalid IL or missing references)
		//IL_01ed: Expected O, but got Unknown
		//IL_01fb: Unknown result type (might be due to invalid IL or missing references)
		//IL_0205: Expected O, but got Unknown
		//IL_0216: Unknown result type (might be due to invalid IL or missing references)
		//IL_0220: Expected O, but got Unknown
		switch (connectionId)
		{
		case 1:
			((FrameworkElement)(MainWindow)target).Loaded += new RoutedEventHandler(Window_Loaded);
			((UIElement)(MainWindow)target).Drop += new DragEventHandler(files_drop);
			((UIElement)(MainWindow)target).PreviewDragOver += new DragEventHandler(files_PreviewDragOver);
			break;
		case 2:
			rgms = (GroupBox)target;
			((Control)rgms).MouseDoubleClick += new MouseButtonEventHandler(rgms_MouseDoubleClick);
			break;
		case 3:
			rgms_grid = (Grid)target;
			break;
		case 4:
			uts = (GroupBox)target;
			((Control)uts).MouseDoubleClick += new MouseButtonEventHandler(uts_MouseDoubleClick);
			break;
		case 5:
			ut_grid = (Grid)target;
			break;
		case 6:
			schs = (GroupBox)target;
			((Control)schs).MouseDoubleClick += new MouseButtonEventHandler(schs_MouseDoubleClick);
			break;
		case 7:
			sch_grid = (Grid)target;
			break;
		case 8:
			task = (TextBox)target;
			break;
		case 9:
			tvz = (TextBox)target;
			break;
		case 10:
			grf_save = (System.Windows.Controls.CheckBox)target;
			break;
		case 11:
			rst_save = (System.Windows.Controls.CheckBox)target;
			break;
		case 12:
			csv_save = (System.Windows.Controls.CheckBox)target;
			break;
		case 13:
			start = (Button)target;
			((ButtonBase)start).Click += new RoutedEventHandler(start_Click);
			break;
		case 14:
			prBar = (ProgressBar)target;
			break;
		default:
			_contentLoaded = true;
			break;
		}
	}
}
