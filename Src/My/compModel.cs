using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using Microsoft.VisualBasic.CompilerServices;
using OfficeOpenXml;

namespace arv_mu_computation.My;

public class compModel : INotifyPropertyChanged
{
	[CompilerGenerated]
	internal sealed class _Closure_0024__11_002D0
	{
		public string _0024VB_0024Local_row;

		public _Closure_0024__11_002D0(_Closure_0024__11_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_row = arg0._0024VB_0024Local_row;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__0(rgmsInfo x)
		{
			return Operators.CompareString(x.name, Path.GetFileNameWithoutExtension(_0024VB_0024Local_row), false) == 0;
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__14_002D0
	{
		public string _0024VB_0024Local_row;

		public _Closure_0024__14_002D0(_Closure_0024__14_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_row = arg0._0024VB_0024Local_row;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__0(utInfo x)
		{
			return Operators.CompareString(x.name, Path.GetFileNameWithoutExtension(_0024VB_0024Local_row), false) == 0;
		}
	}

	[CompilerGenerated]
	internal sealed class _Closure_0024__23_002D0
	{
		public ExcelWorksheet _0024VB_0024Local_ws;

		public int _0024VB_0024Local_j;

		public Func<string, bool> _0024I0;

		public Func<string, bool> _0024I1;

		public _Closure_0024__23_002D0(_Closure_0024__23_002D0 arg0)
		{
			if (arg0 != null)
			{
				_0024VB_0024Local_ws = arg0._0024VB_0024Local_ws;
				_0024VB_0024Local_j = arg0._0024VB_0024Local_j;
			}
		}

		[SpecialName]
		internal bool _Lambda_0024__0(string x)
		{
			return Operators.ConditionalCompareObjectEqual((object)Path.GetFileNameWithoutExtension(x), _0024VB_0024Local_ws.Cells[_0024VB_0024Local_j, 2].Value, false);
		}

		[SpecialName]
		internal bool _Lambda_0024__1(string x)
		{
			return Operators.ConditionalCompareObjectEqual((object)Path.GetFileNameWithoutExtension(x), _0024VB_0024Local_ws.Cells[_0024VB_0024Local_j, 2].Value, false);
		}
	}

	private string[] fls;

	private List<string> rgms;

	private List<string> ut2Files;

	private string schFile;

	private string scnsFile;

	private comInfo xlsFile;

	public string[] files
	{
		get
		{
			return fls;
		}
		set
		{
			fls = value;
			NotifyPropertyChanged("regims");
		}
	}

	public List<string> regims
	{
		get
		{
			string[] array = files;
			foreach (string text in array)
			{
				if ((Operators.CompareString(Path.GetExtension(text), ".rst", false) == 0) & !rgms.Contains(text))
				{
					rgms.Add(text);
				}
			}
			if (xlxs.rgms != null && xlxs.rgms.Count != 0)
			{
				using List<string>.Enumerator enumerator = rgms.GetEnumerator();
				_Closure_0024__11_002D0 closure_0024__11_002D = default(_Closure_0024__11_002D0);
				while (enumerator.MoveNext())
				{
					closure_0024__11_002D = new _Closure_0024__11_002D0(closure_0024__11_002D);
					closure_0024__11_002D._0024VB_0024Local_row = enumerator.Current;
					xlxs.rgms.Where(closure_0024__11_002D._Lambda_0024__0).FirstOrDefault().sName = closure_0024__11_002D._0024VB_0024Local_row;
				}
			}
			return rgms;
		}
		set
		{
			rgms.Clear();
		}
	}

	public List<string> ut2s
	{
		get
		{
			string[] array = files;
			foreach (string text in array)
			{
				if ((Operators.CompareString(Path.GetExtension(text), ".ut2", false) == 0) & !ut2Files.Contains(text))
				{
					ut2Files.Add(text);
				}
			}
			if (xlxs.ut != null && xlxs.ut.Count != 0)
			{
				using List<string>.Enumerator enumerator = ut2Files.GetEnumerator();
				_Closure_0024__14_002D0 closure_0024__14_002D = default(_Closure_0024__14_002D0);
				while (enumerator.MoveNext())
				{
					closure_0024__14_002D = new _Closure_0024__14_002D0(closure_0024__14_002D);
					closure_0024__14_002D._0024VB_0024Local_row = enumerator.Current;
					xlxs.ut.Where(closure_0024__14_002D._Lambda_0024__0).FirstOrDefault().sName = closure_0024__14_002D._0024VB_0024Local_row;
				}
			}
			return ut2Files;
		}
		set
		{
			ut2Files.Clear();
		}
	}

	public string schs
	{
		get
		{
			string[] array = files;
			foreach (string path in array)
			{
				if (Operators.CompareString(Path.GetExtension(path), ".sch", false) == 0)
				{
					schFile = path;
				}
			}
			return schFile;
		}
		set
		{
			schFile = "";
		}
	}

	public string scns
	{
		get
		{
			string[] array = files;
			foreach (string path in array)
			{
				if (Operators.CompareString(Path.GetExtension(path), ".scn", false) == 0)
				{
					scnsFile = path;
				}
			}
			return scnsFile;
		}
		set
		{
		}
	}

	public comInfo xlxs
	{
		get
		{
			string[] array = files;
			checked
			{
				_Closure_0024__23_002D0 closure_0024__23_002D = default(_Closure_0024__23_002D0);
				foreach (string text in array)
				{
					if (!((Operators.CompareString(Path.GetExtension(text), ".xlsx", false) == 0) | (Operators.CompareString(Path.GetExtension(text), ".xls", false) == 0)))
					{
						continue;
					}
					FileInfo newFile = new FileInfo(text);
					ExcelPackage excelPackage = new ExcelPackage(newFile);
					if (!((excelPackage.Workbook.Worksheets["rgms"] != null) & (excelPackage.Workbook.Worksheets["rems"] != null) & (excelPackage.Workbook.Worksheets["ut"] != null) & (excelPackage.Workbook.Worksheets["arv"] != null)))
					{
						continue;
					}
					closure_0024__23_002D = new _Closure_0024__23_002D0(closure_0024__23_002D);
					closure_0024__23_002D._0024VB_0024Local_ws = excelPackage.Workbook.Worksheets["rgms"];
					closure_0024__23_002D._0024VB_0024Local_j = 2;
					List<rgmsInfo> list = new List<rgmsInfo>();
					while (closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 1].Value != null)
					{
						List<int> list2 = new List<int>();
						if (closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 3].Value != null)
						{
							string text2 = Conversions.ToString(closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 3].Value);
							string[] array2 = text2.Split(new char[1] { ';' });
							foreach (string text3 in array2)
							{
								list2.Add(Conversions.ToInteger(text3));
							}
						}
						else
						{
							list2.Add(0);
						}
						list.Add(new rgmsInfo
						{
							num = Conversions.ToInteger(closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 1].Value),
							name = Conversions.ToString(closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 2].Value),
							rId = list2,
							sName = rgms.Where(closure_0024__23_002D._Lambda_0024__0).FirstOrDefault()
						});
						closure_0024__23_002D._0024VB_0024Local_j++;
					}
					closure_0024__23_002D._0024VB_0024Local_ws = excelPackage.Workbook.Worksheets["rems"];
					closure_0024__23_002D._0024VB_0024Local_j = 2;
					List<remsInfo> list3 = new List<remsInfo>();
					while (closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 3].Value != null)
					{
						if (closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 1].Value != null)
						{
							List<elId> list4 = new List<elId>();
							list4.Add(new elId
							{
								ip = Conversions.ToInteger(closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 3].Value),
								iq = Conversions.ToInteger((closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 4].Value != null) ? closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 4].Value : ((object)0)),
								np = Conversions.ToInteger((closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 5].Value != null) ? closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 5].Value : ((object)0))
							});
							for (int k = closure_0024__23_002D._0024VB_0024Local_j + 1; (closure_0024__23_002D._0024VB_0024Local_ws.Cells[k, 1].Value == null) & (closure_0024__23_002D._0024VB_0024Local_ws.Cells[k, 3].Value != null); k++)
							{
								list4.Add(new elId
								{
									ip = Conversions.ToInteger(closure_0024__23_002D._0024VB_0024Local_ws.Cells[k, 3].Value),
									iq = Conversions.ToInteger((closure_0024__23_002D._0024VB_0024Local_ws.Cells[k, 4].Value != null) ? closure_0024__23_002D._0024VB_0024Local_ws.Cells[k, 4].Value : ((object)0)),
									np = Conversions.ToInteger((closure_0024__23_002D._0024VB_0024Local_ws.Cells[k, 5].Value != null) ? closure_0024__23_002D._0024VB_0024Local_ws.Cells[k, 5].Value : ((object)0))
								});
							}
							list3.Add(new remsInfo
							{
								id = Conversions.ToInteger(closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 1].Value),
								name = Conversions.ToString(closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 2].Value),
								elements = list4,
								mpdValue = Conversions.ToDouble((closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 6].Value != null) ? closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 6].Value : ((object)(-1))),
								schId = Conversions.ToInteger((closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 7].Value != null) ? closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 7].Value : ((object)(-1))),
								utId = Conversions.ToInteger((closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 8].Value != null) ? closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 8].Value : ((object)(-1)))
							});
						}
						closure_0024__23_002D._0024VB_0024Local_j++;
					}
					closure_0024__23_002D._0024VB_0024Local_ws = excelPackage.Workbook.Worksheets["ut"];
					closure_0024__23_002D._0024VB_0024Local_j = 2;
					List<utInfo> list5 = new List<utInfo>();
					while (closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 1].Value != null)
					{
						list5.Add(new utInfo
						{
							id = Conversions.ToInteger(closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 1].Value),
							name = Conversions.ToString(closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 2].Value),
							sName = ut2Files.Where(closure_0024__23_002D._Lambda_0024__1).FirstOrDefault()
						});
						closure_0024__23_002D._0024VB_0024Local_j++;
					}
					closure_0024__23_002D._0024VB_0024Local_ws = excelPackage.Workbook.Worksheets["arv"];
					closure_0024__23_002D._0024VB_0024Local_j = 2;
					List<pssInfo> list6 = new List<pssInfo>();
					while (closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 1].Value != null)
					{
						list6.Add(new pssInfo
						{
							table = Conversions.ToString(closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 1].Value),
							param = Conversions.ToString(closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 2].Value),
							id = Conversions.ToInteger(closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 3].Value),
							val = Conversions.ToDouble(closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 4].Value)
						});
						closure_0024__23_002D._0024VB_0024Local_j++;
					}
					closure_0024__23_002D._0024VB_0024Local_ws = excelPackage.Workbook.Worksheets["grf"];
					closure_0024__23_002D._0024VB_0024Local_j = 2;
					List<grfInfo> list7 = new List<grfInfo>();
					while (closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 1].Value != null)
					{
						string text4 = Conversions.ToString(closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 3].Value);
						List<int> list8 = new List<int>();
						string[] array3 = text4.Split(new char[1] { '+' });
						foreach (string text5 in array3)
						{
							list8.Add(Conversions.ToInteger(text5));
						}
						list7.Add(new grfInfo
						{
							table = Conversions.ToString(closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 1].Value),
							param = Conversions.ToString(closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 2].Value),
							id = list8,
							name = Conversions.ToString(closure_0024__23_002D._0024VB_0024Local_ws.Cells[closure_0024__23_002D._0024VB_0024Local_j, 4].Value)
						});
						closure_0024__23_002D._0024VB_0024Local_j++;
					}
					xlsFile.rgms = list;
					xlsFile.rems = list3;
					xlsFile.ut = list5;
					xlsFile.arv = list6;
					xlsFile.grf = list7;
					xlsFile.name = Path.GetFileNameWithoutExtension(text);
				}
				return xlsFile;
			}
		}
		set
		{
		}
	}

	public event PropertyChangedEventHandler PropertyChanged;

	public compModel()
	{
		rgms = new List<string>();
		ut2Files = new List<string>();
		schFile = "";
		scnsFile = "";
		xlsFile = new comInfo();
	}

	private void NotifyPropertyChanged([CallerMemberName] string propertyName = null)
	{
		PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
	}
}
