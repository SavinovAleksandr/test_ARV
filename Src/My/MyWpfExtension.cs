using System;
using System.Collections;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Windows;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.ApplicationServices;
using Microsoft.VisualBasic.CompilerServices;
using Microsoft.VisualBasic.Devices;
using Microsoft.VisualBasic.Logging;

namespace arv_mu_computation.My;

[StandardModule]
[HideModuleName]
internal sealed class MyWpfExtension
{
	[EditorBrowsable(EditorBrowsableState.Never)]
	[MyGroupCollection("System.Windows.Window", "Create__Instance__", "Dispose__Instance__", "My.MyWpfExtenstionModule.Windows")]
	internal sealed class MyWindows
	{
		[ThreadStatic]
		private static Hashtable s_WindowBeingCreated;

		[EditorBrowsable(EditorBrowsableState.Never)]
		public MainWindow m_MainWindow;

		public MainWindow MainWindow
		{
			[DebuggerHidden]
			get
			{
				m_MainWindow = Create__Instance__(m_MainWindow);
				return m_MainWindow;
			}
			[DebuggerHidden]
			set
			{
				if (value != m_MainWindow)
				{
					if (value != null)
					{
						throw new ArgumentException("Property can only be set to Nothing");
					}
					Dispose__Instance__(ref m_MainWindow);
				}
			}
		}

		[DebuggerHidden]
		private static T Create__Instance__<T>(T Instance) where T : Window, new()
		{
			if (Instance == null)
			{
				if (s_WindowBeingCreated != null)
				{
					if (s_WindowBeingCreated.ContainsKey(typeof(T)))
					{
						throw new InvalidOperationException("The window cannot be accessed via My.Windows from the Window constructor.");
					}
				}
				else
				{
					s_WindowBeingCreated = new Hashtable();
				}
				s_WindowBeingCreated.Add(typeof(T), null);
				return new T();
			}
			return Instance;
		}

		[DebuggerHidden]
		private void Dispose__Instance__<T>(ref T instance) where T : Window
		{
			instance = default(T);
		}

		[DebuggerHidden]
		[EditorBrowsable(EditorBrowsableState.Never)]
		public MyWindows()
		{
		}

		[EditorBrowsable(EditorBrowsableState.Never)]
		public override bool Equals(object o)
		{
			return base.Equals(RuntimeHelpers.GetObjectValue(o));
		}

		[EditorBrowsable(EditorBrowsableState.Never)]
		public override int GetHashCode()
		{
			return base.GetHashCode();
		}

		[EditorBrowsable(EditorBrowsableState.Never)]
		internal new Type GetType()
		{
			return typeof(MyWindows);
		}

		[EditorBrowsable(EditorBrowsableState.Never)]
		public override string ToString()
		{
			return base.ToString();
		}
	}

	private static MyProject.ThreadSafeObjectProvider<Computer> s_Computer = new MyProject.ThreadSafeObjectProvider<Computer>();

	private static MyProject.ThreadSafeObjectProvider<User> s_User = new MyProject.ThreadSafeObjectProvider<User>();

	private static MyProject.ThreadSafeObjectProvider<MyWindows> s_Windows = new MyProject.ThreadSafeObjectProvider<MyWindows>();

	private static MyProject.ThreadSafeObjectProvider<Log> s_Log = new MyProject.ThreadSafeObjectProvider<Log>();

	internal static Application Application => (Application)(object)Application.Current;

	internal static Computer Computer => s_Computer.GetInstance;

	internal static User User => s_User.GetInstance;

	internal static Log Log => s_Log.GetInstance;

	internal static MyWindows Windows
	{
		[DebuggerHidden]
		get
		{
			return s_Windows.GetInstance;
		}
	}
}
