#nullable disable
using System.Reflection;

// Заглушки для типов Microsoft.VisualBasic, отсутствующих в пакете для .NET Core/8.
// Сборка только в avr_mu_computation_net8.csproj (в net472 используются реальные типы из Microsoft.VisualBasic).

namespace Microsoft.VisualBasic.ApplicationServices
{
	public class AssemblyInfo
	{
		private readonly Assembly _assembly;
		public AssemblyInfo(Assembly assembly) { _assembly = assembly; }
	}
}

namespace Microsoft.VisualBasic.Devices
{
	public class Computer
	{
	}
}

namespace Microsoft.VisualBasic.ApplicationServices
{
	public class User
	{
	}
}

namespace Microsoft.VisualBasic.Logging
{
	public class Log
	{
	}
}

namespace Microsoft.VisualBasic
{
	public static class FileSystem
	{
		public static string CurDir() => System.Environment.CurrentDirectory;
	}
}
