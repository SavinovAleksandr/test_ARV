using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace ASTRALib;

[ComImport]
[CompilerGenerated]
[Guid("708E3897-3EF5-4BEC-B0C5-F5B13A8D448B")]
[TypeIdentifier]
public interface IFWDynamic
{
	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(1)]
	RastrRetCode Run();
}
