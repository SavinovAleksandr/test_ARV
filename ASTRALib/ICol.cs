using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace ASTRALib;

[ComImport]
[CompilerGenerated]
[Guid("7E534FE0-B0C6-11D5-A75D-005004526FE6")]
[TypeIdentifier]
public interface ICol
{
	void _VtblGap1_5();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(4)]
	[return: MarshalAs(UnmanagedType.Struct)]
	object get_Z(int index);

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(4)]
	void set_Z(int index, [MarshalAs(UnmanagedType.Struct)] object value);

	void _VtblGap2_6();

	[MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
	[DispId(7)]
	void Calc([MarshalAs(UnmanagedType.BStr)] string formula);
}
