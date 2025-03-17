Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

[ComImport]
[Guid("00020400-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]

public interface IDispatch {
    int GetTypeInfoCount(out int pctinfo);
    int GetTypeInfo(int iTInfo, int lcid, out IntPtr ppTInfo);
    int GetIDsOfNames(
        ref Guid riid,
        [MarshalAs(UnmanagedType.LPArray, ArraySubType=UnmanagedType.LPWStr)] string[] rgszNames,
        int cNames,
        int lcid,
        [MarshalAs(UnmanagedType.LPArray)] int[] rgDispId
    );
    int Invoke(
        int dispIdMember,
        ref Guid riid,
        int lcid,
        ushort wFlags,
//      ref System.Runtime.InteropServices.DISPPARAMS          pDispParams,
        ref System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams,
        out object pVarResult,
//      ref System.Runtime.InteropServices.EXCEPINFO           pExcepInfo,
        ref System.Runtime.InteropServices.ComTypes.EXCEPINFO  pExcepInfo,
        out IntPtr pArgErr
    );
}
"@

# Create an example COM object¿here, Excel.
$excel = New-Object -ComObject Excel.Application

# Cast the COM object to our newly defined IDispatch interface.
$idispatch = [IDispatch] $excel

# Prepare an array to store the resulting DISPIDs.
$dispIds = New-Object int[] 1

# Call GetIDsOfNames on property/method names we want to look up.
# For example, we look up the "Visible" property name (LCID = 1033 (0x0409) for US English).
$hr = $idispatch.GetIDsOfNames([ref][Guid]::Empty, @("Visible"), 1, 0x0409, $dispIds)

Write-Host "HRESULT from GetIDsOfNames: $hr"
Write-Host "DISPIDs returned: $($dispIds[0])"

# Clean up Excel COM object
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
