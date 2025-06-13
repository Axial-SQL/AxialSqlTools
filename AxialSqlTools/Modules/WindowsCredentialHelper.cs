using System;
using System.Runtime.InteropServices;
using System.Text;

public static class WindowsCredentialHelper
{
    // Save a secret token to Windows Credential Manager
    public static void SaveToken(string targetName, string userName, string token)
    {
        byte[] tokenBytes = Encoding.Unicode.GetBytes(token);

        IntPtr tokenPtr = Marshal.AllocHGlobal(tokenBytes.Length);
        Marshal.Copy(tokenBytes, 0, tokenPtr, tokenBytes.Length);

        var cred = new CREDENTIAL
        {
            TargetName = targetName,
            UserName = userName,
            CredentialBlob = tokenPtr,
            CredentialBlobSize = (uint)tokenBytes.Length,
            Type = CRED_TYPE.GENERIC,
            Persist = (uint)CRED_PERSIST.LOCAL_MACHINE,
            AttributeCount = 0,
            Attributes = IntPtr.Zero,
            Comment = null,
            TargetAlias = null
        };

        bool success = CredWrite(ref cred, 0);
        Marshal.FreeHGlobal(tokenPtr);

        if (!success)
        {
            int error = Marshal.GetLastWin32Error();
            throw new Exception($"CredWrite failed with error code {error}");
        }
    }

    // Load a saved token from Windows Credential Manager
    public static string LoadToken(string targetName)
    {
        if (!CredRead(targetName, CRED_TYPE.GENERIC, 0, out IntPtr credPtr))
        {
            int error = Marshal.GetLastWin32Error();
            throw new Exception($"CredRead failed with error code {error}");
        }

        var cred = (CREDENTIAL)Marshal.PtrToStructure(credPtr, typeof(CREDENTIAL));
        string token = Marshal.PtrToStringUni(cred.CredentialBlob, (int)cred.CredentialBlobSize / 2);

        CredFree(credPtr);
        return token;
    }

    // Delete a saved token from Windows Credential Manager
    public static void DeleteToken(string targetName)
    {
        if (!CredDelete(targetName, CRED_TYPE.GENERIC, 0))
        {
            int error = Marshal.GetLastWin32Error();
            throw new Exception($"CredDelete failed with error code {error}");
        }
    }

    // --- Native Win32 API ---

    [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool CredWrite([In] ref CREDENTIAL userCredential, [In] uint flags);

    [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool CredRead(string target, CRED_TYPE type, int reservedFlag, out IntPtr credentialPtr);

    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern bool CredDelete(string target, CRED_TYPE type, int flags);

    [DllImport("advapi32.dll", SetLastError = false)]
    private static extern void CredFree([In] IntPtr buffer);

    // Credential structure
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct CREDENTIAL
    {
        public uint Flags;
        public CRED_TYPE Type;
        public string TargetName;
        public string Comment;
        public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
        public uint CredentialBlobSize;
        public IntPtr CredentialBlob;
        public uint Persist;
        public uint AttributeCount;
        public IntPtr Attributes;
        public string TargetAlias;
        public string UserName;
    }

    private enum CRED_TYPE : uint
    {
        GENERIC = 1
    }

    private enum CRED_PERSIST : uint
    {
        SESSION = 1,
        LOCAL_MACHINE = 2,
        ENTERPRISE = 3
    }
}
