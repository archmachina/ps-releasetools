<#
#>
Function Invoke-ImpersonateUser
{
    [CmdletBinding()]
    param(
        [Parameter(mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Username,

        [Parameter(mandatory=$true)]
        [ValidateNotNull()]
        [AllowEmptyString()]
        [string]$PasswordInPlaintext
    )

    process
    {

Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Principal;
using Microsoft.Win32.SafeHandles;
using System.Diagnostics;

namespace CustomProcess
{
    [StructLayout(LayoutKind.Sequential)]
    public struct PROFILEINFO {
        public int dwSize;
        public int dwFlags;
        [MarshalAs(UnmanagedType.LPTStr)]
        public String lpUserName;
        [MarshalAs(UnmanagedType.LPTStr)]
        public String lpProfilePath;
        [MarshalAs(UnmanagedType.LPTStr)]
        public String lpDefaultPath;
        [MarshalAs(UnmanagedType.LPTStr)]
        public String lpServerName;
        [MarshalAs(UnmanagedType.LPTStr)]
        public String lpPolicyPath;
        public IntPtr hProfile;
    }

    public static class ProcessInvoker
    {
        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool LogonUser(String lpszUsername, String lpszDomain, String lpszPassword,
            int dwLogonType, int dwLogonProvider, out IntPtr phToken);

        [DllImport("userenv.dll", SetLastError=true, CharSet=CharSet.Auto)]
        public static extern bool LoadUserProfile(IntPtr hToken, ref PROFILEINFO lpProfileInfo);

        public static string Impersonate(string username, string password)
        {

            const int LOGON32_PROVIDER_DEFAULT = 0;
            const int LOGON32_LOGON_INTERACTIVE = 2;
            IntPtr token = IntPtr.Zero;

            bool returnValue = LogonUser(username, System.Environment.MachineName, password,
                LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT,
                out token);

            if (false == returnValue)
            {
                int ret = Marshal.GetLastWin32Error();
                Console.WriteLine("LogonUser failed with error code : {0}", ret);
                //throw new System.ComponentModel.Win32Exception(ret);
                return ("LogonUser Error: " + ret);
            }

            var profileInfo = new PROFILEINFO();
            profileInfo.lpUserName = username;
            profileInfo.dwSize = Marshal.SizeOf(profileInfo);

            returnValue = LoadUserProfile(token, ref profileInfo);

            if (returnValue == false)
            {
                int ret = Marshal.GetLastWin32Error();
                Console.WriteLine("LogonUser failed with error code : {0}", ret);
                //throw new System.ComponentModel.Win32Exception(ret);
                return ("LoadUserProfile Error: " + ret);
            }

            var identity = new WindowsIdentity(token);

            var context = identity.Impersonate();

            return String.Empty;
        }
    }
}
'@

        [CustomProcess.ProcessInvoker]::Impersonate($Username, $PasswordInPlaintext)
    }
}
