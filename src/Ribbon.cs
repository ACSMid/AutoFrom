using System;
using System.Runtime.InteropServices;

namespace AutoFrom
{
    [ComVisible(true), Guid("000C0396-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility
    {
        [DispId(1)] [return: MarshalAs(UnmanagedType.BStr)] string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string ribbonId);
    }
    [ComVisible(true), Guid("34B48104-B89A-45C3-B5DD-C96447665E98"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IAutoFromRibbon
    {
        [DispId(1)] void OpenSettings([MarshalAs(UnmanagedType.IDispatch)] object control);
        [DispId(2)] void CheckDraft([MarshalAs(UnmanagedType.IDispatch)] object control);
    }
    public static class RibbonMarkup
    {
        public static string For(string ribbonId)
        {
            if (ribbonId != "Microsoft.Outlook.Explorer" && ribbonId != "Microsoft.Outlook.Mail.Compose") return null;
            return "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"><ribbon><tabs>" +
                "<tab id=\"AutoFromTab\" label=\"AutoFrom\"><group id=\"AutoFromRules\" label=\"Sender selection\">" +
                "<button id=\"AutoFromSettings\" label=\"Settings\" size=\"large\" imageMso=\"PropertySheet\" onAction=\"OpenSettings\" " +
                "screentip=\"Configure automatic senders\" supertip=\"Choose sending accounts, shared mailboxes and recipient rules. Test your rules before saving.\"/>" +
                "<button id=\"AutoFromCheck\" label=\"Check draft\" onAction=\"CheckDraft\" screentip=\"Explain sender selection for the open draft\"/>" +
                "</group></tab></tabs></ribbon></customUI>";
        }
    }
}
