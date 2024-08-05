using System.Runtime.InteropServices;

namespace ShortcutTest
{
    internal class Program
    {
        public static string GetLnkTarget(string lnkPath)
        {
            if (!TryGetShellType(out var shellType))
                throw new InvalidOperationException();

            var shellComObject =
                Activator.CreateInstance(shellType) ?? throw new InvalidOperationException();

            var dirName = Path.GetDirectoryName(lnkPath) ?? throw new InvalidOperationException();

            var lnkFile = Path.GetFileName(lnkPath);

            var dirComObject =
                shellType.InvokeMember(
                    "NameSpace",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null,
                    shellComObject,
                    [dirName]
                ) ?? throw new Exception("e");

            var itemsComObject =
                dirComObject
                    .GetType()
                    .InvokeMember(
                        "Items",
                        System.Reflection.BindingFlags.InvokeMethod,
                        null,
                        dirComObject,
                        null
                    ) ?? throw new Exception("f");

            var itemComObject =
                itemsComObject
                    .GetType()
                    .InvokeMember(
                        "Item",
                        System.Reflection.BindingFlags.InvokeMethod,
                        null,
                        itemsComObject,
                        [lnkFile]
                    ) ?? throw new Exception("g");

            var linkComObject =
                itemComObject
                    .GetType()
                    .InvokeMember(
                        "GetLink",
                        System.Reflection.BindingFlags.GetProperty,
                        null,
                        itemComObject,
                        null
                    ) ?? throw new Exception("h");

            var targetPath =
                linkComObject
                    .GetType()
                    .InvokeMember(
                        "Path",
                        System.Reflection.BindingFlags.GetProperty,
                        null,
                        linkComObject,
                        null
                    ) ?? throw new Exception("i");

            return (string)targetPath;
            if (linkComObject != null)
                Marshal.ReleaseComObject(linkComObject);
            if (itemComObject != null)
                Marshal.ReleaseComObject(itemComObject);
            if (itemsComObject != null)
                Marshal.ReleaseComObject(itemsComObject);
            if (dirComObject != null)
                Marshal.ReleaseComObject(dirComObject);
            if (shellComObject != null)
                Marshal.ReleaseComObject(shellComObject);
        }

        public static bool TryGetShellType(out Type shellType)
        {
            shellType = typeof(object);

            if (!OperatingSystem.IsWindows())
                return false;

            if (Type.GetTypeFromProgID("Shell.Application") is not { } st)
                return false;

            shellType = st;
            return true;
        }

        public static bool TryInvokeMethod(
            Type invocationTargetType,
            string invocationName,
            object invocationTarget,
            object[] methodArgs,
            out object ret
        )
        {
            ret = new object();
            if (
                invocationTargetType.InvokeMember(
                    invocationName,
                    System.Reflection.BindingFlags.InvokeMethod,
                    null,
                    invocationTarget,
                    methodArgs
                )
                is not { } r
            )
                return false;
            ret = r;
            return true;
        }

        public static void Main(string[] args)
        {
            var lnkFile = @"C:\Users\tgudl\shortcutsTesting\07-20-24_a_sh.lnk";
            const string lnkFolder = @"C:\Users\tgudl\shortcutsTesting\Screenshots_sh.lnk";
            var tar = GetLnkTarget(lnkFile);
            var tar2 = GetLnkTarget(lnkFolder);
            Console.WriteLine(tar);
            Console.WriteLine(tar2);
        }
    }
}
