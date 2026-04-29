using System;
using System.IO;
using System.Web.Hosting;

namespace backend_print.Utils
{
    public static class PathResolveUtils
    {
        public static string ResolveTemplateBasePath(string configured)
        {
            // 1) 未設定なら既定（従来互換）
            if (string.IsNullOrWhiteSpace(configured))
                return @"C:\app_data\b-templates";

            // 2) 物理パスならそのまま- それ以外（仮想パス/相対パス）は HostingEnvironment.MapPath で物理化
            var baseDir = Path.IsPathRooted(configured)
                ? configured
                : (HostingEnvironment.MapPath(configured) ?? "");

            // パスの末尾の区切り文字(/)を削除
            baseDir = baseDir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            return baseDir;
        }
    }
}

