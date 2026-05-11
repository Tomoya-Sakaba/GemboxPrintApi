using System.Collections.Generic;
using System.IO;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace backend_print.Services
{
    /// <summary>
    /// 複数PDFをページ順に結合する（中間に別PDFを挟む用途）。
    /// </summary>
    public static class PdfMergeService
    {
        /// <summary>与えられた順に全ページを1本のPDFにまとめる。空の要素はスキップ。</summary>
        public static byte[] MergePdfs(IEnumerable<byte[]> pdfParts)
        {
            using (var output = new PdfDocument())
            {
                foreach (var part in pdfParts)
                {
                    if (part == null || part.Length == 0)
                        continue;

                    using (var inputStream = new MemoryStream(part, writable: false))
                    using (var input = PdfReader.Open(inputStream, PdfDocumentOpenMode.Import))
                    {
                        var count = input.PageCount;
                        for (var i = 0; i < count; i++)
                            output.AddPage(input.Pages[i]);
                    }
                }

                using (var outMs = new MemoryStream())
                {
                    output.Save(outMs, false);
                    return outMs.ToArray();
                }
            }
        }
    }
}
