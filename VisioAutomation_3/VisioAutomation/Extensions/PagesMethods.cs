using System.Collections.Generic;
using IVisio = Microsoft.Office.Interop.Visio;
using System.Linq;

namespace VisioAutomation.Extensions
{
    public static class PagesMethods
    {
        public static IEnumerable<IVisio.Page> AsEnumerable(this IVisio.Pages pages)
        {
            for (int i = 0; i < pages.Count; i++)
            {
                yield return pages[i + 1];
            }
        }

        public static string[] GetNamesU(this IVisio.Pages pages)
        {
            System.Array names_sa;
            pages.GetNamesU(out names_sa);
            string[] names = (string[]) names_sa;
            return names;
        }
    }
}