using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableDefinitionCreator.utils
{
    internal static class LinqExtensions
    {
        public static string StringJoin<T>(this IEnumerable<T> source, string separator)
        {
            return string.Join(separator, source);
        }

        public static string StringJoin<T>(this IEnumerable<T> source, string separator, Func<T, string> selector)
        {
            return string.Join(separator, source.Select(selector));
        }
    }
}
