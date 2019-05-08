using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace MercyShipsTimeEntry
{
    public static class EnumExtensions
    {

        public static List<T> ToList<T>(this Array arr)
        {
            var list = new List<T>();
            foreach (var item in arr)
            {
                list.Add((T)item);
            }
            return list;
        }

        public static string ToFriendlyString(this Enum code)
        {
            return Enum.GetName(code.GetType(), code);
        }
        /// <summary>
        /// Allows you to put Description decorators on your Enums
        /// [Description("Option One")]
        /// </summary>
        public static string ToDescriptionString(this Enum This)
        {
            Type type = This.GetType();

            string name = Enum.GetName(type, This);

            MemberInfo member = type.GetMembers()
                .Where(w => w.Name == name)
                .FirstOrDefault();

            DescriptionAttribute attribute = member != null
                ? member.GetCustomAttributes(true)
                    .Where(w => w.GetType() == typeof(DescriptionAttribute))
                    .FirstOrDefault() as DescriptionAttribute
                : null;

            return attribute != null ? attribute.Description : name;
        }
    }

    public static class DictExtensions
    {
        // Use sb.AppendFormat("{0}:{1} \n", nameof(Name), Name);

        public static IEnumerable<KeyValuePair<T1,T2>> ToList<T1,T2>(this Dictionary<T1,T2> dict)
        {
            var result = new List<KeyValuePair<T1, T2>>();
            foreach(var pair in dict)
            {
                result.Add(pair);
            }
            return result;
        }
    }
}
