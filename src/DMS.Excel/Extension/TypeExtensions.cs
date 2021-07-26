using DMS.Common.Extensions;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace DMS.Excel.Extension
{
    /// <summary>
    /// 
    /// </summary>
    public static class TypeExtensions
    {
        /// <summary>
        ///获取显示名
        /// </summary>
        /// <param name="customAttributeProvider"></param>
        /// <param name="inherit"></param>
        /// <returns></returns>
        public static string GetDisplayName(this ICustomAttributeProvider customAttributeProvider, bool inherit = false)
        {
            var displayAttribute = customAttributeProvider.GetAttribute<DisplayAttribute>();
            string displayName;
            if (displayAttribute != null)
            {
                displayName = displayAttribute.Name;
            }
            else
            {
                displayName = customAttributeProvider.GetAttribute<DisplayNameAttribute>()?.DisplayName;
            }
            return displayName;
        }

        /// <summary>
        ///获取程序集属性
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="assembly"></param>
        /// <param name="inherit"></param>
        /// <returns></returns>
        public static T GetAttribute<T>(this ICustomAttributeProvider assembly, bool inherit = false)
            where T : Attribute
        {
            return assembly
                .GetCustomAttributes(typeof(T), inherit)
                .OfType<T>()
                .FirstOrDefault();
        }

        /// <summary>
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static string GetCSharpTypeName(this Type type)
        {
            var sb = new StringBuilder();
            var name = type.Name;
            if (!type.IsGenericType) return name;
            sb.Append(name.Substring(0, name.IndexOf('`')));
            sb.Append("<");
            sb.Append(string.Join(", ", type.GetGenericArguments()
                .Select(t => t.GetCSharpTypeName())));

            sb.Append(">");
            return sb.ToString();
        }

        /// <summary>
        ///是否必填
        /// </summary>
        /// <param name="propertyInfo"></param>
        /// <returns></returns>
        public static bool IsRequired(this PropertyInfo propertyInfo)
        {
            if (propertyInfo.GetAttribute<RequiredAttribute>(true) != null) return true;
            //Boolean、Byte、SByte、Int16、UInt16、Int32、UInt32、Int64、UInt64、Char、Double、Single
            if (propertyInfo.PropertyType.IsPrimitive) return true;
            switch (propertyInfo.PropertyType.Name)
            {
                case "DateTime":
                case "Decimal":
                    return true;
            }

            return false;
        }

        /// <summary>
        ///     是否为可为空类型
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static bool IsNullable(this Type type)
        {
            return Nullable.GetUnderlyingType(type) != null;
        }

        /// <summary>
        ///     获取可为空类型的底层类型
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static Type GetNullableUnderlyingType(this Type type)
        {
            return Nullable.GetUnderlyingType(type);
        }

        /// <summary>
        ///     获取Format
        /// </summary>
        /// <param name="customAttributeProvider"></param>
        /// <returns></returns>
        public static string GetDisplayFormat(this ICustomAttributeProvider customAttributeProvider)
        {
            var formatAttribute = customAttributeProvider.GetAttribute<DisplayFormatAttribute>();
            string displayFormat = string.Empty;
            if (formatAttribute != null)
            {
                displayFormat = formatAttribute.DataFormatString;
            }
            return displayFormat;
        }

        /// <summary>
        ///     获取类型描述
        /// </summary>
        /// <param name="customAttributeProvider"></param>
        /// <param name="inherit"></param>
        /// <returns></returns>
        public static string GetDescription(this ICustomAttributeProvider customAttributeProvider, bool inherit = false)
        {
            var des = string.Empty;
            var desAttribute = customAttributeProvider.GetAttribute<DescriptionAttribute>();
            if (desAttribute != null) des = desAttribute.Description;
            return des;
        }

        /// <summary>
        ///     获取枚举定义列表
        /// </summary>
        /// <returns>返回枚举列表元组（名称、值、显示名、描述）</returns>
        public static IEnumerable<Tuple<string, int, string, string>> GetEnumDefinitionList(this Type type)
        {
            var list = new List<Tuple<string, int, string, string>>();
            var attrType = type;
            if (!attrType.IsEnum) return null;
            var names = Enum.GetNames(attrType);
            var values = Enum.GetValues(attrType);
            var index = 0;
            foreach (var value in values)
            {
                var name = names[index];
                var field = value.GetType().GetField(value.ToString());
                var displayName = field.GetDisplayName();
                var des = field.GetDescription();
                var item = new Tuple<string, int, string, string>(
                    name,
                    Convert.ToInt32(value),
                    displayName.IsNullOrEmpty() ? null : displayName,
                    des.IsNullOrEmpty() ? null : des
                );
                list.Add(item);
                index++;
            }

            return list;
        }
    }
}
