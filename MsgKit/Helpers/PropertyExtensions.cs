using MsgKit.Enums;
using MsgKit.Structures;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MsgKit.Helpers
{
    public static class PropertyExtensions
    {
        /// <summary>
        ///     Get the PropertyId as reconised by Exchange REST Api
        /// </summary>
        public static string RestPropertyId(this NamedPropertyTag prop) =>
             PropertyIdTypeName(prop.Type) + " {" + prop.Guid.ToString() + "} Id " + prop.Id.ToString("X4");

        /// <summary>
        ///     Get the PropertyId as reconised by Exchange REST Api
        /// </summary>
        public static string RestPropertyId(this PropertyTag prop) =>
             PropertyIdTypeName(prop.Type) + prop.Id.ToString("X4");


        private static string PropertyIdTypeName(PropertyType type)
        {
            switch (type)
            {
                case PropertyType.PT_UNSPECIFIED:
                    break;
                case PropertyType.PT_NULL:
                    break;
                case PropertyType.PT_SHORT:
                    return "Integer";
                case PropertyType.PT_LONG:
                    return "Integer";
                case PropertyType.PT_FLOAT:
                    break;
                case PropertyType.PT_DOUBLE:
                    break;
                case PropertyType.PT_APPTIME:
                    break;
                case PropertyType.PT_ERROR:
                    break;
                case PropertyType.PT_BOOLEAN:
                    break;
                case PropertyType.PT_OBJECT:
                    break;
                case PropertyType.PT_LONGLONG:
                    break;
                case PropertyType.PT_UNICODE:
                    break;
                case PropertyType.PT_STRING8:
                    return "String";
                case PropertyType.PT_SYSTIME:
                    break;
                case PropertyType.PT_CLSID:
                    break;
                case PropertyType.PT_SVREID:
                    break;
                case PropertyType.PT_SRESTRICT:
                    break;
                case PropertyType.PT_ACTIONS:
                    break;
                case PropertyType.PT_BINARY:
                    break;
                case PropertyType.PT_MV_SHORT:
                    return "IntegerArray";
                case PropertyType.PT_MV_LONG:
                    return "IntegerArray";
                case PropertyType.PT_MV_FLOAT:
                    break;
                case PropertyType.PT_MV_DOUBLE:
                    break;
                case PropertyType.PT_MV_CURRENCY:
                    break;
                case PropertyType.PT_MV_APPTIME:
                    break;
                case PropertyType.PT_MV_LONGLONG:
                    break;
                case PropertyType.PT_MV_UNICODE:
                    break;
                case PropertyType.PT_MV_STRING8:
                    return "StringArray";
                case PropertyType.PT_MV_SYSTIME:
                    break;
                case PropertyType.PT_MV_CLSID:
                    break;
                case PropertyType.PT_MV_BINARY:
                    break;
                default:
                    break;
            }

            throw new System.Exception($"Do not know property type {type.ToString()}");
            
        }

        /// <summary>
        ///     Converts value to specified property type
        /// </summary>
        /// <param name="propertyType"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static object ParsePropertyType(this string value, PropertyType propertyType)
        {
            switch (propertyType)
            {
                case PropertyType.PT_UNSPECIFIED:
                    throw new NotSupportedException("PT_UNSPECIFIED property type is not supported");
                case PropertyType.PT_NULL:
                    return null;
                case PropertyType.PT_SHORT:
                    return Convert.ToInt16(value);
                case PropertyType.PT_LONG:
                case PropertyType.PT_ERROR:
                    return Convert.ToInt32(value);
                case PropertyType.PT_FLOAT:
                    return (float) Convert.ToDecimal(value);
                case PropertyType.PT_DOUBLE:
                    return Convert.ToDouble(value);
                case PropertyType.PT_APPTIME:
                case PropertyType.PT_SYSTIME:
                    return Convert.ToDateTime(value);
                case PropertyType.PT_BOOLEAN:
                    return Convert.ToBoolean(value);
                case PropertyType.PT_OBJECT:
                case PropertyType.PT_UNICODE:
                case PropertyType.PT_STRING8:
                    return value;
                case PropertyType.PT_I8:
                    return Convert.ToInt64(value);
                case PropertyType.PT_CLSID:
                    return new Guid(value);
                case PropertyType.PT_SVREID:
                    throw new NotSupportedException("PT_SVREID property type is not supported");
                case PropertyType.PT_SRESTRICT:
                    throw new NotSupportedException("PT_SRESTRICT property type is not supported");
                case PropertyType.PT_ACTIONS:
                    throw new NotSupportedException("PT_ACTIONS property type is not supported");
                case PropertyType.PT_BINARY:
                    return Encoding.UTF8.GetBytes(value);
                case PropertyType.PT_MV_SHORT:
                    throw new NotSupportedException("PT_MV_SHORT property type is not supported");
                case PropertyType.PT_MV_LONG:
                    throw new NotSupportedException("PT_MV_LONG property type is not supported");
                case PropertyType.PT_MV_FLOAT:
                    throw new NotSupportedException("PT_MV_FLOAT property type is not supported");
                case PropertyType.PT_MV_DOUBLE:
                    throw new NotSupportedException("PT_MV_DOUBLE property type is not supported");
                case PropertyType.PT_MV_CURRENCY:
                    throw new NotSupportedException("PT_MV_CURRENCY property type is not supported");
                case PropertyType.PT_MV_APPTIME:
                    throw new NotSupportedException("PT_MV_APPTIME property type is not supported");
                case PropertyType.PT_MV_LONGLONG:
                    throw new NotSupportedException("PT_MV_LONGLONG property type is not supported");
                case PropertyType.PT_MV_TSTRING:
                    throw new NotSupportedException("PT_MV_TSTRING property type is not supported");
                case PropertyType.PT_MV_STRING8:
                    throw new NotSupportedException("PT_MV_STRING8 property type is not supported");
                case PropertyType.PT_MV_SYSTIME:
                    throw new NotSupportedException("PT_MV_SYSTIME property type is not supported");
                case PropertyType.PT_MV_CLSID:
                    throw new NotSupportedException("PT_MV_CLSID property type is not supported");
                case PropertyType.PT_MV_BINARY:
                    throw new NotSupportedException("PT_MV_BINARY property type is not supported");
                default:
                    return value;
            }
        }
    }
}
