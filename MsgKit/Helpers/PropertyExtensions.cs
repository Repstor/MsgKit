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

        public static string RestPropertyId(this PropertyTag prop) =>
             PropertyIdTypeName(prop.Type) + prop.Id.ToString("X4");

        /// <summary>
        ///     Get the PropertyId as reconised by Exchange REST Api
        /// </summary>
        //public static string RestPropertyId(this PropertyTag prop) =>
        //     PropertyIdTypeName(prop.Type) + " {" + prop.Guid.ToString() + "} Name " + prop.;

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
    }
}
