using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MsgKit
{
    /// <summary>
    /// Set of constants - common property set ids - [MS-OXPROPS] 1.3.2
    /// </summary>
    internal class PropertySets
    {
        internal static Guid PS_PUBLIC_STRINGS => new Guid("00020329-0000-0000-C000-000000000046");
        internal static Guid PS_INTERNET_HEADERS => new Guid("00020386-0000-0000-C000-000000000046");
        internal static Guid PSETID_Common => new Guid("00062008-0000-0000-C000-000000000046");
        internal static Guid PS_MAPI => new Guid("00020328-0000-0000-C000-000000000046");
        internal static Guid PSETID_Task => new Guid("00062003-0000-0000-C000-000000000046");
        internal static Guid PSETID_Appointment => new  Guid("00062002-0000-0000-C000-000000000046");
        internal static Guid PSETID_Meeting => new  Guid("6ED8DA90-450B-101B-98DA-00AA003F1305");
    }
}
