using System;
using System.Collections.Generic;
using System.Text;

namespace SHSchool.Behavior.StudentExtendControls
{
    public interface ICellInfo<T>
    {
        T OriginValue { get;}
        void SetValue(T value);
        bool IsDirty { get;}
    } 
}
