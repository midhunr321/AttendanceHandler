using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceHander
{
    interface IfileOpenAction
    {
        void okButtonPressed(List<FileInfo> fileinfos);
        void okButtonPressed(FileInfo fileInfo);
        void cancelled();
    }
}
