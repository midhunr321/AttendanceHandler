using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceHander
{
    class ColourHandler
    {

        public static Color get_random_colour()
        {
            var rand = new Random();
            Color c = Color.FromArgb(rand.Next(256),
                rand.Next(256), rand.Next(256));

            return c;

        }

    }
}
