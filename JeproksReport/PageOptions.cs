using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JeproksReport
{
    public enum PageOrientation
    {
        Portrait, Landscape
    }

    public class PageOptions
    {
        public PageOrientation PageOrientation { get; set; }
        public Margins Margins { get; set; }
        public PageSize PageSize { get; set; }

        public PageOptions()
        {
            this.Margins = new Margins();
            this.PageSize = PageSize.A4;
        }
    }
    public class Margins
    {
        public double Top { get; set; }
        public double Right { get; set; }
        public double Bottom { get; set; }
        public double Left { get; set; }

        public Margins()
        {
            this.Top = 30;
            this.Right = 30;
            this.Bottom = 30;
            this.Left = 30;
        }
    }
    public enum PageSize
    {
        A4, Letter
    }
}
