using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace JeproksReport
{
    public enum PageOrientation
    {
        Portrait, Landscape
    }
    public class Margins
    {
        [XmlAttribute]
        public double Top { get; set; }
        public bool ShouldSerializeTop()
        {
            return Top != 30;
        }
        [XmlAttribute]
        public double Right { get; set; }
        public bool ShouldSerializeRight()
        {
            return Right != 30;
        }
        [XmlAttribute]
        public double Bottom { get; set; }
        public bool ShouldSerializeBottom()
        {
            return Bottom != 30;
        }
        [XmlAttribute]
        public double Left { get; set; }
        public bool ShouldSerializeLeft()
        {
            return Left != 30;
        }

        public Margins()
        {
            this.Top = 30;
            this.Right = 30;
            this.Bottom = 30;
            this.Left = 30;
        }
    }
    public class PageOptions
    {
        [XmlAttribute]
        public PageOrientation PageOrientation { get; set; }

        [XmlElement("PageOptions.Margins")]
        public Margins Margins { get; set; }
        public bool ShouldSerializeMargins()
        {
            return Margins.Bottom != 30 ||
                Margins.Left != 30 ||
                Margins.Right != 30 ||
                Margins.Top != 30;
        }
        public PageOptions()
        {
            this.Margins = new Margins();
        }
    }
}
