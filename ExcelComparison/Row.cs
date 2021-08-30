using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelComparison
{
    public class Row
    {
        public string Entry1 { get; set; }
        public string Entry2 { get; set; }
        public string Entry3 { get; set; }
        public string Entry4 { get; set; }
        public string Entry5 { get; set; }
        public string Entry6 { get; set; }
        public string Entry7 { get; set; }
        public string Entry8 { get; set; }
        public string Entry9 { get; set; }
        public string Entry10 { get; set; }
        public string Entry11 { get; set; }
        public string Entry12 { get; set; }
        public string Entry13 { get; set; }
        public string Entry14 { get; set; }
        public string Entry15 { get; set; }
        public string Entry16 { get; set; }
        public string Entry17 { get; set; }
        public string Entry18 { get; set; }
        public string Entry19 { get; set; }
        public string Entry20 { get; set; }
        public string Entry21 { get; set; }
        public string Entry22 { get; set; }
        public string Entry23 { get; set; }
        public string Entry24 { get; set; }
        public string Entry25 { get; set; }
        public string Entry26 { get; set; }
        public string Entry27 { get; set; }
        public string Entry28 { get; set; }
        public string Entry29 { get; set; }
        public string Entry30 { get; set; }

        public override bool Equals(object obj)
        {
            // If the passed object is null
            if (obj == null)
            {
                return false;
            }
            if (!(obj is Row))
            {
                return false;
            }
            return (this.Entry1 == ((Row)obj).Entry1)
                && (this.Entry2 == ((Row)obj).Entry2)
                && (this.Entry3 == ((Row)obj).Entry3)
                && (this.Entry4 == ((Row)obj).Entry4)
                && (this.Entry5 == ((Row)obj).Entry5)
                && (this.Entry6 == ((Row)obj).Entry6)
                && (this.Entry7 == ((Row)obj).Entry7)
                && (this.Entry8 == ((Row)obj).Entry8)
                && (this.Entry9 == ((Row)obj).Entry9)
                && (this.Entry10 == ((Row)obj).Entry10)
                && (this.Entry11 == ((Row)obj).Entry11)
                && (this.Entry12 == ((Row)obj).Entry12)
                && (this.Entry13 == ((Row)obj).Entry13)
                && (this.Entry14 == ((Row)obj).Entry14)
                && (this.Entry15 == ((Row)obj).Entry15)
                && (this.Entry16 == ((Row)obj).Entry16)
                && (this.Entry17 == ((Row)obj).Entry17)
                && (this.Entry18 == ((Row)obj).Entry18)
                && (this.Entry19 == ((Row)obj).Entry19)
                && (this.Entry20 == ((Row)obj).Entry20)
                && (this.Entry21 == ((Row)obj).Entry21)
                && (this.Entry22 == ((Row)obj).Entry22)
                && (this.Entry23 == ((Row)obj).Entry23)
                && (this.Entry24 == ((Row)obj).Entry24)
                && (this.Entry25 == ((Row)obj).Entry25)
                && (this.Entry26 == ((Row)obj).Entry26)
                && (this.Entry27 == ((Row)obj).Entry27)
                && (this.Entry28 == ((Row)obj).Entry28)
                && (this.Entry29 == ((Row)obj).Entry29)
                && (this.Entry30 == ((Row)obj).Entry30);
        }
        public override int GetHashCode()
        {
            return Entry1.GetHashCode() ^ Entry2.GetHashCode() ^ Entry3.GetHashCode() ^ Entry4.GetHashCode()
                ^ Entry5.GetHashCode() ^ Entry6.GetHashCode() ^ Entry7.GetHashCode() ^ Entry8.GetHashCode()
                ^ Entry9.GetHashCode() ^ Entry10.GetHashCode() ^ Entry11.GetHashCode() ^ Entry12.GetHashCode()
                ^ Entry13.GetHashCode() ^ Entry14.GetHashCode() ^ Entry15.GetHashCode() ^ Entry16.GetHashCode()
                ^ Entry17.GetHashCode() ^ Entry18.GetHashCode() ^ Entry19.GetHashCode() ^ Entry20.GetHashCode()
                ^ Entry21.GetHashCode() ^ Entry22.GetHashCode() ^ Entry23.GetHashCode() ^ Entry24.GetHashCode()
                ^ Entry25.GetHashCode() ^ Entry26.GetHashCode() ^ Entry27.GetHashCode() ^ Entry28.GetHashCode()
                ^ Entry29.GetHashCode() ^ Entry30.GetHashCode();
        }

    }
}
