using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace PerfectXL.VbaCodeAnalyzer.Macro
{
    public class MacroTerm
    {
        private static readonly List<MacroTermRating> TermWithRating = new List<MacroTermRating>();

        static MacroTerm()
        {
            TermWithRating.AddRange(new List<MacroTermRating>
            {
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Activate" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="ActiveChart" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="ActiveSheet" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="ActiveWindow" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="ActiveWorkbook" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Add" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="AllowMultiSelect" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Application" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Apply" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Auto_Open"},
                new MacroTermRating {Rate = (decimal) 0.1, Term ="AutoClose" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="AutoExec" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="AutoExit" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="AutoFill" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="AutoNew" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="AutoOpen" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Clear" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Close" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Copy" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Count" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="CutCopyMode" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Delete" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Destination" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Display3DShading" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="DisplayFullScreen" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="DisplayHeadings" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="DropDownLines" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="End" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Formula" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Header" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Hyperlinks" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Insert" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="LinkedCell" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="ListFillRange" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="MatchVase" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="MsgBox" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="msoFileOpen" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Offset" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Open" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Orientation" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="PasteSpecial" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Protect" },
                new MacroTermRating {Rate = (decimal) 0.700, Term ="Range" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Refresh" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Rows" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Save" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="ScreenUpdating" },
                new MacroTermRating {Rate = (decimal) 0.300, Term ="Select" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="SelectedSheets" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Selection" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="SelectionChange" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="SetRange" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Shapes" },
                new MacroTermRating {Rate = (decimal) 0.200, Term ="Sheets" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Show" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="SmallScroll" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="SortFields" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="SortMethod" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Unprotect" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Values" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Windows" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Workbook_Open" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="XValues" },
                new MacroTermRating {Rate = (decimal) 0.1, Term ="Zoom"}
            });


            // XmlTextReader reader = new XmlTextReader(@"C:\Users\HarveyBouva\Projects\PerfectXL\PerfectXL-Rubberduck\PerfectXL-Rubberduck\MacroTermRating.xml");

            //todo implement xml https://www.reddit.com/r/csharp/comments/351hk4/using_linq_to_extract_attributes_from_xml/ 
        }

        public static IEnumerable<string> List()
        {
            return TermWithRating.Select(x => x.Term).ToList();
        }

        public static IEnumerable<MacroTermRating> Rates()
        {
            return TermWithRating;
        }

    }

    public class MacroTermRating
    {
        public string Term { get; set; }
        public decimal Rate { get; set; }
    }

    public class MacroTermPresenter
    {
        public string Module { get; set; }
        public string Function { get; set; }
        public string Term { get; set; }
        public bool Listed { get; set; }
        public int Repeat { get; set; }
        public decimal Percentage { get; set; }
    }

}
