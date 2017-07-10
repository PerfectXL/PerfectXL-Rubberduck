using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace PerfectXL.VbaCodeAnalyzer.Models
{
    internal class Windows : IWindows
    {
        public Windows(IVBE vbe)
        {
            VBE = vbe;
        }

        public IList<IWindow> WindowList { get; } = new List<IWindow>();

        public int Count => WindowList.Count;

        public IWindow this[object index] => index is string ? WindowList.SingleOrDefault(window => window.Caption == (string)index) : WindowList[(int)index];

        IEnumerator IEnumerable.GetEnumerator()
        {
            return WindowList.GetEnumerator();
        }

        IEnumerator<IWindow> IEnumerable<IWindow>.GetEnumerator()
        {
            return WindowList.GetEnumerator();
        }

        public bool Equals(IWindows other)
        {
            return Equals(this, other);
        }

        public object Target { get; }
        public bool IsWrappingNullReference { get; }

        public IVBE VBE { get; set; }

        public IApplication Parent => null;

        public ToolWindowInfo CreateToolWindow(IAddIn addInInst, string progId, string caption, string guidPosition)
        {
            return new ToolWindowInfo(new Window(), null);
        }

        public override int GetHashCode()
        {
            return WindowList.GetHashCode();
        }
    }
}