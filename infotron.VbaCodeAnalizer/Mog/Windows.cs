using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace infotron.VbaCodeAnalizer.Mog
{
    internal class Windows : IWindows
    {
        private readonly IList<IWindow> _windows = new List<IWindow>();

        public Windows(IVBE vbe)
        {
            VBE = vbe;
        }

        public int Count => _windows.Count;

        public IWindow this[object index] => index is string ? _windows.SingleOrDefault(window => window.Caption == (string)index) : _windows[(int)index];

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _windows.GetEnumerator();
        }

        IEnumerator<IWindow> IEnumerable<IWindow>.GetEnumerator()
        {
            return _windows.GetEnumerator();
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
            return _windows.GetHashCode();
        }

        public void Add(IWindow window)
        {
            _windows.Add(window);
        }
    }
}