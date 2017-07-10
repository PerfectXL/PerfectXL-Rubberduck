using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace infotron.VbaCodeAnalizer.Mog
{
    internal class CodeModule : ICodeModule
    {
        private static readonly string[] ModuleBodyTokens = {Tokens.Sub + ' ', Tokens.Function + ' ', Tokens.Property + ' '};
        private List<string> _lines;

        public CodeModule(IVBE vbe,string name, string content, IVBComponent component, ICodePane codePane)
        {
            VBE = vbe;
            Name = name;
            _lines = content.Split(new[] {Environment.NewLine}, StringSplitOptions.None).ToList();
            Parent = component;
            CodePane = codePane;
        }

        public IVBE VBE { get; }
        public IVBComponent Parent { get; }
        public ICodePane CodePane { get; }

        public int CountOfDeclarationLines
        {
            get { return _lines.TakeWhile(line => line.Contains(Tokens.Declare + ' ') || !ModuleBodyTokens.Any(line.Contains)).Count(); }
        }

        public int CountOfLines => _lines.Count;
        public string Name { get; set; }

        public string GetLines(int startLine, int count)
        {
            return string.Join(Environment.NewLine, _lines.Skip(startLine - 1).Take(count));
        }

        public string GetLines(Selection selection)
        {
            return string.Join(Environment.NewLine, _lines.Skip(selection.StartLine - 1).Take(selection.LineCount));
        }

        public void DeleteLines(Selection selection)
        {
            _lines.RemoveRange(selection.StartLine - 1, selection.LineCount);
        }

        public void DeleteLines(int startLine, int count = 1)
        {
            _lines.RemoveRange(startLine - 1, count);
        }

        public QualifiedSelection? GetQualifiedSelection()
        {
            throw new NotImplementedException();
        }

        public string Content()
        {
            return string.Join(Environment.NewLine, _lines);
        }

        public void Clear()
        {
            _lines = new List<string>();
        }

        public string ContentHash()
        {
            throw new NotImplementedException();
        }

        public void AddFromString(string content)
        {
            _lines.AddRange(content.Split(new[] {Environment.NewLine}, StringSplitOptions.None));
        }

        public void AddFromFile(string path)
        {
            throw new NotImplementedException();
        }

        public void InsertLines(int line, string content)
        {
            if (line - 1 >= _lines.Count)
            {
                _lines.AddRange(content.Split(new[] {Environment.NewLine}, StringSplitOptions.None));
            }
            else
            {
                _lines.InsertRange(line - 1, content.Split(new[] {Environment.NewLine}, StringSplitOptions.None));
            }
        }

        public void ReplaceLine(int line, string content)
        {
            _lines[line - 1] = content;
        }

        public int GetProcStartLine(string procName, ProcKind procKind)
        {
            throw new NotImplementedException();
        }

        public int GetProcBodyStartLine(string procName, ProcKind procKind)
        {
            throw new NotImplementedException();
        }

        public int GetProcCountLines(string procName, ProcKind procKind)
        {
            throw new NotImplementedException();
        }

        public string GetProcOfLine(int line)
        {
            throw new NotImplementedException();
        }

        public ProcKind GetProcKindOfLine(int line)
        {
            throw new NotImplementedException();
        }

        public bool Equals(ICodeModule other)
        {
            return Name.Equals(other.Name) && Content().Equals(other.Content());
        }

        public object Target { get; }
        public bool IsWrappingNullReference { get; }

        public override int GetHashCode()
        {
            return Target.GetHashCode();
        }
    }
}