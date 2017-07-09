using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Moq;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace infotron.VbaCodeAnalizer.Mog
{
    /// <summary>
    /// Builds a mock VBE.
    /// </summary>
    internal class MockVbeBuilder
    {
        private const string TestProjectName = "TestProject1";
        private const string TestModuleName = "TestModule1";
        private readonly IVBE _vbe;

        #region standard library paths (referenced in all VBA projects hosted in Microsoft Excel)
        private static readonly string LibraryPathVBA = @"C:\PROGRA~2\COMMON~1\MICROS~1\VBA\VBA7.1\VBE7.DLL";      // standard library, priority locked
        #endregion

        //private Mock<IWindows> _vbWindows;
        private readonly Windows _windows = new Windows();

        private IVBProjects _vbProjects;
        private readonly ICollection<IVBProject> _projects = new List<IVBProject>();
        
        private ICodePanes _vbCodePanes;
        private readonly ICollection<ICodePane> _codePanes = new List<ICodePane>(); 

        public MockVbeBuilder()
        {
            _vbe = CreateVbeMock();
        }

        /// <summary>
        /// Adds a project to the mock VBE.
        /// Use a <see cref="MockProjectBuilder"/> to build the <see cref="project"/>.
        /// </summary>
        /// <param name="project">A mock VBProject.</param>
        /// <returns>Returns the <see cref="MockVbeBuilder"/> instance.</returns>
        public MockVbeBuilder AddProject(IVBProject project)
        {
            // TODO roel This needs to be re-added once we have factored out all the mocking
            //project.SetupGet(m => m.VBE).Returns(_vbe);

            _projects.Add(project);

            foreach (var component in _projects.SelectMany(vbProject => vbProject.VBComponents))
            {
                _codePanes.Add(component.CodeModule.CodePane);
            }

            // TODO roel Do we need this? It is set in the `BuildFromSingleModule` method as well...
            //_vbe.SetupGet(vbe => vbe.ActiveVBProject).Returns(project);

            return this;
        }

        /// <summary>
        /// Creates a <see cref="MockProjectBuilder.Mog.MockProjectBuilder"/> to build a new project.
        /// </summary>
        /// <param name="name">The name of the project to build.</param>
        /// <param name="protection">A value that indicates whether the project is protected.</param>
        private MockProjectBuilder ProjectBuilder(string name, ProjectProtection protection)
        {
            return ProjectBuilder(name, string.Empty, protection);
        }

        private MockProjectBuilder ProjectBuilder(string name, string filename, ProjectProtection protection)
        {
            return new MockProjectBuilder(name, filename, protection, () => _vbe, this);
        }

        private MockProjectBuilder ProjectBuilder(string name, string filename, string projectId, ProjectProtection protection)
        {
            return new MockProjectBuilder(name, filename, projectId, protection, () => _vbe, this);
        }

        /// <summary>
        /// Gets the mock VBE instance.
        /// </summary>
        private IVBE Build()
        {
            return _vbe;
        }

        /// <summary>
        /// Gets a mock VBE instance, 
        /// containing a single "TestProject1" VBProject
        /// and a single "TestModule1" VBComponent, with the specified <see cref="content"/>.
        /// </summary>
        /// <param name="content">The VBA code associated to the component.</param>
        /// <param name="component">The created VBComponent</param>
        /// <param name="module">The created CodeModule</param>
        /// <param name="selection"></param>
        /// <returns></returns>
        public static IVBE BuildFromSingleStandardModule(string content, out IVBComponent component, Selection selection = default(Selection), bool referenceStdLibs = false)
        {
            return BuildFromSingleModule(content, TestModuleName, ComponentType.StandardModule, out component, selection, referenceStdLibs);
        }

        private static IVBE BuildFromSingleStandardModule(string content, string name, out IVBComponent component, Selection selection = default(Selection), bool referenceStdLibs = false)
        {
            return BuildFromSingleModule(content, name, ComponentType.StandardModule, out component, selection, referenceStdLibs);
        }

        private static IVBE BuildFromSingleModule(string content, ComponentType type, out IVBComponent component, Selection selection = default(Selection), bool referenceStdLibs = false)
        {
            return BuildFromSingleModule(content, TestModuleName, type, out component, selection, referenceStdLibs);
        }

        private static IVBE BuildFromSingleModule(string content, string name, ComponentType type, out IVBComponent component, Selection selection = default(Selection), bool referenceStdLibs = false)
        {
            var vbeBuilder = new MockVbeBuilder();

            var builder = vbeBuilder.ProjectBuilder(TestProjectName, ProjectProtection.Unprotected);
            builder.AddComponent(name, type, content, selection);

            if (referenceStdLibs)
            {
                builder.AddReference("VBA", LibraryPathVBA, 4, 1, true);
            }

            var project = builder.Build();
            var vbe = vbeBuilder.AddProject(project).Build();

            component = project.VBComponents[0];

            vbe.ActiveVBProject = project;
            vbe.ActiveCodePane = component.CodeModule.CodePane;

            return vbe;
        }

        private IVBE CreateVbeMock()
        {
            var vbe = new Mock<IVBE>();
            _windows.VBE = vbe.Object;
            vbe.Setup(m => m.Windows).Returns(() => _windows);
            vbe.SetupProperty(m => m.ActiveCodePane);
            vbe.SetupProperty(m => m.ActiveVBProject);
            
            vbe.SetupGet(m => m.SelectedVBComponent).Returns(() => vbe.Object.ActiveCodePane.CodeModule.Parent);
            vbe.SetupGet(m => m.ActiveWindow).Returns(() => vbe.Object.ActiveCodePane.Window);

            var mainWindow = new Mock<IWindow>();
            mainWindow.Setup(m => m.HWnd).Returns(0);

            vbe.SetupGet(m => m.MainWindow).Returns(() => mainWindow.Object);

            _vbProjects = CreateProjectsMock();
            vbe.SetupGet(m => m.VBProjects).Returns(() => _vbProjects);

            _vbCodePanes = CreateCodePanesMock();
            vbe.SetupGet(m => m.CodePanes).Returns(() => _vbCodePanes);

            vbe.SetupGet(m => m.Version).Returns("7.1");
            vbe.SetupGet(m => m.VBProjects).Returns(() => _vbProjects);


            return vbe.Object;
        }

        private IVBProjects CreateProjectsMock()
        {
            var result = new Mock<IVBProjects>();

            result.Setup(m => m.GetEnumerator()).Returns(() => _projects.GetEnumerator());
            result.As<IEnumerable>().Setup(m => m.GetEnumerator()).Returns(() => _projects.GetEnumerator());
            
            result.Setup(m => m[It.IsAny<int>()]).Returns<int>(value => _projects.ElementAt(value));
            result.SetupGet(m => m.Count).Returns(() => _projects.Count);


            return result.Object;
        }

        private ICodePanes CreateCodePanesMock()
        {
            var result = new Mock<ICodePanes>();

            result.Setup(m => m.GetEnumerator()).Returns(() => _codePanes.GetEnumerator());
            result.As<IEnumerable>().Setup(m => m.GetEnumerator()).Returns(() => _codePanes.GetEnumerator());
            
            result.Setup(m => m[It.IsAny<int>()]).Returns<int>(value => _codePanes.ElementAt(value));
            result.SetupGet(m => m.Count).Returns(() => _codePanes.Count);

            return result.Object;
        }
    }
}
