using NUnit.Framework;
using Moq;
using Rubberduck.Settings;
using Rubberduck.SmartIndenter;

namespace RubberduckTests.Settings
{
    public class IndenterSettingsTests
    {
        // Defaults
        private const int DefaultAlignDimColumn = 15;
        private const EndOfLineCommentStyle DefaultEndOfLineCommentStyle = EndOfLineCommentStyle.AlignInColumn;
        private const EmptyLineHandling DefaultEmptyLineHandlingMethod = EmptyLineHandling.Ignore;
        private const int DefaultEndOfLineCommentColumnSpaceAlignment = 50;
        private const int DefaultIndentSpaces = 4;
        private const int DefaultLinesBetweenProcedures = 1;

        // Nondefaults
        private const int NondefaultAlignDimColumn = 16;
        private const EndOfLineCommentStyle NondefaultEndOfLineCommentStyle = EndOfLineCommentStyle.Absolute;
        private const EmptyLineHandling NondefaultEmptyLineHandlingMethod = EmptyLineHandling.Remove;
        private const int NondefaultEndOfLineCommentColumnSpaceAlignment = 60;
        private const int NondefaultIndentSpaces = 2;
        private const int NondefaultLinesBetweenProcedures = 2;

        public static Rubberduck.SmartIndenter.IndenterSettings GetMockIndenterSettings(bool nondefault = false)
        {
            var output = new Mock<Rubberduck.SmartIndenter.IndenterSettings>();

            output.SetupProperty(s => s.IndentEntireProcedureBody);
            output.SetupProperty(s => s.IndentFirstCommentBlock);
            output.SetupProperty(s => s.IndentFirstDeclarationBlock);
            output.SetupProperty(s => s.IgnoreEmptyLinesInFirstBlocks);
            output.SetupProperty(s => s.AlignCommentsWithCode);
            output.SetupProperty(s => s.AlignContinuations);
            output.SetupProperty(s => s.IgnoreOperatorsInContinuations);
            output.SetupProperty(s => s.IndentCase);
            output.SetupProperty(s => s.IndentEnumTypeAsProcedure);
            output.SetupProperty(s => s.ForceDebugPrintInColumn1);
            output.SetupProperty(s => s.ForceDebugAssertInColumn1);
            output.SetupProperty(s => s.ForceStopInColumn1);
            output.SetupProperty(s => s.ForceCompilerDirectivesInColumn1);
            output.SetupProperty(s => s.IndentCompilerDirectives);
            output.SetupProperty(s => s.AlignDims);
            output.SetupProperty(s => s.AlignDimColumn);
            output.SetupProperty(s => s.EndOfLineCommentStyle);
            output.SetupProperty(s => s.EndOfLineCommentColumnSpaceAlignment);
            output.SetupProperty(s => s.EmptyLineHandlingMethod);
            output.SetupProperty(s => s.IndentSpaces);
            output.SetupProperty(s => s.VerticallySpaceProcedures);
            output.SetupProperty(s => s.LinesBetweenProcedures);
            output.SetupProperty(s => s.GroupRelatedProperties);

            output.Object.IndentEntireProcedureBody = !nondefault;
            output.Object.IndentFirstCommentBlock = !nondefault;
            output.Object.IndentFirstDeclarationBlock = !nondefault;
            output.Object.IgnoreEmptyLinesInFirstBlocks = nondefault;
            output.Object.AlignCommentsWithCode = !nondefault;
            output.Object.AlignContinuations = !nondefault;
            output.Object.IgnoreOperatorsInContinuations = !nondefault;
            output.Object.IndentCase = nondefault;
            output.Object.IndentEnumTypeAsProcedure = nondefault;
            output.Object.ForceDebugPrintInColumn1 = nondefault;
            output.Object.ForceDebugAssertInColumn1 = nondefault;
            output.Object.ForceStopInColumn1 = nondefault;
            output.Object.ForceCompilerDirectivesInColumn1 = nondefault;
            output.Object.IndentCompilerDirectives = !nondefault;
            output.Object.AlignDims = nondefault;
            output.Object.AlignDimColumn = nondefault ? NondefaultAlignDimColumn : DefaultAlignDimColumn;
            output.Object.EndOfLineCommentStyle = nondefault ? NondefaultEndOfLineCommentStyle : DefaultEndOfLineCommentStyle;
            output.Object.EndOfLineCommentColumnSpaceAlignment = nondefault ? NondefaultEndOfLineCommentColumnSpaceAlignment : DefaultEndOfLineCommentColumnSpaceAlignment;
            output.Object.EmptyLineHandlingMethod = nondefault ? NondefaultEmptyLineHandlingMethod : DefaultEmptyLineHandlingMethod;
            output.Object.IndentSpaces = nondefault ? NondefaultIndentSpaces : DefaultIndentSpaces;
            output.Object.VerticallySpaceProcedures = !nondefault;
            output.Object.LinesBetweenProcedures = nondefault ? NondefaultLinesBetweenProcedures : DefaultLinesBetweenProcedures;
            output.Object.GroupRelatedProperties = nondefault;

            return output.Object;
        }

    }
}
