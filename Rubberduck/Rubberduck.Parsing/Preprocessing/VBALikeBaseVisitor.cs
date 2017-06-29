//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     ANTLR Version: 4.3
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

// Generated from C:\Users\Splinter\Documents\Visual Studio 2015\Projects\TestProj\TestProj\VBALike.g4 by ANTLR 4.3

// Unreachable code detected

using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;

#pragma warning disable 0162
// The variable '...' is assigned but its value is never used
#pragma warning disable 0219
// Missing XML comment for publicly visible type or member '...'
#pragma warning disable 1591

namespace Rubberduck.Parsing.PreProcessing {
    /// <summary>
/// This class provides an empty implementation of <see cref="IVBALikeVisitor{Result}"/>,
/// which can be extended to create a visitor which only needs to handle a subset
/// of the available methods.
/// </summary>
/// <typeparam name="Result">The return type of the visit operation.</typeparam>
[System.CodeDom.Compiler.GeneratedCode("ANTLR", "4.3")]
[System.CLSCompliant(false)]
public partial class VBALikeBaseVisitor<Result> : AbstractParseTreeVisitor<Result>, IVBALikeVisitor<Result> {
	/// <summary>
	/// Visit a parse tree produced by <see cref="VBALikeParser.likePatternString"/>.
	/// <para>
	/// The default implementation returns the result of calling <see cref="AbstractParseTreeVisitor{Result}.VisitChildren(IRuleNode)"/>
	/// on <paramref name="context"/>.
	/// </para>
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	public virtual Result VisitLikePatternString([NotNull] VBALikeParser.LikePatternStringContext context) { return VisitChildren(context); }

	/// <summary>
	/// Visit a parse tree produced by <see cref="VBALikeParser.compilationUnit"/>.
	/// <para>
	/// The default implementation returns the result of calling <see cref="AbstractParseTreeVisitor{Result}.VisitChildren(IRuleNode)"/>
	/// on <paramref name="context"/>.
	/// </para>
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	public virtual Result VisitCompilationUnit([NotNull] VBALikeParser.CompilationUnitContext context) { return VisitChildren(context); }

	/// <summary>
	/// Visit a parse tree produced by <see cref="VBALikeParser.likePatternCharlistChar"/>.
	/// <para>
	/// The default implementation returns the result of calling <see cref="AbstractParseTreeVisitor{Result}.VisitChildren(IRuleNode)"/>
	/// on <paramref name="context"/>.
	/// </para>
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	public virtual Result VisitLikePatternCharlistChar([NotNull] VBALikeParser.LikePatternCharlistCharContext context) { return VisitChildren(context); }

	/// <summary>
	/// Visit a parse tree produced by <see cref="VBALikeParser.likePatternCharlist"/>.
	/// <para>
	/// The default implementation returns the result of calling <see cref="AbstractParseTreeVisitor{Result}.VisitChildren(IRuleNode)"/>
	/// on <paramref name="context"/>.
	/// </para>
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	public virtual Result VisitLikePatternCharlist([NotNull] VBALikeParser.LikePatternCharlistContext context) { return VisitChildren(context); }

	/// <summary>
	/// Visit a parse tree produced by <see cref="VBALikeParser.likePatternElement"/>.
	/// <para>
	/// The default implementation returns the result of calling <see cref="AbstractParseTreeVisitor{Result}.VisitChildren(IRuleNode)"/>
	/// on <paramref name="context"/>.
	/// </para>
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	public virtual Result VisitLikePatternElement([NotNull] VBALikeParser.LikePatternElementContext context) { return VisitChildren(context); }

	/// <summary>
	/// Visit a parse tree produced by <see cref="VBALikeParser.likePatternChar"/>.
	/// <para>
	/// The default implementation returns the result of calling <see cref="AbstractParseTreeVisitor{Result}.VisitChildren(IRuleNode)"/>
	/// on <paramref name="context"/>.
	/// </para>
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	public virtual Result VisitLikePatternChar([NotNull] VBALikeParser.LikePatternCharContext context) { return VisitChildren(context); }

	/// <summary>
	/// Visit a parse tree produced by <see cref="VBALikeParser.likePatternCharlistElement"/>.
	/// <para>
	/// The default implementation returns the result of calling <see cref="AbstractParseTreeVisitor{Result}.VisitChildren(IRuleNode)"/>
	/// on <paramref name="context"/>.
	/// </para>
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	public virtual Result VisitLikePatternCharlistElement([NotNull] VBALikeParser.LikePatternCharlistElementContext context) { return VisitChildren(context); }

	/// <summary>
	/// Visit a parse tree produced by <see cref="VBALikeParser.likePatternCharlistRange"/>.
	/// <para>
	/// The default implementation returns the result of calling <see cref="AbstractParseTreeVisitor{Result}.VisitChildren(IRuleNode)"/>
	/// on <paramref name="context"/>.
	/// </para>
	/// </summary>
	/// <param name="context">The parse tree.</param>
	/// <return>The visitor result.</return>
	public virtual Result VisitLikePatternCharlistRange([NotNull] VBALikeParser.LikePatternCharlistRangeContext context) { return VisitChildren(context); }
}
} // namespace Rubberduck.Parsing.Like
