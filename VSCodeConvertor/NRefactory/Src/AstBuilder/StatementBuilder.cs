// <file>
//     <copyright see="prj:///doc/copyright.txt"/>
//     <license see="prj:///doc/license.txt"/>
//     <author name="Daniel Grunwald"/>
//     <version>$Revision: 3832 $</version>
// </file>

using System;
using System.Collections.Generic;
using ICSharpCode.NRefactory.Ast;

namespace ICSharpCode.NRefactory.AstBuilder
{
	/// <summary>
	/// Extension methods for NRefactory.Ast.Expression.
	/// </summary>
	public static class StatementBuilder
	{       
        #if NET35
            public static void AddStatement(this BlockStatement block, Statement statement)
        #else
            public static void AddStatement(BlockStatement block, Statement statement)
        #endif
		{
			if (block == null)
				throw new ArgumentNullException("block");
			if (statement == null)
				throw new ArgumentNullException("statement");
			block.AddChild(statement);
			statement.Parent = block;
		}
		
        #if NET35
		    public static void AddStatement(this BlockStatement block, Expression expressionStatement)
        #else
            public static void AddStatement(BlockStatement block, Expression expressionStatement)
        #endif
		{
			if (expressionStatement == null)
				throw new ArgumentNullException("expressionStatement");
			AddStatement(block, new ExpressionStatement(expressionStatement));
		}
		
        #if NET35
            public static void Throw(this BlockStatement block, Expression expression)
        #else
            public static void Throw(BlockStatement block, Expression expression)
        #endif
		{
			if (expression == null)
				throw new ArgumentNullException("expression");
			AddStatement(block, new ThrowStatement(expression));
		}
		#if NET35
            public static void Return(this BlockStatement block, Expression expression)
        #else
            public static void Return(BlockStatement block, Expression expression)
        #endif
            //public static void Return(this BlockStatement block, Expression expression)
		{
			if (expression == null)
				throw new ArgumentNullException("expression");
			AddStatement(block, new ReturnStatement(expression));
		}
		#if NET35
        public static void Assign(this BlockStatement block, Expression left, Expression right)
        #else
            public static void Assign(BlockStatement block, Expression left, Expression right)
        #endif
        //    public static void Assign(this BlockStatement block, Expression left, Expression right)
		{
			if (left == null)
				throw new ArgumentNullException("left");
			if (right == null)
				throw new ArgumentNullException("right");
			AddStatement(block, new AssignmentExpression(left, AssignmentOperatorType.Assign, right));
		}
	}
}
