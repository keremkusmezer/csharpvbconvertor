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
	public static class ExpressionBuilder
	{
		public static IdentifierExpression Identifier(string identifier)
		{
			return new IdentifierExpression(identifier);
		}
		#if NET35
		public static MemberReferenceExpression Member(this Expression targetObject, string memberName)
        #else
            public static MemberReferenceExpression Member(Expression targetObject, string memberName)
        #endif
		{
			if (targetObject == null)
				throw new ArgumentNullException("targetObject");
			return new MemberReferenceExpression(targetObject, memberName);
		}
		#if NET35
		public static InvocationExpression Call(this Expression callTarget, string methodName, params Expression[] arguments)
        #else
            public static InvocationExpression Call(Expression callTarget, string methodName, params Expression[] arguments)
        #endif
		{
			if (callTarget == null)
				throw new ArgumentNullException("callTarget");
			return Call(Member(callTarget, methodName), arguments);
		}
#if NET35
        public static InvocationExpression Call(this Expression callTarget, params Expression[] arguments)
#else
        public static InvocationExpression Call(Expression callTarget, params Expression[] arguments)
#endif
//            public static InvocationExpression Call(this Expression callTarget, params Expression[] arguments)
		{
			if (callTarget == null)
				throw new ArgumentNullException("callTarget");
			if (arguments == null)
				throw new ArgumentNullException("arguments");
			return new InvocationExpression(callTarget, new List<Expression>(arguments));
		}
		#if NET35
        public static ObjectCreateExpression New(this TypeReference createType, params Expression[] arguments)
        #else
        public static ObjectCreateExpression New(TypeReference createType, params Expression[] arguments)
        #endif
        //public static ObjectCreateExpression New(this TypeReference createType, params Expression[] arguments)
		{
			if (createType == null)
				throw new ArgumentNullException("createType");
			if (arguments == null)
				throw new ArgumentNullException("arguments");
			return new ObjectCreateExpression(createType, new List<Expression>(arguments));
		}
		
		public static Expression CreateDefaultValueForType(TypeReference type)
		{
			if (type != null && !type.IsArrayType) {
				switch (type.Type) {
					case "System.SByte":
					case "System.Byte":
					case "System.Int16":
					case "System.UInt16":
					case "System.Int32":
					case "System.UInt32":
					case "System.Int64":
					case "System.UInt64":
					case "System.Single":
					case "System.Double":
						return new PrimitiveExpression(0, "0");
					case "System.Char":
						return new PrimitiveExpression('\0', "'\\0'");
					case "System.Object":
					case "System.String":
						return new PrimitiveExpression(null, "null");
					case "System.Boolean":
						return new PrimitiveExpression(false, "false");
					default:
						return new DefaultValueExpression(type);
				}
			} else {
				return new PrimitiveExpression(null, "null");
			}
		}
#if NET35
		/// <summary>
		/// Just calls the BinaryOperatorExpression constructor,
		/// but being an extension method; this allows for a nicer
		/// infix syntax in some cases.
		/// </summary>
		public static BinaryOperatorExpression Operator(this Expression left, BinaryOperatorType op, Expression right)
		{
			return new BinaryOperatorExpression(left, op, right);
		}
#else
        public static BinaryOperatorExpression Operator(Expression left, BinaryOperatorType op, Expression right)
        {
            return new BinaryOperatorExpression(left, op, right);
        }
#endif
	}
}
