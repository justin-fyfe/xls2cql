/*
 * Licensed under the Apache License, Version 2.0 (the "License"); you 
 * may not use this file except in compliance with the License. You may 
 * obtain a copy of the License at 
 * 
 * http://www.apache.org/licenses/LICENSE-2.0 
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
 * WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the 
 * License for the specific language governing permissions and limitations under 
 * the License.
 * 
 * User: fyfej
 * Date: 2022-3-4
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace Xls2Cql.DecisionTable
{
    /// <summary>
    /// Generic CQL expression
    /// </summary>
    public abstract class CqlExpression
    {

        // Regular expression to extract a clause
        private static readonly Regex clauseExtraction = new Regex(@"^(\""[\s\S]*?\"")\s*?([\=\!\<\>\&\|]{1,2}).\s*?(\""[\s\S]*?\""|TRUE|FALSE|[\d\w\s]*?)$", RegexOptions.IgnoreCase);

        private static readonly Dictionary<String, CqlBinaryOperator> operatorMap = new Dictionary<string, CqlBinaryOperator>()
        {
            { "=", CqlBinaryOperator.Equal },
            { "!", CqlBinaryOperator.NotEqual },
            { "!=", CqlBinaryOperator.NotEqual },
            { "<=", CqlBinaryOperator.LessThanEqualTo },
            { ">=", CqlBinaryOperator.GreaterThanEqualTo },
            { ">", CqlBinaryOperator.GreaterThan },
            { "<", CqlBinaryOperator.LessThan },
            { "&", CqlBinaryOperator.And },
            { "&&", CqlBinaryOperator.And },
            { "and", CqlBinaryOperator.And },
            { "or", CqlBinaryOperator.Or },
            { "|", CqlBinaryOperator.Or },
            { "||", CqlBinaryOperator.Or }
        };

        /// <summary>
        /// Parses a single clause into an expression
        /// </summary>
        public static CqlExpression Parse(String parseCell)
        {

            var match = clauseExtraction.Match(parseCell);
            if(!match.Success)
            {
                match = clauseExtraction.Match(parseCell + " = TRUE"); // Handles case of just a single clause
            }
            if(!match.Success)
            {
                throw new InvalidOperationException($"The expression {parseCell} is not well formed - please use syntax \"attribute\" [=,!=,<,>,>=,<=] [true,false,####,\"other attribute\"]");
            }

            if(!operatorMap.TryGetValue(match.Groups[2].Value, out var op))
            {
                throw new InvalidOperationException($"Operator {match.Groups[2].Value} not supported. Use one of : =, <, <=, >, >=, !=, !, &, |");
            }

            return new CqlBinaryExpression(op, new CqlIdentifier(match.Groups[1].Value), new CqlIdentifier(match.Groups[3].Value));
        }
    }

    /// <summary>
    /// Binary operator
    /// </summary>
    public enum CqlBinaryOperator
    {
        And,
        Or,
        Equal,
        NotEqual,
        GreaterThan,
        GreaterThanEqualTo,
        LessThan,
        LessThanEqualTo
    }


    /// <summary>
    /// A simple identifier
    /// </summary>
    public class CqlIdentifier : CqlExpression
    {

        public CqlIdentifier(String identifier)
        {
            this.Identifier = identifier;
        }

        /// <summary>
        /// Gets the identifier
        /// </summary>
        public String Identifier { get; }

        /// <summary>
        /// Represent as a string
        /// </summary>
        public override string ToString()
        {
            if (this.Identifier.StartsWith("\""))
            {
                return this.Identifier.Replace("\r","").Replace("\n","");
            }
            else
            {
                return this.Identifier.ToLower();
            }
        }
    }

    /// <summary>
    /// CQL binary expression
    /// </summary>
    public class CqlBinaryExpression : CqlExpression
    {

        private static readonly Dictionary<CqlBinaryOperator, String> operatorMap = new Dictionary<CqlBinaryOperator, String>()
        {
            { CqlBinaryOperator.Equal , "=" },
            { CqlBinaryOperator.NotEqual, "<>"  },
            { CqlBinaryOperator.LessThanEqualTo, "<="  },
            { CqlBinaryOperator.GreaterThanEqualTo, ">="  },
            { CqlBinaryOperator.GreaterThan, ">"  },
            { CqlBinaryOperator.LessThan, "<"  },
            { CqlBinaryOperator.And, "and"  },
            { CqlBinaryOperator.Or , "or" }
        };

        /// <summary>
        /// Creates a new binary expression
        /// </summary>
        public CqlBinaryExpression(CqlBinaryOperator op, CqlExpression left, CqlExpression right)
        {
            this.Operator = op;
            this.Left = left;
            this.Right = right;
        }

        /// <summary>
        /// Operator
        /// </summary>
        public CqlBinaryOperator Operator { get; }

        /// <summary>
        /// Left side of expression
        /// </summary>
        public CqlExpression Left { get; }

        /// <summary>
        /// Right side of expression
        /// </summary>
        public CqlExpression Right { get; }

        public override string ToString()
        {
            var sb = new StringBuilder("(");
            sb.Append(this.Left);
            sb.AppendFormat(" {0} ", operatorMap[this.Operator]);
            sb.Append(this.Right);
            sb.Append(")");
            return sb.ToString();
        }
    }
}
