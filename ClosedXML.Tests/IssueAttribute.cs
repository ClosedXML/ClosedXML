using System;
using NUnit.Framework;

namespace ClosedXML.Tests;

/// <summary>
/// Link a test to a GitHub issue for more context about the original problem.
/// </summary>
[AttributeUsage(AttributeTargets.Method | AttributeTargets.Class, AllowMultiple = true)]
public class IssueAttribute(string issue) : PropertyAttribute("Issue", issue);
