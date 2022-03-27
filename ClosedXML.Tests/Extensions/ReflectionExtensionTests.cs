using ClosedXML.Extensions;
using NUnit.Framework;
using System;
using System.Linq;
using System.Reflection;

namespace ClosedXML.Tests.Extensions
{
    public class ReflectionExtensionTests
    {
        private class TestClass
        {
            static TestClass()
            {
            }

            public static int StaticProperty { get; set; }
            public static int StaticField;

            public static event EventHandler<EventArgs> StaticEvent;

            public static void StaticMethod()
            {
            }

            public const int Const = 100;

            public TestClass()
            {
            }

            public int InstanceProperty { get; set; }
            public int InstanceField;

            public event EventHandler<EventArgs> InstanceEvent;

            public void InstanceMethod()
            {
            }
        }

        [TestCase(nameof(TestClass.StaticProperty), true)]
        [TestCase(nameof(TestClass.StaticField), true)]
        [TestCase(nameof(TestClass.StaticEvent), true)]
        [TestCase(nameof(TestClass.StaticMethod), true)]
        [TestCase(nameof(TestClass.Const), true)]
        [TestCase(nameof(TestClass.InstanceProperty), false)]
        [TestCase(nameof(TestClass.InstanceField), false)]
        [TestCase(nameof(TestClass.InstanceEvent), false)]
        [TestCase(nameof(TestClass.InstanceMethod), false)]
        public void IsStatic(string memberName, bool expectedIsStatic)
        {
            var member = typeof(TestClass).GetMember(memberName).Single();
            Assert.AreEqual(expectedIsStatic, member.IsStatic());
        }

        [TestCase(BindingFlags.Static | BindingFlags.NonPublic, true)]
        [TestCase(BindingFlags.Instance | BindingFlags.Public, false)]
        public void ConstructorIsStatic(BindingFlags flag, bool expectedIsStatic)
        {
            var constructors = typeof(TestClass).GetConstructors(flag);
            Assert.AreEqual(expectedIsStatic, constructors.Single().IsStatic());
        }
    }
}
