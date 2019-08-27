// Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

using System;
using System.Linq;
using System.Collections.Generic;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;
using Microsoft.CodeAnalysis.Testing;
using Test.Utilities;
using Text.CSharp.Analyzers;
using Text.VisualBasic.Analyzers;
using Xunit;

namespace Text.Analyzers.UnitTests
{
    public class IdentifiersShouldBeSpelledCorrectlyTests : DiagnosticAnalyzerTestBase
    {

        private const string MisspelledDelegateCode = @"
namespace MyNamespace
{
    public delegate void MyDelegate(string firstNaem);
}";

        private const string MisspelledMemberCode = @"
public class Program
{
    public void SomeMathod()
    {
    }

    public string Naem { get; set; }
}";

        private const string MisspelledMemberParameterCode = @"
public class Program
{
    public void Method(string firstNaem)
    {
    }
}";

        private const string MisspelledMethodTypeParameterCode = @"
public class Program
{
    public void Method<TTipe>(TTipe item)
    {
    }
}";

        private const string MisspelledNamespaceCode = @"
namespace Tests.DoSumthing
{
}";

        private const string MisspelledTypesCode = @"
namespace MyNamespace
{
    public class MyClazz
    {
    }

    public struct MyStroct
    {
    }

    public enum MiiEnumeration
    {
    }

    public interface IIface
    {
    }

    public delegate int MyDelegete();
}";

        private const string MisspelledTypeTypeParameterCode = @"
namespace MyNamespace
{
    public class MyClass<TCorrect, TWroong>
    {
    }

    public struct MyStructure<TWroong>
    {
    }

    public interface IInterface<TWroong>
    {
    }

    public delegate int MyDelegate<TWroong>();
}";

        private static readonly IEnumerable<string> RecognizedMisspellings = new[] {
            "naem",
            "mathod",
            "tipe",
            "sumthing",
            "clazz",
            "stroct",
            "mii",
            "iface",
            "delegete",
            "wroong",
        };

        public static IEnumerable<object[]> TestCodes =>
            new[] {
                new object[] { MisspelledDelegateCode },
                new object[] { MisspelledMemberCode },
                new object[] { MisspelledMemberParameterCode },
                new object[] { MisspelledMethodTypeParameterCode },
                new object[] { MisspelledNamespaceCode },
                new object[] { MisspelledTypesCode },
                new object[] { MisspelledTypeTypeParameterCode },
            };

        [Theory]
        [MemberData(nameof(TestCodes))]
        public void CSharp_CustomXmlDictionary_AllowsMisspelledWords(string testCode)
        {
            var dictionary = CreateCustomXmlDictionary(RecognizedMisspellings);
            VerifyCSharp(new[] { testCode }, additionalFile: dictionary, ReferenceFlags.None);
        }

        [Theory]
        [MemberData(nameof(TestCodes))]
        public void CSharp_CustomDicDictionary_AllowsMisspelledWords(string testCode)
        {
            var dictionary = CreateCustomDicDictionary(RecognizedMisspellings);
            VerifyCSharp(new[] { testCode }, additionalFile: dictionary, ReferenceFlags.None);
        }

        [Fact]
        public void CSharp_CustomXmlDictionary_DisallowsCorrectWords()
        {
            var dictionary = CreateCustomXmlDictionary(null, new[] { "meth", "pass" });
            var testCode = @"
public class Program
{
    public void Meth(string passMe)
    {
    }
}";

            VerifyCSharp(
                new[] { testCode },
                additionalFile: dictionary,
                ReferenceFlags.None,
                GetCA1704MemberResultAt(4, 17, "Meth", "Program.Meth(string)"),
                GetCA1704MemberParameterResultAt(4, 29, "Program.Meth(string)", "pass", "passMe"));
        }

        [Fact]
        public void CSharp_MultipleCustomDictionaries()
        {
            var dictionary1 = CreateCustomXmlDictionary("CodeAnalysisDictionary1.xml", new[] { "obj" }, new[] { "items" });
            var dictionary2 = CreateCustomXmlDictionary("CodeAnalysisDictionary2.xml", new[] { "param" }, new[] { "program" });
            var dictionary3 = CreateCustomDicDictionary(new[] { "tipe" });

            var testCode = @"
namespace MyNamespace
{
    public class Program
    {
        public delegate int CombineItems<TTipe>(TTipe firstParam, object secondObj);
    }
}";

            VerifyCSharp(
                testCode,
                additionalFiles: new[] { dictionary1, dictionary2, dictionary3 },
                expected: new[]
                {
                    GetCA1704TypeResultAt(4, 18, "Program", "MyNamespace.Program"),
                    GetCA1704TypeResultAt(6, 29, "Items", "MyNamespace.Program.CombineItems<TTipe>")
                });
        }

        [Fact]
        public void CSharp_AssemblyMisspelled_EmitsDiagnostic()
        {
            VerifyCSharp(@"
using System;

namespace Test
{
    public class Program
    {
        static void Main(string[] arguments)
        {
        }
    }
}",
            testProjectName: "MyAssambly",
            expected: GetCA1704AssemblyResultAt("Assambly", "MyAssambly"));
        }

        [Fact]
        public void CSharp_AssemblyUnmeaningful_EmitsDiagnostic()
        {
            VerifyCSharp(@"
public class Program
{
    static void Main(string[] arguments)
    {
    }
}",
            testProjectName: "A",
            expected: GetCA1704AssemblyUnmeaningfulResultAt("A"));
        }

        [Fact]
        public void CSharp_DelegateParameterMisspelled_EmitsDiagnostic()
        {

            VerifyCSharp(
                MisspelledDelegateCode,
                GetCA1704DelegateParameterResultAt(4, 44, "MyNamespace.MyDelegate", "Naem", "firstNaem"));
        }

        [Fact]
        public void CSharp_DelegateParameterUnmeaningful_EmitsDiagnostic()
        {

            VerifyCSharp(@"
namespace MyNamespace
{
    public delegate void MyDelegate(string a);
}",
                GetCA1704DelegateParameterUnmeaningfulResultAt(4, 44, "MyNamespace.MyDelegate", "a"));
        }

        [Fact]
        public void CSharp_MemberMisspelled_EmitsDiagnostic()
        {

            VerifyCSharp(
                MisspelledMemberCode,
                GetCA1704MemberResultAt(4, 17, "Mathod", "Program.SomeMathod()"),
                GetCA1704MemberResultAt(8, 19, "Naem", "Program.Naem"));
        }

        [Fact]
        public void CSharp_MemberUnmeaningful_EmitsDiagnostic()
        {
            VerifyCSharp(
                @"
public class Program
{
    public void A()
    {
    }

    public int B(string name, int age) => 0;
}",
                GetCA1704MemberUnmeaningfulResultAt(4, 17, "A"),
                GetCA1704MemberUnmeaningfulResultAt(8, 16, "B"));
        }

        [Fact]
        public void CSharp_MemberParameterMisspelled_EmitsDiagnostic()
        {
            VerifyCSharp(
                MisspelledMemberParameterCode,
                GetCA1704MemberParameterResultAt(4, 31, "Program.Method(string)", "Naem", "firstNaem"));
        }

        [Fact]
        public void CSharp_MemberParameterUnmeaningful_EmitsDiagnostic()
        {
            VerifyCSharp(@"
public class Program
{
    public void Method(string a)
    {
    }

    public string this[int i] => null;
}",
                GetCA1704MemberParameterUnmeaningfulResultAt(4, 31, "Program.Method(string)", "a"),
                GetCA1704MemberParameterUnmeaningfulResultAt(8, 28, "Program.this[int]", "i"));
        }

        [Fact]
        public void CSharp_MethodTypeParameterMisspelled_EmitsDiagnostic()
        {
            VerifyCSharp(
                MisspelledMethodTypeParameterCode,
                GetCA1704MethodTypeParameterResultAt(4, 24, "Program.Method<TTipe>(TTipe)", "Tipe", "TTipe"));
        }

        [Fact]
        public void CSharp_MethodTypeParameterUnmeaningful_EmitsDiagnostic()
        {
            VerifyCSharp(@"
public class Program
{
    public void Method<TA>(TA parameter)
    {
    }
}",
                GetCA1704MethodTypeParameterUnmeaningfulResultAt(4, 24, "Program.Method<TA>(TA)", "TA"));
        }

        [Fact]
        public void CSharp_NamespaceMisspelled_EmitsDiagnostic()
        {
            VerifyCSharp(
                MisspelledNamespaceCode,
                GetCA1704NamespaceResultAt(2, 17, "Sumthing", "Tests.DoSumthing"));
        }

        [Fact]
        public void CSharp_NamespaceUnmeaningful_EmitsDiagnostic()
        {
            VerifyCSharp(@"
namespace A
{
    public class Program
    {
    }
}",
                GetCA1704NamespaceUnmeaningfulResultAt(2, 11, "A"));
        }

        [Fact]
        public void CSharp_TypesMisspelled_EmitsDiagnostics()
        {
            VerifyCSharp(
                MisspelledTypesCode,
                GetCA1704TypeResultAt(4, 18, "Clazz", "MyNamespace.MyClazz"),
                GetCA1704TypeResultAt(8, 19, "Stroct", "MyNamespace.MyStroct"),
                GetCA1704TypeResultAt(12, 17, "Mii", "MyNamespace.MiiEnumeration"),
                GetCA1704TypeResultAt(16, 22, "Iface", "MyNamespace.IIface"),
                GetCA1704TypeResultAt(20, 25, "Delegete", "MyNamespace.MyDelegete"));
        }

        [Fact]
        public void CSharp_TypesUnmeaningful_EmitsDiagnostics()
        {
            VerifyCSharp(@"
namespace MyNamespace
{
    public class A
    {
    }

    public struct B
    {
    }

    public enum C
    {
    }

    public interface ID
    {
    }

    public delegate int E();
}",
                GetCA1704TypeUnmeaningfulResultAt(4, 18, "A"),
                GetCA1704TypeUnmeaningfulResultAt(8, 19, "B"),
                GetCA1704TypeUnmeaningfulResultAt(12, 17, "C"),
                GetCA1704TypeUnmeaningfulResultAt(16, 22, "D"),
                GetCA1704TypeUnmeaningfulResultAt(20, 25, "E"));
        }

        [Fact]
        public void CSharp_TypeTypeParametersMisspelled_EmitsDiagnostics()
        {
            VerifyCSharp(
                MisspelledTypeTypeParameterCode,
                GetCA1704TypeTypeParameterResultAt(4, 36, "MyNamespace.MyClass<TCorrect, TWroong>", "Wroong", "TWroong"),
                GetCA1704TypeTypeParameterResultAt(8, 31, "MyNamespace.MyStructure<TWroong>", "Wroong", "TWroong"),
                GetCA1704TypeTypeParameterResultAt(12, 33, "MyNamespace.IInterface<TWroong>", "Wroong", "TWroong"),
                GetCA1704TypeTypeParameterResultAt(16, 36, "MyNamespace.MyDelegate<TWroong>", "Wroong", "TWroong"));
        }

        [Fact]
        public void CSharp_TypeTypesUnmeaningful_EmitsDiagnostics()
        {
            VerifyCSharp(@"
namespace MyNamespace
{
    public class MyClass<A>
    {
    }

    public struct MyStructure<B>
    {
    }

    public interface IInterface<C>
    {
    }

    public delegate int MyDelegate<D>();
}",
                GetCA1704TypeTypeParameterUnmeaningfulResultAt(4, 26, "MyNamespace.MyClass<A>", "A"),
                GetCA1704TypeTypeParameterUnmeaningfulResultAt(8, 31, "MyNamespace.MyStructure<B>", "B"),
                GetCA1704TypeTypeParameterUnmeaningfulResultAt(12, 33, "MyNamespace.IInterface<C>", "C"),
                GetCA1704TypeTypeParameterUnmeaningfulResultAt(16, 36, "MyNamespace.MyDelegate<D>", "D"));
        }

        private static FileAndSource CreateCustomXmlDictionary(IEnumerable<string> recognizedWords, IEnumerable<string> unrecognizedWords = null) =>
            CreateCustomXmlDictionary("CodeAnalysisDictionary.xml", recognizedWords, unrecognizedWords);

        private static FileAndSource CreateCustomXmlDictionary(string filename, IEnumerable<string> recognizedWords, IEnumerable<string> unrecognizedWords = null)
        {
            var contents = $@"<?xml version=""1.0"" encoding=""utf-8""?>
<Dictionary>
    <Words>
        <Recognized>{CreateXml(recognizedWords)}</Recognized>
        <Unrecognized>{CreateXml(unrecognizedWords)}</Unrecognized>
    </Words>
</Dictionary>";

            return new FileAndSource { FilePath = filename, Source = contents };
            static string CreateXml(IEnumerable<string> words) =>
                string.Join(Environment.NewLine, words?.Select(x => $"<Word>{x}</Word>") ?? Enumerable.Empty<string>());
        }

        private static FileAndSource CreateCustomDicDictionary(IEnumerable<string> recognizedWords)
        {
            var contents = string.Join(Environment.NewLine, recognizedWords);
            return new FileAndSource { FilePath = "CustomDictionary.dic", Source = contents };
        }

        protected override DiagnosticAnalyzer GetBasicDiagnosticAnalyzer() => new BasicIdentifiersShouldBeSpelledCorrectlyAnalyzer();

        protected override DiagnosticAnalyzer GetCSharpDiagnosticAnalyzer() => new CSharpIdentifiersShouldBeSpelledCorrectlyAnalyzer();

        private void VerifyCSharp(string source, string testProjectName, params DiagnosticResult[] expected)
        {
            Verify(source, LanguageNames.CSharp, GetCSharpDiagnosticAnalyzer(), testProjectName, null, expected);
        }

        private void VerifyCSharp(string source, IEnumerable<FileAndSource> additionalFiles, params DiagnosticResult[] expected)
        {
            TestAdditionalDocument ConvertAdditionalFile(FileAndSource additionalFile) => GetAdditionalTextFile(additionalFile.FilePath, additionalFile.Source);
            Verify(source, LanguageNames.CSharp, GetCSharpDiagnosticAnalyzer(), "Test", additionalFiles.Select(ConvertAdditionalFile), expected);
        }

        private void Verify(string source, string language, DiagnosticAnalyzer analyzer, string testProjectName, IEnumerable<TestAdditionalDocument> additionalFiles, DiagnosticResult[] expected)
        {
            var sources = new[] { source };
            var diagnostics = GetSortedDiagnostics(
                sources.ToFileAndSource(),
                language,
                analyzer,
                compilationOptions: null,
                parseOptions: null,
                referenceFlags: ReferenceFlags.None,
                projectName: testProjectName,
                additionalFiles: additionalFiles);
            diagnostics.Verify(analyzer, GetDefaultPath(language), expected);
        }

        private DiagnosticResult GetCA1704AssemblyResultAt(string misspelling, string assemblyName)
        {
            return GetCA1704ResultAt(string.Format(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageAssembly, misspelling, assemblyName));
        }

        private DiagnosticResult GetCA1704AssemblyUnmeaningfulResultAt(string assemblyName)
        {
            return GetCA1704ResultAt(string.Format(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageAssemblyMoreMeaningfulName, assemblyName));
        }

        private DiagnosticResult GetCA1704DelegateParameterResultAt(int line, int column, string delegateName, string misspelling, string parameterName)
        {
            return GetCA1704ResultAt(line, column, string.Format(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageDelegateParameter, delegateName, misspelling, parameterName));
        }

        private DiagnosticResult GetCA1704DelegateParameterUnmeaningfulResultAt(int line, int column, string delegateName, string parameterName)
        {
            return GetCA1704ResultAt(line, column, string.Format(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageDelegateParameterMoreMeaningfulName, delegateName, parameterName));
        }

        private DiagnosticResult GetCA1704MemberResultAt(int line, int column, string misspelling, string memberName)
        {
            return GetCA1704ResultAt(line, column, string.Format(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageMember, misspelling, memberName));
        }

        private DiagnosticResult GetCA1704MemberUnmeaningfulResultAt(int line, int column, string memberName)
        {
            return GetCA1704ResultAt(line, column, string.Format(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageMemberMoreMeaningfulName, memberName));
        }

        private DiagnosticResult GetCA1704MemberParameterResultAt(int line, int column, string memberName, string misspelling, string parameterName)
        {
            return GetCA1704ResultAt(line, column, string.Format(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageMemberParameter, memberName, misspelling, parameterName));
        }

        private DiagnosticResult GetCA1704MemberParameterUnmeaningfulResultAt(int line, int column, string memberName, string parameterName)
        {
            return GetCA1704ResultAt(line, column, string.Format(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageMemberParameterMoreMeaningfulName, memberName, parameterName));
        }

        private DiagnosticResult GetCA1704MethodTypeParameterResultAt(int line, int column, string memberName, string misspelling, string parameterName)
        {
            return GetCA1704ResultAt(line, column, string.Format(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageMethodTypeParameter, memberName, misspelling, parameterName));
        }

        private DiagnosticResult GetCA1704MethodTypeParameterUnmeaningfulResultAt(int line, int column, string memberName, string parameterName)
        {
            return GetCA1704ResultAt(line, column, string.Format(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageMethodTypeParameterMoreMeaningfulName, memberName, parameterName));
        }

        private DiagnosticResult GetCA1704NamespaceResultAt(int line, int column, string misspelling, string namespaceName)
        {
            return GetCA1704ResultAt(line, column, string.Format(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageNamespace, misspelling, namespaceName));
        }

        private DiagnosticResult GetCA1704NamespaceUnmeaningfulResultAt(int line, int column, string namespaceName)
        {
            return GetCA1704ResultAt(line, column, string.Format(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageNamespaceMoreMeaningfulName, namespaceName));
        }

        private DiagnosticResult GetCA1704TypeResultAt(int line, int column, string misspelling, string typeName)
        {
            return GetCA1704ResultAt(line, column, string.Format(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageType, misspelling, typeName));
        }

        private DiagnosticResult GetCA1704TypeUnmeaningfulResultAt(int line, int column, string namespaceName)
        {
            return GetCA1704ResultAt(line, column, string.Format(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageTypeMoreMeaningfulName, namespaceName));
        }

        private DiagnosticResult GetCA1704TypeTypeParameterResultAt(int line, int column, string typeName, string misspelling, string genericTypeName)
        {
            return GetCA1704ResultAt(line, column, string.Format(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageTypeTypeParameter, typeName, misspelling, genericTypeName));
        }

        private DiagnosticResult GetCA1704TypeTypeParameterUnmeaningfulResultAt(int line, int column, string typeName, string genericTypeName)
        {
            return GetCA1704ResultAt(line, column, string.Format(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageTypeTypeParameterMoreMeaningfulName, typeName, genericTypeName));
        }

        private DiagnosticResult GetCA1704ResultAt(string message)
        {
            return GetCSharpResultAt(IdentifiersShouldBeSpelledCorrectlyAnalyzer.RuleId, message);
        }

        private DiagnosticResult GetCA1704ResultAt(int line, int column, string message)
        {
            return GetCSharpResultAt(line, column, IdentifiersShouldBeSpelledCorrectlyAnalyzer.RuleId, message);
        }

        public enum DictionaryType
        {
            Xml,
            Dic
        }
    }
}