// Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using Analyzer.Utilities;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Diagnostics;

namespace Text.Analyzers
{
    /// <summary>
    /// CA1704: Identifiers should be spelled correctly
    /// </summary>
    public abstract class IdentifiersShouldBeSpelledCorrectlyAnalyzer : DiagnosticAnalyzer
    {
        internal const string RuleId = "CA1704";

        private static readonly LocalizableString s_localizableTitle = new LocalizableResourceString(nameof(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyTitle), TextAnalyzersResources.ResourceManager, typeof(TextAnalyzersResources));

        private static readonly LocalizableString s_localizableMessageAssembly = new LocalizableResourceString(nameof(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageAssembly), TextAnalyzersResources.ResourceManager, typeof(TextAnalyzersResources));
        private static readonly LocalizableString s_localizableMessageNamespace = new LocalizableResourceString(nameof(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageNamespace), TextAnalyzersResources.ResourceManager, typeof(TextAnalyzersResources));
        private static readonly LocalizableString s_localizableMessageType = new LocalizableResourceString(nameof(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageType), TextAnalyzersResources.ResourceManager, typeof(TextAnalyzersResources));
        private static readonly LocalizableString s_localizableMessageMember = new LocalizableResourceString(nameof(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageMember), TextAnalyzersResources.ResourceManager, typeof(TextAnalyzersResources));
        private static readonly LocalizableString s_localizableMessageMemberParameter = new LocalizableResourceString(nameof(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageMemberParameter), TextAnalyzersResources.ResourceManager, typeof(TextAnalyzersResources));
        private static readonly LocalizableString s_localizableMessageDelegateParameter = new LocalizableResourceString(nameof(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageDelegateParameter), TextAnalyzersResources.ResourceManager, typeof(TextAnalyzersResources));
        private static readonly LocalizableString s_localizableMessageTypeTypeParameter = new LocalizableResourceString(nameof(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageTypeTypeParameter), TextAnalyzersResources.ResourceManager, typeof(TextAnalyzersResources));
        private static readonly LocalizableString s_localizableMessageMethodTypeParameter = new LocalizableResourceString(nameof(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageMethodTypeParameter), TextAnalyzersResources.ResourceManager, typeof(TextAnalyzersResources));
        private static readonly LocalizableString s_localizableMessageAssemblyMoreMeaningfulName = new LocalizableResourceString(nameof(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageAssemblyMoreMeaningfulName), TextAnalyzersResources.ResourceManager, typeof(TextAnalyzersResources));
        private static readonly LocalizableString s_localizableMessageNamespaceMoreMeaningfulName = new LocalizableResourceString(nameof(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageNamespaceMoreMeaningfulName), TextAnalyzersResources.ResourceManager, typeof(TextAnalyzersResources));
        private static readonly LocalizableString s_localizableMessageTypeMoreMeaningfulName = new LocalizableResourceString(nameof(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageTypeMoreMeaningfulName), TextAnalyzersResources.ResourceManager, typeof(TextAnalyzersResources));
        private static readonly LocalizableString s_localizableMessageMemberMoreMeaningfulName = new LocalizableResourceString(nameof(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageMemberMoreMeaningfulName), TextAnalyzersResources.ResourceManager, typeof(TextAnalyzersResources));
        private static readonly LocalizableString s_localizableMessageMemberParameterMoreMeaningfulName = new LocalizableResourceString(nameof(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageMemberParameterMoreMeaningfulName), TextAnalyzersResources.ResourceManager, typeof(TextAnalyzersResources));
        private static readonly LocalizableString s_localizableMessageDelegateParameterMoreMeaningfulName = new LocalizableResourceString(nameof(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageDelegateParameterMoreMeaningfulName), TextAnalyzersResources.ResourceManager, typeof(TextAnalyzersResources));
        private static readonly LocalizableString s_localizableMessageTypeTypeParameterMoreMeaningfulName = new LocalizableResourceString(nameof(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageTypeTypeParameterMoreMeaningfulName), TextAnalyzersResources.ResourceManager, typeof(TextAnalyzersResources));
        private static readonly LocalizableString s_localizableMessageMethodTypeParameterMoreMeaningfulName = new LocalizableResourceString(nameof(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyMessageMethodTypeParameterMoreMeaningfulName), TextAnalyzersResources.ResourceManager, typeof(TextAnalyzersResources));
        private static readonly LocalizableString s_localizableDescription = new LocalizableResourceString(nameof(TextAnalyzersResources.IdentifiersShouldBeSpelledCorrectlyDescription), TextAnalyzersResources.ResourceManager, typeof(TextAnalyzersResources));

        internal static DiagnosticDescriptor AssemblyRule = new DiagnosticDescriptor(RuleId,
                                                                             s_localizableTitle,
                                                                             s_localizableMessageAssembly,
                                                                             DiagnosticCategory.Naming,
                                                                             DiagnosticHelpers.DefaultDiagnosticSeverity,
                                                                             isEnabledByDefault: false,
                                                                             description: s_localizableDescription,
                                                                             helpLinkUri: "https://docs.microsoft.com/visualstudio/code-quality/ca1704-identifiers-should-be-spelled-correctly",
                                                                             customTags: FxCopWellKnownDiagnosticTags.PortedFxCopRule);
        internal static DiagnosticDescriptor NamespaceRule = new DiagnosticDescriptor(RuleId,
                                                                             s_localizableTitle,
                                                                             s_localizableMessageNamespace,
                                                                             DiagnosticCategory.Naming,
                                                                             DiagnosticHelpers.DefaultDiagnosticSeverity,
                                                                             isEnabledByDefault: false,
                                                                             description: s_localizableDescription,
                                                                             helpLinkUri: "https://docs.microsoft.com/visualstudio/code-quality/ca1704-identifiers-should-be-spelled-correctly",
                                                                             customTags: FxCopWellKnownDiagnosticTags.PortedFxCopRule);
        internal static DiagnosticDescriptor TypeRule = new DiagnosticDescriptor(RuleId,
                                                                             s_localizableTitle,
                                                                             s_localizableMessageType,
                                                                             DiagnosticCategory.Naming,
                                                                             DiagnosticHelpers.DefaultDiagnosticSeverity,
                                                                             isEnabledByDefault: false,
                                                                             description: s_localizableDescription,
                                                                             helpLinkUri: "https://docs.microsoft.com/visualstudio/code-quality/ca1704-identifiers-should-be-spelled-correctly",
                                                                             customTags: FxCopWellKnownDiagnosticTags.PortedFxCopRule);
        internal static DiagnosticDescriptor MemberRule = new DiagnosticDescriptor(RuleId,
                                                                             s_localizableTitle,
                                                                             s_localizableMessageMember,
                                                                             DiagnosticCategory.Naming,
                                                                             DiagnosticHelpers.DefaultDiagnosticSeverity,
                                                                             isEnabledByDefault: false,
                                                                             description: s_localizableDescription,
                                                                             helpLinkUri: "https://docs.microsoft.com/visualstudio/code-quality/ca1704-identifiers-should-be-spelled-correctly",
                                                                             customTags: FxCopWellKnownDiagnosticTags.PortedFxCopRule);
        internal static DiagnosticDescriptor MemberParameterRule = new DiagnosticDescriptor(RuleId,
                                                                             s_localizableTitle,
                                                                             s_localizableMessageMemberParameter,
                                                                             DiagnosticCategory.Naming,
                                                                             DiagnosticHelpers.DefaultDiagnosticSeverity,
                                                                             isEnabledByDefault: false,
                                                                             description: s_localizableDescription,
                                                                             helpLinkUri: "https://docs.microsoft.com/visualstudio/code-quality/ca1704-identifiers-should-be-spelled-correctly",
                                                                             customTags: FxCopWellKnownDiagnosticTags.PortedFxCopRule);
        internal static DiagnosticDescriptor DelegateParameterRule = new DiagnosticDescriptor(RuleId,
                                                                             s_localizableTitle,
                                                                             s_localizableMessageDelegateParameter,
                                                                             DiagnosticCategory.Naming,
                                                                             DiagnosticHelpers.DefaultDiagnosticSeverity,
                                                                             isEnabledByDefault: false,
                                                                             description: s_localizableDescription,
                                                                             helpLinkUri: "https://docs.microsoft.com/visualstudio/code-quality/ca1704-identifiers-should-be-spelled-correctly",
                                                                             customTags: FxCopWellKnownDiagnosticTags.PortedFxCopRule);
        internal static DiagnosticDescriptor TypeTypeParameterRule = new DiagnosticDescriptor(RuleId,
                                                                             s_localizableTitle,
                                                                             s_localizableMessageTypeTypeParameter,
                                                                             DiagnosticCategory.Naming,
                                                                             DiagnosticHelpers.DefaultDiagnosticSeverity,
                                                                             isEnabledByDefault: false,
                                                                             description: s_localizableDescription,
                                                                             helpLinkUri: "https://docs.microsoft.com/visualstudio/code-quality/ca1704-identifiers-should-be-spelled-correctly",
                                                                             customTags: FxCopWellKnownDiagnosticTags.PortedFxCopRule);
        internal static DiagnosticDescriptor MethodTypeParameterRule = new DiagnosticDescriptor(RuleId,
                                                                             s_localizableTitle,
                                                                             s_localizableMessageMethodTypeParameter,
                                                                             DiagnosticCategory.Naming,
                                                                             DiagnosticHelpers.DefaultDiagnosticSeverity,
                                                                             isEnabledByDefault: false,
                                                                             description: s_localizableDescription,
                                                                             helpLinkUri: "https://docs.microsoft.com/visualstudio/code-quality/ca1704-identifiers-should-be-spelled-correctly",
                                                                             customTags: FxCopWellKnownDiagnosticTags.PortedFxCopRule);
        internal static DiagnosticDescriptor AssemblyMoreMeaningfulNameRule = new DiagnosticDescriptor(RuleId,
                                                                             s_localizableTitle,
                                                                             s_localizableMessageAssemblyMoreMeaningfulName,
                                                                             DiagnosticCategory.Naming,
                                                                             DiagnosticHelpers.DefaultDiagnosticSeverity,
                                                                             isEnabledByDefault: false,
                                                                             description: s_localizableDescription,
                                                                             helpLinkUri: "https://docs.microsoft.com/visualstudio/code-quality/ca1704-identifiers-should-be-spelled-correctly",
                                                                             customTags: FxCopWellKnownDiagnosticTags.PortedFxCopRule);
        internal static DiagnosticDescriptor NamespaceMoreMeaningfulNameRule = new DiagnosticDescriptor(RuleId,
                                                                             s_localizableTitle,
                                                                             s_localizableMessageNamespaceMoreMeaningfulName,
                                                                             DiagnosticCategory.Naming,
                                                                             DiagnosticHelpers.DefaultDiagnosticSeverity,
                                                                             isEnabledByDefault: false,
                                                                             description: s_localizableDescription,
                                                                             helpLinkUri: "https://docs.microsoft.com/visualstudio/code-quality/ca1704-identifiers-should-be-spelled-correctly",
                                                                             customTags: FxCopWellKnownDiagnosticTags.PortedFxCopRule);
        internal static DiagnosticDescriptor TypeMoreMeaningfulNameRule = new DiagnosticDescriptor(RuleId,
                                                                             s_localizableTitle,
                                                                             s_localizableMessageTypeMoreMeaningfulName,
                                                                             DiagnosticCategory.Naming,
                                                                             DiagnosticHelpers.DefaultDiagnosticSeverity,
                                                                             isEnabledByDefault: false,
                                                                             description: s_localizableDescription,
                                                                             helpLinkUri: "https://docs.microsoft.com/visualstudio/code-quality/ca1704-identifiers-should-be-spelled-correctly",
                                                                             customTags: FxCopWellKnownDiagnosticTags.PortedFxCopRule);
        internal static DiagnosticDescriptor MemberMoreMeaningfulNameRule = new DiagnosticDescriptor(RuleId,
                                                                             s_localizableTitle,
                                                                             s_localizableMessageMemberMoreMeaningfulName,
                                                                             DiagnosticCategory.Naming,
                                                                             DiagnosticHelpers.DefaultDiagnosticSeverity,
                                                                             isEnabledByDefault: false,
                                                                             description: s_localizableDescription,
                                                                             helpLinkUri: "https://docs.microsoft.com/visualstudio/code-quality/ca1704-identifiers-should-be-spelled-correctly",
                                                                             customTags: FxCopWellKnownDiagnosticTags.PortedFxCopRule);
        internal static DiagnosticDescriptor MemberParameterMoreMeaningfulNameRule = new DiagnosticDescriptor(RuleId,
                                                                             s_localizableTitle,
                                                                             s_localizableMessageMemberParameterMoreMeaningfulName,
                                                                             DiagnosticCategory.Naming,
                                                                             DiagnosticHelpers.DefaultDiagnosticSeverity,
                                                                             isEnabledByDefault: false,
                                                                             description: s_localizableDescription,
                                                                             helpLinkUri: "https://docs.microsoft.com/visualstudio/code-quality/ca1704-identifiers-should-be-spelled-correctly",
                                                                             customTags: FxCopWellKnownDiagnosticTags.PortedFxCopRule);
        internal static DiagnosticDescriptor DelegateParameterMoreMeaningfulNameRule = new DiagnosticDescriptor(RuleId,
                                                                             s_localizableTitle,
                                                                             s_localizableMessageDelegateParameterMoreMeaningfulName,
                                                                             DiagnosticCategory.Naming,
                                                                             DiagnosticHelpers.DefaultDiagnosticSeverity,
                                                                             isEnabledByDefault: false,
                                                                             description: s_localizableDescription,
                                                                             helpLinkUri: "https://docs.microsoft.com/visualstudio/code-quality/ca1704-identifiers-should-be-spelled-correctly",
                                                                             customTags: FxCopWellKnownDiagnosticTags.PortedFxCopRule);
        internal static DiagnosticDescriptor TypeTypeParameterMoreMeaningfulNameRule = new DiagnosticDescriptor(RuleId,
                                                                             s_localizableTitle,
                                                                             s_localizableMessageTypeTypeParameterMoreMeaningfulName,
                                                                             DiagnosticCategory.Naming,
                                                                             DiagnosticHelpers.DefaultDiagnosticSeverity,
                                                                             isEnabledByDefault: false,
                                                                             description: s_localizableDescription,
                                                                             helpLinkUri: "https://docs.microsoft.com/visualstudio/code-quality/ca1704-identifiers-should-be-spelled-correctly",
                                                                             customTags: FxCopWellKnownDiagnosticTags.PortedFxCopRule);
        internal static DiagnosticDescriptor MethodTypeParameterMoreMeaningfulNameRule = new DiagnosticDescriptor(RuleId,
                                                                             s_localizableTitle,
                                                                             s_localizableMessageMethodTypeParameterMoreMeaningfulName,
                                                                             DiagnosticCategory.Naming,
                                                                             DiagnosticHelpers.DefaultDiagnosticSeverity,
                                                                             isEnabledByDefault: false,
                                                                             description: s_localizableDescription,
                                                                             helpLinkUri: "https://docs.microsoft.com/visualstudio/code-quality/ca1704-identifiers-should-be-spelled-correctly",
                                                                             customTags: FxCopWellKnownDiagnosticTags.PortedFxCopRule);

        private static readonly SymbolKind[] SymbolsToCheck = new SymbolKind[]
        {
            SymbolKind.Namespace,
            SymbolKind.NamedType,
            SymbolKind.Method,
            SymbolKind.Property,
            SymbolKind.Event,
            SymbolKind.Field,
            SymbolKind.Parameter,
        };

        private NamingService _namingService;

        public override ImmutableArray<DiagnosticDescriptor> SupportedDiagnostics => ImmutableArray.Create(
            AssemblyRule,
            NamespaceRule,
            TypeRule,
            MemberRule,
            MemberParameterRule,
            DelegateParameterRule,
            TypeTypeParameterRule,
            MethodTypeParameterRule,
            AssemblyMoreMeaningfulNameRule,
            NamespaceMoreMeaningfulNameRule,
            TypeMoreMeaningfulNameRule,
            MemberMoreMeaningfulNameRule,
            MemberParameterMoreMeaningfulNameRule,
            DelegateParameterMoreMeaningfulNameRule,
            TypeTypeParameterMoreMeaningfulNameRule,
            MethodTypeParameterMoreMeaningfulNameRule);

        public override void Initialize(AnalysisContext analysisContext)
        {
            analysisContext.EnableConcurrentExecution();
            analysisContext.ConfigureGeneratedCodeAnalysis(GeneratedCodeAnalysisFlags.None);

            analysisContext.RegisterCompilationAction(AnalyzeAssembly);
            analysisContext.RegisterCompilationStartAction(compilationStartContext =>
            {
                var assemblyName = compilationStartContext.Compilation.AssemblyName;
                _namingService = new NamingService(assemblyName);
                _namingService.LoadDictionaries(compilationStartContext.Options.AdditionalFiles, compilationStartContext.CancellationToken);

                compilationStartContext.RegisterSymbolAction(AnalyzeSymbol, SymbolsToCheck);
            });
        }

        private void AnalyzeAssembly(CompilationAnalysisContext context)
        {
            IAssemblySymbol assembly = context.Compilation.Assembly;
            var diagnostics = GetDiagnosticsForSymbol(assembly, assembly.Name);

            foreach (var diagnostic in diagnostics)
            {
                context.ReportDiagnostic(diagnostic);
            }
        }

        private void AnalyzeSymbol(SymbolAnalysisContext context)
        {
            var typeParameterDiagnostics = Enumerable.Empty<Diagnostic>();
            ISymbol symbol = context.Symbol;
            var symbolName = symbol.Name;

            switch (symbol)
            {
                case IMethodSymbol method:
                    if (method.MethodKind == MethodKind.PropertyGet || method.MethodKind == MethodKind.PropertySet)
                    {
                        return;
                    }
                    foreach (var typeParameter in method.TypeParameters)
                    {
                        typeParameterDiagnostics = GetDiagnosticsForSymbol(typeParameter, RemoveCommonPrefixes(typeParameter.Name));
                    }
                    break;

                case INamedTypeSymbol type:
                    if (type.TypeKind == TypeKind.Interface)
                    {
                        symbolName = RemoveCommonPrefixes(symbolName);
                    }
                    foreach (var typeParameter in type.TypeParameters)
                    {
                        typeParameterDiagnostics = GetDiagnosticsForSymbol(typeParameter, RemoveCommonPrefixes(typeParameter.Name));
                    }
                    break;
            }

            var diagnostics = GetDiagnosticsForSymbol(symbol, symbolName);
            var allDiagnostics = typeParameterDiagnostics.Concat(diagnostics);
            foreach (var diagnostic in allDiagnostics)
            {
                context.ReportDiagnostic(diagnostic);
            }
        }

        private static string RemoveCommonPrefixes(string name)
        {
            return name.StartsWith("I", StringComparison.Ordinal) || name.StartsWith("T", StringComparison.Ordinal)
                ? name.Substring(1)
                : name;
        }

        private IEnumerable<Diagnostic> GetDiagnosticsForSymbol(ISymbol symbol, string symbolName)
        {
            var diagnostics = new List<Diagnostic>();
            if (symbolName.Length == 1)
            {
                diagnostics.AddRange(GetUnmeaningfulIdentifierDiagnostics(symbol, symbolName));
            }
            else
            {
                foreach (var misspelledWord in GetMisspelledWords(symbolName))
                {
                    diagnostics.AddRange(GetMisspelledWordDiagnostics(symbol, misspelledWord));
                }
            }

            return diagnostics;
        }

        private IEnumerable<string> GetMisspelledWords(string symbolName)
        {
            if (!_namingService.SupportsSpelling)
            {
                yield break;
            }

            var parser = new WordParser(symbolName, WordParserOptions.SplitCompoundWords);
            if (parser.PeekWord() != null)
            {
                var word = parser.NextWord();

                do
                {
                    var result = _namingService.CheckSpelling(word);
                    if (result != Spelling.WordSpelling.Unrecognized)
                    {
                        continue;
                    }

                    yield return word;
                }
                while ((word = parser.NextWord()) != null);
            }
        }

        private static IEnumerable<Diagnostic> GetUnmeaningfulIdentifierDiagnostics(ISymbol symbol, string symbolName)
        {
            switch (symbol.Kind)
            {
                case SymbolKind.Assembly:
                    yield return Diagnostic.Create(AssemblyMoreMeaningfulNameRule, Location.None, symbolName);
                    break;

                case SymbolKind.Namespace:
                    yield return Diagnostic.Create(NamespaceMoreMeaningfulNameRule, symbol.Locations.First(), symbolName);
                    break;

                case SymbolKind.NamedType:
                    foreach (var location in symbol.Locations)
                    {
                        yield return Diagnostic.Create(TypeMoreMeaningfulNameRule, location, symbolName);
                    }
                    break;

                case SymbolKind.Method:
                case SymbolKind.Property:
                case SymbolKind.Event:
                case SymbolKind.Field:
                    yield return Diagnostic.Create(MemberMoreMeaningfulNameRule, symbol.Locations.First(), symbolName);
                    break;

                case SymbolKind.Parameter:
                    if (symbol.ContainingType.TypeKind == TypeKind.Delegate)
                    {
                        yield return Diagnostic.Create(DelegateParameterMoreMeaningfulNameRule, symbol.Locations.First(), symbol.ContainingType.ToDisplayString(), symbolName);
                    }
                    else
                    {
                        yield return Diagnostic.Create(MemberParameterMoreMeaningfulNameRule, symbol.Locations.First(), symbol.ContainingSymbol.ToDisplayString(), symbolName);
                    }
                    break;

                case SymbolKind.TypeParameter:
                    if (symbol.ContainingSymbol.Kind == SymbolKind.Method)
                    {
                        yield return Diagnostic.Create(MethodTypeParameterMoreMeaningfulNameRule, symbol.Locations.First(), symbol.ContainingSymbol.ToDisplayString(), symbol.Name);
                    }
                    else
                    {
                        yield return Diagnostic.Create(TypeTypeParameterMoreMeaningfulNameRule, symbol.Locations.First(), symbol.ContainingSymbol.ToDisplayString(), symbol.Name);
                    }
                    break;

                default:
                    throw new NotImplementedException($"Unknown SymbolKind: {symbol.Kind}");
            }
        }

        private static IEnumerable<Diagnostic> GetMisspelledWordDiagnostics(ISymbol symbol, string misspelledWord)
        {
            switch (symbol.Kind)
            {
                case SymbolKind.Assembly:
                    yield return Diagnostic.Create(AssemblyRule, Location.None, misspelledWord, symbol.Name);
                    break;

                case SymbolKind.Namespace:
                    yield return Diagnostic.Create(NamespaceRule, symbol.Locations.First(), misspelledWord, symbol.ToDisplayString());
                    break;

                case SymbolKind.NamedType:
                    foreach (var location in symbol.Locations)
                    {
                        yield return Diagnostic.Create(TypeRule, location, misspelledWord, symbol.ToDisplayString());
                    }
                    break;

                case SymbolKind.Method:
                case SymbolKind.Property:
                case SymbolKind.Event:
                case SymbolKind.Field:
                    yield return Diagnostic.Create(MemberRule, symbol.Locations.First(), misspelledWord, symbol.ToDisplayString());
                    break;

                case SymbolKind.Parameter:
                    if (symbol.ContainingType.TypeKind == TypeKind.Delegate)
                    {
                        yield return Diagnostic.Create(DelegateParameterRule, symbol.Locations.First(), symbol.ContainingType.ToDisplayString(), misspelledWord, symbol.Name);
                    }
                    else
                    {
                        yield return Diagnostic.Create(MemberParameterRule, symbol.Locations.First(), symbol.ContainingSymbol.ToDisplayString(), misspelledWord, symbol.Name);
                    }
                    break;

                case SymbolKind.TypeParameter:
                    if (symbol.ContainingSymbol.Kind == SymbolKind.Method)
                    {
                        yield return Diagnostic.Create(MethodTypeParameterRule, symbol.Locations.First(), symbol.ContainingSymbol.ToDisplayString(), misspelledWord, symbol.Name);
                    }
                    else
                    {
                        yield return Diagnostic.Create(TypeTypeParameterRule, symbol.Locations.First(), symbol.ContainingSymbol.ToDisplayString(), misspelledWord, symbol.Name);
                    }
                    break;

                default:
                    throw new NotImplementedException($"Unknown SymbolKind: {symbol.Kind}");
            }
        }
    }
}