Imports System.Diagnostics

''' <summary>
'''     Basic implementation of iLogic rule
'''     Other rules MUST inherit this class
'''     as base class
''' </summary>
''' <remarks></remarks>
Public MustInherit Class RuleBase
    ''' <summary>Provides a logger that can be used in iLogic rules.</summary>
    ''' <remarks>
    '''     In a rule, this interface Is implemented by the predefined object named Logger.
    ''' </remarks>
    Public Logger As IRuleLogger = Nothing

    ''' <summary>Properties And functions for features (in a part Or assembly).</summary>
    ''' <remarks>
    '''     In a rule, this interface Is implemented by the predefined object named Feature.
    ''' </remarks>
    Public Feature As ICadFeature = Nothing

    ''' <summary>Provides functions properties to read And write data from Excel.</summary>
    ''' <remarks>
    '''     In a rule, this interface Is implemented by the predefined object named GoExcel.
    ''' </remarks>
    Public GoExcel As IGoExcel = Nothing

    ''' <summary>
    '''     Provides functions to find And change the current row of an iFeature Or sheet metal punch tool in a part.
    ''' </summary>
    ''' <remarks>
    '''     In a rule, this interface Is implemented by the predefined object named iFeature.
    ''' </remarks>
    Public iFeature As IiFeatureRowChanger = Nothing

    ''' <summary>Provides functions to show (launch) predefined iLogic forms.</summary>
    ''' <remarks>
    '''     In a rule, this interface Is implemented by the predefined object named iLogicForm.
    ''' </remarks>
    Public iLogicForm As IiLogicForm = Nothing

    ''' <summary>
    '''     Provides properties And functions for access to the Inventor API, running other rules, And miscellaneous functions.
    ''' </summary>
    ''' <remarks>
    '''     In a rule, this interface Is implemented by two predefined object identifiers: iLogicVb And InventorVb. They are
    '''     two references to the same object.
    ''' </remarks>
    Public iLogicVb As ILowLevelSupport = Nothing

    ''' <summary>
    '''     Provides properties And functions for access to the Inventor API, running other rules, And miscellaneous functions.
    ''' </summary>
    ''' <remarks>
    '''     In a rule, this interface Is implemented by two predefined object identifiers: iLogicVb And InventorVb. They are
    '''     two references to the same object.
    ''' </remarks>
    Public InventorVb As ILowLevelSupport = Nothing

    ''' <summary>
    '''     Provides functions to find And change the current row of an iPart Or iAssenbly component.
    ''' </summary>
    ''' <remarks>
    '''     In a rule, this interface Is implemented by the predefined object named iPart And also iAssembly. These are two
    '''     references to the same object.
    ''' </remarks>
    Public iPart As IiPartRowChanger = Nothing

    ''' <summary>
    '''     Provides functions to find And change the current row of an iPart Or iAssenbly component.
    ''' </summary>
    ''' <remarks>
    '''     In a rule, this interface Is implemented by the predefined object named iPart And also iAssembly. These are two
    '''     references to the same object.
    ''' </remarks>
    Public iAssembly As IiPartRowChanger = Nothing

    ''' <summary>
    '''     Provides properties to get And set iProperty values And physical properties (in a part Or assembly).
    ''' </summary>
    ''' <remarks>
    '''     In a rule, this interface Is implemented by the predefined object named iProperties.
    ''' </remarks>
    Public iProperties As IiProperties = Nothing

    ''' <summary>
    '''     Provides functions for a measuring distance, angle, And area (in a part Or assembly).
    ''' </summary>
    ''' <remarks>
    '''     In a rule, this interface Is implemented by the predefined object named Measure.
    ''' </remarks>
    Public Measure As ICadMeasure = Nothing

    ''' <summary>
    '''     Provides properties And functions for multivalue lists (lists of choices of parameter expressions).
    ''' </summary>
    ''' <remarks>
    '''     In a rule, this interface Is implemented by the predefined object named MultiValue.
    ''' </remarks>
    Public MultiValue As IMultiValueParam = Nothing

    ''' <summary>Provides properties to get And set the values of Inventor parameters.</summary>
    ''' <remarks>
    '''     In a rule, this interface Is implemented by the predefined object named Parameter.
    ''' </remarks>
    Public Parameter As IParamDynamic = Nothing

    ''' <summary>
    '''     Provides access to context information provided via the RunRule family of automation methods.
    ''' </summary>
    ''' <remarks>
    '''     In a rule, this interface Is implemented by the predefined object named RuleArguments.
    ''' </remarks>
    Public RuleArguments As IRuleArguments = Nothing

    ''' <summary>
    '''     Provides properties And functions for temporary objects that can be shared between iLogic rules in an Inventor
    '''     session.
    ''' </summary>
    ''' <remarks>
    '''     In a rule, this interface Is implemented by the predefined object named SharedVariable.
    ''' </remarks>
    Public SharedVariable As ISharedVariable = Nothing

    ''' <summary>Provides functions And properties for sheet metal parts.</summary>
    ''' <remarks>
    '''     In a rule, this interface Is implemented by the predefined object named SheetMetal.
    ''' </remarks>
    Public SheetMetal As ISheetMetal = Nothing

    ''' <summary>
    '''     Returns the top-level parent application object.
    '''     When used the context of Inventor, an Application object Is returned.
    '''     When used in the context of Apprentice, an ApprenticeServer object Is returned.
    ''' </summary>
    Public ThisApplication As Application = Nothing

    ''' <summary>InventorServer Object.</summary>
    Public ThisServer As InventorServerObject = Nothing

    ''' <summary>
    '''     Provides properties And functions for access to the Inventor document in which the rule Is running.
    ''' </summary>
    ''' <remarks>
    '''     In a rule, this interface Is implemented by the predefined object named ThisDoc.
    ''' </remarks>
    Public ThisDoc As ICadDoc = Nothing

    ''' <summary>Provides functions for access to the drawing in which the rule Is running.</summary>
    ''' <remarks>
    '''     In a rule, this interface Is implemented by the predefined object named ThisDrawing.
    '''     ActiveSheet will be available as a property of the rule class (so that it can be used as a standalone object), if
    '''     used in the rule.
    ''' </remarks>
    Public ThisDrawing As ICadDrawing = Nothing '.IManagedDrawing = Nothing

    ''' <summary>Provides functions And properties for a drawing sheet.</summary>
    ''' <remarks>
    ''' </remarks>
    Public ActiveSheet As ICadDrawingSheet = Nothing '.IManagedSheet = Nothing

    ''' <summary>
    '''     Gets and sets the name of the language to be used in searches for asset names. Default: Nothing.
    '''     Possible values: "chs", "cht", "csy", "deu", "eng", "esp", "fra", "hun", "ita", "jpn", "kor", "plk", "ptb", or
    '''     "rus".
    ''' </summary>
    Public AssetNameLanguage As String = Nothing

    ''' <summary>Provides properties And functions for assembly component occurrences.</summary>
    ''' <remarks>
    '''     In a rule, this interface Is implemented by the predefined object named Component.
    ''' </remarks>
    Public Component As ICadComponent = Nothing

    ''' <summary>
    '''     Provides functions And properties for adding, modifying, deleting, And managing assembly components.
    ''' </summary>
    ''' <remarks>
    '''     In a rule, this interface Is implemented by the predefined object named Components.
    ''' </remarks>
    Public Components As IManagedComponents = Nothing

    ''' <summary>Provides properties And functions for assembly constraints.</summary>
    ''' <remarks>
    '''     In a rule, this interface Is implemented by the predefined object named Constraint.
    ''' </remarks>
    Public Constraint As IAssemConstraint = Nothing

    ''' <summary>Provides functions to add, edit, And delete assembly constraints.</summary>
    ''' <remarks>
    '''     In a rule, this interface Is implemented by the predefined object named Constraints.
    ''' </remarks>
    Public Constraints As IManagedConstraints = Nothing

    ''' <summary>Provides properties And functions for assembly joints.</summary>
    ''' <remarks>
    '''     In a rule, this interface Is implemented by the predefined object named Joint.
    ''' </remarks>
    Public Joint As IAssemJoint = Nothing

    ''' <summary>
    '''     Provides functions And properties for accessing assembly in which the rule Is running
    ''' </summary>
    ''' <remarks>
    '''     In a rule, this interface Is implemented by the predefined object named ThisAssembly.
    ''' </remarks>
    Public ThisAssembly As IManagedAssembly = Nothing

    ''' <summary>Provides functions for access to the assembly BOM.</summary>
    ''' <remarks>
    '''     In a rule, this interface Is implemented by the predefined object named ThisBom.
    ''' </remarks>
    Public ThisBOM As ICadBom = Nothing

    ''' <summary>Provides functions And properties for managing occurrence patterns.</summary>
    ''' <remarks>
    '''     In a rule, this interface Is implemented by the predefined object named Patterns.
    ''' </remarks>
    Public Patterns As IManagedPatterns = Nothing

    ''' <summary>
    '''     Main entry point for a rule
    ''' </summary>
    ''' <remarks></remarks>
    MustOverride Sub Main()

    ''' <summary>
    '''     Custom implementation of default Break function
    ''' </summary>
    Public Sub Break()
        Debugger.Launch()
        'Debugger.Break()
    End Sub
End Class


