'------------------------------------------------------------------------------
' <auto-generated>
'     Dieser Code wurde von einem Tool generiert.
'     Laufzeitversion:4.0.30319.42000
'
'     Änderungen an dieser Datei können falsches Verhalten verursachen und gehen verloren, wenn
'     der Code erneut generiert wird.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On

Imports System

Namespace My.Resources
    
    'Diese Klasse wurde von der StronglyTypedResourceBuilder automatisch generiert
    '-Klasse über ein Tool wie ResGen oder Visual Studio automatisch generiert.
    'Um einen Member hinzuzufügen oder zu entfernen, bearbeiten Sie die .ResX-Datei und führen dann ResGen
    'mit der /str-Option erneut aus, oder Sie erstellen Ihr VS-Projekt neu.
    '''<summary>
    '''  Eine stark typisierte Ressourcenklasse zum Suchen von lokalisierten Zeichenfolgen usw.
    '''</summary>
    <Global.System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "17.0.0.0"),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.Microsoft.VisualBasic.HideModuleNameAttribute()>  _
    Friend Module Resources
        
        Private resourceMan As Global.System.Resources.ResourceManager
        
        Private resourceCulture As Global.System.Globalization.CultureInfo
        
        '''<summary>
        '''  Gibt die zwischengespeicherte ResourceManager-Instanz zurück, die von dieser Klasse verwendet wird.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend ReadOnly Property ResourceManager() As Global.System.Resources.ResourceManager
            Get
                If Object.ReferenceEquals(resourceMan, Nothing) Then
                    Dim temp As Global.System.Resources.ResourceManager = New Global.System.Resources.ResourceManager("inoVBETools.Resources", GetType(Resources).Assembly)
                    resourceMan = temp
                End If
                Return resourceMan
            End Get
        End Property
        
        '''<summary>
        '''  Überschreibt die CurrentUICulture-Eigenschaft des aktuellen Threads für alle
        '''  Ressourcenzuordnungen, die diese stark typisierte Ressourcenklasse verwenden.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend Property Culture() As Global.System.Globalization.CultureInfo
            Get
                Return resourceCulture
            End Get
            Set
                resourceCulture = value
            End Set
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die There is a problem with the import of &apos;{0}&apos;. ähnelt.
        '''</summary>
        Friend ReadOnly Property CHM_ProblemImport() As String
            Get
                Return ResourceManager.GetString("CHM_ProblemImport", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Import Modules ähnelt.
        '''</summary>
        Friend ReadOnly Property CHMTitleImport() As String
            Get
                Return ResourceManager.GetString("CHMTitleImport", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die The module &apos;{0}&apos; is already imported. ähnelt.
        '''</summary>
        Friend ReadOnly Property CMH_ModuleImported() As String
            Get
                Return ResourceManager.GetString("CMH_ModuleImported", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Do you want to replace it? ähnelt.
        '''</summary>
        Friend ReadOnly Property CMH_Replace() As String
            Get
                Return ResourceManager.GetString("CMH_Replace", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die All existing code modules in &apos;{0}&apos; will be overwritten. ähnelt.
        '''</summary>
        Friend ReadOnly Property CMHOverwrite() As String
            Get
                Return ResourceManager.GetString("CMHOverwrite", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Select folder to export to. Just select the folder and keep the filename &apos;{0}&apos;. ähnelt.
        '''</summary>
        Friend ReadOnly Property ConnectExportTitle() As String
            Get
                Return ResourceManager.GetString("ConnectExportTitle", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Select folder to import from. Just select the folder and keep the filename &apos;{0}. ähnelt.
        '''</summary>
        Friend ReadOnly Property ConnectImportTitle() As String
            Get
                Return ResourceManager.GetString("ConnectImportTitle", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die temporary file name ähnelt.
        '''</summary>
        Friend ReadOnly Property ConnectTemporaryFileName() As String
            Get
                Return ResourceManager.GetString("ConnectTemporaryFileName", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Error in line ähnelt.
        '''</summary>
        Friend ReadOnly Property ErrorInLine() As String
            Get
                Return ResourceManager.GetString("ErrorInLine", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die in procedure ähnelt.
        '''</summary>
        Friend ReadOnly Property ErrorInProcedure() As String
            Get
                Return ResourceManager.GetString("ErrorInProcedure", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Add to stage ähnelt.
        '''</summary>
        Friend ReadOnly Property FrmButtonAddToStage() As String
            Get
                Return ResourceManager.GetString("FrmButtonAddToStage", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Cancel ähnelt.
        '''</summary>
        Friend ReadOnly Property frmButtonCancel() As String
            Get
                Return ResourceManager.GetString("frmButtonCancel", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die OK ähnelt.
        '''</summary>
        Friend ReadOnly Property frmButtonOK() As String
            Get
                Return ResourceManager.GetString("frmButtonOK", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Remove from stage ähnelt.
        '''</summary>
        Friend ReadOnly Property FrmButtonRemoveFromStage() As String
            Get
                Return ResourceManager.GetString("FrmButtonRemoveFromStage", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Git file handling ähnelt.
        '''</summary>
        Friend ReadOnly Property FrmGitCaption() As String
            Get
                Return ResourceManager.GetString("FrmGitCaption", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Current branch: {0} ähnelt.
        '''</summary>
        Friend ReadOnly Property frmGitCurrentBranch() As String
            Get
                Return ResourceManager.GetString("frmGitCurrentBranch", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Commit message ähnelt.
        '''</summary>
        Friend ReadOnly Property FrmGitLblCommitMsg() As String
            Get
                Return ResourceManager.GetString("FrmGitLblCommitMsg", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die inoVBETools settings ähnelt.
        '''</summary>
        Friend ReadOnly Property FrmOptionsCaption() As String
            Get
                Return ResourceManager.GetString("FrmOptionsCaption", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Colour of git file area ähnelt.
        '''</summary>
        Friend ReadOnly Property frmOptionsColourGit() As String
            Get
                Return ResourceManager.GetString("frmOptionsColourGit", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Create backup file prior to import ähnelt.
        '''</summary>
        Friend ReadOnly Property FrmOptionsCreateBackup() As String
            Get
                Return ResourceManager.GetString("FrmOptionsCreateBackup", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Git settings ähnelt.
        '''</summary>
        Friend ReadOnly Property FrmOptionsGitGrp() As String
            Get
                Return ResourceManager.GetString("FrmOptionsGitGrp", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Location of git.exe ähnelt.
        '''</summary>
        Friend ReadOnly Property FrmOptionsGitLocation() As String
            Get
                Return ResourceManager.GetString("FrmOptionsGitLocation", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Import settings for code modules ähnelt.
        '''</summary>
        Friend ReadOnly Property FrmOptionsImportGrp() As String
            Get
                Return ResourceManager.GetString("FrmOptionsImportGrp", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Keep backup files ähnelt.
        '''</summary>
        Friend ReadOnly Property FrmOptionsKeepBackup() As String
            Get
                Return ResourceManager.GetString("FrmOptionsKeepBackup", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Langauge settings will be applied with the next restart of the host application. ähnelt.
        '''</summary>
        Friend ReadOnly Property FrmOptionsLangInfo() As String
            Get
                Return ResourceManager.GetString("FrmOptionsLangInfo", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Language ähnelt.
        '''</summary>
        Friend ReadOnly Property frmOptionsLanguage() As String
            Get
                Return ResourceManager.GetString("frmOptionsLanguage", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Name of GoTo statement ähnelt.
        '''</summary>
        Friend ReadOnly Property FrmOptionsNameOfGoToStatement() As String
            Get
                Return ResourceManager.GetString("FrmOptionsNameOfGoToStatement", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Select path to git.exe ähnelt.
        '''</summary>
        Friend ReadOnly Property FrmOptionsTitelGitSearch() As String
            Get
                Return ResourceManager.GetString("FrmOptionsTitelGitSearch", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Changed ähnelt.
        '''</summary>
        Friend ReadOnly Property GH_Changed() As String
            Get
                Return ResourceManager.GetString("GH_Changed", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die New ähnelt.
        '''</summary>
        Friend ReadOnly Property GH_New() As String
            Get
                Return ResourceManager.GetString("GH_New", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Stashed ähnelt.
        '''</summary>
        Friend ReadOnly Property GH_Stashed() As String
            Get
                Return ResourceManager.GetString("GH_Stashed", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Code export/import ähnelt.
        '''</summary>
        Friend ReadOnly Property menuCodeExportImport() As String
            Get
                Return ResourceManager.GetString("menuCodeExportImport", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Add error handling ähnelt.
        '''</summary>
        Friend ReadOnly Property menuErrorHandling() As String
            Get
                Return ResourceManager.GetString("menuErrorHandling", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Add error handling with debug message ähnelt.
        '''</summary>
        Friend ReadOnly Property menuErrorHandlingDebug() As String
            Get
                Return ResourceManager.GetString("menuErrorHandlingDebug", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Export code ähnelt.
        '''</summary>
        Friend ReadOnly Property menuExportCode() As String
            Get
                Return ResourceManager.GetString("menuExportCode", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Git ähnelt.
        '''</summary>
        Friend ReadOnly Property menuGitExport() As String
            Get
                Return ResourceManager.GetString("menuGitExport", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Import code ähnelt.
        '''</summary>
        Friend ReadOnly Property menuImport() As String
            Get
                Return ResourceManager.GetString("menuImport", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Indentation ähnelt.
        '''</summary>
        Friend ReadOnly Property menuIndentationAll() As String
            Get
                Return ResourceManager.GetString("menuIndentationAll", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Line numbering ähnelt.
        '''</summary>
        Friend ReadOnly Property menuLineNumber1() As String
            Get
                Return ResourceManager.GetString("menuLineNumber1", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Remove line numbering ähnelt.
        '''</summary>
        Friend ReadOnly Property menuLineNumber2() As String
            Get
                Return ResourceManager.GetString("menuLineNumber2", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Line numbering current prodedure ähnelt.
        '''</summary>
        Friend ReadOnly Property menuLineNumber3() As String
            Get
                Return ResourceManager.GetString("menuLineNumber3", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Remove line numbering current procedure ähnelt.
        '''</summary>
        Friend ReadOnly Property menuLineNumber4() As String
            Get
                Return ResourceManager.GetString("menuLineNumber4", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Settings ähnelt.
        '''</summary>
        Friend ReadOnly Property menuSettings() As String
            Get
                Return ResourceManager.GetString("menuSettings", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Increment version number ähnelt.
        '''</summary>
        Friend ReadOnly Property menuVersionNumber() As String
            Get
                Return ResourceManager.GetString("menuVersionNumber", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Hint ähnelt.
        '''</summary>
        Friend ReadOnly Property Msg_Hint() As String
            Get
                Return ResourceManager.GetString("Msg_Hint", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die This action is canceled. ähnelt.
        '''</summary>
        Friend ReadOnly Property msgActionCanceled() As String
            Get
                Return ResourceManager.GetString("msgActionCanceled", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Check wether the code of the following code modules behaves as intended: ähnelt.
        '''</summary>
        Friend ReadOnly Property msgCodeWorksIndetended() As String
            Get
                Return ResourceManager.GetString("msgCodeWorksIndetended", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Do you want to continue? ähnelt.
        '''</summary>
        Friend ReadOnly Property msgContinue() As String
            Get
                Return ResourceManager.GetString("msgContinue", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die The file &apos;{0}&apos; seems to be not a current code modules.{1}Shall it be deleted? ähnelt.
        '''</summary>
        Friend ReadOnly Property msgDeleteFileExport() As String
            Get
                Return ResourceManager.GetString("msgDeleteFileExport", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die The module &apos;{0}&apos; seems to not be not part of the stored code modules.{1}Shall it be deleted? ähnelt.
        '''</summary>
        Friend ReadOnly Property msgDeleteModuleImport() As String
            Get
                Return ResourceManager.GetString("msgDeleteModuleImport", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die Do you want to add a git commit directly? ähnelt.
        '''</summary>
        Friend ReadOnly Property msgGitDirect() As String
            Get
                Return ResourceManager.GetString("msgGitDirect", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die You must define the path to git.exe in settings. ähnelt.
        '''</summary>
        Friend ReadOnly Property msgMissingGit() As String
            Get
                Return ResourceManager.GetString("msgMissingGit", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die No commit message given. ähnelt.
        '''</summary>
        Friend ReadOnly Property msgNoCommitMessageGiven() As String
            Get
                Return ResourceManager.GetString("msgNoCommitMessageGiven", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die The current VBA Project has no specific name. ähnelt.
        '''</summary>
        Friend ReadOnly Property msgProjectHasNoSpecificName() As String
            Get
                Return ResourceManager.GetString("msgProjectHasNoSpecificName", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Sucht eine lokalisierte Zeichenfolge, die To use this function a name is required. ähnelt.
        '''</summary>
        Friend ReadOnly Property msgUseThisFunction() As String
            Get
                Return ResourceManager.GetString("msgUseThisFunction", resourceCulture)
            End Get
        End Property
    End Module
End Namespace
