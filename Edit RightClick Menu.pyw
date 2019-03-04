from winreg import *
from tkinter import *
import os, sys, ctypes

#Functions----------------------------------------------------------------------
def elevatePrivileges(tempFolder = __file__.rsplit("\\", 1)[0]):
    #Fix tempFoler path
    if not tempFolder.endswith("\\"): tempFolder += "\\"
    fileName = __file__.rsplit("\\", 1)[1].rsplit(".", 1)[0]
    
    #Check if program was executed with admin priv
    try:
        is_admin = os.getuid() == 0
    except AttributeError:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0

    if is_admin:
        return 1 #Has admin priv
    else:
        if os.path.isfile(tempFolder + fileName + " - TempFlag.txt"):
            os.remove(tempFolder + fileName + " - TempFlag.txt")
            os.remove(tempFolder + fileName + " - TempVbs.vbs")
            return 0 #Already asked and the asnwer was no
        else:
            #Create VBS
            lines = []
            lines.append('Set objShell = CreateObject("Shell.Application")')
            lines.append('Set fso = CreateObject("Scripting.FileSystemObject")')
            lines.append('If WScript.Arguments.length = 0 Then')
            lines.append('If (fso.FileExists("'+tempFolder+fileName+' - TempFlag.txt")) Then')
            lines.append('objShell.ShellExecute "'+__file__+'", "", "", "open", 1')
            lines.append('Else')
            lines.append('Set MyFile = fso.CreateTextFile("'+tempFolder+fileName+' - TempFlag.txt", True)')
            lines.append('MyFile.Close')
            lines.append('Set MyFile = fso.GetFile("'+tempFolder+fileName+' - TempFlag.txt")')
            lines.append('MyFile.Attributes = MyFile.Attributes + 2')
            lines.append('objShell.ShellExecute "wscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 1')
            lines.append('objShell.ShellExecute WScript.ScriptFullName, "", "", "open", 1')
            lines.append('End If')
            lines.append('Else')
            lines.append('objShell.ShellExecute "'+__file__+'", "", "", "open", 1')
            lines.append('End If')

            fileVb = open(tempFolder + fileName + " - TempVbs.vbs", "w")
            for line in lines:
                fileVb.write(line + "\n")
            fileVb.close()

            #Hide temp files
            ctypes.windll.kernel32.SetFileAttributesW(tempFolder + fileName + " - TempVbs.vbs", 2)
            
            #Run VBS (Asks to run this python file as administrator and creates TempFlag.txt)
            os.startfile(tempFolder + fileName + " - TempVbs.vbs")
            return 2 #Asking for admin priv
        
def listSubkeys(HKey, subKey, startwith):
    out = []
    try:
        aKey = OpenKey(HKey, subKey, 0, KEY_ALL_ACCESS)
        i = 0
        while True:
            asubkey = EnumKey(aKey, i)
            if startwith:
                if asubkey.startswith(startwith):
                    out.append(asubkey)
            else:
                out.append(asubkey)
            i += 1
    except (WindowsError, PermissionError):
        return out

def listValues(HKey, subKey):
    out = []
    try:
        aKey = OpenKey(HKey, subKey, 0, KEY_ALL_ACCESS)
        i = 0
        while True:
            name, val, typ = EnumValue(aKey, i)
            out.append(name)
            i += 1
    except (WindowsError, PermissionError):
        return out

def getMenuItems(sub=r"", HKey=HKEY_CLASSES_ROOT, startwith=""):
    #print("-"+sub)
    out = []
    try:
        aKey = OpenKey(HKey, sub, 0, KEY_ALL_ACCESS)
        i = 0
        while True:
            #Test is in menu
            asubKey = EnumKey(aKey, i)
            if startwith:
                if asubKey.startswith(startwith):
                    out += getMenuItems(sub+"\\"+asubKey)
                    if (asubKey == "ShellNew") and ("Config" not in listSubkeys(HKey, sub+"\\"+asubKey, "")) and (len(listValues(HKEY_CLASSES_ROOT, sub+"\\"+asubKey))!=0):
                        out.append(sub+"\\"+asubKey)
            else:
                out += getMenuItems(sub+"\\"+asubKey)
                if (asubKey == "ShellNew") and ("Config" not in listSubkeys(HKey, sub+"\\"+asubKey, "")) and (len(listValues(HKEY_CLASSES_ROOT, sub+"\\"+asubKey))!=0):
                    out.append(sub+"\\"+asubKey)
            i += 1
    except (WindowsError, PermissionError):
        return out

def basicLabel(root, txt, pos, span=0, pad=(0, 0), stick=None, col=None):
  x, y = pos
  pad1, pad2 = pad
  Label(root, text=txt).grid(row=x, column=y, columnspan=span, padx=(pad1, pad2), sticky=stick, fg=col)

def basicButtion(root, txt, comm, pos, span, wid, pad=(0, 0), stick=None):
  x, y = pos
  pad1, pad2 = pad
  z = Button(root, text=txt, command=comm, width=wid)
  z.grid(row=x, column=y, columnspan=span, padx=(pad1, pad2), pady=(2, 1), sticky=stick)

def basicEntry(root, pos, span, wid, pad, stick=None, comm=None):
  x, y = pos
  z = Entry(root)
  z.config(width=wid, justify="left")
  if comm: z.bind('<Return>', comm)
  z.grid(row=x, column=y, columnspan=span, padx=(pad,0), pady=(4, 4), sticky=stick)
  return z

def buttonAdd(z=None): #Z is because the entry box returns something
    #Get input
    item = entryAdd.get()
    if item.startswith("."): item = item[1:]

    #Check for errors
    error = []
    if "." in item: error.append(".")
    if "\\" in item: error.append("\\")
    if len(item)==0: error.append("No value")

    if len(error):
        labelLastAction.set("AddError(" + str(len(error)) + "):  " + "  ".join(error))

    else:
        item = "." + item
        
        #Create Key and value
        CreateKey(HKEY_CLASSES_ROOT, item+"\\ShellNew")
        with OpenKey(HKEY_CLASSES_ROOT, item+"\\ShellNew", 0, KEY_ALL_ACCESS) as h:
            SetValueEx(h, "NullFile", 0, REG_SZ, None)

        global x, y, lastAction
        lastAction = "Added: " + item #Add to log

        #Get new menu (Close current one)
        x = root.winfo_x()
        y = root.winfo_y()
        root.destroy()

def buttonRemove(targetName):
    #delete all keys with the same name
    for n in range(len(inMenuPaths)):
        item = inMenuPaths[n].split("\\", 2)
        if item[1] == targetName:
            DeleteKey(HKEY_CLASSES_ROOT, item[1]+"\\"+item[2])
            

    global x, y, lastAction
    lastAction = "Removed: " + targetName     #Add to log

    #Get new menu (Close current one)
    x = root.winfo_x()
    y = root.winfo_y()
    root.destroy()

def onClose():
    global run
    run = False
    root.destroy()

#Main---------------------------------------------------------------------------
if elevatePrivileges() == 1:
    run = True
    lastAction = "#Last Action"
    x = 200
    y = 200
    while run:
        #Get the keys of the extentions in the "new" menu
        inMenuPaths = getMenuItems(startwith=".")
        inMenu = []
        for n in range(len(inMenuPaths)):
            item = inMenuPaths[n].split("\\", 2)[1]
            if item not in inMenu: inMenu.append(item)
        inMenu.sort()
    
        #Pack menu
        root = Tk() #Create the window that will be open
        root.wm_title("Edit RightClick/New Menu")
        root.resizable(0,0) #Make window not resizeable
        root.geometry("+"+str(x)+"+"+str(y))
        root.protocol("WM_DELETE_WINDOW", onClose)
    
        labelLastAction = StringVar(root)
        labelLastAction.set(lastAction)

        entryAdd = basicEntry(root, (2, 0), 1, 14, 5, "E", buttonAdd)
        basicButtion(root, "Add", lambda:buttonAdd(), (2, 1), 1, 7, (5, 5), "W")
        basicLabel(root, "", (3, 0), 4)
        Label(root, textvariable=labelLastAction, fg="red").grid(row=2, column=2, columnspan=2, padx=(5, 5), sticky="W")

        iCount = 0
        for key in inMenu:
            i = iCount + 8
            n = [0,5]
            if iCount%2: n = [2, 30]
    
            basicLabel(root, key, (int(i/2), n[0]), 1, (30, 0), "E")
            basicButtion(root, "Remove", lambda d=inMenu[iCount]:buttonRemove(d), (int(i/2), n[0]+1), 1, 7, (5, n[1]), "W")
            iCount += 1

        root.mainloop()
