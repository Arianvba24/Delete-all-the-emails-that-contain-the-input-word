#Those two piece of code may take a few time to finally be executed
#This is the cell that deletes all the emails that contain the word "iberempleos.es"
import win32com.client

Outlook = win32com.client.Dispatch("Outlook.Application")
namespace = Outlook.GetNamespace("MAPI")

carpetas = namespace.Folders

segunda_cuenta = carpetas[1]

bandeja_entrada_segunda_cuenta = segunda_cuenta.Folders("Bandeja de entrada")

counter = 0
for i in bandeja_entrada_segunda_cuenta.Items:
    if counter == 50:
        break
    else:
        if "iberempleos.es" in i.SenderEmailAddress:
            try:
                i.Delete()
                counter = counter + 1
            except Exception as e:
                continue

