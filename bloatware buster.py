from tkinter import *
from tkinter.ttk import *
from tkinter import messagebox
from tkinter import filedialog
from idlelib.tooltip import Hovertip
import subprocess,os,json
import ctypes, sys
import threading,time,webbrowser



def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


window=Tk()
window.geometry("400x527")
window.title("Bloatware Buster")
window.resizable(0,0)
window.iconbitmap(resource_path("trash.ico"))

notebook=Notebook(window)
notebook.pack(fill="both",expand=1,pady=(5,0))

frame1=Frame(notebook)
frame2=Frame(notebook)
frame3=Frame(notebook)

frame1.pack(fill="both",expand=1)
frame3.pack(fill="both",expand=1)

status_frame=Frame(window)
status_frame.pack(side="bottom",fill="x")

notebook.add(frame1,text=" Uninstall ")
notebook.add(frame3,text=" Other ")

# unistall bloat
label=Label(frame1,text="Select the apps you want to uninstall.")
label.pack(side="top",anchor="w",pady=(10,7),padx=5)

chk_btn_frame=Frame(frame1)
chk_btn_frame.pack(side="top",fill="x",padx=(5,5))

status_sep=Separator(status_frame,orient="horizontal")
status_sep.pack(side="top", fill="x",)

status_label=Label(status_frame,text="               ")
status_label.pack(side="left",anchor="w",padx=(5))


log_btn=Label(status_frame,text="Log")
log_btn.pack(side="right",anchor="e",padx=(6,5))
Hovertip(log_btn, " Click to view log ")
status_sep=Separator(status_frame,orient="vertical")
status_sep.pack(side="right", fill="y")

def call(event):
    logwindow.deiconify()
    logwindow.config(takefocus=True)
    logwindow.lift()
log_btn.bind("<Button-1>",call)


sub_frame=Frame(frame1)
sub_frame.pack(side="bottom",fill="x",expand=1,pady=(10,0))

logwindow=Toplevel()
logwindow.title("Log")
logwindow.geometry("526x548+490+100")
logwindow.withdraw()
logwindow.resizable(0,0)

def on_closing():
    logwindow.withdraw()
logwindow.protocol("WM_DELETE_WINDOW", on_closing)

def save_log():
    log=text.get("1.0","end-1c")
    file=filedialog.asksaveasfile(initialfile="Log.txt",mode="w",defaultextension=".txt",filetypes=[("All Files","*.*"),("Text Documents","*.txt")])
    if file is None:
        return
    file.write(log)
    file.close()

def clear():
    text.delete("1.0",END)

mainframe=Frame(logwindow)
mainframe.pack(fill="both",expand=True)

scrollbar=Scrollbar(mainframe,orient="vertical")
scrollbar.pack(side="right",fill="y")

text=Text(mainframe,bg="#f0f0f0",borderwidth=0,font=("Consolas",10),insertborderwidth=2,state="disabled")
text.pack(fill="both",expand=True)
text["yscrollcommand"]=scrollbar.set
scrollbar.config(command=text.yview)

logwin_btn_frame=Frame(logwindow)
logwin_btn_frame.pack(side="bottom",fill="x",)

sep=Separator(logwin_btn_frame,orient="horizontal")
sep.pack(side="top",fill="x")

close_btn=Button(logwin_btn_frame,text="Close",command=logwindow.withdraw)
close_btn.pack(side="right",anchor="e",padx=(10,10),ipadx=2,ipady=1,pady=12)

save_btn=Button(logwin_btn_frame,text="Save",command=save_log)
save_btn.pack(side="right",anchor="e",padx=(10,0),ipadx=2,ipady=1,pady=12)

clear_btn=Button(logwin_btn_frame,text="Clear",command=clear)
clear_btn.pack(side="right",anchor="e",padx=(10,0),ipadx=2,ipady=1,pady=12)




chk_btn_values={"Mail and Calendar":"*microsoft.communicationapps*","Getstarted":"*Microsoft.Getstarted*","3DViewer":"*Microsoft.Microsoft3DViewer*","StickyNotes":"*Microsoft.MicrosoftStickyNotes*","GetHelp":"*Microsoft.GetHelp*","SolitaireCollection":"*Microsoft.MicrosoftSolitaireCollection*","OfficeHub":"*Microsoft.MicrosoftOfficeHub*","Clock":"*Microsoft.WindowsAlarms*","Onenote":"*onenote*","OneDriveSync":"*OneDriveSync*","Paint":"*Paint*","FeedbackHub":"*WindowsFeedbackHub*","Maps":"*maps*","Bingnews":"*bingnews*","People":"*people*","Skype":"*skypeapp*","Weather":"*bingweather*","Soundrecorder":"*soundrecorder*","Xbox":"*xbox*","YourPhone":"*YourPhone*","Cortana":"*Microsoft.549981C3F5F10*","Camera":"Microsoft.WindowsCamera","Photos":"Microsoft.Windows.Photos","MS Store":"Microsoft.WindowsStore","Edge":"Microsoft.MicrosoftEdge.Stable","Calculator":"Microsoft.WindowsCalculator","Groove music":"Microsoft.ZuneMusic","Movies and TV":"Microsoft.ZuneVideo","Snip and Sketch":"Microsoft.ScreenSketch"}
print(len(chk_btn_values))
uninstall_li=[]

var_names_li=[]
for key in chk_btn_values:
    var = IntVar()
    var_names_li.append(var)

pos_li=[] #list of chk buttons which have their variable value to 1
def clicked(value,key,pos):
    global posg
    if value.get() == 1:
        print(key," checked")
        uninstall_li.append(key)
        pos_li.append(pos)
    else:
        try:
            uninstall_li.remove(key)
            pos_li.remove(pos)
        except:
            pass


chk_btn_li=[]
print(len(chk_btn_values))
for index1,(value,var) in enumerate(zip(chk_btn_values,var_names_li)):
    
    key=Checkbutton(chk_btn_frame,text=value,variable=var,command=lambda v=var,k=chk_btn_values[value],p=index1:clicked(v, k, p))
    key.grid(row=index1 // 3,column=index1 % 3,sticky="nsew",padx=9,pady=6)
    chk_btn_li.append(key)


final_li=[]
def uninstall():
    global uninstall_li
    print(uninstall_li)
    text.config(state="normal")
    if len(uninstall_li)==0:
        messagebox.showerror("Error","Please select app(s) to install.")
        return
    pb["value"]=0
    increment=100/len(uninstall_li)

    unistall_btn.config(state="disabled")
    chkall_btn.config(state="disabled")
    unchkall_btn.config(state="disabled")
    recommend_btn.config(state="disabled")

    for btn in chk_btn_frame.winfo_children():
        btn.config(state="disabled")

    for app in enumerate(uninstall_li):
        print(app)
        cmd=f'powershell "Get-AppxPackage -allusers {app[1]} | Remove-AppxPackage"'
        print(cmd)
        final_li.append(app[1])
        result=subprocess.run(cmd, shell=True, capture_output=True)

        if result.returncode:
            status_label.config(text=f"Uninstalled {app[1]}")
            print(result.stdout)
            text.insert(f"{app[0]}.0",f"[INFO] Successfully uninstalled {app[1]}\n")
        else:
            text.insert(f"{app[0]}.0",f"[ERR] Failed to uninstalled {app[1]}\nreason----\n{result.stderr}\n")
            messagebox.showerror("Error",f"Failed to uninstall {app[1]}\n\n{result.stderr}")
            pb["value"]+=increment
            continue
            print(result.stderr)
        if pb["value"]<100:
            pb["value"]+=increment
            print(int(pb["value"]))
    
    text.config(state="disabled")
    unistall_btn.config(state="normal")
    chkall_btn.config(state="normal")
    unchkall_btn.config(state="normal")
    recommend_btn.config(state="normal")

    for btn in (chk_btn_frame.winfo_children()):
        try:
            if btn.state()[0] == "disabled" and btn.state()[1]=="selected":
                continue
        except:
            pass
        btn.config(state="normal")
    
    uninstall_li.clear()
    print(final_li)

def caller(func):
    t = threading.Thread(target=func)
    t.start()
    

def unchkall():
    pos_li.clear()
    uninstall_li.clear()

    for pos ,(var,key,btn) in enumerate(zip(var_names_li,chk_btn_values,chk_btn_li)):
        print(btn.state())
        if var.get() == 0:
            continue
        if btn.state()[0]=="disabled":
            continue
        var.set(0)
        print(key)
        
    print(pos_li)
    print(uninstall_li)
        

def rmd():
    pos_li.clear()
    for pos ,(var,key,btn) in enumerate(zip(var_names_li,chk_btn_values,chk_btn_li)):
        
        if key == "Photos" or key=="MS Store" or key=="Calculator" or key=="Edge" or key=="Snip and Sketch":
            try:
                if btn.state()[0] == "disabled":
                    continue
            except:
                pass
            var.set(0)
        else:
            var.set(1)
            uninstall_li.append(chk_btn_values[key])
            pos_li.append(pos)
    
        
def chkall():
    pos_li.clear()
    uninstall_li.clear()
    for pos ,(var,key,btn) in enumerate(zip(var_names_li,chk_btn_values,chk_btn_li)):
        if var.get() == 1:
            continue
        try:
            if btn.state()[0]=="disabled":
                continue
        except:
            pass
        var.set(1)
        uninstall_li.append(chk_btn_values[key])
        pos_li.append(pos)



pb=Progressbar(sub_frame,orient="horizontal",mode="determinate",length=375)
pb.grid(row=0,column=0,columnspan=4,pady=(10),padx=9)

unchkall_btn=Button(sub_frame,text="Unheck all",command=unchkall)
unchkall_btn.grid(row=1,column=0,sticky="n",padx=(9),ipadx=2,ipady=1,pady=7)
Hovertip(unchkall_btn, " Uncheck all apps ")

chkall_btn=Button(sub_frame,text="Check all",command=chkall)
chkall_btn.grid(row=1,column=1,padx=(7),ipadx=2,ipady=1,pady=7)
Hovertip(chkall_btn, " Check all apps ")

recommend_btn=Button(sub_frame,text="Recommended",command=rmd)
recommend_btn.grid(row=1,column=2,padx=(7),ipadx=2,ipady=1,pady=7)
Hovertip(recommend_btn, " Check the apps which are commonly not used ")

unistall_btn=Button(sub_frame,text="Uninstall",command=lambda f=uninstall:caller(f))
unistall_btn.grid(row=1,column=3,sticky="n",padx=(7),ipadx=2,ipady=1,pady=7)
unistall_btn.focus_set()
Hovertip(unistall_btn, " Uninstall selected apps ")


f1=LabelFrame(frame3,text="Theme")
f1.pack(fill="x",padx=7,pady=7)

label=Label(f1,text="Select light or dark mode.")
label.pack(side="top",pady=(2,5),padx=5,anchor="w")

def change_mode():
    command = ['reg.exe', 'add', 'HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize','/v', f'AppsUseLightTheme', '/t', 'REG_DWORD', '/d', f'{var2.get()}', '/f']
    subprocess.run(command)

    command = ['reg.exe', 'add', 'HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize','/v', f'SystemUsesLightTheme', '/t', 'REG_DWORD', '/d', f'{var2.get()}', '/f'] 
    subprocess.run(command)
    print(var2.get())

var2=IntVar()
# var2.set(1)
lbtn=Radiobutton(f1,text="Light",variable=var2,value=1,command=lambda f=change_mode:caller(f))
lbtn.pack(anchor="w",padx=5,pady=(0,3))
Hovertip(lbtn, " Light mode ")

dbtn=Radiobutton(f1,text="Dark",variable=var2,value=0,command=lambda f=change_mode:caller(f))
dbtn.pack(anchor="w",padx=5,pady=(3,4))
Hovertip(dbtn, " Dark mode ")

f2=LabelFrame(frame3,text="Start menu")
f2.pack(fill="x",padx=7,pady=7)

label=Label(f2,text="Clean start menu, removes all tiles from the start menu including\nthe apps which are not installed.")
label.pack(side="top",pady=(2,5),padx=5,anchor="w")

clean_status=0
def clean_start():
    global clean_status
    clean_btn.config(state="disabled")
    file=resource_path("remove-start-menu-tiles.ps1")
    cmd=f"powershell -ExecutionPolicy Bypass -File {file}"
    result=subprocess.run(cmd, shell=True, capture_output=True)
    status_label.config(text="Start menu was cleaned")
    clean_btn.config(state="normal")
    clean_status=1


clean_btn=Button(f2,text="Remove",command=lambda f=clean_start:caller(f))
clean_btn.pack(side="bottom",anchor="e",padx=10,ipadx=2,ipady=1,pady=7)
Hovertip(clean_btn, " Clean the start menu ")

f3=LabelFrame(frame3,text="Configuration")
f3.pack(fill="x",padx=7,pady=7)



label=Label(f3,text="Import or export your current cofiguration, in order to use it later.")
label.pack(side="top",pady=(2,5),padx=5,anchor="w")

def export_config():
    config={}
    for pos,appid in zip(pos_li,final_li):
        config[pos]=appid
    print("config: ",config)
    config["LightTheme"]=var2.get()
    config["StartMenuCleaned"]=clean_status

    with open(r"C:\Users\Atharva\Documents\Config.json","w") as file:
        json.dump(config,file,indent=4)
    messagebox.showinfo("Exported",r"Configuration exported to 'C:\Users\Atharva\Documents\Config.json'.")

def import_config():
    file=filedialog.askopenfilename(title="Open a file",filetypes=[("json source file","*.json")])
    if file is "":
        return
    result=messagebox.askyesno("File import","Do you want to apply the changes now?. If no, you have to apply them manually.")
    import_btn.config(state="disabled")
    Hovertip(import_btn, " Please wait while we import configuration ")
    with open(file,"r") as file1:
        data=json.load(file1)
    
    pos_li.clear()
    uninstall_li.clear()
    for pos,appid in (data.items()):
        if pos=="LightTheme":
            break
        var_names_li[int(pos)].set(1)
        uninstall_li.append(appid)
    var2.set(int(data["LightTheme"]))

    #call function to apply
    print(YES)
    if result:
        status_label.config(text="Applying config")
        if len(uninstall_li) >= 1:
            caller(uninstall)
        if data["StartMenuCleaned"]:
            caller(clean_start)
        change_mode()
    import_btn.config(state="normal")
    Hovertip(import_btn, "  Import a configuration file and apply it ")
    status_label.config(text="Config applied")
    print(appid)
    

export_btn=Button(f3,text="Export",command=lambda f=export_config:caller(f))
export_btn.pack(side="right",anchor="e",padx=10,ipadx=2,ipady=1,pady=(12,7))
Hovertip(export_btn, " Export your current configuration ")

import_btn=Button(f3,text="Import",command=import_config)
import_btn.pack(side="right",anchor="e",padx=10,ipadx=2,ipady=1,pady=(12,7))
Hovertip(import_btn, " Import a configuration file and apply it ")

f3=LabelFrame(frame3,text="Information")
f3.pack(fill="x",padx=7,pady=7)

def link():
    webbrowser.open("https://github.com/appledino3604")
def info():
    messagebox.showinfo("About Bloatware Buster","Bloatware Buster 1.2.4\n\nA simple utility to remove unwanted bloatware and clean up your start menu, making your system faster and more organized.")
def files(path):
    subprocess.Popen(["notepad.exe",path])

img=PhotoImage(file=resource_path("github_icon.png"))
github_btn=Button(f3,text="Github",image=img,compound="left",command=link)
github_btn.pack(side="bottom",padx=10,ipadx=140,ipady=0,pady=(3,15))
Hovertip(github_btn, "Visit GitHub")

info_btn=Button(f3,text="Info",command=info)
info_btn.pack(anchor="w",side="left",padx=(10,4),ipadx=20,ipady=1,pady=(15,7))
Hovertip(info_btn, " View information about this program ")

help_btn=Button(f3,text="Help",command=lambda p=resource_path("user_manual.txt"):files(p))
help_btn.pack(anchor="e",side="left",padx=4,ipadx=20,ipady=1,pady=(15,7))
Hovertip(help_btn, " View help ")

license_btn=Button(f3,text="License",command=lambda p=resource_path("license.txt"):files(p))
license_btn.pack(anchor="e",side="left",padx=(4,10),ipadx=20,ipady=1,pady=(15,7))
Hovertip(license_btn, " View license ")

window.mainloop()
