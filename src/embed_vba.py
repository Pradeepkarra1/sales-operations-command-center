import subprocess, time, uno, sys, zipfile
from openpyxl import Workbook
from com.sun.star.beans import PropertyValue
SHEETS=["Start_Here","Dashboard","Findings","Sales_Analysis","Ops_Analysis","Statistics","Forecast","What_If","Interview_Guide","Macro_Guide","Forecast_Model","Sales_Data","Dim_Product","Dim_Dealer","Targets"]
wb=Workbook(); wb.remove(wb.active); wb.code_name="ThisWorkbook"
for i,n in enumerate(SHEETS,1):
    ws=wb.create_sheet(n); ws.sheet_properties.codeName=f"Sheet{i}"
wb.vba_archive=zipfile.ZipFile("t.xlsm"); wb.save("seed.xlsm")
def pv(n,v):
    p=PropertyValue(); p.Name=n; p.Value=v; return p
proc=subprocess.Popen(["soffice","--headless","--invisible","--norestore","--accept=socket,host=127.0.0.1,port=2002;urp;","-env:UserInstallation=file:///tmp/lo_vba"],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
ctx=None
for i in range(60):
    try:
        local=uno.getComponentContext(); res=local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",local)
        ctx=res.resolve("uno:socket,host=127.0.0.1,port=2002;urp;StarOffice.ComponentContext"); break
    except Exception: time.sleep(1)
if not ctx: sys.exit("no connect")
desk=ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop",ctx)
doc=desk.loadComponentFromURL("file:///home/claude/proj/vba/seed.xlsm","_blank",0,(pv("Hidden",True),pv("MacroExecutionMode",4)))
libs=doc.BasicLibraries
print("libs",libs.getElementNames(),"compat",libs.VBACompatibilityMode,"proj",libs.ProjectName)
lib=libs.getByName("VBAProject")
for m in lib.getElementNames():
    print(" module",m,"type",lib.getModuleInfo(m).ModuleType, repr(lib.getByName(m)[:60]))
code=open("modCommandCenter.vba").read()
if lib.hasByName("modTest"): lib.removeByName("modTest")
lib.insertByName("modCommandCenter",code)
info=uno.createUnoStruct("com.sun.star.script.ModuleInfo"); info.ModuleType=1
if not lib.hasModuleInfo("modCommandCenter"): lib.insertModuleInfo("modCommandCenter",info)
helper = """Sub AddDocModules
    Dim oLib, oProv, i, names
    Dim oInfo As New com.sun.star.script.ModuleInfo
    oLib = ThisComponent.BasicLibraries.getByName("VBAProject")
    oProv = ThisComponent.createInstance("ooo.vba.VBAObjectModuleObjectProvider")
    names = oProv.getElementNames()
    For i = LBound(names) To UBound(names)
        If Not oLib.hasByName(names(i)) Then
            oInfo.ModuleType = 4
            oInfo.ModuleObject = oProv.getByName(names(i))
            oLib.insertByName(names(i), "Rem Attribute VBA_ModuleType=VBADocumentModule" & Chr(10) & "Option VBASupport 1" & Chr(10))
            oLib.insertModuleInfo(names(i), oInfo)
        End If
    Next
End Sub
"""
std=libs.getByName("Standard")
if std.hasByName("Helper"): std.removeByName("Helper")
std.insertByName("Helper",helper)
sp=doc.getScriptProvider()
sc=sp.getScript("vnd.sun.star.script:Standard.Helper.AddDocModules?language=Basic&location=document")
sc.invoke((),(),())
std.removeByName("Helper")
for m in lib.getElementNames(): print(" mod",m,lib.getModuleInfo(m).ModuleType)
print("final modules",lib.getElementNames())
doc.storeToURL("file:///home/claude/proj/vba/lo_out.xlsm",(pv("FilterName","Calc MS Excel 2007 VBA XML"),))
doc.close(True)
try: desk.terminate()
except Exception: pass
proc.wait(timeout=30)
