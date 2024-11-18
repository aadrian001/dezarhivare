#Requires AutoHotkey v2.0

global Counter := 0
global SleepTime := 85
global LongSleepTime := 2500
global IsPaused := false
global ButtonLU  
global ButtonSA
global IsPausedMessageShown := false 
global query := "" 
global ButtonConfig  
global queryGui
global iniFile := A_MyDocuments  . "\archivedb_config.ini" 
githubScriptUrl := "https://raw.githubusercontent.com/aadrian001/dezarhivare/refs/heads/main/test.ahk"  ; Replace with your actual GitHub URL
tempFilePath := A_Temp . "\temp-script.ahk"

if !FileExist(iniFile)
{
    IniWrite("ROTM2DB03", iniFile, "Database", "DataSource")  
    IniWrite("sensordata_new", iniFile, "Database", "InitialCatalog")
    IniWrite("SSPI", iniFile, "Database", "IntegratedSecurity")
}

if A_LineFile = A_ScriptFullPath && !A_IsCompiled
{
    myGui := Constructor()
    myGui.Show("w204 h350")
}

Constructor()
{    
    global ButtonSA, ButtonLU  
    myGui := Gui()
    myGui.Opt("-Resize -MinimizeBox -MaximizeBox")
    ButtonSA := myGui.Add("Button", "x24 y85 w157 h56 Center", "&Start Unarchiving")
    ButtonLU := myGui.Add("Button", "x24 y155 w157 h56 Center", "&Find Archived LU")
    Edit1 := myGui.Add("Edit", "x24 y48 w156 h21")
    myGui.Add("Text", "x0 y16 w204 h23 +0x200 Center", "Number of SFCs to be unarchived:")
    myGui.Add("Text", "x5 y215 w204 h23 +0x200 Left", "Select speed:")
    global Fast := myGui.Add("Radio", "x24 y240", "Fast (1200ms)")
    global Medium := myGui.Add("Radio", "x24 y260", "Medium (3000ms)")
    global Slow := myGui.Add("Radio", "x24 y280", "Slow (4500ms)")
    myGui.Add("Text", "x5 y300 w204 h23 +0x200 Left", "Pause: CTRL + P")
    myGui.Add("Text", "x5 y320 w204 h23 +0x200 Left", "Quit: CTRL + Q")
    Medium.Value := true
    ButtonCheckUpdate := myGui.Add("Button", "x24 y320 w157 h56 Center", "&Check for Updates")
    ButtonCheckUpdate.OnEvent("Click", (*) => CheckForUpdates())  
    ButtonSA.OnEvent("Click", (*) => StartUnarchiving(Edit1.Value))
    ButtonLU.OnEvent("Click", (*) => OpenLUFinder())
    Edit1.OnEvent("Change", OnEventHandler)
    myGui.OnEvent('Close', (*) => ExitApp())
    
    myGui.Title := "Unarchive SFC"
    
    OnEventHandler(*)
    {
    }
    return myGui
}


CloseLUFinder()
{
    global ButtonSA, dbGui
    ButtonSA.Enabled := true
    dbGui.Destroy()
}

OpenLUFinder()
{   
    global ButtonSA, dbGui, query, ButtonConfig, ButtonSQL
    dbGui := Gui()
    dbGui.Opt("-Resize -MinimizeBox -MaximizeBox")
    dbGui.Add("Text", , "Barcodes:")
    barcodesInput := dbGui.Add("Edit", "w300 h100") 
    dbGui.Add("Text", , "Results:")
    ResultBox := dbGui.Add("Edit", "w300 h200 ReadOnly") 
    ButtonExecute := dbGui.Add("Button", "w300 h30 Center", "Execute")
    ButtonExecute.OnEvent("Click", (*) => ExecuteQuery(ResultBox, barcodesInput.Value))
    ButtonExportQuery := dbGui.Add("Button", "w300 h30 Center", "Export SQL Query log file")  
    ButtonExportQuery.OnEvent("Click", (*) => ExportFinalQuery()) 
    ButtonConfig := dbGui.Add("Button", "w300 h30 Center", "&Database connection config")
    ButtonConfig.OnEvent("Click", (*) => OpenConfigGUI())  
    ButtonSQL := dbGui.Add("Button", "w300 h30 Center", "&Edit SQL Query")
    ButtonSQL.OnEvent("Click", (*) => QueryEdit())  
    dbGui.Show("w320 h500")
    dbGui.OnEvent('Close', (*) => CloseLUFinder()) 
    ButtonSA.Enabled := false
    ButtonSQL.Enabled := false
}
QueryEdit()
{
    global queryGui

    queryGui := Gui()
    queryGui.Opt("-Resize -MinimizeBox -MaximizeBox")
    queryGui.Show("w300 h300")
    InitialQuery := queryGui.Add("Edit", "w280 h140") 

}
GetSelectedSleepTimes()
{
    global Fast, Medium, Slow, SleepTime, LongSleepTime
    if Fast.Value
    {
        LongSleepTime := 1200
    }
    else if Medium.Value
    {
        LongSleepTime := 3000
    }
    else if Slow.Value
    {
        LongSleepTime := 4500
    }
}

StartUnarchiving(value)
{
    global Counter, ButtonLU, ButtonSA


    if !IsNumber(value) || value <= 0
    {
        SoundPlay("C:\Windows\Media\Windows Foreground.wav", "Async")
        MsgBox("Please enter a valid positive number in the field.", "ERROR!")
        Return
    }

    Counter := value  
    ButtonLU.Enabled := false  
    ButtonSA.Enabled := false  
    StartLoop() 
}
StartLoop()
{
    global Counter, IsPaused, IsPausedMessageShown, myGui, errorOccurred  

    errorOccurred := false  

    Loop Counter
    {
        if (IsPaused)
        {
            while (IsPaused)
            {
                Sleep(100) 
            }
            continue
        }

        ButtonLU.Enabled := false
        ButtonSA.Enabled := false
        GetSelectedSleepTimes()
        PerformExcelActions()
        PerformEdgeActions()

        if (errorOccurred)  
        {
            MsgBox("The process was stopped due to an error.")
            myGui.Show()  
            errorOccurred := true  
            ButtonSA.Enabled := true
            break  

        }

        Sleep(SleepTime)
    }

    if (!errorOccurred)
    {
        ShowCompletionGUI()
    }
}
ShowCompletionGUI()
{
    global ButtonLU, completionGui, myGui 

    SoundPlay("C:\Windows\Media\tada.wav", "Async")
    completionGui := Gui()
    completionGui.Opt("-Resize -MinimizeBox -MaximizeBox")
    completionGui.Add("Text", "x55 y20 w250 h30", Format("Unarchived {} SFCs successfully.", Counter))
    ButtonResume := completionGui.Add("Button", "x20 y60 w100 h30", "Resume LU Finder")
    ButtonResume.OnEvent("Click", (*) => ResumeLUFinder())
    ButtonExit := completionGui.Add("Button", "x130 y60 w100 h30", "Exit Application")
    ButtonExit.OnEvent("Click", (*) => ExitApp())
    completionGui.Show("w250 h130")
    completionGui.OnEvent('Close', (*) => ExitApp())
}

ResumeLUFinder()
{
    global Counter, ButtonLU, ButtonSA, completionGui, myGui
    ButtonLU.Enabled := true
    ButtonSA.Enabled := true
    completionGui.Destroy()
    WinActivate("ahk_id " myGui.Hwnd)

    if (Counter <= 0)
    {
        MsgBox("Counter value is invalid. Please enter a valid number to start unarchiving.")
        return
    }

}
PerformExcelActions()
{
    global IsPaused

    If WinExist("ahk_class XLMAIN")
    {
        WinActivate("ahk_class XLMAIN")
        WinWaitActive("ahk_class XLMAIN")

        Send("^c") 
        Sleep(SleepTime)

        Send("{Down}")
        Sleep(SleepTime)
    }
    else
    {
        SoundPlay("C:\Windows\Media\Windows Foreground.wav", "Async")
        MsgBox("Excel is not open.", "ERROR!")
        Return
    }
}
PerformEdgeActions()
{
    global errorOccurred

    If WinExist("ahk_class Chrome_WidgetWin_1 ahk_exe msedge.exe")
    {
        WinActivate("ahk_class Chrome_WidgetWin_1 ahk_exe msedge.exe")
        WinWaitActive("ahk_class Chrome_WidgetWin_1 ahk_exe msedge.exe")

        Sleep(500)  

        activeTabTitle := GetActiveTabTitle()

        if !RegExMatch(activeTabTitle, "Archive - SAP Manufacturing Execution SAP SE( and \d+ more pages?)? - Work - Microsoft​ Edge")
        {
            SoundPlay("C:\Windows\Media\Windows Foreground.wav", "Async")
            MsgBox("The active tab is not Archiving activity. Please open the correct TAB.", "ERROR!")
            errorOccurred := true 
            Return  
        }

        Sleep(SleepTime)  
        Send("{Enter}")
        Sleep(SleepTime)
        Send("{Down}")
        Sleep(SleepTime)
        Send("{Enter}")
        Sleep(SleepTime)
        Send("{Tab}")
        Sleep(SleepTime)
        Send("{Tab}")
        Sleep(SleepTime)
        Send("^v")  
        Sleep(SleepTime)
        Send("{Tab}")
        Sleep(SleepTime)
        Send("{Enter}")
        Sleep(LongSleepTime)  
    }
    else
    {
        MsgBox("Microsoft Edge is not open.", "ERROR!")
        errorOccurred := true 
        Return
    }
}
GetActiveTabTitle()
{
    title := ""
    hwnd := WinExist("ahk_class Chrome_WidgetWin_1 ahk_exe msedge.exe")
    if hwnd
    {       
        title := WinGetTitle(hwnd)        
        title := Trim(title)
    }
    return title
}

IsNumber(value)
{
    return (value ~= "^\d+$")
}


OpenConfigGUI()
{
    global iniFile, dataSourceInput, initialCatalogInput, integratedSecurityInput

    configGui := Gui()
    configGui.Opt("-Resize -MinimizeBox -MaximizeBox")
    configGui.Add("Text", "x20 y20 w150 h25", "Server name:")
    dataSourceInput := configGui.Add("Edit", "x170 y20 w200 h25")
    configGui.Add("Text", "x20 y60 w150 h25", "Database:")
    initialCatalogInput := configGui.Add("Edit", "x170 y60 w200 h25")
    configGui.Add("Text", "x20 y100 w150 h25", "Autenthication:")
    integratedSecurityInput := configGui.Add("Edit", "x170 y100 w200 h25")
    dataSourceInput.Text := IniRead(iniFile, "Database", "DataSource", "")
    initialCatalogInput.Text := IniRead(iniFile, "Database", "InitialCatalog", "")
    integratedSecurityInput.Text := IniRead(iniFile, "Database", "IntegratedSecurity", "")
    ButtonSaveConfig := configGui.Add("Button", "w190 h30 Center", "Save Config")
    ButtonSaveConfig.OnEvent("Click", (*) => SaveDatabaseConfig())


    configGui.Show("w400 h200")
}

SaveDatabaseConfig()
{
    global iniFile, dataSourceInput, initialCatalogInput, integratedSecurityInput
    IniWrite(dataSourceInput.Text, iniFile, "Database", "DataSource")
    IniWrite(initialCatalogInput.Text, iniFile, "Database", "InitialCatalog")
    IniWrite(integratedSecurityInput.Text, iniFile, "Database", "IntegratedSecurity")
    MsgBox("Database configuration saved successfully!")
}

ExecuteQuery(ResultBox, barcodes)
{
    global query, iniFile
    dataSource := IniRead(iniFile, "Database", "DataSource", "")
    initialCatalog := IniRead(iniFile, "Database", "InitialCatalog", "")
    integratedSecurity := IniRead(iniFile, "Database", "IntegratedSecurity", "")

    barcodesList := FormatBarcodes(barcodes)
    baseQuery := "
    (
    SELECT DISTINCT
        [parentdata_sfc],
        [parentdata_barcode]
    FROM [sensordata_new].[dbo].[parentdata]
    WHERE parentdata_barcode IN ({barcodes}) 
    )"
    query := StrReplace(baseQuery, "{barcodes}", barcodesList)

    db := ComObject("ADODB.Connection")
    db.ConnectionString := Format("Provider=SQLOLEDB;Data Source={};Initial Catalog={};Integrated Security={};", dataSource, initialCatalog, integratedSecurity)
    try {
        db.Open()
        if (db.State != 1) 
        {
            if (IsObject(ResultBox))  
                ResultBox.Value := "Failed to connect to database."
            return
        }

        rs := db.Execute(query)

        result := "LU:"  
        result .= "`n" 
        
        if rs.EOF 
        {
            if (IsObject(ResultBox))  
                ResultBox.Value := "No records found."  
        }
        else
        {
            while !rs.EOF {
                sfc := rs.Fields.Item("parentdata_sfc").Value
                result .= sfc "`n"  
                rs.MoveNext()
            }
            if (IsObject(ResultBox))  
                ResultBox.Value := result  
        }
        
        rs.Close()
    } catch {
        errorMsg := "An error occurred during query execution.`n"
        if (ComObjType(db.Errors) = 9)  
        {
            for err in db.Errors
            {
                errorMsg .= "Error: " err.Description "`n"
            }
        }
        else
        {
            errorMsg .= "Check the SQL query or database connection settings."
        }

        if (IsObject(ResultBox))  
            ResultBox.Value := errorMsg
    } finally {
        if (db.State = 1) 
            db.Close()
    }
}

ExportFinalQuery()
{
    if (query = "")
    {
        MsgBox("No SQL query to export. Insert barcodes first!")
        return
    }

    currentTime := A_Now
    formattedTime := SubStr(currentTime, 1, 4) "-" SubStr(currentTime, 5, 2) "-" SubStr(currentTime, 7, 2) " " SubStr(currentTime, 9, 2) ":" SubStr(currentTime, 11, 2) ":" SubStr(currentTime, 13, 2)

    logFile := A_ScriptDir "\query_log.txt"

    FileAppend("Timestamp: " formattedTime "`nQuery: `n" query "`n`n", logFile)

    MsgBox("SQL query successfully exported.")
}

FormatBarcodes(barcodes)
{
    barcodesList := ""
    for barcode in StrSplit(barcodes, "`n")
    {
        barcode := Trim(barcode) 
        if (barcode != "")
            barcodesList .= "'" barcode "', "
    }
    barcodesList := RTrim(barcodesList, ", ")
    return barcodesList
}

^p:: {
    global IsPaused, ButtonLU, IsPausedMessageShown, ButtonSA 
    IsPaused := !IsPaused
    if IsPaused
    {
        SoundPlay("C:\Windows\Media\Windows Notify System Generic.wav", "Async")
        MsgBox("Archiving paused. Use CTRL + P to resume.", "(!) Paused (!)")

        ButtonLU.Enabled := false
        ButtonSA.Enabled := false
        IsPausedMessageShown := true  
    }
    else
    {
        Tooltip("Archiving Resumed", , , 1)
        IsPausedMessageShown := false  
    }
    Sleep(1500)
    Tooltip("") 
}

CheckForUpdates()
{
    global githubScriptUrl, tempFilePath, currentScriptContent, latestScriptContent

    ; Ensure 'currentScriptContent' is initialized as an empty string
    currentScriptContent := ""

    Try
    {
        ; Create a COM object to download the file using WinHttpRequest
        http := ComObject("WinHttp.WinHttpRequest.5.1")
        http.Open("GET", githubScriptUrl, false)  ; Use GET method to fetch the file
        http.Send()  ; Send the request

        ; Check if the request was successful
        If (http.Status = 200)
        {
            ; Save the downloaded content to a temporary file
            FileAppend(http.ResponseText, tempFilePath)

            ; Now, assign the downloaded content to 'latestScriptContent'
            latestScriptContent := http.ResponseText
        }
        Else
        {
            MsgBox("Failed to download the script. Status: " . http.Status)
            Return
        }
    }
    Catch
    {
        MsgBox("An error occurred while downloading the script: ")
        Return
    }

    ; Read the current script content using FileRead
    currentScriptContent := FileRead(A_ScriptFullPath)  ; This reads the content of the current script itself

    ; Ensure that the variable is being compared only if the content is actually different
    If (latestScriptContent != currentScriptContent)
    {
        MsgBox("An update is available. Replacing the current script...")

        ; Overwrite the current script with the latest version from the GitHub repository
        FileDelete(A_ScriptFullPath)  ; Delete the current script
        FileAppend(latestScriptContent, A_ScriptFullPath)  ; Save the new script content

        MsgBox("The script has been updated. Restarting...")

        ; Restart the updated script
        Run(A_ScriptFullPath)
        ExitApp
    }
    Else
    {
        MsgBox("You are using the latest version of the script.")
    }
}



^q::ExitApp
