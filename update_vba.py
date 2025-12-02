import win32com.client
import os
import sys

# The corrected VBA Code
VBA_CODE = r'''Option Explicit
' ================================
' SAFI LAB - Full Patient Report System (Modern Outlook Version)
' Generates HTML, PDF, QR codes, Send Email buttons (report link only)
' Each patient row has a Generate button to do all actions individually
' ================================
Const OUTPUT_FOLDER_NAME As String = "QR_Patients" ' Folder name relative to workbook
Const DOMAIN_HOST As String = "safi-lab-8im.pages.dev"
Const COL_LINK As String = "M"
Const COL_QR As String = "N"
Const COL_BTN As String = "O"
Const COL_GENBTN As String = "P"
Const QR_SIZE As Long = 150

' ================================
' Main entrypoint - Generate all patients
Sub Generate_All_Patient_Reports_Final()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Patients")
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lastRow
        Generate_One_Patient ws, r
    Next r
    MsgBox "All patient reports generated.", vbInformation
End Sub

' ================================
' NEW: Entry point for Python
Sub Generate_From_Python(pid As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Patients")
    Dim r As Long
    r = FindRowByID(ws, pid)
    If r > 0 Then
        Generate_One_Patient ws, r
    Else
        Err.Raise vbObjectError + 513, "Generate_From_Python", "Patient ID " & pid & " not found."
    End If
End Sub

Function FindRowByID(ws As Worksheet, pid As String) As Long
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, "A").Value)) = pid Then
            FindRowByID = r
            Exit Function
        End If
    Next r
    FindRowByID = 0
End Function

' ================================
' Generate a single patient report row
Sub Generate_One_Patient(ws As Worksheet, r As Long)
    Dim pid As String, pname As String, age As String, gender As String
    Dim clinic As String, doctor As String, sample_date As String
    Dim phone As String, emailAddr As String, abss As String, conc As String, trans As String
    Dim html As String, fileHtml As String, filePdf As String
    Dim qrUrl As String, safeName As String, patientFolder As String
    Dim shp As Shape, picSh As Shape, pastedPic As Shape, btn As Shape
    Dim rootPath As String

    pid = Trim(CStr(ws.Cells(r, "A").Value))
    pname = Trim(CStr(ws.Cells(r, "B").Value))
    age = Trim(CStr(ws.Cells(r, "C").Value))
    gender = Trim(CStr(ws.Cells(r, "D").Value))
    clinic = Trim(CStr(ws.Cells(r, "E").Value))
    doctor = Trim(CStr(ws.Cells(r, "F").Value))
    sample_date = Trim(CStr(ws.Cells(r, "G").Value))
    phone = Trim(CStr(ws.Cells(r, "H").Value))
    emailAddr = Trim(CStr(ws.Cells(r, "I").Value))
    abss = Trim(CStr(ws.Cells(r, "J").Value))
    conc = Trim(CStr(ws.Cells(r, "K").Value))
    trans = Trim(CStr(ws.Cells(r, "L").Value))

    safeName = MakeSafeFileName(pname & "_" & pid)
    
    ' DYNAMIC PATH FIX
    rootPath = ThisWorkbook.Path & "\" & OUTPUT_FOLDER_NAME & "\"
    If Dir(rootPath, vbDirectory) = "" Then MkDir rootPath
    
    patientFolder = rootPath & safeName & "\"
    If Dir(patientFolder, vbDirectory) = "" Then MkDir patientFolder

    ' Build HTML content
    html = BuildHTML_Patient(pid, pname, age, gender, clinic, doctor, sample_date, phone, emailAddr, abss, conc, trans)
    fileHtml = patientFolder & "patient_" & pid & ".html"
    filePdf = patientFolder & "patient_" & pid & ".pdf"
    WriteTextFile_UTF8 fileHtml, html
    On Error Resume Next
    ConvertHTMLToPDF_UsingWord fileHtml, filePdf
    On Error GoTo 0

    ' Generate QR Code URL
    qrUrl = "https://api.qrserver.com/v1/create-qr-code/?size=200x200&data=" & _
             SAFI_UrlEncode("https://" & DOMAIN_HOST & "/" & safeName & "/patient_" & pid & ".html")

    ' Insert hyperlink in sheet
    On Error Resume Next
    ws.Hyperlinks.Add Anchor:=ws.Cells(r, COL_LINK), _
                      Address:="https://" & DOMAIN_HOST & "/" & safeName & "/patient_" & pid & ".html", _
                      TextToDisplay:="Open Report"
    On Error GoTo 0

    ' Insert QR Code as picture in cell
    On Error Resume Next
    ws.Shapes("QR_" & r).Delete
    Set picSh = ws.Shapes.AddPicture(qrUrl, False, True, 0, 0, QR_SIZE, QR_SIZE)
    picSh.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    picSh.Delete
    ws.Cells(r, COL_QR).Select
    ws.Paste
    On Error Resume Next
    Set pastedPic = Selection.ShapeRange(1)
    On Error GoTo 0
    If Not pastedPic Is Nothing Then
        With pastedPic
            .name = "QR_" & r
            .LockAspectRatio = msoTrue
            .Width = 70
            .Height = 70
            .Left = ws.Cells(r, COL_QR).Left + (ws.Cells(r, COL_QR).Width - .Width) / 2
            .Top = ws.Cells(r, COL_QR).Top + (ws.Cells(r, COL_QR).Height - .Height) / 2
            .Placement = xlMoveAndSize
        End With
    End If

    ' Save QR Code locally
    DownloadFile qrUrl, patientFolder & "qr_" & pid & ".png"

    ' -------------------------
    ' Send Email button
    On Error Resume Next
    ws.Shapes("SendBtn_" & r).Delete
    On Error GoTo 0
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Cells(r, COL_BTN).Left + 4, ws.Cells(r, COL_BTN).Top + 2, 110, 24)
    With btn
        .name = "SendBtn_" & r
        .TextFrame2.TextRange.text = "Send Email"
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Fill.ForeColor.RGB = RGB(0, 122, 204)
        .Line.ForeColor.RGB = RGB(0, 90, 170)
        .OnAction = "SendBtn_Action"
        .AlternativeText = "https://" & DOMAIN_HOST & "/patient_" & pid & ".html"
    End With

    ' -------------------------
    ' Generate Report button
    On Error Resume Next
    ws.Shapes("GenBtn_" & r).Delete
    On Error GoTo 0
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Cells(r, COL_GENBTN).Left + 4, ws.Cells(r, COL_GENBTN).Top + 2, 110, 24)
    With btn
        .name = "GenBtn_" & r
        .TextFrame2.TextRange.text = "Generate Report"
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Fill.ForeColor.RGB = RGB(34, 177, 76)
        .Line.ForeColor.RGB = RGB(0, 90, 170)
        .OnAction = "Generate_Single_Patient"
    End With
End Sub

' ================================
' Button to generate a single patient row
Sub Generate_Single_Patient()
    Dim ws As Worksheet, shpName As String, r As Long
    Set ws = ThisWorkbook.Sheets("Patients")
    shpName = Application.Caller
    r = CLng(Replace(shpName, "GenBtn_", ""))
    Generate_One_Patient ws, r
    MsgBox "Patient report generated for row " & r, vbInformation
End Sub

' ================================
' Send Email via Outlook - Modified to accept caller
Sub Handler_SendButton(shpName As String)
    Dim ws As Worksheet, rowClicked As Long
    Dim recipient As String, patientName As String, reportLink As String
    Dim mailtoLink As String, pid As String, safeName As String
    Dim bodyText As String

    Set ws = ThisWorkbook.Sheets("Patients")
    If Left(shpName, 8) <> "SendBtn_" Then Exit Sub
    rowClicked = CLng(Replace(shpName, "SendBtn_", ""))
    recipient = Trim(CStr(ws.Cells(rowClicked, "I").Value))
    patientName = Trim(CStr(ws.Cells(rowClicked, "B").Value))
    pid = Trim(CStr(ws.Cells(rowClicked, "A").Value))
    If recipient = "" Then Exit Sub

    safeName = Replace(patientName, " ", "%20")
    reportLink = "https://" & DOMAIN_HOST & "/" & safeName & "_" & pid & "/patient_" & pid & ".html"

    bodyText = "Dear " & patientName & "," & vbCrLf & vbCrLf & _
               "You can access your SAFI LAB report by clicking the link below:" & vbCrLf & _
               reportLink & vbCrLf & vbCrLf & "Best regards," & vbCrLf & "SAFI LAB Team"

    mailtoLink = "mailto:" & recipient & _
                 "?subject=" & SAFI_UrlEncode("SAFI LAB - Your Test Report") & _
                 "&body=" & SAFI_UrlEncode(bodyText)

    Application.Wait (Now + TimeValue("0:00:02"))
    ActiveWorkbook.FollowHyperlink mailtoLink
End Sub

' ================================
' Relay macro for SendBtn
Sub SendBtn_Action()
    Dim shpName As String
    shpName = Application.Caller
    Handler_SendButton shpName
End Sub

' ================================
' HTML Builder
Function BuildHTML_Patient(pid As String, pname As String, age As String, gender As String, clinic As String, _
                           doctor As String, sample_date As String, phone As String, emailAddr As String, _
                           abss As String, conc As String, trans As String) As String
    Dim html As String
    Dim icoPerson As String: icoPerson = "https://twemoji.maxcdn.com/v/latest/72x72/1f464.png"
    Dim icoId As String: icoId = "https://twemoji.maxcdn.com/v/latest/72x72/1f4c3.png"
    Dim icoAge As String: icoAge = "https://twemoji.maxcdn.com/v/latest/72x72/1f382.png"
    Dim icoGender As String: icoGender = "https://twemoji.maxcdn.com/v/latest/72x72/26a7.png"
    Dim icoClinic As String: icoClinic = "https://twemoji.maxcdn.com/v/latest/72x72/1f3e5.png"
    Dim icoDoctor As String: icoDoctor = "https://twemoji.maxcdn.com/v/latest/72x72/1f468-200d-2695-fe0f.png"
    Dim icoCal As String: icoCal = "https://twemoji.maxcdn.com/v/latest/72x72/1f4c5.png"
    Dim icoPhone As String: icoPhone = "https://twemoji.maxcdn.com/v/latest/72x72/1f4de.png"
    Dim icoMail As String: icoMail = "https://twemoji.maxcdn.com/v/latest/72x72/1f4e7.png"
    Dim icoLab As String: icoLab = "https://twemoji.maxcdn.com/v/latest/72x72/1f52c.png"

    html = "<!doctype html><html lang='en'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'>"
    html = html & "<title>SAFI LAB - " & EscapeHtml(pname) & "</title>"
    html = html & "<style>"
    html = html & "body{font-family:Inter,Segoe UI,Arial,sans-serif;background:#f8fbff;margin:0;padding:28px;color:#1b2733}"
    html = html & ".card{max-width:820px;margin:0 auto;background:#fff;border-radius:12px;padding:26px;box-shadow:0 12px 36px rgba(4,22,46,0.06)}"
    html = html & "header{display:flex;align-items:center;justify-content:center;flex-direction:column;margin-bottom:8px}"
    html = html & "h1{color:#063970;margin:6px 0;font-size:22px}"
    html = html & ".subtitle{color:#0b66a3;font-size:14px;margin-bottom:10px}"
    html = html & ".section-title{display:flex;align-items:center;color:#074a8a;font-weight:700;margin-top:18px;margin-bottom:8px}"
    html = html & ".icon{width:20px;height:20px;margin-right:10px;opacity:0.95}"
    html = html & "table{width:100%;border-collapse:collapse;font-size:15px}"
    html = html & "td{padding:10px;border-bottom:1px solid #eef6ff}"
    html = html & "td.label{width:32%;font-weight:700;color:#233b4d}"
    html = html & ".val{color:#0b2f4a}"
    html = html & ".val-strong{color:#007a3d;font-weight:700}"
    html = html & "footer{font-size:13px;color:#6b7b86;text-align:center;margin-top:16px}"
    html = html & "</style></head><body>"
    html = html & "<div class='card'>"
    html = html & "<header><img src='" & icoLab & "' style='width:36px;height:36px;'/><h1>SAFI LAB - Patient Report</h1><div class='subtitle'>Professional Laboratory Report</div></header>"
    html = html & "<div class='section-title'><img src='" & icoId & "' class='icon'/>Identification</div>"
    html = html & "<table>"
    html = html & "<tr><td class='label'><img src='" & icoId & "' class='icon'/> Patient ID</td><td class='val'>" & EscapeHtml(pid) & "</td></tr>"
    html = html & "<tr><td class='label'><img src='" & icoPerson & "' class='icon'/> Full Name</td><td class='val'>" & EscapeHtml(pname) & "</td></tr>"
    html = html & "<tr><td class='label'><img src='" & icoAge & "' class='icon'/> Age</td><td class='val'>" & EscapeHtml(age) & " years</td></tr>"
    html = html & "<tr><td class='label'><img src='" & icoGender & "' class='icon'/> Gender</td><td class='val'>" & EscapeHtml(gender) & "</td></tr>"
    html = html & "<tr><td class='label'><img src='" & icoClinic & "' class='icon'/> Clinic</td><td class='val'>" & EscapeHtml(clinic) & "</td></tr>"
    html = html & "<tr><td class='label'><img src='" & icoDoctor & "' class='icon'/> Doctor</td><td class='val'>" & EscapeHtml(doctor) & "</td></tr>"
    html = html & "<tr><td class='label'><img src='" & icoCal & "' class='icon'/> Sample Date</td><td class='val'>" & EscapeHtml(sample_date) & "</td></tr>"
    html = html & "<tr><td class='label'><img src='" & icoPhone & "' class='icon'/> Phone</td><td class='val'>" & EscapeHtml(phone) & "</td></tr>"
    html = html & "<tr><td class='label'><img src='" & icoMail & "' class='icon'/> Email</td><td class='val'>" & EscapeHtml(emailAddr) & "</td></tr>"
    html = html & "</table>"
    html = html & "<div class='section-title'><img src='" & icoLab & "' class='icon'/> Test Results</div>"
    html = html & "<table>"
    html = html & "<tr><td class='label'>ABS</td><td class='val-strong'>" & EscapeHtml(abss) & "</td></tr>"
    html = html & "<tr><td class='label'>CONC</td><td class='val-strong'>" & EscapeHtml(conc) & "</td></tr>"
    html = html & "<tr><td class='label'>TRANS</td><td class='val-strong'>" & EscapeHtml(trans) & "</td></tr>"
    html = html & "</table>"
    html = html & "<footer>© " & Year(Date) & " SAFI LAB - Confidential</footer>"
    html = html & "</div></body></html>"
    BuildHTML_Patient = html
End Function

' ================================
' Utilities
Function MakeSafeFileName(ByVal name As String) As String
    Dim badChars As Variant, ch As Variant
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In badChars
        name = Replace(name, ch, "_")
    Next ch
    MakeSafeFileName = name
End Function

Sub WriteTextFile_UTF8(filePath As String, txt As String)
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(filePath, True, True)
    ts.Write txt
    ts.Close
End Sub

Function EscapeHtml(s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    EscapeHtml = s
End Function

Function SAFI_UrlEncode(s As String) As String
    Dim i As Long, ch As String, outS As String, hexVal As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case AscW(ch)
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126
                outS = outS & ch
            Case 32
                outS = outS & "%20"
            Case Else
                hexVal = Hex$(AscW(ch))
                If Len(hexVal) = 1 Then hexVal = "0" & hexVal
                outS = outS & "%" & hexVal
        End Select
    Next i
    SAFI_UrlEncode = outS
End Function

Sub DownloadFile(fileURL As String, saveAsPath As String)
    Dim XMLHTTP As Object, stream As Object
    Set XMLHTTP = CreateObject("MSXML2.XMLHTTP")
    XMLHTTP.Open "GET", fileURL, False
    XMLHTTP.Send
    If XMLHTTP.Status = 200 Then
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 1 'binary
        stream.Open
        stream.Write XMLHTTP.responseBody
        stream.SaveToFile saveAsPath, 2
        stream.Close
    End If
End Sub

Sub ConvertHTMLToPDF_UsingWord(fileHtml As String, filePdf As String)
    Dim wdApp As Object, wdDoc As Object
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set wdDoc = wdApp.Documents.Open(fileHtml)
    wdDoc.ExportAsFixedFormat filePdf, 17 'wdExportFormatPDF
    wdDoc.Close False
    wdApp.Quit
End Sub
'''

def update_excel_vba():
    excel_path = os.path.abspath("Patients.xlsm")
    if not os.path.exists(excel_path):
        print(f"Error: {excel_path} not found.")
        return

    print(f"Opening {excel_path}...")
    try:
        xl = win32com.client.Dispatch("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        wb = xl.Workbooks.Open(excel_path)
        
        print("Accessing VBA Project...")
        # This requires "Trust access to the VBA project object model" to be enabled in Excel
        try:
            vb_component = wb.VBProject.VBComponents("Module1")
            print("Found Module1, replacing code...")
            code_module = vb_component.CodeModule
            
            # Delete existing code
            num_lines = code_module.CountOfLines
            if num_lines > 0:
                code_module.DeleteLines(1, num_lines)
            
            # Add new code
            code_module.AddFromString(VBA_CODE)
            print("VBA Code updated successfully.")
            
            wb.Save()
            print("Workbook saved.")
        except Exception as e:
            print(f"Failed to update VBA: {e}")
            print("Ensure 'Trust access to the VBA project object model' is enabled in Excel Trust Center.")
        
        wb.Close()
        xl.Quit()
        print("Done.")
        
    except Exception as e:
        print(f"Excel Error: {e}")

if __name__ == "__main__":
    update_excel_vba()
