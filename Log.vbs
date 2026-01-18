Option Explicit

Dim LOG_FILE
Dim scriptFolder
Dim scriptBaseName
' 获取当前脚本所在目录
scriptFolder = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
' 获取当前脚本的文件名（不含路径和扩展名）
scriptBaseName = CreateObject("Scripting.FileSystemObject").GetBaseName(WScript.ScriptFullName)
' 日志文件 = 脚本同目录 + 脚本名称（不含 .vbs） + .log
LOG_FILE = scriptFolder & "\" & scriptBaseName & ".log"

' 全局对象
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

' ======================
' 核心日志写入函数（内部使用）
' ======================
Private Sub WriteLog(msg, level)
    Dim prefix, line
    Dim logStream
    Dim tsYear, tsMonth, tsDay, tsTime
    
    ' 级别前缀
    Select Case UCase(level)
        Case "ERROR"   : prefix = "[ERROR]"
        Case "WARN"    : prefix = "[WARN ]"
        Case "INFO"    : prefix = "[INFO ]"
        Case "DEBUG"   : prefix = "[DEBUG]"
        Case Else      : prefix = "[?????]"
    End Select
    
    ' 手动格式化日期为 2026-01-18（年-月-日，月日补零）
    tsYear  = Year(Now)
    tsMonth = Right("0" & Month(Now), 2)
    tsDay   = Right("0" & Day(Now), 2)
    ' 时间部分保持原样（时:分:秒）
    tsTime = FormatDateTime(Now, vbLongTime)
    
    ' 拼接完整一行（始终带时间戳）
    line = tsYear & "-" & tsMonth & "-" & tsDay & " " & tsTime & " " & prefix & " " & msg
    
    ' 写文件（每次重新打开 → 写 → 立即关闭）
    On Error Resume Next
    Set logStream = fso.OpenTextFile(LOG_FILE, 8, True)   ' 8 = ForAppending, True = 创建如果不存在
    If Not logStream Is Nothing Then
        logStream.WriteLine line
        logStream.Close
        Set logStream = Nothing
    End If
    On Error Goto 0
    
    ' 同时输出到控制台（始终开启）
    WScript.Echo line
End Sub

' ======================
' 便捷调用函数（推荐使用这些）
' ======================
Sub LogError(msg)
    WriteLog msg, "ERROR"
End Sub

Sub LogWarn(msg)
    WriteLog msg, "WARN"
End Sub

Sub LogInfo(msg)
    WriteLog msg, "INFO"
End Sub

Sub LogDebug(msg)
    WriteLog msg, "DEBUG"
End Sub

