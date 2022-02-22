Attribute VB_Name = "modShaderUtil"
Option Explicit

Global g_Current_Shader As Long
Global g_Current_Shader_Texture_Unit As Long

Private Sub PrintShaderError(Shader As Long)
Dim Compiled As Long
Dim InfoLog As String, ErrLen As Long

glGetObjectParameterivARB Shader, GL_COMPILE_STATUS, Compiled
glGetShaderiv Shader, GL_INFO_LOG_LENGTH, ErrLen
InfoLog = String(ErrLen, 0)
glGetInfoLogARB Shader, ErrLen, ErrLen, InfoLog
InfoLog = Replace(Replace(Replace(InfoLog, vbLf, vbCrLf), vbCr & vbCrLf, vbCrLf), vbNullChar, "")
If Compiled Then
    frmMain.Output InfoLog
    frmMain.Output "[INFO] Compiled successfully."
Else
    frmMain.Output InfoLog
    frmMain.Output "[WARN] Failed to compile. "
    glDeleteShader Shader
    Shader = 0
End If
End Sub

Private Sub PrintShaderLinkageError(Prog As Long)
Dim Linked As Long
Dim InfoLog As String, ErrLen As Long

glGetProgramiv Prog, GL_LINK_STATUS, Linked
glGetProgramiv Prog, GL_INFO_LOG_LENGTH, ErrLen
InfoLog = String(ErrLen, 0)
glGetInfoLogARB Prog, ErrLen, ErrLen, InfoLog
InfoLog = Replace(Replace(Replace(InfoLog, vbLf, vbCrLf), vbCr & vbCrLf, vbCrLf), vbNullChar, "")
If Linked Then
    frmMain.Output InfoLog
    frmMain.Output "[INFO] Linked successfully."
Else
    frmMain.Output InfoLog
    frmMain.Output "[WARN] Failed to link. "
    glDeleteProgram Prog
    Prog = 0
End If
End Sub

Function CreateShaderProgram(VS_code As String, FS_code As String, Optional VS_Name As String = "VS", Optional FS_Name As String = "FS") As Long
If GL_VERSION_2_0 = 0 Then
    frmMain.Output "[WARN] OpenGL 2.0 not supported."
    GoTo Exit_Return
End If

Dim name_VS As Long, name_FS As Long

frmMain.Output "[INFO] Compiling vertex shader: " & VS_Name
name_VS = glCreateShader(GL_VERTEX_SHADER)
If name_VS Then
    glShaderSource name_VS, 1, VS_code, Len(VS_code)
    glCompileShader name_VS
    PrintShaderError name_VS
End If
If name_VS = 0 Then GoTo Exit_Return

frmMain.Output "[INFO] Compiling fragment shader: " & FS_Name
name_FS = glCreateShader(GL_FRAGMENT_SHADER)
If name_FS Then
    glShaderSource name_FS, 1, FS_code, Len(FS_code)
    glCompileShader name_FS
    PrintShaderError name_FS
End If
If name_FS = 0 Then GoTo Exit_Return

If name_VS <> 0 And name_FS <> 0 Then
    frmMain.Output "[INFO] Linking shaders."
    CreateShaderProgram = glCreateProgram()
    glAttachShader CreateShaderProgram, name_VS
    glAttachShader CreateShaderProgram, name_FS
    glLinkProgram CreateShaderProgram
    PrintShaderLinkageError CreateShaderProgram
Else
    CreateShaderProgram = 0
End If

Exit_Return:
If name_VS Then glDeleteShader name_VS
If name_FS Then glDeleteShader name_FS
End Function

Function CreateShaderProgramFromFile(Path_VS As String, Path_FS As String) As Long
Dim VshStr As String, FshStr As String, FF As Long

If GL_VERSION_2_0 = 0 Then
    frmMain.Output "[WARN] OpenGL 2.0 not supported."
    GoTo Exit_Return
End If

frmMain.Output "[INFO] Loading vertex shader from path: " & Path_VS
FF = FreeFile
Open Path_VS For Input Access Read As #FF
VshStr = Input(LOF(FF), #FF)
Close #FF

frmMain.Output "[INFO] Loading fragment shader from path: " & Path_FS
FF = FreeFile
Open Path_FS For Input Access Read As #FF
FshStr = Input(LOF(FF), #FF)
Close #FF

CreateShaderProgramFromFile = CreateShaderProgram(VshStr, FshStr, (Path_VS), (Path_FS))

Exit Function
Exit_Return:
End Function

Sub SelectShader(ByVal ShaderProgram As Long)
g_Current_Shader = ShaderProgram
g_Current_Shader_Texture_Unit = 0
glUseProgram g_Current_Shader
End Sub

Sub SetShaderUniformValue(UniformName As String, ByVal Value As Double)
Dim Location As Long
If g_Current_Shader = 0 Then Exit Sub
Location = glGetUniformLocation(g_Current_Shader, UniformName)
If Location <> -1 Then
    glUniform1f Location, Value
Else
    frmMain.Output "[WARN] Location of uniform ""float " & UniformName & """ not found in the shader program."
End If
End Sub

Sub SetShaderUniformInt(UniformName As String, ByVal Value As Long)
Dim Location As Long
If g_Current_Shader = 0 Then Exit Sub
Location = glGetUniformLocation(g_Current_Shader, UniformName)
If Location <> -1 Then
    glUniform1i Location, Value
Else
    frmMain.Output "[WARN] Location of uniform ""int " & UniformName & """ not found in the shader program."
End If
End Sub

Sub SetShaderUniformVector4(UniformName As String, Vector() As Single)
Dim Location As Long
If g_Current_Shader = 0 Then Exit Sub
Location = glGetUniformLocation(g_Current_Shader, UniformName)
If Location <> -1 Then
    glUniform4f Location, Vector(0), Vector(1), Vector(2), Vector(3)
Else
    frmMain.Output "[WARN] Location of uniform ""vec4 " & UniformName & """ not found in the shader program."
End If
End Sub

Sub SetShaderUniformTexture(UniformName As String, Texture As Long)
Dim Location As Long
If g_Current_Shader = 0 Then Exit Sub
Location = glGetUniformLocation(g_Current_Shader, UniformName)

If Location <> -1 Then
    glActiveTexture GL_TEXTURE0 + g_Current_Shader_Texture_Unit
    glBindTexture GL_TEXTURE_2D, Texture
    glUniform1i Location, g_Current_Shader_Texture_Unit
    g_Current_Shader_Texture_Unit = g_Current_Shader_Texture_Unit + 1
Else
    frmMain.Output "[WARN] Location of uniform ""sampler2D " & UniformName & """ not found in the shader program."
End If
End Sub
