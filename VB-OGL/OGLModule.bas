Attribute VB_Name = "OGLModule"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   This module contains code for OpenGL initialization like Setting up PFD and RC    '
'                            Loading Textures , Lighting, etc                         '
'         (You have)CopyRight © 2003 Saadat Ali Shah, shahji_2000@yahoo.com           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'Declarations...
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Global hglRC1 As Long, hglRC2 As Long 'Rendering context handles
Global hDC1 As Long, hDC2 As Long     'HDC's of viewports
Global QObj As Long                   'GLU Quadric Object
Public Enum Heavens 'Use Enum as way of defining constants
  Sun = 1
  Mercury = 2
  Venus = 3
  Earth = 4
  Moon = 5
End Enum
'Texture variables...
Global TArray(0 To 7) As Long, TArray2(0 To 7) As Long, Bdata() As Byte
Dim hBitmap As Long, Binfo As BITMAPINFO
'Lighting variables...
Dim Amb_Dif_Light(3) As Single '= { 0.5!, 0.5!, 0.5!, 1.0! } 'ambient n diffuse light(both r same!)
Dim SpecularLight(3) As Single '= { 1.0!, 1.0!, 1.0!, 1.0! } ' specular light
Dim Light0Pos!(3), Light1Pos!(3), Light2Pos!(3)

'This Function Loads BMP files. Don't Try to Load pictures whose dimensions r not power
'of 2 e.g: 32,64,1024 etc as there is no testing for it Yet! and ur Program will Crash!
Function LoadBMP(FileName As String) As Boolean
 'load the bitmap into memory...
  hBitmap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
  If hBitmap = 0 Then MsgBox "Loading Texture File: '" & FileName & "' Failed!", vbCritical: LoadBMP = False: Exit Function
  Binfo.bmiHeader.biPlanes = 1
  Binfo.bmiHeader.biBitCount = 0
  Binfo.bmiHeader.biCompression = BI_RGB
  'Load BMP info in Binfo
  GetDIBits hDC1, hBitmap, 0, 0, ByVal 0&, ByVal VarPtr(Binfo), DIB_RGB_COLORS
  If Binfo.bmiHeader.biBitCount = 8 Then 'If Bmp is GrayScale
    Binfo.bmiHeader.biBitCount = 24 'force 24-bit colors, Load Grayscale pic as RGB
    ReDim Bdata(0 To Binfo.bmiHeader.biSizeImage * 3 - 1) 'Allocate enough space for Bdata
    'Get BMP data in Bdata
    GetDIBits hDC1, hBitmap, 0, Binfo.bmiHeader.biHeight, Bdata(0), ByVal VarPtr(Binfo), DIB_RGB_COLORS
  Else
    ReDim Bdata(0 To Binfo.bmiHeader.biHeight * Binfo.bmiHeader.biWidth * 3 - 1)
    GetDIBits hDC1, hBitmap, 0, Binfo.bmiHeader.biHeight, Bdata(0), ByVal VarPtr(Binfo), DIB_RGB_COLORS
  End If
  LoadBMP = True
End Function
Function InitGL() As Boolean
Dim Pfd As PIXELFORMATDESCRIPTOR
Dim Result As Long
  'Setup PDF n RC ...
  Pfd.nSize = Len(Pfd)
  Pfd.nVersion = 1
  Pfd.dwFlags = PFD_SUPPORT_OPENGL Or PFD_DRAW_TO_WINDOW Or PFD_DOUBLEBUFFER Or PFD_TYPE_RGBA
  Pfd.iPixelType = PFD_TYPE_RGBA
  Pfd.cColorBits = 32
  Pfd.cDepthBits = 16
  'Pfd.iLayerType = PFD_MAIN_PLANE   'used in earlier implementaions of Opengl but no longer
  Result = ChoosePixelFormat(hDC1, Pfd)
  If Result = 0 Then
      MsgBox "OpenGL Initialization Failed!", vbCritical
      InitGL = False
      Exit Function
  End If
  SetPixelFormat hDC1, Result, Pfd
  hglRC1 = wglCreateContext(hDC1)

  SetPixelFormat hDC2, Result, Pfd
  hglRC2 = wglCreateContext(hDC2)
  
  'Init OpenGL vars...
  Amb_Dif_Light(0) = 0.5!: Amb_Dif_Light(1) = 0.5!: Amb_Dif_Light(2) = 0.5!: Amb_Dif_Light(3) = 1!
  SpecularLight(0) = 1!: SpecularLight(1) = 1!: SpecularLight(2) = 1!: SpecularLight(3) = 1!
  
  Light0Pos(0) = 0!: Light0Pos(1) = 5!: Light0Pos(2) = -40!: Light0Pos(3) = 1!
  Light1Pos(0) = 0!: Light1Pos(1) = 1!: Light1Pos(2) = -19!: Light1Pos(3) = 1!
  Light2Pos(0) = 0!: Light2Pos(1) = 0!: Light2Pos(2) = -2!: Light2Pos(3) = 1!

  'Init Quadric Object...
  QObj = gluNewQuadric()
  gluQuadricTexture QObj, GL_TRUE
  gluQuadricNormals QObj, GL_SMOOTH
  
  'ViewPort 1 Specific inits...
  wglMakeCurrent hDC1, hglRC1
  glClearColor 0, 0, 0, 0  ' clear view with black color
  glShadeModel (GL_SMOOTH) 'Interpolate colors
  glEnable GL_CULL_FACE    'Do not calculate BackFace of polys
  glFrontFace GL_CCW
  glEnable GL_DEPTH_TEST
  glClearDepth 1
  glDepthFunc cfLEqual
  glBlendFunc GL_SRC_ALPHA, GL_DST_ALPHA
  glEnable GL_COLOR_MATERIAL
  
  glEnable GL_LIGHTING    'enable lighting
  
  'setup LIGHT0, Sunlight for Planets...
  glLightfv GL_LIGHT0, GL_AMBIENT, Amb_Dif_Light(0)
  glLightfv GL_LIGHT0, GL_DIFFUSE, Amb_Dif_Light(0)
  glLightfv GL_LIGHT0, GL_POSITION, Light0Pos(0) ' place the light were Sun will be
  'glLightfv GL_LIGHT0, GL_SPECULAR, SpecularLight  'Light0 has Def. Specular light of (1,1,1) so no need to add
  glEnable GL_LIGHT0  ' Enable the light

  'Setup Light1, lits Sun itself...
  glLightfv GL_LIGHT1, GL_AMBIENT, Amb_Dif_Light(0)
  glLightfv GL_LIGHT1, GL_DIFFUSE, Amb_Dif_Light(0)
  glLightfv GL_LIGHT1, GL_SPECULAR, SpecularLight(0)
  glLightfv GL_LIGHT1, GL_POSITION, Light1Pos(0)    'Place Infront of Sun
  
  'Materials r Specular...
  glMaterialfv GL_FRONT, GL_SPECULAR, SpecularLight(0)
  glMaterialf GL_FRONT, GL_SHININESS, 1!
  glEnable GL_TEXTURE_2D
  
  'ViewPort 2 Specific inits...
  wglMakeCurrent hDC2, hglRC2
  glClearColor 0, 0, 0, 0
  glShadeModel (GL_SMOOTH) 'Interpolate colors
  glEnable GL_CULL_FACE    'Do not calculate BackFace of polys
  glFrontFace GL_CCW
  glEnable GL_DEPTH_TEST
  glClearDepth 1
  glDepthFunc cfLEqual
  glBlendFunc GL_SRC_ALPHA, GL_DST_ALPHA
  glEnable GL_COLOR_MATERIAL
  
  glEnable GL_LIGHTING    'enable lighting
  
  'Setup Light2...
  glLightfv GL_LIGHT2, GL_AMBIENT, Amb_Dif_Light(0)
  glLightfv GL_LIGHT2, GL_DIFFUSE, Amb_Dif_Light(0)
  glLightfv GL_LIGHT2, GL_SPECULAR, SpecularLight(0)
  glLightfv GL_LIGHT2, GL_POSITION, Light2Pos(0)
  glEnable GL_LIGHT2
  
  'Materials r Specular
  glMaterialfv GL_FRONT, GL_SPECULAR, SpecularLight(0)
  glMaterialf GL_FRONT, GL_SHININESS, 1!
  glEnable GL_TEXTURE_2D
  
  'Textures...
  wglMakeCurrent hDC1, hglRC1
  glGenTextures 8, TArray(0)  'Gen Textures for 1st viewport
    
  Binfo.bmiHeader.biSize = Len(Binfo.bmiHeader) 'tell me my size
  'Sun's Flare...
  If (LoadBMP(App.Path & "\data\Flare.bmp")) = False Then InitGL = False: Exit Function
  glBindTexture GL_TEXTURE_2D, TArray(0)
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
  glTexImage2D GL_TEXTURE_2D, 0, 3, Binfo.bmiHeader.biWidth, Binfo.bmiHeader.biHeight, 0, GL_BGR_EXT, GL_UNSIGNED_BYTE, Bdata(0)
  wglMakeCurrent hDC2, hglRC2
  glGenTextures 8, TArray2(0) 'Gen Textures for 2nd viewport
  glBindTexture GL_TEXTURE_2D, TArray2(0)
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
  glTexImage2D GL_TEXTURE_2D, 0, 3, Binfo.bmiHeader.biWidth, Binfo.bmiHeader.biHeight, 0, GL_BGR_EXT, GL_UNSIGNED_BYTE, Bdata(0)
  DeleteObject hBitmap: Erase Bdata

  'Sun...
  LoadBMP App.Path & "\data\Sun.bmp"
  wglMakeCurrent hDC1, hglRC1
  glBindTexture GL_TEXTURE_2D, TArray(Sun)
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
  glTexImage2D GL_TEXTURE_2D, 0, 3, Binfo.bmiHeader.biWidth, Binfo.bmiHeader.biHeight, 0, GL_BGR_EXT, GL_UNSIGNED_BYTE, Bdata(0)
  wglMakeCurrent hDC2, hglRC2
  glBindTexture GL_TEXTURE_2D, TArray2(Sun)
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
  glTexImage2D GL_TEXTURE_2D, 0, 3, Binfo.bmiHeader.biWidth, Binfo.bmiHeader.biHeight, 0, GL_BGR_EXT, GL_UNSIGNED_BYTE, Bdata(0)
  DeleteObject hBitmap: Erase Bdata
  
  LoadBMP App.Path & "\data\Mercury.bmp"
  wglMakeCurrent hDC1, hglRC1
  glBindTexture GL_TEXTURE_2D, TArray(Mercury)
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
  glTexImage2D GL_TEXTURE_2D, 0, 3, Binfo.bmiHeader.biWidth, Binfo.bmiHeader.biHeight, 0, GL_BGR_EXT, GL_UNSIGNED_BYTE, Bdata(0)
  wglMakeCurrent hDC2, hglRC2
  glBindTexture GL_TEXTURE_2D, TArray2(Mercury)
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
  glTexImage2D GL_TEXTURE_2D, 0, 3, Binfo.bmiHeader.biWidth, Binfo.bmiHeader.biHeight, 0, GL_BGR_EXT, GL_UNSIGNED_BYTE, Bdata(0)
  DeleteObject hBitmap: Erase Bdata
  
  LoadBMP App.Path & "\data\Venus2.bmp"
  wglMakeCurrent hDC1, hglRC1
  glBindTexture GL_TEXTURE_2D, TArray(Venus)
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
  glTexImage2D GL_TEXTURE_2D, 0, 3, Binfo.bmiHeader.biWidth, Binfo.bmiHeader.biHeight, 0, GL_BGR_EXT, GL_UNSIGNED_BYTE, Bdata(0)
  wglMakeCurrent hDC2, hglRC2
  glBindTexture GL_TEXTURE_2D, TArray2(Venus)
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
  glTexImage2D GL_TEXTURE_2D, 0, 3, Binfo.bmiHeader.biWidth, Binfo.bmiHeader.biHeight, 0, GL_BGR_EXT, GL_UNSIGNED_BYTE, Bdata(0)
  DeleteObject hBitmap: Erase Bdata
  
  LoadBMP App.Path & "\data\Earth.bmp"
  wglMakeCurrent hDC1, hglRC1
  glBindTexture GL_TEXTURE_2D, TArray(Earth)
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
  glTexImage2D GL_TEXTURE_2D, 0, 3, Binfo.bmiHeader.biWidth, Binfo.bmiHeader.biHeight, 0, GL_BGR_EXT, GL_UNSIGNED_BYTE, Bdata(0)
  wglMakeCurrent hDC2, hglRC2
  glBindTexture GL_TEXTURE_2D, TArray2(Earth)
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
  glTexImage2D GL_TEXTURE_2D, 0, 3, Binfo.bmiHeader.biWidth, Binfo.bmiHeader.biHeight, 0, GL_BGR_EXT, GL_UNSIGNED_BYTE, Bdata(0)
  DeleteObject hBitmap: Erase Bdata
  
  LoadBMP App.Path & "\data\Moon.bmp"
  wglMakeCurrent hDC1, hglRC1
  glBindTexture GL_TEXTURE_2D, TArray(Moon)
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
  glTexImage2D GL_TEXTURE_2D, 0, 3, Binfo.bmiHeader.biWidth, Binfo.bmiHeader.biHeight, 0, GL_BGR_EXT, GL_UNSIGNED_BYTE, Bdata(0)
  wglMakeCurrent hDC2, hglRC2
  glBindTexture GL_TEXTURE_2D, TArray2(Moon)
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
  glTexImage2D GL_TEXTURE_2D, 0, 3, Binfo.bmiHeader.biWidth, Binfo.bmiHeader.biHeight, 0, GL_BGR_EXT, GL_UNSIGNED_BYTE, Bdata(0)
  DeleteObject hBitmap: Erase Bdata
  
  LoadBMP App.Path & "\data\Cloud.bmp"
  wglMakeCurrent hDC1, hglRC1
  glBindTexture GL_TEXTURE_2D, TArray(6)
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
  glTexImage2D GL_TEXTURE_2D, 0, 3, Binfo.bmiHeader.biWidth, Binfo.bmiHeader.biHeight, 0, GL_BGR_EXT, GL_UNSIGNED_BYTE, Bdata(0)
  wglMakeCurrent hDC2, hglRC2
  glBindTexture GL_TEXTURE_2D, TArray2(6)
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
  glTexImage2D GL_TEXTURE_2D, 0, 3, Binfo.bmiHeader.biWidth, Binfo.bmiHeader.biHeight, 0, GL_BGR_EXT, GL_UNSIGNED_BYTE, Bdata(0)
  DeleteObject hBitmap: Erase Bdata
  
  LoadBMP App.Path & "\data\Mhalo.bmp"
  wglMakeCurrent hDC1, hglRC1
  glBindTexture GL_TEXTURE_2D, TArray(7)
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
  glTexImage2D GL_TEXTURE_2D, 0, 3, Binfo.bmiHeader.biWidth, Binfo.bmiHeader.biHeight, 0, GL_BGR_EXT, GL_UNSIGNED_BYTE, Bdata(0)
  wglMakeCurrent hDC2, hglRC2
  glBindTexture GL_TEXTURE_2D, TArray2(7)
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR
  glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR
  glTexImage2D GL_TEXTURE_2D, 0, 3, Binfo.bmiHeader.biWidth, Binfo.bmiHeader.biHeight, 0, GL_BGR_EXT, GL_UNSIGNED_BYTE, Bdata(0)
  DeleteObject hBitmap: Erase Bdata

  InitGL = True
End Function
