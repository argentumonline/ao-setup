Attribute VB_Name = "mod_GameIni"
Option Explicit

Public Enum MODO_MEMORIA
    DX9_MEM_D = 0 'manejo de memoria por defecto
    DX9_MEM_A = 1 'manejo de memoria administrador
    DX9_MEM_S = 2 'manejo de memoria de sistema
End Enum
Public Enum API_grafica
    DX9 = 0
    DX92 = 1
    OGL = 2
    DX8 = 3
End Enum
Public Enum MODO2
    DX9_VH = 0 'usar procesamiento de vertices por hardware
    DX9_VS = 1 'usar procesamiento de vertices por software
    OGL_V = 2 'solo es una guia
End Enum

Public Enum Modo
    DX9_HARD = 0
    DX9_REF = 1
    DX9_SOF = 2
    OGL_ = 2 'solo es una guia
End Enum

Private Const INIT_PATH As String = "\INIT\"


