Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function CriaThread Lib "ThreadServer.dll" (ByVal lngIndexThread As Long, _
                                                           ByVal intPrioridade As Long, _
                                                           ByVal strCodPreEnv As String, _
                                                           ByRef objCaller As clsCallerObj) As Long

Public Declare Function FinalizarThread Lib "ThreadServer.dll" (ByVal lngIndexThread As Long) As Long
Public Declare Function ConsultarStatusThread Lib "ThreadServer.dll" (ByVal lngIndexThread As Long, ByRef strStatus As String) As Long
